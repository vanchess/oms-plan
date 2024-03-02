<?php

use App\Enum\MedicalServicesEnum;
use App\Http\Controllers\CategoryTreeController;
use App\Http\Controllers\PlanReports;
use App\Jobs\InitialChanges;
use App\Jobs\InitialDataLoaded;
use App\Models\CareProfiles;
use App\Models\CareProfilesFoms;
use App\Models\ChangePackage;
use App\Models\CommissionDecision;
use App\Models\HospitalBedProfiles;
use App\Models\Indicator;
use App\Models\IndicatorType;
use App\Models\InitialData;
use App\Models\MedicalAssistanceType;
use Illuminate\Support\Facades\Route;
use App\Models\MedicalInstitution;
use App\Models\MedicalServices;
use App\Models\OmsProgram;
use App\Models\Organization;
use App\Models\Period;

use App\Services\InitialDataService;
use App\Models\PlannedIndicator;
use App\Models\PlannedIndicatorChange;
use App\Models\PumpMonitoringProfiles;
use App\Models\PumpMonitoringProfilesRelationType;
use App\Models\PumpMonitoringProfilesUnit;
use App\Models\PumpUnit;
use App\Models\VmpGroup;
use App\Models\VmpTypes;
use App\Services\DataForContractService;
//use App\Services\NodeService;
use App\Services\Dto\InitialDataValueDto;
use App\Services\InitialDataFixingService;
use App\Services\MoDepartmentsInfoForContractService;
use App\Services\MoInfoForContractService;
use App\Services\PeopleAssignedInfoForContractService;
use App\Services\PlannedIndicatorChangeInitService;
use App\Services\PlanReports\PlanCalculatorService;
use App\Services\PumpMonitoringProfilesTreeService;
use App\Services\RehabilitationProfileService;
use Illuminate\Database\Eloquent\Builder;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Auth;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\Storage;
use PhpOffice\PhpSpreadsheet\Calculation\TextData\Trim;
use PhpOffice\PhpSpreadsheet\Shared\StringHelper;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/

Route::get('/createPeriods/{year}', function (int $year) {
    for ($i = 1; $i <= 12; $i++) {
        $monthNum = str_pad($i, 2, 0, STR_PAD_LEFT);
        $month = new DateTime("{$year}-{$monthNum}-01T00:00:00.000000+0500");
        $from = new DateTime("first day of {$month->format('F')} {$year}+0500");
        $to = new DateTime("last day of {$month->format('F')} {$year}T23:59:59.999999+0500");
        $from->setTimezone(new DateTimeZone('UTC'));
        $to->setTimezone(new DateTimeZone('UTC'));
        Period::firstOrCreate(['from' => $from, 'to' => $to]);
    }
    return 'ok';
});

function printTree($tree, $treeService, $l = 0) {
    if ($tree === null) return;
    foreach ($tree as $k => $t) {
        $mp = PumpMonitoringProfiles::find($k);
        $pu = $mp->profilesUnits;
        echo str_repeat("\t", $l) . $mp->name . ' [' . PumpMonitoringProfilesRelationType::find($mp->relation_type_id)?->name . "]\r\n";

        foreach($pu as $u) {
            $pi = $u->plannedIndicators;
            $piIdsViaChild = $treeService->plannedIndicatorIdsViaChild($mp->id, $u->unit->id);
            $piIds = array_column($pi?->ToArray(), 'id');
            $piIdsViaChild = array_diff($piIdsViaChild, $piIds);
            $piViaChild = PlannedIndicator::whereIn('id', $piIdsViaChild)->get();

            if (count($pi) > 0 || count($piIdsViaChild) > 0) {
                foreach($pi as $i) {
                    $piName = plannedIndicatorName($i);
                    echo str_repeat("\t", $l) . "-{$u->unit->name} |{$piName}|\r\n";
                }
                foreach($piViaChild as $i) {
                    $piName = plannedIndicatorName($i);
                    echo str_repeat("\t", $l) . "-{$u->unit->name} |{$piName}| (УНАСЛЕДОВАНО)\r\n";
                }
            } else {
                echo str_repeat("\t", $l) . "-{$u->unit->name} | не утверждается | \r\n";
            }
        }

        printTree($t, $treeService, $l + 1);
    }
}

function plannedIndicatorName($i) {
    return "({$i->id}) {$i->assistanceType?->name} {$i->careProfile?->name} {$i->indicator?->name} {$i->bedProfile?->name} {$i->service?->name} {$i->vmpGroup?->code} {$i->vmpType?->name} {$i->node->nodePath()}";
}

Route::get('/pump-plan', function (PumpMonitoringProfilesTreeService $treeService) {
    /*
    $baseOmsProgram = OmsProgram::where('name','базовая')->first();
    $monitoringProfiles = PumpMonitoringProfiles::where('oms_program_id',$baseOmsProgram->id)->get();
    foreach ($monitoringProfiles as $p) {
        echo $p->name . '<br>';
    }
*/
    $tree = $treeService->nodeTree(1);
    printTree($tree, $treeService);
    // dd($tree);
    return 'ok';
});

/**
 * Заполнить связь профилей мониторинга ПУМП и наших плановых показателей
 */
Route::get('/fill-pump-monitoring-profiles-planned-indicators-relationships', function () {
    $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader("Xlsx");
    $path = 'xlsx/pump';
    $templateFileName = 'PumpPgg2.xlsx';
    $templateFilePath = $path . DIRECTORY_SEPARATOR . $templateFileName;
    $templateFullFilepath = Storage::path($templateFilePath);

    $spreadsheet = $reader->load($templateFullFilepath);
    $sheet = $spreadsheet->getActiveSheet();
    $startRow = 8;
    $endRow = 548;

    $monitoringProfileCodeCol = 2;
    $monitoringProfileNameCol = 1;
    $plannedIndicatorIdCol = 3;

    $typeFinId = IndicatorType::where('name', 'money')->first()->id;
    $typeQuantId = IndicatorType::where('name', 'volume')->first()->id;

    $num = 1;
    for ($i=$startRow; $i <= $endRow; $i++) {
        $name = trim($sheet->getCell([$monitoringProfileNameCol, $i])->getValue());
        $code = trim($sheet->getCell([$monitoringProfileCodeCol, $i])->getValue());
        $indicatorsTemp = explode(',', trim($sheet->getCell([$plannedIndicatorIdCol, $i])->getValue()));
        $indicators = [];
        for ($k = 0; $k < count($indicatorsTemp); $k++) {
            $ind = trim($indicatorsTemp[$k]);
            if ($ind !== '') {
                array_push($indicators, $ind);
            }
        }
        $indicators = array_unique($indicators, SORT_NUMERIC);
        $monitoringProfile = PumpMonitoringProfiles::where('code', $code)->first();
        $t = str_starts_with($name, $monitoringProfile->name);

        $p = 'ERROR';
        $monitoringProfileUnits = null;
        if (str_ends_with($name, '(сумма)')) {
            $p = 'финансовая часть';
            $monitoringProfileUnits = $monitoringProfile->profilesUnits()->whereHas('unit', function (Builder $query) use ($typeFinId) {
                    $query->where('type_id', $typeFinId);
                })->get();
        } else if (str_ends_with($name, '(кол-во)')) {
            $p = 'количественная часть';
            $monitoringProfileUnits = $monitoringProfile->profilesUnits()->whereHas('unit', function (Builder $query) use ($typeQuantId) {
                $query->where('type_id', $typeQuantId);
            })->get();
        }
        foreach ($monitoringProfileUnits as $u) {
            echo $num++ . ') ' . ($t ? 'OK  ' : 'ERROR  ') . $p . ' ' . $u->unit->name . ' ' . $monitoringProfile->name . ' ' . $code . '<br>';
            if (count($indicators) > 0) {
                $u->plannedIndicators()->attach($indicators);
            }
        }
    }
/**/
    return 'ок';
    //phpinfo();
});

/**
 * Заполнить профили мониторинга ПУМП из файла
 */
Route::get('/fill-pump-monitoring-profiles', function () {
    $FIN_PROFILE_TYPE_STR = "только финансовая часть";
    $FIN_QUANT_PROFILE_TYPE_STR = "финансовая и количественная часть";


    OmsProgram::firstOrCreate(['name' => 'базовая']);
    OmsProgram::firstOrCreate(['name' => 'сверхбазовая']);

    PumpUnit::firstOrCreate(['name' => 'вызов', 'type_id' => 1]);
    PumpUnit::firstOrCreate(['name' => 'законченный случай', 'type_id' => 1]);
    PumpUnit::firstOrCreate(['name' => 'исследование', 'type_id' => 1]);
    PumpUnit::firstOrCreate(['name' => 'койко-день', 'type_id' => 1]);
    PumpUnit::firstOrCreate(['name' => 'комплексное посещение', 'type_id' => 1]);
    PumpUnit::firstOrCreate(['name' => 'обращение', 'type_id' => 1]);
    PumpUnit::firstOrCreate(['name' => 'посещение', 'type_id' => 1]);
    PumpUnit::firstOrCreate(['name' => 'случай госпитализации', 'type_id' => 1]);
    PumpUnit::firstOrCreate(['name' => 'случай лечения', 'type_id' => 1]);
    PumpUnit::firstOrCreate(['name' => 'услуга', 'type_id' => 1]);
    PumpUnit::firstOrCreate(['name' => 'стоимость', 'type_id' => 2]);

    PumpMonitoringProfilesRelationType::firstOrCreate(['name' => 'сумма', 'slug' => 'sum']);
    PumpMonitoringProfilesRelationType::firstOrCreate(['name' => 'в том числе', 'slug' => 'including']);

    $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader("Xlsx");
    $path = 'xlsx/pump';
    $templateFileName = 'PumpMonitoringProfiles_v3.xlsx';
    $templateFilePath = $path . DIRECTORY_SEPARATOR . $templateFileName;
    $templateFullFilepath = Storage::path($templateFilePath);

    $spreadsheet = $reader->load($templateFullFilepath);
    $sheet = $spreadsheet->getActiveSheet();

    $omsProgramCol = 1;
    // $recIdCol = 2;
    // $parentRecIdCol = 3;
    $monitoringProfileCodeCol = 4;
    $monitoringProfileParentCodeCol = 5;
    $monitoringProfileShortNameCol = 6;
    $monitoringProfileNameCol = 7;
    $parentRelationCol = 8;
    $monitoringProfileTypeCol = 9;
    $unitCol = 10;

    $iterator = $sheet->getRowIterator();
    $iterator->next();
    while ($iterator->valid()) {
        $row = $iterator->current();
        // ПрограммаОМС ИД ИДродителя Код КодРодителя КраткоеНаименование ПолноеНаименование ОтношениеКРодителю СоставПоказателя КоличественнаяЕдиница
        $p = new PumpMonitoringProfiles();
        $omsProgramName = trim(mb_strtolower($sheet->getCell([$omsProgramCol, $row->getRowIndex()])->getValue()));
        if ($omsProgramName === '') {
            break;
        }
        $omsProgram = OmsProgram::firstOrCreate(['name' => $omsProgramName]);
        $p->oms_program_id = $omsProgram->id;
        $p->code = trim($sheet->getCell([$monitoringProfileCodeCol, $row->getRowIndex()])->getValue());
        $parentCode = trim($sheet->getCell([$monitoringProfileParentCodeCol, $row->getRowIndex()])->getValue());

        if ($parentCode != '') {
            $parent = PumpMonitoringProfiles::where('code', $parentCode)->first();
            $p->parent_id = $parent->id;
        }
        $p->short_name = trim($sheet->getCell([$monitoringProfileShortNameCol, $row->getRowIndex()])->getValue());
        $p->name = trim($sheet->getCell([$monitoringProfileNameCol, $row->getRowIndex()])->getCalculatedValue());
        $rT =  mb_strtolower($sheet->getCell([$parentRelationCol, $row->getRowIndex()])->getValue());
        if ($rT != '') {
            $relationType = PumpMonitoringProfilesRelationType::firstOrCreate(['name' => $rT]);
            $p->relation_type_id = $relationType->id;
        }
        $profileType = trim(mb_strtolower($sheet->getCell([$monitoringProfileTypeCol, $row->getRowIndex()])->getValue()));
        $p->is_leaf = false;

        $p->save();

        $unit = new PumpMonitoringProfilesUnit();
        $unit->unit_id = PumpUnit::where('name', 'стоимость')->first()->id;
        $unit->monitoring_profile_id = $p->id;
        $unit->save();

        if ($profileType === $FIN_QUANT_PROFILE_TYPE_STR) {
            $unitName = $profileType = trim(mb_strtolower($sheet->getCell([$unitCol, $row->getRowIndex()])->getValue()));
            $unit2 = new PumpMonitoringProfilesUnit();
            $unit2->unit_id = PumpUnit::where('name', $unitName)->first()->id;
            $unit2->monitoring_profile_id = $p->id;
            $unit2->save();
        }

        $iterator->next();
    }

    return 'ок';
});

Route::get('/initial_changes', function () {
    InitialChanges::dispatch(2023);
    return "initial_changes";
});

Route::get('/all_initial_data_loaded', function () {
    InitialDataLoaded::dispatch(1, 2023, 1);
    InitialDataLoaded::dispatch(9, 2023, 1);
    InitialDataLoaded::dispatch(17, 2023, 1);
    InitialDataLoaded::dispatch(39, 2023, 1); // Прикрепление Скорая
    return "all_initial_data_loaded";
});



Route::get('/meeting-minutes/{year}/{commissionDecisionsId}', function (DataForContractService $dataForContractService, int $year, int $commissionDecisionsId) {
    // $year = 2022;
    // $commissionDecisionsId = 19;
    $cd = CommissionDecision::find($commissionDecisionsId);
    $currentlyUsedDate = $cd->date->format('Y-m-d');
    $protocolDate = $cd->date->format('d.m.Y');
    $docName = "протокол заседания КРТП ОМС №$cd->number от $protocolDate";
    $packageIds = $cd->changePackage()->pluck('id')->toArray();
    $indicatorIds = [1, 2, 3, 4, 5, 6, 7, 8, 9];
    $content = $dataForContractService->GetArray($year, $packageIds, $indicatorIds);
    // количество коек на последний месяц года
    $contentNumberOfBeds = $dataForContractService->GetArrayByYearAndMonth($year, 12, $packageIds, [1]);

    $moCollection = MedicalInstitution::WhereRaw("? BETWEEN effective_from AND effective_to", [$currentlyUsedDate])->orderBy('order')->get();

    $path = 'xlsx';
    $templateFileName = 'meetingMinutes.xlsx';
    $templateFilePath = $path . DIRECTORY_SEPARATOR . $templateFileName;
    $templateFullFilepath = Storage::path($templateFilePath);
    $resultFileName = "protocol_№$cd->number($protocolDate).xlsx";
    $strDateTimeNow = date("Y-m-d-His");
    $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . ' ' . $resultFileName;
    $fullResultFilepath = Storage::path($resultFilePath);

    $startRow = 6;
    $endRow = 350;
    bcscale(4);


    $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader("Xlsx");
    $spreadsheet = $reader->load($templateFullFilepath);
    $sheet = $spreadsheet->getSheetByName('Скорая помощь');
    $sheet->setCellValue([1,3], $docName);
    $ordinalRowNum = 0;
    $rowIndex = $startRow - 1;
    $category = 'ambulance';
    $callsIndicatorId = 5; // вызовов
    $costIndicatorId = 4; // стоимость
    $callsAssistanceTypeId = 5; // вызовы
    $thrombolysisAssistanceTypeId = 6;// тромболизис
    foreach($moCollection as $mo) {
        $thrombolysis = $content['mo'][$mo->id][$category][$thrombolysisAssistanceTypeId][$callsIndicatorId] ?? 0;
        $allCalls = ($content['mo'][$mo->id][$category][$callsAssistanceTypeId][$callsIndicatorId] ?? 0) + $thrombolysis;
        $costCalls = $content['mo'][$mo->id][$category][$callsAssistanceTypeId][$costIndicatorId] ?? '0';
        $costThrombolysis = $content['mo'][$mo->id][$category][$thrombolysisAssistanceTypeId][$costIndicatorId] ?? '0';

        if( $thrombolysis === 0 && $allCalls === 0 && $costCalls === '0' && $costThrombolysis === '0' ) {continue;}

        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        $sheet->setCellValue([3,$rowIndex], $allCalls);
        //$sheet->setCellValue([9,$rowIndex], $thrombolysis);
        $sheet->setCellValue([4,$rowIndex], bcadd($costCalls, $costThrombolysis));
    }
    $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);




    $sheet = $spreadsheet->getSheetByName('Диагностика');
    $sheet->setCellValue([1,3], $docName);
    $ordinalRowNum = 0;
    $rowIndex = $startRow;
    $category = 'polyclinic';
    $indicatorId = 6; // услуг
    $costIndicatorId = 4; // стоимость

    foreach($moCollection as $mo) {
        // КТ
        $serviceId = MedicalServicesEnum::KT;
        $kt = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $indicatorId);
        $costKt = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $costIndicatorId);

        // МРТ
        $serviceId = MedicalServicesEnum::MRT;
        $mrt = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $indicatorId);
        $costMrt = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $costIndicatorId);

        // УЗИ ССС
        $serviceId = MedicalServicesEnum::UltrasoundCardio;
        $ultrasoundCardio = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $indicatorId);
        $costUltrasoundCardio = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $costIndicatorId);

        // Эндоскопия
        $serviceId = MedicalServicesEnum::Endoscopy;
        $endoscopy = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $indicatorId);
        $costEndoscopy = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $costIndicatorId);

        // ПАИ
        $serviceId = MedicalServicesEnum::PathologicalAnatomicalBiopsyMaterial;
        $pathologicalAnatomicalBiopsyMaterial = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $indicatorId);
        $costPathologicalAnatomicalBiopsyMaterial = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $costIndicatorId);

        // МГИ
        $serviceId = MedicalServicesEnum::MolecularGeneticDetectionOncological;
        $molecularGeneticDetectionOncological = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $indicatorId);
        $costMolecularGeneticDetectionOncological = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $costIndicatorId);

        // Тест.covid-19
        $serviceId = MedicalServicesEnum::CovidTesting;
        $covidTesting = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $indicatorId);
        $costCovidTesting = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $costIndicatorId);

        if( bccomp($kt,'0') === 0
            && bccomp($mrt, '0') === 0
            && bccomp($ultrasoundCardio, '0') === 0
            && bccomp($endoscopy, '0') === 0
            && bccomp($pathologicalAnatomicalBiopsyMaterial, '0') === 0
            && bccomp($molecularGeneticDetectionOncological, '0') === 0
            && bccomp($covidTesting, '0') === 0
            && bccomp($costKt, '0') === 0
            && bccomp($costMrt, '0') === 0
            && bccomp($costEndoscopy, '0') === 0
            && bccomp($costPathologicalAnatomicalBiopsyMaterial, '0') === 0
            && bccomp($costMolecularGeneticDetectionOncological, '0') === 0
            && bccomp($costCovidTesting, '0') === 0
        ) {continue;}

        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        $sheet->setCellValue([3,$rowIndex], $kt);
        $sheet->setCellValue([4,$rowIndex], $costKt);
        $sheet->setCellValue([5,$rowIndex], $mrt);
        $sheet->setCellValue([6,$rowIndex], $costMrt);
        $sheet->setCellValue([7,$rowIndex], $endoscopy);
        $sheet->setCellValue([8,$rowIndex], $costEndoscopy);
        $sheet->setCellValue([9,$rowIndex], $ultrasoundCardio);
        $sheet->setCellValue([10,$rowIndex], $costUltrasoundCardio);
        $sheet->setCellValue([11,$rowIndex], $pathologicalAnatomicalBiopsyMaterial);
        $sheet->setCellValue([12,$rowIndex], $costPathologicalAnatomicalBiopsyMaterial);
        $sheet->setCellValue([13,$rowIndex], $molecularGeneticDetectionOncological);
        $sheet->setCellValue([14,$rowIndex], $costMolecularGeneticDetectionOncological);
        $sheet->setCellValue([15,$rowIndex], $covidTesting);
        $sheet->setCellValue([16,$rowIndex], $costCovidTesting);
    }
    $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);

    $careProfilesFoms = CareProfilesFoms::all();

    $sheet = $spreadsheet->getSheetByName('ДС при стационаре');
    $sheet->setCellValue([1,3], $docName);
    $ordinalRowNum = 0;
    $rowIndex = $startRow;
    $category = 'hospital';

    $numberOfBedsIndicatorId = 1; // число коек
    $casesOfTreatmentIndicatorId = 2; // случаев лечения
    $patientDaysIndicatorId = 3; // пациенто-дней
    $costIndicatorId = 4; // стоимость

    foreach($moCollection as $mo) {
        $moRowIndex = $rowIndex;
        $inHospitalBedProfiles = $content['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? null;
        $inHospitalBedProfilesNumberOfBeds = $contentNumberOfBeds['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? null;

        if (!$inHospitalBedProfiles) { continue; }

        $ordinalRowNum++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        foreach ($careProfilesFoms as $cpf) {
            $numberOfBeds = '0';
            $casesOfTreatment = '0';
            $patientDays = '0';
            $cost = '0';

            $hbp = $cpf->hospitalBedProfiles;
            foreach($hbp as $bp) {
                $bpData = $inHospitalBedProfiles[$bp->id] ?? null;
                if (!$bpData) { continue; }

                $casesOfTreatment = bcadd($casesOfTreatment, $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                $patientDays = bcadd($patientDays, $bpData[$patientDaysIndicatorId] ?? '0');
                $cost = bcadd($cost, $bpData[$costIndicatorId] ?? '0');
            }
            foreach($hbp as $bp) {
                $bpData = $inHospitalBedProfilesNumberOfBeds[$bp->id] ?? null;
                if (!$bpData) { continue; }

                $numberOfBeds = bcadd($numberOfBeds, $bpData[$numberOfBedsIndicatorId] ?? '0');
            }
            if( bccomp($numberOfBeds,'0') === 0
                && bccomp($casesOfTreatment, '0') === 0
                && bccomp($patientDays, '0') === 0
                && bccomp($cost, '0') === 0
            ) {continue;}

            $sheet->setCellValue([3,$rowIndex], $cpf->name);
            $sheet->setCellValue([4,$rowIndex], $numberOfBeds);
            $sheet->setCellValue([5,$rowIndex], $casesOfTreatment);
            $sheet->setCellValue([6,$rowIndex], $patientDays);
            $sheet->setCellValue([7,$rowIndex], $cost);
            $rowIndex++;
        }
        if ($rowIndex - 1 > $moRowIndex) {
            $sheet->mergeCells([1, $moRowIndex, 1, $rowIndex - 1]);
            $sheet->mergeCells([2, $moRowIndex, 2, $rowIndex - 1]);
        }
    }
    $sheet->removeRow($rowIndex,$endRow-$rowIndex+1);


    $sheet = $spreadsheet->getSheetByName('ДС при поликлинике');
    $sheet->setCellValue([1,3], $docName);
    $ordinalRowNum = 0;
    $rowIndex = $startRow;
    $category = 'hospital';

    $numberOfBedsIndicatorId = 1; // число коек
    $casesOfTreatmentIndicatorId = 2; // случаев лечения
    $patientDaysIndicatorId = 3; // пациенто-дней
    $costIndicatorId = 4; // стоимость

    foreach($moCollection as $mo) {
        $moRowIndex = $rowIndex;
        $inPolyclinicBedProfiles = $content['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? null;
        $inPolyclinicBedProfilesNumberOfBeds = $contentNumberOfBeds['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? null;

        if (!$inPolyclinicBedProfiles) { continue; }

        $ordinalRowNum++;
        // $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        foreach ($careProfilesFoms as $cpf) {
            $numberOfBeds = '0';
            $casesOfTreatment = '0';
            $patientDays = '0';
            $cost = '0';

            $hbp = $cpf->hospitalBedProfiles;
            foreach($hbp as $bp) {
                $bpData = $inPolyclinicBedProfiles[$bp->id] ?? null;
                if (!$bpData) { continue; }

                $casesOfTreatment = bcadd($casesOfTreatment, $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                $patientDays = bcadd($patientDays, $bpData[$patientDaysIndicatorId] ?? '0');
                $cost = bcadd($cost, $bpData[$costIndicatorId] ?? '0');
            }
            foreach($hbp as $bp) {
                $bpData = $inPolyclinicBedProfilesNumberOfBeds[$bp->id] ?? null;
                if (!$bpData) { continue; }

                $numberOfBeds = bcadd($numberOfBeds, $bpData[$numberOfBedsIndicatorId] ?? '0');
            }
            if( bccomp($numberOfBeds,'0') === 0
                && bccomp($casesOfTreatment, '0') === 0
                && bccomp($patientDays, '0') === 0
                && bccomp($cost, '0') === 0
            ) {continue;}

            $sheet->setCellValue([3,$rowIndex], "$cpf->name");
            $sheet->setCellValue([4,$rowIndex], $numberOfBeds);
            $sheet->setCellValue([5,$rowIndex], $casesOfTreatment);
            $sheet->setCellValue([6,$rowIndex], $patientDays);
            $sheet->setCellValue([7,$rowIndex], $cost);
            $rowIndex++;
        }
        if ($rowIndex - 1 > $moRowIndex) {
            $sheet->mergeCells([1, $moRowIndex, 1, $rowIndex - 1]);
            $sheet->mergeCells([2, $moRowIndex, 2, $rowIndex - 1]);
        }
    }
    $sheet->removeRow($rowIndex,$endRow-$rowIndex+1);


    $sheet = $spreadsheet->getSheetByName('КС');
    $sheet->setCellValue([1,3], $docName);
    $ordinalRowNum = 0;
    $rowIndex = $startRow;
    $category = 'hospital';

    $numberOfBedsIndicatorId = 1; // число коек
    $casesOfTreatmentIndicatorId = 7; // госпитализаций
    $patientDaysIndicatorId = 3; // пациенто-дней
    $costIndicatorId = 4; // стоимость

    foreach($moCollection as $mo) {
        $moRowIndex = $rowIndex;
        $roundClockBedProfiles = $content['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? null;
        $roundClockBedProfilesNumberOfBeds = $contentNumberOfBeds['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? null;

        if (!$roundClockBedProfiles) { continue; }

        $ordinalRowNum++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        foreach ($careProfilesFoms as $cpf) {
            $numberOfBeds = '0';
            $casesOfTreatment = '0';
            $patientDays = '0';
            $cost = '0';

            $hbp = $cpf->hospitalBedProfiles;
            foreach($hbp as $bp) {
                $bpData = $roundClockBedProfiles[$bp->id] ?? null;
                if (!$bpData) { continue; }

                $casesOfTreatment = bcadd($casesOfTreatment, $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                $patientDays = bcadd($patientDays, $bpData[$patientDaysIndicatorId] ?? '0');
                $cost = bcadd($cost, $bpData[$costIndicatorId] ?? '0');
            }
            foreach($hbp as $bp) {
                $bpData = $roundClockBedProfilesNumberOfBeds[$bp->id] ?? null;
                if (!$bpData) { continue; }

                $numberOfBeds = bcadd($numberOfBeds, $bpData[$numberOfBedsIndicatorId] ?? '0');
            }
            if( bccomp($numberOfBeds,'0') === 0
                && bccomp($casesOfTreatment, '0') === 0
                && bccomp($patientDays, '0') === 0
                && bccomp($cost, '0') === 0
            ) {continue;}

            $sheet->setCellValue([3,$rowIndex], "$cpf->name");
            $sheet->setCellValue([4,$rowIndex], $numberOfBeds);
            $sheet->setCellValue([5,$rowIndex], $casesOfTreatment);
            $sheet->setCellValue([6,$rowIndex], $patientDays);
            $sheet->setCellValue([7,$rowIndex], $cost);
            $rowIndex++;
        }
        if ($rowIndex - 1 > $moRowIndex) {
            $sheet->mergeCells([1, $moRowIndex, 1, $rowIndex - 1]);
            $sheet->mergeCells([2, $moRowIndex, 2, $rowIndex - 1]);
        }
    }
    $sheet->removeRow($rowIndex,$endRow-$rowIndex+1);


    $sheet = $spreadsheet->getSheetByName('ВМП');
    $sheet->setCellValue([1,3], $docName);
    $ordinalRowNum = 0;
    $rowIndex = $startRow;
    $category = 'hospital';

    $numberOfBedsIndicatorId = 1; // число коек
    $casesOfTreatmentIndicatorId = 7; // госпитализаций
    $patientDaysIndicatorId = 3; // пациенто-дней
    $costIndicatorId = 4; // стоимость

    foreach($moCollection as $mo) {
        $moRowIndex = $rowIndex;
        $careProfiles = $content['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? null;
        $careProfilesNumberOfBeds = $contentNumberOfBeds['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? null;

        if (!$careProfiles) { continue; }

        $ordinalRowNum++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        foreach ($careProfilesFoms as $cpf) {
            $numberOfBeds = '0';
            $casesOfTreatment = '0';
            $patientDays = '0';
            $cost = '0';

            $cpmz = $cpf->careProfilesMz;
            foreach($cpmz as $cp) {
                $vmpGroupsData = $careProfiles[$cp->id] ?? null;
                if (!$vmpGroupsData) { continue; }

                foreach ($vmpGroupsData as $vmpTypes) {
                    foreach ($vmpTypes as $vmpT)
                    {
                        $casesOfTreatment = bcadd($casesOfTreatment, $vmpT[$casesOfTreatmentIndicatorId] ?? '0');
                        $patientDays = bcadd($patientDays, $vmpT[$patientDaysIndicatorId] ?? '0');
                        $cost = bcadd($cost, $vmpT[$costIndicatorId] ?? '0');
                    }
                }
            }
            foreach($cpmz as $cp) {
                $vmpGroupsData = $careProfilesNumberOfBeds[$cp->id] ?? null;
                if (!$vmpGroupsData) { continue; }

                foreach ($vmpGroupsData as $vmpTypes) {
                    foreach ($vmpTypes as $vmpT)
                    {
                        $numberOfBeds = bcadd($numberOfBeds, $vmpT[$numberOfBedsIndicatorId] ?? '0');
                    }
                }
            }

            if( bccomp($numberOfBeds,'0') === 0
                && bccomp($casesOfTreatment, '0') === 0
                && bccomp($patientDays, '0') === 0
                && bccomp($cost, '0') === 0
            ) {continue;}

            $sheet->setCellValue([3,$rowIndex], "$cpf->name");
            $sheet->setCellValue([4,$rowIndex], $numberOfBeds);
            $sheet->setCellValue([5,$rowIndex], $casesOfTreatment);
            $sheet->setCellValue([6,$rowIndex], $patientDays);
            $sheet->setCellValue([7,$rowIndex], $cost);
            $rowIndex++;
        }
        if ($rowIndex - 1 > $moRowIndex) {
            $sheet->mergeCells([1, $moRowIndex, 1, $rowIndex - 1]);
            $sheet->mergeCells([2, $moRowIndex, 2, $rowIndex - 1]);
        }
    }
    $sheet->removeRow($rowIndex,$endRow-$rowIndex+1);


    $sheet = $spreadsheet->getSheetByName('АП (подушевое финансирование)');
    $sheet->setCellValue([1,3], $docName);
    $ordinalRowNum = 0;
    $rowIndex = $startRow - 1;
    $category = 'polyclinic';

    $contactingIndicatorId = 8; // обращений
    $visitIndicatorId = 9; // посещений
    $costIndicatorId = 4; // стоимость
    $emergencyVisitsAssistanceTypeIds = [3]; //	посещения неотложные
    $otherVisitsAssistanceTypeIds = [1, 2]; //	посещения профилактические, посещения разовые по заболеваниям
    $contactingAssistanceTypeIds = [4]; //обращения по заболеваниям
    foreach($moCollection as $mo) {
        $assistanceTypesPerPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'] ?? null;
        if (!$assistanceTypesPerPerson) { continue; }

        $emergencyVisits = '0';
        $costEmergencyVisits = '0';
        foreach($emergencyVisitsAssistanceTypeIds as $typeId) {
            $emergencyVisits = bcadd($emergencyVisits, $assistanceTypesPerPerson[$typeId][$visitIndicatorId] ?? '0');
            $costEmergencyVisits = bcadd($costEmergencyVisits, $assistanceTypesPerPerson[$typeId][$costIndicatorId] ?? '0');
        }

        $otherVisits = '0';
        $costOtherVisits = '0';
        foreach($otherVisitsAssistanceTypeIds as $typeId) {
            $otherVisits = bcadd($otherVisits, $assistanceTypesPerPerson[$typeId][$visitIndicatorId] ?? '0');
            $costOtherVisits = bcadd($costOtherVisits, $assistanceTypesPerPerson[$typeId][$costIndicatorId] ?? '0');
        }

        $contacting = '0';
        $costContacting = '0';
        foreach($contactingAssistanceTypeIds as $typeId) {
            $contacting = bcadd($contacting, $assistanceTypesPerPerson[$typeId][$contactingIndicatorId] ?? '0');
            $costContacting = bcadd($costContacting, $assistanceTypesPerPerson[$typeId][$costIndicatorId] ?? '0');
        }



        if( bccomp($emergencyVisits,'0') === 0
            && bccomp($otherVisits, '0') === 0
            && bccomp($contacting, '0') === 0
            && bccomp($costEmergencyVisits, '0') === 0
            && bccomp($costOtherVisits, '0') === 0
            && bccomp($costContacting, '0') === 0
        ) {continue;}

        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        $sheet->setCellValue([3,$rowIndex], $otherVisits);
        $sheet->setCellValue([4,$rowIndex], $costOtherVisits);
        $sheet->setCellValue([5,$rowIndex], $contacting);
        $sheet->setCellValue([6,$rowIndex], $costContacting);
        $sheet->setCellValue([7,$rowIndex], $emergencyVisits);
        $sheet->setCellValue([8,$rowIndex], $costEmergencyVisits);

    }
    $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);


    $sheet = $spreadsheet->getSheetByName('АП (по тарифу)');

    $sheet->setCellValue([1,3], $docName);
    $ordinalRowNum = 0;
    $rowIndex = $startRow - 1;
    $category = 'polyclinic';

    $contactingIndicatorId = 8; // обращений
    $visitIndicatorId = 9; // посещений
    $costIndicatorId = 4; // стоимость
    $emergencyVisitsAssistanceTypeIds = [3]; //	посещения неотложные
    $otherVisitsAssistanceTypeIds = [1, 2]; //	посещения профилактические, посещения разовые по заболеваниям
    $contactingAssistanceTypeIds = [4]; //обращения по заболеваниям
    foreach($moCollection as $mo) {
        $assistanceTypesPerUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'] ?? null;
        if (!$assistanceTypesPerUnit) { continue; }

        $emergencyVisits = '0';
        $costEmergencyVisits = '0';
        foreach($emergencyVisitsAssistanceTypeIds as $typeId) {
            $emergencyVisits = bcadd($emergencyVisits, $assistanceTypesPerUnit[$typeId][$visitIndicatorId] ?? '0');
            $costEmergencyVisits = bcadd($costEmergencyVisits, $assistanceTypesPerUnit[$typeId][$costIndicatorId] ?? '0');
        }

        $otherVisits = '0';
        $costOtherVisits = '0';
        foreach($otherVisitsAssistanceTypeIds as $typeId) {
            $otherVisits = bcadd($otherVisits, $assistanceTypesPerUnit[$typeId][$visitIndicatorId] ?? '0');
            $costOtherVisits = bcadd($costOtherVisits, $assistanceTypesPerUnit[$typeId][$costIndicatorId] ?? '0');
        }

        $contacting = '0';
        $costContacting = '0';
        foreach($contactingAssistanceTypeIds as $typeId) {
            $contacting = bcadd($contacting, $assistanceTypesPerUnit[$typeId][$contactingIndicatorId] ?? '0');
            $costContacting = bcadd($costContacting, $assistanceTypesPerUnit[$typeId][$costIndicatorId] ?? '0');
        }



        if( bccomp($emergencyVisits,'0') === 0
            && bccomp($otherVisits, '0') === 0
            && bccomp($contacting, '0') === 0
            && bccomp($costEmergencyVisits, '0') === 0
            && bccomp($costOtherVisits, '0') === 0
            && bccomp($costContacting, '0') === 0
        ) {continue;}

        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        $sheet->setCellValue([3,$rowIndex], $otherVisits);
        $sheet->setCellValue([4,$rowIndex], $costOtherVisits);
        $sheet->setCellValue([5,$rowIndex], $contacting);
        $sheet->setCellValue([6,$rowIndex], $costContacting);
        $sheet->setCellValue([7,$rowIndex], $emergencyVisits);
        $sheet->setCellValue([8,$rowIndex], $costEmergencyVisits);

    }
    $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);


    $sheet = $spreadsheet->getSheetByName('АП (ФАП)');

    $sheet->setCellValue([1,3], $docName);
    $ordinalRowNum = 0;
    $rowIndex = $startRow - 1;
    $category = 'polyclinic';

    $contactingIndicatorId = 8; // обращений
    $visitIndicatorId = 9; // посещений
    $costIndicatorId = 4; // стоимость
    $emergencyVisitsAssistanceTypeIds = [3]; //	посещения неотложные
    $otherVisitsAssistanceTypeIds = [1, 2]; //	посещения профилактические, посещения разовые по заболеваниям
    $contactingAssistanceTypeIds = [4]; //обращения по заболеваниям
    foreach($moCollection as $mo) {

        $faps = $content['mo'][$mo->id][$category]['fap'] ?? null;
        if (!$faps) { continue; }

        $emergencyVisits = '0';
        $costEmergencyVisits = '0';
        $otherVisits = '0';
        $costOtherVisits = '0';
        $contacting = '0';
        $costContacting = '0';
        foreach ($faps as $f) {
            $assistanceTypesData = $f['assistanceTypes'];
            foreach($emergencyVisitsAssistanceTypeIds as $typeId) {
                $emergencyVisits = bcadd($emergencyVisits, $assistanceTypesData[$typeId][$visitIndicatorId] ?? '0');
                $costEmergencyVisits = bcadd($costEmergencyVisits, $assistanceTypesData[$typeId][$costIndicatorId] ?? '0');
            }
            foreach($otherVisitsAssistanceTypeIds as $typeId) {
                $otherVisits = bcadd($otherVisits, $assistanceTypesData[$typeId][$visitIndicatorId] ?? '0');
                $costOtherVisits = bcadd($costOtherVisits, $assistanceTypesData[$typeId][$costIndicatorId] ?? '0');
            }
            foreach($contactingAssistanceTypeIds as $typeId) {
                $contacting = bcadd($contacting, $assistanceTypesData[$typeId][$contactingIndicatorId] ?? '0');
                $costContacting = bcadd($costContacting, $assistanceTypesData[$typeId][$costIndicatorId] ?? '0');
            }
        }

        if( bccomp($emergencyVisits,'0') === 0
            && bccomp($otherVisits, '0') === 0
            && bccomp($contacting, '0') === 0
            && bccomp($costEmergencyVisits, '0') === 0
            && bccomp($costOtherVisits, '0') === 0
            && bccomp($costContacting, '0') === 0
        ) {continue;}

        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        $sheet->setCellValue([3,$rowIndex], $otherVisits);
        $sheet->setCellValue([4,$rowIndex], $costOtherVisits);
        $sheet->setCellValue([5,$rowIndex], $contacting);
        $sheet->setCellValue([6,$rowIndex], $costContacting);
        $sheet->setCellValue([7,$rowIndex], $emergencyVisits);
        $sheet->setCellValue([8,$rowIndex], $costEmergencyVisits);

    }
    $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);


    $writer = new Xlsx($spreadsheet);
    $writer->save($fullResultFilepath);
    return Storage::download($resultFilePath);
});

function vitacoreV2PrintRow(
        Worksheet $sheet,
        int $colIndex,
        int $rowIndex,
        string | int $ordinalRowNum,
        string | int $moCode,
        string $moName,
        string $planningSectionName,
        string $planningParamName,
        array $values
    ) {
        $moCodeColOffset = 1;
        $moShortNameColOffset = 2;
        $planningSectionColOffset = 3;
        $paramColOffset = 4;
        $firstValueColOffset = 5;

        $sheet->setCellValue([$colIndex, $rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([$colIndex + $moCodeColOffset, $rowIndex], $moCode);
        $sheet->setCellValue([$colIndex + $moShortNameColOffset, $rowIndex], $moName);
        $sheet->setCellValue([$colIndex + $planningSectionColOffset, $rowIndex], $planningSectionName);
        $sheet->setCellValue([$colIndex + $paramColOffset, $rowIndex], $planningParamName);
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $sheet->setCellValue([$colIndex + $firstValueColOffset + $monthNum - 1, $rowIndex], $values[$monthNum]);
        }
}

function vitacoreV2PrintTableHeader(
        Worksheet $sheet,
        int $colIndex,
        int $rowIndex,
  ) {
    $months = [
        1 => 'Январь',
        2 => 'Февраль',
        3 => 'Март',
        4 => 'Апрель',
        5 => 'Май',
        6 => 'Июнь',
        7 => 'Июль',
        8 => 'Август',
        9 => 'Сентябрь',
        10 => 'Октябрь',
        11 => 'Ноябрь',
        12 => 'Декабрь',
    ];
    vitacoreV2PrintRow($sheet, $colIndex, $rowIndex, "", "", "", "", "", $months);
}

function vitacoreHospitalByProfilePeriodsPrintRow(
    Worksheet $sheet,
    int $colIndex,
    int $rowIndex,
    string | int $ordinalRowNum,
    string | int $moCode,
    string $moName,
    string $planningSectionName,
    string $careProfileName,
    string $careProfileV002,
    string $planningParamName,
    array $values
) {
    $moCodeColOffset = 1;
    $moShortNameColOffset = 2;
    $planningSectionColOffset = 3;
    $careProfileNameOffset = 4;
    $careProfileV002Offset = 5;
    $paramColOffset = 6;
    $firstValueColOffset = 7;

    $sheet->setCellValue([$colIndex, $rowIndex], "$ordinalRowNum");
    $sheet->setCellValue([$colIndex + $moCodeColOffset, $rowIndex], $moCode);
    $sheet->setCellValue([$colIndex + $moShortNameColOffset, $rowIndex], $moName);
    $sheet->setCellValue([$colIndex + $planningSectionColOffset, $rowIndex], $planningSectionName);
    $sheet->setCellValue([$colIndex + $careProfileNameOffset, $rowIndex], $careProfileName);
    $sheet->setCellValue([$colIndex + $careProfileV002Offset, $rowIndex], $careProfileV002);
    $sheet->setCellValue([$colIndex + $paramColOffset, $rowIndex], $planningParamName);
    for ($monthNum = 1; $monthNum <= 12; $monthNum++) {
        $sheet->setCellValue([$colIndex + $firstValueColOffset + $monthNum - 1, $rowIndex], $values[$monthNum]);
    }
}

function vitacoreHospitalByProfilePeriodsPrintTableHeader(
    Worksheet $sheet,
    int $colIndex,
    int $rowIndex,
) {
    $months = [
        1 => 'Январь',
        2 => 'Февраль',
        3 => 'Март',
        4 => 'Апрель',
        5 => 'Май',
        6 => 'Июнь',
        7 => 'Июль',
        8 => 'Август',
        9 => 'Сентябрь',
        10 => 'Октябрь',
        11 => 'Ноябрь',
        12 => 'Декабрь',
    ];
    vitacoreHospitalByProfilePeriodsPrintRow($sheet, $colIndex, $rowIndex, "", "", "", "", "", "", "", $months);
}

Route::get('/hospitalization-portal/{year}/{commissionDecisionsId?}', [PlanReports::class, "NumberOfBeds"]);

Route::get('/vitacore-v2/{year}/{commissionDecisionsId?}', function (DataForContractService $dataForContractService, PeopleAssignedInfoForContractService $peopleAssignedInfoForContractService, int $year, int $commissionDecisionsId = null) {
    $packageIds = null;
    $currentlyUsedDate = $year.'-01-01';
    if ($commissionDecisionsId) {
        $commissionDecisions = CommissionDecision::whereYear('date',$year)->where('id', '<=', $commissionDecisionsId)->get();
        $cd = $commissionDecisions->find($commissionDecisionsId);
        $commissionDecisionIds = $commissionDecisions->pluck('id')->toArray();
        $protocolDate = $cd->date->format('d.m.Y');
        $packageIds = ChangePackage::whereIn('commission_decision_id', $commissionDecisionIds)->orWhere('commission_decision_id', null)->pluck('id')->toArray();

        $currentlyUsedDate = $cd->date->format('Y-m-d');
    } else {
        $packageIds = ChangePackage::where('commission_decision_id', null)->pluck('id')->toArray();
    }

    $path = 'xlsx';
    $resultFileName = 'vitacore.xlsx';
    $strDateTimeNow = date("Y-m-d-His");
    $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . ' ' . $resultFileName;
    $fullResultFilepath = Storage::path($resultFilePath);

    bcscale(4);

    $content = $dataForContractService->GetArray($year, $packageIds);
    $contentByMonth = [];
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $contentByMonth[$monthNum] = $dataForContractService->GetArrayByYearAndMonth($year, $monthNum, $packageIds);
    }

    $peopleAssigned = $peopleAssignedInfoForContractService->GetArray($year, $packageIds);
    $moCollection = MedicalInstitution::WhereRaw("? BETWEEN effective_from AND effective_to", [$currentlyUsedDate])->orderBy('order')->get();

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setCellValue([1, 1], "Год:");
    $sheet->setCellValue([2, 1], "$year");

    $ordinalRowNum = 1;
    $firstTableHeadRowIndex = 2;
    $firstTableDataRowIndex = 3;
    $firstTableColIndex = 1;
    $rowOffset = 0;

    vitacoreV2PrintTableHeader($sheet, $firstTableColIndex, $firstTableHeadRowIndex);

    foreach($moCollection as $mo) {
        $planningSectionName = "Скорая помощь";
        $planningParamName = "объемы, вызовов";
        $category = 'ambulance';
        $indicatorId = 5; // вызовов
        $callsAssistanceTypeId = 5; // вызовы
        $thrombolysisAssistanceTypeId = 6;// тромболизис

        $thrombolysis = [];
        $calls = [];
        $hasValue = false;
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $thrombolysis[$monthNum] = $contentByMonth[$monthNum]['mo'][$mo->id][$category][$thrombolysisAssistanceTypeId][$indicatorId] ?? 0;
            $calls[$monthNum] = $contentByMonth[$monthNum]['mo'][$mo->id][$category][$callsAssistanceTypeId][$indicatorId] ?? 0;
            if( bccomp($thrombolysis[$monthNum],'0') !== 0
                || bccomp($calls[$monthNum],'0') !== 0
            ) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            $values = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $values[$monthNum] = $calls[$monthNum] + $thrombolysis[$monthNum];
            }
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }

        // 2.обращения по заболеваниям
        $planningSectionName = "Амбулаторная помощь в связи с заболеваниями";
        $planningParamName = "объемы, обращений";
        $hasValue = false;
        $category = 'polyclinic';
        $indicatorId = 8; // обращений
        $assistanceTypeId = 4; // обращения по заболеваниям
        $perPerson = [];
        $perUnit = [];
        $fap = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $perPerson[$monthNum] = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
            $perUnit[$monthNum] = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
            $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
            $fap[$monthNum] = 0;
            foreach ($faps as $f) {
                $fap[$monthNum] += $f['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
            }
            if( bccomp($perPerson[$monthNum],'0') !== 0
                || bccomp($perUnit[$monthNum],'0') !== 0
                || bccomp($fap[$monthNum],'0') !== 0
            ) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            $values = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $values[$monthNum] = $perPerson[$monthNum] + $perUnit[$monthNum] + $fap[$monthNum];
            }
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }


        // фин.обесп.
        $category = 'polyclinic';
        $assistanceTypeId = 4; //обращения по заболеваниям
        $indicatorId = 4; // стоимость

        // не заполняется по медицинским организациям имеющим прикрепленное население
        if (!($peopleAssigned[$mo->id][$category]['mo']['peopleAssigned'] ?? false)) {
            $planningSectionName = "Амбулаторная помощь в связи с заболеваниями";
            $planningParamName = "финансовое обеспечение, руб.";
            $hasValue = false;
            $perPerson = [];
            $perUnit = [];
            $fap = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $perPerson[$monthNum] = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
                $perUnit[$monthNum] = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
                $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
                $fap[$monthNum] = 0;
                foreach ($faps as $f) {
                    $fap[$monthNum] = bcadd($fap[$monthNum], $f['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0');
                }
                if( bccomp($perPerson[$monthNum],'0') !== 0
                    || bccomp($perUnit[$monthNum],'0') !== 0
                    || bccomp($fap[$monthNum],'0') !== 0
                ) {
                    $hasValue = true;
                }
            }
            if($hasValue) {
                $values = [];
                for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                    $values[$monthNum] = bcadd(bcadd($perPerson[$monthNum], $perUnit[$monthNum]), $fap[$monthNum]);
                }
                vitacoreV2PrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset,
                    $ordinalRowNum,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $planningParamName,
                    $values
                );
                $rowOffset++;
                $ordinalRowNum++;
            }
        }
        // 3.Посещения с иными целями
        $planningSectionName = "Амбулаторная помощь с профилактическими и иными целями";
        $planningParamName = "объемы, посещений";
        $hasValue = false;
        $category = 'polyclinic';
        $indicatorId = 9; // посещений
        $assistanceTypeIds = [1, 2]; //	посещения профилактические, посещения разовые по заболеваниям
        $perPerson = [];
        $perUnit = [];
        $fap = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $perPerson[$monthNum] = 0;
            $perUnit[$monthNum] = 0;
            $fap[$monthNum] = 0;
            $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
            foreach($assistanceTypeIds as $assistanceTypeId) {
                $perPerson[$monthNum] += $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
                $perUnit[$monthNum] += $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
                foreach ($faps as $f) {
                    $fap[$monthNum] += $f['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
                }
            }
            if( bccomp($perPerson[$monthNum],'0') !== 0
                || bccomp($perUnit[$monthNum],'0') !== 0
                || bccomp($fap[$monthNum],'0') !== 0
            ) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $values[$monthNum] = $perPerson[$monthNum] + $perUnit[$monthNum] + $fap[$monthNum];
            }
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }
        // фин.обесп.
        $category = 'polyclinic';
        $assistanceTypeIds = [1, 2]; //	посещения профилактические, посещения разовые по заболеваниям
        $indicatorId = 4; // стоимость

        // не заполняется по медицинским организациям имеющим прикрепленное население

        if (!($peopleAssigned[$mo->id][$category]['mo']['peopleAssigned'] ?? false)) {
            $planningSectionName = "Амбулаторная помощь с профилактическими и иными целями";
            $planningParamName = "финансовое обеспечение, руб.";
            $hasValue = false;
            $perPerson = [];
            $perUnit = [];
            $fap = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $perPerson[$monthNum] = 0;
                $perUnit[$monthNum] = 0;
                $fap[$monthNum] = 0;
                $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
                foreach($assistanceTypeIds as $assistanceTypeId) {
                    $perPerson[$monthNum] = bcadd($perPerson[$monthNum], $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0');
                    $perUnit[$monthNum]   = bcadd($perUnit[$monthNum],   $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0');
                    foreach ($faps as $f) {
                        $fap[$monthNum] = bcadd($fap[$monthNum], $f['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0');
                    }
                }
                if( bccomp($perPerson[$monthNum],'0') !== 0
                    || bccomp($perUnit[$monthNum],'0') !== 0
                    || bccomp($fap[$monthNum],'0') !== 0
                ) {
                    $hasValue = true;
                }
            }
            if($hasValue) {
                $values = [];
                for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                    $values[$monthNum] = bcadd(bcadd($perPerson[$monthNum], $perUnit[$monthNum]), $fap[$monthNum]);
                }
                vitacoreV2PrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset,
                    $ordinalRowNum,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $planningParamName,
                    $values
                );
                $rowOffset++;
                $ordinalRowNum++;
            }
        }

        // 4 Неотложная помощь
        $planningSectionName = "Амбулаторная помощь с неотложной целью";
        $planningParamName = "объемы, посещений";
        $hasValue = false;
        $category = 'polyclinic';
        $indicatorId = 9; // посещений
        $assistanceTypeIds = [3]; //	посещения неотложные
        $perPerson = [];
        $perUnit = [];
        $fap = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $perPerson[$monthNum] = 0;
            $perUnit[$monthNum] = 0;
            $fap[$monthNum] = 0;
            $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
            foreach($assistanceTypeIds as $assistanceTypeId) {
                $perPerson[$monthNum] += $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
                $perUnit[$monthNum] += $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
                foreach ($faps as $f) {
                    $fap[$monthNum] += $f['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
                }
            }
            if( bccomp($perPerson[$monthNum],'0') !== 0
                || bccomp($perUnit[$monthNum],'0') !== 0
                || bccomp($fap[$monthNum],'0') !== 0
            ) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $values[$monthNum] = $perPerson[$monthNum] + $perUnit[$monthNum] + $fap[$monthNum];
            }
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }
        // фин.обесп.
        $category = 'polyclinic';
        $assistanceTypeIds = [3]; // посещения неотложные
        $indicatorId = 4; // стоимость

        $planningSectionName = "Амбулаторная помощь с неотложной целью";
        $planningParamName = "финансовое обеспечение, руб.";
        $hasValue = false;
        $perPerson = [];
        $perUnit = [];
        $fap = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $perPerson[$monthNum] = 0;
            $perUnit[$monthNum] = 0;
            $fap[$monthNum] = 0;
            $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
            foreach($assistanceTypeIds as $assistanceTypeId) {
                $perPerson[$monthNum] = bcadd($perPerson[$monthNum], $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0');
                $perUnit[$monthNum]   = bcadd($perUnit[$monthNum],   $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0');
                foreach ($faps as $f) {
                    $fap[$monthNum] = bcadd($fap[$monthNum], $f['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0');
                }
            }
            if( bccomp($perPerson[$monthNum],'0') !== 0
                || bccomp($perUnit[$monthNum],'0') !== 0
                || bccomp($fap[$monthNum],'0') !== 0
            ) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            $values = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $values[$monthNum] = bcadd(bcadd($perPerson[$monthNum], $perUnit[$monthNum]), $fap[$monthNum]);
            }
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }

        // 2.2 КТ
        $hasValue = false;
        $planningSectionName = "компьютерная томография";
        $planningParamName = "объемы, услуг";
        $category = 'polyclinic';
        $indicatorId = 6; // услуг
        $serviceId = MedicalServicesEnum::KT;

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }
        // 2.2 КТ
        $hasValue = false;
        $planningSectionName = "компьютерная томография";
        $planningParamName = "финансовое обеспечение, руб.";
        $category = 'polyclinic';
        $indicatorId = 4; // стоимость
        $serviceId = MedicalServicesEnum::KT;

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }


        // 2.3 МРТ
        $hasValue = false;
        $planningSectionName = "магнитно-резонансная томография";
        $planningParamName = "объемы, услуг";
        $category = 'polyclinic';
        $indicatorId = 6; // услуг
        $serviceId = MedicalServicesEnum::MRT;

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }
        // 2.3 МРТ
        $category = 'polyclinic';
        $hasValue = false;
        $planningSectionName = "магнитно-резонансная томография";
        $planningParamName = "финансовое обеспечение, руб.";
        $category = 'polyclinic';
        $indicatorId = 4; // стоимость
        $serviceId = MedicalServicesEnum::MRT;

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }

        // 2.4 УЗИ ССС
        $hasValue = false;
        $planningSectionName = "УЗИ сердечно-сосудистой системы";
        $planningParamName = "объемы, услуг";
        $category = 'polyclinic';
        $indicatorId = 6; // услуг
        $serviceId = MedicalServicesEnum::UltrasoundCardio;

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }
        // 2.4 УЗИ ССС
        $hasValue = false;
        $planningSectionName = "УЗИ сердечно-сосудистой системы";
        $category = 'polyclinic';
        $indicatorId = 4; // стоимость
        $serviceId = MedicalServicesEnum::UltrasoundCardio;
        $planningParamName = "финансовое обеспечение, руб.";
        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }


        // 2.5 Эндоскопия
        $category = 'polyclinic';
        $indicatorId = 6; // услуг
        $serviceId = MedicalServicesEnum::Endoscopy;
        $hasValue = false;
        $planningSectionName = "Эндоскопические исследования";
        $planningParamName = "объемы, услуг";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }
        // 2.5 Эндоскопия
        $category = 'polyclinic';
        $indicatorId = 4; // стоимость
        $serviceId = MedicalServicesEnum::Endoscopy;
        $hasValue = false;
        $planningSectionName = "Эндоскопические исследования";
        $planningParamName = "финансовое обеспечение, руб.";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }

        // 2.6 ПАИ
        $category = 'polyclinic';
        $indicatorId = 6; // услуг
        $serviceId = MedicalServicesEnum::PathologicalAnatomicalBiopsyMaterial;
        $hasValue = false;
        $planningSectionName = "Паталого анатомическое исследование биопсийного материала с целью диагностики онкологических заболеваний";
        $planningParamName = "объемы, услуг";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }
        // финансовое обеспечение
        $category = 'polyclinic';
        $indicatorId = 4; // стоимость
        $serviceId = MedicalServicesEnum::PathologicalAnatomicalBiopsyMaterial;
        $hasValue = false;
        $planningSectionName = "Паталого анатомическое исследование биопсийного материала с целью диагностики онкологических заболеваний";
        $planningParamName = "финансовое обеспечение, руб.";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }


        // 2.7 МГИ
        $category = 'polyclinic';
        $indicatorId = 6; // услуг
        $serviceId = MedicalServicesEnum::MolecularGeneticDetectionOncological;
        $hasValue = false;
        $planningSectionName = "Малекулярно-генетические исследования с целью диагностики онкологических заболеваний";
        $planningParamName = "объемы, услуг";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }
        // финансовое обеспечение
        $category = 'polyclinic';
        $indicatorId = 4; // стоимость
        $serviceId = MedicalServicesEnum::MolecularGeneticDetectionOncological;
        $hasValue = false;
        $planningSectionName = "Малекулярно-генетические исследования с целью диагностики онкологических заболеваний";
        $planningParamName = "финансовое обеспечение, руб.";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }

        // Тест.covid-19
        $category = 'polyclinic';
        $indicatorId = 6; // услуг
        $serviceId = MedicalServicesEnum::CovidTesting;
        $hasValue = false;
        $planningSectionName = "Тестирование на выявление covid-19";
        $planningParamName = "объемы, услуг";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
            if (bccomp($values[$monthNum], '0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }
        // финансовое обеспечение
        $category = 'polyclinic';
        $indicatorId = 4; // стоимость
        $serviceId = MedicalServicesEnum::CovidTesting;
        $hasValue = false;
        $planningSectionName = "Тестирование на выявление covid-19";
        $planningParamName = "финансовое обеспечение, руб.";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }

        // 3.3 УЗИ плода
        $category = 'polyclinic';
        $indicatorId = 6; // услуг
        $serviceId = MedicalServicesEnum::FetalUltrasound;
        $hasValue = false;
        $planningSectionName = "УЗИ плода (1 триместр)";
        $planningParamName = "объемы, услуг";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }
        // финансовое обеспечение
        $category = 'polyclinic';
        $indicatorId = 4; // стоимость
        $serviceId = MedicalServicesEnum::FetalUltrasound;
        $hasValue = false;
        $planningSectionName = "УЗИ плода (1 триместр)";
        $planningParamName = "финансовое обеспечение, руб.";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }

        // 3.4 Компл.иссл. репрод.орг.
        $category = 'polyclinic';
        $indicatorId = 6; // услуг
        $serviceId = MedicalServicesEnum::DiagnosisBackgroundPrecancerousDiseasesReproductiveWomen;
        $hasValue = false;
        $planningSectionName = "комплексное исследование для диагностики фоновых и предраковых заболеваний репродуктивных органов у женщин";
        $planningParamName = "объемы, услуг";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }
        // финансовое обеспечение
        $category = 'polyclinic';
        $indicatorId = 4; // стоимость
        $serviceId = MedicalServicesEnum::DiagnosisBackgroundPrecancerousDiseasesReproductiveWomen;
        $hasValue = false;
        $planningSectionName = "комплексное исследование для диагностики фоновых и предраковых заболеваний репродуктивных органов у женщин";
        $planningParamName = "финансовое обеспечение, руб.";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }

        // 3.5 Опред.антигена D
        $category = 'polyclinic';
        $indicatorId = 6; // услуг
        $serviceId = MedicalServicesEnum::DeterminationAntigenD;
        $hasValue = false;
        $planningSectionName = "определение антигена D системы Резус (резус-фактор плода)";
        $planningParamName = "объемы, услуг";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }
        // финансовое обеспечение
        $category = 'polyclinic';
        $indicatorId = 4; // стоимость
        $serviceId = MedicalServicesEnum::DeterminationAntigenD;
        $hasValue = false;
        $planningSectionName = "определение антигена D системы Резус (резус-фактор плода)";
        $planningParamName = "финансовое обеспечение, руб.";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }

        // Круглосуточный ст. (не включая ВМП и медицинскую реабилитацию)
        $category = 'hospital';
        $indicatorId = 7; // госпитализаций

        $hasValue = false;
        $planningSectionName = "Круглосуточный стационар (не включая ВМП и медицинскую реабилитацию)";
        $planningParamName = "объемы, госпитализаций";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $bedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? [];
            $values[$monthNum] = 0;
            foreach ($bedProfiles as $bpId => $bp) {
                if (RehabilitationProfileService::IsRehabilitationBedProfile($bpId)) {
                    continue;
                }
                $values[$monthNum] += $bp[$indicatorId] ?? 0;
            }
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }

        // 6.ВМП
        $category = 'hospital';
        $indicatorId = 7; // госпитализаций
        $hasValue = false;
        $planningSectionName = "ВМП";
        $planningParamName = "объемы, госпитализаций";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $careProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? [];
            $values[$monthNum] = 0;
            foreach ($careProfiles as $vmpGroups) {
                foreach ($vmpGroups as $vmpTypes) {
                    foreach ($vmpTypes as $vmpT) {
                        $values[$monthNum] += ($vmpT[$indicatorId] ?? 0);
                    }
                }
            }
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }

        // Медреабилитация в КС
        $category = 'hospital';
        $indicatorId = 7; // госпитализаций
        $hasValue = false;
        $planningSectionName = "Медицинская реабилитация";
        $planningParamName = "объемы, госпитализаций";

        $values = [];
        $rehabilitationBedProfileIds = RehabilitationProfileService::GetAllRehabilitationBedProfileIds();
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = 0;
            foreach ($rehabilitationBedProfileIds as $rbpId) {
                $values[$monthNum] += $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'][$rbpId][$indicatorId] ?? 0;
            }
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }
        // 8. Дневные стационары
        $category = 'hospital';
        $indicatorId = 2; // случаев лечения
        $hasValue = false;
        $planningSectionName = "Дневные стационары";
        $planningParamName = "объемы, случаев лечения";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = 0;
            $bedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? [];
            foreach ($bedProfiles as $bp) {
                $values[$monthNum] += $bp[$indicatorId] ?? 0;
            }

            $bedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? [];
            foreach ($bedProfiles as $bp) {
                $values[$monthNum] += $bp[$indicatorId] ?? 0;
            }

            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }

        // 3. ДС, фин.обеспечение
        $category = 'hospital';
        $indicatorId = 4; // стоимость
        $hasValue = false;
        $planningSectionName = "Дневные стационары";
        $planningParamName = "финансовое обеспечение, руб.";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = 0;
            $bedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? [];
            foreach ($bedProfiles as $bp) {
                $values[$monthNum] = bcadd($values[$monthNum], $bp[$indicatorId] ?? '0');
            }

            $bedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? [];
            foreach ($bedProfiles as $bp) {
                $values[$monthNum] = bcadd($values[$monthNum], $bp[$indicatorId] ?? '0');
            }
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }
        // 4 КС, фин.обеспечение (не включая мед.реабилитацию и ВМП)
        $category = 'hospital';
        $indicatorId = 4; // стоимость
        $hasValue = false;
        $planningSectionName = "Круглосуточный стационар (не включая ВМП и медицинскую реабилитацию)";
        $planningParamName = "финансовое обеспечение, руб.";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = 0;
            $bedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? [];
            foreach ($bedProfiles as $bpId => $bp) {
                if (RehabilitationProfileService::IsRehabilitationBedProfile($bpId)) {
                    continue;
                }
                $values[$monthNum] = bcadd($values[$monthNum], $bp[$indicatorId] ?? '0');
            }
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }
        // 5 МР, фин.обеспечение
        $category = 'hospital';
        $indicatorId = 4; // стоимость
        $hasValue = false;
        $planningSectionName = "Медицинская реабилитация";
        $planningParamName = "финансовое обеспечение, руб.";

        $values = [];
        $rehabilitationBedProfileIds = RehabilitationProfileService::GetAllRehabilitationBedProfileIds();
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = '0';
            foreach ($rehabilitationBedProfileIds as $rbpId) {
                bcadd($values[$monthNum], $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'][$rbpId][$indicatorId] ?? '0');
            }
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }
        // 6 ВМП, фин.обеспечение
        $category = 'hospital';
        $indicatorId = 4; // стоимость
        $hasValue = false;
        $planningSectionName = "ВМП";
        $planningParamName = "финансовое обеспечение, руб.";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = 0;
            $careProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? [];
            foreach ($careProfiles as $vmpGroups) {
                foreach ($vmpGroups as $vmpTypes) {
                    foreach ($vmpTypes as $vmpT) {
                        $values[$monthNum] = bcadd($values[$monthNum], $vmpT[$indicatorId] ?? '0');
                    }
                }
            }
            if (bccomp($values[$monthNum],'0') !== 0) {
                $hasValue = true;
            }
        }
        if($hasValue) {
            vitacoreV2PrintRow(
                $sheet,
                $firstTableColIndex,
                $firstTableDataRowIndex + $rowOffset,
                $ordinalRowNum,
                $mo->code,
                $mo->short_name,
                $planningSectionName,
                $planningParamName,
                $values
            );
            $rowOffset++;
            $ordinalRowNum++;
        }

    } // foreach MO


    $writer = new Xlsx($spreadsheet);
    $writer->save($fullResultFilepath);
    return Storage::download($resultFilePath);
});


Route::get('/summary-volume/{year}/{commissionDecisionsId?}', [PlanReports::class, "SummaryVolume"]);


Route::get('/summary-cost/{year}/{commissionDecisionsId?}', [PlanReports::class, "SummaryCost"]);

Route::get('/pump-pgg/{year}/{commissionDecisionsId?}', [PlanReports::class, "PumpPgg"]);


Route::get('/hospital-by-profile/{year}/{commissionDecisionsId?}', function (DataForContractService $dataForContractService, PlannedIndicatorChangeInitService $plannedIndicatorChangeInitService, InitialDataFixingService $initialDataFixingService, int $year, int $commissionDecisionsId = null) {
    $packageIds = null;
    $currentlyUsedDate = $year.'-01-01';
    if ($commissionDecisionsId) {
        $commissionDecisions = CommissionDecision::whereYear('date',$year)->where('id', '<=', $commissionDecisionsId)->get();
        $cd = $commissionDecisions->find($commissionDecisionsId);
        $commissionDecisionIds = $commissionDecisions->pluck('id')->toArray();
        $packageIds = ChangePackage::whereIn('commission_decision_id', $commissionDecisionIds)->orWhere('commission_decision_id', null)->pluck('id')->toArray();

        $currentlyUsedDate = $cd->date->format('Y-m-d');
    } else {
        if ($initialDataFixingService->fixedYear($year)) {
            $packageIds = ChangePackage::where('commission_decision_id', null)->pluck('id')->toArray();
        } else {
            $plannedIndicatorChangeInitService->fromInitialData($year);
        }
    }
    $path = 'xlsx';
    $resultFileName = 'hospital.xlsx';
    $strDateTimeNow = date("Y-m-d-His");
    $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . ' ' . $resultFileName;
    $fullResultFilepath = Storage::path($resultFilePath);

    bcscale(4);
    $indicatorIds = [2, 3, 4, 7];
    $content = $dataForContractService->GetArray($year, $packageIds, $indicatorIds);
    // количество коек на последний месяц года
    $contentNumberOfBeds = $dataForContractService->GetArrayByYearAndMonth($year, 12, $packageIds, [1]);

    $moCollection = MedicalInstitution::WhereRaw("? BETWEEN effective_from AND effective_to", [$currentlyUsedDate])->orderBy('order')->get();

    // только используемые профили
    $bedProfilesUsedInHospital = [];
    $bedProfilesUsedInPolyclinic = [];
    $bedProfilesUsedInRegular = [];
    $careProfilesUsed = [];
    foreach($moCollection as $mo) {
        $category = 'hospital';

        $inHospitalBedProfiles = $content['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? [];
        $inPolyclinicBedProfiles = $content['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? [];
        $regularBedProfiles = $content['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? [];
        $careProfiles = $content['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? [];

        $bedProfilesUsedInHospital = array_unique(array_merge($bedProfilesUsedInHospital, array_keys($inHospitalBedProfiles)));
        $bedProfilesUsedInPolyclinic = array_unique(array_merge($bedProfilesUsedInPolyclinic, array_keys($inPolyclinicBedProfiles)));
        $bedProfilesUsedInRegular = array_unique(array_merge($bedProfilesUsedInRegular, array_keys($regularBedProfiles)));
        $careProfilesUsed = array_unique(array_merge($careProfilesUsed, array_keys($careProfiles)));
    }
    //$careProfilesFoms = CareProfilesFoms::all();
    $order = 'name';
    $careProfilesFomsUsedInHospital = CareProfilesFoms::whereHas('hospitalBedProfiles', function ($query) use ($bedProfilesUsedInHospital) {
        return $query->whereIn('hospital_bed_profile_id', $bedProfilesUsedInHospital);
    })->orderBy($order)->get();
    $careProfilesFomsUsedInPolyclinic = CareProfilesFoms::whereHas('hospitalBedProfiles', function ($query) use ($bedProfilesUsedInPolyclinic) {
        return $query->whereIn('hospital_bed_profile_id', $bedProfilesUsedInPolyclinic);
    })->orderBy($order)->get();
    $careProfilesFomsUsedInRegular = CareProfilesFoms::whereHas('hospitalBedProfiles', function ($query) use ($bedProfilesUsedInRegular) {
        return $query->whereIn('hospital_bed_profile_id', $bedProfilesUsedInRegular);
    })->orderBy($order)->get();
    $careProfilesFomsUsedVmp = CareProfilesFoms::whereHas('careProfilesMz', function ($query) use ($careProfilesUsed) {
        return $query->whereIn('care_profile_id', $careProfilesUsed);
    })->orderBy($order)->get();


    $spreadsheet = new Spreadsheet();
    $startRow = 1;
    $startColoumn = 1;
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setTitle('ДС при стационаре');

    $ordinalRowNum = 0;
    $rowIndex = $startRow + 3;
    $tableHeadRow = $startRow;
    $coloumnIndex = $startColoumn;
    $category = 'hospital';

    $numberOfBedsIndicatorId = 1; // число коек
    $casesOfTreatmentIndicatorId = 2; // случаев лечения
    $patientDaysIndicatorId = 3; // пациенто-дней
    $costIndicatorId = 4; // стоимость

    $profilesIndex = 0;
    foreach ($careProfilesFomsUsedInHospital as $cpf) {
        $d = (4 * $profilesIndex) + 3;
        $sheet->setCellValue([$coloumnIndex + $d,     $tableHeadRow], $cpf->name);
        $sheet->setCellValue([$coloumnIndex + $d,     $tableHeadRow + 1], $cpf->code_v002);
        $sheet->setCellValue([$coloumnIndex + $d,     $tableHeadRow + 2], "Койки");
        $sheet->setCellValue([$coloumnIndex + $d + 1, $tableHeadRow + 2], "Объемы, случаев лечения");
        $sheet->setCellValue([$coloumnIndex + $d + 2, $tableHeadRow + 2], "Объемы, пациенто-дней");
        $sheet->setCellValue([$coloumnIndex + $d + 3, $tableHeadRow + 2], "Финансовое обеспечение, руб.");
        $profilesIndex++;
    }
    foreach($moCollection as $mo) {
        $coloumnIndex = $startColoumn;

        $inHospitalBedProfiles = $content['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? null;
        $inHospitalBedProfilesNumberOfBeds = $contentNumberOfBeds['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? null;

        if (!$inHospitalBedProfiles) { continue; }

        $ordinalRowNum++;

        $sheet->setCellValue([$coloumnIndex++, $rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([$coloumnIndex++, $rowIndex], $mo->code);
        $sheet->setCellValue([$coloumnIndex++, $rowIndex], $mo->short_name);

        foreach ($careProfilesFomsUsedInHospital as $cpf) {
            $numberOfBeds = '0';
            $casesOfTreatment = '0';
            $patientDays = '0';
            $cost = '0';

            $hbp = $cpf->hospitalBedProfiles;
            foreach($hbp as $bp) {
                $bpData = $inHospitalBedProfiles[$bp->id] ?? null;
                if (!$bpData) { continue; }

                $casesOfTreatment = bcadd($casesOfTreatment, $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                $patientDays = bcadd($patientDays, $bpData[$patientDaysIndicatorId] ?? '0');
                $cost = bcadd($cost, $bpData[$costIndicatorId] ?? '0');
            }
            foreach($hbp as $bp) {
                $bpData = $inHospitalBedProfilesNumberOfBeds[$bp->id] ?? null;
                if (!$bpData) { continue; }

                $numberOfBeds = bcadd($numberOfBeds, $bpData[$numberOfBedsIndicatorId] ?? '0');
            }

            $sheet->setCellValue([$coloumnIndex++, $rowIndex], $numberOfBeds);
            $sheet->setCellValue([$coloumnIndex++, $rowIndex], $casesOfTreatment);
            $sheet->setCellValue([$coloumnIndex++, $rowIndex], $patientDays);
            $sheet->setCellValue([$coloumnIndex++, $rowIndex], $cost);
        }
        $rowIndex++;
    }


    $spreadsheet->createSheet();
    $spreadsheet->setActiveSheetIndex(1);
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setTitle('ДС при поликлинике');

    $ordinalRowNum = 0;
    $rowIndex = $startRow + 3;
    $tableHeadRow = $startRow;
    $coloumnIndex = $startColoumn;
    $category = 'hospital';

    $numberOfBedsIndicatorId = 1; // число коек
    $casesOfTreatmentIndicatorId = 2; // случаев лечения
    $patientDaysIndicatorId = 3; // пациенто-дней
    $costIndicatorId = 4; // стоимость

    $profilesIndex = 0;
    foreach ($careProfilesFomsUsedInPolyclinic as $cpf) {
        $d = (4 * $profilesIndex) + 3;
        $sheet->setCellValue([$coloumnIndex + $d,     $tableHeadRow], $cpf->name);
        $sheet->setCellValue([$coloumnIndex + $d,     $tableHeadRow + 1], $cpf->code_v002);
        $sheet->setCellValue([$coloumnIndex + $d,     $tableHeadRow + 2], "Койки");
        $sheet->setCellValue([$coloumnIndex + $d + 1, $tableHeadRow + 2], "Объемы, случаев лечения");
        $sheet->setCellValue([$coloumnIndex + $d + 2, $tableHeadRow + 2], "Объемы, пациенто-дней");
        $sheet->setCellValue([$coloumnIndex + $d + 3, $tableHeadRow + 2], "Финансовое обеспечение, руб.");
        $profilesIndex++;
    }
    foreach($moCollection as $mo) {
        $coloumnIndex = $startColoumn;

        $inPolyclinicBedProfiles = $content['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? null;
        $inPolyclinicBedProfilesNumberOfBeds = $contentNumberOfBeds['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? null;

        if (!$inPolyclinicBedProfiles) { continue; }

        $ordinalRowNum++;

        $sheet->setCellValue([$coloumnIndex++, $rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([$coloumnIndex++, $rowIndex], $mo->code);
        $sheet->setCellValue([$coloumnIndex++, $rowIndex], $mo->short_name);

        foreach ($careProfilesFomsUsedInPolyclinic as $cpf) {
            $numberOfBeds = '0';
            $casesOfTreatment = '0';
            $patientDays = '0';
            $cost = '0';

            $hbp = $cpf->hospitalBedProfiles;
            foreach($hbp as $bp) {
                $bpData = $inPolyclinicBedProfiles[$bp->id] ?? null;
                if (!$bpData) { continue; }

                $casesOfTreatment = bcadd($casesOfTreatment, $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                $patientDays = bcadd($patientDays, $bpData[$patientDaysIndicatorId] ?? '0');
                $cost = bcadd($cost, $bpData[$costIndicatorId] ?? '0');
            }
            foreach($hbp as $bp) {
                $bpData = $inPolyclinicBedProfilesNumberOfBeds[$bp->id] ?? null;
                if (!$bpData) { continue; }

                $numberOfBeds = bcadd($numberOfBeds, $bpData[$numberOfBedsIndicatorId] ?? '0');
            }

            $sheet->setCellValue([$coloumnIndex++, $rowIndex], $numberOfBeds);
            $sheet->setCellValue([$coloumnIndex++, $rowIndex], $casesOfTreatment);
            $sheet->setCellValue([$coloumnIndex++, $rowIndex], $patientDays);
            $sheet->setCellValue([$coloumnIndex++, $rowIndex], $cost);
        }
        $rowIndex++;
    }


    $spreadsheet->createSheet();
    $spreadsheet->setActiveSheetIndex(2);
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setTitle('КС');

    $ordinalRowNum = 0;
    $rowIndex = $startRow + 3;
    $tableHeadRow = $startRow;
    $coloumnIndex = $startColoumn;
    $category = 'hospital';

    $numberOfBedsIndicatorId = 1; // число коек
    $casesOfTreatmentIndicatorId = 7; // госпитализаций
    $patientDaysIndicatorId = 3; // пациенто-дней
    $costIndicatorId = 4; // стоимость

    $profilesIndex = 0;
    foreach ($careProfilesFomsUsedInRegular as $cpf) {
        $d = (4 * $profilesIndex) + 3;
        $sheet->setCellValue([$coloumnIndex + $d,     $tableHeadRow], $cpf->name);
        $sheet->setCellValue([$coloumnIndex + $d,     $tableHeadRow + 1], $cpf->code_v002);
        $sheet->setCellValue([$coloumnIndex + $d,     $tableHeadRow + 2], "Койки");
        $sheet->setCellValue([$coloumnIndex + $d + 1, $tableHeadRow + 2], "Объемы, госпитализаций");
        $sheet->setCellValue([$coloumnIndex + $d + 2, $tableHeadRow + 2], "Объемы, койко-дней");
        $sheet->setCellValue([$coloumnIndex + $d + 3, $tableHeadRow + 2], "Финансовое обеспечение, руб.");
        $profilesIndex++;
    }
    foreach($moCollection as $mo) {
        $coloumnIndex = $startColoumn;

        $regularBedProfiles = $content['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? null;
        $regularBedProfilesNumberOfBeds = $contentNumberOfBeds['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? null;

        if (!$regularBedProfiles) { continue; }

        $ordinalRowNum++;

        $sheet->setCellValue([$coloumnIndex++, $rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([$coloumnIndex++, $rowIndex], $mo->code);
        $sheet->setCellValue([$coloumnIndex++, $rowIndex], $mo->short_name);

        foreach ($careProfilesFomsUsedInRegular as $cpf) {
            $numberOfBeds = '0';
            $casesOfTreatment = '0';
            $patientDays = '0';
            $cost = '0';

            $hbp = $cpf->hospitalBedProfiles;
            foreach($hbp as $bp) {
                $bpData = $regularBedProfiles[$bp->id] ?? null;
                if (!$bpData) { continue; }

                $casesOfTreatment = bcadd($casesOfTreatment, $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                $patientDays = bcadd($patientDays, $bpData[$patientDaysIndicatorId] ?? '0');
                $cost = bcadd($cost, $bpData[$costIndicatorId] ?? '0');
            }
            foreach($hbp as $bp) {
                $bpData = $regularBedProfilesNumberOfBeds[$bp->id] ?? null;
                if (!$bpData) { continue; }

                $numberOfBeds = bcadd($numberOfBeds, $bpData[$numberOfBedsIndicatorId] ?? '0');
            }

            $sheet->setCellValue([$coloumnIndex++, $rowIndex], $numberOfBeds);
            $sheet->setCellValue([$coloumnIndex++, $rowIndex], $casesOfTreatment);
            $sheet->setCellValue([$coloumnIndex++, $rowIndex], $patientDays);
            $sheet->setCellValue([$coloumnIndex++, $rowIndex], $cost);
        }
        $rowIndex++;
    }

    $spreadsheet->createSheet();
    $spreadsheet->setActiveSheetIndex(3);
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setTitle('ВМП');

    $ordinalRowNum = 0;
    $rowIndex = $startRow + 3;
    $tableHeadRow = $startRow;
    $coloumnIndex = $startColoumn;
    $category = 'hospital';

    $numberOfBedsIndicatorId = 1; // число коек
    $casesOfTreatmentIndicatorId = 7; // госпитализаций
    $patientDaysIndicatorId = 3; // пациенто-дней
    $costIndicatorId = 4; // стоимость

    $profilesIndex = 0;
    foreach ($careProfilesFomsUsedVmp as $cpf) {
        $d = (4 * $profilesIndex) + 3;
        $sheet->setCellValue([$coloumnIndex + $d,     $tableHeadRow], $cpf->name);
        $sheet->setCellValue([$coloumnIndex + $d,     $tableHeadRow + 1], $cpf->code_v002);
        $sheet->setCellValue([$coloumnIndex + $d,     $tableHeadRow + 2], "Койки");
        $sheet->setCellValue([$coloumnIndex + $d + 1, $tableHeadRow + 2], "Объемы, госпитализаций");
        $sheet->setCellValue([$coloumnIndex + $d + 2, $tableHeadRow + 2], "Объемы, койко-дней");
        $sheet->setCellValue([$coloumnIndex + $d + 3, $tableHeadRow + 2], "Финансовое обеспечение, руб.");
        $profilesIndex++;
    }
    foreach($moCollection as $mo) {
        $coloumnIndex = $startColoumn;

        $careProfiles = $content['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? null;
        $careProfilesNumberOfBeds = $contentNumberOfBeds['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? null;

        if (!$careProfiles) { continue; }

        $ordinalRowNum++;

        $sheet->setCellValue([$coloumnIndex++, $rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([$coloumnIndex++, $rowIndex], $mo->code);
        $sheet->setCellValue([$coloumnIndex++, $rowIndex], $mo->short_name);

        foreach ($careProfilesFomsUsedVmp as $cpf) {
            $numberOfBeds = '0';
            $casesOfTreatment = '0';
            $patientDays = '0';
            $cost = '0';

            $cpmz = $cpf->careProfilesMz;
            foreach($cpmz as $cp) {
                $vmpGroupsData = $careProfiles[$cp->id] ?? null;
                if (!$vmpGroupsData) { continue; }

                foreach ($vmpGroupsData as $vmpTypes) {
                    foreach ($vmpTypes as $vmpT)
                    {
                        $casesOfTreatment = bcadd($casesOfTreatment, $vmpT[$casesOfTreatmentIndicatorId] ?? '0');
                        $patientDays = bcadd($patientDays, $vmpT[$patientDaysIndicatorId] ?? '0');
                        $cost = bcadd($cost, $vmpT[$costIndicatorId] ?? '0');
                    }
                }
            }

            foreach($cpmz as $cp) {
                $vmpGroupsData = $careProfilesNumberOfBeds[$cp->id] ?? null;
                if (!$vmpGroupsData) { continue; }

                foreach ($vmpGroupsData as $vmpTypes) {
                    foreach ($vmpTypes as $vmpT)
                    {
                        $numberOfBeds = bcadd($numberOfBeds, $vmpT[$numberOfBedsIndicatorId] ?? '0');
                    }
                }
            }

            $sheet->setCellValue([$coloumnIndex++, $rowIndex], $numberOfBeds);
            $sheet->setCellValue([$coloumnIndex++, $rowIndex], $casesOfTreatment);
            $sheet->setCellValue([$coloumnIndex++, $rowIndex], $patientDays);
            $sheet->setCellValue([$coloumnIndex++, $rowIndex], $cost);
        }
        $rowIndex++;
    }

    $writer = new Xlsx($spreadsheet);
    $writer->save($fullResultFilepath);
    return Storage::download($resultFilePath);
});

function vitacoreHospitalByProfilePrintRow(
    Worksheet $sheet,
    int $colIndex,
    int $rowIndex,
    string | int $ordinalRowNum,
    string | int $moCode,
    string $moName,
    string $planningSectionName,
    string $profileName,
    string | int $profileCode,
    string $planningParamName,
    string | int $value,
    string | int $vmpGroupCode = ""
 ) {
    $moCodeColOffset = 1;
    $moNameColOffset = 2;
    $planningSectionColOffset = 3;
    $profileNameColOffset = 4;
    $profileCodeColOffset = 5;
    $vmpGroupNumColOffset = 6;
    $paramColOffset = 7;
    $valueColOffset = 8;

    $sheet->setCellValue([$colIndex, $rowIndex], "$ordinalRowNum");
    $sheet->setCellValue([$colIndex + $moCodeColOffset, $rowIndex], $moCode);
    $sheet->setCellValue([$colIndex + $moNameColOffset, $rowIndex], $moName);
    $sheet->setCellValue([$colIndex + $planningSectionColOffset, $rowIndex], $planningSectionName);
    $sheet->setCellValue([$colIndex + $profileNameColOffset, $rowIndex], $profileName);
    $sheet->setCellValue([$colIndex + $profileCodeColOffset, $rowIndex], $profileCode);
    $sheet->setCellValue([$colIndex + $paramColOffset, $rowIndex], $planningParamName);
    $sheet->setCellValue([$colIndex + $valueColOffset, $rowIndex], $value);

    $sheet->setCellValue([$colIndex + $vmpGroupNumColOffset, $rowIndex], $vmpGroupCode);

}

Route::get('/vitacore-hospital-by-profile/{year}/{commissionDecisionsId?}', function (DataForContractService $dataForContractService, int $year, int $commissionDecisionsId = null) {
    $packageIds = null;
    $currentlyUsedDate = $year.'-01-01';
    if ($commissionDecisionsId) {
        $commissionDecisions = CommissionDecision::whereYear('date',$year)->where('id', '<=', $commissionDecisionsId)->get();
        $cd = $commissionDecisions->find($commissionDecisionsId);
        $commissionDecisionIds = $commissionDecisions->pluck('id')->toArray();
        $packageIds = ChangePackage::whereIn('commission_decision_id', $commissionDecisionIds)->orWhere('commission_decision_id', null)->pluck('id')->toArray();

        $currentlyUsedDate = $cd->date->format('Y-m-d');
    } else {
        $packageIds = ChangePackage::where('commission_decision_id', null)->pluck('id')->toArray();
    }
    $path = 'xlsx';
    $resultFileName = 'hospital.xlsx';
    $strDateTimeNow = date("Y-m-d-His");
    $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . ' ' . $resultFileName;
    $fullResultFilepath = Storage::path($resultFilePath);

    bcscale(4);

    $indicatorIds = [2, 3, 4, 7];
    $content = $dataForContractService->GetArray($year, $packageIds, $indicatorIds);
    // количество коек на последний месяц года
    $contentNumberOfBeds = $dataForContractService->GetArrayByYearAndMonth($year, 12, $packageIds, [1]);

    $moCollection = MedicalInstitution::WhereRaw("? BETWEEN effective_from AND effective_to", [$currentlyUsedDate])->orderBy('order')->get();

    $careProfilesFoms = CareProfilesFoms::all();

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    $ordinalRowNum = 1;
    $firstTableDataRowIndex = 3;
    $firstTableColIndex = 1;
    $rowOffset = 0;

    $category = 'hospital';

    foreach($moCollection as $mo) {
        $planningParamNames = [
            1 => "койки",
            2 => "объемы, случаев лечения",
            3 => "объемы, пациенто-дней",
            4 => "финансовое обеспечение, руб.",
            7 => "объемы, госпитализаций"
        ];

        $inHospitalBedProfiles = $content['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? null;
        $inHospitalBedProfilesNumberOfBeds = $contentNumberOfBeds['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? null;
        if ($inHospitalBedProfiles) {
            $planningSectionName = 'Дневные стационары при стационаре';
            $numberOfBedsIndicatorId = 1; // число коек
            $casesOfTreatmentIndicatorId = 2; // случаев лечения
            $patientDaysIndicatorId = 3; // пациенто-дней
            $costIndicatorId = 4; // стоимость
            foreach ($careProfilesFoms as $cpf) {
                $numberOfBeds = '0';
                $casesOfTreatment = '0';
                $patientDays = '0';
                $cost = '0';

                $hbp = $cpf->hospitalBedProfiles;
                foreach($hbp as $bp) {
                    $bpData = $inHospitalBedProfiles[$bp->id] ?? null;
                    $bpDataNumberOfBeds = $inHospitalBedProfilesNumberOfBeds[$bp->id] ?? null;
                    if (!$bpData && !$bpDataNumberOfBeds) { continue; }

                    $numberOfBeds = bcadd($numberOfBeds, $bpDataNumberOfBeds[$numberOfBedsIndicatorId] ?? '0');
                    $casesOfTreatment = bcadd($casesOfTreatment, $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                    $patientDays = bcadd($patientDays, $bpData[$patientDaysIndicatorId] ?? '0');
                    $cost = bcadd($cost, $bpData[$costIndicatorId] ?? '0');
                }

                if (bccomp($numberOfBeds, '0') !== 0) {
                    vitacoreHospitalByProfilePrintRow(
                        $sheet,
                        $firstTableColIndex,
                        $firstTableDataRowIndex + $rowOffset++,
                        $ordinalRowNum++,
                        $mo->code,
                        $mo->short_name,
                        $planningSectionName,
                        $cpf->name,
                        $cpf->code_v002,
                        $planningParamNames[$numberOfBedsIndicatorId],
                        $numberOfBeds
                    );
                }
                if (bccomp($casesOfTreatment, '0') !== 0) {
                    vitacoreHospitalByProfilePrintRow(
                        $sheet,
                        $firstTableColIndex,
                        $firstTableDataRowIndex + $rowOffset++,
                        $ordinalRowNum++,
                        $mo->code,
                        $mo->short_name,
                        $planningSectionName,
                        $cpf->name,
                        $cpf->code_v002,
                        $planningParamNames[$casesOfTreatmentIndicatorId],
                        $casesOfTreatment
                    );
                }
                if (bccomp($patientDays, '0') !== 0) {
                    vitacoreHospitalByProfilePrintRow(
                        $sheet,
                        $firstTableColIndex,
                        $firstTableDataRowIndex + $rowOffset++,
                        $ordinalRowNum++,
                        $mo->code,
                        $mo->short_name,
                        $planningSectionName,
                        $cpf->name,
                        $cpf->code_v002,
                        $planningParamNames[$patientDaysIndicatorId],
                        $patientDays
                    );
                }
                if (bccomp($cost, '0') !== 0) {
                    vitacoreHospitalByProfilePrintRow(
                        $sheet,
                        $firstTableColIndex,
                        $firstTableDataRowIndex + $rowOffset++,
                        $ordinalRowNum++,
                        $mo->code,
                        $mo->short_name,
                        $planningSectionName,
                        $cpf->name,
                        $cpf->code_v002,
                        $planningParamNames[$costIndicatorId],
                        $cost
                    );
                }
            }
        }

        $inPolyclinicBedProfiles = $content['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? null;
        $inPolyclinicBedProfilesNumberOfBeds = $contentNumberOfBeds['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? null;
        if ($inPolyclinicBedProfiles) {
            $planningSectionName = 'Дневные стационары при поликлинике';
            $numberOfBedsIndicatorId = 1; // число коек
            $casesOfTreatmentIndicatorId = 2; // случаев лечения
            $patientDaysIndicatorId = 3; // пациенто-дней
            $costIndicatorId = 4; // стоимость

            foreach ($careProfilesFoms as $cpf) {
                $numberOfBeds = '0';
                $casesOfTreatment = '0';
                $patientDays = '0';
                $cost = '0';

                $hbp = $cpf->hospitalBedProfiles;
                foreach($hbp as $bp) {
                    $bpData = $inPolyclinicBedProfiles[$bp->id] ?? null;
                    $bpDataNumberOfBeds = $inPolyclinicBedProfilesNumberOfBeds[$bp->id] ?? null;
                    if (!$bpData && !$bpDataNumberOfBeds) { continue; }

                    $numberOfBeds = bcadd($numberOfBeds, $bpDataNumberOfBeds[$numberOfBedsIndicatorId] ?? '0');
                    $casesOfTreatment = bcadd($casesOfTreatment, $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                    $patientDays = bcadd($patientDays, $bpData[$patientDaysIndicatorId] ?? '0');
                    $cost = bcadd($cost, $bpData[$costIndicatorId] ?? '0');
                }

                if (bccomp($numberOfBeds, '0') !== 0) {
                    vitacoreHospitalByProfilePrintRow(
                        $sheet,
                        $firstTableColIndex,
                        $firstTableDataRowIndex + $rowOffset++,
                        $ordinalRowNum++,
                        $mo->code,
                        $mo->short_name,
                        $planningSectionName,
                        $cpf->name,
                        $cpf->code_v002,
                        $planningParamNames[$numberOfBedsIndicatorId],
                        $numberOfBeds
                    );
                }
                if (bccomp($casesOfTreatment, '0') !== 0) {
                    vitacoreHospitalByProfilePrintRow(
                        $sheet,
                        $firstTableColIndex,
                        $firstTableDataRowIndex + $rowOffset++,
                        $ordinalRowNum++,
                        $mo->code,
                        $mo->short_name,
                        $planningSectionName,
                        $cpf->name,
                        $cpf->code_v002,
                        $planningParamNames[$casesOfTreatmentIndicatorId],
                        $casesOfTreatment
                    );
                }
                if (bccomp($patientDays, '0') !== 0) {
                    vitacoreHospitalByProfilePrintRow(
                        $sheet,
                        $firstTableColIndex,
                        $firstTableDataRowIndex + $rowOffset++,
                        $ordinalRowNum++,
                        $mo->code,
                        $mo->short_name,
                        $planningSectionName,
                        $cpf->name,
                        $cpf->code_v002,
                        $planningParamNames[$patientDaysIndicatorId],
                        $patientDays
                    );
                }
                if (bccomp($cost, '0') !== 0) {
                    vitacoreHospitalByProfilePrintRow(
                        $sheet,
                        $firstTableColIndex,
                        $firstTableDataRowIndex + $rowOffset++,
                        $ordinalRowNum++,
                        $mo->code,
                        $mo->short_name,
                        $planningSectionName,
                        $cpf->name,
                        $cpf->code_v002,
                        $planningParamNames[$costIndicatorId],
                        $cost
                    );
                }
            }
        }

        $planningParamNames[3] = "объемы, койко-дней";

        $regularBedProfiles = $content['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? null;
        $regularBedProfilesNumberOfBeds = $contentNumberOfBeds['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? null;
        $planningSectionName = 'Круглосуточный стационар (не включая ВМП)';
        if ($regularBedProfiles) {
            $numberOfBedsIndicatorId = 1; // число коек
            $casesOfTreatmentIndicatorId = 7; // госпитализаций
            $patientDaysIndicatorId = 3; // койко-дней
            $costIndicatorId = 4; // стоимость

            foreach ($careProfilesFoms as $cpf) {
                $numberOfBeds = '0';
                $casesOfTreatment = '0';
                $patientDays = '0';
                $cost = '0';

                $hbp = $cpf->hospitalBedProfiles;
                foreach($hbp as $bp) {
                    $bpData = $regularBedProfiles[$bp->id] ?? null;
                    $bpDataNumberOfBeds = $regularBedProfilesNumberOfBeds[$bp->id] ?? null;
                    if (!$bpData && !$bpDataNumberOfBeds) { continue; }

                    $numberOfBeds = bcadd($numberOfBeds, $bpDataNumberOfBeds[$numberOfBedsIndicatorId] ?? '0');
                    $casesOfTreatment = bcadd($casesOfTreatment, $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                    $patientDays = bcadd($patientDays, $bpData[$patientDaysIndicatorId] ?? '0');
                    $cost = bcadd($cost, $bpData[$costIndicatorId] ?? '0');
                }

                if (bccomp($numberOfBeds, '0') !== 0) {
                    vitacoreHospitalByProfilePrintRow(
                        $sheet,
                        $firstTableColIndex,
                        $firstTableDataRowIndex + $rowOffset++,
                        $ordinalRowNum++,
                        $mo->code,
                        $mo->short_name,
                        $planningSectionName,
                        $cpf->name,
                        $cpf->code_v002,
                        $planningParamNames[$numberOfBedsIndicatorId],
                        $numberOfBeds
                    );
                }
                if (bccomp($casesOfTreatment, '0') !== 0) {
                    vitacoreHospitalByProfilePrintRow(
                        $sheet,
                        $firstTableColIndex,
                        $firstTableDataRowIndex + $rowOffset++,
                        $ordinalRowNum++,
                        $mo->code,
                        $mo->short_name,
                        $planningSectionName,
                        $cpf->name,
                        $cpf->code_v002,
                        $planningParamNames[$casesOfTreatmentIndicatorId],
                        $casesOfTreatment
                    );
                }
                if (bccomp($patientDays, '0') !== 0) {
                    vitacoreHospitalByProfilePrintRow(
                        $sheet,
                        $firstTableColIndex,
                        $firstTableDataRowIndex + $rowOffset++,
                        $ordinalRowNum++,
                        $mo->code,
                        $mo->short_name,
                        $planningSectionName,
                        $cpf->name,
                        $cpf->code_v002,
                        $planningParamNames[$patientDaysIndicatorId],
                        $patientDays
                    );
                }
                if (bccomp($cost, '0') !== 0) {
                    vitacoreHospitalByProfilePrintRow(
                        $sheet,
                        $firstTableColIndex,
                        $firstTableDataRowIndex + $rowOffset++,
                        $ordinalRowNum++,
                        $mo->code,
                        $mo->short_name,
                        $planningSectionName,
                        $cpf->name,
                        $cpf->code_v002,
                        $planningParamNames[$costIndicatorId],
                        $cost
                    );
                }
            }
        }

        $vmpGroups = VmpGroup::all();
        $careProfiles = $content['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? null;
        $careProfilesNumberOfBeds = $contentNumberOfBeds['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? null;
        if ($careProfiles) {
            $planningSectionName = 'ВМП';
            $numberOfBedsIndicatorId = 1; // число коек
            $casesOfTreatmentIndicatorId = 7; // госпитализаций
            $patientDaysIndicatorId = 3; // койко-дней
            $costIndicatorId = 4; // стоимость

            foreach ($careProfilesFoms as $cpf) {

                // $numberOfBeds = '0';
                // $casesOfTreatment = '0';
                // $patientDays = '0';
                // $cost = '0';

                $cpmz = $cpf->careProfilesMz;
                foreach($cpmz as $cp) {
                    $vmpGroupsData = $careProfiles[$cp->id] ?? null;

                    if ($vmpGroupsData) {
                        foreach ($vmpGroupsData as $vmpGroupId => $vmpTypes) {
                            $vmpTValueCasesOfTreatment = '0';
                            $vmpTValuePatientDays = '0';
                            $vmpTValueCost = '0';
                            foreach ($vmpTypes as $vmpT)
                            {
                                //$casesOfTreatment = bcadd($casesOfTreatment, $vmpT[$casesOfTreatmentIndicatorId] ?? '0');
                                //$patientDays = bcadd($patientDays, $vmpT[$patientDaysIndicatorId] ?? '0');
                                //$cost = bcadd($cost, $vmpT[$costIndicatorId] ?? '0');
                                $vmpTValueCasesOfTreatment = bcadd($vmpTValueCasesOfTreatment, $vmpT[$casesOfTreatmentIndicatorId] ?? '0');
                                $vmpTValuePatientDays = bcadd($vmpTValuePatientDays, $vmpT[$patientDaysIndicatorId] ?? '0');
                                $vmpTValueCost = bcadd($vmpTValueCost, $vmpT[$costIndicatorId] ?? '0');
                            }
                            if (bccomp($vmpTValueCasesOfTreatment, '0') !== 0) {
                                vitacoreHospitalByProfilePrintRow(
                                    $sheet,
                                    $firstTableColIndex,
                                    $firstTableDataRowIndex + $rowOffset++,
                                    $ordinalRowNum++,
                                    $mo->code,
                                    $mo->short_name,
                                    $planningSectionName,
                                    $cpf->name,
                                    $cpf->code_v002,
                                    $planningParamNames[$casesOfTreatmentIndicatorId],
                                    $vmpTValueCasesOfTreatment,
                                    $vmpGroups->find($vmpGroupId)->code
                                );
                            }
                            if (bccomp($vmpTValuePatientDays, '0') !== 0) {
                                vitacoreHospitalByProfilePrintRow(
                                    $sheet,
                                    $firstTableColIndex,
                                    $firstTableDataRowIndex + $rowOffset++,
                                    $ordinalRowNum++,
                                    $mo->code,
                                    $mo->short_name,
                                    $planningSectionName,
                                    $cpf->name,
                                    $cpf->code_v002,
                                    $planningParamNames[$patientDaysIndicatorId],
                                    $vmpTValuePatientDays,
                                    $vmpGroups->find($vmpGroupId)->code
                                );
                            }
                            if (bccomp($vmpTValueCost, '0') !== 0) {
                                vitacoreHospitalByProfilePrintRow(
                                    $sheet,
                                    $firstTableColIndex,
                                    $firstTableDataRowIndex + $rowOffset++,
                                    $ordinalRowNum++,
                                    $mo->code,
                                    $mo->short_name,
                                    $planningSectionName,
                                    $cpf->name,
                                    $cpf->code_v002,
                                    $planningParamNames[$costIndicatorId],
                                    $vmpTValueCost,
                                    $vmpGroups->find($vmpGroupId)->code
                                );
                            }
                        }
                    }

                    $vmpGroupsDataNumberOfBeds = $careProfilesNumberOfBeds[$cp->id] ?? null;
                    if ($vmpGroupsDataNumberOfBeds) {
                        foreach ($vmpGroupsDataNumberOfBeds as $vmpGroupId => $vmpTypes) {
                            $vmpGroupValue = "0";
                            foreach ($vmpTypes as $vmpT)
                            {
                                //$numberOfBeds = bcadd($numberOfBeds, $vmpT[$numberOfBedsIndicatorId] ?? '0');
                                $vmpGroupValue = bcadd($vmpGroupValue, $vmpT[$numberOfBedsIndicatorId] ?? '0');

                            }
                            if (bccomp($vmpGroupValue, '0') !== 0) {
                                vitacoreHospitalByProfilePrintRow(
                                    $sheet,
                                    $firstTableColIndex,
                                    $firstTableDataRowIndex + $rowOffset++,
                                    $ordinalRowNum++,
                                    $mo->code,
                                    $mo->short_name,
                                    $planningSectionName,
                                    $cpf->name,
                                    $cpf->code_v002,
                                    $planningParamNames[$numberOfBedsIndicatorId],
                                    $vmpGroupValue,
                                    $vmpGroups->find($vmpGroupId)->code
                                );
                            }
                        }
                    }
                }
            }
        }
    } // foreach MO

    $writer = new Xlsx($spreadsheet);
    $writer->save($fullResultFilepath);
    return Storage::download($resultFilePath);
});


Route::get('/vitacore-hospital-by-profile-periods/{year}/{commissionDecisionsId?}', function (DataForContractService $dataForContractService, int $year, int $commissionDecisionsId = null) {
    $packageIds = null;
    $currentlyUsedDate = $year.'-01-01';
    if ($commissionDecisionsId) {
        $commissionDecisions = CommissionDecision::whereYear('date',$year)->where('id', '<=', $commissionDecisionsId)->get();
        $cd = $commissionDecisions->find($commissionDecisionsId);
        $commissionDecisionIds = $commissionDecisions->pluck('id')->toArray();
        $packageIds = ChangePackage::whereIn('commission_decision_id', $commissionDecisionIds)->orWhere('commission_decision_id', null)->pluck('id')->toArray();

        $currentlyUsedDate = $cd->date->format('Y-m-d');
    } else {
        $packageIds = ChangePackage::where('commission_decision_id', null)->pluck('id')->toArray();
    }
    $path = 'xlsx';
    $resultFileName = 'hospital-periods.xlsx';
    $strDateTimeNow = date("Y-m-d-His");
    $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . ' ' . $resultFileName;
    $fullResultFilepath = Storage::path($resultFilePath);

    bcscale(4);

    $indicatorIds = [1, 2, 3, 4, 7];
    $contentByMonth = [];
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $contentByMonth[$monthNum] = $dataForContractService->GetArrayByYearAndMonth($year, $monthNum, $packageIds, $indicatorIds);
    }

    $moCollection = MedicalInstitution::WhereRaw("? BETWEEN effective_from AND effective_to", [$currentlyUsedDate])->orderBy('order')->get();

    $careProfilesFoms = CareProfilesFoms::all();

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    $ordinalRowNum = 1;
    $firstTableDataRowIndex = 3;
    $firstTableColIndex = 1;
    $rowOffset = 0;
    $firstTableHeadRowIndex = 2;

    $category = 'hospital';

    vitacoreHospitalByProfilePeriodsPrintTableHeader($sheet, $firstTableColIndex, $firstTableHeadRowIndex);

    foreach($moCollection as $mo) {
        $planningParamNames = [
            1 => "койки",
            2 => "объемы, случаев лечения",
            3 => "объемы, пациенто-дней",
            4 => "финансовое обеспечение, руб.",
            7 => "объемы, госпитализаций"
        ];

        // Дневные стационары при стационаре
        $numberOfBeds = [];
        $casesOfTreatment = [];
        $patientDays = [];
        $cost = [];

        foreach ($careProfilesFoms as $cpf) {
            $numberOfBeds[$cpf->id] = [];
            $casesOfTreatment[$cpf->id] = [];
            $patientDays[$cpf->id] = [];
            $cost[$cpf->id] = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $numberOfBeds[$cpf->id][$monthNum] = '0';
                $casesOfTreatment[$cpf->id][$monthNum] = '0';
                $patientDays[$cpf->id][$monthNum] = '0';
                $cost[$cpf->id][$monthNum] = '0';
            }
        }

        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $inHospitalBedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? null;

            if ($inHospitalBedProfiles) {
                $planningSectionName = 'Дневные стационары при стационаре';
                $numberOfBedsIndicatorId = 1; // число коек
                $casesOfTreatmentIndicatorId = 2; // случаев лечения
                $patientDaysIndicatorId = 3; // пациенто-дней
                $costIndicatorId = 4; // стоимость

                foreach ($careProfilesFoms as $cpf) {
                    $hbp = $cpf->hospitalBedProfiles;

                    foreach($hbp as $bp) {
                        $bpData = $inHospitalBedProfiles[$bp->id] ?? null;
                        if (!$bpData) { continue; }

                        $numberOfBeds[$cpf->id][$monthNum] = bcadd($numberOfBeds[$cpf->id][$monthNum], $bpData[$numberOfBedsIndicatorId] ?? '0');
                        $casesOfTreatment[$cpf->id][$monthNum] = bcadd($casesOfTreatment[$cpf->id][$monthNum], $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                        $patientDays[$cpf->id][$monthNum] = bcadd($patientDays[$cpf->id][$monthNum], $bpData[$patientDaysIndicatorId] ?? '0');
                        $cost[$cpf->id][$monthNum] = bcadd($cost[$cpf->id][$monthNum], $bpData[$costIndicatorId] ?? '0');
                    }
                }
            }
        }

        foreach ($careProfilesFoms as $cpf) {

            $cpNumberOfBeds = $numberOfBeds[$cpf->id];
            $cpCasesOfTreatment = $casesOfTreatment[$cpf->id];
            $cpPatientDays = $patientDays[$cpf->id];
            $cpCost = $cost[$cpf->id];

            $numberOfBedsHasValue = false;
            $casesOfTreatmentHasValue = false;
            $patientDaysHasValue = false;
            $costHasValue = false;
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $numberOfBedsHasValue = $numberOfBedsHasValue || bccomp($cpNumberOfBeds[$monthNum], '0') !== 0;
                $casesOfTreatmentHasValue = $casesOfTreatmentHasValue || bccomp($cpCasesOfTreatment[$monthNum], '0') !== 0;
                $patientDaysHasValue = $patientDaysHasValue || bccomp($cpPatientDays[$monthNum], '0') !== 0;
                $costHasValue = $costHasValue || bccomp($cpCost[$monthNum], '0') !== 0;
            }

            if ($numberOfBedsHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $cpf->name,
                    $cpf->code_v002,
                    $planningParamNames[$numberOfBedsIndicatorId],
                    $cpNumberOfBeds
                );
            }
            if ($casesOfTreatmentHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $cpf->name,
                    $cpf->code_v002,
                    $planningParamNames[$casesOfTreatmentIndicatorId],
                    $cpCasesOfTreatment
                );
            }
            if ($patientDaysHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $cpf->name,
                    $cpf->code_v002,
                    $planningParamNames[$patientDaysIndicatorId],
                    $cpPatientDays
                );
            }
            if ($costHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $cpf->name,
                    $cpf->code_v002,
                    $planningParamNames[$costIndicatorId],
                    $cpCost
                );
            }
        }

        // Дневные стационары при поликлинике
        $numberOfBeds = [];
        $casesOfTreatment = [];
        $patientDays = [];
        $cost = [];

        foreach ($careProfilesFoms as $cpf) {
            $numberOfBeds[$cpf->id] = [];
            $casesOfTreatment[$cpf->id] = [];
            $patientDays[$cpf->id] = [];
            $cost[$cpf->id] = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $numberOfBeds[$cpf->id][$monthNum] = '0';
                $casesOfTreatment[$cpf->id][$monthNum] = '0';
                $patientDays[$cpf->id][$monthNum] = '0';
                $cost[$cpf->id][$monthNum] = '0';
            }
        }

        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $inPolyclinicBedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? null;

            if ($inPolyclinicBedProfiles) {
                $planningSectionName = 'Дневные стационары при поликлинике';
                $numberOfBedsIndicatorId = 1; // число коек
                $casesOfTreatmentIndicatorId = 2; // случаев лечения
                $patientDaysIndicatorId = 3; // пациенто-дней
                $costIndicatorId = 4; // стоимость

                foreach ($careProfilesFoms as $cpf) {
                    $hbp = $cpf->hospitalBedProfiles;

                    foreach($hbp as $bp) {
                        $bpData = $inPolyclinicBedProfiles[$bp->id] ?? null;
                        if (!$bpData) { continue; }

                        $numberOfBeds[$cpf->id][$monthNum] = bcadd($numberOfBeds[$cpf->id][$monthNum], $bpData[$numberOfBedsIndicatorId] ?? '0');
                        $casesOfTreatment[$cpf->id][$monthNum] = bcadd($casesOfTreatment[$cpf->id][$monthNum], $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                        $patientDays[$cpf->id][$monthNum] = bcadd($patientDays[$cpf->id][$monthNum], $bpData[$patientDaysIndicatorId] ?? '0');
                        $cost[$cpf->id][$monthNum] = bcadd($cost[$cpf->id][$monthNum], $bpData[$costIndicatorId] ?? '0');
                    }
                }
            }
        }

        foreach ($careProfilesFoms as $cpf) {

            $cpNumberOfBeds = $numberOfBeds[$cpf->id];
            $cpCasesOfTreatment = $casesOfTreatment[$cpf->id];
            $cpPatientDays = $patientDays[$cpf->id];
            $cpCost = $cost[$cpf->id];

            $numberOfBedsHasValue = false;
            $casesOfTreatmentHasValue = false;
            $patientDaysHasValue = false;
            $costHasValue = false;
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $numberOfBedsHasValue = $numberOfBedsHasValue || bccomp($cpNumberOfBeds[$monthNum], '0') !== 0;
                $casesOfTreatmentHasValue = $casesOfTreatmentHasValue || bccomp($cpCasesOfTreatment[$monthNum], '0') !== 0;
                $patientDaysHasValue = $patientDaysHasValue || bccomp($cpPatientDays[$monthNum], '0') !== 0;
                $costHasValue = $costHasValue || bccomp($cpCost[$monthNum], '0') !== 0;
            }

            if ($numberOfBedsHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $cpf->name,
                    $cpf->code_v002,
                    $planningParamNames[$numberOfBedsIndicatorId],
                    $cpNumberOfBeds
                );
            }
            if ($casesOfTreatmentHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $cpf->name,
                    $cpf->code_v002,
                    $planningParamNames[$casesOfTreatmentIndicatorId],
                    $cpCasesOfTreatment
                );
            }
            if ($patientDaysHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $cpf->name,
                    $cpf->code_v002,
                    $planningParamNames[$patientDaysIndicatorId],
                    $cpPatientDays
                );
            }
            if ($costHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $cpf->name,
                    $cpf->code_v002,
                    $planningParamNames[$costIndicatorId],
                    $cpCost
                );
            }
        }

        // Круглосуточный стационар (не включая ВМП)
        $planningParamNames[3] = "объемы, койко-дней";

        $numberOfBeds = [];
        $casesOfTreatment = [];
        $patientDays = [];
        $cost = [];

        foreach ($careProfilesFoms as $cpf) {
            $numberOfBeds[$cpf->id] = [];
            $casesOfTreatment[$cpf->id] = [];
            $patientDays[$cpf->id] = [];
            $cost[$cpf->id] = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $numberOfBeds[$cpf->id][$monthNum] = '0';
                $casesOfTreatment[$cpf->id][$monthNum] = '0';
                $patientDays[$cpf->id][$monthNum] = '0';
                $cost[$cpf->id][$monthNum] = '0';
            }
        }

        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $regularBedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? null;

            if ($regularBedProfiles) {
                $planningSectionName = 'Круглосуточный стационар (не включая ВМП)';
                $numberOfBedsIndicatorId = 1; // число коек
                $casesOfTreatmentIndicatorId = 7; // госпитализаций
                $patientDaysIndicatorId = 3; // койко-дней
                $costIndicatorId = 4; // стоимость

                foreach ($careProfilesFoms as $cpf) {
                    $hbp = $cpf->hospitalBedProfiles;

                    foreach($hbp as $bp) {
                        $bpData = $regularBedProfiles[$bp->id] ?? null;
                        if (!$bpData) { continue; }

                        $numberOfBeds[$cpf->id][$monthNum] = bcadd($numberOfBeds[$cpf->id][$monthNum], $bpData[$numberOfBedsIndicatorId] ?? '0');
                        $casesOfTreatment[$cpf->id][$monthNum] = bcadd($casesOfTreatment[$cpf->id][$monthNum], $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                        $patientDays[$cpf->id][$monthNum] = bcadd($patientDays[$cpf->id][$monthNum], $bpData[$patientDaysIndicatorId] ?? '0');
                        $cost[$cpf->id][$monthNum] = bcadd($cost[$cpf->id][$monthNum], $bpData[$costIndicatorId] ?? '0');
                    }
                }
            }
        }

        foreach ($careProfilesFoms as $cpf) {

            $cpNumberOfBeds = $numberOfBeds[$cpf->id];
            $cpCasesOfTreatment = $casesOfTreatment[$cpf->id];
            $cpPatientDays = $patientDays[$cpf->id];
            $cpCost = $cost[$cpf->id];

            $numberOfBedsHasValue = false;
            $casesOfTreatmentHasValue = false;
            $patientDaysHasValue = false;
            $costHasValue = false;
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $numberOfBedsHasValue = $numberOfBedsHasValue || bccomp($cpNumberOfBeds[$monthNum], '0') !== 0;
                $casesOfTreatmentHasValue = $casesOfTreatmentHasValue || bccomp($cpCasesOfTreatment[$monthNum], '0') !== 0;
                $patientDaysHasValue = $patientDaysHasValue || bccomp($cpPatientDays[$monthNum], '0') !== 0;
                $costHasValue = $costHasValue || bccomp($cpCost[$monthNum], '0') !== 0;
            }

            if ($numberOfBedsHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $cpf->name,
                    $cpf->code_v002,
                    $planningParamNames[$numberOfBedsIndicatorId],
                    $cpNumberOfBeds
                );
            }
            if ($casesOfTreatmentHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $cpf->name,
                    $cpf->code_v002,
                    $planningParamNames[$casesOfTreatmentIndicatorId],
                    $cpCasesOfTreatment
                );
            }
            if ($patientDaysHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $cpf->name,
                    $cpf->code_v002,
                    $planningParamNames[$patientDaysIndicatorId],
                    $cpPatientDays
                );
            }
            if ($costHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $cpf->name,
                    $cpf->code_v002,
                    $planningParamNames[$costIndicatorId],
                    $cpCost
                );
            }
        }

        // ВМП
        $planningParamNames[3] = "объемы, койко-дней";

        $numberOfBeds = [];
        $casesOfTreatment = [];
        $patientDays = [];
        $cost = [];

        foreach ($careProfilesFoms as $cpf) {
            $numberOfBeds[$cpf->id] = [];
            $casesOfTreatment[$cpf->id] = [];
            $patientDays[$cpf->id] = [];
            $cost[$cpf->id] = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $numberOfBeds[$cpf->id][$monthNum] = '0';
                $casesOfTreatment[$cpf->id][$monthNum] = '0';
                $patientDays[$cpf->id][$monthNum] = '0';
                $cost[$cpf->id][$monthNum] = '0';
            }
        }

        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $careProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? null;

            if ($careProfiles) {
                $planningSectionName = 'ВМП';
                $numberOfBedsIndicatorId = 1; // число коек
                $casesOfTreatmentIndicatorId = 7; // госпитализаций
                $patientDaysIndicatorId = 3; // койко-дней
                $costIndicatorId = 4; // стоимость

                foreach ($careProfilesFoms as $cpf) {
                    $cpmz = $cpf->careProfilesMz;

                    foreach($cpmz as $cp) {
                        $vmpGroupsData = $careProfiles[$cp->id] ?? null;
                        if (!$vmpGroupsData) { continue; }

                        foreach ($vmpGroupsData as $vmpTypes) {
                            foreach ($vmpTypes as $vmpT)
                            {
                                $numberOfBeds[$cpf->id][$monthNum] = bcadd($numberOfBeds[$cpf->id][$monthNum], $vmpT[$numberOfBedsIndicatorId] ?? '0');
                                $casesOfTreatment[$cpf->id][$monthNum] = bcadd($casesOfTreatment[$cpf->id][$monthNum], $vmpT[$casesOfTreatmentIndicatorId] ?? '0');
                                $patientDays[$cpf->id][$monthNum] = bcadd($patientDays[$cpf->id][$monthNum], $vmpT[$patientDaysIndicatorId] ?? '0');
                                $cost[$cpf->id][$monthNum] = bcadd($cost[$cpf->id][$monthNum], $vmpT[$costIndicatorId] ?? '0');
                            }
                        }
                    }
                }
            }
        }

        foreach ($careProfilesFoms as $cpf) {
            $cpNumberOfBeds = $numberOfBeds[$cpf->id];
            $cpCasesOfTreatment = $casesOfTreatment[$cpf->id];
            $cpPatientDays = $patientDays[$cpf->id];
            $cpCost = $cost[$cpf->id];

            $numberOfBedsHasValue = false;
            $casesOfTreatmentHasValue = false;
            $patientDaysHasValue = false;
            $costHasValue = false;
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $numberOfBedsHasValue = $numberOfBedsHasValue || bccomp($cpNumberOfBeds[$monthNum], '0') !== 0;
                $casesOfTreatmentHasValue = $casesOfTreatmentHasValue || bccomp($cpCasesOfTreatment[$monthNum], '0') !== 0;
                $patientDaysHasValue = $patientDaysHasValue || bccomp($cpPatientDays[$monthNum], '0') !== 0;
                $costHasValue = $costHasValue || bccomp($cpCost[$monthNum], '0') !== 0;
            }

            if ($numberOfBedsHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $cpf->name,
                    $cpf->code_v002,
                    $planningParamNames[$numberOfBedsIndicatorId],
                    $cpNumberOfBeds
                );
            }
            if ($casesOfTreatmentHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $cpf->name,
                    $cpf->code_v002,
                    $planningParamNames[$casesOfTreatmentIndicatorId],
                    $cpCasesOfTreatment
                );
            }
            if ($patientDaysHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $cpf->name,
                    $cpf->code_v002,
                    $planningParamNames[$patientDaysIndicatorId],
                    $cpPatientDays
                );
            }
            if ($costHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $cpf->name,
                    $cpf->code_v002,
                    $planningParamNames[$costIndicatorId],
                    $cpCost
                );
            }
        }

    } // foreach MO

    $writer = new Xlsx($spreadsheet);
    $writer->save($fullResultFilepath);
    return Storage::download($resultFilePath);
});

Route::get('/vitacore-hospital-by-bed-profile-periods/{year}/{commissionDecisionsId?}', function (DataForContractService $dataForContractService, int $year, int $commissionDecisionsId = null) {
    $packageIds = null;
    $currentlyUsedDate = $year.'-01-01';
    if ($commissionDecisionsId) {
        $commissionDecisions = CommissionDecision::whereYear('date',$year)->where('id', '<=', $commissionDecisionsId)->get();
        $cd = $commissionDecisions->find($commissionDecisionsId);
        $commissionDecisionIds = $commissionDecisions->pluck('id')->toArray();
        $packageIds = ChangePackage::whereIn('commission_decision_id', $commissionDecisionIds)->orWhere('commission_decision_id', null)->pluck('id')->toArray();

        $currentlyUsedDate = $cd->date->format('Y-m-d');
    } else {
        $packageIds = ChangePackage::where('commission_decision_id', null)->pluck('id')->toArray();
    }
    $path = 'xlsx';
    $resultFileName = 'hospital-periods.xlsx';
    $strDateTimeNow = date("Y-m-d-His");
    $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . ' ' . $resultFileName;
    $fullResultFilepath = Storage::path($resultFilePath);

    bcscale(4);

    $indicatorIds = [2, 3, 4, 7, 1]; //  - количество коек (убрано, не грузится витакор)
    $contentByMonth = [];
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $contentByMonth[$monthNum] = $dataForContractService->GetArrayByYearAndMonth($year, $monthNum, $packageIds, $indicatorIds);
    }

    $moCollection = MedicalInstitution::WhereRaw("? BETWEEN effective_from AND effective_to", [$currentlyUsedDate])->orderBy('order')->get();

    $hospitalBedProfiles = HospitalBedProfiles::all();

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    $ordinalRowNum = 1;
    $firstTableDataRowIndex = 3;
    $firstTableColIndex = 1;
    $rowOffset = 0;
    $firstTableHeadRowIndex = 2;

    $category = 'hospital';

    vitacoreHospitalByProfilePeriodsPrintTableHeader($sheet, $firstTableColIndex, $firstTableHeadRowIndex);

    foreach($moCollection as $mo) {
        $planningParamNames = [
            1 => "койки",
            2 => "объемы, случаев лечения",
            3 => "объемы, пациенто-дней",
            4 => "финансовое обеспечение, руб.",
            7 => "объемы, госпитализаций"
        ];

        // Дневные стационары при стационаре
        $numberOfBeds = [];
        $casesOfTreatment = [];
        $patientDays = [];
        $cost = [];

        foreach ($hospitalBedProfiles as $hbp) {
            $numberOfBeds[$hbp->id] = [];
            $casesOfTreatment[$hbp->id] = [];
            $patientDays[$hbp->id] = [];
            $cost[$hbp->id] = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $numberOfBeds[$hbp->id][$monthNum] = '0';
                $casesOfTreatment[$hbp->id][$monthNum] = '0';
                $patientDays[$hbp->id][$monthNum] = '0';
                $cost[$hbp->id][$monthNum] = '0';
            }
        }

        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $inHospitalBedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? null;

            if ($inHospitalBedProfiles) {
                $planningSectionName = 'Дневные стационары при стационаре';
                $numberOfBedsIndicatorId = 1; // число коек
                $casesOfTreatmentIndicatorId = 2; // случаев лечения
                $patientDaysIndicatorId = 3; // пациенто-дней
                $costIndicatorId = 4; // стоимость

                foreach ($hospitalBedProfiles as $hbp) {
                    $bpData = $inHospitalBedProfiles[$hbp->id] ?? null;
                    if (!$bpData) { continue; }

                    $numberOfBeds[$hbp->id][$monthNum] = bcadd($numberOfBeds[$hbp->id][$monthNum], $bpData[$numberOfBedsIndicatorId] ?? '0');
                    $casesOfTreatment[$hbp->id][$monthNum] = bcadd($casesOfTreatment[$hbp->id][$monthNum], $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                    $patientDays[$hbp->id][$monthNum] = bcadd($patientDays[$hbp->id][$monthNum], $bpData[$patientDaysIndicatorId] ?? '0');
                    $cost[$hbp->id][$monthNum] = bcadd($cost[$hbp->id][$monthNum], $bpData[$costIndicatorId] ?? '0');
                }
            }
        }

        foreach ($hospitalBedProfiles as $hbp) {

            $cpNumberOfBeds = $numberOfBeds[$hbp->id];
            $cpCasesOfTreatment = $casesOfTreatment[$hbp->id];
            $cpPatientDays = $patientDays[$hbp->id];
            $cpCost = $cost[$hbp->id];

            $numberOfBedsHasValue = false;
            $casesOfTreatmentHasValue = false;
            $patientDaysHasValue = false;
            $costHasValue = false;
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $numberOfBedsHasValue = $numberOfBedsHasValue || bccomp($cpNumberOfBeds[$monthNum], '0') !== 0;
                $casesOfTreatmentHasValue = $casesOfTreatmentHasValue || bccomp($cpCasesOfTreatment[$monthNum], '0') !== 0;
                $patientDaysHasValue = $patientDaysHasValue || bccomp($cpPatientDays[$monthNum], '0') !== 0;
                $costHasValue = $costHasValue || bccomp($cpCost[$monthNum], '0') !== 0;
            }

            if ($numberOfBedsHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $hbp->name,
                    $hbp->code,
                    $planningParamNames[$numberOfBedsIndicatorId],
                    $cpNumberOfBeds
                );
            }
            if ($casesOfTreatmentHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $hbp->name,
                    $hbp->code,
                    $planningParamNames[$casesOfTreatmentIndicatorId],
                    $cpCasesOfTreatment
                );
            }
            if ($patientDaysHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $hbp->name,
                    $hbp->code,
                    $planningParamNames[$patientDaysIndicatorId],
                    $cpPatientDays
                );
            }
            if ($costHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $hbp->name,
                    $hbp->code,
                    $planningParamNames[$costIndicatorId],
                    $cpCost
                );
            }
        }

        // Дневные стационары при поликлинике
        $numberOfBeds = [];
        $casesOfTreatment = [];
        $patientDays = [];
        $cost = [];

        foreach ($hospitalBedProfiles as $hbp) {
            $numberOfBeds[$hbp->id] = [];
            $casesOfTreatment[$hbp->id] = [];
            $patientDays[$hbp->id] = [];
            $cost[$hbp->id] = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $numberOfBeds[$hbp->id][$monthNum] = '0';
                $casesOfTreatment[$hbp->id][$monthNum] = '0';
                $patientDays[$hbp->id][$monthNum] = '0';
                $cost[$hbp->id][$monthNum] = '0';
            }
        }

        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $inPolyclinicBedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? null;

            if ($inPolyclinicBedProfiles) {
                $planningSectionName = 'Дневные стационары при поликлинике';
                $numberOfBedsIndicatorId = 1; // число коек
                $casesOfTreatmentIndicatorId = 2; // случаев лечения
                $patientDaysIndicatorId = 3; // пациенто-дней
                $costIndicatorId = 4; // стоимость

                foreach ($hospitalBedProfiles as $hbp) {
                        $bpData = $inPolyclinicBedProfiles[$hbp->id] ?? null;
                        if (!$bpData) { continue; }

                        $numberOfBeds[$hbp->id][$monthNum] = bcadd($numberOfBeds[$hbp->id][$monthNum], $bpData[$numberOfBedsIndicatorId] ?? '0');
                        $casesOfTreatment[$hbp->id][$monthNum] = bcadd($casesOfTreatment[$hbp->id][$monthNum], $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                        $patientDays[$hbp->id][$monthNum] = bcadd($patientDays[$hbp->id][$monthNum], $bpData[$patientDaysIndicatorId] ?? '0');
                        $cost[$hbp->id][$monthNum] = bcadd($cost[$hbp->id][$monthNum], $bpData[$costIndicatorId] ?? '0');
                }
            }
        }

        foreach ($hospitalBedProfiles as $hbp) {

            $cpNumberOfBeds = $numberOfBeds[$hbp->id];
            $cpCasesOfTreatment = $casesOfTreatment[$hbp->id];
            $cpPatientDays = $patientDays[$hbp->id];
            $cpCost = $cost[$hbp->id];

            $numberOfBedsHasValue = false;
            $casesOfTreatmentHasValue = false;
            $patientDaysHasValue = false;
            $costHasValue = false;
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $numberOfBedsHasValue = $numberOfBedsHasValue || bccomp($cpNumberOfBeds[$monthNum], '0') !== 0;
                $casesOfTreatmentHasValue = $casesOfTreatmentHasValue || bccomp($cpCasesOfTreatment[$monthNum], '0') !== 0;
                $patientDaysHasValue = $patientDaysHasValue || bccomp($cpPatientDays[$monthNum], '0') !== 0;
                $costHasValue = $costHasValue || bccomp($cpCost[$monthNum], '0') !== 0;
            }

            if ($numberOfBedsHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $hbp->name,
                    $hbp->code,
                    $planningParamNames[$numberOfBedsIndicatorId],
                    $cpNumberOfBeds
                );
            }
            if ($casesOfTreatmentHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $hbp->name,
                    $hbp->code,
                    $planningParamNames[$casesOfTreatmentIndicatorId],
                    $cpCasesOfTreatment
                );
            }
            if ($patientDaysHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $hbp->name,
                    $hbp->code,
                    $planningParamNames[$patientDaysIndicatorId],
                    $cpPatientDays
                );
            }
            if ($costHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $hbp->name,
                    $hbp->code,
                    $planningParamNames[$costIndicatorId],
                    $cpCost
                );
            }
        }

        // Круглосуточный стационар (не включая ВМП)
        $planningParamNames[3] = "объемы, койко-дней";

        $numberOfBeds = [];
        $casesOfTreatment = [];
        $patientDays = [];
        $cost = [];

        foreach ($hospitalBedProfiles as $hbp) {
            $numberOfBeds[$hbp->id] = [];
            $casesOfTreatment[$hbp->id] = [];
            $patientDays[$hbp->id] = [];
            $cost[$hbp->id] = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $numberOfBeds[$hbp->id][$monthNum] = '0';
                $casesOfTreatment[$hbp->id][$monthNum] = '0';
                $patientDays[$hbp->id][$monthNum] = '0';
                $cost[$hbp->id][$monthNum] = '0';
            }
        }

        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $regularBedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? null;

            if ($regularBedProfiles) {
                $planningSectionName = 'Круглосуточный стационар (не включая ВМП)';
                $numberOfBedsIndicatorId = 1; // число коек
                $casesOfTreatmentIndicatorId = 7; // госпитализаций
                $patientDaysIndicatorId = 3; // койко-дней
                $costIndicatorId = 4; // стоимость

                foreach ($hospitalBedProfiles as $hbp) {
                        $bpData = $regularBedProfiles[$hbp->id] ?? null;
                        if (!$bpData) { continue; }

                        $numberOfBeds[$hbp->id][$monthNum] = bcadd($numberOfBeds[$hbp->id][$monthNum], $bpData[$numberOfBedsIndicatorId] ?? '0');
                        $casesOfTreatment[$hbp->id][$monthNum] = bcadd($casesOfTreatment[$hbp->id][$monthNum], $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                        $patientDays[$hbp->id][$monthNum] = bcadd($patientDays[$hbp->id][$monthNum], $bpData[$patientDaysIndicatorId] ?? '0');
                        $cost[$hbp->id][$monthNum] = bcadd($cost[$hbp->id][$monthNum], $bpData[$costIndicatorId] ?? '0');
                }
            }
        }

        foreach ($hospitalBedProfiles as $hbp) {

            $cpNumberOfBeds = $numberOfBeds[$hbp->id];
            $cpCasesOfTreatment = $casesOfTreatment[$hbp->id];
            $cpPatientDays = $patientDays[$hbp->id];
            $cpCost = $cost[$hbp->id];

            $numberOfBedsHasValue = false;
            $casesOfTreatmentHasValue = false;
            $patientDaysHasValue = false;
            $costHasValue = false;
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $numberOfBedsHasValue = $numberOfBedsHasValue || bccomp($cpNumberOfBeds[$monthNum], '0') !== 0;
                $casesOfTreatmentHasValue = $casesOfTreatmentHasValue || bccomp($cpCasesOfTreatment[$monthNum], '0') !== 0;
                $patientDaysHasValue = $patientDaysHasValue || bccomp($cpPatientDays[$monthNum], '0') !== 0;
                $costHasValue = $costHasValue || bccomp($cpCost[$monthNum], '0') !== 0;
            }

            if ($numberOfBedsHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $hbp->name,
                    $hbp->code,
                    $planningParamNames[$numberOfBedsIndicatorId],
                    $cpNumberOfBeds
                );
            }
            if ($casesOfTreatmentHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $hbp->name,
                    $hbp->code,
                    $planningParamNames[$casesOfTreatmentIndicatorId],
                    $cpCasesOfTreatment
                );
            }
            if ($patientDaysHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $hbp->name,
                    $hbp->code,
                    $planningParamNames[$patientDaysIndicatorId],
                    $cpPatientDays
                );
            }
            if ($costHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $hbp->name,
                    $hbp->code,
                    $planningParamNames[$costIndicatorId],
                    $cpCost
                );
            }
        }

        // ВМП
        $planningParamNames[3] = "объемы, койко-дней";
        $planningSectionName = 'ВМП';
        $numberOfBedsIndicatorId = 1; // число коек
        $casesOfTreatmentIndicatorId = 7; // госпитализаций
        $patientDaysIndicatorId = 3; // койко-дней
        $costIndicatorId = 4; // стоимость

        $numberOfBeds = [];
        $casesOfTreatment = [];
        $patientDays = [];
        $cost = [];

        foreach ($hospitalBedProfiles as $hbp) {
            $numberOfBeds[$hbp->id] = [];
            $casesOfTreatment[$hbp->id] = [];
            $patientDays[$hbp->id] = [];
            $cost[$hbp->id] = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $numberOfBeds[$hbp->id][$monthNum] = '0';
                $casesOfTreatment[$hbp->id][$monthNum] = '0';
                $patientDays[$hbp->id][$monthNum] = '0';
                $cost[$hbp->id][$monthNum] = '0';
            }
        }

        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $careProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? null;

            if ($careProfiles) {
                foreach ($careProfiles as $cpId => $vmpGroupsData)
                {
                    if (!$vmpGroupsData) { continue; }

                    foreach ($vmpGroupsData as $vmpGroup => $vmpTypes) {
                        $hbpId = vmpGetBedProfileId($cpId, $mo->code, $vmpGroup);

                        foreach ($vmpTypes as $vmpT)
                        {
                            $numberOfBeds[$hbpId][$monthNum] = bcadd($numberOfBeds[$hbpId][$monthNum], $vmpT[$numberOfBedsIndicatorId] ?? '0');
                            $casesOfTreatment[$hbpId][$monthNum] = bcadd($casesOfTreatment[$hbpId][$monthNum], $vmpT[$casesOfTreatmentIndicatorId] ?? '0');
                            $patientDays[$hbpId][$monthNum] = bcadd($patientDays[$hbpId][$monthNum], $vmpT[$patientDaysIndicatorId] ?? '0');
                            $cost[$hbpId][$monthNum] = bcadd($cost[$hbpId][$monthNum], $vmpT[$costIndicatorId] ?? '0');
                        }
                    }
                }
            }
        }

        foreach ($hospitalBedProfiles as $hbp) {
            $cpNumberOfBeds = $numberOfBeds[$hbp->id];
            $cpCasesOfTreatment = $casesOfTreatment[$hbp->id];
            $cpPatientDays = $patientDays[$hbp->id];
            $cpCost = $cost[$hbp->id];

            $numberOfBedsHasValue = false;
            $casesOfTreatmentHasValue = false;
            $patientDaysHasValue = false;
            $costHasValue = false;
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $numberOfBedsHasValue = $numberOfBedsHasValue || bccomp($cpNumberOfBeds[$monthNum], '0') !== 0;
                $casesOfTreatmentHasValue = $casesOfTreatmentHasValue || bccomp($cpCasesOfTreatment[$monthNum], '0') !== 0;
                $patientDaysHasValue = $patientDaysHasValue || bccomp($cpPatientDays[$monthNum], '0') !== 0;
                $costHasValue = $costHasValue || bccomp($cpCost[$monthNum], '0') !== 0;
            }

            if ($numberOfBedsHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $hbp->name,
                    $hbp->code,
                    $planningParamNames[$numberOfBedsIndicatorId],
                    $cpNumberOfBeds
                );
            }
            if ($casesOfTreatmentHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $hbp->name,
                    $hbp->code,
                    $planningParamNames[$casesOfTreatmentIndicatorId],
                    $cpCasesOfTreatment
                );
            }
            if ($patientDaysHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $hbp->name,
                    $hbp->code,
                    $planningParamNames[$patientDaysIndicatorId],
                    $cpPatientDays
                );
            }
            if ($costHasValue) {
                vitacoreHospitalByProfilePeriodsPrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset++,
                    $ordinalRowNum++,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $hbp->name,
                    $hbp->code,
                    $planningParamNames[$costIndicatorId],
                    $cpCost
                );
            }
        }
    } // foreach MO

    $writer = new Xlsx($spreadsheet);
    $writer->save($fullResultFilepath);
    return Storage::download($resultFilePath);
});

function miacHospitalByProfilePeriodsPrintRow(
    Worksheet $sheet,
    int $colIndex,
    int $rowIndex,
    string | int $moCode,
    string | int $paymentType,
    string | int $hospitalType,
    string | int $bedProfileV020,
    string | int $nMonth,
    string | int $level,
    string | int $value
) {
    $moCodeColOffset = 0;
    $paymentTypeColOffset = 1;
    $hospitalTypeColOffset = 2;
    $bedProfileV020ColOffset = 3;
    $nMonthColOffset = 4;
    $levelColOffset = 5;
    $valueColOffset = 6;

    $sheet->setCellValue([$colIndex + $moCodeColOffset, $rowIndex], $moCode);
    $sheet->setCellValue([$colIndex + $paymentTypeColOffset, $rowIndex], $paymentType);
    $sheet->setCellValue([$colIndex + $hospitalTypeColOffset, $rowIndex], $hospitalType);
    $sheet->setCellValue([$colIndex + $bedProfileV020ColOffset, $rowIndex], $bedProfileV020);
    $sheet->setCellValue([$colIndex + $nMonthColOffset, $rowIndex], $nMonth);
    $sheet->setCellValue([$colIndex + $levelColOffset, $rowIndex], $level);
    $sheet->setCellValue([$colIndex + $valueColOffset, $rowIndex], $value);
}

function miacHospitalByProfilePeriodsPrintTableHeader(
    Worksheet $sheet,
    int $colIndex,
    int $rowIndex,
) {
    miacHospitalByProfilePeriodsPrintRow($sheet, $colIndex, $rowIndex, "MCOD", "OPLAT", "TS", "PROFILE", "N_MONTH", "LEVEL", "PLAN_CASES");
}

function getOplat(string $daytimeOrRoundClock, string $hospitalSubType, int $bedProfileId): int
{
    $isVmp = ($hospitalSubType == 'vmp');
    if ($isVmp) {
        return 6; // ВМП
    }
    if (RehabilitationProfileService::IsRehabilitationBedProfile($bedProfileId)) {
        return 7; // реабилитация
    }
    return 1; // специализированная
}
function getTs(string $daytimeOrRoundClock, string $hospitalSubType): int
{
    // 3 - стационар на дому (у нас нет)
    if ($daytimeOrRoundClock === 'roundClock') {
        return 1; // Круглосуточный стационар
    }
    if ($daytimeOrRoundClock === 'daytime') {
        if ($hospitalSubType === 'inPolyclinic') {
            return 2; // дневной при поликлинике
        } elseif ($hospitalSubType = 'inHospital') {
            return 4; // дневной при стационаре
        }
    }
    return -1;
}

function getLevel(int $monthNum, int $ts, int $moId, int $bedProfileId) : string
{
    $l1 = '1';
    $l2_1 = '2.1';
    $l2_2 = '2.2';
    $l3_1 = '3.1';
    $l3_2 = '3.2';
    $l3_3 = '3.3';

    $lResult = "ERROR";

    if ($ts === 2 || $ts === 4) {
        $lResult = $l1;
    } elseif ($ts === 1) {
        switch ($moId) {
            case 17	/* ГБУ "Далматовская ЦРБ" */:
            case 20	/* ГБУ "Катайская ЦРБ" */:
            case 25	/* ГБУ "Шадринская ЦРБ" */:
            case 45	/* ООО "ЛДК "Центр ДНК" */:
                $lResult = $l1;
                break;

            case 89	/* ГБУ «Межрайонная больница №1» */:
            case 90	/* ГБУ «Межрайонная больница №2» */:
            case 91	/* ГБУ «Межрайонная больница №3» */:
            case 92	/* ГБУ «Межрайонная больница №4» */:
            case 93	/* ГБУ «Межрайонная больница №5» */:
            case 94	/* ГБУ «Межрайонная больница №6» */:
            case 95	/* ГБУ «Межрайонная больница №7» */:
            case 96	/* ГБУ «Межрайонная больница №8» */:
                $lResult = $l2_1;
                break;

            case 7	/* ГБУ "Курганская областная специализированная инфекционная больница" */:
            case 38	/* ЧУЗ "РЖД-Медицина" г. Курган" */:
            case 51	/* ГБУ "Санаторий "Озеро Горькое" */:
            case 13	/* ГБУ «КОКВД» */:
                $lResult = $l2_2;
                break;

            case 97	/* ГБУ «Курганская областная больница №2» */ :
            case 11	/* ГБУ "Курганский областной кардиологический диспансер" */:
            case 1	/* ГБУ "КОКБ" */:
            case 98	/* ГБУ "ШГБ" */:
            case 3	/* ГБУ "Курганская БСМП" */:
            case 40	/* ГБУ "Перинатальный центр" */:
            case 42	/* ФГБУ «НМИЦ ТО имени академика Г.А.Илизарова» Минздрава России */:
            case 12	/* ГБУ «КОДКБ им. Красного Креста» */:
                $lResult = $l3_1;
                break;

            case 2	/* ГБУ "КООД" */:
            case 67	/* ГБУ "КОГВВ" */:
                $lResult = $l3_2;
                break;
        }

        if ($monthNum > 1) {
            switch ($moId) {
                case 13	/* ГБУ «КОКВД» */:
                case 38	/* ЧУЗ "РЖД-Медицина" г. Курган" */:
                case 7	/* ГБУ "Курганская областная специализированная инфекционная больница" */:
                    $lResult = $l2_1;
                    break;

                case 89	/* ГБУ «Межрайонная больница №1» */:
                case 90	/* ГБУ «Межрайонная больница №2» */:
                    $lResult = $l2_2;
                    break;

                case 1	/* ГБУ "КОКБ" */:
                case 11	/* ГБУ "Курганский областной кардиологический диспансер" */:
                case 12	/* ГБУ «КОДКБ им. Красного Креста» */:
                case 40	/* ГБУ "Перинатальный центр" */:
                case 42	/* ФГБУ «НМИЦ ТО имени академика Г.А.Илизарова» Минздрава России */:
                    $lResult = $l3_2;
                    break;

                case 2	/* ГБУ "КООД" */:
                case 67	/* ГБУ "КОГВВ" */:
                    $lResult = $l3_3;
                    break;
            }
        }

        if ($monthNum > 3) {
            switch ($moId) {
                case 3	/* ГБУ "Курганская БСМП" */:
                    $lResult = $l3_2;
                    break;
            }
        }

        if ($monthNum > 8) {
            switch ($moId) {
                case 40	/* ГБУ "Перинатальный центр" */:
                    $lResult = $l3_3;
                    break;
            }
        }
    }
    return $lResult;
}

function vmpGetBedProfileId(int $careProfileId, string $moCode, int $vmpGroup)
{
    if ($moCode === '450004' && ($vmpGroup === 23 || $vmpGroup === 24)) {
        return 36; // Радиологические (V020 код: 64);
    }
    // пропускаем профили:
    //  19 - Кардиохирургические
    //  29 - Для беременных и рожениц
    //  30 - Патологии беременности
    //
    //  Нет в ВМП
    //  34 - Реабилитационные для больных с заболеваниями центральной нервной системы и органов чувств
    //  35 - Реабилитационные для больных с заболеваниями опорно-двигательного аппарата и периферической нервной системы
    $bedProfileNotUsedForVmp = [19, 29, 30, 34, 35];

    $cp = CareProfiles::find($careProfileId);

    $cpFomsCollection = $cp->careProfilesFoms;
    $bpArr = [];
    foreach ($cpFomsCollection as $cpFoms) {
        $addBp = $cpFoms->hospitalBedProfiles->pluck('id')->toArray();
        $bpArr = array_merge($bpArr, $addBp);
    }
    $bpArr = array_unique($bpArr, SORT_NUMERIC);
    $bpArr = array_diff($bpArr, $bedProfileNotUsedForVmp);
    if(count($bpArr) !== 1) {
        throw("Error. Count " + count($bpArr));
    }
    return array_pop($bpArr);
}

Route::get('/miac-hospital-by-bed-profile-periods/{year}/{commissionDecisionsId?}', function (DataForContractService $dataForContractService, int $year, int $commissionDecisionsId = null) {
    $packageIds = null;
    $currentlyUsedDate = $year.'-01-01';
    if ($commissionDecisionsId) {
        $commissionDecisions = CommissionDecision::whereYear('date',$year)->where('id', '<=', $commissionDecisionsId)->get();
        $cd = $commissionDecisions->find($commissionDecisionsId);
        $commissionDecisionIds = $commissionDecisions->pluck('id')->toArray();
        $packageIds = ChangePackage::whereIn('commission_decision_id', $commissionDecisionIds)->orWhere('commission_decision_id', null)->pluck('id')->toArray();

        $currentlyUsedDate = $cd->date->format('Y-m-d');
    } else {
        $packageIds = ChangePackage::where('commission_decision_id', null)->pluck('id')->toArray();
    }
    $path = 'xlsx';
    $resultFileName = 'hospital-periods.csv';
    $strDateTimeNow = date("Y-m-d-His");
    $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . ' ' . $resultFileName;
    $fullResultFilepath = Storage::path($resultFilePath);

    bcscale(4);

    $moCollection = MedicalInstitution::WhereRaw("? BETWEEN effective_from AND effective_to", [$currentlyUsedDate])->orderBy('order')->get();
    $hospitalBedProfiles = HospitalBedProfiles::all();
    $oplatOption = [1,6,7];
    $tsOption = [1,2,3,4];

    $values = [];
    foreach ($moCollection as $mo) {
        $values[$mo->id] = [];
        foreach ($oplatOption as $oplat) {
            $values[$mo->id][$oplat] = [];
            foreach ($tsOption as $ts) {
                $values[$mo->id][$oplat][$ts] = [];
                foreach ($hospitalBedProfiles as $hbp) {
                    $values[$mo->id][$oplat][$ts][$hbp->id] = [];
                    for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                        $values[$mo->id][$oplat][$ts][$hbp->id][$monthNum] = '0';
                    }
                }
            }
        }
    }

    $indicatorIds = [2, 7];
    $contentByMonth = [];
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $contentByMonth[$monthNum] = $dataForContractService->GetArrayByYearAndMonth($year, $monthNum, $packageIds, $indicatorIds);
    }

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    $ordinalRowNum = 1;
    $firstTableDataRowIndex = 2;
    $firstTableColIndex = 1;
    $rowOffset = 0;
    $firstTableHeadRowIndex = 1;

    $category = 'hospital';

    miacHospitalByProfilePeriodsPrintTableHeader($sheet, $firstTableColIndex, $firstTableHeadRowIndex);

    foreach($moCollection as $mo) {
        $planningParamNames = [
            2 => "объемы, случаев лечения",
            7 => "объемы, госпитализаций"
        ];

        // Дневные стационары при стационаре
        $daytimeOrRoundClock = 'daytime';
        $hospitalSubType = 'inHospital';
        $casesOfTreatmentIndicatorId = 2; // случаев лечения
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $inHospitalBedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category][$daytimeOrRoundClock][$hospitalSubType]['bedProfiles'] ?? null;

            if ($inHospitalBedProfiles) {
                foreach ($hospitalBedProfiles as $hbp) {
                    $bpData = $inHospitalBedProfiles[$hbp->id] ?? null;
                    if (!$bpData) { continue; }

                    $oplat = getOplat($daytimeOrRoundClock, $hospitalSubType, $hbp->id);
                    $ts = getTs($daytimeOrRoundClock, $hospitalSubType);

                    $values[$mo->id][$oplat][$ts][$hbp->id][$monthNum] = bcadd($values[$mo->id][$oplat][$ts][$hbp->id][$monthNum], $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                }
            }
        }

        // Дневные стационары при поликлинике
        $daytimeOrRoundClock = 'daytime';
        $hospitalSubType = 'inPolyclinic';
        $casesOfTreatmentIndicatorId = 2; // случаев лечения

        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $inPolyclinicBedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category][$daytimeOrRoundClock][$hospitalSubType]['bedProfiles'] ?? null;

            if ($inPolyclinicBedProfiles) {
                foreach ($hospitalBedProfiles as $hbp) {
                    $bpData = $inPolyclinicBedProfiles[$hbp->id] ?? null;
                    if (!$bpData) { continue; }

                    $oplat = getOplat($daytimeOrRoundClock, $hospitalSubType, $hbp->id);
                    $ts = getTs($daytimeOrRoundClock, $hospitalSubType);

                    $values[$mo->id][$oplat][$ts][$hbp->id][$monthNum] = bcadd($values[$mo->id][$oplat][$ts][$hbp->id][$monthNum], $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                }
            }
        }

        // Круглосуточный стационар (не включая ВМП)
        $daytimeOrRoundClock = 'roundClock';
        $hospitalSubType = 'regular';
        $casesOfTreatmentIndicatorId = 7; // госпитализаций

        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $regularBedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category][$daytimeOrRoundClock][$hospitalSubType]['bedProfiles'] ?? null;

            if ($regularBedProfiles) {
                foreach ($hospitalBedProfiles as $hbp) {
                    $bpData = $regularBedProfiles[$hbp->id] ?? null;
                    if (!$bpData) { continue; }

                    $oplat = getOplat($daytimeOrRoundClock, $hospitalSubType, $hbp->id);
                    $ts = getTs($daytimeOrRoundClock, $hospitalSubType);

                    $values[$mo->id][$oplat][$ts][$hbp->id][$monthNum] = bcadd($values[$mo->id][$oplat][$ts][$hbp->id][$monthNum], $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                }
            }
        }

        // ВМП
        $daytimeOrRoundClock = 'roundClock';
        $hospitalSubType = 'vmp';
        $casesOfTreatmentIndicatorId = 7; // госпитализаций

        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $careProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category][$daytimeOrRoundClock][$hospitalSubType]['careProfiles'] ?? null;

            if ($careProfiles) {
                foreach ($careProfiles as $cpId => $vmpGroupsData)
                {
                    if (!$vmpGroupsData) { continue; }

                    $ts = getTs($daytimeOrRoundClock, $hospitalSubType);
                    foreach ($vmpGroupsData as $vmpGroup => $vmpTypes) {
                        $hbpId = vmpGetBedProfileId($cpId, $mo->code, $vmpGroup);
                        $oplat = getOplat($daytimeOrRoundClock, $hospitalSubType, $hbpId);

                        foreach ($vmpTypes as $vmpT)
                        {
                            $values[$mo->id][$oplat][$ts][$hbpId][$monthNum]
                                = bcadd($values[$mo->id][$oplat][$ts][$hbpId][$monthNum], $vmpT[$casesOfTreatmentIndicatorId] ?? '0');
                        }
                    }
                }
            }
        }
    } // foreach MO

    $row = $firstTableDataRowIndex;
    foreach ($moCollection as $mo) {
        foreach ($oplatOption as $oplat) {
            foreach ($tsOption as $ts) {
                foreach ($hospitalBedProfiles as $hbp) {
                    // Если объемы есть хотя бы на один месяц,
                    // выводим информацию по всем месяцам (включая нулевые значения)
                    $hasValue = false;
                    for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                        if(bccomp($values[$mo->id][$oplat][$ts][$hbp->id][$monthNum], '0') != 0) {
                            $hasValue = true;
                            break;
                        }
                    }
                    if ($hasValue) {
                    for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                            miacHospitalByProfilePeriodsPrintRow(
                                $sheet, $firstTableColIndex, $row++, $mo->code,
                                $oplat, $ts, $hbp->code, $monthNum, getLevel($monthNum, $ts, $mo->id, $hbp->id),
                                $values[$mo->id][$oplat][$ts][$hbp->id][$monthNum]
                            );
                        }
                    }
                }
            }
        }
    }

    $writer = new Csv($spreadsheet);
    $writer->setDelimiter(';');
    $writer->setUseBOM(true);
    $writer->save($fullResultFilepath);
    return Storage::download($resultFilePath);
});

Route::get('/{year}/{commissionDecisionsId?}', function (Request $request, DataForContractService $dataForContractService, MoInfoForContractService $moInfoForContractService, MoDepartmentsInfoForContractService $moDepartmentsInfoForContractService, PeopleAssignedInfoForContractService $peopleAssignedInfoForContractService, int $year, int $commissionDecisionsId = null) {
    $onlyMoModifiedByCommission = boolval($request->exists("onlyModified"));

    $packageIds = null;
    $protocolNumber = 0;
    $protocolDate = '';
    if ($commissionDecisionsId) {
        $commissionDecisions = CommissionDecision::whereYear('date',$year)->where('id', '<=', $commissionDecisionsId)->get();
        $cd = $commissionDecisions->find($commissionDecisionsId);
        $commissionDecisionIds = $commissionDecisions->pluck('id')->toArray();
        $protocolNumber = $cd->number;
        $protocolDate = $cd->date->format('d.m.Y');
        $docName = "к протоколу заседания комиссии по разработке территориальной программы ОМС Курганской области от $protocolDate";
        $packageIds = ChangePackage::whereIn('commission_decision_id', $commissionDecisionIds)->orWhere('commission_decision_id', null)->pluck('id')->toArray();
    } else {
        $packageIds = ChangePackage::where('commission_decision_id', null)->pluck('id')->toArray();
    }

    $moIds = null;
    if ($commissionDecisionsId){
        if ($onlyMoModifiedByCommission) {
            $c = CommissionDecision::find($commissionDecisionsId);
            $pIds = $c->changePackage()->pluck('id')->toArray();
            $moIds = PlannedIndicatorChange::select('mo_id')->where('package_id', $pIds)->groupBy('mo_id')->get()->pluck('mo_id')->toArray();
        }
    }

    $strDateTimeNow = date("Y-m-d-His");
    $path =  $year.DIRECTORY_SEPARATOR.$protocolNumber.'('.$protocolDate.')'.DIRECTORY_SEPARATOR.$strDateTimeNow.DIRECTORY_SEPARATOR;

    $indicators = Indicator::select('id','name')->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put($path.'indicators.json', $indicators);

    $medicalServices = MedicalServices::select('id','name')->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put($path.'medicalServices.json', $medicalServices);

    $medicalAssistanceType = MedicalAssistanceType::select('id','name')->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put($path.'medicalAssistanceType.json', $medicalAssistanceType);

    $hospitalBedProfiles = HospitalBedProfiles::select('tbl_hospital_bed_profiles.id','name', 'care_profile_foms_id')
    ->join('tbl_hospital_bed_profile_care_profile_foms',  'tbl_hospital_bed_profiles.id', '=', 'hospital_bed_profile_id')
    ->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put($path.'hospitalBedProfiles.json', $hospitalBedProfiles);

    $vmpGroup = VmpGroup::select('id','code')->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put($path.'vmpGroup.json', $vmpGroup);

    $vmpTypes = VmpTypes::select('id','name')->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put($path.'vmpTypes.json', $vmpTypes);

    $careProfiles = CareProfiles::select('tbl_care_profiles.id','name','care_profile_foms_id')
    ->join('tbl_care_profile_care_profile_foms',  'tbl_care_profiles.id', '=', 'care_profile_id')
    ->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put($path.'careProfiles.json', $careProfiles);

    $careProfilesFoms = CareProfilesFoms::select('id', 'name', 'code_v002 as code')
    ->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put($path.'careProfilesFoms.json', $careProfilesFoms);

    $mo = $moInfoForContractService->GetJson();
    Storage::put($path.'mo.json', $mo);

    $moDepartments = $moDepartmentsInfoForContractService->GetJson();
    Storage::put($path.'moDepartments.json', $moDepartments);

    $content = $dataForContractService->GetJson($year, $packageIds, $moIds);
    Storage::put($path.'data.json', $content);

    $content = $peopleAssignedInfoForContractService->GetJson($year, $packageIds);
    Storage::put($path.'peopleAssignedData.json', $content);

    $files = Storage::files($path);

    $zip = new ZipArchive();
    $zipFileName = $path.$strDateTimeNow.'_Protokol_'.$protocolNumber.'('.$protocolDate.').dogoms';
    $fullZipFileName = Storage::path($zipFileName);

    if ($zip->open($fullZipFileName, ZipArchive::CREATE) !== TRUE) {
        throw new \Exception('Cannot create a zip file');
    }

    foreach($files as $filepath){
        $fullResultFilepath = Storage::path($filepath);
        $zip->addFile($fullResultFilepath, "Json".DIRECTORY_SEPARATOR.basename($fullResultFilepath));
    }
    $zip->close();

    return Storage::download($zipFileName);;
});

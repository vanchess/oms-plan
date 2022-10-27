<?php

use App\Http\Controllers\CategoryTreeController;
use App\Jobs\InitialChanges;
use App\Jobs\InitialDataLoaded;
use App\Models\CareProfiles;
use App\Models\CareProfilesFoms;
use App\Models\ChangePackage;
use App\Models\CommissionDecision;
use App\Models\HospitalBedProfiles;
use App\Models\Indicator;
use App\Models\MedicalAssistanceType;
use Illuminate\Support\Facades\Route;
use App\Models\MedicalInstitution;
use App\Models\MedicalServices;
use App\Models\Organization;
use App\Models\Period;

use App\Services\InitialDataService;
use App\Models\PlannedIndicator;
use App\Models\PlannedIndicatorChange;
use App\Models\VmpGroup;
use App\Models\VmpTypes;
use App\Services\DataForContractService;
//use App\Services\NodeService;
use App\Services\Dto\InitialDataValueDto;
use App\Services\MoDepartmentsInfoForContractService;
use App\Services\MoInfoForContractService;
use App\Services\PeopleAssignedInfoForContractService;
use Illuminate\Support\Facades\Auth;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\Storage;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
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

class MedicalServicesEnum {
    /// <summary>
        /**
         * Эндоскопические исследования
         */
        public const Endoscopy = 1;
        /**
         * КТ
         */
        public const KT = 2;
        /**
         * МРТ
         */
        public const MRT = 3;
        /**
         * Ультразвуковое исследование сердечно-сосудистой системы
         */
        public const UltrasoundCardio = 4;
        /**
         * Патолого-анатомическое исследование биопсийного материала
         */
        public const PathologicalAnatomicalBiopsyMaterial = 5;
        /**
         * Малекулярно-генетические исследования с целью выявления онкологических заболеваний
         */
        public const MolecularGeneticDetectionOncological = 6;
        /**
         * Тестирование на КОВИД
         */
        public const CovidTesting = 7;
        /**
         * ПЭТ
         */
        public const PET = 8;
        /**
         * Определение антигена D системы Резус (резус-фактор)
         */
        public const DeterminationAntigenD = 9;
        /**
         * Дистанционное наблюдение за показателями артериального давления
         */
        public const RemoteMonitoringBloodPressureIndicators = 10;
        /**
         * Комплексное исследование для диагностики фоновых и предраковых заболевание репродуктивных органов у женщин
         */
        public const DiagnosisBackgroundPrecancerousDiseasesReproductiveWomen = 11;
        /**
         * УЗИ плода
         */
        public const FetalUltrasound = 12;
}

Route::get('/321123', function (InitialDataService $initialDataService) {

    $nodeId = 4;
    $userId = 1;

    $medicalInstitutions = MedicalInstitution::OrderBy('order')->get();


    foreach ($medicalInstitutions as $mo) {
        foreach (PlannedIndicator::all() as $pi) {

            $dto = new InitialDataValueDto(
                    year: 2021,
                    moId: $mo->id,
                    plannedIndicatorId: $pi->id,
                    value: 2021 + $mo->id + $pi->id,
                    userId: $userId
                );

            $initialDataService->setValue($dto);
        }
    }

    return 'OK';
    /*
    $medicalInstitutions = MedicalInstitution::OrderBy('order')->get();
    foreach ($medicalInstitutions as $mo) {
        $org = Organization::Where('inn',$mo->inn)->first();
        $mo->organization_id = $org->id;
        $mo->save();
    }
    */
    return 'OK';

    //phpinfo();
});

Route::get('/initial_changes', function () {
    InitialChanges::dispatch(2022);
    return "initial_changes";
});

Route::get('/all_initial_data_loaded', function () {
    InitialDataLoaded::dispatch(1, 2022, 1);
    InitialDataLoaded::dispatch(9, 2022, 1);
    InitialDataLoaded::dispatch(17, 2022, 1);
    return "all_initial_data_loaded";
});

function medicalServicesSum($content, $moId, $serviceId, $indicatorId, $category = 'polyclinic') {
    bcscale(4);

    $perPerson = $content['mo'][$moId][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? '0';
    $perUnit = $content['mo'][$moId][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? '0';
    $faps = $content['mo'][$moId][$category]['fap'] ?? [];
    $fap = '0';
    foreach ($faps as $f) {
        $fap = bcadd($fap, $f['services'][$serviceId][$indicatorId] ?? '0');
    }
    return bcadd($perPerson, bcadd($perUnit, $fap));
}

Route::get('/meeting-minutes/{year}/{commissionDecisionsId}', function (DataForContractService $dataForContractService, int $year, int $commissionDecisionsId) {
    // $year = 2022;
    // $commissionDecisionsId = 19;
    $cd = CommissionDecision::find($commissionDecisionsId);
    $protocolDate = $cd->date->format('d.m.Y');
    $docName = "протокол заседания КРТП ОМС №$cd->number от $protocolDate";
    $packageIds = $cd->changePackage()->pluck('id')->toArray();
    $indicatorIds = [1, 2, 3, 4, 5, 6, 7, 8, 9];
    $content = $dataForContractService->GetArray($year, $packageIds, $indicatorIds);
    $moCollection = MedicalInstitution::orderBy('order')->get();

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
        $kt = medicalServicesSum($content, $mo->id, $serviceId, $indicatorId);
        $costKt = medicalServicesSum($content, $mo->id, $serviceId, $costIndicatorId);

        // МРТ
        $serviceId = MedicalServicesEnum::MRT;
        $mrt = medicalServicesSum($content, $mo->id, $serviceId, $indicatorId);
        $costMrt = medicalServicesSum($content, $mo->id, $serviceId, $costIndicatorId);

        // УЗИ ССС
        $serviceId = MedicalServicesEnum::UltrasoundCardio;
        $ultrasoundCardio = medicalServicesSum($content, $mo->id, $serviceId, $indicatorId);
        $costUltrasoundCardio = medicalServicesSum($content, $mo->id, $serviceId, $costIndicatorId);

        // Эндоскопия
        $serviceId = MedicalServicesEnum::Endoscopy;
        $endoscopy = medicalServicesSum($content, $mo->id, $serviceId, $indicatorId);
        $costEndoscopy = medicalServicesSum($content, $mo->id, $serviceId, $costIndicatorId);

        // ПАИ
        $serviceId = MedicalServicesEnum::PathologicalAnatomicalBiopsyMaterial;
        $pathologicalAnatomicalBiopsyMaterial = medicalServicesSum($content, $mo->id, $serviceId, $indicatorId);
        $costPathologicalAnatomicalBiopsyMaterial = medicalServicesSum($content, $mo->id, $serviceId, $costIndicatorId);

        // МГИ
        $serviceId = MedicalServicesEnum::MolecularGeneticDetectionOncological;
        $molecularGeneticDetectionOncological = medicalServicesSum($content, $mo->id, $serviceId, $indicatorId);
        $costMolecularGeneticDetectionOncological = medicalServicesSum($content, $mo->id, $serviceId, $costIndicatorId);

        // Тест.covid-19
        $serviceId = MedicalServicesEnum::CovidTesting;
        $covidTesting = medicalServicesSum($content, $mo->id, $serviceId, $indicatorId);
        $costCovidTesting = medicalServicesSum($content, $mo->id, $serviceId, $costIndicatorId);

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
    $rowIndex = $startRow - 1;
    $category = 'hospital';

    $numberOfBedsIndicatorId = 1; // число коек
    $casesOfTreatmentIndicatorId = 2; // случаев лечения
    $patientDaysIndicatorId = 3; // пациенто-дней
    $costIndicatorId = 4; // стоимость

    foreach($moCollection as $mo) {

        $inHospitalBedProfiles = $content['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? null;

        if (!$inHospitalBedProfiles) { continue; }

        $ordinalRowNum++;
        $rowIndex++;
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

                $numberOfBeds = bcadd($numberOfBeds, $bpData[$numberOfBedsIndicatorId] ?? '0');
                $casesOfTreatment = bcadd($casesOfTreatment, $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                $patientDays = bcadd($patientDays, $bpData[$patientDaysIndicatorId] ?? '0');
                $cost = bcadd($cost, $bpData[$costIndicatorId] ?? '0');
            }
            if( bccomp($numberOfBeds,'0') === 0
                && bccomp($casesOfTreatment, '0') === 0
                && bccomp($patientDays, '0') === 0
                && bccomp($cost, '0') === 0
            ) {continue;}

            $rowIndex++;
            $sheet->setCellValue([3,$rowIndex], $cpf->name);
            $sheet->setCellValue([4,$rowIndex], $numberOfBeds);
            $sheet->setCellValue([5,$rowIndex], $casesOfTreatment);
            $sheet->setCellValue([6,$rowIndex], $patientDays);
            $sheet->setCellValue([7,$rowIndex], $cost);
        }

    }
    $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);


    $sheet = $spreadsheet->getSheetByName('ДС при поликлинике');
    $sheet->setCellValue([1,3], $docName);
    $ordinalRowNum = 0;
    $rowIndex = $startRow - 1;
    $category = 'hospital';

    $numberOfBedsIndicatorId = 1; // число коек
    $casesOfTreatmentIndicatorId = 2; // случаев лечения
    $patientDaysIndicatorId = 3; // пациенто-дней
    $costIndicatorId = 4; // стоимость

    foreach($moCollection as $mo) {

        $inPolyclinicBedProfiles = $content['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? null;;

        if (!$inPolyclinicBedProfiles) { continue; }

        $ordinalRowNum++;
        $rowIndex++;
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

                $numberOfBeds = bcadd($numberOfBeds, $bpData[$numberOfBedsIndicatorId] ?? '0');
                $casesOfTreatment = bcadd($casesOfTreatment, $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                $patientDays = bcadd($patientDays, $bpData[$patientDaysIndicatorId] ?? '0');
                $cost = bcadd($cost, $bpData[$costIndicatorId] ?? '0');
            }
            if( bccomp($numberOfBeds,'0') === 0
                && bccomp($casesOfTreatment, '0') === 0
                && bccomp($patientDays, '0') === 0
                && bccomp($cost, '0') === 0
            ) {continue;}

            $rowIndex++;
            $sheet->setCellValue([3,$rowIndex], "$cpf->name");
            $sheet->setCellValue([4,$rowIndex], $numberOfBeds);
            $sheet->setCellValue([5,$rowIndex], $casesOfTreatment);
            $sheet->setCellValue([6,$rowIndex], $patientDays);
            $sheet->setCellValue([7,$rowIndex], $cost);
        }

    }
    $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);


    $sheet = $spreadsheet->getSheetByName('КС');
    $sheet->setCellValue([1,3], $docName);
    $ordinalRowNum = 0;
    $rowIndex = $startRow - 1;
    $category = 'hospital';

    $numberOfBedsIndicatorId = 1; // число коек
    $casesOfTreatmentIndicatorId = 7; // госпитализаций
    $patientDaysIndicatorId = 3; // пациенто-дней
    $costIndicatorId = 4; // стоимость

    foreach($moCollection as $mo) {

        $inPolyclinicBedProfiles = $content['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? null;

        if (!$inPolyclinicBedProfiles) { continue; }

        $ordinalRowNum++;
        $rowIndex++;
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

                $numberOfBeds = bcadd($numberOfBeds, $bpData[$numberOfBedsIndicatorId] ?? '0');
                $casesOfTreatment = bcadd($casesOfTreatment, $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                $patientDays = bcadd($patientDays, $bpData[$patientDaysIndicatorId] ?? '0');
                $cost = bcadd($cost, $bpData[$costIndicatorId] ?? '0');
            }
            if( bccomp($numberOfBeds,'0') === 0
                && bccomp($casesOfTreatment, '0') === 0
                && bccomp($patientDays, '0') === 0
                && bccomp($cost, '0') === 0
            ) {continue;}

            $rowIndex++;
            $sheet->setCellValue([3,$rowIndex], "$cpf->name");
            $sheet->setCellValue([4,$rowIndex], $numberOfBeds);
            $sheet->setCellValue([5,$rowIndex], $casesOfTreatment);
            $sheet->setCellValue([6,$rowIndex], $patientDays);
            $sheet->setCellValue([7,$rowIndex], $cost);
        }

    }
    $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);


    $sheet = $spreadsheet->getSheetByName('ВМП');
    $sheet->setCellValue([1,3], $docName);
    $ordinalRowNum = 0;
    $rowIndex = $startRow - 1;
    $category = 'hospital';

    $numberOfBedsIndicatorId = 1; // число коек
    $casesOfTreatmentIndicatorId = 7; // госпитализаций
    $patientDaysIndicatorId = 3; // пациенто-дней
    $costIndicatorId = 4; // стоимость

    foreach($moCollection as $mo) {
        $careProfiles = $content['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? null;

        if (!$careProfiles) { continue; }

        $ordinalRowNum++;
        $rowIndex++;
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
                        $numberOfBeds = bcadd($numberOfBeds, $vmpT[$numberOfBedsIndicatorId] ?? '0');
                        $casesOfTreatment = bcadd($casesOfTreatment, $vmpT[$casesOfTreatmentIndicatorId] ?? '0');
                        $patientDays = bcadd($patientDays, $vmpT[$patientDaysIndicatorId] ?? '0');
                        $cost = bcadd($cost, $vmpT[$costIndicatorId] ?? '0');
                    }
                }
            }

            if( bccomp($numberOfBeds,'0') === 0
                && bccomp($casesOfTreatment, '0') === 0
                && bccomp($patientDays, '0') === 0
                && bccomp($cost, '0') === 0
            ) {continue;}

            $rowIndex++;
            $sheet->setCellValue([3,$rowIndex], "$cpf->name");
            $sheet->setCellValue([4,$rowIndex], $numberOfBeds);
            $sheet->setCellValue([5,$rowIndex], $casesOfTreatment);
            $sheet->setCellValue([6,$rowIndex], $patientDays);
            $sheet->setCellValue([7,$rowIndex], $cost);
        }

    }
    $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);


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

Route::get('/vitacore', function (DataForContractService $dataForContractService, PeopleAssignedInfoForContractService $peopleAssignedInfoForContractService) {
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

    $path = 'xlsx';
    $resultFileName = 'vitacore.xlsx';
    $strDateTimeNow = date("Y-m-d-His");
    $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . ' ' . $resultFileName;
    $fullResultFilepath = Storage::path($resultFilePath);

    bcscale(4);

    $year = 2022;
    $commit = null;
    $content = $dataForContractService->GetArray($year, $commit);
    $contentByMonth = [];
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $contentByMonth[$monthNum] = $dataForContractService->GetArrayByYearAndMonth($year, $monthNum);
    }

    $peopleAssigned = $peopleAssignedInfoForContractService->GetArray($year, $commit);
    $moCollection = MedicalInstitution::orderBy('order')->get();

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    // MO
    $ordinalRowNum = 0;
    $firstTableHeadRowIndex = 2;
    $firstTableDataRowIndex = 6;
    $firstTableColIndex = 1;
    $rowOffset = 0;
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $colOfset = 0;
        $sheet->setCellValue([$firstTableColIndex + $colOfset, $firstTableDataRowIndex + $rowOffset], "$ordinalRowNum");
        $colOfset++;
        $sheet->setCellValue([$firstTableColIndex + $colOfset, $firstTableDataRowIndex + $rowOffset], $mo->code);
        $colOfset++;
        $sheet->setCellValue([$firstTableColIndex + $colOfset, $firstTableDataRowIndex + $rowOffset], $mo->short_name);
        $colOfset++;
        $rowOffset++;
    }
    //$sheet = $spreadsheet->getSheetByName('1.Скорая помощь');
    $category = 'ambulance';
    $indicatorId = 5; // вызовов
    $callsAssistanceTypeId = 5; // вызовы
    $thrombolysisAssistanceTypeId = 6;// тромболизис
    $rowOffset = 0;

    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Скорая помощь, плановые объемы на $year год");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 23, $firstTableHeadRowIndex]);
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], "Численность прикрепленного населения на 01.01.$year");
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, объемы скорой помощи');
    //$sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 1, $firstTableHeadRowIndex + 1]);
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 2], 'вызовов');
    //$sheet->setCellValue([$curColumn + 1, $firstTableHeadRowIndex + 2], 'в том числе вызовов с проведением тромболитической терапии');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, объемы скорой помощи помесячно');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);
    $sheet->setCellValue([$curColumn + 12, $firstTableHeadRowIndex + 1], 'в том числе вызовов с проведением тромболитической терапии');
    $sheet->mergeCells([$curColumn + 12, $firstTableHeadRowIndex + 1, $curColumn + 23, $firstTableHeadRowIndex + 1]);
    /*
    foreach($moCollection as $mo) {
        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        //$peopleAssignedValue = $peopleAssigned[$mo->id][$category]['peopleAssigned'] ?? 0;
        $thrombolysis   = $content['mo'][$mo->id][$category][$thrombolysisAssistanceTypeId][$indicatorId] ?? 0;
        $calls          = $content['mo'][$mo->id][$category][$callsAssistanceTypeId][$indicatorId] ?? 0;

        // Численность прикрепленного населения на 01.01.xxxx
        //$sheet->setCellValue([$curColumn, $curRow], $peopleAssignedValue);
        // Всего, объемы скорой помощи
        // вызовов
        $sheet->setCellValue([$curColumn, $curRow], $calls + $thrombolysis);
        // в том числе вызовов с проведением тромболитической терапии
        $sheet->setCellValue([$curColumn + 1, $curRow], $thrombolysis);
        $rowOffset++;
    }
    $colOfset += 2;
*/
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);

        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $thrombolysis   = $contentByMonth[$monthNum]['mo'][$mo->id][$category][$thrombolysisAssistanceTypeId][$indicatorId] ?? 0;
            $calls          = $contentByMonth[$monthNum]['mo'][$mo->id][$category][$callsAssistanceTypeId][$indicatorId] ?? 0;

            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            // Всего, объемы скорой помощи
            // вызовов
            $sheet->setCellValue([$curColumn, $curRow], $calls + $thrombolysis);
            $rowOffset++;
        }

    }
    $colOfset += 12;

    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);

        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $thrombolysis   = $contentByMonth[$monthNum]['mo'][$mo->id][$category][$thrombolysisAssistanceTypeId][$indicatorId] ?? 0;

            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            // в том числе вызовов с проведением тромболитической терапии
            $sheet->setCellValue([$curColumn, $curRow], $thrombolysis);
            $rowOffset++;
        }

    }
    $colOfset += 12;

    $indicatorId = 4; // стоимость
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], 'Скорая помощь, финансовое обеспечение');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11, $firstTableHeadRowIndex]);
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, руб.');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Скорая помощь, финансовое обеспечение');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);
/*
    $rowOffset = 0;
    foreach($moCollection as $mo) {
        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $calls = $content['mo'][$mo->id][$category][$callsAssistanceTypeId][$indicatorId] ?? '0';
        $thrombolysis = $content['mo'][$mo->id][$category][$thrombolysisAssistanceTypeId][$indicatorId] ?? '0';
        $sheet->setCellValue([$curColumn, $curRow], bcadd($calls, $thrombolysis));
        $rowOffset++;
    }
    $colOfset += 1;
*/
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);

        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $calls = $contentByMonth[$monthNum]['mo'][$mo->id][$category][$callsAssistanceTypeId][$indicatorId] ?? '0';
            $thrombolysis = $contentByMonth[$monthNum]['mo'][$mo->id][$category][$thrombolysisAssistanceTypeId][$indicatorId] ?? '0';
            $sheet->setCellValue([$curColumn, $curRow], bcadd($calls, $thrombolysis));
            $rowOffset++;
        }
    }
    $colOfset += 12;


    // 2.обращения по заболеваниям
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Плановые объемы медицинской помощи в связи с заболеваниями в амбулаторных условиях на $year год");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11, $firstTableHeadRowIndex]);
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], "Численность прикрепленного населения на 01.01.$year");
    //$sheet->setCellValue([$curColumn + 1, $firstTableHeadRowIndex + 1], 'Всего, обращений');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'обращений');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);
    $category = 'polyclinic';
    $indicatorId = 8; // обращений
    $assistanceTypeId = 4; // обращения по заболеваниям
    /*
    $rowOffset = 0;
    foreach($moCollection as $mo) {

        $peopleAssignedPolyclinicValue = $peopleAssigned[$mo->id][$category]['mo']['peopleAssigned'] ?? 0;
        $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
        $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        $fap = 0;
        foreach ($faps as $f) {
            $fap += $f['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
        }

        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $sheet->setCellValue([$curColumn, $curRow], $peopleAssignedPolyclinicValue);
        $sheet->setCellValue([$curColumn + 1, $curRow], $perPerson + $perUnit + $fap );
        $rowOffset++;
    }
    $colOfset += 2;
    */

    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);

        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;

            $perPerson = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
            $perUnit = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
            $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
            $fap = 0;
            foreach ($faps as $f) {
                $fap += $f['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
            }

            $sheet->setCellValue([$curColumn, $curRow], $perPerson + $perUnit + $fap );
            $rowOffset++;
        }
    }
    $colOfset += 12;

    // 3.Посещения с иными целями
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Плановые объемы медицинской помощи в амбулаторных условиях на $year год, посещения с иными целями");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11, $firstTableHeadRowIndex]);
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, посещений');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'посещений');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);
    $category = 'polyclinic';
    $indicatorId = 9; // посещений
    $assistanceTypeIds = [1, 2]; //	посещения профилактические, посещения разовые по заболеваниям
    /*
    $rowOffset = 0;
    foreach($moCollection as $mo) {
        $v = 0;
        foreach($assistanceTypeIds as $assistanceTypeId) {
            $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
            $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
            $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
            $fap = 0;
            foreach ($faps as $f) {
                $fap += $f['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
            }
            $v += ($perPerson + $perUnit + $fap);
        }

        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $sheet->setCellValue([$curColumn,$curRow], $v);
        $rowOffset++;
    }
    $colOfset += 1;
    */
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);

        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;

            $v = 0;
            foreach($assistanceTypeIds as $assistanceTypeId) {
                $perPerson = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
                $perUnit = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
                $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
                $fap = 0;
                foreach ($faps as $f) {
                    $fap += $f['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
                }
                $v += ($perPerson + $perUnit + $fap);
            }

            $sheet->setCellValue([$curColumn,$curRow], $v);
            $rowOffset++;
        }

    }
    $colOfset += 12;

    // 4 Неотложная помощь
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Плановые объемы медицинской помощи в амбулаторных условиях на $year год, неотложная помощь");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11, $firstTableHeadRowIndex]);
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, посещений');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'посещений');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);

    $category = 'polyclinic';
    $indicatorId = 9; // посещений
    $assistanceTypeIds = [3]; //	посещения неотложные
    /*
    $rowOffset = 0;
    foreach($moCollection as $mo) {
        $v = 0;
        foreach($assistanceTypeIds as $assistanceTypeId) {
            $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
            $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
            $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
            $fap = 0;
            foreach ($faps as $f) {
                $fap += $f['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
            }
            $v += ($perPerson + $perUnit + $fap);
        }

        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $sheet->setCellValue([$curColumn,$curRow], $v);
        $rowOffset++;
    }
    $colOfset += 1;
    */
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);

        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $v = 0;
            foreach($assistanceTypeIds as $assistanceTypeId) {
                $perPerson = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
                $perUnit = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
                $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
                $fap = 0;
                foreach ($faps as $f) {
                    $fap += $f['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
                }
                $v += ($perPerson + $perUnit + $fap);
            }

            $sheet->setCellValue([$curColumn,$curRow], $v);
            $rowOffset++;
        }
    }
    $colOfset += 12;


    // 2.2 КТ
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Плановые объемы медицинской помощи в связи с заболеваниями в амбулаторных условиях на  $year год (компьютерная томография)");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11, $firstTableHeadRowIndex]);
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, услуг');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'услуг');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);

    $category = 'polyclinic';
    $indicatorId = 6; // услуг
    $serviceId = MedicalServicesEnum::KT;
    /*
    $rowOffset = 0;
    foreach($moCollection as $mo) {
        $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        $fap = 0;
        foreach ($faps as $f) {
            $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
        }

        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $sheet->setCellValue([$curColumn,$curRow], ($perPerson + $perUnit + $fap));
        $rowOffset++;
    }
    $colOfset += 1;
    */
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);

        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $perPerson = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
            $perUnit = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
            $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
            $fap = 0;
            foreach ($faps as $f) {
                $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
            }
            $sheet->setCellValue([$curColumn,$curRow], ($perPerson + $perUnit + $fap));
            $rowOffset++;
        }
    }
    $colOfset += 12;

    // 2.3 МРТ
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Плановые объемы медицинской помощи в связи с заболеваниями в амбулаторных условиях на  $year год (магнитно-резонансная томография)");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11, $firstTableHeadRowIndex]);
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, услуг');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'услуг');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);

    $category = 'polyclinic';
    $indicatorId = 6; // услуг
    $serviceId = MedicalServicesEnum::MRT;
    /*
    $rowOffset = 0;
    foreach($moCollection as $mo) {
        $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        $fap = 0;
        foreach ($faps as $f) {
            $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
        }

        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $sheet->setCellValue([$curColumn,$curRow], ($perPerson + $perUnit + $fap));
        $rowOffset++;
    }
    $colOfset += 1;
    */
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);

        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $perPerson = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
            $perUnit = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
            $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
            $fap = 0;
            foreach ($faps as $f) {
                $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
            }

            $sheet->setCellValue([$curColumn,$curRow], ($perPerson + $perUnit + $fap));
            $rowOffset++;
        }
    }
    $colOfset += 12;

    // 2.4 УЗИ ССС
    // УЗИ сердечно-сосудистой системы
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Плановые объемы медицинской помощи в связи с заболеваниями в амбулаторных условиях на  $year год (УЗИ сердечно-сосудистой системы)");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11, $firstTableHeadRowIndex]);
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, услуг');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'услуг');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);

    $category = 'polyclinic';
    $indicatorId = 6; // услуг
    $serviceId = MedicalServicesEnum::UltrasoundCardio;
    /*
    $rowOffset = 0;
    foreach($moCollection as $mo) {
        $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        $fap = 0;
        foreach ($faps as $f) {
            $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
        }

        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $sheet->setCellValue([$curColumn,$curRow], ($perPerson + $perUnit + $fap));
        $rowOffset++;
    }
    $colOfset += 1;
    */
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);
        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $perPerson = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
            $perUnit = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
            $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
            $fap = 0;
            foreach ($faps as $f) {
                $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
            }

            $sheet->setCellValue([$curColumn,$curRow], ($perPerson + $perUnit + $fap));
            $rowOffset++;
        }
    }
    $colOfset += 12;


    // 2.5 Эндоскопия
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Плановые объемы медицинской помощи в связи с заболеваниями в амбулаторных условиях на  $year год (Эндоскопические исследования)");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11, $firstTableHeadRowIndex]);
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, услуг');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'услуг');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);

    $category = 'polyclinic';
    $indicatorId = 6; // услуг
    $serviceId = MedicalServicesEnum::Endoscopy;
    /*
    $rowOffset = 0;
    foreach($moCollection as $mo) {
        $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        $fap = 0;
        foreach ($faps as $f) {
            $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
        }

        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $sheet->setCellValue([$curColumn,$curRow], ($perPerson + $perUnit + $fap));
        $rowOffset++;
    }
    $colOfset += 1;
    */
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);
        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $perPerson = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
            $perUnit = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
            $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
            $fap = 0;
            foreach ($faps as $f) {
                $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
            }

            $sheet->setCellValue([$curColumn,$curRow], ($perPerson + $perUnit + $fap));
            $rowOffset++;
        }

    }
    $colOfset += 12;

    // 2.6 ПАИ
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Плановые объемы медицинской помощи в связи с заболеваниями в амбулаторных условиях на  $year год (Паталого анатомическое исследование биопсийного материала с целью диагностики онкологических заболеваний)");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11, $firstTableHeadRowIndex]);
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, услуг');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'услуг');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);

    $category = 'polyclinic';
    $indicatorId = 6; // услуг
    $serviceId = MedicalServicesEnum::PathologicalAnatomicalBiopsyMaterial;
    /*
    $rowOffset = 0;
    foreach($moCollection as $mo) {
        $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        $fap = 0;
        foreach ($faps as $f) {
            $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
        }

        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $sheet->setCellValue([$curColumn,$curRow], ($perPerson + $perUnit + $fap));
        $rowOffset++;
    }
    $colOfset += 1;
    */
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);
        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $perPerson = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
            $perUnit = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
            $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
            $fap = 0;
            foreach ($faps as $f) {
                $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
            }
            $sheet->setCellValue([$curColumn,$curRow], ($perPerson + $perUnit + $fap));
            $rowOffset++;
        }
    }
    $colOfset += 12;

    // 2.7 МГИ
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Плановые объемы медицинской помощи в связи с заболеваниями в амбулаторных условиях на  $year год (Малекулярно-генетические исследования с целью диагностики онкологических заболеваний)");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11, $firstTableHeadRowIndex]);
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, услуг');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'услуг');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);

    $category = 'polyclinic';
    $indicatorId = 6; // услуг
    $serviceId = MedicalServicesEnum::MolecularGeneticDetectionOncological;
    /*
    $rowOffset = 0;
    foreach($moCollection as $mo) {
        $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        $fap = 0;
        foreach ($faps as $f) {
            $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
        }

        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $sheet->setCellValue([$curColumn,$curRow], ($perPerson + $perUnit + $fap));
        $rowOffset++;
    }
    $colOfset += 1;
    */
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);
        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $perPerson = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
            $perUnit = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
            $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
            $fap = 0;
            foreach ($faps as $f) {
                $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
            }
            $sheet->setCellValue([$curColumn,$curRow], ($perPerson + $perUnit + $fap));
            $rowOffset++;
        }
    }
    $colOfset += 12;

    // 2.8  Тест.covid-19
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Плановые объемы медицинской помощи в связи с заболеваниями в амбулаторных условиях на  $year год (Тестирование на выявление covid-19)");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11, $firstTableHeadRowIndex]);
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, услуг');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'услуг');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);

    $category = 'polyclinic';
    $indicatorId = 6; // услуг
    $serviceId = MedicalServicesEnum::CovidTesting;
    /*
    $rowOffset = 0;
    foreach($moCollection as $mo) {
        $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        $fap = 0;
        foreach ($faps as $f) {
            $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
        }

        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $sheet->setCellValue([$curColumn,$curRow], ($perPerson + $perUnit + $fap));
        $rowOffset++;
    }
    $colOfset += 1;
    */
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);
        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $perPerson = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
            $perUnit = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
            $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
            $fap = 0;
            foreach ($faps as $f) {
                $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
            }
            $sheet->setCellValue([$curColumn,$curRow], ($perPerson + $perUnit + $fap));
            $rowOffset++;
        }
    }
    $colOfset += 12;

    // 3.3 УЗИ плода
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Плановые объемы медицинской помощи в амбулаторных условиях на $year год, УЗИ плода (1 триместр)");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11, $firstTableHeadRowIndex]);
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, услуг');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'услуг');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);

    $category = 'polyclinic';
    $indicatorId = 6; // услуг
    $serviceId = MedicalServicesEnum::FetalUltrasound;
    /*
    $rowOffset = 0;
    foreach($moCollection as $mo) {
        $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        $fap = 0;
        foreach ($faps as $f) {
            $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
        }

        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $sheet->setCellValue([$curColumn,$curRow], ($perPerson + $perUnit + $fap));
        $rowOffset++;
    }
    $colOfset += 1;
    */
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);
        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $perPerson = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
            $perUnit = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
            $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
            $fap = 0;
            foreach ($faps as $f) {
                $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
            }
            $sheet->setCellValue([$curColumn,$curRow], ($perPerson + $perUnit + $fap));
            $rowOffset++;
        }
    }
    $colOfset += 12;

    // 3.4 Компл.иссл. репрод.орг.
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Плановые объемы медицинской помощи в амбулаторных условиях на $year год, комплексное исследование для диагностики фоновых и предраковых заболеваний репродуктивных органов у женщин");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11, $firstTableHeadRowIndex]);
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, услуг');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'услуг');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);

    $category = 'polyclinic';
    $indicatorId = 6; // услуг
    $serviceId = MedicalServicesEnum::DiagnosisBackgroundPrecancerousDiseasesReproductiveWomen;
    /*
    $rowOffset = 0;
    foreach($moCollection as $mo) {
        $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        $fap = 0;
        foreach ($faps as $f) {
            $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
        }

        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $sheet->setCellValue([$curColumn,$curRow], ($perPerson + $perUnit + $fap));
        $rowOffset++;
    }
    $colOfset += 1;
    */
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);
        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $perPerson = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
            $perUnit = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
            $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
            $fap = 0;
            foreach ($faps as $f) {
                $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
            }
            $sheet->setCellValue([$curColumn,$curRow], ($perPerson + $perUnit + $fap));
            $rowOffset++;
        }
    }
    $colOfset += 12;

    // 3.5 Опред.антигена D
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Плановые объемы медицинской помощи в амбулаторных условиях на $year год, определение антигена D системы Резус (резус-фактор плода)");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11, $firstTableHeadRowIndex]);
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, услуг');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'услуг');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);

    $category = 'polyclinic';
    $indicatorId = 6; // услуг
    $serviceId = MedicalServicesEnum::DeterminationAntigenD;
    /*
    $rowOffset = 0;
    foreach($moCollection as $mo) {
        $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        $fap = 0;
        foreach ($faps as $f) {
            $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
        }

        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $sheet->setCellValue([$curColumn,$curRow], ($perPerson + $perUnit + $fap));
        $rowOffset++;
    }
    $colOfset += 1;
    */
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);
        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $perPerson = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
            $perUnit = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
            $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
            $fap = 0;
            foreach ($faps as $f) {
                $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
            }
            $sheet->setCellValue([$curColumn,$curRow], ($perPerson + $perUnit + $fap));
            $rowOffset++;
        }
    }
    $colOfset += 12;

    // 5. Круглосуточный ст.
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Объемы медицинской помощи в условиях круглосуточного стационара (не включая ВМП и медицинскую реабилитацию) на $year год");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11, $firstTableHeadRowIndex]);
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, госпитализаций (не включая медицинскую реабилитацию и ВМП)');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'госпитализаций (не включая медицинскую реабилитацию и ВМП)');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);

    $category = 'hospital';
    $indicatorId = 7; // госпитализаций
    /*
    $rowOffset = 0;
    foreach($moCollection as $mo) {
        $bedProfiles = $content['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? [];
        $v = 0;
        foreach ($bedProfiles as $bp) {
            $v += $bp[$indicatorId];
        }

        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $sheet->setCellValue([$curColumn,$curRow], $v);
        $rowOffset++;
    }
    $colOfset += 1;
    */
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);
        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $bedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? [];
            $v = 0;
            foreach ($bedProfiles as $bp) {
                $v += $bp[$indicatorId] ?? 0;
            }
            $sheet->setCellValue([$curColumn,$curRow], $v);
            $rowOffset++;
        }
    }
    $colOfset += 12;

    // 6.ВМП
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Объемы высокотехнологичной медицинской помощи в условиях круглосуточного стационара на $year год");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11, $firstTableHeadRowIndex]);
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Объемы высокотехнологичной медицинской помощи, всего, госпитализаций');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'госпитализаций');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);

    $category = 'hospital';
    $indicatorId = 7; // госпитализаций
    /*
    $rowOffset = 0;
    foreach($moCollection as $mo) {
        $careProfiles = $content['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? [];
        $v = 0;
        foreach ($careProfiles as $vmpGroups) {
            foreach ($vmpGroups as $vmpTypes) {
                foreach ($vmpTypes as $vmpT)
                $v += $vmpT[$indicatorId];
            }
        }

        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $sheet->setCellValue([$curColumn,$curRow], $v);
        $rowOffset++;
    }
    $colOfset += 1;
    */
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);
        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $careProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? [];
            $v = 0;
            foreach ($careProfiles as $vmpGroups) {
                foreach ($vmpGroups as $vmpTypes) {
                    foreach ($vmpTypes as $vmpT)
                    $v += ($vmpT[$indicatorId] ?? 0);
                }
            }
            $sheet->setCellValue([$curColumn,$curRow], $v);
            $rowOffset++;
        }
    }
    $colOfset += 12;

    // 7. Медреабилитация в КС
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Объемы  медицинской реабилитации в условиях круглосуточного стационара на $year год");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11, $firstTableHeadRowIndex]);
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, случаев лечения');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'случаев лечения');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);

    $category = 'hospital';
    $indicatorId = 7; // госпитализаций
    $bedProfileId = 32; // реабилитационные соматические;
    /*
    $rowOffset = 0;
    foreach($moCollection as $mo) {
        $v = $content['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'][$bedProfileId][$indicatorId] ?? 0;

        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $sheet->setCellValue([$curColumn,$curRow], $v);
        $rowOffset++;
    }
    $colOfset += 1;
    */
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);
        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $v = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'][$bedProfileId][$indicatorId] ?? 0;
            $sheet->setCellValue([$curColumn,$curRow], $v);
            $rowOffset++;
        }
    }
    $colOfset += 12;

    // 8. Дневные стационары
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Объемы медицинской помощи в условиях дневных стационаров на $year год");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11, $firstTableHeadRowIndex]);
    // $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, случаев лечения');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'случаев лечения');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);

    $category = 'hospital';
    $indicatorId = 2; // случаев лечения
    /*
    $rowOffset = 0;
    foreach($moCollection as $mo) {
        $v = 0;
        $bedProfiles = $content['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? [];
        foreach ($bedProfiles as $bp) {
            $v += $bp[$indicatorId];
        }

        $bedProfiles = $content['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? [];
        foreach ($bedProfiles as $bp) {
            $v += $bp[$indicatorId];
        }

        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $sheet->setCellValue([$curColumn,$curRow], $v);
        $rowOffset++;
    }
    $colOfset += 1;
    */
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);
        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $v = 0;
            $bedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? [];
            foreach ($bedProfiles as $bp) {
                $v += $bp[$indicatorId] ?? 0;
            }

            $bedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? [];
            foreach ($bedProfiles as $bp) {
                $v += $bp[$indicatorId] ?? 0;
            }
            $sheet->setCellValue([$curColumn,$curRow], $v);
            $rowOffset++;
        }
    }
    $colOfset += 12;

    // 3. ДС, фин.обеспечение
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Финансовое обеспечение медицинской помощи в условиях дневных стационаров на $year год");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11, $firstTableHeadRowIndex]);
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, руб.');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'руб.');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);
    $category = 'hospital';
    $indicatorId = 4; // стоимость
    /*
    $rowOffset = 0;
    foreach($moCollection as $mo) {
        $v = '0';
        $bedProfiles = $content['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? [];
        foreach ($bedProfiles as $bp) {
            $v = bcadd($v, $bp[$indicatorId]);
        }

        $bedProfiles = $content['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? [];
        foreach ($bedProfiles as $bp) {
            $v = bcadd($v, $bp[$indicatorId]);
        }

        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $sheet->setCellValue([$curColumn,$curRow], $v);
        $rowOffset++;
    }
    $colOfset += 1;
    */
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);
        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $v = '0';
            $bedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? [];
            foreach ($bedProfiles as $bp) {
                $v = bcadd($v, $bp[$indicatorId] ?? '0');
            }

            $bedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? [];
            foreach ($bedProfiles as $bp) {
                $v = bcadd($v, $bp[$indicatorId] ?? '0');
            }
            $sheet->setCellValue([$curColumn,$curRow], $v);
            $rowOffset++;
        }
    }
    $colOfset += 12;

    // 4 КС, фин.обеспечение (не включая мед.реабилитацию и ВМП)
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Финансовое обеспечение  медицинской помощи в условиях круглосуточного стационара на $year год (не включая медицинскую реабилитацию и ВМП)");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11, $firstTableHeadRowIndex]);
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, руб.');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'руб.');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);
    $category = 'hospital';
    $bedProfileId = 32; // реабилитационные соматические;
    $indicatorId = 4; // стоимость
    /*
    $rowOffset = 0;
    foreach($moCollection as $mo) {
        $bedProfiles = $content['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? [];
        $v = '0';
        foreach ($bedProfiles as $bpId => $bp) {
            if($bpId === $bedProfileId) {
                continue;
            }
            $v = bcadd($v, $bp[$indicatorId]);
        }
        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $sheet->setCellValue([$curColumn,$curRow], $v);
        $rowOffset++;
    }
    $colOfset += 1;
    */
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);
        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $bedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? [];
            $v = '0';
            foreach ($bedProfiles as $bpId => $bp) {
                if($bpId === $bedProfileId) {
                    continue;
                }
                $v = bcadd($v, $bp[$indicatorId] ?? '0');
            }
            $sheet->setCellValue([$curColumn,$curRow], $v);
            $rowOffset++;
        }
    }
    $colOfset += 12;

    // 5 МР, фин.обеспечение
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Финансовое обеспечение медицинской реабилитации в условиях круглосуточного стационара на $year год");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11, $firstTableHeadRowIndex]);
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, руб.');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'руб.');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);
    $category = 'hospital';
    $bedProfileId = 32; // реабилитационные соматические;
    $indicatorId = 4; // стоимость
    /*
    $rowOffset = 0;
    foreach($moCollection as $mo) {
        $v = $content['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'][$bedProfileId][$indicatorId] ?? 0;

        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $sheet->setCellValue([$curColumn,$curRow], $v);
        $rowOffset++;
    }
    $colOfset += 1;
    */
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);
        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $v = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'][$bedProfileId][$indicatorId] ?? 0;
            $sheet->setCellValue([$curColumn,$curRow], $v);
            $rowOffset++;
        }
    }
    $colOfset += 12;

    // 6 ВМП, фин.обеспечение
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Финансовое обеспечение ВМП в условиях круглосуточного стационара на $year год");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11, $firstTableHeadRowIndex]);
    //$sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, руб.');
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'руб.');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);
    $category = 'hospital';
    $indicatorId = 4; // стоимость
    /*
    $rowOffset = 0;
    foreach($moCollection as $mo) {
        $careProfiles = $content['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? [];
        $v = '0';
        foreach ($careProfiles as $vmpGroups) {
            foreach ($vmpGroups as $vmpTypes) {
                foreach ($vmpTypes as $vmpT) {
                    $v = bcadd($v, $vmpT[$indicatorId] ?? '0');
                }
            }
        }

        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $sheet->setCellValue([$curColumn, $curRow], $v);
        $rowOffset++;
    }
    $colOfset += 1;
    */
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);
        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $careProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? [];
            $v = '0';
            foreach ($careProfiles as $vmpGroups) {
                foreach ($vmpGroups as $vmpTypes) {
                    foreach ($vmpTypes as $vmpT) {
                        $v = bcadd($v, $vmpT[$indicatorId] ?? '0');
                    }
                }
            }
            $sheet->setCellValue([$curColumn, $curRow], $v);
            $rowOffset++;
        }
    }
    $colOfset += 12;


    // 2. АП фин.обесп.
    $curColumn = $firstTableColIndex + $colOfset;
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex], "Финансовое обеспечение медицинской помощи в  амбулаторных условиях на $year год");
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex, $curColumn + 11 + 12 + 12 + 12, $firstTableHeadRowIndex]);
    /*
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Всего, руб.');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn, $firstTableHeadRowIndex + 3]);
    $sheet->setCellValue([$curColumn + 1, $firstTableHeadRowIndex + 1], 'в том числе');
    $sheet->mergeCells([$curColumn + 1, $firstTableHeadRowIndex + 1, $curColumn + 4, $firstTableHeadRowIndex + 1]);
    $sheet->setCellValue([$curColumn + 1, $firstTableHeadRowIndex + 2], 'Финансовое обеспечение медицинской помощи по подушевому нормативу финансирования на прикрепившихся лиц');
    $sheet->mergeCells([$curColumn + 1, $firstTableHeadRowIndex + 2, $curColumn + 1, $firstTableHeadRowIndex + 3]);
    $sheet->setCellValue([$curColumn + 2, $firstTableHeadRowIndex + 2], 'Финансовое обеспечение медицинской помощи по нормативу финансирования структурного подразделения');
    $sheet->mergeCells([$curColumn + 2, $firstTableHeadRowIndex + 2, $curColumn + 2, $firstTableHeadRowIndex + 3]);
    $sheet->setCellValue([$curColumn + 3, $firstTableHeadRowIndex + 2], 'Финансовое обеспечение медицинской помощи в амбулаторных условиях за единицу объема медицинской помощи');
    $sheet->mergeCells([$curColumn + 3, $firstTableHeadRowIndex + 2, $curColumn + 4, $firstTableHeadRowIndex + 2]);
    $sheet->setCellValue([$curColumn + 3, $firstTableHeadRowIndex + 3], 'проведение диагностических исследований');
    $sheet->setCellValue([$curColumn + 4, $firstTableHeadRowIndex + 3], 'посещения, обращения');
    */
    $sheet->setCellValue([$curColumn, $firstTableHeadRowIndex + 1], 'Финансовое обеспечение медицинской помощи по подушевому нормативу финансирования на прикрепившихся лиц');
    $sheet->mergeCells([$curColumn, $firstTableHeadRowIndex + 1, $curColumn + 11, $firstTableHeadRowIndex + 1]);

    $sheet->setCellValue([$curColumn + 12, $firstTableHeadRowIndex + 1], 'Финансовое обеспечение медицинской помощи по нормативу финансирования структурного подразделения');
    $sheet->mergeCells([$curColumn + 12, $firstTableHeadRowIndex + 1, $curColumn + 11 + 12, $firstTableHeadRowIndex + 1]);

    $sheet->setCellValue([$curColumn + 12 + 12, $firstTableHeadRowIndex + 1], 'Финансовое обеспечение медицинской помощи в амбулаторных условиях за единицу объема медицинской помощи (проведение диагностических исследований)');
    $sheet->mergeCells([$curColumn + 12 + 12, $firstTableHeadRowIndex + 1, $curColumn + 11 + 12 + 12, $firstTableHeadRowIndex + 1]);

    $sheet->setCellValue([$curColumn + 12 + 12 + 12, $firstTableHeadRowIndex + 1], 'Финансовое обеспечение медицинской помощи в амбулаторных условиях за единицу объема медицинской помощи (посещения, обращения)');
    $sheet->mergeCells([$curColumn + 12 + 12 + 12, $firstTableHeadRowIndex + 1, $curColumn + 11 + 12 + 12 + 12, $firstTableHeadRowIndex + 1]);

    $category = 'polyclinic';
    $assistanceTypeId = 4; //обращения по заболеваниям
    $indicatorId = 4; // стоимость
    /*
    $rowOffset = 0;
    foreach($moCollection as $mo) {

        $perPersonSum = '0';
        $assistanceTypesPerPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'] ?? [];
        $servicesPerPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['services'] ?? [];
        foreach ($assistanceTypesPerPerson as $assistanceType) {
            $perPersonSum = bcadd($perPersonSum, $assistanceType[$indicatorId] ?? '0');
        }
        foreach ($servicesPerPerson as $service) {
            $perPersonSum = bcadd($perPersonSum, $service[$indicatorId] ?? '0');
        }

        $perUnitAssistanceTypesSum = '0';
        $perUnitServicesSum = '0';
        $assistanceTypesPerUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'] ?? [];
        $servicesPerPerUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['services'] ?? [];
        foreach ($assistanceTypesPerUnit as $assistanceType) {
            $perUnitAssistanceTypesSum = bcadd($perUnitAssistanceTypesSum, $assistanceType[$indicatorId] ?? '0');
        }
        foreach ($servicesPerPerUnit as $service) {
            $perUnitServicesSum = bcadd($perUnitServicesSum, $service[$indicatorId] ?? '0');
        }

        $fapSum = '0';
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        foreach ($faps as $f) {
            $assistanceTypes = $f['assistanceTypes'] ?? [];
            foreach ($assistanceTypes as $assistanceType) {
                $fapSum = bcadd($fapSum, $assistanceType[$indicatorId] ?? 0);
            }
            $services = $f['services'] ?? [];
            foreach ($services as $service) {
                $fapSum = bcadd($fapSum, $service[$indicatorId] ?? 0);
            }
        }


        $curRow = $firstTableDataRowIndex + $rowOffset;
        $curColumn = $firstTableColIndex + $colOfset;
        $polyclinicTotal = bcadd(bcadd($perPersonSum, $fapSum),bcadd($perUnitServicesSum, $perUnitAssistanceTypesSum));
        $sheet->setCellValue([$curColumn, $curRow], $polyclinicTotal);
        $sheet->setCellValue([$curColumn + 1, $curRow], $perPersonSum);
        $sheet->setCellValue([$curColumn + 2, $curRow], $fapSum);
        $sheet->setCellValue([$curColumn + 3, $curRow], $perUnitServicesSum);
        $sheet->setCellValue([$curColumn + 4, $curRow], $perUnitAssistanceTypesSum);
        $rowOffset++;
    }
    $colOfset += 5;
    */
    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);
        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $perPersonSum = '0';
            $assistanceTypesPerPerson = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'] ?? [];
            $servicesPerPerson = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['services'] ?? [];
            foreach ($assistanceTypesPerPerson as $assistanceType) {
                $perPersonSum = bcadd($perPersonSum, $assistanceType[$indicatorId] ?? '0');
            }
            foreach ($servicesPerPerson as $service) {
                $perPersonSum = bcadd($perPersonSum, $service[$indicatorId] ?? '0');
            }
            $sheet->setCellValue([$curColumn, $curRow], $perPersonSum);
            $rowOffset++;
        }
    }
    $colOfset += 12;

    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);
        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $fapSum = '0';
            $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
            foreach ($faps as $f) {
                $assistanceTypes = $f['assistanceTypes'] ?? [];
                foreach ($assistanceTypes as $assistanceType) {
                    $fapSum = bcadd($fapSum, $assistanceType[$indicatorId] ?? 0);
                }
                $services = $f['services'] ?? [];
                foreach ($services as $service) {
                    $fapSum = bcadd($fapSum, $service[$indicatorId] ?? 0);
                }
            }
            $sheet->setCellValue([$curColumn, $curRow], $fapSum);
            $rowOffset++;
        }
    }
    $colOfset += 12;

    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);
        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;
            $perUnitServicesSum = '0';
            $servicesPerPerUnit = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['services'] ?? [];
            foreach ($servicesPerPerUnit as $service) {
                $perUnitServicesSum = bcadd($perUnitServicesSum, $service[$indicatorId] ?? '0');
            }

            $sheet->setCellValue([$curColumn, $curRow], $perUnitServicesSum);
            $rowOffset++;
        }
    }
    $colOfset += 12;

    // помесячно
    for($monthNum = 1; $monthNum <= 12; $monthNum++)
    {
        $sheet->setCellValue([$firstTableColIndex + $colOfset + $monthNum - 1, $firstTableHeadRowIndex + 2], $months[$monthNum]);
        $rowOffset = 0;
        foreach($moCollection as $mo) {
            $curRow = $firstTableDataRowIndex + $rowOffset;
            $curColumn = $firstTableColIndex + $colOfset + $monthNum - 1;

            $perUnitAssistanceTypesSum = '0';
            $assistanceTypesPerUnit = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'] ?? [];
            foreach ($assistanceTypesPerUnit as $assistanceType) {
                $perUnitAssistanceTypesSum = bcadd($perUnitAssistanceTypesSum, $assistanceType[$indicatorId] ?? '0');
            }
            $sheet->setCellValue([$curColumn, $curRow], $perUnitAssistanceTypesSum);
            $rowOffset++;
        }
    }
    $colOfset += 12;

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


Route::get('/vitacore-v2/{year}/{commissionDecisionsId?}', function (DataForContractService $dataForContractService, PeopleAssignedInfoForContractService $peopleAssignedInfoForContractService, int $year, int $commissionDecisionsId = null) {
    $packageIds = null;
    if ($commissionDecisionsId) {
        $commissionDecisions = CommissionDecision::whereYear('date',$year)->where('id', '<=', $commissionDecisionsId)->get();
        $cd = $commissionDecisions->find($commissionDecisionsId);
        $commissionDecisionIds = $commissionDecisions->pluck('id')->toArray();
        $protocolDate = $cd->date->format('d.m.Y');
        $packageIds = ChangePackage::whereIn('commission_decision_id', $commissionDecisionIds)->orWhere('commission_decision_id', null)->pluck('id')->toArray();
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
    $moCollection = MedicalInstitution::orderBy('order')->get();

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
            $values[$monthNum] = medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
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
            $values[$monthNum] = medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
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
            $values[$monthNum] = medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
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
            $values[$monthNum] = medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
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
            $values[$monthNum] = medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
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
            $values[$monthNum] = medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
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
            $values[$monthNum] = medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
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
            $values[$monthNum] = medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
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
            $values[$monthNum] = medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
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
            $values[$monthNum] = medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
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
            $values[$monthNum] = medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
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
            $values[$monthNum] = medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
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
            $values[$monthNum] = medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
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
            $values[$monthNum] = medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
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
            $values[$monthNum] = medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
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
            $values[$monthNum] = medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
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
            $values[$monthNum] = medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
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
            $values[$monthNum] = medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
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
            $values[$monthNum] = medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
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
            $values[$monthNum] = medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
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
        $rehabilitationBedProfileId = 32; // реабилитационные соматические;
        $hasValue = false;
        $planningSectionName = "Круглосуточный стационар (не включая ВМП и медицинскую реабилитацию)";
        $planningParamName = "объемы, госпитализаций";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $bedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? [];
            $values[$monthNum] = 0;
            foreach ($bedProfiles as $bpId => $bp) {
                if ($bpId === $rehabilitationBedProfileId) {
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
        $rehabilitationBedProfileId = 32; // реабилитационные соматические;
        $hasValue = false;
        $planningSectionName = "Медицинская реабилитация";
        $planningParamName = "объемы, госпитализаций";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'][$rehabilitationBedProfileId][$indicatorId] ?? 0;
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
        $rehabilitationBedProfileId = 32; // реабилитационные соматические;
        $hasValue = false;
        $planningSectionName = "Круглосуточный стационар (не включая ВМП и медицинскую реабилитацию)";
        $planningParamName = "финансовое обеспечение, руб.";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = 0;
            $bedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? [];
            foreach ($bedProfiles as $bpId => $bp) {
                if($bpId === $rehabilitationBedProfileId) {
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
        $rehabilitationBedProfileId = 32; // реабилитационные соматические;
        $indicatorId = 4; // стоимость
        $hasValue = false;
        $planningSectionName = "Медицинская реабилитация";
        $planningParamName = "финансовое обеспечение, руб.";

        $values = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++) {
            $values[$monthNum] = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'][$rehabilitationBedProfileId][$indicatorId] ?? 0;
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


Route::get('/summary-volume/{year}/{commissionDecisionsId?}', function (DataForContractService $dataForContractService, PeopleAssignedInfoForContractService $peopleAssignedInfoForContractService, int $year, int $commissionDecisionsId = null) {
    $packageIds = null;
    if ($commissionDecisionsId) {
        $commissionDecisions = CommissionDecision::whereYear('date',$year)->where('id', '<=', $commissionDecisionsId)->get();
        $cd = $commissionDecisions->find($commissionDecisionsId);
        $commissionDecisionIds = $commissionDecisions->pluck('id')->toArray();
        $protocolDate = $cd->date->format('d.m.Y');
        $docName = "к протоколу заседания комиссии по разработке территориальной программы ОМС Курганской области от $protocolDate";
        $packageIds = ChangePackage::whereIn('commission_decision_id', $commissionDecisionIds)->orWhere('commission_decision_id', null)->pluck('id')->toArray();
    } else {
        $packageIds = ChangePackage::where('commission_decision_id', null)->pluck('id')->toArray();
    }
    $path = 'xlsx';
    $templateFileName = '1.xlsx';
    $templateFilePath = $path . DIRECTORY_SEPARATOR . $templateFileName;
    $templateFullFilepath = Storage::path($templateFilePath);
    $resultFileName = 'объемы.xlsx';
    $strDateTimeNow = date("Y-m-d-His");
    $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . ' ' . $resultFileName;
    $fullResultFilepath = Storage::path($resultFilePath);

    bcscale(4);

    $content = $dataForContractService->GetArray($year, $packageIds);
    $peopleAssigned = $peopleAssignedInfoForContractService->GetArray($year, $packageIds);
    $moCollection = MedicalInstitution::orderBy('order')->get();

    $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader("Xlsx");
    $spreadsheet = $reader->load($templateFullFilepath);
    $sheet = $spreadsheet->getSheetByName('1.Скорая помощь');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'ambulance';
    $indicatorId = 5; // вызовов
    $callsAssistanceTypeId = 5; // вызовы
    $thrombolysisAssistanceTypeId = 6;// тромболизис
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        $sheet->setCellValue([7,$rowIndex], $peopleAssigned[$mo->id][$category]['peopleAssigned'] ?? 0);
        $thrombolysis = $content['mo'][$mo->id][$category][$thrombolysisAssistanceTypeId][$indicatorId] ?? 0;
        $sheet->setCellValue([8,$rowIndex], ($content['mo'][$mo->id][$category][$callsAssistanceTypeId][$indicatorId] ?? 0) + $thrombolysis);
        $sheet->setCellValue([9,$rowIndex], $thrombolysis);
    }

    $sheet = $spreadsheet->getSheetByName('2.обращения по заболеваниям');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'polyclinic';
    $indicatorId = 8; // обращений
    $assistanceTypeId = 4; //обращения по заболеваниям
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        $sheet->setCellValue([7,$rowIndex], $peopleAssigned[$mo->id][$category]['mo']['peopleAssigned'] ?? 0);
        $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
        $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        $fap = 0;
        foreach ($faps as $f) {
            $fap += $f['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
        }

        $sheet->setCellValue([8,$rowIndex], $perPerson + $perUnit + $fap );
    }

    $sheet = $spreadsheet->getSheetByName('3.Посещения с иными целями');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'polyclinic';
    $indicatorId = 9; // посещений
    $assistanceTypeIds = [1, 2]; //	посещения профилактические, посещения разовые по заболеваниям
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        $sheet->setCellValue([7,$rowIndex], $peopleAssigned[$mo->id][$category]['mo']['peopleAssigned'] ?? 0);
        $v = 0;
        foreach($assistanceTypeIds as $assistanceTypeId) {
            $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
            $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
            $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
            $fap = 0;
            foreach ($faps as $f) {
                $fap += $f['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
            }
            $v += ($perPerson + $perUnit + $fap);
        }

        $sheet->setCellValue([8,$rowIndex], $v);
    }

    $sheet = $spreadsheet->getSheetByName('4 Неотложная помощь');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'polyclinic';
    $indicatorId = 9; // посещений
    $assistanceTypeIds = [3]; //	посещения неотложные
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        $sheet->setCellValue([7,$rowIndex], $peopleAssigned[$mo->id][$category]['mo']['peopleAssigned'] ?? 0);
        $v = 0;
        foreach($assistanceTypeIds as $assistanceTypeId) {
            $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
            $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
            $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
            $fap = 0;
            foreach ($faps as $f) {
                $fap += $f['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
            }
            $v += ($perPerson + $perUnit + $fap);
        }

        $sheet->setCellValue([8,$rowIndex], $v);
    }

    $sheet = $spreadsheet->getSheetByName('2.2 КТ');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'polyclinic';
    $indicatorId = 6; // услуг
    $serviceId = MedicalServicesEnum::KT;
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        $sheet->setCellValue([7,$rowIndex], $peopleAssigned[$mo->id][$category]['mo']['peopleAssigned'] ?? 0);

        $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        $fap = 0;
        foreach ($faps as $f) {
            $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
        }

        $sheet->setCellValue([8,$rowIndex], ($perPerson + $perUnit + $fap));
    }

    $sheet = $spreadsheet->getSheetByName('2.3 МРТ');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'polyclinic';
    $indicatorId = 6; // услуг
    $serviceId = MedicalServicesEnum::MRT;
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        $sheet->setCellValue([7,$rowIndex], $peopleAssigned[$mo->id][$category]['mo']['peopleAssigned'] ?? 0);

        $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        $fap = 0;
        foreach ($faps as $f) {
            $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
        }

        $sheet->setCellValue([8,$rowIndex], ($perPerson + $perUnit + $fap));
    }

    $sheet = $spreadsheet->getSheetByName('2.4 УЗИ ССС');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'polyclinic';
    $indicatorId = 6; // услуг
    $serviceId = MedicalServicesEnum::UltrasoundCardio;
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        $sheet->setCellValue([7,$rowIndex], $peopleAssigned[$mo->id][$category]['mo']['peopleAssigned'] ?? 0);

        $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        $fap = 0;
        foreach ($faps as $f) {
            $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
        }

        $sheet->setCellValue([8,$rowIndex], ($perPerson + $perUnit + $fap));
    }

    $sheet = $spreadsheet->getSheetByName('2.5 Эндоскопия');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'polyclinic';
    $indicatorId = 6; // услуг
    $serviceId = MedicalServicesEnum::Endoscopy;
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        $sheet->setCellValue([7,$rowIndex], $peopleAssigned[$mo->id][$category]['mo']['peopleAssigned'] ?? 0);

        $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        $fap = 0;
        foreach ($faps as $f) {
            $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
        }

        $sheet->setCellValue([8,$rowIndex], ($perPerson + $perUnit + $fap));
    }

    $sheet = $spreadsheet->getSheetByName('2.6 ПАИ');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'polyclinic';
    $indicatorId = 6; // услуг
    $serviceId = MedicalServicesEnum::PathologicalAnatomicalBiopsyMaterial;
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        $sheet->setCellValue([7,$rowIndex], $peopleAssigned[$mo->id][$category]['mo']['peopleAssigned'] ?? 0);

        $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        $fap = 0;
        foreach ($faps as $f) {
            $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
        }

        $sheet->setCellValue([8,$rowIndex], ($perPerson + $perUnit + $fap));
    }

    $sheet = $spreadsheet->getSheetByName('2.7 МГИ');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'polyclinic';
    $indicatorId = 6; // услуг
    $serviceId = MedicalServicesEnum::MolecularGeneticDetectionOncological;
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        $sheet->setCellValue([7,$rowIndex], $peopleAssigned[$mo->id][$category]['mo']['peopleAssigned'] ?? 0);

        $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        $fap = 0;
        foreach ($faps as $f) {
            $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
        }

        $sheet->setCellValue([8,$rowIndex], ($perPerson + $perUnit + $fap));
    }

    $sheet = $spreadsheet->getSheetByName('2.8  Тест.covid-19');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'polyclinic';
    $indicatorId = 6; // услуг
    $serviceId = MedicalServicesEnum::CovidTesting;
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        $sheet->setCellValue([7,$rowIndex], $peopleAssigned[$mo->id][$category]['mo']['peopleAssigned'] ?? 0);

        $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        $fap = 0;
        foreach ($faps as $f) {
            $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
        }

        $sheet->setCellValue([8,$rowIndex], ($perPerson + $perUnit + $fap));
    }

    $sheet = $spreadsheet->getSheetByName('3.3 УЗИ плода');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'polyclinic';
    $indicatorId = 6; // услуг
    $serviceId = MedicalServicesEnum::FetalUltrasound;
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        $sheet->setCellValue([7,$rowIndex], $peopleAssigned[$mo->id][$category]['mo']['peopleAssigned'] ?? 0);

        $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        $fap = 0;
        foreach ($faps as $f) {
            $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
        }

        $sheet->setCellValue([8,$rowIndex], ($perPerson + $perUnit + $fap));
    }

    $sheet = $spreadsheet->getSheetByName('3.4 Компл.иссл. репрод.орг.');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'polyclinic';
    $indicatorId = 6; // услуг
    $serviceId = MedicalServicesEnum::DiagnosisBackgroundPrecancerousDiseasesReproductiveWomen;
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        $sheet->setCellValue([7,$rowIndex], $peopleAssigned[$mo->id][$category]['mo']['peopleAssigned'] ?? 0);

        $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        $fap = 0;
        foreach ($faps as $f) {
            $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
        }

        $sheet->setCellValue([8,$rowIndex], ($perPerson + $perUnit + $fap));
    }

    $sheet = $spreadsheet->getSheetByName('3.5 Опред.антигена D');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'polyclinic';
    $indicatorId = 6; // услуг
    $serviceId = MedicalServicesEnum::DeterminationAntigenD;
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        $sheet->setCellValue([7,$rowIndex], $peopleAssigned[$mo->id][$category]['mo']['peopleAssigned'] ?? 0);

        $perPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $perUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? 0;
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        $fap = 0;
        foreach ($faps as $f) {
            $fap += $f['services'][$serviceId][$indicatorId] ?? 0;
        }

        $sheet->setCellValue([8,$rowIndex], ($perPerson + $perUnit + $fap));
    }

    $sheet = $spreadsheet->getSheetByName('5. Круглосуточный ст.');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'hospital';
    $indicatorId = 7; // госпитализаций
    $rehabilitationBedProfileId = 32; // реабилитационные соматические;
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);

        $bedProfiles = $content['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? [];
        $v = 0;
        foreach ($bedProfiles as $bpId => $bp) {
            if($bpId === $rehabilitationBedProfileId) {
                continue;
            }
            $v += ($bp[$indicatorId] ?? 0);
        }

        $sheet->setCellValue([7,$rowIndex], $v);
    }

    $sheet = $spreadsheet->getSheetByName('6.ВМП');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'hospital';
    $indicatorId = 7; // госпитализаций
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);

        $careProfiles = $content['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? [];
        $v = 0;
        foreach ($careProfiles as $vmpGroups) {
            foreach ($vmpGroups as $vmpTypes) {
                foreach ($vmpTypes as $vmpT)
                $v += $vmpT[$indicatorId] ?? 0;
            }
        }

        $sheet->setCellValue([7,$rowIndex], $v);
    }

    $sheet = $spreadsheet->getSheetByName('7. Медреабилитация в КС');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'hospital';
    $indicatorId = 7; // госпитализаций
    $bedProfileId = 32; // реабилитационные соматические;
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);

        $v = $content['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'][$bedProfileId][$indicatorId] ?? 0;

        $sheet->setCellValue([7,$rowIndex], $v);
    }

    $sheet = $spreadsheet->getSheetByName('8. Дневные стационары');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'hospital';
    $indicatorId = 2; // случаев лечения
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);

        $v = 0;
        $bedProfiles = $content['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? [];
        foreach ($bedProfiles as $bp) {
            $v += $bp[$indicatorId] ?? 0;
        }

        $bedProfiles = $content['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? [];
        foreach ($bedProfiles as $bp) {
            $v += $bp[$indicatorId] ?? 0;
        }

        $sheet->setCellValue([7,$rowIndex], $v);
    }

    $sheet = $spreadsheet->getSheetByName("6.1. ВМП в разрезе методов");
    $vmpGroups = VmpGroup::all();
    $vmpTypes = VmpTypes::all();
    $careProfiles = CareProfiles::orderBy('id')->get();
    //dd($careProfiles[1]);
    $rowIndex = 7;
    $vmpGroupCodeColIndex = 1;
    $vmpTypeNameColIndex = 2;
    $category = 'hospital';
    $indicatorId = 7; // госпитализаций
    $indicatorCostId = 4; // стоимость
    $vmp = [];

    foreach($moCollection as $mo) {
        $profiles = $content['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? [];
        foreach ($profiles as $careProfileId => $groups) {
            if (!isset($vmp[$careProfileId])) {
                $vmp[$careProfileId] = [];
            }
            foreach ($groups as $groupId => $types) {
                if (!isset($vmp[$careProfileId][$groupId])) {
                    $vmp[$careProfileId][$groupId] = [];
                }
                foreach ($types as $typeId => $vmpT) {
                    $vmp[$careProfileId][$groupId][] = $typeId;
                }
            }
        }
    }
    foreach($vmp as $careProfileId => $groups) {
        $sheet->setCellValue([$vmpGroupCodeColIndex, $rowIndex], $careProfiles->firstWhere('id', $careProfileId)->name);
        $sheet->mergeCells([$vmpGroupCodeColIndex, $rowIndex, $vmpTypeNameColIndex, $rowIndex]);
        $sheet->getStyle([$vmpGroupCodeColIndex, $rowIndex, $vmpTypeNameColIndex, $rowIndex])
            ->getAlignment()
            ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle([$vmpGroupCodeColIndex, $rowIndex, $vmpTypeNameColIndex, $rowIndex])
            ->getFont()
            ->setBold( true );

        $rowIndex++;
        foreach ($vmpGroups as $vmpGroup) {
            if(!isset($groups[$vmpGroup->id])){
                continue;
            }
            foreach ($vmpTypes as $vmpType) {
                if(!in_array($vmpType->id, $groups[$vmpGroup->id])){
                    continue;
                }
                $sheet->setCellValue([$vmpGroupCodeColIndex, $rowIndex], $vmpGroup->code);
                $sheet->setCellValue([$vmpTypeNameColIndex, $rowIndex], $vmpType->name);
                $sheet->getRowDimension($rowIndex)->setRowHeight(-1);
                $rowIndex++;
            }
        }
    }
    for($i = $rowIndex; $i <= 116; $i++){
        $sheet->getRowDimension($i)->setVisible(false);
    }

    $rowIndex = 5;
    $colIndex = 3;

    foreach($moCollection as $mo) {
        if(!isset($content['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'])) {
            continue;
        }
        $sheet->setCellValue([$colIndex,$rowIndex], $mo->short_name);
        $rowOffset = 2;
        foreach($vmp as $careProfileId => $groups) {
            $rowOffset++;
            foreach ($vmpGroups as $vmpGroup) {
                if(!isset($groups[$vmpGroup->id])){
                    continue;
                }
                foreach ($vmpTypes as $vmpType) {
                    if(!in_array($vmpType->id, $groups[$vmpGroup->id])){
                        continue;
                    }
                    $v = $content['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'][$careProfileId][$vmpGroup->id][$vmpType->id] ?? [];
                    $sheet->setCellValue([$colIndex, $rowIndex + $rowOffset], $v[$indicatorId] ?? 0);
                    $sheet->setCellValue([$colIndex + 1, $rowIndex + $rowOffset], $v[$indicatorCostId] ?? 0);
                    $rowOffset++;
                }
            }
        }
        $colIndex += 2;
    }
    $rowOffset = 2;
    foreach($vmp as $careProfileId => $groups) {
        $rowOffset++;
        foreach ($vmpGroups as $vmpGroup) {
            if(!isset($groups[$vmpGroup->id])){
                continue;
            }
            foreach ($vmpTypes as $vmpType) {
                if(!in_array($vmpType->id, $groups[$vmpGroup->id])){
                    continue;
                }
                $v = [];
                $v[$indicatorId] = '0';
                $v[$indicatorCostId] = '0';
                foreach($moCollection as $mo) {
                    if(!isset($content['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'])) {
                        continue;
                    }
                    $tmp = $content['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'][$careProfileId][$vmpGroup->id][$vmpType->id] ?? [];
                    $v[$indicatorId] = bcadd($v[$indicatorId], $tmp[$indicatorId] ?? '0');
                    $v[$indicatorCostId] = bcadd($v[$indicatorCostId], $tmp[$indicatorCostId] ?? '0');
                }
                $sheet->setCellValue([$colIndex, $rowIndex + $rowOffset], $v[$indicatorId] ?? 0);
                $sheet->setCellValue([$colIndex + 1, $rowIndex + $rowOffset], $v[$indicatorCostId] ?? 0);
                $rowOffset++;
            }
        }
    }


    $writer = new Xlsx($spreadsheet);
    $writer->save($fullResultFilepath);
    return Storage::download($resultFilePath);
});


Route::get('/summary-cost/{year}/{commissionDecisionsId?}', function (DataForContractService $dataForContractService, PeopleAssignedInfoForContractService $peopleAssignedInfoForContractService, int $year, int $commissionDecisionsId = null) {
    $packageIds = null;
    if ($commissionDecisionsId) {
        $commissionDecisions = CommissionDecision::whereYear('date',$year)->where('id', '<=', $commissionDecisionsId)->get();
        $cd = $commissionDecisions->find($commissionDecisionsId);
        $commissionDecisionIds = $commissionDecisions->pluck('id')->toArray();
        $protocolDate = $cd->date->format('d.m.Y');
        $docName = "к протоколу заседания комиссии по разработке территориальной программы ОМС Курганской области от $protocolDate";
        $packageIds = ChangePackage::whereIn('commission_decision_id', $commissionDecisionIds)->orWhere('commission_decision_id', null)->pluck('id')->toArray();
    } else {
        $packageIds = ChangePackage::where('commission_decision_id', null)->pluck('id')->toArray();
    }

    $path = 'xlsx';
    $templateFileName = '2.xlsx';
    $templateFilePath = $path . DIRECTORY_SEPARATOR . $templateFileName;
    $templateFullFilepath = Storage::path($templateFilePath);
    $resultFileName = 'стоимость.xlsx';
    $strDateTimeNow = date("Y-m-d-His");
    $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . ' ' . $resultFileName;
    $fullResultFilepath = Storage::path($resultFilePath);

    bcscale(4);
    $indicatorId = 4; // стоимость

    $content = $dataForContractService->GetArray($year, $packageIds);
    $peopleAssigned = $peopleAssignedInfoForContractService->GetArray($year, $packageIds);
    $moCollection = MedicalInstitution::orderBy('order')->get();

    $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader("Xlsx");
    $spreadsheet = $reader->load($templateFullFilepath);
    $sheet = $spreadsheet->getSheetByName('1.Скорая помощь, фин.обесп.');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'ambulance';

    $callsAssistanceTypeId = 5; // вызовы
    $thrombolysisAssistanceTypeId = 6;// тромболизис
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        $sheet->setCellValue([7,$rowIndex], $peopleAssigned[$mo->id][$category]['peopleAssigned'] ?? 0);
        $calls = $content['mo'][$mo->id][$category][$callsAssistanceTypeId][$indicatorId] ?? '0';
        $thrombolysis = $content['mo'][$mo->id][$category][$thrombolysisAssistanceTypeId][$indicatorId] ?? '0';
        $sheet->setCellValue([8,$rowIndex], bcadd($calls, $thrombolysis));
    }

    $sheet = $spreadsheet->getSheetByName('2. АП фин.обесп.');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'polyclinic';
    $assistanceTypeId = 4; //обращения по заболеваниям
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);
        $sheet->setCellValue([7,$rowIndex], $peopleAssigned[$mo->id][$category]['mo']['peopleAssigned'] ?? 0);

        $perPersonSum = '0';
        $assistanceTypesPerPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'] ?? [];
        $servicesPerPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['services'] ?? [];
        foreach ($assistanceTypesPerPerson as $assistanceType) {
            $perPersonSum = bcadd($perPersonSum, $assistanceType[$indicatorId] ?? '0');
        }
        foreach ($servicesPerPerson as $service) {
            $perPersonSum = bcadd($perPersonSum, $service[$indicatorId] ?? '0');
        }
       // [$assistanceTypeId]
       // [$serviceId][$indicatorId]

        $perUnitAssistanceTypesSum = '0';
        $perUnitServicesSum = '0';
        $assistanceTypesPerUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'] ?? [];
        $servicesPerPerUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['services'] ?? [];
        foreach ($assistanceTypesPerUnit as $assistanceType) {
            $perUnitAssistanceTypesSum = bcadd($perUnitAssistanceTypesSum, $assistanceType[$indicatorId] ?? '0');
        }
        foreach ($servicesPerPerUnit as $service) {
            $perUnitServicesSum = bcadd($perUnitServicesSum, $service[$indicatorId] ?? '0');
        }

        $fapSum = '0';
        $faps = $content['mo'][$mo->id][$category]['fap'] ?? [];
        foreach ($faps as $f) {
            $assistanceTypes = $f['assistanceTypes'] ?? [];
            foreach ($assistanceTypes as $assistanceType) {
                $fapSum = bcadd($fapSum, $assistanceType[$indicatorId] ?? 0);
            }
            $services = $f['services'] ?? [];
            foreach ($services as $service) {
                $fapSum = bcadd($fapSum, $service[$indicatorId] ?? 0);
            }
        }

        $sheet->setCellValue([9,$rowIndex], $perPersonSum);
        $sheet->setCellValue([10,$rowIndex], $fapSum);
        $sheet->setCellValue([11,$rowIndex], $perUnitServicesSum);
        $sheet->setCellValue([12,$rowIndex], $perUnitAssistanceTypesSum);
    }


    $sheet = $spreadsheet->getSheetByName('3. ДС, фин.обеспечение');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'hospital';
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);

        $v = '0';
        $bedProfiles = $content['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? [];
        foreach ($bedProfiles as $bp) {
            $v = bcadd($v, $bp[$indicatorId] ?? '0');
        }

        $bedProfiles = $content['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? [];
        foreach ($bedProfiles as $bp) {
            $v = bcadd($v, $bp[$indicatorId] ?? '0');
        }

        $sheet->setCellValue([7,$rowIndex], $v);
    }

    // КС (не включая мед.реабилитацию и ВМП)
    $sheet = $spreadsheet->getSheetByName('4 КС, фин.обеспечение ');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'hospital';
    $bedProfileId = 32; // реабилитационные соматические;
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);

        $bedProfiles = $content['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? [];
        $v = '0';
        foreach ($bedProfiles as $bpId => $bp) {
            if($bpId === $bedProfileId) {
                continue;
            }
            $v = bcadd($v, $bp[$indicatorId] ?? '0');
        }

        $sheet->setCellValue([7,$rowIndex], $v);
    }

    $sheet = $spreadsheet->getSheetByName('5 МР, фин.обеспечение ');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'hospital';
    $bedProfileId = 32; // реабилитационные соматические;
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);

        $v = $content['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'][$bedProfileId][$indicatorId] ?? 0;

        $sheet->setCellValue([7,$rowIndex], $v);
    }

    $sheet = $spreadsheet->getSheetByName('6 ВМП, фин.обеспечение  ');
    $ordinalRowNum = 0;
    $rowIndex = 6;
    $category = 'hospital';
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([1,$rowIndex], $mo->code);
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);

        $careProfiles = $content['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? [];
        $v = '0';
        foreach ($careProfiles as $vmpGroups) {
            foreach ($vmpGroups as $vmpTypes) {
                foreach ($vmpTypes as $vmpT) {
                    $v = bcadd($v, $vmpT[$indicatorId] ?? '0');
                }
            }
        }

        $sheet->setCellValue([7,$rowIndex], $v);
    }

    $writer = new Xlsx($spreadsheet);
    $writer->save($fullResultFilepath);
    return Storage::download($resultFilePath);
});


Route::get('/hospital-by-profile/{year}/{commissionDecisionsId?}', function (DataForContractService $dataForContractService, int $year, int $commissionDecisionsId = null) {
    $packageIds = null;
    if ($commissionDecisionsId) {
        $commissionDecisions = CommissionDecision::whereYear('date',$year)->where('id', '<=', $commissionDecisionsId)->get();
        $cd = $commissionDecisions->find($commissionDecisionsId);
        $commissionDecisionIds = $commissionDecisions->pluck('id')->toArray();
        $packageIds = ChangePackage::whereIn('commission_decision_id', $commissionDecisionIds)->orWhere('commission_decision_id', null)->pluck('id')->toArray();
    } else {
        $packageIds = ChangePackage::where('commission_decision_id', null)->pluck('id')->toArray();
    }
    $path = 'xlsx';
    $resultFileName = 'hospital.xlsx';
    $strDateTimeNow = date("Y-m-d-His");
    $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . ' ' . $resultFileName;
    $fullResultFilepath = Storage::path($resultFilePath);

    bcscale(4);
    $indicatorIds = [1, 2, 3, 4, 7];
    $content = $dataForContractService->GetArray($year, $packageIds, $indicatorIds);

    $moCollection = MedicalInstitution::orderBy('order')->get();

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

                $numberOfBeds = bcadd($numberOfBeds, $bpData[$numberOfBedsIndicatorId] ?? '0');
                $casesOfTreatment = bcadd($casesOfTreatment, $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                $patientDays = bcadd($patientDays, $bpData[$patientDaysIndicatorId] ?? '0');
                $cost = bcadd($cost, $bpData[$costIndicatorId] ?? '0');
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

                $numberOfBeds = bcadd($numberOfBeds, $bpData[$numberOfBedsIndicatorId] ?? '0');
                $casesOfTreatment = bcadd($casesOfTreatment, $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                $patientDays = bcadd($patientDays, $bpData[$patientDaysIndicatorId] ?? '0');
                $cost = bcadd($cost, $bpData[$costIndicatorId] ?? '0');
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

                $numberOfBeds = bcadd($numberOfBeds, $bpData[$numberOfBedsIndicatorId] ?? '0');
                $casesOfTreatment = bcadd($casesOfTreatment, $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                $patientDays = bcadd($patientDays, $bpData[$patientDaysIndicatorId] ?? '0');
                $cost = bcadd($cost, $bpData[$costIndicatorId] ?? '0');
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
                        $numberOfBeds = bcadd($numberOfBeds, $vmpT[$numberOfBedsIndicatorId] ?? '0');
                        $casesOfTreatment = bcadd($casesOfTreatment, $vmpT[$casesOfTreatmentIndicatorId] ?? '0');
                        $patientDays = bcadd($patientDays, $vmpT[$patientDaysIndicatorId] ?? '0');
                        $cost = bcadd($cost, $vmpT[$costIndicatorId] ?? '0');
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
    string | int $value
 ) {
    $moCodeColOffset = 1;
    $moNameColOffset = 2;
    $planningSectionColOffset = 3;
    $profileNameColOffset = 4;
    $profileCodeColOffset = 5;
    $paramColOffset = 6;
    $valueColOffset = 7;

    $sheet->setCellValue([$colIndex, $rowIndex], "$ordinalRowNum");
    $sheet->setCellValue([$colIndex + $moCodeColOffset, $rowIndex], $moCode);
    $sheet->setCellValue([$colIndex + $moNameColOffset, $rowIndex], $moName);
    $sheet->setCellValue([$colIndex + $planningSectionColOffset, $rowIndex], $planningSectionName);
    $sheet->setCellValue([$colIndex + $profileNameColOffset, $rowIndex], $profileName);
    $sheet->setCellValue([$colIndex + $profileCodeColOffset, $rowIndex], $profileCode);
    $sheet->setCellValue([$colIndex + $paramColOffset, $rowIndex], $planningParamName);
    $sheet->setCellValue([$colIndex + $valueColOffset, $rowIndex], $value);

}

Route::get('/vitacore-hospital-by-profile/{year}/{commissionDecisionsId?}', function (DataForContractService $dataForContractService, int $year, int $commissionDecisionsId = null) {
    $packageIds = null;
    if ($commissionDecisionsId) {
        $commissionDecisions = CommissionDecision::whereYear('date',$year)->where('id', '<=', $commissionDecisionsId)->get();
        $cd = $commissionDecisions->find($commissionDecisionsId);
        $commissionDecisionIds = $commissionDecisions->pluck('id')->toArray();
        $packageIds = ChangePackage::whereIn('commission_decision_id', $commissionDecisionIds)->orWhere('commission_decision_id', null)->pluck('id')->toArray();
    } else {
        $packageIds = ChangePackage::where('commission_decision_id', null)->pluck('id')->toArray();
    }
    $path = 'xlsx';
    $resultFileName = 'hospital.xlsx';
    $strDateTimeNow = date("Y-m-d-His");
    $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . ' ' . $resultFileName;
    $fullResultFilepath = Storage::path($resultFilePath);

    bcscale(4);

    $indicatorIds = [1, 2, 3, 4, 7];
    $content = $dataForContractService->GetArray($year, $packageIds, $indicatorIds);

    $moCollection = MedicalInstitution::orderBy('order')->get();

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
                    if (!$bpData) { continue; }

                    $numberOfBeds = bcadd($numberOfBeds, $bpData[$numberOfBedsIndicatorId] ?? '0');
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
                    if (!$bpData) { continue; }

                    $numberOfBeds = bcadd($numberOfBeds, $bpData[$numberOfBedsIndicatorId] ?? '0');
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
                    if (!$bpData) { continue; }

                    $numberOfBeds = bcadd($numberOfBeds, $bpData[$numberOfBedsIndicatorId] ?? '0');
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

        $careProfiles = $content['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? null;
        if ($careProfiles) {
            $planningSectionName = 'ВМП';
            $numberOfBedsIndicatorId = 1; // число коек
            $casesOfTreatmentIndicatorId = 7; // госпитализаций
            $patientDaysIndicatorId = 3; // койко-дней
            $costIndicatorId = 4; // стоимость

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
                            $numberOfBeds = bcadd($numberOfBeds, $vmpT[$numberOfBedsIndicatorId] ?? '0');
                            $casesOfTreatment = bcadd($casesOfTreatment, $vmpT[$casesOfTreatmentIndicatorId] ?? '0');
                            $patientDays = bcadd($patientDays, $vmpT[$patientDaysIndicatorId] ?? '0');
                            $cost = bcadd($cost, $vmpT[$costIndicatorId] ?? '0');
                        }
                    }
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
    } // foreach MO

    $writer = new Xlsx($spreadsheet);
    $writer->save($fullResultFilepath);
    return Storage::download($resultFilePath);
});


Route::get('/{year}/{commissionDecisionsId?}', function (DataForContractService $dataForContractService, MoInfoForContractService $moInfoForContractService, MoDepartmentsInfoForContractService $moDepartmentsInfoForContractService, PeopleAssignedInfoForContractService $peopleAssignedInfoForContractService, int $year, int $commissionDecisionsId = null) {
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
    // InitialChanges::dispatch(2022);
    // InitialDataLoaded::dispatch(2, 2022, 1);
    // InitialDataLoaded::dispatch(9, 2022, 1);

    //$year = 2022;
    //$commit = null;
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

    $content = $dataForContractService->GetJson($year, $packageIds);
    Storage::put($path.'data.json', $content);

    $content = $peopleAssignedInfoForContractService->GetJson($year, $packageIds);
    Storage::put($path.'peopleAssignedData.json', $content);
    return $content;
});

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
    $path = 'xlsx';

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
                    $v += $vmpT[$indicatorId];
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
                $v = bcadd($v, $bp[$indicatorId]);
            }

            $bedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? [];
            foreach ($bedProfiles as $bp) {
                $v = bcadd($v, $bp[$indicatorId]);
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
                $v = bcadd($v, $bp[$indicatorId]);
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
    $peopleAssigned = $peopleAssignedInfoForContractService->GetArray($year, null);
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
    $peopleAssigned = $peopleAssignedInfoForContractService->GetArray($year, null);
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

    $content = $peopleAssignedInfoForContractService->GetJson($year, null);
    Storage::put($path.'peopleAssignedData.json', $content);
    return $content;
});

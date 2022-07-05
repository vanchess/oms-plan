<?php

use App\Jobs\InitialChanges;
use App\Jobs\InitialDataLoaded;
use App\Models\CareProfiles;
use App\Models\CareProfilesFoms;
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

Route::get('/xlsx1', function (DataForContractService $dataForContractService, PeopleAssignedInfoForContractService $peopleAssignedInfoForContractService) {
    $path = 'xlsx';
    $templateFileName = '1.xlsx';
    $templateFilePath = $path . DIRECTORY_SEPARATOR . $templateFileName;
    $templateFullFilepath = Storage::path($templateFilePath);
    $resultFileName = 'объемы.xlsx';
    $strDateTimeNow = date("Y-m-d-His");
    $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . ' ' . $resultFileName;
    $fullResultFilepath = Storage::path($resultFilePath);

    bcscale(4);

    $year = 2022;
    $commit = null;
    $content = $dataForContractService->GetArray($year, $commit);
    $peopleAssigned = $peopleAssignedInfoForContractService->GetArray($year, $commit);
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
    foreach($moCollection as $mo) {
        $ordinalRowNum++;
        $rowIndex++;
        $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);

        $bedProfiles = $content['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? [];
        $v = 0;
        foreach ($bedProfiles as $bp) {
            $v += $bp[$indicatorId];
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
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);

        $careProfiles = $content['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? [];
        $v = 0;
        foreach ($careProfiles as $vmpGroups) {
            foreach ($vmpGroups as $vmpTypes) {
                foreach ($vmpTypes as $vmpT)
                $v += $vmpT[$indicatorId];
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
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);

        $v = 0;
        $bedProfiles = $content['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? [];
        foreach ($bedProfiles as $bp) {
            $v += $bp[$indicatorId];
        }

        $bedProfiles = $content['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? [];
        foreach ($bedProfiles as $bp) {
            $v += $bp[$indicatorId];
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


Route::get('/xlsx2', function (DataForContractService $dataForContractService, PeopleAssignedInfoForContractService $peopleAssignedInfoForContractService) {
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

    $year = 2022;
    $commit = null;
    $content = $dataForContractService->GetArray($year, $commit);
    $peopleAssigned = $peopleAssignedInfoForContractService->GetArray($year, $commit);
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
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);

        $v = '0';
        $bedProfiles = $content['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? [];
        foreach ($bedProfiles as $bp) {
            $v = bcadd($v, $bp[$indicatorId]);
        }

        $bedProfiles = $content['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? [];
        foreach ($bedProfiles as $bp) {
            $v = bcadd($v, $bp[$indicatorId]);
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
        $sheet->setCellValue([2,$rowIndex], $mo->short_name);

        $bedProfiles = $content['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? [];
        $v = '0';
        foreach ($bedProfiles as $bpId => $bp) {
            if($bpId === $bedProfileId) {
                continue;
            }
            $v = bcadd($v, $bp[$indicatorId]);
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


Route::get('/', function (DataForContractService $dataForContractService, MoInfoForContractService $moInfoForContractService, MoDepartmentsInfoForContractService $moDepartmentsInfoForContractService, PeopleAssignedInfoForContractService $peopleAssignedInfoForContractService) {
    // InitialChanges::dispatch(2022);
    // InitialDataLoaded::dispatch(2, 2022, 1);
    // InitialDataLoaded::dispatch(9, 2022, 1);

    $year = 2022;
    $commit = null;
    $strDateTimeNow = date("Y-m-d-His");

    $indicators = Indicator::select('id','name')->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'indicators.json', $indicators);

    $medicalServices = MedicalServices::select('id','name')->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'medicalServices.json', $medicalServices);

    $medicalAssistanceType = MedicalAssistanceType::select('id','name')->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'medicalAssistanceType.json', $medicalAssistanceType);

    $hospitalBedProfiles = HospitalBedProfiles::select('tbl_hospital_bed_profiles.id','name', 'care_profile_foms_id')
    ->join('tbl_hospital_bed_profile_care_profile_foms',  'tbl_hospital_bed_profiles.id', '=', 'hospital_bed_profile_id')
    ->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'hospitalBedProfiles.json', $hospitalBedProfiles);

    $vmpGroup = VmpGroup::select('id','code')->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'vmpGroup.json', $vmpGroup);

    $vmpTypes = VmpTypes::select('id','name')->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'vmpTypes.json', $vmpTypes);

    $careProfiles = CareProfiles::select('tbl_care_profiles.id','name','care_profile_foms_id')
    ->join('tbl_care_profile_care_profile_foms',  'tbl_care_profiles.id', '=', 'care_profile_id')
    ->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'careProfiles.json', $careProfiles);

    $careProfilesFoms = CareProfilesFoms::select('id', 'name', 'code_v002 as code')
    ->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'careProfilesFoms.json', $careProfilesFoms);

    $mo = $moInfoForContractService->GetJson();
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'mo.json', $mo);

    $moDepartments = $moDepartmentsInfoForContractService->GetJson();
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'moDepartments.json', $moDepartments);

    $content = $dataForContractService->GetJson($year, $commit);
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'data.json', $content);

    $content = $peopleAssignedInfoForContractService->GetJson($year, $commit);
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'peopleAssignedData.json', $content);
    return $content;
});

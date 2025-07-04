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

use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;

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

function printTree($tree, PumpMonitoringProfilesTreeService $treeService, int $l = 0) {
    if ($tree === null) return;
    $marginLeft = ($l * 2) . '0px';
    foreach ($tree as $k => $t) {
        $mp = PumpMonitoringProfiles::find($k);
        $pu = $mp->profilesUnits;
        echo "<p><b><nobr style='margin-left:{$marginLeft};word-wrap: normal;'>{$mp->name} [" . PumpMonitoringProfilesRelationType::find($mp->relation_type_id)?->name . "]</nobr></b></p>\r\n";

        foreach($pu as $u) {
            $pi = $u->plannedIndicators;
            $piIdsViaChild = $treeService->plannedIndicatorIdsViaChild($mp->id, $u->unit->id);
            $piIds = array_column($pi?->ToArray(), 'id');
            $piIdsViaChild = array_diff($piIdsViaChild, $piIds);
            $piViaChild = PlannedIndicator::whereIn('id', $piIdsViaChild)->get();

            if (count($pi) > 0 || count($piIdsViaChild) > 0) {
                foreach($pi as $i) {
                    $piName = plannedIndicatorName($i);
                    echo "<p><nobr style='margin-left:{$marginLeft};word-wrap: normal;'>-{$u->unit->name} |{$piName}|</nobr></p>\r\n";
                }
                foreach($piViaChild as $i) {
                    $piName = plannedIndicatorName($i);
                    echo "<p><i><nobr style='margin-left:{$marginLeft};word-wrap: normal;'>-{$u->unit->name} |{$piName}| (УНАСЛЕДОВАНО)</nobr></i></p>\r\n";
                }
            } else {
                echo "<p><nobr style='margin-left:{$marginLeft};word-wrap: normal;'>-{$u->unit->name} | не утверждается | </nobr></p>\r\n";
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
    echo "<div style='word-wrap: normal;'>";
    printTree($tree, $treeService);
    echo "</div>";
    // dd($tree);
    return 'ok';
});


Route::get('/pump-plan-v2', function (PumpMonitoringProfilesTreeService $treeService) {
    $tree = $treeService->nodeTree(1);
   // dd($tree );
    return view('pump-plan', compact('tree', 'treeService'));
});


/**
 * Выгрузить связь "профилей мониторинга" ПУМП и наших плановых показателей
 *
 */
Route::get('/get-pump-monitoring-profiles-planned-indicators-relationships', function () {
    $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader("Xlsx");
    $path = 'xlsx/pump';
    $templateFileName = 'PumpPgg2025';
    $templateFileNameExt = '.xlsx';
    $templateFilePath = $path . DIRECTORY_SEPARATOR . $templateFileName . $templateFileNameExt;
    $templateFullFilepath = Storage::path($templateFilePath);

    $spreadsheet = $reader->load($templateFullFilepath);
    $sheet = $spreadsheet->getActiveSheet();
    $startRow = 8;
    $endRow = 628;

    $monitoringProfileCodeCol = 2;
    $monitoringProfileNameCol = 1;
    $plannedIndicatorIdCol = 3;

    $typeFinId = IndicatorType::where('name', 'money')->first()->id;
    $typeQuantId = IndicatorType::where('name', 'volume')->first()->id;



    $num = 1;
    for ($i = $startRow; $i <= $endRow; $i++) {
        $name = trim($sheet->getCell([$monitoringProfileNameCol, $i])->getValue());
        $code = trim($sheet->getCell([$monitoringProfileCodeCol, $i])->getValue());

        $monitoringProfileTypeId = null;
        // определить тип профиля мониторинга
        if (str_ends_with($name, '(руб.)')) {
            $monitoringProfileTypeId = $typeFinId;
        } else if (str_ends_with($name, '(кол-во)')) {
            $monitoringProfileTypeId = $typeQuantId;
        }
        if ($monitoringProfileTypeId === null) {
            throw new Exception("Не определен тип показателя $code");
        }

        $monitoringProfile = PumpMonitoringProfiles::where('code', $code)->first();
        $monitoringProfileUnits = $monitoringProfile->profilesUnits;
        if ($monitoringProfileUnits->count() > 2) {
            throw new Exception("$code содержит больше 2 'частей'");
        }
        $plannedIndicatorIds = collect([]);
        $plannedIndicatorNames = collect([]);
        foreach ($monitoringProfileUnits as $mpu) {
            if ($mpu->unit->type_id != $monitoringProfileTypeId) {
                continue;
            }

            $pi0 = $mpu->plannedIndicators;

            // Удалить показатели не соответствующие "профилю мониторинга" ПУМП по типу.
            $pi = $pi0->filter( function (PlannedIndicator $value, int $key) use ($monitoringProfileTypeId) {
                // dd($value->indicator);
                return $value->indicator->type_id === $monitoringProfileTypeId;
            });

            $plannedIndicatorIds = $plannedIndicatorIds->concat($pi->pluck('id'));
            $plannedIndicatorNames = $plannedIndicatorNames->concat($pi->map(function ($item, int $key) {
                return plannedIndicatorName($item);
            }));
        }
        $sheet->setCellValue([$plannedIndicatorIdCol, $i], $plannedIndicatorIds->join(','));
        $sheet->setCellValue([$plannedIndicatorIdCol + 1, $i], $plannedIndicatorNames->join(','  . PHP_EOL));
    }


    $resultFileName = $templateFileName . '' . $templateFileNameExt;
    $strDateTimeNow = date("Y-m-d-His");
    $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . ' ' . $resultFileName;
    $fullResultFilepath = Storage::path($resultFilePath);


    $writer = new Xlsx($spreadsheet);
    $writer->save($fullResultFilepath);
    return Storage::download($resultFilePath);
});


/**
 * Заполнить из файла связь профилей мониторинга ПУМП и наших плановых показателей
 */
Route::get('/fill-pump-monitoring-profiles-planned-indicators-relationships', function () {
    $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader("Xlsx");
    $path = 'xlsx/pump';
    $templateFileName = 'PumpPgg2025_IN.xlsx';
    $templateFilePath = $path . DIRECTORY_SEPARATOR . $templateFileName;
    $templateFullFilepath = Storage::path($templateFilePath);

    $spreadsheet = $reader->load($templateFullFilepath);
    $sheet = $spreadsheet->getActiveSheet();
    $startRow = 8;
    $endRow = 628;

    $monitoringProfileCodeCol = 2;
    $monitoringProfileNameCol = 1;
    $plannedIndicatorIdCol = 3;

    $typeFinId = IndicatorType::where('name', 'money')->first()->id;
    $typeQuantId = IndicatorType::where('name', 'volume')->first()->id;

    $num = 1;
    for ($i=$startRow; $i <= $endRow; $i++) {
        $name = trim($sheet->getCell([$monitoringProfileNameCol, $i])->getValue());
        $code = trim($sheet->getCell([$monitoringProfileCodeCol, $i])->getValue());
        $indicators = [];

        // id показателей перечислены в нескольких столбцах
        // внутри каждого столбца значения могут быть перечислены через запятую
        $d = 0;
        do {
            $indicatorsRead = trim($sheet->getCell([$plannedIndicatorIdCol + $d, $i])->getValue());
            $indicatorsTemp = explode(',', $indicatorsRead);
            for ($k = 0; $k < count($indicatorsTemp); $k++) {
                $ind = trim($indicatorsTemp[$k]);
                if ($ind !== '' && $ind !== 'нет данных' && $ind !== 'что это?' && $ind !== 'не утверждается' && $ind !== 'не нашла') {
                    // проверяем что показатель действующий
                    $date = date('Y-m-d');
                    $pi = PlannedIndicator::where('id', $ind)->whereRaw("? BETWEEN effective_from AND effective_to", [$date])->get();
                    if ($pi->count() === 1) {
                        array_push($indicators, $ind);
                    } else {
                        echo "Плановый показатель $ind отсутвует на $date <br>";
                    }
                }
            }
            $d++;
        } while ($indicatorsRead !== '');


        $indicators = array_unique($indicators, SORT_NUMERIC);
        $monitoringProfile = PumpMonitoringProfiles::where('code', $code)->first();
        $t = str_starts_with($name, $monitoringProfile->name);

        $p = 'ERROR';
        $monitoringProfileUnits = null;
        // if (str_ends_with($name, '(сумма)')) {
        if (str_ends_with($name, '(руб.)')) {
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
        if(!$monitoringProfileUnits) {
            throw new Exception(" $name");
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
 * Коды "профилей мониторинга" 2025 изменены
 * пробуем сопоставить показатели с новым кодом по полному имени
 */
Route::get('/pump-monitoring-profiles-update-codes', function () {
    $FIN_PROFILE_TYPE_STR = "только финансовая часть";
    $FIN_QUANT_PROFILE_TYPE_STR = "финансовая и количественная часть";

    $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader("Xlsx");
    $path = 'xlsx/pump';
    $templateFileName = 'PumpMonitoringProfiles_v6.xlsx';
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
    // $monitoringProfileNestingLevel = 8;
    $parentRelationCol = 9;
    $monitoringProfileTypeCol = 10;
    $unitCol = 11;


    // TODO: Получить список всех программ ОМС описанных в файле
    $omsProgramIds =
    [
        OmsProgram::where('name', 'базовая')->first()->id,
        OmsProgram::where('name', 'сверхбазовая')->first()->id
    ];
    $dtNow = \Carbon\Carbon::now();
    PumpMonitoringProfiles
        ::whereIn('oms_program_id', $omsProgramIds)
        ->where('effective_to', '>', $dtNow)
        ->update(['effective_to' => $dtNow]);

    $iterator = $sheet->getRowIterator();

    // TODO: Проверить заголовки

    $iterator->next();
    while ($iterator->valid()) {
        $row = $iterator->current();
        // ПрограммаОМС ИД ИДродителя Код КодРодителя КраткоеНаименование ПолноеНаименование УровеньВложенности ОтношениеКРодителю СоставПоказателя КоличественнаяЕдиница
        $p = new PumpMonitoringProfiles();
        $omsProgramName = trim(mb_strtolower($sheet->getCell([$omsProgramCol, $row->getRowIndex()])->getValue()));
        if ($omsProgramName === '') {
            break;
        }
        $omsProgram = OmsProgram::where('name', $omsProgramName)->first();
        if ($omsProgram === null) {
            throw("Программа ОМС '$omsProgramName' отсутствует в базе");
        }
        $p->oms_program_id = $omsProgram->id;
        $p->code = trim($sheet->getCell([$monitoringProfileCodeCol, $row->getRowIndex()])->getValue());
        $parentCode = trim($sheet->getCell([$monitoringProfileParentCodeCol, $row->getRowIndex()])->getValue());

        if ($parentCode != '') {
            $parent = PumpMonitoringProfiles::where('code', $parentCode)->first();
            if ($parent !== null) {
                $p->parent_id = $parent->id;
            }
        }
        $p->short_name = trim($sheet->getCell([$monitoringProfileShortNameCol, $row->getRowIndex()])->getValue());
        $p->name = trim($sheet->getCell([$monitoringProfileNameCol, $row->getRowIndex()])->getCalculatedValue());
        $rT =  mb_strtolower($sheet->getCell([$parentRelationCol, $row->getRowIndex()])->getValue());
        if ($rT != '') {
            $relationType = PumpMonitoringProfilesRelationType::where('name', $rT)->first();
            if ($relationType === null) {
                throw("Неизвестный тип отношения к родителю: $rT");
            }
            $p->relation_type_id = $relationType->id;
        }
        $profileType = trim(mb_strtolower($sheet->getCell([$monitoringProfileTypeCol, $row->getRowIndex()])->getValue()));
        $p->is_leaf = false;


        // Ищем в базе показатель с таким полныйм именем
        $pOld = PumpMonitoringProfiles::where('name', $p->name)->first();
        if ($pOld !== null) {
            if ($pOld->oms_program_id !== $p->oms_program_id
                // || $pOld->parentCode !== $p->parentCode
                //|| $pOld->parent_id !== $p->parent_id
                || $pOld->short_name !== $p->short_name
                || $pOld->name !== $p->name
                || $pOld->relation_type_id !== $p->relation_type_id
            ) {
                throw new Exception("Профиль мониторинга с кодом $p->code существует и имеет значения отличные от полученных");
            } else {
                if ($pOld->code !== $p->code) {
                    $byCode = PumpMonitoringProfiles::where('code', $p->code)->first();
                    if ($byCode !== null) {
                        $byCode->code = 'old_' . $byCode->code;
                        $byCode->save();
                    }
                    $pOld->code = $p->code;
                    $pOld->parent_id = $p->parent_id ? $p->parent_id : null;
                    $pOld->save();
                }
            }
        }

        $iterator->next();
    }

    return 'ок';
});

/**
 * Заполнить профили мониторинга ПУМП из файла (Обновить на новую версию)
 */
Route::get('/fill-pump-monitoring-profiles', function () {
    $FIN_PROFILE_TYPE_STR = "только финансовая часть";
    $FIN_QUANT_PROFILE_TYPE_STR = "финансовая и количественная часть";

    /*
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
    */

    $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader("Xlsx");
    $path = 'xlsx/pump';
    $templateFileName = 'PumpMonitoringProfiles_v6.xlsx';
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
    // $monitoringProfileNestingLevel = 8;
    $parentRelationCol = 9;
    $monitoringProfileTypeCol = 10;
    $unitCol = 11;


    // TODO: Получить список всех программ ОМС описанных в файле
    $omsProgramIds =
    [
        OmsProgram::where('name', 'базовая')->first()->id,
        OmsProgram::where('name', 'сверхбазовая')->first()->id
    ];
    $dtNow = \Carbon\Carbon::now();
    PumpMonitoringProfiles
        ::whereIn('oms_program_id', $omsProgramIds)
        ->where('effective_to', '>', $dtNow)
        ->update(['effective_to' => $dtNow]);

    // Все единицы измерения количественной части показателя
    $quantTypeId = IndicatorType::where('name', 'volume')->first()->id;
    $quantUnits = PumpUnit::where('type_id', $quantTypeId)->pluck('id')->toArray();

    $iterator = $sheet->getRowIterator();

    // TODO: Проверить заголовки

    $iterator->next();
    while ($iterator->valid()) {
        $row = $iterator->current();
        // ПрограммаОМС ИД ИДродителя Код КодРодителя КраткоеНаименование ПолноеНаименование УровеньВложенности ОтношениеКРодителю СоставПоказателя КоличественнаяЕдиница
        $p = new PumpMonitoringProfiles();
        $omsProgramName = trim(mb_strtolower($sheet->getCell([$omsProgramCol, $row->getRowIndex()])->getValue()));
        if ($omsProgramName === '') {
            break;
        }
        $omsProgram = OmsProgram::where('name', $omsProgramName)->first();
        if ($omsProgram === null) {
            throw("Программа ОМС '$omsProgramName' отсутствует в базе");
        }
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
            $relationType = PumpMonitoringProfilesRelationType::where('name', $rT)->first();
            if ($relationType === null) {
                throw("Неизвестный тип отношения к родителю: $rT");
            }
            $p->relation_type_id = $relationType->id;
        }
        $profileType = trim(mb_strtolower($sheet->getCell([$monitoringProfileTypeCol, $row->getRowIndex()])->getValue()));
        $p->is_leaf = false;



        // Ищем в базе показатель с таким кодом
        $pOld = PumpMonitoringProfiles::where('code', $p->code)->first();
        if ($pOld !== null) {
            if ($pOld->oms_program_id !== $p->oms_program_id
                || $pOld->parentCode !== $p->parentCode
                || $pOld->parent_id !== $p->parent_id
                || $pOld->short_name !== $p->short_name
                || $pOld->name !== $p->name
                || $pOld->relation_type_id !== $p->relation_type_id
            ) {
                throw new Exception("Профиль мониторинга с кодом $p->code существует и имеет значения отличные от полученных");
            } else {
                // TODO
                // Что делать если показатель вновь появился после перерыва (не использовался какой-то период)
                // менять ли effective_from ???

                $pOld->effective_to = \Carbon\Carbon::create(9999, 12, 31, 23, 59, 59);
                $pOld->save();

                // Проверяем единицы измерения колличественной части показателя
                if ($profileType === $FIN_QUANT_PROFILE_TYPE_STR) {
                    $unitName = trim(mb_strtolower($sheet->getCell([$unitCol, $row->getRowIndex()])->getValue()));
                    $unitId = PumpUnit::where('name', $unitName)->first()->id;
                    $mpu = PumpMonitoringProfilesUnit::where('monitoring_profile_id', $pOld->id)
                        ->whereIn('unit_id', $quantUnits)
                        ->pluck('unit_id')
                        ->toArray();

                    if (count($mpu) > 1) {
                        throw new Exception("Профиль мониторинга ($pOld->id) содержит несколько количественных показателей");
                    }
                    if (count($mpu) === 1) {
                        if ($mpu[0] !== $unitId) {
                            throw new Exception("Профиль мониторинга ($pOld->id). Количественный показатель отличается {$mpu[0]} !== $unitId");
                        }
                    } else {
                        throw new Exception("Профиль мониторинга ($pOld->id). Необходимый количественный показатель отсутствует.");
                    }
                }
            }
        } else {
            $p->save();

            $unit = new PumpMonitoringProfilesUnit();
            $unit->unit_id = PumpUnit::where('name', 'стоимость')->first()->id;
            $unit->monitoring_profile_id = $p->id;
            $unit->save();

            if ($profileType === $FIN_QUANT_PROFILE_TYPE_STR) {
                $unitName = trim(mb_strtolower($sheet->getCell([$unitCol, $row->getRowIndex()])->getValue()));
                $unit2 = new PumpMonitoringProfilesUnit();
                $unit2->unit_id = PumpUnit::where('name', $unitName)->first()->id;
                $unit2->monitoring_profile_id = $p->id;
                $unit2->save();
            }
        }

        $iterator->next();
    }

    return 'ок';
});

/**
 * Выделить ошибки выгрузки ПУМП  цветом
 *
 */
Route::get('/pump-monitoring-profiles-errors-highlight-with-color', function () {
    $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader("Xlsx");
    $path = 'xlsx/pump';
    $errorsFileName = 'PumpPggErrors.txt';
    $errorsFilePath = $path . DIRECTORY_SEPARATOR . $errorsFileName;
    $errorsFullFilepath = Storage::path($errorsFilePath);

    $templateFileName = 'PumpPgg_with_errors.xlsx';
    $templateFilePath = $path . DIRECTORY_SEPARATOR . $templateFileName;
    $templateFullFilepath = Storage::path($templateFilePath);

    $spreadsheet = $reader->load($templateFullFilepath);
    $sheet = $spreadsheet->getActiveSheet();

    $parsedErrors = parseErrors($errorsFullFilepath);
    highlightErrors($sheet, $parsedErrors);

    $resultFileName = $templateFileName;
    $strDateTimeNow = date("Y-m-d-His");
    $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . ' ' . $resultFileName;
    $fullResultFilepath = Storage::path($resultFilePath);


    $writer = new Xlsx($spreadsheet);
    $writer->save($fullResultFilepath);
    return Storage::download($resultFilePath);
});

/**
 * Считать ошибки из файла
 * Создано chatgpt )
 */
function parseErrors($filename) {
    $errors = [];
    $content = file_get_contents($filename);
    $blocks = preg_split("/\n\s*\n/", trim($content)); // Разделение блоков по пустой строке

    foreach ($blocks as $block) {
        preg_match('/Строка:\s*(\d+)/u', $block, $lineMatch);
        preg_match('/Столбец:\s*(\S+)/u', $block, $columnMatch);
        $lines = explode("\n", trim($block));
        $errorDescription = end($lines); // Последняя строка блока

        if ($lineMatch && $columnMatch && $errorDescription) {
            $errors[] = [
                'line' => $lineMatch[1],
                'column' => $columnMatch[1],
                'error' => trim($errorDescription)
            ];
        }
    }
    return $errors;
}
/**
 * Выделить ошибки цветом
 * Создано chatgpt
 */
function highlightErrors($sheet, $errors) {
    foreach ($errors as $error) {
        $cell = $error['column'] . $error['line'];
        $sheet->getStyle($cell)->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_YELLOW));
    }
}

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

Route::get('/meeting-minutes/{year}/{commissionDecisionsId}', [PlanReports::class, "MeetingMinutes"]);
Route::get('/decree-n17-vmp/{year}/{commissionDecisionsId?}', [PlanReports::class, 'DecreeN17Vmp']);

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

Route::get('/vitacore-v3/{year}/{commissionDecisionsId?}', [PlanReports::class, "VitacorePlan"]);

Route::get('/vitacore-v2/{year}/{commissionDecisionsId?}', function (DataForContractService $dataForContractService, PeopleAssignedInfoForContractService $peopleAssignedInfoForContractService, int $year, int|null $commissionDecisionsId = null) {
    $packageIds = null;
    $currentlyUsedDate = $year.'-01-01';
    $protocolNumber = 0;
    $protocolDate = '';
    if ($commissionDecisionsId) {
        $commissionDecisions = CommissionDecision::whereYear('date',$year)->where('id', '<=', $commissionDecisionsId)->get();
        $cd = $commissionDecisions->find($commissionDecisionsId);
        $commissionDecisionIds = $commissionDecisions->pluck('id')->toArray();
        $packageIds = ChangePackage::whereIn('commission_decision_id', $commissionDecisionIds)->orWhere('commission_decision_id', null)->pluck('id')->toArray();
        $protocolNumber = $cd->number;
        $protocolDate = $cd->date->format('d.m.Y');
        $currentlyUsedDate = $cd->date->format('Y-m-d');
    } else {
        $packageIds = ChangePackage::where('commission_decision_id', null)->pluck('id')->toArray();
    }
    $protocolNumberForFileName = preg_replace('/[^a-zа-я\d.]/ui', '_', $protocolNumber);

    $path = 'xlsx';
    $resultFileName = 'vitacore-plan' . ($protocolNumber !== 0 ? '(Protokol_№'.$protocolNumberForFileName.'ot'.$protocolDate.')' : '') . '.xlsx';
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
        $planningSectionName = "Молекулярно-генетические исследования с целью диагностики онкологических заболеваний";
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
        $planningSectionName = "Молекулярно-генетические исследования с целью диагностики онкологических заболеваний";
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


Route::get('/hospital-by-profile/{year}/{commissionDecisionsId?}', function (DataForContractService $dataForContractService, PlannedIndicatorChangeInitService $plannedIndicatorChangeInitService, InitialDataFixingService $initialDataFixingService, int $year, int|null $commissionDecisionsId = null) {
    $packageIds = null;
    $currentlyUsedDate = $year.'-01-01';
    $protocolNumber = 0;
    $protocolDate = '';
    if ($commissionDecisionsId) {
        $commissionDecisions = CommissionDecision::whereYear('date',$year)->where('id', '<=', $commissionDecisionsId)->get();
        $cd = $commissionDecisions->find($commissionDecisionsId);
        $commissionDecisionIds = $commissionDecisions->pluck('id')->toArray();
        $packageIds = ChangePackage::whereIn('commission_decision_id', $commissionDecisionIds)->orWhere('commission_decision_id', null)->pluck('id')->toArray();
        $protocolNumber = $cd->number;
        $protocolDate = $cd->date->format('d.m.Y');
        $currentlyUsedDate = $cd->date->format('Y-m-d');
    } else {
        if ($initialDataFixingService->fixedYear($year)) {
            $packageIds = ChangePackage::where('commission_decision_id', null)->pluck('id')->toArray();
        } else {
            $plannedIndicatorChangeInitService->fromInitialData($year);
        }
    }
    $protocolNumberForFileName = preg_replace('/[^a-zа-я\d.]/ui', '_', $protocolNumber);
    $path = 'xlsx';
    $resultFileName = 'hospital_' . ($protocolNumber !== 0 ? '(Protokol_№'.$protocolNumberForFileName.'ot'.$protocolDate.')' : '') . '.xlsx';
    $strDateTimeNow = date("Y-m-d-His");
    $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . '_' . $resultFileName;
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

Route::get('/vitacore-hospital-by-profile/{year}/{commissionDecisionsId?}', function (DataForContractService $dataForContractService, int $year, int|null $commissionDecisionsId = null) {
    $packageIds = null;
    $currentlyUsedDate = $year.'-01-01';
    $protocolNumber = 0;
    $protocolDate = '';
    if ($commissionDecisionsId) {
        $commissionDecisions = CommissionDecision::whereYear('date',$year)->where('id', '<=', $commissionDecisionsId)->get();
        $cd = $commissionDecisions->find($commissionDecisionsId);
        $commissionDecisionIds = $commissionDecisions->pluck('id')->toArray();
        $packageIds = ChangePackage::whereIn('commission_decision_id', $commissionDecisionIds)->orWhere('commission_decision_id', null)->pluck('id')->toArray();
        $protocolNumber = $cd->number;
        $protocolDate = $cd->date->format('d.m.Y');
        $currentlyUsedDate = $cd->date->format('Y-m-d');
    } else {
        $packageIds = ChangePackage::where('commission_decision_id', null)->pluck('id')->toArray();
    }
    $protocolNumberForFileName = preg_replace('/[^a-zа-я\d.]/ui', '_', $protocolNumber);
    $path = 'xlsx';
    $resultFileName = 'hospital-medical-care-profiles' . ($protocolNumber !== 0 ? '(Protokol_№'.$protocolNumberForFileName.'ot'.$protocolDate.')' : '') . '.xlsx';
    $strDateTimeNow = date("Y-m-d-His");
    $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . '_' . $resultFileName;
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


Route::get('/vitacore-hospital-by-profile-periods/{year}/{commissionDecisionsId?}', function (DataForContractService $dataForContractService, int $year, int|null $commissionDecisionsId = null) {
    $packageIds = null;
    $currentlyUsedDate = $year.'-01-01';
    $protocolNumber = 0;
    $protocolDate = '';
    if ($commissionDecisionsId) {
        $commissionDecisions = CommissionDecision::whereYear('date',$year)->where('id', '<=', $commissionDecisionsId)->get();
        $cd = $commissionDecisions->find($commissionDecisionsId);
        $commissionDecisionIds = $commissionDecisions->pluck('id')->toArray();
        $packageIds = ChangePackage::whereIn('commission_decision_id', $commissionDecisionIds)->orWhere('commission_decision_id', null)->pluck('id')->toArray();
        $protocolNumber = $cd->number;
        $protocolDate = $cd->date->format('d.m.Y');
        $currentlyUsedDate = $cd->date->format('Y-m-d');
    } else {
        $packageIds = ChangePackage::where('commission_decision_id', null)->pluck('id')->toArray();
    }
    $protocolNumberForFileName = preg_replace('/[^a-zа-я\d.]/ui', '_', $protocolNumber);
    $path = 'xlsx';
    $resultFileName = 'hospital-medical-care-profiles-by-period' . ($protocolNumber !== 0 ? '(Protokol_№'.$protocolNumberForFileName.'ot'.$protocolDate.')' : '') . '.xlsx';
    $strDateTimeNow = date("Y-m-d-His");
    $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . '_' . $resultFileName;
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

Route::get('/vitacore-hospital-by-bed-profile-periods/{year}/{commissionDecisionsId?}', function (DataForContractService $dataForContractService, int $year, int|null $commissionDecisionsId = null) {
    $packageIds = null;
    $currentlyUsedDate = $year.'-01-01';
    $protocolNumber = 0;
    $protocolDate = '';
    if ($commissionDecisionsId) {
        $commissionDecisions = CommissionDecision::whereYear('date',$year)->where('id', '<=', $commissionDecisionsId)->get();
        $cd = $commissionDecisions->find($commissionDecisionsId);
        $commissionDecisionIds = $commissionDecisions->pluck('id')->toArray();
        $packageIds = ChangePackage::whereIn('commission_decision_id', $commissionDecisionIds)->orWhere('commission_decision_id', null)->pluck('id')->toArray();
        $protocolNumber = $cd->number;
        $protocolDate = $cd->date->format('d.m.Y');
        $currentlyUsedDate = $cd->date->format('Y-m-d');
    } else {
        $packageIds = ChangePackage::where('commission_decision_id', null)->pluck('id')->toArray();
    }
    $protocolNumberForFileName = preg_replace('/[^a-zа-я\d.]/ui', '_', $protocolNumber);
    $path = 'xlsx';
    $resultFileName = 'hospital-bed-profiles-by-period' . ($protocolNumber !== 0 ? '(Protokol_№'.$protocolNumberForFileName.'ot'.$protocolDate.')' : '') . '.xlsx';
    $strDateTimeNow = date("Y-m-d-His");
    $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . '_' . $resultFileName;
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
                        $hbpId = vmpGetBedProfileId($cpId, $mo->code, $vmpGroup, $year);

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

function getLevel(int $year, int $monthNum, int $ts, int $moId, int $bedProfileId) : string
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

        if ($year === 2025) {
            switch ($moId) {
                case 17	/* ГБУ "Далматовская ЦРБ" */:
                case 20	/* ГБУ "Катайская ЦРБ" */:
                case 25	/* ГБУ "Шадринская ЦРБ" */:
                case 45	/* ООО "ЛДК "Центр ДНК" */:
                    $lResult = $l1;
                    break;


                case 91	/* ГБУ «Межрайонная больница №3» */:
                case 92	/* ГБУ «Межрайонная больница №4» */:
                case 93	/* ГБУ «Межрайонная больница №5» */:
                case 94	/* ГБУ «Межрайонная больница №6» */:
                case 95	/* ГБУ «Межрайонная больница №7» */:
                case 96	/* ГБУ «Межрайонная больница №8» */:
                case 13	/* ГБУ «КОКВД» */:
                case 38	/* ЧУЗ "РЖД-Медицина" г. Курган" */:
                case 7	/* ГБУ "Курганская областная специализированная инфекционная больница" */:
                case 51	/* ГБУ "Санаторий "Озеро Горькое" */:
                    $lResult = $l2_1;
                    break;

                case 89	/* ГБУ «Межрайонная больница №1» */:
                case 90	/* ГБУ «Межрайонная больница №2» */:

                    $lResult = $l2_2;
                    break;

                case 98	/* ГБУ "ШГБ" */:
                case 97	/* ГБУ «Курганская областная больница №2» */ :
                    $lResult = $l3_1;
                    break;

                case 42	/* ФГБУ «НМИЦ ТО имени академика Г.А.Илизарова» Минздрава России */:
                case 1	/* ГБУ "КОКБ" */:
                case 11	/* ГБУ "Курганский областной кардиологический диспансер" */:
                case 12	/* ГБУ «КОДКБ им. Красного Креста» */:
                case 3	/* ГБУ "Курганская БСМП" */:
                    $lResult = $l3_2;
                    break;

                case 2	/* ГБУ "КООД" */:
                case 67	/* ГБУ "КОГВВ" */:
                case 40	/* ГБУ "Перинатальный центр" */:
                    $lResult = $l3_3;
                    break;
            }
        }


        if ($year === 2024) {
            switch ($moId) {
                case 17	/* ГБУ "Далматовская ЦРБ" */:
                case 20	/* ГБУ "Катайская ЦРБ" */:
                case 25	/* ГБУ "Шадринская ЦРБ" */:
                case 45	/* ООО "ЛДК "Центр ДНК" */:
                    $lResult = $l1;
                    break;


                case 91	/* ГБУ «Межрайонная больница №3» */:
                case 92	/* ГБУ «Межрайонная больница №4» */:
                case 93	/* ГБУ «Межрайонная больница №5» */:
                case 94	/* ГБУ «Межрайонная больница №6» */:
                case 95	/* ГБУ «Межрайонная больница №7» */:
                case 96	/* ГБУ «Межрайонная больница №8» */:
                case 13	/* ГБУ «КОКВД» */:
                case 38	/* ЧУЗ "РЖД-Медицина" г. Курган" */:
                case 7	/* ГБУ "Курганская областная специализированная инфекционная больница" */:
                case 51	/* ГБУ "Санаторий "Озеро Горькое" */:
                    $lResult = $l2_1;
                    break;

                case 89	/* ГБУ «Межрайонная больница №1» */:
                case 90	/* ГБУ «Межрайонная больница №2» */:
                case 98	/* ГБУ "ШГБ" */:
                    $lResult = $l2_2;
                    break;

                case 97	/* ГБУ «Курганская областная больница №2» */ :
                case 42	/* ФГБУ «НМИЦ ТО имени академика Г.А.Илизарова» Минздрава России */:
                    $lResult = $l3_1;
                    break;

                case 1	/* ГБУ "КОКБ" */:
                case 11	/* ГБУ "Курганский областной кардиологический диспансер" */:
                case 12	/* ГБУ «КОДКБ им. Красного Креста» */:
                case 3	/* ГБУ "Курганская БСМП" */:
                    $lResult = $l3_2;
                    break;

                case 2	/* ГБУ "КООД" */:
                case 67	/* ГБУ "КОГВВ" */:
                case 40	/* ГБУ "Перинатальный центр" */:
                    $lResult = $l3_3;
                    break;
            }
        } else {
            if ($year === 2023) {
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
        }
    }
    return $lResult;
}

function vmpGetBedProfileId(int $careProfileId, string $moCode, int $vmpGroup, int $year)
{
    // КООД
    if ($moCode === '450004') {
        if ($year < 2024) {
            if ($vmpGroup === 23 || $vmpGroup === 24) {
                return 36; // Радиологические (V020 код: 64);
            }
        }
        if ($year === 2024) {
            if ($vmpGroup === 23 || $vmpGroup === 24 || $vmpGroup === 25 || $vmpGroup === 26) {
                return 36; // Радиологические (V020 код: 64);
            }
        }
        if ($year > 2024) {
            if ($vmpGroup === 25 || $vmpGroup === 26 || $vmpGroup === 27) {
                return 36; // Радиологические (V020 код: 64);
            }
        }
    }
    /**/
    // Красный крест
    if ($moCode === '450002') {
        if ($year > 2024) {
            if ($vmpGroup === 79) {
                return 40; // уроандрологические для детей (V020 код: 21);
            }
        }
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

Route::get('/miac-hospital-by-bed-profile-periods/{year}/{commissionDecisionsId?}', function (DataForContractService $dataForContractService, int $year, int|null $commissionDecisionsId = null) {
    $packageIds = null;
    $currentlyUsedDate = $year.'-01-01';
    $protocolNumber = 0;
    $protocolDate = '';
    if ($commissionDecisionsId) {
        $commissionDecisions = CommissionDecision::whereYear('date',$year)->where('id', '<=', $commissionDecisionsId)->get();
        $cd = $commissionDecisions->find($commissionDecisionsId);
        $commissionDecisionIds = $commissionDecisions->pluck('id')->toArray();
        $packageIds = ChangePackage::whereIn('commission_decision_id', $commissionDecisionIds)->orWhere('commission_decision_id', null)->pluck('id')->toArray();
        $protocolNumber = $cd->number;
        $currentlyUsedDate = $cd->date->format('Y-m-d');
        $protocolDate = $cd->date->format('d.m.Y');
    } else {
        $packageIds = ChangePackage::where('commission_decision_id', null)->pluck('id')->toArray();
    }
    $protocolNumberForFileName = preg_replace('/[^a-zа-я\d.]/ui', '_', $protocolNumber);
    $path = 'xlsx';
    $resultFileName = 'Plan_ProfileBed' . ($protocolNumber !== 0 ? '(Protokol_№'.$protocolNumberForFileName.'ot'.$protocolDate.')' : '') . '.csv';
    $strDateTimeNow = date("Y-m-d-His");
    $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . '_' .  $resultFileName;
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
                        $hbpId = vmpGetBedProfileId($cpId, $mo->code, $vmpGroup, $year);
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
                                $oplat, $ts, $hbp->code, $monthNum, getLevel($year, $monthNum, $ts, $mo->id, $hbp->id),
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

Route::get('/{year}/{commissionDecisionsId?}', function (Request $request, DataForContractService $dataForContractService, MoInfoForContractService $moInfoForContractService, MoDepartmentsInfoForContractService $moDepartmentsInfoForContractService, PeopleAssignedInfoForContractService $peopleAssignedInfoForContractService, int $year, int|null $commissionDecisionsId = null) {
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
    $protocolNumberForFileName = preg_replace('/[^a-zа-я\d.]/ui', '_', $protocolNumber);

    $moIds = null;
    if ($commissionDecisionsId){
        if ($onlyMoModifiedByCommission) {
            $c = CommissionDecision::find($commissionDecisionsId);
            $pIds = $c->changePackage()->pluck('id')->toArray();
            $moIds = PlannedIndicatorChange::select('mo_id')->where('package_id', $pIds)->groupBy('mo_id')->get()->pluck('mo_id')->toArray();
        }
    }

    $strDateTimeNow = date("Y-m-d-His");
    $path =  $year.DIRECTORY_SEPARATOR.$protocolNumberForFileName.'('.$protocolDate.')'.DIRECTORY_SEPARATOR.$strDateTimeNow.DIRECTORY_SEPARATOR;

    $indicators = Indicator::select('id','name')->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put($path.'indicators.json', $indicators);

    $medicalServices = MedicalServices::select('id','name','slug','allocateVolumes','order')->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
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
    $zipFileName = $path.$strDateTimeNow.'_Protokol_'.$protocolNumberForFileName.'('.$protocolDate.').dogoms';
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

<?php
declare(strict_types=1);

namespace App\Services\PlanReports;

use App\Models\ChangePackage;
use App\Models\CommissionDecision;
use App\Models\IndicatorType;
use App\Models\MedicalInstitution;
use App\Models\PumpMonitoringProfiles;
use App\Models\PumpMonitoringProfilesUnit;
use App\Services\DataForContractService;
use App\Services\InitialDataFixingService;
use App\Services\PlannedIndicatorChangeInitService;
use App\Services\PumpMonitoringProfilesTreeService;
use Exception;
use Illuminate\Database\Eloquent\Builder;
use Illuminate\Support\Collection;
use Illuminate\Support\Facades\DB;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class PumpPggReportService
{
    use SummaryReportTrait;
    use HospitalSheetTrait;

    public function __construct(
        private DataForContractService $dataForContractService,
        private PlannedIndicatorChangeInitService $plannedIndicatorChangeInitService,
        private InitialDataFixingService $initialDataFixingService,
        private PumpMonitoringProfilesTreeService $pumpMonitoringProfilesTreeService
    ) {}

    private function getColMap(Worksheet $sheet): array {
        $nameRow = 1;
        $codeRow = 2;
        $startCol = 1;
        $curCol = $startCol;
        $typeFinId = IndicatorType::where('name', 'money')->first()->id;
        $typeQuantId = IndicatorType::where('name', 'volume')->first()->id;
        $name = trim($sheet->getCell([$curCol, $nameRow])->getValue() ?? '');
        $code = trim($sheet->getCell([$curCol, $codeRow])->getValue() ?? '');
        $arr = [];
        while ($code != '' || $name != '') {
            if ($code === '') {
                switch ($name) {
                    case 'Субъект по ОКТМО':
                        $arr['subjectOktmo'] = $curCol;
                        break;
                    case 'Наименование субъекта':
                        $arr['subjectName'] = $curCol;
                        break;
                    case 'Ведомство':
                        $arr['institution'] = $curCol;
                        break;
                    case 'Наименование ведомства':
                        $arr['institutionName'] = $curCol;
                        break;
                    case 'Год':
                        $arr['year'] = $curCol;
                        break;
                    case 'Код МО':
                        $arr['moCode'] = $curCol;
                        break;
                    case 'Наименование МО':
                        $arr['moName'] = $curCol;
                        break;
                    default:
                        throw new Exception('Не заполнен код');
                        break;
                }
            } else {

                $monitoringProfile = PumpMonitoringProfiles::where('code', $code)->first();
                if ($monitoringProfile === null) {
                    throw new Exception('Отсутствует профиль мониторинга с кодом ' . $code);
                }
                if (!str_starts_with($name, $monitoringProfile->name)) {
                    throw new Exception("Код {$code} не соответствует названию профиля. В файле: {$name}; В базе: {$monitoringProfile->name}.");
                }

                $monitoringProfileUnits = null;
                if (str_ends_with($name, '(руб.)')) {
                    $monitoringProfileUnits = $monitoringProfile->profilesUnits()->whereHas('unit', function (Builder $query) use ($typeFinId) {
                            $query->where('type_id', $typeFinId);
                        })->get();
                } else if (str_ends_with($name, '(кол-во)')) {
                    $monitoringProfileUnits = $monitoringProfile->profilesUnits()->whereHas('unit', function (Builder $query) use ($typeQuantId) {
                        $query->where('type_id', $typeQuantId);
                    })->get();
                }
                if ($monitoringProfileUnits === null) {
                    throw new Exception('В базе отсутствует соответствующий профиль мониторинга');
                }
                if ($monitoringProfileUnits->count() > 1) {
                    throw new Exception($monitoringProfileUnits->count() . ' показателей профиля мониторинга соответствуют одному полю документа');
                }
                $arr[$monitoringProfileUnits[0]->id] = $curCol;
            }
            $curCol++;
            $name = trim($sheet->getCell([$curCol, $nameRow])->getValue() ?? '');
            $code = trim($sheet->getCell([$curCol, $codeRow])->getValue() ?? '');
        }
        return $arr;
    }

    private function monitoringProfilesUnitValuesGroupedByMo(array $monitoringProfilesUnitIds, array $contentGroupedByMoAndPlannedIndicator) : array {
        $moIds = array_keys($contentGroupedByMoAndPlannedIndicator['mo']);
        /*
        $tblPumpMonitoringProfilesUnitPlannedIndicators = (new PumpMonitoringProfilesUnit())->plannedIndicators()->getTable();
        $monitoringProfilesUnitPlannedIndicators = DB::table($tblPumpMonitoringProfilesUnitPlannedIndicators)
            ->select('monitoring_profile_unit_id','planned_indicator_id')
            ->get()
            ->groupBy(['monitoring_profile_unit_id'])
            ->toArray();
        */

        // Получаем плановые показатели соответствующие "профилю мониторинга"
        $monitoringProfilesUnitPlannedIndicators = [];
        foreach ($monitoringProfilesUnitIds as $mpuId) {
            if (!is_int($mpuId)) {
                continue;
            }
            $monitoringProfilesUnitPlannedIndicators[$mpuId] = $this->pumpMonitoringProfilesTreeService->plannedIndicatorIdsViaChildByMonitoringProfilesUnitId((int)$mpuId);
        }

        $contentByMonitoringProfilesUnit = [];

        bcscale(4);
        //dd($content);
        foreach ($moIds as $moId) {
            $moContent = $contentGroupedByMoAndPlannedIndicator['mo'][$moId];

            $contentByMonitoringProfilesUnit[$moId] = [];

            foreach ($monitoringProfilesUnitIds as $mpuId) {
                $piIds = $monitoringProfilesUnitPlannedIndicators[$mpuId] ?? null;
                if ($piIds !== null) {
                    $v = '0';
                    foreach ($piIds as $piId) {
                        if (isset($moContent[$piId])) {
                            foreach ($moContent[$piId] as $mcpi) { // для каждого ФАПа (при наличии)
                                $v = bcadd($v, ($mcpi->value ?? '0'));
                            }
                        }
                    }
                    $contentByMonitoringProfilesUnit[$moId][$mpuId] = $v;
                }
            }
        }
        return $contentByMonitoringProfilesUnit;
    }

    private function fillSheet(Worksheet $sheet, array $colMap, int $year, array $monitoringProfilesUnitValuesGroupedByMo, Collection $moCollection) : void {
        $startRow = 3;
        $curRow = $startRow;

        foreach($moCollection as $mo) {
            $sheet->setCellValue([$colMap['subjectOktmo'], $curRow], '37000000');
            $sheet->setCellValue([$colMap['subjectName'], $curRow], 'Курганская область');
            $sheet->setCellValue([$colMap['institution'], $curRow], '');
            $sheet->setCellValue([$colMap['institutionName'], $curRow], '');
            $sheet->setCellValue([$colMap['year'], $curRow], $year);
            $sheet->setCellValue([$colMap['moCode'], $curRow], $mo->code);
            $sheet->setCellValue([$colMap['moName'], $curRow], $mo->short_name);
            foreach ($colMap as $monitoringProfilesUnitId => $curCol) {
                if (!is_int($monitoringProfilesUnitId)) {
                    continue;
                }
                $value = $monitoringProfilesUnitValuesGroupedByMo[$mo->id][$monitoringProfilesUnitId] ?? 0;
                $sheet->setCellValue([$curCol, $curRow], $value);
            }
            $curRow++;
        }
    }

    public function generate(string $templateFullFilepath, int $year, int|null $commissionDecisionsId = null) : Spreadsheet
    {
        $packageIds = null;
        $currentlyUsedDate = $year.'-01-01';
        $docName = "";
        if ($commissionDecisionsId) {
            $commissionDecisions = CommissionDecision::whereYear('date',$year)->where('id', '<=', $commissionDecisionsId)->get();
            $cd = $commissionDecisions->find($commissionDecisionsId);
            $commissionDecisionIds = $commissionDecisions->pluck('id')->toArray();
            $protocolDate = $cd->date->format('d.m.Y');
            $docName = "к протоколу заседания комиссии по разработке территориальной программы ОМС Курганской области от $protocolDate";
            $packageIds = ChangePackage::whereIn('commission_decision_id', $commissionDecisionIds)->orWhere('commission_decision_id', null)->pluck('id')->toArray();

            $currentlyUsedDate = $cd->date->format('Y-m-d');
        } else {
            if ($this->initialDataFixingService->fixedYear($year)) {
                $packageIds = ChangePackage::where('commission_decision_id', null)->pluck('id')->toArray();
            } else {
                $this->plannedIndicatorChangeInitService->fromInitialData($year);
            }
        }

        bcscale(4);

        $content = $this->dataForContractService->GetGroupedByMoAndPlannedIndicatorArray($year, $packageIds);

        $moCollection = MedicalInstitution::WhereRaw("? BETWEEN effective_from AND effective_to", [$currentlyUsedDate])->orderBy('code')->get();

        $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader("Xlsx");
        $spreadsheet = $reader->load($templateFullFilepath);
        $sheet = $spreadsheet->getActiveSheet();

        // получаем соответствие id "профиля" номеру колонки
        $colMap = $this->getColMap($sheet);

        // получаем список id всех профилей присутствующих в шаблоне
        $monitoringProfilesUnitIds = array_keys($colMap);
        $monitoringProfilesUnitValuesGroupedByMo = $this->monitoringProfilesUnitValuesGroupedByMo($monitoringProfilesUnitIds, $content);
        $this->fillSheet($sheet, $colMap, $year, $monitoringProfilesUnitValuesGroupedByMo, $moCollection);
        return $spreadsheet;
    }
}

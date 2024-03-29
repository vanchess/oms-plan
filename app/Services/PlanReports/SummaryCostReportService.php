<?php
declare(strict_types=1);

namespace App\Services\PlanReports;

use App\Enum\RehabilitationBedOptionEnum;
use App\Models\ChangePackage;
use App\Models\CommissionDecision;
use App\Models\MedicalInstitution;
use App\Services\DataForContractService;
use App\Services\InitialDataFixingService;
use App\Services\PeopleAssignedInfoForContractService;
use App\Services\PlannedIndicatorChangeInitService;
use App\Services\RehabilitationProfileService;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class SummaryCostReportService
{
    use SummaryReportTrait;
    use HospitalSheetTrait;

    public function __construct(
        private DataForContractService $dataForContractService,
        private PeopleAssignedInfoForContractService $peopleAssignedInfoForContractService,
        private PlannedIndicatorChangeInitService $plannedIndicatorChangeInitService,
        private InitialDataFixingService $initialDataFixingService
    ) {}

    private function fillAmbulanceSheet(Worksheet $sheet, $content, $contentByMonth, $peopleAssigned, $moCollection, int $startRow, int $indicatorId, string $category = 'ambulance', $endRow = 100) {
        $ordinalRowNum = 0;
        $rowIndex = $startRow - 1;

        $callsAssistanceTypeId = 5; // вызовы
        $thrombolysisAssistanceTypeId = 6;// тромболизис
        foreach($moCollection as $mo) {
            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            // $sheet->setCellValue([1,$rowIndex], $mo->code);
            $sheet->setCellValue([2,$rowIndex], $mo->short_name);
            $sheet->setCellValue([7,$rowIndex], $peopleAssigned[$mo->id][$category]['peopleAssigned'] ?? 0);
            $calls = $content['mo'][$mo->id][$category][$callsAssistanceTypeId][$indicatorId] ?? '0';
            $thrombolysis = $content['mo'][$mo->id][$category][$thrombolysisAssistanceTypeId][$indicatorId] ?? '0';
            $sheet->setCellValue([8,$rowIndex], bcadd($calls, $thrombolysis));

            for($monthNum = 1; $monthNum <= 12; $monthNum++)
            {
                $calls = $contentByMonth[$monthNum]['mo'][$mo->id][$category][$callsAssistanceTypeId][$indicatorId] ?? '0';
                $thrombolysis = $contentByMonth[$monthNum]['mo'][$mo->id][$category][$thrombolysisAssistanceTypeId][$indicatorId] ?? '0';
                $sheet->setCellValue([8 + $monthNum, $rowIndex], bcadd($calls, $thrombolysis));
            }
        }
        $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);
        $this->fillSummaryRow($sheet, $rowIndex+1, 7, 20, $startRow, $rowIndex);
    }

    private function fillPolyclinicSheet(Worksheet $sheet, $content, $contentByMonth, $peopleAssigned, $moCollection, int $startRow, int $indicatorId, string $category = 'polyclinic', $endRow = 100) {
        $ordinalRowNum = 0;
        $rowIndex = $startRow - 1;
        foreach($moCollection as $mo) {
            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            // $sheet->setCellValue([1,$rowIndex], $mo->code);
            $sheet->setCellValue([2,$rowIndex], $mo->short_name);
            $sheet->setCellValue([7,$rowIndex], $peopleAssigned[$mo->id][$category]['mo']['peopleAssigned'] ?? 0);

            $assistanceTypesPerPersonSum = PlanCalculatorService::polyclinicPerPersonAssistanceTypesSum($content, $mo->id, $indicatorId);
            $servicesPerPersonSum = PlanCalculatorService::polyclinicPerPersonServicesSum($content, $mo->id, $indicatorId);
            $perPersonSum = bcadd($assistanceTypesPerPersonSum, $servicesPerPersonSum);

            // $perUnitAssistanceTypesSum = PlanCalculatorService::polyclinicPerUnitAssistanceTypesSum($content, $mo->id, $indicatorId);
            $perUnitServicesSum = PlanCalculatorService::polyclinicPerUnitServicesSum($content, $mo->id, $indicatorId);

            $fapSum = bcadd(
                PlanCalculatorService::polyclinicFapServicesSum($content, $mo->id, $indicatorId),
                PlanCalculatorService::polyclinicFapAssistanceTypesSum($content, $mo->id, $indicatorId)
            );
            $curCol = 8;
            $sheet->setCellValue([++$curCol, $rowIndex], $perPersonSum);
            $sheet->setCellValue([++$curCol, $rowIndex], $fapSum);
            $sheet->setCellValue([++$curCol, $rowIndex], $perUnitServicesSum);
            $sheet->setCellValue([++$curCol, $rowIndex], PlanCalculatorService::polyclinicPerUnitAssistanceTypesOnlyIdsSum($content, $mo->id, $indicatorId, [1,2,3,4])); // посещения профилактические, посещения разовые по заболеваниям, посещения неотложные, обращения по заболеваниям
            $sheet->setCellValue([++$curCol, $rowIndex], PlanCalculatorService::polyclinicPerUnitAssistanceTypesOnlyIdsSum($content, $mo->id, $indicatorId, [9,10,11,12])); // Диспансерное наблюдение
            $sheet->setCellValue([++$curCol, $rowIndex], PlanCalculatorService::polyclinicPerUnitAssistanceTypesOnlyIdsSum($content, $mo->id, $indicatorId, [13])); // Диспансеризация взрослого населения
            $sheet->setCellValue([++$curCol, $rowIndex], PlanCalculatorService::polyclinicPerUnitAssistanceTypesOnlyIdsSum($content, $mo->id, $indicatorId, [14])); // Углубленная диспансеризация
            $sheet->setCellValue([++$curCol, $rowIndex], PlanCalculatorService::polyclinicPerUnitAssistanceTypesOnlyIdsSum($content, $mo->id, $indicatorId, [20,21])); // Диспансеризация для оценки репродуктивного здоровья
            $sheet->setCellValue([++$curCol, $rowIndex], PlanCalculatorService::polyclinicPerUnitAssistanceTypesOnlyIdsSum($content, $mo->id, $indicatorId, [15])); // Диспансеризация сирот
            $sheet->setCellValue([++$curCol, $rowIndex], PlanCalculatorService::polyclinicPerUnitAssistanceTypesOnlyIdsSum($content, $mo->id, $indicatorId, [16])); // Диспансеризация опекаемых
            $sheet->setCellValue([++$curCol, $rowIndex], PlanCalculatorService::polyclinicPerUnitAssistanceTypesOnlyIdsSum($content, $mo->id, $indicatorId, [17])); // Профосмотры взрослых
            $sheet->setCellValue([++$curCol, $rowIndex], PlanCalculatorService::polyclinicPerUnitAssistanceTypesOnlyIdsSum($content, $mo->id, $indicatorId, [18])); // Профосмотры несовершеннолетних
            $sheet->setCellValue([++$curCol, $rowIndex], PlanCalculatorService::polyclinicPerUnitAssistanceTypesOnlyIdsSum($content, $mo->id, $indicatorId, [19])); // Школа сахарного диабета
            $sheet->setCellValue([++$curCol, $rowIndex], PlanCalculatorService::polyclinicPerUnitAssistanceTypesOnlyIdsSum($content, $mo->id, $indicatorId, [8])); // Медицинская реабилитация

            for($monthNum = 1; $monthNum <= 12; $monthNum++)
            {
                $cbm = $contentByMonth[$monthNum];
                $perPersonSum = bcadd(
                    PlanCalculatorService::polyclinicPerPersonAssistanceTypesSum($cbm, $mo->id, $indicatorId),
                    PlanCalculatorService::polyclinicPerPersonServicesSum($cbm, $mo->id, $indicatorId)
                );

                $perUnitSum = bcadd(
                    PlanCalculatorService::polyclinicPerUnitAssistanceTypesSum($cbm, $mo->id, $indicatorId),
                    PlanCalculatorService::polyclinicPerUnitServicesSum($cbm, $mo->id, $indicatorId)
                );

                $fapSum = bcadd(
                    PlanCalculatorService::polyclinicFapServicesSum($cbm, $mo->id, $indicatorId),
                    PlanCalculatorService::polyclinicFapAssistanceTypesSum($cbm, $mo->id, $indicatorId)
                );

                $v = bcadd($perPersonSum, bcadd($perUnitSum, $fapSum));

                $sheet->setCellValue([$curCol + $monthNum, $rowIndex], $v);
            }
        }
        $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);
        $this->fillSummaryRow($sheet, $rowIndex+1, 7, $curCol + 12, $startRow, $rowIndex);
    }

    public function generate(string $templateFullFilepath, int $year, int $commissionDecisionsId = null) : Spreadsheet
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
        $startRow = 7;
        $endRow = 100;

        $costIndicatorId = 4; // стоимость

        $content = $this->dataForContractService->GetArray($year, $packageIds);
        for($monthNum = 1; $monthNum <= 12; $monthNum++)
        {
            $contentByMonth[$monthNum] = $this->dataForContractService->GetArrayByYearAndMonth($year, $monthNum, $packageIds);
        }
        $peopleAssigned = $this->peopleAssignedInfoForContractService->GetArray($year, $packageIds);
        $moCollection = MedicalInstitution::WhereRaw("? BETWEEN effective_from AND effective_to", [$currentlyUsedDate])->orderBy('order')->get();

        $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader("Xlsx");
        $spreadsheet = $reader->load($templateFullFilepath);


        $sheet = $spreadsheet->getSheetByName('1.Скорая помощь, фин.обесп.');
        $sheet->setCellValue([20, 2], $docName);
        $sheet->setCellValue([7, 4], "Численность прикрепленного населения на 01.01.$year");
        $this->fillAmbulanceSheet($sheet, $content, $contentByMonth, $peopleAssigned, $moCollection, $startRow, $costIndicatorId, endRow:$endRow);

        $sheet = $spreadsheet->getSheetByName('2. АП фин.обесп.');
        $sheet->setCellValue([7, 4], "Численность прикрепленного населения по состоянию на 01.01.$year");
        $sheet->setCellValue([2, 3], "Финансовое обеспечение медицинской помощи в амбулаторных условиях на $year год");
        $this->fillPolyclinicSheet($sheet, $content, $contentByMonth, $peopleAssigned, $moCollection, $startRow, $costIndicatorId, endRow:$endRow);

        // ДС (не включая мед.реабилитацию)
        $sheet = $spreadsheet->getSheetByName('3. ДС, фин.обеспечение');
        $sheet->setCellValue([1, 3], "Финансовое обеспечение  медицинской помощи в условиях дневных стационаров на $year год (не включая медицинскую реабилитацию)");
        $this->fillHospitalSheet($sheet, $content, $contentByMonth, $moCollection, $startRow, 'daytime', ['inPolyclinic','inHospital'], RehabilitationBedOptionEnum::WithoutRehabilitation, $costIndicatorId, endRow: $endRow);

        $sheet = $spreadsheet->getSheetByName('7 МР в ДС, фин.обеспечение');
        $sheet->setCellValue([1, 3], "Финансовое обеспечение медицинской реабилитации в условиях дневных стационаров на $year год");
        $this->fillHospitalSheet($sheet, $content, $contentByMonth, $moCollection, $startRow, 'daytime', ['inPolyclinic','inHospital'], RehabilitationBedOptionEnum::OnlyRehabilitation, $costIndicatorId, endRow: $endRow);

        // КС (не включая мед.реабилитацию и ВМП)
        $sheet = $spreadsheet->getSheetByName('4 КС, фин.обеспечение');
        $sheet->setCellValue([1, 3], "Финансовое обеспечение  медицинской помощи в условиях круглосуточного стационара на $year год (не включая медицинскую реабилитацию и ВМП)");
        $this->fillHospitalSheet($sheet, $content, $contentByMonth, $moCollection, $startRow, 'roundClock', ['regular'], RehabilitationBedOptionEnum::WithoutRehabilitation, $costIndicatorId, endRow: $endRow);

        $sheet = $spreadsheet->getSheetByName('5 МР в КС, фин.обеспечение');
        $sheet->setCellValue([1, 3], "Финансовое обеспечение медицинской реабилитации в условиях круглосуточного стационара на $year год");
        $this->fillHospitalSheet($sheet, $content, $contentByMonth, $moCollection, $startRow, 'roundClock', ['regular'], RehabilitationBedOptionEnum::OnlyRehabilitation, $costIndicatorId, endRow: $endRow);

        $sheet = $spreadsheet->getSheetByName('6 ВМП, фин.обеспечение  ');
        $sheet->setCellValue([1, 3], "Финансовое обеспечение ВМП в условиях круглосуточного стационара на $year год");
        $this->fillHospitalVmpSheet($sheet, $content, $contentByMonth, $moCollection, $startRow, 'roundClock', 'vmp', $costIndicatorId, endRow: $endRow);

        return $spreadsheet;
    }
}

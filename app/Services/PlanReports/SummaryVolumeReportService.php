<?php
declare(strict_types=1);

namespace App\Services\PlanReports;

use App\Enum\MedicalServicesEnum;
use App\Enum\RehabilitationBedOptionEnum;
use App\Models\CareProfiles;
use App\Models\Category;
use App\Models\CategoryTreeNodes;
use App\Models\ChangePackage;
use App\Models\CommissionDecision;
use App\Models\Indicator;
use App\Models\IndicatorType;
use App\Models\MedicalAssistanceType;
use App\Models\MedicalInstitution;
use App\Models\MedicalServices;
use App\Models\VmpGroup;
use App\Models\VmpTypes;
use App\Services\DataForContractService;
use App\Services\InitialDataFixingService;
use App\Services\MedicalServicesService;
use App\Services\NodeService;
use App\Services\PeopleAssignedInfoForContractService;
use App\Services\PlannedIndicatorChangeInitService;
use App\Services\RehabilitationProfileService;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class SummaryVolumeReportService {
    use SummaryReportTrait;
    use HospitalSheetTrait;

    private $months = [
        1 => 'январь',
        2 => 'февраль',
        3 => 'март',
        4 => 'апрель',
        5 => 'май',
        6 => 'июнь',
        7 => 'июль',
        8 => 'август',
        9 => 'сентябрь',
        10 => 'октябрь',
        11 => 'ноябрь',
        12 => 'декабрь',
    ];

    public function __construct(
        private DataForContractService $dataForContractService,
        private PeopleAssignedInfoForContractService $peopleAssignedInfoForContractService,
        private InitialDataFixingService $initialDataFixingService,
        private PlannedIndicatorChangeInitService $plannedIndicatorChangeInitService,
        private NodeService $nodeService,
        private MedicalServicesService $medicalServicesService)
    { }

    private function fillPolyclinicSheet(Worksheet $sheet, $content, $contentByMonth, $peopleAssigned, $moCollection, int $startRow, int $serviceId, int $indicatorId, string $category = 'polyclinic', $endRow = 100) {
        $emptyLinesCount = 1; // количество пустых строк (под МТР)
        $ordinalRowNum = 0;
        $rowIndex = $startRow - 1;

        foreach($moCollection as $mo) {
            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            // $sheet->setCellValue([1,$rowIndex], $mo->code);
            $sheet->setCellValue([2,$rowIndex], $mo->short_name);
            $sheet->setCellValue([7,$rowIndex], $peopleAssigned[$mo->id][$category]['mo']['peopleAssigned'] ?? 0);

            $v = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $indicatorId, $category);
            $sheet->setCellValue([8,$rowIndex], $v);

            for($monthNum = 1; $monthNum <= 12; $monthNum++)
            {
                $v = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
                $sheet->setCellValue([8 + $monthNum, $rowIndex],  $v);
            }
        }
        $sheet->removeRow($rowIndex+1+$emptyLinesCount, $endRow-$rowIndex-$emptyLinesCount);
        $this->fillSummaryRow($sheet, $rowIndex+1+$emptyLinesCount, 7, 20, $startRow, $rowIndex+$emptyLinesCount);
    }



    public function generate(string $templateFullFilepath, int $year, int $commissionDecisionsId = null) : Spreadsheet {
        $emptyLinesCount = 1; // количество пустых строк (под МТР)

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

        $content = $this->dataForContractService->GetArray($year, $packageIds);
        $contentByMonth = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++)
        {
            $contentByMonth[$monthNum] = $this->dataForContractService->GetArrayByYearAndMonth($year, $monthNum, $packageIds);
        }
        $peopleAssigned = $this->peopleAssignedInfoForContractService->GetArray($year, $packageIds);
        $moCollection = MedicalInstitution::WhereRaw("? BETWEEN effective_from AND effective_to", [$currentlyUsedDate])->orderBy('order')->get();

        $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader("Xlsx");
        $spreadsheet = $reader->load($templateFullFilepath);


        $endRow = 100;
        $startRow = 7;

        $sheet = $spreadsheet->getSheetByName('1.Скорая помощь');
        $sheet->setCellValue([21, 2], $docName);
        $sheet->setCellValue([1, 3], "Скорая помощь, плановые объемы на $year год");
        $sheet->setCellValue([7, 4], "Численность прикрепленного населения на 01.01.$year");
        $ordinalRowNum = 0;
        $rowIndex = $startRow - 1;
        $category = 'ambulance';
        $indicatorId = 5; // вызовов
        $callsAssistanceTypeId = 5; // вызовы
        $thrombolysisAssistanceTypeId = 6;// тромболизис
        foreach($moCollection as $mo) {
            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            // $sheet->setCellValue([1,$rowIndex], $mo->code);
            $sheet->setCellValue([2,$rowIndex], $mo->short_name);
            $sheet->setCellValue([7,$rowIndex], $peopleAssigned[$mo->id][$category]['peopleAssigned'] ?? 0);
            $thrombolysis = $content['mo'][$mo->id][$category][$thrombolysisAssistanceTypeId][$indicatorId] ?? 0;
            $sheet->setCellValue([8,$rowIndex], ($content['mo'][$mo->id][$category][$callsAssistanceTypeId][$indicatorId] ?? 0) + $thrombolysis);
            $sheet->setCellValue([9,$rowIndex], $thrombolysis);
            for($monthNum = 1; $monthNum <= 12; $monthNum++)
            {
                $thrombolysis = $contentByMonth[$monthNum]['mo'][$mo->id][$category][$thrombolysisAssistanceTypeId][$indicatorId] ?? 0;
                $sheet->setCellValue([9 + $monthNum, $rowIndex],  ($contentByMonth[$monthNum]['mo'][$mo->id][$category][$callsAssistanceTypeId][$indicatorId] ?? 0) + $thrombolysis);
            }
        }
        $sheet->removeRow($rowIndex+1+$emptyLinesCount, $endRow-$rowIndex-$emptyLinesCount);
        $this->fillSummaryRow($sheet, $rowIndex+1+$emptyLinesCount, 7, 21, $startRow, $rowIndex+$emptyLinesCount);


        $sheet = $spreadsheet->getSheetByName('2.обращения по заболеваниям');
        $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в связи с заболеваниями в амбулаторных условиях на $year год");
        $sheet->setCellValue([7, 4], "Численность прикрепленного населения на 01.01.$year");
        $ordinalRowNum = 0;
        $rowIndex = $startRow - 1;
        $category = 'polyclinic';
        $indicatorId = 8; // обращений
        $assistanceTypeId = 4; //обращения по заболеваниям
        foreach($moCollection as $mo) {
            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            // $sheet->setCellValue([1,$rowIndex], $mo->code);
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

            for($monthNum = 1; $monthNum <= 12; $monthNum++)
            {
                $perPerson = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
                $perUnit = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
                $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
                $fap = 0;
                foreach ($faps as $f) {
                    $fap += $f['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? 0;
                }
                $sheet->setCellValue([8 + $monthNum, $rowIndex],  ($perPerson + $perUnit + $fap));
            }
        }
        $sheet->removeRow($rowIndex+1+$emptyLinesCount,$endRow-$rowIndex-$emptyLinesCount);
        $this->fillSummaryRow($sheet, $rowIndex+1+$emptyLinesCount, 7, 20, $startRow, $rowIndex+$emptyLinesCount);

        $sheet = $spreadsheet->getSheetByName('2.1 Мед. реабилитация амб.усл.');
        $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в амбулаторных условиях на $year год (медицинская реабилитация)");
        $sheet->setCellValue([7, 4], "Численность прикрепленного населения на 01.01.$year");
        $ordinalRowNum = 0;
        $rowIndex = $startRow - 1;
        $category = 'polyclinic';
        $indicatorId = 9; // посещений
        $assistanceTypeIds = [8]; //	медицинская реабилитация
        foreach($moCollection as $mo) {
            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            // $sheet->setCellValue([1,$rowIndex], $mo->code);
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

            for($monthNum = 1; $monthNum <= 12; $monthNum++)
            {
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
                $sheet->setCellValue([8 + $monthNum, $rowIndex],  $v);
            }
        }
        $sheet->removeRow($rowIndex+1+$emptyLinesCount,$endRow-$rowIndex-$emptyLinesCount);
        $this->fillSummaryRow($sheet, $rowIndex+1+$emptyLinesCount, 7, 20, $startRow, $rowIndex+$emptyLinesCount);

        ///////////////////////////////
        // 2.2 Диспансерное наблюдение (АП, по тарифу)
        ///////////////////////////////
        $sheetSectionNumber = 2;
        $sheetSubsectionNumber = 2;


        $spreadsheet->getDefaultStyle()->getFont()->setName('Arial');
        $spreadsheet->getDefaultStyle()->getFont()->setSize(12);

        $sheetIndex = 2;

        $sectionName = 'polyclinic';

        $typeQuantId = IndicatorType::where('name', 'volume')->first()->id;
        $polyclinicTariffDispensaryObservationCategory = CategoryTreeNodes::Where('slug', 'polyclinic-tariff-dispensary-observation')->first();
        $category = Category::find($polyclinicTariffDispensaryObservationCategory->category_id);

        $sheetName = \Illuminate\Support\Str::ucfirst($category->name);

        $dispensaryObservationTypeIds = $this->nodeService->medicalAssistanceTypesForNodeId($polyclinicTariffDispensaryObservationCategory->id);
        $dispensaryObservationTypes = MedicalAssistanceType::find($dispensaryObservationTypeIds);
        $indicatorIds = $this->nodeService->indicatorsUsedForNodeId($polyclinicTariffDispensaryObservationCategory->id);
        $indicators = Indicator::find($indicatorIds);
        $quantIndicator = $indicators->firstWhere('type_id', $typeQuantId);
        // Данные таблицы
        $dataRow = 0;
        $arrayData = [];
        $tableHasData = false;
        foreach($moCollection as $mo) {
            $rowHasData = false;
            $dataCol = 0;
            $arrayData[$dataRow] = [];
            $arrayData[$dataRow][++$dataCol] = $dataRow + 1;
            $arrayData[$dataRow][++$dataCol] = $mo->code;
            $arrayData[$dataRow][++$dataCol] = $mo->short_name;
            $assistanceTypesPerUnit = $content['mo'][$mo->id][$sectionName]['perUnit']['all']['assistanceTypes'] ?? null;

            // if (!$assistanceTypesPerUnit) { continue; }
            $tableTotalCol = ++$dataCol;
            foreach($dispensaryObservationTypes as $t) {
                $quantVal = $assistanceTypesPerUnit[$t->id][$quantIndicator->id] ?? '0';
                $v = $arrayData[$dataRow][$tableTotalCol] ?? '0';
                $arrayData[$dataRow][$tableTotalCol] = bcadd($v, $quantVal);

                if($t->name === 'прочее') {
                    continue;
                }

                $arrayData[$dataRow][++$dataCol] = $quantVal;
                if(!$rowHasData) {
                    if (bccomp($quantVal,'0') !== 0) {
                        $tableHasData = true;
                        $rowHasData = true;
                    }
                }
            }
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $assistanceTypesPerUnitByMonth = $contentByMonth[$monthNum]['mo'][$mo->id][$sectionName]['perUnit']['all']['assistanceTypes'] ?? null;
                $c = $dataCol + $monthNum;
                foreach($dispensaryObservationTypes as $t) {
                    $quantVal = $assistanceTypesPerUnitByMonth[$t->id][$quantIndicator->id] ?? '0';
                    $v = $arrayData[$dataRow][$c] ?? '0';
                    $arrayData[$dataRow][$c] = bcadd($v, $quantVal);
                }
            }
            $dataRow++;
        }

        if ($tableHasData) {
            $sheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet, mb_substr("$sheetSectionNumber.$sheetSubsectionNumber $sheetName", 0, 31));
            $spreadsheet->addSheet($sheet, ++$sheetIndex);
            $minimumDataCellWidth = 8;

            $curRow = 4;
            $curCol = 1;
            $tableStartCol = $curCol;
            $tableEndCol = $curCol;
            $tableHeadStartRow = $curRow;
            $tableHeadEndRow = $curRow;
            $staticTableHeadStartCol = $curCol;
            // Статическая часть заголовка таблицы
            $sheet->setCellValue([$curCol, $curRow], '№ п/п');
            $strwidth = mb_strwidth (' № п/п ');
            $sheet->getColumnDimensionByColumn($curCol)->setWidth($strwidth);
            $sheet->setCellValue([++$curCol, $curRow], 'Код МО');
            // $sheet->getColumnDimensionByColumn($curCol)->setAutoSize(true);
            $sheet->setCellValue([++$curCol, $curRow], 'Медицинская организация');
            $sheet->getColumnDimensionByColumn($curCol)->setWidth(50);
            $staticTableHeadEndCol = $curCol;
            $sheet->setCellValue([++$curCol, $curRow], "Всего, $quantIndicator->name");
            $tableTotalCol = $curCol;
            $sheet->getColumnDimensionByColumn($curCol)->setAutoSize(true);
            $sheet->setCellValue([++$curCol, $curRow], 'Из них по поводу');
            $tableIncludingSectionStartCol = $curCol;
            $tableIncludingSectionEndCol = $curCol;

            $curRow++;
            // Динамическая часть заголовка таблицы
            foreach($dispensaryObservationTypes as $t) {
                if($t->name === 'прочее') {
                    continue;
                }

                $sheet->setCellValue([$curCol, $curRow], $t->name);
                $width = $minimumDataCellWidth;
                $wordArr = explode(' ', $t->name);
                foreach ($wordArr as $s) {
                    $width = max($width, round(mb_strwidth ($s)) + 2);
                }

                $sheet->getColumnDimensionByColumn($curCol)->setWidth($width);
                $sheet->getRowDimension($curRow)->setRowHeight(count($wordArr) * 15);
                $sheet->mergeCells([$curCol, $curRow, $curCol, $curRow + 1]);
                $sheet->getStyle([$curCol, $curRow, $curCol, $curRow])->getAlignment()->setWrapText(true);
                $tableIncludingSectionEndCol = $curCol;

                $curCol++;
            }
            $sheet->setCellValue([$curCol, $curRow - 1], 'в том числе поквартально');

            $tablePerMonthSectionStartCol = $curCol;
            for($qr = 1; $qr <= 4; $qr++) {
                $c = $tablePerMonthSectionStartCol + (($qr-1) * 3);
                $sheet->setCellValue([$c, $curRow], "$qr квартал");
                $sheet->mergeCells([$c, $curRow, $c + 2, $curRow]);
            }
            $curRow++;
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $sheet->setCellValue([$curCol++, $curRow], $this->months[$monthNum]);
            }
            $tablePerMonthSectionEndCol = $curCol - 1;
            $tableEndCol = $tablePerMonthSectionEndCol;

            $tableHeadEndRow = $curRow;
            $curRow++;
            $tableBodyStartRow = $curRow;

            $totalRow = $tableBodyStartRow + count($arrayData) + $emptyLinesCount;
            $tableBodyEndRow = $totalRow - 1;
            $tableEndRow = $totalRow;
            // Вставляем данные из массива в таблицу
            $sheet->fromArray($arrayData, null, Coordinate::stringFromColumnIndex($tableStartCol) . $tableBodyStartRow);

            // Строка итогов
            $sheet->setCellValue([$staticTableHeadStartCol, $totalRow], 'Итого');
            for ($c = $staticTableHeadEndCol + 1; $c <= $tableEndCol; $c++) {
                $colStringName = Coordinate::stringFromColumnIndex($c);
                $sheet->setCellValue([$c, $totalRow],'=sum(' . $colStringName . $tableBodyStartRow . ':' . $colStringName . $tableBodyEndRow . ')');
            }
            // Объдинение ячеек и выравнивание такста заголовка и итога
            for($ci = $staticTableHeadStartCol; $ci <= $staticTableHeadEndCol + 1; $ci++) {
                $sheet->mergeCells([$ci, $tableHeadStartRow, $ci, $tableHeadEndRow]);
            }

            $sheet->mergeCells([$tableIncludingSectionStartCol, $tableHeadStartRow, $tableIncludingSectionEndCol, $tableHeadStartRow]);
            $sheet->mergeCells([$tablePerMonthSectionStartCol, $tableHeadStartRow, $tablePerMonthSectionEndCol, $tableHeadStartRow]);
            $sheet->mergeCells([$staticTableHeadStartCol, $totalRow, $staticTableHeadEndCol, $totalRow]);
            $sheet->getStyle([$tableStartCol, $tableHeadStartRow, $tableEndCol, $tableHeadEndRow])
                ->getAlignment()
                ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER)
                ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
            $sheet->getStyle([$tableStartCol, $totalRow, $tableStartCol, $totalRow])
                ->getAlignment()
                ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT)
                ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
            $sheet->getStyle([$tableStartCol, $totalRow, $tableEndCol, $totalRow])->getFont()->setBold(true);
            // Border таблицы
            $styleArray = [
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    ],
                ],
            ];
            $sheet->getStyle([$tableStartCol, $tableHeadStartRow, $tableEndCol, $tableEndRow])->applyFromArray($styleArray);

            // Заголовок листа
            $curRow = 0;
            $curCol = 1;

            $sheet->setCellValue([$tableEndCol,++$curRow], "Таблица $sheetSectionNumber.$sheetSubsectionNumber");
            $sheet->getStyle([$tableEndCol, $curRow, $tableEndCol, $curRow])
                ->getAlignment()
                ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT)
                ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
            ++$curRow;
            $titleCol = $curCol + 1;
            $titleRow = $curRow + 1;
            $sheet->setCellValue([$titleCol, $titleRow], "Плановые объемы медицинской помощи в амбулаторных условиях на $year год ($category->name)");
            $sheet->getStyle([$titleCol, $titleRow, $titleCol, $titleRow])->getFont()->setBold(true);
            $sheet->getRowDimension($titleRow)->setRowHeight(20);
            $sheet->freezePane([$staticTableHeadEndCol + 1, $tableBodyStartRow]);

            $sheetSubsectionNumber++;
        }

        // END 2.2 Диспансерное наблюдение

        //////////////////////////////
        // Диагностические услуги
        //////////////////////////////
        $servicesIndicatorId = 6; // услуг
        $templateSheetName = '_ДиагностическиеУслуги';
        /// список диагностических услуг актуальных на текущий год
        $medicalServiceIds = $this->medicalServicesService->getIdsByYear($year);
        $medicalServices = MedicalServices::whereIn('id', $medicalServiceIds)->orderBy('order')->get();

        // $sheetSectionNumber = 2;
        // $sheetSubsectionNumber = 2;

        foreach ($medicalServices as $ms) {
            $sheet = clone $spreadsheet->getSheetByName($templateSheetName);
            $sheetName = $ms->short_name ?? $ms->name;
            $curRow = 0;
            $tableEndCol = 20;

            $sheet->setCellValue([$tableEndCol,++$curRow], "Таблица $sheetSectionNumber.$sheetSubsectionNumber");
            $sheetTitle = str_replace(['/', '\\', '?', '*'], '_', mb_substr("$sheetSectionNumber.$sheetSubsectionNumber $sheetName", 0, 31));
            $sheet->setTitle($sheetTitle);
            $spreadsheet->addSheet($sheet, ++$sheetIndex);

            $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в связи с заболеваниями в амбулаторных условиях на $year год ($ms->name)");
            $sheet->setCellValue([7, 4], "Численность прикрепленного населения на 01.01.$year");

            $serviceId = $ms->id;
            $this->fillPolyclinicSheet($sheet, $content, $contentByMonth, $peopleAssigned, $moCollection, $startRow, $serviceId, indicatorId: $servicesIndicatorId, endRow: $endRow);

            $sheetSubsectionNumber++;
        }

        // Удаляем лист шаблон
        $templateSheetIndex = $spreadsheet->getIndex(
            $spreadsheet->getSheetByName($templateSheetName)
        );
        $spreadsheet->removeSheetByIndex($templateSheetIndex);


        ++$sheetIndex;
        $sheet = $spreadsheet->getSheetByName('3.Посещения с иными целями');
        $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в амбулаторных условиях на $year год, посещения с иными целями");
        $sheet->setCellValue([7, 4], "Численность прикрепленного населения на 01.01.$year");
        $ordinalRowNum = 0;
        $rowIndex = $startRow - 1;
        $category = 'polyclinic';
        $indicatorId = 9; // посещений
        $assistanceTypeIds = [1, 2]; //	посещения профилактические, посещения разовые по заболеваниям
        foreach($moCollection as $mo) {
            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            // $sheet->setCellValue([1,$rowIndex], $mo->code);
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

            for($monthNum = 1; $monthNum <= 12; $monthNum++)
            {
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
                $sheet->setCellValue([8 + $monthNum, $rowIndex],  $v);
            }
        }
        $sheet->removeRow($rowIndex+1+$emptyLinesCount,$endRow-$rowIndex-$emptyLinesCount);
        $this->fillSummaryRow($sheet, $rowIndex+1+$emptyLinesCount, 7, 20, $startRow, $rowIndex+$emptyLinesCount);


        /////////////////////
        // Диспансеризация
        /////////////////////
        $sheetSectionNumber = 3;
        $sheetSubsectionNumber = 2;
        $tableDataStartCol = 7;
        $tableEndCol = $tableDataStartCol + 12;
        $sheet = clone $spreadsheet->getSheetByName('3.1 Диспансеризация');

        $sheetName = 'Дисп.в.н.';
        $sheet->setTitle(mb_substr("$sheetSectionNumber.$sheetSubsectionNumber $sheetName", 0, 31));
        $spreadsheet->addSheet($sheet, ++$sheetIndex);
        $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в амбулаторных условиях на $year год, диспансеризация взрослого населения");
        $ordinalRowNum = 0;
        $rowIndex = $startRow - 1;
        $category = 'polyclinic';
        $indicatorId = 9; // посещений
        $assistanceTypeIds = [13]; //	диспансеризация
        foreach($moCollection as $mo) {
            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            // $sheet->setCellValue([1,$rowIndex], $mo->code);
            $sheet->setCellValue([2,$rowIndex], $mo->short_name);
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

            $sheet->setCellValue([$tableDataStartCol, $rowIndex], $v);

            for($monthNum = 1; $monthNum <= 12; $monthNum++)
            {
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
                $sheet->setCellValue([$tableDataStartCol + $monthNum, $rowIndex],  $v);
            }
        }
        $sheet->removeRow($rowIndex+1+$emptyLinesCount,$endRow-$rowIndex-$emptyLinesCount);
        $this->fillSummaryRow($sheet, $rowIndex+1+$emptyLinesCount, $tableDataStartCol, $tableEndCol, $startRow, $rowIndex+$emptyLinesCount);
        // Заголовок листа
        $curRow = 0;
        $curCol = 1;

        $sheet->setCellValue([$tableEndCol,++$curRow], "Таблица $sheetSectionNumber.$sheetSubsectionNumber");
        $sheet->getStyle([$tableEndCol, $curRow, $tableEndCol, $curRow])
            ->getAlignment()
            ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT)
            ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);



        $sheetSubsectionNumber++;
        $sheet = clone $spreadsheet->getSheetByName('3.1 Диспансеризация');

        $sheetName = 'Угл.дисп.';
        $sheet->setTitle(mb_substr("$sheetSectionNumber.$sheetSubsectionNumber $sheetName", 0, 31));
        $spreadsheet->addSheet($sheet, ++$sheetIndex);
        $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в амбулаторных условиях на $year год, углубленная диспансеризация");
        $ordinalRowNum = 0;
        $rowIndex = $startRow - 1;
        $category = 'polyclinic';
        $indicatorId = 9; // посещений
        $assistanceTypeIds = [14]; // диспансеризация углубленная
        foreach($moCollection as $mo) {
            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            // $sheet->setCellValue([1,$rowIndex], $mo->code);
            $sheet->setCellValue([2,$rowIndex], $mo->short_name);
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

            $sheet->setCellValue([$tableDataStartCol,$rowIndex], $v);

            for($monthNum = 1; $monthNum <= 12; $monthNum++)
            {
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
                $sheet->setCellValue([$tableDataStartCol + $monthNum, $rowIndex],  $v);
            }
        }
        $sheet->removeRow($rowIndex+1+$emptyLinesCount,$endRow-$rowIndex-$emptyLinesCount);
        $this->fillSummaryRow($sheet, $rowIndex+1+$emptyLinesCount, $tableDataStartCol, $tableEndCol, $startRow, $rowIndex+$emptyLinesCount);
        // Заголовок листа
        $curRow = 0;
        $curCol = 1;

        $sheet->setCellValue([$tableEndCol,++$curRow], "Таблица $sheetSectionNumber.$sheetSubsectionNumber");
        $sheet->getStyle([$tableEndCol, $curRow, $tableEndCol, $curRow])
            ->getAlignment()
            ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT)
            ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);


        $sheetSubsectionNumber++;
        $sheet = clone $spreadsheet->getSheetByName('3.1 Диспансеризация');

        $sheetName = 'Дисп.репрод.';
        $sheet->setTitle(mb_substr("$sheetSectionNumber.$sheetSubsectionNumber $sheetName", 0, 31));
        $spreadsheet->addSheet($sheet, ++$sheetIndex);
        $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в амбулаторных условиях на $year год, диспансеризация для оценки репродуктивного здоровья");
        $ordinalRowNum = 0;
        $rowIndex = $startRow - 1;
        $category = 'polyclinic';
        $indicatorId = 9; // посещений
        $assistanceTypeIds = [20, 21]; //	женщин репродуктивного возраста, мужчин репродуктивного возраста
        foreach($moCollection as $mo) {
            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            // $sheet->setCellValue([1,$rowIndex], $mo->code);
            $sheet->setCellValue([2,$rowIndex], $mo->short_name);
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

            $sheet->setCellValue([$tableDataStartCol, $rowIndex], $v);

            for($monthNum = 1; $monthNum <= 12; $monthNum++)
            {
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
                $sheet->setCellValue([$tableDataStartCol + $monthNum, $rowIndex],  $v);
            }
        }
        $sheet->removeRow($rowIndex+1+$emptyLinesCount,$endRow-$rowIndex-$emptyLinesCount);
        $this->fillSummaryRow($sheet, $rowIndex+1+$emptyLinesCount, $tableDataStartCol, $tableEndCol, $startRow, $rowIndex+$emptyLinesCount);

        // Заголовок листа
        $curRow = 0;
        $curCol = 1;

        $sheet->setCellValue([$tableEndCol,++$curRow], "Таблица $sheetSectionNumber.$sheetSubsectionNumber");
        $sheet->getStyle([$tableEndCol, $curRow, $tableEndCol, $curRow])
            ->getAlignment()
            ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT)
            ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);


        $sheetSubsectionNumber++;
        $sheet = clone $spreadsheet->getSheetByName('3.1 Диспансеризация');

        $sheetName = 'Дисп.сир.';
        $sheet->setTitle(mb_substr("$sheetSectionNumber.$sheetSubsectionNumber $sheetName", 0, 31));
        $spreadsheet->addSheet($sheet, ++$sheetIndex);
        $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в амбулаторных условиях на $year год, диспансеризация детей сирот");
        $ordinalRowNum = 0;
        $rowIndex = $startRow - 1;
        $category = 'polyclinic';
        $indicatorId = 9; // посещений
        $assistanceTypeIds = [15]; // диспансеризация детей-сирот
        foreach($moCollection as $mo) {
            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            // $sheet->setCellValue([1,$rowIndex], $mo->code);
            $sheet->setCellValue([2,$rowIndex], $mo->short_name);
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

            $sheet->setCellValue([$tableDataStartCol, $rowIndex], $v);

            for($monthNum = 1; $monthNum <= 12; $monthNum++)
            {
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
                $sheet->setCellValue([$tableDataStartCol + $monthNum, $rowIndex],  $v);
            }
        }
        $sheet->removeRow($rowIndex+1+$emptyLinesCount,$endRow-$rowIndex-$emptyLinesCount);
        $this->fillSummaryRow($sheet, $rowIndex+1+$emptyLinesCount, $tableDataStartCol, $tableEndCol, $startRow, $rowIndex+$emptyLinesCount);
        // Заголовок листа
        $curRow = 0;
        $curCol = 1;

        $sheet->setCellValue([$tableEndCol,++$curRow], "Таблица $sheetSectionNumber.$sheetSubsectionNumber");
        $sheet->getStyle([$tableEndCol, $curRow, $tableEndCol, $curRow])
            ->getAlignment()
            ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT)
            ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);



        $sheetSubsectionNumber++;
        $sheet = $spreadsheet->getSheetByName('3.1 Диспансеризация');
        $sheetName = 'Дисп.опека';
        $sheet->setTitle(mb_substr("$sheetSectionNumber.$sheetSubsectionNumber $sheetName", 0, 31));
        $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в амбулаторных условиях на $year год, диспансеризация детей под опекой");
        $ordinalRowNum = 0;
        $rowIndex = $startRow - 1;
        $category = 'polyclinic';
        $indicatorId = 9; // посещений
        $assistanceTypeIds = [16]; // диспансеризация опекаемых детей
        foreach($moCollection as $mo) {
            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            // $sheet->setCellValue([1,$rowIndex], $mo->code);
            $sheet->setCellValue([2,$rowIndex], $mo->short_name);
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

            $sheet->setCellValue([$tableDataStartCol,$rowIndex], $v);

            for($monthNum = 1; $monthNum <= 12; $monthNum++)
            {
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
                $sheet->setCellValue([$tableDataStartCol + $monthNum, $rowIndex],  $v);
            }
        }
        $sheet->removeRow($rowIndex+1+$emptyLinesCount,$endRow-$rowIndex-$emptyLinesCount);
        $this->fillSummaryRow($sheet, $rowIndex+1+$emptyLinesCount, $tableDataStartCol, $tableEndCol, $startRow, $rowIndex+$emptyLinesCount);
        // Заголовок листа
        $curRow = 0;
        $curCol = 1;

        $sheet->setCellValue([$tableEndCol,++$curRow], "Таблица $sheetSectionNumber.$sheetSubsectionNumber");
        $sheet->getStyle([$tableEndCol, $curRow, $tableEndCol, $curRow])
            ->getAlignment()
            ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT)
            ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);

        /// End Диспансеризация
        ++$sheetIndex;


        //////////////////////////////
        // Профилактические осмотры
        //////////////////////////////
        $tableEndCol = 8 + 12;
        $sheetSubsectionNumber++;
        $sheet = clone $spreadsheet->getSheetByName('3.2 Профилактические осмотры');

        $sheetName = 'ПО взр.';
        $sheet->setTitle(mb_substr("$sheetSectionNumber.$sheetSubsectionNumber $sheetName", 0, 31));
        $spreadsheet->addSheet($sheet, ++$sheetIndex);
        $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в амбулаторных условиях на $year год, профилактические медицинские осмотры взрослого населения");
        $sheet->setCellValue([7, 4], "Численность прикрепленного населения на 01.01.$year");
        $ordinalRowNum = 0;
        $rowIndex = $startRow - 1;
        $category = 'polyclinic';
        $indicatorId = 9; // посещений
        $assistanceTypeIds = [17]; // взрослые
        foreach($moCollection as $mo) {
            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            // $sheet->setCellValue([1,$rowIndex], $mo->code);
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

            for($monthNum = 1; $monthNum <= 12; $monthNum++)
            {
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
                $sheet->setCellValue([8 + $monthNum, $rowIndex],  $v);
            }
        }
        $sheet->removeRow($rowIndex+1+$emptyLinesCount,$endRow-$rowIndex-$emptyLinesCount);
        $this->fillSummaryRow($sheet, $rowIndex+1+$emptyLinesCount, 7, 20, $startRow, $rowIndex+$emptyLinesCount);
        // Заголовок листа
        $curRow = 0;
        $curCol = 1;

        $sheet->setCellValue([$tableEndCol,++$curRow], "Таблица $sheetSectionNumber.$sheetSubsectionNumber");
        $sheet->getStyle([$tableEndCol, $curRow, $tableEndCol, $curRow])
            ->getAlignment()
            ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT)
            ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);


        ++$sheetIndex;
        $sheetSubsectionNumber++;
        $sheet = $spreadsheet->getSheetByName('3.2 Профилактические осмотры');
        $sheetName = 'ПО дети';
        $sheet->setTitle(mb_substr("$sheetSectionNumber.$sheetSubsectionNumber $sheetName", 0, 31));
        $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в амбулаторных условиях на $year год, профилактические медицинские осмотры несовершеннолетних");
        $sheet->setCellValue([7, 4], "Численность прикрепленного населения на 01.01.$year");
        $ordinalRowNum = 0;
        $rowIndex = $startRow - 1;
        $category = 'polyclinic';
        $indicatorId = 9; // посещений
        $assistanceTypeIds = [18]; // несовершеннолетние
        foreach($moCollection as $mo) {
            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            // $sheet->setCellValue([1,$rowIndex], $mo->code);
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

            for($monthNum = 1; $monthNum <= 12; $monthNum++)
            {
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
                $sheet->setCellValue([8 + $monthNum, $rowIndex],  $v);
            }
        }
        $sheet->removeRow($rowIndex+1+$emptyLinesCount,$endRow-$rowIndex-$emptyLinesCount);
        $this->fillSummaryRow($sheet, $rowIndex+1+$emptyLinesCount, 7, 20, $startRow, $rowIndex+$emptyLinesCount);
        // Заголовок листа
        $curRow = 0;
        $curCol = 1;

        $sheet->setCellValue([$tableEndCol,++$curRow], "Таблица $sheetSectionNumber.$sheetSubsectionNumber");
        $sheet->getStyle([$tableEndCol, $curRow, $tableEndCol, $curRow])
            ->getAlignment()
            ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT)
            ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);

        // END Профилактические осмотры

        ////////////////////////////////////////////////////////////////////////////////////////////////
        ////   Центры здоровья
        ////////////////////////////////////////////////////////////////////////////////////////////////
        //$sheetSectionNumber = 3;

        $includingLabel = 'в т.ч.';

        $spreadsheet->getDefaultStyle()->getFont()->setName('Arial');
        $spreadsheet->getDefaultStyle()->getFont()->setSize(12);

        $sectionName = 'polyclinic';

        $typeQuantId = IndicatorType::where('name', 'volume')->first()->id;
        $polyclinicTariffDispensaryObservationCategory = CategoryTreeNodes::Where('slug', 'polyclinic-tariff-health-centers')->first();
        $category = Category::find($polyclinicTariffDispensaryObservationCategory->category_id);

        $sheetName = \Illuminate\Support\Str::ucfirst($category->name);

        $dispensaryObservationTypeIds = $this->nodeService->medicalAssistanceTypesForNodeId($polyclinicTariffDispensaryObservationCategory->id);
        $dispensaryObservationTypes = MedicalAssistanceType::find($dispensaryObservationTypeIds);
        $indicatorIds = $this->nodeService->indicatorsUsedForNodeId($polyclinicTariffDispensaryObservationCategory->id);
        $indicators = Indicator::find($indicatorIds);
        $quantIndicator = $indicators->firstWhere('type_id', $typeQuantId);
        $tableHasIncluding = $dispensaryObservationTypes->count() > 1;
        // Данные таблицы
        $dataRow = 0;
        $arrayData = [];
        $tableHasData = false;
        foreach($moCollection as $mo) {
            $rowHasData = false;
            $dataCol = 0;
            $arrayData[$dataRow] = [];
            $arrayData[$dataRow][++$dataCol] = $dataRow + 1;
            $arrayData[$dataRow][++$dataCol] = $mo->code;
            $arrayData[$dataRow][++$dataCol] = $mo->short_name;
            $assistanceTypesPerUnit = $content['mo'][$mo->id][$sectionName]['perUnit']['all']['assistanceTypes'] ?? null;

            // if (!$assistanceTypesPerUnit) { continue; }
            $tableTotalCol = ++$dataCol;
            foreach($dispensaryObservationTypes as $t) {
                $quantVal = $assistanceTypesPerUnit[$t->id][$quantIndicator->id] ?? '0';
                $v = $arrayData[$dataRow][$tableTotalCol] ?? '0';
                $arrayData[$dataRow][$tableTotalCol] = bcadd($v, $quantVal);

                if($tableHasIncluding && $t->name !== 'прочее') {
                    $arrayData[$dataRow][++$dataCol] = $quantVal;
                }

                if(!$rowHasData) {
                    if (bccomp($quantVal,'0') !== 0) {
                        $tableHasData = true;
                        $rowHasData = true;
                    }
                }
            }
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $assistanceTypesPerUnitByMonth = $contentByMonth[$monthNum]['mo'][$mo->id][$sectionName]['perUnit']['all']['assistanceTypes'] ?? null;
                $c = $dataCol + $monthNum;
                foreach($dispensaryObservationTypes as $t) {
                    $quantVal = $assistanceTypesPerUnitByMonth[$t->id][$quantIndicator->id] ?? '0';
                    $v = $arrayData[$dataRow][$c] ?? '0';
                    $arrayData[$dataRow][$c] = bcadd($v, $quantVal);
                }
            }
            $dataRow++;
        }

        if ($tableHasData) {
            $sheetSubsectionNumber++;
            $sheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet, mb_substr("$sheetSectionNumber.$sheetSubsectionNumber $sheetName", 0, 31));
            $spreadsheet->addSheet($sheet, ++$sheetIndex);
            $minimumDataCellWidth = 8;

            $curRow = 4;
            $curCol = 1;
            $tableStartCol = $curCol;
            $tableEndCol = $curCol;
            $tableHeadStartRow = $curRow;
            $tableHeadEndRow = $curRow;
            $staticTableHeadStartCol = $curCol;
            // Статическая часть заголовка таблицы
            $sheet->setCellValue([$curCol, $curRow], '№ п/п');
            $strwidth = mb_strwidth (' № п/п ');
            $sheet->getColumnDimensionByColumn($curCol)->setWidth($strwidth);
            $sheet->setCellValue([++$curCol, $curRow], 'Код МО');
            // $sheet->getColumnDimensionByColumn($curCol)->setAutoSize(true);
            $sheet->setCellValue([++$curCol, $curRow], 'Медицинская организация');
            $sheet->getColumnDimensionByColumn($curCol)->setWidth(50);
            $staticTableHeadEndCol = $curCol;
            $sheet->setCellValue([++$curCol, $curRow], "Всего, $quantIndicator->name");
            $tableTotalCol = $curCol;
            $sheet->getColumnDimensionByColumn($curCol)->setAutoSize(true);
            // "в том числе"
            if ($tableHasIncluding) {
                $sheet->setCellValue([++$curCol, $curRow], $includingLabel);
                $tableIncludingSectionStartCol = $curCol;
                $tableIncludingSectionEndCol = $curCol;
            }
            $curRow++;

            // Динамическая часть заголовка таблицы
            if ($tableHasIncluding) {
                foreach($dispensaryObservationTypes as $t) {
                    if($t->name === 'прочее') {
                        continue;
                    }

                    $sheet->setCellValue([$curCol, $curRow], $t->name);
                    $width = $minimumDataCellWidth;
                    $wordArr = explode(' ', $t->name);
                    foreach ($wordArr as $s) {
                        $width = max($width, round(mb_strwidth ($s)) + 2);
                    }

                    $sheet->getColumnDimensionByColumn($curCol)->setWidth($width);
                    $sheet->getRowDimension($curRow)->setRowHeight(count($wordArr) * 15);
                    $sheet->mergeCells([$curCol, $curRow, $curCol, $curRow + 1]);
                    $sheet->getStyle([$curCol, $curRow, $curCol, $curRow])->getAlignment()->setWrapText(true);
                    $tableIncludingSectionEndCol = $curCol;
                }
            }
            $curCol++;
            $sheet->setCellValue([$curCol, $curRow - 1], 'в том числе поквартально');

            $tablePerMonthSectionStartCol = $curCol;
            for($qr = 1; $qr <= 4; $qr++) {
                $c = $tablePerMonthSectionStartCol + (($qr-1) * 3);
                $sheet->setCellValue([$c, $curRow], "$qr квартал");
                $sheet->mergeCells([$c, $curRow, $c + 2, $curRow]);
            }
            $curRow++;
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $sheet->setCellValue([$curCol++, $curRow], $this->months[$monthNum]);
            }
            $tablePerMonthSectionEndCol = $curCol - 1;
            $tableEndCol = $tablePerMonthSectionEndCol;

            $tableHeadEndRow = $curRow;
            $curRow++;
            $tableBodyStartRow = $curRow;

            $totalRow = $tableBodyStartRow + count($arrayData) + $emptyLinesCount;
            $tableBodyEndRow = $totalRow - 1;
            $tableEndRow = $totalRow;
            // Вставляем данные из массива в таблицу
            $sheet->fromArray($arrayData, null, Coordinate::stringFromColumnIndex($tableStartCol) . $tableBodyStartRow);

            // Строка итогов
            $sheet->setCellValue([$staticTableHeadStartCol, $totalRow], 'Итого');
            for ($c = $staticTableHeadEndCol + 1; $c <= $tableEndCol; $c++) {
                $colStringName = Coordinate::stringFromColumnIndex($c);
                $sheet->setCellValue([$c, $totalRow],'=sum(' . $colStringName . $tableBodyStartRow . ':' . $colStringName . $tableBodyEndRow . ')');
            }
            // Объдинение ячеек и выравнивание такста заголовка и итога
            for($ci = $staticTableHeadStartCol; $ci <= $staticTableHeadEndCol + 1; $ci++) {
                $sheet->mergeCells([$ci, $tableHeadStartRow, $ci, $tableHeadEndRow]);
            }
            if ($tableHasIncluding) {
                $sheet->mergeCells([$tableIncludingSectionStartCol, $tableHeadStartRow, $tableIncludingSectionEndCol, $tableHeadStartRow]);
            }
            $sheet->mergeCells([$tablePerMonthSectionStartCol, $tableHeadStartRow, $tablePerMonthSectionEndCol, $tableHeadStartRow]);
            $sheet->mergeCells([$staticTableHeadStartCol, $totalRow, $staticTableHeadEndCol, $totalRow]);
            $sheet->getStyle([$tableStartCol, $tableHeadStartRow, $tableEndCol, $tableHeadEndRow])
                ->getAlignment()
                ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER)
                ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
            $sheet->getStyle([$tableStartCol, $totalRow, $tableStartCol, $totalRow])
                ->getAlignment()
                ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT)
                ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
            $sheet->getStyle([$tableStartCol, $totalRow, $tableEndCol, $totalRow])->getFont()->setBold(true);
            // Border таблицы
            $styleArray = [
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    ],
                ],
            ];
            $sheet->getStyle([$tableStartCol, $tableHeadStartRow, $tableEndCol, $tableEndRow])->applyFromArray($styleArray);

            // Заголовок листа
            $curRow = 0;
            $curCol = 1;

            $sheet->setCellValue([$tableEndCol,++$curRow], "Таблица $sheetSectionNumber.$sheetSubsectionNumber");
            $sheet->getStyle([$tableEndCol, $curRow, $tableEndCol, $curRow])
                ->getAlignment()
                ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT)
                ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
            ++$curRow;
            $titleCol = $curCol + 1;
            $titleRow = $curRow + 1;
            $sheet->setCellValue([$titleCol, $titleRow], "Плановые объемы медицинской помощи в амбулаторных условиях на $year год ($category->name)");
            $sheet->getStyle([$titleCol, $titleRow, $titleCol, $titleRow])->getFont()->setBold(true);
            $sheet->getRowDimension($titleRow)->setRowHeight(20);
            $sheet->freezePane([$staticTableHeadEndCol + 1, $tableBodyStartRow]);
        }

        //////////////////////////////
        // школа сахарного диабета
        //////////////////////////////
        if ($year >= 2025) {
            //$sheetSectionNumber = 3;
            //$sheetSubsectionNumber = 2;
            $sheetSubsectionNumber++;
            $includingLabel = 'в т.ч.';

            $spreadsheet->getDefaultStyle()->getFont()->setName('Arial');
            $spreadsheet->getDefaultStyle()->getFont()->setSize(12);

            // $sheetIndex = 3;

            $sectionName = 'polyclinic';

            $typeQuantId = IndicatorType::where('name', 'volume')->first()->id;
            $polyclinicTariffDispensaryObservationCategory = CategoryTreeNodes::Where('slug', 'polyclinic-tariff-diabetes-schools')->first();
            $category = Category::find($polyclinicTariffDispensaryObservationCategory->category_id);

            $sheetName = \Illuminate\Support\Str::ucfirst($category->name);

            $dispensaryObservationTypeIds = $this->nodeService->medicalAssistanceTypesForNodeId($polyclinicTariffDispensaryObservationCategory->id);
            $dispensaryObservationTypes = MedicalAssistanceType::find($dispensaryObservationTypeIds);
            $indicatorIds = $this->nodeService->indicatorsUsedForNodeId($polyclinicTariffDispensaryObservationCategory->id);
            $indicators = Indicator::find($indicatorIds);

            $quantIndicator = $indicators->firstWhere('type_id', $typeQuantId);
            // Данные таблицы
            $dataRow = 0;
            $arrayData = [];
            $tableHasData = false;
            foreach($moCollection as $mo) {
                $rowHasData = false;
                $dataCol = 0;
                $arrayData[$dataRow] = [];
                $arrayData[$dataRow][++$dataCol] = $dataRow + 1;
                $arrayData[$dataRow][++$dataCol] = $mo->code;
                $arrayData[$dataRow][++$dataCol] = $mo->short_name;
                $assistanceTypesPerUnit = $content['mo'][$mo->id][$sectionName]['perUnit']['all']['assistanceTypes'] ?? null;

                // if (!$assistanceTypesPerUnit) { continue; }
                $tableTotalCol = ++$dataCol;
                foreach($dispensaryObservationTypes as $t) {
                    $quantVal = $assistanceTypesPerUnit[$t->id][$quantIndicator->id] ?? '0';
                    $v = $arrayData[$dataRow][$tableTotalCol] ?? '0';
                    $arrayData[$dataRow][$tableTotalCol] = bcadd($v, $quantVal);

                    if($t->name === 'прочее') {
                        continue;
                    }

                    $arrayData[$dataRow][++$dataCol] = $quantVal;
                    if(!$rowHasData) {
                        if (bccomp($quantVal,'0') !== 0) {
                            $tableHasData = true;
                            $rowHasData = true;
                        }
                    }
                }
                for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                    $assistanceTypesPerUnitByMonth = $contentByMonth[$monthNum]['mo'][$mo->id][$sectionName]['perUnit']['all']['assistanceTypes'] ?? null;
                    $c = $dataCol + $monthNum;
                    foreach($dispensaryObservationTypes as $t) {
                        $quantVal = $assistanceTypesPerUnitByMonth[$t->id][$quantIndicator->id] ?? '0';
                        $v = $arrayData[$dataRow][$c] ?? '0';
                        $arrayData[$dataRow][$c] = bcadd($v, $quantVal);
                    }
                }
                $dataRow++;
            }

            if ($tableHasData) {
                $sheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet, mb_substr("$sheetSectionNumber.$sheetSubsectionNumber $sheetName", 0, 31));
                $spreadsheet->addSheet($sheet, ++$sheetIndex);
                $minimumDataCellWidth = 8;

                $curRow = 4;
                $curCol = 1;
                $tableStartCol = $curCol;
                $tableEndCol = $curCol;
                $tableHeadStartRow = $curRow;
                $tableHeadEndRow = $curRow;
                $staticTableHeadStartCol = $curCol;
                // Статическая часть заголовка таблицы
                $sheet->setCellValue([$curCol, $curRow], '№ п/п');
                $strwidth = mb_strwidth (' № п/п ');
                $sheet->getColumnDimensionByColumn($curCol)->setWidth($strwidth);
                $sheet->setCellValue([++$curCol, $curRow], 'Код МО');
                // $sheet->getColumnDimensionByColumn($curCol)->setAutoSize(true);
                $sheet->setCellValue([++$curCol, $curRow], 'Медицинская организация');
                $sheet->getColumnDimensionByColumn($curCol)->setWidth(50);
                $staticTableHeadEndCol = $curCol;
                $sheet->setCellValue([++$curCol, $curRow], "Всего, $quantIndicator->name");
                $tableTotalCol = $curCol;
                $sheet->getColumnDimensionByColumn($curCol)->setAutoSize(true);
                $sheet->setCellValue([++$curCol, $curRow], $includingLabel);
                $tableIncludingSectionStartCol = $curCol;
                $tableIncludingSectionEndCol = $curCol;

                $curRow++;
                // Динамическая часть заголовка таблицы
                foreach($dispensaryObservationTypes as $t) {
                    if($t->name === 'прочее') {
                        continue;
                    }

                    $sheet->setCellValue([$curCol, $curRow], $t->name);
                    $width = $minimumDataCellWidth;
                    $wordArr = explode(' ', $t->name);
                    foreach ($wordArr as $s) {
                        $width = max($width, round(mb_strwidth ($s)) + 2);
                    }

                    $sheet->getColumnDimensionByColumn($curCol)->setWidth($width);
                    $sheet->getRowDimension($curRow)->setRowHeight(count($wordArr) * 15);
                    $sheet->mergeCells([$curCol, $curRow, $curCol, $curRow + 1]);
                    $sheet->getStyle([$curCol, $curRow, $curCol, $curRow])->getAlignment()->setWrapText(true);
                    $tableIncludingSectionEndCol = $curCol;

                    $curCol++;
                }
                $sheet->setCellValue([$curCol, $curRow - 1], 'в том числе поквартально');

                $tablePerMonthSectionStartCol = $curCol;
                for($qr = 1; $qr <= 4; $qr++) {
                    $c = $tablePerMonthSectionStartCol + (($qr-1) * 3);
                    $sheet->setCellValue([$c, $curRow], "$qr квартал");
                    $sheet->mergeCells([$c, $curRow, $c + 2, $curRow]);
                }
                $curRow++;
                for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                    $sheet->setCellValue([$curCol++, $curRow], $this->months[$monthNum]);
                }
                $tablePerMonthSectionEndCol = $curCol - 1;
                $tableEndCol = $tablePerMonthSectionEndCol;

                $tableHeadEndRow = $curRow;
                $curRow++;
                $tableBodyStartRow = $curRow;

                $totalRow = $tableBodyStartRow + count($arrayData) + $emptyLinesCount;
                $tableBodyEndRow = $totalRow - 1;
                $tableEndRow = $totalRow;
                // Вставляем данные из массива в таблицу
                $sheet->fromArray($arrayData, null, Coordinate::stringFromColumnIndex($tableStartCol) . $tableBodyStartRow);

                // Строка итогов
                $sheet->setCellValue([$staticTableHeadStartCol, $totalRow], 'Итого');
                for ($c = $staticTableHeadEndCol + 1; $c <= $tableEndCol; $c++) {
                    $colStringName = Coordinate::stringFromColumnIndex($c);
                    $sheet->setCellValue([$c, $totalRow],'=sum(' . $colStringName . $tableBodyStartRow . ':' . $colStringName . $tableBodyEndRow . ')');
                }
                // Объдинение ячеек и выравнивание такста заголовка и итога
                for($ci = $staticTableHeadStartCol; $ci <= $staticTableHeadEndCol + 1; $ci++) {
                    $sheet->mergeCells([$ci, $tableHeadStartRow, $ci, $tableHeadEndRow]);
                }

                $sheet->mergeCells([$tableIncludingSectionStartCol, $tableHeadStartRow, $tableIncludingSectionEndCol, $tableHeadStartRow]);
                $sheet->mergeCells([$tablePerMonthSectionStartCol, $tableHeadStartRow, $tablePerMonthSectionEndCol, $tableHeadStartRow]);
                $sheet->mergeCells([$staticTableHeadStartCol, $totalRow, $staticTableHeadEndCol, $totalRow]);
                $sheet->getStyle([$tableStartCol, $tableHeadStartRow, $tableEndCol, $tableHeadEndRow])
                    ->getAlignment()
                    ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER)
                    ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
                $sheet->getStyle([$tableStartCol, $totalRow, $tableStartCol, $totalRow])
                    ->getAlignment()
                    ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT)
                    ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
                $sheet->getStyle([$tableStartCol, $totalRow, $tableEndCol, $totalRow])->getFont()->setBold(true);
                // Border таблицы
                $styleArray = [
                    'borders' => [
                        'allBorders' => [
                            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                        ],
                    ],
                ];
                $sheet->getStyle([$tableStartCol, $tableHeadStartRow, $tableEndCol, $tableEndRow])->applyFromArray($styleArray);

                // Заголовок листа
                $curRow = 0;
                $curCol = 1;

                $sheet->setCellValue([$tableEndCol,++$curRow], "Таблица $sheetSectionNumber.$sheetSubsectionNumber");
                $sheet->getStyle([$tableEndCol, $curRow, $tableEndCol, $curRow])
                    ->getAlignment()
                    ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT)
                    ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
                ++$curRow;
                $titleCol = $curCol + 1;
                $titleRow = $curRow + 1;
                $sheet->setCellValue([$titleCol, $titleRow], "Плановые объемы медицинской помощи в амбулаторных условиях на $year год ($category->name)");
                $sheet->getStyle([$titleCol, $titleRow, $titleCol, $titleRow])->getFont()->setBold(true);
                $sheet->getRowDimension($titleRow)->setRowHeight(20);
                $sheet->freezePane([$staticTableHeadEndCol + 1, $tableBodyStartRow]);

                $sheetSubsectionNumber++;
            }


            // Удаляем лист шаблон используемый для планов до 2025 года
            $templateSheetIndex = $spreadsheet->getIndex(
                $spreadsheet->getSheetByName('3.9 Школа С.Д.')
            );
            $spreadsheet->removeSheetByIndex($templateSheetIndex);

        } else {
            // школа сахарного диабета до 2025 года
            $sheet = $spreadsheet->getSheetByName('3.9 Школа С.Д.');
            $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в амбулаторных условиях на $year год, школа сахарного диабета");
            $sheet->setCellValue([7, 4], "Численность прикрепленного населения на 01.01.$year");
            $ordinalRowNum = 0;
            $rowIndex = $startRow - 1;
            $category = 'polyclinic';
            $indicatorId = 9; // посещений
            $assistanceTypeIds = [19]; // школы для пациентов c сахарным диабетом
            foreach($moCollection as $mo) {
                $ordinalRowNum++;
                $rowIndex++;
                $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
                // $sheet->setCellValue([1,$rowIndex], $mo->code);
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

                for($monthNum = 1; $monthNum <= 12; $monthNum++)
                {
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
                    $sheet->setCellValue([8 + 3 + $monthNum, $rowIndex],  $v);
                }
            }
            $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);
            $this->fillSummaryRow($sheet, $rowIndex+1, 7, 20, $startRow, $rowIndex);
        }


        //

        $sheet = $spreadsheet->getSheetByName('4 Неотложная помощь');
        $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в амбулаторных условиях на $year год, неотложная помощь");
        $sheet->setCellValue([7, 4], "Численность прикрепленного населения на 01.01.$year");
        $ordinalRowNum = 0;
        $rowIndex = $startRow - 1;
        $category = 'polyclinic';
        $indicatorId = 9; // посещений
        $assistanceTypeIds = [3]; //	посещения неотложные
        foreach($moCollection as $mo) {
            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            // $sheet->setCellValue([1,$rowIndex], $mo->code);
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

            for($monthNum = 1; $monthNum <= 12; $monthNum++)
            {
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
                $sheet->setCellValue([8 + $monthNum, $rowIndex], $v);
            }
        }
        $sheet->removeRow($rowIndex+1 + $emptyLinesCount,$endRow-$rowIndex - $emptyLinesCount);
        $this->fillSummaryRow($sheet, $rowIndex+1 + $emptyLinesCount, 7, 20, $startRow, $rowIndex + $emptyLinesCount);





        //////////////////////////////
        // Круглосуточный ст.
        //////////////////////////////
        $hospitalizationsIndicatorId = 7; // госпитализаций
        $sheet = $spreadsheet->getSheetByName('5. Круглосуточный ст.');
        $sheet->setCellValue([1, 3], "Объемы медицинской помощи в условиях круглосуточного стационара (не включая ВМП и медицинскую реабилитацию) на $year год");
        $this->fillHospitalSheet($sheet, $content, $contentByMonth, $moCollection, $startRow, 'roundClock', ['regular'], RehabilitationBedOptionEnum::WithoutRehabilitation, $hospitalizationsIndicatorId, endRow: $endRow);

        $sheet = $spreadsheet->getSheetByName('6.ВМП');
        $sheet->setCellValue([1, 3], "Объемы высокотехнологичной медицинской помощи в условиях круглосуточного стационара на $year год");
        $this->fillHospitalVmpSheet($sheet, $content, $contentByMonth, $moCollection, $startRow, 'roundClock', 'vmp', $hospitalizationsIndicatorId, endRow: $endRow);

        $sheet = $spreadsheet->getSheetByName('7. Медреабилитация в КС');
        $sheet->setCellValue([1, 3], "Объемы медицинской реабилитации в условиях круглосуточного стационара на $year год");
        $this->fillHospitalSheet($sheet, $content, $contentByMonth, $moCollection, $startRow, 'roundClock', ['regular'], RehabilitationBedOptionEnum::OnlyRehabilitation, $hospitalizationsIndicatorId, endRow: $endRow);

        $casesTreatmentIndicatorId = 2; // случаев лечения
        $sheet = $spreadsheet->getSheetByName('8. Дневные стационары');
        $sheet->setCellValue([1, 3], "Объемы медицинской помощи в условиях дневных стационаров (не включая медицинскую реабилитацию) на $year год");
        $this->fillHospitalSheet($sheet, $content, $contentByMonth, $moCollection, $startRow, 'daytime', ['inPolyclinic','inHospital'], RehabilitationBedOptionEnum::WithoutRehabilitation, $casesTreatmentIndicatorId, endRow: $endRow);

        $sheet = $spreadsheet->getSheetByName('9. Медреабилитация в ДС');
        $sheet->setCellValue([1, 3], "Объемы медицинской реабилитации в условиях дневных стационаров на $year год");
        $this->fillHospitalSheet($sheet, $content, $contentByMonth, $moCollection, $startRow, 'daytime', ['inPolyclinic','inHospital'], RehabilitationBedOptionEnum::OnlyRehabilitation, $casesTreatmentIndicatorId, endRow: $endRow);


        $sheet = $spreadsheet->getSheetByName("6.1. ВМП в разрезе методов");
        $sheet->setCellValue([1, 3], "Плановые объемы и финансовое обеспечение высокотехнологичной медицинской помощи (ВМП) в условиях круглосуточного стационара на $year год");
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


        // ОГЛАВЛЕНИЕ
        $sheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet, 'Оглавление');
        $spreadsheet->addSheet($sheet, 0);
        $sheetCount = $spreadsheet->getSheetCount();
        for ($i = 1; $i < $sheetCount; $i++){
            $s = $spreadsheet->getSheet($i);
            $st = $s->getTitle();
            $colName = Coordinate::stringFromColumnIndex(1);
           //echo '=HYPERLINK("#\'' . $st . '\'!' . "A1" . '";"' . $st . '")' . '<br>';
            $sheet->setCellValue([1, $i], $st);
            $spreadsheet->addNamedRange( new \PhpOffice\PhpSpreadsheet\NamedRange('my_named_range_' . $i, $s, 'A1'));
            $sheet->getCell([1, $i])->getHyperlink()->setUrl('sheet://my_named_range_' . $i);

        }
        $spreadsheet->setActiveSheetIndex(0);

        return $spreadsheet;
    }
}

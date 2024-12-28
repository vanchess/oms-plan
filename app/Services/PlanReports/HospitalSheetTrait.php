<?php
declare(strict_types=1);

namespace App\Services\PlanReports;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

trait HospitalSheetTrait
{
    private function fillHospitalSheet(Worksheet $sheet, $content, $contentByMonth, $moCollection, int $startRow, string $daytimeOrRoundClock, array $hospitalSubTypes, int $rehabilitationBedOption, int $indicatorId, string $category = 'hospital', $endRow = 100) {
        bcscale(4);

        $tableDataStartCol = 7;
        $tableEndCol = 19;
        $emptyLinesCount = 1; // количество пустых строк (под МТР)
        $ordinalRowNum = 0;
        $rowIndex = $startRow - 1;

        foreach($moCollection as $mo) {
            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            // $sheet->setCellValue([1,$rowIndex], $mo->code);
            $sheet->setCellValue([2,$rowIndex], $mo->short_name);

            $v = '0';
            foreach($hospitalSubTypes as $hospitalSubType) {
                $addV = PlanCalculatorService::hospitalBedProfilesSum($content, $mo->id, $daytimeOrRoundClock, $hospitalSubType, $rehabilitationBedOption, $indicatorId, $category);
                $v = bcadd($v, $addV);
            }

            $sheet->setCellValue([$tableDataStartCol, $rowIndex], $v);
            for($monthNum = 1; $monthNum <= 12; $monthNum++)
            {
                $v = '0';
                foreach($hospitalSubTypes as $hospitalSubType) {
                    $addV = PlanCalculatorService::hospitalBedProfilesSum($contentByMonth[$monthNum], $mo->id, $daytimeOrRoundClock, $hospitalSubType, $rehabilitationBedOption, $indicatorId, $category);
                    $v = bcadd($v, $addV);
                }
                $sheet->setCellValue([$tableDataStartCol + $monthNum, $rowIndex],  $v);
            }
        }
        $sheet->removeRow($rowIndex+1 + $emptyLinesCount,$endRow-$rowIndex - $emptyLinesCount);
        $this->fillSummaryRow($sheet, $rowIndex+1 + $emptyLinesCount, $tableDataStartCol, $tableEndCol, $startRow, $rowIndex + $emptyLinesCount);
    }

    private function fillHospitalVmpSheet(Worksheet $sheet, $content, $contentByMonth, $moCollection, int $startRow, string $daytimeOrRoundClock, string $hospitalSubType = 'vmp', int $indicatorId = 7, string $category = 'hospital', $endRow = 100) {
        bcscale(4);

        $tableDataStartCol = 7;
        $tableEndCol = 19;
        $emptyLinesCount = 1; // количество пустых строк (под МТР)
        $ordinalRowNum = 0;
        $rowIndex = $startRow - 1;

        foreach($moCollection as $mo) {
            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            // $sheet->setCellValue([1,$rowIndex], $mo->code);
            $sheet->setCellValue([2,$rowIndex], $mo->short_name);

            $v = PlanCalculatorService::hospitalVmpSum($content, $mo->id, $daytimeOrRoundClock, $hospitalSubType, $indicatorId, $category);

            $sheet->setCellValue([$tableDataStartCol, $rowIndex], $v);
            for($monthNum = 1; $monthNum <= 12; $monthNum++)
            {
                $v  = PlanCalculatorService::hospitalVmpSum($contentByMonth[$monthNum], $mo->id, $daytimeOrRoundClock, $hospitalSubType, $indicatorId, $category);
                $sheet->setCellValue([$tableDataStartCol + $monthNum, $rowIndex],  $v);
            }
        }
        $sheet->removeRow($rowIndex+1 + $emptyLinesCount,$endRow-$rowIndex - $emptyLinesCount);
        $this->fillSummaryRow($sheet, $rowIndex+1 + $emptyLinesCount, $tableDataStartCol, $tableEndCol, $startRow, $rowIndex + $emptyLinesCount);
    }
}

<?php
declare(strict_types=1);

namespace App\Services\PlanReports;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

trait HospitalSheetTrait
{
    private function fillHospitalSheet(Worksheet $sheet, $content, $contentByMonth, $moCollection, int $startRow, string $daytimeOrRoundClock, array $hospitalSubTypes, int $rehabilitationBedOption, int $indicatorId, string $category = 'hospital', $endRow = 100) {
        bcscale(4);

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

            $sheet->setCellValue([7,$rowIndex], $v);
            for($monthNum = 1; $monthNum <= 12; $monthNum++)
            {
                $v = '0';
                foreach($hospitalSubTypes as $hospitalSubType) {
                    $addV = PlanCalculatorService::hospitalBedProfilesSum($contentByMonth[$monthNum], $mo->id, $daytimeOrRoundClock, $hospitalSubType, $rehabilitationBedOption, $indicatorId, $category);
                    $v = bcadd($v, $addV);
                }
                $sheet->setCellValue([7 + $monthNum, $rowIndex],  $v);
            }
        }
        $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);
        $this->fillSummaryRow($sheet, $rowIndex+1, 7, 20, $startRow, $rowIndex);
    }

    private function fillHospitalVmpSheet(Worksheet $sheet, $content, $contentByMonth, $moCollection, int $startRow, string $daytimeOrRoundClock, string $hospitalSubType = 'vmp', int $indicatorId = 7, string $category = 'hospital', $endRow = 100) {
        bcscale(4);

        $ordinalRowNum = 0;
        $rowIndex = $startRow - 1;

        foreach($moCollection as $mo) {
            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            // $sheet->setCellValue([1,$rowIndex], $mo->code);
            $sheet->setCellValue([2,$rowIndex], $mo->short_name);

            $v = PlanCalculatorService::hospitalVmpSum($content, $mo->id, $daytimeOrRoundClock, $hospitalSubType, $indicatorId, $category);

            $sheet->setCellValue([7,$rowIndex], $v);
            for($monthNum = 1; $monthNum <= 12; $monthNum++)
            {
                $v  = PlanCalculatorService::hospitalVmpSum($contentByMonth[$monthNum], $mo->id, $daytimeOrRoundClock, $hospitalSubType, $indicatorId, $category);
                $sheet->setCellValue([7 + $monthNum, $rowIndex],  $v);
            }
        }
        $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);
        $this->fillSummaryRow($sheet, $rowIndex+1, 7, 20, $startRow, $rowIndex);
    }
}

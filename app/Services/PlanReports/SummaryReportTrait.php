<?php
declare(strict_types=1);

namespace App\Services\PlanReports;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

trait SummaryReportTrait
{
    private function fillSummaryRow(Worksheet $sheet, int $summaryRowIndex, int $fromCol, int $toCol, int $fromRow, int $toRow){
        for ($col = $fromCol; $col <= $toCol; $col++) {
            $colName = Coordinate::stringFromColumnIndex($col);
            $sheet->setCellValue([$col,$summaryRowIndex], "=SUM({$colName}{$fromRow}:{$colName}{$toRow})");
        }
    }
}

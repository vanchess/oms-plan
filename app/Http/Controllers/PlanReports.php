<?php

namespace App\Http\Controllers;

use App\Models\ChangePackage;
use App\Models\CommissionDecision;
use App\Services\PlannedIndicatorChangeInitService;
use App\Services\InitialDataFixingService;
use App\Services\PlanReports\SummaryCostReportService;
use App\Services\PlanReports\SummaryVolumeReportService as SummaryVolumeReportService;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Storage;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class PlanReports extends Controller
{
    public function SummaryVolume (SummaryVolumeReportService $reportService, int $year, int $commissionDecisionsId = null) {

        $path = 'xlsx';
        $templateFileName = '1.xlsx';
        $templateFilePath = $path . DIRECTORY_SEPARATOR . $templateFileName;
        $templateFullFilepath = Storage::path($templateFilePath);
        $resultFileName = 'объемы.xlsx';
        $strDateTimeNow = date("Y-m-d-His");
        $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . ' ' . $resultFileName;
        $fullResultFilepath = Storage::path($resultFilePath);

        $spreadsheet = $reportService->generate($templateFullFilepath, year: $year, commissionDecisionsId: $commissionDecisionsId);

        $writer = new Xlsx($spreadsheet);
        $writer->save($fullResultFilepath);
        return Storage::download($resultFilePath);
    }

    public function SummaryCost (SummaryCostReportService $reportService, int $year, int $commissionDecisionsId = null) {
        $path = 'xlsx';
        $templateFileName = '2.xlsx';
        $templateFilePath = $path . DIRECTORY_SEPARATOR . $templateFileName;
        $templateFullFilepath = Storage::path($templateFilePath);
        $resultFileName = 'стоимость.xlsx';
        $strDateTimeNow = date("Y-m-d-His");
        $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . ' ' . $resultFileName;
        $fullResultFilepath = Storage::path($resultFilePath);

        $spreadsheet = $reportService->generate($templateFullFilepath, year: $year, commissionDecisionsId: $commissionDecisionsId);

        $writer = new Xlsx($spreadsheet);
        $writer->save($fullResultFilepath);
        return Storage::download($resultFilePath);
    }
}

<?php

namespace App\Http\Controllers;

use App\Models\ChangePackage;
use App\Models\CommissionDecision;
use App\Services\PlannedIndicatorChangeInitService;
use App\Services\InitialDataFixingService;
use App\Services\PlanReports\NumberOfBedsReportService;
use App\Services\PlanReports\SummaryCostReportService;
use App\Services\PlanReports\SummaryVolumeReportService as SummaryVolumeReportService;
use App\Services\PlanReports\PumpPggReportService;
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
        $resultFileName = "объемы($year-$commissionDecisionsId).xlsx";
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
        $resultFileName = "стоимость($year-$commissionDecisionsId).xlsx";
        $strDateTimeNow = date("Y-m-d-His");
        $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . ' ' . $resultFileName;
        $fullResultFilepath = Storage::path($resultFilePath);

        $spreadsheet = $reportService->generate($templateFullFilepath, year: $year, commissionDecisionsId: $commissionDecisionsId);

        $writer = new Xlsx($spreadsheet);
        $writer->save($fullResultFilepath);
        return Storage::download($resultFilePath);
    }

    public function NumberOfBeds(NumberOfBedsReportService $numberOfBedsReportService, int $year, int $commissionDecisionsId = null) {
        $xml = $numberOfBedsReportService->generateXml('plan.reports.xml.numberOfBedsXml', $year, $commissionDecisionsId);
        $strDateTimeNow = date("Y-m-d-His");
        $fileneme = $strDateTimeNow . '_numberOfBeds.xml';

        //dd( $xml);
        return response($xml, 200)
                    ->withHeaders([
                        'Content-Type' => 'text/xml',
                        'Cache-Control' => 'no-cache',
                        'Content-Description' => 'File Transfer',
                        'Content-Disposition' => 'attachment; filename=' . $fileneme,
                        'Content-Transfer-Encoding' => 'binary'
                    ]);
    }

    public function PumpPgg (PumpPggReportService $reportService, int $year, int $commissionDecisionsId = null) {
        $path = 'xlsx' . DIRECTORY_SEPARATOR . 'pump';
        $templateFileName = 'PumpPgg_v6.xlsx';
        $templateFilePath = $path . DIRECTORY_SEPARATOR . $templateFileName;
        $templateFullFilepath = Storage::path($templateFilePath);
        $resultFileName = "Приложение №2.11 Шаблон плановые объёмы v6($year-$commissionDecisionsId).xlsx";
        $strDateTimeNow = date("Y-m-d-His");
        $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . '_' . $resultFileName;
        $fullResultFilepath = Storage::path($resultFilePath);

        $spreadsheet = $reportService->generate($templateFullFilepath, year: $year, commissionDecisionsId: $commissionDecisionsId);

        $writer = new Xlsx($spreadsheet);
        $writer->save($fullResultFilepath);
        return Storage::download($resultFilePath);
    }
}

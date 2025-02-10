<?php
declare(strict_types=1);

namespace App\Http\Controllers;

use App\Models\CommissionDecision;
use App\Services\PlanReports\MeetingMinutesReportService;
use App\Services\PlanReports\NumberOfBedsReportService;
use App\Services\PlanReports\SummaryCostReportService;
use App\Services\PlanReports\SummaryVolumeReportService as SummaryVolumeReportService;
use App\Services\PlanReports\PumpPggReportService;
use App\Services\PlanReports\VitacorePlanReportService;
use Illuminate\Support\Facades\Storage;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class PlanReports extends Controller
{
    public function MeetingMinutes(MeetingMinutesReportService $reportService, int $year, int $commissionDecisionsId) {
        $cd = CommissionDecision::find($commissionDecisionsId);
        $protocolDate = $cd->date->format('d.m.Y');
        $protocolNumber = $cd->number;
        $protocolNumberForFileName = preg_replace('/[^a-zа-я\d.]/ui', '_', $protocolNumber);

        $path = 'xlsx';
        $templateFileName = 'meetingMinutes20250120.xlsx';
        $templateFilePath = $path . DIRECTORY_SEPARATOR . $templateFileName;
        $templateFullFilepath = Storage::path($templateFilePath);

        $resultFileName = "protocol_№$protocolNumberForFileName($protocolDate).xlsx";
        $strDateTimeNow = date("Y-m-d-His");
        $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . ' ' . $resultFileName;
        $fullResultFilepath = Storage::path($resultFilePath);

        $spreadsheet = $reportService->generate($templateFullFilepath, $year, $commissionDecisionsId);

        $writer = new Xlsx($spreadsheet);
        $writer->save($fullResultFilepath);
        return Storage::download($resultFilePath);
    }


    public function SummaryVolume(SummaryVolumeReportService $reportService, int $year, int $commissionDecisionsId = null) {
        $path = 'xlsx';
        $templateFileName = 'templateSummaryVolume20241223.xlsx';
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
        $templateFileName = 'templateSummaryCost20250110.xlsx';
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
        $fileneme = "{$strDateTimeNow}_numberOfBeds($year-$commissionDecisionsId).xml";

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
        $v = 'v7';
        $templateFileName = "PumpPgg_$v.xlsx";
        $templateFilePath = $path . DIRECTORY_SEPARATOR . $templateFileName;
        $templateFullFilepath = Storage::path($templateFilePath);
        $resultFileName = "Приложение №2.11 Шаблон плановые объёмы $v($year-$commissionDecisionsId).xlsx";
        $strDateTimeNow = date("Y-m-d-His");
        $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . '_' . $resultFileName;
        $fullResultFilepath = Storage::path($resultFilePath);

        $spreadsheet = $reportService->generate($templateFullFilepath, year: $year, commissionDecisionsId: $commissionDecisionsId);

        $writer = new Xlsx($spreadsheet);
        $writer->save($fullResultFilepath);
        return Storage::download($resultFilePath);
    }

    public function VitacorePlan(VitacorePlanReportService $reportService, int $year, int $commissionDecisionsId = null) {
        $protocolNumber = '';
        $protocolDate = '';
        if ($commissionDecisionsId) {
            $cd = CommissionDecision::find($commissionDecisionsId);
            $protocolDate = $cd->date->format('d.m.Y');
            $protocolNumber = $cd->number;
        }
        $protocolNumberForFileName = preg_replace('/[^a-zа-я\d.]/ui', '_', $protocolNumber);

        $path = 'xlsx';
        $resultFileName = 'vitacore-plan' . ($protocolNumber !== '' ? '(Protokol_№'.$protocolNumberForFileName.'ot'.$protocolDate.')' : '') . '.xlsx';
        $strDateTimeNow = date("Y-m-d-His");
        $resultFilePath = $path . DIRECTORY_SEPARATOR . $strDateTimeNow . ' ' . $resultFileName;
        $fullResultFilepath = Storage::path($resultFilePath);

        $spreadsheet = $reportService->generate(year: $year, commissionDecisionsId: $commissionDecisionsId);

        $writer = new Xlsx($spreadsheet);
        $writer->save($fullResultFilepath);
        return Storage::download($resultFilePath);
    }
}

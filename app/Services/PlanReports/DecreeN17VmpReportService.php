<?php
declare(strict_types=1);

namespace App\Services\PlanReports;

use App\Models\CareProfiles;
use App\Models\ChangePackage;
use App\Models\CommissionDecision;
use App\Models\MedicalInstitution;
use App\Models\VmpGroup;
use App\Models\VmpTypes;
use App\Services\DataForContractService;
use App\Services\InitialDataFixingService;
use App\Services\PlannedIndicatorChangeInitService;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Выгрузка в vitacore для формирования отчета по приказу №17
 */
class DecreeN17VmpReportService {

    public function __construct(
        private DataForContractService $dataForContractService,
        private InitialDataFixingService $initialDataFixingService,
        private PlannedIndicatorChangeInitService $plannedIndicatorChangeInitService
    ){ }

    function printTableHeader(
            Worksheet $sheet,
            int $colIndex,
            int $rowIndex,
    ) {
        $this->printRow($sheet, $colIndex, $rowIndex, 'Код МО', 'МО','Профиль', 'Номер профиля','номер группы ВМП','объемы, госпитализаций','финансовое обеспечение, руб.');
    }

    function printRow(
        Worksheet $sheet,
        int $colIndex,
        int $rowIndex,
        string | int $moCode,
        string $moShortName,
        string $careProfileName,
        string $careProfileN17Code,
        string $vmpGroupCode,
        string $vol,
        string $fin
    ) {
        $moCodeColOffset = 0;
        $moShortNameColOffset = 1;
        $careProfileNameOffset = 2;
        $careProfileN17CodeOffset = 3;
        $vmpGroupCodeOffset = 4;
        $volOffset = 5;
        $finOffset = 6;

        $sheet->setCellValue([$colIndex + $moCodeColOffset, $rowIndex], $moCode);
        $sheet->setCellValue([$colIndex + $moShortNameColOffset, $rowIndex], $moShortName);
        $sheet->setCellValue([$colIndex + $careProfileNameOffset, $rowIndex], $careProfileName);
        $sheet->setCellValue([$colIndex + $careProfileN17CodeOffset, $rowIndex], $careProfileN17Code);
        $sheet->setCellValue([$colIndex + $vmpGroupCodeOffset, $rowIndex], $vmpGroupCode);
        $sheet->setCellValue([$colIndex + $volOffset, $rowIndex], $vol);
        $sheet->setCellValue([$colIndex + $finOffset, $rowIndex], $fin);
    }

    public function generate(int $year, int|null $commissionDecisionsId = null) : Spreadsheet {

        $packageIds = null;
        $currentlyUsedDate = null;
        $currentlyUsedDateString = $year.'-01-01';
        $docName = "";
        if ($commissionDecisionsId) {
            $commissionDecisions = CommissionDecision::whereYear('date',$year)->where('id', '<=', $commissionDecisionsId)->get();
            $cd = $commissionDecisions->find($commissionDecisionsId);
            $commissionDecisionIds = $commissionDecisions->pluck('id')->toArray();
            $packageIds = ChangePackage::whereIn('commission_decision_id', $commissionDecisionIds)->orWhere('commission_decision_id', null)->pluck('id')->toArray();

            $currentlyUsedDate = $cd->date;
            $currentlyUsedDateString = $currentlyUsedDate->format('Y-m-d');
        } else {
            if ($this->initialDataFixingService->fixedYear($year)) {
                $packageIds = ChangePackage::where('commission_decision_id', null)->pluck('id')->toArray();
            } else {
                $this->plannedIndicatorChangeInitService->fromInitialData($year);
            }
        }


        bcscale(4);

        $moCollection = MedicalInstitution::WhereRaw("? BETWEEN effective_from AND effective_to", [$currentlyUsedDateString])->orderBy('order')->get();
        $content = $this->dataForContractService->GetArray($year, $packageIds);

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        $vmpGroups = VmpGroup::all();
        $vmpTypes = VmpTypes::all();
        $careProfiles = CareProfiles::orderBy('id')->get();
        $rowIndex = 7;
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

        $rowIndex = 1;
        $colIndex = 1;
        $this->printTableHeader($sheet, $colIndex, $rowIndex++);

        foreach($moCollection as $mo) {

            if(!isset($content['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'])) {
                continue;
            }

            foreach($vmp as $careProfileId => $groups) {
                $cp = $careProfiles->firstWhere('id', $careProfileId)->decreeN17($currentlyUsedDate);
                $careProfileName = $cp->name;
                $careProfileN17Code = $cp->code;

                foreach ($vmpGroups as $vmpGroup) {
                    if(!isset($groups[$vmpGroup->id])){
                        continue;
                    }
                    $vVol = '0';
                    $vFin = '0';
                    foreach ($vmpTypes as $vmpType) {
                        if(!in_array($vmpType->id, $groups[$vmpGroup->id])){
                            continue;
                        }

                        $v = $content['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'][$careProfileId][$vmpGroup->id][$vmpType->id] ?? [];
                        $vVol = bcadd($vVol, $v[$indicatorId] ?? '0');
                        $vFin = bcadd($vFin, $v[$indicatorCostId] ?? '0');
                    }
                    if (bccomp($vVol, '0') === 1 || bccomp($vFin, '0') === 1) {
                        $this->printRow($sheet, $colIndex, $rowIndex++, $mo->code, $mo->short_name, $careProfileName, $careProfileN17Code, $vmpGroup->code, $vVol, $vFin);
                    }
                }
            }
        }

        return $spreadsheet;
    }
}

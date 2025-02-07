<?php
declare(strict_types=1);

namespace App\Services\PlanReports;

use App\Models\ChangePackage;
use App\Models\CommissionDecision;
use App\Models\MedicalInstitution;
use App\Models\MedicalServices;
use App\Services\DataForContractService;
use App\Services\MedicalServicesService;
use App\Services\PeopleAssignedInfoForContractService;
use App\Services\RehabilitationProfileService;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class VitacorePlanReportService {

    public function __construct(
        private DataForContractService $dataForContractService,
        private PeopleAssignedInfoForContractService $peopleAssignedInfoForContractService,
        private MedicalServicesService $medicalServicesService
    )
    { }

    public function generate(int $year, int $commissionDecisionsId = null) : Spreadsheet {
        $packageIds = null;
        $currentlyUsedDate = $year.'-01-01';
        if ($commissionDecisionsId) {
            $commissionDecisions = CommissionDecision::whereYear('date',$year)->where('id', '<=', $commissionDecisionsId)->get();
            $cd = $commissionDecisions->find($commissionDecisionsId);
            $commissionDecisionIds = $commissionDecisions->pluck('id')->toArray();
            $packageIds = ChangePackage::whereIn('commission_decision_id', $commissionDecisionIds)->orWhere('commission_decision_id', null)->pluck('id')->toArray();
            $currentlyUsedDate = $cd->date->format('Y-m-d');
        } else {
            $packageIds = ChangePackage::where('commission_decision_id', null)->pluck('id')->toArray();
        }

        bcscale(4);

        $contentByMonth = [];
        for($monthNum = 1; $monthNum <= 12; $monthNum++)
        {
            $contentByMonth[$monthNum] = $this->dataForContractService->GetArrayByYearAndMonth($year, $monthNum, $packageIds);
        }

        $peopleAssigned = $this->peopleAssignedInfoForContractService->GetArray($year, $packageIds);
        $moCollection = MedicalInstitution::WhereRaw("? BETWEEN effective_from AND effective_to", [$currentlyUsedDate])->orderBy('order')->get();

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue([1, 1], "Год:");
        $sheet->setCellValue([2, 1], "$year");

        $ordinalRowNum = 1;
        $firstTableHeadRowIndex = 2;
        $firstTableDataRowIndex = 3;
        $firstTableColIndex = 1;
        $rowOffset = 0;

        vitacoreV2PrintTableHeader($sheet, $firstTableColIndex, $firstTableHeadRowIndex);

        foreach($moCollection as $mo) {
            $planningSectionName = "Скорая помощь";
            $planningParamName = "объемы, вызовов";
            $category = 'ambulance';
            $indicatorId = 5; // вызовов
            $callsAssistanceTypeId = 5; // вызовы
            $thrombolysisAssistanceTypeId = 6;// тромболизис

            $thrombolysis = [];
            $calls = [];
            $hasValue = false;
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $thrombolysis[$monthNum] = $contentByMonth[$monthNum]['mo'][$mo->id][$category][$thrombolysisAssistanceTypeId][$indicatorId] ?? '0';
                $calls[$monthNum] = $contentByMonth[$monthNum]['mo'][$mo->id][$category][$callsAssistanceTypeId][$indicatorId] ?? '0';
                if( bccomp($thrombolysis[$monthNum], '0') !== 0
                    || bccomp($calls[$monthNum], '0') !== 0
                ) {
                    $hasValue = true;
                }
            }
            if($hasValue) {
                $values = [];
                for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                    $values[$monthNum] = bcadd($calls[$monthNum], $thrombolysis[$monthNum]);
                }
                vitacoreV2PrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset,
                    $ordinalRowNum,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $planningParamName,
                    $values
                );
                $rowOffset++;
                $ordinalRowNum++;
            }

            // 2.обращения по заболеваниям
            $planningSectionName = "Амбулаторная помощь в связи с заболеваниями";
            $planningParamName = "объемы, обращений";
            $hasValue = false;
            $category = 'polyclinic';
            $indicatorId = 8; // обращений
            $assistanceTypeId = 4; // обращения по заболеваниям
            $perPerson = [];
            $perUnit = [];
            $fap = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $perPerson[$monthNum] = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0';
                $perUnit[$monthNum] = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0';
                $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
                $fap[$monthNum] = '0';
                foreach ($faps as $f) {
                    $fap[$monthNum] = bcadd($fap[$monthNum], $f['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0');
                }
                if( bccomp($perPerson[$monthNum],'0') !== 0
                    || bccomp($perUnit[$monthNum],'0') !== 0
                    || bccomp($fap[$monthNum],'0') !== 0
                ) {
                    $hasValue = true;
                }
            }
            if($hasValue) {
                $values = [];
                for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                    $values[$monthNum] = bcadd(bcadd($perPerson[$monthNum], $perUnit[$monthNum]), $fap[$monthNum]);
                }
                vitacoreV2PrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset,
                    $ordinalRowNum,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $planningParamName,
                    $values
                );
                $rowOffset++;
                $ordinalRowNum++;
            }


            // фин.обесп.
            $category = 'polyclinic';
            $assistanceTypeId = 4; //обращения по заболеваниям
            $indicatorId = 4; // стоимость

            // не заполняется по медицинским организациям имеющим прикрепленное население
            if (!($peopleAssigned[$mo->id][$category]['mo']['peopleAssigned'] ?? false)) {
                $planningSectionName = "Амбулаторная помощь в связи с заболеваниями";
                $planningParamName = "финансовое обеспечение, руб.";
                $hasValue = false;
                $perPerson = [];
                $perUnit = [];
                $fap = [];
                for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                    $perPerson[$monthNum] = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0';
                    $perUnit[$monthNum] = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0';
                    $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
                    $fap[$monthNum] = '0';
                    foreach ($faps as $f) {
                        $fap[$monthNum] = bcadd($fap[$monthNum], $f['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0');
                    }
                    if( bccomp($perPerson[$monthNum],'0') !== 0
                        || bccomp($perUnit[$monthNum],'0') !== 0
                        || bccomp($fap[$monthNum],'0') !== 0
                    ) {
                        $hasValue = true;
                    }
                }
                if($hasValue) {
                    $values = [];
                    for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                        $values[$monthNum] = bcadd(bcadd($perPerson[$monthNum], $perUnit[$monthNum]), $fap[$monthNum]);
                    }
                    vitacoreV2PrintRow(
                        $sheet,
                        $firstTableColIndex,
                        $firstTableDataRowIndex + $rowOffset,
                        $ordinalRowNum,
                        $mo->code,
                        $mo->short_name,
                        $planningSectionName,
                        $planningParamName,
                        $values
                    );
                    $rowOffset++;
                    $ordinalRowNum++;
                }
            }

            // 3.Посещения с иными целями
            $planningSectionName = "Амбулаторная помощь с профилактическими и иными целями";
            $planningParamName = "объемы, посещений";
            $hasValue = false;
            $category = 'polyclinic';
            $indicatorId = 9; // посещений
            $assistanceTypeIds = [1, 2]; //	посещения профилактические, посещения разовые по заболеваниям
            $perPerson = [];
            $perUnit = [];
            $fap = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $perPerson[$monthNum] = '0';
                $perUnit[$monthNum] = '0';
                $fap[$monthNum] = '0';
                $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
                foreach($assistanceTypeIds as $assistanceTypeId) {
                    $perPerson[$monthNum] = bcadd($perPerson[$monthNum], $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0');
                    $perUnit[$monthNum]   = bcadd($perUnit[$monthNum],   $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0');
                    foreach ($faps as $f) {
                        $fap[$monthNum] = bcadd($fap[$monthNum], $f['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0');
                    }
                }
                if( bccomp($perPerson[$monthNum], '0') !== 0
                    || bccomp($perUnit[$monthNum], '0') !== 0
                    || bccomp($fap[$monthNum], '0') !== 0
                ) {
                    $hasValue = true;
                }
            }
            if($hasValue) {
                for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                    $values[$monthNum] = bcadd(bcadd($perPerson[$monthNum], $perUnit[$monthNum]), $fap[$monthNum]);
                }
                vitacoreV2PrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset,
                    $ordinalRowNum,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $planningParamName,
                    $values
                );
                $rowOffset++;
                $ordinalRowNum++;
            }

            // фин.обесп.
            $category = 'polyclinic';
            $assistanceTypeIds = [1, 2]; //	посещения профилактические, посещения разовые по заболеваниям
            $indicatorId = 4; // стоимость

            // не заполняется по медицинским организациям имеющим прикрепленное население

            if (!($peopleAssigned[$mo->id][$category]['mo']['peopleAssigned'] ?? false)) {
                $planningSectionName = "Амбулаторная помощь с профилактическими и иными целями";
                $planningParamName = "финансовое обеспечение, руб.";
                $hasValue = false;
                $perPerson = [];
                $perUnit = [];
                $fap = [];
                for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                    $perPerson[$monthNum] = '0';
                    $perUnit[$monthNum] = '0';
                    $fap[$monthNum] = '0';
                    $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
                    foreach($assistanceTypeIds as $assistanceTypeId) {
                        $perPerson[$monthNum] = bcadd($perPerson[$monthNum], $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0');
                        $perUnit[$monthNum]   = bcadd($perUnit[$monthNum],   $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0');
                        foreach ($faps as $f) {
                            $fap[$monthNum] = bcadd($fap[$monthNum], $f['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0');
                        }
                    }
                    if( bccomp($perPerson[$monthNum],'0') !== 0
                        || bccomp($perUnit[$monthNum],'0') !== 0
                        || bccomp($fap[$monthNum],'0') !== 0
                    ) {
                        $hasValue = true;
                    }
                }
                if($hasValue) {
                    $values = [];
                    for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                        $values[$monthNum] = bcadd(bcadd($perPerson[$monthNum], $perUnit[$monthNum]), $fap[$monthNum]);
                    }
                    vitacoreV2PrintRow(
                        $sheet,
                        $firstTableColIndex,
                        $firstTableDataRowIndex + $rowOffset,
                        $ordinalRowNum,
                        $mo->code,
                        $mo->short_name,
                        $planningSectionName,
                        $planningParamName,
                        $values
                    );
                    $rowOffset++;
                    $ordinalRowNum++;
                }
            }

            // 4 Неотложная помощь
            $planningSectionName = "Амбулаторная помощь с неотложной целью";
            $planningParamName = "объемы, посещений";
            $hasValue = false;
            $category = 'polyclinic';
            $indicatorId = 9; // посещений
            $assistanceTypeIds = [3]; //	посещения неотложные
            $perPerson = [];
            $perUnit = [];
            $fap = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $perPerson[$monthNum] = '0';
                $perUnit[$monthNum] = '0';
                $fap[$monthNum] = '0';
                $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
                foreach($assistanceTypeIds as $assistanceTypeId) {
                    $perPerson[$monthNum] = bcadd($perPerson[$monthNum], $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0');
                    $perUnit[$monthNum] = bcadd($perUnit[$monthNum], $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0');
                    foreach ($faps as $f) {
                        $fap[$monthNum] = bcadd($fap[$monthNum], $f['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0');
                    }
                }
                if( bccomp($perPerson[$monthNum],'0') !== 0
                    || bccomp($perUnit[$monthNum],'0') !== 0
                    || bccomp($fap[$monthNum],'0') !== 0
                ) {
                    $hasValue = true;
                }
            }
            if($hasValue) {
                for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                    $values[$monthNum] = bcadd(bcadd($perPerson[$monthNum], $perUnit[$monthNum]), $fap[$monthNum]);
                }
                vitacoreV2PrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset,
                    $ordinalRowNum,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $planningParamName,
                    $values
                );
                $rowOffset++;
                $ordinalRowNum++;
            }
            // фин.обесп.
            $category = 'polyclinic';
            $assistanceTypeIds = [3]; // посещения неотложные
            $indicatorId = 4; // стоимость

            $planningSectionName = "Амбулаторная помощь с неотложной целью";
            $planningParamName = "финансовое обеспечение, руб.";
            $hasValue = false;
            $perPerson = [];
            $perUnit = [];
            $fap = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $perPerson[$monthNum] = '0';
                $perUnit[$monthNum] = '0';
                $fap[$monthNum] = '0';
                $faps = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['fap'] ?? [];
                foreach($assistanceTypeIds as $assistanceTypeId) {
                    $perPerson[$monthNum] = bcadd($perPerson[$monthNum], $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0');
                    $perUnit[$monthNum]   = bcadd($perUnit[$monthNum],   $contentByMonth[$monthNum]['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0');
                    foreach ($faps as $f) {
                        $fap[$monthNum] = bcadd($fap[$monthNum], $f['assistanceTypes'][$assistanceTypeId][$indicatorId] ?? '0');
                    }
                }
                if( bccomp($perPerson[$monthNum],'0') !== 0
                    || bccomp($perUnit[$monthNum],'0') !== 0
                    || bccomp($fap[$monthNum],'0') !== 0
                ) {
                    $hasValue = true;
                }
            }
            if($hasValue) {
                $values = [];
                for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                    $values[$monthNum] = bcadd(bcadd($perPerson[$monthNum], $perUnit[$monthNum]), $fap[$monthNum]);
                }
                vitacoreV2PrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset,
                    $ordinalRowNum,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $planningParamName,
                    $values
                );
                $rowOffset++;
                $ordinalRowNum++;
            }

            // Диагностические услуги
            /// список диагностических услуг актуальных на текущий год
            $medicalServiceIds = $this->medicalServicesService->getIdsByYear($year);
            $medicalServices = MedicalServices::whereIn('id', $medicalServiceIds)->orderBy('order')->get();

            foreach ($medicalServices as $ms) {
                $hasValue = false;

                $serviceId = $ms->id;
                $planningSectionName = \Illuminate\Support\Str::ucfirst($ms->name);
                $planningParamName = "объемы, услуг";

                $category = 'polyclinic';
                $indicatorId = 6; // услуг

                $values = [];
                for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                    $values[$monthNum] = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
                    if (!$hasValue && bccomp($values[$monthNum], '0') !== 0) {
                        $hasValue = true;
                    }
                }
                if($hasValue) {
                    vitacoreV2PrintRow(
                        $sheet,
                        $firstTableColIndex,
                        $firstTableDataRowIndex + $rowOffset,
                        $ordinalRowNum,
                        $mo->code,
                        $mo->short_name,
                        $planningSectionName,
                        $planningParamName,
                        $values
                    );
                    $rowOffset++;
                    $ordinalRowNum++;
                }


                $hasValue = false;
                $planningParamName = "финансовое обеспечение, руб.";
                $indicatorId = 4; // стоимость

                $values = [];
                for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                    $values[$monthNum] = PlanCalculatorService::medicalServicesSum($contentByMonth[$monthNum], $mo->id, $serviceId, $indicatorId, $category);
                    if (!$hasValue && bccomp($values[$monthNum],'0') !== 0) {
                        $hasValue = true;
                    }
                }

                if($hasValue) {
                    vitacoreV2PrintRow(
                        $sheet,
                        $firstTableColIndex,
                        $firstTableDataRowIndex + $rowOffset,
                        $ordinalRowNum,
                        $mo->code,
                        $mo->short_name,
                        $planningSectionName,
                        $planningParamName,
                        $values
                    );
                    $rowOffset++;
                    $ordinalRowNum++;
                }
            }

            // Круглосуточный ст. (не включая ВМП и медицинскую реабилитацию)
            $category = 'hospital';
            $indicatorId = 7; // госпитализаций

            $hasValue = false;
            $planningSectionName = "Круглосуточный стационар (не включая ВМП и медицинскую реабилитацию)";
            $planningParamName = "объемы, госпитализаций";

            $values = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $bedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? [];
                $values[$monthNum] = '0';
                foreach ($bedProfiles as $bpId => $bp) {
                    if (RehabilitationProfileService::IsRehabilitationBedProfile($bpId)) {
                        continue;
                    }
                    $values[$monthNum] = bcadd($values[$monthNum], $bp[$indicatorId] ?? '0');
                }
                if (bccomp($values[$monthNum],'0') !== 0) {
                    $hasValue = true;
                }
            }
            if($hasValue) {
                vitacoreV2PrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset,
                    $ordinalRowNum,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $planningParamName,
                    $values
                );
                $rowOffset++;
                $ordinalRowNum++;
            }

            // 6.ВМП
            $category = 'hospital';
            $indicatorId = 7; // госпитализаций
            $hasValue = false;
            $planningSectionName = "ВМП";
            $planningParamName = "объемы, госпитализаций";

            $values = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $careProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? [];
                $values[$monthNum] = '0';
                foreach ($careProfiles as $vmpGroups) {
                    foreach ($vmpGroups as $vmpTypes) {
                        foreach ($vmpTypes as $vmpT) {
                            $values[$monthNum] = bcadd($values[$monthNum], $vmpT[$indicatorId] ?? 0);
                        }
                    }
                }
                if (bccomp($values[$monthNum],'0') !== 0) {
                    $hasValue = true;
                }
            }
            if($hasValue) {
                vitacoreV2PrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset,
                    $ordinalRowNum,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $planningParamName,
                    $values
                );
                $rowOffset++;
                $ordinalRowNum++;
            }

            // Медреабилитация в КС
            $category = 'hospital';
            $indicatorId = 7; // госпитализаций
            $hasValue = false;
            $planningSectionName = "Медицинская реабилитация";
            $planningParamName = "объемы, госпитализаций";

            $values = [];
            $rehabilitationBedProfileIds = RehabilitationProfileService::GetAllRehabilitationBedProfileIds();
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $values[$monthNum] = '0';
                foreach ($rehabilitationBedProfileIds as $rbpId) {
                    $values[$monthNum] = bcadd($values[$monthNum], $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'][$rbpId][$indicatorId] ?? '0');
                }
                if (bccomp($values[$monthNum],'0') !== 0) {
                    $hasValue = true;
                }
            }
            if($hasValue) {
                vitacoreV2PrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset,
                    $ordinalRowNum,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $planningParamName,
                    $values
                );
                $rowOffset++;
                $ordinalRowNum++;
            }
            // 8. Дневные стационары
            $category = 'hospital';
            $indicatorId = 2; // случаев лечения
            $hasValue = false;
            $planningSectionName = "Дневные стационары";
            $planningParamName = "объемы, случаев лечения";

            $values = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $values[$monthNum] = '0';
                $bedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? [];
                foreach ($bedProfiles as $bp) {
                    $values[$monthNum] = bcadd($values[$monthNum], $bp[$indicatorId] ?? '0');
                }

                $bedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? [];
                foreach ($bedProfiles as $bp) {
                    $values[$monthNum] = bcadd($values[$monthNum], $bp[$indicatorId] ?? '0');
                }

                if (bccomp($values[$monthNum], '0') !== 0) {
                    $hasValue = true;
                }
            }
            if($hasValue) {
                vitacoreV2PrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset,
                    $ordinalRowNum,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $planningParamName,
                    $values
                );
                $rowOffset++;
                $ordinalRowNum++;
            }

            // 3. ДС, фин.обеспечение
            $category = 'hospital';
            $indicatorId = 4; // стоимость
            $hasValue = false;
            $planningSectionName = "Дневные стационары";
            $planningParamName = "финансовое обеспечение, руб.";

            $values = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $values[$monthNum] = '0';
                $bedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? [];
                foreach ($bedProfiles as $bp) {
                    $values[$monthNum] = bcadd($values[$monthNum], $bp[$indicatorId] ?? '0');
                }

                $bedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? [];
                foreach ($bedProfiles as $bp) {
                    $values[$monthNum] = bcadd($values[$monthNum], $bp[$indicatorId] ?? '0');
                }
                if (bccomp($values[$monthNum],'0') !== 0) {
                    $hasValue = true;
                }
            }
            if($hasValue) {
                vitacoreV2PrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset,
                    $ordinalRowNum,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $planningParamName,
                    $values
                );
                $rowOffset++;
                $ordinalRowNum++;
            }
            // 4 КС, фин.обеспечение (не включая мед.реабилитацию и ВМП)
            $category = 'hospital';
            $indicatorId = 4; // стоимость
            $hasValue = false;
            $planningSectionName = "Круглосуточный стационар (не включая ВМП и медицинскую реабилитацию)";
            $planningParamName = "финансовое обеспечение, руб.";

            $values = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $values[$monthNum] = '0';
                $bedProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? [];
                foreach ($bedProfiles as $bpId => $bp) {
                    if (RehabilitationProfileService::IsRehabilitationBedProfile($bpId)) {
                        continue;
                    }
                    $values[$monthNum] = bcadd($values[$monthNum], $bp[$indicatorId] ?? '0');
                }
                if (bccomp($values[$monthNum],'0') !== 0) {
                    $hasValue = true;
                }
            }
            if($hasValue) {
                vitacoreV2PrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset,
                    $ordinalRowNum,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $planningParamName,
                    $values
                );
                $rowOffset++;
                $ordinalRowNum++;
            }
            // 5 МР, фин.обеспечение
            $category = 'hospital';
            $indicatorId = 4; // стоимость
            $hasValue = false;
            $planningSectionName = "Медицинская реабилитация";
            $planningParamName = "финансовое обеспечение, руб.";

            $values = [];
            $rehabilitationBedProfileIds = RehabilitationProfileService::GetAllRehabilitationBedProfileIds();
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $values[$monthNum] = '0';
                foreach ($rehabilitationBedProfileIds as $rbpId) {
                    bcadd($values[$monthNum], $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'][$rbpId][$indicatorId] ?? '0');
                }
                if (bccomp($values[$monthNum],'0') !== 0) {
                    $hasValue = true;
                }
            }
            if($hasValue) {
                vitacoreV2PrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset,
                    $ordinalRowNum,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $planningParamName,
                    $values
                );
                $rowOffset++;
                $ordinalRowNum++;
            }
            // 6 ВМП, фин.обеспечение
            $category = 'hospital';
            $indicatorId = 4; // стоимость
            $hasValue = false;
            $planningSectionName = "ВМП";
            $planningParamName = "финансовое обеспечение, руб.";

            $values = [];
            for($monthNum = 1; $monthNum <= 12; $monthNum++) {
                $values[$monthNum] = '0';
                $careProfiles = $contentByMonth[$monthNum]['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? [];
                foreach ($careProfiles as $vmpGroups) {
                    foreach ($vmpGroups as $vmpTypes) {
                        foreach ($vmpTypes as $vmpT) {
                            $values[$monthNum] = bcadd($values[$monthNum], $vmpT[$indicatorId] ?? '0');
                        }
                    }
                }
                if (bccomp($values[$monthNum],'0') !== 0) {
                    $hasValue = true;
                }
            }
            if($hasValue) {
                vitacoreV2PrintRow(
                    $sheet,
                    $firstTableColIndex,
                    $firstTableDataRowIndex + $rowOffset,
                    $ordinalRowNum,
                    $mo->code,
                    $mo->short_name,
                    $planningSectionName,
                    $planningParamName,
                    $values
                );
                $rowOffset++;
                $ordinalRowNum++;
            }

        } // foreach MO

        return $spreadsheet;
    }
}

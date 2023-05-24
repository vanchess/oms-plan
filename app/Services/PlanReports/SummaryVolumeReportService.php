<?php
declare(strict_types=1);

namespace App\Services\PlanReports;

use App\Enum\MedicalServicesEnum;
use App\Enum\RehabilitationBedOptionEnum;
use App\Models\CareProfiles;
use App\Models\ChangePackage;
use App\Models\CommissionDecision;
use App\Models\MedicalInstitution;
use App\Models\VmpGroup;
use App\Models\VmpTypes;
use App\Services\DataForContractService;
use App\Services\InitialDataFixingService;
use App\Services\PeopleAssignedInfoForContractService;
use App\Services\PlannedIndicatorChangeInitService;
use App\Services\RehabilitationProfileService;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class SummaryVolumeReportService {
    use SummaryReportTrait;
    use HospitalSheetTrait;

    public function __construct(
        private DataForContractService $dataForContractService,
        private PeopleAssignedInfoForContractService $peopleAssignedInfoForContractService,
        private InitialDataFixingService $initialDataFixingService,
        private PlannedIndicatorChangeInitService $plannedIndicatorChangeInitService)
    { }

    private function fillPolyclinicSheet(Worksheet $sheet, $content, $contentByMonth, $peopleAssigned, $moCollection, int $startRow, int $serviceId, int $indicatorId, string $category = 'polyclinic', $endRow = 100) {
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
        $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);
        $this->fillSummaryRow($sheet, $rowIndex+1, 7, 20, $startRow, $rowIndex);
    }



    public function generate(string $templateFullFilepath, int $year, int $commissionDecisionsId = null) : Spreadsheet {
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

        $sheet = $spreadsheet->getSheetByName('1.Скорая помощь');
        $sheet->setCellValue([21, 2], $docName);
        $sheet->setCellValue([1, 3], "Скорая помощь, плановые объемы на $year год");
        $sheet->setCellValue([7, 4], "Численность прикрепленного населения на 01.01.$year");
        $ordinalRowNum = 0;
        $startRow = 7;
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
        $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);
        $this->fillSummaryRow($sheet, $rowIndex+1, 7, 21, $startRow, $rowIndex);


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
        $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);
        $this->fillSummaryRow($sheet, $rowIndex+1, 7, 20, $startRow, $rowIndex);


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
        $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);
        $this->fillSummaryRow($sheet, $rowIndex+1, 7, 20, $startRow, $rowIndex);


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
        $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);
        $this->fillSummaryRow($sheet, $rowIndex+1, 7, 20, $startRow, $rowIndex);

        $servicesIndicatorId = 6; // услуг
        $sheet = $spreadsheet->getSheetByName('2.2 КТ');
        $serviceId = MedicalServicesEnum::KT;
        $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в связи с заболеваниями в амбулаторных условиях на $year год (компьютерная томография)");
        $sheet->setCellValue([7, 4], "Численность прикрепленного населения на 01.01.$year");
        $this->fillPolyclinicSheet($sheet, $content, $contentByMonth, $peopleAssigned, $moCollection, $startRow, $serviceId, indicatorId: $servicesIndicatorId, endRow: $endRow);

        $sheet = $spreadsheet->getSheetByName('2.3 МРТ');
        $serviceId = MedicalServicesEnum::MRT;
        $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в связи с заболеваниями в амбулаторных условиях на $year год (магнитно-резонансная томография)");
        $sheet->setCellValue([7, 4], "Численность прикрепленного населения на 01.01.$year");
        $this->fillPolyclinicSheet($sheet, $content, $contentByMonth, $peopleAssigned, $moCollection, $startRow, $serviceId, indicatorId: $servicesIndicatorId, endRow: $endRow);

        $sheet = $spreadsheet->getSheetByName('2.4 УЗИ ССС');
        $serviceId = MedicalServicesEnum::UltrasoundCardio;
        $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в связи с заболеваниями в амбулаторных условиях на $year год (УЗИ сердечно-сосудистой системы)");
        $sheet->setCellValue([7, 4], "Численность прикрепленного населения на 01.01.$year");
        $this->fillPolyclinicSheet($sheet, $content, $contentByMonth, $peopleAssigned, $moCollection, $startRow, $serviceId, indicatorId: $servicesIndicatorId, endRow: $endRow);

        $sheet = $spreadsheet->getSheetByName('2.5 Эндоскопия');
        $serviceId = MedicalServicesEnum::Endoscopy;
        $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в связи с заболеваниями в амбулаторных условиях на $year год (Эндоскопические исследования)");
        $sheet->setCellValue([7, 4], "Численность прикрепленного населения на 01.01.$year");
        $this->fillPolyclinicSheet($sheet, $content, $contentByMonth, $peopleAssigned, $moCollection, $startRow, $serviceId, indicatorId: $servicesIndicatorId, endRow: $endRow);

        $sheet = $spreadsheet->getSheetByName('2.6 ПАИ');
        $serviceId = MedicalServicesEnum::PathologicalAnatomicalBiopsyMaterial;
        $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в связи с заболеваниями в амбулаторных условиях на $year год (Паталого анатомическое исследование биопсийного материала с целью диагностики онкологических заболеваний)");
        $sheet->setCellValue([7, 4], "Численность прикрепленного населения на 01.01.$year");
        $this->fillPolyclinicSheet($sheet, $content, $contentByMonth, $peopleAssigned, $moCollection, $startRow, $serviceId, indicatorId: $servicesIndicatorId, endRow: $endRow);

        $sheet = $spreadsheet->getSheetByName('2.7 МГИ');
        $serviceId = MedicalServicesEnum::MolecularGeneticDetectionOncological;
        $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в связи с заболеваниями в амбулаторных условиях на $year год (Малекулярно-генетические исследования с целью диагностики онкологических заболеваний)");
        $sheet->setCellValue([7, 4], "Численность прикрепленного населения на 01.01.$year");
        $this->fillPolyclinicSheet($sheet, $content, $contentByMonth, $peopleAssigned, $moCollection, $startRow, $serviceId, indicatorId: $servicesIndicatorId, endRow: $endRow);

        $sheet = $spreadsheet->getSheetByName('2.8  Тест.covid-19');
        $serviceId = MedicalServicesEnum::CovidTesting;
        $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в связи с заболеваниями в амбулаторных условиях на $year год (Тестирование на выявление covid-19)");
        $sheet->setCellValue([7, 4], "Численность прикрепленного населения на 01.01.$year");
        $this->fillPolyclinicSheet($sheet, $content, $contentByMonth, $peopleAssigned, $moCollection, $startRow, $serviceId, indicatorId: $servicesIndicatorId, endRow: $endRow);

        $sheet = $spreadsheet->getSheetByName('3.3 УЗИ плода');
        $serviceId = MedicalServicesEnum::FetalUltrasound;
        $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в амбулаторных условиях на $year год, УЗИ плода (1 триместр)");
        $sheet->setCellValue([7, 4], "Численность прикрепленного населения на 01.01.$year");
        $this->fillPolyclinicSheet($sheet, $content, $contentByMonth, $peopleAssigned, $moCollection, $startRow, $serviceId, indicatorId: $servicesIndicatorId, endRow: $endRow);

        $sheet = $spreadsheet->getSheetByName('3.4 Компл.иссл. репрод.орг.');
        $serviceId = MedicalServicesEnum::DiagnosisBackgroundPrecancerousDiseasesReproductiveWomen;
        $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в амбулаторных условиях на $year год, комплексное исследование для диагностики фоновых и предраковых заболеваний репродуктивных органов у женщин");
        $sheet->setCellValue([7, 4], "Численность прикрепленного населения на 01.01.$year");
        $this->fillPolyclinicSheet($sheet, $content, $contentByMonth, $peopleAssigned, $moCollection, $startRow, $serviceId, indicatorId: $servicesIndicatorId, endRow: $endRow);

        $sheet = $spreadsheet->getSheetByName('3.5 Опред.антигена D');
        $serviceId = MedicalServicesEnum::DeterminationAntigenD;
        $sheet->setCellValue([2, 3], "Плановые объемы медицинской помощи в амбулаторных условиях на $year год, определение антигена D системы Резус (резус-фактор плода)");
        $sheet->setCellValue([7, 4], "Численность прикрепленного населения на 01.01.$year");
        $this->fillPolyclinicSheet($sheet, $content, $contentByMonth, $peopleAssigned, $moCollection, $startRow, $serviceId, indicatorId: $servicesIndicatorId, endRow: $endRow);


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

        return $spreadsheet;
    }
}

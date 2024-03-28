<?php
declare(strict_types=1);

namespace App\Services\PlanReports;

use App\Enum\MedicalServicesEnum;
use App\Models\CareProfilesFoms;
use App\Models\Category;
use App\Models\CategoryTreeNodes;
use App\Models\CommissionDecision;
use App\Models\Indicator;
use App\Models\IndicatorType;
use App\Models\MedicalAssistanceType;
use App\Models\MedicalInstitution;
use App\Models\MedicalServices;
use App\Models\PlannedIndicator;
use App\Services\DataForContractService;
use App\Services\NodeService;
use Illuminate\Support\Facades\Storage;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class MeetingMinutesReportService
{
    public function __construct(
        private DataForContractService $dataForContractService,
        private NodeService $nodeService
    ){ }

    public function generate(string $templateFullFilepath, int $year, int $commissionDecisionsId) : Spreadsheet {
        $cd = CommissionDecision::find($commissionDecisionsId);
        $currentlyUsedDate = $cd->date->format('Y-m-d');
        $protocolDate = $cd->date->format('d.m.Y');
        $docName = "протокол заседания КРТП ОМС №$cd->number от $protocolDate";
        $packageIds = $cd->changePackage()->pluck('id')->toArray();
        $indicatorIds = [1, 2, 3, 4, 5, 6, 7, 8, 9];
        $content = $this->dataForContractService->GetArray($year, $packageIds, $indicatorIds);
        // количество коек на последний месяц года
        $contentNumberOfBeds = $this->dataForContractService->GetArrayByYearAndMonth($year, 12, $packageIds, [1]);

        $moCollection = MedicalInstitution::WhereRaw("? BETWEEN effective_from AND effective_to", [$currentlyUsedDate])->orderBy('order')->get();

        $startRow = 6;
        $endRow = 350;
        bcscale(4);


        $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader("Xlsx");
        $spreadsheet = $reader->load($templateFullFilepath);
        $sheet = $spreadsheet->getSheetByName('Скорая помощь');
        $sheet->setCellValue([1,3], $docName);
        $ordinalRowNum = 0;
        $rowIndex = $startRow - 1;
        $category = 'ambulance';
        $callsIndicatorId = 5; // вызовов
        $costIndicatorId = 4; // стоимость
        $callsAssistanceTypeId = 5; // вызовы
        $thrombolysisAssistanceTypeId = 6;// тромболизис
        foreach($moCollection as $mo) {
            $thrombolysis = $content['mo'][$mo->id][$category][$thrombolysisAssistanceTypeId][$callsIndicatorId] ?? 0;
            $allCalls = ($content['mo'][$mo->id][$category][$callsAssistanceTypeId][$callsIndicatorId] ?? 0) + $thrombolysis;
            $costCalls = $content['mo'][$mo->id][$category][$callsAssistanceTypeId][$costIndicatorId] ?? '0';
            $costThrombolysis = $content['mo'][$mo->id][$category][$thrombolysisAssistanceTypeId][$costIndicatorId] ?? '0';

            if( $thrombolysis === 0 && $allCalls === 0 && $costCalls === '0' && $costThrombolysis === '0' ) {continue;}

            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            $sheet->setCellValue([2,$rowIndex], $mo->short_name);
            $sheet->setCellValue([3,$rowIndex], $allCalls);
            //$sheet->setCellValue([9,$rowIndex], $thrombolysis);
            $sheet->setCellValue([4,$rowIndex], bcadd($costCalls, $costThrombolysis));
        }
        $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);




        $sheet = $spreadsheet->getSheetByName('Диагностика');
        $sheet->setCellValue([1,3], $docName);
        $ordinalRowNum = 0;
        $rowIndex = $startRow;
        $category = 'polyclinic';
        $indicatorId = 6; // услуг
        $costIndicatorId = 4; // стоимость

        foreach($moCollection as $mo) {
            // КТ
            $serviceId = MedicalServicesEnum::KT;
            $kt = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $indicatorId);
            $costKt = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $costIndicatorId);

            // МРТ
            $serviceId = MedicalServicesEnum::MRT;
            $mrt = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $indicatorId);
            $costMrt = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $costIndicatorId);

            // УЗИ ССС
            $serviceId = MedicalServicesEnum::UltrasoundCardio;
            $ultrasoundCardio = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $indicatorId);
            $costUltrasoundCardio = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $costIndicatorId);

            // Эндоскопия
            $serviceId = MedicalServicesEnum::Endoscopy;
            $endoscopy = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $indicatorId);
            $costEndoscopy = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $costIndicatorId);

            // ПАИ
            $serviceId = MedicalServicesEnum::PathologicalAnatomicalBiopsyMaterial;
            $pathologicalAnatomicalBiopsyMaterial = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $indicatorId);
            $costPathologicalAnatomicalBiopsyMaterial = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $costIndicatorId);

            // МГИ
            $serviceId = MedicalServicesEnum::MolecularGeneticDetectionOncological;
            $molecularGeneticDetectionOncological = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $indicatorId);
            $costMolecularGeneticDetectionOncological = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $costIndicatorId);

            // Тест.covid-19
            $serviceId = MedicalServicesEnum::CovidTesting;
            $covidTesting = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $indicatorId);
            $costCovidTesting = PlanCalculatorService::medicalServicesSum($content, $mo->id, $serviceId, $costIndicatorId);

            if( bccomp($kt,'0') === 0
                && bccomp($mrt, '0') === 0
                && bccomp($ultrasoundCardio, '0') === 0
                && bccomp($endoscopy, '0') === 0
                && bccomp($pathologicalAnatomicalBiopsyMaterial, '0') === 0
                && bccomp($molecularGeneticDetectionOncological, '0') === 0
                && bccomp($covidTesting, '0') === 0
                && bccomp($costKt, '0') === 0
                && bccomp($costMrt, '0') === 0
                && bccomp($costEndoscopy, '0') === 0
                && bccomp($costPathologicalAnatomicalBiopsyMaterial, '0') === 0
                && bccomp($costMolecularGeneticDetectionOncological, '0') === 0
                && bccomp($costCovidTesting, '0') === 0
            ) {continue;}

            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            $sheet->setCellValue([2,$rowIndex], $mo->short_name);
            $sheet->setCellValue([3,$rowIndex], $kt);
            $sheet->setCellValue([4,$rowIndex], $costKt);
            $sheet->setCellValue([5,$rowIndex], $mrt);
            $sheet->setCellValue([6,$rowIndex], $costMrt);
            $sheet->setCellValue([7,$rowIndex], $endoscopy);
            $sheet->setCellValue([8,$rowIndex], $costEndoscopy);
            $sheet->setCellValue([9,$rowIndex], $ultrasoundCardio);
            $sheet->setCellValue([10,$rowIndex], $costUltrasoundCardio);
            $sheet->setCellValue([11,$rowIndex], $pathologicalAnatomicalBiopsyMaterial);
            $sheet->setCellValue([12,$rowIndex], $costPathologicalAnatomicalBiopsyMaterial);
            $sheet->setCellValue([13,$rowIndex], $molecularGeneticDetectionOncological);
            $sheet->setCellValue([14,$rowIndex], $costMolecularGeneticDetectionOncological);
            $sheet->setCellValue([15,$rowIndex], $covidTesting);
            $sheet->setCellValue([16,$rowIndex], $costCovidTesting);
        }
        $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);

        $careProfilesFoms = CareProfilesFoms::all();

        $sheet = $spreadsheet->getSheetByName('ДС при стационаре');
        $sheet->setCellValue([1,3], $docName);
        $ordinalRowNum = 0;
        $rowIndex = $startRow;
        $category = 'hospital';

        $numberOfBedsIndicatorId = 1; // число коек
        $casesOfTreatmentIndicatorId = 2; // случаев лечения
        $patientDaysIndicatorId = 3; // пациенто-дней
        $costIndicatorId = 4; // стоимость

        foreach($moCollection as $mo) {
            $moRowIndex = $rowIndex;
            $inHospitalBedProfiles = $content['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? null;
            $inHospitalBedProfilesNumberOfBeds = $contentNumberOfBeds['mo'][$mo->id][$category]['daytime']['inHospital']['bedProfiles'] ?? null;

            if (!$inHospitalBedProfiles) { continue; }

            $ordinalRowNum++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            $sheet->setCellValue([2,$rowIndex], $mo->short_name);
            foreach ($careProfilesFoms as $cpf) {
                $numberOfBeds = '0';
                $casesOfTreatment = '0';
                $patientDays = '0';
                $cost = '0';

                $hbp = $cpf->hospitalBedProfiles;
                foreach($hbp as $bp) {
                    $bpData = $inHospitalBedProfiles[$bp->id] ?? null;
                    if (!$bpData) { continue; }

                    $casesOfTreatment = bcadd($casesOfTreatment, $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                    $patientDays = bcadd($patientDays, $bpData[$patientDaysIndicatorId] ?? '0');
                    $cost = bcadd($cost, $bpData[$costIndicatorId] ?? '0');
                }
                foreach($hbp as $bp) {
                    $bpData = $inHospitalBedProfilesNumberOfBeds[$bp->id] ?? null;
                    if (!$bpData) { continue; }

                    $numberOfBeds = bcadd($numberOfBeds, $bpData[$numberOfBedsIndicatorId] ?? '0');
                }
                if( bccomp($numberOfBeds,'0') === 0
                    && bccomp($casesOfTreatment, '0') === 0
                    && bccomp($patientDays, '0') === 0
                    && bccomp($cost, '0') === 0
                ) {continue;}

                $sheet->setCellValue([3,$rowIndex], $cpf->name);
                $sheet->setCellValue([4,$rowIndex], $numberOfBeds);
                $sheet->setCellValue([5,$rowIndex], $casesOfTreatment);
                $sheet->setCellValue([6,$rowIndex], $patientDays);
                $sheet->setCellValue([7,$rowIndex], $cost);
                $rowIndex++;
            }
            if ($rowIndex - 1 > $moRowIndex) {
                $sheet->mergeCells([1, $moRowIndex, 1, $rowIndex - 1]);
                $sheet->mergeCells([2, $moRowIndex, 2, $rowIndex - 1]);
            }
        }
        $sheet->removeRow($rowIndex,$endRow-$rowIndex+1);


        $sheet = $spreadsheet->getSheetByName('ДС при поликлинике');
        $sheet->setCellValue([1,3], $docName);
        $ordinalRowNum = 0;
        $rowIndex = $startRow;
        $category = 'hospital';

        $numberOfBedsIndicatorId = 1; // число коек
        $casesOfTreatmentIndicatorId = 2; // случаев лечения
        $patientDaysIndicatorId = 3; // пациенто-дней
        $costIndicatorId = 4; // стоимость

        foreach($moCollection as $mo) {
            $moRowIndex = $rowIndex;
            $inPolyclinicBedProfiles = $content['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? null;
            $inPolyclinicBedProfilesNumberOfBeds = $contentNumberOfBeds['mo'][$mo->id][$category]['daytime']['inPolyclinic']['bedProfiles'] ?? null;

            if (!$inPolyclinicBedProfiles) { continue; }

            $ordinalRowNum++;
            // $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            $sheet->setCellValue([2,$rowIndex], $mo->short_name);
            foreach ($careProfilesFoms as $cpf) {
                $numberOfBeds = '0';
                $casesOfTreatment = '0';
                $patientDays = '0';
                $cost = '0';

                $hbp = $cpf->hospitalBedProfiles;
                foreach($hbp as $bp) {
                    $bpData = $inPolyclinicBedProfiles[$bp->id] ?? null;
                    if (!$bpData) { continue; }

                    $casesOfTreatment = bcadd($casesOfTreatment, $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                    $patientDays = bcadd($patientDays, $bpData[$patientDaysIndicatorId] ?? '0');
                    $cost = bcadd($cost, $bpData[$costIndicatorId] ?? '0');
                }
                foreach($hbp as $bp) {
                    $bpData = $inPolyclinicBedProfilesNumberOfBeds[$bp->id] ?? null;
                    if (!$bpData) { continue; }

                    $numberOfBeds = bcadd($numberOfBeds, $bpData[$numberOfBedsIndicatorId] ?? '0');
                }
                if( bccomp($numberOfBeds,'0') === 0
                    && bccomp($casesOfTreatment, '0') === 0
                    && bccomp($patientDays, '0') === 0
                    && bccomp($cost, '0') === 0
                ) {continue;}

                $sheet->setCellValue([3,$rowIndex], "$cpf->name");
                $sheet->setCellValue([4,$rowIndex], $numberOfBeds);
                $sheet->setCellValue([5,$rowIndex], $casesOfTreatment);
                $sheet->setCellValue([6,$rowIndex], $patientDays);
                $sheet->setCellValue([7,$rowIndex], $cost);
                $rowIndex++;
            }
            if ($rowIndex - 1 > $moRowIndex) {
                $sheet->mergeCells([1, $moRowIndex, 1, $rowIndex - 1]);
                $sheet->mergeCells([2, $moRowIndex, 2, $rowIndex - 1]);
            }
        }
        $sheet->removeRow($rowIndex,$endRow-$rowIndex+1);


        $sheet = $spreadsheet->getSheetByName('КС');
        $sheet->setCellValue([1,3], $docName);
        $ordinalRowNum = 0;
        $rowIndex = $startRow;
        $category = 'hospital';

        $numberOfBedsIndicatorId = 1; // число коек
        $casesOfTreatmentIndicatorId = 7; // госпитализаций
        $patientDaysIndicatorId = 3; // пациенто-дней
        $costIndicatorId = 4; // стоимость

        foreach($moCollection as $mo) {
            $moRowIndex = $rowIndex;
            $roundClockBedProfiles = $content['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? null;
            $roundClockBedProfilesNumberOfBeds = $contentNumberOfBeds['mo'][$mo->id][$category]['roundClock']['regular']['bedProfiles'] ?? null;

            if (!$roundClockBedProfiles) { continue; }

            $ordinalRowNum++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            $sheet->setCellValue([2,$rowIndex], $mo->short_name);
            foreach ($careProfilesFoms as $cpf) {
                $numberOfBeds = '0';
                $casesOfTreatment = '0';
                $patientDays = '0';
                $cost = '0';

                $hbp = $cpf->hospitalBedProfiles;
                foreach($hbp as $bp) {
                    $bpData = $roundClockBedProfiles[$bp->id] ?? null;
                    if (!$bpData) { continue; }

                    $casesOfTreatment = bcadd($casesOfTreatment, $bpData[$casesOfTreatmentIndicatorId] ?? '0');
                    $patientDays = bcadd($patientDays, $bpData[$patientDaysIndicatorId] ?? '0');
                    $cost = bcadd($cost, $bpData[$costIndicatorId] ?? '0');
                }
                foreach($hbp as $bp) {
                    $bpData = $roundClockBedProfilesNumberOfBeds[$bp->id] ?? null;
                    if (!$bpData) { continue; }

                    $numberOfBeds = bcadd($numberOfBeds, $bpData[$numberOfBedsIndicatorId] ?? '0');
                }
                if( bccomp($numberOfBeds,'0') === 0
                    && bccomp($casesOfTreatment, '0') === 0
                    && bccomp($patientDays, '0') === 0
                    && bccomp($cost, '0') === 0
                ) {continue;}

                $sheet->setCellValue([3,$rowIndex], "$cpf->name");
                $sheet->setCellValue([4,$rowIndex], $numberOfBeds);
                $sheet->setCellValue([5,$rowIndex], $casesOfTreatment);
                $sheet->setCellValue([6,$rowIndex], $patientDays);
                $sheet->setCellValue([7,$rowIndex], $cost);
                $rowIndex++;
            }
            if ($rowIndex - 1 > $moRowIndex) {
                $sheet->mergeCells([1, $moRowIndex, 1, $rowIndex - 1]);
                $sheet->mergeCells([2, $moRowIndex, 2, $rowIndex - 1]);
            }
        }
        $sheet->removeRow($rowIndex,$endRow-$rowIndex+1);


        $sheet = $spreadsheet->getSheetByName('ВМП');
        $sheet->setCellValue([1,3], $docName);
        $ordinalRowNum = 0;
        $rowIndex = $startRow;
        $category = 'hospital';

        $numberOfBedsIndicatorId = 1; // число коек
        $casesOfTreatmentIndicatorId = 7; // госпитализаций
        $patientDaysIndicatorId = 3; // пациенто-дней
        $costIndicatorId = 4; // стоимость

        foreach($moCollection as $mo) {
            $moRowIndex = $rowIndex;
            $careProfiles = $content['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? null;
            $careProfilesNumberOfBeds = $contentNumberOfBeds['mo'][$mo->id][$category]['roundClock']['vmp']['careProfiles'] ?? null;

            if (!$careProfiles) { continue; }

            $ordinalRowNum++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            $sheet->setCellValue([2,$rowIndex], $mo->short_name);
            foreach ($careProfilesFoms as $cpf) {
                $numberOfBeds = '0';
                $casesOfTreatment = '0';
                $patientDays = '0';
                $cost = '0';

                $cpmz = $cpf->careProfilesMz;
                foreach($cpmz as $cp) {
                    $vmpGroupsData = $careProfiles[$cp->id] ?? null;
                    if (!$vmpGroupsData) { continue; }

                    foreach ($vmpGroupsData as $vmpTypes) {
                        foreach ($vmpTypes as $vmpT)
                        {
                            $casesOfTreatment = bcadd($casesOfTreatment, $vmpT[$casesOfTreatmentIndicatorId] ?? '0');
                            $patientDays = bcadd($patientDays, $vmpT[$patientDaysIndicatorId] ?? '0');
                            $cost = bcadd($cost, $vmpT[$costIndicatorId] ?? '0');
                        }
                    }
                }
                foreach($cpmz as $cp) {
                    $vmpGroupsData = $careProfilesNumberOfBeds[$cp->id] ?? null;
                    if (!$vmpGroupsData) { continue; }

                    foreach ($vmpGroupsData as $vmpTypes) {
                        foreach ($vmpTypes as $vmpT)
                        {
                            $numberOfBeds = bcadd($numberOfBeds, $vmpT[$numberOfBedsIndicatorId] ?? '0');
                        }
                    }
                }

                if( bccomp($numberOfBeds,'0') === 0
                    && bccomp($casesOfTreatment, '0') === 0
                    && bccomp($patientDays, '0') === 0
                    && bccomp($cost, '0') === 0
                ) {continue;}

                $sheet->setCellValue([3,$rowIndex], "$cpf->name");
                $sheet->setCellValue([4,$rowIndex], $numberOfBeds);
                $sheet->setCellValue([5,$rowIndex], $casesOfTreatment);
                $sheet->setCellValue([6,$rowIndex], $patientDays);
                $sheet->setCellValue([7,$rowIndex], $cost);
                $rowIndex++;
            }
            if ($rowIndex - 1 > $moRowIndex) {
                $sheet->mergeCells([1, $moRowIndex, 1, $rowIndex - 1]);
                $sheet->mergeCells([2, $moRowIndex, 2, $rowIndex - 1]);
            }
        }
        $sheet->removeRow($rowIndex,$endRow-$rowIndex+1);


        $sheet = $spreadsheet->getSheetByName('АП (подушевое финансирование)');
        $sheet->setCellValue([1,3], $docName);
        $ordinalRowNum = 0;
        $rowIndex = $startRow - 1;
        $category = 'polyclinic';

        $contactingIndicatorId = 8; // обращений
        $visitIndicatorId = 9; // посещений
        $costIndicatorId = 4; // стоимость
        $emergencyVisitsAssistanceTypeIds = [3]; //	посещения неотложные
        $otherVisitsAssistanceTypeIds = [1, 2]; //	посещения профилактические, посещения разовые по заболеваниям
        $contactingAssistanceTypeIds = [4]; //обращения по заболеваниям
        foreach($moCollection as $mo) {
            $assistanceTypesPerPerson = $content['mo'][$mo->id][$category]['perPerson']['all']['assistanceTypes'] ?? null;
            if (!$assistanceTypesPerPerson) { continue; }

            $emergencyVisits = '0';
            $costEmergencyVisits = '0';
            foreach($emergencyVisitsAssistanceTypeIds as $typeId) {
                $emergencyVisits = bcadd($emergencyVisits, $assistanceTypesPerPerson[$typeId][$visitIndicatorId] ?? '0');
                $costEmergencyVisits = bcadd($costEmergencyVisits, $assistanceTypesPerPerson[$typeId][$costIndicatorId] ?? '0');
            }

            $otherVisits = '0';
            $costOtherVisits = '0';
            foreach($otherVisitsAssistanceTypeIds as $typeId) {
                $otherVisits = bcadd($otherVisits, $assistanceTypesPerPerson[$typeId][$visitIndicatorId] ?? '0');
                $costOtherVisits = bcadd($costOtherVisits, $assistanceTypesPerPerson[$typeId][$costIndicatorId] ?? '0');
            }

            $contacting = '0';
            $costContacting = '0';
            foreach($contactingAssistanceTypeIds as $typeId) {
                $contacting = bcadd($contacting, $assistanceTypesPerPerson[$typeId][$contactingIndicatorId] ?? '0');
                $costContacting = bcadd($costContacting, $assistanceTypesPerPerson[$typeId][$costIndicatorId] ?? '0');
            }



            if( bccomp($emergencyVisits,'0') === 0
                && bccomp($otherVisits, '0') === 0
                && bccomp($contacting, '0') === 0
                && bccomp($costEmergencyVisits, '0') === 0
                && bccomp($costOtherVisits, '0') === 0
                && bccomp($costContacting, '0') === 0
            ) {continue;}

            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            $sheet->setCellValue([2,$rowIndex], $mo->short_name);
            $sheet->setCellValue([3,$rowIndex], $otherVisits);
            $sheet->setCellValue([4,$rowIndex], $costOtherVisits);
            $sheet->setCellValue([5,$rowIndex], $contacting);
            $sheet->setCellValue([6,$rowIndex], $costContacting);
            $sheet->setCellValue([7,$rowIndex], $emergencyVisits);
            $sheet->setCellValue([8,$rowIndex], $costEmergencyVisits);

        }
        $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);

        ////////////////////////
        // Поликлиника по тарифу
        ////////////////////////
        $spreadsheet->getDefaultStyle()->getFont()->setName('Arial');
        $spreadsheet->getDefaultStyle()->getFont()->setSize(12);

        $typeFinId = IndicatorType::where('name', 'money')->first()->id;
        $typeQuantId = IndicatorType::where('name', 'volume')->first()->id;

        $polyclinicTariffCategory = CategoryTreeNodes::Where('slug', 'polyclinic-tariff')->first();
        $polyclinicTariffAllNodeIds = $this->nodeService->allChildrenNodeIds($polyclinicTariffCategory->id);
        foreach($polyclinicTariffAllNodeIds as $nodeId) {
            $node = CategoryTreeNodes::find($nodeId);
            $category = Category::find($node->category_id);

            $medicalAssistanceTypeIds = $this->nodeService->medicalAssistanceTypesForNodeId($nodeId);
            $medicalServiceIds = $this->nodeService->medicalServicesForNodeId($nodeId);
            $plannedIndicatorsForNodeId = PlannedIndicator::find($this->nodeService->plannedIndicatorsForNodeId($nodeId));
            $allIndicators = Indicator::all();

            $arr['assistanceTypes'] = MedicalAssistanceType::find($medicalAssistanceTypeIds);
            $arr['services'] = MedicalServices::find($medicalServiceIds);

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
                foreach($arr as $key => $colunms) {
                    $perUnit = $content['mo'][$mo->id]['polyclinic']['perUnit']['all'][$key] ?? null;
                    if (!$perUnit) { continue; }
                    foreach($colunms as $medicalAssistanceTypeOrService) {
                        $indicatorIds = $plannedIndicatorsForNodeId->filter(function ($value) use ($medicalAssistanceTypeOrService, $key) {
                            if ($key == 'assistanceTypes') {
                                return $value->assistance_type_id === $medicalAssistanceTypeOrService->id;
                            } else if ($key == 'services') {
                                return $value->service_id === $medicalAssistanceTypeOrService->id;
                            }

                        })->unique('indicator_id')->pluck('indicator_id');
                        $indicators = $allIndicators->find($indicatorIds);
                        $quantIndicator = $indicators->firstWhere('type_id', $typeQuantId);
                        $finIndicator = $indicators->firstWhere('type_id', $typeFinId);

                        $quantVal = $perUnit[$medicalAssistanceTypeOrService->id][$quantIndicator->id] ?? '0';
                        $finVal = $perUnit[$medicalAssistanceTypeOrService->id][$finIndicator->id] ?? '0';

                        $arrayData[$dataRow][++$dataCol] = $quantVal;
                        $arrayData[$dataRow][++$dataCol] = $finVal;
                        if(!$rowHasData) {
                            if (bccomp($quantVal,'0') !== 0 || bccomp($finVal,'0')) {
                                $tableHasData = true;
                                $rowHasData = true;
                            }
                        }
                    }
                }
                if(!$rowHasData) {
                    unset($arrayData[$dataRow]);
                } else {
                    $dataRow++;
                }
            }
            if ($tableHasData) {
                $sheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet, mb_substr('АП(тариф)' . \Illuminate\Support\Str::ucfirst($category->name), 0, 31));
                $spreadsheet->addSheet($sheet, 0);
                $curRow = 0;
                $curCol = 1;
                $sheet->setCellValue([$curCol,++$curRow], 'Корректировка объемов и финансового обеспечения медицинской помощи');
                $sheet->getRowDimension($curRow)->setRowHeight(20);
                $sheet->setCellValue([1,++$curRow], 'Медицинская помощь в амбулаторных условиях, оплата по тарифу, ' . $category->name);
                $sheet->getRowDimension($curRow)->setRowHeight(20);
                $sheet->setCellValue([1,++$curRow], $docName);
                $sheet->getRowDimension($curRow)->setRowHeight(20);

                $minimumVolumeDataCellWidth = 8;
                $minimumMoneyDataCellWidth = 15;

                $curRow = 4;
                $curCol = 1;
                $step = 2;
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
                $sheet->getColumnDimensionByColumn($curCol)->setAutoSize(true);
                $sheet->setCellValue([++$curCol, $curRow], 'Медицинская организация');
                $sheet->getColumnDimensionByColumn($curCol)->setWidth(50);
                $staticTableHeadEndCol = $curCol;
                $sheet->setCellValue([++$curCol, $curRow], 'корректировка');

                $curRow++;
                // Динамическая часть заголовка таблицы
                foreach($arr as $key => $colunms) {
                    foreach($colunms as $t) {
                        $sheet->setCellValue([$curCol, $curRow], $t->name);
                        $sheet->mergeCells([$curCol, $curRow, $curCol + $step - 1, $curRow]);
                        $halfWidth = round(mb_strwidth ($t->name) / 2) + 2;

                        $indicatorIds = $plannedIndicatorsForNodeId->filter(function ($value) use ($t, $key) {
                            if ($key == 'assistanceTypes') {
                                return $value->assistance_type_id === $t->id;
                            } else if ($key == 'services') {
                                return $value->service_id === $t->id;
                            }

                        })->unique('indicator_id')->pluck('indicator_id');
                        $indicators = $allIndicators->find($indicatorIds);
                        $quantIndicator = $indicators->firstWhere('type_id', $typeQuantId);
                        // $finIndicator = $indicators->firstWhere('type_id', $typeFinId);

                        $sheet->getColumnDimensionByColumn($curCol)->setWidth(max(mb_strwidth($quantIndicator->name)+2, $halfWidth, $minimumVolumeDataCellWidth));
                        $sheet->getColumnDimensionByColumn($curCol + 1)->setWidth(max($halfWidth, $minimumMoneyDataCellWidth));
                        $sheet->setCellValue([$curCol, $curRow + 1], 'объемы, ' . $quantIndicator->name);
                        $sheet->getRowDimension($curRow + 1)->setRowHeight(50);
                        $sheet->setCellValue([$curCol + 1, $curRow + 1], 'финансовое обеспечение, руб.');
                        $sheet->getStyle([$curCol, $curRow + 1, $curCol + 1, $curRow + 1])->getAlignment()->setWrapText(true);

                        $tableEndCol = $curCol + $step - 1;
                        $curCol += $step;
                    }
                }
                $tableHeadEndRow = ++$curRow;
                $curRow++;
                $tableBodyStartRow = $curRow;

                $totalRow = $tableBodyStartRow + count($arrayData);
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
                for($ci = $staticTableHeadStartCol; $ci <= $staticTableHeadEndCol; $ci++) {
                    $sheet->mergeCells([$ci, $tableHeadStartRow, $ci, $tableHeadEndRow]);
                }
                $sheet->mergeCells([$staticTableHeadEndCol + 1, $tableHeadStartRow, $tableEndCol, $tableHeadStartRow]);
                $sheet->mergeCells([$staticTableHeadStartCol, $totalRow, $staticTableHeadEndCol, $totalRow]);
                $sheet->getStyle([$tableStartCol, $tableHeadStartRow, $tableEndCol, $tableHeadEndRow])
                    ->getAlignment()
                    ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER)
                    ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
                $sheet->getStyle([$tableStartCol, $totalRow, $tableStartCol, $totalRow])
                    ->getAlignment()
                    ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT)
                    ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
                $sheet->getStyle([$tableStartCol, $totalRow, $tableStartCol, $totalRow])->getFont()->setBold(true);
                // Border таблицы
                $styleArray = [
                    'borders' => [
                        'allBorders' => [
                            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                        ],
                    ],
                ];
                $sheet->getStyle([$tableStartCol, $tableHeadStartRow, $tableEndCol, $tableEndRow])->applyFromArray($styleArray);
            }
        }

        $sheet = $spreadsheet->getSheetByName('АП (по тарифу)');

        $sheet->setCellValue([1,3], $docName);
        $ordinalRowNum = 0;
        $rowIndex = $startRow - 1;
        $category = 'polyclinic';

        $contactingIndicatorId = 8; // обращений
        $visitIndicatorId = 9; // посещений
        $costIndicatorId = 4; // стоимость
        $emergencyVisitsAssistanceTypeIds = [3]; //	посещения неотложные
        $otherVisitsAssistanceTypeIds = [1, 2]; //	посещения профилактические, посещения разовые по заболеваниям
        $contactingAssistanceTypeIds = [4]; //обращения по заболеваниям
        foreach($moCollection as $mo) {
            $assistanceTypesPerUnit = $content['mo'][$mo->id][$category]['perUnit']['all']['assistanceTypes'] ?? null;
            if (!$assistanceTypesPerUnit) { continue; }

            $emergencyVisits = '0';
            $costEmergencyVisits = '0';
            foreach($emergencyVisitsAssistanceTypeIds as $typeId) {
                $emergencyVisits = bcadd($emergencyVisits, $assistanceTypesPerUnit[$typeId][$visitIndicatorId] ?? '0');
                $costEmergencyVisits = bcadd($costEmergencyVisits, $assistanceTypesPerUnit[$typeId][$costIndicatorId] ?? '0');
            }

            $otherVisits = '0';
            $costOtherVisits = '0';
            foreach($otherVisitsAssistanceTypeIds as $typeId) {
                $otherVisits = bcadd($otherVisits, $assistanceTypesPerUnit[$typeId][$visitIndicatorId] ?? '0');
                $costOtherVisits = bcadd($costOtherVisits, $assistanceTypesPerUnit[$typeId][$costIndicatorId] ?? '0');
            }

            $contacting = '0';
            $costContacting = '0';
            foreach($contactingAssistanceTypeIds as $typeId) {
                $contacting = bcadd($contacting, $assistanceTypesPerUnit[$typeId][$contactingIndicatorId] ?? '0');
                $costContacting = bcadd($costContacting, $assistanceTypesPerUnit[$typeId][$costIndicatorId] ?? '0');
            }



            if( bccomp($emergencyVisits,'0') === 0
                && bccomp($otherVisits, '0') === 0
                && bccomp($contacting, '0') === 0
                && bccomp($costEmergencyVisits, '0') === 0
                && bccomp($costOtherVisits, '0') === 0
                && bccomp($costContacting, '0') === 0
            ) {continue;}

            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            $sheet->setCellValue([2,$rowIndex], $mo->short_name);
            $sheet->setCellValue([3,$rowIndex], $otherVisits);
            $sheet->setCellValue([4,$rowIndex], $costOtherVisits);
            $sheet->setCellValue([5,$rowIndex], $contacting);
            $sheet->setCellValue([6,$rowIndex], $costContacting);
            $sheet->setCellValue([7,$rowIndex], $emergencyVisits);
            $sheet->setCellValue([8,$rowIndex], $costEmergencyVisits);

        }
        $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);


        $sheet = $spreadsheet->getSheetByName('АП (ФАП)');

        $sheet->setCellValue([1,3], $docName);
        $ordinalRowNum = 0;
        $rowIndex = $startRow - 1;
        $category = 'polyclinic';

        $contactingIndicatorId = 8; // обращений
        $visitIndicatorId = 9; // посещений
        $costIndicatorId = 4; // стоимость
        $emergencyVisitsAssistanceTypeIds = [3]; //	посещения неотложные
        $otherVisitsAssistanceTypeIds = [1, 2]; //	посещения профилактические, посещения разовые по заболеваниям
        $contactingAssistanceTypeIds = [4]; //обращения по заболеваниям
        foreach($moCollection as $mo) {

            $faps = $content['mo'][$mo->id][$category]['fap'] ?? null;
            if (!$faps) { continue; }

            $emergencyVisits = '0';
            $costEmergencyVisits = '0';
            $otherVisits = '0';
            $costOtherVisits = '0';
            $contacting = '0';
            $costContacting = '0';
            foreach ($faps as $f) {
                $assistanceTypesData = $f['assistanceTypes'];
                foreach($emergencyVisitsAssistanceTypeIds as $typeId) {
                    $emergencyVisits = bcadd($emergencyVisits, $assistanceTypesData[$typeId][$visitIndicatorId] ?? '0');
                    $costEmergencyVisits = bcadd($costEmergencyVisits, $assistanceTypesData[$typeId][$costIndicatorId] ?? '0');
                }
                foreach($otherVisitsAssistanceTypeIds as $typeId) {
                    $otherVisits = bcadd($otherVisits, $assistanceTypesData[$typeId][$visitIndicatorId] ?? '0');
                    $costOtherVisits = bcadd($costOtherVisits, $assistanceTypesData[$typeId][$costIndicatorId] ?? '0');
                }
                foreach($contactingAssistanceTypeIds as $typeId) {
                    $contacting = bcadd($contacting, $assistanceTypesData[$typeId][$contactingIndicatorId] ?? '0');
                    $costContacting = bcadd($costContacting, $assistanceTypesData[$typeId][$costIndicatorId] ?? '0');
                }
            }

            if( bccomp($emergencyVisits,'0') === 0
                && bccomp($otherVisits, '0') === 0
                && bccomp($contacting, '0') === 0
                && bccomp($costEmergencyVisits, '0') === 0
                && bccomp($costOtherVisits, '0') === 0
                && bccomp($costContacting, '0') === 0
            ) {continue;}

            $ordinalRowNum++;
            $rowIndex++;
            $sheet->setCellValue([1,$rowIndex], "$ordinalRowNum");
            $sheet->setCellValue([2,$rowIndex], $mo->short_name);
            $sheet->setCellValue([3,$rowIndex], $otherVisits);
            $sheet->setCellValue([4,$rowIndex], $costOtherVisits);
            $sheet->setCellValue([5,$rowIndex], $contacting);
            $sheet->setCellValue([6,$rowIndex], $costContacting);
            $sheet->setCellValue([7,$rowIndex], $emergencyVisits);
            $sheet->setCellValue([8,$rowIndex], $costEmergencyVisits);

        }
        $sheet->removeRow($rowIndex+1,$endRow-$rowIndex);

        return $spreadsheet;
    }
}

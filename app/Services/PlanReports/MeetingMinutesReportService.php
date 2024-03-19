<?php
declare(strict_types=1);

namespace App\Services\PlanReports;

use App\Enum\MedicalServicesEnum;
use App\Models\CareProfilesFoms;
use App\Models\CommissionDecision;
use App\Models\MedicalInstitution;
use App\Services\DataForContractService;
use Illuminate\Support\Facades\Storage;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class MeetingMinutesReportService
{
    public function __construct(
        private DataForContractService $dataForContractService,
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

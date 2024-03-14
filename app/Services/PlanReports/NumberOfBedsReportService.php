<?php
declare(strict_types=1);

namespace App\Services\PlanReports;

use App\Models\ChangePackage;
use App\Models\CommissionDecision;
use App\Models\HospitalBedProfiles;
use App\Models\MedicalInstitution;
use App\Services\DataForContractService;
use App\Services\InitialDataFixingService;
use App\Services\PlannedIndicatorChangeInitService;

class NumberOfBedsReportService
{
    public function __construct(
        private DataForContractService $dataForContractService,
        private InitialDataFixingService $initialDataFixingService,
        private PlannedIndicatorChangeInitService $plannedIndicatorChangeInitService,

    ) {}

    public function generateXml(string $bladeTemplateName, int $year, int $commissionDecisionsId = null) {
        $currentlyUsedDate = $year.'-01-01';
        $dateBeg = '01.01.' . $year;
        if ($commissionDecisionsId) {
            $commissionDecisions = CommissionDecision::whereYear('date',$year)->where('id', '<=', $commissionDecisionsId)->get();
            $cd = $commissionDecisions->find($commissionDecisionsId);
            $commissionDecisionIds = $commissionDecisions->pluck('id')->toArray();
            // $protocolDate = $cd->date->format('d.m.Y');
            $packageIds = ChangePackage::whereIn('commission_decision_id', $commissionDecisionIds)->orWhere('commission_decision_id', null)->pluck('id')->toArray();

            $currentlyUsedDate = $cd->date->format('Y-m-d');
            $dateBeg = $cd->date->format('d.m.Y');
        } else {
            if ($this->initialDataFixingService->fixedYear($year)) {
                $packageIds = ChangePackage::where('commission_decision_id', null)->pluck('id')->toArray();
            } else {
                $this->plannedIndicatorChangeInitService->fromInitialData($year);
            }
        }

        $numOfBedIndicatorId = 1; // стоимость

        // Берем число коек на декабрь (12)
        $content = $this->dataForContractService->GetArrayByYearAndMonth($year, 12, $packageIds, [$numOfBedIndicatorId]);

        $moCollection = MedicalInstitution::WhereRaw("? BETWEEN effective_from AND effective_to", [$currentlyUsedDate])->orderBy('order')->get();
        $hospitalBedProfiles = HospitalBedProfiles::all();

        return View($bladeTemplateName)
                ->with("moCollection", $moCollection)
                ->with('content', $content)
                ->with('hospitalBedProfiles', $hospitalBedProfiles)
                ->with('dateBeg', $dateBeg);
    }
}

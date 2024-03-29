<?php
declare(strict_types=1);

namespace App\Services\PlanReports;

use App\Enum\RehabilitationBedOptionEnum;
use App\Services\RehabilitationProfileService;

class PlanCalculatorService {
    // Фин обеспечение
    public static function polyclinicSum($content, int $moId, int $indicatorId, string $category, string $methodOfFinancing, string $polyclinicSubType, array $onlyIds = null ) : string
    {
        $sum = '0';
        $assistanceTypesPerPerson = $content['mo'][$moId][$category][$methodOfFinancing]['all'][$polyclinicSubType] ?? [];
        if ($onlyIds === null) {
            foreach ($assistanceTypesPerPerson as $assistanceType) {
                $sum = bcadd($sum, $assistanceType[$indicatorId] ?? '0');
            }
        } else {
            foreach ($onlyIds as $id) {
                $sum = bcadd($sum, $assistanceTypesPerPerson[$id][$indicatorId] ?? '0');
            }
        }
        return $sum;
    }

    // Фин обеспечение МП по подушевому нормативу финансирования на прикрепившихся лиц (обращения + посещения)
    public static function polyclinicPerPersonAssistanceTypesSum($content, int $moId, int $indicatorId) : string
    {
        return static::polyclinicSum($content, $moId, $indicatorId, 'polyclinic', 'perPerson', 'assistanceTypes');
    }
    // Фин обеспечение МП по подушевому нормативу финансирования на прикрепившихся лиц (услуги)
    public static function polyclinicPerPersonServicesSum($content, int $moId, int $indicatorId) : string
    {
        return static::polyclinicSum($content, $moId, $indicatorId, 'polyclinic', 'perPerson', 'services');
    }
    // Фин обеспечение МП в амбулаторных условиях за единицу объема медицинской помощи (обращения + посещения)
    public static function polyclinicPerUnitAssistanceTypesSum($content, int $moId, int $indicatorId) : string
    {
        return static::polyclinicSum($content, $moId, $indicatorId, 'polyclinic', 'perUnit', 'assistanceTypes');
    }
    // Фин обеспечение МП в амбулаторных условиях за единицу объема медицинской помощи (только указанные ИД из medical_assistance_types)
    public static function polyclinicPerUnitAssistanceTypesOnlyIdsSum($content, int $moId, int $indicatorId, array $onlyIds) : string
    {
        return static::polyclinicSum($content, $moId, $indicatorId, 'polyclinic', 'perUnit', 'assistanceTypes', $onlyIds);
    }
    // Фин обеспечение МП в амбулаторных условиях за единицу объема медицинской помощи (услуги)
    public static function polyclinicPerUnitServicesSum($content, int $moId, int $indicatorId) : string
    {
        return static::polyclinicSum($content, $moId, $indicatorId, 'polyclinic', 'perUnit', 'services');
    }
    // Фин обеспечение МП по нормативу финансирования структурного подразделения (ФАПЫ)
    public static function polyclinicFapSum($content, int $moId, int $indicatorId, string $polyclinicSubType) : string
    {
        $fapSum = '0';
        $faps = $content['mo'][$moId]['polyclinic']['fap'] ?? [];
        foreach ($faps as $f) {
            $services = $f[$polyclinicSubType] ?? [];
            foreach ($services as $service) {
                $fapSum = bcadd($fapSum, $service[$indicatorId] ?? '0');
            }
        }
        return $fapSum;
    }
    // Фин обеспечение МП по нормативу финансирования структурного подразделения (ФАПЫ) (обращения + посещения)
    public static function polyclinicFapServicesSum($content, int $moId, int $indicatorId) : string
    {
        return static::polyclinicFapSum($content, $moId, $indicatorId, 'services');
    }
    // Фин обеспечение МП по нормативу финансирования структурного подразделения (ФАПЫ) (услуги)
    public static function polyclinicFapAssistanceTypesSum($content, int $moId, int $indicatorId) : string
    {
        return static::polyclinicFapSum($content, $moId, $indicatorId, 'assistanceTypes');
    }

    public static function medicalServicesSum($content, int $moId, int $serviceId, int $indicatorId, string $category = 'polyclinic') : string
    {
        bcscale(4);

        $perPerson = $content['mo'][$moId][$category]['perPerson']['all']['services'][$serviceId][$indicatorId] ?? '0';
        $perUnit = $content['mo'][$moId][$category]['perUnit']['all']['services'][$serviceId][$indicatorId] ?? '0';
        $faps = $content['mo'][$moId][$category]['fap'] ?? [];
        $fap = '0';
        foreach ($faps as $f) {
            $fap = bcadd($fap, $f['services'][$serviceId][$indicatorId] ?? '0');
        }
        return bcadd($perPerson, bcadd($perUnit, $fap));
    }

    public static function hospitalBedProfilesSum($content, int $moId, string $daytimeOrRoundClock, string $hospitalSubType, int $reabiletationBedOption, int $indicatorId, string $category = 'hospital') : string
    {
        bcscale(4);

        $bedProfiles = $content['mo'][$moId][$category][$daytimeOrRoundClock][$hospitalSubType]['bedProfiles'] ?? [];
        $v = '0';
        if ($reabiletationBedOption == RehabilitationBedOptionEnum::WithoutRehabilitation) {
            foreach ($bedProfiles as $bpId => $bp) {
                if (RehabilitationProfileService::IsRehabilitationBedProfile($bpId)) {
                    continue;
                }
                $v = bcadd($v, ($bp[$indicatorId] ?? '0'));
            }
        }
        if ($reabiletationBedOption == RehabilitationBedOptionEnum::OnlyRehabilitation) {
            $rehabilitationBedProfileIds = RehabilitationProfileService::GetAllRehabilitationBedProfileIds();
            foreach ($rehabilitationBedProfileIds as $rbpId) {
                $v = bcadd($v, ($bedProfiles[$rbpId][$indicatorId] ?? '0'));
            }
        }
        if ($reabiletationBedOption == RehabilitationBedOptionEnum::WithRehabilitation) {
            foreach ($bedProfiles as $bpId => $bp) {
                $v = bcadd($v, ($bp[$indicatorId] ?? '0'));
            }
        }
        return $v;
    }

    public static function hospitalVmpSum($content, int $moId, string $daytimeOrRoundClock, string $hospitalSubType = 'vmp', int $indicatorId = 7, string $category = 'hospital') : string
    {
        bcscale(4);
        $careProfiles = $content['mo'][$moId][$category][$daytimeOrRoundClock][$hospitalSubType]['careProfiles'] ?? [];
        $v = '0';
        foreach ($careProfiles as $vmpGroups) {
            foreach ($vmpGroups as $vmpTypes) {
                foreach ($vmpTypes as $vmpT) {
                    $v = bcadd($v, ($vmpT[$indicatorId] ?? '0'));
                }
            }
        }
        return $v;
    }
}

<?php
declare(strict_types=1);

namespace App\Services;

class RehabilitationProfileService {

    public static function GetAllRehabilitationBedProfileIds() : array
    {
        // 32 реабилитационные соматические;
        // 34 Реабилитационные для больных с заболеваниями центральной нервной системы и органов чувств;
        // 35 Реабилитационные для больных с заболеваниями опорно-двигательного аппарата и периферической нервной системы;
        return [32, 34, 35];
    }


    // Является реабилитационной койкой
    public static function IsRehabilitationBedProfile(int $bedProfile) : bool
    {
        $rehabilitationBedProfileIds = RehabilitationProfileService::GetAllRehabilitationBedProfileIds();
        return in_array($bedProfile, $rehabilitationBedProfileIds);
    }
}

<?php
declare(strict_types=1);

namespace App\Services;

class VolumeDistributionAlgorithmService
{
    public function __construct(

    ) { }

    public function getAlgorithmId(int $indicatorId)
    {
        if($indicatorId === 1) {
            return 1;
        }
        return 2;
    }
}

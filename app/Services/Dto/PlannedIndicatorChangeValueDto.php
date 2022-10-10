<?php
declare(strict_types=1);

namespace App\Services\Dto;

class PlannedIndicatorChangeValueDto
{
    public function __construct(
        public int $periodId,
        public int $moId,
        public int $plannedIndicatorId,
        public string $value,
        public int $userId,
        public ?int $packageId = null,
        public ?int $moDepartmentId = null,
        public ?int $id = null,
    ) { }
}

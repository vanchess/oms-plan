<?php
declare(strict_types=1);

namespace App\Services\Dto;

class InitialDataValueDto
{
    public function __construct(
        public int $year,
        public int $moId,
        public int $plannedIndicatorId,
        public string $value,
        public int $userId,
        public ?int $moDepartmentId = null,
        public ?int $id = null,
    ) { }
}

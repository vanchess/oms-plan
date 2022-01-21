<?php
declare(strict_types=1);

namespace App\Services\Dto;

class InitialDataValueDto
{
    public function __construct(
        public ?int $id = null,
        public int $year,
        public int $moId,
        //public ?int $moDepartmentId = null,
        public int $plannedIndicatorId,
        public string $value,
        public int $userId
    ) { }
}

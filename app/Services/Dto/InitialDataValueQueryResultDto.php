<?php
declare(strict_types=1);

namespace App\Services\Dto;

use App\Services\Dto\InitialDataValueDto;

class InitialDataValueQueryResultDto
{
    public function __construct(
        public bool $operationError = false,
        public string $operationMessage,
        public bool $hasValue = false,
        public InitialDataValueDto $value
    ) { }
}
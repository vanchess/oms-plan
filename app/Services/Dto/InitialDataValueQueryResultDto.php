<?php
declare(strict_types=1);

namespace App\Services\Dto;

use App\Services\Dto\InitialDataValueDto;

class InitialDataValueQueryResultDto extends QueryResultDto
{
    public function __construct(
        public InitialDataValueDto $value,
        string $operationMessage,
        bool $operationError = false,
        bool $hasValue = false
    ) {
        $this->value = $value;
        $this->operationError = $operationError;
        $this->operationMessage = $operationMessage;
        $this->hasValue = $hasValue;
    }
}

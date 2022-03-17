<?php
declare(strict_types=1);

namespace App\Services\Dto;

class PlannedIndicatorChangeValueQueryResultDto extends QueryResultDto
{
    public function __construct
    (
        public PlannedIndicatorChangeValueDto $value,
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

<?php
declare(strict_types=1);

namespace App\Services\Dto;

class PlannedIndicatorChangeValueCollectionQueryResultDto extends QueryResultDto
{
    public function __construct
    (
        public PlannedIndicatorChangeValueCollectionDto $value,
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

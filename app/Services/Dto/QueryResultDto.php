<?php
declare(strict_types=1);

namespace App\Services\Dto;

abstract class QueryResultDto
{
    public bool $operationError = false;
    public string $operationMessage;
    public bool $hasValue = false;
}

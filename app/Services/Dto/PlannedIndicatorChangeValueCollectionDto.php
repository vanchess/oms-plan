<?php
declare(strict_types=1);

namespace App\Services\Dto;

class PlannedIndicatorChangeValueCollectionDto
{
    private array $items = [];

    public function __construct(PlannedIndicatorChangeValueDto ...$array)
    {
        $this->items = $array;
    }

    public function toArray(): array
    {
        return $this->items;
    }
}

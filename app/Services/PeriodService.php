<?php
declare(strict_types=1);

namespace App\Services;

use App\Models\Period;
use DateTime;
use DateTimeZone;

class PeriodService
{
    public function getAll()//: array
    {
        $periods = Period::all();
        return $periods->toArray();
    }

    public function getByYear(int $year): array
    {
        $from = new DateTime("${year}-01-01+0500");
        $to = new DateTime("${year}-12-31T23:59:59.999999+0500");
        $from->setTimezone(new DateTimeZone('UTC'));
        $to->setTimezone(new DateTimeZone('UTC'));
        $periods = Period::whereBetween('from',[$from->format('Y-m-d H:i:s'), $to->format('Y-m-d H:i:s')])->orderBy('from')->get();
        return $periods->toArray();
    }

    public function getIdsByYear(int $year): array
    {
        return array_column($this->getByYear($year), 'id');
    }

    public function getByYearAndMonth(int $year, int $monthNum)
    {
        $from = new DateTime("${year}-$monthNum-01+0500");
        $lastDay = $from->format( 'Y-m-t' );
        $to = new DateTime("{$lastDay}T23:59:59.999999+0500");
        $from->setTimezone(new DateTimeZone('UTC'));
        $to->setTimezone(new DateTimeZone('UTC'));

        $periods = Period::whereBetween('from',[$from->format('Y-m-d H:i:s'), $to->format('Y-m-d H:i:s')])->orderBy('from')->get();
        return $periods->toArray();
    }

    public function getIdsByYearAndMonth(int $year, int $monthNum)
    {
        return array_column($this->getByYearAndMonth($year, $monthNum), 'id');
    }

    public function getYearById(int $id): int
    {
        $period = Period::findOrFail($id);
        return $period->to->year;
    }
}

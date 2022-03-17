<?php
declare(strict_types=1);

namespace App\Services;

use App\Models\InitialData;
use App\Models\PlannedIndicatorChange;

class PlannedIndicatorChangeInitService
{
    public function __construct(
        private PeriodService $periodService
    ) { }

    /**
     * Сформировать "базовые" изменения планируемых показателей из начальных данных.
     * Ранее обработанные начальные данные (processed_at != NULL) пропускаются.
     * После обработки начальных данных заполняется свойство processed_at.
     */
    public function fromInitialData(int $year)
    {
        bcscale(4);

        $initialData = InitialData::where('year',$year)->whereNull('processed_at')->orderBy('id')->get();
        $periods = $this->periodService->getIdsByYear($year);
        $periodsCount = count($periods);

        foreach ($initialData as $data) {

            $changesForDelete = PlannedIndicatorChange::whereIn('period_id',$periods)
                ->where('mo_id', $data->mo_id)
                ->where('planned_indicator_id', $data->planned_indicator_id);
            if($data->mo_department_id) {
                $changesForDelete = $changesForDelete->where('mo_department_id', $data->mo_department_id);
            } else {
                $changesForDelete = $changesForDelete->whereNull('mo_department_id');
            }
            $changesForDelete->delete();


            if($data->algorithm_id === 2) {
                $v = strval(round($data->value / $periodsCount));
                $t = bcsub($data->value, strval($v * $periodsCount));
                $values = [];

                for ($i=0; $i < $periodsCount; $i++) {
                    $values[$i] = $v;
                }
                $tRound = round(floatval($t), 0, PHP_ROUND_HALF_DOWN);
                $remainder = bcsub($t, strval($tRound));
                $t = $tRound;

                $values[$periodsCount-1] = bcadd($values[$periodsCount-1], $remainder);

                if($t != 0) {
                    $d = ($t > 0) ? '1' : '-1';
                    $t = abs($t);
                    $step = round($periodsCount / $t, 0, PHP_ROUND_HALF_DOWN);
                    for ($j=0; $j < $t; $j++) {
                        $key = $periodsCount - 1 - ($step * $j);
                        $values[$key] = bcadd($values[$key], $d);
                    }
                }

                $i = 0;
                foreach ($periods as $periodId) {
                    $change = new PlannedIndicatorChange();
                    $change->mo_id = $data->mo_id;
                    $change->planned_indicator_id = $data->planned_indicator_id;
                    $change->user_id = $data->user_id;
                    $change->value = $values[$i];
                    $change->mo_department_id = $data->mo_department_id;
                    $change->period_id = $periodId;
                    $change->save();

                    $i++;
                }

            } elseif ($data->algorithm_id === 1) {
                foreach ($periods as $periodId) {
                    $change = new PlannedIndicatorChange();
                    $change->mo_id = $data->mo_id;
                    $change->planned_indicator_id = $data->planned_indicator_id;
                    $change->user_id = $data->user_id;
                    $change->value = $data->value;
                    $change->mo_department_id = $data->mo_department_id;
                    $change->period_id = $periodId;
                    $change->save();
                }
            }
            $data->processed_at = date('Y-m-d H:i:s');
            $data->save();


        }
    }
}

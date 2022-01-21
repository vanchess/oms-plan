<?php

namespace App\Http\Resources;

use Illuminate\Http\Resources\Json\JsonResource;

class InitialDataResource extends JsonResource
{
    /**
     * Transform the resource into an array.
     *
     * @param  \Illuminate\Http\Request  $request
     * @return array
     */
    public function toArray($request)
    {
        return [
            'id'    => $this->id,
            'year'  => $this->year,
            'mo_id' => $this->moId,
            //'mo_department_id' => $this->moDepartmentId,
            'planned_indicator_id' => $this->plannedIndicatorId,
            'value' => rtrim(rtrim($this->value,'0'),'.'),
            'user_id' => $this->userId,
            //'processed_at' => $this->processed_at,
            //'created_at' => $this->created_at,
        ];
    }
}

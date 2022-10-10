<?php

namespace App\Http\Resources;

use Illuminate\Http\Resources\Json\JsonResource;

class PlannedIndicatorChangeResource extends JsonResource
{
    /**
     * Transform the resource into an array.
     *
     * @param  \Illuminate\Http\Request  $request
     * @return array|\Illuminate\Contracts\Support\Arrayable|\JsonSerializable
     */
    public function toArray($request)
    {
        return [
            'id'    => $this->id,
            'period_id' => $this->periodId,
            'mo_id' => $this->moId,
            'planned_indicator_id' => $this->plannedIndicatorId,
            'user_id' => $this->userId,
            'value' => str_contains($this->value,'.')?rtrim(rtrim($this->value,'0'),'.'):$this->value,
            'mo_department_id' => $this->moDepartmentId,
            'package_id' => $this->packageId
            //'created_at' => $this->created_at,
        ];
    }
}

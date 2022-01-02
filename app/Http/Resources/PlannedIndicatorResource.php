<?php

namespace App\Http\Resources;

use Illuminate\Http\Resources\Json\JsonResource;

class PlannedIndicatorResource extends JsonResource
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
            'node_id' => $this->node_id,
            'indicator_id' => $this->indicator_id,
            'service_id' => $this->service_id,
            'profile_id' => $this->profile_id,
            'assistance_type_id' => $this->assistance_type_id,
            'fap_id' => $this->fap_id,
        ];
    }
}

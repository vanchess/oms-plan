<?php

namespace App\Http\Resources;

use Illuminate\Http\Resources\Json\JsonResource;

class MedicalInstitutionResource extends JsonResource
{
    /**
     * Transform the resource into an array.
     *
     * @param  \Illuminate\Http\Request  $request
     * @return array
     */
    public function toArray($request)
    {
        //return parent::toArray($request);
        return [
            //'type'          => 'medicalInstitution',
            'id'            => $this->id,
            //'attributes'    => [
                'code' => $this->code,
                'name' => $this->name,
                'short_name' => $this->short_name,
                //'description' => $this->description,
                //'inn' => $this->inn,
                //'ogrn' => $this->ogrn,
                //'kpp' => $this->kpp,
                //'address' => $this->address,
                //'chief' => $this->chief,
                //'phone' => $this->phone,
                //'email' => $this->email,
                //'license' => $this->license,
                //'inclusion_in_register' => $this->inclusion_in_register,
                //'created_at' => $this->created_at,
                //'updated_at' => $this->updated_at,
                'order' => $this->order
            //],
           // 'relationships' => new EmployeeRelationshipResource($this),
           // 'links'         => [
           //     'self' => route('medical-institution.show', ['medical_institution' => $this->id]),
           // ],
            
        ];
    }
}

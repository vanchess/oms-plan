<?php

namespace App\Http\Controllers;

use App\Models\MedicalInstitution;
use Illuminate\Http\Request;

use App\Http\Resources\MedicalInstitutionCollection;
use App\Http\Resources\MedicalInstitutionResource;

class MedicalInstitutionController extends Controller
{
    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function index(Request $request)
    {
        //$perPage = (int)$request->input('per_page', 0);
        //if($perPage == -1) {
        //    $result = MedicalInstitution::OrderBy('order')->paginate(999999999);
        //    return new MedicalInstitutionCollection($result);
        //}
        return new MedicalInstitutionCollection(MedicalInstitution::OrderBy('order')->get());//->paginate($perPage));
    }

    /**
     * Store a newly created resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @return \Illuminate\Http\Response
     */
    public function store(Request $request)
    {
        //
    }

    /**
     * Display the specified resource.
     *
     * @param  \App\Models\MedicalInstitution  $medicalInstitution
     * @return \Illuminate\Http\Response
     */
    public function show(MedicalInstitution $medicalInstitution)
    {
        // сообщить ресурсу, что мы не хотим что бы он был завернут (элемент не должен иметь ключ верхнего уровня data) 
        MedicalInstitutionResource::withoutWrapping();
        return new MedicalInstitutionResource($medicalInstitution);
    }

    /**
     * Update the specified resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @param  \App\Models\MedicalInstitution  $medicalInstitution
     * @return \Illuminate\Http\Response
     */
    public function update(Request $request, MedicalInstitution $medicalInstitution)
    {
        //
    }

    /**
     * Remove the specified resource from storage.
     *
     * @param  \App\Models\MedicalInstitution  $medicalInstitution
     * @return \Illuminate\Http\Response
     */
    public function destroy(MedicalInstitution $medicalInstitution)
    {
        //
    }
}

<?php

namespace App\Http\Controllers;

use App\Models\MedicalInstitution;
use Illuminate\Http\Request;

use App\Http\Resources\MedicalInstitutionCollection;
use App\Http\Resources\MedicalInstitutionResource;
use DateTime;
use Illuminate\Support\Facades\Validator;

class MedicalInstitutionController extends Controller
{
    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function index(Request $request)
    {
        $validator = Validator::make($request->all(), [
            'date' => 'date|after:2021-12-31',
        ]);
        if ($validator->fails()) {
            return response()->json($validator->errors()->toJson(), 400);
        }

        $validated = $validator->validated();
        $date = date('Y-m-d');
        if (isset($validated['date'])) {
            $date = $validated['date'];
        }

        return new MedicalInstitutionCollection(MedicalInstitution::WhereRaw("? BETWEEN effective_from AND effective_to", [$date])->orderBy('order')->get());
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

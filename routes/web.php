<?php

use App\Jobs\InitialChanges;
use App\Jobs\InitialDataLoaded;
use App\Models\CareProfiles;
use App\Models\CareProfilesFoms;
use App\Models\HospitalBedProfiles;
use App\Models\Indicator;
use App\Models\MedicalAssistanceType;
use Illuminate\Support\Facades\Route;
use App\Models\MedicalInstitution;
use App\Models\MedicalServices;
use App\Models\Organization;
use App\Models\Period;

use App\Services\InitialDataService;
use App\Models\PlannedIndicator;
use App\Models\PlannedIndicatorChange;
use App\Models\VmpGroup;
use App\Models\VmpTypes;
use App\Services\DataForContractService;
//use App\Services\NodeService;
use App\Services\Dto\InitialDataValueDto;
use App\Services\MoDepartmentsInfoForContractService;
use App\Services\MoInfoForContractService;
use App\Services\PeopleAssignedInfoForContractService;
use Illuminate\Support\Facades\Auth;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\Storage;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/

Route::get('/321123', function (InitialDataService $initialDataService) {

    $nodeId = 4;
    $userId = 1;

    $medicalInstitutions = MedicalInstitution::OrderBy('order')->get();


    foreach ($medicalInstitutions as $mo) {
        foreach (PlannedIndicator::all() as $pi) {

            $dto = new InitialDataValueDto(
                    year: 2021,
                    moId: $mo->id,
                    plannedIndicatorId: $pi->id,
                    value: 2021 + $mo->id + $pi->id,
                    userId: $userId
                );

            $initialDataService->setValue($dto);
        }
    }

    return 'OK';
    /*
    $medicalInstitutions = MedicalInstitution::OrderBy('order')->get();
    foreach ($medicalInstitutions as $mo) {
        $org = Organization::Where('inn',$mo->inn)->first();
        $mo->organization_id = $org->id;
        $mo->save();
    }
    */
    return 'OK';

    //phpinfo();
});

Route::get('/', function (DataForContractService $dataForContractService, MoInfoForContractService $moInfoForContractService, MoDepartmentsInfoForContractService $moDepartmentsInfoForContractService, PeopleAssignedInfoForContractService $peopleAssignedInfoForContractService) {
    // InitialChanges::dispatch(2022);
    // InitialDataLoaded::dispatch(2, 2022, 1);
    // InitialDataLoaded::dispatch(9, 2022, 1);

    $year = 2022;
    $commit = null;
    $strDateTimeNow = date("Y-m-d-His");

    $indicators = Indicator::select('id','name')->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'indicators.json', $indicators);

    $medicalServices = MedicalServices::select('id','name')->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'medicalServices.json', $medicalServices);

    $medicalAssistanceType = MedicalAssistanceType::select('id','name')->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'medicalAssistanceType.json', $medicalAssistanceType);

    $hospitalBedProfiles = HospitalBedProfiles::select('tbl_hospital_bed_profiles.id','name', 'care_profile_foms_id')
    ->join('tbl_hospital_bed_profile_care_profile_foms',  'tbl_hospital_bed_profiles.id', '=', 'hospital_bed_profile_id')
    ->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'hospitalBedProfiles.json', $hospitalBedProfiles);

    $vmpGroup = VmpGroup::select('id','code')->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'vmpGroup.json', $vmpGroup);

    $vmpTypes = VmpTypes::select('id','name')->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'vmpTypes.json', $vmpTypes);

    $careProfiles = CareProfiles::select('tbl_care_profiles.id','name','care_profile_foms_id')
    ->join('tbl_care_profile_care_profile_foms',  'tbl_care_profiles.id', '=', 'care_profile_id')
    ->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'careProfiles.json', $careProfiles);

    $careProfilesFoms = CareProfilesFoms::select('id', 'name', 'code_v002 as code')
    ->get()->mapWithKeys(function($item) {return [$item->id => $item];})->toJson();
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'careProfilesFoms.json', $careProfilesFoms);

    $mo = $moInfoForContractService->GetJson();
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'mo.json', $mo);

    $moDepartments = $moDepartmentsInfoForContractService->GetJson();
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'moDepartments.json', $moDepartments);

    $content = $dataForContractService->GetJson($year, $commit);
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'data.json', $content);

    $content = $peopleAssignedInfoForContractService->GetJson($year, $commit);
    Storage::put( $strDateTimeNow.DIRECTORY_SEPARATOR.'peopleAssignedData.json', $content);
    return $content;
});

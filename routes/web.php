<?php

use App\Jobs\InitialChanges;
use App\Jobs\InitialDataLoaded;
use Illuminate\Support\Facades\Route;
use App\Models\MedicalInstitution;
use App\Models\Organization;
use App\Models\Period;

use App\Services\InitialDataService;
use App\Models\PlannedIndicator;
//use App\Services\NodeService;
use App\Services\Dto\InitialDataValueDto;

use Illuminate\Support\Facades\Auth;
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

Route::get('/', function () {
     InitialChanges::dispatch(2022);
    // InitialDataLoaded::dispatch(2, 2022, 1);
    // InitialDataLoaded::dispatch(9, 2022, 1);

    return 'OK12';
});

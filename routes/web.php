<?php

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
    
    $medicalInstitutions = MedicalInstitution::OrderBy('order')->get();
    foreach ($medicalInstitutions as $mo) {
        $mo->year_qty = 150;
        $mo->year_sum = 3500000;
        $mo->Q1_qty = 40;
        $mo->Q1_sum = 1000000;
        $mo->Q2_qty = 30;
        $mo->Q2_sum = 500000;
        $mo->Q3_qty = 40;
        $mo->Q3_sum = 1000000;
        $mo->Q4_qty = 40;
        $mo->Q4_sum = 1000000;
    }
    
    
    $pages = [
        ['id' => 1, 'name' => 'План'],
    ];
    


    
    return view('welcome', [
        'pages' => $pages,
        'medicalInstitutions' => $medicalInstitutions
    ]);
});

Route::get('/medicalInstitution/{id}/{period}', function ($id,$period) {
    
    $medicalInstitution = MedicalInstitution::find($id);
    
    $pages = [
        ['id' => 1, 'name' => 'План'],
    ];
    
    switch ($period) {
        case 'year':
            $period = 'Год';
            break;
        case 'Q1':
            $period = '1 квартал';
            break;
        case 'Q2':
            $period = '2 квартал';
            break;
        case 'Q3':
            $period = '3 квартал';
            break;
        case 'Q4':
            $period = '4 квартал';
            break;
        default:
           $period = '';
    }
    
    
    
    return view('medical-institution-period', [
        'pages' => $pages,
        'medicalInstitution' => $medicalInstitution,
        'period' => $period
    
    ]);
})->name('medicalInstitutionPeriod');

Route::get('/medicalInstitution/{id}', function ($id) {
    $pages = [
        ['id' => 1, 'name' => 'План'],
    ];
    
     $medicalInstitution = MedicalInstitution::find($id);
    
    return view('medical-institution', [
        'pages' => $pages,
        'medicalInstitution' => $medicalInstitution
    
    ]);
})->name('medicalInstitution');

Route::get('/page/{id}', function ($id) {
    
    return redirect('/');
})->name('page');
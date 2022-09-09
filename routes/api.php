<?php

use Illuminate\Http\Request;
use Illuminate\Support\Facades\Route;

use App\Http\Controllers\AuthController;
use App\Http\Controllers\CareProfilesController;
use App\Http\Controllers\CategoryController;
use App\Http\Controllers\CategoryTreeController;
use App\Http\Controllers\CategoryTreeNodesController;
use App\Http\Controllers\MedicalInstitutionController;
use App\Http\Controllers\MedicalInstitutionDepartmentController;
use App\Http\Controllers\MedicalInstitutionIdsController;
use App\Http\Controllers\HospitalBedProfilesController;
use App\Http\Controllers\MedicalAssistanceTypeController;
use App\Http\Controllers\MedicalServicesController;
use App\Http\Controllers\UsedIndicatorsController;
use App\Http\Controllers\UsedHospitalBedProfilesController;
use App\Http\Controllers\UsedMedicalAssistanceTypeController;
use App\Http\Controllers\UsedMedicalServicesController;
use App\Http\Controllers\IndicatorsController;
use App\Http\Controllers\InitialDataController;
use App\Http\Controllers\InitialDataLoadedController;
use App\Http\Controllers\PeriodController;
use App\Http\Controllers\PlannedIndicatorChangeController;
use App\Http\Controllers\PlannedIndicatorController;
use App\Http\Controllers\UsedCareProfilesController;
use App\Http\Controllers\VmpGroupController;
use App\Http\Controllers\VmpTypesController;

/*
|--------------------------------------------------------------------------
| API Routes
|--------------------------------------------------------------------------
|
| Here is where you can register API routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| is assigned the "api" middleware group. Enjoy building your API!
|
*/
Route::group(array('prefix' => 'v1'), function()
{
    Route::middleware('auth:api')->group(function()
    {
        Route::get('user', function (Request $request) {
            return $request->user();
        });


        Route::apiResource('medical-institution', MedicalInstitutionController::class);
        Route::apiResource('medical-institution-department', MedicalInstitutionDepartmentController::class);
        Route::get('/medical-institution-ids/having-departments', [MedicalInstitutionIdsController::class, 'havingDepartments']);
        Route::apiResource('hospital-bed-profiles', HospitalBedProfilesController::class);
        Route::apiResource('medical-assistance-types', MedicalAssistanceTypeController::class);
        Route::apiResource('medical-services', MedicalServicesController::class);
        Route::apiResource('indicators', IndicatorsController::class);
        Route::apiResource('planned-indicators', PlannedIndicatorController::class);
        Route::get('/used-indicators', [UsedIndicatorsController::class, 'indicatorsUsedForNodeId']);
        Route::get('/used-hospital-bed-profiles', [UsedHospitalBedProfilesController::class, 'hospitalBedProfilesUsedForNodeId']);
        Route::get('/used-medical-assistance-types', [UsedMedicalAssistanceTypeController::class, 'medicalAssistanceTypesUsedForNodeId']);
        Route::get('/used-medical-services', [UsedMedicalServicesController::class, 'medicalServicesUsedForNodeId']);
        Route::get('/used-care-profiles',[UsedCareProfilesController::class, 'careProfilesUsedForNodeId']);
        Route::apiResource('node-init-data', InitialDataController::class);
        Route::apiResource('init-data-loaded-nodes', InitialDataLoadedController::class);
        Route::apiResource('care-profiles', CareProfilesController::class);
        Route::apiResource('vmp-types', VmpTypesController::class);
        Route::apiResource('vmp-groups', VmpGroupController::class);
        Route::get('/node/{nodeId}/cildren', [CategoryTreeController::class, 'nodeWithChildren']);
        Route::get('/category-tree/{rootNodeId}', [CategoryTreeController::class, 'getCategoryTree']);
        Route::apiResource('category-tree-nodes', CategoryTreeNodesController::class);
        Route::apiResource('categories', CategoryController::class);
        Route::apiResource('planned-indicator-change', PlannedIndicatorChangeController::class);
        Route::post('/planned-indicator-change-add-values', [PlannedIndicatorChangeController::class, 'incrementValues']);
        Route::apiResource('periods', PeriodController::class);
    });


    Route::get('/users', function (Request $request) {
        return $request->user();
    });

    Route::group(['middleware' => 'api','prefix' => 'auth'], function ($router) {
        Route::post('/login', [AuthController::class, 'login'])->name('login');
        Route::post('/register', [AuthController::class, 'register']);
        Route::post('/logout', [AuthController::class, 'logout']);
        Route::post('/refresh', [AuthController::class, 'refresh']);
        Route::get('/user-profile', [AuthController::class, 'userProfile']);
    });
});

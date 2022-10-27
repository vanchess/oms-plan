<?php
declare(strict_types=1);

namespace App\Services;

use App\Models\Indicator;
use App\Models\PlannedIndicator;
use App\Models\PlannedIndicatorChange;
use Illuminate\Support\Facades\DB;

class DataForContractService
{
    public function __construct(
        private PeriodService $periodService
    ) {}

    public function GetJson(int $year, array $packageIds = null): string {
        return $this->CreateData($year, $packageIds)->toJson();
    }

    public function GetArray(int $year, array $packageIds = null, array $indicatorIds = [2, 4, 5, 6, 7, 8, 9]): array {
        return $this->CreateData($year, $packageIds, $indicatorIds)->toArray();
    }

    public function GetArrayByYearAndMonth(int $year, int $monthNum, array $packageIds = null, array $indicatorIds = [2, 4, 5, 6, 7, 8, 9]): array {
        return $this->CreateDataByYearAndMonth($year, $monthNum, $packageIds, $indicatorIds)->toArray();
    }

    private function CreateData(int $year, array $packageIds = null, array $indicatorIds = [2, 4, 5, 6, 7, 8, 9]) {
        $periodIds = $this->periodService->getIdsByYear($year);
        $data = $this->CreateDataByPeriodIds($periodIds, $packageIds, $indicatorIds);

        return collect([
            'year' => $year,
            'mo' => $data
        ]);
    }

    private function CreateDataByYearAndMonth(int $year, int $monthNum, array $packageIds = null, array $indicatorIds = [2, 4, 5, 6, 7, 8, 9]) {
        $periodIds = $this->periodService->getIdsByYearAndMonth($year, $monthNum);
        $data = $this->CreateDataByPeriodIds($periodIds, $packageIds, $indicatorIds);

        return collect([
            'mo' => $data
        ]);
    }

    private function CreateDataByPeriodIds(array $periodIds, array $packageIds = null, array $indicatorIds = [2, 4, 5, 6, 7, 8, 9]) {

        $hospitalNodeIds = [1,2,3,4,5,6,7];
        $hospitalDaytimeNodeIds = [2,4,5];
        $hospitalInPolyclinicNodeIds = [4];
        $hospitalInHospitalNodeIds = [5];
        $hospitalRoundClockNodeIds = [3,6,7];
        $hospitalRoundClockRegularNodeIds = [6];
        $hospitalRoundClockVmpNodeIds = [7];

        $ambulanceNodeIds = [17];

        $polyclinicNodeIds = [9,10,11,12,28,29,30,31,32,33];
        $polyclinicPerPersonNodeIds = [10,28,32];
        $polyclinicPerUnitNodeIds = [11,30,31];
        $polyclinicFapNodeIds = [12,29,33];

        $dataSql = DB::table((new PlannedIndicator())->getTable().' as pi')
        ->selectRaw('SUM(value) as value, node_id, indicator_id, service_id, profile_id, assistance_type_id, care_profile_id, vmp_group_id, vmp_type_id, mo_id, planned_indicator_id, mo_department_id')
        //->leftJoin((new Indicator())->getTable().' as ind','pi.indicator_id','=','ind.id')
        ->leftJoin((new PlannedIndicatorChange())->getTable().' as pic', 'pi.id', '=', 'pic.planned_indicator_id');
        if ($packageIds) {
            $dataSql = $dataSql->whereIn('package_id',$packageIds);
        }
        //
        $dataSql = $dataSql->whereIn('indicator_id', $indicatorIds)
        ->whereIn('period_id', $periodIds)
        ->whereNull('pic.deleted_at')
        ->groupBy('node_id', 'indicator_id', 'service_id', 'profile_id', 'assistance_type_id', 'care_profile_id', 'vmp_group_id', 'vmp_type_id', 'mo_id', 'planned_indicator_id', 'mo_department_id');

        $data = $dataSql->get();

        $data = $data->groupBy([
            'mo_id',
            function($item, $key) use  ($hospitalNodeIds, $ambulanceNodeIds, $polyclinicNodeIds) {
                $nodeId = $item->node_id;
                if(in_array($nodeId, $hospitalNodeIds)) {
                    return 'hospital';
                };
                if(in_array($nodeId, $ambulanceNodeIds)) {
                    return 'ambulance';
                };
                if(in_array($nodeId, $polyclinicNodeIds)) {
                    return 'polyclinic';
                };
                return 'none';
            }
        ]);

        foreach ($data as $key => $value) {
            if(isset($value['hospital'])) {
                $value['hospital'] = $value['hospital']->groupBy([
                    function($item, $key) use  ($hospitalDaytimeNodeIds, $hospitalRoundClockNodeIds) {
                        $nodeId = $item->node_id;
                        if(in_array($nodeId, $hospitalDaytimeNodeIds)) {
                            return 'daytime';
                        };
                        if(in_array($nodeId, $hospitalRoundClockNodeIds)) {
                            return 'roundClock';
                        };
                        return 'none';
                    },
                    function($item, $key) use  ($hospitalInPolyclinicNodeIds, $hospitalInHospitalNodeIds, $hospitalRoundClockRegularNodeIds, $hospitalRoundClockVmpNodeIds) {
                        $nodeId = $item->node_id;
                        if(in_array($nodeId, $hospitalInPolyclinicNodeIds)) {
                            return 'inPolyclinic';
                        };
                        if(in_array($nodeId, $hospitalInHospitalNodeIds)) {
                            return 'inHospital';
                        };
                        if(in_array($nodeId, $hospitalRoundClockRegularNodeIds)) {
                            return 'regular';
                        };
                        if(in_array($nodeId, $hospitalRoundClockVmpNodeIds)) {
                            return 'vmp';
                        };
                        return 'none';
                    },
                    function($item, $key) {
                        $serviceId = $item->service_id;
                        $profileId = $item->profile_id;
                        $assistanceTypeId = $item->assistance_type_id;
                        $careProfileId = $item->care_profile_id;
                        if($serviceId) {
                            return 'services';
                        };
                        if($profileId) {
                            return 'bedProfiles';
                        };
                        if($assistanceTypeId) {
                            return 'assistanceTypes';
                        };
                        if($careProfileId) {
                            return 'careProfiles';
                        };
                        return 'none';
                    },
                    function($item, $key) {
                        $serviceId = $item->service_id;
                        $profileId = $item->profile_id;
                        $assistanceTypeId = $item->assistance_type_id;
                        $careProfileId = $item->care_profile_id;
                        if($serviceId) {
                            return $serviceId;
                        };
                        if($profileId) {
                            return $profileId;
                        };
                        if($assistanceTypeId) {
                            return $assistanceTypeId;
                        };
                        if($careProfileId) {
                            return $careProfileId;
                        };
                        return 'none';
                    },
                ]);

                if(isset($value['hospital']['roundClock'])) {
                    if(isset($value['hospital']['roundClock']['vmp'])) {
                        $value['hospital']['roundClock']['vmp']['careProfiles'] = $value['hospital']['roundClock']['vmp']['careProfiles']->mapWithKeys(
                            function($item, $key) {
                                return [$key => $item->groupBy(['vmp_group_id','vmp_type_id'])];
                            }
                        );
                        $value['hospital']['roundClock']['vmp']['careProfiles']->transform(
                            function($careProfileData) {
                                return $careProfileData->transform(function($vmpGroupeData) {
                                    return $vmpGroupeData->transform(function($vmpTypeData) {
                                        return $vmpTypeData->mapWithKeys(function($item) {
                                            return [$item->indicator_id => $item->value];
                                        });
                                    });
                                });
                            }
                        );
                    }

                    if(isset($value['hospital']['roundClock']['regular'])) {
                        $value['hospital']['roundClock']['regular']['bedProfiles']->transform(
                            function($bedProfileData) {
                                return $bedProfileData->mapWithKeys(function($item) {
                                    return [$item->indicator_id => $item->value];
                                });
                            }
                        );
                    }
                }
                if(isset($value['hospital']['daytime'])) {
                    if(isset($value['hospital']['daytime']['inPolyclinic'])) {
                        $value['hospital']['daytime']['inPolyclinic']['bedProfiles']->transform(
                            function($bedProfileData) {
                                return $bedProfileData->mapWithKeys(function($item) {
                                    return [$item->indicator_id => $item->value];
                                });
                            }
                        );
                    }
                    if(isset($value['hospital']['daytime']['inHospital'])) {
                        $value['hospital']['daytime']['inHospital']['bedProfiles']->transform(
                            function($bedProfileData) {
                                return $bedProfileData->mapWithKeys(function($item) {
                                    return [$item->indicator_id => $item->value];
                                });
                            }
                        );
                    }
                }
            }

            if(isset($value['polyclinic'])) {
                $value['polyclinic'] = $value['polyclinic']->groupBy([
                    function($item, $key) use  ($polyclinicPerPersonNodeIds, $polyclinicPerUnitNodeIds, $polyclinicFapNodeIds) {
                        $nodeId = $item->node_id;
                        if(in_array($nodeId, $polyclinicFapNodeIds)) {
                            return 'fap';
                        };
                        if(in_array($nodeId, $polyclinicPerPersonNodeIds)) {
                            return 'perPerson';
                        };
                        if(in_array($nodeId, $polyclinicPerUnitNodeIds)) {
                            return 'perUnit';
                        };
                        return 'none';
                    },
                    function($item, $key) use  ($polyclinicFapNodeIds) {
                        $nodeId = $item->node_id;
                        $departmentId = $item->mo_department_id;
                        if(in_array($nodeId, $polyclinicFapNodeIds)) {
                            return $departmentId;
                        };
                        return 'all';
                    },
                    function($item, $key) use  ($polyclinicFapNodeIds) {
                        $serviceId = $item->service_id;
                        $profileId = $item->profile_id;
                        $assistanceTypeId = $item->assistance_type_id;
                        $careProfileId = $item->care_profile_id;
                        if($serviceId) {
                            return 'services';
                        };
                        if($profileId) {
                            return 'bedProfiles';
                        };
                        if($assistanceTypeId) {
                            return 'assistanceTypes';
                        };
                        if($careProfileId) {
                            return 'careProfiles';
                        };
                        return 'none';
                    },
                    function($item, $key) {
                        $serviceId = $item->service_id;
                        $profileId = $item->profile_id;
                        $assistanceTypeId = $item->assistance_type_id;
                        $careProfileId = $item->care_profile_id;
                        if($serviceId) {
                            return $serviceId;
                        };
                        if($profileId) {
                            return $profileId;
                        };
                        if($assistanceTypeId) {
                            return $assistanceTypeId;
                        };
                        if($careProfileId) {
                            return $careProfileId;
                        };
                        return 'none';
                    },
                ]);

                if(isset($value['polyclinic']['fap'])) {
                    $value['polyclinic']['fap']->transform(function($fapData) {
                        if(isset($fapData['services'])) {
                            $fapData['services']->transform(function($serviceData) {
                                return $serviceData->mapWithKeys(function($item) {
                                    return [$item->indicator_id => $item->value];
                                });
                            });
                        }
                        if(isset($fapData['assistanceTypes'])) {
                            $fapData['assistanceTypes']->transform(function($assistanceTypeData) {
                                return $assistanceTypeData->mapWithKeys(function($item) {
                                    return [$item->indicator_id => $item->value];
                                });
                            });
                        }
                        return $fapData;
                    });
                }

                if(isset($value['polyclinic']['perPerson'])) {
                    $moData = $value['polyclinic']['perPerson']['all'];
                    if(isset($moData['services'])) {
                        $moData['services']->transform(function($serviceData) {
                            return $serviceData->mapWithKeys(function($item) {
                                return [$item->indicator_id => $item->value];
                            });
                        });
                    }
                    if(isset($moData['assistanceTypes'])) {
                        $moData['assistanceTypes']->transform(function($assistanceTypeData) {
                            return $assistanceTypeData->mapWithKeys(function($item) {
                                return [$item->indicator_id => $item->value];
                            });
                        });
                    }
                }

                if(isset($value['polyclinic']['perUnit'])) {
                    $moData = $value['polyclinic']['perUnit']['all'];
                    if(isset($moData['services'])) {
                        $moData['services']->transform(function($serviceData) {
                            return $serviceData->mapWithKeys(function($item) {
                                return [$item->indicator_id => $item->value];
                            });
                        });
                    }
                    if(isset($moData['assistanceTypes'])) {
                        $moData['assistanceTypes']->transform(function($assistanceTypeData) {
                            return $assistanceTypeData->mapWithKeys(function($item) {
                                return [$item->indicator_id => $item->value];
                            });
                        });
                    }
                }
            }

            if(isset($value['ambulance'])) {
                $value['ambulance'] = $value['ambulance']->groupBy(['assistance_type_id']);

                foreach ($value['ambulance'] as $assistanceTypeId => $assistanceTypeData) {
                    $value['ambulance'][$assistanceTypeId] = $assistanceTypeData->mapWithKeys(
                        function($item, $key) {
                            return [$item->indicator_id => $item->value];
                        }
                    );
                }
            }
        }
        return $data;
    }
}

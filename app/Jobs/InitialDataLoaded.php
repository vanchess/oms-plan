<?php

namespace App\Jobs;

use App\Services\InitialDataFixingService;
use Illuminate\Bus\Queueable;
use Illuminate\Contracts\Queue\ShouldBeUnique;
use Illuminate\Contracts\Queue\ShouldQueue;
use Illuminate\Foundation\Bus\Dispatchable;
use Illuminate\Queue\InteractsWithQueue;
use Illuminate\Queue\SerializesModels;

class InitialDataLoaded implements ShouldQueue
{
    use Dispatchable, InteractsWithQueue, Queueable, SerializesModels;

    public $timeout = 60;
    protected $year;
    protected $nodeId;
    protected $userId;

    /**
     * Create a new job instance.
     *
     * @return void
     */
    public function __construct(int $nodeId, int $year, int $userId)
    {
        $this->year = $year;
        $this->nodeId = $nodeId;
        $this->userId = $userId;
    }

    /**
     * Execute the job.
     *
     * @return void
     */
    public function handle(InitialDataFixingService $initialDataFixingService)
    {
        $initialDataFixingService->commit($this->nodeId, $this->year, $this->userId);
    }
}

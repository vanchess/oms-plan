<?php

namespace App\Jobs;

use App\Services\PlannedIndicatorChangeInitService;
use Illuminate\Bus\Queueable;
use Illuminate\Contracts\Queue\ShouldBeUnique;
use Illuminate\Contracts\Queue\ShouldQueue;
use Illuminate\Foundation\Bus\Dispatchable;
use Illuminate\Queue\InteractsWithQueue;
use Illuminate\Queue\SerializesModels;

class InitialChanges implements ShouldQueue, ShouldBeUnique
{
    use Dispatchable, InteractsWithQueue, Queueable, SerializesModels;

    public $timeout = 60;
    protected $year;

    /**
     * Create a new job instance.
     *
     * @return void
     */
    public function __construct(int $year)
    {
        //
        $this->year = $year;
        $this->onQueue('one-thread');
    }

    /**
     * Execute the job.
     *
     * @return void
     */
    public function handle(PlannedIndicatorChangeInitService $plannedIndicatorChangeInitService)
    {
        $plannedIndicatorChangeInitService->fromInitialData($this->year);
    }
}

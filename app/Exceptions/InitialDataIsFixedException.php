<?php
declare(strict_types=1);

namespace App\Exceptions;

use Exception;

class InitialDataIsFixedException extends Exception
{
    /**
     * Render the exception as an HTTP response.
     *
     * @param  \Illuminate\Http\Request  $request
     * @return \Illuminate\Http\Response
     */
    public function render($request)
    {
        return response()->json([
            'error' => $this->getMessage()
        ], 423);
    }
}

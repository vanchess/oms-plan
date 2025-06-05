@extends('layouts.app')

@section('content')
    <div style="word-wrap: normal;">
        @include('tree', ['tree' => $tree, 'treeService' => $treeService, 'level' => 0])
    </div>
@endsection

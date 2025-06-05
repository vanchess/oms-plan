@if(is_array($tree) && count($tree) > 0)
    @foreach($tree as $k => $t)
        @php
            $mp = \App\Models\PumpMonitoringProfiles::find($k);
            $pu = $mp->profilesUnits;
            $relationType = \App\Models\PumpMonitoringProfilesRelationType::find($mp->relation_type_id)?->name;
            $marginLeft = ($level * 20) . 'px';
            $treeId = "tree-{$mp->id}";
        @endphp

        <p>
            <nobr style="margin-left: {{ $marginLeft }}; word-wrap: normal;">
                <button class="tree-toggle" onclick="toggleBlock('{{ $treeId }}', this)">▶</button>
                <b>{{ $mp->name }} [{{ $relationType }}]</b>
            </nobr>
        </p>

        <div id="{{ $treeId }}" style="display: none; margin-left: {{ $marginLeft }};">
            @foreach($pu as $u)
                @php
                    $pi = $u->plannedIndicators;
                    $piIdsViaChild = $treeService->plannedIndicatorIdsViaChild($mp->id, $u->unit->id);
                    $piIds = array_column($pi?->toArray(), 'id');
                    $piIdsViaChild = array_diff($piIdsViaChild, $piIds);
                    $piViaChild = \App\Models\PlannedIndicator::whereIn('id', $piIdsViaChild)->get();
                    $hasData = count($pi) > 0 || count($piIdsViaChild) > 0;
                    $unitBlockId = "block-{$mp->id}-{$u->unit->id}";
                @endphp

                <p>
                    <nobr style="margin-left: {{ $marginLeft }}; word-wrap: normal;">
                        <button class="unit-toggle" onclick="toggleBlock('{{ $unitBlockId }}', this)">➤</button>
                        - {{ $u->unit->name }}
                    </nobr>
                </p>

                <div id="{{ $unitBlockId }}" style="display: none; margin-left: {{ $marginLeft }};">
                    @if($hasData)
                        @foreach($pi as $i)
                            <p>- {{ plannedIndicatorName($i) }}</p>
                        @endforeach
                        @foreach($piViaChild as $i)
                            <p><i>- {{ plannedIndicatorName($i) }} (УНАСЛЕДОВАНО)</i></p>
                        @endforeach
                    @else
                        <p>- не утверждается</p>
                    @endif
                </div>
            @endforeach

            @include('tree', ['tree' => $t, 'treeService' => $treeService, 'level' => $level + 1])
        </div>
    @endforeach
@endif

<script>
    function toggleBlock(id, button) {
        let block = document.getElementById(id);

        if (block.style.display === "none") {
            block.style.display = "block";
            button.textContent = button.classList.contains("tree-toggle") ? "▼" : "⬇";
        } else {
            block.style.display = "none";
            button.textContent = button.classList.contains("tree-toggle") ? "▶" : "➤";
        }
    }
</script>

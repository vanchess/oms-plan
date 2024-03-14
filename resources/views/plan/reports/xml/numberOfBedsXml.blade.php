{!! '<'.'?xml version="1.0" encoding="windows-1251"?>' !!}
<Plan_koek>
@foreach ($moCollection as $mo)
    <LPU>
      <MCOD>{{ $mo->code }}</MCOD>
      <TS_DATA>
        <TS>1</TS>
@foreach ($hospitalBedProfiles as $profile)
        <PROFIL>
          <PR_OMS>{{$profile->code}}</PR_OMS>
          <KOIKI>{{sprintf('%d',$content['mo'][$mo->id]['hospital']['roundClock']['regular']['bedProfiles'][$profile->id][1] ?? 0)}}</KOIKI>
          <DATEBEG>{{$dateBeg}}</DATEBEG>
        </PROFIL>
@endforeach
      </TS_DATA>
    </LPU>
@endforeach
</Plan_koek>

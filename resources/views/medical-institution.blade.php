@extends('layouts.app')
@section('sidebar')
    @include('plan.sidebar')
@endsection
@section('content')
<div class="card">
  <div class="card-body">
    <h5 class="card-title"></h5>
    <p> {{ $medicalInstitution['name'] }} </p>
      <table style="width:100%;border-style:double;border-width:5px;">
          
          <tr style="height:100px;">
            <td colspan="2" style="text-align: center;border-style:double;border-width:5px;">
              <a class="nav-link active" href="{{ route('medicalInstitutionPeriod', ['id' => $medicalInstitution['id'], 'period' => 'year' ]) }}">Год</a>
            </td>
          </tr>
          <tr style="height:100px;">
            <td style="text-align: center;border-style:double;border-width:5px;">
              <a class="nav-link active" href="{{ route('medicalInstitutionPeriod', ['id' => $medicalInstitution['id'], 'period' => 'Q1' ]) }}">1 квартал</a>
            </td>
            <td style="text-align: center;border-style:double;border-width:5px;">
              <a class="nav-link active" href="{{ route('medicalInstitutionPeriod', ['id' => $medicalInstitution['id'], 'period' => 'Q2' ]) }}">2 квартал</a>
            </td>
          </tr>
          <tr style="height:100px;">
            <td style="text-align: center;border-style:double;border-width:5px;">
              <a class="nav-link active" href="{{ route('medicalInstitutionPeriod', ['id' => $medicalInstitution['id'], 'period' => 'Q3' ]) }}">3 квартал</a>
            </td>
            <td style="text-align: center;border-style:double;border-width:5px;">
              <a class="nav-link active" href="{{ route('medicalInstitutionPeriod', ['id' => $medicalInstitution['id'], 'period' => 'Q4' ]) }}">4 квартал</a>
            </td>
          </tr>
          
      </table>
  </div>
</div>
@endsection
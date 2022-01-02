@extends('layouts.app')
@section('sidebar')
    @include('plan.sidebar')
@endsection
@section('content')
<div class="card">
  <div class="card-body">
    <h5 class="card-title"></h5>
    <p>{{ $medicalInstitution['name'] }}. {{ $period }}</p>
      <ol>
        <li>АПП проф</li>
        <li>АПП заб (в том числе услуги)</li>
        <li>АПП неотложная</li>
        <li>Дневной стационар при поликлинике</li>
        <li>Дневной стационар при стационаре</li>
        <li>Круглосуточный стационар - обычный</li>
        <li>Круглосуточный стационар - ВМП</li>
        <li>Круглосуточный стационар - реабилитация</li>
        <li>Скорая</li>
      </ol>
      
      
      <p></p>
       <p></p>
      <table border="1" style="width:100%;">
          <tr>
            <th rowspan="3" class="first" style="width:10%; text-align: center;">АПП</th>
            <th colspan="2" style="width:15%; text-align: center;">Всего</th>
            <th colspan="2" style="text-align: center;">проф</th>
            <th colspan="2" style="text-align: center;">заб (в том числе услуги)</th>
            <th colspan="2" style="text-align: center;">неотложная</th>
          </tr>
          <tr>
            <td class="first" style="text-align: center;">Количество</td>
            <td class="first" style="text-align: center;">Сумма</td>
            <td class="first" style="text-align: center;">Количество</td>
            <td class="first" style="text-align: center;">Сумма</td>
            <td class="first" style="text-align: center;">Количество</td>
            <td class="first" style="text-align: center;">Сумма</td>
            <td class="first" style="text-align: center;">Количество</td>
            <td class="first" style="text-align: center;">Сумма</td>
          </tr>
          <tr>
            <td style="text-align: center;">3</td>
            <td style="text-align: center;">3</td>
            <td style="text-align: center;">1</td>
            <td style="text-align: center;">1</td>
            <td style="text-align: center;">1</td>
            <td style="text-align: center;">1</td>
            <td style="text-align: center;">1</td>
            <td style="text-align: center;">1</td>
          </tr>

          
        </table>
        
        <p></p>
       <p></p>
        <table border="1" style="width:100%;">
          <tr>
            <th rowspan="3" class="first" style="width:10%; text-align: center;">Дневной стационар</th>
            <th colspan="2" style="width:15%; text-align: center;">Всего</th>
            <th colspan="2" style="text-align: center;">Дневной стационар при поликлинике</th>
            <th colspan="2" style="text-align: center;">Дневной стационар при стационаре</th>
          </tr>
          <tr>
            <td class="first" style="text-align: center;">Количество</td>
            <td class="first" style="text-align: center;">Сумма</td>
            <td class="first" style="text-align: center;">Количество</td>
            <td class="first" style="text-align: center;">Сумма</td>
            <td class="first" style="text-align: center;">Количество</td>
            <td class="first" style="text-align: center;">Сумма</td>
          </tr>
          <tr>
            <td style="text-align: center;">2</td>
            <td style="text-align: center;">2</td>
            <td style="text-align: center;">1</td>
            <td style="text-align: center;">1</td>
            <td style="text-align: center;">1</td>
            <td style="text-align: center;">1</td>
          </tr>

          
        </table>
        
        <p></p>
       <p></p>
        <table border="1" style="width:100%;"> 
          <tr>
            <th rowspan="3" class="first" style="width:10%; text-align: center;">Круглосуточный стационар</th>
            <th colspan="2" style="width:15%; text-align: center;">Всего</th>
            <th colspan="2" style="text-align: center;">обычный</th>
            <th colspan="2" style="text-align: center;">ВМП</th>
            <th colspan="2" style="text-align: center;">реабилитация</th>
          </tr>
          <tr>
            <td class="first" style="text-align: center;">Количество</td>
            <td class="first" style="text-align: center;">Сумма</td>
            <td class="first" style="text-align: center;">Количество</td>
            <td class="first" style="text-align: center;">Сумма</td>
            <td class="first" style="text-align: center;">Количество</td>
            <td class="first" style="text-align: center;">Сумма</td>
            <td class="first" style="text-align: center;">Количество</td>
            <td class="first" style="text-align: center;">Сумма</td>
          </tr>
          <tr>
            <td style="text-align: center;">3</td>
            <td style="text-align: center;">3</td>
            <td style="text-align: center;">1</td>
            <td style="text-align: center;">1</td>
            <td style="text-align: center;">1</td>
            <td style="text-align: center;">1</td>
            <td style="text-align: center;">1</td>
            <td style="text-align: center;">1</td>
          </tr>

          
        </table>
        
       <p></p>
       <p></p>
        <table border="1" style="width:100%;">
          <tr>
            <th rowspan="3" class="first" style="width:10%; text-align: center;">Скорая</th>
            <th colspan="2" style="width:15%; text-align: center;">Всего</th>
            <th colspan="2" style="text-align: center;"></th>
          </tr>
          <tr>
            <td class="first" style="text-align: center;">Количество</td>
            <td class="first" style="text-align: center;">Сумма</td>
            <td colspan="2" class="first" style="text-align: center;"></td>
          </tr>
          <tr>
            <td style="text-align: center;">1</td>
            <td style="text-align: center;">1</td>
            <td colspan="2" style="text-align: center;"></td>
          </tr>

          
        </table>
        
  </div>
</div>
@endsection
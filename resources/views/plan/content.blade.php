<div class="card">
  <div class="card-body">
    <h5 class="card-title"></h5>
        <table>
          <caption>Таблица объемы</caption> 
          <tr>
            <th rowspan="2" class="first" style="width:100px; text-align: center;">ЛПУ</th>
            <th colspan="2" style="text-align: center;">Год</th>
            <th colspan="2" style="text-align: center;">1 квартал</th>
            <th colspan="2" style="text-align: center;">2 квартал</th>
            <th colspan="2" style="text-align: center;">3 квартал</th>
            <th colspan="2" style="text-align: center;">4 квартал</th>
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
            <td class="first" style="text-align: center;">Количество</td>
            <td class="first" style="text-align: center;">Сумма</td>
          </tr>
          @foreach ($medicalInstitutions as $medicalInstitution)
          <tr>
            <td class="first"><a class="nav-link active" href="{{ route('medicalInstitution', ['id' => $medicalInstitution['id'] ]) }}">{{ $medicalInstitution['short_name'] }}</a></td>
            <td>{{ $medicalInstitution['year_qty'] }}</td>
            <td>{{ $medicalInstitution['year_sum'] }}</td>
            <td>{{ $medicalInstitution['Q1_qty'] }}</td>
            <td>{{ $medicalInstitution['Q1_sum'] }}</td>
            <td>{{ $medicalInstitution['Q2_qty'] }}</td>
            <td>{{ $medicalInstitution['Q2_sum'] }}</td>
            <td>{{ $medicalInstitution['Q3_qty'] }}</td>
            <td>{{ $medicalInstitution['Q3_sum'] }}</td>
            <td>{{ $medicalInstitution['Q4_qty'] }}</td>
            <td>{{ $medicalInstitution['Q4_sum'] }}</td>
          </tr>
          @endforeach
        </table>
    
    
                            
    
  </div>
</div>
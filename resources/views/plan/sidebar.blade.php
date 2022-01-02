                  <div id="mainSidebar">
                    <div class="card">  
                      <div class="card-header" id="headingAdapters">
                        <button class="btn btn-link" type="button" data-toggle="collapse" data-target="#collapseOne" aria-expanded="false" aria-controls="collapseOne">
                          Главная
                        </button>
                      </div>
                      <div id="collapseOne" class="collapse" aria-labelledby="headingOne">
                        <div class="card-body">
                          <ul class="nav flex-column">
                            @foreach ($pages as $page)
                            <li class="nav-item"><a class="nav-link active" href="{{ route('page', ['id' => $page['id'] ]) }}">{{ $page['name'] }}</a></li>
                            @endforeach
                          </ul>
                        </div>
                      </div>
                    </div>
                    
                    <div class="card">
                      <div class="card-header" id="headingInformationType">
                        <button class="btn btn-link" type="button" data-toggle="collapse" data-target="#collapseTwo" aria-expanded="false" aria-controls="collapseTwo">
                            Ещё раздел
                        </button>
                      </div>
                      <div id="collapseTwo" class="collapse" aria-labelledby="headingTwo">
                        <div class="card-body">
                          <ul class="nav flex-column">
                            @foreach ($pages as $page)
                            <li class="nav-item"><a class="nav-link active" href="{{ route('page', ['id' => $page['id'] ]) }}">{{ $page['name'] }}</a></li>
                            @endforeach
                          </ul>
                        </div>
                      </div>
                    </div>
                   </div>
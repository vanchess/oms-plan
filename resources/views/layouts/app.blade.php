<!DOCTYPE html>
<html lang="{{ str_replace('_', '-', app()->getLocale()) }}">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
        
        <!-- Application styles-->
        <link rel="stylesheet" href="/css/app.css">
        
        <!-- Bootstrap -->
        <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/css/bootstrap.min.css" integrity="sha384-9aIt2nRpC12Uk9gS9baDl411NQApFmC26EwAOH8WgZl5MYYxFfc+NcPb1dKGj7Sk" crossorigin="anonymous">
        
        <title>@yield('title', 'Dog')</title>
    </head>
    <body class="d-flex flex-column min-vh-100">
        <header>
        {{-- @include('blocks.header') --}}
        </header>
        <div class="container-fluid flex-grow-1">
          <div class="row">  
            <div class="col-md-3 col-xl-2">
                <nav class="flex-column">
                  @yield('sidebar')
                </nav>
            </div>
            <main class="col-12 col-md-9 col-xl-10">
                @yield('content')
            </main>
          </div>  
        </div>
        <footer class="footer">
          <small>Copyright © <time datetime="2020">2021</time> ТФОМС Курганской области</small>
        </footer>
        
        <!-- Bootstrap -->
        <!-- JS, Popper.js, and jQuery --> 
        <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js" integrity="sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj" crossorigin="anonymous"></script>
        <script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.0/dist/umd/popper.min.js" integrity="sha384-Q6E9RHvbIyZFJoft+2mJbHaEWldlvI9IOYy5n3zV9zzTtmI3UksdQRVvoxMfooAo" crossorigin="anonymous"></script>
        <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/js/bootstrap.min.js" integrity="sha384-OgVRvuATP1z7JjHLkuOU7Xw704+h835Lr+6QL9UvYjZE3Ipu6Tp75j7Bh/kR0JKI" crossorigin="anonymous"></script>
        
    </body>
</html>

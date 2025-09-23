<!doctype html>
<html lang="en">

<head>
    <!-- Required meta tags -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet"
        integrity="sha384-1BmE4kWBq78iYhFldvKuhfTAU6auU8tT94WrHftjDbrCEXSU1oBoqyl2QvZ6jIW3" crossorigin="anonymous">

    <title>Hello, world!</title>
</head>

<body>

    <div class="container">

        <div class="row my-5">
            <div class="col-md-12">
                <h1>Non PLTS (Realisasi)</h1>
            </div>
        </div>
        <div class="row mb-5">
            <div class="col-md-12">
                <ul class="nav nav-pills">
                    <li class="nav-item">
                        <a class="nav-link" href="/">Toll Roads</a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link" href="/non-plts">Non PLTS</a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link" href="/plts">PLTS</a>
                    </li>
                </ul>
            </div>
        </div>
        <div class="row">
            <div class="col-md-5">
                <form action="data-non-plts" method="POST" enctype="multipart/form-data">
                    @csrf
                    <div style="display: flex; align-items: center; " class="mb-4">
                        <div class="me-2">
                            <label for="formFile" class="form-label">File</label>
                            <input class="form-control" type="file" id="formFile" name="file">
                        </div>
                        <div>
                            <label for="formFile" class="form-label">&nbsp;</label> <br>
                            <button type="submit" class="btn btn-primary">Upload</button>
                        </div>

                    </div>

                </form>
            </div>
        </div>
        <div class="row">
            <table class="table table-bordered border-primary">
                <thead>
                    <tr rowspan="1">
                        <th rowspan="2">Parameter</th>
                        <th rowspan="2">Satuan</th>
                        @foreach ($headers as $item)
                            @if ($item['year'] != null)
                                <th colspan="{{ 12 + 3 }}">
                                    {{ $item['year'] }}
                                </th>
                            @endif
                        @endforeach
                    </tr>
                    <tr>
                        @foreach ($headers as $item)
                            <th>
                                {{ $item['month'] }}
                            </th>
                        @endforeach
                    </tr>
                </thead>
                <tbody>
                    @foreach ($final_results as $item)
                        <tr>
                            <th>{{ $item['attribute'] }}</th>
                            <th>{{ $item['unit'] }}</th>
                            @foreach ($item['results'] as $r)
                                <th>
                                    <span style="color: blue">{{ $r['value'] }}</span> <small
                                        style="color:Red">{{ $r['cell_position'] }}</small>
                                    <br>
                                    {{-- <s>{{ $r['col_position'] }} {{ $r['row_position'] }}</s> --}}
                                </th>
                            @endforeach
                        </tr>
                    @endforeach
                </tbody>
            </table>

        </div>
    </div>

    <!-- Optional JavaScript; choose one of the two! -->





    <!-- Option 1: Bootstrap Bundle with Popper -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"
        integrity="sha384-ka7Sk0Gln4gmtz2MlQnikT1wXgYsOg+OMhuP+IlRH9sENBO0LRn5q+8nbTov4+1p" crossorigin="anonymous">
    </script>



    <!-- Option 2: Separate Popper and Bootstrap JS -->
    <!--
    <script src="https://cdn.jsdelivr.net/npm/@popperjs/core@2.10.2/dist/umd/popper.min.js"
        integrity="sha384-7+zCNj/IqJ95wo16oMtfsKbZ9ccEh31eOz1HGyDuCQ6wgnyJNSYdrPa03rtR1zdB" crossorigin="anonymous">
    </script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.min.js"
        integrity="sha384-QJHtvGhmr9XOIpI6YVutG+2QOK9T+ZnN4kzFN1RtK3zEFEIsxhlmWl5/YESvpZ13" crossorigin="anonymous">
    </script>
    -->
</body>

</html>

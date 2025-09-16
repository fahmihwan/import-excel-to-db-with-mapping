<?php

namespace App\Http\Controllers;

use App\Models\ProjectionTollRoad;
use Exception;
use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;




class TableDinamisController extends Controller
{

    public function home()
    {
        $realisation_projection_update_id = 1;

        // headers -> array
        $headerRow = ProjectionTollRoad::selectRaw('year, quartal, month, col_position')
            ->groupBy('month', 'year', 'quartal', 'col_position')
            ->where('realisation_projection_update_id', $realisation_projection_update_id)
            ->orderBy('col_position')
            ->get()
            ->toArray();

        // attributes -> array
        $attributesCol = ProjectionTollRoad::query()
            ->select('row_position', 'attribute')
            ->where('realisation_projection_update_id', $realisation_projection_update_id)
            ->distinct()
            ->orderBy('row_position')
            ->orderBy('attribute')
            ->get()
            ->toArray();

        // results -> array
        $results = ProjectionTollRoad::where('realisation_projection_update_id', $realisation_projection_update_id)
            ->select('col_position', 'row_position', 'cell_position', 'value', 'quartal', 'month')
            ->orderBy('col_position', 'asc')
            ->get()
            ->toArray();

        // index results per row_position
        $resultsByRow = [];
        foreach ($results as $r) {
            $rp = (string) $r['row_position'];
            if (!isset($resultsByRow[$rp])) $resultsByRow[$rp] = [];
            $resultsByRow[$rp][] = $r;
        }

        $final_results = [];

        foreach ($attributesCol as $attr) {
            $rowPos    = (string) $attr['row_position'];
            $attribute = (string) $attr['attribute'];

            $rows = $resultsByRow[$rowPos] ?? [];

            // --- filter (tanpa in_array, pakai switch) ---
            $filtered = [];
            foreach ($rows as $f) {
                $quartal = $f['quartal'];
                $value   = $f['value'];

                switch ($attribute) {
                    case '- Pendapatan Tol Triwulanan (Rp Milliar)':
                    case '- Biaya Operasional & Pemeliharaan Trwiulanan (Rp Milliar)':
                        // FY lolos; selain FY hanya yang bernilai
                        if ($quartal === 'FY' || ($value !== '' && $value !== null)) {
                            $filtered[] = $f;
                        }
                        break;

                    case '- Pendapatan Tol Tahunan (Rp Milliar)':
                        if ($f['quartal'] == "FY") {
                            $filtered[] = $f;
                        }
                        break;
                    case '- Biaya Operasional & Pemeliharaan Tahunan (Rp Milliar)':
                        if ($f['quartal'] == "FY") {
                            $filtered[] = $f;
                        }
                        break;
                    default:
                        $filtered[] = $f;
                }
            }

            // --- map + hitung colspan ---
            $result_data = [];
            foreach ($filtered as $res) {
                $quartal = $res['quartal'];
                $colspan = 0;

                switch ($attribute) {
                    case '- Pendapatan Tol Triwulanan (Rp Milliar)':
                    case '- Biaya Operasional & Pemeliharaan Trwiulanan (Rp Milliar)':
                        if ($quartal !== 'FY') {
                            $colspan = 3;
                        }
                        $result_data[] = [
                            'col_position' => $res['col_position'],
                            'row_position' => $res['row_position'],
                            'cell_position' => $res['cell_position'],
                            'value'        => $res['value'],
                            'quartal'      => $quartal,
                            'colspan'      => $colspan,
                        ];
                        break;

                    case '- Pendapatan Tol Tahunan (Rp Milliar)': //MERGE 12 COL / 12 TAHUN
                        $colspan = 0;
                        $result_data[] = [
                            'col_position' => $res['col_position'],
                            'row_position' => $res['row_position'],
                            'cell_position' => $res['cell_position'],
                            'value'        => $res['value'],
                            'quartal'      => $quartal,
                            'colspan'      => $colspan,
                        ];
                        $result_data[] = [
                            'col_position' => $res['col_position'],
                            'row_position' => $res['row_position'],
                            'cell_position' => $res['cell_position'],
                            'value'        => '',
                            'quartal'      => 'MERGE',
                            'colspan'      => 12,
                        ];
                        break;
                    case '- Biaya Operasional & Pemeliharaan Tahunan (Rp Milliar)':
                        $colspan = 0;
                        $result_data[] = [
                            'col_position' => $res['col_position'],
                            'row_position' => $res['row_position'],
                            'value'        => $res['value'],
                            'cell_position' => $res['cell_position'],
                            'quartal'      => $quartal,
                            'colspan'      => $colspan,
                        ];
                        $result_data[] = [
                            'col_position' => $res['col_position'],
                            'row_position' => $res['row_position'],
                            'cell_position' => $res['cell_position'],
                            'value'        => '',
                            'quartal'      => 'MERGE',
                            'colspan'      => 12,
                        ];
                        break;
                    default:
                        $colspan = 0;
                        $result_data[] = [
                            'col_position' => $res['col_position'],
                            'row_position' => $res['row_position'],
                            'value'        => $res['value'],
                            'cell_position' => $res['cell_position'],
                            'quartal'      => $quartal,
                            'colspan'      => $colspan,
                        ];
                }
            }


            $final_results[] = [
                'row_position' => $rowPos,
                'attribute'    => $attribute,
                'results'      => array_values($result_data),
            ];
        }

        // return $final_results;
        return view('welcome', [
            'attributes'    => $attributesCol,
            'headers'       => $headerRow,
            'results'       => $results,
            'final_results' => $final_results
        ]);
    }

    public function homeV4() //hampir fix, EXCLUDE MERGE
    {

        $realisation_projection_update_id = 1;

        // get headers
        $headerRow = ProjectionTollRoad::selectRaw('year, quartal, month,col_position',)
            ->groupBy('month', 'year', 'quartal', 'col_position')
            ->where('realisation_projection_update_id', $realisation_projection_update_id)
            ->orderBy('col_position') // opsional
            ->get()->toArray();



        //get attributes data column
        $attributesCol = ProjectionTollRoad::query()
            ->select('row_position', 'attribute')
            ->where('realisation_projection_update_id', $realisation_projection_update_id)
            ->distinct()
            ->orderBy('row_position')
            ->orderBy('attribute')
            ->get();

        // CARI QUARTAL

        // get all results by id
        $results = ProjectionTollRoad::where('realisation_projection_update_id', $realisation_projection_update_id)
            ->select('col_position', 'row_position', 'value', 'quartal', 'month')
            ->orderBy('col_position', 'asc')->get();


        $final_results = $attributesCol->map(function ($attr) use ($results) {


            $result_data = $results->filter(function ($r) use ($attr) {
                return $r->row_position == $attr->row_position; //mapping berdasarkan position

            })->filter(function ($f) use ($attr) { //mapping COLSPAN 
                if ($attr->attribute == '- Pendapatan Tol Triwulanan (Rp Milliar)') { //jika triwulan, untuk Q1 dll nilai kosong gausah di teruskan, karena untuk COLSPAN (merge column)
                    return $f->quartal == 'FY'  || $f->value != '';
                } else if ($attr->attribute == '- Biaya Operasional & Pemeliharaan Trwiulanan (Rp Milliar)') {
                    return $f->quartal == 'FY'  || $f->value != '';
                } else if ($attr->attribute == '- Pendapatan Tol Tahunan (Rp Milliar)') {
                    return $f->quartal == 'FY';
                } else {
                    return $f;
                }
            })->map((function ($res) use ($attr) {

                $colspan = 0;

                if ($attr->attribute == '- Pendapatan Tol Triwulanan (Rp Milliar)') { //jika triwulan, untuk Q1 dll nilai kosong gausah di teruskan, karena untuk COLSPAN (merge column)
                    if ($res->quartal != 'FY') {
                        $colspan = 3;
                    }
                } else if ($attr->attribute == '- Biaya Operasional & Pemeliharaan Trwiulanan (Rp Milliar)') {
                    if ($res->quartal != 'FY') {
                        $colspan = 3;
                    }
                } else if ($attr->attribute == '- Pendapatan Tol Tahunan (Rp Milliar)') {
                    // === LOGIC TAHUNAN: satu Q di antara FY jadi colspan 12 ===

                    if ($res->quartal === 'FY') {

                        // buka segmen: Q pertama setelah ini yang akan trigger
                    }
                    // (kalau belum ketemu Q lalu muncul FY lagi, gate tetap true dan reset segmen berikutnya otomatis)
                } else {
                    $colspan = 0;
                }




                return [
                    "col_position" => $res->col_position,
                    "row_position" => $res->row_position,
                    "value" => $res->value,
                    "quartal" => $res->quartal,
                    'colspan' => $colspan,
                ];
            }))->values()->toArray();


            return [
                'quartal' => $attr->quartal,
                'row_position'      => $attr->row_position,
                'attribute'     => $attr->attribute,
                'results'  => $result_data,
            ];
        });

        return $final_results;



        return view('welcome', [
            'attributes' => $attributesCol,
            'headers' => $headerRow,
            'results' => $results,
            'final_results' => $final_results
        ]);
    }


    // untuk saat startRow dan endRow itukan di muali dari  11 sampai 24 ya, dan di mulai nya dari kolom A, 
    // saya ingin di langsung mencari dari cell yang ada kalimat "TAHUN", kalau sudah ketemua nanti baru di pakai untuk batasan
    // jadi "TAHUN" nanti sebagai startRow dan startCol nya


    public function store_import(Request $request)
    {
        /**
         * Cari cell berisi teks (case-insensitive, trim) tanpa getCellByColumnAndRow.
         * Return: [row, colIndex] atau null jika tidak ketemu.
         */
        $funcFindCellByText = function (Worksheet $sheet, string $needle): ?array {
            $highestRow    = $sheet->getHighestRow();
            $highestColIdx = Coordinate::columnIndexFromString($sheet->getHighestColumn());
            $needleNorm    = mb_strtolower(trim($needle));

            for ($r = 1; $r <= $highestRow; $r++) {
                for ($c = 1; $c <= $highestColIdx; $c++) {
                    $coord = Coordinate::stringFromColumnIndex($c) . $r;
                    $val   = $sheet->getCell($coord)->getValue();
                    $valNorm = mb_strtolower(trim((string)$val));
                    if ($valNorm === $needleNorm) {
                        return [$r, $c];
                    }
                }
            }
            return null;
        };

        // file tidak sesuai


        /**
         * Ambil index kolom terakhir yang punya data pada rentang baris & dari kolom tertentu.
         * (sudah OK dan tidak pakai getCellByColumnAndRow)
         */
        $funcGetLastDataColIndex = function (
            Worksheet $sheet,
            int $startRow,
            int $endRow,
            int $startColIndex = 1
        ): int {
            $highestColIdx = Coordinate::columnIndexFromString($sheet->getHighestColumn());
            for ($c = $highestColIdx; $c >= $startColIndex; $c--) {
                $col = Coordinate::stringFromColumnIndex($c);
                $colVals = $sheet->rangeToArray("{$col}{$startRow}:{$col}{$endRow}", null, false, false, false);
                foreach ($colVals as $rowArr) {
                    $v = $rowArr[0] ?? null;
                    if ($v !== null && trim((string)$v) !== '') {
                        return $c;
                    }
                }
            }
            return $startColIndex;
        };

        /**
         * Ambil index baris terakhir yang punya data pada area mulai (startRow, startColIndex).
         * Tanpa getCellByColumnAndRow: pakai getCell($coord).
         */
        $funcGetLastDataRowIndex = function (
            Worksheet $sheet,
            int $startColIndex,
            int $startRow
        ): int {
            $highestRow    = $sheet->getHighestRow();
            $highestColIdx = Coordinate::columnIndexFromString($sheet->getHighestColumn());

            for ($r = $highestRow; $r >= $startRow; $r--) {
                for ($c = $highestColIdx; $c >= $startColIndex; $c--) {
                    $coord = Coordinate::stringFromColumnIndex($c) . $r;
                    $v = $sheet->getCell($coord)->getValue();
                    if ($v !== null && trim((string)$v) !== '') {
                        return $r;
                    }
                }
            }
            return $startRow;
        };




        $funcConvertionAlphabetToNumber = function ($text) {
            $text = strtoupper(trim($text));

            // pecah huruf & angka, contoh "ABZ200" => ["ABZ", "200"]
            if (!preg_match('/^([A-Z]+)(\d+)$/', $text, $matches)) {
                throw new Exception("Format cell tidak valid: $text");
            }

            $letters = $matches[1];    // huruf kolom
            $row     = (int)$matches[2]; // angka baris

            // konversi huruf ke angka kolom (A=1, ..., Z=26, AA=27, ABZ=754, dst.)
            $col = 0;
            for ($i = 0; $i < strlen($letters); $i++) {
                $col = $col * 26 + (ord($letters[$i]) - ord('A') + 1);
            }

            return [
                'letters'   => $letters, // huruf kolom
                'col_position' => $col,     // hasil konversi huruf → angka
                'row_position' => $row      // angka baris
            ];
        };



        /** Baca cell sebagai nilai final (formula → hasil), null jika kosong */
        $funcReadCell = function (Worksheet $sheet, int $colIndex, int $row) {
            $coord = Coordinate::stringFromColumnIndex($colIndex) . $row;
            $cell  = $sheet->getCell($coord);
            $val   = $cell->getValue();

            // RichText → string
            if ($val instanceof RichText) {
                $val = $val->getPlainText();
            }

            // Jika formula, hitung nilainya
            if (is_string($val) && strlen($val) > 0 && $val[0] === '=') {
                try {
                    $val = $cell->getCalculatedValue(); // ambil hasil formula
                } catch (\Exception $e) {
                    // fallback kalau formula gagal dihitung
                    $val = null;
                }
            }

            // Normalisasi kosong
            if ($val === null) return null;
            $valStr = trim((string)$val);
            return ($valStr === '') ? null : $valStr;
        };


        // TOOD
        // nambahin validasi X di file excel 
        // nambahin validasi position_col dan position_row



        // DOCS
        // 1. mengambil data excel, trigger dari text yang ada kalimat "TAHUN" lalu di trigger sebagai penentu untuk ROW START dan COL START
        // 2. untuk collspan, data tetap di insert setiap kolom, namun saat fetch mengambil yang FY saja tanpa Q1,Q2,dst, lalu hanya di sisipkan 1 array kosong yang di  gunakan untuk colspan 12
        // 3. trigger cell B60 isi data hanya 'xx' sebagai penanda file itu dari kita
        $validated =  $request->validate([
            'file' => 'required|mimes:xlsx,xls,csv,pdf,doc,docx|max:2048',
        ]);

        $file = $validated['file'];


        $spreadsheet = IOFactory::load($file->getRealPath());

        // baca tab spreadsheet
        $sheet = $spreadsheet->getSheetByName('Realisasi');
        if (!$sheet) {
            abort(422, 'Sheet "Realisasi" tidak ditemukan.');
        }

        // Baca nilai dari sel B60
        $value = $sheet->getCell('B60')->getValue();
        if ($value != 'xx') {
            abort(422, 'File tidak sesuai".');
        }

        // 1) Cari posisi “TAHUN” → dijadikan startRow & startColIndex
        $pos = $funcFindCellByText($sheet, 'TAHUN');
        if (!$pos) {
            abort(422, 'Label "TAHUN" tidak ditemukan di worksheet.');
        }

        [$startRow, $startColIndex] = $pos;

        // 2) Tentukan endRow & endColIndex secara dinamis
        $endRow      = $funcGetLastDataRowIndex($sheet, $startColIndex, $startRow);
        $endColIndex = $funcGetLastDataColIndex($sheet, $startRow, $endRow, $startColIndex);


        // --- ekstraksi data (kolom → baris) ---
        $data = [];
        for ($col = $startColIndex; $col <= $endColIndex; $col++) {
            $colData = [];
            for ($row = $startRow; $row <= $endRow; $row++) {
                $colData[] = $funcReadCell($sheet, $col, $row);
            }
            $data[] = $colData;
        }

        // $data[0] adalah KOLOM pertama dari area (startColIndex..endColIndex)
        $headerAttribut = $data[0];
        $totalHeader = count($headerAttribut) - 3; // exclude tahun, quartal, bulan (karena kolom ini tidak di store ke DB)
        if ($totalHeader < 1) {
            $totalHeader = 1;
        }

        $result = [];

        for ($i = 1; $i < count($data); $i++) {
            $row = $data[$i];
            $year    = $row[0] ?? null;
            $quartal = $row[1] ?? null;
            $month   = $row[2] ?? null;

            for ($j = 3; $j < count($headerAttribut); $j++) {

                $val = $row[$j] ?? null;
                $val = ($val === null) ? null : trim((string)$val);

                if ($year != null && $headerAttribut[$j] != null && $headerAttribut[$j] != 'xx') { //penanda itu file dari kita ada xx di cell B60

                    // >>> koordinat cell aslinya (mis. "B11")
                    $actualColIndex = $startColIndex + $i;   // i=1 berarti kolom di kanan start
                    $actualRowIndex = $startRow + $j;        // j=3 berarti 3 baris di bawah start
                    $cellRef = Coordinate::stringFromColumnIndex($actualColIndex) . $actualRowIndex;

                    $position =  $funcConvertionAlphabetToNumber($cellRef);

                    $result[] = [
                        'year'      => $year,
                        'quartal'   => $quartal,
                        'month'     => $month,
                        'attribute' => $headerAttribut[$j] ?? null,
                        'value'     => $val,
                        'is_show'   => false,
                        'row_position' =>  $position['row_position'],
                        'col_position'  =>  $position['col_position'],
                        'cell_position' => $cellRef,
                        'realisation_projection_update_id' => 1
                    ];
                }
            }
        }

        $rows = $result;

        // Hitung jumlah kolom dari satu row insert (untuk batas parameter)
        $cols = max(1, count($rows[0] ?? []));

        // Batas aman parameter (SQL Server 2100), margin 80%
        $maxWanted = 1000;
        $chunkSize = max(100, min($maxWanted, (int) floor((2100 * 0.8) / $cols)));

        $x = collect($rows)
            ->chunk($chunkSize)
            ->each(function ($chunk) {
                ProjectionTollRoad::insert($chunk->toArray());
            });

        return $x;
    }





    // public function store_import(Request $request)
    // {
    //     $validated =  $request->validate([
    //         'file' => 'required|mimes:xlsx,xls,csv,pdf,doc,docx|max:2048',
    //     ]);

    //     $file = $validated['file'];
    //     $spreadsheet = IOFactory::load($file->getRealPath());

    //     $sheet = $spreadsheet->getSheetByName('Realisasi');

    //     function getLastDataColIndex(
    //         Worksheet $sheet,
    //         int $startRow,
    //         int $endRow,
    //         int $startColIndex = 1
    //     ): int {
    //         $highestColIdx = Coordinate::columnIndexFromString($sheet->getHighestColumn());
    //         for ($c = $highestColIdx; $c >= $startColIndex; $c--) {
    //             $col = Coordinate::stringFromColumnIndex($c);
    //             $colVals = $sheet->rangeToArray("{$col}{$startRow}:{$col}{$endRow}", null, true, false, false);

    //             foreach ($colVals as $rowArr) {
    //                 $v = $rowArr[0] ?? null;
    //                 if ($v !== null && trim((string)$v) !== '') return $c;
    //             }
    //         }
    //         return $startColIndex;
    //     }


    //     $startRow = 11;
    //     $endRow   = 24;

    //     $endColIndex = getLastDataColIndex($sheet, $startRow, $endRow, 1); // mulai dari kolom A
    //     $data = [];

    //     // loop kolom dulu (kiri ke kanan)
    //     for ($col = 1; $col <= $endColIndex; $col++) {
    //         $colData = [];

    //         // dalam tiap kolom, ambil baris 11–14 (atas ke bawah)
    //         for ($row = $startRow; $row <= $endRow; $row++) {
    //             $cell = $sheet->getCell(Coordinate::stringFromColumnIndex($col) . $row);
    //             $colData[] = $cell->getValue();
    //         }

    //         $data[] = $colData;
    //     }


    //     $headerAttribut = $data[0]; // baris pertama
    //     $totalHeader = count($headerAttribut) - 3; //ambil total header : exclude tahun, bulan, null
    //     $result = [];
    //     $row_position = 0;
    //     $col_position = 1;


    //     for ($i = 1; $i < count($data); $i++) {
    //         $row = $data[$i];
    //         $year = $row[0];
    //         $quartal = $row[1];

    //         $month = $row[2];

    //         for ($j = 3; $j < count($headerAttribut); $j++) {

    //             if ($row_position >= $totalHeader) { // untuk nomor urut sesuai attribute yang ada
    //                 $row_position = 1;
    //                 $col_position++;
    //             } else {
    //                 $row_position++;
    //             }

    //             $result[] = [
    //                 'year'      => $year,
    //                 'quartal'   => $quartal,
    //                 'month' => $month,
    //                 'attribute' => $headerAttribut[$j],
    //                 'value'     => trim((string) $row[$j])  ?? null,
    //                 'is_show' => false,
    //                 'row_position' =>  $row_position,
    //                 'col_position' => $col_position,
    //                 'realisation_projection_update_id' => 1
    //             ];
    //         }
    //     }

    //     $rows = $result;

    //     // Hitung jumlah kolom dari baris pertama
    //     $cols = max(1, count($rows[0] ?? []));

    //     // Pakai batas maksimum 1000 tapi tetap aman dari limit sql server (error limit 2100) (pakai margin 80%)
    //     $maxWanted = 1000;
    //     $chunkSize = max(100, min($maxWanted, (int) floor((2100 * 0.8) / $cols)));

    //     $x = collect($rows)
    //         ->chunk($chunkSize)
    //         ->each(function ($chunk) {
    //             ProjectionTollRoad::insert($chunk->toArray());
    //         });


    //     return $x;
    // }





    // public function store_import_fix_but_not_row(Request $request)
    // {
    //     $validated =  $request->validate([
    //         'file' => 'required|mimes:xlsx,xls,csv,pdf,doc,docx|max:2048',
    //     ]);

    //     $file = $validated['file'];
    //     $spreadsheet = IOFactory::load($file->getRealPath());

    //     $sheet = $spreadsheet->getSheetByName('Realisasi');

    //     function getLastDataColIndex(
    //         Worksheet $sheet,
    //         int $startRow,
    //         int $endRow,
    //         int $startColIndex = 1
    //     ): int {
    //         $highestColIdx = Coordinate::columnIndexFromString($sheet->getHighestColumn());
    //         for ($c = $highestColIdx; $c >= $startColIndex; $c--) {
    //             $col = Coordinate::stringFromColumnIndex($c);
    //             $colVals = $sheet->rangeToArray("{$col}{$startRow}:{$col}{$endRow}", null, true, false, false);

    //             foreach ($colVals as $rowArr) {
    //                 $v = $rowArr[0] ?? null;
    //                 if ($v !== null && trim((string)$v) !== '') return $c;
    //             }
    //         }
    //         return $startColIndex;
    //     }


    //     $startRow = 11;
    //     $endRow   = 24;

    //     $endColIndex = getLastDataColIndex($sheet, $startRow, $endRow, 1); // mulai dari kolom A
    //     $data = [];

    //     // loop kolom dulu (kiri ke kanan)
    //     for ($col = 1; $col <= $endColIndex; $col++) {
    //         $colData = [];

    //         // dalam tiap kolom, ambil baris 11–14 (atas ke bawah)
    //         for ($row = $startRow; $row <= $endRow; $row++) {
    //             $cell = $sheet->getCell(Coordinate::stringFromColumnIndex($col) . $row);
    //             $colData[] = $cell->getValue();
    //         }

    //         $data[] = $colData;
    //     }


    //     $headerAttribut = $data[0]; // baris pertama
    //     $totalHeader = count($headerAttribut) - 3; //ambil total header : exclude tahun, bulan, null
    //     $result = [];
    //     $row_position = 0;
    //     $col_position = 1;


    //     for ($i = 1; $i < count($data); $i++) {
    //         $row = $data[$i];
    //         $year = $row[0];
    //         $quartal = $row[1];

    //         $month = $row[2];

    //         for ($j = 3; $j < count($headerAttribut); $j++) {

    //             if ($row_position >= $totalHeader) { // untuk nomor urut sesuai attribute yang ada
    //                 $row_position = 1;
    //                 $col_position++;
    //             } else {
    //                 $row_position++;
    //             }

    //             $result[] = [
    //                 'year'      => $year,
    //                 'quartal'   => $quartal,
    //                 'month' => $month,
    //                 'attribute' => $headerAttribut[$j],
    //                 'value'     => trim((string) $row[$j])  ?? null,
    //                 'is_show' => false,
    //                 'row_position' =>  $row_position,
    //                 'col_position' => $col_position,
    //                 'realisation_projection_update_id' => 1
    //             ];
    //         }
    //     }

    //     $rows = $result;

    //     // Hitung jumlah kolom dari baris pertama
    //     $cols = max(1, count($rows[0] ?? []));

    //     // Pakai batas maksimum 1000 tapi tetap aman dari limit sql server (error limit 2100) (pakai margin 80%)
    //     $maxWanted = 1000;
    //     $chunkSize = max(100, min($maxWanted, (int) floor((2100 * 0.8) / $cols)));

    //     $x = collect($rows)
    //         ->chunk($chunkSize)
    //         ->each(function ($chunk) {
    //             ProjectionTollRoad::insert($chunk->toArray());
    //         });


    //     return $x;
    // }
}

    




//     public function store_import_hardcode_row(Request $request)
//     {
//         $validated =  $request->validate([
//             'file' => 'required|mimes:xlsx,xls,csv,pdf,doc,docx|max:2048',
//         ]);

//         $file = $validated['file'];
//         $spreadsheet = IOFactory::load($file->getRealPath());

//         $sheet = $spreadsheet->getSheetByName('Realisasi');



//         $startRow = 11;
//         $endRow   = 24;

//         $endColIndex = $this->getLastDataColIndex($sheet, $startRow, $endRow, 1); // mulai dari kolom A

//         $data = [];

//         // loop kolom dulu (kiri ke kanan)
//         for ($col = 1; $col <= $endColIndex; $col++) {
//             $colData = [];

//             // dalam tiap kolom, ambil baris 11–14 (atas ke bawah)
//             for ($row = $startRow; $row <= $endRow; $row++) {
//                 $cell = $sheet->getCell(Coordinate::stringFromColumnIndex($col) . $row);
//                 $colData[] = $cell->getValue();
//             }

//             $data[] = $colData;
//         }


//         $headerAttribut = $data[0]; // baris pertama
//         $totalHeader = count($headerAttribut) - 3; //ambil total header : exclude tahun, bulan, null
//         $result = [];
//         $row_position = 0;
//         $col_position = 1;


//         for ($i = 1; $i < count($data); $i++) {
//             $row = $data[$i];
//             $year = $row[0];
//             $quartal = $row[1];

//             $month = $row[2];

//             for ($j = 3; $j < count($headerAttribut); $j++) {

//                 if ($row_position >= $totalHeader) { // untuk nomor urut sesuai attribute yang ada
//                     $row_position = 1;
//                     $col_position++;
//                 } else {
//                     $row_position++;
//                 }

//                 $result[] = [
//                     'year'      => $year,
//                     'quartal'   => $quartal,
//                     'month' => $month,
//                     'attribute' => $headerAttribut[$j],
//                     'value'     => trim((string) $row[$j])  ?? null,
//                     'is_show' => false,
//                     'row_position' =>  $row_position,
//                     'col_position' => $col_position,
//                     'realisation_projection_update_id' => 1
//                 ];
//             }
//         }

//         $rows = $result;

//         // Hitung jumlah kolom dari baris pertama
//         $cols = max(1, count($rows[0] ?? []));

//         // Pakai batas maksimum 1000 tapi tetap aman dari limit sql server (error limit 2100) (pakai margin 80%)
//         $maxWanted = 1000;
//         $chunkSize = max(100, min($maxWanted, (int) floor((2100 * 0.8) / $cols)));

//         $x = collect($rows)
//             ->chunk($chunkSize)
//             ->each(function ($chunk) {
//                 ProjectionTollRoad::insert($chunk->toArray());
//             });


//         return $x;
//     }

//     public function store_import_hardcode_col_row(Request $request)
//     {

//         $validated =  $request->validate([
//             'file' => 'required|mimes:xlsx,xls,csv,pdf,doc,docx|max:2048',
//         ]);

//         $file = $validated['file'];
//         $spreadsheet = IOFactory::load($file->getRealPath());

//         $sheet = $spreadsheet->getSheetByName('Realisasi');


//         $startRow = 11;
//         $endRow   = 24;
//         // $endCol   = 'MA'; //harus otomatis ngambil data col 
//         $endCol = "N";
//         $endColIndex = Coordinate::columnIndexFromString($endCol);

//         $data = [];

//         // loop kolom dulu (kiri ke kanan)
//         for ($col = 1; $col <= $endColIndex; $col++) {
//             $colData = [];

//             // dalam tiap kolom, ambil baris 11–14 (atas ke bawah)
//             for ($row = $startRow; $row <= $endRow; $row++) {
//                 $cell = $sheet->getCell(Coordinate::stringFromColumnIndex($col) . $row);
//                 $colData[] = $cell->getValue();
//             }

//             $data[] = $colData;
//         }


//         $headerAttribut = $data[0]; // baris pertama
//         $totalHeader = count($headerAttribut) - 3; //ambil total header : exclude tahun, bulan, null
//         $result = [];
//         $row_position = 0;
//         $col_position = 1;


//         for ($i = 1; $i < count($data); $i++) {
//             $row = $data[$i];
//             $year = $row[0];
//             $quartal = $row[1];

//             $month = $row[2];

//             for ($j = 3; $j < count($headerAttribut); $j++) {

//                 if ($row_position >= $totalHeader) { // untuk nomor urut sesuai attribute yang ada
//                     $row_position = 1;
//                     $col_position++;
//                 } else {
//                     $row_position++;
//                 }

//                 $result[] = [
//                     'year'      => $year,
//                     'quartal'   => $quartal,
//                     'month' => $month,
//                     'attribute' => $headerAttribut[$j],
//                     'value'     => (string) $row[$j] ?? null,
//                     'is_show' => false,
//                     'row_position' =>  $row_position,
//                     'col_position' => $col_position,
//                     'realisation_projection_update_id' => 1
//                 ];
//             }
//         }


//         foreach ($result as $x) {
//             ProjectionTollRoad::create($x);
//         }

//         return $result;
//     }
// }



    // function numToLetters($num)
    //     {
    //         $letters = '';
    //         while ($num > 0) {
    //             $mod = ($num - 1) % 26;
    //             $letters = chr(65 + $mod) . $letters;
    //             $num = (int)(($num - $mod) / 26);
    //         }
    //         return $letters;
    //     }

    //     function lettersToNum($letters)
    //     {
    //         $num = 0;
    //         $len = strlen($letters);
    //         for ($i = 0; $i < $len; $i++) {
    //             $num = $num * 26 + (ord($letters[$i]) - 64);
    //         }
    //         return $num;
    //     }


        // numToLetters(1);   // A
        // numToLetters(26);  // Z
        // numToLetters(27);  // AA
        // numToLetters(52);  // AZ
        // numToLetters(326); // MN

        // lettersToNum("A");   // 1
        // lettersToNum("Z");   // 26
        // lettersToNum("AA");  // 27
        // lettersToNum("MN");  // 326
<?php

namespace App\Http\Controllers;

use App\Models\ProjectionElectriciteNonPlts;
use App\Models\ProjectionTollRoad;
use Exception;
use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpParser\Node\Stmt\Return_;

class TableDinamisController extends Controller
{

    public function toll_roads()
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
        return view('toll_roads', [
            'attributes'    => $attributesCol,
            'headers'       => $headerRow,
            'results'       => $results,
            'final_results' => $final_results
        ]);
    }


    // untuk saat startRow dan endRow itukan di muali dari  11 sampai 24 ya, dan di mulai nya dari kolom A, 
    // saya ingin di langsung mencari dari cell yang ada kalimat "TAHUN", kalau sudah ketemua nanti baru di pakai untuk batasan
    // jadi "TAHUN" nanti sebagai startRow dan startCol nya
    public function store_import_toll_roads(Request $request)
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


            $val = $cell->getFormattedValue();

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
        $sheet = $spreadsheet->getSheetByName('Proyeksi Originasi');
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

            // $j di mulai dari 3, karena hanya mengambil dari baris ke 3 saja
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





    public function non_plts()
    {
        $realisation_projection_update_id = 1;
        // headers -> array
        $headerRow = ProjectionElectriciteNonPlts::selectRaw('year, col_position')
            ->groupBy('year', 'col_position')
            ->where('realisation_projection_update_id', $realisation_projection_update_id)
            ->orderBy('col_position')
            ->get()
            ->toArray();
        // return $headerRow;


        // attributes -> array
        $attributesCol = ProjectionElectriciteNonPlts::query()
            ->select('row_position', 'attribute', 'unit')
            ->where('realisation_projection_update_id', $realisation_projection_update_id)
            ->distinct()
            ->orderBy('row_position')
            ->orderBy('attribute')
            ->get()
            ->toArray();

        // results -> array
        $results = ProjectionElectriciteNonPlts::where('realisation_projection_update_id', $realisation_projection_update_id)
            ->select('col_position', 'row_position', 'cell_position', 'value', 'unit')
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
            $unit = (string) $attr['unit'];

            $rows = $resultsByRow[$rowPos] ?? [];

            // --- filter (tanpa in_array, pakai switch) ---
            $filtered = [];
            foreach ($rows as $f) {
                // $quartal = $f['quartal'];
                // $value   = $f['value'];

                switch ($attribute) {
                    // case '- Pendapatan Tol Triwulanan (Rp Milliar)':
                    // case '- Biaya Operasional & Pemeliharaan Trwiulanan (Rp Milliar)':
                    //     // FY lolos; selain FY hanya yang bernilai
                    //     if ($quartal === 'FY' || ($value !== '' && $value !== null)) {
                    //         $filtered[] = $f;
                    //     }
                    //     break;

                    // case '- Pendapatan Tol Tahunan (Rp Milliar)':
                    //     if ($f['quartal'] == "FY") {
                    //         $filtered[] = $f;
                    //     }
                    //     break;
                    // case '- Biaya Operasional & Pemeliharaan Tahunan (Rp Milliar)':
                    //     if ($f['quartal'] == "FY") {
                    //         $filtered[] = $f;
                    //     }
                    //     break;
                    default:
                        $filtered[] = $f;
                }
            }
            // return $filtered;

            // --- map + hitung colspan ---
            $result_data = [];
            foreach ($filtered as $res) {
                // $quartal = $res['quartal'];
                // $colspan = 0;

                switch ($attribute) {
                    // case '- Pendapatan Tol Triwulanan (Rp Milliar)':
                    // case '- Biaya Operasional & Pemeliharaan Trwiulanan (Rp Milliar)':
                    //     if ($quartal !== 'FY') {
                    //         $colspan = 3;
                    //     }
                    //     $result_data[] = [
                    //         'col_position' => $res['col_position'],
                    //         'row_position' => $res['row_position'],
                    //         'cell_position' => $res['cell_position'],
                    //         'value'        => $res['value'],
                    //         'quartal'      => $quartal,
                    //         'colspan'      => $colspan,
                    //     ];
                    //     break;

                    // case '- Pendapatan Tol Tahunan (Rp Milliar)': //MERGE 12 COL / 12 TAHUN
                    //     $colspan = 0;
                    //     $result_data[] = [
                    //         'col_position' => $res['col_position'],
                    //         'row_position' => $res['row_position'],
                    //         'cell_position' => $res['cell_position'],
                    //         'value'        => $res['value'],
                    //         'quartal'      => $quartal,
                    //         'colspan'      => $colspan,
                    //     ];
                    //     $result_data[] = [
                    //         'col_position' => $res['col_position'],
                    //         'row_position' => $res['row_position'],
                    //         'cell_position' => $res['cell_position'],
                    //         'value'        => '',
                    //         'quartal'      => 'MERGE',
                    //         'colspan'      => 12,
                    //     ];
                    //     break;
                    // case '- Biaya Operasional & Pemeliharaan Tahunan (Rp Milliar)':
                    //     $colspan = 0;
                    //     $result_data[] = [
                    //         'col_position' => $res['col_position'],
                    //         'row_position' => $res['row_position'],
                    //         'value'        => $res['value'],
                    //         'cell_position' => $res['cell_position'],
                    //         'quartal'      => $quartal,
                    //         'colspan'      => $colspan,
                    //     ];
                    //     $result_data[] = [
                    //         'col_position' => $res['col_position'],
                    //         'row_position' => $res['row_position'],
                    //         'cell_position' => $res['cell_position'],
                    //         'value'        => '',
                    //         'quartal'      => 'MERGE',
                    //         'colspan'      => 12,
                    //     ];
                    //     break;
                    default:
                        $colspan = 0;
                        $result_data[] = [
                            'col_position' => $res['col_position'],
                            'row_position' => $res['row_position'],
                            'value'        => $res['value'],
                            'cell_position' => $res['cell_position'],
                            // 'quartal'      => $quartal,
                            // 'colspan'      => $colspan,
                        ];
                }
            }

            $final_results[] = [
                'row_position' => $rowPos,
                'attribute'    => $attribute,
                'unit' => $unit,
                'results'      => array_values($result_data),
            ];
        }

        return view('non_plts', [
            'attributes'    => $attributesCol,
            'headers'       => $headerRow,
            'results'       => $results,
            'final_results' => $final_results
        ]);
    }

    public function store_import_non_plts(Request $request)
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


            $val = $cell->getFormattedValue();

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
        $sheet = $spreadsheet->getSheetByName('Proyeksi_nonPLTS');
        if (!$sheet) {
            abort(422, 'Sheet "Proyeksi_nonPLTS" tidak ditemukan.');
        }

        // Baca nilai dari sel B60
        $value = $sheet->getCell('B60')->getValue();
        if ($value != 'xx') {
            abort(422, 'File tidak sesuai".');
        }

        // 1) Cari posisi text “PARAMETER” → dijadikan startRow & startColIndex
        $pos = $funcFindCellByText($sheet, 'Parameter');
        if (!$pos) {
            abort(422, 'Label "Parameter" tidak ditemukan di worksheet.');
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
        $headerAttribut = $data[0]; //attribute
        $headerUnit = $data[1]; //satuan

        $totalHeader = count($headerAttribut) - 2; // exclude parameter (karena kolom ini tidak di store ke DB)

        if ($totalHeader < 1) {
            $totalHeader = 1;
        }

        // return $data; //isi array dari atas ke bawah lalu kesamping. index sudah sesuai row

        $result = [];
        for ($i = 1; $i < count($data); $i++) {
            $row = $data[$i];

            $year   = $row[1] ?? null; //artinya tahun ada di baris ke 2

            // $j di mulai dari 2, karena hanya mengambil dari baris ke 2 saja
            for ($j = 2; $j < count($headerAttribut); $j++) {

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
                        // 'quartal'   => $quartal,
                        // 'month'     => $month,
                        'unit' => $headerUnit[$j] ?? null,
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

        // return $rows;
        // Hitung jumlah kolom dari satu row insert (untuk batas parameter)
        $cols = max(1, count($rows[0] ?? []));

        // Batas aman parameter (SQL Server 2100), margin 80%
        $maxWanted = 1000;
        $chunkSize = max(100, min($maxWanted, (int) floor((2100 * 0.8) / $cols)));

        $x = collect($rows)
            ->chunk($chunkSize)
            ->each(function ($chunk) {
                ProjectionElectriciteNonPlts::insert($chunk->toArray());
            });

        return $x;
    }
}

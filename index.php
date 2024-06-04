<?php

include 'vendor/autoload.php';

$reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
$spreadsheet = $reader->load('products.xlsx');

foreach ($spreadsheet->getSheet(0)->toArray() as $index => $row) {
    if ($index == 0) {
        continue;
    }

    $spreadsheet->getSheet(0)->setCellValue('Z' . ($index + 1), $row[24]);
    $spreadsheet->getSheet(0)->setCellValue('Y' . ($index + 1), '');

    $spreadsheet->getSheet(0)->setCellValue('AB' . ($index + 1), $row[17]);
    $spreadsheet->getSheet(0)->setCellValue('R' . ($index + 1), '');

    if (empty($row[32])) {
        continue;
    }

    $titleRu = mb_substr($row[1], 0, 106) . ' O_o';
    $titleUa = mb_substr($row[2], 0, 126) . ' O_o';

    $spreadsheet->getSheet(0)->setCellValue('B' . ($index + 1), $titleRu);
    $spreadsheet->getSheet(0)->setCellValue('C' . ($index + 1), $titleUa);
}

foreach ($spreadsheet->getSheet(1)->toArray() as $index => $row) {
    if ($index == 0) {
        continue;
    }

    $spreadsheet->getSheet(1)->setCellValue('D' . ($index + 1), $row[0]);
    $spreadsheet->getSheet(1)->setCellValue('A' . ($index + 1), '');
}

$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->save('products_new.xlsx');

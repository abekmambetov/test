<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class XlsxCreator {

    public function __construct(Spreadsheet $spreadsheet) {
        $this->spreadsheet = $spreadsheet;

        # Массив английских символов для столбцов
        $this->alphachar = array_merge(range('A', 'Z'), range('a', 'z'));
    }

    public function parseJson($filename) {
        $string = file_get_contents($filename);
        $json_a = json_decode($string, true);
 
        return $json_a;
    }

    public function createHeaders($headers) {
        $this->spreadsheet->setActiveSheetIndex(0);

        for ($i = 0; $i < count($headers); $i++) {
            $this->spreadsheet->getActiveSheet()->setCellValue($this->alphachar[$i] . "1", $headers[$i]);
        }

        $this->writer = new Xlsx($this->spreadsheet);
        
    }

    public function createRows($data) {
        
        // Начинаем со второй строки потому что первая строка это заголовок ячейки
        $rowNum = 2;
        foreach($data as $row) {
            for ($i = 0; $i < count($row); $i++) {

                $this->spreadsheet->getActiveSheet()->setCellValue($this->alphachar[$i] . $rowNum, $row[$i]);
            }
            $rowNum++;
        }

        $this->writer->save('table.xlsx');
    }

}




$spreadsheet = new Spreadsheet();

$XlsxCreator = new XlsxCreator($spreadsheet);

$headers = $XlsxCreator->parseJson("header.json");
$XlsxCreator->createHeaders($headers);

$data = $XlsxCreator->parseJson("data.json");
$XlsxCreator->createRows($data);

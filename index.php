<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class XlsxFileCreator {

    public function __construct(Spreadsheet $spreadsheet) {
        $this->spreadsheet = $spreadsheet;

        # Массив английских символов для столбцов
        $this->alphachar = array_merge(range('A', 'Z'), range('a', 'z'));
    }

    public function parseJson(string $filename): array {
        $string = file_get_contents($filename);
        $json_a = json_decode($string, true);
 
        return $json_a;
    }

    public function addHeaders(array $headers) {
        
        $this->spreadsheet->setActiveSheetIndex(0);

        for ($i = 0; $i < count($headers); $i++) {
            $this->spreadsheet->getActiveSheet()->setCellValue($this->alphachar[$i] . "1", $headers[$i]);
        }

        $this->countColumns = count($headers);
    }

    public function addRows(array $data) {
        
        // Начинаем со второй строки потому что первая строка это заголовок ячейки
        $rowNum = 2;
        foreach($data as $row) {
            for ($i = 0; $i < count($row); $i++) {

                $this->spreadsheet->getActiveSheet()->setCellValue($this->alphachar[$i] . $rowNum, $row[$i]);
            }
            $rowNum++;
        }
    }

    public function createXlsxFile() {

        // Ставим необходимую ширину для ячеек начиная с A и заканчивая последним литералом из массива $this->alphachar по количеству колонок $this->countColumns в таблице
        foreach(range('A', $this->alphachar[$this->countColumns-1]) as $columnID) {
            $this->spreadsheet->getActiveSheet()->getColumnDimension($columnID)
                ->setAutoSize(true);
        }
        
        $this->writer = new Xlsx($this->spreadsheet);
        $this->writer->save('table.xlsx');
    }
}

$XlsxFileCreator = new XlsxFileCreator(new Spreadsheet());

$headers = $XlsxFileCreator->parseJson(filename: "header.json");
$XlsxFileCreator->addHeaders($headers);

$data = $XlsxFileCreator->parseJson(filename: "data.json");
$XlsxFileCreator->addRows($data);

$XlsxFileCreator->createXlsxFile();

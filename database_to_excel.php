<?php
require 'vendor/autoload.php'; 

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$host = 'localhost:4306';
$dbname = 'college_db';
$username = 'root';
$password = '';

try {
    $pdo = new PDO("mysql:host=$host;dbname=$dbname", $username, $password);
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    $stmt = $pdo->query("SELECT prn, name, address FROM students");
    $data = $stmt->fetchAll(PDO::FETCH_ASSOC);

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    $sheet->setCellValue('A1', 'prn');
    $sheet->setCellValue('B1', 'name');
    $sheet->setCellValue('C1', 'address');

    $rowNumber = 2;
    foreach ($data as $row) {
        $sheet->setCellValue('A' . $rowNumber, $row['prn']);
        $sheet->setCellValue('B' . $rowNumber, $row['name']);
        $sheet->setCellValue('C' . $rowNumber, $row['address']);
        $rowNumber++;
    }

    $writer = new Xlsx($spreadsheet);
    $outputFileName = 'output_file.xlsx';
    $writer->save($outputFileName);

    echo "Data successfully exported to $outputFileName!";
} catch (Exception $e) {
    echo "Error: " . $e->getMessage();
}
?>

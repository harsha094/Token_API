<?php
include_once './config/database.php';
require "../vendor/autoload.php";

use \Firebase\JWT\JWT;
use Firebase\JWT\Key;

// Import PHPExcel classes
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


header("Access-Control-Allow-Origin: *");
header("Content-Type: application/json; charset=UTF-8");
header("Access-Control-Allow-Methods: POST");
header("Access-Control-Max-Age: 3600");
header("Access-Control-Allow-Headers: Content-Type, Access-Control-Allow-Headers, Authorization, X-Requested-With");

$secret_key = "bGS6lzFqvvSQ8ALbOxatm7/Vk7mLQyzqaS34Q4oR1ew=";
$jwt = null;
$databaseService = new DatabaseService();
$conn = $databaseService->getConnection();

$data = json_decode(file_get_contents("php://input"), true);
// var_dump($data);
// die();
$authHeader = $_SERVER['HTTP_AUTHORIZATION'];

$arr = explode(" ", $authHeader);
// print_r($arr);
// die();

$jwt = $arr[1];

if ($jwt) {
    try {
        // $decoded = JWT::decode($jwt, $secret_key, array('HS256'));
        $decoded = JWT::decode($jwt, new Key($secret_key, 'HS256'));
        // var_dump('$secret_key');
        // Access is granted. Add code of the operation here 
        // Convert JSON data to an array
        // $student_data = json_decode(file_get_contents("php://input"), true);
        // print_r('$data');
        // die();
        // Create a new Spreadsheet object
        $spreadsheet = new Spreadsheet();

        // Get the active worksheet
        $worksheet = $spreadsheet->getActiveSheet();

        // Set the column headers
        $worksheet->setCellValue('A1', 'Student Name');
        $worksheet->setCellValue('B1', 'Subject');
        $worksheet->setCellValue('C1', 'Obtained Marks');
        $worksheet->setCellValue('D1', 'Total Marks');

        // Set the data rows
        $row = 2;
        foreach ($data['student name'] as $student) {
            $worksheet->setCellValue('A' . $row, $student['name']);
            foreach ($student['subjects'] as $subject) {
                $worksheet->setCellValue('B' . $row, $subject['subject']);
                $worksheet->setCellValue('C' . $row, $subject['obtained_marks']);
                $worksheet->setCellValue('D' . $row, $subject['total_marks']);
                $row++;
                    var_dump($worksheet->getCell('A' . $row)->getValue());
                    var_dump($worksheet->getCell('B' . $row)->getValue());
            }
        }
        //     var_dump($worksheet->getCell('C' . $row)->getValue());
        //     var_dump($worksheet->getCell('D' . $row)->getValue());
        //     $row++;
        // }

        // Create a new Excel file in the server
        // $writer = new Xlsx($spreadsheet);
        $filename = '/home/harshraghuwanshi/www/phpproject/php-jwt-api/api/API/excel/students.xlsx';
        // $writer->save('/home/harshraghuwanshi/www/phpproject/php-jwt-api/api/API/excel/students.xlsx');
        if (!is_writable($filename)) {
            // echo "Error: $filename is not writable by the web server user";
            chmod($filename, 0644);
            
        }
        $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save('/home/harshraghuwanshi/www/phpproject/php-jwt-api/api/API/excel/students.xlsx');
        // chmod('/home/harshraghuwanshi/www/phpproject/php-jwt-api/api/API/excel/students.xlsx', 0644);
        // var_dump('File saved');
        // Download the file to the client
        // header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        // header('Content-Disposition: attachment;filename="' . $filename . '"');
        // header('Cache-Control: max-age=0');
        // $writer->save('php://output');
        // readfile($filename);

        echo json_encode(array(
            "message" => "Access granted"
        ));
    } catch (Exception $e) {
        http_response_code(401);
        // print_r($e->getTrace());
        // die();
        echo json_encode(array(
            "message" => "Access denied",
            "error" => $e->getMessage()
        ));
    }
} else {
    http_response_code(401);

    echo json_encode(array(
        "message" => "Access denied: token not provided"
    ));
}

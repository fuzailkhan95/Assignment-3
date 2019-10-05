<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$servername = "";
$username = "";
$password = "";
$database = "";

// Create connection

$conn = mysqli_connect($servername, $username, $password, $database);
                                         
if ($conn->connect_error) {
    die("Connection failed: " . $conn->connect_error);
}

//take variable convert the date format to date month year time (00.00.00.00) for examble; 21-AUG-19 01.00.00.00 AM
// $start_date = '14-FEB-19';
// $end_date = '15-FEB-19';

$start_date = $_POST["from"];
$end_date = $_POST["to"];

$date1 = new DateTime($start_date);
$start_date_formatted = $date1->format('Y-m-d');
$date2 = new DateTime($end_date);
$end_date_formatted = $date2->format('Y-m-d');

$spreadsheet = new Spreadsheet();
$spreadsheet->setActiveSheetIndex(0);
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'APP_UID');
$sheet->setCellValue('B1', 'APP_NUMBER');
$sheet->setCellValue('C1', 'APP_STATUS');
$sheet->setCellValue('D1', 'REQUEST_TYPE');
$sheet->setCellValue('E1', 'REQUEST_TYPE_LABEL');
$sheet->setCellValue('F1', 'T2APPROVAL_LABEL');
$sheet->setCellValue('G1', 'CASE_STATUS');
$sheet->setCellValue('H1', 'EMPLID');
$sheet->setCellValue('I1', 'FIRST_NAME');
$sheet->setCellValue('J1', 'LAST_NAME');
$sheet->setCellValue('K1', 'USER_ROLE_ID');
$sheet->setCellValue('L1', 'T2APPROVAL');
$sheet->setCellValue('M1', 'REQUEST_REASON_TEXT');

$styleArray = array(
    'font' => array(
    'bold' => true
    )
);
$sheet->getStyle('A1:M1')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(40);
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(15);
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(15);
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(25);
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(25);
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(25);
$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(15);
$spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('K')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('L')->setWidth(15);
$spreadsheet->getActiveSheet()->getColumnDimension('M')->setWidth(25);


$sql = "select A.* from wf_workflow.pmt_electronic_record_modification A
left join wf_workflow.app_delegation  B ON A.APP_UID = B.APP_UID AND B.TAS_UID = '503173828582c8c6b624f29016336731'
where DEL_FINISH_DATE > '$start_date_formatted 01:00:00' AND DEL_FINISH_DATE < '$end_date_formatted 01:00:00'";

$result=mysqli_query($conn,$sql);

$i = 2;
while ($row = mysqli_fetch_array($result)) {
    //Read Query Results (for loop)
    //While you read the rows, you will insert them into the Excel file
    $sheet->setCellValue('A'.$i, $row['APP_UID']);
    $sheet->setCellValue('B'.$i, $row['APP_NUMBER']);
    $sheet->setCellValue('C'.$i, $row['APP_STATUS']);
    $sheet->setCellValue('D'.$i, $row['REQUEST_TYPE']);
    $sheet->setCellValue('E'.$i, $row['REQUEST_TYPE_LABEL']);
    $sheet->setCellValue('F'.$i, $row['T2APPROVAL_LABEL']);
    $sheet->setCellValue('G'.$i, $row['CASE_STATUS']);
    $sheet->setCellValue('H'.$i, $row['EMPLID']);
    $sheet->setCellValue('I'.$i, $row['FIRST_NAME']);
    $sheet->setCellValue('J'.$i, $row['LAST_NAME']);
    $sheet->setCellValue('K'.$i, $row['USER_ROLE_ID']);
    $sheet->setCellValue('L'.$i, $row['T2APPROVAL']);
    $sheet->setCellValue('M'.$i, $row['REQUEST_REASON_TEXT']);
    $i++;
}

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="eRMR-report.xlsx"');
header('Cache-Control: max-age=0');

$objWriter = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet, 'Xlsx');
$objWriter->save('php://output');
exit;


//Close connection
$mysqli->close();


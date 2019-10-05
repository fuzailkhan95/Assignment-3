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
$sheet->setCellValue('D1', 'STUDENT_FIRSTNAME');
$sheet->setCellValue('E1', 'STUDENT_LASTNAME');
$sheet->setCellValue('F1', 'SEM_LABEL');
$sheet->setCellValue('G1', 'INSTRUCTOR');
$sheet->setCellValue('H1', 'INSTRUCTOR_EMPLID');
$sheet->setCellValue('I1', 'COURSE_LABEL');
$sheet->setCellValue('J1', 'GRADE_CHANGE_FROM');
$sheet->setCellValue('K1', 'GRADE_CHANGE_TO');
$sheet->setCellValue('L1', 'CAREER');
$sheet->setCellValue('M1', 'STUDENT_EMPLID');
$sheet->setCellValue('N1', 'CASE_STATUS');
$sheet->setCellValue('O1', 'T2_APPROVAL_LABEL');
$sheet->setCellValue('P1', 'T3_APPROVAL_LABEL');
$sheet->setCellValue('Q1', 'T4_APPROVAL_LABEL');
$sheet->setCellValue('R1', 'T5_REASONS_LABEL');
$sheet->setCellValue('S1', 'T5_REASONS');
$sheet->setCellValue('T1', 'COMMENTS_LOG');

$styleArray = array(
    'font' => array(
    'bold' => true
    )
);
$sheet->getStyle('A1:T1')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(40);
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(15);
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(25);
$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(15);
$spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('K')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('L')->setWidth(10);
$spreadsheet->getActiveSheet()->getColumnDimension('M')->setWidth(15);
$spreadsheet->getActiveSheet()->getColumnDimension('N')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('O')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('P')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('Q')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('R')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('S')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('T')->setWidth(20);


$sql = "select A.* from wf_workflow.pmt_egrade_to_registrar A
left join wf_workflow.app_delegation  B ON A.APP_UID = B.APP_UID AND B.TAS_UID = '94906616254450b79ae9bc5049067589'
where DEL_FINISH_DATE > '$start_date_formatted 01:00:00' AND DEL_FINISH_DATE < '$end_date_formatted 01:00:00'";

$result=mysqli_query($conn,$sql);

$i = 2;
while ($row = mysqli_fetch_array($result)) {
    //Read Query Results (for loop)
    //While you read the rows, you will insert them into the Excel file
    $sheet->setCellValue('A'.$i, $row['APP_UID']);
    $sheet->setCellValue('B'.$i, $row['APP_NUMBER']);
    $sheet->setCellValue('C'.$i, $row['APP_STATUS']);
    $sheet->setCellValue('D'.$i, $row['STUDENT_FIRSTNAME']);
    $sheet->setCellValue('E'.$i, $row['STUDENT_LASTNAME']);
    $sheet->setCellValue('F'.$i, $row['SEM_LABEL']);
    $sheet->setCellValue('G'.$i, $row['INSTRUCTOR']);
    $sheet->setCellValue('H'.$i, $row['INSTRUCTOR_EMPLID']);
    $sheet->setCellValue('I'.$i, $row['COURSE_LABEL']);
    $sheet->setCellValue('J'.$i, $row['GRADE_CHANGE_FROM']);
    $sheet->setCellValue('K'.$i, $row['GRADE_CHANGE_TO']);
    $sheet->setCellValue('L'.$i, $row['CAREER']);
    $sheet->setCellValue('M'.$i, $row['STUDENT_EMPLID']);
    $sheet->setCellValue('N'.$i, $row['CASE_STATUS']);
    $sheet->setCellValue('O'.$i, $row['T2_APPROVAL_LABEL']);
    $sheet->setCellValue('P'.$i, $row['T3_APPROVAL_LABEL']);
    $sheet->setCellValue('Q'.$i, $row['T4_APPROVAL_LABEL']);
    $sheet->setCellValue('R'.$i, $row['T5_REASONS_LABEL']);
    $sheet->setCellValue('S'.$i, $row['T5_REASONS']);
    $sheet->setCellValue('T'.$i, strip_tags($row['COMMENTS_LOG']));
    $i++;
}

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="eGrade-report.xlsx"');
header('Cache-Control: max-age=0');

$objWriter = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet, 'Xlsx');
$objWriter->save('php://output');
exit;


//Close connection
$mysqli->close();


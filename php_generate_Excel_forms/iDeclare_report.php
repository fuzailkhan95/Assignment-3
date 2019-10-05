<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

//Oracle - production
$dbServername = "";
$o_username="";
$o_password="";
$o_database="";


//Oracle Connection
$conn = oci_connect($o_username, $o_password, $o_database);
                                         
if (!$conn) {
    trigger_error("Could not connect to database", E_USER_ERROR);
}

//take variable convert the date format to date month year time (00.00.00.00) for examble; 21-AUG-19 01.00.00.00 AM
// $start_date = '14-FEB-19';
// $end_date = '15-FEB-19';

$start_date = $_POST["from"];
$end_date = $_POST["to"];

$date1 = new DateTime($start_date);
$start_date_formatted = $date1->format('d-M-y');
$date2 = new DateTime($end_date);
$end_date_formatted = $date2->format('d-M-y');

$spreadsheet = new Spreadsheet();
$spreadsheet->setActiveSheetIndex(0);
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'ID');
$sheet->setCellValue('B1', 'APP_UID');
$sheet->setCellValue('C1', 'APP_NUMBER');
$sheet->setCellValue('D1', 'EMPLID');
$sheet->setCellValue('E1', 'FNAME');
$sheet->setCellValue('F1', 'LNAME');
$sheet->setCellValue('G1', 'MNAME');
$sheet->setCellValue('H1', 'EMAIL');
$sheet->setCellValue('I1', 'PHONE');
$sheet->setCellValue('J1', 'PHONE_ORIGINAL');
$sheet->setCellValue('K1', 'CAREER');
$sheet->setCellValue('L1', 'MHC');
$sheet->setCellValue('M1', 'CHANGE_TYPE');
$sheet->setCellValue('N1', 'DEPARTMENT');
$sheet->setCellValue('O1', 'DEPARTMENT_LABEL');
$sheet->setCellValue('P1', 'MAJOR_PLAN');
$sheet->setCellValue('Q1', 'MAJOR_PLAN_LABEL');
$sheet->setCellValue('R1', 'MAJOR_SUBPLAN');
$sheet->setCellValue('S1', 'MAJOR_SUBPLAN_LABEL');
$sheet->setCellValue('T1', 'STUDENT_SIGNATURE');
$sheet->setCellValue('U1', 'SIGNATURE_DATE');
$sheet->setCellValue('V1', 'STATUS');
$sheet->setCellValue('W1', 'SUBMITTED');
$sheet->setCellValue('X1', 'ADVISOR');
$sheet->setCellValue('Y1', 'ADVISOR_DT');
$sheet->setCellValue('Z1', 'ADVISOR_COMMENT');
$sheet->setCellValue('AA1', 'REGISTRAR');
$sheet->setCellValue('AB1', 'REGISTRAR_DT');
$sheet->setCellValue('AC1', 'REGISTRAR_COMMENT');
$sheet->setCellValue('AD1', 'TPE_ADVISOR');
$sheet->setCellValue('AE1', 'TPE_ADVISOR_DT');
$sheet->setCellValue('AF1', 'TPE_ADVISOR_COMMENT');

$styleArray = array(
    'font' => array(
    'bold' => true
    )
);
$sheet->getStyle('A1:AF1')->applyFromArray($styleArray);

$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(40);
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(15);
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(10);
$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(40);
$spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('K')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('L')->setWidth(6);
$spreadsheet->getActiveSheet()->getColumnDimension('M')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('N')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('O')->setWidth(30);
$spreadsheet->getActiveSheet()->getColumnDimension('P')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('Q')->setWidth(30);
$spreadsheet->getActiveSheet()->getColumnDimension('R')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('S')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('T')->setWidth(30);
$spreadsheet->getActiveSheet()->getColumnDimension('U')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('V')->setWidth(30);
$spreadsheet->getActiveSheet()->getColumnDimension('W')->setWidth(35);
$spreadsheet->getActiveSheet()->getColumnDimension('X')->setWidth(30);
$spreadsheet->getActiveSheet()->getColumnDimension('Y')->setWidth(35);
$spreadsheet->getActiveSheet()->getColumnDimension('Z')->setWidth(30);
$spreadsheet->getActiveSheet()->getColumnDimension('AA')->setWidth(30);
$spreadsheet->getActiveSheet()->getColumnDimension('AB')->setWidth(30);
$spreadsheet->getActiveSheet()->getColumnDimension('AC')->setWidth(30);
$spreadsheet->getActiveSheet()->getColumnDimension('AD')->setWidth(30);
$spreadsheet->getActiveSheet()->getColumnDimension('AE')->setWidth(30);
$spreadsheet->getActiveSheet()->getColumnDimension('AF')->setWidth(30);


if ($_POST['status'] == 'status1') {
    $sql1 = "select * from ideclare where registrar_dt >= '$start_date_formatted 01.00.00 AM' and registrar_dt <= '$end_date_formatted 11.59.00 PM'order by registrar_dt desc";
    $result_oci1 = oci_parse($conn, $sql1);
    oci_execute($result_oci1);

    $i = 2;
    while (($row = oci_fetch_array($result_oci1, OCI_ASSOC+OCI_RETURN_NULLS)) != false) {
        
        //Read Query Results (for loop)
        //While you read the rows, you will insert them into the Excel file
        $sheet->setCellValue('A'.$i, $row['ID']);
        $sheet->setCellValue('B'.$i, $row['APP_UID']);
        $sheet->setCellValue('C'.$i, $row['APP_NUMBER']);
        $sheet->setCellValue('D'.$i, $row['EMPLID']);
        $sheet->setCellValue('E'.$i, $row['FNAME']);
        $sheet->setCellValue('F'.$i, $row['LNAME']);
        $sheet->setCellValue('G'.$i, $row['MNAME']);
        $sheet->setCellValue('H'.$i, $row['EMAIL']);
        $sheet->setCellValue('I'.$i, $row['PHONE']);
        $sheet->setCellValue('J'.$i, $row['PHONE_ORIGINAL']);
        $sheet->setCellValue('K'.$i, $row['CAREER']);
        $sheet->setCellValue('L'.$i, $row['MHC']);
        $sheet->setCellValue('M'.$i, $row['CHANGE_TYPE']);
        $sheet->setCellValue('N'.$i, $row['DEPARTMENT']);
        $sheet->setCellValue('O'.$i, $row['DEPARTMENT_LABEL']);
        $sheet->setCellValue('P'.$i, $row['MAJOR_PLAN']);
        $sheet->setCellValue('Q'.$i, $row['MAJOR_PLAN_LABEL']);
        $sheet->setCellValue('R'.$i, $row['MAJOR_SUBPLAN']);
        $sheet->setCellValue('S'.$i, $row['MAJOR_SUBPLAN_LABEL']);
        $sheet->setCellValue('T'.$i, $row['STUDENT_SIGNATURE']);
        $sheet->setCellValue('U'.$i, $row['SIGNATURE_DATE']);
        $sheet->setCellValue('V'.$i, $row['STATUS']);
        $sheet->setCellValue('W'.$i, $row['SUBMITTED']);
        $sheet->setCellValue('X'.$i, $row['ADVISOR']);
        $sheet->setCellValue('Y'.$i, $row['ADVISOR_DT']);
        $sheet->setCellValue('Z'.$i, $row['ADVISOR_COMMENT']);
        $sheet->setCellValue('AA'.$i, $row['REGISTRAR']);
        $sheet->setCellValue('AB'.$i, $row['REGISTRAR_DT']);
        $sheet->setCellValue('AC'.$i, $row['REGISTRAR_COMMENT']);
        $sheet->setCellValue('AD'.$i, $row['TPE_ADVISOR']);
        $sheet->setCellValue('AE'.$i, $row['TPE_ADVISOR_DT']);
        $sheet->setCellValue('AF'.$i, $row['TPE_ADVISOR_COMMENT']);
        $i++;
    }

    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="ideclare-report-registrar.xlsx"');
    header('Cache-Control: max-age=0');

    $objWriter = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet, 'Xlsx');
    $objWriter->save('php://output');
    exit;

}
elseif ( $_POST['status'] == 'status2') {
    $sql2 = "select * from ideclare where (status in ('PENDING ADVISOR APPROVAL','PENDING ENGLISH COMP. OFFICE APPROVAL') or status is null) and submitted >= '$start_date_formatted 01.00.00 AM' AND submitted <=  '$end_date_formatted 11.59.00 PM' order by id";
    $result_oci2 = oci_parse($conn, $sql2);
    oci_execute($result_oci2);

    $i = 2;
    while (($row = oci_fetch_array($result_oci2, OCI_ASSOC+OCI_RETURN_NULLS)) != false) {
        
        //Read Query Results (for loop)
        //While you read the rows, you will insert them into the Excel file
        $sheet->setCellValue('A'.$i, $row['ID']);
        $sheet->setCellValue('B'.$i, $row['APP_UID']);
        $sheet->setCellValue('C'.$i, $row['APP_NUMBER']);
        $sheet->setCellValue('D'.$i, $row['EMPLID']);
        $sheet->setCellValue('E'.$i, $row['FNAME']);
        $sheet->setCellValue('F'.$i, $row['LNAME']);
        $sheet->setCellValue('G'.$i, $row['MNAME']);
        $sheet->setCellValue('H'.$i, $row['EMAIL']);
        $sheet->setCellValue('I'.$i, $row['PHONE']);
        $sheet->setCellValue('J'.$i, $row['PHONE_ORIGINAL']);
        $sheet->setCellValue('K'.$i, $row['CAREER']);
        $sheet->setCellValue('L'.$i, $row['MHC']);
        $sheet->setCellValue('M'.$i, $row['CHANGE_TYPE']);
        $sheet->setCellValue('N'.$i, $row['DEPARTMENT']);
        $sheet->setCellValue('O'.$i, $row['DEPARTMENT_LABEL']);
        $sheet->setCellValue('P'.$i, $row['MAJOR_PLAN']);
        $sheet->setCellValue('Q'.$i, $row['MAJOR_PLAN_LABEL']);
        $sheet->setCellValue('R'.$i, $row['MAJOR_SUBPLAN']);
        $sheet->setCellValue('S'.$i, $row['MAJOR_SUBPLAN_LABEL']);
        $sheet->setCellValue('T'.$i, $row['STUDENT_SIGNATURE']);
        $sheet->setCellValue('U'.$i, $row['SIGNATURE_DATE']);
        $sheet->setCellValue('V'.$i, $row['STATUS']);
        $sheet->setCellValue('W'.$i, $row['SUBMITTED']);
        $sheet->setCellValue('X'.$i, $row['ADVISOR']);
        $sheet->setCellValue('Y'.$i, $row['ADVISOR_DT']);
        $sheet->setCellValue('Z'.$i, $row['ADVISOR_COMMENT']);
        $sheet->setCellValue('AA'.$i, $row['REGISTRAR']);
        $sheet->setCellValue('AB'.$i, $row['REGISTRAR_DT']);
        $sheet->setCellValue('AC'.$i, $row['REGISTRAR_COMMENT']);
        $sheet->setCellValue('AD'.$i, $row['TPE_ADVISOR']);
        $sheet->setCellValue('AE'.$i, $row['TPE_ADVISOR_DT']);
        $sheet->setCellValue('AF'.$i, $row['TPE_ADVISOR_COMMENT']);
        $i++;
    }
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="ideclare-pending-cases.xlsx"');
    header('Cache-Control: max-age=0');

    $objWriter = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet, 'Xlsx');
    $objWriter->save('php://output');
    exit;
}

//Close connection
oci_close($conn);


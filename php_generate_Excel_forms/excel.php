<?php

if ($_POST['report_type'] == '1') {
     include "iDeclare_report.php";
}elseif($_POST['report_type'] == '2'){
     include "eGrade_report.php";
}elseif($_POST['report_type'] == '3'){
     include "eRMR_report.php";
}

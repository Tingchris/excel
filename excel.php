<?php
    include_once "header.php";
    require 'vendor/autoload.php';
    require_once '/opt/lampp/htdocs/site/mysql.php';

    if($_SERVER['REQUEST_METHOD']=="POST"){
        $account=$_POST['account'];
    }

    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    use PhpOffice\PhpSpreadsheet\IOFactory;
    use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

    $patient = "SELECT * FROM co WHERE account = '$account'";
    $result=mysqli_query($conn,$patient);

    $action1="SELECT * FROM action WHERE account = '$account' AND degree='初階' AND parts='上肢'";
    $query1=mysqli_query($conn,$action1);

    $action2="SELECT * FROM action WHERE account = '$account' AND degree='初階' AND parts='下肢'";
    $query2=mysqli_query($conn,$action2);

    $action3="SELECT * FROM action WHERE account = '$account' AND degree='進階' AND parts='上肢'";
    $query3=mysqli_query($conn,$action3);

    $action4="SELECT * FROM action WHERE account = '$account' AND degree='初階' AND parts='吞嚥'";
    $query4=mysqli_query($conn,$action4);

    $action5="SELECT * FROM action WHERE account = '$account' AND degree='進階' AND parts='下肢'";
    $query5=mysqli_query($conn,$action5);

    $spreadsheet = new Spreadsheet();
    $spreadsheet->setActiveSheetIndex(0);
    $spreadsheet->getActiveSheet()->getStyle('A:S')->getAlignment()->setHorizontal(
        \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER
    );

    $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth('12');
    $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth('9');
    $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth('9');
    $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth('12');
    $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth('9');
    $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth('19');
    $spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth('9');
    $spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth('12');
    $spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth('13');
    $spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth('18');
    $spreadsheet->getActiveSheet()->getColumnDimension('K')->setWidth('12');
    $spreadsheet->getActiveSheet()->getColumnDimension('L')->setWidth('10');
    
    $spreadsheet->getActiveSheet()->setTitle('病人資料');
    $spreadsheet->getActiveSheet()->setCellValue('A1','個案編號');
    $spreadsheet->getActiveSheet()->setCellValue('B1','姓名');
    $spreadsheet->getActiveSheet()->setCellValue('C1','性別');
    $spreadsheet->getActiveSheet()->setCellValue('D1','生日');
    $spreadsheet->getActiveSheet()->setCellValue('E1','年齡');
    $spreadsheet->getActiveSheet()->setCellValue('F1','診斷');
    $spreadsheet->getActiveSheet()->setCellValue('G1','患側');
    $spreadsheet->getActiveSheet()->setCellValue('H1','聯絡電話');
    $spreadsheet->getActiveSheet()->setCellValue('I1','緊急聯絡人');
    $spreadsheet->getActiveSheet()->setCellValue('J1','緊急聯絡人電話');
    $spreadsheet->getActiveSheet()->setCellValue('K1','加入日期');
    $spreadsheet->getActiveSheet()->setCellValue('L1','累積金幣');
    foreach($result as $row){
        $spreadsheet->getActiveSheet()->setCellValue('A2',$row['account']);
        $spreadsheet->getActiveSheet()->setCellValue('B2',$row['name']);
        $spreadsheet->getActiveSheet()->setCellValue('C2',$row['gender']);
        $spreadsheet->getActiveSheet()->setCellValue('D2',$row['birthday']);
        $spreadsheet->getActiveSheet()->setCellValue('E2',$row['age']);
        $spreadsheet->getActiveSheet()->setCellValue('F2',$row['diagnosis']);
        $spreadsheet->getActiveSheet()->setCellValue('G2',$row['affectedside']);
        $spreadsheet->getActiveSheet()->setCellValue('H2',$row['phone']);
        $spreadsheet->getActiveSheet()->setCellValue('I2',$row['urgenname']);
        $spreadsheet->getActiveSheet()->setCellValue('J2',$row['urgenphone']);
        $spreadsheet->getActiveSheet()->setCellValue('K2',$row['joindate']);
        $spreadsheet->getActiveSheet()->setCellValue('L2',$row['coin']);
    }
    $spreadsheet->createSheet();
    $spreadsheet->setActiveSheetIndex(1);
    $spreadsheet->getActiveSheet()->getStyle('A:S')->getAlignment()->setHorizontal(
        \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER
    );

    $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth('15');
    $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth('7');
    $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth('12');
    $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth('5');
    $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth('15');
    $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth('7');
    $spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth('12');
    $spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth('5');
    $spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth('15');
    $spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth('7');
    $spreadsheet->getActiveSheet()->getColumnDimension('K')->setWidth('12');
    $spreadsheet->getActiveSheet()->getColumnDimension('L')->setWidth('5');
    $spreadsheet->getActiveSheet()->getColumnDimension('M')->setWidth('15');
    $spreadsheet->getActiveSheet()->getColumnDimension('N')->setWidth('7');
    $spreadsheet->getActiveSheet()->getColumnDimension('O')->setWidth('12');
    $spreadsheet->getActiveSheet()->getColumnDimension('P')->setWidth('5');
    $spreadsheet->getActiveSheet()->getColumnDimension('Q')->setWidth('15');
    $spreadsheet->getActiveSheet()->getColumnDimension('R')->setWidth('7');
    $spreadsheet->getActiveSheet()->getColumnDimension('S')->setWidth('12');

    $spreadsheet->getActiveSheet()->setTitle('復健記錄');
    $spreadsheet->getActiveSheet()->mergeCells('A1:G1');
    $spreadsheet->getActiveSheet()->mergeCells('A2:C2');
    $spreadsheet->getActiveSheet()->mergeCells('E2:G2');
    $spreadsheet->getActiveSheet()->setCellValue('A1','上肢訓練');
    $spreadsheet->getActiveSheet()->setCellValue('A2','初階動作');
    $spreadsheet->getActiveSheet()->setCellValue('E2','進階動作');
    $spreadsheet->getActiveSheet()->setCellValue('A3','時間');
    $spreadsheet->getActiveSheet()->setCellValue('B3','次數');
    $spreadsheet->getActiveSheet()->setCellValue('C3','動作');

    $i=4;
    foreach($query1 as $row){
        $spreadsheet->getActiveSheet()->setCellValue('A'.$i,$row['time']);
        $spreadsheet->getActiveSheet()->setCellValue('B'.$i,$row['times']);
        $spreadsheet->getActiveSheet()->setCellValue('C'.$i,$row['action']);
        $i=$i+1;
    }
    $spreadsheet->getActiveSheet()->setCellValue('E3','時間');
    $spreadsheet->getActiveSheet()->setCellValue('F3','次數');
    $spreadsheet->getActiveSheet()->setCellValue('G3','動作');
    $i=4;
    foreach($query3 as $row){
        $spreadsheet->getActiveSheet()->setCellValue('E'.$i,$row['time']);
        $spreadsheet->getActiveSheet()->setCellValue('F'.$i,$row['times']);
        $spreadsheet->getActiveSheet()->setCellValue('G'.$i,$row['action']);
        $i=$i+1;
    }
    
    $spreadsheet->getActiveSheet()->mergeCells('I1:O1');
    $spreadsheet->getActiveSheet()->mergeCells('I2:K2');
    $spreadsheet->getActiveSheet()->mergeCells('M2:O2');
    $spreadsheet->getActiveSheet()->setCellValue('I1','下肢訓練');
    $spreadsheet->getActiveSheet()->setCellValue('I2','初階動作');
    $spreadsheet->getActiveSheet()->setCellValue('M2','進階動作');
    $spreadsheet->getActiveSheet()->setCellValue('I3','時間');
    $spreadsheet->getActiveSheet()->setCellValue('J3','次數');
    $spreadsheet->getActiveSheet()->setCellValue('K3','動作');
    $i=4;
    foreach($query2 as $row){
        $spreadsheet->getActiveSheet()->setCellValue('I'.$i,$row['time']);
        $spreadsheet->getActiveSheet()->setCellValue('J'.$i,$row['times']);
        $spreadsheet->getActiveSheet()->setCellValue('K'.$i,$row['action']);
        $i=$i+1;
    }
    $spreadsheet->getActiveSheet()->setCellValue('M3','時間');
    $spreadsheet->getActiveSheet()->setCellValue('N3','次數');
    $spreadsheet->getActiveSheet()->setCellValue('O3','動作');
    $i=4;
    foreach($query5 as $row){
        $spreadsheet->getActiveSheet()->setCellValue('M'.$i,$row['time']);
        $spreadsheet->getActiveSheet()->setCellValue('N'.$i,$row['times']);
        $spreadsheet->getActiveSheet()->setCellValue('O'.$i,$row['action']);
        $i=$i+1;
    }

    $spreadsheet->getActiveSheet()->mergeCells('Q1:S1');
    $spreadsheet->getActiveSheet()->setCellValue('Q1','吞嚥訓練');
    $spreadsheet->getActiveSheet()->setCellValue('Q3','時間');
    $spreadsheet->getActiveSheet()->setCellValue('R3','次數');
    $spreadsheet->getActiveSheet()->setCellValue('S3','動作');
    $i=4;
    foreach($query4 as $row){
        $spreadsheet->getActiveSheet()->setCellValue('Q'.$i,$row['time']);
        $spreadsheet->getActiveSheet()->setCellValue('R'.$i,$row['times']);
        $spreadsheet->getActiveSheet()->setCellValue('S'.$i,$row['action']);
        $i=$i+1;
    }
    $spreadsheet->setActiveSheetIndex(0);
    
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="TEST.xlsx"');
    header('Cache-Control: max-age=0');

    $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
    $writer->save('php://output');

    $spreadsheet->disconnectWorksheets();
    unset($spreadsheet);

    exit;
?>

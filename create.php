<?php
	// die(var_dump($_FILES));
	if (($_FILES["json_file"]["error"]==0)&&(substr($_FILES["json_file"]["name"], strlen($_FILES["json_file"]["name"])-5)==".json")) {
		$time=getdate();
		move_uploaded_file($_FILES["json_file"]["tmp_name"], "input/".$_FILES["json_file"]["name"]);
	}
	$content=json_decode(file_get_contents("input/".$_FILES["json_file"]["name"]),true);

	include 'PHPExcel.php';
	// include 'PHPExcel/Writer/Excel2007.php';
	//或者include 'PHPExcel/Writer/Excel5.php'; 用于输出.xls的
	// 创建一个excel
	$objPHPExcel = new PHPExcel();

	// 设置当前的sheet
	$objPHPExcel->setActiveSheetIndex(0);

	// 设置sheet的name
	$objPHPExcel->getActiveSheet()->setTitle('uyan comments');

	// 设置单元格的值
	$objPHPExcel->getActiveSheet()->setCellValue('A1', '时间');
	$objPHPExcel->getActiveSheet()->setCellValue('B1', '发言人');
	$objPHPExcel->getActiveSheet()->setCellValue('C1', '来自');
	$objPHPExcel->getActiveSheet()->setCellValue('D1', '内容');
	$objPHPExcel->getActiveSheet()->mergeCells('A1:A2');
	$objPHPExcel->getActiveSheet()->mergeCells('B1:B2');
	$objPHPExcel->getActiveSheet()->mergeCells('C1:C2');
	$objPHPExcel->getActiveSheet()->mergeCells('D1:D2');

	$objPHPExcel->getActiveSheet()->setCellValue('E1', '回复');
	$objPHPExcel->getActiveSheet()->mergeCells('E1:H1');
	$objPHPExcel->getActiveSheet()->setCellValue('E2', '时间');
	$objPHPExcel->getActiveSheet()->setCellValue('F2', '发言人');
	$objPHPExcel->getActiveSheet()->setCellValue('G2', '来自');
	$objPHPExcel->getActiveSheet()->setCellValue('H2', '内容');
	
	$i=3;
	foreach ($content as $key => $value) {
		$objPHPExcel->getActiveSheet()->setCellValue('A'.$i,$value['time']);
		$objPHPExcel->getActiveSheet()->setCellValue('B'.$i,$value['uname']);
		$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,$value['ulink']);
		$objPHPExcel->getActiveSheet()->setCellValue('D'.$i,$value['content']);
		$i++;
		if (isset($value['child'])) {
			foreach ($value['child'] as $k => $v) {
				$objPHPExcel->getActiveSheet()->setCellValue('E'.$i,$v['time']);
				$objPHPExcel->getActiveSheet()->setCellValue('F'.$i,$v['uname']);
				$objPHPExcel->getActiveSheet()->setCellValue('G'.$i,$v['ulink']);
				$objPHPExcel->getActiveSheet()->setCellValue('H'.$i,$v['content']);
				$i++;
			}
		}
	}
	echo $i-2;

	// 保存excel—2007格式
	$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
	//或者$objWriter = new PHPExcel_Writer_Excel5($objPHPExcel); 非2007格式
	$objWriter->save("output/".$_POST["excel_name"].".xlsx");
?>
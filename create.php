<?php
	// die(var_dump($_FILES));
	if (($_FILES["json_file"]["error"]==0)&&(substr($_FILES["json_file"]["name"], strlen($_FILES["json_file"]["name"])-5)==".json")) {
		$time=getdate();
		move_uploaded_file($_FILES["json_file"]["tmp_name"], "input/".$_FILES["json_file"]["name"]);
	}
	$content=json_decode(file_get_contents("input/".$_FILES["json_file"]["name"]),true);

	include 'PHPExcel_1.7.9_doc/Classes/PHPExcel.php';
	include 'PHPExcel_1.7.9_doc/Classes/PHPExcel/Writer/Excel2007.php';
	// 或者
	// include 'PHPExcel_1.7.9_doc/Classes/PHPExcel/Writer/Excel5.php'; //用于输出.xls的
	// 创建一个excel
	$objPHPExcel = new PHPExcel();

	// 设置当前的sheet
	$objPHPExcel->setActiveSheetIndex(0);

	// 设置sheet的name
	$objPHPExcel->getActiveSheet()->setTitle('uyan comments');

	// 设置单元格的值
	$objPHPExcel->getActiveSheet()->setCellValue('A1', '序号');
	$objPHPExcel->getActiveSheet()->setCellValue('B1', '时间');
	$objPHPExcel->getActiveSheet()->setCellValue('C1', '发言人');
	$objPHPExcel->getActiveSheet()->setCellValue('D1', '来自');
	$objPHPExcel->getActiveSheet()->setCellValue('E1', '内容');
	// $objPHPExcel->getActiveSheet()->mergeCells('A1:A2');
	// $objPHPExcel->getActiveSheet()->mergeCells('B1:B2');
	// $objPHPExcel->getActiveSheet()->mergeCells('C1:C2');
	// $objPHPExcel->getActiveSheet()->mergeCells('D1:D2');

	$objPHPExcel->getActiveSheet()->setCellValue('F1', '回复');
	$objPHPExcel->getActiveSheet()->setCellValue('G1', '类别');
	$objPHPExcel->getActiveSheet()->setCellValue('H1', '来自求是潮');
	// $objPHPExcel->getActiveSheet()->mergeCells('E1:H1');
	// $objPHPExcel->getActiveSheet()->setCellValue('E2', '时间');
	// $objPHPExcel->getActiveSheet()->setCellValue('F2', '发言人');
	// $objPHPExcel->getActiveSheet()->setCellValue('G2', '来自');
	// $objPHPExcel->getActiveSheet()->setCellValue('H2', '内容');
	
	$i=1;
	$index=0;
	foreach ($content as $key => $value) {
		$index++;		
		$i++;
		$objPHPExcel->getActiveSheet()->setCellValue('A'.$i,$index);
		$objPHPExcel->getActiveSheet()->setCellValue('B'.$i,$value['time']);
		$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,$value['uname']);
		// $objPHPExcel->getActiveSheet()->setCellValue('D'.$i,$value['ulink']);
		if (!$value['ulink']) {
			$objPHPExcel->getActiveSheet()->setCellValue('D'.$i,'游客');
		} elseif (stripos($value['ulink'], 'qq')) {
			$objPHPExcel->getActiveSheet()->setCellValue('D'.$i,'空间');
		} elseif (stripos($value['ulink'], 'weibo')) {
			$objPHPExcel->getActiveSheet()->setCellValue('D'.$i,'微博');
		} elseif (stripos($value['ulink'], 'renren')) {
			$objPHPExcel->getActiveSheet()->setCellValue('D'.$i,'人人');
		} else {
			$objPHPExcel->getActiveSheet()->setCellValue('D'.$i,'其它');
		}
		$objPHPExcel->getActiveSheet()->setCellValue('E'.$i,$value['content']);
		$objPHPExcel->getActiveSheet()->setCellValue('F'.$i,0);
		$fatherindex=$index;
		if (isset($value['child'])) {
			foreach ($value['child'] as $k => $v) {
				$index++;
				$i++;
				$objPHPExcel->getActiveSheet()->setCellValue('A'.$i,$index);
				$objPHPExcel->getActiveSheet()->setCellValue('B'.$i,$v['time']);
				$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,$v['uname']);
				if (!$v['ulink']) {
					$objPHPExcel->getActiveSheet()->setCellValue('D'.$i,'游客');
				} elseif (stripos($v['ulink'], 'qq')) {
					$objPHPExcel->getActiveSheet()->setCellValue('D'.$i,'空间');
				} elseif (stripos($v['ulink'], 'weibo')) {
					$objPHPExcel->getActiveSheet()->setCellValue('D'.$i,'微博');
				} elseif (stripos($v['ulink'], 'renren')) {
					$objPHPExcel->getActiveSheet()->setCellValue('D'.$i,'人人');
				} else {
					$objPHPExcel->getActiveSheet()->setCellValue('D'.$i,'其它');
				}
				$objPHPExcel->getActiveSheet()->setCellValue('E'.$i,$v['content']);
				$objPHPExcel->getActiveSheet()->setCellValue('F'.$i,$fatherindex);
			}
		}
	}
	echo $i-1;

	// 保存excel—2007格式
	$objWriter = new PHPExcel_Writer_Excel5($objPHPExcel);
	//或者$objWriter = new PHPExcel_Writer_Excel5($objPHPExcel); 非2007格式
	$objWriter->save("output/".$_POST["excel_name"].".xls");
	header("Location: "."output/".$_POST["excel_name"].".xls")
	// $file=$_POST["excel_name"].'.xlsx';
	// header("Content-Type:application/force-download");
	// header('Content-Disposition:attachment;filename='.basename($file));
	// $objWriter->save('php://output');


	// $objWriter = new PHPExcel_Writer_Excel5($objPHPExcel);
	// header("Pragma: public");
	// header("Expires: 0");
	// header("Cache-Control:must-revalidate, post-check=0, pre-check=0");
	// header("Content-Type:application/force-download");
	// header("Content-Type:application/vnd.ms-excel");
	// header("Content-Type:application/octet-stream");
	// header("Content-Type:application/download");;
	// header('Content-Disposition:attachment;filename="'.$_POST["excel_name"].'.xls"');
	// header("Content-Transfer-Encoding:binary");
	// $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
	// $objWriter->save('php://output');
?>
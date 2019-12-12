<?php
require_once 'C:/xampp/htdocs/phpexcel/PHPExcel-1.8/PHPExcel-1.8/Classes/PHPExcel.php';

//create PHPExcel object


$conn = mysqli_connect("localhost","guest","guest123","cie");

$excel = new PHPExcel();

$excel->setActiveSheetIndex(0);

$query = "SELECT * FROM result";


$result = mysqli_query($conn,$query);



$i=4;



while($data = mysqli_fetch_object($result)){
	
	$excel->getActiveSheet()
		->setCellValue('A'.$i, $i-3)
		->setCellValue('B'.$i, $data->usn)
		->setCellValue('C'.$i, $data->name)
		->setCellValue('D'.$i, $data->res);
	
	$i++;
	
}	



$excel->getActiveSheet()->getColumnDimension('A')->setWidth(10);
$excel->getActiveSheet()->getColumnDimension('B')->setWidth(30);
$excel->getActiveSheet()->getColumnDimension('C')->setWidth(30);
$excel->getActiveSheet()->getColumnDimension('D')->setWidth(10);


//Header

$excel->getActiveSheet()
	->setCellValue('A1','CIE Eligibility')
	->setCellValue('A3','S. No.')
	->setCellValue('B3','USN')
	->setCellValue('C3','Name')
	->setCellValue('D3','Eligible');
	
	
	
$excel->getActiveSheet()->mergeCells('A1:D2');

$excel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal('center');

$excel->getActiveSheet()->getStyle('A1')->applyFromArray(
	array(
		'font'=>array(
			'size'=>24,
		)
	)
);


$excel->getActiveSheet()->getStyle('A3:D3')->applyFromArray(
	array(
		'font'=>array(
			'bold'=>true
		),
		'borders' =>array(
			'allborders' =>array(
				'style'=>PHPExcel_Style_Border::BORDER_THIN
			)
		)
	)
);


$excel->getActiveSheet()->getStyle('A4:D'.($i-1))->applyFromArray(
	array(
		'borders' =>array(
			'outline'=>array(
				'style'=>PHPExcel_Style_Border::BORDER_THIN
			),
			'vertical'=>array(
				'style'=>PHPExcel_Style_Border::BORDER_THIN
			)
		)
	)
);


ob_end_clean();

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="CIE_Evaluation.xlsx"');




header('Cache-Control: max-age=0');

//write the result to a file
//for excel 2007 format
$file = PHPExcel_IOFactory::createWriter($excel,'Excel2007');

//for excel 2003 format
//$file = PHPExcel_IOFactory::createWriter($excel,'Excel5');

//output to php output instead of filename
$file->save('php://output');



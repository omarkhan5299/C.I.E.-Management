<html>
<body>

<?php

require_once 'C:/xampp/htdocs/phpexcel/PHPExcel-1.8/PHPExcel-1.8/Classes/PHPExcel.php';



//function


function alert($msg){
	
	echo "<script type='text/javascript'>alert('$msg');</script>";
}

function error(){
	global $marks,$attendance,$name,$usn,$i;
	if(!is_double($marks) || $marks<0 || $marks>100){
		alert("Marks should be between 0 and 100.");
		alert("Error in row $i. Reupload from row $i onwards.");
		throw new Exception(0);
	}
	
	if(!is_double($attendance) || $attendance<0 || $attendance>100){
		alert("Attendance should be between 0 and 100.");
		alert("Error in row $i. Reupload from row $i onwards.");
		throw new Exception(0);
	}
	
	if(!is_string($name) ||  ltrim($name)==''){
		alert("Invalid Name.");
		alert("Error in row $i. Reupload from row $i onwards.");
		throw new Exception(0);
	}
	
	
	
}



if(empty($_FILES)){
	echo "
	<form method='post' enctype='multipart/form-data' action='uploadToDatabase.php'>
		<input type='file' name='excel'>
		<br>
		<button type='submit'>Submit</button>
	</form>
	";
}else{
	$excel = PHPExcel_IOFactory::load($_FILES['excel']['tmp_name']);

$excel->setActiveSheetIndex(0);


$i=2;


while($excel->getActiveSheet()->getCell('B'.$i)->getValue()!=""){
	
	$usn=$excel->getActiveSheet()->getCell('B'.$i)->getValue();
	$name=$excel->getActiveSheet()->getCell('C'.$i)->getValue();
	$marks=$excel->getActiveSheet()->getCell('D'.$i)->getValue();
	$attendance=$excel->getActiveSheet()->getCell('E'.$i)->getValue();
	
	
	try{
		error();
	}
	catch(Exception $e){
		break;
	}
	
	
	
	$conn = mysqli_connect("localhost","guest","guest123","cie");
	

	$query = "INSERT INTO eligibility (SNo, usn, name, marks, attendance) VALUES (NULL,'$usn','$name','$marks','$attendance')";
	
	$duplicate="SELECT * FROM eligibility WHERE usn='$usn'";
	
	
	
	$d=mysqli_query($conn,$duplicate) or die(mysqli_error($conn));
	if(mysqli_num_rows($d)){
		alert("Duplicate USN.");
		alert("Error in row $i. Reupload from row $i onwards.");
		break;
	}
	
	mysqli_query($conn,$query) or die(mysqli_error($conn));

	if($marks>85 && $attendance>85)
		$r='Yes';
	else
		$r='No';


	$query = "INSERT INTO result (SNo, usn, name, res) VALUES (NULL,'$usn','$name','$r')";
	

	mysqli_query($conn,$query) or die(mysqli_error($conn));
	
	
	$i++;
		
		
}

	echo "<script type='text/javascript'>
	location.replace('http://localhost:8080/My%20Programs/Web%20Programming%20Project/redirect/homePage.html')
	</script>";

}


?>



</body>
</html>
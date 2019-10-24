<?php
require('excel.php');
require('db.php');
$error="";
$success="";
$success1="";
$ind=0;
if(isset($_FILES["fileToUpload"])){
$target_file = basename($_FILES["fileToUpload"]["name"]);
$allowed =  array('xls','xlsx');
$filename = $_FILES['fileToUpload']['name'];
$ext = pathinfo($filename, PATHINFO_EXTENSION);
if(in_array($ext,$allowed) ) {
  move_uploaded_file($_FILES["fileToUpload"]["tmp_name"], $target_file);
$reader = new Reader();

try {
  $reader->open($target_file);
} catch (\Exception $e) {
  $error=$e;
}
$sheets = $reader->getSheets();
$heade=Array() ;
$rawdata=Array() ;
$csv=Array() ;
$ar=Array() ;
    $reader->changeSheet(1);
    if($sheets[1]->getName()=="Output .csv"){
    foreach($reader as $v=>$row) {
     $rr=Array();
        foreach($row as $i=>$r) {
	         array_push($rr,$r);
     }
    array_push($csv,$rr);
}
$fil = fopen("output.csv","w");
foreach($csv as $l) {
	fputcsv($fil, $l);
}
$success1="Output.csv exported successfully.Click <a href='output.csv' class=alert-link download > here </a> to Download.";

}else{
  $error="Invalid Sheet name!";
}
    $reader->changeSheet(0);
    if($sheets[0]->getName()=="input excel sheet"){
    foreach($reader as $v=>$row) {
      $sql="INSERT INTO input(sub_category,model_no, short_des,qty,des, sales_price, delivery_charge, max_price, warrenty, brand) VALUES ('$row[0]','$row[1]','$row[2]','$row[3]','$row[4]','$row[5]','$row[6]','$row[7]','$row[8]','$row[9]')";
      if(mysqli_query($database,$sql)){
          $ind+=1;
          $error="";
      }else{
        $error="Error while adding data to database";
      }
   }
   $success=$ind." item(s) added to database successfully!";
 }else{
   $error="Invalid Sheet name!";
   }
$reader->close();
unlink($target_file);
}else{
  $error="Invalid File type.";
}
}
?>
<!DOCTYPE html>
<html lang="en">
<head>
  <title>Excel managger (Internship-Task)</title>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.0/css/bootstrap.min.css">
  <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js"></script>
  <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.0/js/bootstrap.min.js"></script>
</head>
<body>

<div class="jumbotron text-center">
  <h1>Excel Manager</h1>
  <p>Export excel file to database just within a single click!</p>
</div>

<div class="container">
<form action="index.php" method="post" enctype="multipart/form-data">
<div class="form-group">
  <label for="fileToUpload">Choose .xlsx file to export into database:</label>
  <input type="file" class="form-control" name="fileToUpload">
</div>
<button type="submit" class="btn btn-primary">Submit</button>
</form>
<br><br>
<div class="alert alert-success" id="success" role="alert">
  <?=$success?>
</div>
<div class="alert alert-success" id="success1" role="alert">
  <?=$success1?>
</div>
<div class="alert alert-danger" id="error" role="alert">
  <?=$error?>
</div>
</div>
<script type="text/javascript">
  $(function() {
    $(".alert").each(function() {
      var ele=$(this);
      if(ele.text().trim()==""){
        ele.addClass("hidden");
      }else{
        ele.removeClass("hidden");
      }
  })
})
</script>
</body>
</html>

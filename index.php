<!doctype html>
<html class="no-js" lang="en">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Point Calculator</title>
</head>
<body>

<?php

require_once('Classes/PHPExcel.php');
require_once('Classes/PHPExcel/IOFactory.php');


$jsonFile = "position";
$excelFile = "djFR2";

$string = file_get_contents($jsonFile.".json");
$json = json_decode($string, true);

//print_r($json[0]["name"]);

$objReader = PHPExcel_IOFactory::createReader('Excel2007');
$objPHPExcel = $objReader->load($excelFile.'.xlsx');
 
//Itrating through all the sheets in the excel workbook and storing the array data
foreach ($objPHPExcel->getWorksheetIterator() as $worksheet) {
    $arrayData[$worksheet->getTitle()] = $worksheet->toArray();
}

/*
 * Loops through the excel spreadsheet,
 *  grab all mouse click positions.
 */
$dataField = $arrayData[$excelFile.'.txt'];
$dataName = $dataField[0];
$i_mouse = 0;
$i_page = 0;

for (; $i_mouse < count($dataName); $i_mouse++)
	if ($dataName[$i_mouse] == "MOUSE_CLICK_POSITION")
		break;

for (; $i_page < count($dataName); $i_page++)
	if ($dataName[$i_page] == "CURRENT_PAGE")
		break;

echo "<pre>";
print_r($json);
echo "</pre>";

for ($i = 1; $i < count($dataField); $i++) {
	$data = $dataField[$i][$i_mouse];
	$page = $dataField[$i][$i_page];

echo "<pre>";
	print_r($data);
	print_r($page);
echo "</pre>";
}

?>

</body>
</html>

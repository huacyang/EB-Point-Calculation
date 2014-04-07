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
require_once('Classes/PHPExcel/Writer/Excel2007.php');

// the local json file name
$jsonFile = "position";
// the local excel file name
$excelFile = "djFR2";
// decodes the local json file 
$string = file_get_contents($jsonFile.".json");
$json = json_decode($string, true);
// reads the local excel file
$objReader = PHPExcel_IOFactory::createReader('Excel2007');
$objPHPExcel = $objReader->load($excelFile.'.xlsx');
// iterates through all the sheets in the excel workbook and storing the array data
foreach ($objPHPExcel->getWorksheetIterator() as $worksheet) {
    $arrayData[$worksheet->getTitle()] = $worksheet->toArray();
}

$dataField = $arrayData[$excelFile.'.txt'];
$dataName = $dataField[0];
$i_mouse = 0;
$i_page = 0;
$last_col = count($dataName);
$dataField[0][$last_col] = "clicked_page";

// searches for the column that has the mouse click position
for (; $i_mouse < count($dataName); $i_mouse++)
	if ($dataName[$i_mouse] == "MOUSE_CLICK_POSITION")
		break;
// searches for the column that has the current page number
for (; $i_page < count($dataName); $i_page++)
	if ($dataName[$i_page] == "CURRENT_PAGE")
		break;
// iterates through the excel file
for ($i = 1; $i < count($dataField); $i++) {
	$special = array("(", ")");
	$mouse = $dataField[$i][$i_mouse];
	$mouse = str_replace($special, "", $mouse);
	$click = explode(",", $mouse);
	$index = $dataField[$i][$i_page];
	$page = $json[$index];
	$position = $page['position'];
	// iterates through the json file
	for ($n = 0; $n < count($position); $n++) {
		$token = explode(",", $position[$n]);

		if ($click[0] > $token[0] &&
			$click[0] < $token[2] &&
			$click[1] > $token[1] &&
			$click[1] < $token[3]) {

			$dataField[$i][$last_col] = $page['link'][$n];
			break;
		}
	}
}

echo "<pre>";
print_r($dataField);
echo "</pre>";

$fp = fopen($excelFile.'.csv', 'w');

foreach ($dataField as $field) {
	fputcsv($fp, $field);
}

fclose($fp);

?>

</body>
</html>

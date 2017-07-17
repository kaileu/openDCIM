<?php
	#require_once( "PHPExcel/PHPExcel.php" );
	#require_once( "PHPExcel/PHPExcel/Writer/Excel2007.php" );
	#require_once( "facilities.inc.php" );
	require_once( "db.inc.php" );
        require_once( "facilities.inc.php" );
	header("Content-Type: charset=utf-8;");
	$counter=0;
	$json=json_decode($_REQUEST["path"], true);
	foreach($json as $item) {
		if($item['DeviceName']!=''){
			if($counter>0 && $counter+1!=sizeof($json))
				if($item['Label']=='')
					$path= $path.'+++'.$item['DeviceName'].' '.$item['PortNumber'];
				else 
					$path= $path.'+++'.$item['DeviceName'].' '.$item['Label'];
			else if($counter+1==sizeof($json))
				if($item['Label']=='')
					$path= $path.'+++'.$item['DeviceName'].'+++'.$item['PortNumber'];
				else
					$path= $path.'+++'.$item['DeviceName'].'+++'.$item['Label'];
			else if($item['Label']=='')
				$path= $path.$item['Notes'].'+++'.$item['DeviceName'].'+++'.$item['PortNumber'];
			else
				$path= $path.$item['Notes'].'+++'.$item['DeviceName'].'+++'.$item['Label'];
			$counter++;
		}
	}
	//echo var_dump($json);
	$paths = explode("||", $path);
	$path = array();
	foreach($paths as $devs){
		$path[]=explode("+++", $devs);
	}
	$columns        = array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O");
        //echo var_dump($path);
        $sheet = new PHPExcel();

        $sheet->getProperties()->setCreator("openDCIM");
        $sheet->getProperties()->setLastModifiedBy("openDCIM");
        $sheet->getProperties()->setTitle("Device Port Connections");
        $sheet->getProperties()->setSubject("Device Port Detail");

        $sheet->setActiveSheetIndex(0);
        $sheet->getActiveSheet()->SetCellValue('A1',"Bezeichnung");
        $sheet->getActiveSheet()->SetCellValue('B1',"Komponente 1");
        $sheet->getActiveSheet()->SetCellValue('C1',"Port Komponente 1");
        $sheet->getActiveSheet()->SetCellValue('D1',"Patch 1");
        $sheet->getActiveSheet()->SetCellValue('E1',"Patch 2");
        $sheet->getActiveSheet()->SetCellValue('F1',"Patch 3");
        $sheet->getActiveSheet()->SetCellValue('G1',"Patch 4");
        $sheet->getActiveSheet()->SetCellValue('H1',"Patch 5");
        $sheet->getActiveSheet()->SetCellValue('I1',"Patch 6");
        $sheet->getActiveSheet()->SetCellValue('J1',"Patch 7");
        $sheet->getActiveSheet()->SetCellValue('K1',"Patch 8");
        $sheet->getActiveSheet()->SetCellValue('L1',"Patch 9");
        $sheet->getActiveSheet()->SetCellValue('M1',"Patch 10");
        $sheet->getActiveSheet()->SetCellValue('N1',"Komponente 2");
        $sheet->getActiveSheet()->SetCellValue('O1',"Port Komponente 2");
	
	foreach( range('A','M') as $columnID) {
			$sheet->getActiveSheet()->getColumnDimension($columnID)->setAutoSize(true);
		}
	//$sheet->getActiveSheet()->setTitle(__("Connections"));
	$sheet->getActiveSheet()->setTitle("Connections");
	$row = 2;
	$devNum = 1;
	
	foreach ($path as $deviceName){
		$counter=0;
		foreach ( $deviceName as $indexDevice=>$device){
			
			if(sizeof($deviceName)==sizeof($columns)){
				$sheet->getActiveSheet()->SetCellValue($columns[$indexDevice] . $row, $device);
			}
			else if(sizeof($deviceName)>sizeof($columns)){
				echo "Ungültiges Label";
				die;
			}
			else if(sizeof($deviceName)<sizeof($columns) && sizeof($deviceName)>1){
				$diff=(sizeof($columns)-sizeof($deviceName))-1;
				//echo $diff-$counter;
				if(sizeof($columns)-$diff-$counter>3)
					$sheet->getActiveSheet()->SetCellValue($columns[$indexDevice] . $row, $device);
				else
					$sheet->getActiveSheet()->SetCellValue($columns[sizeof($columns)-(sizeof($columns)-$diff-$counter-1)] . $row, $device);
				$counter++;
				
			}
			else{
				echo "Ungültige Eingabe";
				die;
			}
		}
		$row++;
	}
	
	
	$writer = new PHPExcel_Writer_Excel5($sheet);
	if(isset($_REQUEST["temp"]) && $_REQUEST["temp"]=="1"){
		$tmpName = @tempnam(PHPExcel_Shared_File::sys_get_temp_dir(),'tmpcnxs');
		$writer->save($tmpName);
	} else {
		header('Content-Type: application/vnd.ms-excel');
		header( "Content-Disposition: attachment;filename=\"openDCIM-dev" . 'Kabelverbindung' . "-connections.xls\"" );
		$writer->save("php://output");
	}

?>


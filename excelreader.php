<?php
/**
 * Created by PhpStorm.
 * User: stephen
 * Date: 6/21/16
 * Time: 2:11 PM
 */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

// DB
define('DB_DRIVER', 'mysqli');
define('DB_HOSTNAME', 'localhost');
define('DB_USERNAME', 'haika');
define('DB_PASSWORD', '991028');
define('DB_DATABASE', 'opencart_test');
define('DB_PORT', '3306');
define('DB_PREFIX', 'oc_');

require_once "/home/stephen/PHPExcel/Classes/PHPExcel/IOFactory.php";

$db = new mysqli(DB_HOSTNAME, DB_USERNAME, DB_PASSWORD, DB_DATABASE, DB_PORT);

$sql = "CREATE TABLE IF NOT EXISTS oc_product_mro (
  product_id  int(11)  NOT NULL AUTO_INCREMENT,
  product_name  varchar(255) NOT NULL,
  part_num varchar(128) NOT NULL,
  escape_part_num varchar(128) NOT NULL,
  mpn varchar(128),
  brandid varchar(128) NOT NULL,
  price decimal(15,4) NOT NULL,
  weight decimal(15,8),
  ship_time int NOT NULL,
  discription text,
  supplier_id int(11) NOT NULL,
  note text,
  product_time datetime NOT NULL,
  url text NOT NULL,
  min_count int NOT NULL,
  spn varchar(128) NOT NULL,
  type_id varchar(128) NOT NULL,
  PRIMARY KEY (product_id))";

if($db->query($sql) != TRUE){
    die("create table fail.");
}

$files = glob('/home/stephen/data/*.xlsx');
foreach($files as $file){
    echo $file."\n";
    readExcel2Mysql($file, $db);
}

function findNum($str=''){
    $str=trim($str);
    if(empty($str)){return '';}
    $result='';
    for($i=0;$i<strlen($str);$i++){
        if(is_numeric($str[$i])){
            $result.=$str[$i];
        }
    }
    return $result;
}

function readExcel2Mysql($file, $db){
    echo "readExcel2Mysql file = " . $file . "\n";
    $objPHPExcel = PHPExcel_IOFactory::load($file);
    $currentSheet = $objPHPExcel->getSheet();
    $maxCol = $currentSheet->getHighestColumn();
    $maxRow = $currentSheet->getHighestRow();

    for($rowIndex = 1; $rowIndex <= $maxRow; $rowIndex++){
//    for($colIndex = 'A'; $colIndex <= $maxCol; $colIndex++){
//        $addr = $colIndex . $rowIndex;
//        $value = $currentSheet->getCell($addr)->getValue();
//        echo $addr . " : " . $value;
//    }
        echo "rowIndex: " . $rowIndex . "\n";

        $name = $currentSheet->getCell("B" . $rowIndex)->getValue();
        if(!$name){
            $name = "MRO NO NAME";
        }
        $num = $currentSheet->getCell("C" . $rowIndex)->getValue();
        if(!$num){
            $num = $name;
        }
        $enum = $currentSheet->getCell("D" . $rowIndex)->getValue();
        if(!$enum){
            $enum = $num;
        }
        $mpn = $currentSheet->getCell("E" . $rowIndex)->getValue();
        if(!$mpn){
            $mpn = $name;
        }
        $bid = $currentSheet->getCell("F" . $rowIndex)->getValue();
        if(!$bid){
            $bid = $name;
        }
        $price = $currentSheet->getCell("G" . $rowIndex)->getValue();
        $weight = $currentSheet->getCell("H" . $rowIndex)->getValue();
        if(!$weight){
            $weight = 0;
        }
        $st = $currentSheet->getCell("I" . $rowIndex)->getValue();
        $stime = findNum($st);
        if(empty($stime)){
            $stime = 1;
        }
        $dis = $currentSheet->getCell("J" . $rowIndex)->getValue();
        if(!$dis){
            $dis = "MRO NO DIS";
        }
        $sid = $currentSheet->getCell("K" . $rowIndex)->getValue();
        if(!$sid){
            $sid = 0;
        }
        $note = $currentSheet->getCell("L" . $rowIndex)->getValue();
        if(!$note){
            $note = "MRO no comment";
        }
        $time = date("y-m-d", time());//$currentSheet->getCell("M" . $rowIndex)->getValue();
        $url = "https://www.mrosupply.com";
        $min =  $currentSheet->getCell("O" . $rowIndex)->getValue();
        if(!$min){
            $min = 1;
        }
        $spn = $currentSheet->getCell("P" . $rowIndex)->getValue();
        if(!$spn){
            $spn = $num;
        }
        $tid = $currentSheet->getCell("Q" . $rowIndex)->getValue();
        if(!$tid){
            $tid = 0;
        }

        $s = "INSERT INTO oc_product_mro(product_name, part_num, escape_part_num, mpn, brandid, price, weight, ship_time, discription, supplier_id, note, product_time, url, min_count, spn, type_id) VALUES".
            "('" . $name . "','" .
            $num . "','" .
            $enum  . "','" .
            $mpn  . "','" .
            $bid  . "','" .
            $price   . "','".
            $weight   . "'," .
            (int)$stime  . ",'" .
            $db->real_escape_string($dis)  . "'," .
            (int)$sid  . ",'".
            $note  . "','" .
            $time  . "','".
            $url  . "'," .
            (int)$min  . ",'" .
            $spn  . "','" .
            $tid . "'" .
            ")";

//        echo "insert row id = " . $rowIndex . " sql = " . $s . "\n";

        try{
            $result = $db->query($s);
        }catch(Exception $e){
            echo "insert exception: " . $e->getMessage() . "\n";
        }

        if(!$result){
            echo "insert failed row id = " . $rowIndex . "\n";
        }
    }

    echo "insert to db finish. row count = " . ($rowIndex - 1) . "\n";
}


$db->close();
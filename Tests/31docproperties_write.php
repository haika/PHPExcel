<?php
/**
 * PHPExcel
 *
 * Copyright (C) 2006 - 2011 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2011 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */

/** Error reporting */
error_reporting(E_ALL);

date_default_timezone_set('Europe/London');

/** PHPExcel */
require_once '../Classes/PHPExcel.php';


// Create new PHPExcel object
echo date('H:i:s') . " Create new PHPExcel object" , PHP_EOL;
$objPHPExcel = new PHPExcel();

// Set properties
echo date('H:i:s') . " Set properties" , PHP_EOL;
$objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
							 ->setLastModifiedBy("Franklin")
							 ->setTitle("Office 2007 XLSX Test Document")
							 ->setSubject("Office 2007 XLSX Test Document")
							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
							 ->setKeywords("office 2007 openxml php")
							 ->setCategory("Test result file");
// Add some data
echo date('H:i:s') . " Add some data" , PHP_EOL;
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', 'Hello')
            ->setCellValue('B2', 'world!');

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);

// Save Excel 2007 file
echo date('H:i:s') . " Write to Excel2007 format" , PHP_EOL;
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save(str_replace('.php', '.xlsx', __FILE__));

echo date('H:i:s') . " Adjust properties" , PHP_EOL;
$objPHPExcel->getProperties()->setTitle("Office 95 XLS Test Document")
							 ->setSubject("Office 95 XLS Test Document")
							 ->setDescription("Test document for Office 95 XLS, generated using PHP classes.")
							 ->setKeywords("office 95 openxml php");

// Save Excel5 file
echo date('H:i:s') . " Write to Excel5 format" , PHP_EOL;
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save(str_replace('.php', '.xls', __FILE__));


// Echo memory peak usage
echo date('H:i:s') . " Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB" , PHP_EOL;

// Echo done
echo date('H:i:s') . " Done writing files." , PHP_EOL;
unset($objWriter);


echo PHP_EOL;
// Reread File
echo date('H:i:s') . " Reread Excel5 file" , PHP_EOL;
$objPHPExcelRead = PHPExcel_IOFactory::load('31docproperties_write.xls');

// Set properties
echo date('H:i:s') . " Get properties" , PHP_EOL;
echo 'Creator : '.$objPHPExcelRead->getProperties()->getCreator() , PHP_EOL;
echo 'LastModifiedBy : '.$objPHPExcelRead->getProperties()->getLastModifiedBy() , PHP_EOL;
echo 'Title : '.$objPHPExcelRead->getProperties()->getTitle() , PHP_EOL;
echo 'Subject : '.$objPHPExcelRead->getProperties()->getSubject() , PHP_EOL;
echo 'Description : '.$objPHPExcelRead->getProperties()->getDescription() , PHP_EOL;
echo 'Keywords : '.$objPHPExcelRead->getProperties()->getKeywords() , PHP_EOL;
echo 'Category : '.$objPHPExcelRead->getProperties()->getCategory() , PHP_EOL;

// Echo memory peak usage
echo date('H:i:s') . " Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB" , PHP_EOL;





?>
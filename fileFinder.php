<?php
/**
 * Created by PhpStorm.
 * User: stephen
 * Date: 6/22/16
 * Time: 3:31 PM
 */

$files = glob('/home/stephen/data/*.xlsx');
foreach($files as $file){
    echo $file . '\n';
}
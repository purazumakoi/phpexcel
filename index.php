<?php
/**
 * Created by PhpStorm.
 * User: kenji
 * Date: 14/09/11
 * Time: 13:49
 */

set_include_path(get_include_path().PATH_SEPARATOR.$_SERVER["DOCUMENT_ROOT"].'/kenji/git_repo/phpexcel/PHPExcel_1.8.0_doc/Classes/');

include_once( 'PHPExcel.php' );


// 新しいエクセルファイルを作成する
$objPHPExcel = new PHPExcel();

/*
この間で、エクセルファイルを編集する。
*/

// "Excel2007" 形式で保存する
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('./sample.xlsx');
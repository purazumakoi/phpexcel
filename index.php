<?php
/**
 * Created by PhpStorm.
 * User: kenji
 * Date: 14/09/11
 * Time: 13:49
 */

set_include_path(get_include_path().PATH_SEPARATOR.$_SERVER["DOCUMENT_ROOT"].'/kenji/git_repo/phpexcel/PHPExcel_1.8.0_doc/Classes/');

include_once( 'PHPExcel.php' );


// �V�����G�N�Z���t�@�C�����쐬����
$objPHPExcel = new PHPExcel();

/*
���̊ԂŁA�G�N�Z���t�@�C����ҏW����B
*/

// "Excel2007" �`���ŕۑ�����
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('./sample.xlsx');
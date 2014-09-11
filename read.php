<?php
/**
 * Created by PhpStorm.
 * User: kenji
 * Date: 14/09/11
 * Time: 13:49
 *
 * via�Fhttp://www.pxt.jp/ja/diary/article/281/
 */

set_include_path(get_include_path().PATH_SEPARATOR.$_SERVER["DOCUMENT_ROOT"].'/kenji/git_repo/phpexcel/PHPExcel_1.8.0_doc/Classes/');

include_once( 'PHPExcel.php' );


// �����t�@�C���̓ǂݍ��݂̏ꍇ
$objPHPExcel = PHPExcel_IOFactory::load("./read.xlsx");


/**
 * �V�[�g�n
 */
		// 0�Ԗڂ̃V�[�g���A�N�e�B�u�ɂ���i�V�[�g�͍����珇�ɁA0�A1�C2�E�E�E�j
		$objPHPExcel->setActiveSheetIndex(0);

		// �A�N�e�B�u�ɂ����V�[�g�̏����擾
		$objSheet = $objPHPExcel->getActiveSheet();

		// �V�[�g���ύX
		$objSheet->setTitle('test');

/**
 * �Z���Ƀf�[�^��������
 */
		// �Z���Ƀf�[�^��������
		$objSheet->getCell('A1')->setValue('�e�X�g');
		//�������̏������ł��悢
		$objSheet->setCellValue('A1', '�e�X�g');

		// �����̏�����
		$objSheet->getCell('B1')->setValue(120);
		$objSheet->getCell('B2')->setValue(523);
		$objSheet->getCell('B3')->setValue('=B1+B2');

/**
 * �Z������f�[�^��ǂݎ��
 */
		// �l�̎擾
		$val = $objSheet->getCell('B3')->getValue();

/**
 * ����
 */
		// �Z���̌���
		$objSheet->mergeCells('A10:D14');

		// �t�H���g
		$objStyle->getFont()->setName('���C���I');

		// �t�H���g�T�C�Y
		$objStyle->getFont()->setSize(14);

		// �Z���̉��Ɍr��������
		$objStyle->getBorders()->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
		$objStyle->getBorders()->getRight()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
		$objStyle->getBorders()->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
		$objStyle->getBorders()->getLeft()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

		// �r���̈ꊇ�w��
		$objStyle->applyFromArray(array(
		'borders' => array(
		'top'     => array('style' => PHPExcel_Style_Border::BORDER_THIN),
		'bottom'  => array('style' => PHPExcel_Style_Border::BORDER_THIN),
		'left'    => array('style' => PHPExcel_Style_Border::BORDER_THIN),
		'right'   => array('style' => PHPExcel_Style_Border::BORDER_THIN)
		)
		));

		// ����
		$objStyle->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
		// �Z���^�[��
		$objStyle->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		// �E��
		$objStyle->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

		// �w�i�F�w��
		$objStyle->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);//��fillType��ݒ肵�Ȃ��Ɣw�i�F�͂��Ȃ��B
		$objStyle->getFill()->getStartColor()->setRGB('dddddd');

/**
 * �ۑ�
 */
// "Excel2007" �`���ŕۑ�����
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('./read.xlsx');
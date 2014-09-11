<?php
/**
 * Created by PhpStorm.
 * User: kenji
 * Date: 14/09/11
 * Time: 13:49
 *
 * via：http://www.pxt.jp/ja/diary/article/281/
 */

set_include_path(get_include_path().PATH_SEPARATOR.$_SERVER["DOCUMENT_ROOT"].'/kenji/git_repo/phpexcel/PHPExcel_1.8.0_doc/Classes/');

include_once( 'PHPExcel.php' );


// 既存ファイルの読み込みの場合
$objPHPExcel = PHPExcel_IOFactory::load("./read.xlsx");


/**
 * シート系
 */
		// 0番目のシートをアクティブにする（シートは左から順に、0、1，2・・・）
		$objPHPExcel->setActiveSheetIndex(0);

		// アクティブにしたシートの情報を取得
		$objSheet = $objPHPExcel->getActiveSheet();

		// シート名変更
		$objSheet->setTitle('test');

/**
 * セルにデータ書き込む
 */
		// セルにデータ書き込み
		$objSheet->getCell('A1')->setValue('テスト');
		//こっちの書き方でもよい
		$objSheet->setCellValue('A1', 'テスト');

		// 数式の書き方
		$objSheet->getCell('B1')->setValue(120);
		$objSheet->getCell('B2')->setValue(523);
		$objSheet->getCell('B3')->setValue('=B1+B2');

/**
 * セルからデータを読み取る
 */
		// 値の取得
		$val = $objSheet->getCell('B3')->getValue();

/**
 * 書式
 */
	// デフォルトのスタイル
	$objStyle = $objSheet->getDefaultStyle();

	// セル個別のスタイル
	$objStyle = $objSheet->getStyle('B1');

		// セルの結合
		$objSheet->mergeCells('A10:D14');

		// フォント
		//$objStyle->getFont()->setName('メイリオ');

		// フォントサイズ
		$objStyle->getFont()->setSize(14);

		// セルの下に罫線を引く
		$objStyle->getBorders()->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
		$objStyle->getBorders()->getRight()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
		$objStyle->getBorders()->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
		$objStyle->getBorders()->getLeft()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

		// 罫線の一括指定
		$objStyle->applyFromArray(array(
		'borders' => array(
		'top'     => array('style' => PHPExcel_Style_Border::BORDER_THIN),
		'bottom'  => array('style' => PHPExcel_Style_Border::BORDER_THIN),
		'left'    => array('style' => PHPExcel_Style_Border::BORDER_THIN),
		'right'   => array('style' => PHPExcel_Style_Border::BORDER_THIN)
		)
		));

		// 左寄せ
		$objStyle->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
		// センター寄せ
		$objStyle->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		// 右寄せ
		$objStyle->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

		// 背景色指定
		$objStyle->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);//←fillTypeを設定しないと背景色はつかない。
		$objStyle->getFill()->getStartColor()->setRGB('dddddd');

/**
 * 保存
 */
// "Excel2007" 形式で保存する
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('./read.xlsx');



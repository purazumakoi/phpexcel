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

/**
 * 元ファイルからの読み込み
 *
 *
 */

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
//		// セルにデータ書き込み
//		$objSheet->getCell('A1')->setValue('テスト');
//		//こっちの書き方でもよい
//		$objSheet->setCellValue('A1', 'テスト');
//
//		// 数式の書き方
//		$objSheet->getCell('B1')->setValue(120);
//		$objSheet->getCell('B2')->setValue(523);
//		$objSheet->getCell('B3')->setValue('=B1+B2');

/**
 * セルからデータを読み取る
 */
// 配列定義
$ary_data = array();

// データの行数
$line = 300;
for($i=2; $i<$line; $i++){
	$cel_a = $objSheet->getCell('A'.$i)->getValue();
	$cel_b = $objSheet->getCell('B'.$i)->getValue();
	$cel_c = $objSheet->getCell('C'.$i)->getValue();
	$cel_d = $objSheet->getCell('D'.$i)->getValue();
	$cel_e = $objSheet->getCell('E'.$i)->getValue();
	$cel_f = $objSheet->getCell('F'.$i)->getValue();

	// 日付がはいっていたらその日に切替える
	if($cel_a){
		$now_date = $cel_a;
	}
	// 区分があれば区分をそれにする
	if($cel_b){
		$now_kubun = $cel_b;
		$key = 0;
	}

	// 配列の数取得
	$key = intval(@max( array_keys((array)$ary_data[$now_date][$now_kubun]) )) + 1;


	// 配列作成
	if($cel_e){
		$ary_data[$now_date][$now_kubun][$key]['calorie'] = $cel_c;
		$ary_data[$now_date][$now_kubun][$key]['price'] = $cel_d;
		$ary_data[$now_date][$now_kubun][$key]['name'] = $cel_e;
		$ary_data[$now_date][$now_kubun][$key]['background-color'] = $cel_f;
	}

}

//print_r($ary_data); print " ";




/**
 * 新ファイルへの書き出し
 *
 *
 */
// 既存ファイルの読み込みの場合
$objPHPExcel = PHPExcel_IOFactory::load("./put.xlsx");

// 0番目のシートをアクティブにする（シートは左から順に、0、1，2・・・）
$objPHPExcel->setActiveSheetIndex(0);

// アクティブにしたシートの情報を取得
$objSheet = $objPHPExcel->getActiveSheet();

// シート名変更
$objSheet->setTitle('第一工場 昼');

/**
 * データ構築
 */

print_r($ary_data); print " ";

// 日付ループカウント
$cnt = 0;


foreach($ary_data as $key => $val){
	// 日付いれ列名
	$date_write_row = chr(98+$cnt);
	$objSheet->setCellValue($date_write_row  ."4", $key);

	// スタート行
	$line = 6;

	foreach($val as $key2 => $val2){


		foreach($val2 as $key3 => $val3){
			if($val3['name']){
				// 実データ入れ列名
				$date_write_data = 98+$cnt;


				// 実データセット

				// 区分
				$objSheet->setCellValue(chr($date_write_data)  .$line, $key2);

				// その他
				$objSheet->setCellValue(chr($date_write_data+1) . $line, $val3['calorie']);
				$objSheet->setCellValue(chr($date_write_data+2) . $line, $val3['price']);
				$objSheet->setCellValue(chr($date_write_data+3) . $line, $val3['name']);
				//$objSheet->setCellValue($date_write_data . $line, $val2['background-color']);

				$line++;

			}
		}
	}

	// 日付かわりました。
	$cnt += 4;
}

$objSheet->setCellValue("B" . $line,'売価には消費税8％が含まれています。');
$objSheet->setCellValue("F" . $line,'仕入れの都合により内容が変更する場合があります。');

/**
 * 保存
 */
// "Excel2007" 形式で保存する
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('./put.xlsx');

echo "<br><br>「put.xlsx」を書き出しました。";


// 権限変えます
chmod("./put.xlsx", 0777);

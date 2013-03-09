<?php
include 'phpexcel/PHPExcel.PHP';
include 'phpexcel/PHPExcel/Writer/Excel2007.PHP';

class MyExcel {

    public static function makeExcel($title, $data, $excle_name) {
        $exclefile = 'test.xlsx';

        $Excel = new PHPExcel();
        $Excel->setActiveSheetIndex(0);
        $Excel->getSheet()->setTitle($title);

        $cell_one = $data[0];
        $j = 0;
        foreach ($cell_one as $k => $v) {
            $Excel->getSheet()->setCellValue(self::getCharByNunber($j) . '1', $k);
            $j++;
        }

        $x = 2;
        foreach ($data as $value) {
            $y = 0;
            foreach ($value as $k => $v) {
                $Excel->getSheet()->setCellValue(self::getCharByNunber($y) . $x, $v);
                $y++;
            }
            $x++;
        }

        $objwriter = new PHPExcel_Writer_Excel2007($Excel);
        $objwriter->save($exclefile);
        // TMDebugUtils::debugLog('make ' . $exclefile);
        return $exclefile;
    }

    protected static function getCharByNunber($num) {
        $num = intval($num);
        $arr = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',);
        return $arr[$num];
    }

}

//导出用户基本信息
function test() {
    $excelTitle = '用户基本信息';
    $excelFileName = 'userbase';
    $data = array();
    // $rs = ....  这里rs是从db读取的数据
	$rs = array(array('name' => 'aa', 'age' => 23));
    for ($i = 0; $i < count($rs); $i++) {         
		$data[$i]['姓名'] = $rs[$i]['name'];
        $data[$i]['年龄'] = $rs[$i]['age'];
        $data[$i]['日期'] = date('Y-m-d');
    }
    MyExcel::makeExcel($excelTitle, $data, $excelFileName);
}

test();
?>
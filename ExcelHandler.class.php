<?php
/**
 * Created by PhpStorm.
 * User: kazuma
 * Date: 2017/12/7
 * Time: 下午2:32
 */

namespace Prouter\lib\ExcelHandle;
Vendor('PHPExcel.PHPExcel');
Vendor('PHPExcel.PHPExcel.IOFactory');

/**
 * Class ExcelHandler
 * @package Prouter\lib\ExcelHandle
 * excel处理类，读取数据与导出数据
 * @caution 目前只支持.xls格式的数据解析,不支持.xlsx格式的
 */
class ExcelHandler
{
    /**
     * @var array
     * excel表列号与字段号的映射
     */
    static $headers = array('A', 'B', 'C', 'D', 'E', 'F', 'G',
        'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T');
    /**
     * @var array
     * self::headers的翻转数组
     */
    static $reversedMap = array('A' => 0, 'B' => 1, 'C' => 2, 'D' => 3, 'E' => 4, 'F' => 5, 'G' => 6,
        'H' => 7, 'I' => 8, 'J' => 9, 'K' => 10, 'L' => 11, 'M' => 12, 'N' => 13, 'O' => 14);

    /**
     * 从源数据生成工作簿输出到页面上
     * @param $bookName
     * @param array $tables
     * @param int $limit
     * $bookName是要生成的文件名称
     * $tables是源数据，多个excelTable对象组成的数组
     * $limit是worksheet的限制
     */
    public static function buildBook($bookName, array $tables, $limit = 1)
    {

        $excelInstance = new \PHPExcel();
        foreach ($tables as $index => $table) {
            if ($index == $limit) break;
            if ($index) $excelInstance->createSheet();
            $excelInstance->setActiveSheetIndex($index);
            self::buildSheet($table, $excelInstance);
        }
        ob_clean();
        header('pragma:public');
        header('Content-type:application/vnd.ms-excel;charset=utf-8;name="' . $bookName . '.xls"');
        header("Content-Disposition:attachment;filename={$bookName}.xls");//attachment新窗口打印inline本窗口打印
        $objWriter = \PHPExcel_IOFactory::createWriter($excelInstance, 'Excel5');
        $objWriter->save('php://output');
    }


    /**
     * @param ExcelTable $table
     * @param \PHPExcel $excelInstance
     * 通过sql查询结果格式的数据来构造workbook
     */
    private static function buildSheet(ExcelTable $table, \PHPExcel $excelInstance)
    {
        $fields = $table->getFields();
        $rowNum = 1;
        $fieldToHeader = array();
        $excelInstance->getActiveSheet()->setTitle($table->getName());
        foreach ($fields as $i => $field) {
            $excelInstance->getActiveSheet()->setCellValue(self::$headers[$i] . $rowNum, $field);
            $fieldToHeader[$field] = self::$headers[$i];
        }
        foreach ($table->getRows() as $row) {
            $rowNum++;
            foreach ($row as $key => $value) {
                $excelInstance->getActiveSheet()->setCellValue($fieldToHeader[$key] . $rowNum, $value);
            }
        }
    }

    /**
     * @param $filename
     * @param int $sheetLimit
     * @return array
     * 从excel workbook文件中读取数据
     * $filename是上传文件保存的tmp_name
     * 返回多个excelTable对象组成的数组
     */
    public static function readTablesFromFile($filename)
    {

        $excelInstance = \PHPExcel_IOFactory::load($filename);
        $sheetCount = $excelInstance->getSheetCount();
        $excelTables = array();
        for ($sheetIndex = 0; $sheetIndex < $sheetCount; $sheetIndex++) {
            $excelInstance->setActiveSheetIndex($sheetIndex);
            $rowsCount = $excelInstance->getActiveSheet()->getHighestRow(); //获取总行数
            $highestColumn = $excelInstance->getActiveSheet()->getHighestColumn();
            $columnsCount = self::$reversedMap[$highestColumn] + 1;
            for ($i = 0; $i < $columnsCount; $i++) {
                $header = self::$headers[$i];
                $field = $excelInstance->getActiveSheet()->getCell("{$header}1")->getValue();
                if ($field) {
                    $fieldsMap[$header] = $field;
                } else {
                    continue;
                }
            }
            $rows = array();
            for ($rowIndex = 2; $rowIndex <= $rowsCount; $rowIndex++) {
                $row = array();
                foreach ($fieldsMap as $header => $field) {
                    $row[$field] = $excelInstance->getActiveSheet()->getCell("{$header}{$rowIndex}")->getValue();
                }
                $rows[] = $row;
            }
            $excelTables[] = new ExcelTable($rows, $excelInstance->getActiveSheet()->getTitle());
        }
        return $excelTables;
    }

}
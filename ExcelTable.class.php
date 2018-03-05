<?php
/**
 * Created by PhpStorm.
 * User: lucio
 * Date: 2017/12/7
 * Time: 上午11:16
 */

namespace Prouter\lib\ExcelHandle;
/**
 * Class ExcelTable
 * @package Prouter\lib\ExcelHandle
 * 该类封装了sqlResult格式的数据，对应一个worksheet的数据
 * @example $this->rows = array(array('uid'=>1,'username'=>'lucio'),array('uid'=>2,'username'=>'jack'))
 * 用一个sqlResult格式的数组和一个字符串可以初始化该类，并提取类的数据
 */
class ExcelTable
{
    /**
     * @var array
     * 基础数据行
     */
    private $rows = array();
    /**
     * @var
     * worksheet名
     */
    private $name;

    /**
     * ExcelHandler constructor.
     * 一个handler处理一个页面
     */
    public function __construct($rows, $name)
    {
        $this->rows = $rows;
        $this->name = $name;
    }

    /**
     * @return array
     * 获取所有字段组成的数组
     * @example array('uid','username')
     */
    public function getFields()
    {
        return array_keys(current($this->rows));
    }

    /**
     * @return int
     * 获取行数，在该示例中返回2
     */
    public function rowsCount()
    {
        return count($this->rows);
    }

    /**
     * @return array
     * @example  array(array('uid'=>1,'username'=>'lucio'),array('uid'=>2,'username'=>'jack'))
     *
     */
    public function getRows()
    {
        return $this->rows;
    }

    /**
     * @return mixed
     * 获取worksheet名
     */
    public function getName()
    {
        return $this->name;
    }

    /**
     * @param $columnParam
     * @return array
     * 根据字段名组成的数组返回每个列的值组成的数组或者单列的数组，视参数而定
     * @example array('uid'=>array(1, 2), 'username'=>array('lucio','jack')) or array(1,2)
     */
    public function getValuesByFields($columnParam)
    {

        if (is_array($columnParam)) {
            $totalValues = array();
            foreach ($this->rows as $index => $row) {
                foreach ($columnParam as $column) {
                    if (in_array($column, $columnParam))
                        $totalValues[$index] = $row[$column];
                }
            }
            return $totalValues;
        }elseif(is_string($columnParam)){
            $values = array();
            foreach($this->rows as $row){
                $values[] = $row[$columnParam];
            }
            return $values;
        }else{
            throw_exception('invalid parameters');
        }
    }

    /**
     * @param $index
     * @return mixed
     * 根据行号获取制定行，index从0开始
     * @example array('uid'=>'1', 'username'=>'lucio')
     */
    public function getRowByIndex($index){
        return $this->rows[(int) $index];
    }

    /**
     * @param $index
     * @param $field
     * @return mixed
     * 根据行号和字段名获取值
     * @example $index = 0, $field = 'username', return 'lucio'
     */
    public function getValueByIndexAndField($index, $field){
        return $this->rows[$index][$field];
    }

}
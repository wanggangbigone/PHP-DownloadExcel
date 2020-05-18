<?php
/**
 * Created by PhpStorm.
 * User: owen
 * Date: 2020/5/18
 * Time: 11:32
 */

require_once  "PHPExcel.php";

class DownExcel
{
    // 封装导出Excel
    public function downloadExcel($title,$cells,$data=array())
    {
        $EXCEL = new \PHPExcel;

        //横向单元格标识
        $cellNum = $this->getCellNum($cells);
        $maxMerge = $this->getMaxMerge($cells);
        $cellName = $this->getCellName($cellNum);
        $tag = implode('',array_keys($cells));
        $this->mergeCells($EXCEL,$cellName,$maxMerge,$cells,0,$tag);
        $this->inputData($EXCEL,$data,$cellName,$maxMerge,$cellNum);

        $objWrite = \PHPExcel_IOFactory::createWriter($EXCEL, 'Excel2007');
        header('pragma:public');
        header("Content-Disposition:attachment;filename=$title.xls");
        $objWrite->save('php://output');exit;

    }


    public function getCellNum($example_data,&$count = 0)
    {
        if(is_array($example_data)) {
            foreach ($example_data as $v){
                if(is_array($v)) {
                    $this->getCellNum($v,$count);
                }else{
                    $count ++;
                }
            }
        }

        return  $count;
    }

    public function getCellName($num)
    {
        $base_cell = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z');
        $times = (int)ceil($num/26);
        if($times == 1 ){
            $return_cell = $base_cell;
        }else {
            $return_cell = $base_cell;
            for ($i=1 ; $i< $times ;$i++){
                foreach ($base_cell as $v){
                    array_push($return_cell,$base_cell[$i-1].$v);
                }
            }

        }

        return $return_cell;

    }

    public function getMaxMerge($example_data)
    {
        if(!is_array($example_data)) {
            return 0;
        }else {
            $max = 0;
            foreach($example_data as $item)
            {
                $t = $this->getMaxMerge($item);
                if( $t > $max) $max = $t;
            }
            return $max + 1;
        }
    }


    /*  格式调整 */
    //  $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(20);//所有单元格（行）默认高度
    //     $objPHPExcel->getActiveSheet()->getDefaultColumnDimension()->setWidth(20);//所有单元格（列）默认宽度
    //     $objPHPExcel->getActiveSheet()->getRowDimension('1')->setRowHeight(30);//设置行高度
    //     $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(30);//设置列宽度
    //     $objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setSize(18);//设置文字大小
    //     $objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);//设置是否加粗
    //     $objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);// 设置文字颜色
    //     $objPHPExcel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);//设置文字居左（HORIZONTAL_LEFT，默认值）中（HORIZONTAL_CENTER）右（HORIZONTAL_RIGHT）
    //     $objPHPExcel->getActiveSheet()->getStyle('A1')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);//垂直居中
    //     $objPHPExcel->getActiveSheet()->getStyle('A1')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);//设置填充颜色
    //     $objPHPExcel->getActiveSheet()->getStyle('A1')->getFill()->getStartColor()->setARGB('FF7F24');//设置填充颜色


    // 自动合并
    public function mergeCells($excel,$cellName,$maxMerge,$example_data,$parr_start = 0)
    {
        $values = array_values($example_data);
        $keys = array_keys($example_data);
        foreach ($values as $k => $val) {
            $start = $maxMerge - $this->getMaxMerge($values) + 1;
            if($start - $parr_start > 1) $start = $parr_start + 1;
            $key = $this->getCellNum(array_slice($values,0,$k));
            $excel->getActiveSheet()->getStyle($cellName[$key].$start)->getFont()->setSize(12);
            $excel->getActiveSheet()->getStyle($cellName[$key].$start)->getFont()->setBold(true);
            $excel->getActiveSheet()->getColumnDimension($cellName[$key])->setWidth(20);
            $excel->getActiveSheet()->getStyle($cellName[$key].$start)->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            $excel->getActiveSheet()->getStyle($cellName[$key].$start)->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);

            if(!is_array($val)){
                $excel->getActiveSheet(0)->mergeCells("$cellName[$key]$start:$cellName[$key]$maxMerge");
                $excel->setActiveSheetIndex(0)->setCellValue($cellName[$key].$start, $val);
            }else{
                $val_count = $this->getCellNum($val);
                $key1 = $key + $val_count - 1;
                $excel->getActiveSheet(0)->mergeCells("$cellName[$key]$start:$cellName[$key1]$start");
                $excel->setActiveSheetIndex(0)->setCellValue($cellName[$key].$start, $keys[$k]);
                $child_cell_name = [];
                for ($i = $key ;$i<$key + $val_count; $i++){
                    $child_cell_name[] = $cellName[$i];
                }

                $this->mergeCells($excel,$child_cell_name,$maxMerge,$val,$start);

            }
        }

    }

    // 能够简单一维合并行的数据填充
    public function inputData($excel, $data,$cellName,$startPosition,$cell_count)
    {
        foreach ($data as $v){           
            $max_num = $this->getChildMaxNum($v);
            $v = array_values($v);
            $position = ++$startPosition;
            for ($i=0; $i<$cell_count ;$i++){
                if(!is_array($v[$i]) && $max_num>1){
                    $end_position = $position + $max_num - 1;
                    $excel->getActiveSheet(0)->mergeCells("$cellName[$i]$position:$cellName[$i]$end_position");
                    $excel->getActiveSheet(0)->getStyle($cellName[$i].$position)->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $excel->getActiveSheet(0)->getStyle($cellName[$i].$position)->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                    $excel->getActiveSheet(0)->setCellValue($cellName[$i].$position, $v[$i]);
                }elseif( is_array($v[$i]) ){
                    $start_point = $position;
                    $start_index = 0;
                    while($start_point < $position + $max_num){
                        $excel->getActiveSheet(0)->getStyle($cellName[$i].$start_point)->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                        $excel->getActiveSheet(0)->setCellValue($cellName[$i].$start_point, $v[$i][$start_index]);
                        $start_point++;
                        $start_index++;
                    }
                }else{
                    $excel->getActiveSheet()->getStyle($cellName[$i].$position)->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                    $excel->getActiveSheet(0)->setCellValue($cellName[$i].$position, $v[$i]);

                }

            }

            if($max_num > 1) $startPosition = $position + $max_num  -1;
        }

    }



}

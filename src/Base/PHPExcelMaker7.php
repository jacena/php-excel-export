<?php

namespace Jacena\PhpExcelMaker\Base;


class PHPExcelMaker7
{
    /**
     *  column auto width
     */
    private $columnAutoSize;


    function __construct($columnAutoSize = false)
    {
        $this->columnAutoSize = $columnAutoSize;
    }

    /**
     * function php array to export excel
     *
     * @param array [topic keys]
     * @param array [topic title map]
     * @param array [data array]
     * @param string [filename]
     * @param boolean [makefile or browser down]
     * @param string [output dir]
     * @return void
     */
    public function exportExcel(array $keys, array $title, array $data, $filename = "data_export", $returnFile = true, $output = '.')
    {
        try {
            if (empty($keys)) {
                $keys = array_keys($data[0]);
            }
            if (empty($title)) {
                $title = array_keys($data[0]);
            }
            if (empty($keys) && empty($title)) {
                echo 'params error';
                exit;
            }
            $objPHPExcel = new \PHPExcel();
            $objPHPExcel->getProperties()->setCreator("jacena")
                ->setLastModifiedBy("jacena")
                ->setTitle("Form Data")
                ->setSubject("Form Subject")
                ->setDescription("Form Description")
                ->setKeywords("Form keyword")
                ->setCategory("Form category");
            $cellTitle = [
                'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
                'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ'
            ]; //最多52列
            $valuesCount = count($data);
            //记录每列最大的长度, 如果最长则让列自适应
            $columnStrLength = [];

            //写第一行title
            $titleCount = count($title);
            $i = 0;
            foreach ($title as $t) {
                $columnStrLength[$i] = strlen($t) + 10;

                $column = $cellTitle[$i++];
                $number = 1;
                $objPHPExcel->setActiveSheetIndex(0)->setCellValue($column . $number, $t);

                //第一排标题需要加粗并剧中
                $objPHPExcel->getActiveSheet()->getStyle($column . $number)->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                $objPHPExcel->getActiveSheet()->getStyle($column . $number)->getFont()->setBold(true);
                $objPHPExcel->getActiveSheet()->getColumnDimension($column)->setWidth($columnStrLength[$i - 1] * 1); //设置宽度自适应
                // $objPHPExcel->getActiveSheet()->getColumnDimension($column)->setAutoSize(true); //设置宽度自适应
            }
            //写数据行
            for ($i = 0; $i < $valuesCount; $i++) { //循环所有行
                $value = is_array($data[$i]) ? $data[$i] : [];

                for ($j = 0; $j < $titleCount; $j++) { //显示每一行的数据
                    //如果数据的key在data中存在则显示, 否则为空
                    if (array_key_exists($keys[$j], $value)) {
                        $column = $cellTitle[$j];
                        $number = $i + 2; //当前的列数
                        $objPHPExcel->setActiveSheetIndex(0)->setCellValue($column . $number, $value[$keys[$j]]); //填充内容
                        if (max($columnStrLength[$j], strlen($value[$keys[$j]])) > $columnStrLength[$j]) {
                            //$columnStrLength[$j] = max($columnStrLength[$i], strlen($value[$keys[$j]]));
                            if ($this->columnAutoSize) {
                                $objPHPExcel->getActiveSheet()->getColumnDimension($column)->setAutoSize(true); //设置宽度自适应
                            } else {
                                $objPHPExcel->getActiveSheet()->getColumnDimension($column)->setWidth($columnStrLength[$j] * 1); //设置宽度自适应
                            }
                        }
                    }
                } //end for j
            } //end for i
            $objPHPExcel->getActiveSheet()->setTitle($filename);
            $objPHPExcel->setActiveSheetIndex(0);
            // Redirect output to a client’s web browser (Excel5)
            $filename = $filename ? $filename : date("Ymd", time());
            $newFileName = $filename . '_' . date('Y-m-d-H:i:s') . '.xlsx';
            if ($returnFile) {
                $dirPath = $output; //注意要设置为绝对路径
                $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
                $objWriter->save($dirPath . '/' . $newFileName);
                return $newFileName;
            } else {
                ob_end_clean();
                header('Content-Type: application/vnd.ms-excel');
                header('Access-Control-Expose-Headers: Content-Disposition');
                header('Content-Disposition: attachment;filename=' . $newFileName);
                header('Cache-Control: max-age=0');
                header('Cache-Control: max-age=1');
                header('Expires: Mon, 26 Jul 1997 05:00:00 GMT');
                header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT');
                header('Cache-Control: cache, must-revalidate');
                header('Pragma: public');

                $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
                $objWriter->save('php://output');
                exit;
            }
        } catch (\Exception $e) {
            return false;
        }
    }
}

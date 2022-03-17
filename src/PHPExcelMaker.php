<?php

namespace Jacena\PhpExcelMaker;
// require '../PhpSpreadsheet/vendor/autoload.php';

use Jacena\PhpExcelMaker\Base\PHPExcelMaker7;
use Jacena\PhpExcelMaker\Base\PHPExcelMaker8;

// require_once  '../PhpSpreadsheet/samples/Bootstrap.php';

class PHPExcelMaker
{
    /**
     *  column auto width
     */
    private $columnAutoSize;
    private $phpVersion;

    function __construct($columnAutoSize = false, $timezone = "Asia/Chongqing")
    {
        $this->phpVersion = $this->getPHPVersion();
        $this->columnAutoSize = $columnAutoSize;
        date_default_timezone_set($timezone);
    }

    /**
     * php version >= php7.3
     *
     * @return bool
     */
    public function getPHPVersion()
    {
        return version_compare(PHP_VERSION, '7.3.0', 'ge') ? true : false;
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
        if (empty($data)) {
            throw new \Exception("[ data empty ]");
            return false;
        }
        if (!$filename) {
            throw new \Exception("[ must input filename ]");
            return false;
        }
        try {
            if ($this->phpVersion) {
                $phpExcel = new PHPExcelMaker8($this->columnAutoSize);
            } else {
                $phpExcel = new PHPExcelMaker7($this->columnAutoSize);
            }
            $phpExcel->exportExcel($keys, $title, $data, $filename, $returnFile, $output);
        } catch (\Exception $e) {
            echo 'Message: ' . $e->getMessage();
            return false;
        }
    }
}

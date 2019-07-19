<?php

namespace Core\Report;


use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Border;

/**
 * Author:Robert Tsang
 *
 * Class export
 * @package Core\Report
 */
class Export
{

    /**
     * Author:Robert Tsang
     *
     * @var array
     */
    protected $outputs = [];

    /**
     * Author:Robert Tsang
     *
     * @var string
     */
    protected $real_name = '';

    /**
     * Author:Robert Tsang
     *
     * @var string
     */
    protected $title = '';

    /**
     * Author:Robert Tsang
     *
     * @var string
     */
    protected $department = '';

    /**
     * Author:Robert Tsang
     *
     * @var
     */
    protected $tpl;

    /**
     * Author:Robert Tsang
     *
     * @var
     */
    protected $debug;

    /**
     *
     */
    const BORDER_COLOR_RGB = '000000';


    /**
     * Export constructor.
     * @param string $tpl
     * @param array $outputs
     * @param bool $debug
     */
    public function __construct(string $tpl, array $outputs, bool $debug = false)
    {
        $this->outputs = $outputs;
        $this->debug = $debug;
        $this->tpl = $tpl;
    }

    /**
     * Author:Robert Tsang
     *
     * @param string $title
     * @return $this
     */
    public function setTitle(string $title)
    {
        $this->title = $title;
        return $this;
    }

    /**
     * Author:Robert Tsang
     *
     * @param string $realName
     * @return $this
     */
    public function setRealName(string $realName)
    {
        $this->real_name = $realName;
        return $this;
    }

    /**
     * Author:Robert Tsang
     *
     * @param string $department
     * @return $this
     */
    public function setDepartment(string $department)
    {
        $this->department = $department;
        return $this;
    }

    /**
     * Author:Robert Tsang
     *
     * @return array
     */
    protected function getStyle(): array
    {
        return [
            'borders' => [
                'outline' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['argb' => self::BORDER_COLOR_RGB],
                ],
            ],
        ];
    }

    /**
     * Author:Robert Tsang
     *
     */
    public function markdown()
    {
        if ($this->outputs) {
            echo PHP_EOL;
            echo '#'.$this->title.'#'.PHP_EOL;
            echo PHP_EOL;
            echo '| 姓名 | 部门 | 填表日期  |'.PHP_EOL;
            echo '|------|------|-----------|'.PHP_EOL;
            echo '|'.$this->real_name.'| '.$this->department.'  |'.date('Y/m/d').' |'.PHP_EOL;
            echo PHP_EOL;
            echo '##本周工作##'.PHP_EOL;
            echo PHP_EOL;
            foreach ($this->outputs as $output) {
                echo '- '.$output.PHP_EOL;
            }
            echo PHP_EOL;
        }
    }

    /**
     * Author:Robert Tsang
     *
     * @param $dist
     * @return mixed
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function save(string $dist)
    {
        $reader = IOFactory::createReaderForFile($this->tpl);
        $spreadsheet = $reader->load($this->tpl);
        $sheet = $spreadsheet->getActiveSheet();
        //设置工作日志表头
        $sheet->getCell("A1")->setValue($this->title);
        $sheet->getCell("F2")->setValue(date('Y/m/d'));
        $sheet->getCell("D2")->setValue($this->department);
        $sheet->getCell("B2")->setValue($this->real_name);
        //写工作日志
        $line = 4;
        foreach ($this->outputs as $i => $output) {
            $sheet->getRowDimension($line)->setRowHeight(40);
            $sheet->getStyle("A$line")->applyFromArray($this->getStyle());
            $sheet->getStyle("E$line")->applyFromArray($this->getStyle());
            $sheet->mergeCells("B$line:D$line")->getStyle("B$line:D$line")->applyFromArray($this->getStyle());
            $sheet->mergeCells("F$line:H$line")->getStyle("F$line:H$line")->applyFromArray($this->getStyle());
            $sheet->getCell("A$line")->setValue($i + 1);
            $sheet->getCell("E$line")->setValue($i + 1);
            $sheet->getCell("B$line")->setValue($output);
            $line++;
        }
        $sheet->getRowDimension($line)->setRowHeight(40);
        $sheet->getStyle("A$line")->applyFromArray($this->getStyle());
        $sheet->mergeCells("B$line:H$line")->getStyle("B$line:H$line")->applyFromArray($this->getStyle());
        $sheet->getCell("A$line")->setValue('工作总结：');
        $sheet->getCell("B$line")->setValue('这个得自己写');
        $sheet->mergeCells("B$line:H$line")->getStyle("B$line:D$line")->applyFromArray($this->getStyle());
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        if (file_exists($dist)) {
            unlink($dist);
        }
        return $writer->save($dist);
    }


}
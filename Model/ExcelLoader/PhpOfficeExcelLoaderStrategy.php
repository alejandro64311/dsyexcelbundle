<?php 

namespace dsarhoya\DSYXLSBundle\Model\ExcelLoader;

use dsarhoya\DSYXLSBundle\XLS\Factory;
use Generator;

class PhpOfficeExcelLoaderStrategy implements ExcelLoaderInterface {

    private $excelObj;
    private $activeSheet;
    private $start_at_index;
    private $column_configurations;
    public $loadDone = false;

    public function __construct(string $filePath)
    {   
        $this->excelObj = Factory::createPHPExcelObject($filePath);
        $this->activeSheet = $this->excelObj->getSheet(0);
        $this->start_at_index = 1;
        $this->column_configurations = [];
    }

    /**
     * Return max row allowed
     *
     * @return integer|null
     */
    public function getMaxRows() : ?int
    {
        return $this->activeSheet->getHighestRow();
    }
    
    /**
     * Set row to start
     *
     * @param integer $index
     * @return void
     */
    public function setStartAtIndex(int $index)
    {
        $this->start_at_index = $index;
    }

    /**
     * Set Configurations Column
     *
     * @param array $configs
     * @return void
     */
    public function setColunnsConfigurations(array $configs)
    {
        $this->column_configurations = $configs;
    }

    /**
     * Get RoW
     *
     * @return Generator|null
     */
    public function getRowIterator() : ?Generator
    {
        $row = [];
        foreach ($this->column_configurations as $column_index => $configuration) {
            $cell = $this->activeSheet->getCellByColumnAndRow($column_index + 1, $this->start_at_index);

            $value = $this->getValue($cell, $configuration);
            if ($configuration->getIndex() && empty($value)) {
                $this->loadDone = true;
                return; // la columna indice no tiene valor
            }
            $row[] = $value;
        }

        $this->start_at_index++;

        yield $row;

    }
    
    private function getValue($cell, $configuration)
    {
        $value = $cell->getCalculatedValue();
        if ($configuration->getTrim()) {
            $value = trim($value);
        }

        /* @var $cell \PHPExcel_Cell */
        // if (\PHPExcel_Shared_Date::isDateTime($cell)) {
        if (\PhpOffice\PhpSpreadsheet\Shared\Date::isDateTime($cell)) {
            return \PHPExcel_Style_NumberFormat::toFormattedString($value, $configuration->getFormat());
        }

        return $value;
    }
}
<?php 

namespace dsarhoya\DSYXLSBundle\Model\ExcelLoader;

use Box\Spout\Common\Entity\Cell;
use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
use Generator;

class SpoutExcelLoaderStrategy implements ExcelLoaderInterface 
{
    private $reader;
    private $start_at_index;
    private $column_configurations;
    public $loadDone = false;

    public function __construct(string $filePath)
    {   
        $this->reader = ReaderEntityFactory::createReaderFromFile($filePath);
        $this->reader->open($filePath);
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
        return null;
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
        foreach ($this->reader->getSheetIterator() as $sheet) {
            $i = 0;
            foreach ($sheet->getRowIterator() as $row) {
                $row_data = [];
                $i++;
                if ($i >= $this->start_at_index) {
                    $cells = $row->getCells();
                    foreach ($this->column_configurations as $column_index => $configuration) {
                        /** @var Cell $cell */
                        $cell = $cells[$column_index];
                        if ($configuration->getIndex() && $cell->isEmpty()) {
                            $this->loadDone = true;
                            $this->reader->close();
                            return null;
                        }
                        
                        $row_data[] = $cell ? $cell->getValue() : null;
                        
                    }
                    $this->start_at_index++;
                    yield $row_data;
                    
                }

            }
            $this->loadDone = true;
            return null;
        }
    }    
}
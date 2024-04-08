<?php 

namespace dsarhoya\DSYXLSBundle\Model\ExcelLoader;

use Generator;

interface ExcelLoaderInterface {

    /**
     * Return max row allowed
     *
     * @return integer|null
     */
    public function getMaxRows() : ?int;
    
    /**
     * Set row to start
     *
     * @param integer $index
     * @return void
     */
    public function setStartAtIndex(int $index);

    /**
     * Set Configurations Column
     *
     * @param array $configs
     * @return void
     */
    public function setColunnsConfigurations(array $configs);

    /**
     * Get RoW
     *
     * @return Generator|null
     */
    public function getRowIterator() : ?Generator;

}


<?php

namespace dsarhoya\DSYXLSBundle\XLS\Legacy\Validators;

use dsarhoya\DSYXLSBundle\XLS\Legacy\Validators\AbstractCellValidator;
use PHPExcel_Cell_DataType;

/**
 * Description of NumberCellValidator
 *
 * @author matias
 */
class NumberCellValidator extends AbstractCellValidator{
    public function validate($cell) {
        $value = ($cell->getDataType()=== PHPExcel_Cell_DataType::TYPE_FORMULA)?$cell->getCalculatedValue():$cell->getValue();
        if(is_numeric($value)) return true;
        $this->addError('Cell value must be number');
        return false;
    }
}

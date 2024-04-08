<?php

namespace dsarhoya\DSYXLSBundle\XLS\Legacy\Validators;

use dsarhoya\DSYXLSBundle\XLS\Legacy\Validators\AbstractCellValidator;

/**
 * Description of NumberCellValidator
 *
 * @author matias
 */
class StringCellValidator extends AbstractCellValidator{
    public function validate($cell) {
        $value = $cell->getValue();
        if(is_string($value)) return true;
        $this->addError('Cell value must be string');
        return false;
    }
}

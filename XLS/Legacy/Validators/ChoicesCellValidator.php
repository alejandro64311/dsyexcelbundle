<?php

namespace dsarhoya\DSYXLSBundle\XLS\Legacy\Validators;

use dsarhoya\DSYXLSBundle\XLS\Legacy\Validators\AbstractCellValidator;

/**
 * Description of NumberCellValidator
 *
 * @author matias
 */
class ChoicesCellValidator extends AbstractCellValidator{
    
    private $choices;
    
    public function __construct($options = null) {
        $this->choices = array();
        if(isset($options['choices']) && is_array($options['choices'])) $this->choices = $options['choices'];
    }
    
    public function validate($cell) {
        $value = $cell->getValue();
        if(in_array($value, $this->choices)) return true;
        $this->addError('Invalid option');
        return false;
    }
}

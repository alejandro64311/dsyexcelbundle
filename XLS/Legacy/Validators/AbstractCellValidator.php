<?php

namespace dsarhoya\DSYXLSBundle\XLS\Legacy\Validators;

/**
 * Description of CellValidator
 *
 * @author matias
 */
abstract class AbstractCellValidator {
    CONST VALIDATOR_TYPE_NUMBER = 'number';
    CONST VALIDATOR_TYPE_STRING = 'string';
    CONST VALIDATOR_TYPE_DATE = 'date';
    CONST VALIDATOR_TYPE_CHOICES = 'choice';
    
    private $errors;
    private $custom_error_message;
    
    public abstract function validate($cell);
    
    public function clearErrors(){
        $this->errors = array();
    }
    public function getErrors(){
        return implode(', ', $this->errors);
    }
    public function addError($error){
        if(!is_string($error)) throw new \Exception('Errors must be strings');
        $this->errors[] = $this->custom_error_message ? $this->custom_error_message : $error;
    }
    public function setCustomErrorMessage($custom_message){
        if(!is_string($custom_message)) throw new \Exception('Error messages must be strings');
        $this->custom_error_message = $custom_message;
    }
}

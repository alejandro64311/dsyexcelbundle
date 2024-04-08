<?php

namespace dsarhoya\DSYXLSBundle\XLS\Legacy;

use dsarhoya\DSYXLSBundle\XLS\Validators\AbstractCellValidator;

/**
 * Description of ColumnConfiguration
 *
 * @author matias
 */
class ColumnConfiguration {
    
    CONST CELL_TYPE_DATE = 'date';
    
    private $name;
    private $type;
    private $validators;
    private $errors;
    private $format = 'dd-mm-yyyy';
    private $allows_empty = false;
    private $index = false;
    
    public function __construct() {
        $this->validators   = array();
    }
    
    public function setName($name){
        $this->name = $name;
        return $this;
    }
    
    public function setType($type){
        $this->type = $type;
        return $this;
    }
    
    public function setFormat($format){
        $this->format = $format;
        return $this;
    }
    
    public function setIndex($index){
        $this->index = $index;
        return $this;
    }
    
    public function setAllowsEmpty($allows_empty){
        $this->allows_empty = $allows_empty;
        return $this;
    }
    
    public function getName(){
        return $this->name;
    }
    
    public function addValidatior(AbstractCellValidator $validator){
        $this->validators[] = $validator;
        return $this;
    }
    
    public function validate($cell){
        $validates = true;
        $this->errors = array();
        if($this->allows_empty && !$this->index){
            if(!$cell->getValue()) return true;
            if($cell->getValue() == '') return true;
        }
        
        foreach ($this->validators as $validator) {
            $validator->clearErrors();
            if(!$validator->validate($cell)){
                $this->errors[$this->name][] = $validator->getErrors();
                $validates = false;
            }
        }
        return $validates;
    }
    
    public function value($cell){
        if($this->type == self::CELL_TYPE_DATE){
            
            if (\PHPExcel_Shared_Date::isDateTime($cell) ) {
                $date= \PHPExcel_Style_NumberFormat::toFormattedString($cell->getCalculatedValue(), $this->format);
            }else{
                $date = $cell->getCalculatedValue();
            }

            $date = str_replace(" ","",$date);

            return $date;
        }
        return $cell->getCalculatedValue();
    }
    
    public function getErrors(){
        return $this->errors;
    }
}
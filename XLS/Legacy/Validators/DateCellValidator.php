<?php

namespace dsarhoya\DSYXLSBundle\XLS\Legacy\Validators;

use dsarhoya\DSYXLSBundle\XLS\Legacy\Validators\AbstractCellValidator;

/**
 * Description of NumberCellValidator
 *
 * @author matias
 */
class DateCellValidator extends AbstractCellValidator{
    
    private $format = 'dd-mm-yyyy';
    
    public function validate($cell) {
        
        if (\PHPExcel_Shared_Date::isDateTime($cell) ) {
            $date= \PHPExcel_Style_NumberFormat::toFormattedString($cell->getValue(), $this->format);
        }else{
            $date = $cell->getValue();
        }
        
        $date = str_replace("/","-",$date);
        $date = str_replace(" ","",$date);
        
        if($date=='' || preg_match('/^\d{2}.\d{2}.\d{4}$/', $date)==0 || count(split("-",$date))!=3 ){
            $this->addError('Cell value must be date');
            return false;
        }
        return true;
    }
}

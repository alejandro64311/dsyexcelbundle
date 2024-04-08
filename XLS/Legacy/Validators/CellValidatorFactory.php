<?php

namespace dsarhoya\DSYXLSBundle\XLS\Legacy\Validators;

use dsarhoya\DSYXLSBundle\XLS\Legacy\Validators\AbstractCellValidator;

/**
 * Description of CellValidatorFactory
 *
 * @author matias
 */
class CellValidatorFactory {
    public static function validatorWithType($type, $options=array()){
        
        $custom_error_message = null;
        if(isset($options['error_message'])) $custom_error_message = $options['error_message'];
        
        //Esto se podría hacer sin el switch, usando los nombres de las clases
        switch ($type) {
            case AbstractCellValidator::VALIDATOR_TYPE_NUMBER:{
                $validator = new NumberCellValidator();
                if($custom_error_message) $validator->setCustomErrorMessage ($custom_error_message);
                return $validator;
            }
            break;
            case AbstractCellValidator::VALIDATOR_TYPE_STRING:{
                $validator = new StringCellValidator();
                if($custom_error_message) $validator->setCustomErrorMessage ($custom_error_message);
                return $validator;
            }
            break;
            case AbstractCellValidator::VALIDATOR_TYPE_DATE:{
                $validator = new DateCellValidator();
                if($custom_error_message) $validator->setCustomErrorMessage ($custom_error_message);
                return $validator;
            }
            break;
            case AbstractCellValidator::VALIDATOR_TYPE_CHOICES:{
                $validator = new ChoicesCellValidator($options);
                if($custom_error_message) $validator->setCustomErrorMessage ($custom_error_message);
                return $validator;
            }
            break;

            default:
                throw new \Exception('Unsupported cell validator type: '.$type);
            break;
        }
    }
}
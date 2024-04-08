<?php

namespace dsarhoya\DSYXLSBundle\XLS\Legacy;

use dsarhoya\DSYXLSBundle\XLS\Legacy\ColumnConfiguration;
use dsarhoya\DSYXLSBundle\XLS\Legacy\Validators\CellValidatorFactory;

/**
 * Description of ExcelLoader
 *
 * @author matias
 */
class ExcelLoader {
    private $phpexcel;
    private $columns_configuration;
    private $has_headers;
    private $stops_on_error;
    private $errors;
    private $loader_index;
    private $max_rows;
    
    public function __construct($phpexcel) {
        $this->columns_configuration    = array();
        $this->has_headers              = true;
        $this->stops_on_error           = true;
        $this->phpexcel                 = $phpexcel;
        $this->loader_index             = 0;
        $this->max_rows                 = 0;
    }
    
    /*
     * Agrega una columna al loader
     * 
     * Los parametros son:
     * 'name': Obligatorio. El nombre de la columna. Se usa para dar el resultado y los errores.
     * 'type': Opcional. Se usa para agregar validadores automáticamente.
     * 'format': Opcional. Formato de salida. Por ahora solo funciona con el tipo date y solo si la celda del excel es tipo date.
     * 'validators': Opcional. Validadores para la columna. Debe ser un arreglo que tenga la llave 'type' o ser un objeto que extienda de AbstractCellValidator.
     * 'allows_empty': Opcional, por defecto false. Permite que una columna sea vacía
     * 
     * @param $configuration
     * @return ExcelLoader
     */
    public function addCell($configuration){
        
        if(!is_array($configuration)) throw new \Exception('Unsupported column configuration type');
        
        $column_configuration = new ColumnConfiguration();
        
        if(!isset($configuration['name'])) throw new \Exception('configuration needs name');
        $column_configuration->setName($configuration['name']);
        
        if(isset($configuration['type'])) {
            
            $options = null;
            //esto funciona porque el tipo de la columna se llama igual que el tipo del validador
            if($configuration['type']==ColumnConfiguration::CELL_TYPE_DATE){
                if(isset($configuration['format'])) $column_configuration->setFormat ($configuration['format']);
            }
            
            if(isset($configuration['error_message'])) $options['error_message'] = $configuration['error_message'];
            if(isset($configuration['choices'])) $options['choices'] = $configuration['choices'];
            
            $column_configuration->setType($configuration['type']);
            $column_configuration->addValidatior(CellValidatorFactory::validatorWithType($configuration['type'], $options));
        }
        
        if(isset($configuration['allows_empty'])) $column_configuration->setAllowsEmpty($configuration['allows_empty']);
        
        if(isset($configuration['validators'])) {
            foreach ($configuration['validators'] as $validator) {
                
                if(is_subclass_of($validator, 'dsarhoya\DSYXLSBundle\XLS\Validators\AbstractCellValidator')){
                    $column_configuration->addValidatior($validator);
                } elseif (!is_array($validator)) { //tiene que heredar del abstract 
                    throw new \Exception('Unsupported validator for column '.$configuration['name']);
                } else{ //array
                    
                    if(!isset($validator['type'])) throw new Exception('Must specify validator type');
                    
                    $options = null;
                    
                    if(isset($validator['error_message'])) $options['error_message'] = $validator['error_message'];
                    
                    $column_configuration->addValidatior(CellValidatorFactory::validatorWithType($validator['type'], $options));
                }
            }
        }
        
        $this->columns_configuration[] = $column_configuration;
        
        return $this;
    }
    
    public function setHasHeaders($has_headers){
        if(is_bool($has_headers)) $this->has_headers = $has_headers;
        return $this;
    }
    
    public function setStopsOnError($stops_on_error){
        if(is_bool($stops_on_error)) $this->stops_on_error = $stops_on_error;
        return $this;
    }
    
    public function setLoaderIndex($index){
        if(is_numeric($index)) $this->loader_index = $index;
        return $this;
    }
    
    public function setMaxRows($max_rows){
        if(!is_int($max_rows)) throw new \Exception('Max rows must be a number');
        if($max_rows < 0) return $this;
        $this->max_rows = $max_rows;
        return $this;
    }
    
    public function load($xls){
        
        $excelObj = $this->phpexcel->createPHPExcelObject($xls);
        $activeSheet = $excelObj->getSheet(0);
        
        foreach ($this->columns_configuration as $index => $configuration) {
            $configuration->setIndex($index == $this->loader_index ? true : false);
        }
        
        if($this->max_rows){
            if($activeSheet->getHighestRow() - ($this->has_headers ? 1 : 0) > $this->max_rows){
                $this->errors[] = sprintf('Máximo de %d filas exedido (%d filas).', $this->max_rows, $activeSheet->getHighestRow() - ($this->has_headers ? 1 : 0));
                return false;
            }
        }
        
        $i = $this->has_headers ? 2 : 1;
        $this->errors = array(); //reset
        $result = array();
        
        do{
            $continue = true;
            $valid_row = true;
            $row_info = array();
            
            foreach ($this->columns_configuration as $column_index => $configuration) {
                
                $cell = $activeSheet->getCellByColumnAndRow($column_index, $i);
                if($configuration->validate($cell)){
                    $row_info[$configuration->getName()] = $configuration->value($cell);
                    $row_info['xls_line_number'] = $i;
                }else{
                    $this->errors['line '.$i] = $configuration->getErrors();
                    if($this->stops_on_error) $continue = false;
                    $valid_row = false;
                }
            }
            
            if($valid_row) $result[] = $row_info;
            $i++;
            
        }while(!is_null($activeSheet->getCellByColumnAndRow($this->loader_index, $i)->getCalculatedValue()) && $continue);
        
        return $result;
    }
    
    public function isValid(){
        return count($this->errors) ? false : true;
    }
    
    public function getErrors(){
        return $this->errors;
    }
}
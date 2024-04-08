<?php

namespace dsarhoya\DSYXLSBundle\XLS\Loader;

use dsarhoya\DSYXLSBundle\Model\ExcelLoader\ExcelLoaderInterface;
use dsarhoya\DSYXLSBundle\Model\ExcelLoader\PhpOfficeExcelLoaderStrategy;
use dsarhoya\DSYXLSBundle\Model\ExcelLoader\SpoutExcelLoaderStrategy;
use Symfony\Component\Validator\Constraints\NotBlank;
use Symfony\Component\Validator\Constraints\NotNull;
use Symfony\Component\Validator\Validator\TraceableValidator;

/**
 * Description of ExcelLoader.
 *
 * @author matias
 */
class ExcelLoader
{
    const COUNT_MAX_VALUE = 20000;

    // private $phpexcel;
    private $validator;
    private $columns_configuration;
    private $has_headers;
    private $stops_on_error;
    private $errors;
    private $loader_index;
    private $max_rows;
    private $starts_at_index;

    /**
     * @var ExcelLoaderInterface
     */
    private $loaderStrategy;

    public function __construct($validator)
    {
        $this->columns_configuration = [];
        $this->has_headers = true;
        $this->stops_on_error = true;
        // $this->phpexcel                 = $phpexcel;
        $this->validator = $validator;
        $this->loader_index = 0;
        $this->max_rows = 0;
    }

    /**
     * Agrega una columna al loader.
     *
     * Los parametros son:
     * 'name': Obligatorio. El nombre de la columna. Se usa para dar el resultado y los errores.
     * 'type': Opcional. Se usa para agregar validadores automáticamente.
     * 'format': Opcional. Formato de salida. Por ahora solo funciona con el tipo date y solo si la celda del excel es tipo date.
     * 'validators': Opcional. Validadores para la columna. Debe ser un arreglo que tenga la llave 'type' o ser un objeto que extienda de AbstractCellValidator.
     * 'allows_empty': Opcional, por defecto false. Permite que una columna sea vacía
     * 'trim': Opcional, por defecto false. elimina espacios vacios de ambos lados del valor de la celda.
     *
     * @param $configuration
     *
     * @return ExcelLoader
     */
    public function addCell(ColumnOptions $configuration)
    {
        $column_configuration = new ColumnConfiguration($configuration->constraints);

        $column_configuration->setName($configuration->name);

        if ($configuration->format) {
            $column_configuration->setFormat($configuration->format);
        }
        $column_configuration->setTrim($configuration->trim);

        $this->columns_configuration[] = $column_configuration;

        return $this;
    }

    public function setHasHeaders($has_headers)
    {
        if (is_bool($has_headers)) {
            $this->has_headers = $has_headers;
        }

        return $this;
    }

    public function setStopsOnError($stops_on_error)
    {
        if (is_bool($stops_on_error)) {
            $this->stops_on_error = $stops_on_error;
        }

        return $this;
    }

    public function setLoaderIndex($index)
    {
        if (is_numeric($index)) {
            $this->loader_index = $index;
        }

        return $this;
    }

    public function setMaxRows($max_rows)
    {
        if (!is_int($max_rows)) {
            throw new \Exception('Max rows must be a number');
        }
        if ($max_rows < 0) {
            return $this;
        }
        $this->max_rows = $max_rows;

        return $this;
    }

    public function startsAtIndex($index)
    {
        if (is_int($index)) {
            $this->starts_at_index = $index;
        }

        return $this;
    }

    public function getMaxRows($xls)
    {
        $strategy = $this->strategyForFile($xls);

        if (null !== $maxRows = $strategy->getMaxRows()) {
            return $maxRows;
        }

        $maxValue = self::COUNT_MAX_VALUE;
        $count = 0;

        foreach ($strategy->getRowIterator() as $key => $value) {
            ++$count;

            if ($count > $maxValue) {
                return $maxValue;
            }
        }

        return $count;
    }

    public function load($xls)
    {
        $this->loaderStrategy = $this->strategyForFile($xls);

        return $this->doLoad();
    }

    /**
     * @return ExcelLoaderInterface
     */
    private function strategyForFile($xls)
    {
        
        $parts = explode('.', $xls);
        $extension = array_pop($parts);

        return 'xls' === $extension ? new PhpOfficeExcelLoaderStrategy($xls) : new SpoutExcelLoaderStrategy($xls);
    }

    private function doLoad()
    {
        foreach ($this->columns_configuration as $index => $configuration) {
            $configuration->setIndex($index == $this->loader_index ? true : false);
        }

        if ($this->max_rows) {
            if (null !== $this->loaderStrategy->getMaxRows()) {
                if ($this->loaderStrategy->getMaxRows() - ($this->has_headers ? 1 : 0) > $this->max_rows) {
                    $this->errors[] = sprintf('Máximo de %d filas exedido (%d filas).', $this->max_rows, $this->loaderStrategy->getMaxRows() - ($this->has_headers ? 1 : 0));

                    return false;
                }
            }
        }

        $i = $this->has_headers ? 2 : 1;

        if (is_int($this->starts_at_index)) {
            $i = $this->starts_at_index;
        }

        $this->loaderStrategy->setStartAtIndex($i);

        $this->loaderStrategy->setColunnsConfigurations($this->columns_configuration);

        $this->errors = [];

        do {
            $continue = true;
            $valid_row = true;
            $row_info = [];

            foreach ($this->loaderStrategy->getRowIterator() as $row) {
                foreach ($row as $key => $cellValue) {
                    if ($this->validate($cellValue, $this->validator, $this->columns_configuration[$key])) {
                        $row_info[$this->columns_configuration[$key]->getName()] = $cellValue;
                        $row_info['xls_line_number'] = $i;
                    } else {
                        $this->errors['line '.$i] = $this->columns_configuration[$key]->getErrors();
                        if ($this->stops_on_error) {
                            $continue = false;
                        }
                        $valid_row = false;
                    }
                }
                ++$i;
                if ($valid_row) {
                    yield $row_info;
                }
            }
        } while (!$this->loaderStrategy->loadDone && $continue);
    }

    public function isValid()
    {
        return count($this->errors) ? false : true;
    }

    public function getErrors()
    {
        return $this->errors;
    }

    public function validate($cellValue, $validator, $configuration)
    {
        $isValid = true;
        $errors = [];

        /** @var ColumnConfiguration $configuration */
        $validations = $configuration->getValidations();

        if ($configuration->getIndex()) {
            $notNull = new NotNull();
            $notNull->message = 'El índice no puede ser nulo';
            $notBlank = new NotBlank();
            $notBlank->message = 'El índice no puede ser nulo';
            $validations[] = $notBlank;
            $validations[] = $notNull;
        }
        $errorList = $validator->validate($cellValue, $validations);

        if (count($errorList) > 0) {
            $errors[$configuration->getName()][] = $errorList[0]->getMessage();
            $configuration->setErrors($errors);
            $isValid = false;
        }

        if ($validator instanceof TraceableValidator) {
            $validator->reset();
        }

        return $isValid;
    }
}

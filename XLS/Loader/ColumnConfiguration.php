<?php

namespace dsarhoya\DSYXLSBundle\XLS\Loader;

use dsarhoya\DSYXLSBundle\XLS\Legacy\Validators\AbstractCellValidator;
use Symfony\Component\Validator\Constraints\Collection;
use Symfony\Component\Validator\Constraints\NotNull;
use Symfony\Component\Validator\Constraints\NotBlank;

/**
 * Description of ColumnConfiguration.
 *
 * @author matias
 */
class ColumnConfiguration
{
    const CELL_TYPE_DATE = 'date';

    private $name;
    private $type;
    private $validators;

    /**
     * @var bool
     */
    private $trim = false;

    /**
     * @var Collection
     */
    private $validation_collection;
    private $errors;
    private $format = 'dd-mm-yyyy';
    private $allows_empty = false;
    private $index = false;

    public function __construct(array $validation_contraints = [])
    {
        $this->validation_collection = $validation_contraints;
    }

    public function setName($name)
    {
        $this->name = $name;

        return $this;
    }

    /**
     * @param string|null $format
     *
     * @return ColumnConfiguration
     */
    public function setType($type)
    {
        $this->type = $type;

        return $this;
    }

    /**
     * @param string|null $format
     *
     * @return ColumnConfiguration
     */
    public function setFormat($format)
    {
        $this->format = $format;

        return $this;
    }

    public function setTrim($trim)
    {
        $this->trim = $trim;

        return $this;
    }

    public function setIndex($index)
    {
        $this->index = $index;

        return $this;
    }

    public function setAllowsEmpty($allows_empty)
    {
        $this->allows_empty = $allows_empty;

        return $this;
    }

    public function getName()
    {
        return $this->name;
    }

    public function addValidatior(AbstractCellValidator $validator)
    {
        $this->validators[] = $validator;

        return $this;
    }

    public function getErrors()
    {
        return $this->errors;
    }

    public function setErrors($errors)
    {
        $this->errors = $errors;

        return $this;
    }

    public function getTrim()
    {
        return $this->trim;
    }

    public function getFormat()
    {
        return $this->format;
    }

    public function getIndex()
    {
        return $this->index;
    }

    public function getValidations()
    {
        return $this->validation_collection;
    }
}

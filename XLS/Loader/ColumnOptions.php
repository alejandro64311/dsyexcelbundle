<?php

namespace dsarhoya\DSYXLSBundle\XLS\Loader;

use Symfony\Component\OptionsResolver\OptionsResolver;

/**
 * Description of ColumnOptions
 *
 * @author matias
 */
class ColumnOptions extends OptionsResolver
{
    public $options;

    public function __construct(array $options = array())
    {
        $this->configureOptions($this);

        $this->options = $this->resolve($options);
    }

    public function __get($name)
    {
        return $this->options[$name];
    }

    public function configureOptions(OptionsResolver $resolver)
    {
        $resolver->setDefaults(array(
            'format' => null,
            'constraints' => [],
            'trim' => false
        ));

        $resolver->setRequired('name');
        $resolver->setAllowedTypes('format', array('string', 'null'));
        $resolver->setAllowedTypes('trim', array('boolean'));
        $resolver->setAllowedTypes('constraints', array('array'));
    }
}

<?php

namespace dsarhoya\DSYXLSBundle\Interfaces;

interface ExcelReadinessInterface {
    public static function getColumnsHeaders($options = null);
    public function getRowData($options = null);
}
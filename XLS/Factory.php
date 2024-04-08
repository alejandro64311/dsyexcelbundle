<?php

namespace dsarhoya\DSYXLSBundle\XLS;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;
use PhpOffice\PhpSpreadsheet\Reader\IReader;
use Symfony\Component\HttpFoundation\StreamedResponse;

class Factory
{

    /**
     * @param string $filename
     * @return IReader
     */
    public function createPHPExcelReaderForFile($filename)
    {
        return IOFactory::createReaderForFile($filename);
    }

    /**
     * @param string $filename
     * @return Spreadsheet
     */
    public function createPHPExcelObject($filename = null)
    {
        return (null === $filename) ? new Spreadsheet() : IOFactory::load($filename);
    }

    /**
     * @param Spreadsheet $spreadsheet
     * @param string $type
     * @return IWriter
     */
    public function createWriter(Spreadsheet $spreadsheet, $type = 'Xls')
    {
        return \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, $type);
    }

    /**
     * @param IWriter $writer
     * @param integer $status
     * @param array $headers
     * @return StreamedResponse
     */
    public function createStreamedResponse(IWriter $writer, $status = 200, $headers = array())
    {
        return new StreamedResponse(
            function () use ($writer) {
                $writer->save('php://output');
            },
            $status,
            $headers
        );
    }
}

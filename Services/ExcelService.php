<?php

namespace dsarhoya\DSYXLSBundle\Services;

use dsarhoya\DSYXLSBundle\XLS\Legacy\ExcelLoader as LegacyLoader;
use dsarhoya\DSYXLSBundle\XLS\Loader\ExcelLoader;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xls\Worksheet;
use Symfony\Component\PropertyAccess\PropertyAccess;
/**
 * Description of UserKeysService
 *
 * @author matias
 */
class ExcelService {
    
    private $em;
    private $phpExcel;
    private $validator;

    public function __construct($entityManager, $phpExcel, $validator) {
        $this->em               = $entityManager;
        $this->phpExcel         = $phpExcel;
        $this->validator        = $validator;
    }
    
    public function getExcelFileFromCollection($headers, $collection,$description = null, $rowDataOptions = null, $sheetOptions = null){
        $phpExcelObject = $this->createExcelFile($description);

        if ($sheetOptions !== null && isset($sheetOptions['sheetName'])) {
            $phpExcelObject->removeSheetByIndex(0);
            $activeSheet = $this->createSheet($phpExcelObject, $sheetOptions['sheetName']);
            // $activeSheet = $phpExcelObject->getSheetByName($sheetOptions['sheetName']);
        } else {
            $phpExcelObject->setActiveSheetIndex(0);
            $activeSheet = $phpExcelObject->getActiveSheet();
        }

        $this->setDataOnSheet($activeSheet, $headers, $collection, $rowDataOptions);
        
        return $phpExcelObject;//return the excel file itself
    }
    
    CONST FILE_TYPE_EXCEL = 'Xlsx';
    CONST FILE_TYPE_CSV = 'Csv';
    
    public function returnExcelFile($excelFile,$excelFileName = 'book', $type = self::FILE_TYPE_EXCEL){
        $phpexcel = $this->phpExcel;
        // create the writer
        $writer = $phpexcel->createWriter($excelFile, $type);
        
        // create the response
        $response = $phpexcel->createStreamedResponse($writer);
        // adding headers
        $response->headers->set('Content-Type', 'text/vnd.ms-excel; charset=utf-8');
        $response->headers->set('Content-Disposition', 'attachment;filename='.$excelFileName.'.'.$this->typeToExtension($type));
        $response->headers->set('Pragma', 'public');
        $response->headers->set('Cache-Control', 'maxage=1');

        return $response;
    }
    
    private function typeToExtension($type){
        switch ($type) {
            case self::FILE_TYPE_CSV:
                return 'csv';
        }
        return 'xlsx';
    }
    
    public function getLegacyExcelLoader(){
        return new LegacyLoader($this->phpExcel);
    }
    
    public function getExcelLoader(){
        return new ExcelLoader($this->validator);
    }

    public function createExcelFile($description)
    {
        $phpExcelObject = $this->phpExcel->createPHPExcelObject();
        $phpExcelObject
                ->getProperties()->setCreator("dsarhoya SpA")
                ->setDescription($description);

        return $phpExcelObject;
    }

    public function createSheet(Spreadsheet $excelFile, $sheetName = 'Worksheet')
    {
        $newSheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($excelFile, $sheetName);
        return $excelFile->addSheet($newSheet);
    }

    public function setDataOnSheet($sheet, $headers, $collection, $rowDataOptions = null)
    {
        //HEADERS
        $colHeadres = 1;
        foreach($headers as $header){
            $sheet->setCellValueByColumnAndRow($colHeadres++,1,$header);
        }
        //every property
        $activeRow = 2;
        $activeCol = 1;
        
        foreach($collection as $row => $object){
            
            if(is_array($object)){
                $rowData = $object['row_data'];
            }else{
                $rowData = $object->getRowData(array_replace(
                        !is_array($rowDataOptions) ? [] : $rowDataOptions, 
                        array('row_number'=>$row+2)));
            }
            foreach($rowData as $data){
                $sheet->setCellValueByColumnAndRow($activeCol++,$activeRow,$data);
            }
            $activeRow++;
            $activeCol = 1;
        }

        $borderStyleArray = 
                array(
                    'borders' => array(
                        'allborders' => array(
                                'style' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                            )
                        )
        );
        $auxRange = 'A1:'.\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex(count($headers)).''.($activeRow-1);

        $this->setStyle($sheet, $auxRange, $borderStyleArray);

        $boldColHeadersStyleArray = 
            array(
                'font' => array(
                    'bold' => true,
                )
            );
        $auxRange = 'A1:'.\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex(count($headers)).'1';
        $this->setStyle($sheet, $auxRange, $boldColHeadersStyleArray);
    }

    public function setStyle($sheet, $cells, $style)
    {
        $sheet->getStyle($cells)->applyFromArray($style);
    }
}

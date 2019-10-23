<?php

namespace Xola\ReportWriterBundle;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * Factory for \PhpOffice\PhpSpreadsheet\Spreadsheet objects, StreamedResponse, and \PhpOffice\PhpSpreadsheet\Writer\IWriter.
 */
class SpreadsheetFactory
{
    private $phpExcelIO;

    public function __construct($phpExcelIO = IOFactory::class)
    {
        $this->phpExcelIO = $phpExcelIO;
    }

    /**
     * Creates an empty \PhpOffice\PhpSpreadsheet\Spreadsheet Object if the filename is empty, otherwise loads the file into the object.
     *
     * @param string $filename
     *
     * @return Spreadsheet
     */
    public function createPhpSpreadsheetObject($filename = null)
    {
        if (null == $filename) {
            $phpExcelObject = new Spreadsheet();
            return $phpExcelObject;
        }
        return call_user_func(array($this->phpExcelIO, 'load'), $filename);
    }

    /**
     * Create a writer given the \PhpOffice\PhpSpreadsheet\SpreadsheetObject and the type,
     *   the type could be one of \PhpOffice\PhpSpreadsheet\IOFactory::$_autoResolveClasses
     *
     * @param Spreadsheet $phpExcelObject
     * @param string    $type
     *
     *
     * @return \PhpOffice\PhpSpreadsheet\Writer\IWriter
     */
    public function createWriter(Spreadsheet $phpExcelObject, $type = 'Xls')
    {
        return call_user_func(array($this->phpExcelIO, 'createWriter'), $phpExcelObject, $type);
    }
}
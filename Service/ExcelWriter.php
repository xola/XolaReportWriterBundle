<?php

namespace Xola\ReportWriterBundle\Service;

use Symfony\Component\DependencyInjection\Container;
use Psr\Log\LoggerInterface;
use Xola\ReportWriterBundle\PHPExcelFactory;

class ExcelWriter extends AbstractWriter
{
    private $phpexcelService;
    /* @var \PHPExcel $handle */
    private $handle;

    public function __construct(Container $container, LoggerInterface $logger, PHPExcelFactory $phpExcel)
    {
        $this->logger = $logger;
        $this->phpexcelService = $phpExcel;
    }

    public function setup()
    {
        $this->handle = $this->phpexcelService->createPHPExcelObject();

        /*
        $excel->getProperties()->setCreator("Creator")
            ->setTitle("Title: Office 2005 XLSX Test Document")
            ->setSubject("Subject: Office 2005 XLSX Test Document")
            ->setDescription("Description: Test document for Office 2005 XLSX, generated using PHP classes.")
            ->setKeywords("office 2005 openxml php")
            ->setCategory("Test result file");
        */
    }

    public function setWorksheet($index, $title)
    {
        $this->handle->createSheet($index);
        $this->handle->setActiveSheetIndex($index);
        $this->setSheetTitle($title);
    }

    public function setSheetTitle($title)
    {
        $this->handle->getActiveSheet()->setTitle($title);
    }

    public function prepare($cacheFile, $sortedHeaders)
    {
        $worksheet = $this->handle->getActiveSheet();

        $column = 'A';
        foreach ($sortedHeaders as $idx => $header) {
            $cell = $column . '1';
            if (!is_array($header)) {
                $worksheet->setCellValue($cell, $header);
                // Assumption that all headers are multi-row, so we merge the rows of non-multirow headers
                $worksheet->mergeCells($column . '1:' . $column . '2');
                //$worksheet->getStyle($cell)->getFill()->setFillType(\PHPExcel_Style_Fill::FILL_SOLID);
                //$worksheet->getStyle($cell)->getFill()->getStartColor()->setRGB('EFEFEF');
                $column++;
            } else {
                // This is a multi-row header, the first row consists of one value merged across several cells and the
                // second row contains the "children".

                // Write the first row of the header
                $arrKeys = array_keys($header);
                $headerName = reset($arrKeys);
                $worksheet->setCellValue($cell, $headerName);

                // Figure out how many cells across to merge
                $mergeLength = count($header[$headerName]) - 1;
                $mergeDestination = $this->incrementColumn($column, $mergeLength);
                $worksheet->mergeCells($column . '1:' . $mergeDestination . '1');

                //$worksheet->getStyle($column . '1:' . $mergeDestination . '2')->getFill()->setFillType(\PHPExcel_Style_Fill::FILL_SOLID);
                //$worksheet->getStyle($column . '1:' . $mergeDestination . '2')->getFill()->getStartColor()->setRGB('EFEFEF');

                // Now write the children's values onto the second row
                foreach ($header[$headerName] as $subHeaderName) {
                    $worksheet->setCellValue($column . '2', $subHeaderName);
                    $column++;
                }
            }
        }

        // TODO: Track memory usage
        $file = new \SplFileObject($cacheFile);
        $rowIdx = 0;
        $excelRow = [];
        while (!$file->eof()) {
            $dataRow = json_decode($file->current(), true);
            if (!is_array($dataRow)) {
                // Invalid json data -- don't process this row
                continue;
            }

            foreach ($sortedHeaders as $idx => $header) {
                if (!is_array($header)) {
                    $excelRow[$rowIdx][] = (isset($dataRow[$header])) ? $dataRow[$header] : null;
                } else {
                    // Multi-row header, so we need to set all values
                    $nestedHeaderName = array_keys($header)[0];
                    $nestedHeaders = $header[$nestedHeaderName];

                    foreach ($nestedHeaders as $nestedHeader) {
                        $excelRow[$rowIdx][] = (isset($dataRow[$nestedHeaderName][$nestedHeader])) ? $dataRow[$nestedHeaderName][$nestedHeader] : null;
                    }
                }
            }

            $rowIdx++;
            $file->next();
        }

        // Write the entire worksheet. TODO: Maybe too much data. perf.
        $worksheet->fromArray($excelRow, null, 'A3');

        // Freeze the headers
        $worksheet->freezePane('A3');

        $file = null; // Get rid of the file handle that SplFileObject has on cache file
        unlink($cacheFile);
    }

    public function finalize($filename)
    {
        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $this->handle->setActiveSheetIndex(0);

        // Write the file to disk
        $writer = $this->phpexcelService->createWriter($this->handle, 'Excel2007');
        $writer->save($filename);
    }

    /**
     * Increment columns. A + 1 = B, A + 2 = C etc...
     *
     * @param     $str
     * @param int $count
     *
     * @return mixed
     */
    private function incrementColumn($str, $count = 1)
    {
        for ($i = 0; $i < $count; $i++) {
            $str++;
        }

        return $str;
    }
}
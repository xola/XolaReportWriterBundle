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
    private $currentRow = 1;

    public function __construct(Container $container, LoggerInterface $logger, PHPExcelFactory $phpExcel)
    {
        $this->logger = $logger;
        $this->phpexcelService = $phpExcel;
    }

    /**
     * Initialize the excel writer
     *
     * @param string $author The author/creator of this file
     * @param string $title  The title of this file
     */
    public function setup($author = '', $title = '')
    {
        $this->handle = $this->phpexcelService->createPHPExcelObject();
        $this->handle->getActiveSheet()->getPageSetup()->setOrientation(\PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);

        $this->handle->getProperties()
            ->setCreator($author)
            ->setTitle($title);
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

    public function writeHeaders($sortedHeaders)
    {
        $worksheet = $this->handle->getActiveSheet();

        $initRow = $this->currentRow;
        $column = 'A';
        foreach ($sortedHeaders as $idx => $header) {
            $cell = $column . $initRow;
            if (!is_array($header)) {
                $worksheet->setCellValue($cell, $header);
                // Assumption that all headers are multi-row, so we merge the rows of non-multirow headers
                $worksheet->mergeCells($column . $initRow . ':' . $column . ($initRow+1));
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
                $worksheet->mergeCells($column . $initRow . ':' . $mergeDestination . $initRow);

                // Now write the children's values onto the second row
                foreach ($header[$headerName] as $subHeaderName) {
                    $worksheet->setCellValue($column . ($initRow + 1), $subHeaderName);
                    $column++;
                }
            }
        }

        $this->currentRow += 2;
    }

    public function prepare($cacheFile, $sortedHeaders)
    {
        $this->writeHeaders($sortedHeaders);

        $file = new \SplFileObject($cacheFile);
        while (!$file->eof()) {
            $dataRow = json_decode($file->current(), true);
            $this->writeRow($dataRow, $sortedHeaders);
            $file->next();
        }

        $file = null; // Get rid of the file handle that SplFileObject has on cache file
        unlink($cacheFile);
    }

    public function writeRows($rows, $headers)
    {
        foreach ($rows as $row) {
            $this->writeRow($row, $headers);
        }
    }

    public function writeRow($dataRow, $headers)
    {
        if (!is_array($dataRow)) {
            // Invalid data -- don't process this row
            return;
        }

        $rowIdx = $this->currentRow;
        $excelRow = [];
        foreach ($headers as $idx => $header) {
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

        $this->handle->getActiveSheet()->fromArray($excelRow, null, 'A' . $rowIdx);
        $this->currentRow++;
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
     * Write ad-hoc set of rows without any dependence on headers
     *
     * @param array  $lines
     * @param int    $row
     * @param string $column
     */
    public function writeRawRows(array $lines, $column = 'A', $row = null)
    {
        if (is_null($row)) {
            $row = $this->currentRow;
        }
        $this->handle->getActiveSheet()->fromArray($lines, null, $column . $row);
        $this->currentRow += count($lines);
    }

    public function freezePane($cell = '')
    {
        if (empty($cell)) {
            $cell = 'A2';
        }
        $this->handle->getActiveSheet()->freezePane($cell);
    }

    public function addPageBreak($cell = '')
    {
        if (empty($cell)) {
            $cell = 'A' . ($this->currentRow - 2);
        }
        $this->handle->getActiveSheet()->setBreak($cell, \PHPExcel_Worksheet::BREAK_ROW);
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
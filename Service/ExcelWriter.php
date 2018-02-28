<?php

namespace Xola\ReportWriterBundle\Service;

use Psr\Log\LoggerInterface;
use Xola\ReportWriterBundle\PHPExcelFactory;

class ExcelWriter extends AbstractWriter
{
    private $phpexcel;
    /* @var \PHPExcel $handle */
    private $handle;
    private $currentRow = 1;

    public function __construct(LoggerInterface $logger, PHPExcelFactory $phpExcel)
    {
        parent::__construct($logger);
        $this->phpexcel = $phpExcel;
        \PHPExcel_Shared_Font::setAutoSizeMethod(\PHPExcel_Shared_Font::AUTOSIZE_METHOD_EXACT);
    }

    /**
     * Initialize the excel writer
     *
     * @param string $filepath
     */
    public function setup($filepath)
    {
        $this->handle = $this->phpexcel->createPHPExcelObject();
        $this->handle->getActiveSheet()->getPageSetup()->setOrientation(\PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
        $this->filepath = $filepath;
    }

    /**
     * Set meta properties for the excel writer
     *
     * @param string $author The author/creator of this file
     * @param string $title  The title of this file
     */
    public function setProperties($author = '', $title = '')
    {
        $this->handle->getProperties()
            ->setCreator($author)
            ->setTitle($title);
    }

    /**
     * Create a new worksheet, and set it as the active one
     *
     * @param int    $index Location at which to create the sheet (NULL for last)
     * @param string $title The title of the sheet
     * @throws \PHPExcel_Exception
     */
    public function setWorksheet($index, $title)
    {
        $this->handle->createSheet($index);
        $this->resetCurrentRow(1);
        $this->handle->setActiveSheetIndex($index);
        $this->setSheetTitle($title);
    }

    /**
     * Set the title for the current active worksheet
     *
     * @param string $title
     */
    public function setSheetTitle($title)
    {
        $this->handle->getActiveSheet()->setTitle($title);
    }

    /**
     * Write the headers (nested or otherwise) to the current active worksheet
     *
     * @param $headers
     * @param $initRow
     *
     * @throws \PHPExcel_Exception
     */
    public function writeHeaders($headers, $initRow = null)
    {
        $worksheet = $this->handle->getActiveSheet();
        $hasMultiRowHeaders = $this->hasMultiRowHeaders($headers);

        if ($initRow) {
            $worksheet->insertNewRowBefore($initRow, 1);
        } else {
            $initRow = $this->currentRow;
        }

        $column = 'A';
        foreach ($headers as $idx => $header) {
            $cell = $column . $initRow;
            if (!is_array($header)) {
                $worksheet->setCellValue($cell, $header);
                $worksheet->getColumnDimension($column)->setAutoSize(true);
                if ($hasMultiRowHeaders) {
                    // These set of headers contain multi-row headers. So this cell needs to be merged with cell in the
                    // row below it.
                    $worksheet->mergeCells($column . $initRow . ':' . $column . ($initRow + 1));
                }
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
                    $worksheet->getColumnDimension($column)->setAutoSize(true);
                    $column++;
                }
            }

            // Mark headers as bold
            $worksheet->getStyle($cell)->getFont()->setBold(true);
        }

        $worksheet->calculateColumnWidths(true);

        $this->currentRow = $initRow + (($hasMultiRowHeaders) ? 2 : 1);
    }

    /**
     * Check if the given array have multi-row headers
     *
     * @param $headers
     *
     * @return bool True if it has multi-row headers
     */
    private function hasMultiRowHeaders($headers)
    {
        foreach ($headers as $idx => $header) {
            if (is_array($header)) {
                // Multi-row header
                return true;
            }
        }

        return false;
    }

    /**
     * Utility method that will data from the cached file per row and write it
     *
     * @param string $cacheFile     Filename where the fetched data can be cached from
     * @param array  $sortedHeaders Headers to write sorted in the order you want them
     * @param bool   $freezeHeaders True if you want to freeze headers (default: false)
     */
    public function prepare($cacheFile, $sortedHeaders, $freezeHeaders = false)
    {
        $this->writeHeaders($sortedHeaders);
        if($freezeHeaders) {
            $this->freezePanes();
        }

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

    public function writeRow($dataRow, $headers = [])
    {
        if (empty($headers)) {
            // No headers to manage. Just write this array of data directly
            $this->writeArray($dataRow);
            return;
        }

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

        $this->writeArrays($excelRow);
    }

    /**
     * Write one or more rows starting at the given row and column
     *
     * @param array  $lines
     */
    private function writeArrays(array $lines)
    {
        $startCell = 'A' . $this->currentRow;
        $this->handle->getActiveSheet()->fromArray($lines, null, $startCell, true);
        $this->currentRow += count($lines);
    }

    /**
     * Write a single row of data
     *
     * @param array $row A single row of data
     */
    private function writeArray(array $row)
    {
        $startCell = 'A' . $this->currentRow;
        $this->handle->getActiveSheet()->fromArray([$row], null, $startCell, true);
        $this->currentRow++;
    }

    /**
     * Freeze panes at the given location so they stay fixed upon scroll
     *
     * @param string $cell
     * @throws \PHPExcel_Exception
     */
    public function freezePanes($cell = '')
    {
        if (empty($cell)) {
            $cell = 'A3';
        }
        $this->handle->getActiveSheet()->freezePane($cell);
    }

    /**
     * Add a horizonal (row) page break for print layout
     *
     * @param string $cell
     * @throws \PHPExcel_Exception
     */
    public function addHorizontalPageBreak($cell = '')
    {
        if (empty($cell)) {
            $cell = 'A' . ($this->currentRow - 2);
        }
        $this->handle->getActiveSheet()->setBreak($cell, \PHPExcel_Worksheet::BREAK_ROW);
    }

    /**
     * Save the current data into an .xlsx file
     *
     * @throws \PHPExcel_Exception
     */
    public function finalize()
    {
        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $this->handle->setActiveSheetIndex(0);

        // Write the file to disk
        $writer = $this->phpexcel->createWriter($this->handle, 'Excel2007');
        $writer->save($this->filepath);
    }

    public function resetCurrentRow($pos)
    {
        $this->currentRow = $pos;
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

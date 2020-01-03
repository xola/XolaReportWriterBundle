<?php

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PHPUnit\Framework\TestCase;
use Psr\Log\LoggerInterface;
use Xola\ReportWriterBundle\Service\ExcelWriter;
use Xola\ReportWriterBundle\SpreadsheetFactory;

class ExcelWriterTest extends TestCase
{
    /* @var Spreadsheet| */
    private $spreadsheet;

    public function setUp()
    {
        $this->spreadsheet = $this->getMockBuilder(Spreadsheet::class)->disableOriginalConstructor()->getMock();
    }

    public function buildService($params = [])
    {
        $defaults = ['logger' => $this->createMock(LoggerInterface::class)];
        if(!isset($params['phpExcel'])) {
            $defaults['phpExcel'] = $this->getSpreadsheetFactoryMock();
        }

        $params = array_merge($defaults, $params);

        return new ExcelWriter($params['logger'], $params['phpExcel']);
    }

    public function testShouldInitializePHPExcelObject()
    {
        // Setup all the mocks
        $pageSetupMock = $this->getMockBuilder(PageSetup::class)->disableOriginalConstructor()->getMock();
        $pageSetupMock->expects($this->once())->method('setOrientation')->with('landscape');
        $worksheetMock = $this->getMockBuilder(Worksheet::class)->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('getPageSetup')->willReturn($pageSetupMock);
        $this->spreadsheet->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $this->buildService()->setup("filename.xlsx");
    }

    public function testShouldSetPropertiesForExcelFile()
    {
        $author = 'foo';
        $title = 'bar';

        // Setup all the mocks
        $propertiesMock = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Document\Properties')->disableOriginalConstructor()->getMock();
        $propertiesMock->expects($this->once())->method('setCreator')->with($author)->willReturn($propertiesMock);
        $propertiesMock->expects($this->once())->method('setTitle')->with($title)->willReturn($propertiesMock);
        $this->spreadsheet->expects($this->once())->method('getProperties')->willReturn($propertiesMock);

        $this->buildService()->setProperties($author, $title);
    }

    public function testShouldSetCurrentWorksheet()
    {
        $index = 0;
        $title = 'Hello World';

        // Setup all the mocks
        $worksheetMock = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('setTitle')->with($title);
        $this->spreadsheet->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);
        $this->spreadsheet->expects($this->once())->method('createSheet')->with($index);
        $this->spreadsheet->expects($this->once())->method('setActiveSheetIndex')->with($index);

        $service = $this->buildService();

        $service->setWorksheet($index, $title);
        $this->assertAttributeEquals(1, 'currentRow', $service, 'Row should be reset to 1 when switching active sheets');
    }

    public function testShouldWriteSingleRowHeaders()
    {
        $columnDimensionMock = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Worksheet\ColumnDimension')->disableOriginalConstructor()->getMock();
        $columnDimensionMock->expects($this->exactly(4))->method('setAutoSize')->with(true);

        $phpExcelStyleMock2 = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Style\Font')->disableOriginalConstructor()->getMock();
        $phpExcelStyleMock2->expects($this->exactly(4))->method('setBold')->with(true);
        $phpExcelStyleMock = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Style\Style')->disableOriginalConstructor()->getMock();
        $phpExcelStyleMock->expects($this->exactly(4))->method('getFont')->willReturn($phpExcelStyleMock2);
        $worksheetMock = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->exactly(4))->method('getStyle')->willReturn($phpExcelStyleMock);
        $worksheetMock->expects($this->exactly(4))->method('setCellValue')->withConsecutive(
            ['A1', 'Alpha'], ['B1', 'Bravo'], ['C1', 'Gamma'], ['D1', 'Delta']
        );
        $worksheetMock->expects($this->exactly(4))->method('getColumnDimension')->withConsecutive(
            ['A'], ['B'], ['C'], ['D']
        )->willReturn($columnDimensionMock);
        $this->spreadsheet->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $headers = ['Alpha', 'Bravo', 'Gamma', 'Delta'];
        $this->buildService()->writeHeaders($headers);
    }

    public function testShouldWriteNestedHeaders()
    {
        $columnDimensionMock = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Worksheet\ColumnDimension')->disableOriginalConstructor()->getMock();
        $columnDimensionMock->expects($this->exactly(6))->method('setAutoSize')->with(true);

        $phpExcelStyleMock2 = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Style\Font')->disableOriginalConstructor()->getMock();
        $phpExcelStyleMock2->expects($this->exactly(5))->method('setBold')->with(true);
        $phpExcelStyleMock = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Style\Style')->disableOriginalConstructor()->getMock();
        $phpExcelStyleMock->expects($this->exactly(5))->method('getFont')->willReturn($phpExcelStyleMock2);

        $worksheetMock = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->exactly(5))->method('getStyle')->willReturn($phpExcelStyleMock);
        $worksheetMock->expects($this->exactly(7))->method('setCellValue')->withConsecutive(
            ['A1', 'Alpha'], ['B1', 'Bravo'], ['C1', 'Gamma'], ['D1', 'Delta'], ['E1', 'Echo'],
            ['E2', 'Foxtrot'], ['F2', 'Hotel']
        );
        $worksheetMock->expects($this->exactly(5))->method('mergeCells')->withConsecutive(
            ['A1:A2'], ['B1:B2'], ['C1:C2'], ['D1:D2'], ['E1:F1']
        );
        $worksheetMock->expects($this->exactly(6))->method('getColumnDimension')->withConsecutive(
            ['A'], ['B'], ['C'], ['D'], ['E'], ['F']
        )->willReturn($columnDimensionMock);
        $this->spreadsheet->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $headers = [0 => 'Alpha', 1 => 'Bravo', 2 => 'Gamma', 3 => 'Delta', 4 => ['Echo' => ['Foxtrot', 'Hotel']]];
        $this->buildService()->writeHeaders($headers);
    }

    public function testShouldWriteNonNestedData()
    {
        $input = [
            'Alpha' => 'This is A',
            'Bravo' => 'This is B',
            'Gamma' => 'This is G, we skipped C',
            'Delta' => 'This is Dee',
            'Charlie' => 'This header does not exist, so value should not show'
        ];
        $expected = [1 => [
            'This is A', 'This is B', 'This is G, we skipped C', 'This is Dee'
        ]];

        $worksheetMock = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('fromArray')->with($expected, null, 'A1');
        $this->spreadsheet->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $headers = [0 => 'Alpha', 1 => 'Bravo', 2 => 'Gamma', 3 => 'Delta'];
        $this->buildService()->writeRow($input, $headers);
    }

    public function testShouldWriteNestedRowData()
    {
        $input = [
            'Alpha' => 'This is A',
            'Bravo' => 'This is B',
            'Gamma' => 'This is G, we skipped C',
            'Delta' => 'This is Dee',
            'Echo' => [
                'Foxtrot' => 'Fancy, F',
                'Hotel' => 'Etch'
            ],
            'India' => 'This header does not exist, so value should not show'
        ];
        $expected = [1 => [
            'This is A', 'This is B', 'This is G, we skipped C', 'This is Dee', 'Fancy, F', 'Etch'
        ]];

        $worksheetMock = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('fromArray')->with($expected, null, 'A1');
        $this->spreadsheet->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $headers = [0 => 'Alpha', 1 => 'Bravo', 2 => 'Gamma', 3 => 'Delta', 4 => ['Echo' => ['Foxtrot', 'Hotel']]];
        $this->buildService()->writeRow($input, $headers);
    }

    public function testShouldWriteArrayAsDataStartingFromTheFirstCell()
    {
        $lines = ['Lorem', 'Ipsum', 'Dolor', 'Sit', 'Amet'];
        $worksheetMock = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('fromArray')->with([$lines], null, 'A1');
        $this->spreadsheet->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $this->buildService()->writeRow($lines);
    }

    public function testShouldFreezePanes()
    {
        $worksheetMock = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('freezePane')->with('A3');
        $this->spreadsheet->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $this->buildService()->freezePanes();
    }

    public function testShouldFreezePanesAtTheGivenLocation()
    {
        $location = 'X3';

        $worksheetMock = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('freezePane')->with($location);
        $this->spreadsheet->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $this->buildService()->freezePanes($location);
    }

    public function testShouldAddRowPageBreak()
    {
        $worksheetMock = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('setBreak')->with('A3', 1);
        $this->spreadsheet->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $service = $this->buildService();
        $service->resetCurrentRow(5);
        $service->addHorizontalPageBreak();
    }

    public function testShouldSaveFileInExcel2007Format()
    {
        $filename = 'export.xlsx';

        $this->spreadsheet->expects($this->once())->method('setActiveSheetIndex')->with(0);
        $writerMock = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Writer\IWriter')->disableOriginalConstructor()->getMock();
        $writerMock->expects($this->once())->method('save')->with($filename);

        // Setup all the mocks
        $pageSetupMock = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup')->disableOriginalConstructor()->getMock();
        $worksheetMock = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('getPageSetup')->willReturn($pageSetupMock);
        $this->spreadsheet->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $phpExcel = $this->getSpreadsheetFactoryMock();
        $phpExcel->expects($this->once())->method('createWriter')
            ->with($this->spreadsheet, 'Xlsx')
            ->willReturn($writerMock);

        $service = $this->buildService(['phpExcel' => $phpExcel]);
        $service->setup($filename);

        $service->finalize();
    }

    private function getSpreadsheetFactoryMock()
    {
        $factory = $this->getMockBuilder(SpreadsheetFactory::class)->disableOriginalConstructor()->getMock();
        $factory->expects($this->any())->method('createPhpSpreadsheetObject')->willReturn($this->spreadsheet);

        return $factory;
    }
}

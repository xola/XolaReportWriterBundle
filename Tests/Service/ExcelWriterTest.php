<?php

use PhpOffice\PhpSpreadsheet\Document\Properties;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Style\Style;
use PhpOffice\PhpSpreadsheet\Worksheet\ColumnDimension;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use Psr\Log\LoggerInterface;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use Xola\ReportWriterBundle\Service\ExcelWriter;

class ExcelWriterTest extends PHPUnit_Framework_TestCase
{
    /* @var \PHPUnit_Framework_MockObject_MockObject */
    private $spreadsheetMock;

    public function setUp()
    {
        $this->spreadsheetMock = $this->getMockBuilder(Spreadsheet::class)->disableOriginalConstructor()->getMock();
    }

    public function buildService($params = [])
    {
        $defaults = ['logger' => $this->getMockBuilder(LoggerInterface::class)->disableOriginalConstructor()->getMock()];
        $params = array_merge($defaults, $params);

        return new ExcelWriter($params['logger'], $this->spreadsheetMock);
    }

    public function testShouldSetPropertiesForExcelFile()
    {
        $author = 'foo';
        $title = 'bar';

        // Setup all the mocks
        $propertiesMock = $this->getMockBuilder(Properties::class)->disableOriginalConstructor()->getMock();
        $propertiesMock->expects($this->once())->method('setCreator')->with($author)->willReturn($propertiesMock);
        $propertiesMock->expects($this->once())->method('setTitle')->with($title)->willReturn($propertiesMock);
        $this->spreadsheetMock->expects($this->once())->method('getProperties')->willReturn($propertiesMock);

        $service = $this->buildService();

        $service->setProperties($author, $title);
    }

    public function testShouldSetCurrentWorksheet()
    {
        $index = 0;
        $title = 'Hello World';

        // Setup all the mocks
        $worksheetMock = $this->getMockBuilder(Worksheet::class)->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('setTitle')->with($title);
        $this->spreadsheetMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);
        $this->spreadsheetMock->expects($this->once())->method('createSheet')->with($index);
        $this->spreadsheetMock->expects($this->once())->method('setActiveSheetIndex')->with($index);

        $service = $this->buildService();

        $service->setWorksheet($index, $title);
        $this->assertAttributeEquals(1, 'currentRow', $service, 'Row should be reset to 1 when switching active sheets');
    }

    public function testShouldWriteSingleRowHeaders()
    {
        $columnDimensionMock = $this->getMockBuilder(ColumnDimension::class)->disableOriginalConstructor()->getMock();
        $columnDimensionMock->expects($this->exactly(4))->method('setAutoSize')->with(true);

        $phpExcelStyleMock2 = $this->getMockBuilder(Font::class)->disableOriginalConstructor()->getMock();
        $phpExcelStyleMock2->expects($this->exactly(4))->method('setBold')->with(true);
        $phpExcelStyleMock = $this->getMockBuilder(Style::class)->disableOriginalConstructor()->getMock();
        $phpExcelStyleMock->expects($this->exactly(4))->method('getFont')->willReturn($phpExcelStyleMock2);
        $worksheetMock = $this->getMockBuilder(Worksheet::class)->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->exactly(4))->method('getStyle')->willReturn($phpExcelStyleMock);
        $worksheetMock->expects($this->exactly(4))->method('setCellValue')->withConsecutive(
            ['A1', 'Alpha'], ['B1', 'Bravo'], ['C1', 'Gamma'], ['D1', 'Delta']
        );
        $worksheetMock->expects($this->exactly(4))->method('getColumnDimension')->withConsecutive(
            ['A'], ['B'], ['C'], ['D']
        )->willReturn($columnDimensionMock);
        $this->spreadsheetMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $headers = ['Alpha', 'Bravo', 'Gamma', 'Delta'];
        $this->buildService()->writeHeaders($headers);
    }

    public function testShouldWriteNestedHeaders()
    {
        $columnDimensionMock = $this->getMockBuilder(ColumnDimension::class)->disableOriginalConstructor()->getMock();
        $columnDimensionMock->expects($this->exactly(6))->method('setAutoSize')->with(true);

        $phpExcelStyleMock2 = $this->getMockBuilder(Font::class)->disableOriginalConstructor()->getMock();
        $phpExcelStyleMock2->expects($this->exactly(5))->method('setBold')->with(true);
        $phpExcelStyleMock = $this->getMockBuilder(Style::class)->disableOriginalConstructor()->getMock();
        $phpExcelStyleMock->expects($this->exactly(5))->method('getFont')->willReturn($phpExcelStyleMock2);

        $worksheetMock = $this->getMockBuilder(Worksheet::class)->disableOriginalConstructor()->getMock();
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
        $this->spreadsheetMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

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

        $worksheetMock = $this->getMockBuilder(Worksheet::class)->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('fromArray')->with($expected, null, 'A1');
        $this->spreadsheetMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

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

        $worksheetMock = $this->getMockBuilder(Worksheet::class)->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('fromArray')->with($expected, null, 'A1');
        $this->spreadsheetMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $headers = [0 => 'Alpha', 1 => 'Bravo', 2 => 'Gamma', 3 => 'Delta', 4 => ['Echo' => ['Foxtrot', 'Hotel']]];
        $this->buildService()->writeRow($input, $headers);
    }

    public function testShouldWriteArrayAsDataStartingFromTheFirstCell()
    {
        $lines = ['Lorem', 'Ipsum', 'Dolor', 'Sit', 'Amet'];
        $worksheetMock = $this->getMockBuilder(Worksheet::class)->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('fromArray')->with([$lines], null, 'A1');
        $this->spreadsheetMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $this->buildService()->writeRow($lines);
    }

    public function testShouldFreezePanes()
    {
        $worksheetMock = $this->getMockBuilder(Worksheet::class)->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('freezePane')->with('A3');
        $this->spreadsheetMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $this->buildService()->freezePanes();
    }

    public function testShouldFreezePanesAtTheGivenLocation()
    {
        $location = 'X3';

        $worksheetMock = $this->getMockBuilder(Worksheet::class)->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('freezePane')->with($location);
        $this->spreadsheetMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $this->buildService()->freezePanes($location);
    }

    public function testShouldAddRowPageBreak()
    {
        $worksheetMock = $this->getMockBuilder(Worksheet::class)->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('setBreak')->with('A3', 1);
        $this->spreadsheetMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $service = $this->buildService();
        $service->resetCurrentRow(5);
        $service->addHorizontalPageBreak();
    }
}

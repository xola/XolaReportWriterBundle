<?php

use Psr\Log\LoggerInterface;
use Xola\ReportWriterBundle\PHPExcelFactory;
use Xola\ReportWriterBundle\Service\ExcelWriter;

class ExcelWriterTest extends PHPUnit_Framework_TestCase
{
    /* @var \PHPUnit_Framework_MockObject_MockObject */
    private $phpExcelHandleMock;

    public function setUp()
    {
        $this->phpExcelHandleMock = $this->getMockBuilder('\PHPExcel')->disableOriginalConstructor()->getMock();
    }

    public function buildService($params = [])
    {
        $defaults = ['logger' => $this->getMockBuilder(LoggerInterface::class)->disableOriginalConstructor()->getMock()];
        if(!isset($params['phpExcel'])) {
            $defaults['phpExcel'] = $this->getMockBuilder(PHPExcelFactory::class)->disableOriginalConstructor()->getMock();
        }

        $params = array_merge($defaults, $params);

        $service = new ExcelWriter($params['logger'], $params['phpExcel']);
        $this->setHandle($service);
        return $service;
    }

    public function testShouldInitializePHPExcelObject()
    {
        // Setup all the mocks
        $pageSetupMock = $this->getMockBuilder('\PHPExcel_Worksheet_PageSetup')->disableOriginalConstructor()->getMock();
        $pageSetupMock->expects($this->once())->method('setOrientation')->with('landscape');
        $worksheetMock = $this->getMockBuilder('\PHPExcel_Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('getPageSetup')->willReturn($pageSetupMock);
        $this->phpExcelHandleMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);
        $phpExcel = $this->getMockBuilder(PHPExcelFactory::class)->disableOriginalConstructor()->getMock();
        $phpExcel->expects($this->once())->method('createPHPExcelObject')->willReturn($this->phpExcelHandleMock);

        $this->buildService(['phpExcel' => $phpExcel])->setup("filename.xlsx");
    }

    public function testShouldSetPropertiesForExcelFile()
    {
        $author = 'foo';
        $title = 'bar';

        // Setup all the mocks
        $propertiesMock = $this->getMockBuilder('\PHPExcel_DocumentProperties')->disableOriginalConstructor()->getMock();
        $propertiesMock->expects($this->once())->method('setCreator')->with($author)->willReturn($propertiesMock);
        $propertiesMock->expects($this->once())->method('setTitle')->with($title)->willReturn($propertiesMock);
        $this->phpExcelHandleMock->expects($this->once())->method('getProperties')->willReturn($propertiesMock);

        $service = $this->buildService();

        $service->setProperties($author, $title);
    }

    public function testShouldSetCurrentWorksheet()
    {
        $index = 0;
        $title = 'Hello World';

        // Setup all the mocks
        $worksheetMock = $this->getMockBuilder('\PHPExcel_Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('setTitle')->with($title);
        $this->phpExcelHandleMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);
        $this->phpExcelHandleMock->expects($this->once())->method('createSheet')->with($index);
        $this->phpExcelHandleMock->expects($this->once())->method('setActiveSheetIndex')->with($index);

        $service = $this->buildService();

        $service->setWorksheet($index, $title);
        $this->assertAttributeEquals(1, 'currentRow', $service, 'Row should be reset to 1 when switching active sheets');
    }

    public function testShouldWriteSingleRowHeaders()
    {
        $columnDimensionMock = $this->getMockBuilder('\PHPExcel_Worksheet_ColumnDimension')->disableOriginalConstructor()->getMock();
        $columnDimensionMock->expects($this->exactly(4))->method('setAutoSize')->with(true);

        $phpExcelStyleMock2 = $this->getMockBuilder('\PHPExcel_Style_Font')->disableOriginalConstructor()->getMock();
        $phpExcelStyleMock2->expects($this->exactly(4))->method('setBold')->with(true);
        $phpExcelStyleMock = $this->getMockBuilder('\PHPExcel_Style')->disableOriginalConstructor()->getMock();
        $phpExcelStyleMock->expects($this->exactly(4))->method('getFont')->willReturn($phpExcelStyleMock2);
        $worksheetMock = $this->getMockBuilder('\PHPExcel_Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->exactly(4))->method('getStyle')->willReturn($phpExcelStyleMock);
        $worksheetMock->expects($this->exactly(4))->method('setCellValue')->withConsecutive(
            ['A1', 'Alpha'], ['B1', 'Bravo'], ['C1', 'Gamma'], ['D1', 'Delta']
        );
        $worksheetMock->expects($this->exactly(4))->method('getColumnDimension')->withConsecutive(
            ['A'], ['B'], ['C'], ['D']
        )->willReturn($columnDimensionMock);
        $this->phpExcelHandleMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $headers = ['Alpha', 'Bravo', 'Gamma', 'Delta'];
        $this->buildService()->writeHeaders($headers);
    }

    public function testShouldWriteNestedHeaders()
    {
        $columnDimensionMock = $this->getMockBuilder('\PHPExcel_Worksheet_ColumnDimension')->disableOriginalConstructor()->getMock();
        $columnDimensionMock->expects($this->exactly(6))->method('setAutoSize')->with(true);

        $phpExcelStyleMock2 = $this->getMockBuilder('\PHPExcel_Style_Font')->disableOriginalConstructor()->getMock();
        $phpExcelStyleMock2->expects($this->exactly(5))->method('setBold')->with(true);
        $phpExcelStyleMock = $this->getMockBuilder('\PHPExcel_Style')->disableOriginalConstructor()->getMock();
        $phpExcelStyleMock->expects($this->exactly(5))->method('getFont')->willReturn($phpExcelStyleMock2);

        $worksheetMock = $this->getMockBuilder('\PHPExcel_Worksheet')->disableOriginalConstructor()->getMock();
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
        $this->phpExcelHandleMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

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

        $worksheetMock = $this->getMockBuilder('\PHPExcel_Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('fromArray')->with($expected, null, 'A1');
        $this->phpExcelHandleMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

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

        $worksheetMock = $this->getMockBuilder('\PHPExcel_Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('fromArray')->with($expected, null, 'A1');
        $this->phpExcelHandleMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $headers = [0 => 'Alpha', 1 => 'Bravo', 2 => 'Gamma', 3 => 'Delta', 4 => ['Echo' => ['Foxtrot', 'Hotel']]];
        $this->buildService()->writeRow($input, $headers);
    }

    public function testShouldWriteArrayAsDataStartingFromTheFirstCell()
    {
        $lines = ['Lorem', 'Ipsum', 'Dolor', 'Sit', 'Amet'];
        $worksheetMock = $this->getMockBuilder('\PHPExcel_Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('fromArray')->with([$lines], null, 'A1');
        $this->phpExcelHandleMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $this->buildService()->writeRow($lines);
    }

    public function testShouldFreezePanes()
    {
        $worksheetMock = $this->getMockBuilder('\PHPExcel_Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('freezePane')->with('A3');
        $this->phpExcelHandleMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $this->buildService()->freezePanes();
    }

    public function testShouldFreezePanesAtTheGivenLocation()
    {
        $location = 'X3';

        $worksheetMock = $this->getMockBuilder('\PHPExcel_Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('freezePane')->with($location);
        $this->phpExcelHandleMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $this->buildService()->freezePanes($location);
    }

    public function testShouldAddRowPageBreak()
    {
        $worksheetMock = $this->getMockBuilder('\PHPExcel_Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('setBreak')->with('A3', 1);
        $this->phpExcelHandleMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $service = $this->buildService();
        $service->resetCurrentRow(5);
        $service->addHorizontalPageBreak();
    }

    public function testShouldSaveFileInExcel2007Format()
    {
        $filename = 'export.xlsx';

        $this->phpExcelHandleMock->expects($this->once())->method('setActiveSheetIndex')->with(0);
        $writerMock = $this->getMockBuilder('\PHPExcel_Writer_IWriter')->disableOriginalConstructor()->getMock();
        $writerMock->expects($this->once())->method('save')->with($filename);

        // Setup all the mocks
        $pageSetupMock = $this->getMockBuilder('\PHPExcel_Worksheet_PageSetup')->disableOriginalConstructor()->getMock();
        $worksheetMock = $this->getMockBuilder('\PHPExcel_Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('getPageSetup')->willReturn($pageSetupMock);
        $this->phpExcelHandleMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $phpExcel = $this->getMockBuilder(PHPExcelFactory::class)->disableOriginalConstructor()->getMock();
        $phpExcel->expects($this->once())->method('createPHPExcelObject')->willReturn($this->phpExcelHandleMock);
        $phpExcel->expects($this->once())->method('createWriter')
            ->with($this->phpExcelHandleMock, 'Excel2007')
            ->willReturn($writerMock);

        $service = $this->buildService(['phpExcel' => $phpExcel]);
        $service->setup($filename);

        $service->finalize();
    }

    private function setHandle($service)
    {
        $reflectionClass = new ReflectionClass(ExcelWriter::class);
        $property = $reflectionClass->getProperty('handle');
        $property->setAccessible(true);
        $property->setValue($service, $this->phpExcelHandleMock);
    }
}

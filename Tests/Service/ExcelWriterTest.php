<?php

use Xola\ReportWriterBundle\Service\AbstractWriter;
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
        $defaults = ['logger' => $this->getMockBuilder('Psr\Log\LoggerInterface')->disableOriginalConstructor()->getMock()];
        if(!isset($params['phpExcel'])) {
            $defaults['phpExcel'] = $this->getPHPExcelMock();
        }

        $params = array_merge($defaults, $params);

        return new ExcelWriter($params['logger'], $params['phpExcel']);
    }

    public function testShouldInitializePHPExcelObject()
    {
        $author = 'foo';
        $title = 'bar';

        // Setup all the mocks
        $pageSetupMock = $this->getMockBuilder('\PHPExcel_Worksheet_PageSetup')->disableOriginalConstructor()->getMock();
        $pageSetupMock->expects($this->once())->method('setOrientation')->with('landscape');
        $propertiesMock = $this->getMockBuilder('\PHPExcel_DocumentProperties')->disableOriginalConstructor()->getMock();
        $propertiesMock->expects($this->once())->method('setCreator')->with($author)->willReturn($propertiesMock);
        $propertiesMock->expects($this->once())->method('setTitle')->with($title)->willReturn($propertiesMock);
        $worksheetMock = $this->getMockBuilder('\PHPExcel_Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('getPageSetup')->willReturn($pageSetupMock);
        $this->phpExcelHandleMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);
        $this->phpExcelHandleMock->expects($this->once())->method('getProperties')->willReturn($propertiesMock);

        $this->buildService()->setup($author, $title);
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

        $this->buildService()->setWorksheet($index, $title);
    }

    public function testShouldWriteSingleRowHeaders()
    {
        $worksheetMock = $this->getMockBuilder('\PHPExcel_Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->exactly(4))->method('setCellValue')->withConsecutive(
            ['A1', 'Alpha'], ['B1', 'Bravo'], ['C1', 'Gamma'], ['D1', 'Delta']
        );
        $worksheetMock->expects($this->exactly(4))->method('mergeCells')->withConsecutive(
            ['A1:A2'], ['B1:B2'], ['C1:C2'], ['D1:D2']
        );
        $this->phpExcelHandleMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $headers = ['Alpha', 'Bravo', 'Gamma', 'Delta'];
        $this->buildService()->writeHeaders($headers);
    }

    public function testShouldWriteNestedHeaders()
    {
        $worksheetMock = $this->getMockBuilder('\PHPExcel_Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->exactly(7))->method('setCellValue')->withConsecutive(
            ['A1', 'Alpha'], ['B1', 'Bravo'], ['C1', 'Gamma'], ['D1', 'Delta'], ['E1', 'Echo'],
            ['E2', 'Foxtrot'], ['F2', 'Hotel']
        );
        $worksheetMock->expects($this->exactly(5))->method('mergeCells')->withConsecutive(
            ['A1:A2'], ['B1:B2'], ['C1:C2'], ['D1:D2'], ['E1:F1']
        );
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

    public function testShouldWriteRawDataStartingFromTheFirstCell()
    {
        $lines = ['Lorem', 'Ipsum', 'Dolor', 'Sit', 'Amet'];
        $worksheetMock = $this->getMockBuilder('\PHPExcel_Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('fromArray')->with($lines, null, 'A1');
        $this->phpExcelHandleMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $this->buildService()->writeRawRows($lines);
    }

    public function testShouldWriteRawDataStartingFromTheGivenCell()
    {
        $lines = ['Lorem', 'Ipsum', 'Dolor', 'Sit', 'Amet'];
        $worksheetMock = $this->getMockBuilder('\PHPExcel_Worksheet')->disableOriginalConstructor()->getMock();
        $worksheetMock->expects($this->once())->method('fromArray')->with($lines, null, 'C3');
        $this->phpExcelHandleMock->expects($this->once())->method('getActiveSheet')->willReturn($worksheetMock);

        $this->buildService()->writeRawRows($lines, 'C', 3);
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

        $phpExcel = $this->getPHPExcelMock();
        $phpExcel->expects($this->once())->method('createWriter')
            ->with($this->phpExcelHandleMock, 'Excel2007')
            ->willReturn($writerMock);

        $service = $this->buildService(['phpExcel' => $phpExcel]);

        $service->finalize($filename);
    }

    private function getPHPExcelMock()
    {
        $phpExcel = $this->getMockBuilder('Xola\ReportWriterBundle\PHPExcelFactory')->disableOriginalConstructor()->getMock();
        $phpExcel->expects($this->once())->method('createPHPExcelObject')->willReturn($this->phpExcelHandleMock);

        return $phpExcel;
    }
}
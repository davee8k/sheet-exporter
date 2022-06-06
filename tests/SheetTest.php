<?php
use SheetExporter\Sheet;

class SheetTest extends \PHPUnit\Framework\TestCase {

	public function testCreateSheet () {
		$sheet = new Sheet('<h2>test</h2>');

		$this->assertEquals('test', $sheet->getName());

		$sheet->setColHeaders(['test','test']);
		$this->assertEquals(['test','test'], $sheet->getHeaders());
		$this->assertEquals(2, $sheet->getColCount());
		$sheet->addColHeader('test');
		$this->assertEquals(['test','test','test'], $sheet->getHeaders());
		$this->assertEquals(3, $sheet->getColCount());

		$sheet->addRow(['text','text']);
		$this->assertEquals(3, $sheet->getColCount());
		$this->assertEquals([['text','text']], $sheet->getRows());
		$this->assertEquals(false, $sheet->getStyle(0));

		$sheet->setColWidth([], 100);
		$this->assertEmpty($sheet->getCols());
		$this->assertEquals(100, $sheet->getDefCol());

		$sheet->setColWidth([50,50], 100);
		$this->assertEquals([50,50], $sheet->getCols());
		$this->assertEquals(100, $sheet->getDefCol());
	}
}
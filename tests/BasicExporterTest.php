<?php
use SheetExporter\Exporter;

class BasicExporterTest extends \PHPUnit\Framework\TestCase {

	public function testEntityConversion () {
		$this->assertEquals('&lt;test&gt;&quot;&apos;&amp;hi', Exporter::xmlEntities('<test>"\'&hi'));
	}

	public function testSizeConversion () {
		$this->assertEquals(10, Exporter::convertSize(10));
		$this->assertEquals(75, Exporter::convertSize('100px'));
		$this->assertEquals(100, Exporter::convertSize('100px', 'px'));
		$this->assertEquals(100, Exporter::convertSize('100px', 'px', 'px'));
		$this->assertEquals(75, Exporter::convertSize('100px', 'px', 'pt'));
		$this->assertEquals(100, Exporter::convertSize(100, 'px', 'px'));
		$this->assertEquals(100, Exporter::convertSize('100px', false, 'px'));

		$this->assertEquals(10, Exporter::convertSize(100, 'mm', 'cm'));
		$this->assertEquals(10, Exporter::convertSize(1, 'cm', 'mm'));
	}

	public function testConvertUnitFailFirst () {
		$this->expectException('InvalidArgumentException', 'Unknown input measure unit: fail');
		Exporter::convertSize(1, 'fail', 'mm');
	}

	public function testConvertUnitFailSecond () {
		$this->expectException('InvalidArgumentException', 'Unknown output measure unit: fail');
		Exporter::convertSize(1, 'mm', 'fail');
	}

	public function testConvertBadValueFail () {
		$this->expectException('InvalidArgumentException', 'Unknown measure value: 10,,');
		Exporter::convertSize('10,,');
	}
}
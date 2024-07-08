<?php declare(strict_types=1);

use SheetExporter\Exporter;

use PHPUnit\Framework\TestCase;

class BasicExporterTest extends TestCase {

	public function testEntityConversion (): void {
		$this->assertEquals('&lt;test&gt;&quot;&apos;&amp;hi', Exporter::xmlEntities('<test>"\'&hi'));
	}

	public function testSizeConversion (): void {
		$this->assertEquals(10, Exporter::convertSize(10));
		$this->assertEquals(75, Exporter::convertSize('100px'));
		$this->assertEquals(100, Exporter::convertSize('100px', 'px'));
		$this->assertEquals(100, Exporter::convertSize('100px', 'px', 'px'));
		$this->assertEquals(75, Exporter::convertSize('100px', 'px', 'pt'));
		$this->assertEquals(100, Exporter::convertSize(100, 'px', 'px'));

		$this->assertEquals(10, Exporter::convertSize(100, 'mm', 'cm'));
		$this->assertEquals(10, Exporter::convertSize(1, 'cm', 'mm'));
	}

	public function testConvertUnitFailFirst (): void {
		$this->expectException('InvalidArgumentException');
		$this->expectExceptionMessage('Unknown input measure unit: fail');
		Exporter::convertSize(1, 'fail', 'mm');
	}

	public function testConvertUnitFailSecond (): void {
		$this->expectException('InvalidArgumentException');
		$this->expectExceptionMessage('Unknown output measure unit: fail');
		Exporter::convertSize(1, 'mm', 'fail');
	}

	public function testConvertBadValueFail (): void {
		$this->expectException('InvalidArgumentException');
		$this->expectExceptionMessage('Unknown measure value: 10,,');
		Exporter::convertSize('10,,');
	}
}

<?php declare(strict_types=1);

use SheetExporter\ExporterCsv;

class ExporterCsvTest extends \PHPUnit\Framework\TestCase {
	use BasicExporterTrait;


	public function testExportBasic (): void {
		$ex = new ExporterCsv('test');
		$this->fillSheetBasic($ex);

$txt = '"one with \' and """,
two,"two, next",
 ,three, , ,last,
 ,,
four, , ,"column, next",
three,last,
';

		$this->expectOutputString($txt);
		$ex->compile();
	}

	public function testExportComplex (): void {
		$ex = new ExporterCsv('test');
		$this->fillSheetComplex($ex);

$txt = '<h1>Nadpis</h1>,Další nadpis,
obsah,text,
a,b,c,d,e,f,g,h,i,j,k,a,b,c,d,e,f,g,h,i,j,k,a,b,c,d,e,f,g,h,i,j,k,a,b,c,d,e,
super radek, , , ,další&znak,
 , , , ,End?,
final,row,
';

		$this->expectOutputString($txt);
		$ex->compile();
	}
}

<?php
declare(strict_types=1);

use SheetExporter\Exporter;
use SheetExporter\ExporterHtml;
use SheetExporter\ExporterXlsx;
use SheetExporter\ExporterOds;
use SheetExporter\ExporterCsv;
use SheetExporter\Sheet;
use PHPUnit\Framework\Attributes\DataProvider;

trait BasicExporterTrait
{

	/**
	 *
	 * @return array<string, array{Exporter}>
	 */
	public static function dataExporters(): array
	{
		return [
			'HTML' => [new ExporterHtml('test')],
			'XLSX' => [new ExporterXlsx('test')],
			'ODS' => [new ExporterOds('test')],
			'CSV' => [new ExporterCsv('test')]
		];
	}

	/**
	 *
	 * @param Exporter $obj
	 * @param string $name
	 * @param mixed[] $args
	 * @return string
	 */
	protected function callPrivateMethod(Exporter $obj, string $name, array $args = [])
	{
		$class = new \ReflectionClass($obj);
		$method = $class->getMethod($name);
		$method->setAccessible(true);
		return $method->invokeArgs($obj, $args);
	}

	protected function fillSheetBasic(Exporter $ex): void
	{
		$ex->setDefault(['SIZE' => 20], ['BORDER' => 10], 20);

		$sheet = $ex->insertSheet('List');

		$sheet->addArray([
			['one with \' and "'],
			[['ROWS' => 3, 'VAL' => 'two'], 'two, next'],
			[['COLS' => 3, 'VAL' => 'three'], 'last'],
			[null],
			[['COLS' => 3, 'VAL' => 'four'], 'column, next'],
			[['ROWS' => 3, 'VAL' => 'three'], 'last']
		]);
	}

	protected function fillSheetComplex(Exporter $ex): void
	{
		$ex->addStyle('one', ['SIZE' => 20, 'ALIGN' => 'right'], ['WIDTH' => 5], 20);
		$ex->addStyle('two', ['COLOR' => '#fff000'], ['LEFT' => ['COLOR' => '#000fff', 'STYLE' => 'dotted', 'WIDTH' => 10]], 40);
		$ex->addStyle('three', [], ['BACKGROUND' => '#6600ff']);

		$sheet = $ex->insertSheet('List');
		$sheet->setColWidth([50, '30pt', '20mm'], 20);
		$sheet->setColHeaders(['<h1>Nadpis</h1>', 'Další nadpis']);

		$sheet->addRow(['obsah', ['VAL' => 'text', 'STYLE' => 'one']]);
		$sheet->addRow(['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'a', 'b', 'c', 'd', 'e']);
		$sheet->addRow([['COLS' => 4, 'ROWS' => 2, 'VAL' => 'super radek'], 'další&znak'], 'two');
		$sheet->addRow(['End?'], 'three');
		$sheet->addRow(['final', 'row']);
	}

	protected function fillSheetFormula(Exporter $ex): void
	{
		$sheet = $ex->insertSheet('List');

		$sheet->addRow([1, 2, 3, 4, "test"]);
		$sheet->addRow([['FORMULA' => 'SUM(A1:D1)'], ['FORMULA' => '$E$1']]);
	}

	#[DataProvider('dataExporters')]
	public function testBaseExport(Exporter $ex): void
	{
		$this->assertEquals('SheetExporter '.Exporter::VERSION, $ex->getVersion());

		$ex->insertSheet();
		$ex->insertSheet();
		$this->assertEquals(2, count($ex->getSheets()));

		$ex->addSheet(new Sheet('Test'));
		$this->assertEquals(3, count($ex->getSheets()));

		$this->expectException('InvalidArgumentException');
		$this->expectExceptionMessage('Sheet name already exists.');
		$ex->addSheet(new Sheet('Test'));
	}

	#[DataProvider('dataExporters')]
	public function testStyleExportFail(Exporter $ex): void
	{
		$ex->setDefault([], ['COLOR' => '#fff000', 'LEFT' => ['COLOR' => '#000fff']]);
		$ex->addStyle('style');
		$this->expectException('InvalidArgumentException');
		$this->expectExceptionMessage('Style mark must by small alphanumeric only.');
		$ex->addStyle('fa.il');
	}
}

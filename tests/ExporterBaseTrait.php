<?php
use SheetExporter\Exporter,
	SheetExporter\ExporterHtml,
	SheetExporter\ExporterXlsx,
	SheetExporter\ExporterOds,
	SheetExporter\Sheet;

trait ExporterBaseTrait {

	public function dataExporters () {
		return [
			'HTML' => [new ExporterHtml('test')],
			'XLSX' => [new ExporterXlsx('test')],
			'ODS' => [new ExporterOds('test')]
		];
	}

	private function fillSheets (&$ex) {
		$ex->addStyle('one', ['SIZE'=>20, 'ALIGN'=>'right'], ['WIDTH'=>5], 20);
		$ex->addStyle('two', ['COLOR'=>'#fff000'], ['LEFT'=>['COLOR'=>'#000fff', 'STYLE'=>'dotted', 'WIDTH'=>10]], 40);

		$sheet = $ex->insertSheet('List');
		$sheet->setColWidth([50, '30pt', '20mm'], 20);
		$sheet->setColHeaders(['<h1>Nadpis</h1>','Další nadpis']);

		$sheet->addRow(['obsah', ['VAL'=>'text', 'STYLE'=>'one']]);
		$sheet->addRow(['a','b','c','d','e','f','g','h','i','j','k','a','b','c','d','e','f','g','h','i','j','k','a','b','c','d','e','f','g','h','i','j','k','a','b','c','d','e']);
		$sheet->addRow([['COLS'=>4, 'ROWS'=>2, 'VAL'=>'super radek'], 'další&znak'], 'two');
		$sheet->addRow(['End?']);
		$sheet->addRow(['final','row']);
	}

	/**
	 *
	 * @dataProvider dataExporters
	 */
	public function test_base_export (Exporter $ex) {
		$this->assertEquals('SheetExporter '.Exporter::VERSION, $ex->getVersion());

		$ex->insertSheet();
		$ex->insertSheet();
		$this->assertEquals(2, sizeof($ex->getSheets()));

		$ex->addSheet(new Sheet('Test'));
		$this->assertEquals(3, sizeof($ex->getSheets()));

		$this->expectException('InvalidArgumentException', 'Sheet name already exists.');
		$ex->addSheet(new Sheet('Test'));
	}

	/**
	 *
	 * @dataProvider dataExporters
	 */
	public function test_style_export (Exporter $ex) {
		$ex->setDefault([], ['COLOR'=>'#fff000', 'LEFT'=>['COLOR'=>'#000fff']]);
		$ex->addStyle('style');
		$this->expectException('InvalidArgumentException', 'Style mark must by small alfanumeric only.');
		$ex->addStyle('fa.il');
	}
}
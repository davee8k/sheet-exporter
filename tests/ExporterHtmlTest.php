<?php declare(strict_types=1);

use SheetExporter\Exporter;
use SheetExporter\ExporterHtml;

use PHPUnit\Framework\TestCase;

class ExporterHtmlTest extends TestCase {
	use BasicExporterTrait;

	public function testExportBasic (): void {
		$ex = new ExporterHtml('test');
		$this->fillSheetBasic($ex);

$txt = '<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" lang="cs" xml:lang="cs">
<head>
	<title>Export test</title>
	<meta charset="utf-8" />
	<meta name="generator" content="SheetExporter '.Exporter::VERSION.'" />
	<style>
		td { border-color: #000000; }
		td.number { text-align: right; }
		th,td { font-size: 20pt; height: 20pt }
	</style>
</head>
<body>
	<table border="1" cellspacing="0" cellpadding="5">
		<caption>List</caption>
		<tbody>
			<tr><td>one with &apos; and &quot;</td></tr>
			<tr><td rowspan="3">two</td><td>two, next</td></tr>
			<tr><td colspan="3">three</td><td>last</td></tr>
			<tr></tr>
			<tr><td colspan="3">four</td><td>column, next</td></tr>
			<tr><td rowspan="3">three</td><td>last</td></tr>
		</tbody>
	</table>
</body>
</html>
';

		$this->expectOutputString($txt);
		$ex->compile();
	}

	public function testExportComplex (): void {
		$ex = new ExporterHtml('test');
		$this->fillSheetComplex($ex);

$txt = '<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" lang="cs" xml:lang="cs">
<head>
	<title>Export test</title>
	<meta charset="utf-8" />
	<meta name="generator" content="SheetExporter '.Exporter::VERSION.'" />
	<style>
		td { border-color: #000000; }
		td.number { text-align: right; }
		.one { font-size: 20pt; text-align: right; border-width: 5pt; height: 20pt }
		.two { color: #fff000; border-left-color: #000fff; border-left-style: dotted; border-left-width: 10pt; height: 40pt }
		.three { background-color: #6600ff }
	</style>
</head>
<body>
	<table border="1" cellspacing="0" cellpadding="5">
		<caption>List</caption>
		<col style="width: 50pt" />
		<col style="width: 30pt" />
		<col style="width: 56.7pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<col style="width: 20pt" />
		<thead>
			<tr>
				<th>&lt;h1&gt;Nadpis&lt;/h1&gt;</th><th>Další nadpis</th>			</tr>
		</thead>
		<tbody>
			<tr><td>obsah</td><td class="one">text</td></tr>
			<tr><td>a</td><td>b</td><td>c</td><td>d</td><td>e</td><td>f</td><td>g</td><td>h</td><td>i</td><td>j</td><td>k</td><td>a</td><td>b</td><td>c</td><td>d</td><td>e</td><td>f</td><td>g</td><td>h</td><td>i</td><td>j</td><td>k</td><td>a</td><td>b</td><td>c</td><td>d</td><td>e</td><td>f</td><td>g</td><td>h</td><td>i</td><td>j</td><td>k</td><td>a</td><td>b</td><td>c</td><td>d</td><td>e</td></tr>
			<tr><td class="two" rowspan="2" colspan="4">super radek</td><td class="two">další&amp;znak</td></tr>
			<tr><td class="three">End?</td></tr>
			<tr><td>final</td><td>row</td></tr>
		</tbody>
	</table>
</body>
</html>
';
		$this->expectOutputString($txt);
		$ex->compile();
	}

	public function testExportFormula (): void {
		$ex = new ExporterHtml('test');
		$this->fillSheetFormula($ex);

$txt = '<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" lang="cs" xml:lang="cs">
<head>
	<title>Export test</title>
	<meta charset="utf-8" />
	<meta name="generator" content="SheetExporter '.Exporter::VERSION.'" />
	<style>
		td { border-color: #000000; }
		td.number { text-align: right; }
	</style>
</head>
<body>
	<table border="1" cellspacing="0" cellpadding="5">
		<caption>List</caption>
		<tbody>
			<tr><td class="number">1</td><td class="number">2</td><td class="number">3</td><td class="number">4</td><td>test</td></tr>
			<tr><td></td><td></td></tr>
		</tbody>
	</table>
</body>
</html>
';

		$this->expectOutputString($txt);
		$ex->compile();
	}
}

<?php
use SheetExporter\Exporter,
	SheetExporter\ExporterHtml;

class ExporterHtmlTest extends \PHPUnit\Framework\TestCase {
	use BasicExporterTrait;

	public function testExport () {
		$ex = new ExporterHtml('test');
		$this->fillSheets($ex);

$txt = '<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" lang="cs" xml:lang="cs">
<head>
	<title>Export test</title>
	<meta charset="utf-8" />
	<meta name="generator" content="SheetExporter '.Exporter::VERSION.'" />
	<style>
		td.number { text-align: right; }
		.one { font-size: 20pt; text-align: right; border-width: 5pt; height: 20pt }
		.two { color: #fff000; border-left-color: #000fff; border-left-style: dotted; border-left-width: 10pt; height: 40pt }
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
			<tr><td>End?</td></tr>
			<tr><td>final</td><td>row</td></tr>
		</tbody>
	</table>
</body>
</html>
';
		$this->expectOutputString($txt);
		$ex->compile();
	}
}
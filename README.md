# SheetExporter

## Description

Simple XLSX/ODS/HTML table exporter

## Requirements

- PHP 7.1+ (PHP 5.3+ for version 0.86 and older)
- for XLSX and ODS require php-zip extension

## Usage

example:

	$ex = new ExporterXlsx('fileName');

	// add custom style with unique name
	$ex->addStyle('one', ['SIZE'=>20, 'ALIGN'=>'right'], ['WIDTH'=>5], 20);
	$ex->addStyle('two', ['COLOR'=>'#fff000'], ['LEFT'=>['COLOR'=>'#000fff', 'STYLE'=>'dotted', 'WIDTH'=>10]], 40);

	// add sheet
	$sheet = $ex->insertSheet('List');
	// set column sizes
	$sheet->setColWidth([50, '30pt', '20mm'] /* first 3 columns */, 20 /* default size */);
	// set first highlighted row
	$sheet->setColHeaders(['Title','Next title']);
	// insert row
	$sheet->addRow(['a','b','c','d','e','f','g','h','i','j','k']);
	// insert column with custom style
	$sheet->addRow([['STYLE'=>'one', 'VAL'=>'text']]);
	// insert custom size cell
	$sheet->addRow([['COLS'=>4, 'ROWS'=>2, 'VAL'=>'big cell'], 'one']);
	// set custom style to every cell in row
	$sheet->addRow(['Last','row'], 'two');

	// generate and download file
	$exporter->download();

## TODO
Add basic formula support

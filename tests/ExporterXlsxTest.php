<?php
use SheetExporter\Exporter,
	SheetExporter\ExporterXlsx;

class ExporterXlsxTest extends \PHPUnit\Framework\TestCase {
	use BasicExporterTrait;

	public function testExportBasic () {
		$ex = new ExporterXlsx('test');
		$this->fillSheetBasic($ex);
		$sheets = $ex->getSheets();

$txt = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetFormatPr defaultRowHeight="20" />
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>one with &apos; and &quot;</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>two</t></is></c><c r="B2" t="inlineStr"><is><t>two, next</t></is></c></row>
    <row r="3"><c r="A3" /><c r="B3" t="inlineStr"><is><t>three</t></is></c><c r="C3" /><c r="D3" /><c r="E3" t="inlineStr"><is><t>last</t></is></c></row>
    <row r="4"><c r="A4" /></row>
    <row r="5"><c r="A5" t="inlineStr"><is><t>four</t></is></c><c r="B5" /><c r="C5" /><c r="D5" t="inlineStr"><is><t>column, next</t></is></c></row>
    <row r="6"><c r="A6" t="inlineStr"><is><t>three</t></is></c><c r="B6" t="inlineStr"><is><t>last</t></is></c></row>
  </sheetData>
  <mergeCells count="4">
    <mergeCell ref="A2:A4"/>
    <mergeCell ref="B3:D3"/>
    <mergeCell ref="A5:C5"/>
    <mergeCell ref="A6:A8"/>
  </mergeCells>
</worksheet>
';
		$this->assertEquals($txt, $this->callPrivateMethod($ex, 'fileSheet', [reset($sheets)]));

$txt = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="20" />
    </font>
  </fonts>
  <fills count="2">
    <fill>
      <patternFill patternType="none" />
    </fill>
    <fill>
      <patternFill patternType="gray125" />
    </fill>
  </fills>
  <borders count="1">
    <border>
      <left />
      <right />
      <top />
      <bottom />
      <diagonal />
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"></xf>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normální" xfId="0" builtinId="0" />
  </cellStyles>
  <dxfs count="0" />
</styleSheet>
';
		$this->assertEquals($txt, $this->callPrivateMethod($ex, 'fileStyles'));

$txt = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="List" sheetId="1" r:id="rId2" />
  </sheets>
</workbook>
';
		$this->assertEquals($txt, $this->callPrivateMethod($ex, 'fileWorkbook'));

$txt = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>
';
		$this->assertEquals($txt, $this->callPrivateMethod($ex, 'fileContentTypes'));

$txt = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" Id="rId1" />
  <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet1.xml" Id="rId2" />
</Relationships>
';
		$this->assertEquals($txt, $this->callPrivateMethod($ex, 'fileRelationships', ['worksheet', $this->callPrivateMethod($ex, 'getSheetRelationships') ]));
	}

	public function testExportComplex () {
		$ex = new ExporterXlsx('test');
		$this->fillSheetComplex($ex);
		$sheets = $ex->getSheets();

$txt = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetFormatPr defaultRowHeight="" />
  <cols>
    <col collapsed="false" hidden="false" min="1" max="1" width="9.52107852" />
    <col collapsed="false" hidden="false" min="2" max="2" width="5.71264744" />
    <col collapsed="false" hidden="false" min="3" max="3" width="10.79690323" />
    <col collapsed="false" hidden="false" min="4" max="38" width="3.80843163" />
  </cols>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>&lt;h1&gt;Nadpis&lt;/h1&gt;</t></is></c><c r="B1" t="inlineStr"><is><t>Další nadpis</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>obsah</t></is></c><c r="B2" s="1" t="inlineStr"><is><t>text</t></is></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>a</t></is></c><c r="B3" t="inlineStr"><is><t>b</t></is></c><c r="C3" t="inlineStr"><is><t>c</t></is></c><c r="D3" t="inlineStr"><is><t>d</t></is></c><c r="E3" t="inlineStr"><is><t>e</t></is></c><c r="F3" t="inlineStr"><is><t>f</t></is></c><c r="G3" t="inlineStr"><is><t>g</t></is></c><c r="H3" t="inlineStr"><is><t>h</t></is></c><c r="I3" t="inlineStr"><is><t>i</t></is></c><c r="J3" t="inlineStr"><is><t>j</t></is></c><c r="K3" t="inlineStr"><is><t>k</t></is></c><c r="L3" t="inlineStr"><is><t>a</t></is></c><c r="M3" t="inlineStr"><is><t>b</t></is></c><c r="N3" t="inlineStr"><is><t>c</t></is></c><c r="O3" t="inlineStr"><is><t>d</t></is></c><c r="P3" t="inlineStr"><is><t>e</t></is></c><c r="Q3" t="inlineStr"><is><t>f</t></is></c><c r="R3" t="inlineStr"><is><t>g</t></is></c><c r="S3" t="inlineStr"><is><t>h</t></is></c><c r="T3" t="inlineStr"><is><t>i</t></is></c><c r="U3" t="inlineStr"><is><t>j</t></is></c><c r="V3" t="inlineStr"><is><t>k</t></is></c><c r="W3" t="inlineStr"><is><t>a</t></is></c><c r="X3" t="inlineStr"><is><t>b</t></is></c><c r="Y3" t="inlineStr"><is><t>c</t></is></c><c r="Z3" t="inlineStr"><is><t>d</t></is></c><c r="AA3" t="inlineStr"><is><t>e</t></is></c><c r="AB3" t="inlineStr"><is><t>f</t></is></c><c r="AC3" t="inlineStr"><is><t>g</t></is></c><c r="AD3" t="inlineStr"><is><t>h</t></is></c><c r="AE3" t="inlineStr"><is><t>i</t></is></c><c r="AF3" t="inlineStr"><is><t>j</t></is></c><c r="AG3" t="inlineStr"><is><t>k</t></is></c><c r="AH3" t="inlineStr"><is><t>a</t></is></c><c r="AI3" t="inlineStr"><is><t>b</t></is></c><c r="AJ3" t="inlineStr"><is><t>c</t></is></c><c r="AK3" t="inlineStr"><is><t>d</t></is></c><c r="AL3" t="inlineStr"><is><t>e</t></is></c></row>
    <row r="4" ht="40"><c r="A4" s="2" t="inlineStr"><is><t>super radek</t></is></c><c r="B4" s="2" /><c r="C4" s="2" /><c r="D4" s="2" /><c r="E4" s="2" t="inlineStr"><is><t>další&amp;znak</t></is></c></row>
    <row r="5"><c r="A5" /><c r="B5" /><c r="C5" /><c r="D5" /><c r="E5" s="3" t="inlineStr"><is><t>End?</t></is></c></row>
    <row r="6"><c r="A6" t="inlineStr"><is><t>final</t></is></c><c r="B6" t="inlineStr"><is><t>row</t></is></c></row>
  </sheetData>
  <mergeCells count="1">
    <mergeCell ref="A4:D5"/>
  </mergeCells>
</worksheet>
';
		$this->assertEquals($txt, $this->callPrivateMethod($ex, 'fileSheet', [reset($sheets)]));

$txt = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font>
      <sz val="10" />
    </font>
    <font>
      <sz val="20" />
    </font>
    <font>
      <sz val="10" />
      <color rgb="FFFFF000" />
    </font>
  </fonts>
  <fills count="3">
    <fill>
      <patternFill patternType="none" />
    </fill>
    <fill>
      <patternFill patternType="gray125" />
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FF6600FF" />
      </patternFill>
    </fill>
  </fills>
  <borders count="3">
    <border>
      <left />
      <right />
      <top />
      <bottom />
      <diagonal />
    </border>
    <border>
      <left style="thick">
        <color rgb="FF000000" />
      </left>
      <right style="thick">
        <color rgb="FF000000" />
      </right>
      <top style="thick">
        <color rgb="FF000000" />
      </top>
      <bottom style="thick">
        <color rgb="FF000000" />
      </bottom>
      <diagonal />
    </border>
    <border>
      <left style="dotted">
        <color rgb="FF000FFF" />
      </left>
      <right />
      <top />
      <bottom />
      <diagonal />
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />
  </cellStyleXfs>
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"></xf>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0">
      <alignment horizontal="right" />
    </xf>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="2" xfId="0"></xf>
    <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0"></xf>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normální" xfId="0" builtinId="0" />
  </cellStyles>
  <dxfs count="0" />
</styleSheet>
';
		$this->assertEquals($txt, $this->callPrivateMethod($ex, 'fileStyles'));
	}
}
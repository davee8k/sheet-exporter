<?php declare(strict_types=1);

use SheetExporter\ExporterOds;

class ExporterOdsTest extends \PHPUnit\Framework\TestCase {
	use BasicExporterTrait;

	public function testExportBasic (): void {
		$ex = new ExporterOds('test');
		$this->fillSheetBasic($ex);

$txt = '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:drawooo="http://openoffice.org/2010/draw" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" xmlns:css3t="http://www.w3.org/TR/css3-text/" office:version="1.2">
  <office:scripts/>
  <office:font-face-decls/>
  <office:automatic-styles>
    <style:style style:name="rdef" style:family="table-row">
      <style:table-row-properties style:row-height="7.054674mm" />
    </style:style>
  </office:automatic-styles>
  <office:body>
    <office:spreadsheet>
      <table:table table:name="List">
        <table:table-column/>
        <table:table-row table:style-name="rdef"><table:table-cell office:value-type="string"><text:p>one with &apos; and &quot;</text:p></table:table-cell></table:table-row>
        <table:table-row table:style-name="rdef"><table:table-cell office:value-type="string" table:number-rows-spanned="3"><text:p>two</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>two, next</text:p></table:table-cell></table:table-row>
        <table:table-row table:style-name="rdef"><table:covered-table-cell /><table:table-cell office:value-type="string" table:number-rows-spanned="1" table:number-columns-spanned="3"><text:p>three</text:p></table:table-cell><table:covered-table-cell table:number-columns-repeated="2" /><table:table-cell office:value-type="string"><text:p>last</text:p></table:table-cell></table:table-row>
        <table:table-row table:style-name="rdef"><table:covered-table-cell /></table:table-row>
        <table:table-row table:style-name="rdef"><table:table-cell office:value-type="string" table:number-rows-spanned="1" table:number-columns-spanned="3"><text:p>four</text:p></table:table-cell><table:covered-table-cell table:number-columns-repeated="2" /><table:table-cell office:value-type="string"><text:p>column, next</text:p></table:table-cell></table:table-row>
        <table:table-row table:style-name="rdef"><table:table-cell office:value-type="string" table:number-rows-spanned="3"><text:p>three</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>last</text:p></table:table-cell></table:table-row>
      </table:table>
      <table:named-expressions/>
    </office:spreadsheet>
  </office:body>
</office:document-content>
';
		$this->assertEquals($txt, $this->callPrivateMethod($ex, 'fileSheet'));

$txt = '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="manifest.rdf" manifest:media-type="application/rdf+xml"/>
  <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="META-INF/manifest.xml" manifest:media-type="text/xml" />
</manifest:manifest>
';
	$this->assertEquals($txt, $this->callPrivateMethod($ex, 'fileMetaInf', ['content.xml']));

	$this->assertEquals('application/vnd.oasis.opendocument.spreadsheet', $this->callPrivateMethod($ex, 'fileMime'));

$txt = '<office:document-styles xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">
  <office:font-face-decls/>
  <office:styles>
    <number:number-style style:name="N0">
      <number:number number:min-integer-digits="1" />
    </number:number-style>
    <style:style style:name="Default" style:family="table-cell" style:data-style-name="N0">
      <style:table-cell-properties style:vertical-align="automatic" fo:background-color="transparent" />
      <style:text-properties fo:font-size="20pt" style:font-size-asian="20pt" style:font-size-complex="20pt" />
      <style:table-cell-properties />
      <style:text-properties fo:color="#000000" />
    </style:style>
  </office:styles>
  <office:automatic-styles />
  <office:master-styles/>
</office:document-styles>
';
		$this->assertEquals($txt, $this->callPrivateMethod($ex, 'fileStyles'));

$txt = '<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
  <rdf:Description rdf:about="content.xml">
    <rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/odf#ContentFile"/>
  </rdf:Description>
  <rdf:Description rdf:about="">
    <ns0:hasPart xmlns:ns0="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#" rdf:resource="content.xml"/>
  </rdf:Description>
  <rdf:Description rdf:about="">
    <rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#Document"/>
  </rdf:Description>
</rdf:RDF>
';
		$this->assertEquals($txt, $this->callPrivateMethod($ex, 'fileManifest', ['content.xml']));

$txt = preg_quote('<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:grddl="http://www.w3.org/2003/g/data-view#" office:version="1.2">
  <office:meta>
    <meta:generator>'.$ex->getVersion().'</meta:generator>
    <meta:creation-date>', '/').'([0-9T\-\+\:]+)'.preg_quote('</meta:creation-date>
    <dc:date>', '/').'([0-9T\-\+\:]+)'.preg_quote('</dc:date>
  </office:meta>
</office:document-meta>
', '/');
		$this->assertMatchesRegularExpression('/'.$txt.'/', $this->callPrivateMethod($ex, 'fileMeta'));
	}

	public function testExportComplex (): void {
		$ex = new ExporterOds('test');
		$this->fillSheetComplex($ex);

$txt = '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:drawooo="http://openoffice.org/2010/draw" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" xmlns:css3t="http://www.w3.org/TR/css3-text/" office:version="1.2">
  <office:scripts/>
  <office:font-face-decls/>
  <office:automatic-styles>
    <style:style style:name="col_0_0" style:family="table-column">
      <style:table-column-properties fo:break-before="auto" style:column-width="17.636684mm" />
    </style:style>
    <style:style style:name="col_0_1" style:family="table-column">
      <style:table-column-properties fo:break-before="auto" style:column-width="10.582011mm" />
    </style:style>
    <style:style style:name="col_0_2" style:family="table-column">
      <style:table-column-properties fo:break-before="auto" style:column-width="20mm" />
    </style:style>
    <style:style style:name="col_0_default" style:family="table-column">
      <style:table-column-properties fo:break-before="auto" style:column-width="7.054674mm" />
    </style:style>
    <style:style style:name="tc_one" style:family="table-cell">
      <style:text-properties fo:font-size="20pt" style:font-size-asian="20pt" style:font-size-complex="20pt" />
      <style:paragraph-properties fo:text-align="end" />
      <style:table-cell-properties fo:border= "5pt solid #000000" />
    </style:style>
    <style:style style:name="ro_one" style:family="table-row">
      <style:table-row-properties style:row-height="7.054674mm" />
    </style:style>
    <style:style style:name="tc_two" style:family="table-cell">
      <style:text-properties fo:color="#fff000" />
      <style:table-cell-properties fo:border-left= "10pt dotted #000fff" />
    </style:style>
    <style:style style:name="ro_two" style:family="table-row">
      <style:table-row-properties style:row-height="14.109347mm" />
    </style:style>
    <style:style style:name="tc_three" style:family="table-cell">
      <style:text-properties />
      <style:table-cell-properties fo:background-color="#6600ff" />
    </style:style>
  </office:automatic-styles>
  <office:body>
    <office:spreadsheet>
      <table:table table:name="List">
        <table:table-column table:style-name="col_0_0" />
        <table:table-column table:style-name="col_0_1" />
        <table:table-column table:style-name="col_0_2" />
        <table:table-column table:style-name="col_0_default" table:number-columns-repeated="35" />
<table:table-row><table:table-cell office:value-type="string"><text:p>&lt;h1&gt;Nadpis&lt;/h1&gt;</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>Další nadpis</text:p></table:table-cell></table:table-row>
        <table:table-row><table:table-cell office:value-type="string"><text:p>obsah</text:p></table:table-cell><table:table-cell table:style-name="tc_one" office:value-type="string"><text:p>text</text:p></table:table-cell></table:table-row>
        <table:table-row><table:table-cell office:value-type="string"><text:p>a</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>b</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>c</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>d</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>e</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>f</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>g</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>h</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>i</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>j</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>k</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>a</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>b</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>c</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>d</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>e</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>f</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>g</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>h</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>i</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>j</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>k</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>a</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>b</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>c</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>d</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>e</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>f</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>g</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>h</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>i</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>j</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>k</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>a</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>b</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>c</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>d</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>e</text:p></table:table-cell></table:table-row>
        <table:table-row table:style-name="ro_two"><table:table-cell table:style-name="tc_two" office:value-type="string" table:number-rows-spanned="2" table:number-columns-spanned="4"><text:p>super radek</text:p></table:table-cell><table:covered-table-cell table:number-columns-repeated="3" /><table:table-cell table:style-name="tc_two" office:value-type="string"><text:p>další&amp;znak</text:p></table:table-cell></table:table-row>
        <table:table-row><table:covered-table-cell table:number-columns-repeated="4" /><table:table-cell table:style-name="tc_three" office:value-type="string"><text:p>End?</text:p></table:table-cell></table:table-row>
        <table:table-row><table:table-cell office:value-type="string"><text:p>final</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>row</text:p></table:table-cell></table:table-row>
      </table:table>
      <table:named-expressions/>
    </office:spreadsheet>
  </office:body>
</office:document-content>
';
		$this->assertEquals($txt, $this->callPrivateMethod($ex, 'fileSheet'));

$txt = '<office:document-styles xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">
  <office:font-face-decls/>
  <office:styles>
    <number:number-style style:name="N0">
      <number:number number:min-integer-digits="1" />
    </number:number-style>
    <style:style style:name="Default" style:family="table-cell" style:data-style-name="N0">
      <style:table-cell-properties style:vertical-align="automatic" fo:background-color="transparent" />
      <style:text-properties fo:color="#000000" />
    </style:style>
  </office:styles>
  <office:automatic-styles />
  <office:master-styles/>
</office:document-styles>
';
		$this->assertEquals($txt, $this->callPrivateMethod($ex, 'fileStyles'));
	}

	public function testExportFormula (): void {
		$ex = new ExporterOds('test');
		$this->fillSheetFormula($ex);

$txt = '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:drawooo="http://openoffice.org/2010/draw" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" xmlns:css3t="http://www.w3.org/TR/css3-text/" office:version="1.2">
  <office:scripts/>
  <office:font-face-decls/>
  <office:automatic-styles>
  </office:automatic-styles>
  <office:body>
    <office:spreadsheet>
      <table:table table:name="List">
        <table:table-column/>
        <table:table-row><table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="3"><text:p>3</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="4"><text:p>4</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>test</text:p></table:table-cell></table:table-row>
        <table:table-row><table:table-cell office:value-type="string" table:formula="of:=SUM([.A1:.D1])"><text:p></text:p></table:table-cell><table:table-cell office:value-type="string" table:formula="of:=[.$E$1]"><text:p></text:p></table:table-cell></table:table-row>
      </table:table>
      <table:named-expressions/>
    </office:spreadsheet>
  </office:body>
</office:document-content>
';

		$this->assertEquals($txt, $this->callPrivateMethod($ex, 'fileSheet'));
	}
}

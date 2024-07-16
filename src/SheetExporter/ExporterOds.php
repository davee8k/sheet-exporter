<?php
declare(strict_types=1);

namespace SheetExporter;

use RuntimeException;
use InvalidArgumentException;
use ZipArchive;

/**
 * Export to Ods
 */
class ExporterOds extends Exporter
{
	/**
	 * @param string $fileName
	 * @throws RuntimeException
	 */
	public function __construct(string $fileName)
	{
		if (!class_exists('ZipArchive')) throw new RuntimeException('Missing ZipArchive extension for ODS.');
		parent::__construct($fileName);
	}

	/**
	 * Create download content
	 */
	public function download(): void
	{
		$tempFile = $this->compile();
		header('Content-Type: application/vnd.oasis.opendocument.spreadsheet; charset=utf-8');
		header('Content-Disposition: attachment; filename="'.$this->fileName.'.ods"');
		readfile($tempFile);
		@unlink($tempFile);
	}

	/**
	 * Generate Ods file
	 * @return string
	 * @throws RuntimeException
	 */
	public function compile(): string
	{
		$zip = new ZipArchive;
		$tempFile = $this->createTemp();
		$res = $zip->open($tempFile, ZipArchive::CREATE | ZipArchive::OVERWRITE);
		if ($res === true) {
			$contentFile = 'content.xml';
			$zip->addFromString('mimetype', $this->fileMime());
			$zip->addFromString('META-INF/manifest.xml', self::XML_HEADER.$this->fileMetaInf($contentFile));
			$zip->addFromString('manifest.rdf', self::XML_HEADER.$this->fileManifest($contentFile));
			$zip->addFromString('meta.xml', self::XML_HEADER.$this->fileMeta());
			$zip->addFromString('styles.xml', self::XML_HEADER.$this->fileStyles());
			$zip->addFromString($contentFile, self::XML_HEADER.$this->fileSheet());

			if ($zip->close()) return $tempFile;
		}
		throw new RuntimeException("Failed to export: ".$res);
	}

	/**
	 *
	 * @return string
	 * @throws InvalidArgumentException
	 */
	private function fileSheet(): string
	{
		ob_start();
		$defHeight = null;
?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:drawooo="http://openoffice.org/2010/draw" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" xmlns:css3t="http://www.w3.org/TR/css3-text/" office:version="1.2">
  <office:scripts/>
  <office:font-face-decls/>
  <office:automatic-styles>
<?php
		$this->printColStyles();

		if (!empty($this->defaultStyle['HEIGHT'])) {
			$defHeight = 'rdef';
?>
    <style:style style:name="<?=$defHeight;?>" style:family="table-row">
      <style:table-row-properties style:row-height="<?=self::convertSize($this->defaultStyle['HEIGHT'], self::UNITS, 'mm');?>mm" />
    </style:style>
<?php
		}

		// row styles
		foreach ($this->styles as $mark => $style) {
?>
    <style:style style:name="tc_<?=$mark;?>" style:family="table-cell">
<?php
		echo $this->getStyle($style);
?>
    </style:style>
<?php
			if (!empty($style['HEIGHT'])) {
?>
    <style:style style:name="ro_<?=$mark;?>" style:family="table-row">
      <style:table-row-properties style:row-height="<?=self::convertSize($style['HEIGHT'], self::UNITS, 'mm');?>mm" />
    </style:style>
<?php
			}
		}
?>
  </office:automatic-styles>
  <office:body>
    <office:spreadsheet>
<?php
		foreach ($this->sheets as $sid => $sheet) {
?>
      <table:table table:name="<?=$sheet->getName();?>">
<?php
			$this->printColWidths($sheet, $sid);

			$this->printHeader($sheet);
			$this->printSheet($sheet, $defHeight);
?>
      </table:table>
<?php
		}
?>
      <table:named-expressions/>
    </office:spreadsheet>
  </office:body>
</office:document-content>
<?php
		return ob_get_clean() ?: '';
	}

	/**
	 * Print sheet header
	 * @param Sheet $sheet
	 */
	private function printHeader(Sheet $sheet): void
	{
		if (!empty($sheet->getHeaders())) {
			echo '<table:table-row>';
			foreach ($sheet->getHeaders() as $name) {
				echo '<table:table-cell office:value-type="string"><text:p>'.self::xmlEntities($name).'</text:p></table:table-cell>';
			}
			echo "</table:table-row>\n";
		}
	}

	/**
	 * Print sheet content
	 * @param Sheet $sheet
	 * @param string|null $defHeight
	 * @throws InvalidArgumentException
	 */
	private function printSheet(Sheet $sheet, ?string $defHeight): void
	{
		$skipPlan = [];
		foreach ($sheet->getRows() as $num => $row) {
			$last = -1;
			$move = 0;
			$class = $sheet->getStyle($num);
			$height = $this->getRowHeight($class, $defHeight);

			echo '        <table:table-row',($height ? ' table:style-name="'.$height.'"' : ''),'>';
			for ($j = 0; $j <= $move + $sheet->getColCount(); $j++) {
				// insert empty cells under merged
				if (isset($skipPlan[$num][$j])) {
					echo '<table:covered-table-cell'.($skipPlan[$num][$j] > 1 ? ' table:number-columns-repeated="'.$skipPlan[$num][$j].'"' : '').' />';
					$move += $skipPlan[$num][$j];
				}
				if (isset($row[$j - $move]) && $last < ($j - $move)) {
					$last = $j - $move;
					echo $this->getCell($row[$last], $num, $j, $class, $skipPlan);
				}
			}
			echo "</table:table-row>\n";

			if (isset($skipPlan[$num])) unset($skipPlan[$num]);
		}
	}

	/**
	 *  Print columns widths
	 * @param Sheet $sheet
	 * @param string|int $sid
	 */
	private function printColWidths(Sheet $sheet, $sid): void
	{
		$count = count($sheet->getCols());
		if ($count === 0) {
?>
        <table:table-column/>
<?php
		} else {
			$keys = array_keys($sheet->getCols());
			foreach ($keys as $num) {
?>
        <table:table-column table:style-name="col_<?=$sid.'_'.$num;?>" />
<?php
			}

			if ($sheet->getDefCol() && $count < $sheet->getColCount()) {
?>
        <table:table-column table:style-name="col_<?=$sid;?>_default" table:number-columns-repeated="<?=$sheet->getColCount() - $count;?>" />
<?php
			}
		}
	}

	/**
	 * Print columns styles
	 */
	private function printColStyles(): void
	{
		// columns widths
		foreach ($this->sheets as $sid => $sheet) {
			foreach ($sheet->getCols() as $num => $col) {
?>
    <style:style style:name="col_<?=$sid.'_'.$num;?>" style:family="table-column">
      <style:table-column-properties fo:break-before="auto" style:column-width="<?=self::convertSize($col, self::UNITS, 'mm');?>mm" />
    </style:style>
<?php
			}

		if ($sheet->getDefCol() && count($sheet->getCols()) < $sheet->getColCount()) {
?>
    <style:style style:name="col_<?=$sid.'_default';?>" style:family="table-column">
      <style:table-column-properties fo:break-before="auto" style:column-width="<?=self::convertSize($sheet->getDefCol(), self::UNITS, 'mm');?>mm" />
    </style:style>
<?php
			}
		}
	}

	/**
	 * Get table cell
	 * @param array<string, mixed>|string|float|int|null $col
	 * @param int $num
	 * @param int $j
	 * @param string|null $class
	 * @param array<int, array<int, int>> $skipPlan
	 * @return string
	 */
	private function getCell($col, int $num, int $j, ?string $class, array &$skipPlan): string
	{
		if (isset($col['COLS']) && $col['COLS'] > 1 || isset($col['ROWS']) && $col['ROWS'] > 1) {
			$cols = $col['COLS'] ?? 1;
			$rows = $col['ROWS'] ?? 1;

			// addd empty cells under merged
			if ($rows > 1) {
				for ($i = 1; $i < $rows; $i++) {
					$skipPlan[$num + $i][$j] = $cols;
				}
			}
			if ($cols > 1) {
				$skipPlan[$num][$j + 1] = $cols - 1;
			}
		};

		if (is_array($col))	{
			return $this->getCellValue($col['VAL'] ?? '', $col['STYLE'] ?? $class, $col);
		} elseif ($col !== null) {
			return $this->getCellValue($col, $class);
		}
		return '';
	}


	/**
	 * Get table cell value
	 * @param string|float|int $val
	 * @param string|null $class
	 * @param array<string, mixed>|null $col
	 * @return string
	 */
	private function getCellValue($val, ?string $class = null, ?array $col = null): string
	{
		return '<table:table-cell'.($class ? ' table:style-name="tc_'.$class.'"' : '').
			' office:value-type='.(is_numeric($val) ? '"float" office:value="'.$val.'"' : '"string"').
			(empty($col['FORMULA']) ? '' : ' table:formula="of:='.$this->reformatFormula($col['FORMULA']).'"').
			(isset($col['COLS']) || isset($col['ROWS']) && $col['ROWS'] > 1 ? ' table:number-rows-spanned="'.($col['ROWS'] ?? 1).'"' : '').
			(isset($col['COLS']) && $col['COLS'] > 1 ? ' table:number-columns-spanned="'.$col['COLS'].'"' : '').
			'><text:p>'.self::xmlEntities($val).'</text:p></table:table-cell>';
	}

	/**
	 * Return row height
	 * @param string|null $class
	 * @param string|null $default
	 * @return string|null
	 * @throws InvalidArgumentException
	 */
	private function getRowHeight(?string $class, ?string $default): ?string
	{
		if ($class) {
			if (isset($this->styles[$class])) {
				if (!empty($this->styles[$class]['HEIGHT'])) return 'ro_'.$class;
			} else {
				throw new InvalidArgumentException('Missing style: '.htmlspecialchars((string) $class, ENT_QUOTES));
			}
		}
		return $default;
	}

	/**
	 *
	 * @param string $formula
	 * @return string
	 */
	private function reformatFormula(string $formula): string
	{
		return (string) preg_replace_callback(
				'/(\$?[A-Z]+\$?[0-9]+)((\:?)(\$?[A-Z]+\$?[0-9]+))?/',
				function($arg) { return '[.'.$arg[1].(isset($arg[4]) ? $arg[3].'.'.$arg[4] : '').']'; },
				$formula
			);
	}

	/**
	 *
	 * @param array<string, mixed> $style
	 * @return string
	 */
	private function getStyle(array $style): string
	{
		$txt = '';
		if (isset($style['FONT'])) {
			$txt .= $this->getStyleFont($style['FONT']);
		}
		if (isset($style['CELL'])) {
			$txt .= '      <style:table-cell-properties'.(isset($style['CELL']['BACKGROUND']) ? ' fo:background-color="'.$style['CELL']['BACKGROUND'].'"' : '');
			if ($this->isBorderStyle($style['CELL'], self::$borderStyles)) {
				$txt .= $this->getBorderStyle($style['CELL']);
			} else {
				foreach (self::$borderTypes as $key => $mark) {
					if (!empty($style['CELL'][$mark])) {
						$txt .= $this->getBorderStyle($style['CELL'][$mark], '-'.$key);
					}
				}
			}
			$txt .= " />\n";
		}
		return $txt;
	}

	/**
	 * Returns font style
	 * @param array<string, mixed> $style
	 * @return string
	 */
	private function getStyleFont(array $style): string
	{
		$txt = '      <style:text-properties';
		if (isset($style['COLOR'])) {
			$txt .= ' fo:color="'.$style['COLOR'].'"';
		}
		if (isset($style['WEIGHT'])) {
			$txt .= ' fo:font-weight="'.$style['WEIGHT'].'" style:font-weight-asian="'.$style['WEIGHT'].'" style:font-weight-complex="'.$style['WEIGHT'].'"';
		}
		if (isset($style['SIZE'])) {
			if (is_numeric($style['SIZE'])) $style['SIZE'] .= self::UNITS;
			$txt .= ' fo:font-size="'.$style['SIZE'].'" style:font-size-asian="'.$style['SIZE'].'" style:font-size-complex="'.$style['SIZE'].'"';
		}
		if (isset($style['FAMILY'])) {
			$txt .= ' style:font-name="'.$style['FAMILY'].'" style:font-name-asian="'.$style['FAMILY'].'" style:font-name-complex="'.$style['FAMILY'].'"';
		}
		if (isset($style['BACKGROUND'])) {
			$txt .= ' fo:background-color="'.$style['BACKGROUND'].'"';
		}
		$txt .= " />\n";

		if (isset($style['ALIGN'])) {
			$txt .= '      <style:paragraph-properties fo:text-align="'.str_ireplace(['left','right'], ['start','end'], $style['ALIGN']).'" />'."\n";
		}
		return $txt;
	}

	/**
	 *
	 * @param array<string, mixed> $style
	 * @param string $side
	 * @return string
	 */
	private function getBorderStyle(array $style, string $side = ''): string
	{
		$txt = ' fo:border'.$side.'= "';
		if (isset($style['WIDTH'])) $txt .= self::convertSize($style['WIDTH'], self::UNITS).self::UNITS.' ';
		return $txt.($style['STYLE'] ?? 'solid').' '.($style['COLOR'] ?? static::$defColor).'"';
	}

	/**
	 *
	 * @return string
	 */
	private function fileStyles(): string
	{
		ob_start();
?>
<office:document-styles xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">
  <office:font-face-decls/>
  <office:styles>
    <number:number-style style:name="N0">
      <number:number number:min-integer-digits="1" />
    </number:number-style>
    <style:style style:name="Default" style:family="table-cell" style:data-style-name="N0">
      <style:table-cell-properties style:vertical-align="automatic" fo:background-color="transparent" />
<?php
		if (!empty($this->defaultStyle)) {
			echo $this->getStyle($this->defaultStyle);
		}
		if (empty($this->defaultStyle) || empty($this->defaultStyle['FONT']['COLOR'])) {
			echo '      <style:text-properties fo:color="'.static::$defColor.'" />',"\n";
		}
?>
    </style:style>
  </office:styles>
  <office:automatic-styles />
  <office:master-styles/>
</office:document-styles>
<?php
		return ob_get_clean() ?: '';
	}

	/**
	 * manifest:version="1.2" - buggy for MSO2007
	 * @param string $contentFile
	 * @return string
	 */
	private function fileMetaInf(string $contentFile): string
	{
		ob_start();
?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
  <manifest:file-entry manifest:full-path="<?=$contentFile;?>" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="manifest.rdf" manifest:media-type="application/rdf+xml"/>
  <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="META-INF/manifest.xml" manifest:media-type="text/xml" />
</manifest:manifest>
<?php
		return ob_get_clean() ?: '';
	}

	/**
	 * Return ODS mime type
	 * @return string
	 */
	private function fileMime(): string
	{
		return 'application/vnd.oasis.opendocument.spreadsheet';
	}

	/**
	 * Return ODS manifest file content
	 * @param string $contentFile
	 * @return string
	 */
	private function fileManifest(string $contentFile): string
	{
		ob_start();
?>
<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
  <rdf:Description rdf:about="<?=$contentFile;?>">
    <rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/odf#ContentFile"/>
  </rdf:Description>
  <rdf:Description rdf:about="">
    <ns0:hasPart xmlns:ns0="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#" rdf:resource="<?=$contentFile;?>"/>
  </rdf:Description>
  <rdf:Description rdf:about="">
    <rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#Document"/>
  </rdf:Description>
</rdf:RDF>
<?php
		return ob_get_clean() ?: '';
	}

	/**
	 * Return ODS meta file content
	 * @return string
	 */
	private function fileMeta(): string
	{
		ob_start();
		$date = date('c');
?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:grddl="http://www.w3.org/2003/g/data-view#" office:version="1.2">
  <office:meta>
    <meta:generator><?=$this->getVersion();?></meta:generator>
    <meta:creation-date><?=$date;?></meta:creation-date>
    <dc:date><?=$date;?></dc:date>
  </office:meta>
</office:document-meta>
<?php
		return ob_get_clean() ?: '';
	}
}

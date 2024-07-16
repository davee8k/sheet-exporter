<?php
declare(strict_types=1);

namespace SheetExporter;

use RuntimeException;
use InvalidArgumentException;
use ZipArchive;

/**
 * Export to Xlsx
 */
class ExporterXlsx extends Exporter
{
	/** @var string */
	private const FONT_SIZE = '10pt';
	/** @var int */
	private const BORDER_MEDIUM = 2,
		BORDER_THICK = 4;

	/**
	 * @param string $fileName
	 * @throws RuntimeException
	 */
	public function __construct(string $fileName)
	{
		if (!class_exists('ZipArchive')) throw new RuntimeException('Missing ZipArchive extension for XLSX.');
		parent::__construct($fileName);
	}

	/**
	 * Create download content
	 */
	public function download(): void
	{
		$tempFile = $this->compile();
		header('Content-Type: application/excel; charset=utf-8');
		header('Content-Disposition: attachment; filename="'.$this->fileName.'.xlsx"');
		readfile($tempFile);
		@unlink($tempFile);
	}

	/**
	 * Generate Xlsx file
	 * @return string
	 * @throws RuntimeException
	 */
	public function compile(): string
	{
		$zip = new ZipArchive;
		$tempFile = $this->createTemp();
		$res = $zip->open($tempFile, ZipArchive::CREATE | ZipArchive::OVERWRITE);
		if ($res === true) {
			$zip->addFromString('[Content_Types].xml', self::XML_HEADER.$this->fileContentTypes());
			$zip->addFromString('_rels/.rels', self::XML_HEADER.$this->fileRelationships('officeDocument', ['rs'.md5($this->fileName) => '/xl/workbook.xml']));
			$zip->addFromString('xl/_rels/workbook.xml.rels', self::XML_HEADER.$this->fileRelationships('worksheet', $this->getSheetRelationships() ));
			$zip->addFromString('xl/workbook.xml', self::XML_HEADER.$this->fileWorkbook());
			$zip->addFromString('xl/styles.xml', self::XML_HEADER.$this->fileStyles());

			foreach ($this->sheets as $num => $sheet) {
				$zip->addFromString('xl/worksheets/sheet'.($num + 1).'.xml', self::XML_HEADER.$this->fileSheet($sheet));
			}

			if ($zip->close()) return $tempFile;
		}
		throw new RuntimeException("Failed to export: ".$res);
	}

	/**
	 * Data sheet
	 * @param Sheet $sheet
	 * @return string
	 * @throws InvalidArgumentException
	 */
	private function fileSheet(Sheet $sheet): string
	{
		$count = count($sheet->getCols());
		ob_start();
?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetFormatPr defaultRowHeight="<?=isset($this->defaultStyle['HEIGHT']) && $this->defaultStyle['HEIGHT'] ? self::convertSize($this->defaultStyle['HEIGHT']) : '';?>" />
<?php
		if ($count !== 0) {
?>
  <cols>
<?php
			foreach ($sheet->getCols() as $num => $col) {
				echo '    <col collapsed="false" hidden="false" min="',($num + 1),'" max="',($num + 1),'" width="',$this->convertColSize($col),'" />',"\n";
			}

			if ($sheet->getDefCol() && $count < $sheet->getColCount()) {
				echo '    <col collapsed="false" hidden="false" min="',($count + 1),
					'" max="',$sheet->getColCount(),'" width="',$this->convertColSize($sheet->getDefCol()),'" />',"\n";
			}
?>
  </cols>
<?php
		}
?>
  <sheetData>
<?php
		$mergeCells = [];
		$line = 0;

		$this->printHeader($sheet, $line);
		$this->printSheet($sheet, $line, $mergeCells);
?>
  </sheetData>
<?php
		if (!empty($mergeCells)) {
?>
  <mergeCells count="<?=count($mergeCells);?>">
<?php
	foreach ($mergeCells as $from => $to) {
?>
    <mergeCell ref="<?=$from.':'.$to;?>"/>
<?php
	}
?>
  </mergeCells>
<?php
		}
?>
</worksheet>
<?php
		return ob_get_clean() ?: '';
	}

	/**
	 * Print sheet header
	 * @param Sheet $sheet
	 * @param int $line
	 */
	private function printHeader(Sheet $sheet, int &$line): void
	{
		if (!empty($sheet->getHeaders())) {
			echo '    <row r="'.++$line.'">';
			foreach ($sheet->getHeaders() as $num => $name) {
				echo '<c r="'.self::toAlpha($num).$line.'" t="inlineStr"><is><t>'.self::xmlEntities($name).'</t></is></c>';
			}
			echo "</row>\n";
		}
	}

	/**
	 * Print sheet content
	 * @param Sheet $sheet
	 * @param int $line
	 * @param array<string, string> $mergeCells
	 * @throws InvalidArgumentException
	 */
	private function printSheet(Sheet $sheet, int &$line, array &$mergeCells): void
	{
		$skipPlan = [];

		$styles = array_keys($this->styles);
		$styles = array_flip($styles);

		foreach ($sheet->getRows() as $num => $row) {
			$last = -1;
			$move = 0;
			$class = $sheet->getStyle($num);
			$height = $this->getRowHeight($class);

			echo '    <row r="',++$line,'"',($height ? ' ht="'.self::convertSize($height).'"' : ''),'>';
			for ($j = 0; $j < $move + $sheet->getColCount(); $j++) {
				// insert empty cells under merged
				if (isset($skipPlan[$num][$j])) {
					for ($i = 0; $i < $skipPlan[$num][$j]; $i++) {
						echo '<c r="'.self::toAlpha($j + $i).$line.'" />';
					}
					$move += $skipPlan[$num][$j];
				}
				if (isset($row[$j - $move]) && $last < ($j - $move)) {
					$last = $j - $move;
					echo $this->getCell($row[$last], $num, $line, $j, $class, $styles, $skipPlan, $mergeCells);
				}
			}
			echo "</row>\n";

			if (isset($skipPlan[$num])) unset($skipPlan[$num]);
		}
	}

	/**
	 * Get table cell
	 * @param array<string, mixed>|string|float|int|null $col
	 * @param int $num
	 * @param int $line
	 * @param int $j
	 * @param string|null $class
	 * @param array<string, int> $styles
	 * @param array<int, array<int, int>> $skipPlan
	 * @param array<string, string> $mergeCells
	 * @return string
	 */
	private function getCell($col, int $num, int $line, int $j, ?string $class, array $styles, array &$skipPlan, array &$mergeCells): string
	{
		if (is_array($col)) {
			if (isset($col['STYLE'])) {
				if (!isset($styles[$col['STYLE']])) {
					throw new InvalidArgumentException('Missing style: '.htmlspecialchars($col['STYLE'], ENT_QUOTES));
				}
				$class = $col['STYLE'];
			}

			if (isset($col['COLS']) && $col['COLS'] > 1 || isset($col['ROWS']) && $col['ROWS'] > 1) {
				$cols = $col['COLS'] ?? 1;
				$rows = $col['ROWS'] ?? 1;

				$mergeCells[self::toAlpha($j).$line] = self::toAlpha($j + $cols - 1).($line + $rows - 1);

				// addd empty cells under merged
				if ($rows > 1) {
					for ($i = 1; $i < $rows; $i++) {
						$skipPlan[$num + $i][$j] = $cols;
					}
				}
				if ($cols > 1) {
					$skipPlan[$num][$j + 1] = $cols - 1;
				}
			}

			return $this->getCellValue($col['VAL'] ?? '', $j, $line, $class ? $styles[$class] + 1 : null, $col);
		} elseif ($col !== null) {
			return $this->getCellValue($col, $j, $line, $class ? $styles[$class] + 1 : null);
		}
		return '';
	}

	/**
	 * Get table cell value
	 * @param string|float|int $val
	 * @param int $num
	 * @param int $line
	 * @param int|null $style
	 * @param array<string, mixed>|null $col
	 * @return string
	 */
	private function getCellValue($val, int $num, int $line, int $style = null, ?array $col = null): string
	{
		$isNum = is_int($val) || is_float($val) || is_numeric($val) && ctype_digit(substr($val, 0, 1));
		return '<c r="'.self::toAlpha($num).$line.'"'.($style ? ' s="'.$style.'"' : '').' t="'.($isNum ? 'n' : 'inlineStr').'">'.
			(empty($col['FORMULA']) ? '' : '<f aca="false">'.htmlspecialchars($col['FORMULA'], ENT_QUOTES).'</f>').
			($isNum ? '<v>'.$val.'</v>' : '<is><t>'.self::xmlEntities($val).'</t></is>').'</c>';
	}

	/**
	 * Try to convert metric units to ms unit
	 * @param string|float $size (optimal - mm)
	 * @return float
	 */
	protected function convertColSize($size): float
	{
		return round((self::convertSize($size, self::UNITS, 'mm') / 1.852383) / (self::convertSize(isset($this->defaultStyle['FONT'], $this->defaultStyle['FONT']['SIZE']) ? $this->defaultStyle['FONT']['SIZE'] : self::FONT_SIZE) / 10), 8);
	}

	/**
	 * Return row height
	 * @param string|int|null $class
	 * @return string|int|null
	 * @throws InvalidArgumentException
	 */
	private function getRowHeight($class)
	{
		if ($class) {
			if (isset($this->styles[$class])) {
				if (!empty($this->styles[$class]['HEIGHT'])) return $this->styles[$class]['HEIGHT'];
			} else {
				throw new InvalidArgumentException('Missing style: '.htmlspecialchars((string) $class, ENT_QUOTES));
			}
		}
		return null;
	}

	/**
	 * XLSX style, can't be empty
	 * @return string
	 */
	protected function fileStyles(): string
	{
		ob_start();
?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<?php
		if (!isset($this->defaultStyle['FONT'], $this->defaultStyle['FONT']['SIZE'])) {
			$this->defaultStyle['FONT']['SIZE'] = self::FONT_SIZE;
		}
		if (!isset($this->defaultStyle['CELL'])) {
			$this->defaultStyle['CELL'] = [];
		}

		$groups = $this->reshuffleStyles(array_merge([$this->defaultStyle], $this->styles));
?>
  <fonts count="<?=count($groups['FONTS']);?>">
<?php
		foreach ($groups['FONTS'] as $font) {
			echo $this->getFontStyle($font);
		}
?>
  </fonts>
  <fills count="<?=count($groups['FILLS']) + 1;?>">
<?php
		foreach ($groups['FILLS'] as $num => $fill) {
?>
    <fill>
<?php
			if ($fill) {
?>
      <patternFill patternType="solid">
        <fgColor rgb="<?=self::convertColor($fill);?>" />
      </patternFill>
<?php
			} else {
?>
      <patternFill patternType="none" />
<?php
			}
?>
    </fill>
<?php
			// FIXING MSO2007 FEATURE - 1 is always gray125
			if ($num == 0) {
?>
    <fill>
      <patternFill patternType="gray125" />
    </fill>
<?php
			}
		}
?>
  </fills>
  <borders count="<?=count($groups['BORDERS']) ;?>">
<?php
		foreach ($groups['BORDERS'] as $border) {
?>
    <border>
<?php
			foreach (self::$borderTypes as $key => $mark) {
				echo $this->getBorderStyle($key, $border[$mark] ?? null);
			}
?>
      <diagonal />
    </border>
<?php
		}
?>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />
  </cellStyleXfs>
  <cellXfs count="<?=count($groups['MAP']);?>">
<?php
		$this->printCellXfs($groups['MAP']);
?>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normální" xfId="0" builtinId="0" />
  </cellStyles>
  <dxfs count="0" />
</styleSheet>
<?php
		return ob_get_clean() ?: '';
	}

	/**
	 * Create font style for xlsx
	 * @param array<string, mixed> $font
	 * @return string
	 */
	private function getFontStyle(array $font): string
	{
		$txt = "    <font>\n";
		if (isset($font['WEIGHT'])) $txt .= '      <b />'."\n";
		if (isset($font['SIZE'])) $txt .= '      <sz val="'.self::convertSize($font['SIZE']).'" />'."\n";
		if (isset($font['FAMILY'])) $txt .= '      <name val="'.$font['FAMILY'].'" />'."\n";
		if (isset($font['COLOR'])) $txt .= '      <color rgb="'.self::convertColor($font['COLOR']).'" />'."\n";
		return $txt ."    </font>\n";
	}

	/**
	 * Create border style and convert border width to 3 types used in xlsx
	 * @param string $elm
	 * @param array<string, mixed> $border
	 * @return string
	 */
	private function getBorderStyle(string $elm, array $border = null): string
	{
		$style = $border['STYLE'] ?? null;
		if ($style === 'solid') $style = 'thin';

		if (isset($border['WIDTH'])) {
			if ($border['WIDTH'] >= self::BORDER_THICK) {
				$style = $this->convertBorderType($style, 'thick');
			} elseif ($border['WIDTH'] >= self::BORDER_MEDIUM) {
				$style = $this->convertBorderType($style, 'medium');
			} elseif ($border['WIDTH'] > 0 && !$style) {
				$style = 'thin';
			}
		}

		if ($style !== null) {
			return '      <'.$elm.' style="'.$style.'">
        <color rgb="'.self::convertColor($border['COLOR'] ?? static::$defColor).'" />
      </'.$elm.">\n";
		}
		return '      <'.$elm." />\n";
	}

	/**
	 * Tries to convert not solid border to adequate type for xlsx
	 * @param string|null $style
	 * @param string $alt
	 * @return string
	 */
	private function convertBorderType(?string $style, string $alt): string
	{
		if ($style) {
			if (in_array($style, ['dashed'])) return 'medium'.ucfirst($style);
			return $style;
		}
		return $alt;
	}

	/**
	 *
	 * @param array<string, mixed> $list
	 */
	private function printCellXfs(array $list): void
	{
		foreach ($list as $item) {
			echo '    <xf numFmtId="0" fontId="', $item['font'] ?? 0,'" fillId="',
					isset($item['fill']) && $item['fill'] > 0 ? $item['fill'] + 1 : 0,'" borderId="', $item['border'] ?? 0,'" xfId="0">';
			if (!empty($item['ALIGN'])) echo "\n      ",'<alignment horizontal="'.$item['ALIGN'].'" />',"\n    ";
			echo "</xf>\n";
		}
	}

	/**
	 * Prepare styles for xlsx structure
	 * @param array<int|string, mixed> $list
	 * @return array<string, mixed>
	 */
	private function reshuffleStyles(array $list): array
	{
		$map = [];
		$fonts = [];
		$fills = [];
		$borders = [];

		foreach ($list as $num => $style) {
			$map[$num] = [];

			if (!empty($style['FONT'])) {
				if (!isset($style['FONT']['SIZE'])) {
					$style['FONT']['SIZE'] = $this->defaultStyle['FONT']['SIZE'];
				}
				if (!isset($style['FONT']['FAMILY']) && isset($this->defaultStyle['FONT']['FAMILY'])) {
					$style['FONT']['FAMILY'] = $this->defaultStyle['FONT']['FAMILY'];
				}
				$this->sortStyle($num, $style['FONT'], 'font', $fonts, $map);

				if (isset($style['FONT']['ALIGN'])) $map[$num]['ALIGN'] = $style['FONT']['ALIGN'];
			}
			if (isset($style['CELL'])) {
				$cell = $style['CELL'];
				$this->sortStyle($num, $cell['BACKGROUND'] ?? null, 'fill', $fills, $map);

				$item = [];
				if ($this->isBorderStyle($cell, self::$borderStyles)) {
					$cell = $this->reshuffleStyleBorder($cell);
				}
				foreach (self::$borderTypes as $side) {
					if (!empty($cell[$side])) {
						$item[$side] = [
							'COLOR' => $cell[$side]['COLOR'] ?? static::$defColor,
							'STYLE' => $cell[$side]['STYLE'] ?? null,
							'WIDTH' => isset($cell[$side]['WIDTH']) ? self::convertSize($cell[$side]['WIDTH']) : 1
						];
					}
				}
				$this->sortStyle($num, $item, 'border', $borders, $map);
			}
		}
		return ['MAP' => $map, 'FONTS' => $fonts, 'FILLS' => $fills, 'BORDERS' => $borders];
	}

	/**
	 * Reformat style for borders
	 * @param array<string, mixed> $cell
	 * @return array<string, mixed>
	 */
	private function reshuffleStyleBorder(array $cell): array
	{
		foreach (self::$borderStyles as $type) {
			if (isset($cell[$type])) {
				foreach (self::$borderTypes as $side) {
					$cell[$side][$type] = $cell[$type];
				}
				unset($cell[$type]);
			}
		}
		return $cell;
	}

	/**
	 * Reorder style
	 * @param string|int $num
	 * @param array<string, mixed>|string|null $item
	 * @param string $type
	 * @param array<string, mixed> $list
	 * @param array<int, array<string, string>> $map
	 */
	private function sortStyle($num, $item, string $type, array &$list, array &$map): void
	{
		$key = array_search($item, $list);
		if ($key === false) {
			$list[] = $item;
			end($list);
			$map[$num][$type] = key($list);
			reset($list);
		} else {
			$map[$num][$type] = $key;
		}
	}

	/**
	 * XLSX workbook info
	 * @return string
	 */
	private function fileWorkbook(): string
	{
		ob_start();
?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
<?php foreach ($this->sheets as $num => $sheet) { ?>
    <sheet name="<?=$sheet->getName();?>" sheetId="<?=$num + 1;?>" r:id="rId<?=$num + 2;?>" />
<?php } ?>
  </sheets>
</workbook>
<?php
		return ob_get_clean() ?: '';
	}

	/**
	 * XLSX basic info file
	 * @return string
	 */
	private function fileContentTypes(): string
	{
		$keys = array_keys($this->sheets);
		ob_start();
?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<?php foreach ($keys as $num) { ?>
  <Override PartName="/xl/worksheets/sheet<?=$num + 1;?>.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
<?php } ?>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>
<?php
		return ob_get_clean() ?: '';
	}

	/**
	 * Additional doc info
	 * @param string $type
	 * @param array<string, string> $data
	 * @return string
	 */
	private function fileRelationships(string $type, array $data): string
	{
		ob_start();
?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" Id="rId1" />
<?php foreach ($data as $id => $target) { ?>
  <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/<?=$type;?>" Target="<?=$target;?>" Id="<?=$id;?>" />
<?php } ?>
</Relationships>
<?php
		return ob_get_clean() ?: '';
	}

	/**
	 * Create sheet list for xlsx
	 * @return array<string, string>
	 */
	private function getSheetRelationships(): array
	{
		$data = [];
		$keys = array_keys($this->sheets);
		foreach ($keys as $num) {
			$data['rId'.($num + 2)] = '/xl/worksheets/sheet'.($num + 1).'.xml';
		}
		return $data;
	}

	/**
	 * Convert color format
	 * @param string $color		color in #format
	 * @return string
	 */
	public static function convertColor(string $color): string
	{
		return 'FF'.(strtoupper(substr($color, 1)));
	}

	/**
	 * Convert column number to xlsx alphanumeric
	 * @param int $num
	 * @return string
	 */
	public static function toAlpha(int $num): string
	{
		$alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
		if ($num < 26) return $alphabet[$num];

		$alpha = '';
		$dividend = $num + 1;

		while ($dividend > 0) {
			$left = ($dividend - 1) % 26;
			$alpha = $alphabet[$left].$alpha;
			$dividend = floor(($dividend - $left) / 26);
		}
		return $alpha;
	}
}

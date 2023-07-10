<?php
namespace SheetExporter;

use RuntimeException,
	InvalidArgumentException,
	ZipArchive;

/**
 * Export to Xlsx
 */
class ExporterXlsx extends Exporter {
	const FONT_SIZE = '10pt';
	const BORDER_MEDIUM = 2;
	const BORDER_THICK = 4;

	/**
	 * @param string $fileName
	 * @throws RuntimeException
	 */
	public function __construct (string $fileName) {
		if (!class_exists('ZipArchive')) throw new RuntimeException('Missing ZipArchive extension for XLSX.');
		parent::__construct($fileName);
	}

	/**
	 * Create download content
	 */
	public function download (): void {
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
	public function compile (): string {
		$zip = new ZipArchive;
		$tempFile = $this->createTemp();
		$res = $zip->open($tempFile, ZipArchive::CREATE|ZipArchive::OVERWRITE);
		if ($res === true) {
			$zip->addFromString('[Content_Types].xml', self::XML_HEADER.$this->fileContentTypes());
			$zip->addFromString('_rels/.rels', self::XML_HEADER.$this->fileRelationships('officeDocument', ['rs'.md5($this->fileName) => '/xl/workbook.xml']));
			$zip->addFromString('xl/_rels/workbook.xml.rels', self::XML_HEADER.$this->fileRelationships('worksheet', $this->getSheetRelationships() ));
			$zip->addFromString('xl/workbook.xml', self::XML_HEADER.$this->fileWorkbook());
			$zip->addFromString('xl/styles.xml', self::XML_HEADER.$this->fileStyles());

			foreach ($this->sheets as $num=>$sheet) {
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
	private function fileSheet (Sheet $sheet): string {
		ob_start();
		$mergeCells = [];
		$skipPlan = [];
		$line = 0;
?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetFormatPr defaultRowHeight="<?=isset($this->defaultStyle['HEIGHT']) && $this->defaultStyle['HEIGHT'] ? self::convertSize($this->defaultStyle['HEIGHT']) : '';?>" />
<?php
		if (count($sheet->getCols()) !== 0) {
?>
  <cols>
<?php
			foreach ($sheet->getCols() as $num=>$col) {
				echo '    <col collapsed="false" hidden="false" min="',($num + 1),'" max="',($num + 1),'" width="',$this->convertColSize($col),'" />',"\n";
			}
			if ($sheet->getDefCol() && count($sheet->getCols()) < $sheet->getColCount()) {
				echo '    <col collapsed="false" hidden="false" min="',(count($sheet->getCols()) + 1),
					'" max="',$sheet->getColCount(),'" width="',$this->convertColSize($sheet->getDefCol()),'" />',"\n";
			}
?>
  </cols>
<?php
		}
?>
  <sheetData>
<?php
		if (count($sheet->getHeaders()) !== 0) {
			echo '    <row r="'.++$line.'">';
			foreach ($sheet->getHeaders() as $num=>$name) echo '<c r="'.self::toAlpha($num).$line.'" t="inlineStr"><is><t>'.self::xmlEntities($name).'</t></is></c>';
			echo "</row>\n";
		}

		$styles = array_keys($this->styles);
		$styles = array_flip($styles);

		foreach ($sheet->getRows() as $num=>$row) {
            $current = 0;
			$class = $sheet->getStyle($num);
			if ($class && !isset($styles[$class])) throw new InvalidArgumentException('Missing style: '.htmlspecialchars($class, ENT_QUOTES));

			echo '    <row r="',++$line,'"',($class && isset($this->styles[$class]['HEIGHT']) && $this->styles[$class]['HEIGHT'] ? ' ht="'.self::convertSize($this->styles[$class]['HEIGHT']).'"' : ''),'>';
			foreach ($row as $j=>$col) {
				$move = 1;
//				$current += $skipPlan[$num][$j] ?? 0;

				// insert empty cells under merged
				if (isset($skipPlan[$num][$j])) {
					for ($i = 0; $i < $skipPlan[$num][$j]; $i++) {
						echo '<c r="'.self::toAlpha($current++).$line.'" />';
					}
				}

				if (is_array($col)) {
					if (isset($col['COLS']) && $col['COLS'] > 1 || isset($col['ROWS']) && $col['ROWS'] > 1) {
						$cols = $col['COLS'] ?? 1;
						$rows = $col['ROWS'] ?? 1;

						$mergeCells[self::toAlpha($current).$line] = self::toAlpha($current + $cols - 1).($line + $rows - 1);

						if ($rows > 1) {
							for ($i = 1; $i < $rows; $i++) {
								$skipPlan[$num + $i][$j] = $cols;
							}
						}

						$move = $cols;
					}

					if (isset($col['STYLE'])) {
						if (!isset($styles[$col['STYLE']])) throw new InvalidArgumentException('Missing style: '.htmlspecialchars($col['STYLE'], ENT_QUOTES));
						$class = $col['STYLE'];
					}
					echo $this->getColumn($current, $line, $col['VAL'], $class ? $styles[$class] + 1 : null);

					// insert empty cells under merged
					if ($move > 1) {
						for ($i = 1; $i < $move; $i++) {
							echo '<c r="'.self::toAlpha(++$current).$line.'"'.($class ? ' s="'.($styles[$class] + 1).'"' : '').' />';
						}
					}
				}
				else if ($col !== null) echo $this->getColumn($current, $line, $col, $class ? $styles[$class] + 1 : null);

                $current++;
				if (isset($skipPlan[$num])) unset($skipPlan[$num]);
			}

			echo "</row>\n";
		}
?>
  </sheetData>
<?php
		if (!empty($mergeCells)) {
?>
  <mergeCells count="<?=count($mergeCells);?>">
<?php
	foreach ($mergeCells as $from=>$to) {
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
		return ob_get_clean();
	}

	/**
	 * XLSX cell value
	 * @param int $num
	 * @param int $line
	 * @param string $val
	 * @param int|null $style
	 * @return string
	 */
	protected function getColumn (int $num, int $line, string $val, int $style = null): string {
		return '<c r="'.self::toAlpha($num).$line.'"'.($style ? ' s="'.$style.'"' : '').
			(is_numeric($val) && ctype_digit(substr($val, 0, 1)) ? ' t="n"><v>'.$val.'</v>' : ' t="inlineStr"><is><t>'.self::xmlEntities($val).'</t></is>').'</c>';
	}

	/**
	 * Try to convert metric units to ms unit
	 * @param string|float $size (optimal - mm)
	 * @return float
	 */
	protected function convertColSize ($size): float {
		return round((self::convertSize($size, self::UNITS, 'mm') / 1.852383) / (self::convertSize(isset($this->defaultStyle['FONT'],$this->defaultStyle['FONT']['SIZE']) ? $this->defaultStyle['FONT']['SIZE'] : self::FONT_SIZE) / 10), 8);
	}

	/**
	 * XLSX style
	 * @return string
	 */
	protected function fileStyles (): string {
		ob_start();
?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<?php
		if (!empty($this->defaultStyle) || !empty($this->styles)) {
			if (!isset($this->defaultStyle['FONT'],$this->defaultStyle['FONT']['SIZE'])) $this->defaultStyle['FONT']['SIZE'] = self::FONT_SIZE;
			if (!isset($this->defaultStyle['CELL'])) $this->defaultStyle['CELL'] = [];

			$d = $this->reshuffleStyles(array_merge([$this->defaultStyle], $this->styles));
?>
  <fonts count="<?=count($d['FONTS']);?>">
<?php
			foreach ($d['FONTS'] as $font) echo $this->getFontStyle($font);
?>
  </fonts>
  <fills count="<?=count($d['FILLS']) + 1;?>">
<?php
			foreach ($d['FILLS'] as $num=>$fill) {
?>
    <fill>
<?php
				if ($fill) {
?>
      <patternFill patternType="solid">
        <fgColor rgb="<?=self::convertColor($fill);?>" />
      </patternFill>
<?php
				}
				else {
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
  <borders count="<?=count($d['BORDERS']) ;?>">
<?php
			foreach ($d['BORDERS'] as $border) {
?>
    <border>
<?php
				foreach (self::$borderTypes as $key=>$mark) {
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
  <cellXfs count="<?=count($d['MAP']);?>">
<?php
			foreach ($d['MAP'] as $i=>$xf) {
				echo '    <xf numFmtId="0" fontId="', $xf['font'] ?? 0,'" fillId="',
						isset($xf['fill']) && $xf['fill'] > 0 ? $xf['fill'] + 1 : 0,'" borderId="', $xf['border'] ?? 0,'" xfId="0">';
				if (!empty($xf['ALIGN'])) echo "\n      ",'<alignment horizontal="'.$xf['ALIGN'].'" />',"\n    ";
				echo "</xf>\n";
			}
?>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normální" xfId="0" builtinId="0" />
  </cellStyles>
  <dxfs count="0" />
<?php
		}
?>
</styleSheet>
<?php
		return ob_get_clean();
	}

	/**
	 * Create font style for xlsx
	 * @param array<string, mixed> $font
	 * @return string
	 */
	private function getFontStyle (array $font): string {
		$t = "    <font>\n";
		if (isset($font['WEIGHT'])) $t .= '      <b />'."\n";
		if (isset($font['SIZE'])) $t .= '      <sz val="'.self::convertSize($font['SIZE']).'" />'."\n";
		if (isset($font['FAMILY'])) $t .= '      <name val="'.$font['FAMILY'].'" />'."\n";
		if (isset($font['COLOR'])) $t .= '      <color rgb="'.self::convertColor($font['COLOR']).'" />'."\n";
		return $t ."    </font>\n";
	}

	/**
	 * Create border style and convert border width to 3 types used in xlsx
	 * @param string $elm
	 * @param array<string, mixed> $border
	 * @return string
	 */
	private function getBorderStyle (string $elm, array $border = null): string {
		if (isset($border['STYLE']) && $border['STYLE'] === 'solid') $border['STYLE'] = 'thin';
		$style = $border['STYLE'] ?? null;
		if (isset($border['WIDTH'])) {
			if ($border['WIDTH'] >= self::BORDER_THICK) {
				if ($style) {
					if (in_array($style, ['dashed'])) $style = 'medium'.ucfirst($style);
				}
				else $style = 'thick';
			}
			else if ($border['WIDTH'] >= self::BORDER_MEDIUM) {
				if ($style) {
					if (in_array($style, ['dashed'])) $style = 'medium'.ucfirst($style);
				}
				else $style = 'medium';
			}
			else if ($border['WIDTH'] > 0 && !$style) $style = 'thin';
		}

		if ($style !== null) {
			return '      <'.$elm.' style="'.$style.'">
        <color rgb="'.self::convertColor($border['COLOR'] ?? static::$defColor).'" />
      </'.$elm.">\n";
		}
		return '      <'.$elm." />\n";
	}

	/**
	 * Prepare styles for xlsx structure
	 * @param array<int|string, mixed> $list
	 * @return array<string, mixed>
	 */
	private function reshuffleStyles (array $list): array {
		$map = [];
		$fonts = [];
		$fills = [];
		$borders = [];

		foreach ($list as $num=>$style) {
			$map[$num] = [];

			if (!empty($style['FONT'])) {
				if (!isset($style['FONT']['SIZE'])) $style['FONT']['SIZE'] = $this->defaultStyle['FONT']['SIZE'];
				if (!isset($style['FONT']['FAMILY']) && isset($this->defaultStyle['FONT']['FAMILY'])) $style['FONT']['FAMILY'] = $this->defaultStyle['FONT']['FAMILY'];
				$this->sortStyle($num, $style['FONT'], 'font', $fonts, $map);

				if (isset($style['FONT']['ALIGN'])) $map[$num]['ALIGN'] = $style['FONT']['ALIGN'];
			}
			if (isset($style['CELL'])) {
				$cell = $style['CELL'];
				$this->sortStyle($num, $cell['BACKGROUND'] ?? null, 'fill', $fills, $map);

				$item = [];
				if (isset($cell['COLOR']) || isset($cell['STYLE']) || isset($cell['WIDTH'])) {
					$cell = $this->reshuffleStyleBorder($cell);
				}

				if (isset($cell['LEFT']) || isset($cell['RIGHT']) || isset($cell['TOP']) || isset($cell['BOTTOM'])) {
					foreach (self::$borderTypes as $side) {
						if (!empty($cell[$side])) {
							$c =& $cell[$side];
							$item[$side] = [
								'COLOR' => $c['COLOR'] ?? static::$defColor,
								'STYLE' => $c['STYLE'] ?? null,
								'WIDTH' => isset($c['WIDTH']) ? self::convertSize($c['WIDTH']) : 1
							];
						}
					}
				}
				$this->sortStyle($num, $item, 'border', $borders, $map);
			}
		}
		return ['MAP'=>$map, 'FONTS'=>$fonts, 'FILLS'=>$fills, 'BORDERS'=>$borders];
	}

	/**
	 * Reformat style for borders
	 * @param array<string, mixed> $cell
	 * @return array<string, mixed>
	 */
	private function reshuffleStyleBorder (array $cell): array {
		foreach (['COLOR','STYLE','WIDTH'] as $type) {
			if (isset($cell[$type])) {
				foreach (self::$borderTypes as $side) $cell[$side][$type] = $cell[$type];
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
	private function sortStyle ($num, $item, string $type, array &$list, array &$map): void {
		$id = array_search($item, $list);
		if ($id === false) {
			$list[] = $item;
			end($list);
			$map[$num][$type] = key($list);
			reset($list);
		}
		else {
			$map[$num][$type] = $id;
		}
	}

	/**
	 * XLSX workbook info
	 * @return string
	 */
	private function fileWorkbook (): string {
		ob_start();
?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
<?php	foreach ($this->sheets as $num=>$sheet) {	?>
    <sheet name="<?=$sheet->getName();?>" sheetId="<?=$num + 1;?>" r:id="rId<?=$num + 2;?>" />
<?php	}	?>
  </sheets>
</workbook>
<?php
		return ob_get_clean();
	}

	/**
	 * XLSX basic info file
	 * @return string
	 */
	private function fileContentTypes (): string {
		ob_start();
?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<?php	foreach ($this->sheets as $num=>$sheet) {	?>
  <Override PartName="/xl/worksheets/sheet<?=$num + 1;?>.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
<?php	}	?>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>
<?php
		return ob_get_clean();
	}

	/**
	 * Additional doc info
	 * @param string $type
	 * @param array<string, string> $data
	 * @return string
	 */
	private function fileRelationships (string $type, array $data): string {
		ob_start();
?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" Id="rId1" />
<?php	foreach ($data as $id=>$target) {	?>
  <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/<?=$type;?>" Target="<?=$target;?>" Id="<?=$id;?>" />
<?php	}	?>
</Relationships>
<?php
		return ob_get_clean();
	}

	/**
	 * Create sheet list for xlsx
	 * @return array<string, string>
	 */
	private function getSheetRelationships (): array {
		$data = [];
		foreach ($this->sheets as $num=>$sheet) {
			$data['rId'.($num + 2)] = '/xl/worksheets/sheet'.($num + 1).'.xml';
		}
		return $data;
	}

	/**
	 * Convert color format
	 * @param string $color		color in #format
	 * @return string
	 */
	public static function convertColor (string $color): string {
		return 'FF'.(strtoupper(substr($color, 1)));
	}

	/**
	 * Convert column number to xlsx alphanumeric
	 * @param int $num
	 * @return string
	 */
	public static function toAlpha (int $num): string {
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
<?php
namespace SheetExporter;

/**
 * Export to Xlsx
 */
class ExporterXlsx extends Exporter {
	const FONT_SIZE = '10pt';
	const BORDER_MEDIUM = 2;
	const BORDER_THICK = 4;

	public function __construct ($fileName) {
		if (!class_exists('\ZipArchive')) throw new RuntimeException('Missing ZipArchive extension for XLSX.');
		parent::__construct($fileName);
	}

	public function download () {
		$tempFile = $this->compile();
		header('Content-Type: application/excel; charset=utf-8');
		header('Content-Disposition: attachment; filename="'.$this->fileName.'.xlsx"');
		readfile($tempFile);
		@unlink($tempFile);
	}

	public function compile () {
		$zip = new \ZipArchive;
		$tempFile = $this->createTemp();
		$res = $zip->open($tempFile, \ZipArchive::CREATE);
		if ($res === true) {
			$zip->addFromString('[Content_Types].xml', self::XML_HEADER.$this->fileContentTypes());
			$zip->addFromString('_rels/.rels', self::XML_HEADER.$this->fileRelationships('officeDocument', array('rs'.md5($this->fileName) => '/xl/workbook.xml') ));
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
	 * @throws \InvalidArgumentException
	 */
	private function fileSheet ($sheet) {
		ob_start();
		$megreCells = array();
		$spaces = array();
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
			$k = 0;
			$move = 0;
			$class = $sheet->getStyle($num);
			if ($class && !isset($styles[$class])) throw new \InvalidArgumentException('Missing style: '.htmlspecialchars($class, ENT_QUOTES));

			echo '    <row r="',++$line,'"',($class && isset($this->styles[$class]['HEIGHT']) && $this->styles[$class]['HEIGHT'] ? ' ht="'.self::convertSize($this->styles[$class]['HEIGHT']).'"' : ''),'>';
			foreach ($row as $col) {
				$num = $move + $k++;
				// insert empty cells under merged
				if (isset($spaces[0][$num])) {
					for ($i = 0; $i < $spaces[0][$num]; $i++) {
						echo '<c r="'.self::toAlpha($num + $i).$line.'" />';
					}
					$num += $spaces[0][$num];
				}
				if (is_array($col)) {
					// prepare merge cells and empty cells
					if (!empty($col['COLS']) && $col['COLS'] > 1) {
						if (!empty($col['ROWS']) && $col['ROWS'] > 1) {
							$megreCells[self::toAlpha($num).$line] = self::toAlpha($num + $col['COLS'] - 1).($line + $col['ROWS'] - 1);
							for ($i = 1; $i < $col['ROWS']; $i++) {
								$spaces[$i][$num] = $col['COLS'];
							}
						}
						else {
							$megreCells[self::toAlpha($num).$line] = self::toAlpha($num + $col['COLS'] - 1).$line;
						}
					}
					else if (!empty($col['ROWS']) && $col['ROWS'] > 1) {
						$megreCells[self::toAlpha($num).$line] = self::toAlpha($num).($line + $col['ROWS'] - 1);
						for ($i = 1; $i < $col['ROWS']; $i++) {
							$spaces[$i][$num] = 1;
						}
					}
					if (isset($col['STYLE'])) {
						if (!isset($styles[$col['STYLE']])) throw new \InvalidArgumentException('Missing style: '.htmlspecialchars($col['STYLE'], ENT_QUOTES));
						$class = $col['STYLE'];
					}
					echo $this->getColumn($num, $line, $col['VAL'], $class ? $styles[$class] + 1 : null);

					if (!empty($col['COLS']) && $col['COLS'] > 1) {
						for ($i = 1; $i < $col['COLS']; $i++) {
							echo '<c r="'.self::toAlpha($num + $i).$line.'"'.($class ? ' s="'.($styles[$class] + 1).'"' : '').' />';
						}
						$move += $col['COLS'] - 1;
					}
				}
				else if ($col !== null) echo $this->getColumn($num, $line, $col, $class ? $styles[$class] + 1 : null);
			}

			// remove used dummy cells
			if (!empty($spaces)) {
				if (isset($spaces[0])) unset($spaces[0]);
				if (!empty($spaces)) $spaces = array_values($spaces);
			}
			echo "</row>\n";
		}
?>
  </sheetData>
<?php
		if (!empty($megreCells)) {
?>
  <mergeCells count="<?=count($megreCells);?>">
<?php
	foreach ($megreCells as $from=>$to) echo '    <mergeCell ref="'.$from.':'.$to.'"/>';
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
	 * @param string $style
	 * @return string
	 */
	protected function getColumn ($num, $line, $val, $style = null) {
		return '<c r="'.self::toAlpha($num).$line.'"'.($style ? ' s="'.$style.'"' : '').
			(is_numeric($val) && ctype_digit(substr($val, 0, 1)) ? ' t="n"><v>'.$val.'</v>' : ' t="inlineStr"><is><t>'.self::xmlEntities($val).'</t></is>').'</c>';
	}

	/**
	 *
	 * @param string|float $size (optimal - mm)
	 * @return float
	 */
	protected function convertColSize ($size) {
		return round((self::convertSize($size, self::UNITS, 'mm') / 1.852383) / (self::convertSize(isset($this->defaultStyle['FONT'],$this->defaultStyle['FONT']['SIZE']) ? $this->defaultStyle['FONT']['SIZE'] : self::FONT_SIZE) / 10), 8);
	}

	/**
	 * XLSX style
	 * @return string
	 */
	protected function fileStyles () {
		ob_start();
?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<?php
		if (!empty($this->defaultStyle) || !empty($this->styles)) {
			if (!isset($this->defaultStyle['FONT'],$this->defaultStyle['FONT']['SIZE'])) $this->defaultStyle['FONT']['SIZE'] = self::FONT_SIZE;
			if (!isset($this->defaultStyle['CELL'])) $this->defaultStyle['CELL'] = array();

			$d = $this->resuffleStyles(array_merge(array($this->defaultStyle), $this->styles));
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
					echo $this->getBorderStyle($key, isset($border[$mark]) ? $border[$mark] : null);
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
				echo '    <xf numFmtId="0" fontId="',isset($xf['font']) ? $xf['font'] : 0,'" fillId="',
						isset($xf['fill']) && $xf['fill'] > 0 ? $xf['fill'] + 1 : 0,'" borderId="',isset($xf['border']) ? $xf['border'] : 0,'" xfId="0">';
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
	 * @param array $font
	 * @return string
	 */
	private function getFontStyle (array $font) {
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
	 * @param array $border
	 * @return string
	 */
	private function getBorderStyle ($elm, array $border = null) {
		if (isset($border['STYLE']) && $border['STYLE'] === 'solid') $border['STYLE'] = 'thin';
		$style = isset($border['STYLE']) ? $border['STYLE'] : null;
		if (isset($border['WIDTH'])) {
			if ($border['WIDTH'] >= self::BORDER_THICK) {
				if ($style) {
					if (in_array($style, array('dashed'))) $style = 'medium'.ucfirst($style);
				}
				else $style = 'thick';
			}
			else if ($border['WIDTH'] >= self::BORDER_MEDIUM) {
				if ($style) {
					if (in_array($style, array('dashed'))) $style = 'medium'.ucfirst($style);
				}
				else $style = 'medium';
			}
			else if ($border['WIDTH'] > 0 && !$style) $style = 'thin';
		}

		if ($style !== null) {
			return '      <'.$elm.' style="'.$style.'">
        <color rgb="'.self::convertColor($border['COLOR']).'" />
      </'.$elm.">\n";
		}
		return '      <'.$elm." />\n";
	}

	/**
	 * Prepare styles for xlsx structure
	 * @param array $list
	 * @return array
	 */
	private function resuffleStyles (array $list) {
		$map = array();
		$fonts = array();
		$fills = array();
		$borders = array();

		foreach ($list as $num=>$style) {
			$map[$num] = array();

			if (!empty($style['FONT'])) {
				if (!isset($style['FONT']['SIZE'])) $style['FONT']['SIZE'] = $this->defaultStyle['FONT']['SIZE'];
				if (!isset($style['FONT']['FAMILY']) && isset($this->defaultStyle['FONT']['FAMILY'])) $style['FONT']['FAMILY'] = $this->defaultStyle['FONT']['FAMILY'];
				$this->sortStyle($num, $style['FONT'], 'font', $fonts, $map);

				if (isset($style['FONT']['ALIGN'])) $map[$num]['ALIGN'] = $style['FONT']['ALIGN'];
			}
			if (isset($style['CELL'])) {
				$cell = $style['CELL'];
				$this->sortStyle($num, isset($cell['BACKGROUND']) ? $cell['BACKGROUND'] : null, 'fill', $fills, $map);

				$item = array();
				if (isset($cell['COLOR']) || isset($cell['STYLE']) || isset($cell['WIDTH'])) {
					$cell = $this->reshuffleStyleBorder($cell);
				}

				if (isset($cell['LEFT']) || isset($cell['RIGHT']) || isset($cell['TOP']) || isset($cell['BOTTOM'])) {
					foreach (self::$borderTypes as $side) {
						if (!empty($cell[$side])) {
							$c =& $cell[$side];
							$item[$side] = array(
								'COLOR' => isset($c['COLOR']) ? $c['COLOR'] : '#000000',
								'STYLE' => isset($c['STYLE']) ? $c['STYLE'] : null,
								'WIDTH' => isset($c['WIDTH']) ? self::convertSize($c['WIDTH']) : 1
							);
						}
					}
				}
				$this->sortStyle($num, $item, 'border', $borders, $map);
			}
		}
		return array('MAP'=>$map, 'FONTS'=>$fonts, 'FILLS'=>$fills, 'BORDERS'=>$borders);
	}

	/**
	 *
	 * @param array $cell
	 * @return array
	 */
	private function reshuffleStyleBorder (array $cell) {
		foreach (array('COLOR','STYLE','WIDTH') as $type) {
			if (isset($cell[$type])) {
				foreach (self::$borderTypes as $side) $cell[$side][$type] = $cell[$type];
				unset($cell[$type]);
			}
		}
		return $cell;
	}

	/**
	 *
	 * @param int $num
	 * @param array $item
	 * @param string $type
	 * @param array $list
	 * @param array $map
	 */
	private function sortStyle ($num, $item, $type, array &$list, array &$map) {
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
	private function fileWorkbook () {
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
	 *
	 * @return string
	 */
	private function fileContentTypes () {
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
	 * @param array $data
	 * @return string
	 */
	private function fileRelationships ($type, array $data) {
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
	 * @return array
	 */
	private function getSheetRelationships () {
		$data = array();
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
	public static function convertColor ($color) {
		return 'FF'.(strtoupper(substr($color, 1)));
	}

	/**
	 * Convert column number to xlsx alphanumeric
	 * @param int $num
	 * @return string
	 */
	public static function toAlpha ($num) {
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
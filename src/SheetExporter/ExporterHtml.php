<?php
namespace SheetExporter;

/**
 * Export to HTML file
 */
class ExporterHtml extends Exporter {

	public function download () {
		header('Content-Type: text/html; charset=utf-8');
		header('Content-Disposition: attachment; filename="'.$this->fileName.'.html"');
		$this->compile();
	}

	public function compile () {
?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" lang="cs" xml:lang="cs">
<head>
	<title>Export <?=$this->fileName;?></title>
	<meta charset="utf-8" />
	<meta name="generator" content="<?=$this->getVersion();?>" />
	<style>
		td.number { text-align: right; }
<?php
		if (!empty($this->defaultStyle)) {
			echo "\t\tth,td { ".implode('; ', $this->getStyle($this->defaultStyle))." }\n";
		}
		foreach ($this->styles as $mark=>$style) {
			echo "\t\t.".$mark." { ".implode('; ', $this->getStyle($style))." }\n";
		}
?>
	</style>
</head>
<body>
<?php
		foreach ($this->sheets as $sheet) {
?>
	<table border="1" cellspacing="0" cellpadding="5">
		<caption><?=$sheet->getName();?></caption>
<?php
		if (count($sheet->getCols()) !== 0) {
			foreach ($sheet->getCols() as $width) {
?>
		<col style="width: <?=self::convertSize($width, self::UNITS).self::UNITS;?>" />
<?php
			}
		}

		if ($sheet->getDefCol() && count($sheet->getCols()) < $sheet->getColCount()) {
			$width = self::convertSize($sheet->getDefCol(), self::UNITS).self::UNITS;
			for ($i = count($sheet->getCols()); $i < $sheet->getColCount(); $i++) {
?>
		<col style="width: <?=$width;?>" />
<?php
			}
		}
		if (count($sheet->getHeaders()) !== 0) {
?>
		<thead>
			<tr>
				<?php foreach ($sheet->getHeaders() as $name) echo '<th>'.self::xmlEntities($name).'</th>';?>
			</tr>
		</thead>
<?php
		}
?>
		<tbody>
<?php

			$skipRow = array();
			foreach ($sheet->getRows() as $num=>$row) {
				$skipNow = $skipRow;
				$class = $sheet->getStyle($num);
				if ($class && !isset($this->styles[$class])) throw new \InvalidArgumentException('Missing style: '.htmlspecialchars($class, ENT_QUOTES));

				echo "\t\t\t<tr>";
				if (self::$overlay) {
					// hide
					foreach ($row as $i=>$col) {
						if (!isset($skipNow[$i]) || $skipNow[$i] === 0) {
							if (is_array($col)) echo $this->getColumn($col['VAL'], isset($col['STYLE']) ? $col['STYLE'] : $class, $col);
							else echo $this->getColumn($col, $class);
						}
						if (is_array($col) && !empty($col['ROWS'])) {
							$cols = !empty($col['COLS']) ? $col['COLS'] : 1;
							for ($j = 0; $j < $cols; $j++) {
								if (!isset($skipRow[$i + $j])) $skipRow[$i + $j] = 0;
								$skipRow[$i + $j] += $col['ROWS'];
							}
						}
					}
					foreach ($skipRow as $i=>$none) {
						if ($skipRow[$i] > 0) $skipRow[$i]--;
					}
				}
				else {
					// render everything
					foreach ($row as $i=>$col) {
						if (is_array($col)) echo $this->getColumn($col['VAL'], isset($col['STYLE']) ? $col['STYLE'] : $class, $col);
						else echo $this->getColumn($col, $class);
					}
				}
				echo "</tr>\n";
			}
?>
		</tbody>
	</table>
<?php
		}
?>
</body>
</html>
<?php
		return true;
	}

	/**
	 *
	 * @param string $val
	 * @param string $class
	 * @param array $col
	 * @return string
	 */
	protected function getColumn ($val, $class = '', array $col = null) {
		$isNum = is_numeric($val);
		return '<td'.($class || $isNum ? ' class="'.$class.($isNum ? ' number' : '').'"' : '').
				(isset($col['ROWS']) && $col['ROWS'] > 1 ? ' rowspan="'.$col['ROWS'].'"' : '').
				(isset($col['COLS']) && $col['COLS'] > 1 ? ' colspan="'.$col['COLS'].'"' : '').
				'>'.self::xmlEntities($val).'</td>';
	}

	/**
	 *
	 * @param array $style
	 * @return array
	 */
	private function getStyle (array $style) {
		$data = array();
		if (isset($style['FONT']['COLOR'])) $data[] = 'color: '.$style['FONT']['COLOR'];
		if (isset($style['FONT']['SIZE'])) $data[] = 'font-size: '.self::convertSize($style['FONT']['SIZE'], self::UNITS).self::UNITS;
		if (isset($style['FONT']['FAMILY'])) $data[] = 'font-family: "'.$style['FONT']['FAMILY'].'"';
		if (isset($style['FONT']['WEIGHT'])) $data[] = 'font-weight: '.$style['FONT']['WEIGHT'];
		if (isset($style['FONT']['ALIGN'])) $data[] = 'text-align: '.$style['FONT']['ALIGN'];

		if (isset($style['CELL']['BACKGROUND'])) $data[] = 'background-color: '.$style['CELL']['BACKGROUND'];

		if (isset($style['CELL']['COLOR']) || isset($style['CELL']['STYLE']) || isset($style['CELL']['WIDTH'])) {
			$data = $this->getBorderStyle($data, 'border', $style['CELL']);
		}
		else {
			if (!empty($style['CELL']['LEFT'])) $data = $this->getBorderStyle($data, 'border-left', $style['CELL']['LEFT']);
			if (!empty($style['CELL']['RIGHT'])) $data = $this->getBorderStyle($data, 'border-right', $style['CELL']['RIGHT']);
			if (!empty($style['CELL']['TOP'])) $data = $this->getBorderStyle($data, 'border-top', $style['CELL']['TOP']);
			if (!empty($style['CELL']['BOTTOM'])) $data = $this->getBorderStyle($data, 'border-bottom', $style['CELL']['BOTTOM']);
		}

		if (isset($style['HEIGHT']) && $style['HEIGHT'] !== null) $data[] = 'height: '.self::convertSize($style['HEIGHT'], self::UNITS).self::UNITS;
		return $data;
	}

	/**
	 *
	 * @param array $data
	 * @param string $side
	 * @param array $style
	 * @return array
	 */
	private function getBorderStyle (array $data, $side, array $style = null) {
		if (isset($style['COLOR'])) $data[] = $side.'-color: '.$style['COLOR'];
		if (isset($style['STYLE'])) $data[] = $side.'-style: '.$style['STYLE'];
		if (isset($style['WIDTH'])) $data[] = $side.'-width: '.self::convertSize($style['WIDTH'], self::UNITS).self::UNITS;
		return $data;
	}
}
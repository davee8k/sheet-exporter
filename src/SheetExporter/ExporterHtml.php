<?php
declare(strict_types=1);

namespace SheetExporter;

use InvalidArgumentException;

/**
 * Export to HTML file
 */
class ExporterHtml extends Exporter
{
	/**
	 * Create download content
	 */
	public function download(): void
	{
		header('Content-Type: text/html; charset=utf-8');
		header('Content-Disposition: attachment; filename="'.$this->fileName.'.html"');
		$this->compile();
	}

	/**
	 * Generate html code
	 * @return string|null
	 * @throws InvalidArgumentException
	 */
	public function compile(): ?string
	{
?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" lang="cs" xml:lang="cs">
<head>
	<title>Export <?=$this->fileName;?></title>
	<meta charset="utf-8" />
	<meta name="generator" content="<?=$this->getVersion();?>" />
	<style>
		td { border-color: <?=static::$defColor;?>; }
		td.number { text-align: right; }
<?php
		if (!empty($this->defaultStyle)) {
			echo "\t\tth,td { ".implode('; ', $this->getStyle($this->defaultStyle))." }\n";
		}
		foreach ($this->styles as $mark => $style) {
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
		$this->printColWidths($sheet);

		$this->printHeader($sheet);
?>
		<tbody>
<?php
		$this->printSheet($sheet);
?>
		</tbody>
	</table>
<?php
		}
?>
</body>
</html>
<?php
		return null;
	}

	/**
	 * Print sheet header
	 * @param Sheet $sheet
	 */
	private function printHeader(Sheet $sheet): void
	{
		if (!empty($sheet->getHeaders())) {
?>
		<thead>
			<tr>
				<?php foreach ($sheet->getHeaders() as $name) echo '<th>'.self::xmlEntities($name).'</th>';?>
			</tr>
		</thead>
<?php
		}
	}

	/**
	 * Print sheet content
	 * @param Sheet $sheet
	 */
	private function printSheet(Sheet $sheet): void
	{
		foreach ($sheet->getRows() as $num => $row) {
			$class = $sheet->getStyle($num);
			if ($class && !isset($this->styles[$class])) {
				throw new InvalidArgumentException('Missing style: '.htmlspecialchars($class, ENT_QUOTES));
			}

			echo "\t\t\t<tr>";
			// render everything
			foreach ($row as $col) {
				if (is_array($col)) echo $this->getCell($col['VAL'] ?? '', $col['STYLE'] ?? $class, $col);
				else echo $this->getCell($col, $class);
			}
			echo "</tr>\n";
		}
	}

	/**
	 * Print columns widths
	 * @param Sheet $sheet
	 */
	private function printColWidths(Sheet $sheet): void
	{
		$count = count($sheet->getCols());
		if ($count !== 0) {
			foreach ($sheet->getCols() as $width) {
?>
		<col style="width: <?=self::convertSize($width, self::UNITS).self::UNITS;?>" />
<?php
			}
		}

		if ($sheet->getDefCol() && $count < $sheet->getColCount()) {
			$width = self::convertSize($sheet->getDefCol(), self::UNITS).self::UNITS;
			for ($i = $count; $i < $sheet->getColCount(); $i++) {
?>
		<col style="width: <?=$width;?>" />
<?php
			}
		}
	}

	/**
	 * Get table cell
	 * @param string|float|int|null $val
	 * @param string|null $class
	 * @param array<string, mixed>|null $col
	 * @return string
	 */
	protected function getCell($val, ?string $class = null, ?array $col = null): string
	{
		if ($val === null) return '';
		if (is_numeric($val)) $class = ($class === null ? '' : $class.' ').'number';
		return '<td'.($class ? ' class="'.$class.'"' : '').
				(isset($col['ROWS']) && $col['ROWS'] > 1 ? ' rowspan="'.$col['ROWS'].'"' : '').
				(isset($col['COLS']) && $col['COLS'] > 1 ? ' colspan="'.$col['COLS'].'"' : '').
				'>'.self::xmlEntities($val).'</td>';
	}

	/**
	 * Format cell style
	 * @param array<string, mixed> $style
	 * @return string[]
	 */
	private function getStyle(array $style): array
	{
		$data = [];
		if (!empty($style['FONT'])) $data = $this->getStyleFont($data, $style['FONT']);
		if (!empty($style['CELL'])) $data = $this->getStyleCell($data, $style['CELL']);

		if (isset($style['HEIGHT']) && $style['HEIGHT'] !== null) {
			$data[] = 'height: '.self::convertSize($style['HEIGHT'], self::UNITS).self::UNITS;
		}
		return $data;
	}

	/**
	 * Returns font style
	 * @param string[] $data
	 * @param array<string, mixed> $style
	 * @return string[]
	 */
	private function getStyleFont(array $data, array $style): array
	{
		if (isset($style['COLOR'])) $data[] = 'color: '.$style['COLOR'];
		if (isset($style['SIZE'])) $data[] = 'font-size: '.self::convertSize($style['SIZE'], self::UNITS).self::UNITS;
		if (isset($style['FAMILY'])) $data[] = 'font-family: "'.$style['FAMILY'].'"';
		if (isset($style['WEIGHT'])) $data[] = 'font-weight: '.$style['WEIGHT'];
		if (isset($style['ALIGN'])) $data[] = 'text-align: '.$style['ALIGN'];
		return $data;
	}

	/**
	 * Returns cells border and background
	 * @param string[] $data
	 * @param array<string, mixed> $style
	 * @return string[]
	 */
	private function getStyleCell(array $data, array $style): array
	{
		if (isset($style['BACKGROUND'])) $data[] = 'background-color: '.$style['BACKGROUND'];

		if ($this->isBorderStyle($style, self::$borderStyles)) {
			$data = $this->getStyleBorder($data, 'border', $style);
		} else {
			foreach (self::$borderTypes as $key => $mark) {
				if (!empty($style[$mark])) $data = $this->getStyleBorder($data, 'border-'.$key, $style[$mark]);
			}
		}
		return $data;
	}

	/**
	 * Format border style settings
	 * @param string[] $data
	 * @param string $side
	 * @param array<string, mixed> $style
	 * @return string[]
	 */
	private function getStyleBorder(array $data, string $side, array $style = null): array
	{
		if (isset($style['COLOR'])) {
			$data[] = $side.'-color: '.$style['COLOR'];
		}
		if (isset($style['STYLE'])) {
			$data[] = $side.'-style: '.$style['STYLE'];
		}
		if (isset($style['WIDTH'])) {
			$data[] = $side.'-width: '.self::convertSize($style['WIDTH'], self::UNITS).self::UNITS;
		}
		return $data;
	}
}

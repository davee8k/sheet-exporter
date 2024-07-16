<?php
declare(strict_types=1);

namespace SheetExporter;

use InvalidArgumentException;

/**
 * Simple XLSX/ODS/HTML table exporter
 * XLSX and ODS require ZipArchive class
 *
 * Basic style support for:
 * $font -[ COLOR, SIZE, FAMILY, WEIGHT, ALIGN ]
 * $cell (possible LEFT, RIGHT, TOP, BOTTOM) - [ BACKGROUND, COLOR, WIDTH, STYLE ]
 *
 * @author DaVee8k
 * @license https://unlicense.org/
 * @version 0.87.4
 */
abstract class Exporter
{
	/** @var float */
	public const VERSION = 0.87;
	/** @var string */
	protected const UNITS = 'pt';
	/** @var string */
	protected const XML_HEADER = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n";

	/** @var array<string, string> */
	public static $borderTypes = ['left' => 'LEFT', 'right' => 'RIGHT', 'top' => 'TOP', 'bottom' => 'BOTTOM'];
	/** @var string[] */
	public static $borderStyles = ['COLOR', 'STYLE', 'WIDTH'];
	/** @var string */
	public static $defColor = '#000000';
	/** @var string */
	protected $fileName;
	/** @var array<string, mixed> */
	protected $defaultStyle = [];
	/** @var Sheet[] */
	protected $sheets = [];
	/** @var array<string, mixed> */
	protected $styles = [];

	/**
	 * Set output file name
	 * @param string $fileName
	 */
	public function __construct(string $fileName)
	{
		$name = filter_var($fileName, FILTER_SANITIZE_URL);
		if ($name === false || $name === '') throw new InvalidArgumentException('Nonexistent name');
		$this->fileName = $name;
	}

	/**
	 * Download file
	 */
	public abstract function download(): void;

	/**
	 * Build and print(plain text) or return(link on zip file) content
	 */
	public abstract function compile(): ?string;

	/**
	 * Return version
	 * @return string
	 */
	public function getVersion(): string
	{
		return 'SheetExporter '.Exporter::VERSION;
	}

	/**
	 * Set default style
	 * @param array<string, mixed> $font
	 * @param array<string, mixed> $cell
	 * @param int|null $height
	 */
	public function setDefault(array $font = [], array $cell = [], int $height = null): void
	{
		$this->defaultStyle = ['FONT' => $font, 'CELL' => $this->prepareBorder($cell), 'HEIGHT' => $height];
	}

	/**
	 * Insert new content style
	 * @param string $mark
	 * @param array<string, mixed> $font
	 * @param array<string, mixed> $cell
	 * @param int|null $height
	 * @throws InvalidArgumentException
	 */
	public function addStyle(string $mark, array $font = [], array $cell = [], int $height = null): void
	{
		if (!preg_match('/^[a-z0-9\-]+$/', $mark))
				throw new InvalidArgumentException("Style mark must by small alphanumeric only.");
		$this->styles[$mark] = ['FONT' => $font, 'CELL' => $this->prepareBorder($cell), 'HEIGHT' => $height];
	}

	/**
	 * Insert sheet, throws error if name exist
	 * @param Sheet $sheet
	 * @throws InvalidArgumentException
	 */
	public function addSheet(Sheet $sheet): void
	{
		$name = $sheet->getName();
		if ($this->checkUniqueName($name) != $name) throw new InvalidArgumentException("Sheet name already exists.");
		$this->sheets[] = $sheet;
	}

	/**
	 * Create new data sheet, if name already exist rename
	 * @param string $name
	 * @return Sheet
	 */
	public function insertSheet(string $name = 'List'): Sheet
	{
		$sheet = new Sheet($this->checkUniqueName(Sheet::filterName($name)));
		$this->sheets[] = $sheet;
		return $sheet;
	}

	/**
	 * Return array of Sheet
	 * @return Sheet[]
	 */
	public function getSheets(): array
	{
		return $this->sheets;
	}

	/**
	 * Create a temporary file in the temporary
	 * @return string
	 * @throws InvalidArgumentException
	 */
	protected function createTemp(): string
	{
		$tmpDir = ini_get('upload_tmp_dir');
		$tmpFile = tempnam($tmpDir ?: sys_get_temp_dir(), $this->fileName);
		if (!$tmpFile) throw new InvalidArgumentException('Failed to create temporary file');
		return $tmpFile;
	}

	/**
	 * Check if chosen name is unique - if not change it
	 * @param string $newName
	 * @param int $count
	 * @return string
	 */
	protected function checkUniqueName(string $newName, int $count = 0): string
	{
		foreach ($this->sheets as $sheet) {
			if ($sheet->getName() == $newName.($count ? '_'.$count : '')) {
				return $this->checkUniqueName($newName, ++$count);
			}
		}
		return $newName.($count ? '_'.$count : '');
	}

	/**
	 * Convert border settings
	 * @param array<string, mixed> $cell
	 * @return array<string, mixed>
	 */
	protected function prepareBorder(array $cell): array
	{
		if (empty($cell)) return [];

		if ($this->isBorderStyle($cell, self::$borderStyles) && $this->isBorderStyle($cell, self::$borderTypes)) {
			foreach (self::$borderStyles as $type) {
				if (isset($cell[$type])) {
					foreach (self::$borderTypes as $side) {
						if (!isset($cell[$side]) || !isset($cell[$side][$type])) $cell[$side][$type] = $cell[$type];
					}
					unset($cell[$type]);
				}
			}
		}
		return $cell;
	}

	/**
	 * Checks border style existence
	 * @param array<string, mixed> $cell
	 * @param array<string|int, string> $params
	 * @return bool
	 */
	protected function isBorderStyle(array $cell, array $params): bool
	{
		foreach ($params as $mark) {
			if (isset($cell[$mark])) return true;
		}
		return false;
	}

	/**
	 * Convert XML entities
	 * @param string|float|int|null $val
	 * @return string
	 */
	public static function xmlEntities($val): string
	{
		if (is_numeric($val) || !$val) return (string) $val;
		return strtr($val, ['<' => '&lt;', '>' => '&gt;', '"' => '&quot;', "'" => '&apos;', '&' => '&amp;']);
	}

	/**
	 * Convert between different measure systems
	 * @param string|float|int $num	Value
	 * @param string $def	Measure unit
	 * @param string|null $out	Measure unit
	 * @return float
	 * @throws InvalidArgumentException
	 */
	public static function convertSize($num, string $def = 'pt', string $out = null): float
	{
		if ($out === null) $out = $def;

		if (is_numeric($num)) {
			if ($def != $out) $num .= $def;
			else return (float) $num;
		} elseif (preg_match('/'.preg_quote($out, '/').'$/i', $num)) {
			return floatval(str_replace($out, '', $num));
		}

		if (preg_match('/[a-z]+$/i', $num, $match)) {
			$units = $match[0];
			$val = self::convertFromUnit($units, floatval(str_replace($units, '', $num)));
			return self::convertToUnit($out, $val);
		}
		throw new InvalidArgumentException("Unknown measure value: ".htmlspecialchars($num, ENT_QUOTES));
	}

	/**
	 *
	 * @param string $unit
	 * @param float $val
	 * @return float
	 * @throws InvalidArgumentException
	 */
	private static function convertFromUnit(string $unit, float $val): float
	{
		switch ($unit) {
			case 'em':
			case 'px': return $val;
			case 'pt': return $val *= 96 / 72;
			case 'pc': return $val *= 16;
			case 'in': return $val *= 96;
			case 'mm': return $val *= 3.78;
			case 'cm': return $val *= 37.8;
		}
		throw new InvalidArgumentException("Unknown input measure unit: ".htmlspecialchars($unit, ENT_QUOTES));
	}

	/**
	 *
	 * @param string $unit
	 * @param float $val
	 * @return float
	 * @throws InvalidArgumentException
	 */
	private static function convertToUnit(string $unit, float $val): float
	{
		switch ($unit) {
			case 'em': return $val;
			case 'px': return round($val, 2);
			case 'pt': return round($val / (96 / 72), 2);
			case 'pc': return round($val / 16, 4);
			case 'in': return round($val / 96, 4);
			case 'mm': return round($val / 3.78, 6);
			case 'cm': return round($val / 37.8, 6);
		}
		throw new InvalidArgumentException("Unknown output measure unit: ".htmlspecialchars($unit, ENT_QUOTES));
	}
}

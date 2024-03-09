<?php declare(strict_types=1);

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
 * @version 0.87.3
 */
abstract class Exporter {
	/** @var float */
	public const VERSION = 0.87;
	/** @var string */
	protected const UNITS = 'pt';
	/** @var string */
	protected const XML_HEADER = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n";

	/** @var array<string, string> */
	public static $borderTypes = ['left'=>'LEFT','right'=>'RIGHT','top'=>'TOP','bottom'=>'BOTTOM'];
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
	public function __construct (string $fileName) {
		$name = filter_var($fileName, FILTER_SANITIZE_URL);
		if ($name === false || $name === '') throw new InvalidArgumentException('Nonexistent name');
		$this->fileName = $name;
	}

	/**
	 * Download file
	 */
	public abstract function download (): void;

	/**
	 * Build and print(plain text) or return(link on zip file) content
	 */
	public abstract function compile (): ?string;

	/**
	 * Return version
	 * @return string
	 */
	public function getVersion (): string {
		return 'SheetExporter '.Exporter::VERSION;
	}

	/**
	 * Set default style
	 * @param array<string, mixed> $font
	 * @param array<string, mixed> $cell
	 * @param int|null $height
	 */
	public function setDefault (array $font = [], array $cell = [], int $height = null): void {
		$this->defaultStyle = ['FONT'=>$font, 'CELL'=>$this->prepareBorder($cell), 'HEIGHT'=>$height];
	}

	/**
	 * Insert new content style
	 * @param string $mark
	 * @param array<string, mixed> $font
	 * @param array<string, mixed> $cell
	 * @param int|null $height
	 * @throws InvalidArgumentException
	 */
	public function addStyle (string $mark, array $font = [], array $cell = [], int $height = null): void {
		if (!preg_match('/^[a-z0-9\-]+$/', $mark)) throw new InvalidArgumentException("Style mark must by small alphanumeric only.");
		$this->styles[$mark] = ['FONT'=>$font, 'CELL'=>$this->prepareBorder($cell), 'HEIGHT'=>$height];
	}

	/**
	 * Insert sheet, throws error if name exist
	 * @param Sheet $sheet
	 * @throws InvalidArgumentException
	 */
	public function addSheet (Sheet $sheet): void {
		$name = $sheet->getName();
		if ($this->checkUniqueName($name) != $name) throw new InvalidArgumentException("Sheet name already exists.");
		$this->sheets[] = $sheet;
	}

	/**
	 * Create new data sheet, if name already exist rename
	 * @param string $name
	 * @return Sheet
	 */
	public function insertSheet (string $name = 'List'): Sheet {
		$sheet = new Sheet($this->checkUniqueName(Sheet::filterName($name)));
		$this->sheets[] = $sheet;
		return $sheet;
	}

	/**
	 * Return array of Sheet
	 * @return Sheet[]
	 */
	public function getSheets (): array {
		return $this->sheets;
	}

	/**
	 * Create a temporary file in the temporary
	 * @return string
	 * @throws InvalidArgumentException
	 */
	protected function createTemp ():string {
		$tmpDir = ini_get('upload_tmp_dir');
		$tmpFile = tempnam($tmpDir ?: sys_get_temp_dir(), $this->fileName);
		if (!$tmpFile) throw new InvalidArgumentException('Failed to create temporary file');
		return $tmpFile;
	}

	/**
	 * Check if chosen name is unique - if not change it
	 * @param string $newName
	 * @param int $i
	 * @return string
	 */
	protected function checkUniqueName (string $newName, int $i = 0): string {
		foreach ($this->sheets as $sheet) {
			if ($sheet->getName() == $newName.($i ? '_'.$i : '')) {
				return $this->checkUniqueName($newName, ++$i);
			}
		}
		return $newName.($i ? '_'.$i : '');
	}

	/**
	 * Convert border settings
	 * @param array<string, mixed> $cell
	 * @return array<string, mixed>
	 */
	protected function prepareBorder (array $cell): array {
		if (empty($cell)) return [];

		$global = isset($cell['COLOR']) || isset($cell['STYLE']) || isset($cell['WIDTH']);
		$local = isset($cell['LEFT']) || isset($cell['RIGHT']) || isset($cell['TOP']) || isset($cell['BOTTOM']);
		if ($global && $local) {
			foreach (['COLOR','STYLE','WIDTH'] as $type) {
				if (isset($cell[$type])) {
					foreach (self::$borderTypes as $side) {
						if (!isset($cell[$side]) || (!empty($cell[$side]) && !isset($cell[$side][$type]))) $cell[$side][$type] = $cell[$type];
					}
					unset($cell[$type]);
				}
			}
		}
		return $cell;
	}

	/**
	 * Convert XML entities
	 * @param string|float|int|null $val
	 * @return string
	 */
	public static function xmlEntities ($val): string {
		if (is_numeric($val) || !$val) return (string) $val;
		return strtr($val, ['<'=>'&lt;', '>'=>'&gt;','"'=>'&quot;', "'"=>'&apos;','&'=>'&amp;']);
	}

	/**
	 * Convert between different measure systems
	 * @param string|float $num	Value
	 * @param string $def	Measure unit
	 * @param string|null $out	Measure unit
	 * @return float
	 * @throws InvalidArgumentException
	 */
	public static function convertSize ($num, string $def = 'pt', string $out = null): float {
		if ($out === null) $out = $def;
		if (is_numeric($num)) {
			if ($def != $out) $num .= $def;
			else return (float)$num;
		}
		else if (preg_match('/'.preg_quote($out, '/').'$/i', $num)) return floatval(str_replace($out, '', $num));

		if (preg_match('/[a-z]+$/i', $num, $match)) {
			$units = $match[0];
			$val = floatval(str_replace($units, '', $num));

			switch ($units) {
				case 'em': return $val;
				case 'px': break;
				case 'pt': $val *= 96/72; break;
				case 'pc': $val *= 16; break;
				case 'in': $val *= 96; break;
				case 'mm': $val *= 3.78; break;
				case 'cm': $val *= 37.8; break;
				default: throw new InvalidArgumentException("Unknown input measure unit: ".htmlspecialchars($units, ENT_QUOTES));
			}

			switch ($out) {
				case 'px': return round($val, 2);
				case 'pt': return round($val / (96/72), 2);
				case 'pc': return round($val / 16, 4);
				case 'in': return round($val / 96, 4);
				case 'mm': return round($val / 3.78, 6);
				case 'cm': return round($val / 37.8, 6);
				default: throw new InvalidArgumentException("Unknown output measure unit: ".htmlspecialchars($out, ENT_QUOTES));
			}
		}
		throw new InvalidArgumentException("Unknown measure value: ".htmlspecialchars($num, ENT_QUOTES));
	}
}

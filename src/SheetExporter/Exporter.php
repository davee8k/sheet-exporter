<?php
namespace SheetExporter;

use InvalidArgumentException;

/**
 * Simple XLSX/ODS/HTML table exporter
 * XLSX and ODS require ZipArchive class
 * basic style support for
 * $font -[ COLOR, SIZE, FAMILY, WEIGHT, ALIGN ]
 * $cell (possible LEFT, RIGHT, TOP, BOTTOM) - [ BACKGROUND, COLOR, WIDTH, STYLE ]
 *
 * @author DaVee8k
 * @license https://unlicense.org/
 * @version 0.85.53
 */
abstract class Exporter {
	const VERSION = 0.855;
	const UNITS = 'pt';
	const XML_HEADER = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n";

	/** @var array */
	public static $borderTypes = array('left'=>'LEFT','right'=>'RIGHT','top'=>'TOP','bottom'=>'BOTTOM');
	/** @var bool */
	public static $overlay = false;

	/** @var string */
	protected $fileName;
	/** @var array */
	protected $defaultStyle = array();
	/** @var Sheet[] */
	protected $sheets = array();
	/** @var array */
	protected $styles = array();

	/**
	 * Set output file name
	 * @param string $fileName
	 */
	public function __construct ($fileName) {
		$this->fileName = filter_var($fileName, FILTER_SANITIZE_URL);
	}

	/**
	 * Download file
	 */
	abstract public function download ();

	/**
	 * Build and retrun content
	 */
	abstract public function compile ();

	/**
	 * Return version
	 * @return string
	 */
	public function getVersion () {
		return 'SheetExporter '.Exporter::VERSION;
	}

	/**
	 * Set default style
	 * @param array $font
	 * @param array $cell
	 * @param int|null $height
	 */
	public function setDefault ($font = array(), $cell = array(), $height = null) {
		$this->defaultStyle = array('FONT'=>$font, 'CELL'=>$this->prepareBorder($cell), 'HEIGHT'=>$height);
	}

	/**
	 * Insert new content style
	 * @param string $mark
	 * @param array $font
	 * @param array $cell
	 * @param int|null $height
	 * @throws InvalidArgumentException
	 */
	public function addStyle ($mark, $font = array(), $cell = array(), $height = null) {
		if (!preg_match('/^[a-z0-9\-]+$/', $mark)) throw new InvalidArgumentException("Style mark must by small alfanumeric only.");
		$this->styles[$mark] = array('FONT'=>$font, 'CELL'=>$this->prepareBorder($cell), 'HEIGHT'=>$height);
	}

	/**
	 * Insert sheet, throws error if name exist
	 * @param Sheet $sheet
	 * @throws InvalidArgumentException
	 */
	public function addSheet (Sheet $sheet) {
		$name = $sheet->getName();
		if ($this->checkUniqueName($name) != $name) throw new InvalidArgumentException("Sheet name already exists.");
		$this->sheets[] = $sheet;
	}

	/**
	 * Create new data sheet, if name already exist rename
	 * @param string $name
	 * @return Sheet
	 */
	public function insertSheet ($name = 'List') {
		$sheet = new Sheet($this->checkUniqueName(Sheet::filterName($name)));
		$this->sheets[] = $sheet;
		return $sheet;
	}

	/**
	 * Return array of Sheet
	 * @return array
	 */
	public function getSheets () {
		return $this->sheets;
	}

	/**
	 * Create a temporary file in the temporary
	 * @return false|string
	 */
	protected function createTemp () {
		$temDir = ini_get('upload_tmp_dir');
		return tempnam($temDir ? $temDir : sys_get_temp_dir(), $this->fileName);
	}

	/**
	 * Check if choosed name is unique - if not change it
	 * @param string $newName
	 * @param int $i
	 * @return string
	 */
	protected function checkUniqueName ($newName, $i = 0) {
		foreach ($this->sheets as $sheet) {
			if ($sheet->getName() == $newName.($i ? '_'.$i : '')) {
				return $this->checkUniqueName($newName, ++$i);
			}
		}
		return $newName.($i ? '_'.$i : '');
	}

	/**
	 * Convert border settings
	 * @param array $cell
	 * @return array
	 */
	protected function prepareBorder ($cell) {
		if (empty($cell)) return array();

		$global = isset($cell['COLOR']) || isset($cell['STYLE']) || isset($cell['WIDTH']);
		$local = isset($cell['LEFT']) || isset($cell['RIGHT']) || isset($cell['TOP']) || isset($cell['BOTTOM']);
		if ($global && $local) {
			foreach (array('COLOR','STYLE','WIDTH') as $type) {
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
	 * @param string $string
	 * @return string
	 */
	public static function xmlEntities ($string) {
		return $string ? strtr($string, array('<'=>'&lt;', '>'=>'&gt;','"'=>'&quot;', "'"=>'&apos;','&'=>'&amp;')) : $string;
	}

	/**
	 * Convert between diferent measure systems
	 * @param string|float $num	value
	 * @param string $def	measure unit
	 * @param string $out	measure unit
	 * @return string
	 * @throws InvalidArgumentException
	 */
	public static function convertSize ($num, $def = 'pt', $out = null) {
		if ($out === null) $out = $def;
		if (is_numeric($num)) {
			if ($def != $out) $num .= $def;
			else return $num;
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
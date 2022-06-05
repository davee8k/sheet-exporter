<?php
namespace SheetExporter;

/**
 * Sheet page in document
 */
class Sheet {
	/** @var string */
	protected $name;
	/** @var string|int */
	protected $colCount = 0;
	/** @var string|float */
	protected $colDefault = null;
	/** @var array */
	protected $cols = array();
	/** @var array */
	protected $headers = array();
	/** @var array */
	protected $rows = array();
	/** @var array */
	protected $styles = array();

	/**
	 * Create sheet with name
	 * @param string $name
	 */
	public function __construct ($name) {
		$this->name = self::filterName($name);
	}

	/**
	 * Add sheet header
	 * @param string $column
	 */
	public function addColHeader ($column) {
		$this->headers[] = $column;
		if (count($this->headers) > $this->colCount) $this->colCount = count($this->headers);
	}

	/**
	 * Set sheet headers
	 * @param array $headers
	 */
	public function setColHeaders (array $headers) {
		$this->headers = $headers;
		if (count($this->headers) > $this->colCount) $this->colCount = count($this->headers);
	}

	/**
	 * Set width of the columns
	 * @param array $widths
	 * @param string|float $default
	 */
	public function setColWidth (array $widths, $default = null) {
		$this->cols = $widths;
		$this->colDefault = $default;
	}

	/**
	 * Add row with style
	 * @param array $row
	 * @param string $style
	 */
	public function addRow (array $row, $style = null) {
		$this->rows[] = $row;
		if (count($row) > $this->colCount) $this->colCount = count($row);
		if ($style !== null) $this->styles[count($this->rows) - 1] = $style;
	}

	/**
	 * Add 2d array
	 * @param array $array
	 * @param string $style
	 */
	public function addArray (array $array, $style = null) {
		foreach ($array as $row) $this->addRow($row, $style);
	}

	/**
	 * Return sheet name
	 * @return string
	 */
	public function getName () {
		return $this->name;
	}

	/**
	 * Return default column width
	 * @return string|float
	 */
	public function getDefCol () {
		return $this->colDefault;
	}

	/**
	 * Return width of the columns
	 * @return array
	 */
	public function getCols () {
		return $this->cols;
	}

	/**
	 * Return max column count
	 * @return int
	 */
	public function getColCount () {
		return $this->colCount;
	}

	/**
	 * Return headers
	 * @return array
	 */
	public function getHeaders () {
		return $this->headers;
	}

	/**
	 * Return all rows
	 * @return array
	 */
	public function getRows () {
		return $this->rows;
	}

	/**
	 * Return style for selected row
	 * @param int $num
	 * @return string|null
	 */
	public function getStyle ($num) {
		return isset($this->styles[$num]) ? $this->styles[$num] : null;
	}

	/**
	 * Normalize sheet name
	 * @param string $name
	 * @return string
	 */
	public static function filterName ($name) {
		return strip_tags($name);
	}
}
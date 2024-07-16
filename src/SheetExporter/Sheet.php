<?php
declare(strict_types=1);

namespace SheetExporter;

/**
 * Sheet page in document
 */
class Sheet
{
	/** @var string */
	protected $name;
	/** @var int */
	protected $colCount = 0;
	/** @var string|float|null */
	protected $colDefault = null;
	/** @var array<string|float> */
	protected $cols = [];
	/** @var string[] */
	protected $headers = [];
	/** @var array<int, mixed[]> */
	protected $rows = [];
	/** @var array<int, string> */
	protected $styles = [];

	/**
	 * Create sheet with name
	 * @param string $name
	 */
	public function __construct(string $name)
	{
		$this->name = self::filterName($name);
	}

	/**
	 * Add sheet header
	 * @param string $column
	 */
	public function addColHeader(string $column): void
	{
		$this->headers[] = $column;
		if (count($this->headers) > $this->colCount) $this->colCount = count($this->headers);
	}

	/**
	 * Set sheet headers
	 * @param string[] $headers
	 */
	public function setColHeaders(array $headers): void
	{
		$this->headers = $headers;
		if (count($this->headers) > $this->colCount) $this->colCount = count($this->headers);
	}

	/**
	 * Set width of the columns
	 * @param array<string|float> $widths
	 * @param string|float $default
	 */
	public function setColWidth(array $widths, $default = null): void
	{
		$this->cols = $widths;
		$this->colDefault = $default;
	}

	/**
	 * Add row with style
	 * @param mixed[] $row
	 * @param string|null $style
	 */
	public function addRow(array $row, string $style = null): void
	{
		$this->rows[] = $row;
		if (count($row) > $this->colCount) $this->colCount = count($row);
		if ($style !== null) $this->styles[count($this->rows) - 1] = $style;
	}

	/**
	 * Add 2d array
	 * @param array<mixed[]> $array
	 * @param string|null $style
	 */
	public function addArray(array $array, string $style = null): void
	{
		foreach ($array as $row) {
			$this->addRow($row, $style);
		}
	}

	/**
	 * Return sheet name
	 * @return string
	 */
	public function getName(): string
	{
		return $this->name;
	}

	/**
	 * Return default column width
	 * @return string|float|null
	 */
	public function getDefCol()
	{
		return $this->colDefault;
	}

	/**
	 * Return width of the columns
	 * @return array<string|float>
	 */
	public function getCols(): array
	{
		return $this->cols;
	}

	/**
	 * Return max column count
	 * @return int
	 */
	public function getColCount(): int
	{
		return $this->colCount;
	}

	/**
	 * Return headers
	 * @return string[]
	 */
	public function getHeaders(): array
	{
		return $this->headers;
	}

	/**
	 * Return all rows
	 * @return array<int, mixed[]>
	 */
	public function getRows(): array
	{
		return $this->rows;
	}

	/**
	 * Return style for selected row
	 * @param int $num
	 * @return string|null
	 */
	public function getStyle(int $num): ?string
	{
		return $this->styles[$num] ?? null;
	}

	/**
	 * Normalize sheet name
	 * @param string $name
	 * @return string
	 */
	public static function filterName(string $name): string
	{
		return strip_tags($name);
	}
}

<?php declare(strict_types=1);

namespace SheetExporter;

use InvalidArgumentException;

/**
 * Export to CSV file
 */
class ExporterCsv extends Exporter {

	/**
	 * Create download content
	 */
	public function download (): void {
		header('Content-Type: text/csv; charset=utf-8');
		header('Content-Disposition: attachment; filename="'.$this->fileName.'.csv"');
		$this->compile();
	}

	/**
	 * Generate csv code
	 * @return string|null
	 * @throws InvalidArgumentException
	 */
	public function compile (): ?string {
		foreach ($this->sheets as $sheet) {
			if (count($sheet->getHeaders()) !== 0) {
				foreach ($sheet->getHeaders() as $name) {
					echo $this->getColumn($name).',';
				}
				echo "\n";
			}
			$skipPlan = [];
			foreach ($sheet->getRows() as $num=>$row) {
				foreach ($row as $j=>$col) {
					if (isset($skipPlan[$num][$j])) {
						echo str_repeat(' ,', $skipPlan[$num][$j]);
					}
					if (is_array($col)) {
						echo $this->getColumn($col['VAL']).',';
						if (isset($col['ROWS']) && $col['ROWS'] > 1) {
							for ($i = 1; $i < $col['ROWS']; $i++) {
								$skipPlan[$num + $i][$j] = $col['COLS'] ?? 1;
							}
						}
						if (isset($col['COLS']) && $col['COLS'] > 1) {
							echo str_repeat(' ,', $col['COLS'] - 1);
						}
					}
					else echo $this->getColumn($col).',';
				}
				echo "\n";
			}
		}
		return null;
	}

	/**
	 * Escape csv data
	 * @param string|float|int|null $val
	 * @return string
	 */
	private function getColumn ($val): string {
		if (is_numeric($val)) return (string)$val;
		if (!$val) return '';
		$str = str_replace('"', '""', $val);
		if (strpos($str, ',') !== false || strpos($str, '"') !== false) return '"'.$str.'"';
		return $str;
	}
}

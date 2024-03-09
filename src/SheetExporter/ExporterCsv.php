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
					echo $this->getCellValue($name).',';
				}
				echo "\n";
			}
			$skipPlan = [];
			foreach ($sheet->getRows() as $num=>$row) {
				$last = -1;
				$move = 0;
				for ($j = 0; $j <= $move + $sheet->getColCount(); $j++) {
					if (isset($skipPlan[$num][$j])) {
						$move += $skipPlan[$num][$j];
						echo str_repeat(' ,', $skipPlan[$num][$j]);
					}
					if (isset($row[$j - $move]) && $last < ($j - $move)) {
						$last = $j - $move;
						echo $this->getCell($row[$last], $num, $j, $skipPlan);
					}
				}
				echo "\n";

				if (isset($skipPlan[$num])) unset($skipPlan[$num]);
			}
		}
		return null;
	}

	/**
	 * Get csv cell
	 * @param array<string, mixed>|string|float|int|null $col
	 * @param int $num
	 * @param int $j
	 * @param array<int, array<int, int>> $skipPlan
	 * @return string
	 */
	private function getCell ($col, int $num, int $j, array &$skipPlan): string {
		if (is_array($col)) {
			if (isset($col['ROWS']) && $col['ROWS'] > 1) {
				for ($i = 1; $i < $col['ROWS']; $i++) {
					$skipPlan[$num + $i][$j] = $col['COLS'] ?? 1;
				}
			}
			if (isset($col['COLS']) && $col['COLS'] > 1) {
				$skipPlan[$num][$j + 1] = $col['COLS'] - 1;
			}
			return $this->getCellValue($col['VAL'] ?? '').',';
		}
		return $this->getCellValue($col).',';
	}

	/**
	 * Escape csv cell value
	 * @param string|float|int|null $val
	 * @return string
	 */
	private function getCellValue ($val): string {
		if (is_numeric($val)) return (string)$val;
		if (!$val) return '';
		$str = str_replace('"', '""', $val);
		if (strpos($str, ',') !== false || strpos($str, '"') !== false) return '"'.$str.'"';
		return $str;
	}
}

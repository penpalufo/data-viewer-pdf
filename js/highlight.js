// 日付は時刻で比較し、それ以外は型も含めて比較します。
export function valuesAreEqual(currentValue, previousValue) {
	if (currentValue instanceof Date && previousValue instanceof Date) {
		return currentValue.getTime() === previousValue.getTime()
	}
	return Object.is(currentValue, previousValue)
}

// 対象列の現在値が、検索結果内の直前行と違うか判定します。
export function isDifferentFromPrevious(filteredRows, selectedColumns, resultIndex, columnIndex) {
	if (resultIndex === 0 || !selectedColumns.includes(columnIndex)) return false
	return !valuesAreEqual(
		filteredRows[resultIndex].row[columnIndex],
		filteredRows[resultIndex - 1].row[columnIndex]
	)
}

// チェックされた列を選択状態へ追加、または選択状態から削除します。
export function toggleSelectedColumn(selectedColumns, columnIndex, checked) {
	const selected = new Set(selectedColumns)
	if (checked) {
		selected.add(columnIndex)
	} else {
		selected.delete(columnIndex)
	}
	return [...selected].sort((a, b) => a - b)
}

// 現在のシートにあるすべての列番号を返します。
export function getAllColumnIndexes(headers) {
	return headers.map((_, index) => index)
}

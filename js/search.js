import { displayValue } from './utils.js'

// 入力された文字列から検索用の正規表現を作成します。
export function createSearchPattern(searchText, caseSensitive) {
	if (!searchText) return null

	try {
		return new RegExp(searchText, caseSensitive ? '' : 'i')
	} catch {
		return null
	}
}

// 正規表現の入力に誤りがある場合、画面へ表示する文言を返します。
export function getSearchError(searchText, searchPattern) {
	if (!searchText || searchPattern) return ''
	return '正規表現が正しくありません'
}

// 固定2条件をAND評価し、両方に一致する行だけを取り出します。
// 入力が空の条件は無視し、不正な正規表現がある場合は絞り込みません。
export function filterRows(rows, conditions) {
	const indexedRows = rows.map((row, sourceIndex) => ({ row, sourceIndex }))
	const activeConditions = conditions.filter(condition => condition.searchText)
	if (!activeConditions.length) return indexedRows
	if (activeConditions.some(condition => !condition.searchPattern)) return indexedRows

	return indexedRows.filter(({ row }) => {
		return activeConditions.every(condition => {
			condition.searchPattern.lastIndex = 0
			return condition.searchPattern.test(displayValue(row[condition.columnIndex]))
		})
	})
}

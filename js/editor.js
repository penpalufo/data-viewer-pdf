// クリックしたセルについて、モーダル表示に必要な情報をまとめます。
export function createEditingCell(rowIndex, columnIndex, address, header) {
	return {
		rowIndex,
		columnIndex,
		address,
		header
	}
}

// 編集内容を数値・真偽値・文字列のいずれかへ変換します。
export function parseEditedValue(text) {
	const value = text.trim()
	if (value === '') return ''
	if (/^-?(?:\d+|\d*\.\d+)$/.test(value)) return Number(value)
	if (/^(true|false)$/i.test(value)) return value.toLowerCase() === 'true'
	return value
}

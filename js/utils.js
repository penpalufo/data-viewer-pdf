// 行ごとのセル数を、見出しの列数に揃えます。
export function normalizeRow(row, columnCount) {
	return Array.from({ length: columnCount }, (_, index) => row[index] ?? '')
}

// セルの値を画面表示用の文字列へ変換します。
export function displayValue(value) {
	if (value === null || value === undefined) return ''
	if (value instanceof Date) {
		return new Intl.DateTimeFormat('ja-JP').format(value)
	}
	return String(value)
}

// 件数を日本語環境の桁区切り形式で表示します。
export function formatNumber(value) {
	return new Intl.NumberFormat('ja-JP').format(value)
}

// 掲載ページの値を、電子カタログで使用するページ番号へ変換します。
export function getCatalogPageUrl(value) {
	if (value === null || value === undefined || String(value).trim() === '') return ''

	const publicationPage = Number(value)
	if (!Number.isInteger(publicationPage) || publicationPage < 1) return ''

	const catalogPage = publicationPage % 2 === 0
		? publicationPage / 2 + 1
		: (publicationPage + 1) / 2 + 1

	return `https://gmd.okamura.jp/iportal/cv.do?c=28444570000&pg=${catalogPage}&v=OKM05`
}

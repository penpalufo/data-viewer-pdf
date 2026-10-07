import { normalizeRow } from './utils.js'

// ExcelファイルをSheetJSで読み込み、ワークブックを返します。
export async function readExcelFile(file, XLSX) {
	const buffer = await file.arrayBuffer()
	return XLSX.read(buffer, {
		type: 'array',
		cellDates: true,
		cellStyles: true,
		cellFormula: true
	})
}

// 指定したシートを、見出しと行データへ変換します。
export function getSheetTable(workbook, sheetName, XLSX) {
	const sheet = workbook?.Sheets[sheetName]
	if (!sheet || !sheet['!ref']) return { headers: [], rows: [] }

	const range = XLSX.utils.decode_range(sheet['!ref'])
	const matrix = XLSX.utils.sheet_to_json(sheet, {
		header: 1,
		defval: '',
		raw: true,
		range: range.s.r
	})

	const columnCount = range.e.c - range.s.c + 1
	const headerRow = matrix[0] || []
	const headers = Array.from({ length: columnCount }, (_, index) => {
		const value = headerRow[index]
		return String(value ?? '').trim() || `列${index + 1}`
	})
	const rows = matrix.slice(1).map(row => normalizeRow(row, columnCount))
	return { headers, rows }
}

// Excel上の行番号と列番号からセル番地を返します。
export function getCellAddress(rowIndex, columnIndex, XLSX) {
	return XLSX.utils.encode_cell({ r: rowIndex + 1, c: columnIndex })
}

// JavaScriptの値からSheetJS用のセルデータを作成します。
function createCell(value) {
	if (typeof value === 'number') return { t: 'n', v: value }
	if (typeof value === 'boolean') return { t: 'b', v: value }
	return { t: 's', v: String(value) }
}

// 既存の書式情報を残しながら、セルの値や数式を書き換えます。
function writeCellValue(sheet, address, value) {
	const cell = { ...(sheet[address] || {}) }
	delete cell.v
	delete cell.w
	delete cell.f

	if (value === '') {
		cell.t = 'z'
		sheet[address] = cell
		return
	}

	if (typeof value === 'string' && value.startsWith('=')) {
		cell.t = 'n'
		cell.f = value.slice(1)
		sheet[address] = cell
		return
	}

	Object.assign(cell, createCell(value))
	sheet[address] = cell
}

// 編集先が現在のシート範囲外の場合、シート範囲を広げます。
function expandSheetRange(sheet, rowIndex, columnIndex, XLSX) {
	const range = sheet['!ref']
		? XLSX.utils.decode_range(sheet['!ref'])
		: { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }
	range.e.r = Math.max(range.e.r, rowIndex)
	range.e.c = Math.max(range.e.c, columnIndex)
	sheet['!ref'] = XLSX.utils.encode_range(range)
}

// 編集値を指定シートへ書き戻し、編集したセル番地を返します。
export function updateWorkbookCell(workbook, sheetName, rowIndex, columnIndex, value, XLSX) {
	const sheet = workbook.Sheets[sheetName]
	const address = getCellAddress(rowIndex, columnIndex, XLSX)
	writeCellValue(sheet, address, value)
	expandSheetRange(sheet, rowIndex + 1, columnIndex, XLSX)
	return address
}

// 全シートを保持したワークブックを編集済みExcelとして保存します。
export function saveWorkbookFile(workbook, fileName, XLSX) {
	const baseName = fileName.replace(/\.(xlsx|xls)$/i, '')
	XLSX.writeFile(workbook, `${baseName}_編集済み.xlsx`, {
		bookType: 'xlsx',
		compression: true,
		cellStyles: true
	})
}

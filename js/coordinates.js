// Excelの列番号を、座標JSONで使用するA、B、...、AA形式へ変換します。
export function getExcelColumnName(columnIndex) {
	let value = columnIndex + 1
	let name = ''
	while (value > 0) {
		value -= 1
		name = String.fromCharCode(65 + value % 26) + name
		value = Math.floor(value / 26)
	}
	return name
}

function toPoint(value) {
	const point = Array.isArray(value) ? value : value?.point
	if (!Array.isArray(point) || point.length !== 2) return null
	const [x, y] = point.map(Number)
	return Number.isFinite(x) && Number.isFinite(y) ? [x, y] : null
}

function toSize(value, fallback) {
	const size = Array.isArray(value) && value.length === 2 ? value : fallback
	if (!Array.isArray(size) || size.length !== 2) return null
	const normalized = size.map(Number)
	return normalized.every(number => Number.isFinite(number) && number > 0) ? normalized : null
}

function getCatalogPageFromPdfName(pdfName) {
	const match = String(pdfName || '').match(/(\d{3,4})(?=\D*\.pdf$)/i)
	return match ? Number(match[1]) : null
}

// v2の単一PDF形式と、セルごとにPDF情報を持てる統合形式を共通の索引へ正規化します。
export function createCoordinateIndex(data) {
	if (!data || ![2, 3].includes(Number(data.v)) || !Array.isArray(data.sheets)) {
		throw new Error('座標JSONの形式またはバージョンに対応していません。')
	}
	if (data.coordinate_system && data.coordinate_system !== 'pdf_points_top_left') {
		throw new Error('座標JSONの座標系に対応していません。')
	}
	if (data.point_rule && data.point_rule !== 'bbox_center') {
		throw new Error('座標JSONの座標ルールに対応していません。')
	}

	const defaultSize = toSize(data.size, null)
	const sheets = new Map()
	for (const sheet of data.sheets) {
		if (!sheet || typeof sheet.name !== 'string' || !Array.isArray(sheet.records)) continue
		const records = new Map()
		for (const record of sheet.records) {
			if (!Number.isInteger(record?.row) || !record.cells || typeof record.cells !== 'object') continue
			const cells = new Map()
			for (const [column, value] of Object.entries(record.cells)) {
				const point = toPoint(value)
				if (!point) continue
				const metadata = Array.isArray(value) ? {} : value
				const pdf = String(metadata.pdf ?? record.pdf ?? sheet.pdf ?? data.pdf ?? '')
				cells.set(column.toUpperCase(), {
					x: point[0],
					y: point[1],
					page: Number(metadata.page ?? record.page ?? sheet.page ?? data.page) || 1,
					sourcePage: Number(metadata.source_page ?? record.source_page ?? sheet.source_page ?? data.source_page) || getCatalogPageFromPdfName(pdf),
					pdf,
					size: toSize(metadata.size ?? record.size ?? sheet.size, defaultSize)
				})
			}
			records.set(record.row, cells)
		}
		sheets.set(sheet.name, records)
	}

	return { sheets, workbook: String(data.workbook || '') }
}

// 表のセルに対応するPDF上の中心座標と、その座標が属する元ページ情報を返します。
export function getCellCoordinate(index, sheetName, sourceIndex, columnIndex) {
	return index?.sheets.get(sheetName)?.get(sourceIndex + 2)?.get(getExcelColumnName(columnIndex)) || null
}


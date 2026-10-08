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

// 座標JSONの基本形式を確認し、シート名と行番号から引ける索引を作ります。
export function createCoordinateIndex(data, expectedNames) {
	if (!data || data.v !== 2 || !Array.isArray(data.sheets)) {
		throw new Error('座標JSONの形式またはバージョンに対応していません。')
	}
	if (data.coordinate_system !== 'pdf_points_top_left' || data.point_rule !== 'bbox_center') {
		throw new Error('座標JSONの座標系に対応していません。')
	}
	if (!Array.isArray(data.size) || data.size.length !== 2) {
		throw new Error('座標JSONにPDFサイズがありません。')
	}
	if (data.pdf && data.pdf !== expectedNames.pdf) {
		throw new Error(`座標JSONが参照するPDFは ${data.pdf} です。`)
	}
	// workbookは生成時の元ファイル名を記録した参考情報です。
	// 配布時に同一ベース名へ改名される場合があるため、照合の必須条件にはしません。

	const sheets = new Map()
	for (const sheet of data.sheets) {
		if (!sheet || typeof sheet.name !== 'string' || !Array.isArray(sheet.records)) continue
		const records = new Map()
		for (const record of sheet.records) {
			if (!Number.isInteger(record?.row) || !record.cells || typeof record.cells !== 'object') continue
			records.set(record.row, record)
		}
		sheets.set(sheet.name, records)
	}

	return {
		page: Number(data.page) || 1,
		size: data.size.map(Number),
		sheets
	}
}

// 表のセルに対応するPDF上の中心座標を返します。
export function getCellCoordinate(index, sheetName, sourceIndex, columnIndex) {
	const record = index?.sheets.get(sheetName)?.get(sourceIndex + 2)
	const point = record?.cells?.[getExcelColumnName(columnIndex)]
	if (!Array.isArray(point) || point.length !== 2) return null
	const [x, y] = point.map(Number)
	return Number.isFinite(x) && Number.isFinite(y) ? { x, y, page: index.page } : null
}


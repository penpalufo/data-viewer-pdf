/* Excel Data Viewer
	* Version history:
	* 0.10.1 (2026-10-08) Fitted the workspace to the browser height with internal scrolling.
	* 0.10.0 (2026-10-08) Added linked PDF display and coordinate markers.
	* 0.3.3 (2026-09-29) Added catalog links to publication page cells.
	* 0.3.2 (2026-09-29) Applied value-change highlighting within filtered results.
	* 0.3.1 (2026-09-29) Added a fixed two-condition AND regex search.
	* 0.3.0 (2026-09-29) Split JavaScript features into ES Modules.
	* 0.2.3 (2026-09-28) Added selectable column highlighting for value changes.
	* 0.2.2 (2026-09-28) Added a modal for viewing and editing complete cell values.
	* 0.2.1 (2026-09-28) Added explanatory comments to the source files.
	* 0.2.0 (2026-09-28) Sheet selection, cell editing, workbook export.
	* 0.1.0 (2026-09-28) Initial release: Excel import and regex search.
	*/

import {
	getCellAddress,
	getSheetTable,
	readExcelFile,
	saveWorkbookFile,
	updateWorkbookCell
} from './js/excel.js'
import { createSearchPattern, filterRows, getSearchError } from './js/search.js'
import { createEditingCell, parseEditedValue } from './js/editor.js'
import {
	getAllColumnIndexes,
	isDifferentFromPrevious,
	toggleSelectedColumn
} from './js/highlight.js'
import { displayValue, formatNumber, getCatalogPageUrl } from './js/utils.js'
import { createCoordinateIndex, getCellCoordinate } from './js/coordinates.js'
import { loadPdfDocument, renderPdfPage } from './js/pdf.js?v=0.10.0-fix1'
import { findSiblingFile, getDroppedFiles } from './js/files.js?v=0.10.0-folder'

// アプリ全体で使用するバージョン情報と表示件数の上限です。
const APP_VERSION = '0.10.1'
const RELEASE_DATE = '2026-10-08'
const DISPLAY_LIMIT = 5000

// Vue 3のOptions APIで、画面の状態と各機能モジュールを連携します。
Vue.createApp({
	// 画面表示やExcel処理で使用するデータです。
	data() {
		return {
			appVersion: APP_VERSION,
			releaseDate: RELEASE_DATE,
			displayLimit: DISPLAY_LIMIT,
			fileName: '',
			workbook: null,
			sheetNames: [],
			selectedSheetName: '',
			headers: [],
			rows: [],
			selectedHeaderIndex1: 0,
			searchText1: '',
			selectedHeaderIndex2: 0,
			searchText2: '',
			caseSensitive: false,
			isDragging: false,
			isLoading: false,
			errorMessage: '',
			editedCellKeys: new Set(),
			highlightColumnsBySheet: {},
			editingCell: null,
			editingValue: '',
			pdfDocument: null,
			pdfFileName: '',
			pdfPageNumber: 1,
			pdfPageCount: 0,
			pdfRenderSize: null,
			pdfMarker: null,
			coordinateIndex: null,
			pdfStatusMessage: 'PDF連動なし',
			selectedRowSourceIndex: null,
			resizeTimer: null
		}
	},

	// 他のデータから自動計算される、検索結果や編集件数です。
	computed: {
		// 検索条件1の入力から、検索に使用する正規表現を作ります。
		searchPattern1() {
			return createSearchPattern(this.searchText1, this.caseSensitive)
		},

		// 検索条件1の正規表現に誤りがあれば、エラーメッセージを返します。
		searchError1() {
			return getSearchError(this.searchText1, this.searchPattern1)
		},

		// 検索条件2の入力から、検索に使用する正規表現を作ります。
		searchPattern2() {
			return createSearchPattern(this.searchText2, this.caseSensitive)
		},

		// 検索条件2の正規表現に誤りがあれば、エラーメッセージを返します。
		searchError2() {
			return getSearchError(this.searchText2, this.searchPattern2)
		},

		// 2つの検索条件をAND評価し、一致した行を返します。
		filteredRows() {
			return filterRows(this.rows, [
				{
					columnIndex: Number(this.selectedHeaderIndex1),
					searchText: this.searchText1,
					searchPattern: this.searchPattern1
				},
				{
					columnIndex: Number(this.selectedHeaderIndex2),
					searchText: this.searchText2,
					searchPattern: this.searchPattern2
				}
			])
		},

		// 描画上限までの検索結果を、画面に表示する行として返します。
		visibleRows() {
			return this.filteredRows.slice(0, this.displayLimit)
		},

		// 現在までに編集したセルの数を返します。
		editedCellCount() {
			return this.editedCellKeys.size
		},

		// 選択中のシートで、変化を強調する列番号を返します。
		highlightedColumnIndexes() {
			return this.highlightColumnsBySheet[this.selectedSheetName] || []
		},

		// 変化を強調する列の選択状態を、画面表示用の文言にします。
		highlightSelectionLabel() {
			const count = this.highlightedColumnIndexes.length
			return count ? `${count}列選択中` : '未選択'
		},

		// PDFと座標JSONの両方を正常に読み込めたかを返します。
		hasPdfViewer() {
			return Boolean(this.pdfDocument && this.coordinateIndex)
		},

		pdfStatus() {
			return this.hasPdfViewer ? 'PDF連動中' : this.pdfStatusMessage
		},

		pdfStageStyle() {
			if (!this.pdfRenderSize) return {}
			return {
				width: `${this.pdfRenderSize.width}px`,
				height: `${this.pdfRenderSize.height}px`
			}
		},

		pdfMarkerStyle() {
			if (!this.pdfMarker || !this.pdfRenderSize || !this.coordinateIndex) return {}
			const scaleX = this.pdfRenderSize.width / this.coordinateIndex.size[0]
			const scaleY = this.pdfRenderSize.height / this.coordinateIndex.size[1]
			return {
				left: `${this.pdfMarker.x * scaleX}px`,
				top: `${this.pdfMarker.y * scaleY}px`
			}
		}
	},

	mounted() {
		window.addEventListener('resize', this.handleWindowResize)
	},

	beforeUnmount() {
		window.removeEventListener('resize', this.handleWindowResize)
		clearTimeout(this.resizeTimer)
		this.pdfDocument?.destroy()
	},

	// 画面イベントを受け取り、機能モジュールの処理を呼び出します。
	methods: {
		// ドラッグ中のファイルが読込領域から離れたとき、強調表示を解除します。
		handleDragLeave(event) {
			if (!event.currentTarget.contains(event.relatedTarget)) this.isDragging = false
		},

		// 読込領域へドロップされたExcelファイルを受け取ります。
		async handleDrop(event) {
			this.isDragging = false
			try {
				this.loadDataSet(await getDroppedFiles(event.dataTransfer))
			} catch {
				this.errorMessage = 'ドロップされたフォルダーを読み込めませんでした。'
			}
		},

		// ファイル選択ボタンで選ばれたExcelファイルを受け取ります。
		handleFileSelect(event) {
			this.loadDataSet([...event.target.files])
			event.target.value = ''
		},

		// 選択ファイルからExcelを探し、同名のPDFとJSONも読み込み対象にします。
		loadDataSet(files) {
			const excelFiles = files.filter(file => /\.(xlsx|xls)$/i.test(file.name))
			if (!excelFiles.length) {
				this.errorMessage = 'Excelファイル（.xlsx または .xls）を選択してください。'
				return
			}
			if (excelFiles.length > 1) {
				this.errorMessage = 'Excelファイルが複数あります。対象のExcelが1つだけ入ったフォルダーを選択してください。'
				return
			}
			this.loadExcel(excelFiles[0], files)
		},

		// Excelファイルを読み込み、シート一覧と最初のシートを準備します。
		async loadExcel(file, selectedFiles = [file]) {
			this.errorMessage = ''

			if (!/\.(xlsx|xls)$/i.test(file.name)) {
				this.errorMessage = 'Excelファイル（.xlsx または .xls）を選択してください。'
				return
			}

			this.isLoading = true
			await this.resetPdfViewer()
			try {
				const workbook = await readExcelFile(file, XLSX)
				if (!workbook.SheetNames.length) throw new Error('シートが見つかりませんでした。')

				this.fileName = file.name
				this.workbook = workbook
				this.sheetNames = [...workbook.SheetNames]
				this.selectedSheetName = this.sheetNames[0]
				this.editedCellKeys = new Set()
				this.highlightColumnsBySheet = {}
				this.selectedRowSourceIndex = null
				this.loadSelectedSheet()
				await this.loadPdfCompanions(file, selectedFiles)
			} catch (error) {
				this.errorMessage = error.message || 'ファイルを読み込めませんでした。'
			} finally {
				this.isLoading = false
			}
		},

		// 選択したフォルダー内から、Excelと同じディレクトリにある補助ファイルを取得します。
		async getCompanionFile(selectedFiles, excelFile, fileName, responseType) {
			const selected = findSiblingFile(selectedFiles, excelFile, fileName)
			if (!selected) return null
			return responseType === 'json' ? JSON.parse(await selected.text()) : selected
		},

		// 同一ベース名のPDFとJSONが両方ある場合だけPDF連動を開始します。
		async loadPdfCompanions(excelFile, selectedFiles) {
			const baseName = excelFile.name.replace(/\.(xlsx|xls)$/i, '')
			const pdfFileName = `${baseName}.pdf`
			const jsonFileName = `${baseName}.json`
			const [pdfSource, coordinateData] = await Promise.all([
				this.getCompanionFile(selectedFiles, excelFile, pdfFileName, 'arrayBuffer'),
				this.getCompanionFile(selectedFiles, excelFile, jsonFileName, 'json')
			])

			if (!pdfSource || !coordinateData) {
				this.pdfStatusMessage = 'PDF連動なし'
				return
			}

			try {
				const coordinateIndex = createCoordinateIndex(coordinateData, {
					pdf: pdfFileName,
					workbook: excelFile.name
				})
				const pdfDocument = await loadPdfDocument(pdfSource)
				if (coordinateIndex.page > pdfDocument.numPages) {
					throw new Error('座標JSONのページがPDFに存在しません。')
				}
				const page = await pdfDocument.getPage(coordinateIndex.page)
				const viewport = page.getViewport({ scale: 1 })
				const [expectedWidth, expectedHeight] = coordinateIndex.size
				if (Math.abs(viewport.width - expectedWidth) > 1 || Math.abs(viewport.height - expectedHeight) > 1) {
					throw new Error('PDFと座標JSONのページサイズが一致しません。')
				}

				this.coordinateIndex = coordinateIndex
				// PDF.jsのprivate fieldをVueのProxyで包まないよう、監視対象外として保持します。
				this.pdfDocument = Vue.markRaw(pdfDocument)
				this.pdfFileName = pdfFileName
				this.pdfPageNumber = coordinateIndex.page
				this.pdfPageCount = pdfDocument.numPages
				await this.$nextTick()
				await this.renderCurrentPdfPage()
			} catch (error) {
				await this.resetPdfViewer()
				this.pdfStatusMessage = `PDF連動なし：${error.message}`
			}
		},

		// PDF連動用の状態を破棄し、Excel単体表示へ戻します。
		async resetPdfViewer() {
			const documentToDestroy = this.pdfDocument
			this.pdfDocument = null
			this.pdfFileName = ''
			this.pdfPageNumber = 1
			this.pdfPageCount = 0
			this.pdfRenderSize = null
			this.pdfMarker = null
			this.coordinateIndex = null
			this.pdfStatusMessage = 'PDF連動なし'
			if (documentToDestroy) await documentToDestroy.destroy()
		},

		// 現在のPDFページを右側の表示幅へ合わせて描画します。
		async renderCurrentPdfPage() {
			if (!this.pdfDocument || !this.$refs.pdfCanvas || !this.$refs.pdfScroll) return
			this.pdfRenderSize = await renderPdfPage(
				this.pdfDocument,
				this.pdfPageNumber,
				this.$refs.pdfCanvas,
				this.$refs.pdfScroll.clientWidth
			)
		},

		// 表セルへマウスを重ねたとき、対応するPDF座標へマーカーを表示します。
		async showPdfMarker(sourceIndex, columnIndex) {
			if (!this.hasPdfViewer) return
			const marker = getCellCoordinate(
				this.coordinateIndex,
				this.selectedSheetName,
				sourceIndex,
				columnIndex
			)
			this.pdfMarker = marker
			if (!marker) return
			await this.$nextTick()
			this.scrollPdfMarkerIntoView()
		},

		clearPdfMarker() {
			this.pdfMarker = null
		},

		// マーカーがPDF表示領域の中央付近に見えるようスクロールします。
		scrollPdfMarkerIntoView() {
			const scroll = this.$refs.pdfScroll
			const marker = this.$refs.pdfMarker
			if (!scroll || !marker) return
			const left = marker.offsetLeft - scroll.clientWidth / 2
			const top = marker.offsetTop - scroll.clientHeight / 2
			scroll.scrollTo({ left: Math.max(0, left), top: Math.max(0, top), behavior: 'smooth' })
		},

		handleWindowResize() {
			clearTimeout(this.resizeTimer)
			this.resizeTimer = setTimeout(() => this.renderCurrentPdfPage(), 150)
		},

		// 選択されたシートから、見出しとデータ行を読み込みます。
		loadSelectedSheet() {
			this.errorMessage = ''
			this.closeCellModal()
			this.clearPdfMarker()
			this.selectedRowSourceIndex = null
			this.clearSearch()
			this.selectedHeaderIndex1 = 0
			this.selectedHeaderIndex2 = 0

			const table = getSheetTable(this.workbook, this.selectedSheetName, XLSX)
			this.headers = table.headers
			this.rows = table.rows
		},

		// 左端の行番号をクリックするたびに、その行の選択状態を切り替えます。
		toggleSelectedRow(sourceIndex) {
			this.selectedRowSourceIndex = this.selectedRowSourceIndex === sourceIndex
				? null
				: sourceIndex
		},

		// 選択列の値が、検索結果内の直前行から変化したかを判定します。
		isDifferentFromPrevious(resultIndex, columnIndex) {
			return isDifferentFromPrevious(
				this.filteredRows,
				this.highlightedColumnIndexes,
				resultIndex,
				columnIndex
			)
		},

		// 指定した列を、変化を強調する対象へ追加または対象外にします。
		toggleHighlightColumn(columnIndex, checked) {
			this.highlightColumnsBySheet = {
				...this.highlightColumnsBySheet,
				[this.selectedSheetName]: toggleSelectedColumn(
					this.highlightedColumnIndexes,
					columnIndex,
					checked
				)
			}
		},

		// 現在のシートにあるすべての列を強調対象にします。
		selectAllHighlightColumns() {
			this.highlightColumnsBySheet = {
				...this.highlightColumnsBySheet,
				[this.selectedSheetName]: getAllColumnIndexes(this.headers)
			}
		},

		// 現在のシートに設定された強調対象をすべて解除します。
		clearHighlightColumns() {
			this.highlightColumnsBySheet = {
				...this.highlightColumnsBySheet,
				[this.selectedSheetName]: []
			}
		},

		// クリックされたセルの全文を、編集用モーダルに表示します。
		openCellModal(rowIndex, columnIndex) {
			const address = getCellAddress(rowIndex, columnIndex, XLSX)
			this.editingCell = createEditingCell(
				rowIndex,
				columnIndex,
				address,
				this.headers[columnIndex]
			)
			this.editingValue = displayValue(this.rows[rowIndex][columnIndex])
			this.$nextTick(() => this.$refs.cellEditor?.focus())
		},

		// セル編集用モーダルを閉じ、入力中の内容を破棄します。
		closeCellModal() {
			this.editingCell = null
			this.editingValue = ''
		},

		// モーダルの入力内容をセルへ反映し、モーダルを閉じます。
		saveCellEdit() {
			if (!this.editingCell) return
			const { rowIndex, columnIndex } = this.editingCell
			this.applyCellValue(rowIndex, columnIndex, this.editingValue)
			this.closeCellModal()
		},

		// 入力文字列を適切な値へ変換し、表とExcelシートの両方へ反映します。
		applyCellValue(rowIndex, columnIndex, text) {
			if (text === displayValue(this.rows[rowIndex][columnIndex])) return

			const value = parseEditedValue(text)
			this.rows[rowIndex][columnIndex] = value
			const address = updateWorkbookCell(
				this.workbook,
				this.selectedSheetName,
				rowIndex,
				columnIndex,
				value,
				XLSX
			)
			this.editedCellKeys.add(`${this.selectedSheetName}!${address}`)
			this.editedCellKeys = new Set(this.editedCellKeys)
		},

		// セルの値を、画面表示に適した文字列へ変換します。
		displayValue(value) {
			return displayValue(value)
		},

		// 指定された列が「掲載ページ」列かどうかを判定します。
		isCatalogPageColumn(columnIndex) {
			return String(this.headers[columnIndex]).trim() === '掲載ページ'
		},

		// 「製品番号」列を横スクロール時の固定対象として判定します。
		isProductNumberColumn(columnIndex) {
			return String(this.headers[columnIndex]).trim() === '製品番号'
		},

		// 掲載ページの値から、電子カタログを開くURLを作ります。
		getCatalogPageUrl(value) {
			return getCatalogPageUrl(value)
		},

		// 編集内容を保持したExcelブックを、新しいファイルとして保存します。
		saveWorkbook() {
			if (!this.workbook) return
			saveWorkbookFile(this.workbook, this.fileName, XLSX)
		},

		// 2つの検索条件に入力された正規表現を消去します。
		clearSearch() {
			this.searchText1 = ''
			this.searchText2 = ''
		},

		// 件数などの数値を、桁区切り付きの文字列へ変換します。
		formatNumber(value) {
			return formatNumber(value)
		}
	}
}).mount('#app')

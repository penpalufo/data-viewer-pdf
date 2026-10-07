/* Excel Data Viewer
	* Version history:
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

// アプリ全体で使用するバージョン情報と表示件数の上限です。
const APP_VERSION = '0.3.3'
const RELEASE_DATE = '2026-09-29'
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
			editingValue: ''
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
		}
	},

	// 画面イベントを受け取り、機能モジュールの処理を呼び出します。
	methods: {
		// ドラッグ中のファイルが読込領域から離れたとき、強調表示を解除します。
		handleDragLeave(event) {
			if (!event.currentTarget.contains(event.relatedTarget)) this.isDragging = false
		},

		// 読込領域へドロップされたExcelファイルを受け取ります。
		handleDrop(event) {
			this.isDragging = false
			const [file] = event.dataTransfer.files
			if (file) this.loadExcel(file)
		},

		// ファイル選択ボタンで選ばれたExcelファイルを受け取ります。
		handleFileSelect(event) {
			const [file] = event.target.files
			if (file) this.loadExcel(file)
			event.target.value = ''
		},

		// Excelファイルを読み込み、シート一覧と最初のシートを準備します。
		async loadExcel(file) {
			this.errorMessage = ''

			if (!/\.(xlsx|xls)$/i.test(file.name)) {
				this.errorMessage = 'Excelファイル（.xlsx または .xls）を選択してください。'
				return
			}

			this.isLoading = true
			try {
				const workbook = await readExcelFile(file, XLSX)
				if (!workbook.SheetNames.length) throw new Error('シートが見つかりませんでした。')

				this.fileName = file.name
				this.workbook = workbook
				this.sheetNames = [...workbook.SheetNames]
				this.selectedSheetName = this.sheetNames[0]
				this.editedCellKeys = new Set()
				this.highlightColumnsBySheet = {}
				this.loadSelectedSheet()
			} catch (error) {
				this.errorMessage = error.message || 'ファイルを読み込めませんでした。'
			} finally {
				this.isLoading = false
			}
		},

		// 選択されたシートから、見出しとデータ行を読み込みます。
		loadSelectedSheet() {
			this.errorMessage = ''
			this.closeCellModal()
			this.clearSearch()
			this.selectedHeaderIndex1 = 0
			this.selectedHeaderIndex2 = 0

			const table = getSheetTable(this.workbook, this.selectedSheetName, XLSX)
			this.headers = table.headers
			this.rows = table.rows
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

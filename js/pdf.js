let pdfjsPromise = null

// PDF.jsはPDFが必要になった時点で読み込み、Excel単体画面の初期化を妨げないようにします。
async function getPdfJs() {
	if (!pdfjsPromise) {
		pdfjsPromise = import('../vendor/pdfjs/pdf.min.js').then(pdfjsLib => {
			pdfjsLib.GlobalWorkerOptions.workerSrc = './vendor/pdfjs/pdf.worker.min.js'
			return pdfjsLib
		})
	}
	return pdfjsPromise
}

// PDFファイルをPDF.jsの文書オブジェクトとして読み込みます。
export async function loadPdfDocument(source) {
	const pdfjsLib = await getPdfJs()
	const data = source instanceof File ? await source.arrayBuffer() : source
	return pdfjsLib.getDocument({ data }).promise
}

// 指定ページを、表示領域の幅に収まる倍率でCanvasへ描画します。
export async function renderPdfPage(pdfDocument, pageNumber, canvas, containerWidth) {
	const page = await pdfDocument.getPage(pageNumber)
	const baseViewport = page.getViewport({ scale: 1 })
	const availableWidth = Math.max(280, containerWidth - 32)
	const scale = availableWidth / baseViewport.width
	const viewport = page.getViewport({ scale })
	const outputScale = window.devicePixelRatio || 1
	const context = canvas.getContext('2d')

	canvas.width = Math.floor(viewport.width * outputScale)
	canvas.height = Math.floor(viewport.height * outputScale)
	canvas.style.width = `${viewport.width}px`
	canvas.style.height = `${viewport.height}px`

	await page.render({
		canvasContext: context,
		viewport,
		transform: outputScale === 1 ? null : [outputScale, 0, 0, outputScale, 0, 0]
	}).promise

	return { width: viewport.width, height: viewport.height, scale }
}


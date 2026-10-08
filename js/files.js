// Fileオブジェクトから、選択したフォルダー内での相対パスを取得します。
export function getRelativeFilePath(file) {
	return String(file.viewerRelativePath || file.webkitRelativePath || file.name).replaceAll('\\', '/')
}

// 相対パスの親ディレクトリ部分を返します。
export function getFileDirectory(file) {
	const path = getRelativeFilePath(file)
	const separatorIndex = path.lastIndexOf('/')
	return separatorIndex < 0 ? '' : path.slice(0, separatorIndex)
}

// Excelと同じディレクトリにある、指定ファイル名のFileを探します。
export function findSiblingFile(files, excelFile, fileName) {
	const directory = getFileDirectory(excelFile).toLowerCase()
	return files.find(file => {
		return getFileDirectory(file).toLowerCase() === directory
			&& file.name.toLowerCase() === fileName.toLowerCase()
	}) || null
}

// ドロップされたファイルまたはフォルダーを再帰的に展開します。
export async function getDroppedFiles(dataTransfer) {
	const items = [...(dataTransfer.items || [])]
	const entries = items
		.map(item => item.webkitGetAsEntry?.())
		.filter(Boolean)

	if (!entries.length) return [...dataTransfer.files]

	const files = []
	for (const entry of entries) {
		files.push(...await readEntry(entry, ''))
	}
	return files
}

async function readEntry(entry, parentPath) {
	const relativePath = parentPath ? `${parentPath}/${entry.name}` : entry.name

	if (entry.isFile) {
		const file = await new Promise((resolve, reject) => entry.file(resolve, reject))
		Object.defineProperty(file, 'viewerRelativePath', {
			value: relativePath,
			configurable: true
		})
		return [file]
	}

	if (!entry.isDirectory) return []
	const reader = entry.createReader()
	const children = []
	while (true) {
		const batch = await new Promise((resolve, reject) => reader.readEntries(resolve, reject))
		if (!batch.length) break
		children.push(...batch)
	}

	const files = []
	for (const child of children) {
		files.push(...await readEntry(child, relativePath))
	}
	return files
}


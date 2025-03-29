import * as path from "path"
import fs from "fs/promises"
import * as XLSX from "xlsx"

export interface ChunkData {
	sheetName: string
	columnName: string
	startRow: number
	endRow: number
	data: string
}

export interface SheetInfo {
	name: string
	rowCount: number
	columnCount: number
	headers: string[]
}

export interface ExcelLensResult {
	type: "init" | "chunk" | "complete" | "error"
	chunkData?: ChunkData
	fileInfo?: {
		fileName: string
		sheets: SheetInfo[]
	}
	message?: string
}

export interface SessionState {
	currentChunkIndex: number
	totalChunks: number
	metadata: {
		fileName: string
		sheets: Record<string, SheetInfo>
	}
	csvChunks: ChunkData[]
}

// 会话管理
const sessions = new Map<string, SessionState>()

// 添加文件类型检查函数
// function isExcelFile(filePath: string): boolean {
// 	const ext = path.extname(filePath).toLowerCase()
// 	return [".xlsx", ".xls", ".xlsm", ".xlsb", ".csv"].includes(ext)
// }

// 增加文件类型检查和错误处理
export async function extractExelFile(
	filePath: string,
	// options?: {
	// 	action?: "continue" | "stop"
	// },
): Promise<string> {
	try {
		await fs.access(filePath)
	} catch (error) {
		throw new Error(`File not found: ${filePath}`)
	}

	try {
		const sessionId = path.resolve(filePath)
		let session = sessions.get(sessionId)

		if (!session) {
			return await initializeSession(filePath)
		}

		return await processNextChunk(sessionId)
		// if (session.currentChunkIndex >= session.totalChunks) {
		// 	sessions.delete(sessionId)
		// 	return JSON.stringify({
		// 		type: "complete",
		// 		message: "excel文件分析完成。",
		// 	} as ExcelLensResult)
		// } else {
		// 	return await processNextChunk(sessionId)
		// }
	} catch (error) {
		sessions.delete(path.resolve(filePath))
		return JSON.stringify({
			type: "error",
			message: error instanceof Error ? error.message : "excel文件分析失败",
		} as ExcelLensResult)
	}
}

function determineChunkSize(columnCount: number): number {
	const targetCellCount = 4000
	const rowCount = Math.floor(targetCellCount / columnCount)
	return rowCount
}

async function initializeSession(filePath: string): Promise<string> {
	const workbook = XLSX.readFile(filePath)
	const fileName = path.basename(filePath)

	const session: SessionState = {
		currentChunkIndex: 0,
		totalChunks: 0,
		metadata: {
			fileName,
			sheets: {},
		},
		csvChunks: [],
	}

	for (const sheetName of workbook.SheetNames) {
		const worksheet = workbook.Sheets[sheetName]

		// 直接转换为CSV格式
		const csvData = XLSX.utils.sheet_to_csv(worksheet, {
			blankrows: false, // 跳过空行
			strip: true, // 去除空格
		})
		const rows = csvData.split("\n").filter((row) => row.trim())
		const columnCount = rows[0]?.split(",").length || 0 // 获取列数
		const rowCount = rows.length // 获取行数

		const columnNames = rows[0]?.split(",").map((col) => col.trim()) || []
		const columnNamesString = `${columnNames.join(",")}` // 添加列名信息

		const chunkSize = determineChunkSize(columnCount)

		for (let startRow = 1; startRow <= rowCount; startRow += chunkSize) {
			const endRow = Math.min(startRow + chunkSize, rowCount)
			const selectedRows = rows.slice(startRow, endRow)
			session.csvChunks.push({
				sheetName: sheetName,
				columnName: columnNamesString,
				startRow: startRow,
				endRow: endRow,
				data: selectedRows.join("\n"),
			})
		}
	}

	session.totalChunks = session.csvChunks.length
	sessions.set(path.resolve(filePath), session)

	if (session.totalChunks === 0) {
		return JSON.stringify({
			type: "complete",
			message: "excel文件没有数据。",
		} as ExcelLensResult)
	}

	return await processNextChunk(path.resolve(filePath))
}

async function processNextChunk(sessionId: string): Promise<string> {
	const session = sessions.get(sessionId)!
	const currentChunk = session.csvChunks[session.currentChunkIndex]
	let result: ExcelLensResult

	if (session.currentChunkIndex + 1 >= session.totalChunks) {
		result = {
			type: "complete",
			chunkData: currentChunk,
			message: `已获取最后部分数据，即excel中的第 ${session.currentChunkIndex + 1} 个数据块，工作表: ${currentChunk.sheetName}, 行范围: ${currentChunk.startRow}-${currentChunk.endRow})。请总结分析当前数据和之前获取的数据。`,
		}
		sessions.delete(sessionId)
	} else {
		result = {
			type: "chunk",
			chunkData: currentChunk,
			message: `请分析excel中的第 ${session.currentChunkIndex + 1} 个数据块，工作表: ${currentChunk.sheetName}, 行范围: ${currentChunk.startRow}-${currentChunk.endRow})。需继续调用当前工具以获取全部数据进行分析。`,
		}
	}

	session.currentChunkIndex++
	return JSON.stringify(result)
}

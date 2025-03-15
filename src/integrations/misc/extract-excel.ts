import * as path from "path"
// @ts-ignore-next-line
import pdf from "pdf-parse/lib/pdf-parse"
import mammoth from "mammoth"
import fs from "fs/promises"
import { isBinaryFile } from "isbinaryfile"

import * as XLSX from "xlsx"

export async function extractExelFile(filePath: string): Promise<string> {
	try {
		await fs.access(filePath)
	} catch (error) {
		throw new Error(`File not found: ${filePath}`)
	}
	const fileExtension = path.extname(filePath).toLowerCase()

	// switch (fileExtension) {
	//     case ".xlsx":
	//         const workbook = await XLSX.readFile(filePath)
	//         const sheetName = workbook.SheetNames[0]
	//         const worksheet = workbook.Sheets[sheetName]
	//         const jsonData = XLSX.utils.sheet_to_json(worksheet)
	//         return JSON.stringify(jsonData)
	// 	default:
	// 		throw new Error(`Cannot read excel for file type: ${fileExtension}`)

	switch (fileExtension) {
		case ".xlsx":
		case ".xls":
		case ".xlsm":
		case ".xlsb":
		case ".csv":
			const workbook = XLSX.readFile(filePath)
			const sheetNames = workbook.SheetNames
			const allSheetsData: Record<string, any[]> = {}
			for (const sheetName of sheetNames) {
				const worksheet = workbook.Sheets[sheetName]
				const sheetData = XLSX.utils.sheet_to_json(worksheet, {
					header: 1,
					defval: "",
					raw: false,
				}) as any[]
				const filteredData = sheetData.filter((row) => Array.isArray(row) && row.some((cell) => cell !== ""))
				const limitedData = filteredData.length > 20 ? filteredData.slice(0, 21) : filteredData
				allSheetsData[sheetName] = limitedData
			}
			return JSON.stringify(allSheetsData)
		default:
			throw new Error(`Cannot read excel for file type: ${fileExtension}`)
	}

	// switch (fileExtension) {
	// 	case ".pdf":
	// 		return extractTextFromPDF(filePath)
	// 	case ".docx":
	// 		return extractTextFromDOCX(filePath)
	// 	case ".ipynb":
	// 		return extractTextFromIPYNB(filePath)
	// 	default:
	// 		const isBinary = await isBinaryFile(filePath).catch(() => false)
	// 		if (!isBinary) {
	// 			return await fs.readFile(filePath, "utf8")
	// 		} else {
	// 			throw new Error(`Cannot read text for file type: ${fileExtension}`)
	// 		}
	// }
}

// async function extractTextFromPDF(filePath: string): Promise<string> {
// 	const dataBuffer = await fs.readFile(filePath)
// 	const data = await pdf(dataBuffer)
// 	return data.text
// }

// async function extractTextFromDOCX(filePath: string): Promise<string> {
// 	const result = await mammoth.extractRawText({ path: filePath })
// 	return result.value
// }

// async function extractTextFromIPYNB(filePath: string): Promise<string> {
// 	const data = await fs.readFile(filePath, "utf8")
// 	const notebook = JSON.parse(data)
// 	let extractedText = ""

// 	for (const cell of notebook.cells) {
// 		if ((cell.cell_type === "markdown" || cell.cell_type === "code") && cell.source) {
// 			extractedText += cell.source.join("\n") + "\n"
// 		}
// 	}

// 	return extractedText
// }

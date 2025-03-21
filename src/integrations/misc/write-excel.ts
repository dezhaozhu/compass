import * as path from "path"
import * as XLSX from "xlsx"
import fs from "fs/promises"
import ExcelJS from "exceljs"

/**
 * 将字符串转换为Record<string, any[]>格式
 * @param input 字符串或已经是Record<string, any[]>的对象
 * @returns 转换后的Record<string, any[]>对象，如果转换失败则返回undefined
 */
export function parseColumnData(input: string | undefined) {
	if (!input) {
		return undefined
	}

	// 尝试解析字符串为JSON对象
	try {
		const parsed = JSON.parse(input)

		// 验证解析后的对象是否符合Record<string, any[]>格式
		if (typeof parsed === "object" && parsed !== null) {
			const isValid = Object.values(parsed).every((value) => Array.isArray(value))
			if (isValid) {
				return parsed as Record<string, any[]>
			}
		}

		console.error("写入Excel的数据格式不正确，应为{列名: 数组值}格式", parsed)
		return undefined
	} catch (error) {
		// console.error("解析写入Excel的数据时出错:", error)
		return undefined
	}
}

/**
 * 打开 Excel 文件并根据 JSON 数据在多个列写入内容
 * @param filePath Excel 文件路径
 * @param columnData JSON 对象或字符串，键为列名，值为要写入的数据数组
 * @param sheetName 可选，工作表名称，默认使用第一个工作表
 */
export async function writeExcelFile(filePath: string, columnData: string, sheetName?: string) {
	try {
		await fs.access(filePath)
	} catch (error) {
		throw new Error(`File not found: ${filePath}`)
	}

	// 解析或验证输入数据
	let parsedData = parseColumnData(columnData)
	if (!parsedData) {
		return
		// throw new Error("未定义的写入数据")
		// parsedData = {};
	}

	const fileExtension = path.extname(filePath).toLowerCase()
	switch (fileExtension) {
		case ".csv":
		case ".xls":
		case ".xlsx":
			// 读取 Excel 文件
			const workbook = new ExcelJS.Workbook()
			let worksheet: ExcelJS.Worksheet | undefined

			// 尝试读取现有文件
			try {
				await workbook.xlsx.readFile(filePath)
				worksheet = workbook.getWorksheet(1)
				if (!worksheet) {
					worksheet = workbook.addWorksheet("Sheet1")
				}
			} catch (error) {
				return
			}

			const columnNames = Object.keys(parsedData)

			// 首先检查是否有任何列已存在
			for (const columnName of columnNames) {
				const col = worksheet.columns.find((c: Partial<ExcelJS.Column>) => c.header === columnName)
				if (col) {
					// console.log(`列 ${columnName} 在文件 ${filePath} 中已存在，跳过整个文件的写入`)
					return // 直接返回，完全跳过写入操作
				} else {
					columnNames.forEach((columnName) => {
						// 添加新列
						const nextCol = worksheet.columns.length + 1
						worksheet.getColumn(nextCol).header = columnName
						const col = worksheet.getColumn(nextCol)

						// 写入数据
						const values = parsedData[columnName]
						values.forEach((value, index) => {
							worksheet.getCell(index + 2, col.number).value = value
						})
					})
				}
			}

			// 保存文件
			await workbook.xlsx.writeFile(filePath)
			console.log(`成功将数据写入 ${filePath} 的列`)
			break
		default:
			throw new Error(`Cannot write excel for file type: ${fileExtension}`)
	}
}

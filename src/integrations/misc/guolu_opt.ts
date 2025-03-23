import { runPythonScript } from "../../utils/run_py"
import * as path from "path"

export async function guoluOpt(filePath: string): Promise<string> {
	return new Promise((reject) => {
		// 脚本路径目前要用绝对路径
		runPythonScript(
			path.join(__dirname, "..", "src", "py_scripts", "jsp_calculate.py"),
			["--file_path", filePath],
			(error, data) => {
				if (error) {
					console.error(error)
					reject(error)
				} else {
					// data 就是 python 脚本打印的 json 字符串
					console.log(data)
					// console.log(`锅炉排程优化完成`)
					// if (data) {
					// 	try {
					// 		// 先解析 Python 输出的 JSON 字符串为 JavaScript 对象
					// 		const parsedData = JSON.parse(data)
					// 		// 再将对象转换回 JSON 字符串，但不转义中文字符
					// 		resolve(JSON.stringify(parsedData, null, 2))
					// 	} catch (e) {
					// 		console.error("Error parsing JSON data:", e)
					// 		reject(new Error("Failed to parse JSON data from Python script"))
					// 	}
					// } else {
					// 	reject(new Error("No data returned from Python script"))
					// }
				}
			},
		)
	})
}

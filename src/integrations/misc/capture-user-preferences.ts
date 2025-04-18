import { runPythonScript } from "../../utils/run_py"
import * as path from "path"
import { ClineProvider } from "../../core/webview/ClineProvider"

export async function captureUserPreferences(filePath: string): Promise<string> {
	// 读取json, json 在 当前文件的同级目录
	// const data = {"A63": {"绿叶": 0.5, "金科尔": 0.5}, "B52": {"德海": 0.24, "汇能": 0.24, "绿叶": 0.24, "申港": 0.12, "祥安": 0.09, "都江电力": 0.06, "六安祥安": 0.03}, "C52": {"绿叶": 0.33, "汇能": 0.24, "德海": 0.18, "祥安": 0.11, "华益": 0.07, "四方": 0.04, "太湖": 0.02}, "C53": {"源能力通": 0.67, "源能力通,上锅": 0.33}, "D42": {"绿叶": 0.57, "德海": 0.17, "华益": 0.13, "中科": 0.04, "金科尔": 0.04, "银锅": 0.04}, "D52": {"汇能": 0.25, "绿叶": 0.25, "德海": 0.2, "华益": 0.14, "上锅": 0.06, "武锅": 0.04, "申港": 0.02, "祥安": 0.02, "源能力通": 0.01}, "D53": {"德海": 0.5, "源能力通": 0.38, "汇能": 0.12}, "E32": {"华益": 0.33, "申港": 0.22, "绿叶": 0.22, "汇能": 0.17, "源能力通": 0.06}, "E42": {"绿叶": 0.38, "汇能": 0.16, "华益": 0.14, "德海": 0.14, "中科": 0.07, "银锅、绿叶": 0.04, "上锅": 0.03, "银锅、汇能": 0.03, "祥安": 0.01}, "F12": {"上锅": 0.42, "汇能": 0.25, "绿叶": 0.17, "德耐特": 0.08, "德海": 0.04, "华益": 0.02, "杭富": 0.02}, "F13": {"源能力通": 0.67, "汇能": 0.33}, "F15": {"上锅": 0.34, "绿叶": 0.27, "汇能": 0.17, "德海": 0.05, "德耐特": 0.05, "申港": 0.05, "中科": 0.02, "祥安": 0.02, "上锅、江阴七星": 0.01, "源能力通": 0.01, "源能力通,华益": 0.01}, "FW2": {"上锅": 0.33, "汇能": 0.2, "绿叶": 0.18, "德海": 0.08, "新桠欣": 0.07, "中科": 0.03, "华益": 0.03, "金科尔": 0.03, "上锅;华益": 0.02, "申港": 0.02, "外购": 0.01, "汇能、祥川": 0.01, "都江电力": 0.01}, "G11": {"上锅": 0.39, "汇能": 0.22, "绿叶": 0.18, "新桠欣": 0.05, "德海": 0.04, "德耐特": 0.04, "申港": 0.04, "都江电力": 0.03, "武锅": 0.02}, "G15": {"绿叶": 0.33, "上锅": 0.25, "汇能": 0.21, "申港": 0.15, "德海": 0.03, "德耐特": 0.02, "武锅": 0.01, "新桠欣": 0.0}, "GC2": {"上锅": 1.0}, "GJ2": {"上锅": 1.0}, "JG1": {"太湖": 1.0}, "JG2": {"太湖": 1.0}, "KY2": {"绿叶": 1.0}, "SH2": {"德海": 0.39, "祥安": 0.22, "汇能": 0.17, "哈尔滨科能": 0.09, "上锅": 0.04, "中科": 0.04, "四方": 0.04}, "TA2": {"汇能": 1.0}}
	// const preferences = JSON.stringify(data)
	// return preferences
	// loadEnv()

	let preferencesPath: string | undefined

	// 创建一个 WeakRef 来引用 ClineProvider
	const provider = ClineProvider.getVisibleInstance()
	if (provider?.context?.globalStorageUri) {
		const globalStoragePath = provider.context.globalStorageUri.fsPath
		preferencesPath = path.join(globalStoragePath, "preferences.json")
	} else {
		preferencesPath = process.env.PREFERENCES_PATH || "./preference.json"
	}
	// 去掉双引号
	preferencesPath = String(preferencesPath).replace(/"/g, "")

	console.log("执行captureUserPreference, 参数:preferencesPath", preferencesPath, "filePath:", filePath)
	return new Promise((resolve, reject) => {
		// 脚本路径目前要用绝对路径
		runPythonScript(
			path.join(__dirname, "..", "src", "py_scripts", "capture_user_preferences.py"),
			["--file_path", filePath, "--preferences_path", preferencesPath],
			(error, data) => {
				if (error) {
					console.error(error)
					reject(error)
				} else {
					// data 就是 python 脚本打印的 json 字符串
					console.log(data)
					if (data) {
						try {
							// 先解析 Python 输出的 JSON 字符串为 JavaScript 对象
							const parsedData = JSON.parse(data)
							// 再将对象转换回 JSON 字符串，但不转义中文字符
							resolve(JSON.stringify(parsedData, null, 2))
						} catch (e) {
							console.error("Error parsing JSON data:", e)
							reject(new Error("Failed to parse JSON data from Python script"))
						}
					} else {
						reject(new Error("No data returned from Python script"))
					}
				}
			},
		)
	})
}

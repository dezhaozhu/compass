import { runPythonScript } from "../../utils/run_py"
import * as path from "path"

interface ProductItem {
	product_id: string
	target_date: string
	target_factory: string
}

type ProductList = ProductItem[]

export async function adjustJsp(filePath: string, productToDateOrder: ProductList): Promise<string> {
	// 测试
	// productToDateOrder = [
	//     {
	//         product_id: "Z-1400F15-260-740(7)",
	//         target_date: "2025-03-25 10:00:00",
	//         target_factory: "绿叶"
	//     }
	// ]

	const productToDateOrder_string = JSON.stringify(productToDateOrder)
	console.log("执行 adjustJsp", filePath, productToDateOrder_string)
	return new Promise((resolve, reject) => {
		// 脚本路径目前要用绝对路径
		runPythonScript(
			path.join(__dirname, "..", "src", "py_scripts", "adjust_jsp.py"),
			["--file_path", filePath, "--product_to_date_order", productToDateOrder_string],
			(error, data) => {
				if (error) {
					console.error(error)
					reject(error)
				} else {
					// data 就是 python 脚本执行的返回值
					console.log("adjustJsp 返回值", data)
					if (data) {
						resolve(data)
					} else {
						reject(new Error("No data returned from Python script"))
					}
				}
			},
		)
	})
}

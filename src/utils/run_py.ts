import { spawn } from "child_process"
import * as vscode from 'vscode';
import { execa } from 'execa';

/**
 * Runs a Python script and returns the output.
 * @param scriptPath - The path to the Python script.
 * @param args - The arguments to pass to the Python script.
 * @param callback - The callback function to handle the output.
 */

export function runPythonScript(
	scriptPath: string,
	args: string[],
	callback: (error: string | null, data: string | null) => void,
) {

	const pythonPath = vscode.workspace.getConfiguration('python').get<string>('defaultInterpreterPath') || 
	vscode.workspace.getConfiguration('python').get<string>('pythonPath') || 
	'python'; // 默认使用系统 PATH 中的 python 命令

	const pythonProcess = spawn(
		// "D:\\miniconda3\\envs\\dl_torch2\\python.exe", // 指定 Python 解释器路径
		pythonPath,
		[scriptPath].concat(args),
		{ env: { ...process.env, PYTHONIOENCODING: "utf-8" } }, // 确保 Python 使用 UTF-8 编码输出
	)
	let data = ""
	pythonProcess.stdout.on("data", (chunk) => {
		data += chunk.toString("utf8") // 明确指定使用 UTF-8 编码
	})

	pythonProcess.stderr.on("data", (error) => {
		console.error(`stderr: ${error}`)
	})

	pythonProcess.on("close", (code) => {
		if (code !== 0) {
			console.log(`Python script exited with code ${code}`)
			callback(`Error: Script exited with code ${code}`, null)
		} else {
			console.log("Python script executed successfully")
			callback(null, data)
		}
	})
}


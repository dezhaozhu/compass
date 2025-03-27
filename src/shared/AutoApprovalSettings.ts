export interface AutoApprovalSettings {
	// Whether auto-approval is enabled
	enabled: boolean
	// Individual action permissions
	actions: {
		readFiles: boolean // Read files and directories
		readExcel: boolean // Read files and directories
		captureUserPreferences: boolean
		adjustJsp: boolean
		editFiles: boolean // Edit files
		executeCommands: boolean // Execute safe commands
		useBrowser: boolean // Use browser
		useMcp: boolean // Use MCP servers
	}
	// Global settings
	maxRequests: number // Maximum number of auto-approved requests
	enableNotifications: boolean // Show notifications for approval and task completion
}

export const DEFAULT_AUTO_APPROVAL_SETTINGS: AutoApprovalSettings = {
	enabled: false,
	actions: {
		readFiles: false,
		readExcel: false,
		captureUserPreferences: false,
		adjustJsp: false,
		editFiles: false,
		executeCommands: false,
		useBrowser: false,
		useMcp: false,
	},
	maxRequests: 20,
	enableNotifications: false,
}

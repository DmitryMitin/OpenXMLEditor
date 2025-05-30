// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import * as os from 'os';

import { OpenXMLTreeDataProvider, OpenXMLTreeItem } from './openXmlTreeProvider';
import { OpenXMLFileSystemProvider } from './openXmlFileSystem';
import { XMLFormatter } from './services/xmlFormatter';
import { COMMANDS, CONTEXT_KEYS, EXTENSION_CONFIG, UI_CONFIG } from './constants';
import { logger } from './utils/logger';
import { FileUtils } from './utils/fileUtils';

let treeDataProvider: OpenXMLTreeDataProvider;
let fileSystemProvider: OpenXMLFileSystemProvider;
let xmlFormatter: XMLFormatter;
let openedOpenXMLFiles: Set<string> = new Set(); // Track multiple opened files

// This method is called when your extension is activated
// Your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {
	logger.info('OpenXML Editor extension is now active!');

	// Initialize providers
	treeDataProvider = new OpenXMLTreeDataProvider();
	fileSystemProvider = new OpenXMLFileSystemProvider();
	xmlFormatter = new XMLFormatter();

	// Register file system provider with detailed logging
	logger.debug('Registering OpenXML file system provider');
	const fsRegistration = vscode.workspace.registerFileSystemProvider(EXTENSION_CONFIG.scheme, fileSystemProvider, {
		isCaseSensitive: true,
		isReadonly: false
	});
	context.subscriptions.push(fsRegistration);
	logger.success('OpenXML file system provider registered successfully');

	// Register tree data provider
	const treeRegistration = vscode.window.registerTreeDataProvider('openxmlStructure', treeDataProvider);
	context.subscriptions.push(treeRegistration);
	logger.success('OpenXML tree data provider registered successfully');

	// Set up document save listener for temp files
	const documentSaveListener = vscode.workspace.onDidSaveTextDocument(async (document) => {
		// Check if this is a temp file from our extension
		const filePath = document.uri.fsPath;
		const tempDir = os.tmpdir();
		
		if (filePath.startsWith(tempDir) && filePath.includes(EXTENSION_CONFIG.tempDirPrefix)) {
			logger.debug('VS Code saved temp file, changes will be synced via file watcher', { filePath });
		}
	});
	context.subscriptions.push(documentSaveListener);
	logger.success('Document save listener registered');

	// Test file system provider registration
	console.log('🧪 Testing file system provider registration...');
	const testProviders = vscode.workspace.fs;
	console.log('Available file systems:', Object.getOwnPropertyNames(testProviders));

	// Register commands
	const openFileCommand = vscode.commands.registerCommand('openxmleditor.openFile', async (uri: vscode.Uri) => {
		await openOpenXMLFile(uri);
	});

	const saveChangesCommand = vscode.commands.registerCommand('openxmleditor.saveChanges', async () => {
		await saveAllChanges();
	});

	const refreshTreeCommand = vscode.commands.registerCommand('openxmleditor.refreshTree', async () => {
		await refreshTreeView();
	});

	const openXmlFileCommand = vscode.commands.registerCommand('openxmleditor.openXmlFile', async (treeItem: OpenXMLTreeItem) => {
		await openXmlFileFromTree(treeItem);
	});

	const formatXMLCommand = vscode.commands.registerCommand('openxmleditor.formatXML', async () => {
		await formatCurrentXML();
	});

	const showFileInfoCommand = vscode.commands.registerCommand('openxmleditor.showFileInfo', async (uri: vscode.Uri) => {
		await showOpenXMLFileInfo(uri);
	});

	const closeOpenXMLFileCommand = vscode.commands.registerCommand('openxmleditor.closeOpenXMLFile', async (treeItem: OpenXMLTreeItem) => {
		if (treeItem.openXMLPath) {
			await closeOpenXMLFile(treeItem.openXMLPath);
		}
	});

	const openInSystemCommand = vscode.commands.registerCommand('openxmleditor.openInSystem', async (target: vscode.Uri | OpenXMLTreeItem) => {
		if (target && typeof target === 'object' && 'openXMLPath' in target) {
			// Called from tree view - target is OpenXMLTreeItem
			const treeItem = target as OpenXMLTreeItem;
			if (treeItem.openXMLPath) {
				await openInSystemApplication(vscode.Uri.file(treeItem.openXMLPath));
			}
		} else {
			// Called from explorer context menu - target is Uri
			await openInSystemApplication(target as vscode.Uri);
		}
	});

	context.subscriptions.push(
		openFileCommand,
		saveChangesCommand,
		refreshTreeCommand,
		openXmlFileCommand,
		formatXMLCommand,
		showFileInfoCommand,
		closeOpenXMLFileCommand,
		openInSystemCommand
	);

	// Set up status bar
	const statusBarItem = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Left, 100);
	context.subscriptions.push(statusBarItem);

	// Update status bar when tree view changes
	treeDataProvider.onDidChangeTreeData(() => {
		updateStatusBar(statusBarItem);
	});
}

async function openOpenXMLFile(uri: vscode.Uri): Promise<void> {
	try {
		const filePath = uri.fsPath;
		
		// Check if file is already opened
		if (openedOpenXMLFiles.has(filePath)) {
			vscode.window.showInformationMessage(`OpenXML file already opened: ${path.basename(filePath)}`);
			return;
		}
		
		// Load the OpenXML file into our virtual file system
		await fileSystemProvider.loadOpenXMLFile(filePath);
		
		// Add to opened files
		openedOpenXMLFiles.add(filePath);
		
		// Add to tree structure
		await treeDataProvider.addOpenXMLFile(filePath);
		
		// Show the tree view
		await vscode.commands.executeCommand('setContext', 'openxmlFileOpen', true);
		await vscode.commands.executeCommand('openxmlStructure.focus');
		
		vscode.window.showInformationMessage(`OpenXML file loaded: ${path.basename(filePath)}`);
		
	} catch (error) {
		vscode.window.showErrorMessage(`Failed to open OpenXML file: ${error}`);
	}
}

async function saveAllChanges(): Promise<void> {
	if (openedOpenXMLFiles.size === 0) {
		vscode.window.showWarningMessage('No OpenXML files are currently open');
		return;
	}

	try {
		console.log('💾 Manual save requested for all OpenXML files');
		let savedCount = 0;
		let totalFiles = openedOpenXMLFiles.size;
		
		for (const filePath of openedOpenXMLFiles) {
			console.log(`💾 Processing save for: ${filePath}`);
			
			// Always try to save (this will force sync temp files first)
			try {
				await fileSystemProvider.saveChanges(filePath);
				savedCount++;
				console.log(`✅ Successfully saved: ${filePath}`);
			} catch (error) {
				console.error(`❌ Failed to save ${filePath}:`, error);
				vscode.window.showErrorMessage(`Failed to save ${path.basename(filePath)}: ${error}`);
			}
		}
		
		if (savedCount > 0) {
			vscode.window.showInformationMessage(`Saved ${savedCount} of ${totalFiles} OpenXML file(s)`);
		} else {
			vscode.window.showInformationMessage('No changes found to save');
		}
		
		// Update tree to reflect saved state
		treeDataProvider.refresh();
		
	} catch (error) {
		console.error('❌ Save all operation failed:', error);
		vscode.window.showErrorMessage(`Failed to save changes: ${error}`);
	}
}

async function refreshTreeView(): Promise<void> {
	try {
		await treeDataProvider.refreshAll();
	} catch (error) {
		vscode.window.showErrorMessage(`Failed to refresh tree view: ${error}`);
	}
}

async function openXmlFileFromTree(treeItem: OpenXMLTreeItem): Promise<void> {
	if (treeItem.isDirectory) {
		return;
	}

	try {
		console.log('Opening XML file from tree:', treeItem.fullPath);
		console.log('OpenXML file:', treeItem.openXMLPath);
		
		// Get the temporary file path instead of creating virtual URI
		const tempFilePath = fileSystemProvider.getTempFilePath(treeItem.openXMLPath!, treeItem.fullPath);
		
		if (!tempFilePath) {
			vscode.window.showErrorMessage(`Temporary file not found for: ${treeItem.fullPath}`);
			return;
		}
		
		console.log('Opening temporary file:', tempFilePath);
		
		// Open the temporary file as a regular file
		const tempUri = vscode.Uri.file(tempFilePath);
		const document = await vscode.workspace.openTextDocument(tempUri);
		
		// Ensure language mode is set to XML
		const fileExtension = path.extname(treeItem.fullPath).toLowerCase();
		if (fileExtension === '.xml' || fileExtension === '.rels') {
			if (document.languageId !== 'xml') {
				await vscode.languages.setTextDocumentLanguage(document, 'xml');
			}
		}
		
		// Open the document in editor
		const editor = await vscode.window.showTextDocument(document, {
			preview: false, // Open in new tab, not preview
			viewColumn: vscode.ViewColumn.Active
		});
		
		console.log('Successfully opened temporary file:', tempFilePath);
		console.log('Document language ID:', document.languageId);
		console.log('Document URI scheme:', document.uri.scheme);
		console.log('Document URI path:', document.uri.path);
		
	} catch (error) {
		console.error('Failed to open XML file:', error);
		vscode.window.showErrorMessage(`Failed to open file: ${error}`);
	}
}

async function closeOpenXMLFile(filePath: string): Promise<void> {
	try {
		if (!openedOpenXMLFiles.has(filePath)) {
			return;
		}
		
		// Check for unsaved changes
		if (fileSystemProvider.hasUnsavedChanges(filePath)) {
			const result = await vscode.window.showWarningMessage(
				`File ${path.basename(filePath)} has unsaved changes. Save before closing?`,
				'Save and Close',
				'Close without Saving',
				'Cancel'
			);
			
			if (result === 'Save and Close') {
				await fileSystemProvider.saveChanges(filePath);
			} else if (result === 'Cancel') {
				return;
			}
		}
		
		// Clean up temporary files
		await fileSystemProvider.cleanupTempFiles(filePath);
		
		// Remove from opened files
		openedOpenXMLFiles.delete(filePath);
		
		// Remove from tree
		await treeDataProvider.removeOpenXMLFile(filePath);
		
		// Update context if no files are open
		if (openedOpenXMLFiles.size === 0) {
			await vscode.commands.executeCommand('setContext', 'openxmlFileOpen', false);
		}
		
		vscode.window.showInformationMessage(`Closed OpenXML file: ${path.basename(filePath)}`);
		
	} catch (error) {
		vscode.window.showErrorMessage(`Failed to close OpenXML file: ${error}`);
	}
}

async function openInSystemApplication(uri: vscode.Uri): Promise<void> {
	try {
		const filePath = uri.fsPath;
		const fileName = path.basename(filePath);
		
		// Check if file exists
		if (!fs.existsSync(filePath)) {
			vscode.window.showErrorMessage(`File not found: ${fileName}`);
			return;
		}
		
		// Open the file with the default system application
		await vscode.env.openExternal(vscode.Uri.file(filePath));
		
		vscode.window.showInformationMessage(`Opening ${fileName} in system application...`);
		
	} catch (error) {
		vscode.window.showErrorMessage(`Failed to open file in system application: ${error}`);
	}
}

async function formatCurrentXML(): Promise<void> {
	const activeEditor = vscode.window.activeTextEditor;
	if (!activeEditor) {
		vscode.window.showErrorMessage('No active editor found');
		return;
	}

	try {
		const document = activeEditor.document;
		const text = document.getText();
		
		// Enhanced XML file detection
		const isXMLFile = document.languageId === 'xml' || 
						  document.uri.scheme === EXTENSION_CONFIG.scheme ||
						  FileUtils.isXMLFile(document.fileName) ||
						  FileUtils.isXMLContent(text);
		
		if (!isXMLFile) {
			vscode.window.showWarningMessage('Current file does not appear to be an XML file');
			return;
		}

		// Check if file is empty or has no XML content
		if (!text.trim() || !text.includes('<')) {
			vscode.window.showWarningMessage('No XML content found to format');
			return;
		}
		
		// Show progress for large files
		if (text.length > UI_CONFIG.progressThreshold) {
			vscode.window.withProgress({
				location: vscode.ProgressLocation.Notification,
				title: "Formatting XML...",
				cancellable: false
			}, async (progress) => {
				progress.report({ increment: 50 });
				const formatted = xmlFormatter.formatXML(text);
				progress.report({ increment: 100 });
				return formatted;
			}).then(async (formatted) => {
				await applyFormattedText(activeEditor, text, formatted!);
			});
		} else {
			// Format immediately for smaller files
			const formatted = xmlFormatter.formatXML(text);
			await applyFormattedText(activeEditor, text, formatted);
		}

	} catch (error) {
		logger.error('Failed to format XML', error);
		vscode.window.showErrorMessage(`Failed to format XML: ${error}`);
	}
}

async function applyFormattedText(editor: vscode.TextEditor, originalText: string, formattedText: string): Promise<void> {
	// Check if formatting actually changed anything
	if (originalText === formattedText) {
		vscode.window.showInformationMessage('XML is already properly formatted');
		return;
	}

	// Save current selection
	const originalSelection = editor.selection;
	
	// Apply the formatting
	const edit = new vscode.WorkspaceEdit();
	const fullRange = new vscode.Range(
		editor.document.positionAt(0),
		editor.document.positionAt(originalText.length)
	);
	edit.replace(editor.document.uri, fullRange, formattedText);
	
	const success = await vscode.workspace.applyEdit(edit);
	
	if (success) {
		// Try to restore cursor position intelligently
		try {
			const originalOffset = editor.document.offsetAt(originalSelection.start);
			const lines = formattedText.split('\n');
			let newOffset = 0;
			let targetLine = 0;
			
			// Find approximately the same position in formatted text
			for (let i = 0; i < lines.length && newOffset < originalOffset; i++) {
				if (newOffset + lines[i].length >= originalOffset) {
					targetLine = i;
					break;
				}
				newOffset += lines[i].length + 1; // +1 for newline
				targetLine = i;
			}
			
			// Set cursor to start of the target line
			const newPosition = new vscode.Position(targetLine, 0);
			editor.selection = new vscode.Selection(newPosition, newPosition);
			editor.revealRange(new vscode.Range(newPosition, newPosition), vscode.TextEditorRevealType.InCenter);
			
		} catch (positionError) {
			// Fallback: move to beginning
			const newPosition = new vscode.Position(0, 0);
			editor.selection = new vscode.Selection(newPosition, newPosition);
		}
		
		vscode.window.showInformationMessage('XML formatted successfully');
	} else {
		vscode.window.showErrorMessage('Failed to apply XML formatting');
	}
}

async function showOpenXMLFileInfo(uri: vscode.Uri): Promise<void> {
	try {
		const filePath = uri.fsPath;
		const stats = fs.statSync(filePath);
		const fileName = path.basename(filePath);
		const fileExt = path.extname(filePath).toLowerCase();

		const fileInfo = {
			fileName: fileName,
			fileSize: stats.size,
			fileType: FileUtils.getFileTypeDescription(fileExt),
			createdDate: stats.birthtime,
			modifiedDate: stats.mtime,
			hasUnsavedChanges: openedOpenXMLFiles.has(filePath) ? fileSystemProvider.hasUnsavedChanges(filePath) : false,
			modifiedFiles: openedOpenXMLFiles.has(filePath) ? fileSystemProvider.getModifiedFiles(filePath) : []
		};

		// Show information in a webview
		const panel = vscode.window.createWebviewPanel(
			'fileInfo',
			`File Info: ${fileName}`,
			vscode.ViewColumn.Beside,
			{}
		);

		panel.webview.html = generateFileInfoHTML(fileInfo);

	} catch (error) {
		logger.error('Failed to get file information', error);
		vscode.window.showErrorMessage(`Failed to get file information: ${error}`);
	}
}

function generateFileInfoHTML(fileInfo: any): string {
	const modifiedFilesHtml = fileInfo.modifiedFiles.length > 0 
		? `<tr><td>Modified Files</td><td>${fileInfo.modifiedFiles.join('<br>')}</td></tr>`
		: '';

	return `
		<!DOCTYPE html>
		<html>
		<head>
			<style>
				body { font-family: Arial, sans-serif; padding: 20px; }
				.info-table { border-collapse: collapse; width: 100%; }
				.info-table th, .info-table td { 
					border: 1px solid #ddd; 
					padding: 8px; 
					text-align: left; 
				}
				.info-table th { background-color: #f2f2f2; }
				.header { color: #333; margin-bottom: 20px; }
				.modified { color: #ff6b35; font-weight: bold; }
				.saved { color: #28a745; font-weight: bold; }
			</style>
		</head>
		<body>
			<h2 class="header">File Information</h2>
			<table class="info-table">
				<tr><th>Property</th><th>Value</th></tr>
				<tr><td>File Name</td><td>${fileInfo.fileName}</td></tr>
				<tr><td>File Type</td><td>${fileInfo.fileType}</td></tr>
				<tr><td>File Size</td><td>${FileUtils.formatFileSize(fileInfo.fileSize)}</td></tr>
				<tr><td>Created</td><td>${fileInfo.createdDate?.toLocaleString() || 'Unknown'}</td></tr>
				<tr><td>Modified</td><td>${fileInfo.modifiedDate?.toLocaleString() || 'Unknown'}</td></tr>
				<tr><td>Status</td><td class="${fileInfo.hasUnsavedChanges ? 'modified' : 'saved'}">${fileInfo.hasUnsavedChanges ? 'Has unsaved changes' : 'All changes saved'}</td></tr>
				${modifiedFilesHtml}
			</table>
		</body>
		</html>
	`;
}

function updateStatusBar(statusBarItem: vscode.StatusBarItem): void {
	if (openedOpenXMLFiles.size > 0) {
		let totalModified = 0;
		for (const filePath of openedOpenXMLFiles) {
			if (fileSystemProvider.hasUnsavedChanges(filePath)) {
				totalModified += fileSystemProvider.getModifiedFiles(filePath).length;
			}
		}
		
		statusBarItem.text = `$(file-zip) ${openedOpenXMLFiles.size} OpenXML file(s)${totalModified > 0 ? ` (${totalModified} modified)` : ''}`;
		statusBarItem.tooltip = totalModified > 0 
			? `${openedOpenXMLFiles.size} OpenXML files open with ${totalModified} unsaved changes - click to save all`
			: `${openedOpenXMLFiles.size} OpenXML files open - all changes saved`;
		statusBarItem.command = totalModified > 0 ? 'openxmleditor.saveChanges' : undefined;
		statusBarItem.show();
	} else {
		statusBarItem.hide();
	}
}

// This method is called when your extension is deactivated
export function deactivate() {
	// Clean up all temporary files
	for (const filePath of openedOpenXMLFiles) {
		fileSystemProvider.cleanupTempFiles(filePath).catch(error => {
			console.error('Failed to cleanup temp files during deactivation:', error);
		});
	}
	
	// Clean up context
	vscode.commands.executeCommand('setContext', 'openxmlFileOpen', false);
}
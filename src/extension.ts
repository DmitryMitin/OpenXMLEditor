// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import { OpenXMLTreeDataProvider, OpenXMLTreeItem } from './openXmlTreeProvider';
import { OpenXMLFileSystemProvider } from './openXmlFileSystem';

let treeDataProvider: OpenXMLTreeDataProvider;
let fileSystemProvider: OpenXMLFileSystemProvider;
let openedOpenXMLFiles: Set<string> = new Set(); // Track multiple opened files

interface OpenXMLFileInfo {
	fileName: string;
	fileSize: number;
	fileType: string;
	createdDate?: Date;
	modifiedDate?: Date;
	author?: string;
	title?: string;
}

// This method is called when your extension is activated
// Your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {
	console.log('OpenXML Editor extension is now active!');

	// Initialize providers
	treeDataProvider = new OpenXMLTreeDataProvider();
	fileSystemProvider = new OpenXMLFileSystemProvider();

	// Register file system provider with detailed logging
	console.log('Registering OpenXML file system provider...');
	const fsRegistration = vscode.workspace.registerFileSystemProvider('openxml', fileSystemProvider, {
		isCaseSensitive: true,
		isReadonly: false
	});
	context.subscriptions.push(fsRegistration);
	console.log('✅ OpenXML file system provider registered successfully');

	// Track OpenXML documents for better display
	vscode.workspace.onDidOpenTextDocument(document => {
		if (document.uri.scheme === 'openxml') {
			console.log('📄 Opened OpenXML document:', document.uri.toString());
			
			// Ensure the document is properly recognized as XML for other extensions
			const fileName = document.uri.path;
			if (fileName.endsWith('.xml') || fileName.endsWith('.rels')) {
				// Force language mode to XML if not already set
				if (document.languageId !== 'xml') {
					vscode.languages.setTextDocumentLanguage(document, 'xml').then(() => {
						console.log('✅ Set language mode to XML for:', fileName);
					});
				}
				
				// Set additional context variables that other extensions might check
				vscode.commands.executeCommand('setContext', 'resourceExtname', fileName.endsWith('.rels') ? '.rels' : '.xml');
				vscode.commands.executeCommand('setContext', 'resourceLangId', 'xml');
			}
		}
	});

	// Track active editor changes to maintain proper context for third-party extensions
	vscode.window.onDidChangeActiveTextEditor(editor => {
		if (editor && editor.document.uri.scheme === 'openxml') {
			const fileName = editor.document.uri.path;
			if (fileName.endsWith('.xml') || fileName.endsWith('.rels')) {
				// Refresh context variables when switching to openxml files
				const fileExtension = fileName.endsWith('.rels') ? '.rels' : '.xml';
				vscode.commands.executeCommand('setContext', 'resourceExtname', fileExtension);
				vscode.commands.executeCommand('setContext', 'resourceLangId', 'xml');
				vscode.commands.executeCommand('setContext', 'resourceScheme', 'openxml');
				
				console.log('🔄 Refreshed XML context for active editor:', fileName);
			}
		}
	});

	// Register tree data provider
	const treeRegistration = vscode.window.registerTreeDataProvider('openxmlStructure', treeDataProvider);
	context.subscriptions.push(treeRegistration);
	console.log('✅ OpenXML tree data provider registered successfully');

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
		let savedCount = 0;
		for (const filePath of openedOpenXMLFiles) {
			if (fileSystemProvider.hasUnsavedChanges(filePath)) {
				await fileSystemProvider.saveChanges(filePath);
				savedCount++;
			}
		}
		
		if (savedCount > 0) {
			vscode.window.showInformationMessage(`Saved changes in ${savedCount} OpenXML file(s)`);
		} else {
			vscode.window.showInformationMessage('No unsaved changes found');
		}
		
		// Update tree to reflect saved state
		treeDataProvider.refresh();
		
	} catch (error) {
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
		
		// Create virtual URI for the file
		const virtualUri = fileSystemProvider.createUri(treeItem.openXMLPath!, treeItem.fullPath);
		console.log('Created virtual URI:', virtualUri.toString());
		
		// Open the file in a NEW editor (not replacing current)
		const document = await vscode.workspace.openTextDocument(virtualUri);
		const editor = await vscode.window.showTextDocument(document, {
			preview: false, // Open in new tab, not preview
			viewColumn: vscode.ViewColumn.Active
		});
		
		// Ensure proper language mode and context for XML files
		const fileExtension = path.extname(treeItem.fullPath).toLowerCase();
		if (fileExtension === '.xml' || fileExtension === '.rels') {
			// Set language mode to XML - this is crucial for other extensions
			await vscode.languages.setTextDocumentLanguage(document, 'xml');
			
			// Set context variables that other XML extensions typically check
			await vscode.commands.executeCommand('setContext', 'resourceExtname', fileExtension);
			await vscode.commands.executeCommand('setContext', 'resourceLangId', 'xml');
			await vscode.commands.executeCommand('setContext', 'resourceScheme', 'openxml');
			
			// Also ensure the editor is focused and active for extension detection
			await vscode.window.showTextDocument(document, {
				preview: false,
				preserveFocus: false,
				viewColumn: editor.viewColumn
			});
			
			// Trigger manual events that might help other extensions detect the file
			setTimeout(() => {
				// Fire events that XML extensions might listen to
				vscode.commands.executeCommand('workbench.action.files.revert');
				vscode.commands.executeCommand('editor.action.redetectLanguage');
			}, 100);
			
			console.log('✅ XML language mode and contexts set for:', treeItem.fullPath);
		}
		
		console.log('Successfully opened file in new tab:', treeItem.fullPath);
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
		
		// Enhanced XML file detection for virtual files
		const isXMLFile = document.languageId === 'xml' || 
						  document.uri.scheme === 'openxml' ||
						  document.fileName.endsWith('.xml') || 
						  document.fileName.endsWith('.rels') ||
						  text.trim().startsWith('<?xml') ||
						  text.trim().startsWith('<');
		
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
		if (text.length > 50000) {
			vscode.window.withProgress({
				location: vscode.ProgressLocation.Notification,
				title: "Formatting XML...",
				cancellable: false
			}, async (progress) => {
				progress.report({ increment: 50 });
				const formatted = formatXML(text);
				progress.report({ increment: 100 });
				return formatted;
			}).then(async (formatted) => {
				await applyFormattedText(activeEditor, text, formatted!);
			});
		} else {
			// Format immediately for smaller files
			const formatted = formatXML(text);
			await applyFormattedText(activeEditor, text, formatted);
		}

	} catch (error) {
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

function formatXML(xml: string): string {
	try {
		// Clean up the input
		let formatted = xml.trim();
		
		// Remove all existing whitespace between tags
		formatted = formatted.replace(/>\s*</g, '><');
		
		// Split into array for processing
		const result: string[] = [];
		let depth = 0;
		const indent = '  '; // 2 spaces
		
		// Process character by character to build a token stream
		let i = 0;
		const tokens: Array<{type: 'tag' | 'text' | 'comment' | 'cdata', content: string}> = [];
		
		while (i < formatted.length) {
			const char = formatted[i];
			const nextChars4 = formatted.substring(i, i + 4);
			const nextChars9 = formatted.substring(i, i + 9);
			
			// Handle CDATA sections
			if (nextChars9 === '<![CDATA[') {
				const cdataEnd = formatted.indexOf(']]>', i);
				if (cdataEnd !== -1) {
					tokens.push({
						type: 'cdata',
						content: formatted.substring(i, cdataEnd + 3)
					});
					i = cdataEnd + 3;
					continue;
				}
			}
			
			// Handle comments
			if (nextChars4 === '<!--') {
				const commentEnd = formatted.indexOf('-->', i);
				if (commentEnd !== -1) {
					tokens.push({
						type: 'comment',
						content: formatted.substring(i, commentEnd + 3)
					});
					i = commentEnd + 3;
					continue;
				}
			}
			
			// Handle tags
			if (char === '<') {
				const tagEnd = formatted.indexOf('>', i);
				if (tagEnd !== -1) {
					tokens.push({
						type: 'tag',
						content: formatted.substring(i, tagEnd + 1)
					});
					i = tagEnd + 1;
					continue;
				}
			}
			
			// Handle text content
			let textContent = '';
			while (i < formatted.length && formatted[i] !== '<') {
				textContent += formatted[i];
				i++;
			}
			
			if (textContent.trim()) {
				tokens.push({
					type: 'text',
					content: textContent.trim()
				});
			}
		}
		
		// Now process tokens intelligently
		for (let tokenIndex = 0; tokenIndex < tokens.length; tokenIndex++) {
			const token = tokens[tokenIndex];
			
			if (token.type === 'comment' || token.type === 'cdata') {
				result.push(indent.repeat(depth) + token.content);
			} else if (token.type === 'tag') {
				const tag = token.content;
				
				if (tag.startsWith('<?') || tag.startsWith('<!')) {
					// XML declaration or DOCTYPE - no indentation change
					result.push(tag);
				} else if (tag.startsWith('</')) {
					// Closing tag - decrease depth first, then add
					depth = Math.max(0, depth - 1);
					result.push(indent.repeat(depth) + tag);
				} else if (tag.endsWith('/>')) {
					// Self-closing tag - no depth change
					result.push(indent.repeat(depth) + tag);
				} else {
					// Opening tag - check if it's a simple text-only element
					const isSimpleElement = tokenIndex + 2 < tokens.length &&
						tokens[tokenIndex + 1].type === 'text' &&
						tokens[tokenIndex + 2].type === 'tag' &&
						tokens[tokenIndex + 2].content.startsWith('</');
					
					if (isSimpleElement) {
						// Simple element: <tag>text</tag> - keep on one line
						const textToken = tokens[tokenIndex + 1];
						const closingTag = tokens[tokenIndex + 2];
						
						result.push(indent.repeat(depth) + tag + textToken.content + closingTag.content);
						
						// Skip the next two tokens since we processed them
						tokenIndex += 2;
					} else {
						// Complex element - add tag and increase depth
						result.push(indent.repeat(depth) + tag);
						depth++;
					}
				}
			} else if (token.type === 'text') {
				// Standalone text (not part of a simple element)
				result.push(indent.repeat(depth) + token.content);
			}
		}
		
		// Join and clean up
		return result
			.filter(line => line.trim()) // Remove empty lines
			.map(line => line.trimEnd()) // Remove trailing spaces
			.join('\n');
			
	} catch (error) {
		console.error('Error formatting XML:', error);
		// Fallback to simple formatting
		return xml
			.replace(/>\s*</g, '>\n<')
			.split('\n')
			.map(line => line.trim())
			.filter(line => line)
			.join('\n');
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
			fileType: getFileTypeDescription(fileExt),
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
		vscode.window.showErrorMessage(`Failed to get file information: ${error}`);
	}
}

function getFileTypeDescription(extension: string): string {
	const types: { [key: string]: string } = {
		'.docx': 'Microsoft Word Document',
		'.xlsx': 'Microsoft Excel Spreadsheet',
		'.pptx': 'Microsoft PowerPoint Presentation',
		'.dotx': 'Microsoft Word Template',
		'.xltx': 'Microsoft Excel Template',
		'.potx': 'Microsoft PowerPoint Template'
	};
	return types[extension] || 'OpenXML Document';
}

function generateFileInfoHTML(fileInfo: any): string {
	const formatSize = (bytes: number): string => {
		const sizes = ['B', 'KB', 'MB', 'GB'];
		if (bytes === 0) {
			return '0 B';
		}
		const i = Math.floor(Math.log(bytes) / Math.log(1024));
		return Math.round(bytes / Math.pow(1024, i) * 100) / 100 + ' ' + sizes[i];
	};

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
				<tr><td>File Size</td><td>${formatSize(fileInfo.fileSize)}</td></tr>
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
	// Clean up
	vscode.commands.executeCommand('setContext', 'openxmlFileOpen', false);
}
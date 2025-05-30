import * as vscode from 'vscode';
import * as yauzl from 'yauzl';
import * as yazl from 'yazl';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';

import { EXTENSION_CONFIG, SYNC_CONFIG, MIME_TYPES } from './constants';
import { logger } from './utils/logger';
import { FileUtils } from './utils/fileUtils';

interface OpenXMLTempData {
    originalPath: string;
    tempDir: string;
    files: Map<string, Uint8Array>;
    tempFiles: Map<string, string>; // internal path -> temp file path
    modified: Set<string>;
    watchers: fs.FSWatcher[];
    sourceFileWatcher?: fs.FSWatcher; // Watch the original OpenXML file for external changes
    lastModified: number; // Track when the original file was last modified
}

export class OpenXMLFileSystemProvider implements vscode.FileSystemProvider {
    private readonly _emitter = new vscode.EventEmitter<vscode.FileChangeEvent[]>();
    readonly onDidChangeFile: vscode.Event<vscode.FileChangeEvent[]> = this._emitter.event;

    private readonly openXMLFiles = new Map<string, OpenXMLTempData>();
    private readonly pathMappings = new Map<string, string>(); // Map short IDs to full paths

    watch(uri: vscode.Uri, options: { recursive: boolean; excludes: string[] }): vscode.Disposable {
        // For now, we'll implement a simple no-op watcher
        return new vscode.Disposable(() => {});
    }

    async stat(uri: vscode.Uri): Promise<vscode.FileStat> {
        const { openXMLPath, filePath } = this.parseUri(uri);
        const openXMLData = this.openXMLFiles.get(openXMLPath);
        
        if (!openXMLData) {
            throw vscode.FileSystemError.FileNotFound(uri);
        }

        if (!openXMLData.files.has(filePath)) {
            throw vscode.FileSystemError.FileNotFound(uri);
        }

        const content = openXMLData.files.get(filePath)!;
        const fileName = filePath.toLowerCase();
        
        // Log MIME type information for debugging
        if (FileUtils.isXMLFile(fileName)) {
            const mimeType = fileName.endsWith('.rels') ? MIME_TYPES.rels : MIME_TYPES.xml;
            logger.debug(`XML file stat: ${filePath}`, { mimeType });
        }
        
        return {
            type: vscode.FileType.File,
            ctime: Date.now(),
            mtime: openXMLData.modified.has(filePath) ? Date.now() : Date.now() - 10000,
            size: content.length
        };
    }

    async readDirectory(uri: vscode.Uri): Promise<[string, vscode.FileType][]> {
        // For this implementation, we'll focus on individual file access
        // Directory listing is handled by the tree provider
        return [];
    }

    createDirectory(uri: vscode.Uri): void {
        // Not implemented for this use case
        throw vscode.FileSystemError.NoPermissions(uri);
    }

    async readFile(uri: vscode.Uri): Promise<Uint8Array> {
        logger.debug('ReadFile called for URI', { uri: uri.toString() });
        
        try {
            const { openXMLPath, filePath } = this.parseUri(uri);
            logger.debug('Parsed URI', { openXMLPath, filePath });
            
            // Try to find the OpenXML data with different path formats
            let openXMLData = this.openXMLFiles.get(openXMLPath);
            
            if (!openXMLData) {
                openXMLData = this.findOpenXMLDataByPath(openXMLPath);
            }
            
            if (!openXMLData) {
                logger.error('OpenXML data not found', { 
                    requestedPath: openXMLPath,
                    availablePaths: Array.from(this.openXMLFiles.keys())
                });
                throw vscode.FileSystemError.FileNotFound(uri);
            }

            logger.debug('Found OpenXML data, checking for internal file', { 
                filePath,
                availableFiles: Array.from(openXMLData.files.keys())
            });

            // Check if file exists
            let content = openXMLData.files.get(filePath);
            
            if (!content) {
                content = this.findFileByPath(openXMLData, filePath);
            }
            
            if (!content) {
                logger.error('File not found in archive', { filePath });
                throw vscode.FileSystemError.FileNotFound(uri);
            }

            logger.success('Successfully read file', { filePath, size: content.length });
            return content;
            
        } catch (error) {
            logger.error('Error in readFile', error);
            if (error instanceof vscode.FileSystemError) {
                throw error;
            }
            throw vscode.FileSystemError.FileNotFound(uri);
        }
    }

    private findOpenXMLDataByPath(openXMLPath: string): OpenXMLTempData | undefined {
        // Try normalized Windows path
        const normalizedPath = openXMLPath.replace(/\//g, '\\');
        let data = this.openXMLFiles.get(normalizedPath);
        if (data) {
            logger.debug('Found with normalized Windows path', { normalizedPath });
            return data;
        }
        
        // Try normalized Unix path
        const unixPath = openXMLPath.replace(/\\/g, '/');
        data = this.openXMLFiles.get(unixPath);
        if (data) {
            logger.debug('Found with normalized Unix path', { unixPath });
            return data;
        }
        
        return undefined;
    }

    private findFileByPath(openXMLData: OpenXMLTempData, filePath: string): Uint8Array | undefined {
        logger.debug('File not found, trying case-insensitive search');
        const lowerFilePath = filePath.toLowerCase();
        
        for (const [key, value] of openXMLData.files) {
            if (key.toLowerCase() === lowerFilePath) {
                logger.success('Found file with different case', { original: key, requested: filePath });
                return value;
            }
        }
        
        return undefined;
    }

    async writeFile(uri: vscode.Uri, content: Uint8Array, options: { create: boolean; overwrite: boolean }): Promise<void> {
        const { openXMLPath, filePath } = this.parseUri(uri);
        let openXMLData = this.openXMLFiles.get(openXMLPath);
        
        if (!openXMLData) {
            throw vscode.FileSystemError.FileNotFound(uri);
        }

        // Store the modified content
        openXMLData.files.set(filePath, content);
        openXMLData.modified.add(filePath);

        // Fire change event
        this._emitter.fire([{
            type: vscode.FileChangeType.Changed,
            uri: uri
        }]);

        // Auto-save after a short delay
        setTimeout(() => this.autoSave(openXMLPath), 1000);
    }

    delete(uri: vscode.Uri, options: { recursive: boolean }): void {
        // Not implemented for this use case
        throw vscode.FileSystemError.NoPermissions(uri);
    }

    rename(oldUri: vscode.Uri, newUri: vscode.Uri, options: { overwrite: boolean }): void {
        // Not implemented for this use case
        throw vscode.FileSystemError.NoPermissions(oldUri);
    }

    async loadOpenXMLFile(originalPath: string): Promise<void> {
        logger.info('Loading OpenXML file', { path: originalPath });
        
        // Remove existing data if reloading
        if (this.openXMLFiles.has(originalPath)) {
            logger.debug('Removing existing data for reload');
            await this.cleanupTempFiles(originalPath);
            this.openXMLFiles.delete(originalPath);
        }
        
        // Get the current modification time of the source file
        const sourceStats = fs.statSync(originalPath);
        const lastModified = sourceStats.mtime.getTime();
        
        // Create temporary directory for this OpenXML file
        const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), EXTENSION_CONFIG.tempDirPrefix));
        logger.debug('Created temp directory', { tempDir });
        
        const openXMLData: OpenXMLTempData = {
            originalPath: originalPath,
            tempDir: tempDir,
            files: new Map<string, Uint8Array>(),
            tempFiles: new Map<string, string>(),
            modified: new Set<string>(),
            watchers: [],
            lastModified: lastModified
        };

        // Load all files from the OpenXML archive and create temp files
        await this.extractArchiveFiles(originalPath, openXMLData);

        // Set up source file watcher to detect external changes
        this.setupSourceFileWatcher(openXMLData);

        this.openXMLFiles.set(originalPath, openXMLData);
        logger.success('OpenXML file loaded successfully', {
            path: originalPath,
            filesCount: openXMLData.files.size,
            tempFilesCount: openXMLData.tempFiles.size
        });
    }

    private setupSourceFileWatcher(openXMLData: OpenXMLTempData): void {
        try {
            logger.debug('Setting up source file watcher', { path: openXMLData.originalPath });
            
            let debounceTimer: NodeJS.Timeout | null = null;
            
            openXMLData.sourceFileWatcher = fs.watch(openXMLData.originalPath, (eventType, filename) => {
                logger.debug('Source file event detected', { 
                    eventType, 
                    filename, 
                    path: openXMLData.originalPath 
                });
                
                // Debounce the file change events
                if (debounceTimer) {
                    clearTimeout(debounceTimer);
                }
                
                debounceTimer = setTimeout(async () => {
                    try {
                        await this.handleSourceFileChange(openXMLData);
                    } catch (error) {
                        logger.error('Failed to handle source file change', error);
                    }
                }, SYNC_CONFIG.debounceDelayMs);
            });
            
            openXMLData.sourceFileWatcher.on('error', (error) => {
                logger.error('Source file watcher error', error, { path: openXMLData.originalPath });
            });
            
            logger.success('Source file watcher set up successfully', { path: openXMLData.originalPath });
            
        } catch (error) {
            logger.error('Failed to set up source file watcher', error, { path: openXMLData.originalPath });
        }
    }

    private async handleSourceFileChange(openXMLData: OpenXMLTempData): Promise<void> {
        try {
            // Check if the source file still exists
            if (!fs.existsSync(openXMLData.originalPath)) {
                logger.warn('Source file was deleted', { path: openXMLData.originalPath });
                vscode.window.showWarningMessage(
                    `Source file was deleted: ${path.basename(openXMLData.originalPath)}`
                );
                return;
            }
            
            // Check if the file was actually modified
            const currentStats = fs.statSync(openXMLData.originalPath);
            const currentModified = currentStats.mtime.getTime();
            
            if (currentModified <= openXMLData.lastModified) {
                logger.debug('Source file not actually modified, ignoring event');
                return;
            }
            
            logger.info('Source file was modified externally', { 
                path: openXMLData.originalPath,
                lastModified: new Date(openXMLData.lastModified),
                currentModified: new Date(currentModified)
            });
            
            // Check if we have unsaved changes
            const hasUnsavedChanges = openXMLData.modified.size > 0;
            
            if (hasUnsavedChanges) {
                // Ask user what to do with conflicting changes
                const result = await vscode.window.showWarningMessage(
                    `The file "${path.basename(openXMLData.originalPath)}" was modified externally but you have unsaved changes. What would you like to do?`,
                    {
                        modal: true,
                        detail: `Modified files: ${Array.from(openXMLData.modified).join(', ')}`
                    },
                    'Reload (Lose Changes)',
                    'Keep Changes',
                    'Save First, Then Reload'
                );
                
                switch (result) {
                    case 'Reload (Lose Changes)':
                        await this.reloadFromSource(openXMLData);
                        vscode.window.showInformationMessage('File reloaded from external changes. Local changes were discarded.');
                        break;
                        
                    case 'Keep Changes':
                        // Update the lastModified time but keep our changes
                        openXMLData.lastModified = currentModified;
                        vscode.window.showInformationMessage('Keeping local changes. External changes ignored.');
                        break;
                        
                    case 'Save First, Then Reload':
                        await this.saveChanges(openXMLData.originalPath);
                        await this.reloadFromSource(openXMLData);
                        vscode.window.showInformationMessage('Changes saved, then file reloaded with external changes.');
                        break;
                        
                    default:
                        // User cancelled - keep current state
                        break;
                }
            } else {
                // No unsaved changes - safe to reload automatically
                await this.reloadFromSource(openXMLData);
                vscode.window.showInformationMessage(
                    `File "${path.basename(openXMLData.originalPath)}" was updated with external changes.`
                );
            }
            
        } catch (error) {
            logger.error('Failed to handle source file change', error);
            vscode.window.showErrorMessage(`Failed to handle external file changes: ${error}`);
        }
    }

    private async reloadFromSource(openXMLData: OpenXMLTempData): Promise<void> {
        logger.info('Reloading from source file', { path: openXMLData.originalPath });
        
        try {
            // Update last modified time
            const sourceStats = fs.statSync(openXMLData.originalPath);
            openXMLData.lastModified = sourceStats.mtime.getTime();
            
            // Clear current data
            openXMLData.files.clear();
            openXMLData.modified.clear();
            
            // Clean up existing temp files
            for (const [internalPath, tempFilePath] of openXMLData.tempFiles) {
                try {
                    if (fs.existsSync(tempFilePath)) {
                        fs.unlinkSync(tempFilePath);
                    }
                } catch (error) {
                    logger.warn('Failed to delete temp file during reload', error, { tempFilePath });
                }
            }
            openXMLData.tempFiles.clear();
            
            // Close existing watchers (except source file watcher)
            for (const watcher of openXMLData.watchers) {
                watcher.close();
            }
            openXMLData.watchers = [];
            
            // Reload archive files
            await this.extractArchiveFiles(openXMLData.originalPath, openXMLData);
            
            logger.success('Successfully reloaded from source', { 
                path: openXMLData.originalPath,
                filesCount: openXMLData.files.size
            });
            
            // Notify VS Code that files may have changed
            this._emitter.fire([{
                type: vscode.FileChangeType.Changed,
                uri: vscode.Uri.file(openXMLData.originalPath)
            }]);
            
        } catch (error) {
            logger.error('Failed to reload from source', error);
            throw error;
        }
    }

    private async extractArchiveFiles(originalPath: string, openXMLData: OpenXMLTempData): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            yauzl.open(originalPath, { lazyEntries: true }, (err: Error | null, zipfile?: yauzl.ZipFile) => {
                if (err || !zipfile) {
                    logger.error('Failed to open OpenXML file', err);
                    reject(err || new Error('Failed to open OpenXML file'));
                    return;
                }

                let totalFiles = 0;
                let processedFiles = 0;
                let pendingReads = 0;

                const checkComplete = () => {
                    if (pendingReads === 0 && processedFiles === totalFiles) {
                        logger.success('Successfully loaded files from OpenXML archive', { 
                            totalFiles: openXMLData.files.size 
                        });
                        resolve();
                    }
                };

                zipfile.readEntry();

                zipfile.on('entry', (entry: yauzl.Entry) => {
                    if (!entry.fileName.endsWith('/')) {
                        // It's a file
                        totalFiles++;
                        pendingReads++;
                        logger.debug('Processing file', { fileName: entry.fileName });
                        
                        zipfile.openReadStream(entry, (err: Error | null, readStream?: NodeJS.ReadableStream) => {
                            if (err || !readStream) {
                                logger.error('Failed to read entry', err, { fileName: entry.fileName });
                                pendingReads--;
                                processedFiles++;
                                zipfile.readEntry();
                                checkComplete();
                                return;
                            }

                            const chunks: Buffer[] = [];
                            
                            readStream.on('data', (chunk: Buffer) => {
                                chunks.push(chunk);
                            });

                            readStream.on('end', async () => {
                                const content = Buffer.concat(chunks);
                                openXMLData.files.set(entry.fileName, new Uint8Array(content));
                                
                                // Create temporary file on disk for XML files
                                if (FileUtils.isXMLFile(entry.fileName)) {
                                    await this.createTempFile(openXMLData, entry.fileName, content);
                                }
                                
                                logger.debug('Loaded file', { 
                                    fileName: entry.fileName, 
                                    size: content.length 
                                });
                                
                                pendingReads--;
                                processedFiles++;
                                zipfile.readEntry();
                                checkComplete();
                            });

                            readStream.on('error', (err: Error) => {
                                logger.error('Stream error for file', err, { fileName: entry.fileName });
                                pendingReads--;
                                processedFiles++;
                                zipfile.readEntry();
                                checkComplete();
                            });
                        });
                    } else {
                        // It's a directory, skip
                        logger.debug('Skipping directory', { dirName: entry.fileName });
                        zipfile.readEntry();
                    }
                });

                zipfile.on('end', () => {
                    logger.debug('Finished scanning archive', { totalFiles });
                    checkComplete();
                });

                zipfile.on('error', (err: Error) => {
                    logger.error('ZipFile error', err);
                    reject(err);
                });
            });
        });
    }

    private async createTempFile(openXMLData: OpenXMLTempData, internalPath: string, content: Buffer): Promise<void> {
        try {
            // Create directory structure in temp folder
            const tempFilePath = path.join(openXMLData.tempDir, internalPath);
            const tempDir = path.dirname(tempFilePath);
            
            // Ensure directory exists
            fs.mkdirSync(tempDir, { recursive: true });
            
            // Write file to temp location
            fs.writeFileSync(tempFilePath, content);
            
            // Store mapping
            openXMLData.tempFiles.set(internalPath, tempFilePath);
            
            // Set up Node.js file watcher with debouncing
            let timeoutId: NodeJS.Timeout | null = null;
            
            const watcher = fs.watch(tempFilePath, (eventType, filename) => {
                logger.debug('Temp file event', { eventType, path: tempFilePath });
                
                // Debounce - wait before syncing
                if (timeoutId) {
                    clearTimeout(timeoutId);
                }
                
                timeoutId = setTimeout(async () => {
                    try {
                        await this.syncTempFileBack(openXMLData, internalPath);
                    } catch (error) {
                        logger.error('Failed to sync after file change', error);
                    }
                }, SYNC_CONFIG.debounceDelayMs);
            });
            
            watcher.on('error', (error) => {
                logger.error('Watcher error', error, { path: tempFilePath });
            });
            
            openXMLData.watchers.push(watcher);
            
            logger.debug('Created temp file with watcher', { 
                internalPath, 
                tempFilePath 
            });
            
        } catch (error) {
            logger.error('Failed to create temp file', error, { internalPath });
        }
    }

    private async syncTempFileBack(openXMLData: OpenXMLTempData, internalPath: string): Promise<void> {
        try {
            const tempFilePath = openXMLData.tempFiles.get(internalPath);
            if (!tempFilePath || !fs.existsSync(tempFilePath)) {
                logger.warn('Temp file not found for sync', { tempFilePath });
                return;
            }
            
            // Get file stats to check if it was recently modified
            const stats = fs.statSync(tempFilePath);
            const now = Date.now();
            const modifiedTime = stats.mtime.getTime();
            
            // Only sync if file was modified recently
            if (now - modifiedTime > SYNC_CONFIG.fileAgeThresholdMs) {
                logger.debug('Temp file too old, skipping sync', { tempFilePath });
                return;
            }
            
            // Read updated content from temp file
            const updatedContent = fs.readFileSync(tempFilePath);
            const originalContent = openXMLData.files.get(internalPath);
            
            // Compare content to see if it actually changed
            if (originalContent && Buffer.from(originalContent).equals(updatedContent)) {
                logger.debug('No content changes detected', { internalPath });
                return;
            }
            
            // Update in-memory data
            openXMLData.files.set(internalPath, new Uint8Array(updatedContent));
            openXMLData.modified.add(internalPath);
            
            logger.success('Synced temp file back', { 
                internalPath, 
                size: updatedContent.length 
            });
            
            // Auto-save after delay for batch processing
            setTimeout(() => this.autoSave(openXMLData.originalPath), SYNC_CONFIG.autoSaveDelayMs);
            
        } catch (error) {
            logger.error('Failed to sync temp file', error, { internalPath });
        }
    }

    async cleanupTempFiles(openXMLPath: string): Promise<void> {
        const openXMLData = this.openXMLFiles.get(openXMLPath);
        if (!openXMLData) {
            return;
        }
        
        try {
            // Close source file watcher
            if (openXMLData.sourceFileWatcher) {
                openXMLData.sourceFileWatcher.close();
                openXMLData.sourceFileWatcher = undefined;
                logger.debug('Closed source file watcher', { path: openXMLPath });
            }
            
            // Dispose temp file watchers
            openXMLData.watchers.forEach(watcher => watcher.close());
            openXMLData.watchers = [];
            
            // Remove temp directory
            if (fs.existsSync(openXMLData.tempDir)) {
                fs.rmSync(openXMLData.tempDir, { recursive: true, force: true });
                logger.info('Cleaned up temp directory', { tempDir: openXMLData.tempDir });
            }
            
        } catch (error) {
            logger.error('Failed to cleanup temp files', error);
        }
    }

    getTempFilePath(openXMLPath: string, internalPath: string): string | undefined {
        const openXMLData = this.openXMLFiles.get(openXMLPath);
        return openXMLData?.tempFiles.get(internalPath);
    }

    async saveChanges(openXMLPath: string): Promise<void> {
        const openXMLData = this.openXMLFiles.get(openXMLPath);
        if (!openXMLData) {
            logger.warn('No OpenXML data found for path', { openXMLPath });
            return;
        }

        logger.info('Starting save process', { 
            openXMLPath,
            modifiedFilesBefore: Array.from(openXMLData.modified),
            tempFilesCount: openXMLData.tempFiles.size
        });

        // Force sync all temporary files before saving
        await this.syncAllTempFiles(openXMLData);

        logger.debug('Modified files after sync', { 
            modifiedFilesAfter: Array.from(openXMLData.modified) 
        });

        if (openXMLData.modified.size === 0) {
            logger.info('No changes to save');
            return;
        }

        // Create a backup
        const backupPath = openXMLPath + '.backup';
        fs.copyFileSync(openXMLPath, backupPath);
        logger.debug('Created backup', { backupPath });

        try {
            // Create new ZIP file
            const zipFile = new yazl.ZipFile();
            
            logger.debug('Adding files to ZIP');
            // Add all files to the new ZIP
            for (const [filePath, content] of openXMLData.files) {
                zipFile.addBuffer(Buffer.from(content), filePath);
                const isModified = openXMLData.modified.has(filePath);
                logger.debug('Added file to ZIP', { 
                    filePath, 
                    size: content.length, 
                    isModified 
                });
            }

            zipFile.end();

            // Write to temporary file first
            const tempPath = openXMLPath + '.tmp';
            const writeStream = fs.createWriteStream(tempPath);
            
            await new Promise<void>((resolve, reject) => {
                zipFile.outputStream.pipe(writeStream);
                writeStream.on('close', resolve);
                writeStream.on('error', reject);
            });

            // Replace original file
            fs.unlinkSync(openXMLPath);
            fs.renameSync(tempPath, openXMLPath);

            // Update the last modified time to prevent source watcher from triggering
            const newStats = fs.statSync(openXMLPath);
            openXMLData.lastModified = newStats.mtime.getTime();

            // Clean up backup
            fs.unlinkSync(backupPath);

            // Clear modified set
            openXMLData.modified.clear();

            logger.success('OpenXML file saved successfully', { openXMLPath });
            vscode.window.showInformationMessage(`OpenXML file saved: ${path.basename(openXMLPath)}`);

        } catch (error) {
            logger.error('Save failed', error);
            // Restore from backup
            if (fs.existsSync(backupPath)) {
                fs.copyFileSync(backupPath, openXMLPath);
                fs.unlinkSync(backupPath);
                logger.info('Restored from backup');
            }
            throw error;
        }
    }

    private async syncAllTempFiles(openXMLData: OpenXMLTempData): Promise<void> {
        logger.progress('Force syncing all temporary files...');
        
        for (const [internalPath, tempFilePath] of openXMLData.tempFiles) {
            try {
                if (fs.existsSync(tempFilePath)) {
                    const originalContent = openXMLData.files.get(internalPath);
                    
                    if (originalContent) {
                        // Compare file modification time or content
                        const updatedContent = fs.readFileSync(tempFilePath);
                        
                        // Compare content to detect changes
                        if (!Buffer.from(originalContent).equals(updatedContent)) {
                            logger.debug('Detected changes in temp file', { internalPath });
                            openXMLData.files.set(internalPath, new Uint8Array(updatedContent));
                            openXMLData.modified.add(internalPath);
                        } else {
                            logger.debug('No changes in temp file', { internalPath });
                        }
                    }
                } else {
                    logger.warn('Temp file not found', { tempFilePath });
                }
            } catch (error) {
                logger.error('Failed to sync temp file', error, { internalPath });
            }
        }
        
        logger.success('Finished syncing temporary files');
    }

    private async autoSave(openXMLPath: string): Promise<void> {
        try {
            await this.saveChanges(openXMLPath);
        } catch (error) {
            logger.error('Auto-save failed', error);
            vscode.window.showWarningMessage(`Auto-save failed: ${error}`);
        }
    }

    private parseUri(uri: vscode.Uri): { openXMLPath: string; filePath: string } {
        // Parse URI format: openxml:/PresentationName[id]/internal-file-path
        try {
            logger.debug('Parsing URI', { uri: uri.toString() });
            logger.debug('URI path', { path: uri.path });
            
            // Remove leading slash and split path
            const pathParts = uri.path.substring(1).split('/');
            
            if (pathParts.length < 2) {
                throw new Error('Invalid URI format - not enough path parts');
            }
            
            // First part is the readable ID (PresentationName[hash])
            const readableId = pathParts[0];
            // Rest of path parts form the internal file path
            const filePath = pathParts.slice(1).join('/');
            
            // Look up the full path from our mapping
            const openXMLPath = this.pathMappings.get(readableId);
            
            if (!openXMLPath) {
                logger.error('Readable ID not found in mappings', { readableId });
                logger.debug('Available mappings', { 
                    mappings: Array.from(this.pathMappings.entries()) 
                });
                throw new Error(`Readable ID not found: ${readableId}`);
            }
            
            logger.success('Parsed readable URI', {
                readableId,
                openXMLPath,
                filePath
            });
            
            return { openXMLPath, filePath };
            
        } catch (error) {
            logger.error('Failed to parse URI', error, { uri: uri.toString() });
            throw new Error(`Invalid URI format: ${uri.toString()}`);
        }
    }

    createUri(openXMLPath: string, filePath: string): vscode.Uri {
        // Create human-readable URI format for beautiful tab names
        // Goal: Tab shows "PresentationName → path/to/file.xml"
        
        // Get the presentation name without extension
        const presentationName = FileUtils.extractPresentationName(openXMLPath);
        
        // Create a unique identifier for mapping (short and clean)
        const pathHash = FileUtils.createShortHash(openXMLPath);
        
        // Create a readable identifier
        const readableId = `${presentationName}[${pathHash}]`;
        
        // Store the mapping for internal resolution
        this.pathMappings.set(readableId, openXMLPath);
        
        // Create beautiful URI that will show nicely in tabs
        // Format: openxml:/PresentationName[id]/path/to/file.xml
        // Tab will display: "PresentationName[id]/path/to/file.xml"
        
        // Clean up the file path for display (remove leading slashes)
        const cleanFilePath = filePath.startsWith('/') ? filePath.substring(1) : filePath;
        
        const uri = `openxml:/${readableId}/${cleanFilePath}`;
        
        logger.debug('Creating beautiful URI', {
            presentation: presentationName,
            filePath: cleanFilePath,
            tabDisplay: `${readableId}/${cleanFilePath}`,
            fullUri: uri
        });
        
        return vscode.Uri.parse(uri);
    }

    hasUnsavedChanges(openXMLPath: string): boolean {
        const openXMLData = this.openXMLFiles.get(openXMLPath);
        return openXMLData ? openXMLData.modified.size > 0 : false;
    }

    getModifiedFiles(openXMLPath: string): string[] {
        const openXMLData = this.openXMLFiles.get(openXMLPath);
        return openXMLData ? Array.from(openXMLData.modified) : [];
    }
} 
import * as vscode from 'vscode';
import * as yauzl from 'yauzl';
import * as yazl from 'yazl';
import * as fs from 'fs';
import * as path from 'path';

export class OpenXMLFileSystemProvider implements vscode.FileSystemProvider {
    private _emitter = new vscode.EventEmitter<vscode.FileChangeEvent[]>();
    readonly onDidChangeFile: vscode.Event<vscode.FileChangeEvent[]> = this._emitter.event;

    private openXMLFiles = new Map<string, { originalPath: string; files: Map<string, Uint8Array>; modified: Set<string> }>();
    private pathMappings = new Map<string, string>(); // Map short IDs to full paths

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
        
        // Enhanced file stat with proper type detection for XML files
        let fileType = vscode.FileType.File;
        
        // For XML files, ensure they are properly recognized
        const fileName = filePath.toLowerCase();
        if (fileName.endsWith('.xml') || fileName.endsWith('.rels')) {
            // Mark as regular file but with specific properties that might help other extensions
            fileType = vscode.FileType.File;
        }
        
        return {
            type: fileType,
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
        console.log('🔍 ReadFile called for URI:', uri.toString());
        
        try {
            const { openXMLPath, filePath } = this.parseUri(uri);
            console.log('📂 Looking for OpenXML file:', openXMLPath);
            console.log('📄 Looking for internal file:', filePath);
            
            // Try to find the OpenXML data with different path formats
            let openXMLData = this.openXMLFiles.get(openXMLPath);
            
            if (!openXMLData) {
                // Try normalized Windows path
                const normalizedPath = openXMLPath.replace(/\//g, '\\');
                openXMLData = this.openXMLFiles.get(normalizedPath);
                console.log('🔄 Trying normalized Windows path:', normalizedPath);
            }
            
            if (!openXMLData) {
                // Try normalized Unix path
                const unixPath = openXMLPath.replace(/\\/g, '/');
                openXMLData = this.openXMLFiles.get(unixPath);
                console.log('🔄 Trying normalized Unix path:', unixPath);
            }
            
            if (!openXMLData) {
                console.error('❌ OpenXML data not found for any path format');
                console.log('🗂️ Available OpenXML files:');
                Array.from(this.openXMLFiles.keys()).forEach((key, index) => {
                    console.log(`  ${index + 1}. "${key}"`);
                });
                throw vscode.FileSystemError.FileNotFound(uri);
            }

            console.log('✅ Found OpenXML data, checking for internal file:', filePath);
            console.log('📁 Available files in archive:');
            Array.from(openXMLData.files.keys()).forEach((key, index) => {
                console.log(`  ${index + 1}. "${key}"`);
            });

            // Check if file exists
            let content = openXMLData.files.get(filePath);
            
            if (!content) {
                console.log('🔄 File not found, trying case-insensitive search...');
                // Try case-insensitive search
                const lowerFilePath = filePath.toLowerCase();
                for (const [key, value] of openXMLData.files) {
                    if (key.toLowerCase() === lowerFilePath) {
                        console.log(`✅ Found file with different case: "${key}" vs "${filePath}"`);
                        content = value;
                        break;
                    }
                }
            }
            
            if (!content) {
                console.error('❌ File not found even with case-insensitive search:', filePath);
                throw vscode.FileSystemError.FileNotFound(uri);
            }

            console.log('🎉 Successfully read file:', filePath, 'size:', content.length, 'bytes');
            return content;
            
        } catch (error) {
            console.error('💥 Error in readFile:', error);
            if (error instanceof vscode.FileSystemError) {
                throw error;
            }
            throw vscode.FileSystemError.FileNotFound(uri);
        }
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
        console.log('Loading OpenXML file:', originalPath);
        
        // Remove existing data if reloading
        if (this.openXMLFiles.has(originalPath)) {
            console.log('Removing existing data for reload');
            this.openXMLFiles.delete(originalPath);
        }
        
        const openXMLData = {
            originalPath: originalPath,
            files: new Map<string, Uint8Array>(),
            modified: new Set<string>()
        };

        // Load all files from the OpenXML archive
        await new Promise<void>((resolve, reject) => {
            yauzl.open(originalPath, { lazyEntries: true }, (err: Error | null, zipfile?: yauzl.ZipFile) => {
                if (err || !zipfile) {
                    console.error('Failed to open OpenXML file:', err);
                    reject(err || new Error('Failed to open OpenXML file'));
                    return;
                }

                let totalFiles = 0;
                let processedFiles = 0;
                let pendingReads = 0;

                const checkComplete = () => {
                    if (pendingReads === 0 && processedFiles === totalFiles) {
                        console.log(`Successfully loaded ${openXMLData.files.size} files from OpenXML archive`);
                        resolve();
                    }
                };

                zipfile.readEntry();

                zipfile.on('entry', (entry: yauzl.Entry) => {
                    if (!entry.fileName.endsWith('/')) {
                        // It's a file
                        totalFiles++;
                        pendingReads++;
                        console.log(`Processing file: ${entry.fileName}`);
                        
                        zipfile.openReadStream(entry, (err: Error | null, readStream?: NodeJS.ReadableStream) => {
                            if (err || !readStream) {
                                console.error('Failed to read entry:', entry.fileName, err);
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

                            readStream.on('end', () => {
                                const content = Buffer.concat(chunks);
                                openXMLData.files.set(entry.fileName, new Uint8Array(content));
                                console.log(`✅ Loaded file: ${entry.fileName} (${content.length} bytes)`);
                                
                                pendingReads--;
                                processedFiles++;
                                zipfile.readEntry();
                                checkComplete();
                            });

                            readStream.on('error', (err: Error) => {
                                console.error('❌ Stream error for file:', entry.fileName, err);
                                pendingReads--;
                                processedFiles++;
                                zipfile.readEntry();
                                checkComplete();
                            });
                        });
                    } else {
                        // It's a directory, skip
                        console.log(`Skipping directory: ${entry.fileName}`);
                        zipfile.readEntry();
                    }
                });

                zipfile.on('end', () => {
                    console.log(`📁 Finished scanning archive. Found ${totalFiles} files total.`);
                    checkComplete();
                });

                zipfile.on('error', (err: Error) => {
                    console.error('❌ ZipFile error:', err);
                    reject(err);
                });
            });
        });

        this.openXMLFiles.set(originalPath, openXMLData);
        console.log('🎉 OpenXML file stored with key:', originalPath);
        console.log('📋 Final loaded files count:', openXMLData.files.size);
        console.log('📁 All loaded files:');
        Array.from(openXMLData.files.keys()).forEach((key, index) => {
            console.log(`  ${index + 1}. "${key}"`);
        });
    }

    private async loadFileFromOpenXML(openXMLPath: string, filePath: string): Promise<void> {
        const openXMLData = this.openXMLFiles.get(openXMLPath);
        if (!openXMLData) {
            return;
        }

        // File should already be loaded, but just in case
        if (!openXMLData.files.has(filePath)) {
            return new Promise<void>((resolve, reject) => {
                yauzl.open(openXMLData.originalPath, { lazyEntries: true }, (err: Error | null, zipfile?: yauzl.ZipFile) => {
                    if (err || !zipfile) {
                        reject(err || new Error('Failed to open OpenXML file'));
                        return;
                    }

                    zipfile.readEntry();

                    zipfile.on('entry', (entry: yauzl.Entry) => {
                        if (entry.fileName === filePath) {
                            zipfile.openReadStream(entry, (err: Error | null, readStream?: NodeJS.ReadableStream) => {
                                if (err || !readStream) {
                                    reject(err || new Error('Failed to read file stream'));
                                    return;
                                }

                                const chunks: Buffer[] = [];
                                readStream.on('data', (chunk: Buffer) => {
                                    chunks.push(chunk);
                                });

                                readStream.on('end', () => {
                                    const content = Buffer.concat(chunks);
                                    openXMLData.files.set(filePath, new Uint8Array(content));
                                    resolve();
                                });
                            });
                        } else {
                            zipfile.readEntry();
                        }
                    });

                    zipfile.on('end', () => {
                        reject(new Error('File not found in archive'));
                    });

                    zipfile.on('error', (err: Error) => {
                        reject(err);
                    });
                });
            });
        }
    }

    async saveChanges(openXMLPath: string): Promise<void> {
        const openXMLData = this.openXMLFiles.get(openXMLPath);
        if (!openXMLData || openXMLData.modified.size === 0) {
            return;
        }

        // Create a backup
        const backupPath = openXMLPath + '.backup';
        fs.copyFileSync(openXMLPath, backupPath);

        try {
            // Create new ZIP file
            const zipFile = new yazl.ZipFile();
            
            // Add all files to the new ZIP
            for (const [filePath, content] of openXMLData.files) {
                zipFile.addBuffer(Buffer.from(content), filePath);
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

            // Clean up backup
            fs.unlinkSync(backupPath);

            // Clear modified set
            openXMLData.modified.clear();

            vscode.window.showInformationMessage(`OpenXML file saved: ${path.basename(openXMLPath)}`);

        } catch (error) {
            // Restore from backup
            if (fs.existsSync(backupPath)) {
                fs.copyFileSync(backupPath, openXMLPath);
                fs.unlinkSync(backupPath);
            }
            throw error;
        }
    }

    private async autoSave(openXMLPath: string): Promise<void> {
        try {
            await this.saveChanges(openXMLPath);
        } catch (error) {
            console.error('Auto-save failed:', error);
            vscode.window.showWarningMessage(`Auto-save failed: ${error}`);
        }
    }

    private parseUri(uri: vscode.Uri): { openXMLPath: string; filePath: string } {
        // Parse URI format: openxml:/PresentationName[id]/internal-file-path
        try {
            console.log('🔍 Parsing URI:', uri.toString());
            console.log('📋 URI path:', uri.path);
            
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
                console.error('❌ Readable ID not found in mappings:', readableId);
                console.log('🗂️ Available mappings:');
                Array.from(this.pathMappings.entries()).forEach(([id, path], index) => {
                    console.log(`  ${index + 1}. "${id}" -> "${path}"`);
                });
                throw new Error(`Readable ID not found: ${readableId}`);
            }
            
            console.log('✅ Parsed readable URI:');
            console.log('  📝 Readable ID:', readableId);
            console.log('  📁 Resolved OpenXML path:', openXMLPath);
            console.log('  📄 Internal file path:', filePath);
            
            return { openXMLPath, filePath };
            
        } catch (error) {
            console.error('💥 Failed to parse URI:', uri.toString(), error);
            throw new Error(`Invalid URI format: ${uri.toString()}`);
        }
    }

    createUri(openXMLPath: string, filePath: string): vscode.Uri {
        // Create human-readable URI format for beautiful tab names
        // Goal: Tab shows "PresentationName → path/to/file.xml"
        
        // Get the presentation name without extension
        const presentationName = path.basename(openXMLPath, path.extname(openXMLPath));
        
        // Create a unique identifier for mapping (short and clean)
        const pathHash = Buffer.from(openXMLPath, 'utf8').toString('base64')
            .replace(/[/+=]/g, (match) => {
                switch (match) {
                    case '/': return '_';
                    case '+': return '-';
                    case '=': return '';
                    default: return match;
                }
            }).substring(0, 6); // Very short ID
        
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
        
        console.log('Creating beautiful URI:');
        console.log('  📁 Presentation:', presentationName);
        console.log('  📄 File path:', cleanFilePath);
        console.log('  📋 Tab will show:', `${readableId}/${cleanFilePath}`);
        console.log('  🔗 Full URI:', uri);
        
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
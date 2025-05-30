import * as vscode from 'vscode';
import * as path from 'path';
import * as yauzl from 'yauzl';

export interface OpenXMLTreeItem {
    label: string;
    fullPath: string;
    isDirectory: boolean;
    size?: number;
    children?: OpenXMLTreeItem[];
    openXMLPath?: string; // Add reference to parent OpenXML file
    isOpenXMLRoot?: boolean; // Mark if this is a root OpenXML file node
}

export class OpenXMLTreeDataProvider implements vscode.TreeDataProvider<OpenXMLTreeItem> {
    private _onDidChangeTreeData: vscode.EventEmitter<OpenXMLTreeItem | undefined | null | void> = new vscode.EventEmitter<OpenXMLTreeItem | undefined | null | void>();
    readonly onDidChangeTreeData: vscode.Event<OpenXMLTreeItem | undefined | null | void> = this._onDidChangeTreeData.event;

    private openXMLFiles: Map<string, OpenXMLTreeItem[]> = new Map(); // Store multiple files

    constructor() {}

    refresh(): void {
        this._onDidChangeTreeData.fire();
    }

    async addOpenXMLFile(filePath: string): Promise<void> {
        console.log('Adding OpenXML file to tree:', filePath);
        try {
            const treeData = await this.parseOpenXMLStructure(filePath);
            this.openXMLFiles.set(filePath, treeData);
            this.refresh();
            console.log('✅ OpenXML file added to tree successfully');
        } catch (error) {
            console.error('❌ Failed to add OpenXML file to tree:', error);
            throw error;
        }
    }

    async removeOpenXMLFile(filePath: string): Promise<void> {
        console.log('Removing OpenXML file from tree:', filePath);
        this.openXMLFiles.delete(filePath);
        this.refresh();
        console.log('✅ OpenXML file removed from tree');
    }

    async refreshAll(): Promise<void> {
        console.log('Refreshing all OpenXML files in tree');
        const filePaths = Array.from(this.openXMLFiles.keys());
        this.openXMLFiles.clear();
        
        for (const filePath of filePaths) {
            try {
                await this.addOpenXMLFile(filePath);
            } catch (error) {
                console.error('Failed to refresh file:', filePath, error);
            }
        }
    }

    // Legacy method for compatibility
    async loadOpenXMLFile(filePath: string): Promise<void> {
        await this.addOpenXMLFile(filePath);
    }

    getTreeItem(element: OpenXMLTreeItem): vscode.TreeItem {
        const treeItem = new vscode.TreeItem(
            element.label,
            element.isDirectory || element.isOpenXMLRoot ? vscode.TreeItemCollapsibleState.Collapsed : vscode.TreeItemCollapsibleState.None
        );

        if (element.isOpenXMLRoot) {
            // This is a root OpenXML file node
            treeItem.iconPath = new vscode.ThemeIcon('file-zip');
            treeItem.contextValue = 'openxmlFile';
            treeItem.tooltip = `OpenXML File: ${element.openXMLPath}`;
            treeItem.description = path.basename(element.openXMLPath || '');
        } else if (!element.isDirectory) {
            // This is a file inside the archive - set specific icons based on file type
            treeItem.command = {
                command: 'openxmleditor.openXmlFile',
                title: 'Open File',
                arguments: [element]
            };
            treeItem.contextValue = 'xmlFile';
            
            // Set icon based on file type and location
            treeItem.iconPath = this.getFileIcon(element.fullPath);
            
            // Add size info to description
            if (element.size !== undefined) {
                treeItem.description = this.formatSize(element.size);
            }
        } else {
            // This is a directory - set icon based on directory type
            treeItem.iconPath = this.getDirectoryIcon(element.fullPath);
            treeItem.contextValue = 'folder';
        }

        return treeItem;
    }

    getChildren(element?: OpenXMLTreeItem): Thenable<OpenXMLTreeItem[]> {
        if (!element) {
            // Return root nodes (one for each opened OpenXML file)
            const rootNodes: OpenXMLTreeItem[] = [];
            
            for (const [filePath, treeData] of this.openXMLFiles) {
                const rootNode: OpenXMLTreeItem = {
                    label: path.basename(filePath),
                    fullPath: filePath,
                    isDirectory: false,
                    isOpenXMLRoot: true,
                    openXMLPath: filePath,
                    children: treeData
                };
                rootNodes.push(rootNode);
            }
            
            return Promise.resolve(rootNodes);
        }

        if (element.isOpenXMLRoot) {
            // Return children of the OpenXML file
            return Promise.resolve(element.children || []);
        }

        return Promise.resolve(element.children || []);
    }

    private async parseOpenXMLStructure(filePath: string): Promise<OpenXMLTreeItem[]> {
        return new Promise((resolve, reject) => {
            const items: OpenXMLTreeItem[] = [];
            const pathMap = new Map<string, OpenXMLTreeItem>();

            yauzl.open(filePath, { lazyEntries: true }, (err: Error | null, zipfile?: yauzl.ZipFile) => {
                if (err || !zipfile) {
                    reject(err || new Error('Failed to open zip file'));
                    return;
                }

                zipfile.readEntry();

                zipfile.on('entry', (entry: yauzl.Entry) => {
                    const fullPath = entry.fileName;
                    const isDirectory = fullPath.endsWith('/');
                    
                    if (isDirectory) {
                        // Handle directory
                        const dirItem: OpenXMLTreeItem = {
                            label: path.basename(fullPath.slice(0, -1)) || fullPath.slice(0, -1),
                            fullPath: fullPath,
                            isDirectory: true,
                            openXMLPath: filePath,
                            children: []
                        };
                        
                        pathMap.set(fullPath, dirItem);
                        
                        // Add to parent or root
                        const parentPath = this.getParentPath(fullPath);
                        if (parentPath && pathMap.has(parentPath)) {
                            pathMap.get(parentPath)!.children!.push(dirItem);
                        } else if (!parentPath) {
                            items.push(dirItem);
                        }
                    } else {
                        // Handle file
                        const fileItem: OpenXMLTreeItem = {
                            label: path.basename(fullPath),
                            fullPath: fullPath,
                            isDirectory: false,
                            size: entry.uncompressedSize,
                            openXMLPath: filePath
                        };
                        
                        // Ensure parent directories exist
                        const parentPath = this.getParentPath(fullPath + '/');
                        if (parentPath) {
                            this.ensureParentExists(parentPath, pathMap, items, filePath);
                            if (pathMap.has(parentPath)) {
                                pathMap.get(parentPath)!.children!.push(fileItem);
                            }
                        } else {
                            items.push(fileItem);
                        }
                    }

                    zipfile.readEntry();
                });

                zipfile.on('end', () => {
                    // Sort items: directories first, then files
                    this.sortTreeItems(items);
                    resolve(items);
                });

                zipfile.on('error', (err: Error) => {
                    reject(err);
                });
            });
        });
    }

    private getParentPath(fullPath: string): string | null {
        const parts = fullPath.split('/').filter(p => p);
        if (parts.length <= 1) {
            return null;
        }
        return parts.slice(0, -1).join('/') + '/';
    }

    private ensureParentExists(parentPath: string, pathMap: Map<string, OpenXMLTreeItem>, rootItems: OpenXMLTreeItem[], openXMLPath: string): void {
        if (pathMap.has(parentPath)) {
            return;
        }

        const parts = parentPath.split('/').filter(p => p);
        let currentPath = '';
        
        for (let i = 0; i < parts.length; i++) {
            currentPath += parts[i] + '/';
            
            if (!pathMap.has(currentPath)) {
                const dirItem: OpenXMLTreeItem = {
                    label: parts[i],
                    fullPath: currentPath,
                    isDirectory: true,
                    openXMLPath: openXMLPath,
                    children: []
                };
                
                pathMap.set(currentPath, dirItem);
                
                if (i === 0) {
                    rootItems.push(dirItem);
                } else {
                    const parentDir = parts.slice(0, i).join('/') + '/';
                    if (pathMap.has(parentDir)) {
                        pathMap.get(parentDir)!.children!.push(dirItem);
                    }
                }
            }
        }
    }

    private sortTreeItems(items: OpenXMLTreeItem[]): void {
        items.sort((a, b) => {
            // Directories first
            if (a.isDirectory && !b.isDirectory) {
                return -1;
            }
            if (!a.isDirectory && b.isDirectory) {
                return 1;
            }
            // Then alphabetically
            return a.label.localeCompare(b.label);
        });

        // Recursively sort children
        items.forEach(item => {
            if (item.children) {
                this.sortTreeItems(item.children);
            }
        });
    }

    private formatSize(bytes: number): string {
        const sizes = ['B', 'KB', 'MB', 'GB'];
        if (bytes === 0) {
            return '0 B';
        }
        const i = Math.floor(Math.log(bytes) / Math.log(1024));
        return Math.round(bytes / Math.pow(1024, i) * 100) / 100 + ' ' + sizes[i];
    }

    getCurrentFilePaths(): string[] {
        return Array.from(this.openXMLFiles.keys());
    }

    // Legacy method for compatibility
    getCurrentFilePath(): string | undefined {
        const paths = this.getCurrentFilePaths();
        return paths.length > 0 ? paths[0] : undefined;
    }

    private getFileIcon(filePath: string): vscode.ThemeIcon {
        const fileName = path.basename(filePath).toLowerCase();
        const fileExt = path.extname(filePath).toLowerCase();
        const dirPath = path.dirname(filePath).toLowerCase();
        
        // Special icons for different types of XML files
        
        // Relationship files
        if (fileExt === '.rels' || fileName.includes('.rels')) {
            return new vscode.ThemeIcon('link', new vscode.ThemeColor('charts.blue'));
        }
        
        // PowerPoint specific files
        if (dirPath.includes('ppt')) {
            if (dirPath.includes('slides') && fileName.startsWith('slide')) {
                return new vscode.ThemeIcon('preview', new vscode.ThemeColor('charts.green'));
            }
            if (dirPath.includes('slideLayouts')) {
                return new vscode.ThemeIcon('layout', new vscode.ThemeColor('charts.orange'));
            }
            if (dirPath.includes('slideMasters')) {
                return new vscode.ThemeIcon('symbol-ruler', new vscode.ThemeColor('charts.purple'));
            }
            if (fileName === 'presentation.xml') {
                return new vscode.ThemeIcon('device-desktop', new vscode.ThemeColor('charts.red'));
            }
        }
        
        // Word specific files
        if (dirPath.includes('word')) {
            if (fileName === 'document.xml') {
                return new vscode.ThemeIcon('file-text', new vscode.ThemeColor('charts.blue'));
            }
            if (fileName === 'styles.xml') {
                return new vscode.ThemeIcon('symbol-color', new vscode.ThemeColor('charts.orange'));
            }
            if (fileName === 'settings.xml') {
                return new vscode.ThemeIcon('settings-gear', new vscode.ThemeColor('charts.yellow'));
            }
        }
        
        // Excel specific files
        if (dirPath.includes('xl')) {
            if (dirPath.includes('worksheets')) {
                return new vscode.ThemeIcon('table', new vscode.ThemeColor('charts.green'));
            }
            if (fileName === 'workbook.xml') {
                return new vscode.ThemeIcon('book', new vscode.ThemeColor('charts.blue'));
            }
            if (fileName === 'sharedStrings.xml') {
                return new vscode.ThemeIcon('symbol-string', new vscode.ThemeColor('charts.purple'));
            }
        }
        
        // Theme files
        if (dirPath.includes('theme')) {
            return new vscode.ThemeIcon('paintcan', new vscode.ThemeColor('charts.purple'));
        }
        
        // Media files directory
        if (dirPath.includes('media')) {
            if (fileName.includes('image')) {
                return new vscode.ThemeIcon('file-media', new vscode.ThemeColor('charts.green'));
            }
        }
        
        // Core properties and metadata
        if (fileName === 'core.xml' || fileName === 'app.xml') {
            return new vscode.ThemeIcon('info', new vscode.ThemeColor('charts.blue'));
        }
        
        // Content types
        if (fileName === '[content_types].xml') {
            return new vscode.ThemeIcon('symbol-misc', new vscode.ThemeColor('charts.orange'));
        }
        
        // Default XML file icon
        return new vscode.ThemeIcon('file-code', new vscode.ThemeColor('charts.gray'));
    }

    private getDirectoryIcon(dirPath: string): vscode.ThemeIcon {
        const dirName = path.basename(dirPath).toLowerCase();
        const fullPath = dirPath.toLowerCase();
        
        // PowerPoint directories
        if (fullPath.includes('ppt')) {
            if (dirName === 'slides') {
                return new vscode.ThemeIcon('folder', new vscode.ThemeColor('charts.green'));
            }
            if (dirName === 'slidelayouts') {
                return new vscode.ThemeIcon('folder', new vscode.ThemeColor('charts.orange'));
            }
            if (dirName === 'slidemasters') {
                return new vscode.ThemeIcon('folder', new vscode.ThemeColor('charts.purple'));
            }
            if (dirName === 'theme') {
                return new vscode.ThemeIcon('folder', new vscode.ThemeColor('charts.purple'));
            }
        }
        
        // Word directories
        if (fullPath.includes('word') && dirName === 'word') {
            return new vscode.ThemeIcon('folder', new vscode.ThemeColor('charts.blue'));
        }
        
        // Excel directories
        if (fullPath.includes('xl')) {
            if (dirName === 'worksheets') {
                return new vscode.ThemeIcon('folder', new vscode.ThemeColor('charts.green'));
            }
            if (dirName === 'xl') {
                return new vscode.ThemeIcon('folder', new vscode.ThemeColor('charts.green'));
            }
        }
        
        // Special directories
        if (dirName === '_rels') {
            return new vscode.ThemeIcon('folder', new vscode.ThemeColor('charts.blue'));
        }
        
        if (dirName === 'media') {
            return new vscode.ThemeIcon('folder', new vscode.ThemeColor('charts.yellow'));
        }
        
        if (dirName === 'docprops') {
            return new vscode.ThemeIcon('folder', new vscode.ThemeColor('charts.blue'));
        }
        
        // Default folder icon
        return new vscode.ThemeIcon('folder');
    }
} 
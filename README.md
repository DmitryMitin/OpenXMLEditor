# OpenXML Editor

Advanced Visual Studio Code extension for viewing, editing, and analyzing OpenXML documents (Microsoft Office files: Word, Excel, PowerPoint) with temporary file extraction for full compatibility with other XML extensions.

## ✨ Key Features

### 🌳 **File Structure Tree**
- **Sidebar structure display**: View the complete internal structure of OpenXML files
- **File navigation**: Open any XML file from the archive with a single click
- **Size information**: See the size of each file in the archive
- **Hierarchical display**: Folders and files organized in an intuitive tree structure

### 💾 **Temporary File System**
- **Real file extraction**: XML files are extracted to temporary directories as real files
- **Full XML extension compatibility**: Works seamlessly with all VS Code XML extensions
- **Automatic synchronization**: Changes to temporary files are automatically synced back
- **Smart file watchers**: Node.js file watchers with debouncing for reliable change detection
- **Auto-cleanup**: Temporary files are automatically cleaned up when closing

### 🔄 **Advanced Change Tracking**
- **Multi-layer sync**: Node.js file watchers + VS Code save events
- **Smart debouncing**: Prevents excessive sync operations
- **Content comparison**: Only syncs when actual content changes
- **Backup creation**: Automatic backup creation when saving

## 🚀 Usage

### Opening an OpenXML File

1. **Right-click** on a `.docx`, `.xlsx`, or `.pptx` file in VS Code Explorer
2. Select **"Open in OpenXML Editor"** from the context menu
3. The file structure will appear in the **"OpenXML Structure"** sidebar
4. The status bar will show information about the loaded file

### Editing XML Files

1. **Click on any XML file** in the structure tree
2. The file will open in the editor with syntax highlighting
3. **Edit the content** like regular text
4. Changes are **automatically saved** to the original OpenXML file after 1 second

### Opening in Office Applications

#### **From File Explorer**
1. **Right-click** on any OpenXML file in VS Code Explorer
2. Select **"Open in Office Application"**
3. File opens in the appropriate Office program (Word/Excel/PowerPoint)

#### **From Structure Tree**
1. **Right-click** on the root OpenXML file node in the structure tree
2. Select **"Open in Office Application"**
3. Original file opens in the corresponding Office application

### Saving Changes

- **Auto-save**: Happens automatically 1 second after changes
- **Manual save**: Click the "💾" button in the structure panel header
- **Via Command Palette**: "OpenXML: Save Changes to OpenXML"

## 📋 Supported Formats

- **📄 Word**: `.docx`, `.dotx` (documents and templates)
- **📊 Excel**: `.xlsx`, `.xltx` (spreadsheets and templates)  
- **🎯 PowerPoint**: `.pptx`, `.potx` (presentations and templates)

## 🎯 Commands

All commands are available through Command Palette (`Ctrl+Shift+P`):

- `OpenXML: Open in OpenXML Editor` - Open OpenXML file in editor
- `OpenXML: Save Changes to OpenXML` - Save all changes
- `OpenXML: Open in Office Application` - Open file in system Office application
- `OpenXML: Show File Information` - Show detailed file information

## 🔧 Interface

### "OpenXML Structure" Sidebar
- 🗂️ **Hierarchical tree** of all files and folders
- 📏 **File sizes** for each element
- 🔄 **Refresh button** to reload structure
- 💾 **Save button** for manual saving

### Status Bar
- 📁 **File name** with ZIP archive icon
- ⚠️ **Number of modified files** (if there are unsaved changes)
- 💾 **Click to save** (if there are changes)

### Information Panels
- 📊 **Detailed file information**
- 📝 **List of modified files**
- ✅ **Save status**

## 🔗 Third-Party Extension Compatibility

**OpenXML Editor is designed to work seamlessly with other VS Code XML extensions:**

- ✅ **Language Mode**: Files are automatically set to `xml` language mode
- ✅ **Context Variables**: Proper context variables are set for extension detection
- ✅ **File Type Detection**: XML files are properly recognized by other extensions
- ✅ **IntelliSense**: XML autocompletion and validation work with supported extensions
- ✅ **Formatting Tools**: Third-party XML formatters can be used alongside built-in formatting

**Compatible with popular XML extensions:**
- XML Language Support
- XML Tools
- Auto Rename Tag  
- Bracket Pair Colorizer
- Path Intellisense (for XML paths)

## ⚙️ Technical Details

### Architecture
- **Temporary File System**: Extracts XML files to real temporary directories
- **Multi-layer Sync**: Node.js file watchers + VS Code save events
- **Tree Data Provider**: For displaying structure in sidebar with inherited TreeItem classes
- **File System Provider**: For URI handling and file operations
- **Debounced Auto-save**: Smart saving with change detection
- **Object-Oriented Design**: OpenXMLTreeItem inherits from vscode.TreeItem for better VS Code integration

### Sync Mechanisms
1. **Node.js File Watchers**: Real-time file change detection with debouncing
2. **VS Code Save Events**: Triggered when user saves documents

### Dependencies
- **yauzl**: Reading ZIP archives (OpenXML files are ZIP archives)
- **yazl**: Creating ZIP archives for saving changes
- **Node.js fs**: File system operations and watchers
- **VS Code API**: For editor integration

### Security
- ✅ **Automatic backups** before saving
- ✅ **Integrity checks** when writing
- ✅ **Recovery from backup** on errors

## 🚧 Development Installation

1. Clone the repository
```bash
git clone <repository-url>
cd openxmleditor
```

2. Install dependencies
```bash
npm install
```

3. Compile TypeScript
```bash
npm run compile
```

4. Run in development mode
- Press `F5` to launch Extension Development Host
- Or use "Run and Debug" in VS Code

## 🔨 Development

### Build
```bash
npm run compile
```

### Watch for changes
```bash
npm run watch
```

### Linting
```bash
npm run lint
```

### Testing
```bash
npm test
```

## 🐛 Known Limitations

- Large OpenXML files (>50MB) may work slowly
- Binary data (images) are displayed as Base64

## 📈 Development Plans

- 🔍 **Content search** within XML files
- 🎨 **Visual editor** for common elements
- 📦 **Export individual files** from archive
- 🔗 **Navigation between linked files**
- 🛠️ **XML schema validation**

## 🤝 Contributing

We welcome contributions to the project:
- 🐛 Bug reports
- 💡 Feature suggestions
- 🔧 Pull requests
- 📖 Documentation improvements

## 📄 License

This project is distributed as open source. Check the LICENSE file for details.

---

**🎉 Explore the internal structure of your Office documents with ease!** 📄✨

### 🌟 Why is this awesome?

- ⚡ **Instant access** to structure without extracting files
- 🎯 **Precision editing** of XML content
- 🔄 **Seamless workflow** - edit and save in one click
- 🛡️ **Security** - automatic backups and checks
- 🎨 **Modern UI** - intuitive interface

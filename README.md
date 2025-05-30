# OpenXML Editor

Advanced Visual Studio Code extension for viewing, editing, and analyzing OpenXML documents (Microsoft Office files: Word, Excel, PowerPoint) directly in the editor without extracting to disk.

## ✨ Key Features

### 🌳 **Virtual Structure Tree**
- **File structure display in sidebar**: View the complete internal structure of OpenXML files
- **File navigation**: Open any XML file from the archive with a single click
- **Size information**: See the size of each file in the archive
- **Hierarchical display**: Folders and files organized in an intuitive tree structure

### 💾 **Virtual File System**
- **Edit without extraction**: Modify XML files directly in memory
- **Auto-save**: Changes are automatically saved back to the original OpenXML file
- **Change tracking**: See which files have been modified
- **Backup creation**: Automatic backup creation when saving

### 🔄 **Automatic Packaging**
- **Transparent saving**: All changes are automatically packaged back into OpenXML
- **Change status**: Status bar shows the number of unsaved changes
- **Safety**: Create backup copies before saving

### 🎨 **XML Formatting**
- **Smart auto-formatting**: Intelligent XML formatting with proper indentation and structure
- **Advanced algorithm**: Handles complex XML with CDATA, comments, and nested elements
- **Syntax highlighting**: Full XML syntax support
- **Smart editing**: Auto-closing tags and brackets
- **Progress indication**: Shows progress for large files
- **Cursor preservation**: Maintains approximate cursor position after formatting
- **Third-party compatibility**: Works with other XML extensions and tools

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

### Formatting XML

1. Open any XML file in the editor
2. Use **Command Palette** (`Ctrl+Shift+P`) → **"OpenXML: Format XML Content"**
3. Or use keyboard shortcut: **`Ctrl+Alt+F`** (Windows/Linux) / **`Cmd+Alt+F`** (Mac)
4. Or right-click in editor → **"Format XML Content"**

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
- `OpenXML: Format XML Content` - Format current XML file with proper indentation
- `OpenXML: Open in Office Application` - Open file in system Office application
- `OpenXML: Show File Information` - Show detailed file information

## ⌨️ Keyboard Shortcuts

When editing OpenXML virtual files:

- **`Ctrl+Alt+F`** (Windows/Linux) / **`Cmd+Alt+F`** (Mac) - Format XML with proper indentation

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
- **Virtual File System Provider**: For working with files in memory
- **Tree Data Provider**: For displaying structure in sidebar
- **Custom URI Scheme**: `openxml://` for virtual files
- **Auto-save mechanism**: Delayed saving of changes

### Dependencies
- **yauzl**: Reading ZIP archives (OpenXML files are ZIP archives)
- **yazl**: Creating ZIP archives for saving changes
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
- Some specific OpenXML elements may require manual formatting

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

# Changelog

All notable changes to the OpenXML Editor extension will be documented in this file.

## [0.2.0] - 2024-12-19

### 🔧 Major Refactoring

#### ✨ New Architecture
- **Centralized Constants**: Created `src/constants.ts` with all configuration values
- **Logging System**: Implemented proper logging utility with different log levels
- **File Utilities**: Extracted file operations into reusable utility class
- **XML Formatter Service**: Moved XML formatting logic into dedicated service class
- **Better Error Handling**: Comprehensive error handling throughout the codebase
- **TreeItem Inheritance**: OpenXMLTreeItem now inherits from vscode.TreeItem for better integration and less code duplication

#### 🧹 Code Cleanup
- **Removed Code Duplication**: Eliminated duplicate logic across files, especially in tree item creation
- **Removed Dead Code**: Deleted unused methods and interfaces
- **Consistent Naming**: Applied consistent naming conventions
- **Type Safety**: Improved TypeScript usage with proper types and class inheritance
- **Magic Numbers**: Replaced all magic numbers with named constants
- **Simplified TreeDataProvider**: getTreeItem() method now simply returns the element directly

#### 📁 New File Structure
```
src/
├── constants.ts              # All configuration constants
├── utils/
│   ├── logger.ts            # Centralized logging system
│   └── fileUtils.ts         # File operation utilities
├── services/
│   └── xmlFormatter.ts      # XML formatting service
├── extension.ts             # Main extension (refactored)
├── openXmlFileSystem.ts     # File system provider (refactored)
└── openXmlTreeProvider.ts   # Tree provider
```

#### 🔄 Improvements
- **Better Logging**: Replaced console.log with structured logging
- **Configuration Management**: All settings centralized in constants
- **Service Architecture**: Separated concerns into dedicated services
- **Error Handling**: Comprehensive error handling with proper logging
- **Code Readability**: Improved code structure and documentation

#### 🗑️ Removed
- **Unused Methods**: Removed `loadFileFromOpenXML` and other dead code
- **Duplicate Functions**: Consolidated duplicate file type detection logic
- **Magic Values**: Replaced hardcoded values with constants
- **Legacy Code**: Removed compatibility methods that were no longer needed

#### 📊 Metrics
- **File Count**: Increased from 3 to 7 files (better organization)
- **Package Size**: Slightly increased to 86.87 KB (due to better structure)
- **Code Quality**: Significantly improved maintainability and readability

## [0.1.3] - 2024-12-19

### 🔧 Simplified Sync Mechanism
- **Removed Periodic Timer**: Eliminated background sync timer
- **Removed Force Sync Command**: Simplified user interface
- **Streamlined Architecture**: Focus on file watcher + VS Code save events

## [0.1.2] - 2024-12-19

### 🔧 Improved File Watching
- **Node.js File Watchers**: Replaced VS Code watchers with Node.js fs.watch
- **Debouncing**: Added smart debouncing to prevent excessive sync operations
- **Multiple Sync Mechanisms**: File watchers + periodic sync + save events
- **Better Error Handling**: Improved sync reliability

## [0.1.1] - 2024-12-19

### 🔧 Enhanced Sync System
- **Multi-layer Sync**: Added multiple synchronization mechanisms
- **Force Sync Command**: Manual sync trigger for reliability
- **Better Change Detection**: Content comparison for accurate sync

## [0.1.0] - 2024-12-19

### 🎉 Initial Release
- **Temporary File System**: Extract XML files to real temporary directories
- **Full XML Extension Compatibility**: Works with all VS Code XML extensions
- **File Structure Tree**: Sidebar view of OpenXML internal structure
- **XML Formatting**: Advanced XML formatting with proper indentation
- **Auto-save**: Automatic saving of changes back to OpenXML files
- **Office Integration**: Open files in Office applications

### 📋 Supported Formats
- Microsoft Word (.docx, .dotx)
- Microsoft Excel (.xlsx, .xltx)
- Microsoft PowerPoint (.pptx, .potx)
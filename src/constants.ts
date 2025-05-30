/**
 * Application constants and configuration
 */

export const EXTENSION_CONFIG = {
  name: 'openxmleditor',
  displayName: 'OpenXML Editor',
  scheme: 'openxml',
  tempDirPrefix: 'openxml-',
} as const;

export const FILE_EXTENSIONS = {
  word: ['.docx', '.dotx'],
  excel: ['.xlsx', '.xltx'],
  powerpoint: ['.pptx', '.potx'],
  xml: ['.xml', '.rels'],
  all: ['.docx', '.dotx', '.xlsx', '.xltx', '.pptx', '.potx'],
} as const;

export const SYNC_CONFIG = {
  debounceDelayMs: 500,
  autoSaveDelayMs: 2000,
  fileAgeThresholdMs: 10000,
  largeFileThreshold: 50000,
} as const;

export const UI_CONFIG = {
  progressThreshold: 50000,
  defaultIndent: '  ',
} as const;

export const COMMANDS = {
  openFile: 'openxmleditor.openFile',
  openInSystem: 'openxmleditor.openInSystem',
  saveChanges: 'openxmleditor.saveChanges',
  refreshTree: 'openxmleditor.refreshTree',
  openXmlFile: 'openxmleditor.openXmlFile',
  showFileInfo: 'openxmleditor.showFileInfo',
  closeOpenXMLFile: 'openxmleditor.closeOpenXMLFile',
} as const;

export const TREE_ITEMS = {
  xmlFile: 'xmlFile',
  openxmlFile: 'openxmlFile',
} as const;

export const CONTEXT_KEYS = {
  openxmlFileOpen: 'openxmlFileOpen',
} as const;

export const MIME_TYPES = {
  xml: 'application/xml',
  textXml: 'text/xml',
  rels: 'application/vnd.openxmlformats-package.relationships+xml',
} as const; 
/**
 * File utilities and validation functions
 */

import * as path from 'path';
import { FILE_EXTENSIONS } from '../constants';

export class FileUtils {
  /**
   * Check if a file is an OpenXML document
   */
  static isOpenXMLFile(filePath: string): boolean {
    const ext = path.extname(filePath).toLowerCase();
    return FILE_EXTENSIONS.all.includes(ext as any);
  }

  /**
   * Check if a file is an XML file
   */
  static isXMLFile(filePath: string): boolean {
    const ext = path.extname(filePath).toLowerCase();
    return FILE_EXTENSIONS.xml.includes(ext as any);
  }

  /**
   * Get file type description based on extension
   */
  static getFileTypeDescription(extension: string): string {
    const types: Record<string, string> = {
      '.docx': 'Microsoft Word Document',
      '.xlsx': 'Microsoft Excel Spreadsheet',
      '.pptx': 'Microsoft PowerPoint Presentation',
      '.dotx': 'Microsoft Word Template',
      '.xltx': 'Microsoft Excel Template',
      '.potx': 'Microsoft PowerPoint Template',
    };
    return types[extension] || 'OpenXML Document';
  }

  /**
   * Format file size in human-readable format
   */
  static formatFileSize(bytes: number): string {
    const sizes = ['B', 'KB', 'MB', 'GB'];
    if (bytes === 0) return '0 B';
    
    const i = Math.floor(Math.log(bytes) / Math.log(1024));
    return Math.round(bytes / Math.pow(1024, i) * 100) / 100 + ' ' + sizes[i];
  }

  /**
   * Normalize file path for cross-platform compatibility
   */
  static normalizePath(filePath: string): string {
    return filePath.replace(/[/\\]/g, path.sep);
  }

  /**
   * Create a short hash for file identification
   */
  static createShortHash(input: string, length: number = 6): string {
    return Buffer.from(input, 'utf8')
      .toString('base64')
      .replace(/[/+=]/g, (match) => {
        switch (match) {
          case '/': return '_';
          case '+': return '-';
          case '=': return '';
          default: return match;
        }
      })
      .substring(0, length);
  }

  /**
   * Check if content appears to be XML
   */
  static isXMLContent(content: string): boolean {
    const trimmed = content.trim();
    return trimmed.startsWith('<?xml') || 
           trimmed.startsWith('<') || 
           trimmed.includes('<');
  }

  /**
   * Extract presentation name from file path
   */
  static extractPresentationName(filePath: string): string {
    return path.basename(filePath, path.extname(filePath));
  }
} 
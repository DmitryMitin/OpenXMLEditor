/**
 * XML formatting service with advanced formatting capabilities
 */

import { UI_CONFIG } from '../constants';
import { logger } from '../utils/logger';

interface FormatToken {
  type: 'tag' | 'text' | 'comment' | 'cdata';
  content: string;
}

export class XMLFormatter {
  private readonly indent: string;

  constructor(indent: string = UI_CONFIG.defaultIndent) {
    this.indent = indent;
  }

  /**
   * Format XML content with proper indentation
   */
  formatXML(xml: string): string {
    try {
      logger.debug('Starting XML formatting', { length: xml.length });
      
      // Clean up the input
      let formatted = xml.trim();
      
      // Remove all existing whitespace between tags
      formatted = formatted.replace(/>\s*</g, '><');
      
      // Parse into tokens
      const tokens = this.parseTokens(formatted);
      logger.debug('Parsed tokens', { count: tokens.length });
      
      // Format tokens
      const result = this.formatTokens(tokens);
      
      logger.success('XML formatting completed');
      return result;
      
    } catch (error) {
      logger.error('XML formatting failed, using fallback', error);
      return this.fallbackFormat(xml);
    }
  }

  /**
   * Parse XML content into tokens
   */
  private parseTokens(xml: string): FormatToken[] {
    const tokens: FormatToken[] = [];
    let i = 0;

    while (i < xml.length) {
      const char = xml[i];
      const nextChars4 = xml.substring(i, i + 4);
      const nextChars9 = xml.substring(i, i + 9);

      // Handle CDATA sections
      if (nextChars9 === '<![CDATA[') {
        const cdataEnd = xml.indexOf(']]>', i);
        if (cdataEnd !== -1) {
          tokens.push({
            type: 'cdata',
            content: xml.substring(i, cdataEnd + 3)
          });
          i = cdataEnd + 3;
          continue;
        }
      }

      // Handle comments
      if (nextChars4 === '<!--') {
        const commentEnd = xml.indexOf('-->', i);
        if (commentEnd !== -1) {
          tokens.push({
            type: 'comment',
            content: xml.substring(i, commentEnd + 3)
          });
          i = commentEnd + 3;
          continue;
        }
      }

      // Handle tags
      if (char === '<') {
        const tagEnd = xml.indexOf('>', i);
        if (tagEnd !== -1) {
          tokens.push({
            type: 'tag',
            content: xml.substring(i, tagEnd + 1)
          });
          i = tagEnd + 1;
          continue;
        }
      }

      // Handle text content
      let textContent = '';
      while (i < xml.length && xml[i] !== '<') {
        textContent += xml[i];
        i++;
      }

      if (textContent.trim()) {
        tokens.push({
          type: 'text',
          content: textContent.trim()
        });
      }
    }

    return tokens;
  }

  /**
   * Format parsed tokens with proper indentation
   */
  private formatTokens(tokens: FormatToken[]): string {
    const result: string[] = [];
    let depth = 0;

    for (let tokenIndex = 0; tokenIndex < tokens.length; tokenIndex++) {
      const token = tokens[tokenIndex];

      if (token.type === 'comment' || token.type === 'cdata') {
        result.push(this.indent.repeat(depth) + token.content);
      } else if (token.type === 'tag') {
        const tag = token.content;

        if (tag.startsWith('<?') || tag.startsWith('<!')) {
          // XML declaration or DOCTYPE - no indentation change
          result.push(tag);
        } else if (tag.startsWith('</')) {
          // Closing tag - decrease depth first, then add
          depth = Math.max(0, depth - 1);
          result.push(this.indent.repeat(depth) + tag);
        } else if (tag.endsWith('/>')) {
          // Self-closing tag - no depth change
          result.push(this.indent.repeat(depth) + tag);
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

            result.push(this.indent.repeat(depth) + tag + textToken.content + closingTag.content);

            // Skip the next two tokens since we processed them
            tokenIndex += 2;
          } else {
            // Complex element - add tag and increase depth
            result.push(this.indent.repeat(depth) + tag);
            depth++;
          }
        }
      } else if (token.type === 'text') {
        // Standalone text (not part of a simple element)
        result.push(this.indent.repeat(depth) + token.content);
      }
    }

    // Join and clean up
    return result
      .filter(line => line.trim()) // Remove empty lines
      .map(line => line.trimEnd()) // Remove trailing spaces
      .join('\n');
  }

  /**
   * Fallback formatting for when advanced parsing fails
   */
  private fallbackFormat(xml: string): string {
    logger.warn('Using fallback XML formatting');
    return xml
      .replace(/>\s*</g, '>\n<')
      .split('\n')
      .map(line => line.trim())
      .filter(line => line)
      .join('\n');
  }
} 
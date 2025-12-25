/**
 * Document Redaction Module
 * 
 * This module provides robust regex patterns and logic for detecting
 * sensitive information in documents including:
 * - Email addresses
 * - Phone numbers (various formats)
 * - Social Security Numbers (full and partial - last 4 digits)
 */

// Redaction marker
const REDACTION_MARKER = '[REDACTED]';

/**
 * Regex patterns for sensitive information detection
 * These patterns are designed to be comprehensive and handle various formats
 */
const PATTERNS = {
  /**
   * Email pattern - matches standard email formats
   * Examples: john.doe@example.com, user+tag@domain.co.uk
   */
  email: /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b/g,

  /**
   * Phone number patterns - multiple formats supported
   * Formats covered:
   * - (123) 456-7890
   * - (123)456-7890
   * - 123-456-7890
   * - 123.456.7890
   * - 123 456 7890
   * - +1 123 456 7890
   * - +1-123-456-7890
   * - +1(123)456-7890
   * - +11234567890 (with country code)
   * - 1234567890
   * Note: Captures optional leading + for international numbers
   */
  phone: /\+?1?[-.\s]?\(?[0-9]{3}\)?[-.\s]?[0-9]{3}[-.\s]?[0-9]{4}\b/g,

  /**
   * Full Social Security Number patterns
   * Formats covered:
   * - 123-45-6789
   * - 123 45 6789
   * - 123.45.6789
   * - 123456789 (9 consecutive digits)
   */
  ssnFull: /\b\d{3}[-.\s]?\d{2}[-.\s]?\d{4}\b/g,

  /**
   * Masked SSN formats
   * Formats covered:
   * - xxxx1234, XXXX1234, ****1234
   * - xxx-xx-1234, XXX-XX-1234
   * - *xx-xx-1234
   */
  ssnMasked: /\b[xX*]{3,4}[-.\s]?[xX*]{0,2}[-.\s]?\d{4}\b/g,

  /**
   * Partial SSN - Last 4 digits mentioned in context
   * Matches patterns like:
   * - "last four digits ... are 1234"
   * - "last 4 digits ... 5678"
   * - "SSN ending in 9012"
   * - "ends in 3456"
   */
  ssnPartialContext: /(?:last\s+(?:four|4)\s+digits?(?:\s+\w+){0,10}\s+(?:are|is|:)?\s*)(\d{4})\b/gi,
  
  /**
   * SSN ending pattern
   * Matches: "ending in 1234", "ends in 5678"
   */
  ssnEnding: /(?:(?:ssn|social\s*security(?:\s*number)?)\s+)?(?:ending|ends)\s+in\s+(\d{4})\b/gi,
};

/**
 * Result of a redaction operation
 */
export interface RedactionResult {
  success: boolean;
  emailsRedacted: number;
  phonesRedacted: number;
  ssnsRedacted: number;
  totalRedacted: number;
  trackingEnabled: boolean;
  headerAdded: boolean;
  error?: string;
}

/**
 * Validates if a potential SSN is actually an SSN
 * SSNs have specific rules: 
 * - Area number (first 3 digits) cannot be 000, 666, or 900-999
 * - Group number (middle 2 digits) cannot be 00
 * - Serial number (last 4 digits) cannot be 0000
 */
function isValidSSN(ssn: string): boolean {
  const digits = ssn.replace(/\D/g, '');
  if (digits.length !== 9) return false;
  
  const area = parseInt(digits.substring(0, 3), 10);
  const group = parseInt(digits.substring(3, 5), 10);
  const serial = parseInt(digits.substring(5, 9), 10);
  
  // Invalid area numbers
  if (area === 0 || area === 666 || area >= 900) return false;
  // Invalid group number
  if (group === 0) return false;
  // Invalid serial number
  if (serial === 0) return false;
  
  return true;
}

/**
 * Check if Word API version supports Track Changes (1.5+)
 */
function isTrackChangesSupported(): boolean {
  return Office.context.requirements.isSetSupported('WordApi', '1.5');
}

/**
 * Extract partial SSN (last 4 digits) from contextual mentions
 */
function findPartialSSNs(text: string): string[] {
  const partialSSNs: string[] = [];
  
  // Pattern 1: "last four digits ... are XXXX"
  const contextPattern = /last\s+(?:four|4)\s+digits?(?:\s+\w+){0,10}\s+(?:are|is|:)?\s*(\d{4})\b/gi;
  let match;
  while ((match = contextPattern.exec(text)) !== null) {
    partialSSNs.push(match[1]);
  }
  
  // Pattern 2: "ending in XXXX" or "ends in XXXX"
  const endingPattern = /(?:ending|ends)\s+in\s+(\d{4})\b/gi;
  while ((match = endingPattern.exec(text)) !== null) {
    partialSSNs.push(match[1]);
  }
  
  // Pattern 3: "SSN: XXX-XX-XXXX" or just the last 4 after SSN context
  const ssnContextPattern = /(?:ssn|social\s*security(?:\s*number)?)[:\s]+(?:\d{3}[-.\s]?\d{2}[-.\s]?)?(\d{4})\b/gi;
  while ((match = ssnContextPattern.exec(text)) !== null) {
    partialSSNs.push(match[1]);
  }
  
  return [...new Set(partialSSNs)];
}

/**
 * Main redaction function that processes the entire document
 */
export async function redactDocument(): Promise<RedactionResult> {
  const result: RedactionResult = {
    success: false,
    emailsRedacted: 0,
    phonesRedacted: 0,
    ssnsRedacted: 0,
    totalRedacted: 0,
    trackingEnabled: false,
    headerAdded: false,
  };

  try {
    await Word.run(async (context) => {
      const document = context.document;
      const trackChangesSupported = isTrackChangesSupported();
      
      // First, DISABLE track changes so redactions appear clean (no strikethrough)
      // We'll enable it AFTER making changes
      if (trackChangesSupported) {
        document.changeTrackingMode = Word.ChangeTrackingMode.off;
        await context.sync();
      }

      // Get the document body
      const body = context.document.body;
      body.load('text');
      await context.sync();

      const originalText = body.text;

      // Count occurrences before redaction
      result.emailsRedacted = countMatches(originalText, PATTERNS.email);
      
      // For phones, we need to be more careful to avoid false positives
      const phoneMatches = originalText.match(PATTERNS.phone) || [];
      result.phonesRedacted = phoneMatches.filter(match => {
        // Filter out matches that are too short (less than 10 digits with area code)
        const digits = match.replace(/\D/g, '');
        return digits.length >= 10 && digits.length <= 11;
      }).length;
      
      // For full SSNs, validate each match
      const ssnFullMatches = originalText.match(PATTERNS.ssnFull) || [];
      const validFullSSNs = ssnFullMatches.filter(isValidSSN).length;
      
      // For masked SSNs (xxxx1234 format)
      const maskedSSNs = originalText.match(PATTERNS.ssnMasked) || [];
      
      // For partial SSNs (last 4 digits in context)
      const partialSSNs = findPartialSSNs(originalText);
      result.ssnsRedacted = validFullSSNs + maskedSSNs.length + partialSSNs.length;

      // Perform redactions using Word's search and replace

      // Redact emails
      await redactEmails(context, body);
      
      // Redact phone numbers with validation
      await redactPhoneNumbers(context, body);
      
      // Redact full SSNs with validation
      await redactFullSSNs(context, body);
      
      // Redact masked SSNs (xxxx1234 format)
      await redactMaskedSSNs(context, body);
      
      // Redact partial SSNs (last 4 digits in context)
      await redactPartialSSNs(context, body, partialSSNs);

      // Add confidential header
      const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
      header.insertParagraph('CONFIDENTIAL DOCUMENT', Word.InsertLocation.start);
      
      // Style the header
      const headerParagraph = header.paragraphs.getFirst();
      headerParagraph.font.bold = true;
      headerParagraph.font.size = 14;
      headerParagraph.font.color = '#C00000';
      headerParagraph.alignment = Word.Alignment.centered;
      
      result.headerAdded = true;

      await context.sync();

      // NOW enable Track Changes so future modifications will be tracked
      // The redactions we just made will appear clean (no strikethrough)
      if (trackChangesSupported) {
        document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        result.trackingEnabled = true;
        await context.sync();
      }

      result.totalRedacted = result.emailsRedacted + result.phonesRedacted + result.ssnsRedacted;
      result.success = true;
    });
  } catch (error) {
    result.error = error instanceof Error ? error.message : 'An unknown error occurred';
    result.success = false;
  }

  return result;
}

/**
 * Counts matches for a given pattern in text
 */
function countMatches(text: string, pattern: RegExp): number {
  const matches = text.match(pattern);
  return matches ? matches.length : 0;
}

/**
 * Redacts emails using Word's search
 */
async function redactEmails(
  context: Word.RequestContext,
  body: Word.Body
): Promise<void> {
  body.load('text');
  await context.sync();

  const text = body.text;
  const matches = text.match(PATTERNS.email) || [];
  
  // Use a Set to avoid duplicate searches
  const uniqueMatches = [...new Set(matches)];

  for (const match of uniqueMatches) {
    const searchResults = body.search(match, {
      matchCase: false,
      matchWholeWord: false,
    });
    searchResults.load('items');
    await context.sync();

    for (const item of searchResults.items) {
      item.insertText(REDACTION_MARKER, Word.InsertLocation.replace);
    }
    await context.sync();
  }
}

/**
 * Redacts phone numbers with additional validation
 */
async function redactPhoneNumbers(
  context: Word.RequestContext,
  body: Word.Body
): Promise<void> {
  body.load('text');
  await context.sync();

  const text = body.text;
  const matches = text.match(PATTERNS.phone) || [];
  
  // Filter and deduplicate - require at least 10 digits (with area code)
  const validMatches = [...new Set(
    matches.filter(match => {
      const digits = match.replace(/\D/g, '');
      return digits.length >= 10 && digits.length <= 11;
    })
  )];

  for (const match of validMatches) {
    const searchResults = body.search(match, {
      matchCase: true,
      matchWholeWord: false,
    });
    searchResults.load('items');
    await context.sync();

    for (const item of searchResults.items) {
      item.insertText(REDACTION_MARKER, Word.InsertLocation.replace);
    }
    await context.sync();
  }
}

/**
 * Redacts full SSNs (9 digits) with validation
 */
async function redactFullSSNs(
  context: Word.RequestContext,
  body: Word.Body
): Promise<void> {
  body.load('text');
  await context.sync();

  const text = body.text;
  const matches = text.match(PATTERNS.ssnFull) || [];
  
  // Filter valid SSNs and deduplicate
  const validSSNs = [...new Set(matches.filter(isValidSSN))];

  for (const match of validSSNs) {
    const searchResults = body.search(match, {
      matchCase: true,
      matchWholeWord: false,
    });
    searchResults.load('items');
    await context.sync();

    for (const item of searchResults.items) {
      item.insertText(REDACTION_MARKER, Word.InsertLocation.replace);
    }
    await context.sync();
  }
}

/**
 * Redacts masked SSNs (xxxx1234 format)
 */
async function redactMaskedSSNs(
  context: Word.RequestContext,
  body: Word.Body
): Promise<void> {
  body.load('text');
  await context.sync();

  const text = body.text;
  const matches = text.match(PATTERNS.ssnMasked) || [];
  
  // Deduplicate
  const uniqueMatches = [...new Set(matches)];

  for (const match of uniqueMatches) {
    const searchResults = body.search(match, {
      matchCase: false,
      matchWholeWord: false,
    });
    searchResults.load('items');
    await context.sync();

    for (const item of searchResults.items) {
      item.insertText(REDACTION_MARKER, Word.InsertLocation.replace);
    }
    await context.sync();
  }
}

/**
 * Redacts partial SSNs (last 4 digits mentioned in context)
 */
async function redactPartialSSNs(
  context: Word.RequestContext,
  body: Word.Body,
  partialSSNs: string[]
): Promise<void> {
  // For each partial SSN found in context, search and redact just the 4 digits
  for (const digits of partialSSNs) {
    // Search for the exact 4-digit sequence
    const searchResults = body.search(digits, {
      matchCase: true,
      matchWholeWord: true,
    });
    searchResults.load('items');
    await context.sync();

    for (const item of searchResults.items) {
      item.insertText(REDACTION_MARKER, Word.InsertLocation.replace);
    }
    await context.sync();
  }
}

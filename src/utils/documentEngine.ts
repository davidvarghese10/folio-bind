import { PDFDocument, StandardFonts, rgb, degrees } from 'pdf-lib';
import JSZip from 'jszip';
import mammoth from 'mammoth';
import {
  QueuedDocument,
  MergeConfiguration,
  MergedArtifact,
  StructuredBlock,
  SlideContent,
  SourceFileType,
  RejectedFileError,
  MergeProgressState,
} from '../types/document';

const ALLOWED_PDF_MIMES = new Set([
  'application/pdf',
  'application/x-pdf',
  'application/acrobat',
  'applications/vnd.pdf',
  'text/pdf',
  'text/x-pdf',
]);

const ALLOWED_DOCX_MIMES = new Set([
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/zip',
  'application/x-zip-compressed',
  'application/octet-stream',
]);

const ALLOWED_PPTX_MIMES = new Set([
  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  'application/mspowerpoint',
  'application/powerpoint',
  'application/vnd.ms-powerpoint',
  'application/x-mspowerpoint',
  'application/zip',
  'application/x-zip-compressed',
  'application/octet-stream',
]);

/**
 * Validates uploaded files to strictly accept only valid PDF (.pdf), Word (.docx), and PowerPoint (.pptx) files.
 * Returns both accepted files and detailed error descriptions for any unsupported files.
 */
export function validateUploadedFiles(fileList: FileList | File[]): {
  validFiles: File[];
  rejectedFiles: RejectedFileError[];
} {
  const validFiles: File[] = [];
  const rejectedFiles: RejectedFileError[] = [];
  const files = Array.from(fileList);

  for (const file of files) {
    const rawName = file.name || 'Unnamed_File';
    const lastDotIndex = rawName.lastIndexOf('.');
    const ext = lastDotIndex !== -1 ? rawName.slice(lastDotIndex).toLowerCase() : '';
    const mime = (file.type || '').toLowerCase().trim();
    const detectedLabel = ext ? ext.toUpperCase() : mime || 'Unknown Format';

    if (file.size === 0) {
      rejectedFiles.push({
        id: `err_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
        fileName: rawName,
        detectedType: detectedLabel,
        sizeBytes: 0,
        reason: `File "${rawName}" is empty (0 bytes). Please upload a valid non-empty .pdf, .docx, or .pptx document.`,
      });
      continue;
    }

    if (ext === '.doc') {
      rejectedFiles.push({
        id: `err_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
        fileName: rawName,
        detectedType: '.DOC (Legacy Word 97-2003)',
        sizeBytes: file.size,
        reason: `Legacy Word ".doc" format is not supported for "${rawName}". Please save or export the file as ".docx" or ".pdf" and try again.`,
      });
      continue;
    }

    if (ext === '.ppt') {
      rejectedFiles.push({
        id: `err_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
        fileName: rawName,
        detectedType: '.PPT (Legacy PowerPoint 97-2003)',
        sizeBytes: file.size,
        reason: `Legacy PowerPoint ".ppt" format is not supported for "${rawName}". Please save or export the presentation as ".pptx" or ".pdf" and try again.`,
      });
      continue;
    }

    const isPdfExt = ext === '.pdf';
    const isDocxExt = ext === '.docx';
    const isPptxExt = ext === '.pptx';

    if (!isPdfExt && !isDocxExt && !isPptxExt) {
      rejectedFiles.push({
        id: `err_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
        fileName: rawName,
        detectedType: detectedLabel,
        sizeBytes: file.size,
        reason: `Unsupported file type (${ext || mime || 'unknown'}) for "${rawName}". Only PDF (.pdf), Word (.docx), and PowerPoint (.pptx) files are allowed.`,
      });
      continue;
    }

    if (
      isPdfExt &&
      mime &&
      !ALLOWED_PDF_MIMES.has(mime) &&
      !mime.includes('pdf') &&
      mime !== 'application/octet-stream'
    ) {
      rejectedFiles.push({
        id: `err_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
        fileName: rawName,
        detectedType: `${ext} (${mime})`,
        sizeBytes: file.size,
        reason: `Invalid MIME type "${mime}" detected for "${rawName}". Expected a valid Adobe PDF (.pdf) document.`,
      });
      continue;
    }

    if (
      isDocxExt &&
      mime &&
      !ALLOWED_DOCX_MIMES.has(mime) &&
      !mime.includes('word') &&
      !mime.includes('officedocument')
    ) {
      rejectedFiles.push({
        id: `err_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
        fileName: rawName,
        detectedType: `${ext} (${mime})`,
        sizeBytes: file.size,
        reason: `Invalid MIME type "${mime}" detected for "${rawName}". Expected a valid Microsoft Word (.docx) document.`,
      });
      continue;
    }

    if (
      isPptxExt &&
      mime &&
      !ALLOWED_PPTX_MIMES.has(mime) &&
      !mime.includes('presentation') &&
      !mime.includes('powerpoint') &&
      !mime.includes('officedocument')
    ) {
      rejectedFiles.push({
        id: `err_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
        fileName: rawName,
        detectedType: `${ext} (${mime})`,
        sizeBytes: file.size,
        reason: `Invalid MIME type "${mime}" detected for "${rawName}". Expected a valid Microsoft PowerPoint (.pptx) presentation.`,
      });
      continue;
    }

    validFiles.push(file);
  }

  return { validFiles, rejectedFiles };
}

/**
 * Reads a File into an ArrayBuffer while reporting real-time byte progress (0..1)
 */
function readFileBufferWithProgress(
  file: File,
  onProgress?: (fraction: number) => void
): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onprogress = (event) => {
      if (event.lengthComputable && onProgress) {
        onProgress(Math.min(1, event.loaded / Math.max(1, event.total)));
      }
    };
    reader.onload = () => {
      if (onProgress) onProgress(1);
      if (reader.result instanceof ArrayBuffer) {
        resolve(reader.result);
      } else {
        reject(new Error('Failed to read file binary buffer.'));
      }
    };
    reader.onerror = () => {
      reject(new Error(`Unable to read file "${file.name}".`));
    };
    reader.readAsArrayBuffer(file);
  });
}

function yieldToUi(ms = 22): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Ensures text is safe for pdf-lib StandardFonts (WinAnsiEncoding)
 */
export function sanitizeForWinAnsi(input: string): string {
  if (!input) return '';
  return input
    .replace(/[\u2018\u2019\u201A\u201B]/g, "'")
    .replace(/[\u201C\u201D\u201E\u201F]/g, '"')
    .replace(/[\u2013\u2014\u2015]/g, ' - ')
    .replace(/\u2026/g, '...')
    .replace(/[\u2022\u2023\u25E6\u2043\u2219]/g, '-')
    .replace(/\u00A0/g, ' ')
    .replace(/[^\x20-\x7E\xA1-\xFF\n\r\t]/g, '');
}

/**
 * Escapes XML special characters and strips illegal XML 1.0 control chars for OpenXML (.docx / .pptx) generation
 */
function escapeXml(unsafe: string): string {
  return (unsafe || '')
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Formats byte size into human-readable string with tabular precision
 */
export function formatBytes(bytes: number): string {
  if (bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  const val = bytes / Math.pow(k, i);
  return `${val.toFixed(i === 0 ? 0 : 1)} ${sizes[i]}`;
}

/**
 * Parses a user page/slide range string (e.g., "All", "1-3, 5") into 0-based indices
 */
export function parsePageRange(rangeStr: string, totalPages: number): number[] {
  const cleaned = rangeStr.trim().toLowerCase();
  if (!cleaned || cleaned === 'all' || cleaned === '*') {
    return Array.from({ length: totalPages }, (_, i) => i);
  }

  const indices = new Set<number>();
  const parts = cleaned.split(',').map((p) => p.trim()).filter(Boolean);

  for (const part of parts) {
    if (part.includes('-')) {
      const [startRaw, endRaw] = part.split('-').map((s) => parseInt(s.trim(), 10));
      if (!isNaN(startRaw) && !isNaN(endRaw)) {
        const start = Math.max(1, Math.min(startRaw, endRaw));
        const end = Math.min(totalPages, Math.max(startRaw, endRaw));
        for (let p = start; p <= end; p++) {
          indices.add(p - 1);
        }
      }
    } else {
      const pageNum = parseInt(part, 10);
      if (!isNaN(pageNum) && pageNum >= 1 && pageNum <= totalPages) {
        indices.add(pageNum - 1);
      }
    }
  }

  const result = Array.from(indices).sort((a, b) => a - b);
  return result.length > 0 ? result : Array.from({ length: totalPages }, (_, i) => i);
}

/**
 * Parses Mammoth HTML into structured blocks for previewing and cross-format rendering
 */
function parseHtmlToBlocks(html: string, fallbackText: string): StructuredBlock[] {
  const blocks: StructuredBlock[] = [];
  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    const elements = doc.body.querySelectorAll('h1, h2, h3, h4, p, li, tr');

    elements.forEach((el) => {
      const tag = el.tagName.toLowerCase();
      if ((tag === 'p' && el.closest('li')) || (tag === 'p' && el.closest('tr'))) {
        return;
      }
      if (tag === 'tr') {
        const cellElements = Array.from(el.querySelectorAll(':scope > th, :scope > td'));
        const effectiveCells =
          cellElements.length > 0
            ? cellElements
            : Array.from(el.querySelectorAll('th, td'));
        const cells = effectiveCells.map((c) =>
          (c.textContent || '').replace(/\s+/g, ' ').trim()
        );
        if (cells.some((c) => c.length > 0)) {
          const isHeaderRow =
            el.querySelectorAll('th').length > 0 ||
            Boolean(el.closest('thead')) ||
            (blocks.length > 0 && blocks[blocks.length - 1].type !== 'table-row');
          blocks.push({
            type: 'table-row',
            text: cells.join(' | '),
            cells,
            isHeaderRow,
          });
        }
        return;
      }

      const text = (el.textContent || '').replace(/\s+/g, ' ').trim();
      if (!text) return;

      if (tag === 'h1') blocks.push({ type: 'h1', text });
      else if (tag === 'h2') blocks.push({ type: 'h2', text });
      else if (tag === 'h3' || tag === 'h4') blocks.push({ type: 'h3', text });
      else if (tag === 'li') blocks.push({ type: 'li', text });
      else {
        blocks.push({ type: 'p', text });
      }
    });
  } catch {
    // Fallback to raw text line splitting
  }

  if (blocks.length === 0 && fallbackText.trim()) {
    fallbackText
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter(Boolean)
      .forEach((line, idx) => {
        blocks.push({
          type: idx === 0 ? 'h2' : 'p',
          text: line,
        });
      });
  }

  return blocks;
}

/**
 * Extracts readable text literals from raw PDF bytes when available
 */
function extractBasicTextFromPdfBuffer(buffer: ArrayBuffer): StructuredBlock[] {
  const blocks: StructuredBlock[] = [];
  try {
    const decoder = new TextDecoder('latin1');
    const raw = decoder.decode(new Uint8Array(buffer.slice(0, Math.min(buffer.byteLength, 250000))));
    const matches = raw.match(/\(([^()\\]{4,120})\)\s*Tj/g);
    if (matches && matches.length > 0) {
      const lines = matches
        .slice(0, 40)
        .map((m) => m.replace(/^\(|\)\s*Tj$/g, '').trim())
        .filter((s) => /[a-zA-Z]{2,}/.test(s));
      if (lines.length > 0) {
        lines.forEach((line, i) => {
          blocks.push({ type: i === 0 ? 'h2' : 'p', text: line });
        });
      }
    }
  } catch {
    // Ignore decoding issues
  }

  return blocks;
}

/**
 * Extracts native Word headers (word/header*.xml) and footers (word/footer*.xml) from a .docx zip archive
 * so they can be merged as true OpenXML <w:hdr> and <w:ftr> parts rather than plain body text.
 */
async function extractDocxHeadersAndFooters(zip: JSZip): Promise<{
  headers: string[];
  footers: string[];
}> {
  const headerEntries: Array<{ num: number; path: string }> = [];
  const footerEntries: Array<{ num: number; path: string }> = [];

  zip.forEach((relativePath) => {
    const hMatch = relativePath.match(/^word\/header(\d+)\.xml$/i);
    if (hMatch) {
      headerEntries.push({ num: parseInt(hMatch[1], 10), path: relativePath });
    }
    const fMatch = relativePath.match(/^word\/footer(\d+)\.xml$/i);
    if (fMatch) {
      footerEntries.push({ num: parseInt(fMatch[1], 10), path: relativePath });
    }
  });

  headerEntries.sort((a, b) => a.num - b.num);
  footerEntries.sort((a, b) => a.num - b.num);

  const parser = new DOMParser();
  const extractLinesFromXmlPart = async (path: string): Promise<string[]> => {
    const fileObj = zip.file(path);
    if (!fileObj) return [];
    const xmlText = await fileObj.async('text');
    const xmlDoc = parser.parseFromString(xmlText, 'application/xml');
    const pNodes = Array.from(xmlDoc.getElementsByTagName('w:p'));
    const lines: string[] = [];

    for (const pNode of pNodes) {
      const tNodes = Array.from(pNode.getElementsByTagName('w:t'));
      const text = tNodes
        .map((t) => t.textContent || '')
        .join('')
        .replace(/\s+/g, ' ')
        .trim();
      if (text) {
        lines.push(text);
      }
    }
    return lines;
  };

  const headersSet = new Set<string>();
  for (const hEntry of headerEntries) {
    const lines = await extractLinesFromXmlPart(hEntry.path);
    lines.forEach((l) => headersSet.add(l));
  }

  const footersSet = new Set<string>();
  for (const fEntry of footerEntries) {
    const lines = await extractLinesFromXmlPart(fEntry.path);
    lines.forEach((l) => footersSet.add(l));
  }

  return {
    headers: Array.from(headersSet),
    footers: Array.from(footersSet),
  };
}

/**
 * Extracts slide content, background colors, and table structures from an OpenXML PowerPoint (.pptx) archive
 */
async function extractSlidesFromPptxBuffer(
  arrayBuffer: ArrayBuffer,
  onProgress?: (detail: string) => void
): Promise<{
  slides: SlideContent[];
  blocks: StructuredBlock[];
  rawText: string;
  dimensionsLabel: string;
}> {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const slideEntries: Array<{ num: number; path: string }> = [];

  zip.forEach((relativePath) => {
    const match = relativePath.match(/^ppt\/slides\/slide(\d+)\.xml$/i);
    if (match) {
      slideEntries.push({
        num: parseInt(match[1], 10),
        path: relativePath,
      });
    }
  });

  slideEntries.sort((a, b) => a.num - b.num);

  let dimensionsLabel = '16:9 Widescreen Slides';
  try {
    const presFile = zip.file('ppt/presentation.xml');
    if (presFile) {
      const presXml = await presFile.async('text');
      const cxMatch = presXml.match(/<p:sldSz[^>]*\bcx="(\d+)"/);
      const cyMatch = presXml.match(/<p:sldSz[^>]*\bcy="(\d+)"/);
      if (cxMatch && cyMatch) {
        const cx = parseInt(cxMatch[1], 10);
        const cy = parseInt(cyMatch[1], 10);
        if (cx > 0 && cy > 0) {
          const ratio = cx / cy;
          if (ratio > 1.68) dimensionsLabel = '16:9 Widescreen Slides';
          else if (ratio > 1.5) dimensionsLabel = '16:10 Presentation Slides';
          else dimensionsLabel = '4:3 Standard Slides';
        }
      }
    }
  } catch {
    // Ignore presentation dimension parsing error
  }

  // Parse theme colors from ppt/theme/theme1.xml so schemeClr backgrounds resolve accurately
  const schemeColors: Record<string, string> = {
    dk1: '000000',
    lt1: 'FFFFFF',
    dk2: '1F497D',
    lt2: 'EEECE1',
    accent1: '4F81BD',
    accent2: 'C0504D',
    accent3: '9BBB59',
    accent4: '8064A2',
    accent5: '4BACC6',
    accent6: 'F79646',
    bg1: 'FFFFFF',
    bg2: 'EEECE1',
    tx1: '000000',
    tx2: '1F497D',
  };

  const parser = new DOMParser();

  try {
    const themeFile = zip.file('ppt/theme/theme1.xml');
    if (themeFile) {
      const themeXml = await themeFile.async('text');
      const themeDoc = parser.parseFromString(themeXml, 'application/xml');
      const clrScheme = themeDoc.getElementsByTagName('a:clrScheme')[0];
      if (clrScheme) {
        for (const key of [
          'dk1',
          'lt1',
          'dk2',
          'lt2',
          'accent1',
          'accent2',
          'accent3',
          'accent4',
          'accent5',
          'accent6',
        ]) {
          const node = clrScheme.getElementsByTagName(`a:${key}`)[0];
          if (node) {
            const srgb = node.getElementsByTagName('a:srgbClr')[0]?.getAttribute('val');
            const sysLast = node.getElementsByTagName('a:sysClr')[0]?.getAttribute('lastClr');
            const hex = (srgb || sysLast || '').replace(/[^0-9A-Fa-f]/g, '').toUpperCase();
            if (hex.length === 6) {
              schemeColors[key] = hex;
            }
          }
        }
        schemeColors.bg1 = schemeColors.lt1;
        schemeColors.bg2 = schemeColors.lt2;
        schemeColors.tx1 = schemeColors.dk1;
        schemeColors.tx2 = schemeColors.dk2;
      }
    }
  } catch {
    // Fallback to standard Office theme map
  }

  const extractBgHexFromXmlDoc = (xmlDoc: Document): string | undefined => {
    const bgNodes = xmlDoc.getElementsByTagName('p:bg');
    if (bgNodes.length === 0) return undefined;
    const bgNode = bgNodes[0];

    const srgbNodes = bgNode.getElementsByTagName('a:srgbClr');
    if (srgbNodes.length > 0) {
      const val = (srgbNodes[0].getAttribute('val') || '').replace(/[^0-9A-Fa-f]/g, '').toUpperCase();
      if (val.length === 6) return val;
    }

    const schemeNodes = bgNode.getElementsByTagName('a:schemeClr');
    if (schemeNodes.length > 0) {
      const val = (schemeNodes[0].getAttribute('val') || '').trim();
      if (val && schemeColors[val]) {
        return schemeColors[val];
      }
    }
    return undefined;
  };

  let masterBgHex: string | undefined;
  try {
    const masterFile = zip.file('ppt/slideMasters/slideMaster1.xml');
    if (masterFile) {
      const masterXml = await masterFile.async('text');
      const masterDoc = parser.parseFromString(masterXml, 'application/xml');
      masterBgHex = extractBgHexFromXmlDoc(masterDoc);
    }
  } catch {
    // Ignore master bg error
  }

  const slides: SlideContent[] = [];
  const allBlocks: StructuredBlock[] = [];

  for (let i = 0; i < slideEntries.length; i++) {
    const entry = slideEntries[i];
    onProgress?.(`Parsing slide ${i + 1} of ${slideEntries.length}...`);
    const fileObj = zip.file(entry.path);
    if (!fileObj) continue;

    const xmlText = await fileObj.async('text');
    const xmlDoc = parser.parseFromString(xmlText, 'application/xml');

    // Determine slide background color: slide <p:bg> -> slideLayout <p:bg> -> slideMaster <p:bg>
    let bgColorHex = extractBgHexFromXmlDoc(xmlDoc);
    if (!bgColorHex) {
      try {
        const relsPath = `ppt/slides/_rels/slide${entry.num}.xml.rels`;
        const relsFile = zip.file(relsPath);
        if (relsFile) {
          const relsXml = await relsFile.async('text');
          const layoutMatch = relsXml.match(/Target="\.\.\/slideLayouts\/(slideLayout\d+\.xml)"/i);
          if (layoutMatch) {
            const layoutFile = zip.file(`ppt/slideLayouts/${layoutMatch[1]}`);
            if (layoutFile) {
              const layoutXml = await layoutFile.async('text');
              const layoutDoc = parser.parseFromString(layoutXml, 'application/xml');
              bgColorHex = extractBgHexFromXmlDoc(layoutDoc);
            }
          }
        }
      } catch {
        // Ignore layout bg lookup error
      }
    }
    if (!bgColorHex && masterBgHex) {
      bgColorHex = masterBgHex;
    }

    // Check if there is a full-slide background rectangle shape with solidFill when <p:bg> wasn't used
    const shapes = Array.from(xmlDoc.getElementsByTagName('p:sp'));
    if (!bgColorHex && shapes.length > 0) {
      const firstSp = shapes[0];
      const spText = (firstSp.textContent || '').trim();
      if (!spText) {
        const spPr = firstSp.getElementsByTagName('p:spPr')[0];
        if (spPr) {
          const solidFill = spPr.getElementsByTagName('a:solidFill')[0];
          if (solidFill) {
            const srgb = solidFill.getElementsByTagName('a:srgbClr')[0]?.getAttribute('val');
            const scheme = solidFill.getElementsByTagName('a:schemeClr')[0]?.getAttribute('val');
            if (srgb && /^[0-9A-Fa-f]{6}$/.test(srgb)) {
              bgColorHex = srgb.toUpperCase();
            } else if (scheme && schemeColors[scheme]) {
              bgColorHex = schemeColors[scheme];
            }
          }
        }
      }
    }

    let textColorHex = '0F172A';
    if (bgColorHex && bgColorHex.length === 6) {
      const r = parseInt(bgColorHex.slice(0, 2), 16);
      const g = parseInt(bgColorHex.slice(2, 4), 16);
      const b = parseInt(bgColorHex.slice(4, 6), 16);
      const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
      if (luminance < 0.48) {
        textColorHex = 'F8FAFC';
      }
    }

    const slideBlocks: StructuredBlock[] = [];
    let slideTitle = '';

    if (shapes.length > 0) {
      for (const sp of shapes) {
        const phNodes = sp.getElementsByTagName('p:ph');
        let isTitleShape = false;
        if (phNodes.length > 0) {
          const phType = (phNodes[0].getAttribute('type') || '').toLowerCase();
          if (phType === 'title' || phType === 'ctrtitle') {
            isTitleShape = true;
          }
        }

        const paragraphs = Array.from(sp.getElementsByTagName('a:p'));
        for (const p of paragraphs) {
          const textNodes = Array.from(p.getElementsByTagName('a:t'));
          const line = textNodes
            .map((t) => t.textContent || '')
            .join('')
            .replace(/\s+/g, ' ')
            .trim();
          if (!line) continue;

          const hasBullet =
            p.getElementsByTagName('a:buChar').length > 0 ||
            p.getElementsByTagName('a:buAutoNum').length > 0;

          if (isTitleShape && !slideTitle) {
            slideTitle = line;
            slideBlocks.push({ type: 'h1', text: line });
          } else if (!slideTitle && slideBlocks.length === 0) {
            slideTitle = line;
            slideBlocks.push({ type: 'h2', text: line });
          } else if (hasBullet) {
            slideBlocks.push({ type: 'li', text: line });
          } else {
            slideBlocks.push({ type: 'p', text: line });
          }
        }
      }
    }

    const tables = Array.from(xmlDoc.getElementsByTagName('a:tbl'));
    for (const tbl of tables) {
      const tableRows = Array.from(tbl.getElementsByTagName('a:tr'));
      tableRows.forEach((tr, rowIdx) => {
        const cells = Array.from(tr.getElementsByTagName('a:tc')).map((tc) =>
          Array.from(tc.getElementsByTagName('a:t'))
            .map((t) => t.textContent || '')
            .join('')
            .replace(/\s+/g, ' ')
            .trim()
        );
        if (cells.some((c) => c.length > 0)) {
          slideBlocks.push({
            type: 'table-row',
            text: cells.join(' | '),
            cells,
            isHeaderRow: rowIdx === 0,
          });
        }
      });
    }

    if (slideBlocks.length === 0) {
      const allParas = Array.from(xmlDoc.getElementsByTagName('a:p'));
      for (const p of allParas) {
        const line = Array.from(p.getElementsByTagName('a:t'))
          .map((t) => t.textContent || '')
          .join('')
          .replace(/\s+/g, ' ')
          .trim();
        if (!line) continue;
        if (!slideTitle) {
          slideTitle = line;
          slideBlocks.push({ type: 'h2', text: line });
        } else {
          slideBlocks.push({ type: 'p', text: line });
        }
      }
    }

    slides.push({
      slideNumber: i + 1,
      title: slideTitle,
      blocks: slideBlocks,
      bgColorHex,
      textColorHex,
    });
    allBlocks.push(...slideBlocks);
  }

  const rawText = allBlocks.map((b) => b.text).join('\n\n');
  return {
    slides,
    blocks: allBlocks,
    rawText,
    dimensionsLabel,
  };
}

/**
 * Inspects an uploaded PDF, DOCX, or PPTX file and returns a QueuedDocument
 */
export async function inspectUploadedFile(
  file: File,
  onFileProgress?: (percentWithinFile: number, stepDetail: string) => void
): Promise<QueuedDocument> {
  const id = `doc_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
  const now = new Date();
  const addedAt = now.toTimeString().slice(0, 8);
  const lowerName = file.name.toLowerCase();
  const extension: SourceFileType = lowerName.endsWith('.pptx')
    ? 'pptx'
    : lowerName.endsWith('.docx')
    ? 'docx'
    : 'pdf';

  try {
    onFileProgress?.(10, `Reading binary stream for ${file.name}...`);
    const arrayBuffer = await readFileBufferWithProgress(file, (fraction) => {
      const streamPercent = Math.round(10 + fraction * 35);
      onFileProgress?.(
        streamPercent,
        `Uploading & reading ${file.name} (${Math.round(fraction * 100)}%)...`
      );
    });
    await yieldToUi(18);

    if (extension === 'pdf') {
      onFileProgress?.(55, `Validating PDF header & page geometry in ${file.name}...`);
      const pdfDoc = await PDFDocument.load(arrayBuffer, { ignoreEncryption: true });
      const pageCount = Math.max(1, pdfDoc.getPageCount());
      const firstPage = pdfDoc.getPage(0);
      const { width, height } = firstPage.getSize();
      const wPt = Math.round(width);
      const hPt = Math.round(height);
      let standardName = 'Custom';
      if (Math.abs(wPt - 612) <= 6 && Math.abs(hPt - 792) <= 6) standardName = 'US Letter';
      else if (Math.abs(wPt - 595) <= 6 && Math.abs(hPt - 842) <= 6) standardName = 'ISO A4';

      onFileProgress?.(
        82,
        `Extracting text preview from ${pageCount} PDF ${pageCount === 1 ? 'page' : 'pages'}...`
      );
      await yieldToUi(15);

      const blocks = extractBasicTextFromPdfBuffer(arrayBuffer);
      const rawText = blocks.map((b) => b.text).join('\n\n');
      const wordCount = rawText.split(/\s+/).filter(Boolean).length;

      const previewHtml = blocks
        .map((b) =>
          b.type === 'h2'
            ? `<h3 class="font-semibold mb-2">${escapeXml(b.text)}</h3>`
            : `<p class="mb-2">${escapeXml(b.text)}</p>`
        )
        .join('');

      onFileProgress?.(100, `Inspection complete for ${file.name}`);

      return {
        id,
        file,
        name: file.name,
        extension: 'pdf',
        sizeBytes: file.size,
        addedAt,
        status: 'ready',
        pageCount,
        wordCount: Math.max(wordCount, pageCount * 120),
        dimensionsLabel: `${wPt} × ${hPt} pt (${standardName})`,
        included: true,
        pageRangeInput: 'All',
        rotationDegrees: 0,
        arrayBuffer,
        previewHtml,
        rawText,
        blocks,
      };
    } else if (extension === 'pptx') {
      onFileProgress?.(55, `Unpacking OpenXML PowerPoint (.pptx) slides for ${file.name}...`);
      const { slides, blocks, rawText, dimensionsLabel } = await extractSlidesFromPptxBuffer(
        arrayBuffer,
        (detail) => onFileProgress?.(78, detail)
      );
      await yieldToUi(15);

      const slideCount = Math.max(1, slides.length);
      const words = rawText.trim().split(/\s+/).filter(Boolean);
      const wordCount = words.length;

      const previewHtml = slides
        .map(
          (s) =>
            `<div class="mb-3"><h4 class="font-bold">Slide ${s.slideNumber}${
              s.title ? `: ${escapeXml(s.title)}` : ''
            }</h4></div>`
        )
        .join('');

      onFileProgress?.(100, `Inspection complete for ${file.name} (${slideCount} slides)`);

      return {
        id,
        file,
        name: file.name,
        extension: 'pptx',
        sizeBytes: file.size,
        addedAt,
        status: 'ready',
        pageCount: slideCount,
        wordCount,
        dimensionsLabel,
        included: true,
        pageRangeInput: 'All',
        rotationDegrees: 0,
        arrayBuffer,
        previewHtml,
        rawText,
        blocks,
        slides,
      };
    } else {
      onFileProgress?.(55, `Unpacking OpenXML (.docx) container for ${file.name}...`);
      const [htmlResult, textResult] = await Promise.all([
        mammoth.convertToHtml({ arrayBuffer }),
        mammoth.extractRawText({ arrayBuffer }),
      ]);

      onFileProgress?.(80, `Parsing structured paragraphs, headers & footers in ${file.name}...`);
      await yieldToUi(15);

      let rawText = textResult.value || '';
      let blocks = parseHtmlToBlocks(htmlResult.value || '', rawText);
      let docxHeaders: string[] = [];
      let docxFooters: string[] = [];

      let estimatedPages = 1;
      try {
        const zip = await JSZip.loadAsync(arrayBuffer);

        // Extract native Word headers and footers from word/header*.xml and word/footer*.xml
        const extractedHF = await extractDocxHeadersAndFooters(zip);
        docxHeaders = extractedHF.headers;
        docxFooters = extractedHF.footers;

        // If any header/footer lines accidentally appeared in body blocks, filter them out of body blocks
        // so they are merged strictly as native headers and footers rather than plain body text
        if (docxHeaders.length > 0 || docxFooters.length > 0) {
          const hfSet = new Set([...docxHeaders, ...docxFooters].map((s) => s.trim()));
          const filteredBlocks = blocks.filter((b) => !hfSet.has(b.text.trim()));
          if (filteredBlocks.length > 0) {
            blocks = filteredBlocks;
          }
        }

        // Also parse word/document.xml directly to ensure all <w:tbl> tables and <w:p> paragraphs are captured with exact cell arrays
        const docXmlFile = zip.file('word/document.xml');
        if (docXmlFile) {
          const docXml = await docXmlFile.async('text');
          const parser = new DOMParser();
          const xmlDoc = parser.parseFromString(docXml, 'application/xml');
          const bodyNode = xmlDoc.getElementsByTagName('w:body')[0];

          if (bodyNode) {
            const xmlBlocks: StructuredBlock[] = [];
            const children = Array.from(bodyNode.childNodes);
            for (const child of children) {
              if (child.nodeType !== 1) continue;
              const el = child as Element;
              const tagName = el.tagName || el.nodeName;
              if (tagName === 'w:p') {
                const pStyle = el.getElementsByTagName('w:pStyle')[0]?.getAttribute('w:val') || '';
                const numPr = el.getElementsByTagName('w:numPr')[0];
                const tNodes = Array.from(el.getElementsByTagName('w:t'));
                const lineText = tNodes
                  .map((t) => t.textContent || '')
                  .join('')
                  .replace(/\s+/g, ' ')
                  .trim();
                if (!lineText) continue;
                if (/heading\s*1|^h1$/i.test(pStyle)) {
                  xmlBlocks.push({ type: 'h1', text: lineText });
                } else if (/heading\s*2|^h2$/i.test(pStyle)) {
                  xmlBlocks.push({ type: 'h2', text: lineText });
                } else if (/heading\s*[3-6]|^h[3-6]$/i.test(pStyle)) {
                  xmlBlocks.push({ type: 'h3', text: lineText });
                } else if (numPr) {
                  xmlBlocks.push({ type: 'li', text: lineText });
                } else {
                  xmlBlocks.push({ type: 'p', text: lineText });
                }
              } else if (tagName === 'w:tbl') {
                const trNodes = Array.from(el.getElementsByTagName('w:tr'));
                trNodes.forEach((tr, rowIdx) => {
                  const tcNodes = Array.from(tr.childNodes).filter(
                    (n) => n.nodeType === 1 && (n as Element).tagName === 'w:tc'
                  ) as Element[];
                  const effectiveTc =
                    tcNodes.length > 0 ? tcNodes : Array.from(tr.getElementsByTagName('w:tc'));
                  const cells = effectiveTc.map((tc) => {
                    const cellParas = Array.from(tc.getElementsByTagName('w:p'));
                    if (cellParas.length > 0) {
                      return cellParas
                        .map((cp) =>
                          Array.from(cp.getElementsByTagName('w:t'))
                            .map((t) => t.textContent || '')
                            .join('')
                            .trim()
                        )
                        .filter(Boolean)
                        .join(' ');
                    }
                    return Array.from(tc.getElementsByTagName('w:t'))
                      .map((t) => t.textContent || '')
                      .join('')
                      .trim();
                  });
                  if (cells.some((c) => c.length > 0)) {
                    const hasTblHeader = tr.getElementsByTagName('w:tblHeader').length > 0;
                    xmlBlocks.push({
                      type: 'table-row',
                      text: cells.join(' | '),
                      cells,
                      isHeaderRow: rowIdx === 0 || hasTblHeader,
                    });
                  }
                });
              }
            }

            if (xmlBlocks.length > 0) {
              const hfSet = new Set([...docxHeaders, ...docxFooters].map((s) => s.trim()));
              const cleanedXmlBlocks =
                hfSet.size > 0
                  ? xmlBlocks.filter((b) => b.type === 'table-row' || !hfSet.has(b.text.trim()))
                  : xmlBlocks;
              if (cleanedXmlBlocks.length > 0) {
                blocks = cleanedXmlBlocks;
              }
              if (!rawText.trim()) {
                rawText = blocks.map((b) => b.text).join('\n\n');
              }
            }
          }
        }

        const words = rawText.trim().split(/\s+/).filter(Boolean);
        estimatedPages = Math.max(1, Math.ceil(words.length / 260));

        const appXmlFile = zip.file('docProps/app.xml');
        if (appXmlFile) {
          const appXml = await appXmlFile.async('text');
          const match = appXml.match(/<Pages>(\d+)<\/Pages>/);
          if (match && parseInt(match[1], 10) > 0) {
            estimatedPages = Math.max(estimatedPages, parseInt(match[1], 10));
          }
        }
      } catch {
        // Fallback to word count page calculation
      }

      const words = rawText.trim().split(/\s+/).filter(Boolean);
      const wordCount = words.length;

      onFileProgress?.(100, `Inspection complete for ${file.name}`);

      return {
        id,
        file,
        name: file.name,
        extension: 'docx',
        sizeBytes: file.size,
        addedAt,
        status: 'ready',
        pageCount: estimatedPages,
        wordCount,
        dimensionsLabel: 'OpenXML Flow (US Letter)',
        included: true,
        pageRangeInput: 'All',
        rotationDegrees: 0,
        arrayBuffer,
        previewHtml: htmlResult.value || `<p>${escapeXml(rawText)}</p>`,
        rawText,
        blocks,
        docxHeaders,
        docxFooters,
      };
    }
  } catch (err) {
    const message = err instanceof Error ? err.message : 'Unable to parse file structure';
    onFileProgress?.(100, `Failed to parse ${file.name}`);
    return {
      id,
      file,
      name: file.name,
      extension,
      sizeBytes: file.size,
      addedAt,
      status: 'error',
      errorMessage: message,
      pageCount: 0,
      wordCount: 0,
      dimensionsLabel: 'Unreadable',
      included: false,
      pageRangeInput: 'All',
      rotationDegrees: 0,
      arrayBuffer: new ArrayBuffer(0),
      previewHtml: '',
      rawText: '',
      blocks: [],
    };
  }
}

/**
 * Wraps a line of text to fit within maxWidth in pdf-lib
 */
function wrapTextLines(
  text: string,
  font: Awaited<ReturnType<PDFDocument['embedFont']>>,
  fontSize: number,
  maxWidth: number
): string[] {
  const sanitized = sanitizeForWinAnsi(text);
  const paragraphs = sanitized.split(/\r?\n/);
  const lines: string[] = [];

  for (const para of paragraphs) {
    const words = para.split(/\s+/).filter(Boolean);
    if (words.length === 0) {
      lines.push('');
      continue;
    }
    let currentLine = words[0];
    for (let i = 1; i < words.length; i++) {
      const candidate = `${currentLine} ${words[i]}`;
      const width = font.widthOfTextAtSize(candidate, fontSize);
      if (width <= maxWidth) {
        currentLine = candidate;
      } else {
        lines.push(currentLine);
        currentLine = words[i];
      }
    }
    lines.push(currentLine);
  }
  return lines;
}

/**
 * Merges active QueuedDocuments into a single PDF, DOCX, or PPTX artifact
 */
export async function mergeDocuments(
  docs: QueuedDocument[],
  config: MergeConfiguration,
  onMergeProgress?: (progress: Omit<MergeProgressState, 'phase'>) => void
): Promise<MergedArtifact> {
  const startTime = performance.now();
  const activeDocs = docs.filter((d) => d.included && d.status === 'ready');

  if (activeDocs.length === 0) {
    throw new Error('No active documents selected for merging.');
  }

  let resolvedFormat: 'pdf' | 'docx' | 'pptx' = 'pdf';
  if (config.outputFormat === 'docx') {
    resolvedFormat = 'docx';
  } else if (config.outputFormat === 'pptx') {
    resolvedFormat = 'pptx';
  } else if (config.outputFormat === 'pdf') {
    resolvedFormat = 'pdf';
  } else {
    const allDocx = activeDocs.every((d) => d.extension === 'docx');
    const allPptx = activeDocs.every((d) => d.extension === 'pptx');
    if (allDocx) resolvedFormat = 'docx';
    else if (allPptx) resolvedFormat = 'pptx';
    else resolvedFormat = 'pdf';
  }

  const baseFilename = (config.outputFilename || 'Merged_Document')
    .trim()
    .replace(/\.(pdf|docx|pptx)$/i, '')
    .replace(/[^a-zA-Z0-9_\-. ]/g, '_');

  if (resolvedFormat === 'pdf') {
    return await buildUnifiedPdf(activeDocs, config, baseFilename, startTime, onMergeProgress);
  } else if (resolvedFormat === 'pptx') {
    return await buildUnifiedPptx(activeDocs, config, baseFilename, startTime, onMergeProgress);
  } else {
    return await buildUnifiedDocx(activeDocs, config, baseFilename, startTime, onMergeProgress);
  }
}

async function buildUnifiedPdf(
  activeDocs: QueuedDocument[],
  config: MergeConfiguration,
  baseFilename: string,
  startTime: number,
  onMergeProgress?: (progress: Omit<MergeProgressState, 'phase'>) => void
): Promise<MergedArtifact> {
  onMergeProgress?.({
    percent: 8,
    processedDocs: 0,
    totalDocs: activeDocs.length,
    currentDocName: activeDocs[0]?.name || '',
    stepDetail: 'Initializing PDF document...',
  });
  await yieldToUi(20);

  const mergedPdf = await PDFDocument.create();
  const fontRegular = await mergedPdf.embedFont(StandardFonts.Helvetica);
  const fontBold = await mergedPdf.embedFont(StandardFonts.HelveticaBold);
  const fontSerif = await mergedPdf.embedFont(StandardFonts.TimesRoman);

  const stdPageWidth = config.pageSize === 'a4' ? 595.28 : 612;
  const stdPageHeight = config.pageSize === 'a4' ? 841.89 : 792;

  // Slide dimensions when rendering PPTX slides into PDF
  const slideWidthPt = 792;
  const slideHeightPt =
    config.slideRatio === '4:3'
      ? 594
      : config.slideRatio === '16:10'
      ? 495
      : 445.5;

  const manifest: MergedArtifact['manifest'] = [];
  let totalWords = 0;

  const parseHexToRgb = (hex?: string, fallback: [number, number, number] = [1, 1, 1]) => {
    if (!hex) return rgb(fallback[0], fallback[1], fallback[2]);
    const clean = hex.replace(/[^0-9A-Fa-f]/g, '');
    if (clean.length !== 6) return rgb(fallback[0], fallback[1], fallback[2]);
    return rgb(
      parseInt(clean.slice(0, 2), 16) / 255,
      parseInt(clean.slice(2, 4), 16) / 255,
      parseInt(clean.slice(4, 6), 16) / 255
    );
  };

  const renderBlocksOntoPdf = (
    blocksToRender: StructuredBlock[],
    pWidth: number,
    pHeight: number,
    headerLines?: string[],
    footerLines?: string[],
    bgColorHex?: string,
    textColorHex?: string
  ) => {
    const margin = 54;
    const contentWidth = pWidth - margin * 2;
    const isDarkBg = textColorHex === 'F8FAFC';

    const paintPageBackgroundAndHeaders = (page: ReturnType<PDFDocument['addPage']>) => {
      if (bgColorHex && bgColorHex.length === 6 && bgColorHex.toUpperCase() !== 'FFFFFF') {
        page.drawRectangle({
          x: 0,
          y: 0,
          width: pWidth,
          height: pHeight,
          color: parseHexToRgb(bgColorHex, [1, 1, 1]),
        });
      }
      if (headerLines && headerLines.length > 0) {
        const headerText = sanitizeForWinAnsi(headerLines.join('  ·  ')).slice(0, 110);
        if (headerText) {
          page.drawText(headerText, {
            x: margin,
            y: pHeight - 30,
            size: 8.5,
            font: fontRegular,
            color: isDarkBg ? rgb(0.78, 0.84, 0.92) : rgb(0.42, 0.47, 0.55),
          });
        }
      }
      if (footerLines && footerLines.length > 0) {
        const footerText = sanitizeForWinAnsi(footerLines.join('  ·  ')).slice(0, 110);
        if (footerText) {
          page.drawText(footerText, {
            x: margin,
            y: 28,
            size: 8.5,
            font: fontRegular,
            color: isDarkBg ? rgb(0.78, 0.84, 0.92) : rgb(0.42, 0.47, 0.55),
          });
        }
      }
    };

    let currentPage = mergedPdf.addPage([pWidth, pHeight]);
    paintPageBackgroundAndHeaders(currentPage);
    if (blocksToRender.length === 0) return;

    let cursorY = pHeight - margin;

    let bIdx = 0;
    while (bIdx < blocksToRender.length) {
      const block = blocksToRender[bIdx];

      // Group consecutive table-row blocks into a bordered multi-column table
      if (block.type === 'table-row') {
        const tableRows: StructuredBlock[] = [];
        while (bIdx < blocksToRender.length && blocksToRender[bIdx].type === 'table-row') {
          tableRows.push(blocksToRender[bIdx]);
          bIdx++;
        }

        const numCols = Math.max(
          1,
          ...tableRows.map((r) =>
            r.cells && r.cells.length > 0
              ? r.cells.length
              : r.text.split(/\s*\|\s*/).filter(Boolean).length
          )
        );
        const colWidth = contentWidth / numCols;
        const cellPadX = 6;
        const cellPadY = 5;
        const cellFontSize = 9.5;
        const cellLineHeight = 13;

        cursorY -= 4;

        for (let rIdx = 0; rIdx < tableRows.length; rIdx++) {
          const row = tableRows[rIdx];
          const rawCells =
            row.cells && row.cells.length > 0
              ? row.cells
              : row.text.split(/\s*\|\s*/);
          const isHeader = Boolean(row.isHeaderRow || rIdx === 0);
          const rowFont = isHeader ? fontBold : fontRegular;

          const wrappedCells: string[][] = [];
          let maxLinesInRow = 1;
          for (let c = 0; c < numCols; c++) {
            const cellStr = rawCells[c] ?? '';
            const lines = wrapTextLines(
              cellStr,
              rowFont,
              cellFontSize,
              Math.max(24, colWidth - cellPadX * 2)
            );
            wrappedCells.push(lines);
            if (lines.length > maxLinesInRow) {
              maxLinesInRow = lines.length;
            }
          }

          const rowHeight = maxLinesInRow * cellLineHeight + cellPadY * 2;

          if (cursorY - rowHeight < margin + 32) {
            currentPage = mergedPdf.addPage([pWidth, pHeight]);
            paintPageBackgroundAndHeaders(currentPage);
            cursorY = pHeight - margin;
          }

          const rowTopY = cursorY;
          const rowBottomY = cursorY - rowHeight;

          for (let c = 0; c < numCols; c++) {
            const cellX = margin + c * colWidth;

            currentPage.drawRectangle({
              x: cellX,
              y: rowBottomY,
              width: colWidth,
              height: rowHeight,
              color: isHeader
                ? isDarkBg
                  ? rgb(0.18, 0.24, 0.34)
                  : rgb(0.93, 0.95, 0.98)
                : undefined,
              borderColor: isDarkBg ? rgb(0.45, 0.55, 0.68) : rgb(0.68, 0.74, 0.82),
              borderWidth: 0.75,
            });

            const cellLines = wrappedCells[c] || [];
            for (let lIdx = 0; lIdx < cellLines.length; lIdx++) {
              const lineStr = cellLines[lIdx];
              if (lineStr.trim().length > 0) {
                currentPage.drawText(lineStr, {
                  x: cellX + cellPadX,
                  y: rowTopY - cellPadY - cellFontSize - lIdx * cellLineHeight + 1,
                  size: cellFontSize,
                  font: rowFont,
                  color: isDarkBg
                    ? rgb(0.96, 0.98, 1)
                    : isHeader
                    ? rgb(0.06, 0.09, 0.16)
                    : rgb(0.15, 0.2, 0.28),
                });
              }
            }
          }

          cursorY -= rowHeight;
        }

        cursorY -= 10;
        continue;
      }

      let fontSize = 11;
      let lineHeight = 16;
      let blockFont = fontSerif;
      let indent = 0;
      let spaceAfter = 8;
      let prefix = '';

      if (block.type === 'h1') {
        fontSize = 18;
        lineHeight = 24;
        blockFont = fontBold;
        spaceAfter = 12;
        cursorY -= 6;
      } else if (block.type === 'h2') {
        fontSize = 14;
        lineHeight = 20;
        blockFont = fontBold;
        spaceAfter = 10;
        cursorY -= 4;
      } else if (block.type === 'h3') {
        fontSize = 12;
        lineHeight = 17;
        blockFont = fontBold;
        spaceAfter = 8;
      } else if (block.type === 'li') {
        fontSize = 11;
        lineHeight = 15.5;
        blockFont = fontRegular;
        indent = 14;
        prefix = '-  ';
        spaceAfter = 5;
      } else {
        fontSize = 11;
        lineHeight = 16;
        blockFont = fontRegular;
        spaceAfter = 9;
      }

      const lines = wrapTextLines(
        prefix + block.text,
        blockFont,
        fontSize,
        contentWidth - indent
      );

      for (const line of lines) {
        if (cursorY - lineHeight < margin + 32) {
          currentPage = mergedPdf.addPage([pWidth, pHeight]);
          paintPageBackgroundAndHeaders(currentPage);
          cursorY = pHeight - margin;
        }
        if (line.trim().length > 0) {
          currentPage.drawText(line, {
            x: margin + indent,
            y: cursorY - fontSize,
            size: fontSize,
            font: blockFont,
            color: isDarkBg
              ? rgb(0.96, 0.98, 1)
              : block.type === 'h1' || block.type === 'h2' || block.type === 'h3'
              ? rgb(0.06, 0.09, 0.16)
              : rgb(0.15, 0.2, 0.28),
          });
        }
        cursorY -= lineHeight;
      }
      cursorY -= spaceAfter;
      bIdx++;
    }
  };

  for (let idx = 0; idx < activeDocs.length; idx++) {
    const doc = activeDocs[idx];
    const stepBasePercent = Math.round(12 + (idx / activeDocs.length) * 72);
    onMergeProgress?.({
      percent: stepBasePercent,
      processedDocs: idx,
      totalDocs: activeDocs.length,
      currentDocName: doc.name,
      stepDetail: `Merging document ${idx + 1} of ${activeDocs.length}: ${doc.name}...`,
    });
    await yieldToUi(24);

    totalWords += doc.wordCount;
    const startPage = mergedPdf.getPageCount() + 1;

    if (doc.extension === 'pdf') {
      const srcPdf = await PDFDocument.load(doc.arrayBuffer, { ignoreEncryption: true });
      const totalSrcPages = srcPdf.getPageCount();
      const selectedIndices = parsePageRange(doc.pageRangeInput, totalSrcPages);

      const copiedPages = await mergedPdf.copyPages(srcPdf, selectedIndices);
      for (const page of copiedPages) {
        if (doc.rotationDegrees !== 0) {
          const currentRotation = page.getRotation().angle;
          page.setRotation(degrees((currentRotation + doc.rotationDegrees) % 360));
        }
        mergedPdf.addPage(page);
      }
    } else if (doc.extension === 'pptx') {
      const slides = doc.slides && doc.slides.length > 0 ? doc.slides : [];
      if (slides.length > 0) {
        const selectedSlideIndices = parsePageRange(doc.pageRangeInput, slides.length);
        for (const sIdx of selectedSlideIndices) {
          const slide = slides[sIdx];
          if (!slide) continue;
          renderBlocksOntoPdf(
            slide.blocks,
            slideWidthPt,
            slideHeightPt,
            undefined,
            undefined,
            slide.bgColorHex,
            slide.textColorHex
          );
        }
      } else if (doc.blocks.length > 0) {
        renderBlocksOntoPdf(doc.blocks, slideWidthPt, slideHeightPt);
      }
    } else {
      const blocksToRender: StructuredBlock[] =
        doc.blocks.length > 0
          ? doc.blocks
          : doc.rawText.trim()
          ? [{ type: 'p', text: doc.rawText.trim() }]
          : [];
      renderBlocksOntoPdf(
        blocksToRender,
        stdPageWidth,
        stdPageHeight,
        doc.docxHeaders,
        doc.docxFooters
      );
    }

    const endPage = Math.max(startPage, mergedPdf.getPageCount());
    manifest.push({
      docId: doc.id,
      name: doc.name,
      extension: doc.extension,
      pagesIncluded: Math.max(1, endPage - startPage + 1),
      startPage,
      endPage,
    });
  }

  if (mergedPdf.getPageCount() === 0) {
    mergedPdf.addPage([stdPageWidth, stdPageHeight]);
  }

  const finalTotalPages = mergedPdf.getPageCount();
  if (config.stampUnifiedPageNumbers) {
    const pages = mergedPdf.getPages();
    for (let i = 0; i < pages.length; i++) {
      const p = pages[i];
      const { width } = p.getSize();
      const footerText = sanitizeForWinAnsi(`Page ${i + 1} of ${finalTotalPages}`);
      const textW = fontRegular.widthOfTextAtSize(footerText, 8.5);
      p.drawText(footerText, {
        x: Math.max(24, (width - textW) / 2),
        y: 16,
        size: 8.5,
        font: fontRegular,
        color: rgb(0.45, 0.5, 0.58),
      });
    }
  }

  onMergeProgress?.({
    percent: 94,
    processedDocs: activeDocs.length,
    totalDocs: activeDocs.length,
    currentDocName: `${baseFilename}.pdf`,
    stepDetail: 'Serializing unified PDF binary stream...',
  });
  await yieldToUi(16);

  const pdfBytes = await mergedPdf.save();
  const blob = new Blob([pdfBytes as unknown as BlobPart], { type: 'application/pdf' });
  const blobUrl = URL.createObjectURL(blob);
  const durationMs = Math.max(12, Math.round(performance.now() - startTime));

  onMergeProgress?.({
    percent: 100,
    processedDocs: activeDocs.length,
    totalDocs: activeDocs.length,
    currentDocName: `${baseFilename}.pdf`,
    stepDetail: `Merged ${activeDocs.length} documents (${finalTotalPages} pages) into ${baseFilename}.pdf`,
  });

  return {
    id: `merged_${Date.now()}`,
    filename: `${baseFilename}.pdf`,
    format: 'pdf',
    blob,
    blobUrl,
    sizeBytes: blob.size,
    totalPages: finalTotalPages,
    totalWords,
    sourceCount: activeDocs.length,
    mergedAt: new Date().toTimeString().slice(0, 8),
    durationMs,
    manifest,
  };
}

/**
 * Builds a valid OpenXML (.docx) file combining all active documents, merging headers and footers
 * as true OpenXML <w:hdr> (word/header*.xml) and <w:ftr> (word/footer*.xml) parts.
 */
async function buildUnifiedDocx(
  activeDocs: QueuedDocument[],
  config: MergeConfiguration,
  baseFilename: string,
  startTime: number,
  onMergeProgress?: (progress: Omit<MergeProgressState, 'phase'>) => void
): Promise<MergedArtifact> {
  onMergeProgress?.({
    percent: 10,
    processedDocs: 0,
    totalDocs: activeDocs.length,
    currentDocName: activeDocs[0]?.name || '',
    stepDetail: 'Initializing OpenXML (.docx) archive structure...',
  });
  await yieldToUi(20);

  const zip = new JSZip();
  const manifest: MergedArtifact['manifest'] = [];
  let totalWords = 0;
  let runningPage = 1;

  const pageWidthTwips = config.pageSize === 'a4' ? 11906 : 12240;
  const pageHeightTwips = config.pageSize === 'a4' ? 16838 : 15840;

  const bodyXmlParts: string[] = [];
  const docRels: string[] = [];
  const contentTypeOverrides: string[] = [];

  // Track inherited header/footer across sections (Word "Link to Previous" behavior)
  let currentHeaderLines: string[] = [];
  let currentFooterLines: string[] = [];

  // Pre-seed with the first document that has headers or footers if the very first doc doesn't have any
  for (const d of activeDocs) {
    if (d.docxHeaders && d.docxHeaders.length > 0 && currentHeaderLines.length === 0) {
      currentHeaderLines = d.docxHeaders;
    }
    if (d.docxFooters && d.docxFooters.length > 0 && currentFooterLines.length === 0) {
      currentFooterLines = d.docxFooters;
    }
  }

  const sectionHeaderRefs: Array<{ hdrRelId?: string; ftrRelId?: string }> = [];

  for (let i = 0; i < activeDocs.length; i++) {
    const doc = activeDocs[i];
    const stepPercent = Math.round(15 + (i / activeDocs.length) * 68);
    onMergeProgress?.({
      percent: stepPercent,
      processedDocs: i,
      totalDocs: activeDocs.length,
      currentDocName: doc.name,
      stepDetail: `Merging document ${i + 1} of ${activeDocs.length}: ${doc.name}...`,
    });
    await yieldToUi(24);

    if (doc.docxHeaders && doc.docxHeaders.length > 0) {
      currentHeaderLines = doc.docxHeaders;
    }
    if (doc.docxFooters && doc.docxFooters.length > 0) {
      currentFooterLines = doc.docxFooters;
    }

    let hdrRelId: string | undefined;
    let ftrRelId: string | undefined;

    if (currentHeaderLines.length > 0) {
      const hdrIdx = i + 1;
      hdrRelId = `rIdHdr${hdrIdx}`;
      const hdrFileName = `header${hdrIdx}.xml`;
      const hdrParas = currentHeaderLines
        .map((line) =>
          createDocxParagraph(line, {
            sizeHalfPt: 18,
            color: '64748B',
            spacingAfter: 60,
          })
        )
        .join('\n  ');

      zip.file(
        `word/${hdrFileName}`,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${hdrParas}
</w:hdr>`
      );
      docRels.push(
        `  <Relationship Id="${hdrRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="${hdrFileName}"/>`
      );
      contentTypeOverrides.push(
        `  <Override PartName="/word/${hdrFileName}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`
      );
    }

    if (currentFooterLines.length > 0 || config.stampUnifiedPageNumbers) {
      const ftrIdx = i + 1;
      ftrRelId = `rIdFtr${ftrIdx}`;
      const ftrFileName = `footer${ftrIdx}.xml`;
      const ftrParas: string[] = currentFooterLines.map((line) =>
        createDocxParagraph(line, {
          sizeHalfPt: 18,
          color: '64748B',
          spacingAfter: 60,
        })
      );

      if (config.stampUnifiedPageNumbers) {
        ftrParas.push(`<w:p>
    <w:pPr><w:jc w:val="center"/></w:pPr>
    <w:r><w:rPr><w:color w:val="64748B"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t xml:space="preserve">Page </w:t></w:r>
    <w:fldSimple w:instr="PAGE"><w:r><w:rPr><w:color w:val="64748B"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>
    <w:r><w:rPr><w:color w:val="64748B"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t xml:space="preserve"> of </w:t></w:r>
    <w:fldSimple w:instr="NUMPAGES"><w:r><w:rPr><w:color w:val="64748B"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>
  </w:p>`);
      }

      zip.file(
        `word/${ftrFileName}`,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${ftrParas.join('\n  ')}
</w:ftr>`
      );
      docRels.push(
        `  <Relationship Id="${ftrRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="${ftrFileName}"/>`
      );
      contentTypeOverrides.push(
        `  <Override PartName="/word/${ftrFileName}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`
      );
    }

    sectionHeaderRefs.push({ hdrRelId, ftrRelId });

    totalWords += doc.wordCount;
    const startPage = runningPage;
    const estimatedSpan = Math.max(1, doc.pageCount);
    const endPage = startPage + estimatedSpan - 1;
    runningPage = endPage + 1;

    manifest.push({
      docId: doc.id,
      name: doc.name,
      extension: doc.extension,
      pagesIncluded: estimatedSpan,
      startPage,
      endPage,
    });

    if (doc.extension === 'docx') {
      try {
        const srcZip = await JSZip.loadAsync(doc.arrayBuffer);
        const hfSet = new Set(
          [...(doc.docxHeaders || []), ...(doc.docxFooters || [])].map((s) => s.trim())
        );
        const richParts = await extractCleanOpenXmlFromDocx(srcZip, hfSet);
        if (richParts.length > 0) {
          bodyXmlParts.push(...richParts);
        } else {
          appendStructuredBlocksToDocx(doc.blocks, bodyXmlParts);
        }
      } catch {
        appendStructuredBlocksToDocx(doc.blocks, bodyXmlParts);
      }
    } else {
      const blocks =
        doc.blocks.length > 0
          ? doc.blocks
          : doc.rawText.trim()
          ? [{ type: 'p' as const, text: doc.rawText.trim() }]
          : [{ type: 'p' as const, text: '' }];
      appendStructuredBlocksToDocx(blocks, bodyXmlParts);
    }

    // Add a section break for each intermediate document so its native header/footer applies properly
    if (i < activeDocs.length - 1) {
      const breakType = config.insertPageBreakBetweenDocs ? 'nextPage' : 'continuous';
      bodyXmlParts.push(`<w:p>
      <w:pPr>
        <w:sectPr>
          ${hdrRelId ? `<w:headerReference w:type="default" r:id="${hdrRelId}"/>` : ''}
          ${ftrRelId ? `<w:footerReference w:type="default" r:id="${ftrRelId}"/>` : ''}
          <w:type w:val="${breakType}"/>
          <w:pgSz w:w="${pageWidthTwips}" w:h="${pageHeightTwips}"/>
          <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
        </w:sectPr>
      </w:pPr>
    </w:p>`);
    }
  }

  if (
    bodyXmlParts.length === 0 ||
    bodyXmlParts[bodyXmlParts.length - 1].trim().startsWith('<w:tbl')
  ) {
    bodyXmlParts.push(createDocxParagraph(''));
  }

  const lastSectionRefs = sectionHeaderRefs[sectionHeaderRefs.length - 1] || {};

  // Include standard word/settings.xml and word/fontTable.xml for full MS Word Desktop compatibility
  zip.file(
    'word/settings.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>`
  );
  docRels.push(
    `  <Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>`
  );
  contentTypeOverrides.push(
    `  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>`
  );

  zip.file(
    'word/fontTable.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:font w:name="Calibri">
    <w:panose1 w:val="020F0502020204030204"/>
    <w:charset w:val="00"/>
    <w:family w:val="swiss"/>
    <w:pitch w:val="variable"/>
  </w:font>
</w:fonts>`
  );
  docRels.push(
    `  <Relationship Id="rIdFontTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>`
  );
  contentTypeOverrides.push(
    `  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>`
  );

  // Always include a clean, self-contained word/styles.xml so MS Word Desktop renders headings and tables cleanly
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="120" w:line="276" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">
    <w:name w:val="Default Paragraph Font"/>
    <w:uiPriority w:val="1"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
  </w:style>
  <w:style w:type="table" w:default="1" w:styleId="TableNormal">
    <w:name w:val="Normal Table"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:tblPr>
      <w:tblInd w:w="0" w:type="dxa"/>
      <w:tblCellMar>
        <w:top w:w="0" w:type="dxa"/>
        <w:left w:w="108" w:type="dxa"/>
        <w:bottom w:w="0" w:type="dxa"/>
        <w:right w:w="108" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
  </w:style>
</w:styles>`
  );
  docRels.push(
    `  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`
  );
  contentTypeOverrides.push(
    `  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`
  );

  onMergeProgress?.({
    percent: 88,
    processedDocs: activeDocs.length,
    totalDocs: activeDocs.length,
    currentDocName: `${baseFilename}.docx`,
    stepDetail: 'Packaging OpenXML relationships, headers, footers & compressing .docx archive...',
  });
  await yieldToUi(20);

  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
${contentTypeOverrides.join('\n')}
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`
  );

  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`
  );

  const totalPages = Math.max(1, runningPage - 1);

  zip.file(
    'docProps/app.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office Word</Application>
  <Pages>${totalPages}</Pages>
  <Words>${totalWords}</Words>
  <AppVersion>16.0000</AppVersion>
</Properties>`
  );

  zip.file(
    'docProps/core.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${escapeXml(baseFilename)}</dc:title>
  <dc:creator>FolioBind</dc:creator>
  <cp:lastModifiedBy>FolioBind</cp:lastModifiedBy>
  <cp:revision>1</cp:revision>
</cp:coreProperties>`
  );

  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${docRels.join('\n')}
</Relationships>`
  );

  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    ${bodyXmlParts.join('\n    ')}
    <w:sectPr>
      ${
        lastSectionRefs.hdrRelId
          ? `<w:headerReference w:type="default" r:id="${lastSectionRefs.hdrRelId}"/>`
          : ''
      }
      ${
        lastSectionRefs.ftrRelId
          ? `<w:footerReference w:type="default" r:id="${lastSectionRefs.ftrRelId}"/>`
          : ''
      }
      <w:pgSz w:w="${pageWidthTwips}" w:h="${pageHeightTwips}"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`
  );

  const blob = await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  const blobUrl = URL.createObjectURL(blob);
  const durationMs = Math.max(14, Math.round(performance.now() - startTime));

  onMergeProgress?.({
    percent: 100,
    processedDocs: activeDocs.length,
    totalDocs: activeDocs.length,
    currentDocName: `${baseFilename}.docx`,
    stepDetail: `Merged ${activeDocs.length} documents (${totalPages} pages) into ${baseFilename}.docx`,
  });

  return {
    id: `merged_${Date.now()}`,
    filename: `${baseFilename}.docx`,
    format: 'docx',
    blob,
    blobUrl,
    sizeBytes: blob.size,
    totalPages,
    totalWords,
    sourceCount: activeDocs.length,
    mergedAt: new Date().toTimeString().slice(0, 8),
    durationMs,
    manifest,
  };
}

/**
 * Parses word/document.xml from a source .docx archive and reconstructs 100% ECMA-376 / MS Word Desktop
 * compliant <w:p> and <w:tbl> XML nodes without any undeclared namespace prefixes (like w14:, w15:, wp:)
 * or broken relationship references, while preserving table grid columns, cell borders, cell background shading,
 * column spans, alignments, bold/italic/underline/color runs, headings, and lists.
 */
async function extractCleanOpenXmlFromDocx(
  srcZip: JSZip,
  hfSet: Set<string>
): Promise<string[]> {
  const docXmlFile = srcZip.file('word/document.xml');
  if (!docXmlFile) return [];
  const docXml = await docXmlFile.async('text');
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(docXml, 'application/xml');
  if (xmlDoc.getElementsByTagName('parsererror').length > 0) {
    return [];
  }

  const bodyNode = xmlDoc.getElementsByTagName('w:body')[0];
  if (!bodyNode) return [];

  const serializeRunNode = (
    rNode: Element,
    inheritedBold: boolean,
    inheritedSizeHalfPt?: number,
    inheritedColor?: string
  ): string => {
    const rPrNode = rNode.getElementsByTagName('w:rPr')[0];
    const isBold =
      inheritedBold || Boolean(rPrNode && rPrNode.getElementsByTagName('w:b').length > 0);
    const isItalic = Boolean(rPrNode && rPrNode.getElementsByTagName('w:i').length > 0);
    const isUnderline = Boolean(rPrNode && rPrNode.getElementsByTagName('w:u').length > 0);
    const isStrike = Boolean(rPrNode && rPrNode.getElementsByTagName('w:strike').length > 0);

    let colorVal = inheritedColor || '';
    const colorEl = rPrNode?.getElementsByTagName('w:color')[0];
    if (colorEl) {
      const rawClr = (colorEl.getAttribute('w:val') || '').trim();
      if (/^[0-9A-Fa-f]{6}$/.test(rawClr) || rawClr === 'auto') {
        colorVal = rawClr;
      }
    }

    let szVal = inheritedSizeHalfPt ? String(inheritedSizeHalfPt) : '';
    const szEl = rPrNode?.getElementsByTagName('w:sz')[0];
    if (szEl) {
      const rawSz = (szEl.getAttribute('w:val') || '').trim();
      if (/^\d+$/.test(rawSz)) {
        szVal = rawSz;
      }
    }

    let shdFill = '';
    const shdEl = rPrNode?.getElementsByTagName('w:shd')[0];
    if (shdEl) {
      const rawFill = (shdEl.getAttribute('w:fill') || '').trim();
      if (/^[0-9A-Fa-f]{6}$/.test(rawFill)) {
        shdFill = rawFill;
      }
    }

    // Strict CT_RPr element sequence per ECMA-376 wml.xsd: b -> i -> strike -> color -> sz -> szCs -> u -> shd
    const rPrParts: string[] = [];
    if (isBold) rPrParts.push('<w:b/>');
    if (isItalic) rPrParts.push('<w:i/>');
    if (isStrike) rPrParts.push('<w:strike/>');
    if (colorVal) rPrParts.push(`<w:color w:val="${colorVal}"/>`);
    if (szVal) {
      rPrParts.push(`<w:sz w:val="${szVal}"/>`);
      rPrParts.push(`<w:szCs w:val="${szVal}"/>`);
    }
    if (isUnderline) rPrParts.push('<w:u w:val="single"/>');
    if (shdFill) rPrParts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${shdFill}"/>`);

    const runContentParts: string[] = [];
    const childNodes = Array.from(rNode.childNodes);
    for (const c of childNodes) {
      if (c.nodeType !== 1) continue;
      const cEl = c as Element;
      const cTag = cEl.tagName || cEl.nodeName;
      if (cTag === 'w:t') {
        const txt = cEl.textContent || '';
        if (txt.length > 0) {
          runContentParts.push(`<w:t xml:space="preserve">${escapeXml(txt)}</w:t>`);
        }
      } else if (cTag === 'w:br') {
        const brType = cEl.getAttribute('w:type');
        runContentParts.push(brType === 'page' ? '<w:br w:type="page"/>' : '<w:br/>');
      } else if (cTag === 'w:tab') {
        runContentParts.push('<w:tab/>');
      }
    }

    if (runContentParts.length === 0) return '';
    return `<w:r>${rPrParts.length > 0 ? `<w:rPr>${rPrParts.join('')}</w:rPr>` : ''}${runContentParts.join('')}</w:r>`;
  };

  const serializeParagraphNode = (pNode: Element, insideTable = false): string | null => {
    const allText = Array.from(pNode.getElementsByTagName('w:t'))
      .map((t) => t.textContent || '')
      .join('')
      .replace(/\s+/g, ' ')
      .trim();

    if (!insideTable && allText && hfSet.has(allText)) {
      return null;
    }

    const pPrNode = Array.from(pNode.childNodes).find(
      (n) => n.nodeType === 1 && (n as Element).tagName === 'w:pPr'
    ) as Element | undefined;

    const pStyleVal = pPrNode?.getElementsByTagName('w:pStyle')[0]?.getAttribute('w:val') || '';
    const isH1 = /heading\s*1|^h1$|^title$/i.test(pStyleVal);
    const isH2 = /heading\s*2|^h2$|^subtitle$/i.test(pStyleVal);
    const isH3 = /heading\s*[3-6]|^h[3-6]$/i.test(pStyleVal);
    const hasNumPr = Boolean(pPrNode && pPrNode.getElementsByTagName('w:numPr').length > 0);

    const jcVal = pPrNode?.getElementsByTagName('w:jc')[0]?.getAttribute('w:val') || '';
    const validJc = ['left', 'center', 'right', 'both'].includes(jcVal) ? jcVal : '';

    const spacingEl = pPrNode?.getElementsByTagName('w:spacing')[0];
    let spacingBefore = isH1 ? '240' : isH2 ? '200' : isH3 ? '160' : insideTable ? '40' : '0';
    let spacingAfter = isH1 ? '120' : isH2 ? '100' : isH3 ? '80' : insideTable ? '40' : '120';
    if (spacingEl) {
      const bAttr = spacingEl.getAttribute('w:before');
      const aAttr = spacingEl.getAttribute('w:after');
      if (bAttr && /^\d+$/.test(bAttr)) spacingBefore = bAttr;
      if (aAttr && /^\d+$/.test(aAttr)) spacingAfter = aAttr;
    }

    const indEl = pPrNode?.getElementsByTagName('w:ind')[0];
    let leftInd = hasNumPr ? '360' : '';
    if (indEl) {
      const lAttr = indEl.getAttribute('w:left');
      if (lAttr && /^\d+$/.test(lAttr)) leftInd = lAttr;
    }

    const shdEl = pPrNode?.getElementsByTagName('w:shd')[0];
    const pShdFill = shdEl?.getAttribute('w:fill') || '';

    // Strict CT_PPr element sequence per ECMA-376 wml.xsd: shd -> spacing -> ind -> jc
    const pPrXmlParts: string[] = [];
    if (/^[0-9A-Fa-f]{6}$/.test(pShdFill)) {
      pPrXmlParts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${pShdFill}"/>`);
    }
    pPrXmlParts.push(`<w:spacing w:before="${spacingBefore}" w:after="${spacingAfter}"/>`);
    if (leftInd) pPrXmlParts.push(`<w:ind w:left="${leftInd}"/>`);
    if (validJc) pPrXmlParts.push(`<w:jc w:val="${validJc}"/>`);

    const inheritedBold = isH1 || isH2 || isH3;
    const inheritedSize = isH1 ? 32 : isH2 ? 26 : isH3 ? 23 : undefined;
    const inheritedColor = isH1 || isH2 ? '0F172A' : undefined;

    const rElements: Element[] = [];
    const collectRuns = (parent: Element) => {
      for (const ch of Array.from(parent.childNodes)) {
        if (ch.nodeType !== 1) continue;
        const chEl = ch as Element;
        const tag = chEl.tagName || chEl.nodeName;
        if (tag === 'w:r') {
          rElements.push(chEl);
        } else if (tag === 'w:hyperlink' || tag === 'w:smartTag' || tag === 'w:sdtContent' || tag === 'w:ins') {
          collectRuns(chEl);
        }
      }
    };
    collectRuns(pNode);

    const serializedRuns: string[] = [];
    if (hasNumPr && rElements.length > 0) {
      serializedRuns.push(
        `<w:r><w:t xml:space="preserve">•  </w:t></w:r>`
      );
    }

    for (const rEl of rElements) {
      const rXml = serializeRunNode(rEl, inheritedBold, inheritedSize, inheritedColor);
      if (rXml) serializedRuns.push(rXml);
    }

    if (serializedRuns.length === 0 && !insideTable) {
      // Check if paragraph has a page break
      const hasPageBreak = pNode.getElementsByTagName('w:br').length > 0;
      if (!hasPageBreak) return null;
    }

    return `<w:p><w:pPr>${pPrXmlParts.join('')}</w:pPr>${serializedRuns.join('')}</w:p>`;
  };

  const serializeTableNode = (tblNode: Element): string | null => {
    const trNodes = Array.from(tblNode.childNodes).filter(
      (n) => n.nodeType === 1 && (n as Element).tagName === 'w:tr'
    ) as Element[];
    if (trNodes.length === 0) return null;

    const tblGridNode = Array.from(tblNode.childNodes).find(
      (n) => n.nodeType === 1 && (n as Element).tagName === 'w:tblGrid'
    ) as Element | undefined;

    let gridColWidths: number[] = [];
    if (tblGridNode) {
      const gridCols = Array.from(tblGridNode.getElementsByTagName('w:gridCol'));
      gridColWidths = gridCols
        .map((gc) => parseInt(gc.getAttribute('w:w') || '0', 10))
        .filter((w) => w > 0);
    }

    const maxColsInRows = Math.max(
      1,
      ...trNodes.map(
        (tr) =>
          Array.from(tr.childNodes).filter(
            (n) => n.nodeType === 1 && (n as Element).tagName === 'w:tc'
          ).length
      )
    );
    const numCols = Math.max(gridColWidths.length, maxColsInRows, 1);
    const totalTableWidthTwips = 9360;
    const defaultColWidth = Math.floor(totalTableWidthTwips / numCols);

    const finalGridColsXml = Array.from({ length: numCols }, (_, i) => {
      const w = gridColWidths[i] && gridColWidths[i] > 200 ? gridColWidths[i] : defaultColWidth;
      return `      <w:gridCol w:w="${w}"/>`;
    }).join('\n');

    // Inspect table borders from source <w:tblPr>, defaulting to clean single borders
    let borderColor = '94A3B8';
    let borderSize = '6';
    const tblPrNode = Array.from(tblNode.childNodes).find(
      (n) => n.nodeType === 1 && (n as Element).tagName === 'w:tblPr'
    ) as Element | undefined;
    if (tblPrNode) {
      const topBdr = tblPrNode.getElementsByTagName('w:top')[0];
      if (topBdr) {
        const c = topBdr.getAttribute('w:color') || '';
        const sz = topBdr.getAttribute('w:sz') || '';
        if (/^[0-9A-Fa-f]{6}$/.test(c)) borderColor = c;
        if (/^\d+$/.test(sz) && parseInt(sz, 10) > 0) borderSize = sz;
      }
    }

    const serializedRows: string[] = [];
    trNodes.forEach((trNode, rowIdx) => {
      const trPrNode = Array.from(trNode.childNodes).find(
        (n) => n.nodeType === 1 && (n as Element).tagName === 'w:trPr'
      ) as Element | undefined;
      const isHeaderRow =
        Boolean(trPrNode && trPrNode.getElementsByTagName('w:tblHeader').length > 0) ||
        rowIdx === 0;

      const tcNodes = Array.from(trNode.childNodes).filter(
        (n) => n.nodeType === 1 && (n as Element).tagName === 'w:tc'
      ) as Element[];

      if (tcNodes.length === 0) return;

      const serializedCells = tcNodes.map((tcNode, colIdx) => {
        const tcPrNode = Array.from(tcNode.childNodes).find(
          (n) => n.nodeType === 1 && (n as Element).tagName === 'w:tcPr'
        ) as Element | undefined;

        const tcWEl = tcPrNode?.getElementsByTagName('w:tcW')[0];
        const rawW = parseInt(tcWEl?.getAttribute('w:w') || '0', 10);
        const cellWidth =
          rawW > 200
            ? rawW
            : gridColWidths[colIdx] && gridColWidths[colIdx] > 200
            ? gridColWidths[colIdx]
            : defaultColWidth;

        const gridSpanEl = tcPrNode?.getElementsByTagName('w:gridSpan')[0];
        const gridSpanVal = gridSpanEl?.getAttribute('w:val') || '';

        const vMergeEl = tcPrNode?.getElementsByTagName('w:vMerge')[0];
        const vMergeVal = vMergeEl ? vMergeEl.getAttribute('w:val') || 'continue' : '';

        const vAlignEl = tcPrNode?.getElementsByTagName('w:vAlign')[0];
        const vAlignVal = vAlignEl?.getAttribute('w:val') || '';

        const shdEl = tcPrNode?.getElementsByTagName('w:shd')[0];
        let cellFill = shdEl?.getAttribute('w:fill') || '';
        if (!/^[0-9A-Fa-f]{6}$/.test(cellFill) && isHeaderRow) {
          cellFill = 'F1F5F9';
        }

        // Strict CT_TcPr sequence per ECMA-376 wml.xsd: tcW -> gridSpan -> vMerge -> shd -> tcMar (top, left, bottom, right) -> vAlign
        const tcPrParts: string[] = [`<w:tcW w:w="${cellWidth}" w:type="dxa"/>`];
        if (gridSpanVal && /^\d+$/.test(gridSpanVal) && parseInt(gridSpanVal, 10) > 1) {
          tcPrParts.push(`<w:gridSpan w:val="${gridSpanVal}"/>`);
        }
        if (vMergeVal) {
          tcPrParts.push(
            vMergeVal === 'restart' ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>'
          );
        }
        if (/^[0-9A-Fa-f]{6}$/.test(cellFill)) {
          tcPrParts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${cellFill}"/>`);
        }
        tcPrParts.push(`<w:tcMar>
            <w:top w:w="100" w:type="dxa"/>
            <w:left w:w="140" w:type="dxa"/>
            <w:bottom w:w="100" w:type="dxa"/>
            <w:right w:w="140" w:type="dxa"/>
          </w:tcMar>`);
        if (['top', 'center', 'bottom'].includes(vAlignVal)) {
          tcPrParts.push(`<w:vAlign w:val="${vAlignVal}"/>`);
        }

        const cellParas = Array.from(tcNode.childNodes).filter(
          (n) => n.nodeType === 1 && (n as Element).tagName === 'w:p'
        ) as Element[];

        const cellParaXmls = cellParas
          .map((cp) => serializeParagraphNode(cp, true))
          .filter((x): x is string => Boolean(x));

        // OpenXML schema strictly requires every <w:tc> to contain at least one <w:p>
        if (cellParaXmls.length === 0) {
          cellParaXmls.push('<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr></w:p>');
        }

        return `      <w:tc>
        <w:tcPr>${tcPrParts.join('')}</w:tcPr>
        ${cellParaXmls.join('\n        ')}
      </w:tc>`;
      });

      serializedRows.push(`    <w:tr>
      ${isHeaderRow ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}
${serializedCells.join('\n')}
    </w:tr>`);
    });

    if (serializedRows.length === 0) return null;

    return `<w:tbl>
    <w:tblPr>
      <w:tblW w:w="${totalTableWidthTwips}" w:type="dxa"/>
      <w:tblBorders>
        <w:top w:val="single" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>
        <w:left w:val="single" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>
        <w:bottom w:val="single" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>
        <w:right w:val="single" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>
        <w:insideH w:val="single" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>
        <w:insideV w:val="single" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>
      </w:tblBorders>
    </w:tblPr>
    <w:tblGrid>
${finalGridColsXml}
    </w:tblGrid>
${serializedRows.join('\n')}
  </w:tbl>`;
  };

  const resultParts: string[] = [];
  for (const child of Array.from(bodyNode.childNodes)) {
    if (child.nodeType !== 1) continue;
    const el = child as Element;
    const tag = el.tagName || el.nodeName;
    if (tag === 'w:p') {
      const pXml = serializeParagraphNode(el, false);
      if (pXml) resultParts.push(pXml);
    } else if (tag === 'w:tbl') {
      const tblXml = serializeTableNode(el);
      if (tblXml) resultParts.push(tblXml);
    }
  }

  return resultParts;
}

function appendStructuredBlocksToDocx(blocks: StructuredBlock[], bodyXmlParts: string[]) {
  let idx = 0;
  while (idx < blocks.length) {
    const block = blocks[idx];
    if (block.type === 'table-row') {
      const tableRows: StructuredBlock[] = [];
      while (idx < blocks.length && blocks[idx].type === 'table-row') {
        tableRows.push(blocks[idx]);
        idx++;
      }
      bodyXmlParts.push(createDocxTable(tableRows));
      continue;
    }

    if (block.type === 'h1') {
      bodyXmlParts.push(
        createDocxParagraph(block.text, {
          bold: true,
          sizeHalfPt: 32,
          color: '0F172A',
          spacingBefore: 240,
          spacingAfter: 120,
        })
      );
    } else if (block.type === 'h2') {
      bodyXmlParts.push(
        createDocxParagraph(block.text, {
          bold: true,
          sizeHalfPt: 26,
          color: '0F172A',
          spacingBefore: 200,
          spacingAfter: 100,
        })
      );
    } else if (block.type === 'h3') {
      bodyXmlParts.push(
        createDocxParagraph(block.text, {
          bold: true,
          sizeHalfPt: 23,
          color: '1E293B',
          spacingBefore: 160,
          spacingAfter: 80,
        })
      );
    } else if (block.type === 'li') {
      bodyXmlParts.push(
        createDocxParagraph(`•  ${block.text}`, {
          sizeHalfPt: 22,
          color: '1E293B',
          leftIndentTwips: 360,
          spacingAfter: 80,
        })
      );
    } else {
      bodyXmlParts.push(
        createDocxParagraph(block.text, {
          sizeHalfPt: 22,
          color: '1E293B',
          spacingAfter: 140,
        })
      );
    }
    idx++;
  }
}

function createDocxTable(rows: StructuredBlock[]): string {
  const numCols = Math.max(
    1,
    ...rows.map((r) =>
      r.cells && r.cells.length > 0 ? r.cells.length : r.text.split(/\s*\|\s*/).length
    )
  );
  const totalWidthTwips = 9360;
  const colWidthTwips = Math.floor(totalWidthTwips / numCols);

  const gridCols = Array.from(
    { length: numCols },
    () => `      <w:gridCol w:w="${colWidthTwips}"/>`
  ).join('\n');

  const trXml = rows
    .map((row, rIdx) => {
      const rawCells =
        row.cells && row.cells.length > 0 ? row.cells : row.text.split(/\s*\|\s*/);
      const isHeader = Boolean(row.isHeaderRow || rIdx === 0);

      const tcXml = Array.from({ length: numCols }, (_, cIdx) => {
        const cellText = rawCells[cIdx] ?? '';
        return `      <w:tc>
        <w:tcPr>
          <w:tcW w:w="${colWidthTwips}" w:type="dxa"/>
          ${isHeader ? '<w:shd w:val="clear" w:color="auto" w:fill="F1F5F9"/>' : ''}
          <w:tcMar>
            <w:top w:w="100" w:type="dxa"/>
            <w:left w:w="140" w:type="dxa"/>
            <w:bottom w:w="100" w:type="dxa"/>
            <w:right w:w="140" w:type="dxa"/>
          </w:tcMar>
        </w:tcPr>
        <w:p>
          <w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>
          <w:r>
            <w:rPr>
              ${isHeader ? '<w:b/>' : ''}
              <w:color w:val="${isHeader ? '0F172A' : '1E293B'}"/>
              <w:sz w:val="20"/>
              <w:szCs w:val="20"/>
            </w:rPr>
            <w:t xml:space="preserve">${escapeXml(cellText)}</w:t>
          </w:r>
        </w:p>
      </w:tc>`;
      }).join('\n');

      return `    <w:tr>
      ${isHeader ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}
${tcXml}
    </w:tr>`;
    })
    .join('\n');

  return `<w:tbl>
    <w:tblPr>
      <w:tblW w:w="${totalWidthTwips}" w:type="dxa"/>
      <w:tblBorders>
        <w:top w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>
        <w:left w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>
        <w:bottom w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>
        <w:right w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>
        <w:insideH w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>
        <w:insideV w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>
      </w:tblBorders>
    </w:tblPr>
    <w:tblGrid>
${gridCols}
    </w:tblGrid>
${trXml}
  </w:tbl>`;
}

/**
 * Builds a 100% ECMA-376 / MS PowerPoint Desktop compliant OpenXML Presentation (.pptx)
 * preserving native slide XML, backgrounds (<p:bg>), embedded images, and tables (<a:tbl>)
 * with configurable aspect ratios (16:9, 4:3, 16:10).
 */
async function buildUnifiedPptx(
  activeDocs: QueuedDocument[],
  config: MergeConfiguration,
  baseFilename: string,
  startTime: number,
  onMergeProgress?: (progress: Omit<MergeProgressState, 'phase'>) => void
): Promise<MergedArtifact> {
  onMergeProgress?.({
    percent: 10,
    processedDocs: 0,
    totalDocs: activeDocs.length,
    currentDocName: activeDocs[0]?.name || '',
    stepDetail: 'Initializing OpenXML PowerPoint (.pptx) presentation...',
  });
  await yieldToUi(20);

  const zip = new JSZip();
  const manifest: MergedArtifact['manifest'] = [];
  let totalWords = 0;

  // Aspect ratio geometry in EMUs (1 inch = 914400 EMUs)
  let sldCx = 9144000;
  let sldCy = 5143500;
  let sldType = 'screen16x9';

  if (config.slideRatio === '4:3') {
    sldCx = 9144000;
    sldCy = 6858000;
    sldType = 'screen4x3';
  } else if (config.slideRatio === '16:10') {
    sldCx = 9144000;
    sldCy = 5715000;
    sldType = 'screen16x10';
  }

  const marginX = 548640; // 0.6 in
  const boxWidth = sldCx - marginX * 2;
  const titleY = 457200; // 0.5 in
  const titleH = 914400; // 1.0 in
  const bodyY = 1524000; // 1.67 in
  const bodyH = sldCy - bodyY - 457200;

  interface CompiledSlideEntry {
    rawSlideXml?: string;
    extraRels?: Array<{ id: string; type: string; target: string }>;
    title: string;
    blocks: StructuredBlock[];
    bgColorHex?: string;
    textColorHex?: string;
  }

  const compiledSlides: CompiledSlideEntry[] = [];
  const mediaExtensions = new Set<string>();
  let mediaCounter = 1;

  // Check if any active document is a PPTX so we can preserve its theme1.xml if available
  let customThemeXml: string | null = null;

  for (let i = 0; i < activeDocs.length; i++) {
    const doc = activeDocs[i];
    const stepPercent = Math.round(15 + (i / activeDocs.length) * 68);
    onMergeProgress?.({
      percent: stepPercent,
      processedDocs: i,
      totalDocs: activeDocs.length,
      currentDocName: doc.name,
      stepDetail: `Merging presentation ${i + 1} of ${activeDocs.length}: ${doc.name}...`,
    });
    await yieldToUi(24);

    totalWords += doc.wordCount;
    const startSlide = compiledSlides.length + 1;

    if (doc.extension === 'pptx' && doc.slides && doc.slides.length > 0) {
      const selectedIndices = parsePageRange(doc.pageRangeInput, doc.slides.length);
      let srcZip: JSZip | null = null;
      const slideFilePaths: string[] = [];

      try {
        srcZip = await JSZip.loadAsync(doc.arrayBuffer);
        if (!customThemeXml) {
          const themeFile = srcZip.file('ppt/theme/theme1.xml');
          if (themeFile) {
            customThemeXml = await themeFile.async('text');
          }
        }
        const entries: Array<{ num: number; path: string }> = [];
        srcZip.forEach((relPath) => {
          const m = relPath.match(/^ppt\/slides\/slide(\d+)\.xml$/i);
          if (m) {
            entries.push({ num: parseInt(m[1], 10), path: relPath });
          }
        });
        entries.sort((a, b) => a.num - b.num);
        entries.forEach((e) => slideFilePaths.push(e.path));
      } catch {
        srcZip = null;
      }

      for (const sIdx of selectedIndices) {
        const s = doc.slides[sIdx];
        if (!s) continue;

        let rawSlideXml: string | undefined;
        const extraRels: Array<{ id: string; type: string; target: string }> = [];

        if (srcZip && slideFilePaths[sIdx]) {
          try {
            const slidePath = slideFilePaths[sIdx];
            const slideNumMatch = slidePath.match(/slide(\d+)\.xml$/i);
            const origSlideNum = slideNumMatch ? slideNumMatch[1] : String(sIdx + 1);
            const rawXml = await srcZip.file(slidePath)?.async('text');

            if (rawXml) {
              let modifiedXml = rawXml;
              // If the slide inherited its background from its slideLayout or slideMaster and doesn't have an explicit <p:bg> on <p:cSld>, inject it so the background color is preserved
              if (
                s.bgColorHex &&
                s.bgColorHex.length === 6 &&
                !/<p:bg[\s>]/i.test(modifiedXml)
              ) {
                modifiedXml = modifiedXml.replace(
                  /(<p:cSld[^>]*>)/i,
                  `$1<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${s.bgColorHex}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`
                );
              }

              // Copy any referenced image media from ppt/slides/_rels/slideX.xml.rels
              const relsPath = `ppt/slides/_rels/slide${origSlideNum}.xml.rels`;
              const relsFile = srcZip.file(relsPath);
              if (relsFile) {
                const relsText = await relsFile.async('text');
                const parser = new DOMParser();
                const relsDoc = parser.parseFromString(relsText, 'application/xml');
                const relNodes = Array.from(relsDoc.getElementsByTagName('Relationship'));
                for (const relNode of relNodes) {
                  const origId = relNode.getAttribute('Id') || '';
                  const relType = relNode.getAttribute('Type') || '';
                  const target = relNode.getAttribute('Target') || '';
                  if (relType.endsWith('/image') && target) {
                    const cleanTarget = target.replace(/^\.\.\//, 'ppt/');
                    const mediaFile = srcZip.file(cleanTarget);
                    if (mediaFile) {
                      const extMatch = cleanTarget.match(/\.([a-zA-Z0-9]+)$/);
                      const ext = (extMatch ? extMatch[1] : 'png').toLowerCase();
                      const newMediaName = `media_${mediaCounter++}.${ext}`;
                      const mediaBytes = await mediaFile.async('uint8array');
                      zip.file(`ppt/media/${newMediaName}`, mediaBytes);
                      mediaExtensions.add(ext);
                      const newRelId = `rIdImg${mediaCounter}`;
                      extraRels.push({
                        id: newRelId,
                        type: relType,
                        target: `../media/${newMediaName}`,
                      });
                      // Replace r:embed="origId" with newRelId in slide XML
                      modifiedXml = modifiedXml
                        .split(`r:embed="${origId}"`)
                        .join(`r:embed="${newRelId}"`);
                    }
                  }
                }
              }

              rawSlideXml = modifiedXml;
            }
          } catch {
            rawSlideXml = undefined;
          }
        }

        compiledSlides.push({
          rawSlideXml,
          extraRels,
          title: s.title || s.blocks[0]?.text || '',
          blocks: s.blocks,
          bgColorHex: s.bgColorHex,
          textColorHex: s.textColorHex,
        });
      }
    } else {
      const blocks = doc.blocks.length > 0 ? doc.blocks : [];
      if (blocks.length === 0) {
        compiledSlides.push({ title: '', blocks: [] });
      } else {
        for (let b = 0; b < blocks.length; b += 6) {
          const slice = blocks.slice(b, b + 6);
          const first = slice[0];
          compiledSlides.push({
            title: first.type !== 'table-row' ? first.text : '',
            blocks: slice,
          });
        }
      }
    }

    const endSlide = Math.max(startSlide, compiledSlides.length);
    manifest.push({
      docId: doc.id,
      name: doc.name,
      extension: doc.extension,
      pagesIncluded: Math.max(1, endSlide - startSlide + 1),
      startPage: startSlide,
      endPage: endSlide,
    });
  }

  if (compiledSlides.length === 0) {
    compiledSlides.push({ title: '', blocks: [] });
  }

  onMergeProgress?.({
    percent: 88,
    processedDocs: activeDocs.length,
    totalDocs: activeDocs.length,
    currentDocName: `${baseFilename}.pptx`,
    stepDetail: 'Packaging ECMA-376 OpenXML presentation & compressing .pptx archive...',
  });
  await yieldToUi(20);

  const mediaDefaultTypes = Array.from(mediaExtensions)
    .map((ext) => {
      const mime =
        ext === 'jpg' || ext === 'jpeg'
          ? 'image/jpeg'
          : ext === 'gif'
          ? 'image/gif'
          : ext === 'svg'
          ? 'image/svg+xml'
          : 'image/png';
      return `  <Default Extension="${ext}" ContentType="${mime}"/>`;
    })
    .join('\n');

  const slideOverrides = compiledSlides
    .map(
      (_, idx) =>
        `  <Override PartName="/ppt/slides/slide${
          idx + 1
        }.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
    )
    .join('\n');

  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
${mediaDefaultTypes ? `${mediaDefaultTypes}\n` : ''}  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>
  <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>
  <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
${slideOverrides}
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`
  );

  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`
  );

  const sldIdLstXml = compiledSlides
    .map((_, idx) => `    <p:sldId id="${256 + idx}" r:id="rId${idx + 6}"/>`)
    .join('\n');

  zip.file(
    'ppt/presentation.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
${sldIdLstXml}
  </p:sldIdLst>
  <p:sldSz cx="${sldCx}" cy="${sldCy}" type="${sldType}"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:defaultTextStyle>
    <a:defPPr>
      <a:defRPr lang="en-US"/>
    </a:defPPr>
    <a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
      <a:defRPr sz="1800" kern="1200">
        <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>
        <a:latin typeface="+mn-lt"/>
        <a:ea typeface="+mn-ea"/>
        <a:cs typeface="+mn-cs"/>
      </a:defRPr>
    </a:lvl1pPr>
  </p:defaultTextStyle>
</p:presentation>`
  );

  const presSlideRelsXml = compiledSlides
    .map(
      (_, idx) =>
        `  <Relationship Id="rId${
          idx + 6
        }" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${
          idx + 1
        }.xml"/>`
    )
    .join('\n');

  zip.file(
    'ppt/_rels/presentation.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>
${presSlideRelsXml}
</Relationships>`
  );

  zip.file(
    'ppt/presProps.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`
  );

  zip.file(
    'ppt/viewProps.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:normalViewPr>
    <p:restoredLeft sz="15620"/>
    <p:restoredTop sz="94660"/>
  </p:normalViewPr>
</p:viewPr>`
  );

  zip.file(
    'ppt/tableStyles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`
  );

  zip.file(
    'ppt/theme/theme1.xml',
    customThemeXml ||
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
        <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
        <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>`
  );

  zip.file(
    'ppt/slideMasters/slideMaster1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgRef idx="1001">
        <a:schemeClr val="bg1"/>
      </p:bgRef>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
  <p:txStyles>
    <p:titleStyle>
      <a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
        <a:defRPr sz="4400" kern="1200">
          <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>
          <a:latin typeface="+mj-lt"/>
          <a:ea typeface="+mj-ea"/>
          <a:cs typeface="+mj-cs"/>
        </a:defRPr>
      </a:lvl1pPr>
    </p:titleStyle>
    <p:bodyStyle>
      <a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
        <a:defRPr sz="2400" kern="1200">
          <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>
          <a:latin typeface="+mn-lt"/>
          <a:ea typeface="+mn-ea"/>
          <a:cs typeface="+mn-cs"/>
        </a:defRPr>
      </a:lvl1pPr>
    </p:bodyStyle>
    <p:otherStyle>
      <a:defPPr>
        <a:defRPr lang="en-US"/>
      </a:defPPr>
    </p:otherStyle>
  </p:txStyles>
</p:sldMaster>`
  );

  zip.file(
    'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`
  );

  zip.file(
    'ppt/slideLayouts/slideLayout1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank">
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sldLayout>`
  );

  zip.file(
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`
  );

  for (let idx = 0; idx < compiledSlides.length; idx++) {
    const slideData = compiledSlides[idx];
    if (slideData.rawSlideXml) {
      zip.file(`ppt/slides/slide${idx + 1}.xml`, slideData.rawSlideXml);
      const extraRelsXml = (slideData.extraRels || [])
        .map((r) => `  <Relationship Id="${r.id}" Type="${r.type}" Target="${r.target}"/>`)
        .join('\n');
      zip.file(
        `ppt/slides/_rels/slide${idx + 1}.xml.rels`,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
${extraRelsXml ? `${extraRelsXml}\n` : ''}</Relationships>`
      );
      continue;
    }

    const fgHex = slideData.textColorHex || '1E293B';
    const titleHex = slideData.textColorHex || '0F172A';
    const bgXml =
      slideData.bgColorHex && slideData.bgColorHex.length === 6
        ? `    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="${slideData.bgColorHex}"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>\n`
        : '';

    const nonTableBlocks = slideData.blocks.filter(
      (b, bIdx) => b.type !== 'table-row' && !(bIdx === 0 && b.text === slideData.title)
    );
    const tableBlocks = slideData.blocks.filter((b) => b.type === 'table-row');

    const bodyParasXml =
      nonTableBlocks.length > 0
        ? nonTableBlocks
            .map((b) => {
              const line = b.type === 'li' ? `• ${b.text}` : b.text;
              return `<a:p>
            <a:r>
              <a:rPr lang="en-US" sz="1800" dirty="0">
                <a:solidFill><a:srgbClr val="${fgHex}"/></a:solidFill>
              </a:rPr>
              <a:t>${escapeXml(line)}</a:t>
            </a:r>
          </a:p>`;
            })
            .join('\n          ')
        : `<a:p><a:endParaRPr lang="en-US" sz="1800" dirty="0"/></a:p>`;

    let tableFrameXml = '';
    if (tableBlocks.length > 0) {
      const numCols = Math.max(
        1,
        ...tableBlocks.map((r) =>
          r.cells && r.cells.length > 0 ? r.cells.length : r.text.split(/\s*\|\s*/).length
        )
      );
      const colW = Math.floor(boxWidth / numCols);
      const gridColsXml = Array.from(
        { length: numCols },
        () => `              <a:gridCol w="${colW}"/>`
      ).join('\n');
      const rowsXml = tableBlocks
        .map((r, rIdx) => {
          const cells =
            r.cells && r.cells.length > 0 ? r.cells : r.text.split(/\s*\|\s*/);
          const isHeader = Boolean(r.isHeaderRow || rIdx === 0);
          const tcXml = Array.from({ length: numCols }, (_, cIdx) => {
            const cellVal = cells[cIdx] ?? '';
            return `              <a:tc>
                <a:txBody>
                  <a:bodyPr/>
                  <a:lstStyle/>
                  <a:p>
                    <a:r>
                      <a:rPr lang="en-US" sz="1400" ${isHeader ? 'b="1"' : ''} dirty="0">
                        <a:solidFill><a:srgbClr val="${fgHex}"/></a:solidFill>
                      </a:rPr>
                      <a:t>${escapeXml(cellVal)}</a:t>
                    </a:r>
                  </a:p>
                </a:txBody>
                <a:tcPr/>
              </a:tc>`;
          }).join('\n');
          return `            <a:tr h="370840">\n${tcXml}\n            </a:tr>`;
        })
        .join('\n');

      const tblY = nonTableBlocks.length > 0 ? bodyY + Math.floor(bodyH * 0.45) : bodyY;
      const tblH = Math.min(bodyH - (tblY - bodyY), tableBlocks.length * 380000);

      tableFrameXml = `
      <p:graphicFrame>
        <p:nvGraphicFramePr>
          <p:cNvPr id="4" name="SlideTable"/>
          <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
          <p:nvPr/>
        </p:nvGraphicFramePr>
        <p:xfrm>
          <a:off x="${marginX}" y="${tblY}"/>
          <a:ext cx="${boxWidth}" cy="${tblH}"/>
        </p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
            <a:tbl>
              <a:tblPr firstRow="1" bandRow="1"/>
              <a:tblGrid>
${gridColsXml}
              </a:tblGrid>
${rowsXml}
            </a:tbl>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>`;
    }

    zip.file(
      `ppt/slides/slide${idx + 1}.xml`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
${bgXml}    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="TitleTextBox"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="${marginX}" y="${titleY}"/>
            <a:ext cx="${boxWidth}" cy="${titleH}"/>
          </a:xfrm>
          <a:prstGeom prst="rect">
            <a:avLst/>
          </a:prstGeom>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" rtlCol="0"/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="2800" b="1" dirty="0">
                <a:solidFill><a:srgbClr val="${titleHex}"/></a:solidFill>
              </a:rPr>
              <a:t>${escapeXml(slideData.title)}</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="ContentTextBox"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="${marginX}" y="${bodyY}"/>
            <a:ext cx="${boxWidth}" cy="${bodyH}"/>
          </a:xfrm>
          <a:prstGeom prst="rect">
            <a:avLst/>
          </a:prstGeom>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" rtlCol="0"/>
          <a:lstStyle/>
          ${bodyParasXml}
        </p:txBody>
      </p:sp>${tableFrameXml}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>`
    );

    zip.file(
      `ppt/slides/_rels/slide${idx + 1}.xml.rels`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`
    );
  }

  const totalSlides = compiledSlides.length;

  zip.file(
    'docProps/app.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office PowerPoint</Application>
  <PresentationFormat>Widescreen</PresentationFormat>
  <Slides>${totalSlides}</Slides>
  <Notes>0</Notes>
  <HiddenSlides>0</HiddenSlides>
  <MMClips>0</MMClips>
  <ScaleCrop>false</ScaleCrop>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>`
  );

  zip.file(
    'docProps/core.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${escapeXml(baseFilename)}</dc:title>
  <dc:creator>FolioBind</dc:creator>
  <cp:lastModifiedBy>FolioBind</cp:lastModifiedBy>
  <cp:revision>1</cp:revision>
</cp:coreProperties>`
  );

  const blob = await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });
  const blobUrl = URL.createObjectURL(blob);
  const durationMs = Math.max(14, Math.round(performance.now() - startTime));

  onMergeProgress?.({
    percent: 100,
    processedDocs: activeDocs.length,
    totalDocs: activeDocs.length,
    currentDocName: `${baseFilename}.pptx`,
    stepDetail: `Merged ${activeDocs.length} files (${totalSlides} slides · ${config.slideRatio}) into ${baseFilename}.pptx`,
  });

  return {
    id: `merged_${Date.now()}`,
    filename: `${baseFilename}.pptx`,
    format: 'pptx',
    blob,
    blobUrl,
    sizeBytes: blob.size,
    totalPages: totalSlides,
    totalWords,
    sourceCount: activeDocs.length,
    mergedAt: new Date().toTimeString().slice(0, 8),
    durationMs,
    manifest,
  };
}

function createDocxParagraph(
  text: string,
  opts: {
    bold?: boolean;
    sizeHalfPt?: number;
    color?: string;
    leftIndentTwips?: number;
    spacingBefore?: number;
    spacingAfter?: number;
  } = {}
): string {
  const {
    bold = false,
    sizeHalfPt = 22,
    color = '1E293B',
    leftIndentTwips = 0,
    spacingBefore = 0,
    spacingAfter = 120,
  } = opts;

  return `<w:p>
      <w:pPr>
        <w:spacing w:before="${spacingBefore}" w:after="${spacingAfter}"/>
        ${leftIndentTwips > 0 ? `<w:ind w:left="${leftIndentTwips}"/>` : ''}
      </w:pPr>
      <w:r>
        <w:rPr>
          ${bold ? '<w:b/>' : ''}
          <w:color w:val="${color}"/>
          <w:sz w:val="${sizeHalfPt}"/>
          <w:szCs w:val="${sizeHalfPt}"/>
        </w:rPr>
        <w:t xml:space="preserve">${escapeXml(text)}</w:t>
      </w:r>
    </w:p>`;
}

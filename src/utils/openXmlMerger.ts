import JSZip from 'jszip';
import { StructuredBlock } from '../types/document';

function escapeXml(unsafe: string): string {
  return unsafe
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Extracts a map of theme color names (dk1, lt1, accent1..6, bg1, tx1, etc.) -> 6-char hex RGB
 * from a PPTX archive's ppt/theme/theme1.xml and ppt/slideMasters/slideMaster1.xml
 */
export async function extractPptxThemeColorMap(zip: JSZip): Promise<Record<string, string>> {
  const map: Record<string, string> = {
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
    hlink: '0000FF',
    folHlink: '800080',
    bg1: 'FFFFFF',
    tx1: '000000',
    bg2: 'EEECE1',
    tx2: '1F497D',
  };

  try {
    const themeFile = zip.file('ppt/theme/theme1.xml');
    if (themeFile) {
      const xmlText = await themeFile.async('text');
      const parser = new DOMParser();
      const doc = parser.parseFromString(xmlText, 'application/xml');
      const clrScheme = doc.getElementsByTagName('a:clrScheme')[0];
      if (clrScheme) {
        const keys = [
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
          'hlink',
          'folHlink',
        ];
        for (const k of keys) {
          const node = clrScheme.getElementsByTagName(`a:${k}`)[0];
          if (node) {
            const srgb = node.getElementsByTagName('a:srgbClr')[0];
            const sys = node.getElementsByTagName('a:sysClr')[0];
            const val =
              srgb?.getAttribute('val') ||
              sys?.getAttribute('lastClr') ||
              '';
            if (/^[0-9A-Fa-f]{6}$/.test(val)) {
              map[k] = val.toUpperCase();
            }
          }
        }
      }
    }

    const masterFile = zip.file('ppt/slideMasters/slideMaster1.xml');
    if (masterFile) {
      const masterXml = await masterFile.async('text');
      const parser = new DOMParser();
      const doc = parser.parseFromString(masterXml, 'application/xml');
      const clrMap = doc.getElementsByTagName('p:clrMap')[0];
      if (clrMap) {
        for (const alias of ['bg1', 'tx1', 'bg2', 'tx2']) {
          const target = clrMap.getAttribute(alias);
          if (target && map[target]) {
            map[alias] = map[target];
          }
        }
      } else {
        map.bg1 = map.lt1;
        map.tx1 = map.dk1;
        map.bg2 = map.lt2;
        map.tx2 = map.dk2;
      }
    } else {
      map.bg1 = map.lt1;
      map.tx1 = map.dk1;
      map.bg2 = map.lt2;
      map.tx2 = map.dk2;
    }
  } catch {
    // Ignore theme parse errors
  }

  return map;
}

/**
 * Extracts the resolved background hex color for a given slide XML (checking slide, layout, then master)
 */
export async function resolveSlideBackgroundHex(
  zip: JSZip,
  slidePath: string,
  themeMap: Record<string, string>
): Promise<{ bgHex: string; textHex: string }> {
  const parser = new DOMParser();

  const findBgHexInXml = (xmlText: string): string | null => {
    const doc = parser.parseFromString(xmlText, 'application/xml');
    const bgNodes = doc.getElementsByTagName('p:bg');
    if (bgNodes.length === 0) return null;
    const bg = bgNodes[0];

    // Direct srgbClr inside p:bg
    const srgb = bg.getElementsByTagName('a:srgbClr')[0];
    if (srgb) {
      const val = srgb.getAttribute('val') || '';
      if (/^[0-9A-Fa-f]{6}$/.test(val)) return val.toUpperCase();
    }

    // schemeClr inside p:bg
    const scheme = bg.getElementsByTagName('a:schemeClr')[0];
    if (scheme) {
      const key = scheme.getAttribute('val') || '';
      if (themeMap[key] && key !== 'phClr') return themeMap[key];
    }

    return null;
  };

  try {
    const slideFile = zip.file(slidePath);
    if (slideFile) {
      const slideXml = await slideFile.async('text');
      const direct = findBgHexInXml(slideXml);
      if (direct) {
        return { bgHex: direct, textHex: isDarkHex(direct) ? 'F8FAFC' : '0F172A' };
      }
    }

    // Check slide layout via .rels
    const slideFileName = slidePath.split('/').pop() || '';
    const relsPath = `ppt/slides/_rels/${slideFileName}.rels`;
    const relsFile = zip.file(relsPath);
    if (relsFile) {
      const relsXml = await relsFile.async('text');
      const layoutMatch = relsXml.match(/Target="\.\.\/(slideLayouts\/slideLayout\d+\.xml)"/i);
      if (layoutMatch) {
        const layoutFile = zip.file(`ppt/${layoutMatch[1]}`);
        if (layoutFile) {
          const layoutXml = await layoutFile.async('text');
          const fromLayout = findBgHexInXml(layoutXml);
          if (fromLayout) {
            return { bgHex: fromLayout, textHex: isDarkHex(fromLayout) ? 'F8FAFC' : '0F172A' };
          }
        }
      }
    }

    // Check slideMaster1.xml
    const masterFile = zip.file('ppt/slideMasters/slideMaster1.xml');
    if (masterFile) {
      const masterXml = await masterFile.async('text');
      const fromMaster = findBgHexInXml(masterXml);
      if (fromMaster) {
        return { bgHex: fromMaster, textHex: isDarkHex(fromMaster) ? 'F8FAFC' : '0F172A' };
      }
    }
  } catch {
    // Fallback
  }

  const fallbackBg = themeMap.bg1 || 'FFFFFF';
  return {
    bgHex: fallbackBg,
    textHex: isDarkHex(fallbackBg) ? 'F8FAFC' : '0F172A',
  };
}

export function isDarkHex(hex: string): boolean {
  const clean = hex.replace('#', '');
  if (clean.length !== 6) return false;
  const r = parseInt(clean.slice(0, 2), 16) / 255;
  const g = parseInt(clean.slice(2, 4), 16) / 255;
  const b = parseInt(clean.slice(4, 6), 16) / 255;
  const lum = 0.2126 * r + 0.7152 * g + 0.0722 * b;
  return lum < 0.45;
}

/**
 * Extracts raw inner XML children of <w:body> (excluding <w:sectPr>) from a .docx arrayBuffer,
 * preserving all native <w:tbl> tables, borders, shadings, <w:p> paragraphs, and runs.
 */
export async function extractDocxBodyXmlAndMedia(
  arrayBuffer: ArrayBuffer,
  docIndex: number,
  targetZip: JSZip
): Promise<{
  bodyChildrenXml: string;
  extraRelEntries: string[];
  extraContentTypeDefaults: Array<{ ext: string; contentType: string }>;
  styleNodesXml: string[];
}> {
  const srcZip = await JSZip.loadAsync(arrayBuffer);
  const docXmlFile = srcZip.file('word/document.xml');
  if (!docXmlFile) {
    return {
      bodyChildrenXml: '',
      extraRelEntries: [],
      extraContentTypeDefaults: [],
      styleNodesXml: [],
    };
  }

  let docXml = await docXmlFile.async('text');
  const extraRelEntries: string[] = [];
  const extraContentTypeDefaults: Array<{ ext: string; contentType: string }> = [];
  const styleNodesXml: string[] = [];

  // Copy styles from word/styles.xml so table styles and heading styles are preserved
  const stylesFile = srcZip.file('word/styles.xml');
  if (stylesFile) {
    const stylesXml = await stylesFile.async('text');
    const styleMatches = stylesXml.match(/<w:style\b[\s\S]*?<\/w:style>/g);
    if (styleMatches) {
      styleNodesXml.push(...styleMatches);
    }
  }

  // Remap media/image relationships in word/_rels/document.xml.rels
  const relsFile = srcZip.file('word/_rels/document.xml.rels');
  if (relsFile) {
    const relsXml = await relsFile.async('text');
    const parser = new DOMParser();
    const relsDoc = parser.parseFromString(relsXml, 'application/xml');
    const relNodes = Array.from(relsDoc.getElementsByTagName('Relationship'));

    for (const rel of relNodes) {
      const oldId = rel.getAttribute('Id') || '';
      const type = rel.getAttribute('Type') || '';
      const target = rel.getAttribute('Target') || '';
      const targetMode = rel.getAttribute('TargetMode') || '';

      if (!oldId || !target) continue;

      if (type.endsWith('/image') && !targetMode) {
        const cleanTarget = target.replace(/^\/?word\//, '').replace(/^\.\.\//, '');
        const mediaFile = srcZip.file(`word/${cleanTarget}`);
        if (mediaFile) {
          const baseName = cleanTarget.split('/').pop() || 'img.png';
          const ext = (baseName.split('.').pop() || 'png').toLowerCase();
          const newMediaName = `media/doc${docIndex}_${oldId}_${baseName}`;
          const newId = `rId_d${docIndex}_${oldId}`;

          const imgBytes = await mediaFile.async('uint8array');
          targetZip.file(`word/${newMediaName}`, imgBytes);

          extraRelEntries.push(
            `  <Relationship Id="${newId}" Type="${type}" Target="${newMediaName}"/>`
          );

          const mimeMap: Record<string, string> = {
            png: 'image/png',
            jpg: 'image/jpeg',
            jpeg: 'image/jpeg',
            gif: 'image/gif',
            svg: 'image/svg+xml',
            bmp: 'image/bmp',
            emf: 'image/x-emf',
            wmf: 'image/x-wmf',
          };
          if (mimeMap[ext]) {
            extraContentTypeDefaults.push({ ext, contentType: mimeMap[ext] });
          }

          // Replace r:embed="oldId" and r:id="oldId" in docXml
          docXml = docXml
            .replace(new RegExp(`r:embed="${oldId}"`, 'g'), `r:embed="${newId}"`)
            .replace(new RegExp(`r:id="${oldId}"`, 'g'), `r:id="${newId}"`);
        }
      } else if (targetMode === 'External') {
        const newId = `rId_d${docIndex}_${oldId}`;
        extraRelEntries.push(
          `  <Relationship Id="${newId}" Type="${type}" Target="${escapeXml(
            target
          )}" TargetMode="External"/>`
        );
        docXml = docXml.replace(new RegExp(`r:id="${oldId}"`, 'g'), `r:id="${newId}"`);
      }
    }
  }

  // Extract inner content of <w:body>...</w:body>
  const bodyMatch = docXml.match(/<w:body[^>]*>([\s\S]*?)<\/w:body>/);
  if (!bodyMatch) {
    return {
      bodyChildrenXml: '',
      extraRelEntries,
      extraContentTypeDefaults,
      styleNodesXml,
    };
  }

  let bodyInner = bodyMatch[1];
  // Strip any <w:sectPr> blocks so the caller can insert clean section breaks with headers/footers
  bodyInner = bodyInner.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g, '');

  return {
    bodyChildrenXml: bodyInner.trim(),
    extraRelEntries,
    extraContentTypeDefaults,
    styleNodesXml,
  };
}

/**
 * Converts StructuredBlock[] (including table-row blocks with cells) into OpenXML Word (<w:p> and <w:tbl>)
 * for non-DOCX sources (PDF / PPTX) when exporting to .docx
 */
export function blocksToDocxXml(blocks: StructuredBlock[]): string {
  const parts: string[] = [];
  let i = 0;

  while (i < blocks.length) {
    const b = blocks[i];

    if (b.type === 'table-row') {
      const tableRows: StructuredBlock[] = [];
      while (i < blocks.length && blocks[i].type === 'table-row') {
        tableRows.push(blocks[i]);
        i++;
      }
      parts.push(buildDocxTableXml(tableRows));
      continue;
    }

    if (b.type === 'h1') {
      parts.push(
        buildSimpleDocxPara(b.text, { bold: true, sizeHalfPt: 32, color: '0F172A', spaceAfter: 120 })
      );
    } else if (b.type === 'h2') {
      parts.push(
        buildSimpleDocxPara(b.text, { bold: true, sizeHalfPt: 26, color: '0F172A', spaceAfter: 100 })
      );
    } else if (b.type === 'h3') {
      parts.push(
        buildSimpleDocxPara(b.text, { bold: true, sizeHalfPt: 23, color: '1E293B', spaceAfter: 80 })
      );
    } else if (b.type === 'li') {
      parts.push(
        buildSimpleDocxPara(`•  ${b.text}`, {
          sizeHalfPt: 22,
          color: '1E293B',
          leftIndentTwips: 360,
          spaceAfter: 80,
        })
      );
    } else {
      parts.push(buildSimpleDocxPara(b.text, { sizeHalfPt: 22, color: '1E293B', spaceAfter: 140 }));
    }
    i++;
  }

  return parts.join('\n    ');
}

function buildDocxTableXml(rows: StructuredBlock[]): string {
  const parsedRows = rows.map((r) =>
    r.cells && r.cells.length > 0
      ? r.cells
      : r.text.split('|').map((c) => c.trim())
  );
  const colCount = Math.max(1, ...parsedRows.map((r) => r.length));
  const colWidthTwips = Math.floor(9360 / colCount);

  const gridCols = Array.from(
    { length: colCount },
    () => `<w:gridCol w:w="${colWidthTwips}"/>`
  ).join('');

  const trXml = parsedRows
    .map((cells, rIdx) => {
      const isHeader = rIdx === 0;
      const tcXml = Array.from({ length: colCount }, (_, cIdx) => {
        const cellText = cells[cIdx] || '';
        return `<w:tc>
          <w:tcPr>
            <w:tcW w:w="${colWidthTwips}" w:type="dxa"/>
            ${isHeader ? '<w:shd w:val="clear" w:color="auto" w:fill="E2E8F0"/>' : ''}
          </w:tcPr>
          <w:p>
            <w:pPr><w:spacing w:before="60" w:after="60"/></w:pPr>
            <w:r>
              <w:rPr>
                ${isHeader ? '<w:b/>' : ''}
                <w:color w:val="0F172A"/>
                <w:sz w:val="20"/>
              </w:rPr>
              <w:t xml:space="preserve">${escapeXml(cellText)}</w:t>
            </w:r>
          </w:p>
        </w:tc>`;
      }).join('\n        ');

      return `<w:tr>
        ${tcXml}
      </w:tr>`;
    })
    .join('\n      ');

  return `<w:tbl>
      <w:tblPr>
        <w:tblW w:w="9360" w:type="dxa"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="6" w:space="0" w:color="94A3B8"/>
          <w:left w:val="single" w:sz="6" w:space="0" w:color="94A3B8"/>
          <w:bottom w:val="single" w:sz="6" w:space="0" w:color="94A3B8"/>
          <w:right w:val="single" w:sz="6" w:space="0" w:color="94A3B8"/>
          <w:insideH w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
          <w:insideV w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid>${gridCols}</w:tblGrid>
      ${trXml}
    </w:tbl>`;
}

function buildSimpleDocxPara(
  text: string,
  opts: {
    bold?: boolean;
    sizeHalfPt?: number;
    color?: string;
    leftIndentTwips?: number;
    spaceAfter?: number;
  } = {}
): string {
  const {
    bold = false,
    sizeHalfPt = 22,
    color = '1E293B',
    leftIndentTwips = 0,
    spaceAfter = 120,
  } = opts;
  return `<w:p>
      <w:pPr>
        <w:spacing w:after="${spaceAfter}"/>
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

/**
 * Extracts a self-contained slide XML + its media images from a source .pptx archive,
 * preserving the slide's background (<p:bg>), placeholder coordinates (<a:xfrm>),
 * theme colors, shapes, tables, and embedded images.
 */
export async function extractNativePptxSlideWithMedia(
  srcZip: JSZip,
  slideNumber1Based: number,
  docIndex: number,
  newSlideIndex1Based: number,
  targetZip: JSZip,
  themeMap: Record<string, string>,
  targetCx: number,
  targetCy: number
): Promise<{
  slideXml: string;
  slideRelsXml: string;
  imageExtensions: Set<string>;
} | null> {
  const slidePath = `ppt/slides/slide${slideNumber1Based}.xml`;
  const slideFile = srcZip.file(slidePath);
  if (!slideFile) return null;

  let slideXml = await slideFile.async('text');
  const imageExtensions = new Set<string>();
  const keptRels: string[] = [
    `  <Relationship Id="rIdLayout1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`,
  ];

  // Read source presentation slide size to scale shape coordinates if aspect ratio changed
  let srcCx = targetCx;
  let srcCy = targetCy;
  try {
    const presFile = srcZip.file('ppt/presentation.xml');
    if (presFile) {
      const presXml = await presFile.async('text');
      const cxM = presXml.match(/<p:sldSz[^>]*\bcx="(\d+)"/);
      const cyM = presXml.match(/<p:sldSz[^>]*\bcy="(\d+)"/);
      if (cxM && cyM) {
        srcCx = parseInt(cxM[1], 10) || targetCx;
        srcCy = parseInt(cyM[1], 10) || targetCy;
      }
    }
  } catch {
    // Ignore
  }

  // Locate slideLayout and slideMaster in srcZip to inherit <p:bg> and placeholder <a:xfrm>
  let layoutXml = '';
  let layoutRelsXml = '';
  let masterXml = '';
  let masterRelsXml = '';

  const slideRelsFile = srcZip.file(`ppt/slides/_rels/slide${slideNumber1Based}.xml.rels`);
  const slideRelsText = slideRelsFile ? await slideRelsFile.async('text') : '';

  const layoutMatch = slideRelsText.match(
    /Target="\.\.\/(slideLayouts\/(slideLayout\d+\.xml))"/i
  );
  if (layoutMatch) {
    const lFile = srcZip.file(`ppt/${layoutMatch[1]}`);
    if (lFile) layoutXml = await lFile.async('text');
    const lRelsFile = srcZip.file(`ppt/slideLayouts/_rels/${layoutMatch[2]}.rels`);
    if (lRelsFile) layoutRelsXml = await lRelsFile.async('text');
  }

  const mFile = srcZip.file('ppt/slideMasters/slideMaster1.xml');
  if (mFile) masterXml = await mFile.async('text');
  const mRelsFile = srcZip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels');
  if (mRelsFile) masterRelsXml = await mRelsFile.async('text');

  // Helper to copy an image from a .rels file into targetZip
  const copyImagesFromRels = async (relsText: string, prefixTag: string, xmlToUpdate: string) => {
    if (!relsText) return xmlToUpdate;
    const parser = new DOMParser();
    const relsDoc = parser.parseFromString(relsText, 'application/xml');
    const relNodes = Array.from(relsDoc.getElementsByTagName('Relationship'));
    let updated = xmlToUpdate;

    for (const rel of relNodes) {
      const oldId = rel.getAttribute('Id') || '';
      const type = rel.getAttribute('Type') || '';
      const target = rel.getAttribute('Target') || '';
      const targetMode = rel.getAttribute('TargetMode') || '';

      if (!oldId || !target) continue;

      if (type.endsWith('/image') && !targetMode) {
        const cleanMediaPath = target.replace(/^\.\.\//, 'ppt/');
        const mediaObj = srcZip.file(cleanMediaPath);
        if (mediaObj) {
          const baseName = cleanMediaPath.split('/').pop() || 'img.png';
          const ext = (baseName.split('.').pop() || 'png').toLowerCase();
          const newMediaName = `s${newSlideIndex1Based}_${prefixTag}_${oldId}_${baseName}`;
          const newRelId = `rId_${prefixTag}_${oldId}`;

          const bytes = await mediaObj.async('uint8array');
          targetZip.file(`ppt/media/${newMediaName}`, bytes);
          imageExtensions.add(ext);

          keptRels.push(
            `  <Relationship Id="${newRelId}" Type="${type}" Target="../media/${newMediaName}"/>`
          );

          updated = updated
            .replace(new RegExp(`r:embed="${oldId}"`, 'g'), `r:embed="${newRelId}"`)
            .replace(new RegExp(`r:id="${oldId}"`, 'g'), `r:id="${newRelId}"`);
        }
      } else if (targetMode === 'External' && type.endsWith('/hyperlink')) {
        const newRelId = `rId_${prefixTag}_${oldId}`;
        keptRels.push(
          `  <Relationship Id="${newRelId}" Type="${type}" Target="${escapeXml(
            target
          )}" TargetMode="External"/>`
        );
        updated = updated.replace(new RegExp(`r:id="${oldId}"`, 'g'), `r:id="${newRelId}"`);
      }
    }
    return updated;
  };

  slideXml = await copyImagesFromRels(slideRelsText, `d${docIndex}s${slideNumber1Based}`, slideXml);

  // If slideXml does not have <p:bg>, inherit <p:bg> from layoutXml or masterXml
  if (!/<p:bg\b/.test(slideXml)) {
    const layoutBgMatch = layoutXml.match(/<p:bg\b[\s\S]*?<\/p:bg>/);
    const masterBgMatch = masterXml.match(/<p:bg\b[\s\S]*?<\/p:bg>/);

    if (layoutBgMatch) {
      const bgXml = await copyImagesFromRels(
        layoutRelsXml,
        `d${docIndex}lay`,
        layoutBgMatch[0]
      );
      slideXml = slideXml.replace(/(<p:cSld[^>]*>)/, `$1\n    ${bgXml}`);
    } else if (masterBgMatch) {
      const bgXml = await copyImagesFromRels(
        masterRelsXml,
        `d${docIndex}mst`,
        masterBgMatch[0]
      );
      slideXml = slideXml.replace(/(<p:cSld[^>]*>)/, `$1\n    ${bgXml}`);
    } else {
      const bgHex = themeMap.bg1 || 'FFFFFF';
      slideXml = slideXml.replace(
        /(<p:cSld[^>]*>)/,
        `$1\n    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="${bgHex}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`
      );
    }
  }

  // Convert <p:bgRef> to explicit <p:bgPr> using themeMap so custom master backgrounds always display
  slideXml = slideXml.replace(
    /<p:bg>\s*<p:bgRef[^>]*>[\s\S]*?<a:schemeClr\s+val="([^"]+)"[^>]*\/?>[\s\S]*?<\/p:bgRef>\s*<\/p:bg>/g,
    (_full, schemeKey) => {
      const resolvedHex = themeMap[schemeKey] || themeMap.bg1 || 'FFFFFF';
      return `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${resolvedHex}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`;
    }
  );

  // Replace simple self-closing <a:schemeClr val="..."/> with resolved <a:srgbClr val="..."/> from this deck's theme
  slideXml = slideXml.replace(/<a:schemeClr\s+val="([^"]+)"\s*\/>/g, (full, key) => {
    if (themeMap[key] && key !== 'phClr') {
      return `<a:srgbClr val="${themeMap[key]}"/>`;
    }
    return full;
  });

  // Ensure every <p:sp> shape on the slide has explicit <a:xfrm> and <a:prstGeom> so blank layout works cleanly
  slideXml = ensureSlideShapesHaveGeometry(
    slideXml,
    layoutXml,
    masterXml,
    srcCx,
    srcCy,
    targetCx,
    targetCy
  );

  // Strip uncopied relationship references (custDataLst, extLst, timing, transition with external rels)
  slideXml = slideXml.replace(/<p:custDataLst\b[\s\S]*?<\/p:custDataLst>/g, '');

  const slideRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${keptRels.join('\n')}
</Relationships>`;

  return {
    slideXml,
    slideRelsXml,
    imageExtensions,
  };
}

function ensureSlideShapesHaveGeometry(
  slideXml: string,
  layoutXml: string,
  masterXml: string,
  srcCx: number,
  srcCy: number,
  targetCx: number,
  targetCy: number
): string {
  // Build map of placeholder xfrms from layout and master
  const phXfrmMap = new Map<string, string>();

  const collectPhXfrms = (xml: string) => {
    if (!xml) return;
    const spMatches = xml.match(/<p:sp\b[\s\S]*?<\/p:sp>/g) || [];
    for (const sp of spMatches) {
      const phMatch = sp.match(/<p:ph\b([^>]*)\/>/);
      const xfrmMatch = sp.match(/<a:xfrm\b[\s\S]*?<\/a:xfrm>/);
      if (phMatch && xfrmMatch) {
        const attrs = phMatch[1];
        const typeM = attrs.match(/\btype="([^"]+)"/);
        const idxM = attrs.match(/\bidx="([^"]+)"/);
        if (typeM && !phXfrmMap.has(`type:${typeM[1]}`)) {
          phXfrmMap.set(`type:${typeM[1]}`, xfrmMatch[0]);
        }
        if (idxM && !phXfrmMap.has(`idx:${idxM[1]}`)) {
          phXfrmMap.set(`idx:${idxM[1]}`, xfrmMatch[0]);
        }
      }
    }
  };

  collectPhXfrms(layoutXml);
  collectPhXfrms(masterXml);

  const scaleX = srcCx > 0 ? targetCx / srcCx : 1;
  const scaleY = srcCy > 0 ? targetCy / srcCy : 1;

  return slideXml.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, (spBlock) => {
    let updatedSp = spBlock;
    const phMatch = updatedSp.match(/<p:ph\b([^>]*)\/>/);

    // If shape has a placeholder <p:ph> and no <a:xfrm> in <p:spPr>, inject inherited or default <a:xfrm>
    if (phMatch && !/<p:spPr>[\s\S]*?<a:xfrm\b/.test(updatedSp)) {
      const attrs = phMatch[1];
      const typeM = attrs.match(/\btype="([^"]+)"/);
      const idxM = attrs.match(/\bidx="([^"]+)"/);
      const phType = typeM ? typeM[1] : '';
      const phIdx = idxM ? idxM[1] : '';

      let inheritedXfrm =
        (phIdx && phXfrmMap.get(`idx:${phIdx}`)) ||
        (phType && phXfrmMap.get(`type:${phType}`)) ||
        '';

      if (!inheritedXfrm) {
        if (phType === 'title' || phType === 'ctrTitle') {
          inheritedXfrm = `<a:xfrm><a:off x="548640" y="457200"/><a:ext cx="${
            srcCx - 1097280
          }" cy="1097280"/></a:xfrm>`;
        } else {
          inheritedXfrm = `<a:xfrm><a:off x="548640" y="1645920"/><a:ext cx="${
            srcCx - 1097280
          }" cy="${Math.max(1800000, srcCy - 2194560)}"/></a:xfrm>`;
        }
      }

      updatedSp = updatedSp.replace(
        /<p:spPr(\s*\/>|>)/,
        (_m, closePart) =>
          closePart.includes('/>')
            ? `<p:spPr>${inheritedXfrm}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`
            : `<p:spPr>${inheritedXfrm}`
      );
    }

    // Ensure <p:spPr> has <a:prstGeom> or <a:custGeom>
    if (
      /<p:spPr>[\s\S]*?<a:xfrm\b/.test(updatedSp) &&
      !/<a:(prstGeom|custGeom)\b/.test(updatedSp)
    ) {
      updatedSp = updatedSp.replace(
        /(<\/a:xfrm>)/,
        `$1<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`
      );
    }

    // Remove <p:ph .../> from <p:nvPr> so shape does not depend on missing placeholder in blank slideLayout1
    updatedSp = updatedSp.replace(/<p:nvPr>[\s\S]*?<\/p:nvPr>/g, '<p:nvPr/>');

    // Scale coordinates if slide aspect ratio differs
    if (Math.abs(scaleX - 1) > 0.01 || Math.abs(scaleY - 1) > 0.01) {
      updatedSp = updatedSp.replace(
        /<a:off\s+x="(\d+)"\s+y="(\d+)"\s*\/>/g,
        (_m, xStr, yStr) =>
          `<a:off x="${Math.round(parseInt(xStr, 10) * scaleX)}" y="${Math.round(
            parseInt(yStr, 10) * scaleY
          )}"/>`
      );
      updatedSp = updatedSp.replace(
        /<a:ext\s+cx="(\d+)"\s+cy="(\d+)"\s*\/>/g,
        (_m, cxStr, cyStr) =>
          `<a:ext cx="${Math.round(parseInt(cxStr, 10) * scaleX)}" cy="${Math.round(
            parseInt(cyStr, 10) * scaleY
          )}"/>`
      );
    }

    return updatedSp;
  });
}

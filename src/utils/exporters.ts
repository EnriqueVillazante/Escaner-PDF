import {
  AlignmentType,
  Document,
  FrameAnchorType,
  FrameWrap,
  HorizontalPositionRelativeFrom,
  ImageRun,
  Packer,
  PageOrientation,
  Paragraph,
  TextRun,
  TextWrappingType,
  VerticalPositionRelativeFrom,
} from 'docx';
import { degrees, PDFDocument, rgb, StandardFonts } from 'pdf-lib';
import type { ScannedDocument, ScannedPage, TextAlign } from '../types/document';

function stripExtension(fileName: string): string {
  return fileName.replace(/\.[^/.]+$/, '');
}

function hexToRgbUnit(hex: string): { r: number; g: number; b: number } {
  const cleaned = (hex || '#ffffff').replace('#', '').trim();
  const full =
    cleaned.length === 3
      ? cleaned
          .split('')
          .map((c) => c + c)
          .join('')
      : cleaned.padEnd(6, '0').slice(0, 6);
  const num = parseInt(full, 16);
  return {
    r: ((num >> 16) & 255) / 255,
    g: ((num >> 8) & 255) / 255,
    b: (num & 255) / 255,
  };
}

function normalizeHex6(hex?: string, fallback = '0F172A'): string {
  const cleaned = (hex || '').replace('#', '').trim();
  if (!/^[0-9a-fA-F]{3,6}$/.test(cleaned)) return fallback;
  if (cleaned.length === 3) {
    return cleaned
      .split('')
      .map((c) => c + c)
      .join('')
      .toUpperCase();
  }
  return cleaned.slice(0, 6).toUpperCase();
}

/**
 * Limpia caracteres fuera de WinAnsi para que StandardFonts de pdf-lib nunca falle
 */
function sanitizeWinAnsi(text: string): string {
  return text
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/[\u2013\u2014]/g, '-')
    .replace(/\u2026/g, '...')
    .replace(/\u00A0/g, ' ')
    .replace(/[^\x20-\x7E\xA0-\xFF\n]/g, '');
}

/**
 * Invierte una matriz afín 2D [a, b, c, d, e, f] para mapear coordenadas exactas
 * desde el viewport visual (pantalla) de vuelta al espacio nativo del PDF.
 */
function invertMatrix(
  m: [number, number, number, number, number, number]
): [number, number, number, number, number, number] {
  const [a, b, c, d, e, f] = m;
  const det = a * d - b * c;
  if (Math.abs(det) < 1e-8) {
    return [1, 0, 0, -1, 0, 0];
  }
  const invDet = 1 / det;
  return [
    d * invDet,
    -b * invDet,
    -c * invDet,
    a * invDet,
    (c * f - d * e) * invDet,
    (b * e - a * f) * invDet,
  ];
}

function applyMatrix(
  m: [number, number, number, number, number, number],
  x: number,
  y: number
): { x: number; y: number } {
  return {
    x: m[0] * x + m[2] * y + m[4],
    y: m[1] * x + m[3] * y + m[5],
  };
}

function triggerDownload(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

/**
 * Exporta el PDF editado usando pdf-lib sobre el archivo PDF ORIGINAL:
 * Conserva 100% intactos los logos, gráficos, bordes de tabla y orientación original.
 */
export async function exportDocumentToPdf(scannedDoc: ScannedDocument): Promise<void> {
  const pdfDoc = await PDFDocument.load(scannedDoc.originalPdfBytes, {
    ignoreEncryption: true,
  });

  const fontRegular = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
  const pdfPages = pdfDoc.getPages();

  scannedDoc.pages.forEach((scannedPage, idx) => {
    const pdfPage = pdfPages[idx];
    if (!pdfPage) return;

    const invTransform = invertMatrix(scannedPage.viewportTransform);
    const pageRot = scannedPage.pageRotation || 0;

    for (const field of scannedPage.fields) {
      if (!field.isModified && !field.isDeleted && !field.isNew) {
        continue;
      }

      const pad = 1.6;
      const vx1 = field.x - pad;
      const vy1 = field.y - pad;
      const vx2 = field.x + field.width + pad;
      const vy2 = field.y + field.height + pad;

      // 1. Tapar el texto original transformando las 4 esquinas al espacio nativo PDF
      if (!field.isNew) {
        const corners = [
          applyMatrix(invTransform, vx1, vy1),
          applyMatrix(invTransform, vx2, vy1),
          applyMatrix(invTransform, vx1, vy2),
          applyMatrix(invTransform, vx2, vy2),
        ];
        const minX = Math.min(...corners.map((c) => c.x));
        const maxX = Math.max(...corners.map((c) => c.x));
        const minY = Math.min(...corners.map((c) => c.y));
        const maxY = Math.max(...corners.map((c) => c.y));

        const bg = hexToRgbUnit(field.bgColor);
        pdfPage.drawRectangle({
          x: minX,
          y: minY,
          width: Math.max(4, maxX - minX),
          height: Math.max(6, maxY - minY),
          color: rgb(bg.r, bg.g, bg.b),
        });
      }

      // 2. Si no está borrado, dibujar el texto nuevo respetando rotación de página y campo
      if (!field.isDeleted && field.text.trim().length > 0) {
        const font = field.isBold ? fontBold : fontRegular;
        const size = Math.max(6, field.fontSize);
        const tc = hexToRgbUnit(field.textColor);
        const cleanText = sanitizeWinAnsi(field.text);
        const textWidth = font.widthOfTextAtSize(cleanText, size);

        let baseVx = field.x;
        let baseVy = field.y + size * 0.88;

        if (field.rotation === 0) {
          if (field.align === 'right') {
            baseVx = field.x + Math.max(0, field.width - textWidth);
          } else if (field.align === 'center') {
            baseVx = field.x + Math.max(0, (field.width - textWidth) / 2);
          }
          baseVy = field.y + size * 0.88;
        } else if (field.rotation === -90) {
          baseVx = field.x + size * 0.88;
          if (field.align === 'center') {
            baseVy = field.y + field.height - Math.max(0, (field.height - textWidth) / 2);
          } else if (field.align === 'right') {
            baseVy = field.y + textWidth;
          } else {
            baseVy = field.y + field.height;
          }
        } else if (field.rotation === 90) {
          baseVx = field.x + Math.max(2, field.width - size * 0.88);
          if (field.align === 'center') {
            baseVy = field.y + Math.max(0, (field.height - textWidth) / 2);
          } else if (field.align === 'right') {
            baseVy = field.y + Math.max(0, field.height - textWidth);
          } else {
            baseVy = field.y;
          }
        }

        const pdfPos = applyMatrix(invTransform, baseVx, baseVy);
        const pdfAngleDeg = ((-field.rotation + pageRot) % 360 + 360) % 360;

        pdfPage.drawText(cleanText, {
          x: pdfPos.x,
          y: pdfPos.y,
          size,
          font,
          color: rgb(tc.r, tc.g, tc.b),
          rotate: degrees(pdfAngleDeg),
        });
      }
    }
  });

  const savedBytes = await pdfDoc.save();
  const arrayCopy = new Uint8Array(savedBytes);
  const blob = new Blob([arrayCopy.buffer], { type: 'application/pdf' });
  const outName = `${stripExtension(scannedDoc.fileName)}-editado.pdf`;
  triggerDownload(blob, outName);
}

function mapDocxAlignment(align?: TextAlign) {
  switch (align) {
    case 'center':
      return AlignmentType.CENTER;
    case 'right':
      return AlignmentType.RIGHT;
    default:
      return AlignmentType.LEFT;
  }
}

function dataUrlToUint8Array(dataUrl: string): Uint8Array {
  const base64 = dataUrl.split(',')[1] ?? '';
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

/**
 * Genera la imagen de fondo exacta de la página para Word (.docx):
 * - Conserva al 100% logos, bordes de tablas, líneas, colores y diseño del PDF original.
 * - Limpia (tapa con su bgColor real) el texto viejo de los campos horizontales porque encima
 *   van como texto real 100% editable en Word en sus coordenadas exactas.
 * - Para los campos verticales/rotados que fueron modificados o borrados, actualiza directamente
 *   el texto rotado sobre el fondo para que en Word se vea idéntico sin descuadrarse.
 */
async function buildWordPageBackgroundPng(page: ScannedPage): Promise<Uint8Array> {
  const img = await new Promise<HTMLImageElement>((resolve, reject) => {
    const el = new Image();
    el.onload = () => resolve(el);
    el.onerror = reject;
    el.src = page.originalImageUrl;
  });

  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, img.naturalWidth);
  canvas.height = Math.max(1, img.naturalHeight);
  const ctx = canvas.getContext('2d');
  if (!ctx) {
    return dataUrlToUint8Array(page.originalImageUrl);
  }

  ctx.drawImage(img, 0, 0);
  const scaleX = canvas.width / page.width;
  const scaleY = canvas.height / page.height;

  for (const field of page.fields) {
    const padX = 1.6;
    const padY = 1.4;

    if (field.rotation === 0) {
      // Limpiar texto original horizontal en el fondo para colocar encima el texto editable de Word
      if (!field.isNew) {
        ctx.fillStyle = field.bgColor || '#ffffff';
        ctx.fillRect(
          Math.max(0, (field.x - padX) * scaleX),
          Math.max(0, (field.y - padY) * scaleY),
          Math.max(4, (field.width + padX * 2) * scaleX),
          Math.max(6, (field.height + padY * 2) * scaleY)
        );
      }
    } else if (field.isModified || field.isDeleted || field.isNew) {
      // Campos verticales/rotados modificados: tapar original y redibujar en su ángulo exacto
      if (!field.isNew) {
        ctx.fillStyle = field.bgColor || '#ffffff';
        ctx.fillRect(
          Math.max(0, (field.x - padX) * scaleX),
          Math.max(0, (field.y - padY) * scaleY),
          Math.max(4, (field.width + padX * 2) * scaleX),
          Math.max(6, (field.height + padY * 2) * scaleY)
        );
      }

      if (!field.isDeleted && field.text.trim().length > 0) {
        ctx.save();
        const cx = (field.x + field.width / 2) * scaleX;
        const cy = (field.y + field.height / 2) * scaleY;
        ctx.translate(cx, cy);
        ctx.rotate((field.rotation * Math.PI) / 180);
        ctx.fillStyle = field.textColor || '#0f172a';
        ctx.font = `${field.isBold ? 'bold ' : ''}${Math.round(field.fontSize * scaleY)}px Arial, Helvetica, sans-serif`;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(field.text, 0, 0);
        ctx.restore();
      }
    }
  }

  return dataUrlToUint8Array(canvas.toDataURL('image/png'));
}

/**
 * Exporta el documento editado a Word (.docx) con fidelidad visual 100% idéntica al PDF:
 * - Una sección por página respetando el tamaño y orientación exacta (vertical u horizontal).
 * - Imagen de fondo detrás del documento (logos, bordes de tabla, líneas, colores intactos).
 * - Cada campo de texto horizontal posicionado en sus coordenadas exactas (twips) como
 *   texto real editable en Microsoft Word mediante frames absolutos anclados a la página.
 */
export async function exportDocumentToWord(scannedDoc: ScannedDocument): Promise<void> {
  if (scannedDoc.pages.length === 0) return;

  const sections = [];

  for (const page of scannedDoc.pages) {
    const isLandscape = page.width > page.height;
    const shortSideTwips = Math.round(Math.min(page.width, page.height) * 20);
    const longSideTwips = Math.round(Math.max(page.width, page.height) * 20);

    // En docx, transformation usa píxeles a 96 DPI (1 pt PDF = 96 / 72 px)
    const pageWidthPx96 = Math.round((page.width * 96) / 72);
    const pageHeightPx96 = Math.round((page.height * 96) / 72);

    const bgBytes = await buildWordPageBackgroundPng(page);
    const pageChildren: Paragraph[] = [];

    // 1. Fondo exacto del PDF (logos, bordes de tablas, líneas, colores) detrás del texto editable
    pageChildren.push(
      new Paragraph({
        spacing: { before: 0, after: 0 },
        children: [
          new ImageRun({
            data: bgBytes,
            transformation: {
              width: pageWidthPx96,
              height: pageHeightPx96,
            },
            type: 'png',
            floating: {
              horizontalPosition: {
                relative: HorizontalPositionRelativeFrom.PAGE,
                offset: 0,
              },
              verticalPosition: {
                relative: VerticalPositionRelativeFrom.PAGE,
                offset: 0,
              },
              behindDocument: true,
              allowOverlap: true,
              lockAnchor: true,
              wrap: {
                type: TextWrappingType.NONE,
              },
            },
          }),
        ],
      })
    );

    // 2. Campos de texto editables en sus coordenadas exactas sobre el diseño original
    for (const field of page.fields) {
      if (field.rotation !== 0) continue;
      if (field.isDeleted || !field.text.trim()) continue;

      const xTwips = Math.max(0, Math.round(field.x * 20));
      const yTwips = Math.max(0, Math.round(field.y * 20));
      // Un pequeño margen extra en ancho para evitar saltos de línea no deseados en Word
      const wTwips = Math.max(240, Math.round((field.width + 8) * 20));
      const hTwips = Math.max(150, Math.round((field.height + 2) * 20));
      const fontSizeHalfPts = Math.max(12, Math.min(72, Math.round(field.fontSize * 2)));

      pageChildren.push(
        new Paragraph({
          alignment: mapDocxAlignment(field.align),
          spacing: {
            before: 0,
            after: 0,
            line: Math.max(140, Math.round(field.fontSize * 21)),
          },
          frame: {
            type: 'absolute',
            position: {
              x: xTwips,
              y: yTwips,
            },
            width: wTwips,
            height: hTwips,
            anchor: {
              horizontal: FrameAnchorType.PAGE,
              vertical: FrameAnchorType.PAGE,
            },
            wrap: FrameWrap.NONE,
          },
          children: [
            new TextRun({
              text: field.text,
              bold: field.isBold,
              size: fontSizeHalfPts,
              font: 'Arial',
              color: normalizeHex6(field.textColor, '0F172A'),
            }),
          ],
        })
      );
    }

    sections.push({
      properties: {
        page: {
          size: {
            orientation: isLandscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT,
            width: shortSideTwips,
            height: longSideTwips,
          },
          margin: {
            top: 0,
            bottom: 0,
            left: 0,
            right: 0,
            header: 0,
            footer: 0,
          },
        },
      },
      children: pageChildren,
    });
  }

  const doc = new Document({
    sections,
  });

  const blob = await Packer.toBlob(doc);
  const outName = `${stripExtension(scannedDoc.fileName)}-editado.docx`;
  triggerDownload(blob, outName);
}

// ============================================================================
// Generador nativo de OpenDocument Text (.odt) editable y 100% fiel al PDF
// ============================================================================

function escapeXml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

const CRC32_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i += 1) {
    let c = i;
    for (let k = 0; k < 8; k += 1) {
      c = (c & 1) !== 0 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[i] = c >>> 0;
  }
  return table;
})();

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (let i = 0; i < bytes.length; i += 1) {
    crc = CRC32_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

/**
 * Empaquetador ZIP estándar (método STORE = 0), ideal para .odt donde mimetype debe ir primero y sin comprimir
 */
function buildZipStore(entries: { name: string; data: Uint8Array }[]): Uint8Array {
  const encoder = new TextEncoder();
  const localParts: Uint8Array[] = [];
  const centralParts: Uint8Array[] = [];
  let offset = 0;

  for (const entry of entries) {
    const nameBytes = encoder.encode(entry.name);
    const data = entry.data;
    const checksum = crc32(data);

    // Local File Header (30 bytes + name)
    const localHeader = new Uint8Array(30 + nameBytes.length);
    const lv = new DataView(localHeader.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true); // version needed
    lv.setUint16(6, 0x0800, true); // UTF-8 filename flag
    lv.setUint16(8, 0, true); // method: STORE
    lv.setUint16(10, 0, true);
    lv.setUint16(12, 0, true);
    lv.setUint32(14, checksum, true);
    lv.setUint32(18, data.length, true);
    lv.setUint32(22, data.length, true);
    lv.setUint16(26, nameBytes.length, true);
    lv.setUint16(28, 0, true);
    localHeader.set(nameBytes, 30);

    localParts.push(localHeader, data);

    // Central Directory Header (46 bytes + name)
    const centralHeader = new Uint8Array(46 + nameBytes.length);
    const cv = new DataView(centralHeader.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 20, true); // version made by
    cv.setUint16(6, 20, true); // version needed
    cv.setUint16(8, 0x0800, true); // UTF-8
    cv.setUint16(10, 0, true); // method: STORE
    cv.setUint16(12, 0, true);
    cv.setUint16(14, 0, true);
    cv.setUint32(16, checksum, true);
    cv.setUint32(20, data.length, true);
    cv.setUint32(24, data.length, true);
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint16(30, 0, true);
    cv.setUint16(32, 0, true);
    cv.setUint16(34, 0, true);
    cv.setUint16(36, 0, true);
    cv.setUint32(38, 0, true);
    cv.setUint32(42, offset, true);
    centralHeader.set(nameBytes, 46);

    centralParts.push(centralHeader);
    offset += localHeader.length + data.length;
  }

  const centralSize = centralParts.reduce((s, p) => s + p.length, 0);

  // End of Central Directory Record (22 bytes)
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(4, 0, true);
  ev.setUint16(6, 0, true);
  ev.setUint16(8, entries.length, true);
  ev.setUint16(10, entries.length, true);
  ev.setUint32(12, centralSize, true);
  ev.setUint32(16, offset, true);
  ev.setUint16(20, 0, true);

  const totalLen = offset + centralSize + eocd.length;
  const out = new Uint8Array(totalLen);
  let pos = 0;
  for (const part of localParts) {
    out.set(part, pos);
    pos += part.length;
  }
  for (const part of centralParts) {
    out.set(part, pos);
    pos += part.length;
  }
  out.set(eocd, pos);
  return out;
}

/**
 * Exporta el documento editado a OpenDocument Text (.odt) editable y fiel al PDF:
 * - Cada página conserva su tamaño y orientación exacta (horizontal o vertical).
 * - Fondo exacto del PDF (logos, bordes de tabla, líneas, colores y texto vertical actualizado).
 * - Cuadros de texto nativos (<draw:text-box>) en coordenadas exactas editables en LibreOffice / Word.
 */
export async function exportDocumentToOdt(scannedDoc: ScannedDocument): Promise<void> {
  if (scannedDoc.pages.length === 0) return;

  const encoder = new TextEncoder();
  const entries: { name: string; data: Uint8Array }[] = [];
  const manifestPictures: string[] = [];

  // 1. mimetype (siempre primero y sin comprimir según estándar OASIS ODF)
  entries.push({
    name: 'mimetype',
    data: encoder.encode('application/vnd.oasis.opendocument.text'),
  });

  // 2. Generar fondos de página PNG y recopilar estilos/elementos XML
  const pageLayoutsXml: string[] = [];
  const masterPagesXml: string[] = [];
  const contentAutoStylesXml: string[] = [];
  const bodyParagraphsXml: string[] = [];

  contentAutoStylesXml.push(
    `<style:style style:name="fr_bg" style:family="graphic">
      <style:graphic-properties style:run-through="background" style:wrap="run-through" style:horizontal-pos="from-left" style:horizontal-rel="page" style:vertical-pos="from-top" style:vertical-rel="page" draw:stroke="none" draw:fill="none" fo:border="none" fo:padding="0pt"/>
    </style:style>`,
    `<style:style style:name="fr_txt" style:family="graphic">
      <style:graphic-properties style:run-through="foreground" style:wrap="none" style:horizontal-pos="from-left" style:horizontal-rel="page" style:vertical-pos="from-top" style:vertical-rel="page" draw:stroke="none" draw:fill="none" fo:border="none" fo:padding="0pt"/>
    </style:style>`
  );

  let fieldStyleCounter = 0;

  for (let pIdx = 0; pIdx < scannedDoc.pages.length; pIdx += 1) {
    const page = scannedDoc.pages[pIdx];
    const pageNum = pIdx + 1;
    const isLandscape = page.width > page.height;
    const picPath = `Pictures/page-${pageNum}.png`;

    const bgPng = await buildWordPageBackgroundPng(page);
    entries.push({ name: picPath, data: bgPng });
    manifestPictures.push(
      `  <manifest:file-entry manifest:full-path="${picPath}" manifest:media-type="image/png"/>`
    );

    pageLayoutsXml.push(
      `<style:page-layout style:name="pm${pageNum}">
        <style:page-layout-properties fo:page-width="${page.width.toFixed(2)}pt" fo:page-height="${page.height.toFixed(2)}pt" style:print-orientation="${isLandscape ? 'landscape' : 'portrait'}" fo:margin-top="0pt" fo:margin-bottom="0pt" fo:margin-left="0pt" fo:margin-right="0pt"/>
      </style:page-layout>`
    );

    masterPagesXml.push(
      `<style:master-page style:name="mp${pageNum}" style:page-layout-name="pm${pageNum}"/>`
    );

    contentAutoStylesXml.push(
      `<style:style style:name="P_PAGE_${pageNum}" style:family="paragraph" style:master-page-name="mp${pageNum}">
        <style:paragraph-properties fo:margin-top="0pt" fo:margin-bottom="0pt" fo:break-before="${pageNum > 1 ? 'page' : 'auto'}"/>
      </style:style>`
    );

    const framesXml: string[] = [];

    // Imagen de fondo cubriendo el 100% de la página
    framesXml.push(
      `<draw:frame draw:style-name="fr_bg" draw:name="FondoPag${pageNum}" text:anchor-type="page" text:anchor-page-number="${pageNum}" draw:z-index="0" svg:x="0pt" svg:y="0pt" svg:width="${page.width.toFixed(2)}pt" svg:height="${page.height.toFixed(2)}pt">
        <draw:image xlink:href="${picPath}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>
      </draw:frame>`
    );

    // Campos de texto editables posicionados en coordenadas exactas
    for (const field of page.fields) {
      if (field.rotation !== 0) continue;
      if (field.isDeleted || !field.text.trim()) continue;

      fieldStyleCounter += 1;
      const stName = `P_F_${fieldStyleCounter}`;
      const odtAlign =
        field.align === 'center' ? 'center' : field.align === 'right' ? 'end' : 'start';
      const colorHex = normalizeHex6(field.textColor, '0F172A');

      contentAutoStylesXml.push(
        `<style:style style:name="${stName}" style:family="paragraph">
          <style:paragraph-properties fo:text-align="${odtAlign}" fo:margin-top="0pt" fo:margin-bottom="0pt" fo:line-height="115%"/>
          <style:text-properties fo:font-family="Arial" fo:font-size="${Math.max(6, field.fontSize).toFixed(1)}pt" fo:font-weight="${field.isBold ? 'bold' : 'normal'}" fo:color="#${colorHex}"/>
        </style:style>`
      );

      const xPt = Math.max(0, field.x).toFixed(2);
      const yPt = Math.max(0, field.y).toFixed(2);
      const wPt = Math.max(14, field.width + 8).toFixed(2);
      const hPt = Math.max(10, field.height + 2).toFixed(2);

      framesXml.push(
        `<draw:frame draw:style-name="fr_txt" draw:name="Campo_${fieldStyleCounter}" text:anchor-type="page" text:anchor-page-number="${pageNum}" draw:z-index="1" svg:x="${xPt}pt" svg:y="${yPt}pt" svg:width="${wPt}pt" svg:height="${hPt}pt">
          <draw:text-box>
            <text:p text:style-name="${stName}">${escapeXml(field.text)}</text:p>
          </draw:text-box>
        </draw:frame>`
      );
    }

    bodyParagraphsXml.push(
      `<text:p text:style-name="P_PAGE_${pageNum}">${framesXml.join('')}</text:p>`
    );
  }

  const stylesXml = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.2">
  <office:automatic-styles>
    ${pageLayoutsXml.join('\n    ')}
  </office:automatic-styles>
  <office:master-styles>
    ${masterPagesXml.join('\n    ')}
  </office:master-styles>
</office:document-styles>`;

  const contentXml = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.2">
  <office:automatic-styles>
    ${contentAutoStylesXml.join('\n    ')}
  </office:automatic-styles>
  <office:body>
    <office:text>
      ${bodyParagraphsXml.join('\n      ')}
    </office:text>
  </office:body>
</office:document-content>`;

  const metaXml = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" office:version="1.2">
  <office:meta>
    <dc:title>${escapeXml(stripExtension(scannedDoc.fileName))}</dc:title>
  </office:meta>
</office:document-meta>`;

  const manifestXml = `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
  <manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
${manifestPictures.join('\n')}
</manifest:manifest>`;

  entries.push(
    { name: 'styles.xml', data: encoder.encode(stylesXml) },
    { name: 'content.xml', data: encoder.encode(contentXml) },
    { name: 'meta.xml', data: encoder.encode(metaXml) },
    { name: 'META-INF/manifest.xml', data: encoder.encode(manifestXml) }
  );

  const zipBytes = buildZipStore(entries);
  const copy = new Uint8Array(zipBytes);
  const blob = new Blob([copy.buffer], {
    type: 'application/vnd.oasis.opendocument.text',
  });
  const outName = `${stripExtension(scannedDoc.fileName)}-editado.odt`;
  triggerDownload(blob, outName);
}

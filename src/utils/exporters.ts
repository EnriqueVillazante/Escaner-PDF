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

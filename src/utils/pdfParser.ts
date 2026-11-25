import * as pdfjsLib from 'pdfjs-dist';
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import Tesseract from 'tesseract.js';
import type {
  DetectedTable,
  EditableField,
  ScannedDocument,
  ScannedPage,
  ScanProgress,
} from '../types/document';

// Configurar el worker de PDF.js compatible con Vite
pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

interface RawTextItem {
  text: string;
  x: number;
  y: number;
  width: number;
  height: number;
  fontSize: number;
  isBold: boolean;
  /** Ángulo normalizado en grados: 0 (horizontal), -90 (vertical arriba), 90 (vertical abajo), 180, u otro */
  rotation: number;
}

interface LineSegment {
  text: string;
  x: number;
  xEnd: number;
  y: number;
  yEnd: number;
  height: number;
  fontSize: number;
  isBold: boolean;
  rotation: number;
}

interface TextLine {
  y: number;
  fontSize: number;
  height: number;
  segments: LineSegment[];
}

let idCounter = 1;
export function createId(prefix = 'fld'): string {
  idCounter += 1;
  return `${prefix}-${Date.now().toString(36)}-${idCounter}`;
}

function detectBoldFont(fontName: string, fontFamily?: string): boolean {
  const combined = `${fontName} ${fontFamily ?? ''}`.toLowerCase();
  return (
    combined.includes('bold') ||
    combined.includes('black') ||
    combined.includes('heavy') ||
    combined.includes('demi') ||
    combined.includes('semibold') ||
    combined.includes('-bd') ||
    combined.includes(',b')
  );
}

function looksNumeric(text: string): boolean {
  const cleaned = text.trim();
  if (!cleaned) return false;
  return /^[+-]?[$€£¥S/]*\s*\d[\d.,\s]*%?$/.test(cleaned);
}

function normalizeAngleDeg(deg: number): number {
  let a = ((deg + 180) % 360 + 360) % 360 - 180;
  // Ajustar ángulos casi exactos a 0, 90, -90, 180
  if (Math.abs(a) <= 4) return 0;
  if (Math.abs(a - 90) <= 4) return 90;
  if (Math.abs(a + 90) <= 4) return -90;
  if (Math.abs(Math.abs(a) - 180) <= 4) return 180;
  return Math.round(a);
}

function rgbToHex(r: number, g: number, b: number): string {
  const toHex = (n: number) =>
    Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0');
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

function relativeLuminance(r: number, g: number, b: number): number {
  return (0.299 * r + 0.587 * g + 0.114 * b) / 255;
}

/**
 * Detecta si existe una línea vertical de borde de tabla en el canvas entre x1Pt y x2Pt.
 * Si existe una línea divisoria real, jamás se fusionan dos palabras de columnas distintas.
 */
function hasVerticalSeparatorLine(
  ctx: CanvasRenderingContext2D | null,
  canvasWidth: number,
  canvasHeight: number,
  x1Pt: number,
  x2Pt: number,
  yPt: number,
  heightPt: number,
  scale: number
): boolean {
  if (!ctx || x2Pt - x1Pt < 1.5) return false;

  const startX = Math.max(0, Math.min(canvasWidth - 1, Math.round(x1Pt * scale)));
  const endX = Math.max(0, Math.min(canvasWidth - 1, Math.round(x2Pt * scale)));
  const startY = Math.max(0, Math.min(canvasHeight - 1, Math.round(yPt * scale)));
  const hPx = Math.max(6, Math.min(canvasHeight - startY, Math.round(heightPt * scale)));
  const wPx = endX - startX;

  if (wPx < 2 || hPx < 6) return false;

  const imgData = ctx.getImageData(startX, startY, wPx, hPx).data;

  // Buscar una columna de píxeles donde >= 70% de la altura sea más oscura que el fondo blanco/claro
  for (let col = 0; col < wPx; col += 1) {
    let darkCount = 0;
    for (let row = 0; row < hPx; row += 1) {
      const idx = (row * wPx + col) * 4;
      const lum = relativeLuminance(imgData[idx], imgData[idx + 1], imgData[idx + 2]);
      if (lum < 0.86) {
        darkCount += 1;
      }
    }
    if (darkCount / hPx >= 0.72) {
      return true;
    }
  }
  return false;
}

/**
 * Detecta si existe una línea horizontal de borde de tabla en el canvas entre y1Pt y y2Pt
 * (útil para separar celdas de texto vertical apiladas).
 */
function hasHorizontalSeparatorLine(
  ctx: CanvasRenderingContext2D | null,
  canvasWidth: number,
  canvasHeight: number,
  xPt: number,
  widthPt: number,
  y1Pt: number,
  y2Pt: number,
  scale: number
): boolean {
  if (!ctx || y2Pt - y1Pt < 1.5) return false;

  const startX = Math.max(0, Math.min(canvasWidth - 1, Math.round(xPt * scale)));
  const wPx = Math.max(6, Math.min(canvasWidth - startX, Math.round(widthPt * scale)));
  const startY = Math.max(0, Math.min(canvasHeight - 1, Math.round(y1Pt * scale)));
  const endY = Math.max(0, Math.min(canvasHeight - 1, Math.round(y2Pt * scale)));
  const hPx = endY - startY;

  if (wPx < 6 || hPx < 2) return false;

  const imgData = ctx.getImageData(startX, startY, wPx, hPx).data;
  for (let row = 0; row < hPx; row += 1) {
    let darkCount = 0;
    for (let col = 0; col < wPx; col += 1) {
      const idx = (row * wPx + col) * 4;
      const lum = relativeLuminance(imgData[idx], imgData[idx + 1], imgData[idx + 2]);
      if (lum < 0.86) {
        darkCount += 1;
      }
    }
    if (darkCount / wPx >= 0.72) {
      return true;
    }
  }
  return false;
}

/**
 * Muestrea el color de fondo real y color de texto alrededor de la caja en el canvas del PDF.
 */
function sampleColorsAtBox(
  ctx: CanvasRenderingContext2D | null,
  canvasWidth: number,
  canvasHeight: number,
  xPt: number,
  yPt: number,
  wPt: number,
  hPt: number,
  scale: number
): { bgColor: string; textColor: string } {
  if (!ctx) return { bgColor: '#ffffff', textColor: '#0f172a' };

  const px = Math.max(0, Math.min(canvasWidth - 1, Math.round(xPt * scale)));
  const py = Math.max(0, Math.min(canvasHeight - 1, Math.round(yPt * scale)));
  const pw = Math.max(2, Math.min(canvasWidth - px, Math.round(wPt * scale)));
  const ph = Math.max(2, Math.min(canvasHeight - py, Math.round(hPt * scale)));

  const samplePoints: [number, number][] = [];
  const topY = Math.max(0, py - 2);
  const botY = Math.min(canvasHeight - 1, py + ph + 1);
  const steps = Math.min(8, Math.max(3, Math.floor(pw / 10)));

  for (let i = 0; i < steps; i += 1) {
    const sx = Math.min(canvasWidth - 1, px + Math.round(((i + 0.5) / steps) * pw));
    samplePoints.push([sx, topY], [sx, botY]);
  }
  const midY = Math.min(canvasHeight - 1, py + Math.round(ph / 2));
  if (px >= 2) samplePoints.push([px - 2, midY]);
  if (px + pw + 2 < canvasWidth) samplePoints.push([px + pw + 2, midY]);

  const rs: number[] = [];
  const gs: number[] = [];
  const bs: number[] = [];

  for (const [sx, sy] of samplePoints) {
    const pixel = ctx.getImageData(sx, sy, 1, 1).data;
    rs.push(pixel[0]);
    gs.push(pixel[1]);
    bs.push(pixel[2]);
  }

  if (rs.length === 0) return { bgColor: '#ffffff', textColor: '#0f172a' };

  rs.sort((a, b) => a - b);
  gs.sort((a, b) => a - b);
  bs.sort((a, b) => a - b);

  const mid = Math.floor(rs.length / 2);
  const bgR = rs[mid];
  const bgG = gs[mid];
  const bgB = bs[mid];
  const bgColor = rgbToHex(bgR, bgG, bgB);
  const lum = relativeLuminance(bgR, bgG, bgB);

  return {
    bgColor,
    textColor: lum < 0.52 ? '#ffffff' : '#0f172a',
  };
}

/**
 * Agrupa ítems horizontales (rotation === 0) en líneas y celdas respetando líneas divisorias de tabla
 */
function buildHorizontalLines(
  items: RawTextItem[],
  ctx: CanvasRenderingContext2D | null,
  canvasWidth: number,
  canvasHeight: number,
  renderScale: number
): TextLine[] {
  const validItems = items
    .map((it) => ({ ...it, text: it.text.replace(/\s+/g, ' ').trim() }))
    .filter((it) => it.text.length > 0);

  if (validItems.length === 0) return [];

  validItems.sort((a, b) => {
    const dy = a.y - b.y;
    if (Math.abs(dy) > Math.min(a.fontSize, b.fontSize) * 0.42) {
      return dy;
    }
    return a.x - b.x;
  });

  const rawLines: RawTextItem[][] = [];
  for (const item of validItems) {
    let matched: RawTextItem[] | null = null;
    for (let i = rawLines.length - 1; i >= Math.max(0, rawLines.length - 6); i -= 1) {
      const ref = rawLines[i][0];
      const tol = Math.max(ref.fontSize, item.fontSize) * 0.48;
      if (Math.abs(item.y - ref.y) <= tol) {
        matched = rawLines[i];
        break;
      }
    }
    if (matched) {
      matched.push(item);
    } else {
      rawLines.push([item]);
    }
  }

  const lines: TextLine[] = rawLines.map((rowItems) => {
    rowItems.sort((a, b) => a.x - b.x);
    const avgY = rowItems.reduce((s, i) => s + i.y, 0) / rowItems.length;
    const avgFontSize = rowItems.reduce((s, i) => s + i.fontSize, 0) / rowItems.length;
    const maxHeight = Math.max(...rowItems.map((i) => i.height || i.fontSize));

    const segments: LineSegment[] = [];
    for (const item of rowItems) {
      const itemEnd = item.x + Math.max(item.width, item.text.length * item.fontSize * 0.42);
      const last = segments[segments.length - 1];

      const gap = last ? item.x - last.xEnd : Infinity;
      const gapThreshold = Math.max(avgFontSize * 0.68, 6.2);

      // Si hay una línea vertical de borde entre las dos palabras, NO unirlas jamás
      const blockedByBorder =
        last && gap > 0.5 && gap <= gapThreshold * 1.4
          ? hasVerticalSeparatorLine(
              ctx,
              canvasWidth,
              canvasHeight,
              last.xEnd,
              item.x,
              Math.min(last.y, item.y),
              Math.max(last.height, item.height),
              renderScale
            )
          : false;

      if (last && gap <= gapThreshold && !blockedByBorder) {
        const needsSpace =
          gap > avgFontSize * 0.1 && !last.text.endsWith(' ') && !item.text.startsWith(' ');
        last.text = `${last.text}${needsSpace ? ' ' : ''}${item.text}`;
        last.xEnd = Math.max(last.xEnd, itemEnd);
        last.height = Math.max(last.height, item.height || item.fontSize);
        last.yEnd = Math.max(last.yEnd, item.y + item.height);
        last.fontSize = (last.fontSize + item.fontSize) / 2;
        last.isBold = last.isBold || item.isBold;
      } else {
        segments.push({
          text: item.text,
          x: item.x,
          xEnd: itemEnd,
          y: item.y,
          yEnd: item.y + Math.max(item.height || item.fontSize, item.fontSize * 1.1),
          height: Math.max(item.height || item.fontSize, item.fontSize * 1.1),
          fontSize: item.fontSize,
          isBold: item.isBold,
          rotation: 0,
        });
      }
    }

    return {
      y: avgY,
      fontSize: avgFontSize,
      height: maxHeight,
      segments,
    };
  });

  lines.sort((a, b) => a.y - b.y);
  return lines;
}

/**
 * Agrupa ítems verticales (-90° o 90°) por columna X y cercanía en Y, respetando bordes horizontales
 */
function buildVerticalSegments(
  items: RawTextItem[],
  ctx: CanvasRenderingContext2D | null,
  canvasWidth: number,
  canvasHeight: number,
  renderScale: number
): LineSegment[] {
  const validItems = items
    .map((it) => ({ ...it, text: it.text.replace(/\s+/g, ' ').trim() }))
    .filter((it) => it.text.length > 0);

  if (validItems.length === 0) return [];

  // Agrupar por columna X y misma rotación
  const columns: RawTextItem[][] = [];
  for (const item of validItems) {
    let matched: RawTextItem[] | null = null;
    for (const col of columns) {
      const ref = col[0];
      if (ref.rotation === item.rotation && Math.abs(item.x - ref.x) <= Math.max(ref.fontSize, item.fontSize) * 0.6) {
        matched = col;
        break;
      }
    }
    if (matched) {
      matched.push(item);
    } else {
      columns.push([item]);
    }
  }

  const results: LineSegment[] = [];
  for (const col of columns) {
    const rot = col[0].rotation;
    // Si es -90° (se lee de abajo hacia arriba), ordenar por Y descendente; si es 90° (arriba hacia abajo), Y ascendente
    col.sort((a, b) => (rot === -90 ? b.y - a.y : a.y - b.y));

    for (const item of col) {
      const last = results[results.length - 1];
      const sameCol =
        last &&
        last.rotation === rot &&
        Math.abs(last.x - item.x) <= Math.max(last.fontSize, item.fontSize) * 0.6;

      let gap = Infinity;
      let sepY1 = 0;
      let sepY2 = 0;
      if (sameCol) {
        if (rot === -90) {
          // last está abajo, item está arriba
          gap = last.y - (item.y + item.height);
          sepY1 = item.y + item.height;
          sepY2 = last.y;
        } else {
          // last está arriba, item está abajo
          gap = item.y - last.yEnd;
          sepY1 = last.yEnd;
          sepY2 = item.y;
        }
      }

      const gapThreshold = Math.max(item.fontSize * 0.72, 6.5);
      const blocked =
        sameCol && gap > 0.5 && gap <= gapThreshold * 1.4
          ? hasHorizontalSeparatorLine(
              ctx,
              canvasWidth,
              canvasHeight,
              Math.min(last.x, item.x),
              Math.max(last.xEnd - last.x, item.width),
              sepY1,
              sepY2,
              renderScale
            )
          : false;

      if (sameCol && gap <= gapThreshold && gap >= -item.fontSize * 0.4 && !blocked) {
        const needsSpace = gap > item.fontSize * 0.1;
        last.text = `${last.text}${needsSpace ? ' ' : ''}${item.text}`;
        last.x = Math.min(last.x, item.x);
        last.xEnd = Math.max(last.xEnd, item.x + item.width);
        last.y = Math.min(last.y, item.y);
        last.yEnd = Math.max(last.yEnd, item.y + item.height);
        last.height = last.yEnd - last.y;
        last.isBold = last.isBold || item.isBold;
      } else {
        results.push({
          text: item.text,
          x: item.x,
          xEnd: item.x + item.width,
          y: item.y,
          yEnd: item.y + item.height,
          height: item.height,
          fontSize: item.fontSize,
          isBold: item.isBold,
          rotation: rot,
        });
      }
    }
  }

  return results;
}

/**
 * Construye todos los EditableField (horizontales, verticales y rotados) + tablas detectadas
 */
function buildFieldsAndTables(
  items: RawTextItem[],
  ctx: CanvasRenderingContext2D | null,
  canvasWidth: number,
  canvasHeight: number,
  renderScale: number
): { fields: EditableField[]; tables: DetectedTable[] } {
  const horizontalItems = items.filter((it) => it.rotation === 0);
  const verticalItems = items.filter((it) => it.rotation === -90 || it.rotation === 90);
  const otherAngleItems = items.filter(
    (it) => it.rotation !== 0 && it.rotation !== -90 && it.rotation !== 90
  );

  const lines = buildHorizontalLines(horizontalItems, ctx, canvasWidth, canvasHeight, renderScale);
  const verticalSegments = buildVerticalSegments(
    verticalItems,
    ctx,
    canvasWidth,
    canvasHeight,
    renderScale
  );

  const fields: EditableField[] = [];
  const tables: DetectedTable[] = [];

  // 1. Crear campos horizontales
  const lineFieldMap: EditableField[][] = lines.map((line) =>
    line.segments.map((seg) => {
      const width = Math.max(12, seg.xEnd - seg.x);
      const height = Math.max(seg.fontSize * 1.18, seg.height);
      const { bgColor, textColor } = sampleColorsAtBox(
        ctx,
        canvasWidth,
        canvasHeight,
        seg.x,
        seg.y,
        width,
        height,
        renderScale
      );

      const field: EditableField = {
        id: createId('fld'),
        originalText: seg.text,
        text: seg.text,
        x: seg.x,
        y: seg.y,
        width,
        height,
        fontSize: Math.max(6, Math.min(36, Math.round(seg.fontSize * 10) / 10)),
        isBold: seg.isBold,
        align: looksNumeric(seg.text) ? 'right' : 'left',
        rotation: 0,
        bgColor,
        textColor,
        isModified: false,
      };
      fields.push(field);
      return field;
    })
  );

  // 2. Crear campos verticales (-90° / 90°)
  for (const seg of verticalSegments) {
    const width = Math.max(seg.fontSize * 1.18, seg.xEnd - seg.x);
    const height = Math.max(14, seg.yEnd - seg.y);
    const { bgColor, textColor } = sampleColorsAtBox(
      ctx,
      canvasWidth,
      canvasHeight,
      seg.x,
      seg.y,
      width,
      height,
      renderScale
    );

    fields.push({
      id: createId('vfld'),
      originalText: seg.text,
      text: seg.text,
      x: seg.x,
      y: seg.y,
      width,
      height,
      fontSize: Math.max(6, Math.min(36, Math.round(seg.fontSize * 10) / 10)),
      isBold: seg.isBold,
      align: 'center',
      rotation: seg.rotation,
      bgColor,
      textColor,
      isModified: false,
    });
  }

  // 3. Crear campos con otros ángulos (180°, diagonales, etc.)
  for (const it of otherAngleItems) {
    const text = it.text.replace(/\s+/g, ' ').trim();
    if (!text) continue;
    const { bgColor, textColor } = sampleColorsAtBox(
      ctx,
      canvasWidth,
      canvasHeight,
      it.x,
      it.y,
      it.width,
      it.height,
      renderScale
    );

    fields.push({
      id: createId('rfld'),
      originalText: text,
      text,
      x: it.x,
      y: it.y,
      width: Math.max(14, it.width),
      height: Math.max(12, it.height),
      fontSize: Math.max(6, Math.min(36, Math.round(it.fontSize * 10) / 10)),
      isBold: it.isBold,
      align: 'left',
      rotation: it.rotation,
      bgColor,
      textColor,
      isModified: false,
    });
  }

  // 4. Detectar tablas sobre las líneas horizontales
  let i = 0;
  while (i < lines.length) {
    if (lines[i].segments.length >= 2) {
      const groupIndices: number[] = [i];
      let j = i + 1;
      while (j < lines.length) {
        const prev = lines[groupIndices[groupIndices.length - 1]];
        const next = lines[j];
        const dy = next.y - prev.y;
        const maxGap = Math.max(prev.fontSize * 3.2, 34);
        if (dy <= maxGap && next.segments.length >= 2) {
          groupIndices.push(j);
          j += 1;
        } else {
          break;
        }
      }

      if (groupIndices.length >= 2) {
        const tableFields = groupIndices.flatMap((idx) => lineFieldMap[idx]);
        const intervals = tableFields
          .map((f) => ({ start: f.x, end: f.x + f.width }))
          .sort((a, b) => a.start - b.start);

        const avgFs =
          groupIndices.reduce((s, idx) => s + lines[idx].fontSize, 0) / groupIndices.length;
        const mergeTol = Math.max(avgFs * 0.4, 3.5);
        const cols: { start: number; end: number }[] = [];

        for (const iv of intervals) {
          const last = cols[cols.length - 1];
          if (last && iv.start <= last.end + mergeTol) {
            last.end = Math.max(last.end, iv.end);
          } else {
            cols.push({ start: iv.start, end: iv.end });
          }
        }

        if (cols.length >= 2) {
          const tableId = createId('tbl');
          const cellFieldIds: (string | null)[][] = [];

          groupIndices.forEach((lineIdx, rIdx) => {
            const rowIds: (string | null)[] = new Array(cols.length).fill(null);
            for (const fld of lineFieldMap[lineIdx]) {
              let bestCol = 0;
              let bestScore = -Infinity;
              cols.forEach((col, cIdx) => {
                const overlap = Math.min(fld.x + fld.width, col.end) - Math.max(fld.x, col.start);
                const dist = Math.abs(fld.x + fld.width / 2 - (col.start + col.end) / 2);
                const score = overlap > 0 ? overlap + 1000 - dist * 0.1 : -dist;
                if (score > bestScore) {
                  bestScore = score;
                  bestCol = cIdx;
                }
              });

              fld.tableId = tableId;
              fld.rowIndex = rIdx;
              fld.colIndex = bestCol;
              if (!rowIds[bestCol]) {
                rowIds[bestCol] = fld.id;
              }
            }
            cellFieldIds.push(rowIds);
          });

          const rawWidths = cols.map((col, idx) => {
            const nextStart = idx < cols.length - 1 ? cols[idx + 1].start : col.end;
            return Math.max(col.end - col.start, nextStart - col.start, 24);
          });
          const totalSpan = rawWidths.reduce((s, w) => s + w, 0) || 1;
          const columnWidths = rawWidths.map((w) => Math.max(6, Math.round((w / totalSpan) * 100)));
          const sumPct = columnWidths.reduce((s, w) => s + w, 0);
          if (sumPct !== 100 && columnWidths.length > 0) {
            columnWidths[columnWidths.length - 1] += 100 - sumPct;
          }

          const firstRowFields = lineFieldMap[groupIndices[0]];
          const headerBg =
            firstRowFields.find((f) => f.bgColor !== '#ffffff')?.bgColor ?? '#1e3a8a';

          tables.push({
            id: tableId,
            rowCount: groupIndices.length,
            colCount: cols.length,
            columnWidths,
            cellFieldIds,
            y: lines[groupIndices[0]].y,
            hasHeader: true,
            headerBg,
          });

          i = j;
          continue;
        }
      }
    }
    i += 1;
  }

  return { fields, tables };
}

/**
 * Recorta el área superior con logo/membrete si contiene gráficos relevantes
 */
function extractTopLogoBanner(
  canvas: HTMLCanvasElement,
  pageWidth: number,
  pageHeight: number,
  tables: DetectedTable[],
  renderScale: number
): { dataUrl?: string; widthPt?: number; heightPt?: number } {
  const firstTableY = tables.length > 0 ? Math.min(...tables.map((t) => t.y)) : pageHeight * 0.2;
  const bannerHeightPt = Math.min(Math.max(firstTableY - 12, 0), pageHeight * 0.22);
  if (bannerHeightPt < 36) return {};

  const cropH = Math.min(canvas.height, Math.round(bannerHeightPt * renderScale));
  const cropW = canvas.width;
  if (cropH < 30 || cropW < 30) return {};

  const tempCanvas = document.createElement('canvas');
  tempCanvas.width = cropW;
  tempCanvas.height = cropH;
  const tCtx = tempCanvas.getContext('2d');
  if (!tCtx) return {};

  tCtx.drawImage(canvas, 0, 0, cropW, cropH, 0, 0, cropW, cropH);

  const imgData = tCtx.getImageData(0, 0, cropW, cropH).data;
  let coloredPixels = 0;
  for (let p = 0; p < imgData.length; p += 16) {
    const r = imgData[p];
    const g = imgData[p + 1];
    const b = imgData[p + 2];
    const maxDiff = Math.max(Math.abs(r - g), Math.abs(g - b), Math.abs(r - b));
    if (maxDiff > 22) {
      coloredPixels += 1;
    }
  }

  if (coloredPixels < 120) return {};

  return {
    dataUrl: tempCanvas.toDataURL('image/png'),
    widthPt: pageWidth,
    heightPt: bannerHeightPt,
  };
}

/**
 * Preprocesa la imagen de una página (escala de grises + contraste) para mejorar la precisión de Tesseract OCR
 */
function createHighContrastOcrDataUrl(sourceCanvas: HTMLCanvasElement): string {
  const out = document.createElement('canvas');
  out.width = sourceCanvas.width;
  out.height = sourceCanvas.height;
  const ctx = out.getContext('2d');
  if (!ctx) return sourceCanvas.toDataURL('image/png');

  ctx.drawImage(sourceCanvas, 0, 0);
  const img = ctx.getImageData(0, 0, out.width, out.height);
  const d = img.data;
  for (let i = 0; i < d.length; i += 4) {
    const gray = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
    // Aumentar contraste suave para texto nítido sin perder texto sobre encabezados oscuros
    const contrasted = Math.max(0, Math.min(255, (gray - 128) * 1.35 + 128));
    d[i] = contrasted;
    d[i + 1] = contrasted;
    d[i + 2] = contrasted;
  }
  ctx.putImageData(img, 0, 0);
  return out.toDataURL('image/png');
}

/**
 * Ejecuta OCR con Tesseract.js y detecta también palabras verticales (cuando altura >> ancho)
 */
async function extractItemsWithOcr(
  imageDataUrl: string,
  pageWidth: number,
  pageHeight: number,
  imgWidth: number,
  imgHeight: number
): Promise<RawTextItem[]> {
  const result = await Tesseract.recognize(imageDataUrl, 'spa+eng');
  const scaleX = pageWidth / imgWidth;
  const scaleY = pageHeight / imgHeight;

  const items: RawTextItem[] = [];
  const data = result.data as unknown as {
    blocks?: Array<{
      paragraphs?: Array<{
        lines?: Array<{
          words?: Array<{
            text: string;
            bbox: { x0: number; y0: number; x1: number; y1: number };
          }>;
        }>;
      }>;
    }>;
  };

  for (const block of data.blocks ?? []) {
    for (const para of block.paragraphs ?? []) {
      for (const line of para.lines ?? []) {
        for (const word of line.words ?? []) {
          const text = word.text?.trim();
          if (!text) continue;
          const x = word.bbox.x0 * scaleX;
          const y = word.bbox.y0 * scaleY;
          const width = Math.max(4, (word.bbox.x1 - word.bbox.x0) * scaleX);
          const height = Math.max(8, (word.bbox.y1 - word.bbox.y0) * scaleY);

          // Si una palabra de >= 3 caracteres es mucho más alta que ancha, es texto vertical
          const isVerticalWord = text.length >= 3 && height > width * 1.9;

          items.push({
            text,
            x,
            y,
            width,
            height,
            fontSize: Math.max(7, Math.min(28, (isVerticalWord ? width : height) * 0.82)),
            isBold: false,
            rotation: isVerticalWord ? -90 : 0,
          });
        }
      }
    }
  }

  return items;
}

/**
 * Escanea un archivo PDF usando la matriz exacta del viewport de pdfjs-dist
 * (maneja páginas horizontales, verticales, rotadas 90/180/270 y texto vertical/inclinado).
 */
export async function scanPdfFile(
  file: File,
  onProgress?: (progress: ScanProgress) => void
): Promise<ScannedDocument> {
  onProgress?.({
    stage: 'loading',
    currentPage: 0,
    totalPages: 0,
    message: 'Cargando archivo PDF original...',
  });

  const arrayBuffer = await file.arrayBuffer();
  const originalPdfBytes = new Uint8Array(arrayBuffer.slice(0));
  const pdfJsBytes = new Uint8Array(arrayBuffer.slice(0));

  const loadingTask = pdfjsLib.getDocument({ data: pdfJsBytes });
  const pdf = await loadingTask.promise;
  const totalPages = pdf.numPages;

  const pages: ScannedPage[] = [];
  const renderScale = 2.0;

  for (let pageNum = 1; pageNum <= totalPages; pageNum += 1) {
    onProgress?.({
      stage: 'parsing',
      currentPage: pageNum,
      totalPages,
      message: `Analizando página ${pageNum} de ${totalPages} (orientación, tablas y texto)...`,
    });

    const page = await pdf.getPage(pageNum);
    const baseViewport = page.getViewport({ scale: 1 });
    const renderViewport = page.getViewport({ scale: renderScale });

    const canvas = document.createElement('canvas');
    canvas.width = Math.round(renderViewport.width);
    canvas.height = Math.round(renderViewport.height);
    const ctx = canvas.getContext('2d', { willReadFrequently: true });

    if (ctx) {
      await page.render({
        canvasContext: ctx,
        viewport: renderViewport,
        canvas,
      }).promise;
    }

    const originalImageUrl = canvas.toDataURL('image/png');

    // Extraer texto nativo transformando CADA ítem con la matriz exacta del viewport
    const textContent = await page.getTextContent();
    const styles = textContent.styles as Record<string, { fontFamily?: string }>;
    const rawItems: RawTextItem[] = [];

    for (const item of textContent.items) {
      if (!('str' in item)) continue;
      const str = item.str;
      if (!str || !str.trim()) continue;

      // Multiplicar viewport.transform × item.transform -> coordenadas exactas en pantalla (maneja horizontal/rotado/CropBox)
      const m = pdfjsLib.Util.transform(baseViewport.transform, item.transform);
      const fontSize = Math.hypot(m[2], m[3]) || Math.hypot(m[0], m[1]) || item.height || 10;
      const angleDeg = normalizeAngleDeg((Math.atan2(m[1], m[0]) * 180) / Math.PI);
      const advanceLen =
        item.width > 0 ? item.width : Math.max(fontSize * 0.6, str.length * fontSize * 0.5);

      let x = m[4];
      let y = m[5] - fontSize * 0.85;
      let width = advanceLen;
      let height = Math.max(fontSize * 1.15, item.height || fontSize);

      if (angleDeg === -90) {
        // Texto vertical leyendo de abajo hacia arriba
        width = Math.max(fontSize * 1.18, item.height || fontSize);
        height = advanceLen;
        x = m[4] - fontSize * 0.85;
        y = m[5] - advanceLen;
      } else if (angleDeg === 90) {
        // Texto vertical leyendo de arriba hacia abajo
        width = Math.max(fontSize * 1.18, item.height || fontSize);
        height = advanceLen;
        x = m[4] - fontSize * 0.2;
        y = m[5];
      } else if (angleDeg === 180) {
        // Texto invertido
        x = m[4] - advanceLen;
        y = m[5] - fontSize * 0.2;
      }

      const fontStyle = styles?.[item.fontName];
      rawItems.push({
        text: str,
        x,
        y,
        width,
        height,
        fontSize,
        isBold: detectBoldFont(item.fontName ?? '', fontStyle?.fontFamily),
        rotation: angleDeg,
      });
    }

    const charCount = rawItems.reduce((s, it) => s + it.text.trim().length, 0);
    let extractionMode: 'text' | 'ocr' = 'text';
    let itemsToUse = rawItems;

    // Si es PDF escaneado como foto/imagen, ejecutar OCR con preprocesado de contraste
    if (charCount < 15) {
      extractionMode = 'ocr';
      onProgress?.({
        stage: 'ocr',
        currentPage: pageNum,
        totalPages,
        message: `Página ${pageNum} escaneada como imagen. Ejecutando OCR (Tesseract)...`,
      });

      try {
        const ocrImage = createHighContrastOcrDataUrl(canvas);
        itemsToUse = await extractItemsWithOcr(
          ocrImage,
          baseViewport.width,
          baseViewport.height,
          canvas.width,
          canvas.height
        );
      } catch (err) {
        console.error('Error ejecutando OCR:', err);
      }
    }

    const { fields, tables } = buildFieldsAndTables(
      itemsToUse,
      ctx,
      canvas.width,
      canvas.height,
      renderScale
    );

    const banner = extractTopLogoBanner(
      canvas,
      baseViewport.width,
      baseViewport.height,
      tables,
      renderScale
    );

    const vt = baseViewport.transform as [number, number, number, number, number, number];

    pages.push({
      pageNumber: pageNum,
      width: Math.round(baseViewport.width * 100) / 100,
      height: Math.round(baseViewport.height * 100) / 100,
      pageRotation: page.rotate ?? 0,
      viewportTransform: [vt[0], vt[1], vt[2], vt[3], vt[4], vt[5]],
      originalImageUrl,
      logoBannerDataUrl: banner.dataUrl,
      logoBannerWidth: banner.widthPt,
      logoBannerHeight: banner.heightPt,
      extractionMode,
      fields,
      tables,
    });
  }

  onProgress?.({
    stage: 'done',
    currentPage: totalPages,
    totalPages,
    message: 'Escaneo completado.',
  });

  return {
    fileName: file.name,
    originalPdfBytes,
    pages,
  };
}

/**
 * Re-escanea una página con Tesseract OCR (incluye preprocesado de contraste)
 */
export async function rescanPageWithOcr(page: ScannedPage): Promise<ScannedPage> {
  const img = await new Promise<HTMLImageElement>((resolve, reject) => {
    const el = new Image();
    el.onload = () => resolve(el);
    el.onerror = reject;
    el.src = page.originalImageUrl;
  });

  const canvas = document.createElement('canvas');
  canvas.width = img.naturalWidth || Math.round(page.width * 2);
  canvas.height = img.naturalHeight || Math.round(page.height * 2);
  const ctx = canvas.getContext('2d', { willReadFrequently: true });
  if (ctx) {
    ctx.drawImage(img, 0, 0);
  }

  const renderScale = canvas.width / page.width;
  const ocrImage = createHighContrastOcrDataUrl(canvas);
  const ocrItems = await extractItemsWithOcr(
    ocrImage,
    page.width,
    page.height,
    canvas.width,
    canvas.height
  );

  const { fields, tables } = buildFieldsAndTables(
    ocrItems,
    ctx,
    canvas.width,
    canvas.height,
    renderScale
  );

  return {
    ...page,
    extractionMode: 'ocr',
    fields,
    tables,
  };
}

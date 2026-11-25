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

interface VisualTableGrid {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  colLines: number[]; // coordenadas X en puntos PDF ordenadas
  rowLines: number[]; // coordenadas Y en puntos PDF ordenadas
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
  const a = (((deg + 180) % 360) + 360) % 360 - 180;
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
 * Agrupa coordenadas numéricas cercanas (por ejemplo, líneas de borde de 2-3 px de grosor)
 */
function clusterPositions(values: number[], tolerance: number): number[] {
  if (values.length === 0) return [];
  const sorted = [...values].sort((a, b) => a - b);
  const clusters: number[][] = [[sorted[0]]];
  for (let i = 1; i < sorted.length; i += 1) {
    const last = clusters[clusters.length - 1];
    if (sorted[i] - last[last.length - 1] <= tolerance) {
      last.push(sorted[i]);
    } else {
      clusters.push([sorted[i]]);
    }
  }
  return clusters.map((c) => c.reduce((s, v) => s + v, 0) / c.length);
}

/**
 * Detecta rejillas reales de tablas (líneas horizontales y verticales dibujadas) directamente en el canvas.
 * Funciona tanto en PDFs digitales como en PDFs escaneados con tablas con bordes.
 */
function detectVisualTableGrids(
  ctx: CanvasRenderingContext2D | null,
  canvasWidth: number,
  canvasHeight: number,
  renderScale: number
): VisualTableGrid[] {
  if (!ctx || canvasWidth < 40 || canvasHeight < 40) return [];

  const imgData = ctx.getImageData(0, 0, canvasWidth, canvasHeight).data;
  const isDark = (x: number, y: number): boolean => {
    const idx = (y * canvasWidth + x) * 4;
    return relativeLuminance(imgData[idx], imgData[idx + 1], imgData[idx + 2]) < 0.82;
  };

  // 1. Detectar segmentos horizontales largos (>= 22% del ancho de página)
  const minHorizLen = Math.max(80, Math.round(canvasWidth * 0.22));
  const hSegments: { y: number; x1: number; x2: number }[] = [];

  for (let y = 2; y < canvasHeight - 2; y += 1) {
    let runStart = -1;
    let gapAllowed = 0;
    for (let x = 0; x < canvasWidth; x += 1) {
      if (isDark(x, y)) {
        if (runStart < 0) runStart = x;
        gapAllowed = 0;
      } else if (runStart >= 0) {
        gapAllowed += 1;
        if (gapAllowed > 2) {
          const runEnd = x - gapAllowed;
          if (runEnd - runStart >= minHorizLen) {
            hSegments.push({ y, x1: runStart, x2: runEnd });
          }
          runStart = -1;
          gapAllowed = 0;
        }
      }
    }
    if (runStart >= 0 && canvasWidth - 1 - runStart >= minHorizLen) {
      hSegments.push({ y, x1: runStart, x2: canvasWidth - 1 });
    }
  }

  if (hSegments.length < 3) return [];

  // Fusionar líneas horizontales adyacentes (grosor de borde 1-4 px)
  const mergedH: { y: number; x1: number; x2: number }[] = [];
  for (const seg of hSegments) {
    const prev = mergedH[mergedH.length - 1];
    if (
      prev &&
      seg.y - prev.y <= 4 &&
      Math.abs(seg.x1 - prev.x1) <= 18 &&
      Math.abs(seg.x2 - prev.x2) <= 18
    ) {
      prev.y = (prev.y + seg.y) / 2;
      prev.x1 = Math.min(prev.x1, seg.x1);
      prev.x2 = Math.max(prev.x2, seg.x2);
    } else {
      mergedH.push({ ...seg });
    }
  }

  // Agrupar líneas horizontales que pertenecen a la misma tabla (mismo rango X y separación razonable)
  const hGroups: { y: number; x1: number; x2: number }[][] = [];
  for (const line of mergedH) {
    let placed = false;
    for (const grp of hGroups) {
      const ref = grp[grp.length - 1];
      const overlap = Math.min(line.x2, ref.x2) - Math.max(line.x1, ref.x1);
      const minW = Math.min(line.x2 - line.x1, ref.x2 - ref.x1);
      const dy = line.y - ref.y;
      if (overlap >= minW * 0.75 && dy > 6 && dy <= canvasHeight * 0.25) {
        grp.push(line);
        placed = true;
        break;
      }
    }
    if (!placed) {
      hGroups.push([line]);
    }
  }

  const grids: VisualTableGrid[] = [];

  for (const grp of hGroups) {
    if (grp.length < 3) continue; // Mínimo 2 filas (3 líneas horizontales)

    const topPx = Math.round(grp[0].y);
    const botPx = Math.round(grp[grp.length - 1].y);
    const leftPx = Math.round(Math.min(...grp.map((g) => g.x1)));
    const rightPx = Math.round(Math.max(...grp.map((g) => g.x2)));
    const tableHeightPx = botPx - topPx;
    if (tableHeightPx < 24) continue;

    // 2. Buscar líneas verticales dentro del rectángulo [leftPx..rightPx, topPx..botPx]
    const vCandidates: number[] = [];
    for (let x = Math.max(0, leftPx - 4); x <= Math.min(canvasWidth - 1, rightPx + 4); x += 1) {
      let darkPixels = 0;
      for (let y = topPx; y <= botPx; y += 1) {
        if (isDark(x, y)) darkPixels += 1;
      }
      // Si >= 55% de la altura de la tabla tiene píxeles oscuros en esta X, es un separador de columna
      if (darkPixels / Math.max(1, tableHeightPx) >= 0.55) {
        vCandidates.push(x);
      }
    }

    const vLinesPx = clusterPositions(vCandidates, 5);
    if (vLinesPx.length < 3) continue; // Mínimo 2 columnas (3 líneas verticales)

    // Asegurar bordes izquierdo y derecho incluidos
    const colLinesPt = vLinesPx.map((x) => x / renderScale);
    const rowLinesPt = clusterPositions(
      grp.map((g) => g.y / renderScale),
      3
    );

    if (colLinesPt.length >= 3 && rowLinesPt.length >= 3) {
      grids.push({
        x1: colLinesPt[0],
        y1: rowLinesPt[0],
        x2: colLinesPt[colLinesPt.length - 1],
        y2: rowLinesPt[rowLinesPt.length - 1],
        colLines: colLinesPt,
        rowLines: rowLinesPt,
      });
    }
  }

  return grids;
}

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

      const blockedByBorder =
        last && gap > 0.5 && gap <= gapThreshold * 1.6
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

  const columns: RawTextItem[][] = [];
  for (const item of validItems) {
    let matched: RawTextItem[] | null = null;
    for (const col of columns) {
      const ref = col[0];
      if (
        ref.rotation === item.rotation &&
        Math.abs(item.x - ref.x) <= Math.max(ref.fontSize, item.fontSize) * 0.6
      ) {
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
          gap = last.y - (item.y + item.height);
          sepY1 = item.y + item.height;
          sepY2 = last.y;
        } else {
          gap = item.y - last.yEnd;
          sepY1 = last.yEnd;
          sepY2 = item.y;
        }
      }

      const gapThreshold = Math.max(item.fontSize * 0.72, 6.5);
      const blocked =
        sameCol && gap > 0.5 && gap <= gapThreshold * 1.5
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
 * Construye todos los EditableField + detecta tablas con máxima precisión:
 * 1) Primero usa las rejillas visuales reales (líneas horizontales y verticales) detectadas en el canvas.
 * 2) Para tablas sin bordes visibles, usa clustering inteligente de columnas tolerando celdas vacías/multilínea.
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

  // 3. Crear campos con otros ángulos
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

  // 4A. Detección exacta de tablas por rejilla visual (bordes reales dibujados en el PDF/escaneo)
  const assignedFieldIds = new Set<string>();
  const visualGrids = detectVisualTableGrids(ctx, canvasWidth, canvasHeight, renderScale);

  for (const grid of visualGrids) {
    const rowCount = grid.rowLines.length - 1;
    const colCount = grid.colLines.length - 1;
    if (rowCount < 2 || colCount < 2) continue;

    const tableId = createId('tbl');
    const cellFieldIds: string[][][] = Array.from({ length: rowCount }, () =>
      Array.from({ length: colCount }, () => [])
    );

    let matchedFields = 0;

    for (const fld of fields) {
      if (assignedFieldIds.has(fld.id)) continue;
      const cx = fld.x + fld.width / 2;
      const cy = fld.y + fld.height / 2;

      if (cx < grid.x1 - 3 || cx > grid.x2 + 3 || cy < grid.y1 - 3 || cy > grid.y2 + 3) {
        continue;
      }

      // Encontrar fila exacta
      let rIdx = -1;
      for (let r = 0; r < rowCount; r += 1) {
        if (cy >= grid.rowLines[r] - 2 && cy <= grid.rowLines[r + 1] + 2) {
          rIdx = r;
          break;
        }
      }
      // Encontrar columna exacta
      let cIdx = -1;
      for (let c = 0; c < colCount; c += 1) {
        if (cx >= grid.colLines[c] - 2 && cx <= grid.colLines[c + 1] + 2) {
          cIdx = c;
          break;
        }
      }

      if (rIdx >= 0 && cIdx >= 0) {
        cellFieldIds[rIdx][cIdx].push(fld.id);
        fld.tableId = tableId;
        fld.rowIndex = rIdx;
        fld.colIndex = cIdx;
        assignedFieldIds.add(fld.id);
        matchedFields += 1;
      }
    }

    if (matchedFields < 3) continue;

    // Filtrar filas completamente vacías al inicio/final si existieran por ruido de borde
    const nonEmptyRows = cellFieldIds.filter((row) => row.some((cell) => cell.length > 0));
    if (nonEmptyRows.length < 2) continue;

    const totalW = Math.max(1, grid.x2 - grid.x1);
    const columnWidths = Array.from({ length: colCount }, (_, c) =>
      Math.max(4, Math.round(((grid.colLines[c + 1] - grid.colLines[c]) / totalW) * 100))
    );
    const sumPct = columnWidths.reduce((s, v) => s + v, 0);
    if (sumPct !== 100 && columnWidths.length > 0) {
      columnWidths[columnWidths.length - 1] += 100 - sumPct;
    }

    const firstRowIds = nonEmptyRows[0].flat();
    const headerField = fields.find(
      (f) => firstRowIds.includes(f.id) && f.bgColor !== '#ffffff'
    );

    tables.push({
      id: tableId,
      rowCount: nonEmptyRows.length,
      colCount,
      columnWidths,
      cellFieldIds: nonEmptyRows,
      y: grid.y1,
      hasHeader: true,
      headerBg: headerField?.bgColor ?? '#1e3a8a',
      detectionMethod: 'grid',
    });
  }

  // 4B. Respaldo para tablas SIN bordes dibujados (clustering inteligente de columnas tolerando filas con 1 celda)
  const unassignedLineIndices: number[] = [];
  lineFieldMap.forEach((rowFields, idx) => {
    const remaining = rowFields.filter((f) => !assignedFieldIds.has(f.id));
    if (remaining.length > 0) {
      unassignedLineIndices.push(idx);
    }
  });

  let ptr = 0;
  while (ptr < unassignedLineIndices.length) {
    const lineIdx = unassignedLineIndices[ptr];
    const rowFlds = lineFieldMap[lineIdx].filter((f) => !assignedFieldIds.has(f.id));

    if (rowFlds.length >= 2) {
      const groupIndices: number[] = [lineIdx];
      let nextPtr = ptr + 1;

      while (nextPtr < unassignedLineIndices.length) {
        const prevIdx = groupIndices[groupIndices.length - 1];
        const candIdx = unassignedLineIndices[nextPtr];
        const dy = lines[candIdx].y - lines[prevIdx].y;
        const maxGap = Math.max(lines[prevIdx].fontSize * 3.4, 36);
        const candFlds = lineFieldMap[candIdx].filter((f) => !assignedFieldIds.has(f.id));

        if (dy > maxGap) break;

        if (candFlds.length >= 2) {
          groupIndices.push(candIdx);
          nextPtr += 1;
        } else if (
          candFlds.length === 1 &&
          nextPtr + 1 < unassignedLineIndices.length &&
          lines[unassignedLineIndices[nextPtr + 1]].y - lines[candIdx].y <= maxGap &&
          lineFieldMap[unassignedLineIndices[nextPtr + 1]].filter(
            (f) => !assignedFieldIds.has(f.id)
          ).length >= 2
        ) {
          // Permitir filas intermedias con 1 celda sin romper la tabla
          groupIndices.push(candIdx);
          nextPtr += 1;
        } else {
          break;
        }
      }

      if (groupIndices.length >= 2) {
        const tableFields = groupIndices.flatMap((idx) =>
          lineFieldMap[idx].filter((f) => !assignedFieldIds.has(f.id))
        );
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
          const cellFieldIds: string[][][] = [];

          groupIndices.forEach((gIdx, rIdx) => {
            const rowCells: string[][] = Array.from({ length: cols.length }, () => []);
            const fldsInLine = lineFieldMap[gIdx].filter((f) => !assignedFieldIds.has(f.id));
            for (const fld of fldsInLine) {
              let bestCol = 0;
              let bestScore = -Infinity;
              cols.forEach((col, cIdx) => {
                const overlap =
                  Math.min(fld.x + fld.width, col.end) - Math.max(fld.x, col.start);
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
              assignedFieldIds.add(fld.id);
              rowCells[bestCol].push(fld.id);
            }
            cellFieldIds.push(rowCells);
          });

          const rawWidths = cols.map((col, idx) => {
            const nextStart = idx < cols.length - 1 ? cols[idx + 1].start : col.end;
            return Math.max(col.end - col.start, nextStart - col.start, 24);
          });
          const totalSpan = rawWidths.reduce((s, w) => s + w, 0) || 1;
          const columnWidths = rawWidths.map((w) => Math.max(5, Math.round((w / totalSpan) * 100)));
          const sumPct = columnWidths.reduce((s, w) => s + w, 0);
          if (sumPct !== 100 && columnWidths.length > 0) {
            columnWidths[columnWidths.length - 1] += 100 - sumPct;
          }

          const firstRowFlds = lineFieldMap[groupIndices[0]];
          const headerBg =
            firstRowFlds.find((f) => f.bgColor !== '#ffffff')?.bgColor ?? '#1e3a8a';

          tables.push({
            id: tableId,
            rowCount: cellFieldIds.length,
            colCount: cols.length,
            columnWidths,
            cellFieldIds,
            y: lines[groupIndices[0]].y,
            hasHeader: true,
            headerBg,
            detectionMethod: 'alignment',
          });

          ptr = nextPtr;
          continue;
        }
      }
    }
    ptr += 1;
  }

  tables.sort((a, b) => a.y - b.y);
  return { fields, tables };
}

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
    const contrasted = Math.max(0, Math.min(255, (gray - 128) * 1.45 + 128));
    d[i] = contrasted;
    d[i + 1] = contrasted;
    d[i + 2] = contrasted;
  }
  ctx.putImageData(img, 0, 0);
  return out.toDataURL('image/png');
}

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
 * Escanea un archivo PDF (digital o escaneado como imagen) usando pdfjs-dist + Tesseract OCR
 * y detector de rejillas reales de tablas.
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
      message: `Analizando página ${pageNum} de ${totalPages} (rejillas de tablas y texto)...`,
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

    const textContent = await page.getTextContent();
    const styles = textContent.styles as Record<string, { fontFamily?: string }>;
    const rawItems: RawTextItem[] = [];

    for (const item of textContent.items) {
      if (!('str' in item)) continue;
      const str = item.str;
      if (!str || !str.trim()) continue;

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
        width = Math.max(fontSize * 1.18, item.height || fontSize);
        height = advanceLen;
        x = m[4] - fontSize * 0.85;
        y = m[5] - advanceLen;
      } else if (angleDeg === 90) {
        width = Math.max(fontSize * 1.18, item.height || fontSize);
        height = advanceLen;
        x = m[4] - fontSize * 0.2;
        y = m[5];
      } else if (angleDeg === 180) {
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

    // Si es un PDF escaneado (foto/imagen sin texto seleccionable), ejecutar OCR en alta resolución (2.5x)
    if (charCount < 15) {
      extractionMode = 'ocr';
      onProgress?.({
        stage: 'ocr',
        currentPage: pageNum,
        totalPages,
        message: `Página ${pageNum} detectada como escaneo/imagen. Reconociendo texto con OCR...`,
      });

      try {
        const ocrScale = 2.5;
        const ocrViewport = page.getViewport({ scale: ocrScale });
        const ocrCanvas = document.createElement('canvas');
        ocrCanvas.width = Math.round(ocrViewport.width);
        ocrCanvas.height = Math.round(ocrViewport.height);
        const ocrCtx = ocrCanvas.getContext('2d', { willReadFrequently: true });
        if (ocrCtx) {
          await page.render({
            canvasContext: ocrCtx,
            viewport: ocrViewport,
            canvas: ocrCanvas,
          }).promise;
        }
        const ocrImage = createHighContrastOcrDataUrl(ocrCanvas);
        itemsToUse = await extractItemsWithOcr(
          ocrImage,
          baseViewport.width,
          baseViewport.height,
          ocrCanvas.width,
          ocrCanvas.height
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

    const vt = baseViewport.transform as [number, number, number, number, number, number];

    pages.push({
      pageNumber: pageNum,
      width: Math.round(baseViewport.width * 100) / 100,
      height: Math.round(baseViewport.height * 100) / 100,
      pageRotation: page.rotate ?? 0,
      viewportTransform: [vt[0], vt[1], vt[2], vt[3], vt[4], vt[5]],
      originalImageUrl,
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

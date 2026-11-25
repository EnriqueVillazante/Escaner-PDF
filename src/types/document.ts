// Tipos principales para el editor PDF con superposición exacta (pdfjs-dist + tesseract.js + pdf-lib)

export type TextAlign = 'left' | 'center' | 'right';

export interface EditableField {
  id: string;
  /** Texto originalmente detectado en el PDF */
  originalText: string;
  /** Texto actual (editable por el usuario) */
  text: string;
  /** Coordenadas exactas en puntos del viewport de la página (origen superior-izquierdo) */
  x: number;
  y: number;
  width: number;
  height: number;
  fontSize: number;
  isBold: boolean;
  align: TextAlign;
  /** Ángulo de rotación del texto en grados respecto a la vista (0 = horizontal, -90 = vertical hacia arriba, 90 = vertical hacia abajo) */
  rotation: number;
  /** Color de fondo muestreado del PDF original alrededor del texto (ej. #ffffff o #1e3a8a) */
  bgColor: string;
  /** Color del texto inferido por contraste/muestreo (ej. #0f172a o #ffffff) */
  textColor: string;
  /** Indica si el usuario modificó el texto o estilo respecto al original */
  isModified: boolean;
  /** Indica si el campo fue agregado manualmente por el usuario */
  isNew?: boolean;
  /** Indica si el usuario borró este texto (se tapa con bgColor sin escribir nada) */
  isDeleted?: boolean;
  /** Metadatos opcionales si el campo pertenece a una tabla detectada */
  tableId?: string;
  rowIndex?: number;
  colIndex?: number;
}

export interface DetectedTable {
  id: string;
  rowCount: number;
  colCount: number;
  columnWidths: number[];
  /** Matriz [fila][columna] con TODOS los IDs de EditableField que caen dentro de esa celda (soporta multilínea y vacías) */
  cellFieldIds: string[][][];
  y: number;
  hasHeader: boolean;
  headerBg: string;
  /** Método con el que se detectó la tabla: rejilla de bordes reales o alineación de columnas */
  detectionMethod?: 'grid' | 'alignment';
}

export interface ScannedPage {
  pageNumber: number;
  /** Dimensiones visibles de la página en puntos PDF (ya considerando orientación horizontal/vertical) */
  width: number;
  height: number;
  /** Rotación nativa de la página PDF en grados (0, 90, 180, 270) */
  pageRotation: number;
  /** Matriz viewport de pdfjs-dist [a, b, c, d, e, f] a escala 1 para inversión exacta en pdf-lib */
  viewportTransform: [number, number, number, number, number, number];
  /** Imagen renderizada en alta resolución del PDF original (con logos, líneas, bordes, todo) */
  originalImageUrl: string;
  /** Imagen recortada del área de logo/encabezado si se detectó (para incluir en Word) */
  logoBannerDataUrl?: string;
  logoBannerWidth?: number;
  logoBannerHeight?: number;
  /** Método utilizado para extraer los campos */
  extractionMode: 'text' | 'ocr';
  /** Todos los campos de texto/celdas detectados en sus coordenadas exactas */
  fields: EditableField[];
  /** Tablas detectadas que referencian a los campos anteriores */
  tables: DetectedTable[];
}

export interface ScannedDocument {
  fileName: string;
  /** Bytes originales del archivo PDF para que pdf-lib modifique el original sin perder logos ni diseño */
  originalPdfBytes: Uint8Array;
  pages: ScannedPage[];
}

export interface ScanProgress {
  stage: 'loading' | 'parsing' | 'ocr' | 'done';
  currentPage: number;
  totalPages: number;
  message: string;
}

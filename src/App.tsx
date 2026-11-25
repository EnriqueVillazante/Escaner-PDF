import { useMemo, useRef, useState } from 'react';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  Bold,
  Download,
  Eye,
  EyeOff,
  FileSpreadsheet,
  FileText,
  FileUp,
  Minus,
  MousePointerClick,
  Plus,
  Redo2,
  RotateCcw,
  RotateCw,
  Table2,
  Trash2,
  Undo2,
  X,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';
import { EditablePage } from './components/EditablePage';
import { UploadDropzone } from './components/UploadDropzone';
import type {
  EditableField,
  ScannedDocument,
  ScannedPage,
  ScanProgress,
  TextAlign,
} from './types/document';
import { exportDocumentToPdf, exportDocumentToWord } from './utils/exporters';
import { scanPdfFile } from './utils/pdfParser';

type ViewMode = 'overlay' | 'tables';

export function App() {
  const [doc, setDoc] = useState<ScannedDocument | null>(null);
  const [historyPast, setHistoryPast] = useState<ScannedDocument[]>([]);
  const [historyFuture, setHistoryFuture] = useState<ScannedDocument[]>([]);
  const [viewMode, setViewMode] = useState<ViewMode>('overlay');
  const [zoom, setZoom] = useState<number>(100);
  const [showOutlines, setShowOutlines] = useState<boolean>(true);
  const [activePageNumber, setActivePageNumber] = useState<number>(1);
  const [selectedFieldId, setSelectedFieldId] = useState<string | null>(null);
  const [scanProgress, setScanProgress] = useState<ScanProgress | null>(null);
  const [exportingType, setExportingType] = useState<'pdf' | 'word' | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const pageRefs = useRef<Record<number, HTMLDivElement | null>>({});

  // Campo actualmente seleccionado (y la página a la que pertenece) para mostrarlo en el panel lateral
  const selectedContext = useMemo(() => {
    if (!doc || !selectedFieldId) return null;
    for (const page of doc.pages) {
      const field = page.fields.find((f) => f.id === selectedFieldId);
      if (field) {
        return { page, field };
      }
    }
    return null;
  }, [doc, selectedFieldId]);

  const totalTables = useMemo(() => {
    if (!doc) return 0;
    return doc.pages.reduce((sum, p) => sum + p.tables.length, 0);
  }, [doc]);

  const commitDocumentChange = (nextDoc: ScannedDocument) => {
    if (doc) {
      setHistoryPast((prev) => [...prev.slice(-29), doc]);
      setHistoryFuture([]);
    }
    setDoc(nextDoc);
  };

  const handleUndo = () => {
    if (historyPast.length === 0 || !doc) return;
    const previous = historyPast[historyPast.length - 1];
    setHistoryPast((prev) => prev.slice(0, -1));
    setHistoryFuture((prev) => [doc, ...prev.slice(0, 29)]);
    setDoc(previous);
  };

  const handleRedo = () => {
    if (historyFuture.length === 0 || !doc) return;
    const next = historyFuture[0];
    setHistoryFuture((prev) => prev.slice(1));
    setHistoryPast((prev) => [...prev.slice(-29), doc]);
    setDoc(next);
  };

  const processPdfFile = async (file: File) => {
    setSelectedFieldId(null);
    setScanProgress({
      stage: 'loading',
      currentPage: 0,
      totalPages: 0,
      message: 'Cargando archivo PDF original...',
    });

    try {
      const scanned = await scanPdfFile(file, (progress) => {
        setScanProgress(progress);
      });
      setDoc(scanned);
      setHistoryPast([]);
      setHistoryFuture([]);
      setActivePageNumber(1);
      setViewMode('overlay');
    } catch (error) {
      console.error('Error al escanear PDF:', error);
      alert(
        'No se pudo leer el archivo PDF. Verifica que no esté dañado o protegido con contraseña.'
      );
    } finally {
      setScanProgress(null);
    }
  };

  const handleUpdatePage = (updatedPage: ScannedPage) => {
    if (!doc) return;
    const nextDoc: ScannedDocument = {
      ...doc,
      pages: doc.pages.map((p) =>
        p.pageNumber === updatedPage.pageNumber ? updatedPage : p
      ),
    };
    commitDocumentChange(nextDoc);
  };

  const updateSelectedField = (updater: (f: EditableField) => EditableField) => {
    if (!selectedContext) return;
    const { page, field } = selectedContext;
    handleUpdatePage({
      ...page,
      fields: page.fields.map((f) => (f.id === field.id ? updater(f) : f)),
    });
  };

  const handleDeleteSelectedField = () => {
    if (!selectedContext) return;
    const { page, field } = selectedContext;
    if (field.isNew) {
      handleUpdatePage({
        ...page,
        fields: page.fields.filter((f) => f.id !== field.id),
      });
    } else {
      handleUpdatePage({
        ...page,
        fields: page.fields.map((f) =>
          f.id === field.id ? { ...f, text: '', isDeleted: true, isModified: true } : f
        ),
      });
    }
  };

  const handleTableCellEdit = (pageNumber: number, fieldId: string, newText: string) => {
    if (!doc) return;
    const page = doc.pages.find((p) => p.pageNumber === pageNumber);
    if (!page) return;
    const updatedPage: ScannedPage = {
      ...page,
      fields: page.fields.map((f) =>
        f.id === fieldId
          ? {
              ...f,
              text: newText,
              isDeleted: false,
              isModified: f.isNew ? true : newText !== f.originalText,
            }
          : f
      ),
    };
    handleUpdatePage(updatedPage);
  };

  const handleDownloadPdf = async () => {
    if (!doc) return;
    setExportingType('pdf');
    try {
      await exportDocumentToPdf(doc);
    } catch (err) {
      console.error('Error al generar PDF con pdf-lib:', err);
      alert('No se pudo generar el PDF modificado.');
    } finally {
      setExportingType(null);
    }
  };

  const handleDownloadWord = async () => {
    if (!doc) return;
    setExportingType('word');
    try {
      await exportDocumentToWord(doc);
    } catch (err) {
      console.error('Error al exportar a Word:', err);
      alert('No se pudo generar el documento Word (.docx).');
    } finally {
      setExportingType(null);
    }
  };

  const scrollToPage = (pageNumber: number) => {
    setActivePageNumber(pageNumber);
    const el = pageRefs.current[pageNumber];
    if (el) {
      el.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  };

  const getOrientationLabel = (rotation: number): string => {
    if (rotation === 0) return 'Horizontal (0°)';
    if (rotation === -90) return 'Vertical hacia arriba (-90°)';
    if (rotation === 90) return 'Vertical hacia abajo (90°)';
    if (rotation === 180) return 'Invertido (180°)';
    return `Rotado (${rotation}°)`;
  };

  return (
    <div className="app-shell">
      <input
        ref={fileInputRef}
        type="file"
        accept="application/pdf,.pdf"
        style={{ display: 'none' }}
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) void processPdfFile(f);
          e.target.value = '';
        }}
      />

      {/* Barra superior */}
      <header className="topbar">
        <div className="topbar-left">
          <div className="brand">
            <div className="brand-icon">
              <FileSpreadsheet size={18} />
            </div>
            <div>
              <div className="brand-title">Escáner PDF Editable</div>
              <div className="brand-file">{doc ? doc.fileName : 'Sube tu PDF'}</div>
            </div>
          </div>

          {doc && (
            <button
              type="button"
              className="btn"
              onClick={() => fileInputRef.current?.click()}
            >
              <FileUp size={15} />
              Otro PDF
            </button>
          )}
        </div>

        {doc && (
          <>
            <div className="topbar-center">
              <div className="segmented" role="tablist" aria-label="Modo de edición">
                <button
                  type="button"
                  className={`segmented-btn ${viewMode === 'overlay' ? 'active' : ''}`}
                  onClick={() => setViewMode('overlay')}
                >
                  <FileText size={14} />
                  PDF Exacto
                </button>
                <button
                  type="button"
                  className={`segmented-btn ${viewMode === 'tables' ? 'active' : ''}`}
                  onClick={() => setViewMode('tables')}
                >
                  <Table2 size={14} />
                  Tablas ({totalTables})
                </button>
              </div>

              {viewMode === 'overlay' && (
                <>
                  <button
                    type="button"
                    className={`btn ${showOutlines ? 'btn-active-soft' : ''}`}
                    title="Mostrar u ocultar recuadros de campos detectados"
                    onClick={() => setShowOutlines((v) => !v)}
                  >
                    {showOutlines ? <Eye size={14} /> : <EyeOff size={14} />}
                    Campos
                  </button>

                  <div className="zoom-group">
                    <button
                      type="button"
                      title="Alejar"
                      onClick={() => setZoom((z) => Math.max(55, z - 10))}
                    >
                      <ZoomOut size={14} />
                    </button>
                    <span className="zoom-label">{zoom}%</span>
                    <button
                      type="button"
                      title="Acercar"
                      onClick={() => setZoom((z) => Math.min(160, z + 10))}
                    >
                      <ZoomIn size={14} />
                    </button>
                  </div>
                </>
              )}

              <button
                type="button"
                className="btn btn-icon"
                title="Deshacer"
                disabled={historyPast.length === 0}
                onClick={handleUndo}
              >
                <Undo2 size={15} />
              </button>
              <button
                type="button"
                className="btn btn-icon"
                title="Rehacer"
                disabled={historyFuture.length === 0}
                onClick={handleRedo}
              >
                <Redo2 size={15} />
              </button>
            </div>

            <div className="topbar-right">
              <button
                type="button"
                className="btn btn-primary"
                disabled={exportingType !== null}
                onClick={() => void handleDownloadPdf()}
              >
                <Download size={15} />
                {exportingType === 'pdf' ? 'Guardando PDF...' : 'Descargar PDF'}
              </button>
              <button
                type="button"
                className="btn btn-word"
                disabled={exportingType !== null}
                onClick={() => void handleDownloadWord()}
              >
                <Download size={15} />
                {exportingType === 'word' ? 'Generando Word...' : 'Descargar Word'}
              </button>
            </div>
          </>
        )}
      </header>

      {/* Contenido principal */}
      {!doc ? (
        <UploadDropzone onSelectFile={(f) => void processPdfFile(f)} />
      ) : (
        <div className="workspace-body">
          {/* Panel lateral izquierdo: solo páginas si hay varias */}
          {doc.pages.length > 1 && (
            <aside className="sidebar sidebar-compact">
              <h2 className="sidebar-section-title">Páginas</h2>
              <div className="pages-list">
                {doc.pages.map((page) => (
                  <button
                    key={page.pageNumber}
                    type="button"
                    className={`page-thumb-card ${
                      activePageNumber === page.pageNumber ? 'active' : ''
                    }`}
                    onClick={() => scrollToPage(page.pageNumber)}
                  >
                    <img
                      src={page.originalImageUrl}
                      alt={`Página ${page.pageNumber}`}
                      className="page-thumb-img"
                    />
                    <div className="page-thumb-meta">
                      <span className="page-thumb-title">Página {page.pageNumber}</span>
                    </div>
                  </button>
                ))}
              </div>
            </aside>
          )}

          {/* Lienzo principal */}
          <main className="canvas-scroll" onClick={() => setSelectedFieldId(null)}>
            {viewMode === 'overlay' ? (
              doc.pages.map((page) => (
                <section
                  key={page.pageNumber}
                  ref={(el) => {
                    pageRefs.current[page.pageNumber] = el as HTMLDivElement | null;
                  }}
                  className="page-pair-wrapper"
                  onMouseEnter={() => setActivePageNumber(page.pageNumber)}
                >
                  <EditablePage
                    page={page}
                    zoom={zoom}
                    showOutlines={showOutlines}
                    addTextMode={false}
                    selectedFieldId={selectedFieldId}
                    onSelectField={setSelectedFieldId}
                    onExitAddTextMode={() => {}}
                    onUpdatePage={handleUpdatePage}
                  />
                </section>
              ))
            ) : (
              <div className="tables-view-container">
                {doc.pages.map((page) => {
                  const fieldMap = new Map<string, EditableField>();
                  page.fields.forEach((f) => fieldMap.set(f.id, f));

                  return (
                    <div key={page.pageNumber} className="tables-page-card">
                      <h3 className="tables-page-title">
                        Página {page.pageNumber} — {page.tables.length}{' '}
                        {page.tables.length === 1 ? 'tabla detectada' : 'tablas detectadas'}
                      </h3>

                      {page.tables.length === 0 ? (
                        <div className="empty-page-hint">
                          No se detectaron tablas estructuradas en esta página. Edita cada
                          campo directamente en la vista <strong>PDF Exacto</strong>.
                        </div>
                      ) : (
                        page.tables.map((tbl, tIdx) => (
                          <div key={tbl.id} className="sync-table-wrap">
                            <div className="sync-table-label">
                              Tabla {tIdx + 1} ({tbl.rowCount} filas × {tbl.colCount} columnas)
                            </div>
                            <table className="doc-table">
                              <tbody>
                                {tbl.cellFieldIds.map((rowIds, rIdx) => {
                                  const isHeader = tbl.hasHeader && rIdx === 0;
                                  return (
                                    <tr
                                      key={rIdx}
                                      className={isHeader ? 'doc-table-header-row' : undefined}
                                      style={
                                        isHeader
                                          ? { backgroundColor: tbl.headerBg || '#1e3a8a' }
                                          : undefined
                                      }
                                    >
                                      {rowIds.map((fid, cIdx) => {
                                        const fld = fid ? fieldMap.get(fid) : undefined;
                                        return (
                                          <td key={cIdx}>
                                            {fld ? (
                                              <input
                                                type="text"
                                                value={fld.isDeleted ? '' : fld.text}
                                                className="cell-input"
                                                style={{
                                                  textAlign: fld.align,
                                                  fontWeight: isHeader || fld.isBold ? 700 : 400,
                                                }}
                                                onFocus={() => setSelectedFieldId(fld.id)}
                                                onChange={(e) =>
                                                  handleTableCellEdit(
                                                    page.pageNumber,
                                                    fld.id,
                                                    e.target.value
                                                  )
                                                }
                                              />
                                            ) : (
                                              <span className="empty-cell-dash">—</span>
                                            )}
                                          </td>
                                        );
                                      })}
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        ))
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </main>

          {/* Panel lateral derecho: Inspector y editor cómodo del campo seleccionado (siempre en horizontal) */}
          <aside className="inspector-panel">
            {selectedContext ? (
              <div className="inspector-content">
                <div className="inspector-header">
                  <div>
                    <h2 className="inspector-title">Campo seleccionado</h2>
                    <span className="inspector-subtitle">
                      Página {selectedContext.page.pageNumber} ·{' '}
                      {getOrientationLabel(selectedContext.field.rotation)}
                    </span>
                  </div>
                  <button
                    type="button"
                    className="btn btn-icon"
                    title="Cerrar selección"
                    onClick={() => setSelectedFieldId(null)}
                  >
                    <X size={14} />
                  </button>
                </div>

                {/* Caja de edición siempre horizontal y legible */}
                <div className="inspector-group">
                  <label className="inspector-label" htmlFor="inspector-text-input">
                    Contenido del campo (edita aquí cómodamente)
                  </label>
                  <textarea
                    id="inspector-text-input"
                    className="inspector-textarea"
                    rows={4}
                    value={selectedContext.field.isDeleted ? '' : selectedContext.field.text}
                    placeholder="Campo vacío (texto borrado)"
                    onChange={(e) => {
                      const nextVal = e.target.value;
                      updateSelectedField((f) => ({
                        ...f,
                        text: nextVal,
                        isDeleted: false,
                        isModified: f.isNew ? true : nextVal !== f.originalText,
                      }));
                    }}
                  />
                </div>

                {!selectedContext.field.isNew && (
                  <div className="inspector-original-box">
                    <span className="inspector-original-label">Texto original:</span>
                    <div className="inspector-original-value">
                      {selectedContext.field.originalText}
                    </div>
                  </div>
                )}

                {/* Controles de formato */}
                <div className="inspector-group">
                  <span className="inspector-label">Formato y estilo</span>
                  <div className="inspector-controls-row">
                    <button
                      type="button"
                      className={`btn ${selectedContext.field.isBold ? 'btn-active-soft' : ''}`}
                      onClick={() =>
                        updateSelectedField((f) => ({
                          ...f,
                          isBold: !f.isBold,
                          isModified: true,
                        }))
                      }
                    >
                      <Bold size={14} />
                      Negrita
                    </button>

                    <div className="zoom-group">
                      <button
                        type="button"
                        title="Reducir tamaño"
                        onClick={() =>
                          updateSelectedField((f) => ({
                            ...f,
                            fontSize: Math.max(6, Math.round((f.fontSize - 0.5) * 10) / 10),
                            isModified: true,
                          }))
                        }
                      >
                        <Minus size={13} />
                      </button>
                      <span className="zoom-label">
                        {Math.round(selectedContext.field.fontSize)}pt
                      </span>
                      <button
                        type="button"
                        title="Aumentar tamaño"
                        onClick={() =>
                          updateSelectedField((f) => ({
                            ...f,
                            fontSize: Math.min(36, Math.round((f.fontSize + 0.5) * 10) / 10),
                            isModified: true,
                          }))
                        }
                      >
                        <Plus size={13} />
                      </button>
                    </div>
                  </div>
                </div>

                {/* Alineación y orientación */}
                <div className="inspector-group">
                  <span className="inspector-label">Alineación y orientación</span>
                  <div className="inspector-controls-row">
                    {(['left', 'center', 'right'] as TextAlign[]).map((align) => (
                      <button
                        key={align}
                        type="button"
                        className={`btn btn-icon ${
                          selectedContext.field.align === align ? 'btn-active-soft' : ''
                        }`}
                        title={
                          align === 'left'
                            ? 'Izquierda'
                            : align === 'center'
                              ? 'Centro'
                              : 'Derecha'
                        }
                        onClick={() =>
                          updateSelectedField((f) => ({
                            ...f,
                            align,
                            isModified: true,
                          }))
                        }
                      >
                        {align === 'left' ? (
                          <AlignLeft size={14} />
                        ) : align === 'center' ? (
                          <AlignCenter size={14} />
                        ) : (
                          <AlignRight size={14} />
                        )}
                      </button>
                    ))}

                    <button
                      type="button"
                      className="btn"
                      title="Cambiar orientación del texto en el PDF"
                      onClick={() => {
                        const curr = selectedContext.field.rotation;
                        const nextRot = curr === 0 ? -90 : curr === -90 ? 90 : 0;
                        const swap = Math.abs(curr) === 90 !== (Math.abs(nextRot) === 90);
                        updateSelectedField((f) => ({
                          ...f,
                          rotation: nextRot,
                          width: swap ? f.height : f.width,
                          height: swap ? f.width : f.height,
                          isModified: true,
                        }));
                      }}
                    >
                      <RotateCw size={14} />
                      Rotar ({selectedContext.field.rotation}°)
                    </button>
                  </div>
                </div>

                {/* Colores */}
                <div className="inspector-group">
                  <span className="inspector-label">Colores del campo</span>
                  <div className="inspector-colors-row">
                    <label className="color-picker-label">
                      <span>Texto</span>
                      <input
                        type="color"
                        value={selectedContext.field.textColor}
                        onChange={(e) =>
                          updateSelectedField((f) => ({
                            ...f,
                            textColor: e.target.value,
                            isModified: true,
                          }))
                        }
                      />
                    </label>
                    <label className="color-picker-label">
                      <span>Fondo</span>
                      <input
                        type="color"
                        value={selectedContext.field.bgColor}
                        onChange={(e) =>
                          updateSelectedField((f) => ({
                            ...f,
                            bgColor: e.target.value,
                            isModified: true,
                          }))
                        }
                      />
                    </label>
                  </div>
                </div>

                {/* Acciones finales */}
                <div className="inspector-footer-actions">
                  {!selectedContext.field.isNew &&
                    (selectedContext.field.isModified || selectedContext.field.isDeleted) && (
                      <button
                        type="button"
                        className="btn"
                        onClick={() =>
                          updateSelectedField((f) => ({
                            ...f,
                            text: f.originalText,
                            isModified: false,
                            isDeleted: false,
                          }))
                        }
                      >
                        <RotateCcw size={14} />
                        Restaurar original
                      </button>
                    )}

                  <button
                    type="button"
                    className="btn btn-danger-ghost"
                    onClick={handleDeleteSelectedField}
                  >
                    <Trash2 size={14} />
                    Borrar texto
                  </button>
                </div>
              </div>
            ) : (
              <div className="inspector-empty">
                <div className="inspector-empty-icon">
                  <MousePointerClick size={24} />
                </div>
                <p className="inspector-empty-title">Selecciona un campo</p>
                <p className="inspector-empty-desc">
                  Haz clic en cualquier texto o celda del PDF para verlo y editarlo aquí
                  cómodamente, sin importar si está en vertical o en otro sentido.
                </p>
              </div>
            )}
          </aside>
        </div>
      )}

      {/* Indicador de progreso durante escaneo / OCR */}
      {scanProgress && (
        <div className="loading-overlay">
          <div className="loading-box">
            <div className="spinner" />
            <h3 style={{ margin: '0 0 6px 0', fontSize: 16 }}>Escaneando documento</h3>
            <p style={{ margin: 0, fontSize: 13.5, color: '#475569' }}>
              {scanProgress.message}
            </p>
            {scanProgress.totalPages > 0 && (
              <p style={{ margin: '10px 0 0 0', fontSize: 12, color: '#64748b' }}>
                Página {scanProgress.currentPage} de {scanProgress.totalPages}
              </p>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

export default App;

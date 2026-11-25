import React, { useRef } from 'react';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  Bold,
  Minus,
  Plus,
  RotateCcw,
  RotateCw,
  Trash2,
} from 'lucide-react';
import type { EditableField, ScannedPage, TextAlign } from '../types/document';
import { createId } from '../utils/pdfParser';

interface EditablePageProps {
  page: ScannedPage;
  zoom: number;
  showOutlines: boolean;
  addTextMode: boolean;
  selectedFieldId: string | null;
  onSelectField: (fieldId: string | null) => void;
  onExitAddTextMode: () => void;
  onUpdatePage: (updatedPage: ScannedPage) => void;
}

export const EditablePage: React.FC<EditablePageProps> = ({
  page,
  zoom,
  showOutlines,
  addTextMode,
  selectedFieldId,
  onSelectField,
  onExitAddTextMode,
  onUpdatePage,
}) => {
  const sheetRef = useRef<HTMLDivElement>(null);
  // Escala visual: 1 pt PDF = 1.25 px al 100% de zoom
  const displayScale = 1.25 * (zoom / 100);
  const sheetWidthPx = Math.round(page.width * displayScale);
  const sheetHeightPx = Math.round(page.height * displayScale);

  const updateField = (fieldId: string, updater: (f: EditableField) => EditableField) => {
    onUpdatePage({
      ...page,
      fields: page.fields.map((f) => (f.id === fieldId ? updater(f) : f)),
    });
  };

  const removeNewOrDeleteField = (field: EditableField) => {
    if (field.isNew) {
      onUpdatePage({
        ...page,
        fields: page.fields.filter((f) => f.id !== field.id),
      });
    } else {
      updateField(field.id, (f) => ({
        ...f,
        text: '',
        isDeleted: true,
        isModified: true,
      }));
    }
    onSelectField(null);
  };

  const restoreField = (fieldId: string) => {
    updateField(fieldId, (f) => ({
      ...f,
      text: f.originalText,
      isDeleted: false,
      isModified: false,
    }));
  };

  const cycleRotation = (field: EditableField) => {
    const nextRot = field.rotation === 0 ? -90 : field.rotation === -90 ? 90 : 0;
    // Al cambiar entre horizontal y vertical, intercambiar ancho y alto para mantener proporción
    const wasVertical = Math.abs(field.rotation) === 90;
    const willBeVertical = Math.abs(nextRot) === 90;
    const swapDims = wasVertical !== willBeVertical;

    updateField(field.id, (f) => ({
      ...f,
      rotation: nextRot,
      width: swapDims ? f.height : f.width,
      height: swapDims ? f.width : f.height,
      isModified: true,
    }));
  };

  const handleSheetClick = (e: React.MouseEvent<HTMLDivElement>) => {
    if (!addTextMode) {
      onSelectField(null);
      return;
    }

    const rect = sheetRef.current?.getBoundingClientRect();
    if (!rect) return;

    const clickXPt = (e.clientX - rect.left) / displayScale;
    const clickYPt = (e.clientY - rect.top) / displayScale;

    const newField: EditableField = {
      id: createId('new'),
      originalText: '',
      text: 'Nuevo texto',
      x: Math.max(8, clickXPt - 4),
      y: Math.max(8, clickYPt - 7),
      width: 110,
      height: 15,
      fontSize: 10,
      isBold: false,
      align: 'left',
      rotation: 0,
      bgColor: '#ffffff',
      textColor: '#0f172a',
      isModified: true,
      isNew: true,
    };

    onUpdatePage({
      ...page,
      fields: [...page.fields, newField],
    });
    onSelectField(newField.id);
    onExitAddTextMode();
  };

  return (
    <div
      ref={sheetRef}
      className={`document-sheet overlay-sheet ${addTextMode ? 'add-text-cursor' : ''}`}
      style={{
        width: `${sheetWidthPx}px`,
        height: `${sheetHeightPx}px`,
      }}
      onClick={handleSheetClick}
    >
      {/* Fondo 100% exacto del PDF original: logos, tablas, líneas, colores y diseño intactos */}
      <img
        src={page.originalImageUrl}
        alt={`Página ${page.pageNumber} del PDF original`}
        className="original-sheet-img"
        draggable={false}
      />

      {/* Capa de campos editables posicionados en las coordenadas exactas */}
      <div className="overlay-fields-layer">
        {page.fields.map((field) => {
          const isSelected = selectedFieldId === field.id;
          const isCovered = isSelected || field.isModified || field.isDeleted || field.isNew;
          const isVertical = Math.abs(field.rotation) === 90;

          const leftPx = field.x * displayScale;
          const topPx = field.y * displayScale;
          const widthPx = Math.max(14, field.width * displayScale);
          const heightPx = Math.max(14, field.height * displayScale);
          const fontSizePx = Math.max(7.5, field.fontSize * displayScale);

          return (
            <div
              key={field.id}
              className={[
                'overlay-field',
                showOutlines ? 'show-outline' : '',
                isSelected ? 'is-selected' : '',
                field.isModified && !field.isDeleted ? 'is-modified' : '',
                field.isDeleted ? 'is-deleted' : '',
                isVertical ? 'is-vertical-field' : '',
              ]
                .filter(Boolean)
                .join(' ')}
              style={{
                left: `${leftPx - (isCovered ? 1.5 : 0)}px`,
                top: `${topPx - (isCovered ? 1.5 : 0)}px`,
                width: `${widthPx + (isCovered ? 3 : 0)}px`,
                height: `${heightPx + (isCovered ? 3 : 0)}px`,
                backgroundColor: isCovered ? field.bgColor : 'transparent',
              }}
              onClick={(e) => {
                if (addTextMode) return;
                e.stopPropagation();
                onSelectField(field.id);
              }}
            >
              {/* Barra flotante de formato cuando el campo está seleccionado */}
              {isSelected && (
                <div
                  className="field-floating-bar"
                  onClick={(e) => e.stopPropagation()}
                  onMouseDown={(e) => e.preventDefault()}
                >
                  <button
                    type="button"
                    title="Negrita"
                    className={field.isBold ? 'active' : ''}
                    onClick={() =>
                      updateField(field.id, (f) => ({
                        ...f,
                        isBold: !f.isBold,
                        isModified: true,
                      }))
                    }
                  >
                    <Bold size={12} />
                  </button>
                  <button
                    type="button"
                    title="Reducir fuente"
                    onClick={() =>
                      updateField(field.id, (f) => ({
                        ...f,
                        fontSize: Math.max(6, Math.round((f.fontSize - 0.5) * 10) / 10),
                        isModified: true,
                      }))
                    }
                  >
                    <Minus size={11} />
                  </button>
                  <span className="bar-size-label">{Math.round(field.fontSize)}pt</span>
                  <button
                    type="button"
                    title="Aumentar fuente"
                    onClick={() =>
                      updateField(field.id, (f) => ({
                        ...f,
                        fontSize: Math.min(36, Math.round((f.fontSize + 0.5) * 10) / 10),
                        isModified: true,
                      }))
                    }
                  >
                    <Plus size={11} />
                  </button>
                  <span className="bar-divider" />
                  <button
                    type="button"
                    title={`Rotar texto (actual: ${field.rotation}°)`}
                    className={field.rotation !== 0 ? 'active' : ''}
                    onClick={() => cycleRotation(field)}
                  >
                    <RotateCw size={12} />
                  </button>
                  <span className="bar-divider" />
                  {(['left', 'center', 'right'] as TextAlign[]).map((align) => (
                    <button
                      key={align}
                      type="button"
                      title={
                        align === 'left'
                          ? 'Izquierda'
                          : align === 'center'
                            ? 'Centro'
                            : 'Derecha'
                      }
                      className={field.align === align ? 'active' : ''}
                      onClick={() =>
                        updateField(field.id, (f) => ({
                          ...f,
                          align,
                          isModified: true,
                        }))
                      }
                    >
                      {align === 'left' ? (
                        <AlignLeft size={12} />
                      ) : align === 'center' ? (
                        <AlignCenter size={12} />
                      ) : (
                        <AlignRight size={12} />
                      )}
                    </button>
                  ))}
                  {!field.isNew && (field.isModified || field.isDeleted) && (
                    <>
                      <span className="bar-divider" />
                      <button
                        type="button"
                        title="Restaurar texto original"
                        onClick={() => restoreField(field.id)}
                      >
                        <RotateCcw size={12} />
                      </button>
                    </>
                  )}
                  <span className="bar-divider" />
                  <button
                    type="button"
                    title={field.isNew ? 'Eliminar campo' : 'Borrar texto del PDF'}
                    className="danger"
                    onClick={() => removeNewOrDeleteField(field)}
                  >
                    <Trash2 size={12} />
                  </button>
                </div>
              )}

              {/* Input editable: soporta orientación horizontal, vertical (-90° / 90°) y rotada */}
              {isVertical ? (
                <input
                  type="text"
                  value={field.isDeleted ? '' : field.text}
                  aria-label={`Editar campo vertical ${field.originalText || 'nuevo'}`}
                  className="overlay-input overlay-input-rotated"
                  style={{
                    width: `${heightPx}px`,
                    height: `${widthPx}px`,
                    fontSize: `${fontSizePx}px`,
                    fontWeight: field.isBold ? 700 : 400,
                    textAlign: field.align,
                    color: isCovered ? field.textColor : 'transparent',
                    caretColor: field.textColor,
                    transform: `translate(-50%, -50%) rotate(${field.rotation}deg)`,
                  }}
                  onFocus={() => onSelectField(field.id)}
                  onChange={(e) => {
                    const nextVal = e.target.value;
                    updateField(field.id, (f) => ({
                      ...f,
                      text: nextVal,
                      isDeleted: false,
                      isModified: f.isNew ? true : nextVal !== f.originalText,
                    }));
                  }}
                />
              ) : (
                <input
                  type="text"
                  value={field.isDeleted ? '' : field.text}
                  aria-label={`Editar campo ${field.originalText || 'nuevo'}`}
                  className="overlay-input"
                  style={{
                    fontSize: `${fontSizePx}px`,
                    fontWeight: field.isBold ? 700 : 400,
                    textAlign: field.align,
                    color: isCovered ? field.textColor : 'transparent',
                    caretColor: field.textColor,
                    lineHeight: `${heightPx}px`,
                    transform: field.rotation !== 0 ? `rotate(${field.rotation}deg)` : undefined,
                  }}
                  onFocus={() => onSelectField(field.id)}
                  onChange={(e) => {
                    const nextVal = e.target.value;
                    updateField(field.id, (f) => ({
                      ...f,
                      text: nextVal,
                      isDeleted: false,
                      isModified: f.isNew ? true : nextVal !== f.originalText,
                    }));
                  }}
                />
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
};

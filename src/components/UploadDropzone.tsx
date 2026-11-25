import React, { useRef, useState } from 'react';
import { FileUp } from 'lucide-react';

interface UploadDropzoneProps {
  onSelectFile: (file: File) => void;
}

export const UploadDropzone: React.FC<UploadDropzoneProps> = ({ onSelectFile }) => {
  const inputRef = useRef<HTMLInputElement>(null);
  const [isDragging, setIsDragging] = useState(false);

  const handleFiles = (files: FileList | null) => {
    if (!files || files.length === 0) return;
    const file = files[0];
    if (file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')) {
      onSelectFile(file);
    } else {
      alert('Por favor selecciona un archivo en formato PDF (.pdf).');
    }
  };

  return (
    <div className="landing-wrap">
      <div className="landing-card">
        <h1 className="landing-title">Escáner PDF con Previsualización Editable</h1>
        <p className="landing-subtitle">
          Sube un archivo PDF con tablas y registros. Conservamos el diseño exacto con sus
          logos y bordes para que edites cada campo y descargues en <strong>PDF</strong> o{' '}
          <strong>Word (.docx)</strong>.
        </p>

        <input
          ref={inputRef}
          type="file"
          accept="application/pdf,.pdf"
          style={{ display: 'none' }}
          onChange={(e) => {
            handleFiles(e.target.files);
            e.target.value = '';
          }}
        />

        <div
          className={`dropzone ${isDragging ? 'dragging' : ''}`}
          onClick={() => inputRef.current?.click()}
          onDragOver={(e) => {
            e.preventDefault();
            setIsDragging(true);
          }}
          onDragLeave={() => setIsDragging(false)}
          onDrop={(e) => {
            e.preventDefault();
            setIsDragging(false);
            handleFiles(e.dataTransfer.files);
          }}
        >
          <div className="dropzone-icon">
            <FileUp size={28} />
          </div>
          <p className="dropzone-title">Arrastra tu archivo PDF aquí</p>
          <p className="dropzone-hint">
            Compatible con PDFs digitales y escaneados (verticales, horizontales y texto
            rotado)
          </p>

          <div className="landing-actions" onClick={(e) => e.stopPropagation()}>
            <button
              type="button"
              className="btn btn-primary"
              onClick={() => inputRef.current?.click()}
            >
              <FileUp size={16} />
              Seleccionar archivo PDF
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

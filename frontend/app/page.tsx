'use client';

import React, { useState, useEffect } from 'react';
import { 
  Upload, 
  CheckCircle, 
  AlertCircle, 
  Loader2, 
  Download, 
  FileSpreadsheet, 
  FileText, 
  Sparkles, 
  Search, 
  RefreshCw 
} from 'lucide-react';

export default function Home() {
  const [paso, setPaso] = useState<number>(1);

  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [wordFile, setWordFile] = useState<File | null>(null);

  const [codigo, setCodigo] = useState('');
  const [ordenTrabajo, setOrdenTrabajo] = useState('');
  const [tecnico1, setTecnico1] = useState('');
  const [tecnico2, setTecnico2] = useState('');
  const [tecnico3, setTecnico3] = useState('');
  const [tecnico4, setTecnico4] = useState('');
  const [fecha, setFecha] = useState('');
  const [horas, setHoras] = useState('');

  const [uploadingExcel, setUploadingExcel] = useState(false);
  const [processingWord, setProcessingWord] = useState(false);
  const [searching, setSearching] = useState(false);
  const [loadingDoc, setLoadingDoc] = useState(false);
  const [statusMessage, setStatusMessage] = useState<{ type: 'success' | 'error' | 'info' | ''; text: string }>({ type: '', text: '' });
  const [jsReady, setJsReady] = useState(false);

  useEffect(() => {
    setJsReady(true);
  }, []);

  const mapearDatosFormulario = (data: any) => {
    setOrdenTrabajo(data.orden_trabajo || '');
    setTecnico1(data.tecnico1 || '');
    setTecnico2(data.tecnico2 || '');
    setTecnico3(data.tecnico3 || '');
    setTecnico4(data.tecnico4 || '');
    setFecha(data.fecha || '');
    setHoras(data.horas || '');
  };

  const handleExcelUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      setExcelFile(file);
      setUploadingExcel(true);
      setStatusMessage({ type: '', text: '' });

      const formData = new FormData();
      formData.append('excel', file);

      try {
        const response = await fetch('http://localhost:8000/api/cargar-excel', {
          method: 'POST',
          body: formData,
        });

        const data = await response.json();
        if (!response.ok) throw new Error(data.detail || 'Error al procesar el Excel.');

        setStatusMessage({ type: 'success', text: '¡Matriz Excel procesada con éxito! Procede con el Paso 2.' });
        setPaso(2);
      } catch (error: any) {
        setStatusMessage({ type: 'error', text: error.message || 'Error de conexión con el backend.' });
      } finally {
        setUploadingExcel(false);
      }
    }
  };

  const handleWordUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      if (!file.name.endsWith('.docx')) {
        setStatusMessage({ type: 'error', text: 'Selecciona un archivo Word válido (.docx).' });
        return;
      }

      setWordFile(file);
      setProcessingWord(true);
      setStatusMessage({ type: '', text: '' });

      const formData = new FormData();
      formData.append('word', file);

      try {
        const response = await fetch('http://localhost:8000/api/procesar-word-auto', {
          method: 'POST',
          body: formData,
        });

        const data = await response.json();
        if (!response.ok) throw new Error(data.detail || 'Error al leer la plantilla Word.');

        if (data.autodetectado && data.data) {
          setCodigo(data.codigoDetectado || '');
          mapearDatosFormulario(data.data);
          setStatusMessage({ 
            type: 'success', 
            text: `✨ ¡Máquina detectada! Orden de Trabajo (${data.data.orden_trabajo}) y Técnicos autocompletados.` 
          });
        } else {
          setCodigo(data.codigoDetectado || '');
          setStatusMessage({ 
            type: 'info', 
            text: 'Confirma el nombre de la máquina en el Paso 3 y presiona ENTER.' 
          });
        }

        setPaso(3);
      } catch (error: any) {
        setStatusMessage({ type: 'error', text: error.message });
      } finally {
        setProcessingWord(false);
      }
    }
  };

  const handleSearchManual = async (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      if (!codigo.trim()) return;

      setSearching(true);
      setStatusMessage({ type: '', text: '' });

      try {
        const response = await fetch('http://localhost:8000/api/consultar-maquina', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ codigo: codigo.trim() }),
        });

        const data = await response.json();
        if (!response.ok) throw new Error(data.detail || 'No se encontró la máquina.');

        mapearDatosFormulario(data.data);
        setStatusMessage({ type: 'success', text: '¡Datos de la máquina encontrados y autocompletados!' });
      } catch (error: any) {
        setStatusMessage({ type: 'error', text: error.message || 'Máquina no encontrada.' });
      } finally {
        setSearching(false);
      }
    }
  };

  const handleSubmitFinal = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!wordFile) return;

    setLoadingDoc(true);
    setStatusMessage({ type: '', text: '' });

    const formData = new FormData();
    formData.append('file', wordFile);
    formData.append('codigo', codigo);
    formData.append('orden_trabajo', ordenTrabajo);
    formData.append('tecnico1', tecnico1);
    formData.append('tecnico2', tecnico2);
    formData.append('tecnico3', tecnico3);
    formData.append('tecnico4', tecnico4);
    formData.append('fecha', fecha);
    formData.append('horas', horas);

    try {
      const response = await fetch('http://localhost:8000/api/generar-ficha', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({ detail: 'Error al generar documento.' }));
        throw new Error(errorData.detail);
      }

      const blob = await response.blob();
      const downloadUrl = window.URL.createObjectURL(blob);
      
      const a = document.createElement('a');
      a.href = downloadUrl;
      a.download = `${ordenTrabajo || 'Ficha'}_ACTUALIZADO.docx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(downloadUrl);

      setStatusMessage({ type: 'success', text: '¡Ficha en Word generada y descargada con éxito!' });
    } catch (error: any) {
      setStatusMessage({ type: 'error', text: error.message });
    } finally {
      setLoadingDoc(false);
    }
  };

  return (
    <main className="min-h-screen bg-[#D5EAD8] bg-linear-to-tr from-[#C1E1A7]/40 via-[#D5EAD8] to-[#E6F0E3] flex flex-col items-center justify-center p-4 md:p-6 font-sans">
      
      {/* Tarjeta Central Ampliada (max-w-3xl) */}
      <div className="w-full max-w-3xl bg-white rounded-2xl shadow-xl shadow-slate-900/5 border border-[#148D8D]/20 px-6 py-6 md:px-10 md:py-8 relative transition-all duration-300">
        
        {/* Encabezado */}
        <div className="text-center mb-6 border-b border-[#148D8D]/20 pb-6">
          <div className="inline-flex items-center justify-center w-16 h-16 rounded-full bg-linear-to-br from-[#B78A62] via-[#E5C29F] to-[#8C6239] mb-3 shadow-md border border-[#E5C29F]/40 shadow-slate-900/10">
            <span className="text-[#3A2512] font-serif font-extrabold text-base tracking-wider drop-shadow-sm">
              BCRP
            </span>
          </div>

          <h1 className="text-2xl font-extrabold text-[#0E2C40] tracking-tight md:text-3xl">
            Fichas de Mantenimiento
          </h1>

          <div className="flex flex-col items-center mt-2 text-xs font-bold tracking-widest text-[#148D8D] uppercase">
            <span>Gerencia de Gestión del Circulante</span>
            <div className="w-72 border-t border-[#148D8D]/30 my-1" />
            <span>Casa Nacional de Moneda</span>
          </div>

          <div className="mt-3">
            <span className={`inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-[10px] font-bold uppercase tracking-wider border transition-all duration-300 ${
              jsReady
                ? 'bg-emerald-600/10 text-emerald-700 border-emerald-600/20'
                : 'bg-amber-600/10 text-amber-700 border-amber-600/20'
            }`}>
              <span className={`w-1.5 h-1.5 rounded-full ${jsReady ? 'bg-emerald-500 animate-pulse' : 'bg-amber-500'}`} />
              {jsReady ? 'Sistema Operativo' : 'Sincronizando'}
            </span>
          </div>
        </div>

        <div className="space-y-6">
          
          {/* PASO 1: EXCEL */}
          <div>
            <label className="block text-[11px] font-bold uppercase tracking-widest text-[#1A4A5A] mb-2 flex items-center justify-between">
              <span>PASO 1: Subir Matriz del Plan Anual (Excel .xlsx)</span>
              {uploadingExcel && <Loader2 className="w-3.5 h-3.5 animate-spin text-[#148D8D]" />}
            </label>
            <input
              type="file"
              accept=".xlsx"
              onChange={handleExcelUpload}
              className="w-full text-sm text-slate-600 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-xs file:font-bold file:bg-[#1A4A5A] file:text-white hover:file:bg-[#0E2C40] file:cursor-pointer bg-slate-50 border border-slate-200 rounded-xl p-1.5 transition-all hover:border-slate-300 shadow-sm"
            />
            {excelFile && (
              <p className="text-xs text-[#148D8D] font-semibold mt-2 flex items-center gap-1">
                <FileSpreadsheet className="w-4 h-4" /> Matriz activa: {excelFile.name}
              </p>
            )}
          </div>

          {/* PASO 2: WORD */}
          {paso >= 2 && (
            <div className="animate-fade-in border-t border-[#148D8D]/15 pt-4">
              <label className="block text-[11px] font-bold uppercase tracking-widest text-[#1A4A5A] mb-2 flex items-center justify-between">
                <span>PASO 2: Subir Plantilla Word (.docx) del Año Anterior</span>
                {processingWord && <Loader2 className="w-3.5 h-3.5 animate-spin text-[#148D8D]" />}
              </label>
              <input
                type="file"
                accept=".docx"
                onChange={handleWordUpload}
                className="w-full text-sm text-slate-600 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-xs file:font-bold file:bg-[#1A4A5A] file:text-white hover:file:bg-[#0E2C40] file:cursor-pointer bg-slate-50 border border-slate-200 rounded-xl p-1.5 transition-all hover:border-slate-300 shadow-sm"
              />
              {wordFile && (
                <p className="text-xs text-[#148D8D] font-semibold mt-2 flex items-center gap-1">
                  <FileText className="w-4 h-4" /> Plantilla cargada: {wordFile.name}
                </p>
              )}
            </div>
          )}

          {/* PASO 3: FORMULARIO ESPACIOSO Y CENTRADO */}
          {paso >= 3 && (
            <form onSubmit={handleSubmitFinal} className="space-y-6 animate-fade-in border-t border-[#148D8D]/15 pt-4">
              
              <div className="flex items-center gap-1.5 text-xs font-bold text-[#148D8D] uppercase tracking-wider">
                <Sparkles className="w-4 h-4" />
                <span>PASO 3: Confirmación de Datos Operativos</span>
              </div>

              <div>
                <label className="block text-[10px] font-bold uppercase tracking-widest text-[#1A4A5A] mb-1">
                  Código o Nombre de Máquina (Presiona ENTER para refrescar)
                </label>
                <div className="relative">
                  <input
                    type="text"
                    placeholder="Ej. MONTACARGA YALE"
                    value={codigo}
                    onChange={(e) => setCodigo(e.target.value)}
                    onKeyDown={handleSearchManual}
                    className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-[#0E2C40] font-semibold text-sm focus:outline-none focus:ring-2 focus:ring-[#148D8D] shadow-sm pr-10"
                  />
                  <div className="absolute right-3 top-2.5 text-[#148D8D]">
                    {searching ? <RefreshCw className="w-4 h-4 animate-spin" /> : <Search className="w-4 h-4" />}
                  </div>
                </div>
              </div>

              {/* Grid Ampliado con texto centrado en campos clave */}
              <div className="bg-[#148D8D]/5 p-5 rounded-xl border border-[#148D8D]/20 space-y-4">
                
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  <div>
                    <label className="block text-[10px] font-bold uppercase tracking-widest text-[#1A4A5A] mb-1 text-center">
                      Orden de Trabajo (OT)
                    </label>
                    <input
                      type="text"
                      value={ordenTrabajo}
                      onChange={(e) => setOrdenTrabajo(e.target.value)}
                      className="w-full px-3 py-2.5 bg-white border border-slate-200 rounded-lg text-[#0E2C40] font-extrabold text-sm text-center outline-none focus:border-[#148D8D] shadow-sm"
                    />
                  </div>

                  <div>
                    <label className="block text-[10px] font-bold uppercase tracking-widest text-[#1A4A5A] mb-1 text-center">
                      Fecha Ejecución
                    </label>
                    <input
                      type="text"
                      value={fecha}
                      onChange={(e) => setFecha(e.target.value)}
                      className="w-full px-3 py-2.5 bg-white border border-slate-200 rounded-lg text-[#0E2C40] font-semibold text-sm text-center outline-none focus:border-[#148D8D] shadow-sm"
                    />
                  </div>

                  <div>
                    <label className="block text-[10px] font-bold uppercase tracking-widest text-[#1A4A5A] mb-1 text-center">
                      Horas Reales
                    </label>
                    <input
                      type="text"
                      value={horas}
                      onChange={(e) => setHoras(e.target.value)}
                      className="w-full px-3 py-2.5 bg-white border border-slate-200 rounded-lg text-[#0E2C40] font-semibold text-sm text-center outline-none focus:border-[#148D8D] shadow-sm"
                    />
                  </div>
                </div>

                {/* Sección de los 4 Técnicos */}
                <div className="border-t border-[#148D8D]/15 pt-4">
                  <label className="block text-[10px] font-bold uppercase tracking-widest text-[#148D8D] mb-3 text-center">
                    Técnicos Responsables Asignados
                  </label>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    <div>
                      <label className="block text-[9px] font-bold text-slate-500 uppercase mb-1">Técnico 1</label>
                      <input
                        type="text"
                        value={tecnico1}
                        onChange={(e) => setTecnico1(e.target.value)}
                        className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg text-[#0E2C40] font-medium text-xs outline-none focus:border-[#148D8D] shadow-sm"
                      />
                    </div>

                    <div>
                      <label className="block text-[9px] font-bold text-slate-500 uppercase mb-1">Técnico 2</label>
                      <input
                        type="text"
                        value={tecnico2}
                        placeholder="Opcional"
                        onChange={(e) => setTecnico2(e.target.value)}
                        className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg text-[#0E2C40] font-medium text-xs outline-none focus:border-[#148D8D] shadow-sm"
                      />
                    </div>

                    <div>
                      <label className="block text-[9px] font-bold text-slate-500 uppercase mb-1">Técnico 3</label>
                      <input
                        type="text"
                        value={tecnico3}
                        placeholder="Opcional"
                        onChange={(e) => setTecnico3(e.target.value)}
                        className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg text-[#0E2C40] font-medium text-xs outline-none focus:border-[#148D8D] shadow-sm"
                      />
                    </div>

                    <div>
                      <label className="block text-[9px] font-bold text-slate-500 uppercase mb-1">Técnico 4</label>
                      <input
                        type="text"
                        value={tecnico4}
                        placeholder="Opcional"
                        onChange={(e) => setTecnico4(e.target.value)}
                        className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg text-[#0E2C40] font-medium text-xs outline-none focus:border-[#148D8D] shadow-sm"
                      />
                    </div>
                  </div>
                </div>

              </div>

              {/* Botón Principal */}
              <button
                type="submit"
                disabled={!wordFile || loadingDoc}
                className={`w-full py-4 rounded-xl text-xs font-extrabold tracking-wider uppercase transition-all duration-300 shadow-md flex items-center justify-center gap-2 ${
                  !wordFile || loadingDoc
                    ? 'bg-slate-200 text-slate-400 cursor-not-allowed shadow-none'
                    : 'bg-[#EFBC75] text-[#0E2C40] hover:bg-[#e0b06b] hover:shadow-lg active:scale-[0.995]'
                }`}
              >
                {loadingDoc ? (
                  <>
                    <Loader2 className="w-4 h-4 animate-spin text-[#0E2C40]" />
                    <span>Generando Documento Word...</span>
                  </>
                ) : (
                  <>
                    <Download className="w-4 h-4" />
                    <span>Generar y Descargar Ficha en Word (.docx)</span>
                  </>
                )}
              </button>
            </form>
          )}

          {/* Mensajes de Feedback */}
          {statusMessage.text && (
            <div className={`mt-4 text-xs font-bold p-4 rounded-xl border transition-all duration-300 text-center flex items-center justify-center gap-2 ${
              statusMessage.type === 'success'
                ? 'bg-emerald-100 border-emerald-300 text-emerald-800'
                : statusMessage.type === 'error'
                ? 'bg-rose-100 border-rose-300 text-rose-800'
                : 'bg-[#148D8D]/10 border-[#148D8D]/30 text-[#1A4A5A]'
            }`}>
              {statusMessage.type === 'success' ? (
                <CheckCircle className="w-4 h-4 shrink-0" />
              ) : statusMessage.type === 'error' ? (
                <AlertCircle className="w-4 h-4 shrink-0" />
              ) : (
                <Sparkles className="w-4 h-4 shrink-0" />
              )}
              <span>{statusMessage.text}</span>
            </div>
          )}

        </div>

      </div>
    </main>
  );
}
'use client';

import React, { useState, useEffect, useRef } from 'react';
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
  const [diagnostico, setDiagnostico] = useState('');
  
  // Estados para Tipo de Mantenimiento, Nota y Área
  const [tipoMantenimiento, setTipoMantenimiento] = useState<'INTERNO' | 'EXTERNO' | ''>('');
  const [nota, setNota] = useState('');
  const [area, setArea] = useState('');

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
  const [listaMaquinas, setListaMaquinas] = useState<string[]>([]);
  
  // Estados y Referencias para el Buscador Predictivo
  const [mostrarSugerencias, setMostrarSugerencias] = useState(false);
  const [verTodo, setVerTodo] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);

  const [backendUrl, setBackendUrl] = useState<string>('http://localhost:8000');

  // EFECTO 1: Detectar URL del backend
  useEffect(() => {
    setJsReady(true);
    let urlEnvLocal = process.env.NEXT_PUBLIC_API_URL;
    if (urlEnvLocal && urlEnvLocal.trim() !== '') {
      urlEnvLocal = urlEnvLocal.trim().replace(/\/$/, '');
      setBackendUrl(urlEnvLocal);
    } else {
      setBackendUrl('http://localhost:8000');
    }
  }, []);

  // EFECTO 2: Detector de clics fuera del cuadro de búsqueda para ocultarlo
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setMostrarSugerencias(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

// MOTOR DE BÚSQUEDA PREDICTIVO
  const maquinasFiltradas = listaMaquinas.filter((maq) =>
    maq.toLowerCase().includes(codigo.toLowerCase())
  );
  // Si 'verTodo' está activo, mostramos el catálogo entero. Si no, mostramos lo que filtró.
  const opcionesAMostrar = verTodo ? listaMaquinas : maquinasFiltradas;

  const mapearDatosFormulario = (data: any) => {
    setOrdenTrabajo(data.orden_trabajo || '');
    setTecnico1(data.tecnico1 || '');
    setTecnico2(data.tecnico2 || '');
    setTecnico3(data.tecnico3 || '');
    setTecnico4(data.tecnico4 || '');
    setFecha(data.fecha || '');
    setHoras(data.horas || '');

    if (data.diagnostico !== undefined) setDiagnostico(data.diagnostico || '');
    if (data.nota !== undefined) setNota(data.nota || '');
    if (data.area !== undefined) setArea(data.area || '');
    if (data.tipo_mantenimiento) setTipoMantenimiento(data.tipo_mantenimiento as 'INTERNO' | 'EXTERNO');
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
        const response = await fetch(`${backendUrl}/api/cargar-excel`, {
          method: 'POST',
          body: formData,
        });

        const data = await response.json();
        if (!response.ok) throw new Error(data.detail || 'Error al procesar el Excel.');

        if (data.maquinas) {
          setListaMaquinas(data.maquinas);
        }

        setStatusMessage({ type: 'success', text: '¡Matriz Excel procesada con éxito! Puedes continuar con el Paso 2.' });
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
        const response = await fetch(`${backendUrl}/api/procesar-word-auto`, {
          method: 'POST',
          body: formData,
        });

        const data = await response.json();
        if (!response.ok) throw new Error(data.detail || 'Error al leer la plantilla Word.');

        if (data.autodetectado && data.data) {
          setCodigo(data.codigoDetectado || '');
          mapearDatosFormulario(data.data);

          if (data.data.alerta_vacio) {
            setStatusMessage({
              type: 'info',
              text: `⚠️ Máquina detectada en Excel, pero NO tiene datos registrados. Complétalos manualmente.`
            });
          } else {
            const textoOt = data.data.orden_trabajo ? `Orden de Trabajo (${data.data.orden_trabajo})` : 'datos operativos';
            setStatusMessage({
              type: 'success',
              text: `✨ ¡Máquina detectada! ${textoOt} autocompletados correctamente.`
            });
          }
        } else {
          setCodigo(data.codigoDetectado || '');
          setStatusMessage({
            type: 'info',
            text: 'Confirma el nombre de la máquina en el Paso 3 y presiona ENTER para buscarla.'
          });
        }
      } catch (error: any) {
        setStatusMessage({ type: 'error', text: error.message });
      } finally {
        setProcessingWord(false);
      }
    }
  };

  const buscarMaquina = async (codigoBuscado: string) => {
    if (!codigoBuscado.trim()) return;

    setSearching(true);
    setStatusMessage({ type: '', text: '' });

    try {
      const response = await fetch(`${backendUrl}/api/consultar-maquina`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ codigo: codigoBuscado.trim() }),
      });

      const data = await response.json();
      if (!response.ok) throw new Error(data.detail || 'No se encontró la máquina.');

      mapearDatosFormulario(data.data);

      if (data.data.alerta_vacio) {
        setStatusMessage({
          type: 'info',
          text: '⚠️ La máquina existe, pero sus columnas están vacías en la matriz Excel.'
        });
      } else {
        setStatusMessage({
          type: 'success',
          text: '¡Datos de la máquina encontrados y autocompletados!'
        });
      }
    } catch (error: any) {
      setStatusMessage({ type: 'error', text: error.message || 'Máquina no encontrada en el Excel.' });
    } finally {
      setSearching(false);
    }
  };

  const handleSearchManual = async (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      setMostrarSugerencias(false); // Ocultamos lista al presionar enter
      buscarMaquina(codigo);
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
    formData.append('diagnostico', diagnostico);
    formData.append('tipo_mantenimiento', tipoMantenimiento);
    formData.append('nota', nota);
    formData.append('area', area); // Área incluida aquí

    try {
      const response = await fetch(`${backendUrl}/api/generar-ficha`, {
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
      a.download = `${ordenTrabajo || 'Ficha'} (${codigo || 'Maquina'}).docx`;
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

      {/* Tarjeta Central (max-w-6xl) */}
      <div className="w-full max-w-[1400px] bg-white rounded-2xl shadow-xl shadow-slate-900/5 border border-[#148D8D]/20 px-6 py-6 md:px-10 md:py-8 relative transition-all duration-300">

        {/* Encabezado Principal */}
        <div className="text-center mb-8 border-b border-[#148D8D]/20 pb-6">
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
            <span className={`inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-[10px] font-bold uppercase tracking-wider border transition-all duration-300 ${jsReady
              ? 'bg-emerald-600/10 text-emerald-700 border-emerald-600/20'
              : 'bg-amber-600/10 text-amber-700 border-amber-600/20'
              }`}>
              <span className={`w-1.5 h-1.5 rounded-full ${jsReady ? 'bg-emerald-500 animate-pulse' : 'bg-amber-500'}`} />
              {jsReady ? 'Conectado' : 'Sincronizando'}
            </span>
          </div>
        </div>

        {/* SISTEMA DE CUADRÍCULA (GRID) */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start">

          {/* COLUMNA IZQUIERDA: Carga de Archivos */}
          <div className="lg:col-span-5 space-y-6">
            <div className="bg-slate-50 p-5 rounded-xl border border-slate-200 shadow-sm">
              <label className="block text-[11px] font-bold uppercase tracking-widest text-[#1A4A5A] mb-3 flex items-center justify-between">
                <span>PASO 1: Plan Anual (.xlsx)</span>
                {uploadingExcel && <Loader2 className="w-3.5 h-3.5 animate-spin text-[#148D8D]" />}
              </label>
              <input
                type="file"
                accept=".xlsx"
                onChange={handleExcelUpload}
                className="w-full text-sm text-slate-600 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-[11px] file:font-bold file:bg-[#1A4A5A] file:text-white hover:file:bg-[#0E2C40] file:cursor-pointer bg-white border border-slate-200 rounded-xl p-1.5 transition-all hover:border-slate-300 shadow-sm"
              />
              {excelFile && (
                <p className="text-[11px] text-[#148D8D] font-bold mt-3 flex items-center gap-1.5 bg-[#148D8D]/5 p-2 rounded-lg">
                  <FileSpreadsheet className="w-4 h-4 shrink-0" /> <span className="truncate">{excelFile.name}</span>
                </p>
              )}
            </div>

            <div className="bg-slate-50 p-5 rounded-xl border border-slate-200 shadow-sm animate-fade-in">
              <label className="block text-[11px] font-bold uppercase tracking-widest text-[#1A4A5A] mb-3 flex items-center justify-between">
                <span>PASO 2: Plantilla Anterior (.docx)</span>
                {processingWord && <Loader2 className="w-3.5 h-3.5 animate-spin text-[#148D8D]" />}
              </label>
              <input
                type="file"
                accept=".docx"
                onChange={handleWordUpload}
                className="w-full text-sm text-slate-600 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-[11px] file:font-bold file:bg-[#1A4A5A] file:text-white hover:file:bg-[#0E2C40] file:cursor-pointer bg-white border border-slate-200 rounded-xl p-1.5 transition-all hover:border-slate-300 shadow-sm"
              />
              {wordFile && (
                <p className="text-[11px] text-[#148D8D] font-bold mt-3 flex items-center gap-1.5 bg-[#148D8D]/5 p-2 rounded-lg">
                  <FileText className="w-4 h-4 shrink-0" /> <span className="truncate">{wordFile.name}</span>
                </p>
              )}
            </div>
          </div>

          {/* COLUMNA DERECHA: Formulario */}
          <div className="lg:col-span-7">
            <form onSubmit={handleSubmitFinal} className="animate-fade-in h-full flex flex-col justify-between">

              <div className="flex items-center gap-1.5 text-xs font-bold text-[#148D8D] uppercase tracking-wider mb-5">
                <Sparkles className="w-4 h-4" />
                <span>PASO 3: Confirmación de Datos Operativos</span>
              </div>

{/* BARRA DE BÚSQUEDA PREDICTIVA */}
              <div className="mb-5 relative" ref={dropdownRef}>
                <label className="block text-[10px] font-bold uppercase tracking-widest text-[#1A4A5A] mb-1.5">
                  Código o Nombre de Máquina
                </label>
                <div className="flex gap-2 relative">
                  <div className="relative flex-1">
                    <input
                      type="text"
                      placeholder="Escribe para buscar (Ej. ACUÑADORA)..."
                      value={codigo}
                      onChange={(e) => {
                        setCodigo(e.target.value);
                        setMostrarSugerencias(true);
                        setVerTodo(false); // Si empieza a escribir, vuelve a filtrar
                      }}
                      onFocus={() => {
                        setMostrarSugerencias(true);
                        setVerTodo(false); // Si hace clic en la barra, empieza a filtrar
                      }}
                      onKeyDown={handleSearchManual}
                      className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-[#0E2C40] font-semibold text-sm focus:outline-none focus:ring-2 focus:ring-[#148D8D] shadow-sm pr-10"
                    />
                    <div className="absolute right-3 top-2.5 text-[#148D8D]">
                      {searching ? <RefreshCw className="w-4 h-4 animate-spin" /> : <Search className="w-4 h-4" />}
                    </div>

                    {/* CUADRO FLOTANTE DE SUGERENCIAS */}
                    {mostrarSugerencias && listaMaquinas.length > 0 && (
                      <ul className="absolute z-50 w-full bg-white border border-[#148D8D]/30 rounded-xl mt-2 max-h-60 overflow-y-auto shadow-2xl overflow-hidden divide-y divide-slate-100">
                        {opcionesAMostrar.length > 0 ? (
                          opcionesAMostrar.map((maq, index) => {
                            const esSeleccionada = maq === codigo;
                            return (
                              <li
                                key={index}
                                id={esSeleccionada ? 'maq-activa' : undefined} // ID especial para enfocarla
                                className={`px-4 py-3 text-[11px] font-bold cursor-pointer transition-colors ${
                                  esSeleccionada
                                    ? 'bg-[#148D8D] text-white' // Si es la seleccionada, se pinta de verde
                                    : 'text-[#0E2C40] hover:bg-[#148D8D]/10 hover:text-[#148D8D]'
                                }`}
                                onClick={() => {
                                  setCodigo(maq);
                                  setMostrarSugerencias(false);
                                  setVerTodo(false);
                                  buscarMaquina(maq);
                                }}
                              >
                                {maq}
                              </li>
                            );
                          })
                        ) : (
                          <li className="px-4 py-3 text-[11px] text-slate-400 italic text-center">
                            No se encontraron coincidencias
                          </li>
                        )}
                      </ul>
                    )}
                  </div>

{/* BOTÓN PARA ABRIR/CERRAR CATÁLOGO (INTERRUPTOR) */}
                  <button
                    type="button"
                    className="relative w-12 flex-shrink-0 bg-slate-100 border border-slate-200 hover:bg-[#148D8D]/10 hover:border-[#148D8D]/30 rounded-xl flex items-center justify-center transition-colors cursor-pointer group"
                    title="Alternar lista completa de máquinas"
                    onClick={() => {
                      if (!mostrarSugerencias) {
                        setVerTodo(true); 
                        setMostrarSugerencias(true);
                        
                        // Salto instantáneo y directo (sin animaciones lentas)
                        setTimeout(() => {
                          const activeEl = document.getElementById('maq-activa');
                          if (activeEl) {
                            // 'auto' hace que salte de golpe en lugar de deslizarse
                            activeEl.scrollIntoView({ behavior: 'auto', block: 'center' });
                          }
                        }, 0); // Tiempo 0 para que ocurra en el mismo instante que se abre
                      } else {
                        setMostrarSugerencias(false);
                        setVerTodo(false);
                      }
                    }}
                  >
                    <svg xmlns="http://www.w3.org/2000/svg" className="w-6 h-6 text-slate-500 group-hover:text-[#148D8D]" viewBox="0 0 20 20" fill="currentColor">
                      <path fillRule="evenodd" d="M5.293 7.293a1 1 0 011.414 0L10 10.586l3.293-3.293a1 1 0 111.414 1.414l-4 4a1 1 0 01-1.414 0l-4-4a1 1 0 010-1.414z" clipRule="evenodd" />
                    </svg>
                  </button>
                </div>
              </div>

              {/* Contenedor Verde del Formulario */}
              <div className="bg-[#148D8D]/5 p-5 rounded-xl border border-[#148D8D]/20 mb-6">
                
                {/* Clasificación del Mantenimiento */}
                <div className="flex flex-col items-center justify-center mb-4 border-b border-[#148D8D]/15 pb-4">
                  <label className="block text-[10px] font-bold uppercase tracking-widest text-[#1A4A5A] mb-2">
                    Clasificación del Mantenimiento
                  </label>
                  <div className="flex gap-6 p-2.5 bg-white border border-slate-200 rounded-lg shadow-sm w-full md:w-auto justify-center">
                    <label className="flex items-center gap-2 cursor-pointer">
                      <input type="radio" name="tipoMantenimiento" checked={tipoMantenimiento === 'INTERNO'} onChange={() => setTipoMantenimiento('INTERNO')} className="w-3.5 h-3.5 accent-[#148D8D]" />
                      <span className="text-[11px] font-bold text-[#0E2C40]">TIPO I - INTERNO</span>
                    </label>
                    <label className="flex items-center gap-2 cursor-pointer">
                      <input type="radio" name="tipoMantenimiento" checked={tipoMantenimiento === 'EXTERNO'} onChange={() => setTipoMantenimiento('EXTERNO')} className="w-3.5 h-3.5 accent-[#148D8D]" />
                      <span className="text-[11px] font-bold text-[#0E2C40]">TIPO II - EXTERNO</span>
                    </label>
                  </div>
                </div>

                {/* Fila 1: OT, Área, Fecha, Horas */}
                <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-4">
                  <div>
                    <label className="block text-[10px] font-bold uppercase tracking-widest text-[#1A4A5A] mb-1 text-center">
                      Orden (OT)
                    </label>
                    <input type="text" value={ordenTrabajo} onChange={(e) => setOrdenTrabajo(e.target.value)} className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg text-[#0E2C40] font-extrabold text-sm text-center outline-none focus:border-[#148D8D] shadow-sm" />
                  </div>
                  <div>
                    <label className="block text-[10px] font-bold uppercase tracking-widest text-[#1A4A5A] mb-1 text-center">
                      Área
                    </label>
                    <input type="text" value={area} onChange={(e) => setArea(e.target.value)} className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg text-[#0E2C40] font-semibold text-sm text-center outline-none focus:border-[#148D8D] shadow-sm" />
                  </div>
                  <div>
                    <label className="block text-[10px] font-bold uppercase tracking-widest text-[#1A4A5A] mb-1 text-center">
                      Fecha
                    </label>
                    <input type="text" value={fecha} onChange={(e) => setFecha(e.target.value)} className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg text-[#0E2C40] font-semibold text-sm text-center outline-none focus:border-[#148D8D] shadow-sm" />
                  </div>
                  <div>
                    <label className="block text-[10px] font-bold uppercase tracking-widest text-[#1A4A5A] mb-1 text-center">
                      Horas Reales
                    </label>
                    <input type="text" value={horas} onChange={(e) => setHoras(e.target.value)} className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg text-[#0E2C40] font-semibold text-sm text-center outline-none focus:border-[#148D8D] shadow-sm" />
                  </div>
                </div>

                {/* Fila 2: Técnicos */}
                <div className="border-t border-[#148D8D]/15 pt-4 mb-4">
                  <label className="block text-[10px] font-bold uppercase tracking-widest text-[#148D8D] mb-3 text-center">
                    Técnicos Responsables
                  </label>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    <div>
                      <label className="block text-[9px] font-bold text-slate-500 uppercase mb-1">Técnico 1</label>
                      <input type="text" value={tecnico1} onChange={(e) => setTecnico1(e.target.value)} className="w-full px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-[#0E2C40] font-medium text-xs outline-none focus:border-[#148D8D]" />
                    </div>
                    <div>
                      <label className="block text-[9px] font-bold text-slate-500 uppercase mb-1">Técnico 2</label>
                      <input type="text" value={tecnico2} placeholder="Opcional" onChange={(e) => setTecnico2(e.target.value)} className="w-full px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-[#0E2C40] font-medium text-xs outline-none focus:border-[#148D8D]" />
                    </div>
                    <div>
                      <label className="block text-[9px] font-bold text-slate-500 uppercase mb-1">Técnico 3</label>
                      <input type="text" value={tecnico3} placeholder="Opcional" onChange={(e) => setTecnico3(e.target.value)} className="w-full px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-[#0E2C40] font-medium text-xs outline-none focus:border-[#148D8D]" />
                    </div>
                    <div>
                      <label className="block text-[9px] font-bold text-slate-500 uppercase mb-1">Técnico 4</label>
                      <input type="text" value={tecnico4} placeholder="Opcional" onChange={(e) => setTecnico4(e.target.value)} className="w-full px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-[#0E2C40] font-medium text-xs outline-none focus:border-[#148D8D]" />
                    </div>
                  </div>
                </div>

                {/* Fila 3: Diagnóstico */}
                <div className="border-t border-[#148D8D]/15 pt-4">
                  <label className="block text-[10px] font-bold uppercase tracking-widest text-[#1A4A5A] mb-2 text-center">
                    Diagnóstico de Falla
                  </label>
                  <textarea
                    value={diagnostico}
                    onChange={(e) => setDiagnostico(e.target.value)}
                    rows={3}
                    placeholder="El diagnóstico autodetectado aparecerá aquí..."
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg text-[#0E2C40] font-medium text-xs outline-none focus:border-[#148D8D] resize-none"
                  />
                </div>

                {/* Fila 4: Nota Adicional */}
                <div className="border-t border-[#148D8D]/15 pt-4 mt-4">
                  <label className="block text-[10px] font-bold uppercase tracking-widest text-[#1A4A5A] mb-2 text-center">
                    Nota Adicional del Documento
                  </label>
                  <textarea
                    value={nota}
                    onChange={(e) => setNota(e.target.value)}
                    rows={2}
                    placeholder="La nota inferior de la ficha base se extraerá automáticamente aquí..."
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-[#0E2C40] font-medium text-xs outline-none focus:border-[#148D8D] resize-none"
                  />
                </div>
              </div>

              {/* Botón Principal (Abajo) */}
              <button
                type="submit"
                disabled={!wordFile || loadingDoc}
                className={`w-full mt-auto py-4 rounded-xl text-xs font-extrabold tracking-wider uppercase transition-all duration-300 shadow-md flex items-center justify-center gap-2 ${!wordFile || loadingDoc
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
          </div>
        </div>

        {/* MENSAJES INFERIORES */}
        {statusMessage.text && (
          <div className={`mt-6 text-xs font-bold p-4 rounded-xl border transition-all duration-300 text-center flex items-center justify-center gap-2 ${statusMessage.type === 'success'
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
    </main>
  );
}
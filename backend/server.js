console.log("🔄 Cargando backend de Fichas de Mantenimiento...");

const express = require('express');
const cors = require('cors');
const multer = require('multer');
const XLSX = require('xlsx');
const mammoth = require('mammoth');

// Configuración para variables de entorno (Opcional en local, útil en Render)
const dotenv = require('dotenv');
dotenv.config(); 
dotenv.config({ path: '.env.local' }); 

const { reconociendoMaquinaInteligente, consultarDatosMaquina, procesarFicha, obtenerListaMaquinas } = require('./automatizar_fichas');

const app = express();
const PORT = process.env.PORT || 8000;

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// Permitir conexiones desde cualquier Frontend
app.use(cors({
    origin: process.env.FRONTEND_URL || '*', 
    methods: ['GET', 'POST', 'PUT', 'DELETE'],
    credentials: true
}));

app.use(express.json());

let excelCacheJson = null;

app.get('/', (req, res) => {
    res.json({ status: "online", service: "API Fichas de Mantenimiento", version: "3.1.0" });
});

// PASO 1: Cargar Excel
app.post('/api/cargar-excel', upload.single('excel'), (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ detail: "Sube un archivo Excel (.xlsx) válido." });

        const workbook = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
        if (!workbook.SheetNames.includes('PLAN.MANT')) return res.status(400).json({ detail: "El Excel no contiene la hoja 'PLAN.MANT'." });

        const sheet = workbook.Sheets['PLAN.MANT'];
        excelCacheJson = XLSX.utils.sheet_to_json(sheet, { header: 1, range: 9 });
        const maquinasDisponibles = obtenerListaMaquinas(excelCacheJson);

        return res.json({ 
            success: true, 
            message: "Matriz Excel procesada correctamente.",
            totalFilas: excelCacheJson.length,
            maquinas: maquinasDisponibles
        });
    } catch (error) {
        return res.status(500).json({ detail: `Error procesando Excel: ${error.message}` });
    }
});

// PASO 2: Recibir Word y reconocer máquina automáticamente
app.post('/api/procesar-word-auto', upload.single('word'), async (req, res) => {
    try {
        if (!excelCacheJson) {
            return res.status(400).json({ detail: "Primero debes cargar la matriz Excel en el Paso 1." });
        }

        if (!req.file) {
            return res.status(400).json({ detail: "Sube un archivo Word (.docx) válido." });
        }

        const textResult = await mammoth.extractRawText({ buffer: req.file.buffer });
        const textoDoc = textResult.value || "";
        const nombreArchivo = req.file.originalname || "";

        const resultado = reconociendoMaquinaInteligente(textoDoc, nombreArchivo, excelCacheJson);

        // 1. EXTRAER LA NOTA
        let notaExtraida = "";
        const matchNota = textoDoc.match(/Nota:\s*([\s\S]+)$/i); 
        if (matchNota) {
            notaExtraida = matchNota[1].trim(); 
        }

        // 2. EXTRAER EL TIPO DE MANTENIMIENTO (LECTURA QUIRÚRGICA DE XML)
        let tipoExtraido = "";
        try {
            const PizZip = require('pizzip');
            const zip = new PizZip(req.file.buffer);
            let documentXml = zip.file("word/document.xml").asText();
            
            // Aislamos el encabezado para no confundirnos con tablas del cuerpo
            let headerEndIndex = documentXml.search(/ORDEN(?:<[^>]+>|[\s\t\r\n])*DE(?:<[^>]+>|[\s\t\r\n])*TRABAJO/i);
            if (headerEndIndex === -1) headerEndIndex = Math.min(4000, documentXml.length);

            let headerXml = documentXml.substring(0, headerEndIndex);
            let cajitasEncontradas = 0;
            
            // Buscamos todas las celdas o cajas de texto del encabezado
            const regexCajas = /<(w:tc|w:txbxContent)(>| [^>]*>)([\s\S]*?)<\/\1>/g;
            let matchCaja;
            
            while ((matchCaja = regexCajas.exec(headerXml)) !== null) {
                // Limpiamos el texto que está adentro de la celda
                const textoLimpio = matchCaja[3].replace(/<[^>]+>/g, '').trim().toUpperCase();
                
                // Si la celda está vacía o tiene una "X", la consideramos un checkbox
                if (textoLimpio === '' || textoLimpio === 'X' || textoLimpio === '[X]') {
                    cajitasEncontradas++;
                    
                    // La primera caja que encontremos con "X" es INTERNO, la segunda es EXTERNO
                    if (cajitasEncontradas === 1 && textoLimpio.includes('X')) {
                        tipoExtraido = "INTERNO";
                    } else if (cajitasEncontradas === 2 && textoLimpio.includes('X')) {
                        tipoExtraido = "EXTERNO";
                    }
                }
            }
        } catch (err) {
            console.error("No se pudo escanear el XML para detectar el tipo de mantenimiento:", err);
        }

        // 3. EMPAQUETAR Y ENVIAR AL FRONTEND
        // Aseguramos de enviar los datos aunque la máquina no se haya autodetectado
        const datosParaEnviar = resultado.datos || {};
        datosParaEnviar.nota = notaExtraida;
        datosParaEnviar.tipo_mantenimiento = tipoExtraido;

        if (resultado.encontrado) {
            return res.json({
                success: true,
                autodetectado: true,
                codigoDetectado: resultado.maquinaNombre,
                metodo: resultado.metodo,
                data: datosParaEnviar
            });
        } else {
            return res.json({
                success: true,
                autodetectado: false,
                codigoDetectado: "",
                message: resultado.mensaje,
                data: datosParaEnviar
            });
        }

    } catch (error) {
        return res.status(500).json({ detail: `Error procesando el documento Word: ${error.message}` });
    }
});

// PASO 3: Búsqueda manual
app.post('/api/consultar-maquina', (req, res) => {
    try {
        const { codigo } = req.body;
        if (!codigo) return res.status(400).json({ detail: "Ingresa el código o nombre de la máquina." });

        const datos = consultarDatosMaquina(codigo, excelCacheJson);
        return res.json({ success: true, data: datos });
    } catch (error) {
        return res.status(404).json({ detail: error.message });
    }
});

// PASO FINAL: Generar documento Word
app.post('/api/generar-ficha', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ detail: "No se ha subido el documento Word base." });

        // --- NUEVO: RECEPCIÓN DE TIPO Y NOTA ---
        const datosFormulario = {
            codigo: req.body.codigo || '',
            orden_trabajo: req.body.orden_trabajo || '',
            tecnico1: req.body.tecnico1 || '',
            tecnico2: req.body.tecnico2 || '',
            tecnico3: req.body.tecnico3 || '',
            tecnico4: req.body.tecnico4 || '',
            fecha: req.body.fecha || '',
            horas: req.body.horas || '',
            diagnostico: req.body.diagnostico || '',
            tipo_mantenimiento: req.body.tipo_mantenimiento || '',
            nota: req.body.nota || ''
        };

        const resultado = await procesarFicha(req.file.buffer, datosFormulario);

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename="${resultado.nombreSalida}.docx"`);
        return res.send(resultado.bufferWord);

    } catch (error) {
        console.error("Error generando ficha:", error);
        return res.status(500).json({ detail: `Error interno: ${error.message}` });
    }
});

// ESTO MANTIENE EL SERVIDOR CORRIENDO:
app.listen(PORT, () => {
    console.log(`🚀 Servidor Node.js corriendo en el puerto ${PORT}`);
});
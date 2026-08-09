console.log("🔄 Cargando backend de Fichas de Mantenimiento...");

const express = require('express');
const cors = require('cors');
const multer = require('multer');
const XLSX = require('xlsx');
const mammoth = require('mammoth');
const { reconociendoMaquinaInteligente, consultarDatosMaquina, procesarFicha, obtenerListaMaquinas } = require('./automatizar_fichas');

const app = express();
const PORT = process.env.PORT || 8000;

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });



app.use(cors());
app.use(express.json());

let excelCacheJson = null;

// Endpoint de prueba
app.get('/', (req, res) => {
    res.json({ status: "online", service: "API Fichas de Mantenimiento", version: "3.0.0" });
});

// PASO 1: Cargar Excel
app.post('/api/cargar-excel', upload.single('excel'), (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ detail: "Sube un archivo Excel (.xlsx) válido." });
        }

        const workbook = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
        if (!workbook.SheetNames.includes('PLAN.MANT')) {
            return res.status(400).json({ detail: "El Excel no contiene la hoja 'PLAN.MANT'." });
        }

        const sheet = workbook.Sheets['PLAN.MANT'];
        excelCacheJson = XLSX.utils.sheet_to_json(sheet, { header: 1, range: 9 });

        // Extraemos la lista de máquinas usando la nueva función
        const maquinasDisponibles = obtenerListaMaquinas(excelCacheJson);

        return res.json({ 
            success: true, 
            message: "Matriz Excel procesada correctamente.",
            totalFilas: excelCacheJson.length,
            maquinas: maquinasDisponibles // Enviamos la lista al frontend
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

        if (resultado.encontrado) {
            return res.json({
                success: true,
                autodetectado: true,
                codigoDetectado: resultado.maquinaNombre,
                metodo: resultado.metodo,
                data: resultado.datos
            });
        } else {
            return res.json({
                success: true,
                autodetectado: false,
                codigoDetectado: "",
                message: resultado.mensaje
            });
        }

    } catch (error) {
        return res.status(500).json({ detail: `Error procesando el documento Word: ${error.message}` });
    }
});

// PASO 3: Búsqueda manual si presiona ENTER en el formulario
app.post('/api/consultar-maquina', (req, res) => {
    try {
        const { codigo } = req.body;
        if (!codigo) {
            return res.status(400).json({ detail: "Ingresa el código o nombre de la máquina." });
        }

        const datos = consultarDatosMaquina(codigo, excelCacheJson);
        return res.json({ success: true, data: datos });
    } catch (error) {
        return res.status(404).json({ detail: error.message });
    }
});

// PASO FINAL: Generar documento Word
app.post('/api/generar-ficha', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ detail: "No se ha subido el documento Word base." });
        }

        const datosFormulario = {
            codigo: req.body.codigo || '',
            orden_trabajo: req.body.orden_trabajo || '',
            tecnico1: req.body.tecnico1 || '',
            tecnico2: req.body.tecnico2 || '',
            tecnico3: req.body.tecnico3 || '',
            tecnico4: req.body.tecnico4 || '',
            fecha: req.body.fecha || '',
            horas: req.body.horas || '',
            diagnostico: req.body.diagnostico || ''
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
    console.log(`🚀 Servidor Node.js corriendo en http://localhost:${PORT}`);
});
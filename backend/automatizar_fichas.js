const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const DocxTemplater = require('docx-templates');

const EXCEL_LOCAL_PATH = path.join(__dirname, 'docex', 'Prueba_Plan_Mantenimiento_2025.xlsx');
const CARPETA_SALIDA = path.join(__dirname, 'Fichas_Generadas');

/**
 * Normaliza textos (sin tildes, minúsculas/mayúsculas ni espacios extra)
 */
function normalizarTexto(str) {
    return String(str || '')
        .toUpperCase()
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
        .replace(/_/g, ' ')
        .replace(/[^\w\s\(\)]/gi, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}

/**
 * Carga la hoja 'PLAN.MANT' desde el Excel
 */
function obtenerDataframeExcel(localPath = EXCEL_LOCAL_PATH) {
    if (!fs.existsSync(localPath)) {
        throw new Error(`No se encontró el archivo Excel en la ruta: ${localPath}`);
    }

    const workbook = XLSX.readFile(localPath, { cellDates: true });
    if (!workbook.SheetNames.includes('PLAN.MANT')) {
        throw new Error("La hoja 'PLAN.MANT' no existe en el archivo Excel.");
    }

    const sheet = workbook.Sheets['PLAN.MANT'];
    // Carga todas las filas sin recortar encabezados fijos
    return XLSX.utils.sheet_to_json(sheet, { header: 1 });
}

/**
 * Extrae los datos mapeando dinámicamente las columnas según los encabezados
 */
function extraerCamposFila(fila, nombreMaquina, jsonRows) {
    // 1. Identificar índices de columnas según los encabezados (generalmente en fila 8 o 9)
    let idxObs = -1;
    let idxHoras = -1;
    let idxFecha = -1;
    let idxTecnico1 = -1, idxTecnico2 = -1, idxTecnico3 = -1, idxTecnico4 = -1;

    for (let r = 0; r < Math.min(12, jsonRows.length); r++) {
        const rowHeader = jsonRows[r];
        if (!rowHeader) continue;

        // Búsqueda prioritaria de la columna de Horas por su encabezado exacto
        rowHeader.forEach((cell, idx) => {
            const cellNorm = normalizarTexto(cell);

            if (cellNorm.includes("TECNICO 1")) idxTecnico1 = idx;
            else if (cellNorm.includes("TECNICO 2")) idxTecnico2 = idx;
            else if (cellNorm.includes("TECNICO 3")) idxTecnico3 = idx;
            else if (cellNorm.includes("TECNICO 4")) idxTecnico4 = idx;
            else if (cellNorm.includes("OBSERVACION")) idxObs = idx;
            // Prioridad a 'HORAS-HOMBRE' o 'H.H REALES'
            else if (cellNorm.includes("HORAS-HOMBRE") || cellNorm.includes("HORAS HOMBRE") || cellNorm.includes("H.H REALES")) idxHoras = idx;
            else if (cellNorm.includes("EJECUCION") || cellNorm.includes("FECHA")) idxFecha = idx;
        });
    }

    // Fallbacks por posición si no se hallaron por encabezado exacto
    if (idxObs === -1) idxObs = fila.length - 1; // La celda 'Realizado con OSS...' suele ser la última
    if (idxHoras === -1) idxHoras = 10;

    // 2. Extracción de Orden de Trabajo (Busca OSS seguido de dígitos)
    let ordenTrabajo = "S/N";
    // Buscar en la celda de observación identificada o en toda la fila si no
    for (let cellVal of fila) {
        if (cellVal && typeof cellVal === 'string') {
            const otMatch = cellVal.match(/OSS\d+/i);
            if (otMatch) {
                ordenTrabajo = otMatch[0].toUpperCase();
                break;
            }
        }
    }

    // 3. Horas Utilizadas
    // Obtener el valor de horas garantizando que sea la columna de Horas-Hombre
    const horasVal = (idxHoras !== -1 && fila[idxHoras] !== undefined) ? fila[idxHoras] : "0";

    // Si el valor extraído accidentalmente contiene texto como 'OSS' o 'Realizado', lo limpia
    let horasLimpio = String(horasVal).trim();
    if (horasLimpio.toUpperCase().includes("OSS") || horasLimpio.toUpperCase().includes("REALIZADO")) {
        horasLimpio = "0";
    }

    const horasStr = `${horasLimpio} HORAS`;

    // 4. Fecha de Ejecución
    let fechaStr = "DICIEMBRE - 2025";
    const fechaVal = idxFecha !== -1 ? fila[idxFecha] : null;
    if (fechaVal instanceof Date) {
        const meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
            "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];
        fechaStr = `${meses[fechaVal.getMonth()]} - ${fechaVal.getFullYear()}`;
    }

    // 5. Extracción Limpia de los 4 Técnicos Responsables
    const limpiarTecnico = (val) => {
        if (!val || typeof val !== 'string') return "";
        const str = val.trim();
        if (str.toLowerCase().includes("realizado con") || str.toLowerCase().includes("oss")) return "";
        return str;
    };

    let t1 = idxTecnico1 !== -1 ? limpiarTecnico(fila[idxTecnico1]) : "";
    let t2 = idxTecnico2 !== -1 ? limpiarTecnico(fila[idxTecnico2]) : "";
    let t3 = idxTecnico3 !== -1 ? limpiarTecnico(fila[idxTecnico3]) : "";
    let t4 = idxTecnico4 !== -1 ? limpiarTecnico(fila[idxTecnico4]) : "";

    // Si la búsqueda por encabezado no halló los índices exactos, buscamos celdas consecutivas con nombres
    if (!t1 && !t2) {
        const posiblesTecnicos = [];
        fila.forEach(cell => {
            if (cell && typeof cell === 'string') {
                const c = cell.trim();
                // Si parece un nombre propio (al menos 2 palabras en mayúsculas) y no es estado u OSS
                if (c.length > 5 && c.includes(' ') && !c.includes('OSS') && !c.includes('Realizado') && !c.includes('MONTACARGA')) {
                    posiblesTecnicos.push(c);
                }
            }
        });
        t1 = posiblesTecnicos[0] || "";
        t2 = posiblesTecnicos[1] || "";
        t3 = posiblesTecnicos[2] || "";
        t4 = posiblesTecnicos[3] || "";
    }

    return {
        codigo: nombreMaquina,
        orden_trabajo: ordenTrabajo,
        tecnico1: t1 || "MIGUEL GUERRERO",
        tecnico2: t2,
        tecnico3: t3,
        tecnico4: t4,
        fecha: fechaStr,
        horas: horasStr
    };
}

/**
 * Motor de Reconocimiento Inteligente
 */
function reconociendoMaquinaInteligente(textoWord, nombreArchivo, excelFuente = EXCEL_LOCAL_PATH) {
    let jsonRows = Array.isArray(excelFuente) ? excelFuente : obtenerDataframeExcel(excelFuente);
    const fuenteTextoCompleta = normalizarTexto(`${textoWord} ${nombreArchivo}`);

    const maquinasExcel = [];
    jsonRows.forEach((row, idx) => {
        if (row && row[1] && typeof row[1] === 'string') {
            const nombreOriginal = row[1].trim();
            const nombreNorm = normalizarTexto(nombreOriginal);
            if (nombreNorm && !nombreNorm.includes("MAQUINA") && !nombreNorm.includes("MANTENIMIENTO")) {
                maquinasExcel.push({ index: idx, row: row, nombreOriginal: nombreOriginal, nombreNorm: nombreNorm });
            }
        }
    });

    if (maquinasExcel.length === 0) {
        return { encontrado: false, mensaje: "La matriz Excel no contiene máquinas válidas." };
    }

    // CAPA 1: Código numérico o identificador en paréntesis
    const codigosEnWord = fuenteTextoCompleta.match(/\(([A-Z0-9\s-]+)\)/g) || [];
    for (let codRaw of codigosEnWord) {
        let cod = codRaw.replace(/[\(\)]/g, '').trim();
        if (cod.length >= 3) {
            const match = maquinasExcel.find(item => item.nombreNorm.includes(cod));
            if (match) {
                return {
                    encontrado: true,
                    metodo: "Código entre paréntesis",
                    maquinaNombre: match.nombreOriginal,
                    datos: extraerCamposFila(match.row, match.nombreOriginal, jsonRows)
                };
            }
        }
    }

    // CAPA 2: Coincidencia por Subcadena
    for (let item of maquinasExcel) {
        if (fuenteTextoCompleta.includes(item.nombreNorm) || item.nombreNorm.includes(normalizarTexto(nombreArchivo.replace(/\.docx$/i, '')))) {
            return {
                encontrado: true,
                metodo: "Nombre exacto / subcadena",
                maquinaNombre: item.nombreOriginal,
                datos: extraerCamposFila(item.row, item.nombreOriginal, jsonRows)
            };
        }
    }

    // CAPA 3: Palabras Clave
    const stopwords = new Set(["DE", "LA", "EL", "LOS", "LAS", "EN", "DEL", "Y", "O", "PARA", "TIPO", "EQUIPO", "MAQUINA", "CODIGO", "OSS"]);
    const tokensWord = normalizarTexto(fuenteTextoCompleta).split(/\s+/).filter(w => w.length >= 3 && !stopwords.has(w));

    let mejorCandidato = null;
    let maxPuntaje = 0;

    for (let item of maquinasExcel) {
        const tokensExcel = normalizarTexto(item.nombreNorm).split(/\s+/).filter(w => w.length >= 3 && !stopwords.has(w));
        if (tokensExcel.length === 0) continue;

        let coincidenciaCount = 0;
        tokensExcel.forEach(token => {
            if (tokensWord.includes(token)) coincidenciaCount += 1;
        });

        const puntaje = coincidenciaCount / tokensExcel.length;
        if (puntaje > maxPuntaje) {
            maxPuntaje = puntaje;
            mejorCandidato = item;
        }
    }

    if (mejorCandidato && maxPuntaje >= 0.35) {
        return {
            encontrado: true,
            metodo: "Coincidencia de palabras clave",
            maquinaNombre: mejorCandidato.nombreOriginal,
            datos: extraerCamposFila(mejorCandidato.row, mejorCandidato.nombreOriginal, jsonRows)
        };
    }

    return {
        encontrado: false,
        mensaje: "Ingresa el nombre/código de la máquina en el Paso 3 y presiona ENTER."
    };
}

/**
 * Consulta manual en Paso 3
 */
function consultarDatosMaquina(identificadorMaquina, excelFuente = EXCEL_LOCAL_PATH) {
    let jsonRows = Array.isArray(excelFuente) ? excelFuente : obtenerDataframeExcel(excelFuente);
    const idLimpio = normalizarTexto(identificadorMaquina);

    let filaEncontrada = null;
    let nombreEnc = identificadorMaquina;

    for (const row of jsonRows) {
        if (!row || !row[1]) continue;
        const nombreRow = normalizarTexto(row[1]);
        if (nombreRow.includes(idLimpio) || idLimpio.includes(nombreRow)) {
            filaEncontrada = row;
            nombreEnc = row[1];
            break;
        }
    }

    if (!filaEncontrada) {
        throw new Error(`No se encontró la máquina '${identificadorMaquina}' en la matriz Excel.`);
    }

    return extraerCamposFila(filaEncontrada, nombreEnc, jsonRows);
}

/**
 * Reemplaza etiquetas en el archivo Word final
 */
async function actualizarPlantillaWord(docxBuffer, datos) {
    const listaTecnicos = [datos.tecnico1, datos.tecnico2, datos.tecnico3, datos.tecnico4]
        .filter(t => t && t.trim() !== '')
        .join(' / ');

    const result = await DocxTemplater.createReport({
        template: docxBuffer,
        data: datos,
        cmdDelimiter: ['{{', '}}'],
        replaceVars: {
            "ORDEN DE TRABAJO": `ORDEN DE TRABAJO: ${datos.orden_trabajo}`,
            "TECNICO RESPONSABLE": `TECNICO RESPONSABLE\t${listaTecnicos || 'MIGUEL GUERRERO'}`,
            "FECHA DE EJECUCION": `FECHA DE EJECUCION\t   ${datos.fecha}`,
            "HORAS UTILIZADAS": `HORAS UTILIZADAS\t  ${datos.horas}`
        }
    });

    return Buffer.from(result);
}

async function procesarFicha(docxBufferInput, datosPersonalizados) {
    if (!fs.existsSync(CARPETA_SALIDA)) {
        fs.mkdirSync(CARPETA_SALIDA, { recursive: true });
    }

    const nuevoBufferWord = await actualizarPlantillaWord(docxBufferInput, datosPersonalizados);
    const nombreSalida = `${datosPersonalizados.orden_trabajo} (${datosPersonalizados.codigo || 'Ficha'})_ACTUALIZADO`;
    const rutaDocx = path.join(CARPETA_SALIDA, `${nombreSalida}.docx`);

    fs.writeFileSync(rutaDocx, nuevoBufferWord);

    return {
        nombreSalida,
        rutaDocx,
        bufferWord: nuevoBufferWord,
        datos: datosPersonalizados
    };
}

module.exports = {
    reconociendoMaquinaInteligente,
    consultarDatosMaquina,
    procesarFicha,
    CARPETA_SALIDA
};
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const PizZip = require('pizzip');

const EXCEL_LOCAL_PATH = path.join(__dirname, 'docex', 'Prueba_Plan_Mantenimiento_2025.xlsx');
const CARPETA_SALIDA = path.join(__dirname, 'Fichas_Generadas');

function normalizarTexto(str) {
    return String(str || '')
        .toUpperCase()
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
        .replace(/_/g, ' ')
        .replace(/[^\w\s\(\)]/gi, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}

function obtenerDataframeExcel(localPath = EXCEL_LOCAL_PATH) {
    if (!fs.existsSync(localPath)) {
        throw new Error(`No se encontró el archivo Excel en la ruta: ${localPath}`);
    }

    const workbook = XLSX.readFile(localPath, { cellDates: true });
    if (!workbook.SheetNames.includes('PLAN.MANT')) {
        throw new Error("La hoja 'PLAN.MANT' no existe en el archivo Excel.");
    }

    const sheet = workbook.Sheets['PLAN.MANT'];
    return XLSX.utils.sheet_to_json(sheet, { header: 1 });
}

/**
 * Extrae datos de la fila de Excel con búsqueda mejorada de Horas y Fecha Término
 */
function extraerCamposFila(fila, nombreMaquina, jsonRows) {
    let idxObs = -1;
    let idxHoras = -1;
    let idxFechaTermino = -1;
    let idxTecnico1 = -1, idxTecnico2 = -1, idxTecnico3 = -1, idxTecnico4 = -1;

    // 1. Escaneo de Encabezados en las primeras filas
    for (let r = 0; r < Math.min(12, jsonRows.length); r++) {
        const rowHeader = jsonRows[r];
        if (!rowHeader) continue;

        rowHeader.forEach((cell, idx) => {
            const cellNorm = normalizarTexto(cell);
            
            if (cellNorm.includes("TECNICO 1")) idxTecnico1 = idx;
            else if (cellNorm.includes("TECNICO 2")) idxTecnico2 = idx;
            else if (cellNorm.includes("TECNICO 3")) idxTecnico3 = idx;
            else if (cellNorm.includes("TECNICO 4")) idxTecnico4 = idx;
            else if (cellNorm.includes("OBSERVACION")) idxObs = idx;
            else if (cellNorm.includes("HORAS") || cellNorm.includes("H.H") || cellNorm.includes("REALES")) {
                if (idxHoras === -1) idxHoras = idx;
            }
            else if (cellNorm.includes("TERMINO") || cellNorm.includes("FECHA TERMINO")) {
                idxFechaTermino = idx;
            }
        });
    }

    // Fallbacks
    if (idxFechaTermino === -1) idxFechaTermino = 17; // Columna R por defecto

    // 2. Orden de Trabajo (Busca OSS + Números)
    let ordenTrabajo = "S/N";
    for (let cellVal of fila) {
        if (cellVal && typeof cellVal === 'string') {
            const otMatch = cellVal.match(/OSS\d+/i);
            if (otMatch) {
                ordenTrabajo = otMatch[0].toUpperCase();
                break;
            }
        }
    }

    // 3. Extracción Garantizada de Horas Reales
    let horasVal = (idxHoras !== -1 && fila[idxHoras] !== undefined) ? fila[idxHoras] : null;
    
    // Si no se ubicó por encabezado, buscar el valor numérico/decimal de horas en la fila
    if (horasVal === null || horasVal === undefined || String(horasVal).toUpperCase().includes("OSS")) {
        for (let cellVal of fila) {
            if (typeof cellVal === 'number' && cellVal > 0 && cellVal < 1000) {
                horasVal = cellVal;
                break;
            }
        }
    }

    let horasLimpio = (horasVal !== null && horasVal !== undefined) ? String(horasVal).trim() : "0";
    if (horasLimpio.toUpperCase().includes("OSS") || horasLimpio.toUpperCase().includes("REALIZADO")) {
        horasLimpio = "0";
    }
    const horasStr = `${horasLimpio} HORAS`;

    // 4. Fecha de Ejecución (Basada en Fecha Término - Columna R)
    let fechaStr = "AGOSTO - 2025";
    const fechaVal = fila[idxFechaTermino];

    if (fechaVal instanceof Date) {
        const meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", 
                       "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];
        fechaStr = `${meses[fechaVal.getMonth()]} - ${fechaVal.getFullYear()}`;
    } else if (typeof fechaVal === 'string' && fechaVal.trim() !== '') {
        fechaStr = fechaVal.trim().toUpperCase();
    }

    // 5. Técnicos
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

    if (!t1 && !t2) {
        const posibles = [];
        fila.forEach(cell => {
            if (cell && typeof cell === 'string') {
                const c = cell.trim();
                if (c.length > 5 && c.includes(' ') && !c.includes('OSS') && !c.includes('Realizado') && !c.includes('MONTACARGA')) {
                    posibles.push(c);
                }
            }
        });
        t1 = posibles[0] || "";
        t2 = posibles[1] || "";
        t3 = posibles[2] || "";
        t4 = posibles[3] || "";
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

    // CAPA 1: Código en paréntesis
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

    // CAPA 2: Subcadena
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
 * Reemplaza etiquetas tanto en texto inline como dentro de cuadros/celdas
 */
/**
 * Reemplaza etiquetas tanto en texto inline como dentro de cuadros/celdas
 */
async function actualizarPlantillaWord(docxBuffer, datos) {
    const listaTecnicosArr = [datos.tecnico1, datos.tecnico2, datos.tecnico3, datos.tecnico4]
        .filter(t => t && String(t).trim() !== '');

    // Formato con saltos de línea XML (<w:br/>) para cuadros
    const tecnicosXml = listaTecnicosArr.length > 0
        ? listaTecnicosArr.join('</w:t><w:br/><w:t>')
        : 'MIGUEL GUERRERO';

    try {
        const zip = new PizZip(docxBuffer);
        let documentXml = zip.file("word/document.xml").asText();

        // 1. REEMPLAZO DE ORDEN DE TRABAJO
        // SOLUCIÓN: Terminar con "\d" obliga a que se detenga en el último número de la OT.
        // Así NO se comerá la etiqueta de Word que da el salto de línea hacia "AREA: FIELATURA".
        documentXml = documentXml.replace(/OSS(?:<[^>]+>|\d)*\d/gi, datos.orden_trabajo);

        // 2. REEMPLAZO DE TÉCNICOS (TU CÓDIGO ORIGINAL SIN CAMBIOS)
        documentXml = documentXml.replace(/MIGUEL GUERRERO[\s\S]*?JORGE LOZANO/gi, tecnicosXml);
        documentXml = documentXml.replace(/(TECNICO RESPONSABLE\s*[\t\s]*)([^<\r\n]+)/gi, `$1${listaTecnicosArr.join(' / ')}`);

        // 3. REEMPLAZO DE FECHA DE EJECUCIÓN
        // SOLUCIÓN: Igual que arriba, terminamos con "\d" para que se detenga en el año exacto.
        const mesesRegex = "ENERO|FEBRERO|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|SETIEMBRE|SEPTIEMBRE|OCTUBRE|NOVIEMBRE|DICIEMBRE";
        const regexFechaXml = new RegExp(`(?:${mesesRegex})(?:<[^>]+>|[\\s\\t\\r\\n\\-\\u00a0]|&#160;)*\\d(?:<[^>]+>|\\d)*\\d`, 'gi');
        documentXml = documentXml.replace(regexFechaXml, datos.fecha);

        // 4. REEMPLAZO DE HORAS REALES (TU CÓDIGO ORIGINAL SIN CAMBIOS)
        documentXml = documentXml.replace(/(\d+(\.\d+)?)\s*HORAS/gi, datos.horas);

        zip.file("word/document.xml", documentXml);

        return zip.generate({
            type: 'nodebuffer',
            compression: 'DEFLATE'
        });

    } catch (error) {
        console.error(" Error al actualizar el documento Word:", error);
        throw new Error("No se pudo reemplazar el contenido en el documento Word.");
    }
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
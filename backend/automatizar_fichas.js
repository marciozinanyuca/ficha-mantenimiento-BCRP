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
 * Extrae datos de la fila de Excel respetando celdas vacías y detectando falta de información
 */
/**
 * Extrae datos de la fila de Excel de forma estricta (sin inventar datos si está vacío)
 */
function extraerCamposFila(fila, nombreMaquina, jsonRows) {
    let idxObs = -1;
    let idxHoras = -1;
    let idxFechaTermino = -1;
    let idxTecnico1 = -1, idxTecnico2 = -1, idxTecnico3 = -1, idxTecnico4 = -1;

    // 1. Escaneo de Encabezados
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

    if (idxFechaTermino === -1) idxFechaTermino = 17;

    // 2. Orden de Trabajo
    let ordenTrabajo = "";
    for (let cellVal of fila) {
        if (cellVal && typeof cellVal === 'string') {
            const otMatch = cellVal.match(/OSS\d+/i);
            if (otMatch) {
                ordenTrabajo = otMatch[0].toUpperCase();
                break;
            }
        }
    }

    // 3. Horas Reales (EXTRACCIÓN ESTRICTA)
    let horasStr = "";
    if (idxHoras !== -1 && fila[idxHoras] !== undefined && fila[idxHoras] !== null) {
        let horasLimpio = String(fila[idxHoras]).trim();
        // Validamos que no sea un texto equivocado
        if (horasLimpio !== "" && !horasLimpio.toUpperCase().includes("OSS") && !horasLimpio.toUpperCase().includes("REALIZADO")) {
            horasStr = `${horasLimpio} HORAS`;
        }
    }

    // 4. Fecha de Ejecución
    let fechaStr = "";
    const fechaVal = fila[idxFechaTermino];
    if (fechaVal instanceof Date) {
        const meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];
        fechaStr = `${meses[fechaVal.getMonth()]} - ${fechaVal.getFullYear()}`;
    } else if (typeof fechaVal === 'string' && fechaVal.trim() !== '') {
        fechaStr = fechaVal.trim().toUpperCase();
    }

    // 5. Técnicos (ESTRICTO)
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

    // 6. Validación de Fila Vacía
    const datosEstanVacios = !ordenTrabajo && !horasStr && !t1;

    return {
        codigo: nombreMaquina,
        orden_trabajo: ordenTrabajo,
        tecnico1: t1,
        tecnico2: t2,
        tecnico3: t3,
        tecnico4: t4,
        fecha: fechaStr,
        horas: horasStr,
        alerta_vacio: datosEstanVacios // Enviamos la bandera al frontend
    };
}

function reconociendoMaquinaInteligente(textoWord, nombreArchivo, excelFuente = EXCEL_LOCAL_PATH) {
    let jsonRows = Array.isArray(excelFuente) ? excelFuente : obtenerDataframeExcel(excelFuente);
    const fuenteTextoCompleta = normalizarTexto(`${textoWord} ${nombreArchivo}`);

    // --- 1. NUEVO: ESCANEAR Y EXTRAER DIAGNÓSTICO DEL WORD ---
    let diagnosticoDetectado = "";
    // Busca todo lo que esté entre "DIAGNOSTICO DE FALLA" y "OPERACION" (con o sin tilde)
    const regexDiagnostico = /DIAGNOSTICO\s+DE\s+FALLA\s*([\s\S]*?)(?=OPERACI[OÓ]N)/i;
    const matchDiagnostico = textoWord.match(regexDiagnostico);
    
    if (matchDiagnostico && matchDiagnostico[1]) {
        // Reemplaza múltiples espacios y saltos de línea por un solo espacio limpio
        diagnosticoDetectado = matchDiagnostico[1].replace(/\s+/g, ' ').trim();
    }

    // --- 2. NUEVO: FUNCIÓN AUXILIAR PARA EMPAQUETAR EL DIAGNÓSTICO ---
    const inyectarDiagnostico = (datosExtraidos) => {
        datosExtraidos.diagnostico = diagnosticoDetectado;
        return datosExtraidos;
    };

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
                    // --- 3. NUEVO: APLICAMOS LA INYECCIÓN EN LOS DATOS ---
                    datos: inyectarDiagnostico(extraerCamposFila(match.row, match.nombreOriginal, jsonRows))
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
                // --- 3. NUEVO: APLICAMOS LA INYECCIÓN EN LOS DATOS ---
                datos: inyectarDiagnostico(extraerCamposFila(item.row, item.nombreOriginal, jsonRows))
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
            // --- 3. NUEVO: APLICAMOS LA INYECCIÓN EN LOS DATOS ---
            datos: inyectarDiagnostico(extraerCamposFila(mejorCandidato.row, mejorCandidato.nombreOriginal, jsonRows))
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
 * Reemplazo de Zona Aislada: Aniquila números fragmentados (ej. "20" y "24") 
 * delimitando el área exacta de la fecha para evitar cruces.
 */
/**
 * Reemplazo Quirúrgico: Si el formulario está vacío, borra la información 
 * del documento original dejándolo en blanco.
 */
/**
 * Reemplazo Quirúrgico con Mayúsculas Automáticas: Todo texto ingresado 
 * se convertirá a mayúsculas antes de insertarse en el documento.
 */
async function actualizarPlantillaWord(docxBuffer, datos) {
    try {
        const zip = new PizZip(docxBuffer);
        let documentXml = zip.file("word/document.xml").asText();

        const conservarEtiquetas = (str) => (str || "").replace(/(^|>)[^<]+(<|$)/g, '$1$2');

        // 0. REEMPLAZO DE CÓDIGO O NOMBRE DE MÁQUINA (Seguro y delimitado)
        if (datos.codigo && String(datos.codigo).trim() !== "") {
            const codigoNuevo = String(datos.codigo).trim().toUpperCase();
            const regexKeywordMaq = /CODIGO(?:<[^>]+>|[\s\t\r\n])*DE(?:<[^>]+>|[\s\t\r\n])*LA(?:<[^>]+>|[\s\t\r\n])*MAQUINA(?:<[^>]+>|[\s\t\r\n])*O(?:<[^>]+>|[\s\t\r\n])*EQUIPO/i;

            const keywordMaqMatch = documentXml.match(regexKeywordMaq);
            if (keywordMaqMatch) {
                const partes = documentXml.split(keywordMaqMatch[0]);
                if (partes.length > 1) {
                    let subXml = partes[1];
                    let primerNodoEncontrado = false;
                    let contadorNodosModificados = 0;

                    // SOLO modificamos dentro de un radio seguro (primeros 400 caracteres de código)
                    subXml = subXml.replace(/(<w:t[^>]*>)([^<]*?)(<\/w:t>)/gi, (match, openTag, contenido, closeTag, offset) => {
                        if (offset < 400) {
                            if (contenido.trim() !== "" && !/^[:\s]*$/.test(contenido)) {
                                if (!primerNodoEncontrado) {
                                    primerNodoEncontrado = true;
                                    contadorNodosModificados++;
                                    return `${openTag}${codigoNuevo}${closeTag}`; // Inserta el nuevo nombre
                                } else if (contadorNodosModificados < 3) {
                                    // Limpia solo si el nombre viejo estaba partido en nodos vecinos cercanos
                                    contadorNodosModificados++;
                                    return `${openTag}${closeTag}`;
                                }
                            }
                        }
                        return match;
                    });

                    documentXml = partes[0] + keywordMaqMatch[0] + subXml;
                }
            }
        }




        // 1. REEMPLAZO DE ORDEN DE TRABAJO (Convertido a Mayúsculas)
        const otNueva = datos.orden_trabajo && String(datos.orden_trabajo).trim() !== ""
            ? String(datos.orden_trabajo).trim().toUpperCase()
            : " ";
        documentXml = documentXml.replace(/O(?:<[^>]+>)*S(?:<[^>]+>)*S(?:<[^>]+>|\d)*\d/gi, (match) => {
            return otNueva + conservarEtiquetas(match);
        });





// 2. REEMPLAZO DE TÉCNICOS (Método de Ancla con Expansión Automática de Cuadro)
        const listaTecnicosArr = [datos.tecnico1, datos.tecnico2, datos.tecnico3, datos.tecnico4]
            .filter(t => t && String(t).trim() !== '')
            .map(t => String(t).trim().toUpperCase());

        const nombresViejos = /DANIEL CELIS|MIGUEL GUERRERO|JORGE LOZANO|RAUL CONSTANTINI|C[EÉ]SAR VELASQUEZ/gi;
        
        // PASO 1: Limpiamos los nombres viejos y dejamos el Ancla.
        let primerReemplazoHecho = false;
        documentXml = documentXml.replace(nombresViejos, (match) => {
            if (!primerReemplazoHecho) {
                primerReemplazoHecho = true;
                return "[[ANCLA_TEC]]" + conservarEtiquetas(match); 
            } else {
                return conservarEtiquetas(match); 
            }
        });

        // PASO 2: Operamos desde el Ancla
        if (documentXml.includes("[[ANCLA_TEC]]")) {
            const partesAncla = documentXml.split("[[ANCLA_TEC]]");
            const tec1 = listaTecnicosArr.length > 0 ? listaTecnicosArr[0] : " ";
            
            // --- NUEVO PASO ESTRELLA: HACER QUE EL CUADRO SE ESTIRE HACIA ABAJO ---
            // Buscamos hacia atrás dónde empieza la fila de la tabla (<w:tr)
            const indexUltimaFila = partesAncla[0].lastIndexOf('<w:tr'); 
            if (indexUltimaFila !== -1) {
                let fragmentoFila = partesAncla[0].substring(indexUltimaFila);
                // Cambiamos la altura de "exacta" a "automática" solo para esta fila
                fragmentoFila = fragmentoFila.replace(/w:hRule="exact"/gi, 'w:hRule="auto"');
                partesAncla[0] = partesAncla[0].substring(0, indexUltimaFila) + fragmentoFila;
            }
            // ----------------------------------------------------------------------

            // Extraer el formato original (tamaño, tipo de letra)
            const matchRPr = partesAncla[0].match(/<w:rPr>[\s\S]*?<\/w:rPr>/ig);
            const rPrText = matchRPr ? matchRPr[matchRPr.length - 1] : "";

            let xmlPosterior = partesAncla[1];

            // Si hay más de un técnico, creamos los ENTERS reales
            if (listaTecnicosArr.length > 1) {
                let parrafosExtra = "";
                for (let i = 1; i < listaTecnicosArr.length; i++) {
                    parrafosExtra += `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>${rPrText}<w:t>${listaTecnicosArr[i]}</w:t></w:r></w:p>`;
                }
                
                // Buscamos el final de la línea actual para inyectarlos debajo
                const indexFinParrafo = xmlPosterior.indexOf('</w:p>');
                if (indexFinParrafo !== -1) {
                    const posCorte = indexFinParrafo + '</w:p>'.length;
                    xmlPosterior = xmlPosterior.substring(0, posCorte) + parrafosExtra + xmlPosterior.substring(posCorte);
                }
            }

            // PASO 3: Unimos todo con el primer técnico en su lugar
            documentXml = partesAncla[0] + tec1 + xmlPosterior;
        }
        

        // 3. REEMPLAZO DE FECHA DE EJECUCIÓN (Convertido a Mayúsculas)
        const fechaNueva = datos.fecha && String(datos.fecha).trim() !== ""
            ? String(datos.fecha).trim().toUpperCase()
            : " ";
        const regexBloque = /(FECHA(?:<[^>]+>|[\s\t\r\n])*DE(?:<[^>]+>|[\s\t\r\n])*EJECUCION[\s\S]*?)(HORAS|<\/w:tr>|<\/w:txbxContent>)/i;

        documentXml = documentXml.replace(regexBloque, (match, bloqueFecha, limiteFinal) => {
            let fechaInsertada = false;

            const nuevoBloque = bloqueFecha.replace(/(<w:t[^>]*>)([^<]*?)(<\/w:t>)/gi, (m, openTag, texto, closeTag) => {
                if (texto.trim() === "") return m;
                const txtUpper = texto.toUpperCase();

                if (/FECHA|EJECUCION/i.test(txtUpper)) return m;

                const tieneMes = /ENERO|FEBRERO|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|SETIEMBRE|SEPTIEMBRE|OCTUBRE|NOVIEMBRE|DICIEMBRE/i.test(txtUpper);
                const tieneNumeroOGuion = /[0-9\-]/.test(txtUpper);

                if (tieneMes || tieneNumeroOGuion) {
                    if (!fechaInsertada) {
                        fechaInsertada = true;
                        return `${openTag}${fechaNueva}${closeTag}`;
                    } else {
                        return `${openTag}${closeTag}`;
                    }
                }
                return m;
            });
            return nuevoBloque + limiteFinal;
        });

        // 4. REEMPLAZO DE HORAS REALES (Por si se escriben letras)
        let horasNueva = " ";
        if (datos.horas && String(datos.horas).trim() !== "") {
            const horasStr = String(datos.horas).trim().toUpperCase();
            const matchHoras = horasStr.match(/\d+(?:\.\d+)?/);
            horasNueva = matchHoras ? matchHoras[0] : horasStr;
        }

        const regexKeyword = /HORAS(?:<[^>]+>|[\s\t\r\n])*(?:REALES(?:<[^>]+>|[\s\t\r\n])*)?UTILIZADA(?:S)?/i;
        const keywordMatch = documentXml.match(regexKeyword);

        if (keywordMatch) {
            const partes = documentXml.split(keywordMatch[0]);
            if (partes.length > 1) {
                partes[1] = partes[1].replace(/>([\s\t\r\n]*)\d+(?:\.\d+)?/i, `>$1${horasNueva}`);
                documentXml = partes.join(keywordMatch[0]);
            }
        } else {
            documentXml = documentXml.replace(/(\d+(\.\d+)?)\s*HORAS/gi, (match) => {
                return horasNueva + " " + conservarEtiquetas(match).replace(/\d+(\.\d+)?\s*/g, '');
            });
        }

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
    CARPETA_SALIDA,
    obtenerListaMaquinas
};

/**
 * Extrae todos los nombres de las máquinas de la columna "Máquina" en el Excel,
 * omitiendo de forma inteligente los encabezados y los TÍTULOS de secciones.
 */
function obtenerListaMaquinas(jsonRows) {
    let idxMaquina = -1;
    let filaInicio = 0;
    let maquinas = [];

    // 1. Buscar en qué columna y en qué fila está el encabezado "MÁQUINA"
    for (let r = 0; r < Math.min(12, jsonRows.length); r++) {
        const rowHeader = jsonRows[r];
        if (!rowHeader) continue;

        rowHeader.forEach((cell, idx) => {
            if (typeof cell === 'string' && /M[AÁ]QUINA/i.test(cell)) {
                idxMaquina = idx;
                filaInicio = r + 1; // Empezamos a leer desde la fila de abajo
            }
        });
        if (idxMaquina !== -1) break;
    }

    if (idxMaquina === -1) idxMaquina = 1;

    // 2. Extraer los nombres reales de las máquinas
    for (let r = filaInicio; r < jsonRows.length; r++) {
        const celda = jsonRows[r][idxMaquina];
        if (celda && typeof celda === 'string') {
            const nombreLimpio = celda.trim();
            const nombreUpper = nombreLimpio.toUpperCase();

            // Filtros de limpieza
            const esVacio = nombreLimpio === "";
            const esEncabezado = nombreUpper === "MÁQUINA" || nombreUpper === "MAQUINA";

            // FILTRO NUEVO: Bloquea los títulos de agrupaciones del Excel
            const esTituloSeccion = nombreUpper.includes("MANTENIMIENTO") ||
                nombreUpper.includes("ELABORADO") ||
                nombreUpper.includes("FECHA") ||
                nombreUpper.includes("ÁREA") ||
                nombreUpper.includes("AREA");


            // Solo si pasa todas las pruebas de seguridad, se agrega a la lista
            if (!esVacio && !esEncabezado && !esTituloSeccion) {
                maquinas.push(nombreLimpio);
            }
        }
    }

    // Retornamos la lista eliminando cualquier máquina duplicada
    return [...new Set(maquinas)];
}
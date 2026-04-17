/**
 * Función de limpieza única: Elimina la protección global de hoja 
 * con el nombre específico en todos los cursos excepto 1A y 1B.
 */
function limpiezaProteccionGlobalUnica() {
  const ss = SpreadsheetApp.openById(CFG.SS_DATA);
  
  // Eliminamos la referencia a la celda-consola y al panel
  Logger.log("🧹 Iniciando limpieza de 'Protección de Libro de Temas'...");

  const mXC = obtenerObjetoMaterias(ss.getSheetByName("masdatos"), "A2:B");
  const folder = DriveApp.getFolderById(CFG.FLD);
  const cursosExcluidos = [];

  for (let curso in mXC) {
    // Saltamos cursos excluidos si es necesario
    if (cursosExcluidos.includes(curso.toUpperCase())) continue;

    const files = folder.getFilesByName(getNm(curso));
    if (!files.hasNext()) continue;
    
    try {
      const ssCurso = SpreadsheetApp.openById(files.next().getId());
      Logger.log(`🔎 Revisando archivo: ${curso}`);

      const hojas = ssCurso.getSheets();
      hojas.forEach(hoja => {
        // Buscamos protecciones de tipo HOJA (Sheet)
        const proteccionesHoja = hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET);
        
        proteccionesHoja.forEach(p => {
          // Si el nombre coincide con el que causaba problemas, la eliminamos
          if (p.getDescription() === "Protección de Libro de Temas") {
            p.remove();
            Logger.log(`   ✨ Eliminada en materia: ${hoja.getName()}`);
          }
        });
      });
    } catch (e) {
      Logger.log(`❌ Error en curso ${curso}: ${e.message}`);
    }
  }
  Logger.log("🏁 Limpieza finalizada. Ya puedes usar 'actualizarPermisosSemanales'.");
}


/**
 * Revisa la hoja 'dias' y vuelca observaciones (columna F) 
 * en los Libros de Temas si las celdas de contenido están vacías.
 */
function sincronizarObservacionesFeriados() {
  const ssBase = SpreadsheetApp.openById(CFG.SS_DATA);
  const hojaDias = ssBase.getSheetByName("dias");
  
  if (!hojaDias) {
    console.error("❌ Error: No se encontró la hoja 'dias'");
    return;
  }

  // 1. Cargar motivos desde la columna F (Índice 1 del rango E:F)
  // Rango: Fila 2, Columna 5 (E), por 213 filas y 2 columnas (E y F)
  const datosDias = hojaDias.getRange(2, 5, 213, 3).getValues(); 
  const dictMotivos = {};
  
  datosDias.forEach(fila => {
    let fecha = fila[0]; // Columna E
    let motivo = fila[2]; // Columna F
    if (fecha instanceof Date && motivo && motivo.toString().trim() !== "") {
      fecha.setHours(0,0,0,0);
      dictMotivos[fecha.getTime()] = motivo.toString().trim();
    }
  });

  const mXC = obtenerObjetoMaterias(ssBase.getSheetByName("masdatos"), "A2:B");
  const folder = DriveApp.getFolderById(CFG.FLD);

  for (let curso in mXC) {
    const files = folder.getFilesByName(getNm(curso));
    if (!files.hasNext()) continue;
    
    try {
      const ssCurso = SpreadsheetApp.openById(files.next().getId());
      console.log(`📂 Sincronizando motivos: ${curso}`);

      mXC[curso].forEach(nombreM => {
        const hoja = ssCurso.getSheetByName(nombreM);
        if (!hoja) return;

        const rango = hoja.getRange("A3:E214");
        const valores = rango.getValues();
        let huboCambio = false;

        for (let i = 0; i < valores.length; i++) {
          let fechaCelda = valores[i][0];
          
          if (fechaCelda instanceof Date) {
            fechaCelda.setHours(0,0,0,0);
            let motivoDb = dictMotivos[fechaCelda.getTime()];

            if (motivoDb) {
              // Verificamos si la fila está "disponible" para poner el motivo (B, C, D vacíos)
              let colB = valores[i][1].toString().trim();
              let colC = valores[i][2].toString().trim();
              let colD = valores[i][3].toString().trim();
              let colE = valores[i][4].toString().trim();

              // Solo escribimos si no hay datos del docente y el motivo es diferente al actual
              if (colB === "" && colC === "" && colD === "" && colE !== motivoDb) {
                valores[i][4] = motivoDb; 
                huboCambio = true;
              }
            }
          }
        }

        if (huboCambio) {
          rango.setValues(valores);
        }
      });
    } catch (e) {
      console.error(`❌ Error en curso ${curso}: ${e.message}`);
    }
  }
  console.log("🏁 Sincronización de feriados/paros completada.");
}


/**
 * Escanea todos los cursos y materias en busca de fechas duplicadas.
 * No modifica datos, solo informa por la consola del Panel de Control.
 */
function reportarFechasDuplicadas() {
  const ss = SpreadsheetApp.openById(CFG.SS_DATA);
  const panel = ss.getSheetByName("Panel de Control");
  const hojaMasDatos = ss.getSheetByName("masdatos");
  
  if (!panel || !hojaMasDatos) return console.error("Faltan hojas críticas.");

  const celdaConsola = panel.getRange("B6");
  celdaConsola.clearContent();
  
  const log = (msg) => {
    const valorActual = celdaConsola.getValue();
    celdaConsola.setValue(valorActual ? valorActual + "\n" + msg : msg);
    SpreadsheetApp.flush();
  };

  log("🔍 Iniciando escaneo de duplicados...");

  const mXC = obtenerObjetoMaterias(hojaMasDatos, "A2:B");
  const folder = DriveApp.getFolderById(CFG.FLD);
  let totalDuplicadosEncontrados = 0;

  for (let curso in mXC) {
    const files = folder.getFilesByName(getNm(curso));
    if (!files.hasNext()) continue;

    try {
      const ssCurso = SpreadsheetApp.openById(files.next().getId());
      log(`--- 📂 Curso: ${curso} ---`);

      mXC[curso].forEach(nombreM => {
        const hojaMat = ssCurso.getSheetByName(nombreM);
        if (!hojaMat) return;

        const fechas = hojaMat.getRange("A3:A214").getValues();
        const registroFechas = {}; // { 'YYYYMMDD': [filas] }
        let duplicadosEnEstaMateria = [];

        fechas.forEach((fila, index) => {
          let f = fila[0];
          if (f instanceof Date) {
            // Normalizar fecha para comparar
            const k = Utilities.formatDate(f, "GMT", "yyyyMMdd");
            const numFila = index + 3;

            if (!registroFechas[k]) {
              registroFechas[k] = [numFila];
            } else {
              registroFechas[k].push(numFila);
              duplicadosEnEstaMateria.push({
                fecha: Utilities.formatDate(f, "GMT", "dd/MM/yyyy"),
                filas: registroFechas[k]
              });
            }
          }
        });

        if (duplicadosEnEstaMateria.length > 0) {
          log(`⚠️ ${nombreM}:`);
          duplicadosEnEstaMateria.forEach(dup => {
            log(`   • Fecha ${dup.fecha} repetida en filas: ${dup.filas.join(", ")}`);
            totalDuplicadosEncontrados++;
          });
        }
      });
    } catch (e) {
      log(`❌ Error en curso ${curso}: ${e.message}`);
    }
  }

  log("---------------------------------------");
  log(totalDuplicadosEncontrados > 0 
    ? `✅ Escaneo finalizado. Se encontraron ${totalDuplicadosEncontrados} duplicados.` 
    : "✅ Escaneo finalizado. No se encontraron duplicados.");
}



/**
 * Quita exclusivamente los bloqueos de "Mantenimiento_Pasado".
 * Uso esporádico para permitir ediciones históricas.
 */
function liberarBloqueosPasados() {
  const ss = SpreadsheetApp.openById(CFG.SS_DATA);
  const panel = ss.getSheetByName("Panel de Control");
  const celdaConsola = panel ? panel.getRange("B6") : null;

  const log = (msg) => {
    if (celdaConsola) {
      const v = celdaConsola.getValue();
      celdaConsola.setValue(v ? v + "\n" + msg : msg);
      SpreadsheetApp.flush();
    }
    console.log(msg);
  };

  if (celdaConsola) celdaConsola.clearContent();
  log("🔓 Iniciando liberación selectiva de protecciones...");

  const mXC = obtenerObjetoMaterias(ss.getSheetByName("masdatos"), "A2:B");
  const folder = DriveApp.getFolderById(CFG.FLD);

  for (let curso in mXC) {
    const files = folder.getFilesByName(getNm(curso));
    if (!files.hasNext()) continue;

    try {
      const ssCurso = SpreadsheetApp.openById(files.next().getId());
      log(`📂 Abriendo: ${curso}`);

      mXC[curso].forEach(nombreM => {
        const hoja = ssCurso.getSheetByName(nombreM);
        if (!hoja) return;

        const protecciones = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        let contadorHoja = 0;

        protecciones.forEach(p => {
          // Verificamos que la descripción coincida exactamente con la de tu motor
          if (p.getDescription() === "Mantenimiento_Pasado") {
            p.remove();
            contadorHoja++;
          }
        });

        if (contadorHoja > 0) {
          log(`   ✅ ${nombreM}: Se quitaron ${contadorHoja} bloqueos.`);
        }
      });
    } catch (e) {
      log(`❌ Error en curso ${curso}: ${e.message}`);
    }
  }
  log("🏁 Proceso de liberación finalizado.");
}



/**
 * Bloquea el pasado de forma eficiente, agrupando filas consecutivas que tienen datos.
 * Si hay una fila vacía en el medio, se salta, creando rangos de bloqueo separados.
 */




/**
 * Agrupa la información de Educación Física por Profesor.
 * Muestra cantidad de cursos por día y detalle de cursos/horarios.
 */
function listarEdFisicaPorProfesor() {
  const ssActual = SpreadsheetApp.getActiveSpreadsheet();
  let hojaDestino = ssActual.getSheetByName("edfisica");
  
  if (!hojaDestino) {
    hojaDestino = ssActual.insertSheet("edfisica");
  } else {
    hojaDestino.clear();
  }

  const ssData = SpreadsheetApp.openById(CFG.SS_DATA);
  const mXC = obtenerObjetoMaterias(ssData.getSheetByName("masdatos"), "A2:B");
  const ssHorario = SpreadsheetApp.openById(CFG.HORARIO);
  
  // Objeto para agrupar: { "Nombre Prof": { 1: {count: 0, detalle: []}, ... } }
  const reporte = {};
  const nombresDias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"];

  console.log("Analizando horarios por profesor...");

  for (let curso in mXC) {
    const hojaH = ssHorario.getSheetByName(curso);
    if (!hojaH) continue;

    // Buscamos específicamente en la zona de Ed. Física según tu lógica (Fila 22 y 23)
    const dataH = hojaH.getRange(22, 1, 2, 7).getValues();
    const profe = dataH[0][4] ? dataH[0][4].toString().trim() : "Sin Profesor";
    
    if (profe === "Sin Profesor" || profe === "") continue;

    if (!reporte[profe]) {
      reporte[profe] = { 1: [], 2: [], 3: [], 4: [], 5: [] };
    }

    // Revisamos columnas de Lunes (2) a Viernes (6)
    for (let col = 2; col <= 6; col++) {
      const diaNum = col - 1;
      const celdaHora = dataH[1][col].toString().trim(); // Celda donde suele ir el horario
      
      if (celdaHora !== "") {
        reporte[profe][diaNum].push(`${curso} (${celdaHora})`);
      }
    }
  }

  // Preparar matriz para volcar a la hoja
  const filasSalida = [["PROFESOR", "DÍA", "CANT. CURSOS", "DETALLE (CURSO Y HORARIO)"]];

  // Ordenar profesores alfabéticamente
  const profesoresOrdenados = Object.keys(reporte).sort();

  profesoresOrdenados.forEach(p => {
    for (let d = 1; d <= 5; d++) {
      const clases = reporte[p][d];
      if (clases.length > 0) {
        filasSalida.push([
          p, 
          nombresDias[d-1], 
          clases.length, 
          clases.join(" | ")
        ]);
      }
    }
  });

  // Escribir en la hoja
  if (filasSalida.length > 1) {
    hojaDestino.getRange(1, 1, filasSalida.length, 4).setValues(filasSalida);
    
    // Formato
    hojaDestino.getRange("A1:D1").setBackground("#4A86E8").setFontColor("white").setFontWeight("bold");
    hojaDestino.setFrozenRows(1);
    hojaDestino.autoResizeColumns(1, 4);
    // Combinar celdas del nombre del profesor para que sea más legible
    hojaDestino.getRange(2, 1, hojaDestino.getLastRow()-1, 1).setVerticalAlignment("middle");
    
    console.log("✅ Reporte por profesor generado en 'edfisica'.");
  }
}



/**
 * Ejecuta una sola vez para establecer el rango B3:B de todas las materias
 * como "Texto sin formato". Útil para evitar que fechas o números se autodefinan.
 */
function configurarFormatoTextoMasivo() {
  const ss = SpreadsheetApp.openById(CFG.SS_DATA);
  const mXC = obtenerObjetoMaterias(ss.getSheetByName("masdatos"), "A2:B");
  const folder = DriveApp.getFolderById(CFG.FLD);
  
  console.log("Iniciando cambio de formato a 'Texto sin formato'...");

  for (let curso in mXC) {
    const files = folder.getFilesByName(getNm(curso));
    if (!files.hasNext()) continue;

    try {
      const ssCurso = SpreadsheetApp.openById(files.next().getId());
      console.log(`Working on: ${curso}`);

      mXC[curso].forEach(nombreM => {
        const hoja = ssCurso.getSheetByName(nombreM);
        if (hoja) {
          // Aplicamos el formato "@" (Plain Text) al rango de la columna B
          hoja.getRange("B3:B214").setNumberFormat("@");
        }
      });
      
      // Forzar guardado de cambios antes de pasar al siguiente archivo
      SpreadsheetApp.flush(); 
      
    } catch (e) {
      console.error(`Error en curso ${curso}: ${e.message}`);
    }
  }
  
  console.log("✅ Proceso de formato finalizado.");
}
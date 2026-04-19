const scriptProperties = PropertiesService.getScriptProperties();

const CFG = {
  PRE: "L. Temas",
  ANIO: "26",
  FLD:  scriptProperties.getProperty('fld'),
  SS_DATA: scriptProperties.getProperty('ssdata'),
  HORARIO: scriptProperties.getProperty('horario'),
  PLANTILLA: scriptProperties.getProperty('plantilla')
};

const getNm = (c) => {
  const p = c.includes("ESA") ? CFG.PRE.replace(/\s/g, "") : CFG.PRE;
  return `${p} ${c} ${CFG.ANIO}`;
};
function buscarDiasMateriaantigua(materia, curso) {
  try {
    const ss = SpreadsheetApp.openById(CFG.HORARIO), h = ss.getSheetByName(curso);
    if (!h) return [];
    const mat = materia.toUpperCase().trim(), dias = new Set();
    if (mat === "ED. FÍSICA") {
      const data = h.getRange(23, 1, 1, 7).getValues()[0]; 
      for (let c = 2; c < 7; c++) {
        if (data[c].toString().trim() !== "") dias.add(c - 1);
      }
    } else {
      const data = h.getRange(1, 1, 20, 7).getValues(), filas = [3, 5, 8, 10, 13, 15, 18, 20], mB = mat.toLowerCase();
      for (let c = 2; c < 7; c++) {
        filas.forEach(f => { if (data[f - 1][c].toString().toLowerCase().trim() === mB) dias.add(c - 1); });
      }
    }
    return [...dias].sort();
  } catch(e) { return []; }
}
function obtenerDatosPorDias(diasSel) {
  const hoja = SpreadsheetApp.openById(CFG.SS_DATA).getSheetByName("dias");
  if (!hoja || !diasSel.length) return [];
  return hoja.getRange(2, 4, 213, 5).getValues()
    .filter(f => !isNaN(parseInt(f[0] - 1)) && diasSel.includes(parseInt(f[0] - 1)))
    .map(f => ({ fecha: f[1], esLaborable: f[3], motivo: f[4] }));
}
function procesarMateriaYCurso(curso, materia, hojaDirecta = null) {
  let h = hojaDirecta;
  if (!h) {
    const f = DriveApp.getFolderById(CFG.FLD).getFilesByName(getNm(curso));
    if (!f.hasNext()) return console.warn("No se halló el archivo: " + getNm(curso));
    const ss = SpreadsheetApp.open(f.next());
    h = ss.getSheetByName(materia);
  }
  if (!h) return;
  const d = obtenerDatosPorDias(buscarDiasMateria(materia, curso));
  if (!d.length) return console.warn(`   - ${materia}: Sin fechas para volcar.`);
  h.getRange("A3:A214").clearContent();
  h.getRange("E3:E214").clearContent();
  h.getRange(3, 1, d.length, 1).setValues(d.map(x => [x.fecha]));
  h.getRange(3, 5, d.length, 1).setValues(d.map(x => [x.motivo]));
  console.log(`   - ${materia}: ${d.length} fechas escritas.`);
}
function procesarTodosLosArchivosDeCursos() {
  console.log("Iniciando actualización masiva...");
  const files = DriveApp.getFolderById(CFG.FLD).getFiles();
  const preEsc = CFG.PRE.replace(/\./g, "\\.").replace(/\s/g, "\\s?");
  const reg = new RegExp(`^${preEsc}\\s+(.+)\\s+${CFG.ANIO}$`);
  let count = 0;
  while (files.hasNext()) {
    let f = files.next();
    let n = f.getName();
    let m = n.match(reg);
    if (m) {
      let curso = m[1].trim();
      console.log(`>>> ARCHIVO: ${n} (Curso: ${curso})`);
      const ss = SpreadsheetApp.openById(f.getId());
      const hojas = ss.getSheets();
      hojas.forEach(hoja => {
        let nombreH = hoja.getName();
        if (!["copiar", "Config"].includes(nombreH)) {
          procesarMateriaYCurso(curso, nombreH, hoja);
        }
      });
      count++;
      SpreadsheetApp.flush(); 
    }
  }
  console.log(`>>> FINALIZADO: ${count} archivos procesados.`);
}
function generarCopiasCursosYMaterias() {
  console.log("Generando archivos nuevos...");
  const ss = SpreadsheetApp.openById(CFG.SS_DATA), h = ss.getSheetByName("masdatos");
  const dM = h.getRange("D2:E").getValues().reduce((a, [k, v]) => (k ? (a[k] = v, a) : a), {});
  const dT = h.getRange("G2:H").getValues().reduce((a, [k, v]) => (k ? (a[k] = v, a) : a), {});
  const mXC = obtenerObjetoMaterias(h, "A2:B");
  const folder = DriveApp.getFolderById(CFG.FLD), plant = DriveApp.getFileById(CFG.PLANTILLA);
  for (let c in mXC) {
    let nNom = getNm(c);
    console.log(`> Creando: ${nNom}`);
    let ssC = SpreadsheetApp.open(plant.makeCopy(nNom, folder));
    let hP = ssC.getSheetByName("copiar");
    mXC[c].forEach(m => {
      let mat = m.trim();
      let nH = hP.copyTo(ssC).setName(mat);
      nH.getRange("B1").setValue(`${dM[mat] || mat} ${c} T${dT[c] || ""}`);
    });
    ssC.deleteSheet(hP);
  }
  console.log("Copia finalizada.");
}
function obtenerObjetoMaterias(h, r) {
  const v = h.getRange(r).getValues();
  let m = {}, cA = null;
  v.forEach(f => {
    if (f[0]) { cA = f[0].toString().trim(); m[cA] = f[1] ? [f[1]] : []; }
    else if (cA && f[1]) m[cA].push(f[1]);
  });
  return m;
}
function actualizarDesdePanelControl() {
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
  log("⏳ Iniciando proceso ordenado...");
  let fechaInicio = panel.getRange("B3").getValue();
  if (!(fechaInicio instanceof Date) || isNaN(fechaInicio)) {
    fechaInicio = new Date();
    log("ℹ️ B3 vacío. Usando fecha actual.");
  }
  fechaInicio.setHours(0,0,0,0);
  const filtroC3 = panel.getRange("C3").getValue().toString().trim();
  const filtroD3 = panel.getRange("D3").getValue().toString().trim();
  const filtroCursos = filtroC3 ? filtroC3.split(",").map(s => s.trim().toUpperCase()) : [];
  const filtroMaterias = filtroD3 ? filtroD3.split(",").map(s => s.trim().toUpperCase()) : [];
  const mXC = obtenerObjetoMaterias(hojaMasDatos, "A2:B");
  const folder = DriveApp.getFolderById(CFG.FLD);
  for (let curso in mXC) {
    if (filtroCursos.length > 0 && !filtroCursos.includes(curso.toUpperCase())) continue;
    const nombreArchivo = getNm(curso);
    const files = folder.getFilesByName(nombreArchivo);
    if (files.hasNext()) {
      log(`📂 Procesando: ${curso}...`);
      const file = files.next();
      try {
        const ssCurso = SpreadsheetApp.openById(file.getId());
        mXC[curso].forEach(nombreM => {
          if (filtroMaterias.length > 0 && !filtroMaterias.includes(nombreM.toUpperCase())) return;
          const hojaMat = ssCurso.getSheetByName(nombreM);
          if (hojaMat) {
            aplicarReemplazoSelectivoConLogs(curso, nombreM, hojaMat, fechaInicio, log);
          }
        });
      } catch (e) {
        log(`❌ ERROR en ${curso}: ${e.message}`);
      }
    } else {
      log(`⚠️ Archivo no encontrado: ${nombreArchivo}`);
    }
  }
  log("✅ Proceso finalizado.");
}
function buscarDiasMateria(materia, curso) {
  try {
    const ss = SpreadsheetApp.openById(CFG.HORARIO), h = ss.getSheetByName(curso);
    if (!h) return { dias: [], profesor: "" };
    const mat = materia.toUpperCase().trim(), dias = new Set();
    let profesor = "";
    const data = h.getRange(1, 1, 25, 7).getValues(); 
    if (mat === "ED. FÍSICA") {
      profesor = data[21][4]; 
      const filaEF = 22; 
      for (let c = 2; c < 7; c++) {
        if (data[filaEF][c].toString().trim() !== "") {
          dias.add(c - 1);
        }
      }
    } else {
      const filas = [3, 5, 8, 10, 13, 15, 18, 20], mB = mat.toLowerCase();
      for (let c = 2; c < 7; c++) {
        filas.forEach(f => { 
          if (data[f - 1][c].toString().toLowerCase().trim() === mB) {
            dias.add(c - 1);
            if (!profesor) profesor = data[f][c]; 
          }
        });
      }
    }
    return { dias: [...dias].sort(), profesor: profesor.toString().trim() };
  } catch(e) { 
    return { dias: [], profesor: "" }; 
  }
}
function aplicarReemplazoSelectivoConLogs(curso, materia, hoja, fechaInicio, logFn) {
  const infoMateria = buscarDiasMateria(materia, curso);
  const todasLasFechas = obtenerDatosPorDias(infoMateria.dias);
  if (infoMateria.profesor) {
    hoja.getRange("E1").setValue(infoMateria.profesor);
  }
  const fechasNuevas = todasLasFechas.filter(d => {
    let fClase = new Date(d.fecha);
    fClase.setHours(0,0,0,0);
    return fClase >= fechaInicio;
  });
  if (fechasNuevas.length === 0) return;
  const rango = hoja.getRange("A3:E214");
  const valores = rango.getValues();
  let indexHoja = -1;
  for (let i = 0; i < valores.length; i++) {
    let fCelda = valores[i][0];
    if (fCelda instanceof Date) {
      fCelda.setHours(0,0,0,0);
      if (fCelda >= fechaInicio) { indexHoja = i; break; }
    } else if (fCelda === "") { indexHoja = i; break; }
  }
  if (indexHoja === -1) return;
  let fN = 0; 
  for (let i = indexHoja; i < valores.length && fN < fechasNuevas.length; i++) {
    const colB = valores[i][1].toString().trim();
    const colC = valores[i][2].toString().trim();
    const colD = valores[i][3].toString().trim();
    if (colB === "" && colC === "" && colD === "") {
      valores[i][0] = fechasNuevas[fN].fecha;
      valores[i][4] = fechasNuevas[fN].motivo;
      fN++; 
    } else {
      let fOcupada = fechasNuevas[fN].fecha;
      let fStr = (fOcupada instanceof Date) ? fOcupada.toLocaleDateString() : fOcupada;
      logFn(`⏭️ Salto: ${curso} > ${materia} | Fila ${i + 3} ocupada (${fStr})`);
    }
  }
  rango.setValues(valores);
}
function actualizarPermisosDiarios() {
  const FECHA_TEST_STR = null; 
  const EMAIL_ALERTA = "luisoroverde@gmail.com";
  const ss = SpreadsheetApp.openById(CFG.SS_DATA);
  const panel = ss.getSheetByName("Panel de Control");
  const celdaConsola = panel ? panel.getRange("B6") : null;
  let hoy;
  if (FECHA_TEST_STR) {
    const partes = FECHA_TEST_STR.split("/");
    hoy = new Date(partes[2], partes[1] - 1, partes[0]);
  } else {
    hoy = new Date();
  }
  hoy.setHours(0, 0, 0, 0);
  const log = (msg) => {
    if (celdaConsola) {
      const valorActual = celdaConsola.getValue();
      celdaConsola.setValue(valorActual ? valorActual + "\n" + msg : msg);
      SpreadsheetApp.flush();
    }
    console.log(msg);
  };
  if (celdaConsola) celdaConsola.clearContent();
  log(`🚀 Iniciando actualización diaria: ${hoy.toLocaleDateString()}`);
  const mXC = obtenerObjetoMaterias(ss.getSheetByName("masdatos"), "A2:B");
  const folder = DriveApp.getFolderById(CFG.FLD);
  let archivosNoEncontrados = [];
  for (let curso in mXC) {
    const nombreArchivo = getNm(curso);
    const files = folder.getFilesByName(nombreArchivo);
    if (!files.hasNext()) {
      log(`⚠️ ERROR: No se encontró el archivo ${nombreArchivo}`);
      archivosNoEncontrados.push(nombreArchivo);
      continue;
    }
    try {
      const ssCurso = SpreadsheetApp.openById(files.next().getId());
      log(`📂 Procesando curso: ${curso}`);
      mXC[curso].forEach(nombreM => {
        const hoja = ssCurso.getSheetByName(nombreM);
        if (!hoja) return;
        const fechas = hoja.getRange("A3:A214").getValues();
        let indiceEncontrado = -1;
        let esMatchExacto = false;
        for (let i = 0; i < fechas.length; i++) {
          let fCelda = fechas[i][0];
          if (fCelda instanceof Date) {
            fCelda.setHours(0,0,0,0);
            if (fCelda.getTime() === hoy.getTime()) {
              indiceEncontrado = i;
              esMatchExacto = true;
              break; 
            } else if (fCelda < hoy) {
              indiceEncontrado = i;
            } else { break; }
          }
        }
        if (indiceEncontrado === -1) return;
        let rangosLibres = [hoja.getRange("E1")]; 
        let filaActual = indiceEncontrado + 3;
        if (esMatchExacto) {
          if (filaActual >= 3) rangosLibres.push(hoja.getRange(`B${filaActual}:E${filaActual}`));
          if (indiceEncontrado > 0) {
            let filaPrevia = filaActual - 1;
            rangosLibres.push(hoja.getRange(`B${filaPrevia}:E${filaPrevia}`));
          }
        } else {
          if (filaActual >= 3) rangosLibres.push(hoja.getRange(`B${filaActual}:E${filaActual}`));
        }
        const proteccion = hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
        if (proteccion) {
          proteccion.setUnprotectedRanges(rangosLibres);
        }
      });
      log(`   ✅ ${curso} actualizado.`);
    } catch (e) {
      log(`❌ Error crítico en ${curso}: ${e.message}`);
    }
  }
  if (archivosNoEncontrados.length > 0) {
    const cuerpoEmail = `ATENCIÓN: Situación de riesgo detectada.\n\nLos siguientes archivos de Libro de Temas NO fueron encontrados en la carpeta configurada:\n\n- ${archivosNoEncontrados.join("\n- ")}`;
    MailApp.sendEmail({
      to: EMAIL_ALERTA,
      subject: "🚨 ALERTA: Archivos de Libros de Temas Faltantes",
      body: cuerpoEmail
    });
    log("📧 Email de alerta enviado a " + EMAIL_ALERTA);
  }
  log("🏁 Proceso finalizado.");
}


function auditoriaSemanalCargas() {
  const FECHA_TEST_STR = null;
  const MODO_PRUEBA = false;
  const MODO_PRUEBA_2 = false;
  const LIMITE_CURSOS_PRUEBA = null;
  const EMAIL_ADMIN = "luisoroverde@gmail.com";
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
  if (MODO_PRUEBA) log("🧪 MODO PRUEBA 1: Solo Consola.");
  else if (MODO_PRUEBA_2) log("📩 MODO PRUEBA 2: Mails de docentes redirigidos a Luis.");
  else log("🚀 EJECUCIÓN REAL: Mails saldrán hacia los docentes.");
  let fechaRef = FECHA_TEST_STR ? new Date(FECHA_TEST_STR.split("/")[2], FECHA_TEST_STR.split("/")[1]-1, FECHA_TEST_STR.split("/")[0]) : new Date();
  fechaRef.setHours(0,0,0,0);
  const diaSem = fechaRef.getDay() === 0 ? 7 : fechaRef.getDay();
  const lunesPasado = new Date(fechaRef);
  lunesPasado.setDate(fechaRef.getDate() - diaSem - 6);
  const domingoPasado = new Date(lunesPasado);
  domingoPasado.setDate(lunesPasado.getDate() + 6);
  log(`📅 Semana analizada: ${lunesPasado.toLocaleDateString()} al ${domingoPasado.toLocaleDateString()}`);
  const ssHorario = SpreadsheetApp.openById(CFG.HORARIO);
  const hPers = ssHorario.getSheetByName("personal");
  const dPers = hPers.getRange("B2:C" + hPers.getLastRow()).getValues();
  const limpiarTexto = (t) => t.toString().trim().toUpperCase().replace(/\s+/g, " ");
  const dictEmails = dPers.reduce((acc, [nom, mail]) => {
    if (nom && mail) acc[limpiarTexto(nom)] = mail.toString().trim();
    return acc;
  }, {});
  let reporteFaltas = {};
  const mXC = obtenerObjetoMaterias(ss.getSheetByName("masdatos"), "A2:B");
  const folder = DriveApp.getFolderById(CFG.FLD);
  let cursosContador = 0;
  for (let curso in mXC) {
    if (LIMITE_CURSOS_PRUEBA && cursosContador >= LIMITE_CURSOS_PRUEBA) break;
    cursosContador++;
    const nombreArchivo = getNm(curso);
    const files = folder.getFilesByName(nombreArchivo);
    if (!files.hasNext()) continue;
    try {
      const ssC = SpreadsheetApp.openById(files.next().getId());
      log(`📂 [${cursosContador}] Analizando: ${curso}`);
      mXC[curso].forEach(nombreM => {
        const h = ssC.getSheetByName(nombreM);
        if (!h) return;
        const profRaw = h.getRange("E1").getValue();
        const profNormal = limpiarTexto(profRaw);
        if (!profNormal) return;
        const data = h.getRange("A3:E214").getValues();
        data.forEach(fila => {
          const fCelda = fila[0];
          if (fCelda instanceof Date && fCelda >= lunesPasado && fCelda <= domingoPasado) {
            const colB = fila[1].toString().trim();
            const colC = fila[2].toString().trim();
            const colD = fila[3].toString().trim();
            const colE = fila[4].toString().trim();
            if (colB === "" && colC === "" && colD === "" && colE === "") {
              if (!reporteFaltas[profNormal]) reporteFaltas[profNormal] = { nombreReal: profRaw, faltas: [] };
              reporteFaltas[profNormal].faltas.push(`${Utilities.formatDate(fCelda, "GMT-3", "dd/MM")} (${curso} - ${nombreM})`);
            }
          }
        });
      });
    } catch (e) { log(`❌ Error en ${curso}: ${e.message}`); }
  }
  log(`-----------------------------------`);
  let informeAdmin = "📊 RESUMEN AUDITORÍA SEMANAL\n\n";
  for (let pKey in reporteFaltas) {
    const p = reporteFaltas[pKey];
    const emailRealDocente = dictEmails[pKey];
    let listaHtml = p.faltas.map(f => "<li>" + f + "</li>").join("");
    let cuerpoTexto = `Estimado/a ${p.nombreReal},\n\nFaltan registros en el Libro de Temas de la semana pasada:\n${p.faltas.join("\n")}\n\nRecuerde completar las planillas en Drive o el Formulario.`;
    let cuerpoHtml = `<p>Estimado/a <strong>${p.nombreReal}</strong>,</p><p>Se ha detectado que faltan registros en el <strong>Libro de Temas Digital</strong> correspondientes a la semana pasada:</p><ul>${listaHtml}</ul><br><div style="font-size: 11px; color: #666; border-top: 1px solid #eee; padding-top: 10px;"><p>Recuerde que hay dos formas de cargar el libro de temas digital:</p><ol><li><strong>Por las planillas en Drive</strong>: <a href="https://drive.google.com/drive/folders/1FaPpRKnZWH9efcGb94S6SOm5BW9EaQpg">Acceder a Carpetas Drive</a></li><li><strong>Para cargas más antiguas</strong>, utilice el Formulario: <a href="https://forms.gle/ByzE5khGpbauaYpR8">Acceder al Formulario</a></li></ol><p><em>Ante cualquier duda, consultar en Secretaría.</em></p></div>`;
    if (MODO_PRUEBA) {
      log(`📝 SIMULACIÓN CONSOLA: ${p.nombreReal} -> ${emailRealDocente || "SIN MAIL"}`);
    } else {
      let destinatario = MODO_PRUEBA_2 ? EMAIL_ADMIN : emailRealDocente;
      let asunto = (MODO_PRUEBA_2 ? `[TEST PARA: ${p.nombreReal}] ` : "") + "Recordatorio: Carga de Libro de Temas";
      if (destinatario) {
        MailApp.sendEmail({ to: destinatario, subject: asunto, body: cuerpoTexto, htmlBody: cuerpoHtml });
        if (MODO_PRUEBA_2) log(`📧 Redirigido a Luis: Falta de ${p.nombreReal}`);
      } else {
        log(`⚠️ No se pudo enviar mail a ${p.nombreReal} (Email no encontrado)`);
      }
    }
    informeAdmin += `👤 PROF: ${p.nombreReal}\n${p.faltas.join("\n")}\n📧 Email original: ${emailRealDocente || "⚠️ NO ENCONTRADO"}\n\n`;
  }
  let asuntoAdmin = (MODO_PRUEBA ? "[SIMULACIÓN] " : (MODO_PRUEBA_2 ? "[PRUEBA REDIRIGIDA] " : "")) + "Resumen Auditoría Semanal";
  MailApp.sendEmail(EMAIL_ADMIN, asuntoAdmin, informeAdmin);
  log(`🏁 Proceso finalizado.`);
}

function actualizarDesdePanelControl3() {
  const ss = SpreadsheetApp.openById(CFG.SS_DATA);
  const panel = ss.getSheetByName("Panel de Control");
  const hojaMasDatos = ss.getSheetByName("masdatos");
  const hojaDias = ss.getSheetByName("dias");
  const celdaConsola = panel.getRange("B6");
  celdaConsola.clearContent();
  const log = (msg) => {
    const v = celdaConsola.getValue();
    celdaConsola.setValue(v ? v + "\n" + msg : msg);
    SpreadsheetApp.flush();
  };
  log("🚀 Iniciando Sincronización Total...");
  const filtroCursos = panel.getRange("C3").getValue() ? panel.getRange("C3").getValue().toString().toUpperCase().split(",") : [];
  const filtroMaterias = panel.getRange("D3").getValue() ? panel.getRange("D3").getValue().toString().toUpperCase().split(",") : [];
  const añoActual = new Date().getFullYear();
  const inicioClases = new Date(añoActual, 2, 1); 
  const finClases = new Date(añoActual, 11, 20);
  const datosDias = hojaDias.getRange(2, 5, 213, 4).getValues(); 
  const dictObs = {};
  datosDias.forEach(f => {
    if (f[0] instanceof Date) {
      const key = Utilities.formatDate(f[0], "GMT-3", "yyyyMMdd");
      dictObs[key] = f[3] ? f[3].toString().trim() : ""; 
    }
  });
  const mXC = obtenerObjetoMaterias(hojaMasDatos, "A2:B");
  const folder = DriveApp.getFolderById(CFG.FLD);
  for (let curso in mXC) {
    if (filtroCursos.length > 0 && !filtroCursos.includes(curso.toUpperCase())) continue;
    const files = folder.getFilesByName(getNm(curso));
    if (!files.hasNext()) continue;
    const ssCurso = SpreadsheetApp.openById(files.next().getId());
    log(`📂 Procesando Curso: ${curso}`);
    mXC[curso].forEach(nombreM => {
      if (filtroMaterias.length > 0 && !filtroMaterias.includes(nombreM.toUpperCase())) return;
      const hoja = ssCurso.getSheetByName(nombreM);
      if (!hoja) return;
      const infoHorario = buscarDiasMateria(nombreM, curso);
      if (infoHorario.dias.length === 0) return;
      if (infoHorario.profesor) hoja.getRange("E1").setValue(infoHorario.profesor);
      const fechasIdeales = new Set();
      let fLoop = new Date(inicioClases);
      while (fLoop <= finClases) {
        if (infoHorario.dias.includes(fLoop.getDay())) {
          fechasIdeales.add(Utilities.formatDate(fLoop, "GMT-3", "yyyyMMdd"));
        }
        fLoop.setDate(fLoop.getDate() + 1);
      }
      const rango = hoja.getRange("A3:E214");
      const valoresActuales = rango.getValues();
      const resultadoFinal = [];
      let huboCambio = false;
      valoresActuales.forEach(fila => {
        if (!(fila[0] instanceof Date)) return;
        const key = Utilities.formatDate(fila[0], "GMT-3", "yyyyMMdd");
        const tieneDatos = (fila[2].toString().trim() !== "" || fila[3].toString().trim() !== "");
        if (fechasIdeales.has(key)) {
          if (!tieneDatos && fila[4] !== dictObs[key]) {
            fila[4] = dictObs[key] || "";
            huboCambio = true;
          }
          resultadoFinal.push(fila);
          fechasIdeales.delete(key);
        } else {
          if (tieneDatos) resultadoFinal.push(fila);
          else huboCambio = true;
        }
      });
      if (fechasIdeales.size > 0) {
        huboCambio = true;
        fechasIdeales.forEach(keyStr => {
          const y = parseInt(keyStr.substring(0,4)), m = parseInt(keyStr.substring(4,6))-1, d = parseInt(keyStr.substring(6,8));
          resultadoFinal.push([new Date(y, m, d), "", "", "", dictObs[keyStr] || ""]);
        });
      }
      if (huboCambio) {
        log(`   ✨ ${nombreM}: Cronograma y Profesor actualizados.`);
        resultadoFinal.sort((a, b) => a[0].getTime() - b[0].getTime());
        rango.clearContent();
        if (resultadoFinal.length > 0) {
          hoja.getRange(3, 1, resultadoFinal.length, 5).setValues(resultadoFinal.slice(0, 212));
        }
      }
    });
  }
  log("✅ Proceso finalizado exitosamente.");
}

function actualizarPermisosSemanalesEficiente() {
  actualizarSoloProfesores();

  const ss = SpreadsheetApp.openById(CFG.SS_DATA);
  const adminEmail = Session.getEffectiveUser().getEmail();
  
  // 1. CÁLCULO DE FECHA DE CORTE (Lunes de la semana pasada)
  // Si hoy es Miércoles 18/03/2026, el "lunes de la semana pasada" fue el 09/03/2026.
  // Todo lo anterior a esa fecha se bloquea.
  const hoy = new Date();
  const diaSemana = hoy.getDay(); 
  const diferenciaAlLunesActual = (diaSemana === 0 ? 6 : diaSemana - 1);
  const lunesSemanaPasada = new Date(hoy);
  lunesSemanaPasada.setDate(hoy.getDate() - diferenciaAlLunesActual - 7);
  lunesSemanaPasada.setHours(0, 0, 0, 0);

  const toVal = (d) => {
    if (!(d instanceof Date)) return 0;
    return parseInt(Utilities.formatDate(d, "GMT-3", "yyyyMMdd"));
  };

  const valCorte = toVal(lunesSemanaPasada);
  const timeStamp = Utilities.formatDate(hoy, "GMT-3", "dd/MM HH:mm");

  Logger.log(`🚀 Iniciando Bloqueo Semanal. Fecha de corte: ${valCorte}`);

  const mXC = obtenerObjetoMaterias(ss.getSheetByName("masdatos"), "A2:B");
  const folder = DriveApp.getFolderById(CFG.FLD);

  for (let curso in mXC) {
    const files = folder.getFilesByName(getNm(curso));
    if (!files.hasNext()) continue;

    try {
      const ssCurso = SpreadsheetApp.openById(files.next().getId());
      Logger.log(`📂 Procesando: ${curso}`);

      mXC[curso].forEach(nombreM => {
        const hoja = ssCurso.getSheetByName(nombreM);
        if (!hoja) return;

        // Marca temporal en G1
        hoja.getRange("G1").setValue(timeStamp);

        // Limpieza de protecciones previas para no duplicar
        const protExistentes = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        protExistentes.forEach(p => {
          if (p.getDescription() === "Mantenimiento_Pasado") p.remove();
        });

        // Lectura de datos (A=0, B=1, C=2, D=3)
        const datos = hoja.getRange("A3:D214").getValues();
        let rangosABloquear = [];
        let inicioRango = null;

        // 2. LÓGICA DE AGRUPACIÓN EFICIENTE
        for (let i = 0; i < datos.length; i++) {
          const valFila = toVal(datos[i][0]);
          const tieneDatos = datos[i][2].toString().trim() !== "" || 
                            datos[i][3].toString().trim() !== "";

          // REQUISITOS: Anterior a la semana pasada Y tener datos en C o D
          if (valFila > 0 && valFila < valCorte && tieneDatos) {
            if (inicioRango === null) inicioRango = i + 3; // +3 por el offset de filas
          } else {
            if (inicioRango !== null) {
              rangosABloquear.push(`B${inicioRango}:E${i + 2}`);
              inicioRango = null;
            }
          }
        }
        
        // Cerrar último rango si quedó abierto
        if (inicioRango !== null) {
          rangosABloquear.push(`B${inicioRango}:E${datos.length + 2}`);
        }

        // 3. APLICACIÓN DE BLOQUEOS (Solo por grupos unificados)
        rangosABloquear.forEach(rangoStr => {
          const p = hoja.getRange(rangoStr).protect().setDescription("Mantenimiento_Pasado");
          p.removeEditors(p.getEditors());
          if (p.canDomainEdit()) p.setDomainEdit(false);
          try { p.addEditor(adminEmail); } catch(e) {}
        });
      });
    } catch (e) { 
      Logger.log(`❌ Error en ${curso}: ${e.message}`); 
    }
  }
  Logger.log("🏁 Proceso finalizado.");
}


function actualizarSoloProfesores() {
  const ss = SpreadsheetApp.openById(CFG.SS_DATA);
  const mXC = obtenerObjetoMaterias(ss.getSheetByName("masdatos"), "A2:B");
  const folder = DriveApp.getFolderById(CFG.FLD);

  for (let curso in mXC) {
    const files = folder.getFilesByName(getNm(curso));
    if (!files.hasNext()) continue;

    try {
      const ssCurso = SpreadsheetApp.openById(files.next().getId());
      mXC[curso].forEach(nombreM => {
        const hoja = ssCurso.getSheetByName(nombreM);
        if (!hoja) return;

        const info = buscarDiasMateria(nombreM, curso);
        const celdaE1 = hoja.getRange("E1");
        const anterior = celdaE1.getValue();

        if (info.profesor && info.profesor !== anterior) {
          celdaE1.setValue(info.profesor);
          console.log(`[CAMBIO] ${curso} - ${nombreM}: ${anterior || "Vacío"} -> ${info.profesor}`);
        }
      });
    } catch (e) {
      console.error(`Error en ${curso}: ${e.message}`);
    }
  }
}



function verificarIntegridadArchivosDiaria() {
  const EMAIL_ADMIN = "luisoroverde@gmail.com";
  const ss = SpreadsheetApp.openById(CFG.SS_DATA);
  const hojaDatos = ss.getSheetByName("masdatos");
  
  if (!hojaDatos) {
    Logger.log("❌ Error: No se encontró la hoja 'masdatos'.");
    return;
  }

  // 1. Obtener la lista de materias/cursos desde tu objeto de configuración
  // Usamos tu función existente obtenerObjetoMaterias
  const mXC = obtenerObjetoMaterias(hojaDatos, "A2:B");
  const folder = DriveApp.getFolderById(CFG.FLD);
  
  let archivosFaltantes = [];
  Logger.log("🔍 Iniciando verificación diaria de archivos...");

  // 2. Revisar uno por uno si el archivo existe en Drive
  for (let curso in mXC) {
    const nombreEsperado = getNm(curso); // Tu función que traduce el código al nombre del archivo
    const files = folder.getFilesByName(nombreEsperado);
    
    if (!files.hasNext()) {
      Logger.log(`⚠️ FALTANTE: ${nombreEsperado}`);
      archivosFaltantes.push(`- ${nombreEsperado} (Curso: ${curso})`);
    }
  }

  // 3. Si hay faltantes, enviar UN solo correo consolidado
  if (archivosFaltantes.length > 0) {
    const asunto = "⚠️ ALERTA DIARIA: Libros de Temas no encontrados";
    const cuerpo = "Se ha detectado que los siguientes archivos no están en la carpeta de Drive o tienen un nombre incorrecto:\n\n" + 
                   archivosFaltantes.join("\n") + 
                   "\n\n---\nEste es un reporte automático del sistema de Libro de Temas.";
    
    MailApp.sendEmail(EMAIL_ADMIN, asunto, cuerpo);
    Logger.log("📩 Alerta enviada a " + EMAIL_ADMIN);
  } else {
    Logger.log("✅ Todos los archivos están presentes.");
  }
}
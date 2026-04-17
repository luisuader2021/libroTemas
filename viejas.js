function actualizarPermisosFlexibles() {
  const FECHA_MANUAL = null;
  const CURSOS_FILTRO = null;
  const PRIMERA_VEZ = false;
  const ss = SpreadsheetApp.openById(CFG.SS_DATA);
  const panel = ss.getSheetByName("Panel de Control");
  const celdaConsola = panel ? panel.getRange("B6") : null;
  const toVal = (d) => {
    if (!(d instanceof Date)) return 0;
    return parseInt(Utilities.formatDate(d, "GMT-3", "yyyyMMdd"));
  };
  let fechaRef = FECHA_MANUAL ? new Date(FECHA_MANUAL.split("/")[2], FECHA_MANUAL.split("/")[1] - 1, FECHA_MANUAL.split("/")[0]) : new Date();
  const valRef = toVal(fechaRef);
  const log = (msg) => {
    if (celdaConsola) {
      const v = celdaConsola.getValue();
      celdaConsola.setValue(v ? v + "\n" + msg : msg);
      SpreadsheetApp.flush();
    }
    console.log(msg);
  };
  if (celdaConsola) celdaConsola.clearContent();
  log(`🛡️ Ref: ${valRef} (GMT-3) | Pasado Bloqueado | Futuro Libre`);
  const mXC = obtenerObjetoMaterias(ss.getSheetByName("masdatos"), "A2:B");
  const folder = DriveApp.getFolderById(CFG.FLD);
  const adminEmail = Session.getEffectiveUser().getEmail();
  for (let curso in mXC) {
    if (CURSOS_FILTRO && !CURSOS_FILTRO.split(",").includes(curso.toUpperCase())) continue;
    const files = folder.getFilesByName(getNm(curso));
    if (!files.hasNext()) continue;
    try {
      const ssCurso = SpreadsheetApp.openById(files.next().getId());
      log(`📂 Curso: ${curso}`);
      mXC[curso].forEach(nombreM => {
        const hoja = ssCurso.getSheetByName(nombreM);
        if (!hoja) return;
        const protExistentes = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        if (PRIMERA_VEZ) {
          protExistentes.forEach(p => p.remove());
        } else {
          protExistentes.forEach(p => {
            const range = p.getRange();
            if (range.getColumn() >= 2 && range.getLastColumn() <= 5 && range.getRow() >= 3) {
              p.remove();
            }
          });
        }
        const aplicarProteccion = (rango, desc) => {
          const p = rango.protect().setDescription(desc);
          p.removeEditors(p.getEditors());
          if (p.canDomainEdit()) p.setDomainEdit(false);
          try { p.addEditor(adminEmail); } catch(e) {}
        };
        if (PRIMERA_VEZ) {
          aplicarProteccion(hoja.getRange("A1:A"), "Fijo_A");
          aplicarProteccion(hoja.getRange("B1:D2"), "Fijo_B1D2");
          aplicarProteccion(hoja.getRange("E2"), "Fijo_E2");
          aplicarProteccion(hoja.getRange("F1:G"), "Fijo_Estructura");
        }
        const datosA = hoja.getRange("A3:A214").getValues();
        let indiceUltimaOcurrida = -1;
        for (let i = 0; i < datosA.length; i++) {
          let valFila = toVal(datosA[i][0]);
          if (valFila > 0 && valFila <= valRef) {
            indiceUltimaOcurrida = i;
          } else if (valFila > valRef) {
            break; 
          }
        }
        if (indiceUltimaOcurrida !== -1) {
          let filaPenultima = -1;
          for (let j = indiceUltimaOcurrida - 1; j >= 0; j--) {
            if (datosA[j][0] instanceof Date) {
              filaPenultima = j + 3;
              break;
            }
          }
          let filaLimiteBloqueo = (filaPenultima !== -1) ? filaPenultima - 1 : indiceUltimaOcurrida + 2;
          if (filaLimiteBloqueo >= 3) {
            aplicarProteccion(hoja.getRange(`B3:E${filaLimiteBloqueo}`), "Mantenimiento_Pasado");
          }
        }
      });
    } catch (e) { log(`❌ Error: ${e.message}`); }
  }
  log("🏁 Finalizado.");
}
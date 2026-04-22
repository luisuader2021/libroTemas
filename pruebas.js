function auditoriaDiferenciasHorario() {
  const ss = SpreadsheetApp.openById(CFG.SS_DATA);
  const panel = ss.getSheetByName("Panel de Control");
  const hojaMasDatos = ss.getSheetByName("masdatos");
  const celdaConsola = panel.getRange("B6");
  
  celdaConsola.clearContent();
  const log = (msg) => {
    const v = celdaConsola.getValue();
    celdaConsola.setValue(v ? v + "\n" + msg : msg);
    SpreadsheetApp.flush();
  };

  log("🔍 Iniciando Auditoría de Horarios (Hoy en adelante)... \n");

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  
  const mXC = obtenerObjetoMaterias(hojaMasDatos, "A2:B");
  const folder = DriveApp.getFolderById(CFG.FLD);
  let huboDiferencias = false;

  for (let curso in mXC) {
    const files = folder.getFilesByName(getNm(curso));
    if (!files.hasNext()) continue;

    const ssCurso = SpreadsheetApp.openById(files.next().getId());
    
    mXC[curso].forEach(nombreM => {
      const hoja = ssCurso.getSheetByName(nombreM);
      if (!hoja) return;

      // 1. Obtener días según el HORARIO MAESTRO
      const horarioMaestro = buscarDiasMateria(nombreM, curso);
      const diasMaestro = horarioMaestro.dias; // Array de números [1, 3] (Lunes, Miércoles)

      // 2. Obtener días que tiene cargados el LIBRO actualmente
      const fechasLibro = hoja.getRange("A3:A214").getValues();
      const diasEnLibro = new Set();
      
      fechasLibro.forEach(f => {
        if (f[0] instanceof Date) {
          let fechaFila = f[0];
          fechaFila.setHours(0,0,0,0);
          if (fechaFila >= hoy) {
            diasEnLibro.add(fechaFila.getDay());
          }
        }
      });

      const arrayDiasLibro = [...diasEnLibro].sort();

      // 3. Comparar ambos arrays
      const sonIguales = (diasMaestro.length === arrayDiasLibro.length) && 
                         diasMaestro.every((val, index) => val === arrayDiasLibro[index]);

      if (!sonIguales && arrayDiasLibro.length > 0) {
        huboDiferencias = true;
        const nombresDias = ["Dom", "Lun", "Mar", "Mié", "Jue", "Vie", "Sáb"];
        const txtMaestro = diasMaestro.map(d => nombresDias[d]).join(", ");
        const txtLibro = arrayDiasLibro.map(d => nombresDias[d]).join(", ");
        
        log(`⚠️ [${curso}] ${nombreM}:`);
        log(`   - En Horario: ${txtMaestro}`);
        log(`   - En Libro:   ${txtLibro}`);
      }
    });
  }

  if (!huboDiferencias) {
    log("✅ Todos los libros coinciden con el horario maestro.");
  } else {
    log("\n🏁 Auditoría finalizada. Se encontraron discrepancias.");
  }
}
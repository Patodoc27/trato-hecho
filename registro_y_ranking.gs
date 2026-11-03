/*
 * Trato Hecho — Apps Script de registro + ranking
 * =====================================================
 * REEMPLAZA al código actual del script de la planilla
 * (el que solo grababa filas). Mantiene el doPost igual:
 * cada partida terminada agrega UNA fila por estudiante.
 * Agrega el doGet para que el juego pueda LEER el ranking.
 *
 * Cómo actualizar SIN cambiar la URL /exec:
 *   1. En la planilla: Extensiones → Apps Script.
 *   2. Borrá el código viejo y pegá TODO este archivo.
 *   3. Guardar (disquete).
 *   4. Implementar → Administrar implementaciones → lápiz
 *      → Versión: "Nueva versión" → Implementar.
 *   5. Probá en incógnito:
 *      https://.../exec?top=5            → JSON de filas
 *      https://.../exec?top=5&callback=x → x([...]) (JSONP)
 *
 * IMPORTANTE: debe quedar "Ejecutar como: Yo" y
 * "Quién puede acceder: Cualquier persona".
 */

// Orden de columnas que usa el juego (si la planilla ya tiene
// encabezados en la fila 1, se respetan tal cual están).
const CAMPOS = ["fecha","nombre","apellido","seccion","avatar","modo",
  "aciertos","respondidas","descuento","multiplicador","montoFinal","resultado"];

function hoja_(){ return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; }

function encabezados_(sh){
  const ultCol = Math.max(sh.getLastColumn(), CAMPOS.length);
  if(sh.getLastRow()===0){ return null; }
  const fila1 = sh.getRange(1,1,1,ultCol).getValues()[0]
    .map(v=>String(v).toLowerCase().trim());
  // ¿La fila 1 es encabezado? (tiene "apellido" o "nombre")
  if(fila1.indexOf("apellido")>=0 || fila1.indexOf("nombre")>=0) return fila1;
  return null;
}

function doPost(e){
  const sh = hoja_();
  const d = JSON.parse(e.postData.contents||"{}");
  let head = encabezados_(sh);
  if(!head){
    if(sh.getLastRow()===0){ sh.getRange(1,1,1,CAMPOS.length).setValues([CAMPOS]); head=CAMPOS; }
    else { head=CAMPOS; } // planilla sin encabezados: se usa el orden CAMPOS
  }
  const fila = head.map(h=>{
    if(h==="fecha") return new Date();
    return d[h]!==undefined ? d[h] : "";
  });
  sh.appendRow(fila);
  return ContentService.createTextOutput(JSON.stringify({ok:true}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e){
  const sh = hoja_();
  const p = e && e.parameter ? e.parameter : {};
  const top = Math.max(1, Math.min(2000, parseInt(p.top||"500",10)||500));
  let head = encabezados_(sh);
  let desde = 2;
  if(!head){ head=CAMPOS; desde=1; }
  const ult = sh.getLastRow();
  let salida = [];
  if(ult>=desde){
    const datos = sh.getRange(desde,1,ult-desde+1,head.length).getValues();
    salida = datos.map(fila=>{
      const o={};
      head.forEach((h,i)=>{ if(h) o[h]=fila[i]; });
      return o;
    }).filter(o=>o.nombre||o.apellido);
  }
  salida = salida.slice(-top); // las más recientes
  const json = JSON.stringify(salida);
  if(p.callback){
    return ContentService.createTextOutput(p.callback+"("+json+")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

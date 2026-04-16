// ===============================================================
// 📋 CONTROL REMITOS + ALERTAS DE STOCK - VERSIÓN FINAL
// ===============================================================

// ---------------------------------------------------------------
// ⚙️ CONFIGURACIÓN GLOBAL
// ---------------------------------------------------------------
const MI_EMAIL        = "mikeaspulido80@gmail.com";
const ID_HOJA_SALIDAS = "10sIFWNX3UtOHJD_26isE6n_lAH5osU7Y6NWVHxrUUmU";
const INVENTARIO_ID   = "10sIFWNX3UtOHJD_26isE6n_lAH5osU7Y6NWVHxrUUmU";

const WEBHOOK_REMITOS = "https://discordapp.com/api/webhooks/1436819050326917200/RS1h3JfJ4_NuYnSr2_78-eyjYHXp_Q2SOCLq6efg4vcYG0XW8ZQUBD8k2RK2FRWSdyv9";
const WEBHOOK_STOCK   = "https://discordapp.com/api/webhooks/1443581720799608884/4_ZsO_RB5xK9I2Ju2zeEZrwC0sxzrDIdkmNeaEHxkwyScV0VYedbZr_05WDWFrSnuYHj";

const COLOR_FALTANTE = "#ff0000";
const COLOR_NORMAL   = "#434343";

// ---------------------------------------------------------------
// 📍 MAPEO DE FILAS EN HOJA IMPRIMIR
// ---------------------------------------------------------------
const MAPEO_FILAS = [
  { fila: 8,  nombre: 'C-SM1' },
  { fila: 9,  nombre: 'C-SM2' },
  { fila: 10, nombre: 'C-SM3' },
  { fila: 11, nombre: 'C-SM4' },
  { fila: 12, nombre: 'C-SM5' },
  { fila: 14, nombre: 'C-AL1' },
  { fila: 15, nombre: 'C-AL2' },
  { fila: 16, nombre: 'C-AL3' },
  { fila: 17, nombre: 'C-AL4' },
  { fila: 18, nombre: 'C-AL5' },
  { fila: 19, nombre: 'C-AL6' },
  { fila: 20, nombre: 'C-AL7' },
  { fila: 22, nombre: 'C-SL1' },
  { fila: 23, nombre: 'C-SL2' },
  { fila: 24, nombre: 'C-SL3' },
  { fila: 25, nombre: 'C-SL4' },
  { fila: 26, nombre: 'C-SL5' },
  { fila: 27, nombre: 'C-SL6' }
];

const MAPA_NOMBRE_A_FILA = {};
MAPEO_FILAS.forEach(item => { MAPA_NOMBRE_A_FILA[item.nombre] = item.fila; });

const NOMBRES_CONTROLADOS = MAPEO_FILAS.map(i => i.nombre)
  .concat(["REMITO AL", "REMITO SL", "REMITO SM", "RESUMEN DS"]);

var procesando = false;

function esTrue(val) {
  return val === true || val === 'TRUE' || val === 'true';
}


// ===============================================================
// ✏️  onEdit PRINCIPAL
// ===============================================================
function onEdit(e) {
  if (procesando || !e || !e.range) return;
  procesando = true;

  try {
    const range      = e.range;
    const hoja       = range.getSheet();
    const nombreHoja = hoja.getName();
    const celda      = range.getA1Notation();
    const fila       = range.getRow();
    const col        = range.getColumn();
    const valor      = e.value;

    // ===============================================================
    // HOJA "IMPRIMIR"
    // ===============================================================
    if (nombreHoja === "IMPRIMIR") {

      // B2: SELLAR (BLANCO) O DES-SELLAR (GRIS)
      if (celda === 'B2') {
        const marcado = esTrue(valor);
        hoja.getRange("A1:B2").setFontColor(marcado ? "#ffffff" : COLOR_GRIS);
        if (marcado) {
          hoja.getRange("B2").setValue(true);
          hoja.toast("🎨 Remito sellado.", "ESTÉTICA");
        } else {
          hoja.toast("🔓 Remito liberado.", "ESTÉTICA");
        }
        SpreadsheetApp.flush();
        return;
      }

      // A1: MOSTRAR HOJAS SELECCIONADAS
      if (celda === 'A1' && esTrue(valor)) {
        const ultima = hoja.getLastRow();
        if (ultima > 1) {
          const datosAC = hoja.getRange(2, 1, ultima - 1, 3).getValues();
          for (let i = 0; i < datosAC.length; i++) {
            if (esTrue(datosAC[i][0]) && !esTrue(datosAC[i][2])) {
              hoja.getRange(i + 2, 1).setValue(false);
            }
          }
        }
        hoja.getRange('A1').setValue(false);
        Utilities.sleep(150);
        if (typeof mostrarHojasSeleccionadas === 'function') mostrarHojasSeleccionadas();
        return;
      }

      // B1: DESMARCAR TODOS LOS B2
      if (celda === 'B1' && esTrue(valor)) {
        if (typeof desmarcarTodosLosB2 === 'function') desmarcarTodosLosB2();
        hoja.getRange('B1').setValue(false);
        return;
      }

      // ===============================================================
      // MANEJO UNIFICADO DE CELDA C1 (BORRADO O NOMBRE)
      // ===============================================================
      if (celda === 'C1' && valor !== undefined) {
        
        // OPCIÓN 1: Si es el Checkbox de borrado (TRUE)
        if (esTrue(valor)) {
          const ss = SpreadsheetApp.getActiveSpreadsheet();
          const hojas = ss.getSheets();
          let borradas = 0;

          hojas.forEach(h => {
            const n = h.getName();
            // 🛡️ NUEVA CONDICIÓN: Solo si la hoja NO está oculta
            if (!h.isSheetHidden()) {
              // Filtra solo las hojas que empiezan con los prefijos indicados
              if (n.startsWith("C-AL") || n.startsWith("C-SM") || n.startsWith("C-SL")) {
                const b2Val = h.getRange("B2").getValue();
                
                // Si B2 no es true (no está sellado), se elimina
                if (!esTrue(b2Val)) {
                  ss.deleteSheet(h);
                  borradas++;
                }
              }
            }
          });

          hoja.getRange('C1').setValue(false); // Resetea el checkbox
          hoja.toast("🔥 Se eliminaron " + borradas + " hojas visibles.", "LIMPIEZA");
        } 
        
        // OPCIÓN 2: Si es el nombre del cliente (TEXTO)
        else if (typeof valor === 'string' && valor.trim() !== "") {
          const nombreFormateado = valor.split(" ")
            .map(p => p.charAt(0).toUpperCase() + p.slice(1).toLowerCase())
            .join(" ");
          
          if (valor !== nombreFormateado) {
            hoja.getRange("C1").setValue(nombreFormateado);
          }
        }
        return;
      }

      // CHECKS EN IMPRIMIR (A o C, fila >= 2)
      if ((col === 1 || col === 3) && fila >= 2) {
        if (col === 3 && esTrue(valor) && fila >= 8) {
          const nombreEnB = hoja.getRange(fila, 2).getValue();
          if (typeof restaurarNombreEnImprimir === 'function') {
            restaurarNombreEnImprimir(String(nombreEnB).trim());
          }
        }
        Utilities.sleep(150);
        if (typeof mostrarHojasSeleccionadas === 'function') mostrarHojasSeleccionadas();
        return;
      }

      return;
    }

// ===============================================================
    // HOJAS DE REMITOS (C-AL, C-SM, C-SL) - AUTO-ACTUALIZAR IMPRESIÓN
    // ===============================================================
    const esRemito = (nombreHoja.startsWith("C-AL") || 
                      nombreHoja.startsWith("C-SM") || 
                      nombreHoja.startsWith("C-SL"));
    
    if (esRemito) {
      const filaImprimir = MAPA_NOMBRE_A_FILA[nombreHoja];
      
      if (filaImprimir) {
        const hojaImp = obtenerHojaImprimir();
        
        if (hojaImp) {
          // 1. SI SE EDITA C1 (CLIENTE)
          if (celda === 'C1') {
            let textoC1 = valor ? String(valor).trim().toLowerCase() : "";
            
            // Convertir a Formato De Nombre Propio (Ej: juan perez -> Juan Perez)
            const nombreFormateado = textoC1.replace(/(^\w|\s\w)/g, m => m.toUpperCase());
            
            hojaImp.getRange(filaImprimir, 2).setValue(nombreFormateado); // Columna B
            
            if (nombreFormateado !== "") {
              hojaImp.getRange(filaImprimir, 1).setValue(true);         // Columna A (Check)
            }
          } 
          
          // 2. SI SE CARGAN PRODUCTOS (COL 1 o 2, FILA >=4)
          else if ((col === 1 || col === 2) && fila >= 4 && valor && String(valor).trim() !== "") {
            // Solo marcamos el check si no lo está
            if (hojaImp.getRange(filaImprimir, 1).getValue() !== true) {
              hojaImp.getRange(filaImprimir, 1).setValue(true);
            }
          }
        }
      }
    }

    // B2 EN REMITO: SELLAR Y PINTAR TEXTO DE BLANCO (A1:B1 y B2)
    if (celda === 'B2' && esTrue(valor)) {
      hoja.getRange("A1:B1").setFontColor("#ffffff");
      hoja.getRange("B2").setFontColor("#ffffff");
      hoja.getRange("B2").setValue(true);
      SpreadsheetApp.flush();
      hoja.toast("🎨 Remito sellado.", "ESTÉTICA");
      return;
    }

    // FORMATEAR CLIENTE (C1)
    if (celda === 'C1' && valor && valor !== "") {
      const nombreFormateado = valor.split(" ")
        .map(palabra => palabra.charAt(0).toUpperCase() + palabra.slice(1).toLowerCase())
        .join(" ");
      if (valor !== nombreFormateado) {
        hoja.getRange("C1").setValue(nombreFormateado);
      }
      return;
    }

   // [Dentro de function onEdit(e)]

// B1: GUARDAR REMITO + PINTAR TEXTO DE BLANCO (A1:B1 y B2)
    if (celda === 'B1' && esTrue(valor)) {
      try {
        const exito = enviarASalidas(hoja);
        
        if (exito) {
          const cliente = hoja.getRange('C1').getValue() || '(sin nombre)';
          enviarResumenDiscord(hoja, cliente);
          
          // Pintar texto blanco A1:B1 y B2
          hoja.getRange("A1:B1").setFontColor("#ffffff");
          hoja.getRange("B2").setFontColor("#ffffff");
          hoja.getRange("B1").setValue(true);
          hoja.getRange("B2").setValue(true);
          SpreadsheetApp.flush();
          
          hoja.toast("✅ Guardado en SALIDAS y remito sellado.", "ÉXITO", 5);
        }
      } catch (errorEnvio) {
        // NUEVO: Atrapamos el error y avisamos en pantalla
        hoja.getRange("B1").setValue(false); // Reseteamos el check
        hoja.toast("❌ " + errorEnvio.message, "ERROR DE ENVÍO", 8);
        Logger.log("Error B1: " + errorEnvio.message);
      }
      return;
    }

// ===============================================================
  // 📋 SEPARACIÓN DE CÓDIGO:CANTIDAD (FIX PARA NÚMEROS > 24)
  // ===============================================================
  // Solo actuamos si la edición es en la Columna A (1) y filas 4 a 22
  if (col === 1 && fila >= 4 && fila <= 22) {
    
    // IMPORTANTE: Usamos getDisplayValues para leer el texto real "28:20"
    const valoresVisibles = range.getDisplayValues(); 
    let huboCambio = false;

    for (let i = 0; i < valoresVisibles.length; i++) {
      let textoOriginal = valoresVisibles[i][0].trim();

      // Si el texto contiene el separador ":"
      if (textoOriginal.indexOf(':') !== -1) {
        let partes = textoOriginal.split(':');
        
        // Tomamos el primer grupo (código) y el segundo (cantidad)
        let codEncontrado = partes[0].trim();
        let cantEncontrada = partes[1].trim();

        if (codEncontrado !== "" && cantEncontrada !== "") {
          let filaDestino = fila + i;

          // 1. Limpiamos el formato de la celda para que deje de ser "Hora"
          hoja.getRange(filaDestino, 1).setNumberFormat("0");
          hoja.getRange(filaDestino, 2).setNumberFormat("0");

          // 2. Seteamos los valores reales
          hoja.getRange(filaDestino, 1).setValue(codEncontrado);
          hoja.getRange(filaDestino, 2).setValue(cantEncontrada);
          
          huboCambio = true;
        }
      }
    }
    
    if (huboCambio) {
      SpreadsheetApp.flush(); // Aplicamos cambios para que la validación de stock lea los datos nuevos
    }
 
      // 2. Lógica de Stock
      const celdaB     = hoja.getRange(fila, 2);
      const codigo     = hoja.getRange(fila, 1).getValue();
      const contenidoB = celdaB.getDisplayValue();

      if (codigo !== "" && contenidoB !== "") { 
        const stockReal = buscarStockReal(codigo);
        const pedido    = parseFloat(contenidoB.replace(',', '.'));
        const producto  = hoja.getRange(fila, 3).getValue() || String(codigo);

        if (stockReal !== -1 && !isNaN(pedido)) {
          if (pedido > stockReal) {
            celdaB.setFontColor("#ff0000"); // Rojo si falta
            enviarAExcedentes(nombreHoja, codigo, producto, pedido, stockReal, "");
            if (typeof enviarAlertaStockDiscordAzul === 'function') {
              enviarAlertaStockDiscordAzul(hoja, producto, contenidoB, stockReal);
            }
          } else {
            celdaB.setFontColor("#434343").setFontWeight("normal"); // Normal si hay stock
          }
        }
        SpreadsheetApp.flush();
      }
    } // <-- Cierra el if principal de (col 1 o 2)

  } catch (err) {
    Logger.log("Error en onEdit: " + err.message);
  } finally {
    procesando = false; // Libera el script para la siguiente edición
    // Si usas LockService, aquí iría: bloqueo.releaseLock();
  }
} // <-- ESTA ES LA LLAVE QUE CIERRA TODA LA FUNCIÓN onEdit


function enviarASalidas(hojaOrigen) {
  const startTime = new Date().getTime(); 

  // 1. Conexión Inteligente
  const ssActual = SpreadsheetApp.getActiveSpreadsheet();
  // Nota: Asegúrate de que ID_HOJA_SALIDAS esté definida globalmente
  let ssDestino = (ssActual.getId() === ID_HOJA_SALIDAS) ? ssActual : SpreadsheetApp.openById(ID_HOJA_SALIDAS);
  const hojaSalidas = ssDestino.getSheetByName('SALIDAS');
  
  if (!hojaSalidas) throw new Error("Pestaña 'SALIDAS' no encontrada.");

  // 2. Lectura en bloque (Bulk Read)
  const cliente = hojaOrigen.getRange('C1').getValue();
  if (!cliente) throw new Error("Falta el nombre del Cliente en C1");
  
  const rangoOrigen = hojaOrigen.getRange("A4:B22").getValues();

  // Filtrar filas vacías y preparar el bloque (Solo 3 columnas: Cliente, A y B)
  const filasParaEnviar = rangoOrigen
    .filter(fila => fila[0] !== "" && fila[1] !== "")
    .map(fila => [cliente, fila[0], fila[1]]); // <-- ELIMINADO: new Date()

  if (filasParaEnviar.length === 0) return false;

  // 3. Escritura en bloque (Bulk Write)
  const ultimaFila = hojaSalidas.getLastRow();
  
  // Ajustado a 3 columnas de ancho (A, B, C)
  hojaSalidas.getRange(ultimaFila + 1, 1, filasParaEnviar.length, 3)
    .setValues(filasParaEnviar);

  console.log("Tiempo de ejecución: " + (new Date().getTime() - startTime) + "ms");
  return true;
}

// ===============================================================
// 📦  ENVIAR RESUMEN A DISCORD
// ===============================================================
function enviarResumenDiscord(hoja, cliente) {
  const productos = hoja.getRange("A4:B22").getValues().filter(f => f[0] !== "");
  let lista = "";
  productos.forEach(p => { lista += `📦 **${p[0]}**: ${p[1]} unidades\n`; });

  const payload = {
    content: `:package: **Nuevo Remito: ${cliente}**`,
    embeds: [{
      title: "Detalle de Salida",
      description: lista || "(sin productos)",
      color: 3066993
    }]
  };

  UrlFetchApp.fetch(WEBHOOK_REMITOS, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}


// ===============================================================
// 🚨  ALERTA DE STOCK
// ===============================================================
function enviarAlertaStock(sheet) {
  const cliente        = sheet.getRange("C1").getValue() || "(sin nombre)";
  const ultFila        = sheet.getLastRow();
  if (ultFila < 4) return;

  const spreadsheetUrl  = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const remitoPestanaUrl = spreadsheetUrl + "#gid=" + sheet.getSheetId();
  const datosPedido      = sheet.getRange(4, 1, ultFila - 3, 2).getValues();

  const mapaInv    = new Map();
  const mapaNombres = new Map();

  try {
    const inv     = SpreadsheetApp.openById(INVENTARIO_ID).getSheetByName("INVENTARIO");
    if (!inv) return;
    const invData = inv.getRange(2, 1, inv.getLastRow() - 1, 5).getValues();
    invData.forEach(f => {
      const cod = String(f[0]).trim();
      if (cod) {
        mapaInv.set(cod,     Number(f[4]) || 0);
        mapaNombres.set(cod, String(f[1]).trim());
      }
    });
  } catch (err) {
    Logger.log("❌ Error leyendo inventario: " + err.message);
    return;
  }

  const faltantes = [];
  let primerCodigo = null;

  datosPedido.forEach(row => {
    const codigo     = String(row[0]).trim();
    const cantPedida = Number(row[1]);
    if (!codigo || isNaN(cantPedida) || cantPedida <= 0) return;

    const stockReal = mapaInv.get(codigo) || 0;
    if (cantPedida > stockReal) {
      faltantes.push(
        `📦 **${codigo}** (${mapaNombres.get(codigo) || "Desconocido"})\n` +
        `   Pedido: ${cantPedida} | Stock: ${stockReal} | **Falta: ${cantPedida - stockReal}**`
      );
      if (!primerCodigo) primerCodigo = codigo;
    }
  });

  if (faltantes.length === 0) return;

  let textoEnlace = `Revisar el [Inventario General](https://docs.google.com/spreadsheets/d/${INVENTARIO_ID}/edit)`;
  if (primerCodigo) {
    const res = buscarYGenerarEnlaceInventario(primerCodigo);
    if (res) textoEnlace = `Ir al [producto en inventario](${res.url})`;
  }

  const msg =
    `🚨 **ALERTA DE STOCK BAJO**\n` +
    `👤 **Cliente:** ${cliente}\n` +
    `📄 **Hoja:** [${sheet.getName()}](${remitoPestanaUrl})\n\n` +
    `⚠️ **Productos con faltante:**\n` +
    faltantes.join("\n\n") +
    `\n\n🔗 ${textoEnlace}`;

  enviarMensajeDiscord(msg, WEBHOOK_STOCK);
}


// ===============================================================
// 🔔  ALERTA DE STOCK EXCEDIDO
// ===============================================================
function enviarAlertaStockDiscordAzul(hoja, producto, pedido, real) {
  const nombreHoja = hoja.getName();
  const cliente = hoja.getRange("C1").getValue() || "Sin nombre";
  const ahora = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm:ss");

  let colorAlerta = 3447003;
  if (nombreHoja.startsWith("C-AL")) {
    colorAlerta = 15158332;
  } else if (nombreHoja.startsWith("C-SM")) {
    colorAlerta = 16776960;
  } else if (nombreHoja.startsWith("C-SL")) {
    colorAlerta = 3066993;
  }

  const embed = {
    title: "🔹 ALERTA: STOCK EXCEDIDO",
    description: "El pedido supera el stock disponible en inventario.",
    color: colorAlerta,
    fields: [
      { name: "👤 Cliente",      value: String(cliente),        inline: true  },
      { name: "📋 Hoja",         value: String(nombreHoja),     inline: true  },
      { name: "📦 Producto",     value: String(producto),       inline: false },
      { name: "📊 Stock",        value: `Pedido: ${pedido} | Hay: ${real}`, inline: false },
      { name: "⏰ Fecha y Hora", value: ahora,                  inline: false }
    ],
    footer: { text: "Control de Stock Automático" }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ embeds: [embed] }),
    muteHttpExceptions: true
  };

  UrlFetchApp.fetch(WEBHOOK_STOCK, options);
}


// ===============================================================
// 📤  ENVIAR A EXCEDENTES
// ===============================================================
function enviarAExcedentes(nombrePestana, codigo, producto, cantPedida, stockReal, notaParaInsertar) {
  const ID_EXTERNO = "10sIFWNX3UtOHJD_26isE6n_lAH5osU7Y6NWVHxrUUmU";
  try {
    const ssExterno = SpreadsheetApp.openById(ID_EXTERNO);
    let hojaExcedentes = ssExterno.getSheetByName("EXCEDENTES");
    
    hojaExcedentes.insertRowBefore(2);
    
    const ssActual = SpreadsheetApp.getActiveSpreadsheet();
    const cliente = ssActual.getSheetByName(nombrePestana).getRange("C1").getValue() || "Sin Nombre";
    
    const fila = [
      "",
      nombrePestana,
      cliente,
      codigo,
      cantPedida,
      notaParaInsertar,
      producto,
      new Date()
    ];

    hojaExcedentes.getRange(2, 1, 1, 8).setValues([fila]);
    hojaExcedentes.getRange(2, 8).setNumberFormat("dd/mm/yyyy HH:mm:ss");
    
  } catch (e) {
    Logger.log("Error en envío: " + e);
  }
}


// ===============================================================
// ✉️  ENVIAR MENSAJE DISCORD
// ===============================================================
function enviarMensajeDiscord(mensaje, webhook) {
  UrlFetchApp.fetch(webhook, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ content: mensaje }),
    muteHttpExceptions: true
  });
}


// ===============================================================
// 👁️  MOSTRAR / OCULTAR HOJAS
// ===============================================================
function mostrarHojasSeleccionadas() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const hojaImprimir = obtenerHojaImprimir();
  if (!hojaImprimir) return;

  const datos = hojaImprimir.getRange("A2:B50").getValues();
  const mostrar = [];
  const ocultar = [];

  datos.forEach(fila => {
    const check  = esTrue(fila[0]);
    const nombre = String(fila[1]).trim();
    if (!nombre) return;
    if (check) mostrar.push(nombre);
    else       ocultar.push(nombre);
  });

  ss.getSheets().forEach(hoja => {
    const nombre = hoja.getName();
    if (nombre === hojaImprimir.getName()) return;
    if (!NOMBRES_CONTROLADOS.includes(nombre)) return;

    if (mostrar.includes(nombre))      hoja.showSheet();
    else if (ocultar.includes(nombre)) hoja.hideSheet();
  });
}


// ===============================================================
// 🗑️  RESTAURAR NOMBRES EN IMPRIMIR
// ===============================================================
function restaurarNombreEnImprimir(nombreHoja) {
  const hojaImprimir = obtenerHojaImprimir();
  if (!hojaImprimir) return;

  MAPEO_FILAS.forEach(item => {
    const valorB = String(hojaImprimir.getRange(item.fila, 2).getValue()).trim();
    if (valorB === nombreHoja) {
      hojaImprimir.getRange(item.fila, 2).setValue(item.nombre);
      hojaImprimir.getRange(item.fila, 1).setValue(false);
      hojaImprimir.getRange(item.fila, 3).setValue(false);
      Logger.log(`✅ IMPRIMIR fila ${item.fila}: "${nombreHoja}" → "${item.nombre}"`);
    }
  });
}


// ===============================================================
// 🔲  DESMARCAR TODOS LOS B2
// ===============================================================
function desmarcarTodosLosB2() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const excluidas = ['REMITO AL', 'REMITO SL', 'REMITO SM', 'CLIENTES', 'SALIDAS', 'INVENTARIO',
                     obtenerHojaImprimir()?.getName()];
  let contador = 0;

  ss.getSheets().forEach(h => {
    const nombre = h.getName();
    if (excluidas.includes(nombre) || nombre.toLowerCase().includes('remito')) return;
    try {
      if (esTrue(h.getRange('B2').getValue())) {
        h.getRange('B2').setValue(false);
        contador++;
      }
    } catch (err) {}
  });

  Logger.log(`✅ ${contador} celdas B2 desmarcadas`);
}


// ===============================================================
// 🔎  OBTENER HOJA IMPRIMIR
// ===============================================================
function obtenerHojaImprimir() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const nombre of ['IMPRIMIR', 'Imprimir', 'imprimir']) {
    const h = ss.getSheetByName(nombre);
    if (h) return h;
  }
  return null;
}


// ===============================================================
// 📋  SINCRONIZAR LISTA DE HOJAS
// ===============================================================
function obtenerListaHojasParaImprimir() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const imprimir = obtenerHojaImprimir();
  if (!imprimir) return;

  MAPEO_FILAS.forEach(item => {
    const hoja = ss.getSheetByName(item.nombre);
    if (hoja) {
      imprimir.getRange(item.fila, 2).setValue(hoja.getName());
      Logger.log(`✅ IMPRIMIR fila ${item.fila}: "${hoja.getName()}"`);
    }
  });
}


// ===============================================================
// 🔗  BUSCAR ENLACE EN INVENTARIO
// ===============================================================
function buscarYGenerarEnlaceInventario(codigoProducto) {
  try {
    const ssInv    = SpreadsheetApp.openById(INVENTARIO_ID);
    const codigoStr = String(codigoProducto).trim();

    for (const sheet of ssInv.getSheets()) {
      const lastRow = sheet.getLastRow();
      if (lastRow < 1) continue;
      const data = sheet.getRange(1, 1, lastRow, 1).getValues();
      for (let r = 0; r < data.length; r++) {
        if (String(data[r][0]).trim() === codigoStr) {
          return {
            url: ssInv.getUrl() + `#gid=${sheet.getSheetId()}&range=A${r + 1}`,
            sheetName: sheet.getName()
          };
        }
      }
    }
  } catch (err) {
    Logger.log("❌ buscarYGenerarEnlaceInventario: " + err.message);
  }
  return null;
}


// ===============================================================
// 🔍  BUSCAR STOCK REAL
// ===============================================================
function buscarStockReal(codigo) {
  if (!codigo) return -1;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaInv = ss.getSheetByName("INVENTARIO");
    if (!hojaInv) return -1;
    const datos = hojaInv.getDataRange().getValues();
    const busca = String(codigo).trim().toUpperCase();
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][0]).trim().toUpperCase() === busca) return Number(datos[i][4]) || 0;
    }
    return -1;
  } catch (err) {
    return -1;
  }
}


// ===============================================================
// 🎨  COLOREAR STOCK
// ===============================================================
function verificarYColorearStock(sheet) {
  if (!sheet) return;
  if (/^C-AL[1-7]/.test(sheet.getName())) return;

  const stockMap = new Map();
  try {
    const inv     = SpreadsheetApp.openById(INVENTARIO_ID).getSheetByName("INVENTARIO");
    const invData = inv.getRange(2, 1, inv.getLastRow() - 1, 5).getValues();
    invData.forEach(row => {
      if (row[0]) stockMap.set(String(row[0]).trim(), Number(row[4]) || 0);
    });
  } catch (err) { return; }

  const ultFila = sheet.getLastRow();
  if (ultFila < 4) return;

  const datos   = sheet.getRange(4, 1, ultFila - 3, 2).getValues();
  const colores = datos.map(row => {
    const cod      = String(row[0]).trim();
    const cantidad = Number(row[1]);
    const stock    = stockMap.get(cod) || 0;
    return [cantidad > stock ? COLOR_FALTANTE : COLOR_NORMAL];
  });

  sheet.getRange(4, 2, colores.length, 1).setFontColors(colores);
}


// ===============================================================
// 🏗️  VERIFICAR Y CREAR HOJAS
// ===============================================================
function verificarHojasCopia() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Verificar hojas principales
  ['RESUMEN DS', 'IMPRIMIR'].forEach(nombre => {
    if (!ss.getSheetByName(nombre)) {
      ss.insertSheet(nombre);
      Logger.log(`⚠️ Hoja recreada: ${nombre}`);
    }
  });

  const grupos = {
    'REMITO SM': ['C-SM1','C-SM2','C-SM3','C-SM4','C-SM5'],
    'REMITO AL': ['C-AL1','C-AL2','C-AL3','C-AL4','C-AL5','C-AL6','C-AL7'],
    'REMITO SL': ['C-SL1','C-SL2','C-SL3','C-SL4','C-SL5','C-SL6']
  };

  // Obtener nombres actuales para no repetir
  let existentes = ss.getSheets().map(s => s.getName());

  for (const [plantilla, nombres] of Object.entries(grupos)) {
    const hojaBase = ss.getSheetByName(plantilla);
    if (!hojaBase) {
      Logger.log(`❌ No se encontró la plantilla: ${plantilla}`);
      continue;
    }

    nombres.forEach(n => {
      if (!existentes.includes(n)) {
        try {
          // Copiar la hoja
          const nueva = hojaBase.copyTo(ss).setName(n);
          
          // Limpiar contenidos (A4:B22 son datos, C1 es nombre cliente)
          nueva.getRangeList(['A4:B22', 'C1']).clearContent();
          
          // Configurar celdas de control (A1, B1, B2)
          ['A1', 'B1', 'B2'].forEach(celda => {
            const rango = nueva.getRange(celda);
            rango.clearDataValidations(); // Borra cualquier regla previa
            rango.insertCheckboxes();    // Inserta el checkbox limpio
            rango.setValue(false);       // Lo pone en falso
          });

          // Estética: Si es una hoja nueva, asegurar que el texto sea visible
          nueva.getRange("A1:B2").setFontColor("#000000"); 

          Logger.log(`🟢 Hoja creada exitosamente: ${n}`);
          existentes.push(n); // Actualizar lista local para evitar duplicados en el mismo loop
        } catch (e) {
          Logger.log(`Error al crear la hoja ${n}: ${e.message}`);
        }
      }
    });
  }
}


// ===============================================================
// 📂  ORDENAR HOJAS POR PRIORIDAD
// ===============================================================
function ordenarHojasPrioridad() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  function peso(nombre) {
    if (nombre.toUpperCase() === "IMPRIMIR") return 0;
    if (nombre.startsWith("C-SM")) return 1;
    if (nombre.startsWith("C-AL")) return 2;
    if (nombre.startsWith("C-SL")) return 3;
    return 4;
  }

  sheets.sort((a, b) => {
    const pa = peso(a.getName()), pb = peso(b.getName());
    if (pa !== pb) return pa - pb;
    return a.getName().localeCompare(b.getName(), undefined, { numeric: true, sensitivity: 'base' });
  });

  sheets.forEach((h, i) => {
    ss.setActiveSheet(h);
    ss.moveActiveSheet(i + 1);
  });

  const imp = ss.getSheetByName("IMPRIMIR");
  if (imp && !imp.isSheetHidden()) ss.setActiveSheet(imp);
}


// ===============================================================
// 🔧  CORREGIR NOMBRES EN A2
// ===============================================================
function corregirNombresEnA2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  MAPEO_FILAS.forEach(item => {
    const hoja = ss.getSheetByName(item.nombre);
    if (hoja) {
      hoja.getRange('A2').setValue(item.nombre);
      Logger.log(`✅ ${item.nombre}: A2 = "${item.nombre}"`);
    }
  });
  Logger.log('✅ corregirNombresEnA2 completado');
}


// ===============================================================
// 📋  SEPARAR CÓDIGO:CANTIDAD AL PEGAR
// ===============================================================
function separarCodigosAlPegar(e) {
  const range        = e.range;
  const valorOriginal = range.getValue();
  let texto = "";

  if (valorOriginal instanceof Date) {
    texto = valorOriginal.getHours() + ":" + valorOriginal.getMinutes();
  } else {
    texto = String(valorOriginal).trim();
  }

  if (texto.indexOf(':') === -1) return;

  const partes  = texto.split(':');
  const cod     = partes[0].trim();
  const cant    = partes[1].trim();
  const codNum  = (isNaN(cod)  || cod  === "") ? cod  : Number(cod);
  const cantNum = (isNaN(cant) || cant === "") ? cant : Number(cant);

  range.setNumberFormat("@").setValue(codNum).setNumberFormat("0");
  range.offset(0, 1).setValue(cantNum).setNumberFormat("0");
}


// ===============================================================
// 🔄  SINCRONIZAR NOTAS CON EXCEDENTES (CRON)
// ===============================================================
function cronSincronizarNotas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojas = ss.getSheets();
  const fechaHoy = new Date().toLocaleDateString();
  const ID_EXTERNO = "10sIFWNX3UtOHJD_26isE6n_lAH5osU7Y6NWVHxrUUmU";
  
  try {
    const ssExterno = SpreadsheetApp.openById(ID_EXTERNO);
    let hojaExcedentes = ssExterno.getSheetByName("EXCEDENTES");
    if (!hojaExcedentes) return;

    const datosExcedentes = hojaExcedentes.getDataRange().getValues();

    hojas.forEach(hoja => {
      const nombreHoja = hoja.getName();
      
      if (nombreHoja.indexOf("C-") === 0) {
        const ultimaFila = Math.min(hoja.getLastRow(), 100); 
        if (ultimaFila < 4) return;

        const rangoDatos = hoja.getRange(4, 1, ultimaFila - 3, 3).getValues();
        const notas = hoja.getRange(4, 2, ultimaFila - 3, 1).getNotes();

        for (let i = 0; i < notas.length; i++) {
          let notaActual = notas[i][0];
          if (!notaActual || notaActual.trim() === "") continue;

          let codigoArt = rangoDatos[i][0];
          let cantIngresada = rangoDatos[i][1];
          let nombreProd = rangoDatos[i][2];

          let filaEncontradaIdx = -1;

          for (let j = 1; j < datosExcedentes.length; j++) {
            let fechaFila = datosExcedentes[j][7] instanceof Date ? 
                            datosExcedentes[j][7].toLocaleDateString() : "";
            
            if (datosExcedentes[j][1] === nombreHoja && 
                datosExcedentes[j][3].toString() === codigoArt.toString() && 
                fechaFila === fechaHoy) {
              filaEncontradaIdx = j + 1;
              break;
            }
          }

          if (filaEncontradaIdx !== -1) {
            hojaExcedentes.getRange(filaEncontradaIdx, 6).setValue(notaActual);
            hojaExcedentes.getRange(filaEncontradaIdx, 7).setValue(nombreProd);
          } else {
            enviarAExcedentes(nombreHoja, codigoArt, nombreProd, cantIngresada, 0, notaActual);
            datosExcedentes.push(["", nombreHoja, "", codigoArt, cantIngresada, notaActual, nombreProd, new Date()]);
          }
        }
      }
    });
  } catch (e) {
    Logger.log("Error en CRON: " + e);
  }
}

/* Actualiza el día en D1. Si es Domingo, salta a Lunes.
 */
function actualizarDiaEnRemitos() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hojas = libro.getSheets();
  
  const fechaActual = new Date();
  let numeroDia = fechaActual.getDay(); // 0 es Domingo, 6 es Sábado
  
  let diaSemana;

  // REGLA: Si es Domingo (0), poner Lunes.
  if (numeroDia === 0) {
    diaSemana = "Lunes";
  } else {
    const opciones = { weekday: 'long' };
    diaSemana = new Intl.DateTimeFormat('es-ES', opciones).format(fechaActual);
    // Capitalizar (lunes -> Lunes)
    diaSemana = diaSemana.charAt(0).toUpperCase() + diaSemana.slice(1);
  }

  hojas.forEach(hoja => {
    const nombre = hoja.getName();
    if (nombre.startsWith("C-AL") || nombre.startsWith("C-SM") || nombre.startsWith("C-SL")) {
      hoja.getRange("D1").setValue(diaSemana);
    }
  });
  
  console.log("Día actualizado: " + diaSemana);
}

function testEnvioManual() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getActiveSheet();
  
  console.log("Hoja: " + hoja.getName());
  console.log("Cliente (C1): " + hoja.getRange("C1").getValue());
  console.log("ID archivo: " + ss.getId());
  console.log("ID_HOJA_SALIDAS: " + ID_HOJA_SALIDAS);
  
  const hojaSalidas = ss.getSheetByName("SALIDAS");
  console.log("Pestaña SALIDAS existe: " + (hojaSalidas ? "SÍ" : "NO"));
  
  try {
    const resultado = enviarASalidas(hoja);
    console.log("Resultado: " + resultado);
  } catch(e) {
    console.error("ERROR: " + e.message);
  }
}
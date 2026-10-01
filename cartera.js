/**
 * cartera.gs - Cálculo automático de cartera cripto
 *
 * Reemplaza las fórmulas de F:I y la tabla de "Posibles ventas".
 * Vos cargás A:E a mano, el resto lo escribe el script como valores.
 *
 * Instalación:
 *   1. Extensiones > Apps Script
 *   2. Pegar este archivo completo (reemplazando lo que haya)
 *   3. Guardar (Ctrl+S)
 *   4. Recargar la planilla -> aparece el menú "Cartera"
 *   5. Cartera > Recalcular todo
 */

// ============================================================
//  CONFIGURACIÓN
// ============================================================
const CFG = {
  HOJA: 'Criptos',

  FILA_INICIO: 2,

  COL_FECHA:    1,
  COL_CRIPTO:   2,
  COL_TIPO:     3,
  COL_CANTIDAD: 4,
  COL_PRECIO:   5,

  COL_USD:      6,
  COL_ACUM:     7,
  COL_PROM:     8,
  COL_GAN:      9,

  TOL: 1e-9,

  HOJA_REBAL: 'Rebalanceo',
  HOJA_HISTORIAL: 'Historial',   // no se limpia nunca, solo se le suman filas

  CASILLEROS: [
    { nombre: 'USDT (pólvora seca)', pct: 20, monedas: [] },
    { nombre: 'Núcleo',              pct: 50, monedas: ['SOL', 'LINK'] },
    { nombre: 'Satélite',            pct: 20, monedas: ['NEAR', 'INJ'] },
    { nombre: 'Apuestas',            pct: 10, monedas: ['*'] }
  ],

  BANDA_PP:      5,
  MIN_OPERACION: 25,

  // Recomprar cuando el precio cae este % o más desde el precio de la ULTIMA venta
  // registrada de esa moneda (la foto la toma sola de la columna E de la fila de venta)
  UMBRAL_RECOMPRA: 0.30,

  IGNORAR_EN_RESUMEN: ['USDT', 'USDC', 'DAI'],

  COINGECKO: {
    SOL: 'solana',      LINK: 'chainlink',  NEAR: 'near',
    INJ: 'injective-protocol',              ADA: 'cardano',
    ARB: 'arbitrum',    DOGE: 'dogecoin',   PEPE: 'pepe',
    BTC: 'bitcoin',     ETH: 'ethereum',    RENDER: 'render-token',
    VET: 'vechain',     USUAL: 'usual',     USDT: 'tether',
    XNO: 'nano'
  },

  EMAIL: {
    destinatario: 'alonso.nahuel.2002@gmail.com',
    diaDelMes:    1
  }
};


// ============================================================
//  MENÚ Y TRIGGERS
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cartera')
    .addItem('Recalcular todo', 'recalcular')
    .addItem('Validar datos', 'validar')
    .addSeparator()
    .addItem('Rebalancear (trae precios)', 'rebalancear')
    .addItem('Enviar resumen por mail ahora', 'enviarResumenEmail')
    .addSeparator()
    .addItem('Instalar automatización', 'instalarTriggers')
    .addItem('Quitar automatización', 'borrarTriggers')
    .addSeparator()
    .addItem('Diagnóstico', 'diagnostico')
    .addItem('Probar fuentes de precios', 'probarPrecios')
    .addToUi();
}

function alEditar(e) {
  if (!e || !e.range) return;
  if (e.range.getSheet().getName() !== CFG.HOJA) return;

  const c1 = e.range.getColumn();
  const c2 = c1 + e.range.getNumColumns() - 1;
  if (c2 < CFG.COL_FECHA || c1 > CFG.COL_PRECIO) return;
  if (e.range.getRow() + e.range.getNumRows() - 1 < CFG.FILA_INICIO) return;

  recalcular();
}

function instalarTriggers() {
  borrarTriggers();

  const ss = SpreadsheetApp.getActive();

  ScriptApp.newTrigger('alEditar')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  ScriptApp.newTrigger('recalcular')
    .timeBased()
    .everyMinutes(15)
    .create();

  ScriptApp.newTrigger('enviarResumenEmail')
    .timeBased()
    .onMonthDay(CFG.EMAIL.diaDelMes)
    .atHour(9)
    .create();

  const msg = 'Listo. Se instalaron 3 automatizaciones:\n\n'
            + '- Al editar A:E desde la web -> recalcula al instante\n'
            + '- Cada 15 minutos -> cubre las cargas desde el celular\n'
            + '- Dia ' + CFG.EMAIL.diaDelMes + ' de cada mes, 9 AM -> mail con el resumen';
  try { SpreadsheetApp.getUi().alert(msg); } catch (err) { Logger.log(msg); }
}

function borrarTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (t) { ScriptApp.deleteTrigger(t); });
  Logger.log('Triggers borrados: ' + triggers.length);
}

function diagnostico() {
  const ss = SpreadsheetApp.getActive();
  const hoja = ss.getSheetByName(CFG.HOJA);
  const triggers = ScriptApp.getProjectTriggers();

  var txt = 'Planilla: ' + ss.getName() + '\n'
          + 'Hoja "' + CFG.HOJA + '": ' + (hoja ? 'OK' : 'NO EXISTE') + '\n'
          + 'Hojas disponibles: ' + ss.getSheets().map(function (h) { return h.getName(); }).join(', ') + '\n\n'
          + 'Triggers instalados: ' + triggers.length + '\n';

  triggers.forEach(function (t) {
    txt += '  - ' + t.getHandlerFunction() + ' (' + t.getEventType() + ')\n';
  });

  if (hoja) {
    txt += '\nUltima fila con datos: ' + hoja.getLastRow();
  }

  try { SpreadsheetApp.getUi().alert(txt); } catch (err) { Logger.log(txt); }
  return txt;
}


// ============================================================
//  PRINCIPAL
// ============================================================

function recalcular() {
  const hoja = SpreadsheetApp.getActive().getSheetByName(CFG.HOJA);
  if (!hoja) throw new Error('No existe la hoja "' + CFG.HOJA + '"');

  const movs = leerMovimientos(hoja);
  procesar(movs);

  escribirFilas(hoja, movs);

  const errores = movs.filter(function (m) { return m.error; }).length;
  SpreadsheetApp.getActive().toast(
    movs.length + ' movimientos' + (errores ? ' - ' + errores + ' con problemas' : ' - OK'),
    'Cartera', 4
  );
}

function leerMovimientos(hoja) {
  const ultima = hoja.getLastRow();
  if (ultima < CFG.FILA_INICIO) return [];

  const n = ultima - CFG.FILA_INICIO + 1;
  const datos = hoja.getRange(CFG.FILA_INICIO, CFG.COL_FECHA, n, CFG.COL_PRECIO).getValues();

  const movs = [];
  datos.forEach(function (fila, i) {
    const cripto = String(fila[CFG.COL_CRIPTO - 1] || '').trim().toUpperCase();
    if (!cripto) return;

    movs.push({
      hoja_fila: CFG.FILA_INICIO + i,
      fecha:     aFecha(fila[CFG.COL_FECHA - 1]),
      cripto:    cripto,
      tipo:      String(fila[CFG.COL_TIPO - 1] || '').trim().toUpperCase(),
      cantidad:  aNumero(fila[CFG.COL_CANTIDAD - 1]),
      precio:    aNumero(fila[CFG.COL_PRECIO   - 1]),
      error:     ''
    });
  });

  return movs;
}

function procesar(movs) {
  const pos = {};
  const orden = [];
  var ultimaFecha = 0;

  movs.forEach(function (m) {
    m.usd = m.cantidad * m.precio;

    if (!pos[m.cripto]) {
      pos[m.cripto] = { cantidad: 0, costo: 0, realizado: 0 };
      orden.push(m.cripto);
    }
    const p = pos[m.cripto];

    if (m.fecha && ultimaFecha && m.fecha < ultimaFecha) {
      m.error = 'Fecha anterior a la fila de arriba (el promedio se calcula en orden de fila)';
    }
    if (m.fecha) ultimaFecha = m.fecha;

    if (m.tipo === 'COMPRA') {
      p.cantidad += m.cantidad;
      p.costo    += m.usd;
      m.ganancia  = 0;

    } else if (m.tipo === 'VENTA') {
      const promPrevio = p.cantidad > CFG.TOL ? p.costo / p.cantidad : 0;

      if (m.cantidad > p.cantidad + CFG.TOL) {
        m.error = 'Vendes ' + m.cantidad + ' pero tenes ' + p.cantidad;
        m.acum = p.cantidad;
        m.prom = promPrevio;
        m.ganancia = 0;
        return;
      }

      m.ganancia   = (m.precio - promPrevio) * m.cantidad;
      p.cantidad  -= m.cantidad;
      p.costo     -= m.cantidad * promPrevio;
      p.realizado += m.ganancia;

    } else if (m.tipo === '') {
      m.ganancia = 0;
    } else {
      m.error = 'Tipo desconocido: "' + m.tipo + '"';
      m.ganancia = 0;
    }

    if (Math.abs(p.cantidad) < CFG.TOL) {
      p.cantidad = 0;
      p.costo    = 0;
    }

    m.acum = p.cantidad;
    m.prom = p.cantidad > CFG.TOL ? p.costo / p.cantidad : 0;
  });

  return { pos: pos, orden: orden };
}

function escribirFilas(hoja, movs) {
  const desde = CFG.FILA_INICIO;
  const hasta = hoja.getLastRow();
  if (hasta < desde) return;

  const porFila = {};
  movs.forEach(function (m) { porFila[m.hoja_fila] = m; });

  const salida = [];
  const fondos = [];
  for (var f = desde; f <= hasta; f++) {
    const m = porFila[f];
    if (m) {
      salida.push([m.usd, m.acum, m.prom, m.ganancia]);
      fondos.push([m.error ? '#f4c7c3' : null]);
    } else {
      salida.push(['', '', '', '']);
      fondos.push([null]);
    }
  }

  hoja.getRange(desde, CFG.COL_USD, salida.length, 4).setValues(salida);
  hoja.getRange(desde, CFG.COL_TIPO, fondos.length, 1).setBackgrounds(fondos);
}

function validar() {
  const hoja = SpreadsheetApp.getActive().getSheetByName(CFG.HOJA);
  if (!hoja) throw new Error('No existe la hoja "' + CFG.HOJA + '"');
  const movs = leerMovimientos(hoja);
  procesar(movs);

  const problemas = movs
    .filter(function (m) { return m.error; })
    .map(function (m) { return 'Fila ' + m.hoja_fila + ' (' + m.cripto + '): ' + m.error; });

  SpreadsheetApp.getUi().alert(
    problemas.length
      ? problemas.join('\n')
      : 'Todo OK - ' + movs.length + ' movimientos procesados.'
  );
}

// ============================================================
//  HELPERS
// ============================================================

function aNumero(v) {
  if (typeof v === 'number') return v;
  var s = String(v || '').trim();
  if (!s) return 0;
  if (s.indexOf(',') > -1 && s.indexOf('.') > -1) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.indexOf(',') > -1) {
    s = s.replace(',', '.');
  }
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function aFecha(v) {
  if (v instanceof Date) return v.getTime();
  var s = String(v || '').trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]).getTime();
  var d = new Date(s);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}


// ============================================================
//  REBALANCEO
// ============================================================

/**
 * Prueba los proveedores EN CASCADA, combinando resultados: lo que un
 * proveedor no consigue, lo intenta el siguiente solo para lo que falta.
 * Nunca tira error: devuelve lo que haya podido conseguir, aunque sea
 * parcial. El llamador decide si lo que falta es crítico o no (por ejemplo,
 * el precio de una moneda vendida hace tiempo y poco líquida como USUAL no
 * debería tumbar todo el rebalanceo de las posiciones que sí están abiertas).
 */
function traerPrecios(monedas) {
  if (!monedas.length) return {};

  var precios = {};
  var pendientes = monedas.slice();
  const proveedores = [precios_Coinbase, precios_CoinGecko];

  for (var i = 0; i < proveedores.length && pendientes.length; i++) {
    try {
      const p = proveedores[i](pendientes);
      Object.keys(p).forEach(function (m) { precios[m] = p[m]; });
      pendientes = pendientes.filter(function (m) { return !precios[m]; });
    } catch (err) {
      // Este proveedor falló entero (ej: 429, red caída). Seguimos con el resto.
    }
  }

  return precios;   // puede venir incompleto
}

function precios_Coinbase(monedas) {
  const precios = {};
  monedas.forEach(function (m) {
    try {
      const resp = UrlFetchApp.fetch(
        'https://api.coinbase.com/v2/prices/' + m + '-USD/spot',
        { muteHttpExceptions: true }
      );
      if (resp.getResponseCode() !== 200) return;
      const d = JSON.parse(resp.getContentText());
      if (d && d.data && d.data.amount) precios[m] = parseFloat(d.data.amount);
    } catch (err) { }
  });
  return precios;
}

function precios_CoinGecko(monedas) {
  const ids = [];
  const sinId = [];
  monedas.forEach(function (m) {
    const id = CFG.COINGECKO[m];
    if (id) ids.push(id); else sinId.push(m);
  });

  if (sinId.length) {
    throw new Error('Falta el id de CoinGecko para: ' + sinId.join(', ') +
                    '. Agregalo a CFG.COINGECKO.');
  }

  const url = 'https://api.coingecko.com/api/v3/simple/price'
            + '?ids=' + encodeURIComponent(ids.join(','))
            + '&vs_currencies=usd';

  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    throw new Error('HTTP ' + resp.getResponseCode());
  }

  const datos = JSON.parse(resp.getContentText());
  const precios = {};
  monedas.forEach(function (m) {
    const d = datos[CFG.COINGECKO[m]];
    if (d && d.usd) precios[m] = d.usd;
  });
  return precios;
}

function probarPrecios() {
  const hoja = SpreadsheetApp.getActive().getSheetByName(CFG.HOJA);
  const estado = procesar(leerMovimientos(hoja));
  const monedas = estado.orden.filter(function (c) {
    return CFG.IGNORAR_EN_RESUMEN.indexOf(c) === -1 && estado.pos[c].cantidad > CFG.TOL;
  });

  var txt = 'Monedas a cotizar: ' + monedas.join(', ') + '\n\n';

  [precios_Coinbase, precios_CoinGecko].forEach(function (fn) {
    try {
      const p = fn(monedas);
      const ok = monedas.filter(function (m) { return p[m]; });
      txt += fn.name + ': ' + ok.length + '/' + monedas.length + ' OK\n';
      ok.forEach(function (m) { txt += '   ' + m + ' = ' + p[m] + '\n'; });
      const falta = monedas.filter(function (m) { return !p[m]; });
      if (falta.length) txt += '   FALTAN: ' + falta.join(', ') + '\n';
    } catch (err) {
      txt += fn.name + ': ERROR -> ' + err + '\n';
    }
    txt += '\n';
  });

  try { SpreadsheetApp.getUi().alert(txt); } catch (e) { Logger.log(txt); }
  return txt;
}

/**
 * "Foto" de venta por moneda: precio y fecha de la ULTIMA venta registrada
 * de cada cripto. No se guarda en ningún lado aparte: se deriva del propio
 * log cada vez que se corre. Si vendés la misma moneda de nuevo, la foto
 * vieja se pisa sola con la nueva.
 */
function ultimaVenta(movs) {
  const ventas = {};
  movs.forEach(function (m) {
    if (m.tipo === 'VENTA' && !m.error && m.precio > 0) {
      ventas[m.cripto] = { precio: m.precio, fecha: m.fecha };
    }
  });
  return ventas;
}


/**
 * Compara el precio de hoy contra la foto de venta y arma la lista de
 * candidatos a recompra (caída >= CFG.UMBRAL_RECOMPRA desde esa foto).
 *
 * Ojo: que caiga 30% desde el precio de VENTA no significa que este por
 * debajo de tu PROMEDIO de costo (si vendiste con mucha ganancia, puede
 * seguir estando arriba de lo que pagaste vos). Por eso se agrega
 * bajaPromedio, que compara contra el promedio vigente de la posicion.
 */
function candidatosRecompra(movs, precios, pos) {
  const ventas = ultimaVenta(movs);
  const monedas = Object.keys(ventas).filter(function (m) { return precios[m]; });

  return monedas.map(function (m) {
    const precioVenta = ventas[m].precio;
    const precioHoy = precios[m];
    const caida = 1 - (precioHoy / precioVenta);

    const p = pos[m];
    const promedioActual = (p && p.cantidad > CFG.TOL) ? p.costo / p.cantidad : null;
    const bajaPromedio = promedioActual !== null ? precioHoy < promedioActual : null;

    return {
      moneda: m,
      precioVenta: precioVenta,
      precioHoy: precioHoy,
      caida: caida,
      candidato: caida >= CFG.UMBRAL_RECOMPRA,
      promedioActual: promedioActual,
      bajaPromedio: bajaPromedio
    };
  });
}


function calcularUsdt(movs) {
  var saldo = 0;
  movs.forEach(function (m) {
    if (m.tipo !== 'COMPRA' && m.tipo !== 'VENTA') return;   // filas sin tipo o desconocido no mueven saldo
    if (m.tipo === 'VENTA' && m.error) return;               // venta rechazada por exceder la posicion

    const esStable = CFG.IGNORAR_EN_RESUMEN.indexOf(m.cripto) > -1;
    if (esStable) {
      saldo += (m.tipo === 'COMPRA' ? m.cantidad : -m.cantidad);
    } else {
      saldo += (m.tipo === 'VENTA' ? m.usd : -m.usd);
    }
  });
  return saldo;
}

function calcularCajas(estado, precios, usdt) {
  const asignadas = {};
  CFG.CASILLEROS.forEach(function (cas) {
    cas.monedas.forEach(function (m) { if (m !== '*') asignadas[m] = true; });
  });

  const abiertas = estado.orden.filter(function (c) {
    return CFG.IGNORAR_EN_RESUMEN.indexOf(c) === -1 && estado.pos[c].cantidad > CFG.TOL;
  });

  const cajas = CFG.CASILLEROS.map(function (cas) {
    var miembros;
    if (!cas.monedas.length)          miembros = [];
    else if (cas.monedas[0] === '*')  miembros = abiertas.filter(function (c) { return !asignadas[c]; });
    else                              miembros = cas.monedas.filter(function (c) {
      return precios[c] && estado.pos[c] && estado.pos[c].cantidad > CFG.TOL;
    });

    // Solo el casillero sin monedas (USDT) vale el saldo libre; uno con monedas
    // pero sin posicion abierta vale 0.
    const valor = !cas.monedas.length
      ? usdt
      : miembros.reduce(function (a, c) { return a + estado.pos[c].cantidad * precios[c]; }, 0);

    return { nombre: cas.nombre, pct: cas.pct, miembros: miembros, valor: valor };
  });

  const total = cajas.reduce(function (a, c) { return a + c.valor; }, 0);

  cajas.forEach(function (c) {
    c.peso     = total ? 100 * c.valor / total : 0;
    c.objetivo = total * c.pct / 100;
    c.delta    = c.objetivo - c.valor;
    c.desvio   = c.peso - c.pct;

    if (Math.abs(c.desvio) <= CFG.BANDA_PP)            c.accion = 'Dentro de banda';
    else if (Math.abs(c.delta) < CFG.MIN_OPERACION)    c.accion = 'Fuera de banda, pero < minimo';
    else if (c.delta > 0)                              c.accion = 'COMPRAR';
    else                                               c.accion = 'VENDER';
  });

  return cajas;
}

/**
 * Suma una fila por moneda a la hoja Historial (nunca la limpia, solo APPEND).
 * Se llama desde rebalancear() y desde enviarResumenEmail(), asi que cada
 * corrida manual (tipicamente cuando estas por decidir algo) y cada mail
 * mensual dejan una foto completa de precios, no solo de la moneda que
 * estes mirando en ese momento.
 */
function registrarHistorial(ss, estado, precios, usdt) {
  var hoja = ss.getSheetByName(CFG.HOJA_HISTORIAL);
  if (!hoja) {
    hoja = ss.insertSheet(CFG.HOJA_HISTORIAL);
    hoja.appendRow(['Fecha', 'Moneda', 'Precio', 'Cantidad', 'Valor USD', 'Costo prom', 'P&L %']);
    hoja.getRange(1, 1, 1, 7).setFontWeight('bold');
  }

  const fecha = new Date();
  const filas = [];

  estado.orden.forEach(function (c) {
    if (CFG.IGNORAR_EN_RESUMEN.indexOf(c) > -1) return;
    const p = estado.pos[c];
    if (p.cantidad <= CFG.TOL || !precios[c]) return;   // solo posiciones abiertas con precio

    const costoProm = p.costo / p.cantidad;
    const valor = p.cantidad * precios[c];
    const pnlPct = costoProm ? (precios[c] / costoProm - 1) * 100 : 0;
    filas.push([fecha, c, precios[c], p.cantidad, valor, costoProm, pnlPct]);
  });

  filas.push([fecha, 'USDT', 1, usdt, usdt, 1, 0]);

  if (filas.length) {
    hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, 7).setValues(filas);
  }
}


/**
 * Lee la hoja, calcula posiciones y trae precios (de lo abierto y de lo que
 * alguna vez vendiste, para comparar contra la foto de venta). Falla si falta
 * el precio de una posicion abierta; lo vendido sin precio simplemente se omite.
 */
function prepararCartera(hoja) {
  const movs = leerMovimientos(hoja);
  const estado = procesar(movs);
  const usdt = calcularUsdt(movs);

  const abiertas = estado.orden.filter(function (c) {
    return CFG.IGNORAR_EN_RESUMEN.indexOf(c) === -1 && estado.pos[c].cantidad > CFG.TOL;
  });

  const vendidas = Object.keys(ultimaVenta(movs)).filter(function (c) {
    return CFG.IGNORAR_EN_RESUMEN.indexOf(c) === -1 && abiertas.indexOf(c) === -1;
  });

  const precios = traerPrecios(abiertas.concat(vendidas));

  const faltantes = abiertas.filter(function (c) { return !precios[c]; });
  if (faltantes.length) {
    throw new Error('Sin precio para: ' + faltantes.join(', ') +
                    '. Revisa CFG.COINGECKO.');
  }

  return { movs: movs, estado: estado, usdt: usdt, abiertas: abiertas, precios: precios };
}

function rebalancear() {
  const ss = SpreadsheetApp.getActive();
  const hoja = ss.getSheetByName(CFG.HOJA);
  if (!hoja) throw new Error('No existe la hoja "' + CFG.HOJA + '"');

  const suma = CFG.CASILLEROS.reduce(function (a, c) { return a + c.pct; }, 0);
  if (Math.abs(suma - 100) > 0.01) {
    throw new Error('Los pesos objetivo suman ' + suma + '%, tienen que sumar 100.');
  }

  const d = prepararCartera(hoja);
  const movs = d.movs, estado = d.estado, usdt = d.usdt, abiertas = d.abiertas, precios = d.precios;

  const recompra = candidatosRecompra(movs, precios, estado.pos);

  const detalle = abiertas.map(function (c) {
    const p = estado.pos[c];
    return {
      moneda: c,
      cantidad: p.cantidad,
      costoProm: p.costo / p.cantidad,
      costo: p.costo,
      precio: precios[c],
      valor: p.cantidad * precios[c]
    };
  });

  const cajas = calcularCajas(estado, precios, usdt);
  const total = cajas.reduce(function (a, c) { return a + c.valor; }, 0);

  escribirRebalanceo(ss, cajas, detalle, total, usdt, estado.pos, recompra);
  registrarHistorial(ss, estado, precios, usdt);
  ss.toast('Cartera: ' + total.toFixed(2) + ' USD', 'Rebalanceo', 5);
}

function escribirRebalanceo(ss, cajas, detalle, total, usdt, estadoPos, recompra) {
  var hoja = ss.getSheetByName(CFG.HOJA_REBAL);
  if (!hoja) hoja = ss.insertSheet(CFG.HOJA_REBAL);
  hoja.clear();

  const filas = [];

  filas.push(['REBALANCEO', '', '', '', '', '', '']);
  filas.push(['Actualizado', new Date(), '', '', '', '', '']);
  filas.push(['Valor total', total, '', 'USDT libre', usdt, '', '']);
  filas.push(['', '', '', '', '', '', '']);

  filas.push(['CASILLERO', 'Valor USD', 'Peso %', 'Objetivo %', 'Desvio pp', 'Accion', 'USD a mover']);
  cajas.forEach(function (c) {
    filas.push([
      c.nombre + (c.miembros.length ? ' (' + c.miembros.join(', ') + ')' : ''),
      c.valor, c.peso, c.pct, c.desvio, c.accion,
      c.accion === 'COMPRAR' || c.accion === 'VENDER' ? Math.abs(c.delta) : ''
    ]);
  });

  filas.push(['', '', '', '', '', '', '']);
  filas.push(['DETALLE', 'Cantidad', 'Costo prom', 'Precio hoy', 'Valor USD', 'Costo USD', 'P&L no realizado']);

  var costoTotal = 0, valorTotal = 0;
  detalle.sort(function (a, b) { return b.valor - a.valor; }).forEach(function (d) {
    costoTotal += d.costo;
    valorTotal += d.valor;
    filas.push([d.moneda, d.cantidad, d.costoProm, d.precio, d.valor, d.costo, d.valor - d.costo]);
  });

  filas.push(['TOTAL (posiciones abiertas)', '', '', '', valorTotal, costoTotal, valorTotal - costoTotal]);

  const realizadoTotal = Object.keys(estadoPos).reduce(
    function (a, c) { return a + estadoPos[c].realizado; }, 0
  );
  filas.push(['', '', '', '', '', '', '']);
  filas.push(['P&L ya realizado (posiciones cerradas)', '', '', '', '', '', realizadoTotal]);
  filas.push(['TOTAL NETO INVERTIDO (historico)', '', '', '', '', costoTotal - realizadoTotal, '']);
  filas.push(['CAPITAL TOTAL APORTADO (cripto + USDT libre)', '', '', '', '', costoTotal - realizadoTotal + usdt, '']);

  var filaRecompra = -1;
  if (recompra && recompra.length) {
    filas.push(['', '', '', '', '', '', '']);
    filaRecompra = filas.length + 1;
    filas.push(['RECOMPRA (foto = precio de la ultima venta)', 'Precio venta', 'Precio hoy', 'Caida', 'Tu promedio', 'Candidato', 'Baja tu promedio?']);
    recompra.forEach(function (r) {
      var estadoBaja = '-';
      if (r.promedioActual !== null) {
        estadoBaja = r.bajaPromedio ? 'SI' : 'NO (sigue arriba)';
      }
      filas.push([
        r.moneda, r.precioVenta, r.precioHoy, r.caida, r.promedioActual,
        r.candidato ? 'SI, bajo ' + (CFG.UMBRAL_RECOMPRA * 100) + '%' : 'no todavia',
        estadoBaja
      ]);
    });
  }

  hoja.getRange(1, 1, filas.length, 7).setValues(filas);
  hoja.getRange(1, 1, 1, 7).setFontWeight('bold');
  hoja.getRange(5, 1, 1, 7).setFontWeight('bold');
  hoja.getRange(5 + cajas.length + 2, 1, 1, 7).setFontWeight('bold');
  if (filaRecompra > 0) hoja.getRange(filaRecompra, 1, 1, 7).setFontWeight('bold');
  hoja.setColumnWidth(1, 260);
}


// ============================================================
//  RECORDATORIO POR MAIL
// ============================================================

function enviarResumenEmail() {
  const ss = SpreadsheetApp.getActive();
  const hoja = ss.getSheetByName(CFG.HOJA);
  if (!hoja) throw new Error('No existe la hoja "' + CFG.HOJA + '"');

  const d = prepararCartera(hoja);
  const movs = d.movs, estado = d.estado, usdt = d.usdt, precios = d.precios;

  const recompra = candidatosRecompra(movs, precios, estado.pos).filter(function (r) { return r.candidato; });
  registrarHistorial(ss, estado, precios, usdt);

  const cajas = calcularCajas(estado, precios, usdt);
  const total = cajas.reduce(function (a, c) { return a + c.valor; }, 0);

  const fueraDeBanda = cajas.filter(function (c) {
    return c.accion === 'COMPRAR' || c.accion === 'VENDER';
  });

  var partesAsunto = [];
  if (fueraDeBanda.length) partesAsunto.push(fueraDeBanda.length + ' casillero(s) fuera de banda');
  if (recompra.length)     partesAsunto.push(recompra.length + ' candidato(s) a recompra');
  const asunto = 'Cartera cripto: ' + (partesAsunto.length ? partesAsunto.join(', ') : 'todo dentro de banda');

  var cuerpo = 'Valor total: ' + total.toFixed(2) + ' USD\n'
             + 'USDT libre: ' + usdt.toFixed(2) + ' USD\n\n';

  cajas.forEach(function (c) {
    cuerpo += c.nombre + '  '
            + 'Peso ' + c.peso.toFixed(1) + '%'
            + '  Obj ' + c.pct + '%'
            + '  Desvio ' + (c.desvio >= 0 ? '+' : '') + c.desvio.toFixed(1) + 'pp'
            + '  -> ' + c.accion
            + (c.accion === 'COMPRAR' || c.accion === 'VENDER'
                ? ' ' + Math.abs(c.delta).toFixed(0) + ' USD' : '')
            + '\n';
  });

  if (recompra.length) {
    cuerpo += '\nCANDIDATOS A RECOMPRA (cayeron ' + (CFG.UMBRAL_RECOMPRA * 100) + '%+ desde que las vendiste):\n';
    recompra.forEach(function (r) {
      var nota = r.promedioActual !== null
        ? (r.bajaPromedio ? ' [baja tu promedio de ' + r.promedioActual.toFixed(4) + ']'
                           : ' [OJO: sigue arriba de tu promedio de ' + r.promedioActual.toFixed(4) + ']')
        : '';
      cuerpo += '  ' + r.moneda + '  vendida a ' + r.precioVenta + '  hoy ' + r.precioHoy
              + '  (-' + (r.caida * 100).toFixed(1) + '%)' + nota + '\n';
    });
  }

  cuerpo += '\nPlanilla: ' + ss.getUrl();

  MailApp.sendEmail(CFG.EMAIL.destinatario, asunto, cuerpo);
}
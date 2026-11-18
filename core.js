/* Núcleo: lectura del Excel de TIEMPO, validaciones y generación de los TXT de ARCA
   (Libro IVA Digital - Compras). Funciona en el navegador y en Node (para pruebas). */
(function (root) {
  'use strict';

  // ---------- Tablas ----------
  // Tipo TIEMPO (tc + clase) -> código de comprobante ARCA
  const TIPOS = {
    'FC A': '001', 'ND A': '002', 'NC A': '003', 'RE A': '004',
    'FC B': '006', 'ND B': '007', 'NC B': '008', 'RE B': '009',
    'FC C': '011', 'ND C': '012', 'NC C': '013',
    'FC E': '019', 'ND E': '020', 'NC E': '021',
    'LI A': '063', 'LI B': '064',
    'TK A': '081', 'TK B': '082',
    'FC M': '051', 'ND M': '052', 'NC M': '053',
  };
  const NOMBRES_TIPO = {
    '001': 'Factura A', '002': 'Nota de Débito A', '003': 'Nota de Crédito A', '004': 'Recibo A',
    '006': 'Factura B', '007': 'Nota de Débito B', '008': 'Nota de Crédito B', '009': 'Recibo B',
    '011': 'Factura C', '012': 'Nota de Débito C', '013': 'Nota de Crédito C',
    '019': 'Factura E', '020': 'Nota de Débito E', '021': 'Nota de Crédito E',
    '051': 'Factura M', '052': 'Nota de Débito M', '053': 'Nota de Crédito M',
    '063': 'Liquidación A', '064': 'Liquidación B',
    '081': 'Tique Factura A', '082': 'Tique Factura B',
  };
  // Comprobantes que discriminan IVA (llevan registro de alícuotas)
  const CON_ALICUOTAS = new Set(['001', '002', '003', '004', '051', '052', '053', '063', '081']);
  // Alícuotas: clave interna -> [código ARCA, tasa]
  const ALICUOTAS = [['0', '0003', 0], ['105', '0004', 0.105], ['21', '0005', 0.21], ['27', '0006', 0.27], ['5', '0008', 0.05], ['25', '0009', 0.025]];
  const PREFIJOS_CUIT = ['20', '23', '24', '25', '26', '27', '30', '33', '34'];
  const MESES = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12,
    ene: 1, abr: 4, ago: 8, set: 9, dic: 12 };

  const r2 = (x) => Math.round((Number(x) || 0) * 100) / 100;
  const abs2 = (x) => r2(Math.abs(Number(x) || 0));

  // ---------- Lector de Excel antiguo (BIFF2/3/4, el que exporta TIEMPO) ----------
  function leerBIFF(buf) {
    const u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    const dec = new TextDecoder('windows-1252');
    const id0 = dv.getUint16(0, true);
    if (dv.getUint32(0, true) === 0xE011CFD0) throw new Error('Este Excel es formato 97-2003/2007 (OLE). Guardalo como CSV o exportalo de nuevo desde TIEMPO.');
    if (![0x0009, 0x0209, 0x0409].includes(id0)) throw new Error('No reconozco el formato del archivo Excel.');
    const cells = [];
    const put = (r, c, v) => { (cells[r] = cells[r] || [])[c] = v; };
    const rk = (v) => {
      let n;
      if (v & 2) n = v >> 2;
      else { const b = new DataView(new ArrayBuffer(8)); b.setUint32(4, v & 0xFFFFFFFC, true); n = b.getFloat64(0, true); }
      return (v & 1) ? n / 100 : n;
    };
    let p = 0;
    while (p + 4 <= u8.length) {
      const id = dv.getUint16(p, true), len = dv.getUint16(p + 2, true), d = p + 4;
      switch (id) {
        case 0x0003: put(dv.getUint16(d, true), dv.getUint16(d + 2, true), dv.getFloat64(d + 7, true)); break; // NUMBER (BIFF2)
        case 0x0002: put(dv.getUint16(d, true), dv.getUint16(d + 2, true), dv.getUint16(d + 7, true)); break; // INTEGER
        case 0x0004: { const n = u8[d + 7]; put(dv.getUint16(d, true), dv.getUint16(d + 2, true), dec.decode(u8.subarray(d + 8, d + 8 + n))); break; } // LABEL
        case 0x0006: put(dv.getUint16(d, true), dv.getUint16(d + 2, true), dv.getFloat64(d + 7, true)); break; // FORMULA (BIFF2)
        case 0x0203: put(dv.getUint16(d, true), dv.getUint16(d + 2, true), dv.getFloat64(d + 6, true)); break; // NUMBER (BIFF3/4)
        case 0x0204: { const n = dv.getUint16(d + 6, true); put(dv.getUint16(d, true), dv.getUint16(d + 2, true), dec.decode(u8.subarray(d + 8, d + 8 + n))); break; }
        case 0x027E: put(dv.getUint16(d, true), dv.getUint16(d + 2, true), rk(dv.getUint32(d + 6, true))); break; // RK
        case 0x0206: case 0x0406: put(dv.getUint16(d, true), dv.getUint16(d + 2, true), dv.getFloat64(d + 6, true)); break;
      }
      if (id === 0x000A) break; // EOF
      p = d + len;
    }
    const filas = [];
    for (let i = 0; i < cells.length; i++) {
      const row = cells[i] || [];
      const out = [];
      for (let j = 0; j < row.length; j++) out[j] = row[j] === undefined ? '' : row[j];
      filas.push(out);
    }
    return filas;
  }

  // CSV (por si guardan el Excel como CSV). Separador ; o ,
  function leerCSV(texto) {
    const linea1 = texto.split(/\r?\n/)[0];
    const sep = (linea1.split(';').length >= linea1.split(',').length) ? ';' : ',';
    const filas = []; let fila = [], campo = '', q = false;
    for (let i = 0; i < texto.length; i++) {
      const ch = texto[i];
      if (q) { if (ch === '"') { if (texto[i + 1] === '"') { campo += '"'; i++; } else q = false; } else campo += ch; }
      else if (ch === '"') q = true;
      else if (ch === sep) { fila.push(campo); campo = ''; }
      else if (ch === '\n') { fila.push(campo.replace(/\r$/, '')); filas.push(fila); fila = []; campo = ''; }
      else campo += ch;
    }
    if (campo || fila.length) { fila.push(campo); filas.push(fila); }
    return filas.map((f) => f.map((v) => {
      const t = v.trim();
      if (/^-?[\d.]+,\d+$/.test(t)) return Number(t.replace(/\./g, '').replace(',', '.'));
      if (/^-?\d+(\.\d+)?$/.test(t) && !/^0\d/.test(t)) return Number(t);
      return t;
    }));
  }

  // ---------- Utilidades ----------
  function fechaISO(v) {
    if (typeof v === 'number') { // serial Excel
      const d = new Date(Math.round((v - 25569) * 86400000));
      return d.toISOString().slice(0, 10);
    }
    const s = String(v).trim();
    let m = s.match(/^(\d{1,2})[-\/ ]([A-Za-z]{3})[-\/ ](\d{2,4})$/);
    if (m) { const mes = MESES[m[2].toLowerCase()]; let a = +m[3]; if (a < 100) a += 2000; return `${a}-${String(mes).padStart(2, '0')}-${m[1].padStart(2, '0')}`; }
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (m) { let a = +m[3]; if (a < 100) a += 2000; return `${a}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`; }
    m = s.match(/^(\d{4})-?(\d{2})-?(\d{2})$/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    return '';
  }

  function cuitValido(c) {
    const s = String(c || '').replace(/\D/g, '');
    if (s.length !== 11) return false;
    const w = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2];
    let t = 0; for (let i = 0; i < 10; i++) t += +s[i] * w[i];
    let dv = 11 - (t % 11); if (dv === 11) dv = 0; if (dv === 10) dv = 9;
    return dv === +s[10];
  }

  let _seq = 0;
  const nuevoId = () => 'r' + Date.now().toString(36) + (++_seq).toString(36);

  function filaVacia() {
    return {
      id: nuevoId(), incluir: true, origen: 'manual',
      fecha: '', tipo: '001', pv: 0, nro: 0, cuit: '', nombre: '',
      neto0: 0, neto105: 0, iva105: 0, neto21: 0, iva21: 0, neto27: 0, iva27: 0, neto5: 0, iva5: 0, neto25: 0, iva25: 0,
      netoBC: 0, nograv: 0, exento: 0, perIVA: 0, perNac: 0, perIIBB: 0, perMun: 0, impInt: 0, otros: 0,
      total: 0, codOp: 'auto', cfcManual: null, nota: '',
    };
  }

  // ---------- Excel TIEMPO -> filas ----------
  function desdeTiempo(filas) {
    const enc = filas[0].map((h) => String(h).trim().toLowerCase());
    const col = (n) => enc.indexOf(n);
    const req = ['fecha', 'tc', 'clase', 'estab', 'nc', 'razon', 'cuit', 'columna1', 'columna16'];
    const falta = req.filter((n) => col(n) < 0);
    if (falta.length) throw new Error('El archivo no tiene las columnas de TIEMPO esperadas: ' + falta.join(', '));
    const g = (row, n) => { const i = col(n); return i < 0 ? '' : row[i]; };
    const num = (row, n) => { const v = g(row, n); return typeof v === 'number' ? v : (Number(String(v).replace(',', '.')) || 0); };
    const out = [];
    for (let i = 1; i < filas.length; i++) {
      const row = filas[i];
      if (!row || row.every((v) => v === '' || v === undefined)) continue;
      const tc = String(g(row, 'tc')).trim().toUpperCase(), cl = String(g(row, 'clase')).trim().toUpperCase();
      const f = filaVacia();
      f.origen = 'tiempo';
      f.tiempo = { tc, clase: cl, cuenta: g(row, 'cuenta') };
      f.fecha = fechaISO(g(row, 'fecha'));
      f.tipo = TIPOS[tc + ' ' + cl] || '';
      f.pv = parseInt(String(g(row, 'estab')).replace(/\D/g, '') || '0', 10);
      f.nro = parseInt(String(g(row, 'nc')).replace(/\D/g, '') || '0', 10);
      f.cuit = String(g(row, 'cuit')).replace(/\D/g, '');
      f.nombre = String(g(row, 'razon')).trim();
      const c = (k) => abs2(num(row, 'columna' + k));
      const neto = c(1);
      f.iva21 = c(2); f.iva105 = c(3); f.iva27 = c(4);
      f.exento = c(5); f.perIIBB = c(6); f.perMun = c(8); f.nograv = c(9);
      f.perIVA = c(11); f.impInt = c(12); f.perNac = c(14);
      f.total = c(16);
      // Columnas que nunca vimos con datos: se guardan para avisar
      const raras = [7, 10, 13, 15, 17].filter((k) => c(k) !== 0);
      if (raras.length) f.nota = 'Atención: TIEMPO trae importes en columna' + raras.join(', columna') + ' (no se exportan).';
      repartirNeto(f, neto);
      if (tc === 'LI' && (!cuitValido(f.cuit) || !PREFIJOS_CUIT.includes(f.cuit.slice(0, 2)))) { f.incluir = false; f.nota = (f.nota ? f.nota + ' ' : '') + 'Liquidación con CUIT inválido (despacho de importación, ARCA ya lo trae): excluido.'; }
      out.push(f);
    }
    return out;
  }

  // TIEMPO trae un solo neto. Si hay más de una alícuota se separa a partir del IVA de cada una;
  // la de mayor importe se queda con la diferencia para que la suma cierre con el neto original.
  function repartirNeto(f, neto) {
    ['0', '105', '21', '27', '5', '25'].forEach((k) => { f['neto' + k] = 0; });
    if (!CON_ALICUOTAS.has(f.tipo)) { f.neto0 = 0; f.netoBC = neto; return; }
    const con = [['105', 0.105], ['21', 0.21], ['27', 0.27], ['5', 0.05], ['25', 0.025]].filter(([k]) => f['iva' + k] !== 0);
    if (con.length === 0) { f.neto0 = neto; return; }
    if (con.length === 1) { f['neto' + con[0][0]] = neto; return; }
    let mayor = con[0];
    con.forEach((a) => { if (f['iva' + a[0]] > f['iva' + mayor[0]]) mayor = a; });
    let resto = neto;
    con.forEach(([k, t]) => { if (k !== mayor[0]) { f['neto' + k] = r2(f['iva' + k] / t); resto -= f['neto' + k]; } });
    f['neto' + mayor[0]] = r2(resto);
  }

  // ---------- Cálculos ----------
  function netoGravado(f) { return r2(ALICUOTAS.reduce((s, [k]) => s + (f['neto' + k] || 0), 0)); }
  function ivaTotal(f) { return r2(ALICUOTAS.reduce((s, [k]) => s + (k === '0' ? 0 : (f['iva' + k] || 0)), 0)); }
  function netoNoDiscr(f) { return f.netoBC || 0; }
  function sumaComponentes(f) {
    return r2(netoGravado(f) + ivaTotal(f) + netoNoDiscr(f) + f.nograv + f.exento + f.perIVA + f.perNac + f.perIIBB + f.perMun + f.impInt + f.otros);
  }
  function alicuotasDe(f) {
    if (!CON_ALICUOTAS.has(f.tipo)) return [];
    const res = [];
    ALICUOTAS.forEach(([k, cod]) => {
      const neto = f['neto' + k] || 0, iva = k === '0' ? 0 : (f['iva' + k] || 0);
      if (neto !== 0 || iva !== 0) res.push({ cod, neto, iva });
    });
    // ARCA exige al menos una alícuota en comprobantes A: si todo es exento/no gravado va 0% en cero
    if (res.length === 0) res.push({ cod: '0003', neto: 0, iva: 0 });
    return res;
  }
  // Crédito fiscal computable: por defecto la suma del IVA; se puede fijar a mano (cfcManual)
  function creditoFiscal(f) {
    if (!CON_ALICUOTAS.has(f.tipo)) return 0;
    return (f.cfcManual !== null && f.cfcManual !== undefined && f.cfcManual !== '') ? r2(f.cfcManual) : ivaTotal(f);
  }
  function codigoOperacion(f) {
    if (f.codOp && f.codOp !== 'auto') return f.codOp === '-' ? ' ' : f.codOp;
    if (f.nograv > 0) return 'N';
    if (f.exento > 0 && netoGravado(f) === 0) return 'E';
    return ' ';
  }

  function validar(f, periodo) {
    const err = [], av = [];
    if (!f.tipo || !NOMBRES_TIPO[f.tipo]) err.push('Tipo de comprobante desconocido' + (f.tiempo ? ` (${f.tiempo.tc} ${f.tiempo.clase})` : ''));
    if (!/^\d{4}-\d{2}-\d{2}$/.test(f.fecha)) err.push('Fecha inválida');
    if (!cuitValido(f.cuit)) av.push('CUIT con dígito verificador inválido');
    else if (!PREFIJOS_CUIT.includes(f.cuit.slice(0, 2))) av.push('CUIT de exterior / genérico');
    if (!f.nro) err.push('Falta número');
    if (!f.nombre) av.push('Falta razón social');
    const dif = r2(f.total - sumaComponentes(f));
    if (Math.abs(dif) > 0.05) av.push(`El total no cierra con la suma de los importes (diferencia ${dif.toFixed(2)})`);
    ALICUOTAS.forEach(([k, , t]) => {
      if (k === '0') return;
      const n = f['neto' + k] || 0, i = f['iva' + k] || 0;
      if ((n || i) && Math.abs(r2(n * t) - i) > 1) av.push(`IVA ${t * 100}% no coincide con el neto (esperado ${r2(n * t).toFixed(2)})`);
    });
    if (periodo && f.fecha && f.fecha.slice(0, 7) !== periodo) av.push('Fecha fuera del período (se informa igual)');
    return { err, av };
  }

  // ---------- Generación TXT ----------
  const imp = (v) => String(Math.round(Math.abs(Number(v) || 0) * 100)).padStart(15, '0');
  const num = (v, n) => String(v).replace(/\D/g, '').padStart(n, '0').slice(-n);
  function texto(s, n) {
    // Solo caracteres latin-1; el resto se reemplaza
    const limpio = String(s || '').normalize('NFC').replace(/[^\x20-\x7E\xA0-\xFF]/g, '?');
    return limpio.slice(0, n).padEnd(n, ' ');
  }

  function ordenar(filas) {
    return filas.slice().sort((a, b) => a.tipo.localeCompare(b.tipo) || a.pv - b.pv || a.nro - b.nro || a.cuit.localeCompare(b.cuit));
  }

  function generarTXT(filas) {
    const lista = ordenar(filas.filter((f) => f.incluir));
    const cb = [], al = [];
    lista.forEach((f) => {
      const alis = alicuotasDe(f);
      const cfc = creditoFiscal(f);
      cb.push(
        f.fecha.replace(/-/g, '') + num(f.tipo, 3) + num(f.pv, 5) + num(f.nro, 20) + ' '.repeat(16) +
        '80' + num(f.cuit, 20) + texto(f.nombre, 30) +
        imp(f.total) + imp(f.nograv) + imp(f.exento) + imp(f.perIVA) + imp(f.perNac) + imp(f.perIIBB) +
        imp(f.perMun) + imp(f.impInt) + 'PES' + '0001000000' + String(Math.min(alis.length, 9)) +
        codigoOperacion(f) + imp(cfc) + imp(f.otros) + '00000000000' + ' '.repeat(30) + imp(0)
      );
      alis.forEach((a) => {
        al.push(num(f.tipo, 3) + num(f.pv, 5) + num(f.nro, 20) + '80' + num(f.cuit, 20) + imp(a.neto) + a.cod + imp(a.iva));
      });
    });
    const unir = (arr) => arr.length ? arr.join('\r\n') + '\r\n' : '';
    return { cbte: unir(cb), alicuotas: unir(al), cantidad: lista.length };
  }

  // Texto -> bytes ISO-8859-1 (ARCA espera ese encoding)
  function aLatin1(s) {
    const b = new Uint8Array(s.length);
    for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); b[i] = c < 256 ? c : 63; }
    return b;
  }

  // ---------- Leer TXT de ARCA (para reabrir y corregir uno ya generado) ----------
  function desdeTXT(cbteTxt, aliTxt) {
    const tasa = { '0003': '0', '0004': '105', '0005': '21', '0006': '27', '0008': '5', '0009': '25' };
    const mapa = new Map();
    const v = (s) => Number(s) / 100;
    cbteTxt.split(/\r?\n/).filter((l) => l.trim().length >= 300).forEach((l) => {
      const f = filaVacia();
      f.origen = 'txt';
      f.fecha = `${l.slice(0, 4)}-${l.slice(4, 6)}-${l.slice(6, 8)}`;
      f.tipo = l.slice(8, 11); f.pv = +l.slice(11, 16); f.nro = +l.slice(16, 36);
      f.cuit = String(+l.slice(54, 74)); f.nombre = l.slice(74, 104).trim();
      f.total = v(l.slice(104, 119)); f.nograv = v(l.slice(119, 134)); f.exento = v(l.slice(134, 149));
      f.perIVA = v(l.slice(149, 164)); f.perNac = v(l.slice(164, 179)); f.perIIBB = v(l.slice(179, 194));
      f.perMun = v(l.slice(194, 209)); f.impInt = v(l.slice(209, 224)); f.otros = v(l.slice(254, 269));
      const op = l[238]; f.codOp = op === ' ' ? 'auto' : op;
      f._cfc = v(l.slice(239, 254));
      if (!CON_ALICUOTAS.has(f.tipo)) f.netoBC = r2(f.total - f.nograv - f.exento - f.perIVA - f.perNac - f.perIIBB - f.perMun - f.impInt - f.otros);
      const k = [f.tipo, f.pv, f.nro, f.cuit].join('|');
      if (!mapa.has(k)) mapa.set(k, []);
      mapa.get(k).push(f);
    });
    (aliTxt || '').split(/\r?\n/).filter((l) => l.trim().length >= 80).forEach((l) => {
      const k = [l.slice(0, 3), +l.slice(3, 8), +l.slice(8, 28), String(+l.slice(30, 50))].join('|');
      const fs = mapa.get(k); if (!fs) return;
      const t = tasa[l.slice(65, 69)]; if (!t) return;
      // si hay comprobantes repetidos con la misma clave, asigna al primero que no tenga esa alícuota
      const f = fs.find((x) => !x['_a' + t]) || fs[0];
      f['_a' + t] = 1;
      f['neto' + t] = r2(f['neto' + t] + v(l.slice(50, 65)));
      if (t !== '0') f['iva' + t] = r2(f['iva' + t] + v(l.slice(69, 84)));
    });
    const out = [];
    mapa.forEach((fs) => fs.forEach((f) => { Object.keys(f).filter((x) => x.startsWith('_a')).forEach((x) => delete f[x]); out.push(f); }));
    // el código de operación se respeta tal cual venía
    out.forEach((f) => {
      if (f.codOp === 'auto' && codigoOperacion(f) !== ' ') f.codOp = '-';
      if (CON_ALICUOTAS.has(f.tipo) && Math.abs(f._cfc - ivaTotal(f)) > 0.004) f.cfcManual = f._cfc;
      delete f._cfc;
    });
    return out;
  }

  const api = { TIPOS, NOMBRES_TIPO, CON_ALICUOTAS, ALICUOTAS, leerBIFF, leerCSV, desdeTiempo, desdeTXT, repartirNeto,
    filaVacia, netoGravado, ivaTotal, sumaComponentes, alicuotasDe, codigoOperacion, creditoFiscal, validar, generarTXT, aLatin1, cuitValido, fechaISO, r2, ordenar };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.IvaCore = api;
})(this);


const express = require('express');
const ExcelJS = require('exceljs');

const app = express();
app.use(express.json({ limit: '20mb' }));

const LOGO_BURGAS_BASE64 = '/9j/4QAYRXhpZgAASUkqAAgAAAAAAAAAAAAAAP/sABFEdWNreQABAAQAAAA8AAD/4QMraHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLwA8P3hwYWNrZXQgYmVnaW49Iu+7vyIgaWQ9Ilc1TTBNcENlaGlIenJlU3pOVGN6a2M5ZCI/PiA8eDp4bXBtZXRhIHhtbG5zOng9ImFkb2JlOm5zOm1ldGEvIiB4OnhtcHRrPSJBZG9iZSBYTVAgQ29yZSA1LjAtYzA2MSA2NC4xNDA5NDksIDIwMTAvMTIvMDctMTA6NTc6MDEgICAgICAgICI+IDxyZGY6UkRGIHhtbG5zOnJkZj0iaHR0cDovL3d3dy53My5vcmcvMTk5OS8wMi8yMi1yZGYtc3ludGF4LW5zIyI+IDxyZGY6RGVzY3JpcHRpb24gcmRmOmFib3V0PSIiIHhtbG5zOnhtcD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLyIgeG1sbnM6eG1wTU09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9tbS8iIHhtbG5zOnN0UmVmPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VSZWYjIiB4bXA6Q3JlYXRvclRvb2w9IkFkb2JlIFBob3Rvc2hvcCBDUzUuMSBXaW5kb3dzIiB4bXBNTTpJbnN0YW5jZUlEPSJ4bXAuaWlkOjIxM0RDNjc5NkVEOTExRTNCNUM0QTVENTQ4REY1NzAxIiB4bXBNTTpEb2N1bWVudElEPSJ4bXAuZGlkOjIxM0RDNjdBNkVEOTExRTNCNUM0QTVENTQ4REY1NzAxIj4gPHhtcE1NOkRlcml2ZWRGcm9tIHN0UmVmOmluc3RhbmNlSUQ9InhtcC5paWQ6MjEzREM2Nzc2RUQ5MTFFM0I1QzRBNUQ1NDhERjU3MDEiIHN0UmVmOmRvY3VtZW50SUQ9InhtcC5kaWQ6MjEzREM2Nzg2RUQ5MTFFM0I1QzRBNUQ1NDhERjU3MDEiLz4gPC9yZGY6RGVzY3JpcHRpb24+IDwvcmRmOlJERj4gPC94OnhtcG1ldGE+IDw/eHBhY2tldCBlbmQ9InIiPz7/7gAOQWRvYmUAZMAAAAAB/9sAhAAGBAQEBQQGBQUGCQYFBgkLCAYG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OkkGWWQn969p3jZsaD4F16PSRWc24yztsy0r4XwLuyyw7Twhp3w6SOYYCSplc3wj0W/pavH9dOdn6OjVybsuNoICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICDn3EfR81TN+M0MZkflDayJoxdg31ZAOXAbCu/wBJvx8sstlfi518KeZd/cywn4Q8ydxg+EPMncYPhDzJ3GD4Q8ydxg+EPMncYPhDzJ3GD4Q8ydxg+EPMncYXtn05XXasbS0keJJHWSEeixvK5xVNm2KxmUxXLttqttPbLdT0FP8AwqdgYCd5PK4+EnavGvebTMy6IjC7VUiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIMTXaU0/XSGSejZ1rvWezFhPjykLWu69eUqzWFp/IOmO7P7WTzq/ub9Tsg/kHTHdn9rJ509zfqdkH8g6Y7s/tZPOnub9Tsg/kHTHdn9rJ509zfqdkH8g6Y7s/tZPOnub9Tsg/kHTHdn9rJ509zfqdkH8g6Y7s/tZPOnub9TshLNB6Ya4H4Uuw5HSSEfpUe5v1OyGZo6GjooRDSQsgiH1WADp51la0zzTEPdVSICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg+N3wV1z1bqnraTUlzbT3SoZGbLI4xxNM0noyYxy4bvR3L63MV1040j5Y/k4eczzZDTtTerFxA0zHQsv9mbX1QgqW3+Q9VPE4hrmMaI4wdhw5duG5Z7YrfVbPZbEf7U1zFo5/q7B/TRWT1XDyd087p5mXKpa4yPL3AYMIBxJPKvK+rViNvD+sN9E/KxXFatqo+MWn4IaiRkZsl0fJEx7gCRSVRa4tB527CtPRVj29px/ur/AOYV2T88fk4vpiu0RNZYZL/c9VC6Ev64W7q302AcQzI6R2b1cMfCvY3V2Rb5Y19v482FZjHHLZdE6orLXqO61uha271tjt1oq6q7svZYQyaOJ5iORhc3+IGYcu8bsVzeo0xakRtisWm0Y7VqWxPy8sMtRcPn3HhvBr+8XnUN5ulf++mobZKC89ZMY8Gtc2Q+hvP6As7eq7d3irWlax8ZWima90zLxqeM0lj4e3DSdho7rbrrQMH/AD91cx1UxtROM+5rC1+WX0DhsCmPQd+2NlprNZ+EcuEI8uK4hrVNT2Z9PE6rptcy1j2A1EseQMdIRi5zQ5rnYY7sSuiZtnhOrCvD/wCzMW/UHErUPCast1FNca5lnvDIKzqc34k63ujcercRmccjxtG3w+iFlbVp174me2O6v6ZTFrTX9WNfbbWxrnvtOvWRsBc57ntADRtJP7nkC0i9uur/AF+qMf8AJ9GcH6uw1XD61vsVXVVlA0SNbLXEGpD+sc57JcuzFpOGzZhgvn/XVtG2e6Iifw5OrVMdvBua5GggICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIOPT/wBOzfxa5XGg1ddbYbnUyVU8NI7qm5pHufgcjm5subAYr1o+qfLETSs4jHFh4OPNfaf4C0NBqCgvV31Dc7/JbH9dQ09dIXRslBBD9pcdhAOAw2jaqbfqU2pNa1rXPPCa6cTmZy8br/T1a5LvWXCw6gueno6+QzVNFQyFsJkJxJaAWkDbuOOHJsU0+qW7Yi1a2x1J0xnhOF9pfgXZbLU19xq7pW3m81tLLRMuNa/O+GKVhY7IDj6WBwxJ3bOdU3fUbXiIiIrWJziE10xDaNAaKpNHaXpdP09Q+rhpXSObPK1rXHrZDIdg2bMy5vU+onbebzGMrUp2xhh63hJaZdcSaro6qSidWwuprzbWMa6mrYpBlkErThhnbhjhyjHeta+utGrxzGccp+MInXGctYb/AE5U9G+SOyavvNpt5e58VDBMQyPMcSAWuZj5Riun7rM/ypW09VPB0mVw7+nPTbtN11rfc6ya43OeKorrzPllqH9SSWsGOwNJdieU8pVfut++LYjEcoPBGMOrwxiKFkQOIY0NxPLgMF5kzlu5seCUUdtvNHQ6hrqCS83P8VlqqbCKVhweDECwtxaes+heh9w41maxPbXDHxc+LFP/AKfbk9jmP1/fnMcC1zTM4gg7CCM61+5x/wDnRHh/GXRdF6PtOkNO0titQeaamxJkkIMkkjzme95AAxJ5gvP9RvttvNrc2tKxWMQzixWEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEH//Z';
const TARIFA_OFICIAL_DEFECTO = 30.75;
const TARIFA_AYUDANTE_DEFECTO = 30.75;
const TARIFA_OFICINA_DEFECTO = 30.75;
const TARIFA_DESPLAZAMIENTO_DEFECTO = 15;
const IVA_DEFECTO = 0.21;

const FILAS = {
  equipo: { inicio: 16, fin: 20 },
  materiales: { inicio: 24, fin: 48 },
  manoObra: { inicio: 52, fin: 61 }
};

const COLS = [
  { key: 'tipo', width: 14 },
  { key: 'codigo', width: 15 },
  { key: 'descripcion', width: 42 },
  { key: 'cantidad', width: 11 },
  { key: 'unidad', width: 9 },
  { key: 'precio', width: 13 },
  { key: 'descuento', width: 12 },
  { key: 'total', width: 14 },
  { key: 'notas', width: 26 }
];

function texto(valor) {
  return valor === null || valor === undefined ? '' : String(valor);
}

function normalizar(valor) {
  return texto(valor).normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase();
}

function numero(valor, defecto = '') {
  if (valor === null || valor === undefined || valor === '') return defecto;
  if (typeof valor === 'number') return Number.isFinite(valor) ? valor : defecto;
  const raw = String(valor).trim().replace(/\s/g, '');
  if (!raw) return defecto;
  const normalizado = raw.includes(',') ? raw.replace(/\./g, '').replace(',', '.') : raw.replace(/[^0-9.-]/g, '');
  const n = Number(normalizado);
  return Number.isFinite(n) ? n : defecto;
}

function porcentaje(valor, defecto = 0) {
  const n = numero(valor, null);
  if (n === null || n === '') return defecto;
  return Math.abs(n) > 1 ? n / 100 : n;
}

function primerValor(objeto, campos, defecto = '') {
  for (const campo of campos) {
    if (objeto && objeto[campo] !== undefined && objeto[campo] !== null && objeto[campo] !== '') return objeto[campo];
  }
  return defecto;
}

function validarEntrada(body) {
  if (!body || typeof body !== 'object') throw new Error('El cuerpo de la petición debe ser un objeto JSON.');
}

function obtenerPayload(body) {
  validarEntrada(body);
  // Compatible con tu n8n actual: { presupuesto: {...}, lineas: [...] }
  // También acepta que el body sea directamente el presupuesto.
  const presupuesto = body.presupuesto && typeof body.presupuesto === 'object' ? body.presupuesto : body;
  const lineas = Array.isArray(body.lineas)
    ? body.lineas
    : Array.isArray(body.lineas_presupuesto)
      ? body.lineas_presupuesto
      : Array.isArray(presupuesto.lineas)
        ? presupuesto.lineas
        : [];
  return { presupuesto, lineas };
}

function aplicarBordes(cell, color = 'D9E2F3') {
  cell.border = {
    top: { style: 'thin', color: { argb: color } },
    left: { style: 'thin', color: { argb: color } },
    bottom: { style: 'thin', color: { argb: color } },
    right: { style: 'thin', color: { argb: color } }
  };
}

function configurarHoja(ws) {
  ws.name = 'PRESSUPOST';
  ws.views = [{ showGridLines: false }];
  ws.pageSetup = { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  ws.properties.defaultRowHeight = 20;

  COLS.forEach((col, i) => { ws.getColumn(i + 1).width = col.width; });
  ws.getColumn('F').numFmt = '#,##0.00 €';
  ws.getColumn('G').numFmt = '0.00%';
  ws.getColumn('H').numFmt = '#,##0.00 €';

  ws.mergeCells('A1:I1');
  ws.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '2F4FA3' } };
  ws.getRow(1).height = 12;

  ws.mergeCells('A2:C5');
  ws.mergeCells('D2:I2');
  ws.getCell('D2').value = 'PRESSUPOST INSTAL·LACIÓ';
  ws.getCell('D2').font = { bold: true, size: 22, color: { argb: '1F3763' } };
  ws.getCell('D2').alignment = { horizontal: 'right', vertical: 'middle' };

  ws.mergeCells('D3:F3'); ws.getCell('D3').value = 'Núm. pressupost';
  ws.getCell('G3').value = '';
  ws.mergeCells('H3:I3'); ws.getCell('H3').value = 'Data';
  ws.mergeCells('D4:F4'); ws.getCell('D4').value = 'Estat';
  ws.getCell('G4').value = 'Pendent de revisió';
  ws.mergeCells('H4:I4'); ws.getCell('H4').value = '';

  ['D3','D4','H3','H4'].forEach(c => {
    ws.getCell(c).font = { bold: true, color: { argb: '1F3763' } };
    ws.getCell(c).alignment = { horizontal: 'right' };
  });
  ws.getCell('I3').numFmt = 'dd/mm/yyyy';
  // H4 se deja libre para futuras notas internas; no debe contener una fórmula.

  try {
    const logoId = ws.workbook.addImage({ base64: LOGO_BURGAS_BASE64, extension: 'jpeg' });
    ws.addImage(logoId, { tl: { col: 0, row: 1.15 }, ext: { width: 240, height: 65 } });
  } catch (e) {
    console.warn('No se pudo insertar el logo:', e.message);
  }

  crearBloqueDatos(ws);
  crearSeccionTabla(ws, 14, 'EQUIP PRINCIPAL');
  crearSeccionTabla(ws, 22, 'MATERIALS COMPLEMENTARIS');
  crearSeccionTabla(ws, 50, 'MÀ D’OBRA I ALTRES');
  crearResumen(ws);

  ws.getRow(12).height = 45;
  ws.freezePanes = { row: 14 };
}

function crearBloqueDatos(ws) {
  ws.mergeCells('A6:I6');
  ws.getCell('A6').value = 'DADES DEL CLIENT I DE LA INSTAL·LACIÓ';
  ws.getCell('A6').font = { bold: true, color: { argb: 'FFFFFF' } };
  ws.getCell('A6').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1F3763' } };
  ws.getCell('A6').alignment = { horizontal: 'left' };

  const campos = [
    ['A7','Client','B7'], ['E7','Empresa','F7'],
    ['A8','Telèfon','B8'], ['E8','Email','F8'],
    ['A9','Adreça','B9'],
    ['A10','Tipus instal·lació','B10'], ['E10','Marca','F10'], ['G10','Gamma','H10'],
    ['A11','Unitats','B11'], ['C11','Superfície m²','D11'], ['E11','Metres tuberia','F11'], ['G11','Data visita','H11'],
    ['A12','Observacions','B12']
  ];
  ws.mergeCells('B9:I9');
  ws.mergeCells('B12:I12');
  campos.forEach(([labelCell, label]) => {
    const c = ws.getCell(labelCell);
    c.value = label;
    c.font = { bold: true, color: { argb: '1F3763' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EAF0FA' } };
  });
  ['A7:I12'].forEach(r => {
    for (let row = 7; row <= 12; row++) for (let col = 1; col <= 9; col++) aplicarBordes(ws.getCell(row, col));
  });
  ws.getCell('B12').alignment = { wrapText: true, vertical: 'top' };
}

function crearSeccionTabla(ws, headerRow, titulo) {
  ws.mergeCells(`A${headerRow}:I${headerRow}`);
  const title = ws.getCell(`A${headerRow}`);
  title.value = titulo;
  title.font = { bold: true, color: { argb: 'FFFFFF' } };
  title.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1F3763' } };

  const row = headerRow + 1;
  const headers = ['Tipus', 'Codi', 'Descripció', 'Quantitat', 'Unitat', 'Preu unitari', 'Descompte', 'Total', 'Notes'];
  headers.forEach((h, i) => {
    const cell = ws.getCell(row, i + 1);
    cell.value = h;
    cell.font = { bold: true, color: { argb: '1F3763' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EAF0FA' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    aplicarBordes(cell);
  });
}

function prepararFilasTabla(ws, inicio, fin) {
  for (let r = inicio; r <= fin; r++) {
    ws.getRow(r).height = 22;
    for (let c = 1; c <= 9; c++) {
      const cell = ws.getCell(r, c);
      aplicarBordes(cell, 'E7EDF8');
      cell.alignment = { vertical: 'middle', wrapText: c === 3 || c === 9 };
    }
    ws.getCell(`H${r}`).value = { formula: `IF(OR(D${r}="",F${r}=""),"",ROUND(D${r}*F${r}*(1-IF(G${r}="",0,G${r})),2))` };
  }
}

function crearResumen(ws) {
  ws.mergeCells('G64:I64');
  ws.getCell('G64').value = 'RESUM ECONÒMIC';
  ws.getCell('G64').font = { bold: true, color: { argb: 'FFFFFF' } };
  ws.getCell('G64').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1F3763' } };

  const rows = [
    [65, 'Subtotal equip', '', { formula: 'SUM(H16:H20)' }],
    [66, 'Subtotal materials', '', { formula: 'SUM(H24:H48)' }],
    [67, 'Descompte materials', 0, { formula: 'ROUND(I66*H67,2)' }],
    [68, 'Subtotal materials final', '', { formula: 'I66-I67' }],
    [69, 'Subtotal mà d’obra i altres', '', { formula: 'SUM(H52:H61)' }],
    [70, 'Increment material', 0, { formula: 'ROUND(I68*H70,2)' }],
    [71, 'Base imposable', '', { formula: 'SUM(I65,I68,I69,I70)' }],
    [72, 'IVA', IVA_DEFECTO, { formula: 'ROUND(I71*H72,2)' }],
    [73, 'TOTAL PRESSUPOST', '', { formula: 'I71+I72' }]
  ];
  rows.forEach(([r, label, pct, value]) => {
    ws.getCell(`G${r}`).value = label;
    ws.getCell(`H${r}`).value = pct;
    ws.getCell(`I${r}`).value = value;
    ['G','H','I'].forEach(col => aplicarBordes(ws.getCell(`${col}${r}`)));
    ws.getCell(`G${r}`).font = { bold: r === 73 };
    ws.getCell(`I${r}`).font = { bold: r === 73, size: r === 73 ? 14 : 11 };
    ws.getCell(`H${r}`).numFmt = '0.00%';
    ws.getCell(`I${r}`).numFmt = '#,##0.00 €';
    if (r === 73) {
      ['G','H','I'].forEach(col => {
        ws.getCell(`${col}${r}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D9EAF7' } };
      });
    }
  });

  ws.mergeCells('A65:F73');
  ws.getCell('A65').value = 'Document generat automàticament. Revisar i ajustar imports si és necessari abans d’enviar al client.';
  ws.getCell('A65').font = { italic: true, color: { argb: '666666' } };
  ws.getCell('A65').alignment = { wrapText: true, vertical: 'top' };
}

function crearWorkbookBase() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Burgas Excel Generator';
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.calcProperties.fullCalcOnLoad = true;
  workbook.calcProperties.forceFullCalc = true;
  const ws = workbook.addWorksheet('PRESSUPOST');
  configurarHoja(ws);
  prepararFilasTabla(ws, FILAS.equipo.inicio, FILAS.equipo.fin);
  prepararFilasTabla(ws, FILAS.materiales.inicio, FILAS.materiales.fin);
  prepararFilasTabla(ws, FILAS.manoObra.inicio, FILAS.manoObra.fin);
  return workbook;
}

function escribir(ws, celda, valor) { ws.getCell(celda).value = valor; }

function mapearCabecera(ws, presupuesto) {
  escribir(ws, 'G3', texto(primerValor(presupuesto, ['numero_presupuesto', 'numero', 'id'])));
  const fecha = primerValor(presupuesto, ['fecha', 'created_at', 'fecha_presupuesto'], '');
  escribir(ws, 'I3', fecha ? new Date(fecha) : new Date());
  ws.getCell('I3').numFmt = 'dd/mm/yyyy';
  escribir(ws, 'G4', texto(primerValor(presupuesto, ['estado'], 'Pendent de revisió')));
  escribir(ws, 'B7', texto(primerValor(presupuesto, ['cliente_nombre', 'cliente', 'nombre_cliente'])));
  escribir(ws, 'F7', texto(primerValor(presupuesto, ['empresa'])));
  escribir(ws, 'B8', texto(primerValor(presupuesto, ['telefono', 'teléfono', 'telefon'])));
  escribir(ws, 'F8', texto(primerValor(presupuesto, ['email', 'correo'])));
  escribir(ws, 'B9', texto(primerValor(presupuesto, ['direccion', 'dirección', 'adreca'])));
  escribir(ws, 'B10', texto(primerValor(presupuesto, ['tipo_instalacion', 'tipo'])));
  escribir(ws, 'F10', texto(primerValor(presupuesto, ['marca'])));
  escribir(ws, 'H10', texto(primerValor(presupuesto, ['gama'])));
  escribir(ws, 'B11', numero(primerValor(presupuesto, ['numero_unidades', 'unidades']), ''));
  escribir(ws, 'D11', numero(primerValor(presupuesto, ['superficie', 'm2', 'metros_cuadrados']), ''));
  escribir(ws, 'F11', numero(primerValor(presupuesto, ['metros_tuberia', 'metros_tubería']), ''));
  escribir(ws, 'H11', texto(primerValor(presupuesto, ['fecha_visita', 'data_visita'])));
  escribir(ws, 'B12', texto(primerValor(presupuesto, ['observaciones', 'observacions', 'notas'])));
}

function esLineaEquipo(linea) {
  const contenido = normalizar([linea.tipo, linea.categoria, linea.familia, linea.referencia, linea.codigo, linea.descripcion].filter(Boolean).join(' '));
  return contenido.includes('EQUIPO') || contenido.includes('EQUIP') || contenido.includes('MAQUINA') || contenido.includes('UNITAT') || contenido.includes('UNIDAD') || contenido.includes('SPLIT') || contenido.includes('DAIKIN') || contenido.includes('KOSNER') || contenido.includes('MITSUBISHI');
}
function descripcionLinea(linea) { return texto(primerValor(linea, ['descripcion', 'descripcio', 'descripción', 'nombre', 'referencia'])); }
function codigoLinea(linea) { return texto(primerValor(linea, ['codigo', 'código', 'codigo_saltoki', 'referencia', 'ref'])); }
function unidadLinea(linea, defecto = 'ud') { return texto(primerValor(linea, ['unidad', 'unitat'], defecto)); }
function tipoLinea(linea, defecto) { return texto(primerValor(linea, ['tipo', 'categoria', 'familia'], defecto)); }
function descuentoLinea(linea) { return porcentaje(primerValor(linea, ['descuento_pct', 'descompte_pct', 'descuento_porcentaje', 'descompte_porcentaje'], 0), 0); }
function precioUnitarioLinea(linea) { return numero(primerValor(linea, ['precio_unitario', 'preu_unitari', 'precio', 'preu', 'precio_eur'], ''), ''); }
function escribirLinea(ws, fila, linea, tipoDefecto) {
  escribir(ws, `A${fila}`, tipoLinea(linea, tipoDefecto));
  escribir(ws, `B${fila}`, codigoLinea(linea));
  escribir(ws, `C${fila}`, descripcionLinea(linea));
  escribir(ws, `D${fila}`, numero(primerValor(linea, ['cantidad', 'quantitat', 'qty'], 1), 1));
  escribir(ws, `E${fila}`, unidadLinea(linea));
  escribir(ws, `F${fila}`, precioUnitarioLinea(linea));
  escribir(ws, `G${fila}`, descuentoLinea(linea));
  escribir(ws, `I${fila}`, texto(primerValor(linea, ['notas', 'notes', 'motivo', 'justificacion', 'justificacio'], '')));
}


function valorNumericoCelda(ws, celda, defecto = 0) {
  const valor = ws.getCell(celda).value;
  if (valor && typeof valor === 'object' && Object.prototype.hasOwnProperty.call(valor, 'result')) {
    return numero(valor.result, defecto);
  }
  return numero(valor, defecto);
}

function ponerFormulaConResultado(ws, celda, formula, resultado, numFmt) {
  ws.getCell(celda).value = { formula, result: resultado };
  if (numFmt) ws.getCell(celda).numFmt = numFmt;
}

function recalcularResultadosFormulas(ws) {
  // ExcelJS escribe fórmulas, pero no las calcula. Si el archivo se abre con Vista Previa/QuickLook
  // o en herramientas que no recalculan al abrir, los totales se ven vacíos si no guardamos
  // también el resultado cacheado. Aquí guardamos fórmula + resultado inicial.
  const rangos = [FILAS.equipo, FILAS.materiales, FILAS.manoObra];

  for (const rango of rangos) {
    for (let r = rango.inicio; r <= rango.fin; r++) {
      const cantidad = valorNumericoCelda(ws, `D${r}`, null);
      const precio = valorNumericoCelda(ws, `F${r}`, null);
      const descuento = valorNumericoCelda(ws, `G${r}`, 0);
      const formula = `IF(OR(D${r}="",F${r}=""),"",ROUND(D${r}*F${r}*(1-IF(G${r}="",0,G${r})),2))`;
      const resultado = cantidad === null || precio === null ? '' : Math.round(cantidad * precio * (1 - descuento) * 100) / 100;
      ponerFormulaConResultado(ws, `H${r}`, formula, resultado, '#,##0.00 €');
    }
  }

  const sumarResultados = (inicio, fin) => {
    let total = 0;
    for (let r = inicio; r <= fin; r++) total += valorNumericoCelda(ws, `H${r}`, 0);
    return Math.round(total * 100) / 100;
  };

  const subtotalEquip = sumarResultados(FILAS.equipo.inicio, FILAS.equipo.fin);
  const subtotalMaterials = sumarResultados(FILAS.materiales.inicio, FILAS.materiales.fin);
  const pctDescompte = valorNumericoCelda(ws, 'H67', 0);
  const descompteMaterials = Math.round(subtotalMaterials * pctDescompte * 100) / 100;
  const subtotalMaterialsFinal = Math.round((subtotalMaterials - descompteMaterials) * 100) / 100;
  const subtotalMaObra = sumarResultados(FILAS.manoObra.inicio, FILAS.manoObra.fin);
  const pctIncrement = valorNumericoCelda(ws, 'H70', 0);
  const incrementMaterial = Math.round(subtotalMaterialsFinal * pctIncrement * 100) / 100;
  const baseImposable = Math.round((subtotalEquip + subtotalMaterialsFinal + subtotalMaObra + incrementMaterial) * 100) / 100;
  const pctIva = valorNumericoCelda(ws, 'H72', IVA_DEFECTO);
  const iva = Math.round(baseImposable * pctIva * 100) / 100;
  const total = Math.round((baseImposable + iva) * 100) / 100;

  ponerFormulaConResultado(ws, 'I65', 'SUM(H16:H20)', subtotalEquip, '#,##0.00 €');
  ponerFormulaConResultado(ws, 'I66', 'SUM(H24:H48)', subtotalMaterials, '#,##0.00 €');
  ponerFormulaConResultado(ws, 'I67', 'ROUND(I66*H67,2)', descompteMaterials, '#,##0.00 €');
  ponerFormulaConResultado(ws, 'I68', 'I66-I67', subtotalMaterialsFinal, '#,##0.00 €');
  ponerFormulaConResultado(ws, 'I69', 'SUM(H52:H61)', subtotalMaObra, '#,##0.00 €');
  ponerFormulaConResultado(ws, 'I70', 'ROUND(I68*H70,2)', incrementMaterial, '#,##0.00 €');
  ponerFormulaConResultado(ws, 'I71', 'SUM(I65,I68,I69,I70)', baseImposable, '#,##0.00 €');
  ponerFormulaConResultado(ws, 'I72', 'ROUND(I71*H72,2)', iva, '#,##0.00 €');
  ponerFormulaConResultado(ws, 'I73', 'I71+I72', total, '#,##0.00 €');
}
function siguienteFilaLibre(filaActual, limite, bloque) {
  if (filaActual > limite) throw new Error(`La plantilla no tiene más filas disponibles para ${bloque}.`);
  return filaActual;
}
function mapearLineas(ws, lineas, presupuesto) {
  const entrada = (Array.isArray(lineas) ? lineas : []).map(linea => ({ ...linea, marca: linea.marca ?? presupuesto.marca, gama: linea.gama ?? presupuesto.gama }));
  let filaEquipo = FILAS.equipo.inicio;
  let filaMaterial = FILAS.materiales.inicio;
  if (entrada.length === 0 && presupuesto.marca) {
    entrada.push({ tipo: 'Equip', descripcion: `${presupuesto.marca} ${presupuesto.gama || ''}`, cantidad: primerValor(presupuesto, ['numero_unidades'], 1), unidad: 'ud', precio_unitario: primerValor(presupuesto, ['subtotal_materiales', 'subtotal'], '') });
  }
  const indiceEquipoDetectado = entrada.findIndex(linea => esLineaEquipo(linea));
  const indiceEquipo = indiceEquipoDetectado >= 0 ? indiceEquipoDetectado : 0;
  entrada.forEach((linea, indice) => {
    if (indice === indiceEquipo) {
      filaEquipo = siguienteFilaLibre(filaEquipo, FILAS.equipo.fin, 'equip principal');
      escribirLinea(ws, filaEquipo, linea, 'Equip');
      filaEquipo += 1;
    } else {
      filaMaterial = siguienteFilaLibre(filaMaterial, FILAS.materiales.fin, 'materials complementaris');
      escribirLinea(ws, filaMaterial, linea, 'Material');
      filaMaterial += 1;
    }
  });
}
function tarifaDesdeCosteTotal(presupuesto, campoTotal, campoHoras, tarifaDefecto) {
  const total = numero(presupuesto[campoTotal], null);
  const horas = numero(presupuesto[campoHoras], null);
  if (total !== null && horas !== null && horas > 0) return total / horas;
  return tarifaDefecto;
}
function mapearManoObraYOtros(ws, presupuesto) {
  const lineas = [
    { tipo: "Mà d'obra", descripcion: 'Hores Oficial 1a', cantidad: numero(primerValor(presupuesto, ['horas_oficial_primera', 'hores_oficial_primera'], 0), 0), unidad: 'h', precio_unitario: numero(primerValor(presupuesto, ['tarifa_oficial', 'precio_hora_oficial'], tarifaDesdeCosteTotal(presupuesto, 'coste_oficial', 'horas_oficial_primera', TARIFA_OFICIAL_DEFECTO)), TARIFA_OFICIAL_DEFECTO) },
    { tipo: "Mà d'obra", descripcion: 'Hores Ajudant', cantidad: numero(primerValor(presupuesto, ['horas_ayudante', 'hores_ajudant'], 0), 0), unidad: 'h', precio_unitario: numero(primerValor(presupuesto, ['tarifa_ayudante', 'precio_hora_ayudante'], tarifaDesdeCosteTotal(presupuesto, 'coste_ayudante', 'horas_ayudante', TARIFA_AYUDANTE_DEFECTO)), TARIFA_AYUDANTE_DEFECTO) },
    { tipo: "Mà d'obra", descripcion: 'Hores Oficina', cantidad: numero(primerValor(presupuesto, ['horas_oficina', 'hores_oficina'], 0), 0), unidad: 'h', precio_unitario: numero(primerValor(presupuesto, ['tarifa_oficina', 'precio_hora_oficina'], tarifaDesdeCosteTotal(presupuesto, 'coste_oficina', 'horas_oficina', TARIFA_OFICINA_DEFECTO)), TARIFA_OFICINA_DEFECTO) },
    { tipo: 'Altres', descripcion: 'Desplaçament', cantidad: numero(primerValor(presupuesto, ['numero_desplazamientos', 'desplazamientos', 'desplacaments'], 1), 1), unidad: 'ud', precio_unitario: numero(primerValor(presupuesto, ['tarifa_desplazamiento', 'precio_desplazamiento'], tarifaDesdeCosteTotal(presupuesto, 'coste_desplazamientos', 'numero_desplazamientos', TARIFA_DESPLAZAMIENTO_DEFECTO)), TARIFA_DESPLAZAMIENTO_DEFECTO) }
  ];
  lineas.forEach((linea, indice) => escribirLinea(ws, FILAS.manoObra.inicio + indice, linea, linea.tipo));
}
function mapearPorcentajesResumen(ws, presupuesto) {
  const subtotalMateriales = numero(presupuesto.subtotal_materiales, null);
  const descuentoMateriales = numero(presupuesto.descuento_materiales, null);
  let descuentoPct = porcentaje(primerValor(presupuesto, ['descuento_materiales_pct', 'descompte_materials_pct'], null), null);
  if ((descuentoPct === null || descuentoPct === '') && subtotalMateriales && descuentoMateriales) descuentoPct = descuentoMateriales / subtotalMateriales;
  escribir(ws, 'H67', descuentoPct === null || descuentoPct === '' ? 0 : descuentoPct);
  escribir(ws, 'H70', porcentaje(primerValor(presupuesto, ['incremento_material_pct', 'increment_material_pct'], 0), 0));
  escribir(ws, 'H72', porcentaje(primerValor(presupuesto, ['iva_pct', 'iva_porcentaje'], IVA_DEFECTO), IVA_DEFECTO));
}

app.get('/', (req, res) => res.send('Burgas Excel Generator funcionando - v4'));
app.post('/generar', async (req, res) => {
  try {
    const { presupuesto, lineas } = obtenerPayload(req.body);
    const workbook = crearWorkbookBase();
    const worksheet = workbook.getWorksheet('PRESSUPOST');
    mapearCabecera(worksheet, presupuesto);
    mapearLineas(worksheet, lineas, presupuesto);
    mapearManoObraYOtros(worksheet, presupuesto);
    mapearPorcentajesResumen(worksheet, presupuesto);
    recalcularResultadosFormulas(worksheet);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=pressupost-burgas.xlsx');
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error('[POST /generar]', error);
    res.status(500).json({ error: error.message });
  }
});
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor iniciado en puerto ${PORT}`));

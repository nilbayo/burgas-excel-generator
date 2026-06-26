const express = require('express');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(express.json({ limit: '20mb' }));

const TARIFA_OFICIAL_DEFECTO = 30.75;
const TARIFA_AYUDANTE_DEFECTO = 30.75;
const TARIFA_OFICINA_DEFECTO = 30.75;
const TARIFA_DESPLAZAMIENTO_DEFECTO = 15;
const IVA_DEFECTO = 0.21;

const FILAS = {
  equipo: { inicio: 16, fin: 20 },
  materiales: { inicio: 24, fin: 70 },
  trabajos: { inicio: 74, fin: 83 },
  manoObra: { inicio: 87, fin: 96 }
};

const RESUMEN = {
  header: 99,
  subtotalEquip: 100,
  subtotalMaterials: 101,
  descuentoMaterials: 102,
  subtotalMaterialsFinal: 103,
  subtotalTrabajos: 104,
  subtotalMaObra: 105,
  incrementoMaterial: 106,
  baseImposable: 107,
  iva: 108,
  total: 109
};

const COLS = [
  { key: 'tipo', width: 14 },
  { key: 'codigo', width: 17 },
  { key: 'descripcion', width: 44 },
  { key: 'cantidad', width: 11 },
  { key: 'unidad', width: 9 },
  { key: 'precio', width: 13 },
  { key: 'descuento', width: 12 },
  { key: 'total', width: 14 },
  { key: 'notas', width: 30 }
];

function texto(valor) {
  return valor === null || valor === undefined ? '' : String(valor);
}

function normalizar(valor) {
  return texto(valor)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .trim();
}

function numero(valor, defecto = '') {
  if (valor === null || valor === undefined || valor === '') {
    return defecto;
  }

  if (typeof valor === 'number') {
    return Number.isFinite(valor) ? valor : defecto;
  }

  const raw = String(valor).trim().replace(/\s/g, '');

  if (!raw) {
    return defecto;
  }

  const normalizado = raw.includes(',')
    ? raw.replace(/\./g, '').replace(',', '.')
    : raw.replace(/[^0-9.-]/g, '');

  const n = Number(normalizado);

  return Number.isFinite(n) ? n : defecto;
}

function porcentaje(valor, defecto = 0) {
  const n = numero(valor, null);

  if (n === null || n === '') {
    return defecto;
  }

  return Math.abs(n) > 1 ? n / 100 : n;
}

function primerValor(objeto, campos, defecto = '') {
  for (const campo of campos) {
    if (
      objeto &&
      objeto[campo] !== undefined &&
      objeto[campo] !== null &&
      objeto[campo] !== ''
    ) {
      return objeto[campo];
    }
  }

  return defecto;
}

function validarEntrada(body) {
  if (!body || typeof body !== 'object') {
    throw new Error('El cuerpo de la petición debe ser un objeto JSON.');
  }
}

function obtenerPayload(body) {
  validarEntrada(body);

  const presupuesto =
    body.presupuesto && typeof body.presupuesto === 'object'
      ? body.presupuesto
      : body;

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

function buscarLogo() {
  const candidatos = [
    { file: 'logo-burgas.jpg', extension: 'jpeg' },
    { file: 'logo-burgas.jpeg', extension: 'jpeg' },
    { file: 'logo-burgas.png', extension: 'png' },
    { file: 'image001.jpg', extension: 'jpeg' },
    { file: 'image001.jpeg', extension: 'jpeg' },
    { file: 'image001.png', extension: 'png' }
  ];

  for (const candidato of candidatos) {
    const ruta = path.join(__dirname, candidato.file);

    if (fs.existsSync(ruta)) {
      const buffer = fs.readFileSync(ruta);

      return {
        base64: buffer.toString('base64'),
        extension: candidato.extension
      };
    }
  }

  if (process.env.LOGO_BURGAS_BASE64) {
    return {
      base64: process.env.LOGO_BURGAS_BASE64,
      extension: process.env.LOGO_BURGAS_EXTENSION || 'jpeg'
    };
  }

  return null;
}

function configurarHoja(ws) {
  ws.name = 'PRESSUPOST';
  ws.views = [{ showGridLines: false }];
  ws.pageSetup = {
    paperSize: 9,
    orientation: 'landscape',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0
  };
  ws.properties.defaultRowHeight = 20;

  COLS.forEach((col, i) => {
    ws.getColumn(i + 1).width = col.width;
  });

  ws.getColumn('F').numFmt = '#,##0.00 €';
  ws.getColumn('G').numFmt = '0.00%';
  ws.getColumn('H').numFmt = '#,##0.00 €';

  ws.mergeCells('A1:I1');
  ws.getCell('A1').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '2F4FA3' }
  };
  ws.getRow(1).height = 12;

  ws.mergeCells('A2:C5');
  ws.mergeCells('D2:I2');

  ws.getCell('D2').value = 'PRESSUPOST INSTAL·LACIÓ';
  ws.getCell('D2').font = {
    bold: true,
    size: 22,
    color: { argb: '1F3763' }
  };
  ws.getCell('D2').alignment = {
    horizontal: 'right',
    vertical: 'middle'
  };

  ws.mergeCells('D3:F3');
  ws.getCell('D3').value = 'Núm. pressupost';
  ws.getCell('G3').value = '';

  ws.mergeCells('H3:I3');
  ws.getCell('H3').value = 'Data';

  ws.mergeCells('D4:F4');
  ws.getCell('D4').value = 'Estat';
  ws.getCell('G4').value = 'Pendent de revisió';

  ws.mergeCells('H4:I4');
  ws.getCell('H4').value = '';

  ['D3', 'D4', 'H3', 'H4'].forEach(c => {
    ws.getCell(c).font = {
      bold: true,
      color: { argb: '1F3763' }
    };
    ws.getCell(c).alignment = { horizontal: 'right' };
  });

  ws.getCell('I3').numFmt = 'dd/mm/yyyy';

  const logo = buscarLogo();

  if (logo) {
    try {
      const logoId = ws.workbook.addImage({
        base64: logo.base64,
        extension: logo.extension
      });

      ws.addImage(logoId, {
        tl: { col: 0, row: 1.15 },
        ext: { width: 240, height: 65 }
      });
    } catch (error) {
      console.warn('No se pudo insertar el logo:', error.message);
    }
  }

  crearBloqueDatos(ws);

  crearSeccionTabla(ws, 14, 'EQUIP PRINCIPAL');
  crearSeccionTabla(ws, 22, 'MATERIALS COMPLEMENTARIS');
  crearSeccionTabla(ws, 72, 'TREBALLS ADDICIONALS / REVISIÓ');
  crearSeccionTabla(ws, 85, 'MÀ D’OBRA I ALTRES');

  crearResumen(ws);

  ws.getRow(12).height = 45;
}

function crearBloqueDatos(ws) {
  ws.mergeCells('A6:I6');
  ws.getCell('A6').value = 'DADES DEL CLIENT I DE LA INSTAL·LACIÓ';
  ws.getCell('A6').font = {
    bold: true,
    color: { argb: 'FFFFFF' }
  };
  ws.getCell('A6').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '1F3763' }
  };
  ws.getCell('A6').alignment = { horizontal: 'left' };

  const campos = [
    ['A7', 'Client'],
    ['E7', 'Empresa'],
    ['A8', 'Telèfon'],
    ['E8', 'Email'],
    ['A9', 'Adreça'],
    ['A10', 'Tipus instal·lació'],
    ['E10', 'Marca'],
    ['G10', 'Gamma'],
    ['A11', 'Unitats'],
    ['C11', 'Superfície m²'],
    ['E11', 'Metres tuberia'],
    ['G11', 'Data visita'],
    ['A12', 'Observacions']
  ];

  ws.mergeCells('B9:I9');
  ws.mergeCells('B12:I12');

  campos.forEach(([labelCell, label]) => {
    const c = ws.getCell(labelCell);
    c.value = label;
    c.font = {
      bold: true,
      color: { argb: '1F3763' }
    };
    c.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'EAF0FA' }
    };
  });

  for (let row = 7; row <= 12; row++) {
    for (let col = 1; col <= 9; col++) {
      aplicarBordes(ws.getCell(row, col));
    }
  }

  ws.getCell('B12').alignment = {
    wrapText: true,
    vertical: 'top'
  };
}

function crearSeccionTabla(ws, headerRow, titulo) {
  ws.mergeCells(`A${headerRow}:I${headerRow}`);

  const title = ws.getCell(`A${headerRow}`);
  title.value = titulo;
  title.font = {
    bold: true,
    color: { argb: 'FFFFFF' }
  };
  title.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: titulo.includes('REVISIÓ') ? '9C6500' : '1F3763' }
  };

  const row = headerRow + 1;
  const headers = [
    'Tipus',
    'Codi',
    'Descripció',
    'Quantitat',
    'Unitat',
    'Preu unitari',
    'Descompte',
    'Total',
    'Notes'
  ];

  headers.forEach((h, i) => {
    const cell = ws.getCell(row, i + 1);
    cell.value = h;
    cell.font = {
      bold: true,
      color: { argb: '1F3763' }
    };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'EAF0FA' }
    };
    cell.alignment = {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true
    };
    aplicarBordes(cell);
  });
}

function prepararFilasTabla(ws, inicio, fin) {
  for (let r = inicio; r <= fin; r++) {
    ws.getRow(r).height = 22;

    for (let c = 1; c <= 9; c++) {
      const cell = ws.getCell(r, c);
      aplicarBordes(cell, 'E7EDF8');
      cell.alignment = {
        vertical: 'middle',
        wrapText: c === 3 || c === 9
      };
    }

    ws.getCell(`H${r}`).value = {
      formula: `IF(OR(D${r}="",F${r}=""),"",ROUND(D${r}*F${r}*(1-IF(G${r}="",0,G${r})),2))`,
      result: ''
    };
  }
}

function crearResumen(ws) {
  ws.mergeCells(`G${RESUMEN.header}:I${RESUMEN.header}`);
  ws.getCell(`G${RESUMEN.header}`).value = 'RESUM ECONÒMIC';
  ws.getCell(`G${RESUMEN.header}`).font = {
    bold: true,
    color: { argb: 'FFFFFF' }
  };
  ws.getCell(`G${RESUMEN.header}`).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '1F3763' }
  };

  const rows = [
    [RESUMEN.subtotalEquip, 'Subtotal equip', '', `SUM(H${FILAS.equipo.inicio}:H${FILAS.equipo.fin})`],
    [RESUMEN.subtotalMaterials, 'Subtotal materials', '', `SUM(H${FILAS.materiales.inicio}:H${FILAS.materiales.fin})`],
    [RESUMEN.descuentoMaterials, 'Descompte materials', 0, `ROUND(I${RESUMEN.subtotalMaterials}*H${RESUMEN.descuentoMaterials},2)`],
    [RESUMEN.subtotalMaterialsFinal, 'Subtotal materials final', '', `I${RESUMEN.subtotalMaterials}-I${RESUMEN.descuentoMaterials}`],
    [RESUMEN.subtotalTrabajos, 'Subtotal treballs addicionals', '', `SUM(H${FILAS.trabajos.inicio}:H${FILAS.trabajos.fin})`],
    [RESUMEN.subtotalMaObra, 'Subtotal mà d’obra i altres', '', `SUM(H${FILAS.manoObra.inicio}:H${FILAS.manoObra.fin})`],
    [RESUMEN.incrementoMaterial, 'Increment material', 0, `ROUND(I${RESUMEN.subtotalMaterialsFinal}*H${RESUMEN.incrementoMaterial},2)`],
    [RESUMEN.baseImposable, 'Base imposable', '', `SUM(I${RESUMEN.subtotalEquip},I${RESUMEN.subtotalMaterialsFinal},I${RESUMEN.subtotalTrabajos},I${RESUMEN.subtotalMaObra},I${RESUMEN.incrementoMaterial})`],
    [RESUMEN.iva, 'IVA', IVA_DEFECTO, `ROUND(I${RESUMEN.baseImposable}*H${RESUMEN.iva},2)`],
    [RESUMEN.total, 'TOTAL PRESSUPOST', '', `I${RESUMEN.baseImposable}+I${RESUMEN.iva}`]
  ];

  rows.forEach(([r, label, pct, formula]) => {
    ws.getCell(`G${r}`).value = label;
    ws.getCell(`H${r}`).value = pct;
    ws.getCell(`I${r}`).value = {
      formula,
      result: ''
    };

    ['G', 'H', 'I'].forEach(col => {
      aplicarBordes(ws.getCell(`${col}${r}`));
    });

    ws.getCell(`G${r}`).font = { bold: r === RESUMEN.total };
    ws.getCell(`I${r}`).font = {
      bold: r === RESUMEN.total,
      size: r === RESUMEN.total ? 14 : 11
    };

    ws.getCell(`H${r}`).numFmt = '0.00%';
    ws.getCell(`I${r}`).numFmt = '#,##0.00 €';

    if (r === RESUMEN.total) {
      ['G', 'H', 'I'].forEach(col => {
        ws.getCell(`${col}${r}`).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'D9EAF7' }
        };
      });
    }
  });

  ws.mergeCells(`A${RESUMEN.header}:F${RESUMEN.total}`);
  ws.getCell(`A${RESUMEN.header}`).value =
    'Document generat automàticament. Revisar imports i línies marcades com a revisió abans d’enviar al client.';
  ws.getCell(`A${RESUMEN.header}`).font = {
    italic: true,
    color: { argb: '666666' }
  };
  ws.getCell(`A${RESUMEN.header}`).alignment = {
    wrapText: true,
    vertical: 'top'
  };
}

function observacionesPrincipales(presupuesto) {
  const observaciones = texto(
    primerValor(presupuesto, ['observaciones', 'observacions', 'notas'], '')
  );

  if (!observaciones) {
    return '';
  }

  const marcador = '--- DADES TÈCNIQUES FORMULARI ---';

  if (observaciones.includes(marcador)) {
    return observaciones.split(marcador)[0].trim();
  }

  return observaciones;
}

function valorTecnico(objeto, ruta, defecto = '') {
  try {
    return ruta.split('.').reduce((acc, key) => acc?.[key], objeto) ?? defecto;
  } catch (error) {
    return defecto;
  }
}

function escribirFilaTecnica(ws, fila, seccion, campo, valor) {
  ws.getCell(`A${fila}`).value = seccion;
  ws.getCell(`B${fila}`).value = campo;
  ws.getCell(`C${fila}`).value =
    valor === null || valor === undefined || valor === '' ? '' : valor;

  ['A', 'B', 'C'].forEach(col => {
    aplicarBordes(ws.getCell(`${col}${fila}`), 'D9E2F3');
  });

  ws.getCell(`A${fila}`).font = {
    bold: true,
    color: { argb: '1F3763' }
  };

  ws.getCell(`A${fila}`).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'EAF0FA' }
  };

  ws.getCell(`B${fila}`).font = {
    bold: true
  };

  ws.getCell(`C${fila}`).alignment = {
    wrapText: true,
    vertical: 'top'
  };
}

function crearHojaDatosTecnicos(workbook, presupuesto) {
  const datos = presupuesto.datos_tecnicos || {};
  const avisos = Array.isArray(presupuesto.avisos_revision)
    ? presupuesto.avisos_revision
    : [];

  const ws = workbook.addWorksheet('DADES TÈCNIQUES');

  ws.views = [{ showGridLines: false }];
  ws.properties.defaultRowHeight = 20;

  ws.getColumn('A').width = 26;
  ws.getColumn('B').width = 38;
  ws.getColumn('C').width = 48;
  ws.getColumn('D').width = 18;

  ws.mergeCells('A1:D1');
  ws.getCell('A1').value = 'DADES TÈCNIQUES DE LA INSTAL·LACIÓ';
  ws.getCell('A1').font = {
    bold: true,
    size: 16,
    color: { argb: 'FFFFFF' }
  };
  ws.getCell('A1').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '1F3763' }
  };
  ws.getCell('A1').alignment = {
    horizontal: 'center',
    vertical: 'middle'
  };

  ws.getRow(1).height = 28;

  ws.getCell('A3').value = 'Secció';
  ws.getCell('B3').value = 'Camp';
  ws.getCell('C3').value = 'Valor';

  ['A3', 'B3', 'C3'].forEach(celda => {
    ws.getCell(celda).font = {
      bold: true,
      color: { argb: 'FFFFFF' }
    };
    ws.getCell(celda).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2F4FA3' }
    };
    ws.getCell(celda).alignment = {
      horizontal: 'center'
    };
    aplicarBordes(ws.getCell(celda));
  });

  let fila = 4;

  const filas = [
    ['Configuració', 'Configuració equip', valorTecnico(datos, 'configuracio.configuracion_equipo')],
    ['Configuració', 'Unitats interiors', valorTecnico(datos, 'configuracio.numero_unidades_interiores')],
    ['Configuració', 'Unitats exteriors', valorTecnico(datos, 'configuracio.numero_unidades_exteriores')],
    ['Configuració', 'Potència manual kW', valorTecnico(datos, 'configuracio.potencia_manual_kw')],

    ['Preinstal·lació', 'Hi ha preinstal·lació', valorTecnico(datos, 'preinstalacion.hay_preinstalacion')],
    ['Preinstal·lació', 'Preinstal·lació aprofitable', valorTecnico(datos, 'preinstalacion.preinstalacion_aprovechable')],
    ['Preinstal·lació', 'Retirar màquina existent', valorTecnico(datos, 'preinstalacion.retirar_maquina_existente')],

    ['Ubicació', 'Unitat interior', valorTecnico(datos, 'ubicacion.ubicacion_unidad_interior')],
    ['Ubicació', 'Unitat exterior', valorTecnico(datos, 'ubicacion.ubicacion_unidad_exterior')],
    ['Ubicació', 'Planta unitat interior', valorTecnico(datos, 'ubicacion.planta_unidad_interior')],
    ['Ubicació', 'Planta unitat exterior', valorTecnico(datos, 'ubicacion.planta_unidad_exterior')],
    ['Ubicació', 'Dificultat accés exterior', valorTecnico(datos, 'ubicacion.dificultad_acceso_exterior')],

    ['Canalització', 'Metres tuberia frigorífica', valorTecnico(datos, 'canalizacion.metros_tuberia')],
    ['Canalització', 'Metres canaleta 90x60', valorTecnico(datos, 'canalizacion.metros_canaleta_90x60')],
    ['Canalització', 'Angles interiors 90x60', valorTecnico(datos, 'canalizacion.num_angulos_interiores_90x60')],
    ['Canalització', 'Angles exteriors 90x60', valorTecnico(datos, 'canalizacion.num_angulos_exteriores_90x60')],
    ['Canalització', 'Tapes finals 90x60', valorTecnico(datos, 'canalizacion.num_tapas_finales_90x60')],
    ['Canalització', 'Unions canaleta 90x60', valorTecnico(datos, 'canalizacion.num_uniones_canaleta_90x60')],

    ['Electricitat', 'Cable interconnexió m', valorTecnico(datos, 'electricitat.metros_cable_interconexion')],
    ['Electricitat', 'Necessita línia elèctrica', valorTecnico(datos, 'electricitat.necesita_linea_electrica')],
    ['Electricitat', 'Cable alimentació m', valorTecnico(datos, 'electricitat.metros_cable_alimentacion')],
    ['Electricitat', 'Tipus canal elèctrica', valorTecnico(datos, 'electricitat.tipo_canal_electrica')],
    ['Electricitat', 'Metres canal elèctrica', valorTecnico(datos, 'electricitat.metros_canal_electrica')],
    ['Electricitat', 'Necessita protecció elèctrica', valorTecnico(datos, 'electricitat.necesita_proteccion_electrica')],
    ['Electricitat', 'Tipus protecció elèctrica', valorTecnico(datos, 'electricitat.tipo_proteccion_electrica')],
    ['Electricitat', 'Amperatge magnetotèrmic', valorTecnico(datos, 'electricitat.amperaje_magnetotermico')],

    ['Desguàs / condensats', 'Hi ha desguàs a prop', valorTecnico(datos, 'desguas_condensats.hay_desague_cerca')],
    ['Desguàs / condensats', 'Necessita bomba condensats', valorTecnico(datos, 'desguas_condensats.necesita_bomba_condensados')],
    ['Desguàs / condensats', 'Metres desguàs PVC', valorTecnico(datos, 'desguas_condensats.metros_desague_pvc')],
    ['Desguàs / condensats', 'Diàmetre desguàs PVC', valorTecnico(datos, 'desguas_condensats.diametro_desague_pvc')],

    ['Suport exterior', 'Tipus suport exterior', valorTecnico(datos, 'suport_exterior.tipo_soporte_exterior')],
    ['Suport exterior', 'Silentblocks', valorTecnico(datos, 'suport_exterior.necesita_silentblocks')],

    ['Gas', 'Càrrega suplementària', valorTecnico(datos, 'gas.necesita_carga_gas')],
    ['Gas', 'Tipus gas', valorTecnico(datos, 'gas.tipo_gas')],
    ['Gas', 'Metres extra gas', valorTecnico(datos, 'gas.metros_extra_gas')],

    ['Paleteria', 'Necessita paleteria', valorTecnico(datos, 'paleteria.necesita_paleteria')],
    ['Paleteria', 'Tipus paleteria', valorTecnico(datos, 'paleteria.tipo_paleteria')],
    ['Paleteria', 'Cost paleteria manual', valorTecnico(datos, 'paleteria.coste_paleteria_manual')],
    ['Paleteria', 'Observacions paleteria', valorTecnico(datos, 'paleteria.observaciones_paleteria')],

    ['Seguretat', 'Treball en alçada', valorTecnico(datos, 'seguretat.trabajo_en_altura')],
    ['Seguretat', 'Necessita línia de vida', valorTecnico(datos, 'seguretat.necesita_linea_vida')],
    ['Seguretat', 'Necessita bastida', valorTecnico(datos, 'seguretat.necesita_andamio')],
    ['Seguretat', 'Necessita plataforma', valorTecnico(datos, 'seguretat.necesita_plataforma')],
    ['Seguretat', 'Observacions seguretat', valorTecnico(datos, 'seguretat.observaciones_seguridad')],

    ['Mà d’obra', 'Hores oficial 1a', valorTecnico(datos, 'ma_obra.horas_oficial_primera')],
    ['Mà d’obra', 'Hores oficial 2a', valorTecnico(datos, 'ma_obra.horas_oficial_segunda')],
    ['Mà d’obra', 'Hores ajudant', valorTecnico(datos, 'ma_obra.horas_ayudante')],
    ['Mà d’obra', 'Hores oficina', valorTecnico(datos, 'ma_obra.horas_oficina')],
    ['Mà d’obra', 'Desplaçaments', valorTecnico(datos, 'ma_obra.numero_desplazamientos')]
  ];

  filas.forEach(([seccion, campo, valor]) => {
    escribirFilaTecnica(ws, fila, seccion, campo, valor);
    fila += 1;
  });

  fila += 2;

  ws.mergeCells(`A${fila}:D${fila}`);
  ws.getCell(`A${fila}`).value = 'AVISOS I LÍNIES PENDENTS DE REVISIÓ';
  ws.getCell(`A${fila}`).font = {
    bold: true,
    color: { argb: 'FFFFFF' }
  };
  ws.getCell(`A${fila}`).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '9C6500' }
  };

  fila += 1;

  ws.getCell(`A${fila}`).value = 'Referència';
  ws.getCell(`B${fila}`).value = 'Descripció';
  ws.getCell(`C${fila}`).value = 'Motiu';
  ws.getCell(`D${fila}`).value = 'Total';

  ['A', 'B', 'C', 'D'].forEach(col => {
    ws.getCell(`${col}${fila}`).font = {
      bold: true,
      color: { argb: 'FFFFFF' }
    };
    ws.getCell(`${col}${fila}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'C65911' }
    };
    aplicarBordes(ws.getCell(`${col}${fila}`));
  });

  fila += 1;

  if (!avisos.length) {
    ws.mergeCells(`A${fila}:D${fila}`);
    ws.getCell(`A${fila}`).value = 'No hi ha avisos de revisió.';
    ws.getCell(`A${fila}`).font = {
      italic: true,
      color: { argb: '666666' }
    };
  } else {
    avisos.forEach(aviso => {
      ws.getCell(`A${fila}`).value = aviso.referencia || '';
      ws.getCell(`B${fila}`).value = aviso.descripcion || '';
      ws.getCell(`C${fila}`).value = aviso.motivo || '';
      ws.getCell(`D${fila}`).value = numero(aviso.total, 0);
      ws.getCell(`D${fila}`).numFmt = '#,##0.00 €';

      ['A', 'B', 'C', 'D'].forEach(col => {
        aplicarBordes(ws.getCell(`${col}${fila}`), 'E7EDF8');
        ws.getCell(`${col}${fila}`).alignment = {
          wrapText: true,
          vertical: 'top'
        };
      });

      fila += 1;
    });
  }
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
  prepararFilasTabla(ws, FILAS.trabajos.inicio, FILAS.trabajos.fin);
  prepararFilasTabla(ws, FILAS.manoObra.inicio, FILAS.manoObra.fin);

  return workbook;
}

function escribir(ws, celda, valor) {
  ws.getCell(celda).value = valor;
}

function mapearCabecera(ws, presupuesto) {
  escribir(
    ws,
    'G3',
    texto(primerValor(presupuesto, ['numero_presupuesto', 'numero', 'id']))
  );

  const fecha = primerValor(
    presupuesto,
    ['fecha', 'created_at', 'fecha_presupuesto'],
    ''
  );

  escribir(ws, 'I3', fecha ? new Date(fecha) : new Date());
  ws.getCell('I3').numFmt = 'dd/mm/yyyy';

  escribir(
    ws,
    'G4',
    texto(primerValor(presupuesto, ['estado'], 'Pendent de revisió'))
  );

  escribir(
    ws,
    'B7',
    texto(primerValor(presupuesto, ['cliente_nombre', 'cliente', 'nombre_cliente']))
  );

  escribir(ws, 'F7', texto(primerValor(presupuesto, ['empresa'])));
  escribir(ws, 'B8', texto(primerValor(presupuesto, ['telefono', 'teléfono', 'telefon'])));
  escribir(ws, 'F8', texto(primerValor(presupuesto, ['email', 'correo'])));
  escribir(ws, 'B9', texto(primerValor(presupuesto, ['direccion', 'dirección', 'adreca'])));

  escribir(
    ws,
    'B10',
    texto(primerValor(presupuesto, ['tipo_instalacion', 'tipo']))
  );

  escribir(ws, 'F10', texto(primerValor(presupuesto, ['marca'])));
  escribir(ws, 'H10', texto(primerValor(presupuesto, ['gama'])));

  escribir(
    ws,
    'B11',
    numero(primerValor(presupuesto, ['numero_unidades', 'unidades']), '')
  );

  escribir(
    ws,
    'D11',
    numero(primerValor(presupuesto, ['superficie', 'm2', 'metros_cuadrados']), '')
  );

  escribir(
    ws,
    'F11',
    numero(primerValor(presupuesto, ['metros_tuberia', 'metros_tubería']), '')
  );

  escribir(
    ws,
    'H11',
    texto(primerValor(presupuesto, ['fecha_visita', 'data_visita']))
  );

  escribir(
    ws,
    'B12',
    observacionesPrincipales(presupuesto)
  );
}

function valorTipoNormalizado(linea) {
  return normalizar(
    primerValor(linea, ['tipo', 'categoria', 'familia'], '')
  );
}

function clasificarLinea(linea, indice) {
  const codigo = texto(primerValor(linea, ['codigo', 'referencia', 'ref'], ''));

  if (codigo.startsWith('MANUAL-')) {
    return 'trabajo';
  }

  if (linea.requiere_revision === true && texto(linea.origen_precio) === 'manual_formulario') {
    return 'trabajo';
  }

  const tipo = valorTipoNormalizado(linea);

  if (
    tipo.includes('EQUIPO') ||
    tipo.includes('EQUIP') ||
    tipo.includes('MAQUINA') ||
    tipo.includes('MÀQUINA')
  ) {
    return 'equipo';
  }

  if (
    tipo.includes('OBRA') ||
    tipo.includes('ALTRES') ||
    tipo.includes('OTROS') ||
    tipo.includes('TREBALL')
  ) {
    return 'trabajo';
  }

  if (
    tipo.includes('MATERIAL') ||
    tipo.includes('MATERIALS') ||
    tipo.includes('COMPLEMENTARI')
  ) {
    return 'material';
  }

  return indice === 0 ? 'equipo' : 'material';
}

function descripcionLinea(linea) {
  return texto(
    primerValor(
      linea,
      [
        'descripcion',
        'descripcion_manual',
        'descripcio',
        'descripción',
        'nombre',
        'referencia'
      ]
    )
  );
}

function codigoLinea(linea) {
  return texto(
    primerValor(
      linea,
      ['codigo', 'código', 'codigo_saltoki', 'referencia', 'ref']
    )
  );
}

function unidadLinea(linea, defecto = 'ud') {
  return texto(primerValor(linea, ['unidad', 'unitat'], defecto));
}

function descuentoLinea(linea) {
  return porcentaje(
    primerValor(
      linea,
      [
        'descuento_pct',
        'descompte_pct',
        'descuento_porcentaje',
        'descompte_porcentaje'
      ],
      0
    ),
    0
  );
}

function precioUnitarioLinea(linea) {
  return numero(
    primerValor(
      linea,
      [
        'precio_unitario',
        'preu_unitari',
        'precio',
        'preu',
        'precio_eur'
      ],
      0
    ),
    0
  );
}

function cantidadLinea(linea) {
  return numero(
    primerValor(linea, ['cantidad', 'quantitat'], 1),
    1
  );
}

function notasLinea(linea) {
  return texto(
    primerValor(
      linea,
      ['notas', 'notes', 'motivo', 'justificacion', 'justificacio'],
      ''
    )
  );
}

function escribirLinea(ws, fila, linea, tipoPorDefecto) {
  escribir(
    ws,
    `A${fila}`,
    texto(primerValor(linea, ['tipo', 'categoria', 'familia'], tipoPorDefecto))
  );

  escribir(ws, `B${fila}`, codigoLinea(linea));
  escribir(ws, `C${fila}`, descripcionLinea(linea));
  escribir(ws, `D${fila}`, cantidadLinea(linea));
  escribir(ws, `E${fila}`, unidadLinea(linea));
  escribir(ws, `F${fila}`, precioUnitarioLinea(linea));
  escribir(ws, `G${fila}`, descuentoLinea(linea));
  escribir(ws, `I${fila}`, notasLinea(linea));

  ws.getCell(`F${fila}`).numFmt = '#,##0.00 €';
  ws.getCell(`G${fila}`).numFmt = '0.00%';
  ws.getCell(`H${fila}`).numFmt = '#,##0.00 €';

  if (codigoLinea(linea).startsWith('MANUAL-') || linea.requiere_revision === true) {
    for (let c = 1; c <= 9; c++) {
      ws.getCell(fila, c).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF2CC' }
      };
    }
  }
}

function siguienteFilaLibre(filaActual, limite, bloque) {
  if (filaActual > limite) {
    throw new Error(`La plantilla no tiene más filas disponibles para ${bloque}.`);
  }

  return filaActual;
}

function mapearLineas(ws, lineas, presupuesto) {
  const entrada = (Array.isArray(lineas) ? lineas : [])
    .filter(linea => linea && typeof linea === 'object')
    .map(linea => ({
      ...linea,
      marca: linea.marca ?? presupuesto.marca,
      gama: linea.gama ?? presupuesto.gama
    }));

  if (entrada.length === 0 && presupuesto.marca) {
    entrada.push({
      tipo: 'equipo',
      descripcion: `${presupuesto.marca} ${presupuesto.gama || ''}`.trim(),
      cantidad: primerValor(presupuesto, ['numero_unidades'], 1),
      unidad: 'ud',
      precio_unitario: primerValor(
        presupuesto,
        ['subtotal_materiales', 'subtotal'],
        ''
      )
    });
  }

  let filaEquipo = FILAS.equipo.inicio;
  let filaMaterial = FILAS.materiales.inicio;
  let filaTrabajo = FILAS.trabajos.inicio;

  entrada.forEach((linea, indice) => {
    const clasificacion = clasificarLinea(linea, indice);

    if (clasificacion === 'equipo') {
      filaEquipo = siguienteFilaLibre(
        filaEquipo,
        FILAS.equipo.fin,
        'equip principal'
      );

      escribirLinea(
        ws,
        filaEquipo,
        { ...linea, tipo: primerValor(linea, ['tipo'], 'Equip') },
        'Equip'
      );

      filaEquipo += 1;
      return;
    }

    if (clasificacion === 'trabajo') {
      filaTrabajo = siguienteFilaLibre(
        filaTrabajo,
        FILAS.trabajos.fin,
        'treballs addicionals'
      );

      escribirLinea(
        ws,
        filaTrabajo,
        { ...linea, tipo: 'Treball / revisió' },
        'Treball / revisió'
      );

      filaTrabajo += 1;
      return;
    }

    filaMaterial = siguienteFilaLibre(
      filaMaterial,
      FILAS.materiales.fin,
      'materials complementaris'
    );

    escribirLinea(
      ws,
      filaMaterial,
      { ...linea, tipo: primerValor(linea, ['tipo'], 'Material') },
      'Material'
    );

    filaMaterial += 1;
  });
}

function tarifaDesdeCosteTotal(presupuesto, campoTotal, campoHoras, tarifaDefecto) {
  const total = numero(presupuesto[campoTotal], null);
  const horas = numero(presupuesto[campoHoras], null);

  if (total !== null && horas !== null && horas > 0) {
    return total / horas;
  }

  return tarifaDefecto;
}

function mapearManoObraYOtros(ws, presupuesto) {
  const lineas = [
    {
      tipo: "Mà d'obra",
      descripcion: 'Hores Oficial 1a',
      cantidad: numero(
        primerValor(presupuesto, ['horas_oficial_primera', 'hores_oficial_primera'], 0),
        0
      ),
      unidad: 'h',
      precio_unitario: numero(
        primerValor(
          presupuesto,
          ['tarifa_oficial', 'precio_hora_oficial'],
          tarifaDesdeCosteTotal(
            presupuesto,
            'coste_oficial',
            'horas_oficial_primera',
            TARIFA_OFICIAL_DEFECTO
          )
        ),
        TARIFA_OFICIAL_DEFECTO
      )
    },
    {
      tipo: "Mà d'obra",
      descripcion: 'Hores Oficial 2a',
      cantidad: numero(
        primerValor(presupuesto, ['horas_oficial_segunda', 'hores_oficial_segona'], 0),
        0
      ),
      unidad: 'h',
      precio_unitario: numero(
        primerValor(
          presupuesto,
          ['tarifa_oficial_segunda', 'precio_hora_oficial_segunda'],
          tarifaDesdeCosteTotal(
            presupuesto,
            'coste_oficial_segunda',
            'horas_oficial_segunda',
            TARIFA_OFICIAL_DEFECTO
          )
        ),
        TARIFA_OFICIAL_DEFECTO
      )
    },
    {
      tipo: "Mà d'obra",
      descripcion: 'Hores Ajudant',
      cantidad: numero(
        primerValor(presupuesto, ['horas_ayudante', 'hores_ajudant'], 0),
        0
      ),
      unidad: 'h',
      precio_unitario: numero(
        primerValor(
          presupuesto,
          ['tarifa_ayudante', 'precio_hora_ayudante'],
          tarifaDesdeCosteTotal(
            presupuesto,
            'coste_ayudante',
            'horas_ayudante',
            TARIFA_AYUDANTE_DEFECTO
          )
        ),
        TARIFA_AYUDANTE_DEFECTO
      )
    },
    {
      tipo: "Mà d'obra",
      descripcion: 'Hores Oficina',
      cantidad: numero(
        primerValor(presupuesto, ['horas_oficina', 'hores_oficina'], 0),
        0
      ),
      unidad: 'h',
      precio_unitario: numero(
        primerValor(
          presupuesto,
          ['tarifa_oficina', 'precio_hora_oficina'],
          tarifaDesdeCosteTotal(
            presupuesto,
            'coste_oficina',
            'horas_oficina',
            TARIFA_OFICINA_DEFECTO
          )
        ),
        TARIFA_OFICINA_DEFECTO
      )
    },
    {
      tipo: 'Altres',
      descripcion: 'Desplaçament',
      cantidad: numero(
        primerValor(
          presupuesto,
          ['numero_desplazamientos', 'desplazamientos', 'desplacaments'],
          1
        ),
        1
      ),
      unidad: 'ud',
      precio_unitario: numero(
        primerValor(
          presupuesto,
          ['tarifa_desplazamiento', 'precio_desplazamiento'],
          tarifaDesdeCosteTotal(
            presupuesto,
            'coste_desplazamientos',
            'numero_desplazamientos',
            TARIFA_DESPLAZAMIENTO_DEFECTO
          )
        ),
        TARIFA_DESPLAZAMIENTO_DEFECTO
      )
    }
  ];

  lineas.forEach((linea, indice) => {
    escribirLinea(
      ws,
      FILAS.manoObra.inicio + indice,
      linea,
      linea.tipo
    );
  });
}

function mapearPorcentajesResumen(ws, presupuesto) {
  const subtotalMateriales = numero(presupuesto.subtotal_materiales, null);
  const descuentoMateriales = numero(presupuesto.descuento_materiales, null);

  let descuentoPct = porcentaje(
    primerValor(
      presupuesto,
      ['descuento_materiales_pct', 'descompte_materials_pct'],
      null
    ),
    null
  );

  if (
    (descuentoPct === null || descuentoPct === '') &&
    subtotalMateriales &&
    descuentoMateriales
  ) {
    descuentoPct = descuentoMateriales / subtotalMateriales;
  }

  escribir(ws, `H${RESUMEN.descuentoMaterials}`, descuentoPct === null || descuentoPct === '' ? 0 : descuentoPct);

  escribir(
    ws,
    `H${RESUMEN.incrementoMaterial}`,
    porcentaje(
      primerValor(
        presupuesto,
        ['incremento_material_pct', 'increment_material_pct'],
        0
      ),
      0
    )
  );

  escribir(
    ws,
    `H${RESUMEN.iva}`,
    porcentaje(
      primerValor(presupuesto, ['iva_pct', 'iva_porcentaje'], IVA_DEFECTO),
      IVA_DEFECTO
    )
  );
}

function valorNumericoCelda(ws, celda, defecto = 0) {
  const valor = ws.getCell(celda).value;

  if (
    valor &&
    typeof valor === 'object' &&
    Object.prototype.hasOwnProperty.call(valor, 'result')
  ) {
    return numero(valor.result, defecto);
  }

  return numero(valor, defecto);
}

function ponerFormulaConResultado(ws, celda, formula, resultado, numFmt) {
  ws.getCell(celda).value = {
    formula,
    result: resultado
  };

  if (numFmt) {
    ws.getCell(celda).numFmt = numFmt;
  }
}

function recalcularResultadosFormulas(ws) {
  const rangos = [
    FILAS.equipo,
    FILAS.materiales,
    FILAS.trabajos,
    FILAS.manoObra
  ];

  for (const rango of rangos) {
    for (let r = rango.inicio; r <= rango.fin; r++) {
      const cantidad = valorNumericoCelda(ws, `D${r}`, null);
      const precio = valorNumericoCelda(ws, `F${r}`, null);
      const descuento = valorNumericoCelda(ws, `G${r}`, 0);

      const formula =
        `IF(OR(D${r}="",F${r}=""),"",ROUND(D${r}*F${r}*(1-IF(G${r}="",0,G${r})),2))`;

      const resultado =
        cantidad === null || precio === null
          ? ''
          : Math.round(cantidad * precio * (1 - descuento) * 100) / 100;

      ponerFormulaConResultado(
        ws,
        `H${r}`,
        formula,
        resultado,
        '#,##0.00 €'
      );
    }
  }

  const sumarResultados = (inicio, fin) => {
    let total = 0;

    for (let r = inicio; r <= fin; r++) {
      total += valorNumericoCelda(ws, `H${r}`, 0);
    }

    return Math.round(total * 100) / 100;
  };

  const subtotalEquip = sumarResultados(
    FILAS.equipo.inicio,
    FILAS.equipo.fin
  );

  const subtotalMaterials = sumarResultados(
    FILAS.materiales.inicio,
    FILAS.materiales.fin
  );

  const subtotalTreballs = sumarResultados(
    FILAS.trabajos.inicio,
    FILAS.trabajos.fin
  );

  const subtotalMaObra = sumarResultados(
    FILAS.manoObra.inicio,
    FILAS.manoObra.fin
  );

  const pctDescompte = valorNumericoCelda(ws, `H${RESUMEN.descuentoMaterials}`, 0);
  const descompteMaterials =
    Math.round(subtotalMaterials * pctDescompte * 100) / 100;

  const subtotalMaterialsFinal =
    Math.round((subtotalMaterials - descompteMaterials) * 100) / 100;

  const pctIncrement = valorNumericoCelda(ws, `H${RESUMEN.incrementoMaterial}`, 0);
  const incrementMaterial =
    Math.round(subtotalMaterialsFinal * pctIncrement * 100) / 100;

  const baseImposable =
    Math.round(
      (
        subtotalEquip +
        subtotalMaterialsFinal +
        subtotalTreballs +
        subtotalMaObra +
        incrementMaterial
      ) * 100
    ) / 100;

  const pctIva = valorNumericoCelda(ws, `H${RESUMEN.iva}`, IVA_DEFECTO);
  const iva = Math.round(baseImposable * pctIva * 100) / 100;
  const total = Math.round((baseImposable + iva) * 100) / 100;

  ponerFormulaConResultado(ws, `I${RESUMEN.subtotalEquip}`, `SUM(H${FILAS.equipo.inicio}:H${FILAS.equipo.fin})`, subtotalEquip, '#,##0.00 €');
  ponerFormulaConResultado(ws, `I${RESUMEN.subtotalMaterials}`, `SUM(H${FILAS.materiales.inicio}:H${FILAS.materiales.fin})`, subtotalMaterials, '#,##0.00 €');
  ponerFormulaConResultado(ws, `I${RESUMEN.descuentoMaterials}`, `ROUND(I${RESUMEN.subtotalMaterials}*H${RESUMEN.descuentoMaterials},2)`, descompteMaterials, '#,##0.00 €');
  ponerFormulaConResultado(ws, `I${RESUMEN.subtotalMaterialsFinal}`, `I${RESUMEN.subtotalMaterials}-I${RESUMEN.descuentoMaterials}`, subtotalMaterialsFinal, '#,##0.00 €');
  ponerFormulaConResultado(ws, `I${RESUMEN.subtotalTrabajos}`, `SUM(H${FILAS.trabajos.inicio}:H${FILAS.trabajos.fin})`, subtotalTreballs, '#,##0.00 €');
  ponerFormulaConResultado(ws, `I${RESUMEN.subtotalMaObra}`, `SUM(H${FILAS.manoObra.inicio}:H${FILAS.manoObra.fin})`, subtotalMaObra, '#,##0.00 €');
  ponerFormulaConResultado(ws, `I${RESUMEN.incrementoMaterial}`, `ROUND(I${RESUMEN.subtotalMaterialsFinal}*H${RESUMEN.incrementoMaterial},2)`, incrementMaterial, '#,##0.00 €');
  ponerFormulaConResultado(ws, `I${RESUMEN.baseImposable}`, `SUM(I${RESUMEN.subtotalEquip},I${RESUMEN.subtotalMaterialsFinal},I${RESUMEN.subtotalTrabajos},I${RESUMEN.subtotalMaObra},I${RESUMEN.incrementoMaterial})`, baseImposable, '#,##0.00 €');
  ponerFormulaConResultado(ws, `I${RESUMEN.iva}`, `ROUND(I${RESUMEN.baseImposable}*H${RESUMEN.iva},2)`, iva, '#,##0.00 €');
  ponerFormulaConResultado(ws, `I${RESUMEN.total}`, `I${RESUMEN.baseImposable}+I${RESUMEN.iva}`, total, '#,##0.00 €');
}

// =====================================================
// SALTOKI ONLINE - CONSULTA PRECIOS REALES BURGAS
// =====================================================

app.post('/saltoki/precios', async (req, res) => {
  try {
    const apiKey = req.headers['x-api-key'];

    if (process.env.SALTOKI_API_KEY && apiKey !== process.env.SALTOKI_API_KEY) {
      return res.status(401).json({
        ok: false,
        error: 'No autorizado'
      });
    }

    const codigosInput = req.body?.codigos;

    if (!Array.isArray(codigosInput) || codigosInput.length === 0) {
      return res.status(400).json({
        ok: false,
        error: 'El body debe incluir codigos como array',
        ejemplo: {
          codigos: ['4050056112', '4050056712']
        }
      });
    }

    const codigos = [...new Set(
      codigosInput
        .map(codigo => String(codigo || '').trim())
        .filter(Boolean)
    )];

    const resultados = [];
    const errores = [];

    for (const codigo of codigos) {
      try {
        const precio = await obtenerPrecioSaltoki(codigo);
        resultados.push(precio);
      } catch (error) {
        errores.push({
          codigo,
          error: error.message
        });
      }

      await sleep(300);
    }

    return res.json({
      ok: true,
      precios: resultados,
      errores,
      total_codigos: codigos.length,
      total_precios: resultados.length,
      fecha_consulta: new Date().toISOString()
    });

  } catch (error) {
    console.error('[POST /saltoki/precios]', error);

    return res.status(500).json({
      ok: false,
      error: 'Error general consultando precios Saltoki',
      detalle: error.message
    });
  }
});

app.post('/saltoki/precio', async (req, res) => {
  try {
    const apiKey = req.headers['x-api-key'];

    if (process.env.SALTOKI_API_KEY && apiKey !== process.env.SALTOKI_API_KEY) {
      return res.status(401).json({
        ok: false,
        error: 'No autorizado'
      });
    }

    const codigo = String(req.body?.codigo || '').trim();

    if (!codigo) {
      return res.status(400).json({
        ok: false,
        error: 'Falta el campo codigo'
      });
    }

    const precio = await obtenerPrecioSaltoki(codigo);

    return res.json({
      ok: true,
      ...precio
    });

  } catch (error) {
    console.error('[POST /saltoki/precio]', error);

    return res.status(500).json({
      ok: false,
      error: 'Error consultando precio Saltoki',
      detalle: error.message
    });
  }
});

async function obtenerPrecioSaltoki(codigo) {
  const productUrl = `https://online.saltoki.com/producto/x/${encodeURIComponent(codigo)}`;

  const fichaResponse = await fetch(productUrl, {
    method: 'GET',
    headers: buildSaltokiHeaders({
      accept: 'text/html, application/xhtml+xml',
      referer: 'https://online.saltoki.com/'
    })
  });

  if (!fichaResponse.ok) {
    throw new Error(`No se pudo cargar ficha producto ${codigo}. Status: ${fichaResponse.status}`);
  }

  const fichaHtml = await fichaResponse.text();

  if (isLoginPage(fichaHtml)) {
    throw new Error('Sesión Saltoki caducada o no autenticada');
  }

  const tokenMatch = fichaHtml.match(/data-prices-skus-and-quantities-value="([^"]+)"/);

  if (!tokenMatch) {
    throw new Error(`No se encontró token de precios para ${codigo}`);
  }

  const pricesToken = tokenMatch[1];

  const priceUrl = `https://online.saltoki.com/prices/${pricesToken}`;

  const priceResponse = await fetch(priceUrl, {
    method: 'GET',
    headers: buildSaltokiHeaders({
      accept: 'application/json, text/plain, */*',
      referer: productUrl,
      requestedWith: true
    })
  });

  if (!priceResponse.ok) {
    throw new Error(`No se pudo consultar /prices para ${codigo}. Status: ${priceResponse.status}`);
  }

  const priceHtml = await priceResponse.text();

  if (isLoginPage(priceHtml)) {
    throw new Error('Sesión Saltoki caducada al consultar precios');
  }

  const precios = parseSaltokiPrices(priceHtml);
  const precioProducto = precios.find(p => p.codigo === codigo);

  if (!precioProducto) {
    throw new Error(`No se encontró precio neto para ${codigo}`);
  }

  return {
    ...precioProducto,
    product_url: productUrl
  };
}

function buildSaltokiHeaders({ accept, referer, requestedWith = false }) {
  if (!process.env.SALTOKI_COOKIE) {
    throw new Error('Falta variable SALTOKI_COOKIE en Railway');
  }

  const headers = {
    'Accept': accept,
    'Cookie': process.env.SALTOKI_COOKIE,
    'User-Agent': process.env.SALTOKI_USER_AGENT || 'Mozilla/5.0',
    'Referer': referer || 'https://online.saltoki.com/',
    'Accept-Language': 'es-ES,es;q=0.9,en;q=0.8'
  };

  if (requestedWith) {
    headers['X-Requested-With'] = 'XMLHttpRequest';
  }

  if (process.env.SALTOKI_CSRF_TOKEN) {
    headers['X-CSRF-TOKEN'] = process.env.SALTOKI_CSRF_TOKEN;
  }

  if (process.env.SALTOKI_XSRF_TOKEN) {
    headers['X-XSRF-TOKEN'] = process.env.SALTOKI_XSRF_TOKEN;
  }

  return headers;
}

function parseSaltokiPrices(html) {
  const prices = [];
  const used = new Set();

  const priceRegex = /data-sku="([^"]+)"[\s\S]*?data-net-price="([^"]+)"/g;

  let match;

  while ((match = priceRegex.exec(html)) !== null) {
    const codigo = String(match[1]).trim();
    const precioNeto = Number(String(match[2]).replace(',', '.'));

    if (!codigo || Number.isNaN(precioNeto)) {
      continue;
    }

    if (used.has(codigo)) {
      continue;
    }

    used.add(codigo);

    const descuentoPct = extractDiscountForSku(html, codigo);
    const precioTotalLinea = extractTotalForSku(html, codigo);

    prices.push({
      codigo,
      precio_neto: precioNeto,
      precio_unitario: precioNeto,
      precio_total_linea: precioTotalLinea,
      descuento_pct: descuentoPct,
      moneda: 'EUR',
      fuente_precio: 'saltoki_online',
      fecha_actualizacion: new Date().toISOString()
    });
  }

  return prices;
}

function extractDiscountForSku(html, codigo) {
  const skuIndex = html.indexOf(`data-sku="${codigo}"`);

  if (skuIndex === -1) {
    return null;
  }

  const previousChunk = html.slice(Math.max(0, skuIndex - 700), skuIndex);
  const discountMatch = previousChunk.match(/Descuento:<\/span>\s*-([0-9]+(?:[,.][0-9]+)?)%/);

  if (!discountMatch) {
    return null;
  }

  const value = Number(String(discountMatch[1]).replace(',', '.'));

  return Number.isFinite(value) ? value : null;
}

function extractTotalForSku(html, codigo) {
  const regex = new RegExp(
    `data-sku="${codigo}"[\\s\\S]*?data-total-net-price="([^"]+)"`,
    'i'
  );

  const match = html.match(regex);

  if (!match) {
    return null;
  }

  const value = Number(String(match[1]).replace(',', '.'));

  return Number.isFinite(value) ? value : null;
}

function isLoginPage(html) {
  const text = String(html || '').toLowerCase();

  return (
    text.includes('login') &&
    (
      text.includes('password') ||
      text.includes('contraseña') ||
      text.includes('iniciar sesión') ||
      text.includes('correo electrónico')
    )
  );
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

app.get('/', (req, res) => {
  res.send('Burgas Excel Generator funcionando - v6.5 secciones ampliadas');
});

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
    crearHojaDatosTecnicos(workbook, presupuesto);

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );

    res.setHeader(
      'Content-Disposition',
      'attachment; filename=pressupost-burgas.xlsx'
    );

    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error('[POST /generar]', error);
    res.status(500).json({
      error: error.message
    });
  }
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log(`Servidor iniciado en puerto ${PORT}`);
});
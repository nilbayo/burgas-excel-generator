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
  crearSeccionTabla(ws, 50, 'MÀ D’OBRA I ALTRES');
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
    fgColor: { argb: '1F3763' }
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
  ws.mergeCells('G64:I64');
  ws.getCell('G64').value = 'RESUM ECONÒMIC';
  ws.getCell('G64').font = {
    bold: true,
    color: { argb: 'FFFFFF' }
  };
  ws.getCell('G64').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '1F3763' }
  };

  const rows = [
    [65, 'Subtotal equip', '', 'SUM(H16:H20)'],
    [66, 'Subtotal materials', '', 'SUM(H24:H48)'],
    [67, 'Descompte materials', 0, 'ROUND(I66*H67,2)'],
    [68, 'Subtotal materials final', '', 'I66-I67'],
    [69, 'Subtotal mà d’obra i altres', '', 'SUM(H52:H61)'],
    [70, 'Increment material', 0, 'ROUND(I68*H70,2)'],
    [71, 'Base imposable', '', 'SUM(I65,I68,I69,I70)'],
    [72, 'IVA', IVA_DEFECTO, 'ROUND(I71*H72,2)'],
    [73, 'TOTAL PRESSUPOST', '', 'I71+I72']
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

    ws.getCell(`G${r}`).font = { bold: r === 73 };
    ws.getCell(`I${r}`).font = {
      bold: r === 73,
      size: r === 73 ? 14 : 11
    };

    ws.getCell(`H${r}`).numFmt = '0.00%';
    ws.getCell(`I${r}`).numFmt = '#,##0.00 €';

    if (r === 73) {
      ['G', 'H', 'I'].forEach(col => {
        ws.getCell(`${col}${r}`).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'D9EAF7' }
        };
      });
    }
  });

  ws.mergeCells('A65:F73');
  ws.getCell('A65').value =
    'Document generat automàticament. Revisar i ajustar imports si és necessari abans d’enviar al client.';
  ws.getCell('A65').font = {
    italic: true,
    color: { argb: '666666' }
  };
  ws.getCell('A65').alignment = {
    wrapText: true,
    vertical: 'top'
  };
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
    texto(primerValor(presupuesto, ['observaciones', 'observacions', 'notas']))
  );
}

function valorTipoNormalizado(linea) {
  return normalizar(
    primerValor(linea, ['tipo', 'categoria', 'familia'], '')
  );
}

function clasificarLinea(linea, indice) {
  const tipo = valorTipoNormalizado(linea);

  if (
    tipo.includes('MATERIAL') ||
    tipo.includes('MATERIALS') ||
    tipo.includes('COMPLEMENTARI')
  ) {
    return 'material';
  }

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
    tipo.includes('OTROS')
  ) {
    return 'otros';
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

  entrada.forEach((linea, indice) => {
    const clasificacion = clasificarLinea(linea, indice);

    if (clasificacion === 'otros') {
      return;
    }

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

  escribir(ws, 'H67', descuentoPct === null || descuentoPct === '' ? 0 : descuentoPct);

  escribir(
    ws,
    'H70',
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
    'H72',
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

  const pctDescompte = valorNumericoCelda(ws, 'H67', 0);
  const descompteMaterials =
    Math.round(subtotalMaterials * pctDescompte * 100) / 100;

  const subtotalMaterialsFinal =
    Math.round((subtotalMaterials - descompteMaterials) * 100) / 100;

  const subtotalMaObra = sumarResultados(
    FILAS.manoObra.inicio,
    FILAS.manoObra.fin
  );

  const pctIncrement = valorNumericoCelda(ws, 'H70', 0);
  const incrementMaterial =
    Math.round(subtotalMaterialsFinal * pctIncrement * 100) / 100;

  const baseImposable =
    Math.round(
      (
        subtotalEquip +
        subtotalMaterialsFinal +
        subtotalMaObra +
        incrementMaterial
      ) * 100
    ) / 100;

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

app.get('/', (req, res) => {
  res.send('Burgas Excel Generator funcionando - v5');
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

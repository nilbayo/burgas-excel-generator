const path = require('path');
const express = require('express');
const ExcelJS = require('exceljs');

const app = express();

app.use(express.json({ limit: '10mb' }));

const PLANTILLA = path.join(__dirname, 'plantilla.xlsx');

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

const CELDAS = {
  numero: 'G3',
  fecha: 'I3',
  estado: 'G4',
  ivaCabecera: 'I4',
  cliente: 'B7',
  empresa: 'F7',
  telefono: 'B8',
  email: 'F8',
  direccion: 'B9',
  tipoInstalacion: 'B10',
  marca: 'F10',
  gama: 'H10',
  unidades: 'B11',
  superficie: 'D11',
  metrosTuberia: 'F11',
  fechaVisita: 'H11',
  observaciones: 'B12',
  descuentoMaterialesPct: 'H67',
  incrementoMaterialPct: 'H70',
  ivaPct: 'H72'
};

function texto(valor) {
  return valor === null || valor === undefined ? '' : String(valor);
}

function normalizar(valor) {
  return texto(valor)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase();
}

function numero(valor, defecto = '') {
  if (valor === null || valor === undefined || valor === '') return defecto;
  if (typeof valor === 'number') return Number.isFinite(valor) ? valor : defecto;

  const raw = String(valor).trim().replace(/\s/g, '');
  if (!raw) return defecto;

  const normalizado = raw.includes(',')
    ? raw.replace(/\./g, '').replace(',', '.')
    : raw.replace(/[^0-9.-]/g, '');

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
    if (objeto && objeto[campo] !== undefined && objeto[campo] !== null && objeto[campo] !== '') {
      return objeto[campo];
    }
  }
  return defecto;
}

function escribir(ws, celda, valor) {
  ws.getCell(celda).value = valor;
}

function crearErrorUsuario(mensaje) {
  const error = new Error(mensaje);
  error.statusCode = 400;
  return error;
}

function limpiarRango(ws, inicio, fin) {
  for (let fila = inicio; fila <= fin; fila += 1) {
    // A:G e I son celdas editables/de entrada. H mantiene la fórmula del total de línea.
    ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'I'].forEach((columna) => {
      ws.getCell(`${columna}${fila}`).value = '';
    });
  }
}

function restaurarFormulasLineas(ws, inicio, fin) {
  for (let fila = inicio; fila <= fin; fila += 1) {
    ws.getCell(`H${fila}`).value = {
      formula: `IF(OR(D${fila}="",F${fila}=""),"",ROUND(D${fila}*F${fila}*(1-IF(G${fila}="",0,G${fila})),2))`
    };
  }
}

function restaurarFormulasResumen(ws) {
  ws.getCell('I65').value = { formula: 'SUM(H16:H20)' };
  ws.getCell('I66').value = { formula: 'SUM(H24:H48)' };
  ws.getCell('I67').value = { formula: 'ROUND(I66*H67,2)' };
  ws.getCell('I68').value = { formula: 'I66-I67' };
  ws.getCell('I69').value = { formula: 'SUM(H52:H61)' };
  ws.getCell('I70').value = { formula: 'ROUND(I68*H70,2)' };
  ws.getCell('I71').value = { formula: 'SUM(I65,I68,I69,I70)' };
  ws.getCell('I72').value = { formula: 'ROUND(I71*H72,2)' };
  ws.getCell('I73').value = { formula: 'I71+I72' };
  ws.getCell(CELDAS.ivaCabecera).value = { formula: 'H72' };
}

function limpiarPlantilla(ws) {
  limpiarRango(ws, FILAS.equipo.inicio, FILAS.equipo.fin);
  limpiarRango(ws, FILAS.materiales.inicio, FILAS.materiales.fin);
  limpiarRango(ws, FILAS.manoObra.inicio, FILAS.manoObra.fin);

  restaurarFormulasLineas(ws, FILAS.equipo.inicio, FILAS.equipo.fin);
  restaurarFormulasLineas(ws, FILAS.materiales.inicio, FILAS.materiales.fin);
  restaurarFormulasLineas(ws, FILAS.manoObra.inicio, FILAS.manoObra.fin);
  restaurarFormulasResumen(ws);
}

function mapearCabecera(ws, presupuesto) {
  escribir(ws, CELDAS.numero, texto(primerValor(presupuesto, ['numero_presupuesto', 'numero', 'id'])));

  const fecha = primerValor(presupuesto, ['fecha', 'created_at', 'fecha_presupuesto'], '');
  escribir(ws, CELDAS.fecha, fecha ? new Date(fecha) : new Date());
  ws.getCell(CELDAS.fecha).numFmt = 'dd/mm/yyyy';

  escribir(ws, CELDAS.estado, texto(primerValor(presupuesto, ['estado'], 'Pendent de revisió')));
  escribir(ws, CELDAS.cliente, texto(primerValor(presupuesto, ['cliente_nombre', 'cliente', 'nombre_cliente'])));
  escribir(ws, CELDAS.empresa, texto(primerValor(presupuesto, ['empresa'])));
  escribir(ws, CELDAS.telefono, texto(primerValor(presupuesto, ['telefono', 'teléfono', 'telefon'])));
  escribir(ws, CELDAS.email, texto(primerValor(presupuesto, ['email', 'correo'])));
  escribir(ws, CELDAS.direccion, texto(primerValor(presupuesto, ['direccion', 'dirección', 'adreca'])));
  escribir(ws, CELDAS.tipoInstalacion, texto(primerValor(presupuesto, ['tipo_instalacion', 'tipo'])));
  escribir(ws, CELDAS.marca, texto(primerValor(presupuesto, ['marca'])));
  escribir(ws, CELDAS.gama, texto(primerValor(presupuesto, ['gama'])));
  escribir(ws, CELDAS.unidades, numero(primerValor(presupuesto, ['numero_unidades', 'unidades']), ''));
  escribir(ws, CELDAS.superficie, numero(primerValor(presupuesto, ['superficie', 'm2', 'metros_cuadrados']), ''));
  escribir(ws, CELDAS.metrosTuberia, numero(primerValor(presupuesto, ['metros_tuberia', 'metros_tubería']), ''));
  escribir(ws, CELDAS.fechaVisita, texto(primerValor(presupuesto, ['fecha_visita', 'data_visita'])));
  escribir(ws, CELDAS.observaciones, texto(primerValor(presupuesto, ['observaciones', 'observacions', 'notas'])));
}

function esLineaEquipo(linea) {
  const contenido = normalizar([
    linea.tipo,
    linea.categoria,
    linea.familia,
    linea.referencia,
    linea.codigo,
    linea.descripcion
  ].filter(Boolean).join(' '));

  return (
    contenido.includes('EQUIPO') ||
    contenido.includes('EQUIP') ||
    contenido.includes('MAQUINA') ||
    contenido.includes('MÀQUINA') ||
    contenido.includes('UNIDAD') ||
    contenido.includes('UNITAT') ||
    contenido.includes('SPLIT') ||
    contenido.includes('DAIKIN') ||
    contenido.includes('KOSNER') ||
    contenido.includes('MITSUBISHI')
  );
}

function descripcionLinea(linea) {
  return texto(primerValor(linea, ['descripcion', 'descripcio', 'descripción', 'nombre', 'referencia']));
}

function codigoLinea(linea) {
  return texto(primerValor(linea, ['codigo', 'código', 'codigo_saltoki', 'referencia', 'ref']));
}

function unidadLinea(linea, defecto = 'ud') {
  return texto(primerValor(linea, ['unidad', 'unitat'], defecto));
}

function tipoLinea(linea, defecto) {
  return texto(primerValor(linea, ['tipo', 'categoria', 'familia'], defecto));
}

function descuentoLinea(linea) {
  return porcentaje(primerValor(linea, ['descuento_pct', 'descompte_pct', 'descuento_porcentaje', 'descompte_porcentaje'], 0), 0);
}

function precioUnitarioLinea(linea) {
  return numero(primerValor(linea, ['precio_unitario', 'preu_unitari', 'precio', 'preu', 'precio_eur'], ''), '');
}

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

function siguienteFilaLibre(filaActual, limite, bloque) {
  if (filaActual > limite) {
    throw crearErrorUsuario(`La plantilla no tiene más filas disponibles para ${bloque}.`);
  }
  return filaActual;
}

function mapearLineas(ws, lineas, presupuesto) {
  const entrada = (Array.isArray(lineas) ? lineas : []).map((linea) => ({
    ...linea,
    marca: linea.marca ?? presupuesto.marca,
    gama: linea.gama ?? presupuesto.gama
  }));

  let filaEquipo = FILAS.equipo.inicio;
  let filaMaterial = FILAS.materiales.inicio;

  const indiceEquipoDetectado = entrada.findIndex((linea) => esLineaEquipo(linea));
  const indiceEquipo = indiceEquipoDetectado >= 0 ? indiceEquipoDetectado : 0;

  entrada.forEach((linea, indice) => {
    if (indice === indiceEquipo) {
      filaEquipo = siguienteFilaLibre(filaEquipo, FILAS.equipo.fin, 'equip principal');
      escribirLinea(ws, filaEquipo, linea, 'Equip');
      filaEquipo += 1;
      return;
    }

    filaMaterial = siguienteFilaLibre(filaMaterial, FILAS.materiales.fin, 'materials complementaris');
    escribirLinea(ws, filaMaterial, linea, 'Material');
    filaMaterial += 1;
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
    {
      tipo: 'Mà d\'obra',
      descripcion: 'Hores Oficial 1a',
      cantidad: numero(primerValor(presupuesto, ['horas_oficial_primera', 'hores_oficial_primera'], 0), 0),
      unidad: 'h',
      precio_unitario: numero(primerValor(presupuesto, ['tarifa_oficial', 'precio_hora_oficial'], tarifaDesdeCosteTotal(presupuesto, 'coste_oficial', 'horas_oficial_primera', TARIFA_OFICIAL_DEFECTO)), TARIFA_OFICIAL_DEFECTO)
    },
    {
      tipo: 'Mà d\'obra',
      descripcion: 'Hores Ajudant',
      cantidad: numero(primerValor(presupuesto, ['horas_ayudante', 'hores_ajudant'], 0), 0),
      unidad: 'h',
      precio_unitario: numero(primerValor(presupuesto, ['tarifa_ayudante', 'precio_hora_ayudante'], tarifaDesdeCosteTotal(presupuesto, 'coste_ayudante', 'horas_ayudante', TARIFA_AYUDANTE_DEFECTO)), TARIFA_AYUDANTE_DEFECTO)
    },
    {
      tipo: 'Mà d\'obra',
      descripcion: 'Hores Oficina',
      cantidad: numero(primerValor(presupuesto, ['horas_oficina', 'hores_oficina'], 0), 0),
      unidad: 'h',
      precio_unitario: numero(primerValor(presupuesto, ['tarifa_oficina', 'precio_hora_oficina'], tarifaDesdeCosteTotal(presupuesto, 'coste_oficina', 'horas_oficina', TARIFA_OFICINA_DEFECTO)), TARIFA_OFICINA_DEFECTO)
    },
    {
      tipo: 'Altres',
      descripcion: 'Desplaçament',
      cantidad: numero(primerValor(presupuesto, ['numero_desplazamientos', 'desplazamientos', 'desplacaments'], 1), 1),
      unidad: 'ud',
      precio_unitario: numero(primerValor(presupuesto, ['tarifa_desplazamiento', 'precio_desplazamiento'], tarifaDesdeCosteTotal(presupuesto, 'coste_desplazamientos', 'numero_desplazamientos', TARIFA_DESPLAZAMIENTO_DEFECTO)), TARIFA_DESPLAZAMIENTO_DEFECTO)
    }
  ];

  lineas.forEach((linea, indice) => {
    const fila = FILAS.manoObra.inicio + indice;
    escribirLinea(ws, fila, linea, linea.tipo);
  });
}

function mapearPorcentajesResumen(ws, presupuesto) {
  const subtotalMateriales = numero(presupuesto.subtotal_materiales, null);
  const descuentoMateriales = numero(presupuesto.descuento_materiales, null);

  let descuentoPct = porcentaje(primerValor(presupuesto, ['descuento_materiales_pct', 'descompte_materials_pct'], null), null);
  if ((descuentoPct === null || descuentoPct === '') && subtotalMateriales && descuentoMateriales) {
    descuentoPct = descuentoMateriales / subtotalMateriales;
  }

  escribir(ws, CELDAS.descuentoMaterialesPct, descuentoPct === null || descuentoPct === '' ? 0 : descuentoPct);
  escribir(ws, CELDAS.incrementoMaterialPct, porcentaje(primerValor(presupuesto, ['incremento_material_pct', 'increment_material_pct'], 0), 0));
  escribir(ws, CELDAS.ivaPct, porcentaje(primerValor(presupuesto, ['iva_pct', 'iva_porcentaje'], IVA_DEFECTO), IVA_DEFECTO));
}

function validarEntrada(body) {
  if (!body || typeof body !== 'object') {
    throw crearErrorUsuario('El cuerpo de la petición debe ser un objeto JSON.');
  }

  if (!body.presupuesto || typeof body.presupuesto !== 'object') {
    throw crearErrorUsuario('Falta el objeto presupuesto.');
  }

  if (body.lineas !== undefined && !Array.isArray(body.lineas)) {
    throw crearErrorUsuario('El campo lineas debe ser un array.');
  }

  if (Array.isArray(body.lineas)) {
    body.lineas.forEach((linea, indice) => {
      if (!linea || typeof linea !== 'object' || Array.isArray(linea)) {
        throw crearErrorUsuario(`La línea ${indice + 1} debe ser un objeto.`);
      }
    });
  }
}

app.get('/', (req, res) => {
  res.send('Burgas Excel Generator funcionando');
});

app.post('/generar', async (req, res) => {
  try {
    validarEntrada(req.body);

    const presupuesto = req.body.presupuesto;
    const lineas = req.body.lineas ?? [];

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(PLANTILLA);

    workbook.calcProperties.fullCalcOnLoad = true;
    workbook.calcProperties.forceFullCalc = true;

    const worksheet = workbook.getWorksheet('PRESSUPOST') || workbook.worksheets[0];

    limpiarPlantilla(worksheet);
    mapearCabecera(worksheet, presupuesto);
    mapearLineas(worksheet, lineas, presupuesto);
    mapearManoObraYOtros(worksheet, presupuesto);
    mapearPorcentajesResumen(worksheet, presupuesto);

    restaurarFormulasResumen(worksheet);

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

    const estado = error.statusCode || 500;

    res.status(estado).json({
      error: error.message
    });
  }
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log(`Servidor iniciado en puerto ${PORT}`);
});

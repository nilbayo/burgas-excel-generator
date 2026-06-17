const path = require('path');
const express = require('express');
const ExcelJS = require('exceljs');

const app = express();

app.use(express.json({ limit: '10mb' }));

const PLANTILLA = path.join(__dirname, 'plantilla.xlsx');

const COLUMNAS_LINEA = {
    cantidad: 'A',
    descripcion: 'B',
    precioUnitario: 'C',
    subtotal: 'E'
};

const TEXTOS = {
    equipo: 'Equip',
    subtotalMateriales: 'Subtotal materials',
    descuentoMateriales: 'Descompte materials',
    horasOficial: 'Hores Oficial 1a',
    horasAyudante: 'Hores Ajudant',
    horasOficina: 'Hores Oficina',
    desplazamiento: 'Desplaçament',
    incrementoMaterial: 'Increment material',
    subtotal: 'Subtotal',
    iva: 'IVA',
    total: 'Total',
    material: 'Material'
};

const BLOQUE_ACTIVO = {
    nombre: 'ACTIVO',
    filasLimpieza: rango(13, 31),
    equipo: {
        datos: 13,
        importe: 15
    },
    lineas: [16, 17, 18, 19, 20, 21],
    textos: {
        equipo: 'B13',
        subtotalMateriales: 'C23',
        descuentoMateriales: 'C24',
        horasOficial: 'B25',
        horasAyudante: 'B26',
        horasOficina: 'B27',
        desplazamiento: 'B28',
        incrementoMaterial: 'B29',
        subtotal: 'C30',
        iva: 'D31',
        total: 'C31'
    },
    costes: {
        subtotalMateriales: 'E23',
        descuentoMateriales: 'E24',
        horasOficial: 'A25',
        costeOficial: 'E25',
        horasAyudante: 'A26',
        costeAyudante: 'E26',
        horasOficina: 'A27',
        costeOficina: 'E27',
        desplazamientos: 'A28',
        costeDesplazamientos: 'E28',
        incrementoMaterialPct: 'C29',
        subtotalMaterialesFinal: 'E30',
        subtotal: 'E30',
        iva: 'F31',
        total: 'E31'
    }
};

const BLOQUE_INACTIVO = {
    nombre: 'INACTIVO',
    filasLimpieza: rango(35, 53)
};

function rango(inicio, fin) {
    return Array.from({ length: fin - inicio + 1 }, (_, indice) => inicio + indice);
}

function texto(valor) {
    return valor === null || valor === undefined ? '' : String(valor);
}

function numeroRecibido(valor) {
    if (valor === null || valor === undefined || valor === '') {
        return '';
    }

    const numero = Number(String(valor).replace(',', '.'));

    return Number.isFinite(numero) ? numero : valor;
}

function normalizar(valor) {
    return texto(valor)
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toUpperCase();
}

function valorPresupuesto(presupuesto, campo) {
    return numeroRecibido(presupuesto[campo]);
}

function escribir(worksheet, celda, valor) {
    worksheet.getCell(celda).value = valor;
}

function escribirTextoSiVacio(worksheet, celda, valor) {
    const cell = worksheet.getCell(celda);

    if (cell.value === null || cell.value === undefined || cell.value === '') {
        cell.value = valor;
    }
}

function crearErrorUsuario(mensaje) {
    const error = new Error(mensaje);
    error.statusCode = 400;
    return error;
}

// Limpia valores de un bloque completo sin tocar estilos, bordes, colores ni anchos.
function limpiarBloque(worksheet, bloque) {
    bloque.filasLimpieza.forEach((fila) => {
        ['A', 'B', 'C', 'D', 'E', 'F', 'G'].forEach((columna) => {
            worksheet.getCell(`${columna}${fila}`).value = '';
        });
    });
}

function limpiarBloqueActivo(worksheet) {
    limpiarBloque(worksheet, BLOQUE_ACTIVO);
}

function limpiarBloqueInactivo(worksheet) {
    limpiarBloque(worksheet, BLOQUE_INACTIVO);
}

// Cabecera del presupuesto. Solo se escriben datos recibidos, sin calculos.
function mapearCabecera(worksheet, presupuesto) {
    escribir(worksheet, 'D2', texto(presupuesto.cliente_nombre));
    escribir(worksheet, 'D3', texto(presupuesto.direccion));
    escribir(worksheet, 'D4', texto(presupuesto.empresa));
    escribir(worksheet, 'C7', texto(presupuesto.telefono));
    escribir(worksheet, 'C8', texto(presupuesto.email));
    escribir(worksheet, 'C9', texto(presupuesto.marca));
    escribir(worksheet, 'C10', texto(presupuesto.gama));
}

function esLineaEquipo(linea) {
    const contenido = normalizar([
        linea.tipo,
        linea.categoria,
        linea.familia,
        linea.referencia,
        linea.descripcion
    ].filter(Boolean).join(' '));

    return (
        contenido.includes('EQUIPO') ||
        contenido.includes('MAQUINA') ||
        contenido.includes('UNIDAD')
    );
}

function descripcionLinea(linea, textoPorDefecto = '', prefijo = '') {
    const referencia = texto(linea.referencia).trim();
    const descripcion = texto(linea.descripcion).trim();
    const base = descripcion || referencia || textoPorDefecto;

    if (referencia && descripcion) {
        return prefijo ? `${prefijo} - ${referencia} - ${descripcion}` : `${referencia} - ${descripcion}`;
    }

    return prefijo && base ? `${prefijo} - ${base}` : base;
}

// Cada linea se vuelca con importes ya calculados por n8n/Supabase.
function escribirLinea(worksheet, fila, linea) {
    const cantidad = numeroRecibido(linea.cantidad);
    const precioUnitario = numeroRecibido(linea.precio_unitario);
    const totalLinea = numeroRecibido(linea.total);

    escribir(worksheet, `${COLUMNAS_LINEA.cantidad}${fila}`, cantidad);
    escribir(worksheet, `${COLUMNAS_LINEA.descripcion}${fila}`, descripcionLinea(linea, TEXTOS.material));
    escribir(worksheet, `${COLUMNAS_LINEA.precioUnitario}${fila}`, precioUnitario);
    escribir(worksheet, `${COLUMNAS_LINEA.subtotal}${fila}`, totalLinea);
}

function escribirEquipo(worksheet, bloque, linea) {
    const marca = texto(linea.marca).trim();
    const gama = texto(linea.gama).trim();
    const prefijo = [marca, gama].filter(Boolean).join(' ');
    const cantidad = numeroRecibido(linea.cantidad);
    const precioUnitario = numeroRecibido(linea.precio_unitario);
    const totalLinea = numeroRecibido(linea.total);

    escribir(worksheet, `A${bloque.equipo.datos}`, cantidad);
    escribir(worksheet, `B${bloque.equipo.datos}`, descripcionLinea(linea, TEXTOS.equipo, prefijo));
    escribir(worksheet, `C${bloque.equipo.importe}`, precioUnitario);
    escribir(worksheet, `E${bloque.equipo.importe}`, totalLinea);
}

function mapearTextosCostes(worksheet, bloque) {
    const textos = bloque.textos;

    escribirTextoSiVacio(worksheet, textos.equipo, TEXTOS.equipo);
    escribirTextoSiVacio(worksheet, textos.subtotalMateriales, TEXTOS.subtotalMateriales);
    escribirTextoSiVacio(worksheet, textos.descuentoMateriales, TEXTOS.descuentoMateriales);
    escribirTextoSiVacio(worksheet, textos.horasOficial, TEXTOS.horasOficial);
    escribirTextoSiVacio(worksheet, textos.horasAyudante, TEXTOS.horasAyudante);
    escribirTextoSiVacio(worksheet, textos.horasOficina, TEXTOS.horasOficina);
    escribirTextoSiVacio(worksheet, textos.desplazamiento, TEXTOS.desplazamiento);
    escribirTextoSiVacio(worksheet, textos.incrementoMaterial, TEXTOS.incrementoMaterial);
    escribirTextoSiVacio(worksheet, textos.subtotal, TEXTOS.subtotal);
    escribirTextoSiVacio(worksheet, textos.iva, TEXTOS.iva);
    escribirTextoSiVacio(worksheet, textos.total, TEXTOS.total);
}

// No se asumen tipos de material: se respeta el orden que llega en req.body.lineas.
function mapearLineas(worksheet, lineas, bloque, presupuesto) {
    const lineasEntrada = (Array.isArray(lineas) ? lineas : []).map((linea) => ({
        ...linea,
        marca: linea.marca ?? presupuesto.marca,
        gama: linea.gama ?? presupuesto.gama
    }));
    const filasMateriales = [...bloque.lineas];
    const indiceEquipoDetectado = lineasEntrada.findIndex((linea) => esLineaEquipo(linea));
    const indiceEquipo = indiceEquipoDetectado >= 0 ? indiceEquipoDetectado : 0;

    lineasEntrada.forEach((linea, indice) => {
        if (indice === indiceEquipo) {
            escribirEquipo(worksheet, bloque, linea);
            return;
        }

        const fila = filasMateriales.shift();

        if (fila) {
            escribirLinea(worksheet, fila, linea);
            return;
        }

        throw crearErrorUsuario('La plantilla no tiene mas filas de materiales disponibles.');
    });
}

// Costes finales recibidos ya calculados. Railway solo los coloca en celdas.
function mapearCostes(worksheet, presupuesto, bloque) {
    const costes = bloque.costes;
    const incrementoMaterialPct = numeroRecibido(presupuesto.incremento_material_pct);

    mapearTextosCostes(worksheet, bloque);

    escribir(worksheet, costes.subtotalMateriales, valorPresupuesto(presupuesto, 'subtotal_materiales'));
    escribir(worksheet, costes.descuentoMateriales, valorPresupuesto(presupuesto, 'descuento_materiales'));
    escribir(worksheet, costes.subtotalMaterialesFinal, valorPresupuesto(presupuesto, 'subtotal_materiales_final'));

    escribir(worksheet, costes.horasOficial, valorPresupuesto(presupuesto, 'horas_oficial_primera'));
    escribir(worksheet, costes.costeOficial, valorPresupuesto(presupuesto, 'coste_oficial'));

    escribir(worksheet, costes.horasAyudante, valorPresupuesto(presupuesto, 'horas_ayudante'));
    escribir(worksheet, costes.costeAyudante, valorPresupuesto(presupuesto, 'coste_ayudante'));

    escribir(worksheet, costes.horasOficina, valorPresupuesto(presupuesto, 'horas_oficina'));
    escribir(worksheet, costes.costeOficina, valorPresupuesto(presupuesto, 'coste_oficina'));

    escribir(worksheet, costes.desplazamientos, valorPresupuesto(presupuesto, 'numero_desplazamientos'));
    escribir(worksheet, costes.costeDesplazamientos, valorPresupuesto(presupuesto, 'coste_desplazamientos'));

    escribir(
        worksheet,
        costes.incrementoMaterialPct,
        typeof incrementoMaterialPct === 'number' ? incrementoMaterialPct / 100 : incrementoMaterialPct
    );

    escribir(worksheet, costes.subtotal, valorPresupuesto(presupuesto, 'subtotal'));
    escribir(worksheet, costes.iva, valorPresupuesto(presupuesto, 'iva'));
    escribir(worksheet, costes.total, valorPresupuesto(presupuesto, 'total'));
}

function validarEntrada(body) {
    if (!body || typeof body !== 'object') {
        throw crearErrorUsuario('El cuerpo de la peticion debe ser un objeto JSON.');
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
                throw crearErrorUsuario(`La linea ${indice + 1} debe ser un objeto.`);
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
        const bloqueActivo = BLOQUE_ACTIVO;

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(PLANTILLA);

        const worksheet = workbook.worksheets[0];

        // Se limpia el unico bloque activo y se borra el bloque sobrante de la plantilla.
        limpiarBloqueActivo(worksheet);
        limpiarBloqueInactivo(worksheet);

        mapearCabecera(worksheet, presupuesto);
        mapearLineas(worksheet, lineas, bloqueActivo, presupuesto);
        mapearCostes(worksheet, presupuesto, bloqueActivo);

        res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );

        res.setHeader(
            'Content-Disposition',
            'attachment; filename=presupuesto-burgas.xlsx'
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

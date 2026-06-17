const express = require('express');
const ExcelJS = require('exceljs');

const app = express();

app.use(express.json({ limit: '10mb' }));

app.get('/', (req, res) => {
    res.send('Burgas Excel Generator funcionando');
});

app.post('/generar', async (req, res) => {
    try {

        const workbook = new ExcelJS.Workbook();

        await workbook.xlsx.readFile('./plantilla.xlsx');

        const worksheet = workbook.worksheets[0];
const presupuesto = req.body.presupuesto;

// Cabecera
worksheet.getCell('D2').value =
    presupuesto.cliente_nombre ?? '';

worksheet.getCell('D3').value =
    presupuesto.direccion ?? '';

// Mano de obra
worksheet.getCell('A25').value =
    Number(presupuesto.horas_oficial_primera ?? 0);

worksheet.getCell('A26').value =
    Number(presupuesto.horas_ayudante ?? 0);

worksheet.getCell('A27').value =
    Number(presupuesto.horas_oficina ?? 0);

worksheet.getCell('A28').value =
    Number(presupuesto.numero_desplazamientos ?? 0);

// Incremento material
worksheet.getCell('C29').value =
    Number(presupuesto.incremento_material_pct ?? 0) / 100;

        res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );

        res.setHeader(
            'Content-Disposition',
            'attachment; filename=prueba.xlsx'
        );

        await workbook.xlsx.write(res);

        res.end();

    } catch (error) {

        console.error(error);

        res.status(500).json({
            error: error.message
        });

    }
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
    console.log("Servidor iniciado en puerto " + PORT);
});
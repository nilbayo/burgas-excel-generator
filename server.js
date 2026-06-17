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

        // PRUEBA: escribir texto en la plantilla
        worksheet.getCell('F2').value = 'PRUEBA RAILWAY';

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
    console.log(⁠'Servidor iniciado en puerto ${PORT}' ⁠);
});
const express = require('express');

const app = express();

app.use(express.json());

app.get('/', (req, res) => {
    res.send('Burgas Excel Generator funcionando');
});

app.post('/generar', async (req, res) => {

    console.log(req.body);

    res.json({
        ok: true,
        mensaje: 'Servidor funcionando'
    });

});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
    console.log(`Servidor iniciado en puerto ${PORT}`);
});
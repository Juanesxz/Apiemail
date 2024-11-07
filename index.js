require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const XLSX = require('xlsx');
const nodemailer = require('nodemailer');
const fs = require('fs');
const cors = require('cors');

const app = express();
app.use(bodyParser.json());
app.use(cors());

// Configura el transporte de correo (SMTP)
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS
  }
});

// Ruta para recibir los datos y enviar el correo
app.post('/enviar-excel', async (req, res) => {
  try {
    const datos = req.body;

    if (!Array.isArray(datos) || datos.length === 0) {
      return res.status(400).json({ success: false, message: 'Datos inválidos o vacíos' });
    }

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(datos);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Datos');

    const excelPath = `./datos_${Date.now()}.xlsx`;
    XLSX.writeFile(workbook, excelPath);

    const mailOptions = {
      from: process.env.EMAIL_USER,
      to: process.env.EMAIL_TO,
      subject: `Datos en Excel De Cierres de Caja En La Fecha ${datos[0].fechaRegistro}`,
      text: 'Adjunto encontrarás el archivo con los datos en formato Excel.',
      attachments: [{ filename: `Cierres de Caja-${datos[0].fechaRegistro}.xlsx`, path: excelPath }]
    };

    await transporter.sendMail(mailOptions);
    
    // Eliminar el archivo después de enviar el correo
    fs.unlink(excelPath, (err) => {
      if (err) console.error('Error al eliminar el archivo:', err);
    });

    res.status(200).json({ success: true, message: 'Correo enviado con éxito' });
  } catch (error) {
    console.error('Error al enviar el correo:', error);
    res.status(500).json({ success: false, message: 'Error al enviar el correo. Inténtalo de nuevo más tarde.' });
  }
});

const PORT = 3000;
app.listen(PORT, () => {
  console.log(`Servidor ejecutándose en http://localhost:${PORT}`);
});

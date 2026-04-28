const express = require('express');
const twilio = require('twilio');
const dotenv = require('dotenv');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

dotenv.config();

const app = express();

// ────── CONFIGURACIÓN ──────
const accountSid = process.env.TWILIO_ACCOUNT_SID;
const authToken = process.env.TWILIO_AUTH_TOKEN;
const twilioPhone = process.env.TWILIO_PHONE_NUMBER;
const lawyerPhone = process.env.LAWYER_PHONE;
const port = process.env.PORT || 3001;

// ────── MIDDLEWARE ──────
app.use(express.json());

const allowedOrigins = (process.env.ALLOWED_ORIGIN || '*')
  .split(',')
  .map(o => o.trim());

app.use(cors({
  origin: (origin, callback) => {
    if (allowedOrigins.includes('*') || !origin || allowedOrigins.includes(origin)) {
      callback(null, true);
    } else {
      callback(new Error(`CORS: origen no permitido: ${origin}`));
    }
  },
  credentials: true
}));

// Configurar multer para subir archivos (usar memoria en lugar de disco para cloud)
const upload = multer({
  storage: multer.memoryStorage(),
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === '.xlsx' || ext === '.xls' || ext === '.csv') {
      cb(null, true);
    } else {
      cb(new Error('Solo se permite Excel (.xlsx, .xls) o CSV'));
    }
  }
});

// Inicializar cliente Twilio
const client = twilio(accountSid, authToken);

// ────── RUTAS ──────

/**
 * POST /api/upload-excel
 * Sube un archivo Excel y extrae los números de teléfono
 */
app.post('/api/upload-excel', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No se subió ningún archivo' });
    }

    // Leer archivo desde memoria
    const workbook = XLSX.read(req.file.buffer);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet);

    // Extraer números de teléfono (busca en la primera columna o columna con nombre "telefono", "phone", "numero", etc)
    const phones = [];
    data.forEach(row => {
      const phoneKeys = Object.keys(row).filter(k =>
        /telefono|phone|numero|celular|movil|whatsapp/i.test(k)
      );
      const searchKeys = phoneKeys.length > 0 ? phoneKeys : Object.keys(row);

      for (let key of searchKeys) {
        const value = String(row[key]).trim();
        if (/^\+?[0-9]{10,15}$/.test(value)) {
          phones.push(value);
          break;
        }
      }
    });

    // Formatear números a internacional si no tienen +57
    const formattedPhones = phones.map(phone => {
      const digits = phone.replace(/\D/g, '');
      if (phone.startsWith('+')) return '+' + digits;
      if (digits.startsWith('57') && digits.length >= 12) return '+' + digits;
      return '+57' + digits;
    });

    res.json({
      success: true,
      total: formattedPhones.length,
      phones: formattedPhones,
      message: `Se extrajeron ${formattedPhones.length} números de teléfono`
    });

  } catch (error) {
    res.status(500).json({ 
      error: 'Error al procesar el archivo',
      details: error.message 
    });
  }
});

/**
 * POST /api/send-messages
 * Envía mensajes masivos a una lista de números
 */
app.post('/api/send-messages', async (req, res) => {
  try {
    const { phones, messageTemplate } = req.body;

    if (!phones || !Array.isArray(phones) || phones.length === 0) {
      return res.status(400).json({ error: 'Lista de teléfonos vacía' });
    }

    if (!client) {
      return res.status(500).json({ error: 'Cliente Twilio no configurado' });
    }

    const results = [];
    let successful = 0;
    let failed = 0;

    // Enviar mensaje a cada número
    for (const phone of phones) {
      try {
        const message = await client.messages.create({
          body: messageTemplate || `¡Hola! Te estamos contactando sobre tu cartera. Comunícate con nosotros al: ${lawyerPhone}`,
          from: twilioPhone,
          to: phone
        });

        results.push({
          phone,
          status: 'enviado',
          messageId: message.sid,
          timestamp: new Date()
        });
        successful++;
      } catch (error) {
        results.push({
          phone,
          status: 'error',
          error: error.message,
          timestamp: new Date()
        });
        failed++;
      }
      
      // Pequeña pausa entre envíos (para evitar rate limiting)
      await new Promise(resolve => setTimeout(resolve, 100));
    }

    res.json({
      success: true,
      summary: {
        total: phones.length,
        successful,
        failed
      },
      results
    });

  } catch (error) {
    res.status(500).json({
      error: 'Error al enviar mensajes',
      details: error.message
    });
  }
});

/**
 * GET /api/test
 * Endpoint de prueba para verificar que el servidor está activo
 */
app.get('/api/test', (req, res) => {
  res.json({
    status: 'ok',
    message: 'Servidor funcionando correctamente',
    twilioConfigured: !!accountSid && !!authToken,
    lawyerPhone
  });
});

// ────── MANEJO DE ERRORES ──────
app.use((err, req, res, next) => {
  console.error(err);
  res.status(500).json({
    error: 'Error interno del servidor',
    details: err.message
  });
});

// ────── INICIAR SERVIDOR ──────
app.listen(port, () => {
  console.log(`🚀 Servidor ejecutándose en http://localhost:${port}`);
  console.log(`📞 Teléfono de la abogada: ${lawyerPhone}`);
  console.log(`📱 Número Twilio: ${twilioPhone}`);
  console.log(`✅ Twilio ${accountSid ? 'configurado' : 'NO CONFIGURADO'}`);
});

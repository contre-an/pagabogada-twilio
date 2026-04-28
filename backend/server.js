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
const whatsappPhone = process.env.TWILIO_WHATSAPP_NUMBER;
const whatsappTemplateSid = process.env.TWILIO_TEMPLATE_SID;
const lawyerPhone = process.env.LAWYER_PHONE;
const lawyerName = process.env.LAWYER_NAME || 'Luisa Fernanda Ossa';
const entityName = process.env.ENTITY_NAME || 'Finagro';
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

    const contacts = [];
    data.forEach(row => {
      const nombreKey = Object.keys(row).find(k => /nombre/i.test(k));
      const bancoKey = Object.keys(row).find(k => /intermediario|banco|entidad/i.test(k));
      const phoneKeys = Object.keys(row).filter(k => /telefono|phone|numero|celular|movil|whatsapp/i.test(k));
      const searchKeys = phoneKeys.length > 0 ? phoneKeys : Object.keys(row);

      for (let key of searchKeys) {
        const value = String(row[key]).trim();
        if (/^\+?[0-9]{10,15}$/.test(value)) {
          const digits = value.replace(/\D/g, '');
          let phone;
          if (value.startsWith('+')) phone = '+' + digits;
          else if (digits.startsWith('57') && digits.length >= 12) phone = '+' + digits;
          else phone = '+57' + digits;

          contacts.push({
            phone,
            nombre: nombreKey ? String(row[nombreKey]).trim() : '',
            banco: bancoKey ? String(row[bancoKey]).trim() : ''
          });
          break;
        }
      }
    });

    res.json({
      success: true,
      total: contacts.length,
      phones: contacts.map(c => c.phone),
      contacts,
      message: `Se extrajeron ${contacts.length} contactos`
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
    const { phones, contacts, messageTemplate } = req.body;

    if (!phones || !Array.isArray(phones) || phones.length === 0) {
      return res.status(400).json({ error: 'Lista de teléfonos vacía' });
    }

    if (!client) {
      return res.status(500).json({ error: 'Cliente Twilio no configurado' });
    }

    const results = [];
    let successful = 0;
    let failed = 0;

    for (let i = 0; i < phones.length; i++) {
      const phone = phones[i];
      const contact = (contacts && contacts[i]) || {};
      const nombre = contact.nombre || 'Cliente';
      const banco = contact.banco || 'la entidad financiera';
      let sent = false;

      // Intentar primero por WhatsApp con template aprobado
      if (whatsappPhone && whatsappTemplateSid) {
        try {
          const message = await client.messages.create({
            from: `whatsapp:${whatsappPhone}`,
            to: `whatsapp:${phone}`,
            contentSid: whatsappTemplateSid,
            contentVariables: JSON.stringify({ '1': nombre, '2': lawyerName, '3': entityName, '4': banco })
          });
          results.push({ phone, nombre, status: 'enviado', canal: 'whatsapp', messageId: message.sid, timestamp: new Date() });
          successful++;
          sent = true;
        } catch (err) {
          // WhatsApp falló, intentar SMS
        }
      }

      // Si WhatsApp falló o no está configurado, enviar SMS
      if (!sent) {
        try {
          const smsBody = messageTemplate ||
            `Buenas tardes señor/a ${nombre}, le escribe ${lawyerName}, abogada externa de ${entityName}. Me gustaría comentarle las alternativas de pago disponibles respecto a la obligación vencida que tiene con ${banco}. Si le interesa, comuníquese conmigo al ${lawyerPhone} o vía WhatsApp.`;
          const message = await client.messages.create({
            body: smsBody,
            from: twilioPhone,
            to: phone
          });
          results.push({ phone, nombre, status: 'enviado', canal: 'sms', messageId: message.sid, timestamp: new Date() });
          successful++;
        } catch (error) {
          results.push({ phone, nombre, status: 'error', canal: 'sms', error: error.message, timestamp: new Date() });
          failed++;
        }
      }

      await new Promise(resolve => setTimeout(resolve, 500));
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

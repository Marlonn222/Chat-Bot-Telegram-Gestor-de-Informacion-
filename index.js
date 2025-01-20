const TelegramBot = require('node-telegram-bot-api');
const XLSX = require('xlsx');

// Configuración del bot de Telegram
const token = '7624251158:AAHrPukk_2bh_rY9F3FSSrqHCkj5KrC9Kog';
const bot = new TelegramBot(token, { polling: true });

// Ruta al archivo Excel
const excelPath = 'C:\\Users\\mabeb\\Documents\\Nectim_Terceros.xlsx';

// Función para convertir un número de Excel a una fecha legible
function convertirExcelFecha(numeroExcel) {
    if (!numeroExcel || isNaN(numeroExcel)) return 'Fecha no válida';
    const fechaBase = new Date(1899, 11, 30); // Fecha base para Excel
    const dias = Math.floor(numeroExcel);
    fechaBase.setDate(fechaBase.getDate() + dias);
    return fechaBase.toISOString().split('T')[0]; // Formato YYYY-MM-DD
}

// Función para formatear números como precios con formato de moneda
function formatearComoPrecio(valor) {
    if (!valor || isNaN(valor)) return 'N/A';
    return `$${parseInt(valor).toLocaleString('es-CO')}`;
}

// Función para buscar en el archivo Excel
function buscarEnExcel(otp, oth) {
    try {
        // Cargar el archivo Excel
        const workbook = XLSX.readFile(excelPath);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        // Convertir la hoja a JSON y verificar la estructura
        const rows = XLSX.utils.sheet_to_json(sheet);
        console.log('Primera fila de datos:', rows[0]); // Verifica la estructura de los datos

        // Buscar la fila que coincida con la OTP y OTH
        const resultado = rows.find(row => {
            console.log('Revisando fila:', row); // Para depurar la búsqueda
            return row.Title === otp && row['OTH '] === oth;
        });

        if (resultado) {
            return {
                OTP: resultado.Title,
                CIUDAD: resultado.CIUDAD,
                MRC_TERCERO: formatearComoPrecio(resultado['MRC TERCERO']),
                OBSERVACIONES: resultado.OBSERVACIONES,
                FECHA_ASIGNADO: convertirExcelFecha(resultado['FECHA  ASIGNADO_x002']),
                DIRECCION: resultado.DIRECCION,
                CLIENTE: resultado.CLIENTE,
                TIPO_SOLICITUD: resultado['TIPO SOLICITUD'],
                COD_SERVICIO: resultado['COD SERVICIO'],
                OTH: resultado['OTH '],
                DEPARTAMENTO: resultado.DEPARTAMENTO,
                REGIONAL: resultado.REGIONAL,
                ID_3RO: resultado['ID 3RO'],
                FECHA_PROG_ES: convertirExcelFecha(resultado['FECHA PROG ES']),
                FECHA_ENTREGA_UM: convertirExcelFecha(resultado['FECHA ENTREGA UM']),
                ESTADO_UM: resultado['ESTADO UM'],
                OData_3RO: resultado.OData_3RO,
                CODIGOS_RESOLUCION: resultado['CODIGOS DE RESOLUCIO']
            };
        } else {
            return null;
        }
    } catch (error) {
        console.error('Error al buscar en Excel:', error);
        return null;
    }
}

// Manejar mensajes del bot
bot.on('message', async (msg) => {
    const chatId = msg.chat.id;
    const messageText = msg.text ? msg.text.toLowerCase() : '';

    if (messageText === 'hola') {
        bot.sendMessage(chatId, 'OTP:?\nOTH:?');
        return;
    }

    // Verificar si el mensaje contiene OTP y OTH
    const otpMatch = messageText.match(/otp:\s*(\w+)/i);
    const othMatch = messageText.match(/oth:\s*(\w+)/i);

    if (otpMatch && othMatch) {
        const otp = otpMatch[1];
        const oth = othMatch[1];

        console.log('Valores a buscar:', otp, oth); // Verifica los valores extraídos
        bot.sendMessage(chatId, '🔍 Buscando información, por favor espere...');

        try {
            const resultado = await buscarEnExcel(otp, oth);
            
            if (resultado) {
                const mensaje = `✅ Resultados encontrados:\n\n` +
                    `📝 OTP: ${resultado.OTP}\n` +
                    `🔄 OTH: ${resultado.OTH}\n` +
                    `💼 CLIENTE: ${resultado.CLIENTE}\n` +
                    `💵 MRC: ${resultado.MRC_TERCERO}\n` +
                    `🗓️ FECHA ASIGNADO: ${resultado.FECHA_ASIGNADO}\n` +
                    `📍 DIRECCION: ${resultado.DIRECCION}\n` +
                    `🌆 CIUDAD: ${resultado.CIUDAD}\n` +
                    `🗺️ DEPARTAMENTO: ${resultado.DEPARTAMENTO}\n` +
                    `🌍 REGIONAL: ${resultado.REGIONAL}\n` +
                    `📋 TIPO SOLICITUD: ${resultado.TIPO_SOLICITUD}\n` +
                    `🏭 TERCERO: ${resultado.OData_3RO}\n` +
                    `🔢 COD SERVICIO: ${resultado.COD_SERVICIO}\n` +
                    `🏢 ID 3RO: ${resultado.ID_3RO}\n` +
                    `📊 ESTADO UM: ${resultado.ESTADO_UM}\n` +
                    `📝 CODIGOS DE RESOLUCION: ${resultado.CODIGOS_RESOLUCION}\n` +
                    `📄 OBSERVACIONES: ${resultado.OBSERVACIONES}\n` +
                    `📅 FECHA ENTREGA UM: ${resultado.FECHA_ENTREGA_UM}\n` +
                    `📆 FECHA PROG ES: ${resultado.FECHA_PROG_ES}`;

                bot.sendMessage(chatId, mensaje);
            } else {
                bot.sendMessage(chatId, '❌ No se encontraron resultados para la OTP y OTH proporcionadas.');
            }
        } catch (error) {
            bot.sendMessage(chatId, '❌ Ocurrió un error al buscar la información. Por favor, intente nuevamente.');
            console.error('Error completo:', error);
        }
    }
});

console.log('Bot iniciado... Esperando mensajes');

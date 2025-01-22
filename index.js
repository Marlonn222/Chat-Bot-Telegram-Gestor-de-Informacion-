const TelegramBot = require('node-telegram-bot-api');
const XLSX = require('xlsx');

// Configuración del bot de Telegram
const token = '7624251158:AAHrPukk_2bh_rY9F3FSSrqHCkj5KrC9Kog';
const bot = new TelegramBot(token, { polling: true });

// Ruta al archivo Excel
const excelPath = 'C:\\Users\\mabeb\\Documents\\Nectim_Terceros.xlsx';

// Almacén de estados de los usuarios
const userStates = {};

// Función para convertir un número de Excel a una fecha legible
function convertirExcelFecha(numeroExcel) {
    if (!numeroExcel || isNaN(numeroExcel)) return 'Pendiente De Fecha';
    const fechaBase = new Date(1899, 11, 30); // Fecha base para Excel
    const dias = Math.floor(numeroExcel);
    fechaBase.setDate(fechaBase.getDate() + dias);
    return fechaBase.toISOString().split('T')[0]; // Formato YYYY-MM-DD
}

// Función para buscar OTH asociadas a una OTP en el archivo Excel
function buscarOTHPorOTP(otp) {
    try {
        const workbook = XLSX.readFile(excelPath);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet);

        console.log("🔎 Buscando OTH para la OTP:", otp); // Log para ver qué OTP se busca

        const othList = rows
            .filter(row => row.Title === otp) // Ajusta esto si el nombre de la columna no es "Title"
            .map(row => ([
                row['OTH '],          // Verifica que esta columna esté bien nombrada
                row['OData_3RO'],     // Verifica el nombre exacto de esta columna
                row['ESTADO UM']      // Verifica el nombre exacto de esta columna
            ]));

        console.log("📋 Resultados de la búsqueda de OTH:", othList); // Log para ver qué OTHs se encuentran

        return othList;
    } catch (error) {
        console.error('🚨🚨 Error al buscar OTH por OTP:', error);
        return [];
    }
}

// Función para buscar información por OTP y OTH en el archivo Excel
function buscarPorOTH(oth) {
    try {
        const workbook = XLSX.readFile(excelPath);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet);

        console.log("🔎 Buscando por OTH:", oth); // Log para ver qué OTH se busca

        const resultado = rows.find(row => row['OTH '] === oth); // Búsqueda solo por OTH

        if (resultado) {
            return {
                OTP: resultado.Title,
                CIUDAD: resultado.CIUDAD,
                MRC_TERCERO: `${parseInt(resultado['MRC TERCERO'] || 0).toLocaleString('es-CO')}`,
                OBSERVACIONES: resultado.OBSERVACIONES,
                FECHA_ASIGNADO: convertirExcelFecha(resultado['FECHA  ASIGNADO_x002']),
                DIRECCION: resultado.DIRECCION,
                CLIENTE: resultado.CLIENTE,
                TIPO_SOLICITUD: resultado['TIPO SOLICITUD'],
                OData_3RO: resultado.OData_3RO,
                COD_SERVICIO: resultado['COD SERVICIO'],
                OTH: resultado['OTH '],
                DEPARTAMENTO: resultado.DEPARTAMENTO,
                REGIONAL: resultado.REGIONAL,
                ID_3RO: resultado['ID 3RO'],
                FECHA_PROG_ES: convertirExcelFecha(resultado['FECHA PROG ES']),
                FECHA_ENTREGA_UM: convertirExcelFecha(resultado['FECHA ENTREGA UM']),
                ESTADO_UM: resultado['ESTADO UM'],
                CODIGOS_RESOLUCION: resultado['CODIGOS DE RESOLUCIO']
            };
        } else {
            return null;
        }
    } catch (error) {
        console.error('🚨🚨 Error al buscar por OTH:', error);
        return null;
    }
}

// Manejar mensajes del bot en el chat
bot.onText(/.*/, (msg) => {
    const chatId = msg.chat.id;
    const messageText = msg.text ? msg.text.trim() : '';

    // Verificar el estado actual del usuario
    const userState = userStates[chatId] || { step: 'start' };

    if (userState.step === 'start') {
        const menuOptions = {
            reply_markup: {
                keyboard: [
                    [{ text: 'OTP' }, { text: 'OTH' }]
                ],
                one_time_keyboard: true
            }
        };
        bot.sendMessage(chatId, '👋🏽 Hola, ¿cómo estás? Soy Enrique y te hablo de NECTIM 🇨🇴\n 🙏🏽 Por favor, ¿qué tipo de orden deseas consultar?', menuOptions);
        userStates[chatId] = { step: 'menu' };
    } else if (userState.step === 'menu') {
        if (messageText === 'OTP') {
            bot.sendMessage(chatId, '🙏🏽 Por favor, ingrese la OTP:');
            userStates[chatId] = { step: 'awaitingOTP' };
        } else if (messageText === 'OTH') {
            bot.sendMessage(chatId, '🙏🏽 Por favor, ingrese la OTH:');
            userStates[chatId] = { step: 'awaitingOTH' };
        } else {
            bot.sendMessage(chatId, 'Por favor seleccione una opción válida.');
        }
    } else if (userState.step === 'awaitingOTP') {
        const otp = messageText;

        // Buscar las OTH asociadas
        const othList = buscarOTHPorOTP(otp);

        if (othList.length > 0) {
            userStates[chatId] = { step: 'awaitingOTH', otp };

            let tablaTexto = `🔎 Estas son las OTHs asociadas a la OTP ${otp}:\n`;
            tablaTexto += '-----------------------------------------------------------------------------------------------------\n';
            tablaTexto += '     🔄 OTH      |      📊 ESTADO UM       |     🏭 TERCERO     \n';
            tablaTexto += '-----------------------------------------------------------------------------------------------------\n';
            othList.forEach(item => {
                tablaTexto += `✅ ${item[0]} | ${item[1]} | ${item[2]}\n`;
            });
            tablaTexto += '-----------------------------------------------------------------------------------------------------\n';

            bot.sendMessage(chatId, `${tablaTexto}🙏🏽 Por favor, ingrese una OTH para obtener más detalles o ingrese una nueva OTH.`);
        } else {
            bot.sendMessage(chatId, `🚨🚨 No se encontraron OTH asociadas a la OTP ${otp}. 💱 Intente nuevamente.`);
        }
    } else if (userState.step === 'awaitingOTH') {
        const oth = messageText;

        // Buscar solo por OTH
        const resultado = buscarPorOTH(oth);

        if (resultado) {
            const mensaje = `✅ Información encontrada:\n` +
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
                `📝 CODIGOS DE RESOLUCIO: ${resultado.CODIGOS_RESOLUCION}\n` +
                `📄 OBSERVACIONES: ${resultado.OBSERVACIONES}\n` +
                `📅 FECHA ENTREGA UM: ${resultado.FECHA_ENTREGA_UM}\n` +
                `📆 FECHA PROG ES: ${resultado.FECHA_PROG_ES}`;
            bot.sendMessage(chatId, mensaje);
        } else {
            bot.sendMessage(chatId, `🚨🚨 No se encontró información para la OTH ${oth}.`);
        }

        // Reiniciar el estado del usuario
        userStates[chatId] = { step: 'start' };
    }
});

console.log('Bot iniciado... Esperando mensajes.');

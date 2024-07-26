const fs = require('fs');
const xlsx = require('xlsx');
const Docxtemplater = require('docxtemplater');
const PizZip = require("pizzip");
const { toWords } = require('num-words');
const Logica = require('./Logica.js');

const { createBot, createProvider, createFlow, addKeyword } = require('@bot-whatsapp/bot')

const QRPortalWeb = require('@bot-whatsapp/portal')
const WPPConnectProviderClass = require('@bot-whatsapp/provider/wppconnect')
const MockAdapter = require('@bot-whatsapp/database/mock')
let arc = 'C:/Users/tecnologo.operacag/Desktop/chat/PDFs/archivo_unificado.pdf'
const { unificarPDFs } = require("./event.js");

let lista = [];

const flowConfirmacion_colilla = addKeyword('Generacion_documento_carta_activo')
.addAnswer('Si la colilla de pago abre correctamente escribe *Si*, por el contrario escribe *No*',
{capture:true},async(ctx, {fallBack,flowDynamic,state}) => {
    const myState = state.getMyState()
    if(ctx.body == "Si" || ctx.body == "si"){
        await flowDynamic([{ body: "Gracias por la confirmación"}]);
    }else if (ctx.body == "No" || ctx.body == "no"){
        const path = `C:/Users/tecnologo.operacag/Desktop/chatv2.0/PDFs/${myState.NumeroCedula}.pdf`;
        await flowDynamic([{ body: "archivo", media: path }]);
        //Logica.writeToExcelMultiple(myState.ND[0],3)
    }else{
        return fallBack();
    }

});

const flowgeneracion_colilla = addKeyword('Generacion_documento_carta_activo')
.addAnswer('Procederemos a generar tu colilla por favor espera un momento',
{delay:2000},async(_, {gotoFlow,flowDynamic,state}) => {
    
    const myState = state.getMyState()
    const resultado = await Logica.consultarCedulaEnExcel_colilla(myState.NumeroCedula);
    //console.log('flujo espera')
    console.log(resultado)
    //let cc = myState.NumeroCedula
    let ifExist = resultado[0]
    let nombres = null
    let ND = null
    if (resultado[1]){
        nombres = resultado[1]
    }

    if (ifExist) {                  
        try {
            const path = `C:/Users/tecnologo.operacag/Desktop/chatv2.0/PDFs/${myState.NumeroCedula}.pdf`;
            //await state.update({ path: path })
            //console.log('Ruta del PDF:', path);
            //console.log('Archivos PDF unificados con éxito');
            console.log('path')
            //console.log(path)
            
            await flowDynamic([{ body: "archivo",media: path }]);
            return gotoFlow(flowConfirmacion_colilla)
            
        } catch (error) {
            console.error('Error al unificar archivos PDF:', error);
        }
        //return gotoFlow(flowEspera1);
        // // eslint-disable-next-line bot-whatsapp/func-prefix-dynamic-flow-await
        //await flowDynamic([{body:'enviadno...',media:rutaPDF}])

    } else {
        return gotoFlow(flowUserNoExiste);
    }
    //}
    lista = [];
    //return gotoFlow(flowEspera1);
    //for (const archivo of lista) {
    //await flowDynamic([{ body: "entrega", media: `C:/Users/tecnologo.operacag/Desktop/chat/PDFs/${archivo}.pdf` }]);
    //}
    
})

const flowautenticacionExitosa_colilla = addKeyword('Personal_autenticacion_exitosa_activo').addAnswer('Autenticacion exitosa.', 
{delay:1000}, async (ctx, {gotoFlow, state}) => {
    const myState = state.getMyState()
    const pass = await Logica.ReadExcel_activo(myState.NumeroCedula,3)
    console.log(pass)
    //let pass = true
    if (pass){
        return gotoFlow(flowgeneracion_colilla)
    }else{
        //console.log('no permitido el ingreso')
        return gotoFlow(flowIntentosAgotados)
        
    }
})

const flowActivos_colillas = addKeyword(['Personal_colillas'])
.addAnswer('📄 Vamos a proceder a verificar tu identidad. Por favor escribe tu *número interno*.')
.addAnswer(' Si no lo conoces puedes solicitarlo con gestión humana al número de teléfono: 300 4827432'
,{capture:true},async(ctx, {gotoFlow,state}) => {
    const myState = state.getMyState()
    let pass = await Logica.consultarNumeroInterno(myState.NumeroCedula,ctx.body)
    console.log(pass)
    if (pass){
        console.log('autenticado')
        return gotoFlow(flowautenticacionExitosa_colilla)
    }else{
        console.log('no autenticado')
        return gotoFlow(flowautenticacionFallida)
    }
}
)

const flowcolillas = addKeyword(['2'])
.addAnswer('📝Para continuar, ingresa tu número de documento sin puntos, comas o guiones: *(Ejem.: 1234567890)* para generar tu colilla de pago',
{capture:true,delay:1000},async(ctx, {fallBack,gotoFlow,endFlow,state}) => {
    
    if(ctx.body.length >= 7 && !isNaN(ctx.body)){
        try {
            let activo = await Logica.excel_colillas(ctx.body)
            await state.update({ NumeroCedula: ctx.body })

            if (activo){
                //console.log('activo colillas ')
                return gotoFlow(flowActivos_colillas)
            }else{
                //console.log('no activo colillas')flowUserNoExiste
                return gotoFlow(flowUserNoExiste)
            }
            
        } catch (error) {
            console.error('Error al consultar número en Excel0:', error);
            return fallBack();
        }
    }else{
        return fallBack();
    }
    
}
)

const flowgeneracion_documento_Activo = addKeyword('Generacion_documento_carta_activo').addAnswer('Procederemos a generar tu documento por favor espera un momento',
{delay:2000},async(_, {gotoFlow,flowDynamic,state}) => {
    
    //console.log('ctx')
    //console.log(ctx)
    //console.log('ctx.body')
    //console.log(ctx.body)
    const myState = state.getMyState()
    
    const resultado = await Logica.consultarCedulaEnExcel(myState.NumeroCedula);
    console.log('ctx2')

    let ifExist = resultado[0]
    let nombres = null
    if (resultado[1]){
        nombres = resultado[1]
    }
    console.log('ctx3')
    if (ifExist) {  
        console.log('ctx4')                
        // const path = await unificarPDFs(nombres,NumeroCedula)
        //     .then(() => console.log('Archivos PDF unificados con éxito'))
        //     .catch(error => console.error('Error al unificar archivos PDF:', error));
        try {
            const path = await unificarPDFs(nombres, myState.NumeroCedula);
            //await state.update({ path: path })
            //console.log('Ruta del PDF:', path);
            //console.log('Archivos PDF unificados con éxito');
            console.log('path')
            console.log(path)
            
            await flowDynamic([{ body: "entrega", media: path }]);
        } catch (error) {
            console.error('Error al unificar archivos PDF:', error);
        }
        //return gotoFlow(flowEspera1);
        // // eslint-disable-next-line bot-whatsapp/func-prefix-dynamic-flow-await
        //await flowDynamic([{body:'enviadno...',media:rutaPDF}])

    } else {
        return gotoFlow(flowUserNoExiste);
    }
    //}
    lista = [];
    //return gotoFlow(flowEspera1);
    //for (const archivo of lista) {
    //await flowDynamic([{ body: "entrega", media: `C:/Users/tecnologo.operacag/Desktop/chat/PDFs/${archivo}.pdf` }]);
    //}
    
})

const flowautenticacionExitosa_activo = addKeyword('Personal_autenticacion_exitosa_activo').addAnswer('Autenticacion exitosa.', 
{delay:1000}, async (ctx, {gotoFlow, state}) => {
    const myState = state.getMyState()
    const pass = await Logica.ReadExcel_activo(myState.NumeroCedula,2)

    if (pass){
        return gotoFlow(flowgeneracion_documento_Activo)
    }else{
        return gotoFlow(flowIntentosAgotados)
    }
})

const flowautenticacionFallida = addKeyword('Personal_autenticacion_fallida')
.addAnswer('la autenticación ha fallado, por favor verifica tus datos, O escribe a gestión humana al número de teléfono: 300 4827432')
.addAnswer('escribe *Cancelar* para reiniciar la conversacion.')

const flowActivos = addKeyword('Personal_activo_excel')
    .addAnswer('📄 Vamos a proceder a verificar tu identidad, Por favor escribe los ultimos 4 digitos de tu cuenta de ahorros asociado a la empresa,')
    .addAnswer(' Si no lo conoces puedes solicitarlo con gestión humana al número de teléfono: 300 4827432'
    ,{capture:true,delay:1000},async(ctx, {gotoFlow,state,endFlow}) => {
        const myState = state.getMyState()
        let pass = await Logica.consultarNumeroInterno(myState.NumeroCedula,ctx.body)
        //console.log('paso 2')
        //console.log(pass)
        if (pass){
            //console.log('paso 2.1')
            return gotoFlow(flowautenticacionExitosa_activo)
        }else{
            //console.log('paso 2.3')
            return gotoFlow(flowautenticacionFallida)
        }
})

const flowenvio_correo = addKeyword('Envio_correo_GH')
.addAnswer('muy bien, ahora procederemos a enviar la información a gestión humana, la recepcion del correo puede tomar unas horas por favor se paciente',
null,async(_, {gotoFlow,state}) => {
    const myState = state.getMyState()
    console.log('sender de email1')
    await Logica.send_mail(correo,myState.NumeroCedula)
    console.log('sender de email2')
})

const flowvalidacion_correo = addKeyword('Validacion_correo_personal_inactivo')
    .addAnswer('por favor escribe la direccion de correo electronico al que quieres que te llegue el correo.'
    ,{capture:true},async(ctx, { fallBack,gotoFlow }) => {
        if (!ctx.body.includes('@')) {
          return fallBack()
        } else {
            correo = ctx.body
            console.log('validador de email')
            return gotoFlow(flowenvio_correo)
          // Lógica para procesar el correo electrónico del usuario
        }
})

const flowgeneracion_documentos_inactivos = addKeyword('Genracion_documento_carta_activo').addAnswer('Procederemos a generar tu documento por favor espera un momento',
null,async(_, {gotoFlow,endFlow,state}) => {
    const myState = state.getMyState()
    const resultado = await Logica.consultarCedulaEnExcel(myState.NumeroCedula);
    let ifExist = resultado[0]
    let nombres = null
    if (resultado[1]){
        nombres = resultado[1]
    }

    if (ifExist) {                      
        await Logica.unificarPDFs(nombres)
            .then(() => console.log('Archivos PDF unificados con éxito'))
            .catch(error => console.error('Error al unificar archivos PDF:', error));
        return gotoFlow(flowvalidacion_correo)
    } else {
        return gotoFlow(flowUserNoExiste);
    }
})
      
const flowautenticacionExitosa_Inactivo = addKeyword('Personal_autenticacion_exitosa_inactivo').addAnswer('Autenticacion exitosa.', 
null, async (ctx, {gotoFlow,state}) => {
    const myState = state.getMyState()
    const pass = await Logica.ReadExcel_Inactivo(myState.NumeroCedula)

    if (pass){
        return gotoFlow(flowgeneracion_documentos_inactivos)
    }else{
        return gotoFlow(flowIntentosAgotados)
    }
})

const flowInactivos = addKeyword('Personal_inactivo_excel')
    .addAnswer(['📄 Vamos a proceder a verificar tu identidad, Por favor escribe los ultimos 4 digitos de tu cuenta de ahorros asociado a la empresa,',
    'Si no lo conoces puedes conocerlo con gestión humana']
    ,{capture:true},async(ctx, {endFlow,gotoFlow,state}) => {
        const myState = state.getMyState()
        let pass = await Logica.consultarNumeroInterno(myState.NumeroCedula,ctx.body)
        
        if (pass){
            console.log('autenticado')
            return gotoFlow(flowautenticacionExitosa_Inactivo)
        }else{
            console.log('no autenticado')
            return gotoFlow(flowautenticacionFallida)
        }
})

const flowIntentosAgotados = addKeyword('Intentos_agotados_por_persona').addAnswer(
    [
        '🙌 Lo siento, parece que has excedido el límite de intentos.',
        'Si necesitas ayuda adicional o tienes alguna pregunta, no dudes en contactar a gestión humana. Estamos aquí para ayudarte.',
        'escribe *Cancelar* para reiniciar la conversacion.',
    ],
    {delay:2000},
    null,
    [flowcolillas]
)

const flowGracias = addKeyword(['gracias', 'grac']).addAnswer(
    [
        '🚀 Puedes aportar tu granito de arena a este proyecto',
        '[*opencollective*] https://opencollective.com/bot-whatsapp',
        '[*buymeacoffee*] https://www.buymeacoffee.com/leifermendez',
        '[*patreon*] https://www.patreon.com/leifermendez',
        '\n*2* Para siguiente paso.',
    ],
    null,
    null,
    [flowcolillas]
)

const flowFinalizarSesion = addKeyword(['Cancelar','cancelar','cerrar','cerra','cierra','salir'])
.addAnswer('Si quieres volver a intentar con algun proceso, por favor escribe *Hola*')

const flowCartas = addKeyword(['1']).addAnswer(
    '📝Para continuar, ingresa tu número de documento sin puntos, comas o guiones: *(Ejem.: 1234567890)* para generar tu carta'
    ,{capture:true,delay:1000},async(ctx, {fallBack,gotoFlow,state}) => {

        if(ctx.body.length >= 7 && !isNaN(ctx.body)){
            try {
                lista.push(ctx.body);
                let activo = await Logica.activo_excel(ctx.body)
                await state.update({ NumeroCedula: ctx.body })
                
                if (activo){
                    console.log('activo carta')
                    return gotoFlow(flowActivos)
                }else{
                    console.log('inactivo carta')
                    return gotoFlow(flowInactivos)
                }
                
            } catch (error) {
                console.error('Error al consultar número en Excel0:', error);
                return fallBack();
            }
        }else{
            return fallBack()
        }
})

const flowUserNoExiste = addKeyword(['Flujo_Usuario_No_existe']).addAnswer('No me apareces en la base de datos, por favor valida tu información con gestión humana al número de teléfono: 300 4827432')

const flowEspera = addKeyword('Flujo_espera_respuesta')
.addAnswer('Procederemos a generar tu documento por favor espera un momento',
{delay:1000},async(_, {gotoFlow,flowDynamic}) => {

    const rutaPDF = `../PDFs/${NumeroCedula}.pdf`;
    console.log('flujo espera')
    await flowDynamic([{media:rutaPDF}])
    console.log('flujo espera')
})

const flowEspera1 = addKeyword('Flujo_espera_respuesta')
.addAnswer('espera1',
{delay:1000}, async (ctx, {flowDynamic}) => {
    await flowDynamic([{ body: "entrega", media: `C:/Users/tecnologo.operacag/Desktop/chat/PDFs/${archivo}.pdf` }]);
})

const flowPrincipal = addKeyword(['hola', 'ole', 'alo','como vas','como estas','ola','cmo estas','que haces','qwe','asd','ofsjv'
,'wienv','quorv','que hases','hello','hi','helo','buena','buen día','buena tarde','buena noche','buenas','como estas'],{delay:2000})
    .addAnswer('¡Hola! 👋 soy EIA-Asistente 🤖, tu asistente virtual de EIA. Aquí puedes solictar colillas de pago y cartas laborales.')
    //.addAnswer('en cualquier momento puedes escribir *Cancelar* para poder reiniciar la conversacion')
    .addAnswer('Por favor, escribir el número de la opcion a solicitar: \n*1* Cartas Laborales \n*2* Colillas de Pago ',
    {capture:true, delay:1000}, async(ctx, { fallBack,gotoFlow }) => {
        if (ctx.body == '1' ) {
          return gotoFlow(flowCartas)
        }else if(ctx.body =='2'){
            return gotoFlow(flowcolillas)
        }else{
            return fallBack()
        }
})

const main = async () => {
    const adapterDB = new MockAdapter()
    const adapterFlow = createFlow([flowPrincipal,flowUserNoExiste,flowautenticacionExitosa_activo,flowautenticacionFallida,flowgeneracion_documento_Activo,
        flowCartas,flowvalidacion_correo,flowenvio_correo,flowgeneracion_documentos_inactivos,flowEspera1,flowActivos_colillas,flowautenticacionExitosa_colilla,
        flowActivos,flowEspera,flowInactivos,flowFinalizarSesion,flowGracias,flowIntentosAgotados,flowcolillas,flowautenticacionExitosa_Inactivo,
        flowgeneracion_colilla,flowConfirmacion_colilla])
        const adapterProvider = createProvider(WPPConnectProviderClass)

        createBot({
            flow: adapterFlow,
            provider: adapterProvider,
            database: adapterDB,
        })
    
        QRPortalWeb()
    }
    
    main()
const fs = require('fs');
const xlsx = require('xlsx');
const Docxtemplater = require('docxtemplater');
const PizZip = require("pizzip");
const numeroALetras = require('numero-a-letras');
//var docxConverter = require('docx-pdf');
const path = require('path');
const PDFDocumentKit = require('pdfkit');
const { PDFDocument, rgb } = require('pdf-lib');
const ExcelJS = require('exceljs');
const nodemailer = require('nodemailer');
const mammoth = require("mammoth");

const rutaPlantillaRetiros = '../Plantillas/PLANTILLA RETIROS.docx';
const rutaPlantillacolillas = '../Plantillas/Plantilla_colillas.docx';
const rutaPlantillaActivos = '../Plantillas/PLANTILLA ACTIVOS.docx';
const Lugar_wordsActivos = '../Word´s/';
const Lugar_Pdfs = '../PDFs/'; 
const archivoExcel = 'C:/Users/tecnologo.operacag/Desktop/chatv2.0/BASE DE DATOS UNIFICADA.xls';
const Excel_ubicacion_colillas = 'C:/Users/tecnologo.operacag/Desktop/chatv2.0/PAGO NOMINA.xls';
const filePath = 'C:/Users/tecnologo.operacag/Desktop/chatv2.0/BD Chat Bot.xlsx';
let fechaActual = new Date();

async function consultarNumeroCuenta(CC,cuenta) {
    // Lee el archivo Excel
    const workbook = xlsx.readFile(archivoExcel);

    // Obtén la primera hoja del libro de trabajo
    const segundaHoja = workbook.SheetNames[2];

    // Obtén los datos de la primera hoja como objeto
    const datos = xlsx.utils.sheet_to_json(workbook.Sheets[segundaHoja]);

    // Filtrar las filas que contengan el número en la columna "IDENTIFICACION"
    const filasConNumero = datos.filter(fila => fila.IDENTIFICACION === CC && fila.CUENTA_BANCARIA === cuenta);
    //console.log(filasConNumero.length)
    console.log('inactivo')
    if (filasConNumero > 0){
        return true
    }else{
        return false
    }
}

async function excel_colillas(numero){
    // Lee el archivo Excel
    const workbook = xlsx.readFile(Excel_ubicacion_colillas);

    // Obtén la primera hoja del libro de trabajo
    const primeraHoja = workbook.SheetNames[0];
    numero=parseInt(numero);
    // Obtén los datos de la primera hoja como objeto
    const datos = xlsx.utils.sheet_to_json(workbook.Sheets[primeraHoja]);
    //console.log(datos)
    // Filtrar las filas que contengan el número en la columna "IDENTIFICACION"
    const filasFiltradas = datos.filter(fila => fila.IDENTIFICACION === numero);
    //console.log(filasFiltradas)
    //console.log(numero)
    let filas = filasFiltradas.length
    console.log('filas colillas')
    //console.log(filas)
    if (filas > 0){
        return true
    }else{
        return false
    }

}

async function activo_excel(numero){
    // Lee el archivo Excel
    const workbook = xlsx.readFile(archivoExcel);

    // Obtén la primera hoja del libro de trabajo
    const primeraHoja = workbook.SheetNames[0];

    // Obtén los datos de la primera hoja como objeto
    const datos = xlsx.utils.sheet_to_json(workbook.Sheets[primeraHoja]);

    // Filtrar las filas que contengan el número en la columna "IDENTIFICACION"
    const filasFiltradas = datos.filter(fila => fila.IDENTIFICACION === numero && fila.ACTIVO === "Si");
    //const filasFiltradasNO = datos.filter(fila => fila.IDENTIFICACION === numero);
    console.log(filasFiltradas)
    console.log('activo')
    if (filasFiltradas.length > 0){
        console.log('paso activo 1')
        return true
    }else{
        console.log('paso activo 2 false')
        return false
    }

}

async function consultarNumeroInterno(CC,numero) {
    // Lee el archivo Excel
    const workbook = xlsx.readFile(archivoExcel);

    // Obtén la primera hoja del libro de trabajo
    const segundaHoja = workbook.SheetNames[1];
    console.log(CC,'cedula')
    CC = CC.toString();
    console.log(numero,'numero')
    //let numeroEntero = parseInt(numero, 10);
    // Obtén los datos de la primera hoja como objeto
    const datos = xlsx.utils.sheet_to_json(workbook.Sheets[segundaHoja]);
    //const fila = datos.filter(fila => fila.IDENTIFICACION === CC);
    //console.log(fila)

    // Filtrar las filas que contengan el número en la columna "IDENTIFICACION"
    const filasConNumero = datos.filter(fila => fila.IDENTIFICACION === CC && fila.CODIGO_INGRESO === numero);
    console.log('paso 1')
    console.log(filasConNumero)
    let filas = filasConNumero.length
    if (filas > 0){
        console.log('paso 1.2')
        return true
    }else if (filas == 0){
        console.log('paso 1.3')
        return false
    }
}

async function consultarCedulaEnExcel_colilla(numero) {
    // Lee el archivo Excel
    const workbook = xlsx.readFile(Excel_ubicacion_colillas);

    // Obtén la primera hoja del libro de trabajo
    const primeraHoja = workbook.SheetNames[0];

    // Obtén los datos de la primera hoja como objeto
    const datos = xlsx.utils.sheet_to_json(workbook.Sheets[primeraHoja]);
    numero = parseInt(numero);
    // Filtrar las filas que contengan el número en la columna "IDENTIFICACION"
    const filasConNumero = datos.filter(fila => fila.IDENTIFICACION === numero);
    filas = datos.filter(fila => fila.IDENTIFICACION === numero);
    let contador = 0;
    let encontrado = false;
    let nombres  = [];
    //console.log(filasConNumero)
    if (filasConNumero.length > 0) {

        const datos = {
            Pago: [],
            Des: []
        };
        let TP = 0 
        let TD = 0
        for (const elemento of filasConNumero) {

            let PositivoPago = Math.abs(elemento['ValorPago']);
            let valor = {
                "CON": elemento['CONCEPTO'].trim(),
                "CAN": elemento['CANTIDAD'],
                "ValorPago": PositivoPago.toLocaleString(),
                "HORAS": elemento['HORAS']
            };
            
            if (elemento['NATURALEZA'] == 'Pago     ') {
                datos.Pago.push(valor);
                TP += elemento['ValorPago']
            } else if (elemento['NATURALEZA'] == 'Descuento') {
                datos.Des.push(valor);
                TD += PositivoPago
            }
        }
        let Neto = TP - TD
        //console.log(diccionario1)
        let ND = {
            NOMINA: filasConNumero[0].NOMINA.trim(),
            IDENTIFICACION: filasConNumero[0].IDENTIFICACION,
            'NOMBRE COMPLETO': filasConNumero[0]['NOMBRE COMPLETO'].trim(),
            'FECHA INGRESO': numeroSerieAfecha(filasConNumero[0]['FECHA INGRESO']),
            'CENTRO DE COSTO': filasConNumero[0]['NOMINA'].trim(),
            'TP':TP.toLocaleString(),
            'TD':TD.toLocaleString(),
            'Neto':Neto.toLocaleString(),
            'PERIODO PAGO':filasConNumero[0]['PERIODO PAGO '],
            'SALARIO':filasConNumero[0]['SALARIO'].toLocaleString(),
            'CARGO':filasConNumero[0]['CARGO'],
            'FECHA DE PAGO':numeroSerieAfecha(filasConNumero[0]['FECHA DE PAGO'])
        };
        let nombre_persona = ND['NOMBRE COMPLETO'];
        let nombre_archivo = `${nombre_persona}_${contador}.docx`;
        
        ND.Pago =datos.Pago
        ND.Des = datos.Des
        console.log(ND)
        generarDocumentoPlantillaColillas(ND, nombre_archivo);
        let pdf = await ConvertToPDF_colilla(nombre_archivo,datos,numero);
        nombres.push(pdf);
        encontrado = true;
        //writeToExcelMultiple(filasConNumero[0],3)
    }
    else{
        encontrado = false;
    }
    return [encontrado,nombres];
}

async function consultarCedulaEnExcel(numero) {
    // Lee el archivo Excel
    const workbook = xlsx.readFile(archivoExcel);

    // Obtén la primera hoja del libro de trabajo
    const primeraHoja = workbook.SheetNames[0];

    // Obtén los datos de la primera hoja como objeto
    const datos = xlsx.utils.sheet_to_json(workbook.Sheets[primeraHoja]);

    // Filtrar las filas que contengan el número en la columna "IDENTIFICACION"
    const filasConNumero = datos.filter(fila => fila.IDENTIFICACION === numero);
    filas = datos.filter(fila => fila.IDENTIFICACION === numero);
    let contador = 0;
    let encontrado = false;
    let nombres  = [];
    let consecutivo = await FindConsecutivo()

    if (filasConNumero.length > 0) {
        for (const fila of filasConNumero) {
            let Fila = TuplaExcel(fila);
            let nombre_persona = Fila['NOMBRE COMPLETO'].trim();
            let nombre_archivo = `${nombre_persona}_${contador}.docx`;
            contador++;
            let new_conse =consecutivo+contador

            let nuevoTexto = new_conse.toString();
            Fila['Consecutivo1']=nuevoTexto
            Fila['Consecutivo2']=nuevoTexto
            
            if (Fila['ACTIVO'] === 'No') {
                generarDocumentoPlantilla(Fila, nombre_archivo);
                console.log('retiro');
                let pdf = await ConvertToPDF(nombre_archivo);
                nombres.push(pdf);
                encontrado = true;

            } else if (Fila['ACTIVO'] === 'Si') {
                console.log('activo');
                generarDocumentoPlantillaactivo(Fila, nombre_archivo);
                let pdf = await ConvertToPDF(nombre_archivo);
                nombres.push(pdf);
                encontrado = true;
            }
        }
        //writeToExcelMultiple(filasConNumero[0],2)
    }
    return [encontrado,nombres];
}

function numeroSerieAfecha(numeroSerie) {
    // Definir la fecha de origen de Excel (1 de enero de 1900)
    var fechaOrigen = new Date('1899-12-30');
    
    // Calcular la cantidad de días desde la fecha de origen
    var dias = numeroSerie - 1;
    
    // Calcular la cantidad de milisegundos desde la fecha de origen
    var milisegundos = dias * 24 * 60 * 60 * 1000;
    
    // Crear y devolver el objeto de fecha sumando los milisegundos a la fecha de origen
    var fecha = new Date(fechaOrigen.getTime() + milisegundos);
    fecha.setUTCHours(0, 0, 0, 0); // Establecer la hora en 00:00:00

    // Obtener día, mes y año
    var dia = fecha.getUTCDate();
    var mes = fecha.getUTCMonth() + 1; // Los meses van de 0 a 11, por eso se suma 1
    var anio = fecha.getUTCFullYear();

    // Formatear la fecha como día/mes/año
    var fechaFormateada = `${dia}/${mes}/${anio}`;

    // Devolver la fecha formateada
    return fechaFormateada;
}

function generarDocumentoPlantillaColillas(datos, nombreArchivoSalida) {
    try {
        
        // Verificar si la plantilla existe
        if (!fs.existsSync(rutaPlantillacolillas)) {
            throw new Error(`La plantilla "${rutaPlantillacolillas}" no existe.`);
        }
        const contenidoPlantilla = fs.readFileSync(rutaPlantillacolillas,"binary");

        const zip = new PizZip(contenidoPlantilla);
        const generadorPlantillas = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });

        generadorPlantillas.render(datos);

        const buf = generadorPlantillas.getZip().generate({
            type: "nodebuffer",
            compression: "DEFLATE",
        });
        let lugar_nombre = Lugar_wordsActivos+nombreArchivoSalida
        fs.writeFileSync(lugar_nombre, buf);

        console.log(`El documento "${lugar_nombre}" ha sido generado con éxito.`);
        return lugar_nombre
    } catch (error) {

        console.error('Error al generar el documento:', error.message);
    }
}

function generarDocumentoPlantilla(datos, nombreArchivoSalida) {
    try {
        
        // Verificar si la plantilla existe
        if (!fs.existsSync(rutaPlantillaRetiros)) {
            throw new Error(`La plantilla "${rutaPlantillaRetiros}" no existe.`);
        }
        const contenidoPlantilla = fs.readFileSync(rutaPlantillaRetiros,"binary");

        const zip = new PizZip(contenidoPlantilla);
        const generadorPlantillas = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });

        generadorPlantillas.render(datos);

        const buf = generadorPlantillas.getZip().generate({
            type: "nodebuffer",
            compression: "DEFLATE",
        });
        let lugar_nombre = Lugar_wordsActivos+nombreArchivoSalida
        fs.writeFileSync(lugar_nombre, buf);

        //console.log(`El documento "${lugar_nombre}" ha sido generado con éxito.`);
        return lugar_nombre
    } catch (error) {

        console.error('Error al generar el documento:', error.message);
    }
}

function generarDocumentoPlantillaactivo(datos, nombreArchivoSalida) {
    try {
        
        // Verificar si la plantilla existe
        if (!fs.existsSync(rutaPlantillaActivos)) {
            throw new Error(`La plantilla "${rutaPlantillaActivos}" no existe.`);
        }
        const contenidoPlantilla = fs.readFileSync(rutaPlantillaActivos,"binary");

        const zip = new PizZip(contenidoPlantilla);
        const generadorPlantillas = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });

        generadorPlantillas.render(datos);

        const buf = generadorPlantillas.getZip().generate({
            type: "nodebuffer",
            compression: "DEFLATE",
        });
        let lugar_nombre = Lugar_wordsActivos+nombreArchivoSalida
        fs.writeFileSync(lugar_nombre, buf);

        //console.log(`El documento "${lugar_nombre}" ha sido generado con éxito.`);
        return lugar_nombre
    } catch (error) {

        console.error('Error al generar el documento:', error.message);
    }
}

function formatearFecha(fecha) {
    // Dividir la cadena de fecha en día, mes y año
    let partes = fecha.split('/');
    let dia = parseInt(partes[0], 10);
    let mes = parseInt(partes[1], 10) - 1; // Restamos 1 al mes ya que en JavaScript los meses van de 0 a 11
    let anio = parseInt(partes[2], 10);
  
    // Crear un nuevo objeto de fecha
    let fechaObj = new Date(anio, mes, dia);
    //console.log('****')
    //console.log(fechaObj)
    // Verificar si la fecha es válida
    if (isNaN(fechaObj.getTime())) {
      return "Fecha no válida";
    }
  
    // Definir opciones de formato
    let opciones = { year: 'numeric', month: 'long', day: 'numeric' };
  
    // Formatear la fecha
    let fechaFormatoLargo = fechaObj.toLocaleDateString('es-ES', opciones);
  
    return fechaFormatoLargo;
}
  

function TuplaExcel(tupla){
    
    tupla['FECHA INGRESO'] = numeroSerieAfecha(tupla['FECHA INGRESO']);
    tupla['FECHA RETIRO'] = numeroSerieAfecha(tupla['FECHA RETIRO']);
    let SUELDO_LETRAS = numeroALetras.NumerosALetras(tupla['SUELDO']);
    let partes = SUELDO_LETRAS.split(' ');
    let resultadoFinal = partes.slice(0, -2).join(' ');
    tupla['SUELDO LETRAS'] = resultadoFinal;

    if (tupla['PROYECTO'] == 'UNE') {
        tupla['TIPO CONTRATO']='en la ejecución del contrato No. 4220001314 suscrito entre Energía Integral Andina S.A. y UNE EPM Telecomunicaciones S.A, '
    }
    else if(tupla['PROYECTO'] == 'EDATEL'){
        tupla['TIPO CONTRATO']=''
    }
    else{
        tupla['TIPO CONTRATO']=''
    }
    tupla['CARGO'] = tupla['CARGO'].trim();
    tupla['CONTRATO'] = tupla['CONTRATO'].trim();
    tupla['NOMINA'] = tupla['NOMINA'].trim();
    if (tupla['MOTIVO']){
        tupla['MOTIVO'] = tupla['MOTIVO'].trim();
    }

    let opciones = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    let fechaFormatoLargo = fechaActual.toLocaleDateString('es-ES', opciones);
    tupla['FECHA ACTUAL'] = fechaFormatoLargo
    tupla['FECHA INGRESO']= formatearFecha(tupla['FECHA INGRESO'])
    tupla['FECHA RETIRO'] = formatearFecha(tupla['FECHA RETIRO'])
    
    if (tupla['MOTIVO'] == 'TERMINACION CONTRATO JUSTA CAUSA' || tupla['MOTIVO'] == 'TERMINACION CONTRATO SIN JUSTA CAUSA' 
            || tupla['MOTIVO'] == 'FALLECIMIENTO'){
        tupla['MOTIVO'] = ''
        tupla['ANTE MOTIVO']=''
    }else{
        tupla['ANTE MOTIVO']=', Motivo del retiro: '
    }
    const añoActual = fechaActual.getFullYear();
    tupla['año'] = añoActual
    console.log(tupla)

    return tupla
}

async function extractAndSplitText(docxFilePath) {
    return mammoth.extractRawText({ path: docxFilePath })
        .then((result) => {
            const text = result.value; // Texto extraído del documento DOCX
            const lines = text.split("\n"); // Dividir el texto en líneas separadas
            // Filtrar líneas que contienen información significativa
            const filteredLines = lines.filter(line => line.trim() !== '');
            return filteredLines;
        })
        .catch((err) => {
            console.error("Error al extraer texto del documento DOCX:", err);
            throw err;
        });
}

async function extractAndSplitTextColilla(docxFilePath) {
    let mainList = [];
    let secondList = [];
    let isAddingToSecondList = false;
  
    return mammoth.extractRawText({ path: docxFilePath })
        .then((result) => {
            const text = result.value; // Texto extraído del documento DOCX
            const lines = text.split("\n"); // Dividir el texto en líneas separadas
            
            // Filtrar líneas que contienen información significativa
            const filteredLines = lines.filter(line => line.trim() !== '');
  
            for (const line of filteredLines) {
                if (line.includes('VALOR ') || line.includes('TOTAL PAGOS ')) {
                    mainList.push(line); // Agregar 'VALOR ' y 'TOTAL PAGOS ' a mainList
                    isAddingToSecondList = line.includes('VALOR '); // Activa la bandera para secondList
                } else if (isAddingToSecondList) {
                    secondList.push(line);
                } else {
                    mainList.push(line);
                }
            }
            //console.log(mainList)
            //console.log(secondList)
            return { mainList, secondList };
        })
        .catch((err) => {
            console.error("Error al extraer texto del documento DOCX:", err);
            throw err;
        });
}

async function create_pdf_colilla(dictionarymain,result,lugar_pdf){

    return new Promise((resolve, reject) => {
        const doc = new PDFDocumentKit();
        lugar_pdf=lugar_pdf+'.pdf'
        //console.log(lugar_pdf)
        // Pipe el PDF a un archivo
        const stream = fs.createWriteStream(lugar_pdf);
        doc.pipe(stream);
        doc.lineJoin('miter')//cuadrado
        .rect(50, 27, 520, 400) //ancho largo
        .stroke();
        
        //doc.moveDown();
        doc.font('Helvetica-Bold')//Energia
            .fontSize(8)
            .text(dictionarymain['line_1'],0,30, {
            width: 590,
            align: 'center'
        });

        doc.font('Helvetica-Bold')//nomina
            .fontSize(8)
            .text(dictionarymain['line_2'],0,50, {
            width: 490,
            align: 'center'
        });
    
        doc.lineJoin('miter')//linea
        .rect(50, 60, 520, 0)
        .stroke();
    
        doc.lineJoin('miter')//linea
        .rect(50, 105, 520, 0)
        .stroke();
    
        doc.font('Helvetica-Bold')//codigo
            .fontSize(8)
            .text(dictionarymain['line_3'],0,65, {
        width: 505,
        align: 'center'
        });
    
        doc.font('Helvetica-Bold')//identificacion
            .fontSize(8)
            .text(dictionarymain['line_4'],0,80, {
        width: 470,
        align: 'center'
        });
    
        doc.font('Helvetica-Bold')//cargo
            .fontSize(8)
            .text(dictionarymain['line_5'],0,95, {
        width: 290,
        align: 'center'
        });
    
        doc.lineJoin('miter')//linea
        .rect(50, 135, 520, 0)
        .stroke();
    
        doc.font('Helvetica-Bold')//pago
            .fontSize(8)
            .text(dictionarymain['line_6'],0,110, {
        width: 370,
        align: 'center'
        });
    
        doc.font('Helvetica-Bold')//Descuentos
            .fontSize(8)
            .text(dictionarymain['line_7'],0,110, {
        width: 960,
        align: 'center'
        });
    
        doc.font('Helvetica-Bold')//info
            .fontSize(8)
            .text(dictionarymain['line_8'],0,125, {
        width: 130,
        align: 'center'
        });
    
        doc.font('Helvetica-Bold')//DESCRIPCION
            .fontSize(8)
            .text(dictionarymain['line_9'],0,125, {
        width: 240,
        align: 'center'
        });
    
        doc.font('Helvetica-Bold')//Valor
            .fontSize(8)
            .text(dictionarymain['line_10'],0,125, {
        width: 560,
        align: 'center'
        });
    
        doc.font('Helvetica-Bold')//DESCRIPCION
            .fontSize(8)
            .text(dictionarymain['line_9'],0,125, {
        width: 710,
        align: 'center'
        });
    
        doc.font('Helvetica-Bold')//Valor
            .fontSize(8)
            .text(dictionarymain['line_10'],0,125, {
        width: 1060,
        align: 'center'
        });
    
        doc.lineJoin('miter')//linea_media
        .rect(320, 135, 0, 254)
        .stroke();
        let pago = result.Pago
        let Des = result.Des
        x = 150
        // Usando un bucle forEach para recorrer la lista
        pago.forEach(diccionario => {
        
        doc.font('Helvetica-Bold')//CONCEPTO
            .fontSize(8)
            .text(diccionario['CON'],95,x, {
            width: 500,
            align: 'left'
        });
    
        doc.font('Helvetica-Bold')//CANTIDAD
        .fontSize(8)
        .text(diccionario['CAN'],60,x, {
        width: 500,
        align: 'left'
        });
    
        doc.font('Helvetica-Bold')//VALOR PAGO 
        .text(diccionario['ValorPago'],267,x, {
        width: 500,
        align: 'left'
        });
    
        x+= 15
        });
        x = 150
        
        Des.forEach(diccionario => {
        //console.log(diccionario)
        doc.font('Helvetica-Bold')//Concepto
        .text(diccionario['CON'],330,x, {
        width: 500,
        align: 'left'
        });
    
        //console.log(diccionario)
        doc.font('Helvetica-Bold')//valor Descuento
        .text(diccionario['ValorPago'],516,x, {
        width: 500,
        align: 'left'
        });
    
        x+= 15
        });
    
        doc.font('Helvetica-Bold')//TOTAL PAGOS
            .fontSize(8)
            .text(dictionarymain['line_13'],0,380, {
        width: 351,
        align: 'center'
        });
    
        doc.font('Helvetica-Bold')//TOTAL DESCUENTOS
            .fontSize(8)
            .text(dictionarymain['line_14'],0,380, {
        width: 880,
        align: 'center'
        });
    
        doc.font('Helvetica-Bold')//NETO A PAGAR
            .fontSize(8)
            .text(dictionarymain['line_15'],0,400, {
        width: 390,
        align: 'center'
        });
    
        doc.lineJoin('miter')//linea
        .rect(50, 390, 520, 0)
        .stroke();
    
        doc.end();
    
        stream.on('finish', () => {
            console.log('El archivo PDF se ha generado correctamente.');
            resolve();
        });

        stream.on('error', (error) => {
            console.error('Error al generar el archivo PDF:', error);
            reject(error);
        });
    });
}

function ConvertToPDF_colilla(nombre_archivo,datos,cc) {

    return new Promise((resolve, reject) => {
        // Construir rutas de archivos
        const lugar_word = Lugar_wordsActivos + nombre_archivo;
        const lugar_pdf = Lugar_Pdfs + cc;

        //Uso de la función para extraer y dividir el texto del documento DOCX
        extractAndSplitTextColilla(lugar_word)
            .then((lines) => {
                console.log("Líneas de texto extraídas del documento DOCX:");
        
                //console.log(lines);
                const mainList =lines['mainList']
                //console.log(lines1);
                const dictionarymain = {}; 
                mainList.forEach((line, index) => {
                    // Asignar cada línea a una clave única en el diccionario
                    dictionarymain[`line_${index + 1}`] = line;
                });

                //const result = consultarCedulaEnExcel_colilla('1038136076')

                create_pdf_colilla(dictionarymain,datos,lugar_pdf)
                resolve(lugar_pdf);
            })
            .catch((err) => {
                console.error("Error al extraer y dividir el texto:", err);
                reject(err)
            });
    });
}

async function create_pdf(dictionary, lugar_pdf) {
    return new Promise((resolve, reject) => {
        const doc = new PDFDocumentKit();
        //console.log(dictionary)
        // Pipe el PDF a un archivo
        const stream = fs.createWriteStream(lugar_pdf);
        doc.pipe(stream);

        doc.image('../Plantillas/encabezado.png', 0, 34, { width: 620, height: 120 });
        doc.moveDown(5);
        doc.font('Helvetica-Bold')//consecutivo1
            .fontSize(12)
            .text(dictionary['line_1'], {
                width: 500,
                align: 'right',
            });
        doc.moveDown(3);
        doc.font('Helvetica-Bold')//TITULO
            .fontSize(13)
            .text(dictionary['line_2'], {
                width: 470,
                align: 'center',
            });

        doc.font('Helvetica-Bold')//nit
            .fontSize(13)
            .text(dictionary['line_3'], {
                width: 470,
                align: 'center',
            });

        doc.moveDown(3);
        doc.font('Helvetica-Bold')//certifica
            .fontSize(13)
            .text(dictionary['line_4'], {
                width: 470,
                align: 'center',
            });

        doc.moveDown(2);
        doc.font('Helvetica')//texto completo
            .fontSize(11)
            .text(dictionary['line_5'], {
                width: 410,
                align: 'justify'
        });

        doc.moveDown(1);
        doc.font('Helvetica')//texto auxiliar
        .fontSize(11)
        .text(dictionary['line_6'], {
          width: 410,
          align: 'justify'
        }
        );
     
        doc.moveDown(2);
        doc.font('Helvetica')//Coordial
        .fontSize(11)
        .text(dictionary['line_7'], {
          width: 410,
          align: 'justify'
        }
        );
     
        doc.image('../Plantillas/Imagen2.png', 60, 500, {width: 100, height: 50})
        doc.text('___________________________________________', 70, 550)
     
        doc.font('Helvetica')//Nombre
        .fontSize(10)
        .text(dictionary['line_9'], {
          width: 410,
          align: 'justify'
        }
        );
     
        doc.font('Helvetica')//Consecutivo2
        .fontSize(11)
        .text(dictionary['line_10'], {
          width: 410,
          align: 'justify'
        }
        );
     
        doc.font('Helvetica')//cargo
        .fontSize(11)
        .text(dictionary['line_11'], {
          width: 410,
          align: 'justify'
        }
        );
     
        doc.font('Helvetica')//proyecto
        .fontSize(11)
        .text(dictionary['line_12'], {
          width: 410,
          align: 'justify'
        }
        );
     
        doc.font('Helvetica')//EIA
        .fontSize(11)
        .text(dictionary['line_13'], {
          width: 410,
          align: 'justify'
        }
        );
     
        doc.moveDown(4);
        doc.font('Helvetica')//validación
        .fontSize(10)
        .text(dictionary['line_14'], {
          width: 410,
          align: 'justify'
        }
        );

        doc.end();

        stream.on('finish', () => {
            console.log('El archivo PDF se ha generado correctamente.');
            resolve();
        });

        stream.on('error', (error) => {
            console.error('Error al generar el archivo PDF:', error);
            reject(error);
        });
    });
}

function buildDictionary(lines) {
    const dictionary = {};
    lines.forEach((line, index) => {
        dictionary[`line_${index + 1}`] = line;
    });
    return dictionary;
}

function ConvertToPDF(nombre_archivo) {
    return new Promise((resolve, reject) => {
        // Construir rutas de archivos
        const lugar_word = Lugar_wordsActivos + nombre_archivo;
        const lugar_pdf = Lugar_Pdfs + nombre_archivo;

        // Extraer texto del archivo Word y crear PDF
        extractAndSplitText(lugar_word)
            .then((lines) => {
                // Construir diccionario de líneas de texto
                const dictionary = buildDictionary(lines);
                // Crear PDF a partir del diccionario y la ubicación del PDF
                return create_pdf(dictionary, lugar_pdf);
            })
            .then(() => {
                console.log("PDF creado exitosamente en:", lugar_pdf);
                resolve(lugar_pdf);
            })
            .catch((err) => {
                console.error("Error al convertir el archivo a PDF:", err);
                reject(err);
            });
    });
}

const unificarPDFs = async (archivos, CC) => {
    try {
        const pdfDoc = await PDFDocument.create();
        
        for (const archivo of archivos) {
            const pdfBytes = fs.readFileSync(archivo);
            const pdf = await PDFDocument.load(pdfBytes);
            const paginas = await pdfDoc.copyPages(pdf, pdf.getPageIndices());
            paginas.forEach((pagina) => pdfDoc.addPage(pagina));
        }
        const rutaPDF = `../PDFs/${CC}.pdf`;
        const pdfBytes = await pdfDoc.save();
        //console.log('ruta de pdf');
        //console.log(rutaPDF);
        if (!fs.existsSync(rutaPDF)) {
            // Si el archivo no existe, crearlo con contenido vacío
            //console.log('ruta creada');
            fs.writeFileSync(rutaPDF, '');
        }
        fs.writeFileSync(rutaPDF, pdfBytes);
        console.log('Archivos PDF unificados con éxito');
        return rutaPDF;
    } catch (error) {
        const rutaPDF = `../PDFs/${CC}.pdf`;
        console.error('Error al unificar archivos PDF:', error);
        return rutaPDF;
    }
};

async function ReadExcel_activo(identification,num) {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const worksheet = workbook.getWorksheet(num);
        let rowCount = 0;
        if (num == 2){
            worksheet.eachRow((row, rowNumber) => {
                if (rowNumber > 1) { // Ignorar la primera fila de encabezados
                    const id = row.getCell('A').value;
                    let id_num=id.toString();
                    const fechaCellValue = row.getCell('C').value;
                    // Convertir el valor de la celda de fecha a un objeto Date de JavaScript
                    const parts = fechaCellValue.split('/');
                    const monthString = parts[1];
                    const month = parseInt(monthString, 10) - 1;
                    const today = new Date();
                    const currentMonth = today.getMonth();
                    // Filtrar por identificación y fecha del mes actual
                    if (id_num === identification && month === currentMonth) {
                        rowCount++;
                    }
                }
            });
            if(rowCount >= 1){
                return false
            }else{
                return true
            }
        }else if(num == 3){
            // Definimos la variable identification que contiene el número de identificación a filtrar
            //let identification = '123456789'; // Por ejemplo, el número de identificación a filtrar

            // Definimos las variables para almacenar la información del último registro encontrado
            let ultimoRegistro = null;

            worksheet.eachRow((row, rowNumber) => {
                // Obtener el número de identificación de la celda A
                const id = row.getCell('A').value;
                let id_num=id.toString();
                // Verificar si el número de identificación coincide con el que estamos buscando
                if (id_num === identification) {
                    // Si coincide, guardamos toda la fila como el último registro
                    ultimoRegistro = row;
                }
            });

            if (ultimoRegistro) {
                // Obtener la fecha de la celda C del último registro
                const fechaCelda = new Date(ultimoRegistro.getCell('C').value);
                const fechaActual = new Date();

                // Determinar el corte actual
                const diaActual = fechaActual.getDate();
                const corteActual = diaActual <= 15 ? 1 : 2;

                // Determinar el corte de la fecha de la celda
                const diaCelda = fechaCelda.getDate();
                const corteCelda = diaCelda <= 15 ? 1 : 2;
                let flag
                // Verificar si pertenecen al mismo corte
                const mismoCorte = corteActual === corteCelda;
                if (mismoCorte){
                    flag = false
                }
                else{
                    flag = true
                }
                // Mostrar información
                //console.log(`La fecha de la celda pertenece al corte ${corteCelda}`);
                //console.log(`La fecha actual pertenece al corte ${corteActual}`);
                //console.log(`¿Pertenecen al mismo corte? ${mismoCorte ? 'Sí' : 'No'}`);
                return flag
            } else {
                //console.log('No se encontraron registros para la identificación dada.');
                return true;
            }
        }

    } catch (error) {
        console.error('Error al consultar datos:', error);
        return false;
    }
}

async function ReadExcel_Inactivo(identification) {
    try {

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const worksheet = workbook.getWorksheet(2);
        let hasRecordInLastTwoMonths = false;
        
        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) { // Ignorar la primera fila de encabezados
                const id = row.getCell('A').value;
                const fechaCellValue = row.getCell('C').value;
                // Convertir el valor de la celda de fecha a un objeto Date de JavaScript
                const parts = fechaCellValue.split('/');
                const monthString = parts[1];
                const month = parseInt(monthString, 10) - 1;
                const today = new Date();
        
                // Filtrar por identificación y fecha de los últimos 2 meses
                const isWithinLastTwoMonths =
                    (today - new Date(parts[2], month, parts[0])) / (1000 * 60 * 60 * 24 * 30.44) < 2;
        
                if (id === identification && isWithinLastTwoMonths) {
                    hasRecordInLastTwoMonths = true;
                }
            }
        });
        
        // Si tiene al menos un registro en los últimos 2 meses, devuelve false; de lo contrario, devuelve true
        return !hasRecordInLastTwoMonths;

    } catch (error) {
        console.error('Error al consultar datos:', error);
        return false;
    }
}

async function FindConsecutivo() {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const worksheet = workbook.getWorksheet(2);

        let ultimaFila = null;

        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) { // Ignorar la primera fila de encabezados
                ultimaFila = row;
            }
        });

        if (ultimaFila) {
            const valorColumnaA = ultimaFila.getCell('D').value;
            let numero = parseInt(valorColumnaA);
            return  numero;
        } else {
            console.log('No se encontraron filas en el archivo.');
            return null;
        }
    } catch (error) {
        console.error('Error al leer el archivo Excel:', error);
        return null;
    }
}

async function writeToExcelMultiple(dataList,num) {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const worksheet = workbook.getWorksheet(num);
        //console.log('datalist')
        console.log('????????????')
        console.log(dataList)
        console.log('????????????')

        if (dataList.length > 1) {
            
            //console.log(dataList)
            for (const data of dataList) {
                // Obtén la última fila con datos y agrega 1 para obtener la siguiente fila
                const newRowNumber = worksheet.rowCount + 1;
                let infoFecha = {
                    dia: fechaActual.getDate(),
                    mes: fechaActual.getMonth() + 1,
                    anio: fechaActual.getFullYear()
                };
                let fechaEnTexto = `${infoFecha.dia}/${infoFecha.mes}/${infoFecha.anio}`;

                // Agrega datos a la nueva fila
                const newRow = worksheet.getRow(newRowNumber);
                newRow.getCell('A').value = data.IDENTIFICACION;
                newRow.getCell('B').value = data['NOMBRE COMPLETO'].trim();
                newRow.getCell('C').value = fechaEnTexto;
                newRow.getCell('D').value = data['Consecutivo1'];
            }
        } else
        {
            const newRowNumber = worksheet.rowCount + 1;
            let infoFecha = {
                dia: fechaActual.getDate(),
                mes: fechaActual.getMonth() + 1,
                anio: fechaActual.getFullYear()
            };
            let fechaEnTexto = `${infoFecha.dia}/${infoFecha.mes}/${infoFecha.anio}`;

            // Agrega datos a la nueva fila
            const newRow = worksheet.getRow(newRowNumber);
            newRow.getCell('A').value = dataList.IDENTIFICACION;
            newRow.getCell('B').value = dataList['NOMBRE COMPLETO'];
            newRow.getCell('C').value = fechaEnTexto;
        }

        // Guarda los cambios en el archivo Excel
        await workbook.xlsx.writeFile(filePath);
        console.log('Todos los datos han sido agregados exitosamente.');
    } catch (error) {
        console.error('Error al agregar datos:', error);
    }
}

async function send_mail(correo,cc){ //revisar el adjunto
    // Configuración del transporte SMTP con la contraseña de aplicación
    //console.log('send_mail')
    let transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
            user: 'andres.genes@eiasa.com.co', // Coloca aquí tu dirección de correo electrónico
            pass: 'ndwy tbjb bprq kdhr' // Coloca aquí tu contraseña de aplicación
        }
    });
    let primerElemento = filas[0];
    //console.log(primerElemento)
    let mensaje = `la persona con nombre ${primerElemento['NOMBRE COMPLETO']} y cedula: ${primerElemento.IDENTIFICACION}, requiere una carta laborar en el siguiente correo: ${correo}`;
    
    // Opciones del correo electrónico
    let mailOptions = {
        from: 'andres.genes@eiasa.com.co', // Remitente
        to: 'Gestionhumana.med@eiasa.com.co', // Destinatario
        subject: 'Solicitud de envio de carta laborar', // Asunto
        text: mensaje, // Cuerpo del correo
        attachments: [
            {
                filename: 'archivo_adjunto.pdf', // Nombre del archivo adjunto
                path: 'C:/Users/tecnologo.operacag/Desktop/chatv2.0/PDFs/${cc}.pdf' // Ruta al archivo adjunto
            }//C:/Users/tecnologo.operacag/Desktop/chatv2.0/PDFs/${myState.NumeroCedula}.pdf
        ]
    };

    // Envío del correo electrónico
    transporter.sendMail(mailOptions, function(error, info){
        if (error) {
            console.log(error);
        } else {
            console.log('Correo electrónico enviado: ' + info.response);
        }
    });
}

module.exports = {
    generarDocumentoPlantilla: generarDocumentoPlantilla,
    numeroSerieAfecha: numeroSerieAfecha,
    consultarCedulaEnExcel: consultarCedulaEnExcel,
    activo_excel: activo_excel,
    unificarPDFs:unificarPDFs,
    ReadExcel_activo:ReadExcel_activo,
    consultarNumeroInterno:consultarNumeroInterno,
    consultarNumeroCuenta,consultarNumeroCuenta,
    ReadExcel_Inactivo:ReadExcel_Inactivo,
    send_mail:send_mail,
    excel_colillas:excel_colillas,
    consultarCedulaEnExcel_colilla:consultarCedulaEnExcel_colilla,
    writeToExcelMultiple:writeToExcelMultiple
};


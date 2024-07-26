const fs = require('node:fs')
const { PDFDocument, rgb } = require('pdf-lib');
/**
 *
 * @param {*} voiceId clone voice vwfl76D5KBjKuSGfTbLB
 * @returns
 */
const unificarPDFs = async (archivos, CC) => {
    try {
        const pdfDoc = await PDFDocument.create();
        
        for (const archivo of archivos) {
            try {
                const pdfBytes = fs.readFileSync(archivo);
                const pdf = await PDFDocument.load(pdfBytes);
                const paginas = await pdfDoc.copyPages(pdf, pdf.getPageIndices());
                paginas.forEach((pagina) => pdfDoc.addPage(pagina));
            } catch (error) {
                console.error(`Error al procesar el archivo ${archivo}:`, error);
            }
        }

        const rutaPDF = `C:/Users/tecnologo.operacag/Desktop/chat/PDFs/${CC}.pdf`;
        const pdfBytes = await pdfDoc.save();
        console.log('ruta de pdf');
        console.log(rutaPDF);
        if (!fs.existsSync(rutaPDF)) {
            // Si el archivo no existe, crearlo con contenido vacío
            console.log('ruta creada');
            fs.writeFileSync(rutaPDF, '');
        }
        fs.writeFileSync(rutaPDF, pdfBytes);
        //console.log('Archivos PDF unificados con éxito');
        return rutaPDF;
    } catch (error) {
        
        const rutaPDF = `C:/Users/tecnologo.operacag/Desktop/chat/PDFs/${CC}.pdf`;
        console.error('Error al unificar archivos PDF:', error);
        return rutaPDF;
        //throw error;
    }
};


module.exports = { unificarPDFs };
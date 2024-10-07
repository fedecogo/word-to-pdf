// Importa le librerie necessarie
const Docxtemplater = require("docxtemplater");
const PizZip = require("pizzip");
const fs = require("fs");
const path = require("path"); 
const mammoth = require("mammoth");
const puppeteer = require("puppeteer");

// Carica il file 
const content = fs.readFileSync(
    path.resolve(__dirname, "input.docx"),
    "binary" // Leggi il file come contenuto binario
);

// Decompatta il file
const zip = new PizZip(content);

// Analizza il template. 
const doc = new Docxtemplater(zip, {
    paragraphLoop: true, // Permette di iterare sui paragrafi
    linebreaks: true,    // Gestisce i ritorni a capo
});

// Render del documento sostituendo i segnaposto con i valori forniti
doc.render({
    first_name: "Mario",
    last_name: "Rossi",
});

// Ottiene il documento come un file zip (i docx sono file compressi)
// e lo genera come un buffer di Node.js e scrive il buffer Node.js in un file DOCX
const buf = doc.getZip().generate({
    type: "nodebuffer",  // Tipo di output come buffer Node.js
    compression: "DEFLATE", // Applica la compressione per ridurre le dimensioni del file
});
const docxPath = path.resolve(__dirname, "output.docx");
fs.writeFileSync(docxPath, buf); // Salva il file DOCX generato


// Funzione per convertire un file .docx in HTML utilizzando Mammoth
async function convertDocxToHtml(docxPath) {
    const result = await mammoth.convertToHtml({ path: docxPath }); // Converte in HTML
    return result.value; // Ritorna la stringa HTML
}

// Funzione per convertire HTML in PDF utilizzando Puppeteer
async function convertHtmlToPdf(html, pdfPath) {
    const browser = await puppeteer.launch(); // Avvia il browser headless
    const page = await browser.newPage(); // Crea una nuova pagina
    
    // Carica l'HTML nel browser headless
    await page.setContent(html, { waitUntil: "domcontentloaded" });
    
    // Genera il PDF dal contenuto HTML
    await page.pdf({
        path: pdfPath,        // Percorso di salvataggio del PDF
        format: "A4",        // Formato del PDF
        printBackground: true, // Stampa anche lo sfondo
    });
    
    await browser.close(); // Chiude il browser
}

// Funzione principale per gestire la conversione da .docx a .pdf
async function convertDocxToPdf(docxPath) {
    const pdfPath = path.resolve(__dirname, "output.pdf"); // Percorso di salvataggio del PDF
    try {
        const html = await convertDocxToHtml(docxPath); // Converte DOCX in HTML
        await convertHtmlToPdf(html, pdfPath); // Converte HTML in PDF
        console.log(`File PDF salvato in: ${pdfPath}`); // Messaggio di conferma
    } catch (error) {
        console.error("Errore durante la conversione:", error); // Gestione degli errori
    }
}

// Avvia la conversione da DOCX a PDF
convertDocxToPdf(docxPath);

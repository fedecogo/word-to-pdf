const Docxtemplater = require("docxtemplater");
const PizZip = require("pizzip");
const fs = require("fs");
const path = require("path");
const libre = require("libreoffice-convert"); // Per convertire con LibreOffice

// Carica il file .docx
const content = fs.readFileSync(
    path.resolve(__dirname, "input.docx"),
    "binary"
);

// Decompatta il file .docx
const zip = new PizZip(content);

// Analizza il template
const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
});

// Sostituisce i segnaposto con i valori forniti
doc.render({
   // subscription.user.name
   first_name: "Mario",
   // subscription.user.surname
   last_name: "Rossi",
   // subscription.user.tax_code
   tax_code: "11111111111",
  
  // subscription.registry.surname
   register_surname: "Verdi",

   // subscription.registry.name
   register_name: "Giuseppe",

   // subscription.registry.birth_address.country_of_birth.name
   register_birth_country: "Italia",

   // subscription.registry.date_birth
   register_date_birth: "01/01/2000",

   // subscription.registry.tax_id_code
   register_tax_code: "222222",

   // subscription.registry.address.city.name
   register_city: "Roma",

   // subscription.registry.address.province.id
   register_province: "RO",

   // subscription.registry.address.street
   register_street: "Via del Corso", 

   // subscription.registry.address.street_number
   register_street_number: "1",
});

// Genera il documento come un buffer e salva il .docx
const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
});

const docxPath = path.resolve(__dirname, "output.docx");
fs.writeFileSync(docxPath, buf); // Salva il file .docx generato

// Funzione per convertire un file .docx in PDF utilizzando LibreOffice-convert
async function convertDocxToPdf(docxPath) {
    const pdfPath = path.resolve(__dirname, "output.pdf"); // Percorso di output PDF
    const docxBuffer = fs.readFileSync(docxPath); // Leggi il file .docx

    // Esegui la conversione usando libreoffice-convert
    libre.convert(docxBuffer, '.pdf', undefined, (err, pdfBuffer) => {
        if (err) {
            console.error(`Errore durante la conversione: ${err.message}`);
            return;
        }

        // Salva il buffer PDF in un file
        fs.writeFileSync(pdfPath, pdfBuffer);
        console.log(`File PDF salvato in: ${pdfPath}`);
    });
}

// Avvia la conversione da .docx a PDF
convertDocxToPdf(docxPath);

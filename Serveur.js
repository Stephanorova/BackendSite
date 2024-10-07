const express = require('express');
const mongoose = require('mongoose');
const multer = require('multer');
const xlsx = require('xlsx');
const PDFDocument = require('pdfkit');
const fs = require('fs');
const bodyParser = require('body-parser')
require('dotenv/config')
const app = express();
const cors = require('cors')
const path = require('path');
app.use(cors())
app.use(express.json())
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));    
app.use(express.json());

const upload = multer({ dest: 'uploads/' });

const certificatSchema = new mongoose.Schema({
    NomEtPrenomEtudient: String,
    DateFormation: String,
    Ref: String,
    Num: String,
    Formation: String,
    Societe: String
});

const Certificat = mongoose.model('certificat', certificatSchema);

const pdfDir = path.join(__dirname, 'pdfs');
if (!fs.existsSync(pdfDir)){
    fs.mkdirSync(pdfDir);
}

app.post('/import', upload.single('file'), async (req, res) => {
  try {
    const workbook = xlsx.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet);

    let importedCount = 0;
    let errorCount = 0;

    for (let row of data) {
      try {
        // Vérifiez que toutes les valeurs nécessaires sont présentes
        if (!row.NomEtPrenomEtudient || !row.DateFormation || !row.Ref || !row.Num || !row.Formation || !row.Societe) {
          console.log('Données manquantes pour la ligne:', row);
          errorCount++;
          continue; // Passez à la ligne suivante si des données sont manquantes
        }

        const pdfDoc = new PDFDocument();
        const pdfPath = path.join(pdfDir, `${row.Num}.pdf`);
        pdfDoc.pipe(fs.createWriteStream(pdfPath));


        pdfDoc.text(`NomEtPrenomEtudient: ${row.NomEtPrenomEtudient}`);
        pdfDoc.text(`Date de formation: ${row.DateFormation}`);
        pdfDoc.text(`Ref: ${row.Ref}`);
        pdfDoc.text(`Num: ${row.Num}`);
        pdfDoc.text(`Formation: ${row.Formation}`);
        pdfDoc.text(`Societe: ${row.Societe}`);
        pdfDoc.end();

        const certificat = new Certificat({
          NomEtPrenomEtudient: row.NomEtPrenomEtudient,
          DateFormation: row.DateFormation,
          Ref: row.Ref,
          Num: row.Num,
          Formation: row.Formation,
          Societe: row.Societe,
          pdfDir
        });

        await certificat.save();
        importedCount++;
      } catch (error) {
        console.error('Erreur lors du traitement de la ligne:', row, error);
        errorCount++;
      }
    }

    res.json({ message: 'Import terminé', importedCount, errorCount });
  } catch (error) {
    console.error('Erreur lors de l\'import:', error);
    res.status(500).json({ message: 'Erreur lors de l\'import', error: error.message });
  }
});

app.get('/verifier/:Ref', async (req, res) => {
    const certificat = await Certificat.findOne({ Ref: req.params.Ref });
    if (certificat) {
      res.json(certificat);
    } else {
      res.status(404).send('Certificat non trouvé');
    }
  });

app.get('/telecharger/:Ref', async (req, res) => {
  const certificat = await Certificat.findOne({ Ref: req.params.Ref });
  if (certificat) {
    res.download(certificat.pdfPath);
  } else {
    res.status(404).send('Certificat non trouvé');
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Serveur démarré sur le port ${PORT}`));

// Connexion à MongoDB
const url = process.env.DATABASE;
mongoose.connect(url, {})
.then(()=>{
    console.log("connexion mongodb avec success");
    
})
.catch((error)=>{
    console.log(error);
    
})
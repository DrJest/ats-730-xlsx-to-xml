const express = require('express');
const multer = require('multer');
const archiver = require('archiver');
const xlsx = require('node-xlsx');
const builder = require('xmlbuilder');
const moment = require('moment');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

const publicKey = fs.readFileSync(path.resolve(__dirname + '/SanitelCF.cer'), 'utf8');

const parseDateExcel = excelTimestamp => {
  const excelEpoch = new Date(1899, 11, 31);
  let newDate = moment(excelEpoch).add(excelTimestamp - 1, 'days');
  return newDate.format('YYYY-MM-DD');
};

const cifraCF = cf => {
  var buffer = Buffer.from(cf);
  var encrypted = crypto.publicEncrypt(publicKey, buffer);
  return encrypted.toString("base64");
};

const formatCifra = cifra => {
  return Number(cifra).toFixed(2);
};

const convertXlsxToXml = (fileBuffer, codiceRegione, codiceAsl, codiceSSA) => {
  const src = xlsx.parse(fileBuffer);

  var root = builder.create('precompilata');

  var proprietario = root.ele('proprietario');
  proprietario.ele('codiceRegione', {}, codiceRegione);
  proprietario.ele('codiceAsl', {}, codiceAsl);
  proprietario.ele('codiceSSA', {}, codiceSSA);

  src.forEach(foglio => {
    foglio.data.forEach((riga, i) => {
      // 0 Item
      // 1 cfProprietario
      // 2 dataPagamento
      // 3 PagAnt
      // 4 FlagOperazione
      // 5 cfCittadino
      // 6 tipoSpesa
      // 7 Importo
      // 8 idRimborso
      // 9 pIva
      // 10 data emissione
      // 11 dispositivo
      // 12 NumDocumento

      if (!i || !riga[0]) return;
      var documentoSpesa = root.ele('documentoSpesa');
      var idSpesa = documentoSpesa.ele('idSpesa');
      idSpesa.ele('pIva', {}, riga[9]);
      idSpesa.ele('dataEmissione', {}, parseDateExcel(riga[10]));
      var numDocumentoFiscale = idSpesa.ele('numDocumentoFiscale');
      numDocumentoFiscale.ele('dispositivo', {}, riga[11]);
      numDocumentoFiscale.ele('numDocumento', {}, riga[12].replace('.', ''));
      if (riga[8]) {
        var idRimborso = documentoSpesa.ele('idRimborso');
        idRimborso.ele('pIva', {}, riga[9]);
        idRimborso.ele('dataEmissione', {}, parseDateExcel(riga[10]));
        var numDocumentoFiscaleRimborso = idRimborso.ele('numDocumentoFiscale');
        numDocumentoFiscaleRimborso.ele('dispositivo', {}, riga[11]);
        numDocumentoFiscaleRimborso.ele('numDocumento', {}, riga[8].replace('.', ''));
      }
      if (riga[3]) {
        documentoSpesa.ele('flagPagamentoAnticipato', {}, '1');
      }
      documentoSpesa.ele('dataPagamento', {}, parseDateExcel(riga[2]));
      documentoSpesa.ele('flagOperazione', {}, riga[4]);
      documentoSpesa.ele('cfCittadino', {}, cifraCF(riga[5]));
      const voceSpesa = documentoSpesa.ele('voceSpesa');
      voceSpesa.ele('tipoSpesa', riga[6]);
      voceSpesa.ele('importo', formatCifra(riga[7]));
    });
  });

  return root.end({ pretty: true });
};

app.use(express.static('public'));

app.post('/convert', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'Nessun file caricato' });
  }

  try {
    const codiceRegione = req.body.codiceRegione || '030';
    const codiceAsl = req.body.codiceAsl || '322';
    const codiceSSA = req.body.codiceSSA || '000763';

    const xmlContent = convertXlsxToXml(req.file.buffer, codiceRegione, codiceAsl, codiceSSA);
    const filename = `${codiceRegione}_${codiceAsl}_${codiceSSA}_730.xml`;

    // Creare archivio ZIP in memoria
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename="export.zip"`);

    const archive = archiver('zip', { zlib: { level: 9 } });
    archive.on('error', (err) => {
      res.status(500).json({ error: 'Errore nella creazione del ZIP' });
    });

    archive.pipe(res);
    archive.append(xmlContent, { name: filename });
    archive.finalize();

  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Errore durante la conversione: ' + error.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server in ascolto su http://localhost:${PORT}`);
});

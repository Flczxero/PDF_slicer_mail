const fs = require('fs');
const path = require('path');
const { PDFDocument } = require('pdf-lib');

async function testRemove() {
  try {
    const mainPdfPath = path.join(__dirname, '..', '1-s2.0-S2772443321000027-main.pdf');
    console.log('Reading PDF...');
    const pdfBytes = fs.readFileSync(mainPdfPath);
    const sourcePdf = await PDFDocument.load(pdfBytes);
    const totalPages = sourcePdf.getPageCount();
    
    // We want to keep ONLY page 0. Remove pages 1 to totalPages-1
    // Must remove in reverse order to avoid shifting indices!
    for (let i = totalPages - 1; i > 0; i--) {
      sourcePdf.removePage(i);
    }

    const targetPdfBytes = await sourcePdf.save();
    console.log('Test remove successful. Size:', targetPdfBytes.length);

  } catch (err) {
    console.error('Test failed:', err);
  }
}

testRemove();

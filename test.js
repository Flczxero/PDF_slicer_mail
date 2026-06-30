const fs = require('fs');
const path = require('path');
const { PDFDocument } = require('pdf-lib');
const xlsx = require('xlsx');

async function test() {
  try {
    const mainPdfPath = path.join(__dirname, '..', '1-s2.0-S2772443321000027-main.pdf');
    const dataFilePath = path.join(__dirname, '..', 'mail_lists.xlsx');
    const columnName = 'Email'; // Guessing a column name, we'll see if it works
    const pagesPerSlice = 1;

    console.log('Reading Excel data...');
    const workbook = xlsx.readFile(dataFilePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = xlsx.utils.sheet_to_json(worksheet);
    
    if (jsonData.length > 0) {
      console.log('Available columns in Excel:', Object.keys(jsonData[0]));
    } else {
      console.log('Excel file is empty');
      return;
    }

    const actualColumn = Object.keys(jsonData[0])[0];
    console.log('Using column for filenames:', actualColumn);

    const filenames = jsonData.map(row => {
      let name = row[actualColumn];
      if (name) {
        name = name.toString().trim();
        if (!name.toLowerCase().endsWith('.pdf')) {
          name += '.pdf';
        }
        return name;
      }
      return null;
    }).filter(Boolean);

    console.log(`Found ${filenames.length} filenames.`);
    if (filenames.length > 0) {
      console.log('First filename:', filenames[0]);
    }

    console.log('Reading PDF...');
    const sourcePdf = await PDFDocument.load(fs.readFileSync(mainPdfPath));
    const totalPages = sourcePdf.getPageCount();
    console.log(`Main PDF has ${totalPages} pages.`);

    // Just slice the first one to test
    const targetPdf = await PDFDocument.create();
    const endPage = Math.min(pagesPerSlice, totalPages);
    const pageIndices = Array.from({ length: endPage }, (_, index) => index);
    const copiedPages = await targetPdf.copyPages(sourcePdf, pageIndices);
    copiedPages.forEach((page) => targetPdf.addPage(page));

    const targetPdfBytes = await targetPdf.save();
    console.log('Test slice successful. Size:', targetPdfBytes.length);

  } catch (err) {
    console.error('Test failed:', err);
  }
}

test();

import { test } from '@playwright/test';
import fs from 'node:fs/promises';
import * as XLSX from 'xlsx';
import PDFParser from "pdf2json";

let workbook;

async function downloadPdf(page, firstPageUrl, fileName) {
  // const downloadPromise = await page.waitForEvent('download');

  await page.evaluate(async ([firstPageUrl, fileName]) => {
    const response = await fetch(firstPageUrl);
    const blob = await response.blob();
    const link = document.createElement('a');
    link.href = window.URL.createObjectURL(blob);
    link.setAttribute('download', fileName);
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    link.remove();
  }, [firstPageUrl, fileName]);

  await page.waitForEvent('download')
    .then(async download => await download.saveAs(`./downloads/${fileName}.pdf`));
    // .catch(error => {
    //   console.log('error downloading');
    //   console.log(error);
    // });
  // await download.saveAs('./downloads/' + download.suggestedFilename());
}

async function searchAbstractInPdf(articleIndex: number, fileName: string, firstWorksheet: XLSX.WorkSheet) {
  const pdfParser = new PDFParser(this, 1);

  let finishedProcessing = false;

  await pdfParser.loadPDF(`./downloads/${fileName}.pdf`);

  pdfParser.on("pdfParser_dataReady", pdfData => {
    const text = pdfParser.getRawTextContent();

    const lines = text.split('\n');
    const abstractLines: string[] = [];

    let abstractLineStart: number | null = null;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];

      if (/[0-9]+$/.test(line.trim()) && line.includes('Abstract')) {
        continue;
      }

      if (abstractLineStart === null && line.includes('Abstract')) {
        abstractLineStart = i;
      }

      if (abstractLineStart !== null && abstractLineStart === i - 1 && 
        (line.trim() === '') || line.trim().replace('.', '').replace(':', '') === 'Abstract') {

        continue;
      }

      if (abstractLineStart !== null && (line.trim() === '' || line.includes('Keywords'))) {
        break;
      }

      if (abstractLineStart !== null) {
        abstractLines.push(line.replace(/(\r\n|\n|\r)/gm, ''));
      }
    }

    XLSX.utils.sheet_add_aoa(firstWorksheet, [[abstractLines.join('')]], { origin: `E${articleIndex + 1}` });
    finishedProcessing = true;
  });

  while (!finishedProcessing) {
    await new Promise(resolve => setTimeout(resolve, 1000));
  }
}

test('has title', async ({ page }) => {
  const data = await fs.readFile('./input/articole-abstracte.xlsx');
  workbook = XLSX.read(data);

  const firstWorksheet: XLSX.WorkSheet = workbook.Sheets[workbook.SheetNames[0]];

  const rawData = XLSX.utils.sheet_to_json(firstWorksheet, { header: 1, defval: '-' });

  for (let i = 1; i < rawData.length; i++) {  
  // for (let i = 1; i <= 2; i++) {

    const fileName: string = rawData[i][2];
    const author: string = rawData[i][1];

    await page.goto(`https://google.com/search?q=${author} ${fileName}`);
    // await page.goto(`https://google.com/search?q=${fileName}`);

    const firstPageUrl = await page.locator('#rso > div:not(.ULSxyf) >> a').first().getAttribute('href');

    if (!firstPageUrl) {
      return;
    }

    if (firstPageUrl.endsWith('.pdf')) {
      try {
        await downloadPdf(page, firstPageUrl, fileName);
        await searchAbstractInPdf(i, fileName, firstWorksheet);
      } catch (error) {
        console.log('error downloading or searching pdf');
        console.log(error);
      }


    } else {
      console.log(firstPageUrl);
    }
  }
});

test.beforeAll(async () => {
  fs.readdir('./downloads').then(files => {
    for (const file of files) {
      fs.unlink('./downloads/' + file);
    }
  });

  fs.readdir('./output').then(files => {
    for (const file of files) {
      fs.unlink('./output/' + file);
    }
  });
});

test.afterAll(async () => {
  XLSX.writeFile(workbook, "./output/articole+abstracte.xlsx", { compression: true });
});
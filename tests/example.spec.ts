import { test } from '@playwright/test';
import fs from 'node:fs/promises';
import * as XLSX from 'xlsx';
import PDFParser from "pdf2json";

let workbook;

test('has title', async ({ page }) => {
  const pdfParser = new PDFParser(this, 1);

  const data = await fs.readFile('./input/articole-abstracte.xlsx');
  workbook = XLSX.read(data);

  const firstWorksheet = workbook.Sheets[workbook.SheetNames[0]];

  const rawData = XLSX.utils.sheet_to_json(firstWorksheet, { header: 1, defval: '-' });

  await page.goto(`https://google.com/search?q=${rawData[1][2]}`);

  const firstPageUrl = await page.locator('#rso >> a').first().getAttribute('href');

  if (!firstPageUrl) {
    return;
  }

  const downloadPromise = page.waitForEvent('download');

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
  }, [firstPageUrl, rawData[1][2]]);

  const download = await downloadPromise;
  await download.saveAs('./input/' + download.suggestedFilename());

  let finishedProcessing = false;

  pdfParser.loadPDF("./input/A Guide for Hiring Mature Employees at Ascentria.pdf");

  pdfParser.on("pdfParser_dataReady", pdfData => {
    const text = pdfParser.getRawTextContent();

    const lines = text.split('\n');
    const abstractLines: string[] = [];

    let abstractLineStart: number | null = null;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];

      if (abstractLineStart === null && line.includes('Abstract') && lines[i + 1].trim() === '') {
        abstractLineStart = i + 2;
      }

      if (abstractLineStart !== null && i < abstractLineStart) {
        continue;
      }

      if (abstractLineStart && line.trim() === '') {
        break
      }

      if (abstractLineStart !== null) {
        abstractLines.push(line.replace(/(\r\n|\n|\r)/gm, ''));
      }
    }

    XLSX.utils.sheet_add_aoa(firstWorksheet, [[abstractLines.join('')]], { origin: "E2" });
    finishedProcessing = true;
  });

  while (!finishedProcessing) {
    await new Promise(resolve => setTimeout(resolve, 1000));
  }
});

test.afterAll(async () => {
  XLSX.writeFile(workbook, "./input/articole+abstracte.xlsx", { compression: true });
});
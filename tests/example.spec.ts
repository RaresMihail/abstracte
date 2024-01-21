import { test } from '@playwright/test';
import fs from 'node:fs/promises';
import * as XLSX from 'xlsx';
import PDFParser from "pdf2json";

let workbook;

async function downloadPdf(page, firstPageUrl, fileName) {
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

function urlContainsKeyword(url: string, keyword: string): boolean {
  return url.includes(`https://${keyword}`) || url.includes(`https://www.${keyword}`);
} 

async function getSelectedPageURL(page): Promise<string> {
  let selectedPageUrl: string = '';
  const allUrlResults: string[] = [];
  const searchResults = await page.locator('#b_results >> li.b_algo');

  for (let i = 0; i < 3; i++) {
    const result = await searchResults.nth(i);
    const resultUrl: string = await result.locator('a').first().getAttribute('href');

    if (resultUrl === null) {
      continue;
    }

    if (resultUrl.endsWith('.pdf')) {
      // PDF search for abstract does not always work:
      // there are some files that fail to be fetched
      // and in some cases the files are protected by cloudflare
      continue;
    }

    allUrlResults.push(resultUrl);

    // preffered urls
    if (
      // urlContainsKeyword(resultUrl, 'semanticscholar') ||
      urlContainsKeyword(resultUrl, 'pubmed') ||
      urlContainsKeyword(resultUrl, 'papers.ssrn') ||
      urlContainsKeyword(resultUrl, 'emerald') ||
      urlContainsKeyword(resultUrl, 'psycnet') ||
      urlContainsKeyword(resultUrl, 'journals.sagepub') ||
      urlContainsKeyword(resultUrl, 'academic.oup') ||
      urlContainsKeyword(resultUrl, 'cambridge.org') ||
      urlContainsKeyword(resultUrl, 'springer') ||
      urlContainsKeyword(resultUrl, 'onlinelibrary.wiley') ||
      urlContainsKeyword(resultUrl, 'sciencedirect') ||
      urlContainsKeyword(resultUrl, 'taylorfrancis') ||
      urlContainsKeyword(resultUrl, 'research.manchester') ||
      urlContainsKeyword(resultUrl, 'ncbi.nlm') ||
      urlContainsKeyword(resultUrl, 'pure.ulster') ||
      urlContainsKeyword(resultUrl, 'researchportal.bath')) {
        selectedPageUrl = resultUrl;
        return selectedPageUrl;
    }
  }

  // this url is not preffered
  // for (const resultUrl of allUrlResults) {
  //   if (urlContainsKeyword(resultUrl, 'researchgate')) {
  //     selectedPageUrl = resultUrl;
  //     return selectedPageUrl;
  //   }
  // }

  return selectedPageUrl;
}


test('has title', async ({ page }) => {
  const data = await fs.readFile('./input/articole-abstracte.xlsx');
  workbook = XLSX.read(data);

  const firstWorksheet: XLSX.WorkSheet = workbook.Sheets[workbook.SheetNames[0]];

  const rawData = XLSX.utils.sheet_to_json(firstWorksheet, { header: 1, defval: '-' });

  for (let i = 1; i < rawData.length; i++) {  
  // for (let i = 31; i < 32; i++) {  

    const link: string = rawData[i][5];

    // There is scopus LINK
    if (link && link.includes('scopus')) {
      await page.goto(link);

      const abstractLocator = page.locator('#abstractSection');

      if (!await abstractLocator.isVisible()) {
        continue;
      }

      const abstract = await page.locator('#abstractSection').textContent();

      if(!abstract) {
        continue;
      }

      XLSX.utils.sheet_add_aoa(firstWorksheet, [[abstract.replace(/\n/g, '').trim()]], { origin: `E${i + 1}` });
      continue;
    }

    const fileName: string = rawData[i][2];
    const shortFileName = fileName.substring(0, 40);
    const author: string = rawData[i][1];

    const currentUrl = page.url();
    if (!currentUrl.includes('bing.com')) {
      await page.goto('https://www.bing.com/search?q=a');
    }

    const acceptCookiesButton = page.locator('button#bnp_btn_accept');

    if (await acceptCookiesButton.isVisible()) {
      await acceptCookiesButton.click();
    }

    await page.locator('#sb_form_q').fill(`${author} ${fileName}`);
    await page.locator('#sb_form_go').click();

    const selectedPageUrl = await getSelectedPageURL(page);

    if (selectedPageUrl === '') {
      continue;
    }

    if (selectedPageUrl.endsWith('.pdf')) {
      // PDF search for abstract does not always work:
      // there are some files that fail to be fetched
      // and in some cases the files are protected by cloudflare
      continue;

      await downloadPdf(page, selectedPageUrl, shortFileName);
      await searchAbstractInPdf(i, shortFileName, firstWorksheet);

    } else {
      await page.goto(selectedPageUrl);

      XLSX.utils.sheet_add_aoa(firstWorksheet, [[selectedPageUrl]], { origin: `F${i + 1}` });

      if (urlContainsKeyword(selectedPageUrl, 'semanticscholar')) {
        const expandButton = page.locator('[data-test-id="text-truncator-toggle"]');

        if (await expandButton.isVisible()) {
          await expandButton.click();
        }

        const abstractLocator = page.locator('[data-test-id="no-highlight-abstract-text"]');

        if (!await abstractLocator.isVisible()) {
          continue;
        }

        const abstract = await page.locator('[data-test-id="no-highlight-abstract-text"]').textContent();

        if(!abstract) {
          continue;
        }

        XLSX.utils.sheet_add_aoa(firstWorksheet, [[abstract.trim()]], { origin: `E${i + 1}` });
      }

      if (urlContainsKeyword(selectedPageUrl, 'pubmed')) {
        const abstractLocator = page.locator('#eng-abstract');

        if (!await abstractLocator.isVisible()) {
          continue;
        }

        const abstract = await page.locator('#eng-abstract').textContent();

        if(!abstract) {
          continue;
        }

        XLSX.utils.sheet_add_aoa(firstWorksheet, [[abstract.replace(/\n/g, '').trim()]], { origin: `E${i + 1}` });
      }

      if (urlContainsKeyword(selectedPageUrl, 'researchgate')) {
        const abstractLocator = page.locator('[itemprop="description"]');

        if (!await abstractLocator.isVisible()) {
          continue;
        }

        const abstract = await page.locator('[itemprop="description"]').textContent();

        if(!abstract) {
          continue;
        }

        XLSX.utils.sheet_add_aoa(firstWorksheet, [[abstract.replace(/\n/g, '').trim()]], { origin: `E${i + 1}` });
      }

      if (urlContainsKeyword(selectedPageUrl, 'papers.ssrn')) {
        const abstractLocator = page.locator('div.abstract-text > p').first();

        if (!await abstractLocator.isVisible()) {
          continue;
        }

        const abstract = await page.locator('div.abstract-text > p').first().textContent();

        if(!abstract) {
          continue;
        }

        XLSX.utils.sheet_add_aoa(firstWorksheet, [[abstract.replace(/\n/g, '').trim()]], { origin: `E${i + 1}` });
      }

      if (urlContainsKeyword(selectedPageUrl, 'emerald')) {
        const abstractLocator = page.locator('#abstract');

        if (!await abstractLocator.isVisible()) {
          continue;
        }

        const abstract = await page.locator('#abstract').textContent();

        if(!abstract) {
          continue;
        }

        XLSX.utils.sheet_add_aoa(firstWorksheet, [[abstract.replace(/\n/g, '').trim()]], { origin: `E${i + 1}` });
      }

      if (urlContainsKeyword(selectedPageUrl, 'psycnet')) {
        const abstractLocator = page.locator('abstract');

        if (!await abstractLocator.isVisible()) {
          continue;
        }

        const abstract = await page.locator('abstract').textContent();

        if(!abstract) {
          continue;
        }

        XLSX.utils.sheet_add_aoa(firstWorksheet, [[abstract.replace(/\n/g, '').trim()]], { origin: `E${i + 1}` });
      }

      if (urlContainsKeyword(selectedPageUrl, 'journals.sagepub')) {
        const abstractLocator = page.locator('#abstract');

        if (!await abstractLocator.isVisible()) {
          continue;
        }

        const abstract = await page.locator('#abstract').textContent();

        if(!abstract) {
          continue;
        }

        XLSX.utils.sheet_add_aoa(firstWorksheet, [[abstract.trim()]], { origin: `E${i + 1}` });
      }

      if (urlContainsKeyword(selectedPageUrl, 'academic.oup')) {
        const abstractLocator = page.locator('section.abstract');

        if (!await abstractLocator.isVisible()) {
          continue;
        }

        const abstract = await page.locator('section.abstract').textContent();

        if(!abstract) {
          continue;
        }

        XLSX.utils.sheet_add_aoa(firstWorksheet, [[abstract.trim()]], { origin: `E${i + 1}` });
      }

      if (urlContainsKeyword(selectedPageUrl, 'cambridge.org')) {
        const abstractLocator = page.locator('div.abstract-content');

        if (!await abstractLocator.isVisible()) {
          continue;
        }

        const abstract = await page.locator('div.abstract-content').textContent();

        if(!abstract) {
          continue;
        }

        XLSX.utils.sheet_add_aoa(firstWorksheet, [[abstract.trim()]], { origin: `E${i + 1}` });
      }

      if (urlContainsKeyword(selectedPageUrl, 'springer')) {
        const abstractLocator = page.locator('#Abs1-content');

        if (!await abstractLocator.isVisible()) {
          continue;
        }

        const abstract = await page.locator('#Abs1-content').textContent();

        if(!abstract) {
          continue;
        }

        XLSX.utils.sheet_add_aoa(firstWorksheet, [[abstract.trim()]], { origin: `E${i + 1}` });
      }

      if (urlContainsKeyword(selectedPageUrl, 'onlinelibrary.wiley')) {
        const abstractLocator = page.locator('#section-1-en');

        if (!await abstractLocator.isVisible()) {
          continue;
        }

        const abstract = await page.locator('#section-1-en').textContent();

        if(!abstract) {
          continue;
        }

        XLSX.utils.sheet_add_aoa(firstWorksheet, [[abstract.replace(/\n/g, '').trim()]], { origin: `E${i + 1}` });
      }

      if (urlContainsKeyword(selectedPageUrl, 'sciencedirect')) {
        const abstractLocator = page.locator('#abst0010');

        if (!await abstractLocator.isVisible()) {
          continue;
        }

        const abstract = await page.locator('#abst0010').textContent();

        if(!abstract) {
          continue;
        }

        XLSX.utils.sheet_add_aoa(firstWorksheet, [[abstract.trim()]], { origin: `E${i + 1}` });
      }

      if (urlContainsKeyword(selectedPageUrl, 'taylorfrancis')) {
        const abstractLocator = page.locator('#collapseContent');

        if (!await abstractLocator.isVisible()) {
          continue;
        }

        const abstract = await page.locator('#collapseContent').textContent();

        if(!abstract) {
          continue;
        }

        XLSX.utils.sheet_add_aoa(firstWorksheet, [[abstract.trim()]], { origin: `E${i + 1}` });
      }

      if (urlContainsKeyword(selectedPageUrl, 'research.manchester')) {
        const abstractLocator = page.locator('div.rendering_abstractportal');

        if (!await abstractLocator.isVisible()) {
          continue;
        }

        const abstract = await page.locator('div.rendering_abstractportal').textContent();

        if(!abstract) {
          continue;
        }

        XLSX.utils.sheet_add_aoa(firstWorksheet, [[abstract.trim()]], { origin: `E${i + 1}` });
      }

      if (urlContainsKeyword(selectedPageUrl, 'ncbi.nlm')) {
        const abstractLocator = page.locator('[id="abstract-a.f.b.p"]');

        if (!await abstractLocator.isVisible()) {
          continue;
        }

        const abstract = await page.locator('[id="abstract-a.f.b.p"]').textContent();

        if(!abstract) {
          continue;
        }

        XLSX.utils.sheet_add_aoa(firstWorksheet, [[abstract.replace(/\n/g, '').trim()]], { origin: `E${i + 1}` });
      }

      if (urlContainsKeyword(selectedPageUrl, 'pure.ulster')) {
        const abstractLocator = page.locator('div.rendering_abstractportal');

        if (!await abstractLocator.isVisible()) {
          continue;
        }

        const abstract = await page.locator('div.rendering_abstractportal').textContent();

        if(!abstract) {
          continue;
        }

        XLSX.utils.sheet_add_aoa(firstWorksheet, [[abstract.trim()]], { origin: `E${i + 1}` });
      }

      if (urlContainsKeyword(selectedPageUrl, 'researchportal.bath')) {
        const abstractLocator = page.locator('div.rendering_abstractportal');

        if (!await abstractLocator.isVisible()) {
          continue;
        }

        const abstract = await page.locator('div.rendering_abstractportal').textContent();

        if(!abstract) {
          continue;
        }

        XLSX.utils.sheet_add_aoa(firstWorksheet, [[abstract.trim()]], { origin: `E${i + 1}` });
      }
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
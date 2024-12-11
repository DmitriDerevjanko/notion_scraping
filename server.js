import express from 'express';
import bodyParser from 'body-parser';
import fetch from 'node-fetch';
import * as cheerio from 'cheerio';
import path from 'path';
import { fileURLToPath } from 'url';
import exceljs from 'exceljs';
import emtakCodes from './emtakCodes.js';
import countyMap from './countyMap.js';
import PDFDocument from 'pdfkit';
import fs from 'fs';
import puppeteer from 'puppeteer';



const app = express();
const port = 3000;

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

app.use(bodyParser.json());
app.use(express.static(__dirname));
app.use(express.static('public'));
app.use('/files', express.static(path.join(__dirname, 'files')));

app.get('/download/:filename', (req, res) => {
    const filename = req.params.filename;
    const filepath = path.join(__dirname, filename);
    res.download(filepath); 
  });
app.post('/upload', async (req, res) => {
    try {
        const data = req.body.data;
        const results = await processData(data);
        res.json({
            message: 'Data processed successfully',
            matchingFileData: results.matchingFileData,
            nonMatchingFileData: results.nonMatchingFileData
        });
    } catch (error) {
        console.error('Server error:', error);
        res.status(500).json({ message: 'Server error', error: error.message });
    }
});


function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}
function normalizeCompanyName(name) {
    let processedName = name.trim();
    const prefixRegex = /^(AS |OÜ |SAS |MTÜ )[\s\.-]*/i;
    const patterns = [
        { regex: /\baktsiaselts\b/i, replacement: "AS" },
        { regex: /\bosaühing\b/i, replacement: "OÜ" },
        { regex: /\bsihtasutus\b/i, replacement: "SAS" },
        { regex: /\bmittetulundusühing\b/i, replacement: "MTÜ" }
    ];

    let suffix = "";
    const prefixMatch = processedName.match(prefixRegex);
    if (prefixMatch) {
        processedName = processedName.replace(prefixRegex, ''); 
        suffix = prefixMatch[0].trim();
    }
    patterns.forEach(pattern => {
        if (pattern.regex.test(processedName)) {
            processedName = processedName.replace(pattern.regex, '').trim();
            suffix = pattern.replacement;
        }
    });

    if (suffix) {
        processedName = `${processedName}, ${suffix}`;
    }

    processedName = processedName.replace(/,/g, '');
    return processedName;
}

function createPDFWithLinks(data) {
    const doc = new PDFDocument();
    const filePath = path.join(__dirname, 'clients_turnover.pdf');

    doc.pipe(fs.createWriteStream(filePath));

    doc.fontSize(12).text('Clients Turnover', { align: 'center' });
    doc.moveDown();

    data.forEach(company => {
        if (company.pdfUrl) {
            const fullPdfUrl = `https://ariregister.rik.ee${company.pdfUrl}`;
            doc.fontSize(10).text(`${company.name}: ${fullPdfUrl}`, {
                link: fullPdfUrl,
                underline: true
            }).moveDown(0.5);
        }
    });

    doc.end();
}

async function processData(data) {
    console.log('Starting to process data. Number of rows:', data.length);
    const matchingEMTAK = [];
    const nonMatchingEMTAK = [];

    const matchingWorkbook = new exceljs.Workbook();
    const nonMatchingWorkbook = new exceljs.Workbook();
    const matchingWorksheet = matchingWorkbook.addWorksheet('Matching EMTAK');
    const nonMatchingWorksheet = nonMatchingWorkbook.addWorksheet('Non-Matching EMTAK');
    matchingWorksheet.addRow([
        'Property', 'Kliendihaldur', 'People', 'Registry Code',
        'Location', 'EMTAK', 'Main field of activity', 'Employees', 'Turnover', 'Link'
    ]);

    nonMatchingWorksheet.addRow([
        'Property', 'Kliendihaldur', 'People', 'Registry Code',
        'Location', 'EMTAK', 'Main field of activity', 'Employees', 'Turnover', 'Link'
    ]);

    if (data[0].length < 9) {
        const headersToAdd = [
            'Property', 'Kliendihaldur', 'People', 'Registry Code',
            'Location', 'EMTAK', 'Main field of activity', 'Employees', 'Turnover'
        ];
        data[0] = headersToAdd.concat(data[0].slice(2));
    }

    for (let i = 1; i < data.length; i++) {
        console.log(`Processing row ${i}:`, data[i]);
        if (data[i].length === 0 || !data[i][0]) {
            console.log(`Skipping empty or invalid row ${i}:`, data[i]);
            continue;
        }

        let newRow = data[i].slice();
        let registryCode = String(newRow[1]);
        if (registryCode.startsWith('EE')) {
            registryCode = registryCode.substring(2); 
        }        
        newRow[1] = "";
        newRow[3] = registryCode;
        const url = `https://ariregister.rik.ee/est/company/${registryCode}`;
        newRow.push(url);
        console.log(`Fetching URL for row ${i}: ${url}`);
        

        try {
            await sleep(3000);
            console.log('Normalizing company name:', newRow[0]);
            newRow[0] = normalizeCompanyName(newRow[0]);
            const response = await fetch(url);
            const body = await response.text();
            const $ = cheerio.load(body);

            const emtakCodeElement = $('td.text-nowrap.px-1').first();
            newRow[5] = emtakCodeElement.length ? emtakCodeElement.text().trim() : 'Not found';

            const mainFieldOfActivityElement = emtakCodeElement.closest('tr').find('a[title]').first();
            newRow[6] = mainFieldOfActivityElement.length ? mainFieldOfActivityElement.attr('title').trim() : 'Not found';

            const locationElement = $('div.col.font-weight-bold').filter(function () {
                return $(this).prev().text().trim() === 'Aadress';
            }).first();

            if (locationElement.length > 0) {
                let fullAddress = locationElement.text().trim().split(',')[0];
                let countyName = fullAddress.replace(' maakond', '').trim();
                let shortenedCounty = countyMap[countyName] || countyName + 'maa';
                newRow[4] = shortenedCounty;
            } else {
                newRow[4] = 'Not found';
            }

            const browser = await puppeteer.launch();
            const page = await browser.newPage();
            await page.goto(url, { waitUntil: 'networkidle0' });
            console.log(`Page opened for row ${i}: ${url}`);
            
            try {
                await page.waitForSelector('div.col-md-6.text-muted', { timeout: 5000 }); 
                console.log(`Element found for row ${i}`);
            } catch (error) {
                console.error(`Element not found for row ${i}:`, error);
                newRow[7] = 'Element not found';
                await browser.close();
                continue; 
            }
            
            const employeesCount = await page.evaluate(() => {
                const employeesLabel = Array.from(document.querySelectorAll('div.col-md-6.text-muted'))
                                            .find(el => el.textContent.trim() === 'Töötajate arv');
                if (employeesLabel) {
                    const employeesValueElement = employeesLabel.nextElementSibling;
                    if (employeesValueElement) {
                        const employeesValue = employeesValueElement.textContent.trim();
                        const parsedValue = parseInt(employeesValue, 10);
                        if (!isNaN(parsedValue)) {
                            return parsedValue;
                        } else {
                            return `Parsed value is NaN for text: ${employeesValue}`;
                        }
                    } else {
                        return 'Next sibling element (employee count) not found';
                    }
                } else {
                    return 'Employee label not found';
                }
            });
            console.log(`Employees count for row ${i}:`, employeesCount);
            await browser.close();

            if (!isNaN(employeesCount) && employeesCount > 0 && employeesCount <= 5000) {
                newRow[7] = employeesCount;
            } else {
                newRow[7] = 'Invalid number';
            }

            data[i] = newRow;
            if (emtakCodes.includes(newRow[5])) {
                matchingEMTAK.push(newRow);
            } else {
                nonMatchingEMTAK.push(newRow);
            }
        } catch (error) {
            console.error(`Error processing row ${i}:`, error);
            newRow.fill('Error', 4, newRow.length);
            data[i] = newRow;
        }
        console.log(`Finished processing row ${i}`);
    }

    console.log('Finished processing all data');
    matchingWorksheet.addRows(matchingEMTAK);
    nonMatchingWorksheet.addRows(nonMatchingEMTAK);

    const matchingFileData = await matchingWorkbook.xlsx.writeBuffer();
    const nonMatchingFileData = await nonMatchingWorkbook.xlsx.writeBuffer();

    const matchingFileBase64 = matchingFileData.toString('base64');
    const nonMatchingFileBase64 = nonMatchingFileData.toString('base64');
    
    return { matchingFileData: matchingFileBase64, nonMatchingFileData: nonMatchingFileBase64 };
    
}

app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
});
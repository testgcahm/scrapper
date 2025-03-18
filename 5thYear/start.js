const puppeteer = require('puppeteer');
const fs = require('fs');
const csv = require('csv-parser');

const URL = 'https://results.uhs.edu.pk/Results/main';
const INPUT_CSV = 'names.csv';
const OUTPUT_FILE = 'results.csv';
const ROLLNUMBERS_FILE = 'rollnumbers.csv';

async function readNamesFromCSV(file) {
    return new Promise((resolve, reject) => {
        const names = [];
        fs.createReadStream(file)
            .pipe(csv({ headers: false }))
            .on('data', (row) => names.push(row[0].trim()))
            .on('end', () => resolve(names))
            .on('error', reject);
    });
}

async function readProcessedResults(file) {
    if (!fs.existsSync(file)) return new Set();
    return new Set(fs.readFileSync(file, 'utf8').split('\n').map(line => line.split(',')[0]));
}

async function readProcessedNamesResults(file) {
    if (!fs.existsSync(file)) return new Set();
    return new Set(fs.readFileSync(file, 'utf8').split('\n').map(line => line.split(',')[1]));
}


async function readRollNumbersFromCSV(file) {
    return new Promise((resolve, reject) => {
        const rollNumbers = [];
        fs.createReadStream(file)
            .pipe(csv({ headers: false }))
            .on('data', (row) => rollNumbers.push(row[0].trim()))
            .on('end', () => resolve(rollNumbers))
            .on('error', reject);
    });
}

async function appendToResultsFile(rollNo, name, result) {
    fs.appendFileSync(OUTPUT_FILE, `${rollNo},${name},${result}\n`);
}

async function scrapeResults() {
    const browser = await puppeteer.launch({ headless: false });
    const page = await browser.newPage();
    await page.goto(URL, { waitUntil: 'networkidle2' });

    // Initial navigation steps
    await page.click('#MainContent_Image1');
    await page.waitForNavigation({ waitUntil: 'networkidle2' });
    await page.waitForSelector("a[href=\"javascript:__doPostBack('ctl00$MainContent$gvCrs','Select$24')\"]", { visible: true });
    await page.evaluate(() => __doPostBack('ctl00$MainContent$gvCrs', 'Select$24'));
    await page.waitForNavigation({ waitUntil: 'networkidle2' });

    const names = await readNamesFromCSV(INPUT_CSV);
    const processed = await readProcessedResults(OUTPUT_FILE);
    let rollNumbers = await readRollNumbersFromCSV(ROLLNUMBERS_FILE);
    const processedNames = await readProcessedNamesResults(OUTPUT_FILE);
    const processedRollNumbers = new Set();

    for (const name of names) {
        let found = false;

        if (processedNames.has(name)) {
            console.log(`Skipping processed: ${name}`);
            continue;
        }

        for (const rollNoStr of [...rollNumbers]) {
            if (processed.has(rollNoStr) || processedRollNumbers.has(rollNoStr)) {
                continue;
            }

            // Clear form fields
            await page.evaluate(() => {
                document.getElementById('MainContent_rno').value = '';
                document.getElementById('MainContent_stdnm').value = '';
            });

            // Fill form
            await page.type('#MainContent_rno', rollNoStr);
            await page.type('#MainContent_stdnm', name);
            
            // Handle navigation
            const navigationPromise = page.waitForNavigation({ timeout: 5000 }).catch(() => null);
            await page.click('#MainContent_Button1');
            await navigationPromise;

            // Check for result page
            if (page.url().includes('/Results/dtl')) {
                const resultElement = await page.$('#MainContent_GTot');
                const MedicineElement = await page.$('#MainContent_MedicineAlliedMO');
                const MedicineTotalElement = await page.$('#MainContent_MedicineAlliedTM');
                const MedicineResultElement = await page.$('#MainContent_MedicineAlliedResult');
                const SurgeryElement = await page.$('#MainContent_SurgeryAlliedMO');
                const SurgeryTotalElement = await page.$('#MainContent_SurgeryAlliedTM');
                const SurgeryResultElement = await page.$('#MainContent_SurgeryAlliedResult');
                const ObsGynElement = await page.$('#MainContent_ObsGynMO');
                const ObsGynTotalElement = await page.$('#MainContent_ObsGynTM');
                const ObsGynResultElement = await page.$('#MainContent_ObsGynResult');
                const PaedsElement = await page.$('#MainContent_PaedsMO');
                const PaedsTotalElement = await page.$('#MainContent_PaedsTM');
                const PaedsResultElement = await page.$('#MainContent_PaedsResult');
                if (resultElement && MedicineElement && MedicineTotalElement && MedicineResultElement && SurgeryElement && SurgeryTotalElement && SurgeryResultElement && ObsGynElement && ObsGynTotalElement && ObsGynResultElement && PaedsElement && PaedsTotalElement && PaedsResultElement) {
                    const result = await page.evaluate(el => el.textContent.trim(), resultElement);
                    const Medicine = await page.evaluate(el => el.textContent.trim(), MedicineElement);
                    const MedicineTotal = await page.evaluate(el => el.textContent.trim(), MedicineTotalElement);
                    const MedicineResult = await page.evaluate(el => el.textContent.trim(), MedicineResultElement);
                    const Surgery = await page.evaluate(el => el.textContent.trim(), SurgeryElement);
                    const SurgeryTotal = await page.evaluate(el => el.textContent.trim(), SurgeryTotalElement);
                    const SurgeryResult = await page.evaluate(el => el.textContent.trim(), SurgeryResultElement);
                    const ObsGyn = await page.evaluate(el => el.textContent.trim(), ObsGynElement);
                    const ObsGynTotal = await page.evaluate(el => el.textContent.trim(), ObsGynTotalElement);
                    const ObsGynResult = await page.evaluate(el => el.textContent.trim(), ObsGynResultElement);
                    const Paeds = await page.evaluate(el => el.textContent.trim(), PaedsElement);
                    const PaedsTotal = await page.evaluate(el => el.textContent.trim(), PaedsTotalElement);
                    const PaedsResult = await page.evaluate(el => el.textContent.trim(), PaedsResultElement);
                    console.log(`FOUND: ${name} - ${rollNoStr}: ${result || 'Failed'} - ${Medicine} - ${MedicineTotal} - ${MedicineResult} - ${Surgery} - ${SurgeryTotal} - ${SurgeryResult} - ${ObsGyn} - ${ObsGynTotal} - ${ObsGynResult} - ${Paeds} - ${PaedsTotal} - ${PaedsResult}`);
                    await appendToResultsFile(rollNoStr, name, `${result || 'Failed'},${Medicine},${MedicineTotal},${MedicineResult},${Surgery},${SurgeryTotal},${SurgeryResult},${ObsGyn},${ObsGynTotal},${ObsGynResult},${Paeds},${PaedsTotal},${PaedsResult}`);
                    found = true;
                    processed.add(rollNoStr);
                    processedRollNumbers.add(rollNoStr);
                    //  remove rollnumber from rollnumbers.csv
                    rollNumbers = rollNumbers.filter(roll => roll !== rollNoStr);
                    fs.writeFileSync(ROLLNUMBERS_FILE, rollNumbers.join('\n'));
                    await page.goBack({ waitUntil: 'networkidle2' });
                    break;
                }
            } else {
                // Check for error message
                const error = await page.$eval('#MainContent_Label2', el => el.textContent.trim());
                if (error === 'No Record Found!') {
                    console.log(`No match: ${name} - ${rollNoStr}`);
                    
                }
            }
        }
        if (!found) {
            console.log(`No valid roll number found for: ${name}`);
        }
    }

    await browser.close();
}

scrapeResults().catch(console.error);
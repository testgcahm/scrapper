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
    await page.waitForSelector("a[href=\"javascript:__doPostBack('ctl00$MainContent$gvCrs','Select$0')\"]", { visible: true });
    await page.evaluate(() => __doPostBack('ctl00$MainContent$gvCrs', 'Select$0'));
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
                const cmedElement = await page.$('#MainContent_CommunityMedicineMO');
                const cmedResultElement = await page.$('#MainContent_CommunityMedicineResult');
                const cmedTotalElement = await page.$('#MainContent_CommunityMedicineTM');
                const entElement = await page.$('#MainContent_OtorhinolaryngologyMO');
                const entResultElement = await page.$('#MainContent_OtorhinolaryngologyResult');
                const entTotalElement = await page.$('#MainContent_OtorhinolaryngologyTM');
                const eyeElement = await page.$('#MainContent_OphthalmologyMO');
                const eyeResultElement = await page.$('#MainContent_OphthalmologyResult');
                const eyeTotalElement = await page.$('#MainContent_OphthalmologyTM');
                const pathoElement = await page.$('#MainContent_SpecialPathologyMO');
                const pathoResultElement = await page.$('#MainContent_SpecialPathologyResult');
                const pathoTotalElement = await page.$('#MainContent_SpecialPathologyTM');
                if (resultElement && cmedElement && entElement && eyeElement && pathoElement && cmedResultElement && entResultElement && eyeResultElement && pathoResultElement) {
                    const result = await page.evaluate(el => el.textContent.trim(), resultElement);
                    const cmed = await page.evaluate(el => el.textContent.trim(), cmedElement);
                    const ent = await page.evaluate(el => el.textContent.trim(), entElement);
                    const eye = await page.evaluate(el => el.textContent.trim(), eyeElement);
                    const patho = await page.evaluate(el => el.textContent.trim(), pathoElement);
                    const cmedResult = await page.evaluate(el => el.textContent.trim(), cmedResultElement);
                    const entResult = await page.evaluate(el => el.textContent.trim(), entResultElement);
                    const eyeResult = await page.evaluate(el => el.textContent.trim(), eyeResultElement);
                    const pathoResult = await page.evaluate(el => el.textContent.trim(), pathoResultElement);
                    const cmedTotal = await page.evaluate(el => el.textContent.trim(), cmedTotalElement);
                    const entTotal = await page.evaluate(el => el.textContent.trim(), entTotalElement);
                    const eyeTotal = await page.evaluate(el => el.textContent.trim(), eyeTotalElement);
                    const pathoTotal = await page.evaluate(el => el.textContent.trim(), pathoTotalElement);
                    console.log(`FOUND: ${name} - ${rollNoStr}: ${result || 'Failed'} - ${cmed} - ${cmedResult} - ${cmedTotal} - ${ent} - ${entResult} - ${entTotal} - ${eye} - ${eyeResult} - ${eyeTotal} - ${patho} - ${pathoResult} - ${pathoTotal}`);
                    await appendToResultsFile(rollNoStr, name, `${result || 'Failed'},${cmed},${cmedResult},${cmedTotal},${ent},${entResult},${entTotal},${eye},${eyeResult},${eyeTotal},${patho},${pathoResult},${pathoTotal}`);
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
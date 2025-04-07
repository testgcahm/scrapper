const puppeteer = require('puppeteer');
const fs = require('fs');
const csv = require('csv-parser');

const URL = "https://results.uhs.edu.pk/Results/main";
const SELECT_VALUE = process.argv[2]; // Pass Select$X value as an argument
const YEAR = process.argv[3]; // Pass MBBS year as an argument
const INPUT_CSV = `${YEAR}/names.csv`;
const OUTPUT_FILE = `${YEAR}/results_year_${YEAR}.csv`;
const ROLLNUMBERS_FILE = `${YEAR}/rollnumbers.csv`;

async function readProcessedRollNumbers(file) {
    if (!fs.existsSync(file)) return new Set();
    return new Set(fs.readFileSync(file, 'utf8').split('\n').map(line => line.split(',')[0]));
}

async function readProcessedNames(file) {
    if (!fs.existsSync(file)) return new Set();
    return new Set(fs.readFileSync(file, 'utf8').split('\n').map(line => line.split(',')[1]));
}

async function readCSV(file) {
    return new Promise((resolve, reject) => {
        const data = [];
        fs.createReadStream(file)
            .pipe(csv({ headers: false }))
            .on('data', (row) => data.push(row[0].trim()))
            .on('end', () => resolve(data))
            .on('error', reject);
    });
}

async function appendToResultsFile(rollNo, name, result, resultData) {
    const data = `${rollNo},${name},${result},${resultData.join(',')}\n`;
    fs.writeFileSync(OUTPUT_FILE, data, { flag: 'a' }); // Ensures appending
}

async function scrapeResults() {
    const browser = await puppeteer.launch({ headless: false });
    const page = await browser.newPage();
    await page.goto(URL, { waitUntil: 'networkidle2' });

    // Navigate to results page
    await page.click('#MainContent_Image1');
    await page.waitForNavigation({ waitUntil: 'networkidle2' });

    // Select the correct year
    await page.waitForSelector(`a[href="javascript:__doPostBack('ctl00$MainContent$gvCrs','Select$${SELECT_VALUE}')"]`);
    await page.evaluate((selectValue) => {
        __doPostBack('ctl00$MainContent$gvCrs', `Select$${selectValue}`);
    }, SELECT_VALUE);
    await page.waitForNavigation({ waitUntil: 'networkidle2' });

    const names = await readCSV(INPUT_CSV);
    let rollNumbers = await readCSV(ROLLNUMBERS_FILE);
    const processedRollNumbersFromResult = await readProcessedRollNumbers(OUTPUT_FILE);
    const processedNames = await readProcessedNames(OUTPUT_FILE);
    const processedRollNumbers = new Set();

    for (const name of names) {
        let found = false;

        if (processedNames.has(name)) {
            console.log(`Skipping processed: ${name}`);
            continue;
        }

        for (const rollNoStr of [...rollNumbers]) {
            if (processedRollNumbersFromResult.has(rollNoStr) || processedRollNumbers.has(rollNoStr)) {
                continue;
            }
            await page.evaluate(() => {
                document.getElementById('MainContent_rno').value = '';
                document.getElementById('MainContent_stdnm').value = '';
            });

            await page.type('#MainContent_rno', rollNoStr);
            await page.type('#MainContent_stdnm', name);

            const navigationPromise = page.waitForNavigation({ timeout: 6000 }).catch(() => null);
            await page.click('#MainContent_Button1');
            await navigationPromise;

            if (page.url().includes('/Results/dtl')) {
                const resultElement = await page.$('#MainContent_GTot');
                const totalResult = await page.evaluate(el => el.textContent.trim(), resultElement);
                
                // Extracting result table
                const results = await page.evaluate((year) => {
                    const subjects = [];
                    const rows = Array.from(document.querySelectorAll('table.table.table-responsive.table-striped tbody tr'));

                    rows.slice(8).forEach((row) => {
                        const columns = row.querySelectorAll('td');
                        if (columns.length >= 7) {
                            const subjectName = columns[1]?.textContent.trim();
                            const obtainedMarks = columns[parseInt(year) === 5 ? 5 : 4]?.textContent.trim();
                            const totalMarks = columns[parseInt(year) === 5 ? 6 : 5]?.textContent.trim();
                            const result = columns[parseInt(year) === 5 ? 7 : 6]?.textContent.trim();
                            subjects.push(`${subjectName}:${obtainedMarks}/${totalMarks}(${result})`);
                        }
                    });

                    return subjects;
                }, YEAR);            

                if (results.length > 0) {
                    console.log(`FOUND: ${name} - ${rollNoStr}: ${totalResult || 'Failed'} - ${results.join(' | ')}`);
                    await appendToResultsFile(rollNoStr, name, totalResult || 'Failed', results);
                    found = true;
                    processedRollNumbersFromResult.add(rollNoStr);
                    processedRollNumbers.add(rollNoStr);
                    //  remove rollnumber from rollnumbers.csv
                    rollNumbers = rollNumbers.filter(roll => roll !== rollNoStr);
                    fs.writeFileSync(ROLLNUMBERS_FILE, rollNumbers.join('\n'));
                }
                await page.goBack({ waitUntil: 'networkidle2' });
                break;
            } else {
                // Check for error message
                const error = await page.$eval('#MainContent_Label2', el => el.textContent.trim()).catch(() => '');
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

// Run script with provided inputs
scrapeResults().catch(console.error);

const fs = require('fs');
const csv = require('csv-parser');

const RESULTS_CSV = 'results.csv';
const OUTPUT_FILE = 'rollnumbers.csv';
const START_ROLL = 36378;
const END_ROLL = 36378;
const CUSTOM_ROLLS = [];

const USE_CUSTOM = process.argv.includes('-r');

async function getExistingRolls() {
    const rolls = new Set();

    if (!fs.existsSync(RESULTS_CSV)) {
        console.log(`${RESULTS_CSV} not found. Generating all roll numbers.`);
        return null; // Indicate that the file does not exist
    }

    return new Promise((resolve, reject) => {
        fs.createReadStream(RESULTS_CSV)
            .pipe(csv({ headers: false }))
            .on('data', (row) => {
                if (row[0]) {
                    rolls.add(row[0].trim());
                }
            })
            .on('end', () => resolve(rolls))
            .on('error', reject);
    });
}

async function findMissingRolls() {
    try {
        const existingRolls = await getExistingRolls();
        const missing = [];

        if (!USE_CUSTOM) {
            for (let roll = START_ROLL; roll <= END_ROLL; roll++) {
                const rollStr = roll.toString().padStart(6, '0');
                if (!existingRolls || !existingRolls.has(rollStr)) {
                    missing.push(rollStr);
                }
            }
        }

        if (CUSTOM_ROLLS) {
            for (const roll of CUSTOM_ROLLS) {
                const rollStr = roll.toString().padStart(6, '0');
                if (!existingRolls || !existingRolls.has(rollStr)) {
                    missing.push(rollStr);
                }
            }
        }

        fs.writeFileSync(OUTPUT_FILE, missing.join('\n'));
        console.log(`Saved ${missing.length} roll numbers to ${OUTPUT_FILE}`);
    } catch (error) {
        console.error('Error:', error);
    }
}

findMissingRolls();

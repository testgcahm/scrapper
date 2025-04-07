const fs = require('fs');
const xlsx = require('xlsx-js-style');
const csv = require('csv-parser');

const YEAR = process.argv[2]; // Pass MBBS year as an argument
const INPUT_CSV = `${YEAR}/results_year_${YEAR}.csv`;
const OUTPUT_XLSX = `${YEAR}/results_year_${YEAR}.xlsx`;

const subjectMapping = {
    "General Pathology & Microbiology": "Patho",
    "Pharmacology & Therapeutics": "Pharma",
    "Forensic Medicine & Taxicology": "Forensic",
    "Forensic Medicine & Toxicology": "Forensic",
    "Behavioural Sciences": "B.Sciences",
    "Islamic Studies / Ethics & Pak Studies": "IslPak",
    "Community Medicine": "Cmed",
    "Special Pathology": "Patho",
    "Otorhinolaryngology": "Ent",
    "Ophthalmology": "Eye",
    "Medicine & Allied": "Medicine",
    "Surgery & Allied": "Surgery",
    "Obstetrics & Gynaecology": "GynObs",
    "Paediatrics": "Paeds",
};

function calculateEightyFivePercent(number) {
    return number * 0.85;
}

async function convertCSVtoExcel() {
    const data = [];
    const distinctionCounts = {};
    const supplyCounts = {};
    let subjects = [];

    fs.createReadStream(INPUT_CSV)
        .pipe(csv({ headers: false }))
        .on('data', (row) => {
            const entries = Object.values(row);
            const rollNumber = entries[0];
            const studentName = entries[1];
            const result = entries[2];
            const subjectData = entries.slice(3);

            let rawResultMatch = result.match(/(\d+)\/(\d+)/);
            let rawResult = rawResultMatch ? parseInt(rawResultMatch[1]) : 0;

            let studentRecord = {
                "Roll No.": rollNumber,
                "Names": studentName,
                "Result": result,
                "RawResult": rawResult
            };

            let distinctions = [];
            let supplySubjects = [];

            subjectData.forEach((subjectString) => {
                let subjectMatch = subjectString.match(/(.+?):(\d+)\/(\d+)\((Pass|Fail)\)/);
                if (subjectMatch) {
                    let subject = subjectMatch[1].trim();
                    let obtainedMarks = parseInt(subjectMatch[2]);
                    let totalMarks = parseInt(subjectMatch[3]);
                    let subjectResult = subjectMatch[4];
                    let mappedSubject = subjectMapping[subject] || subject;

                    if (!subjects.includes(mappedSubject)) {
                        subjects.push(mappedSubject);
                        distinctionCounts[mappedSubject] = 0;
                        supplyCounts[mappedSubject] = 0;
                    }

                    studentRecord[mappedSubject] = obtainedMarks;
                    studentRecord[`${mappedSubject}_Total`] = totalMarks;
                    studentRecord[`${mappedSubject}_Result`] = subjectResult;

                    if (subject !== "Islamic Studies / Ethics & Pak Studies" && obtainedMarks >= calculateEightyFivePercent(totalMarks) && result !== 'Failed') {
                        distinctions.push(mappedSubject.toUpperCase());
                        distinctionCounts[mappedSubject]++;
                    }
                    if (subjectResult === 'Fail') {
                        supplySubjects.push(mappedSubject.toUpperCase());
                        supplyCounts[mappedSubject]++;
                    }
                }
            });

            studentRecord['Remarks'] = supplySubjects.length > 0 ? supplySubjects.join(', ') : distinctions.length > 0 ? distinctions.join(', ') : '-';
            data.push(studentRecord);
        })
        .on('end', () => {
            data.sort((a, b) => {
                if (a.Result === 'Failed') return 1;
                if (b.Result === 'Failed') return -1;
                return b.RawResult - a.RawResult;
            });

            const finalData = data.map(({ RawResult, ...rest }, index) => {
                const filteredData = {};
                Object.keys(rest).forEach((key) => {
                    if (!key.includes("_Total") && !key.includes("_Result")) {
                        filteredData[key] = rest[key];
                    }
                });

                return {
                    "Sr#": index + 1,
                    ...filteredData
                };
            });

            const wb = xlsx.utils.book_new();
            const ws = xlsx.utils.json_to_sheet(finalData);

            const styles = {
                headerStyle: {
                    font: { bold: true, color: { rgb: 'FFFFFF' } },
                    fill: { fgColor: { rgb: '351C75' } },
                    alignment: { horizontal: 'center' },
                    border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
                },
                failedStyle: {
                    font: { color: { rgb: 'FF0000' } }, alignment: { horizontal: 'left' },
                    border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
                },
                distinctionStyle: {
                    font: { color: { rgb: '008000' } }, alignment: { horizontal: 'left' },
                    border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
                },
                distinctionBox: {
                    font: { color: { rgb: '008000' } }, alignment: { horizontal: 'left' }, fill: { fgColor: { rgb: 'C4FFBD' } },
                    border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
                },
                failedBox: {
                    font: { color: { rgb: 'FF0000' } }, alignment: { horizontal: 'left' }, fill: { fgColor: { rgb: 'FFBDBD' } },
                    border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
                },
                defaultStyle: {
                    alignment: { horizontal: 'left' },
                    border: {
                        top: { style: 'thin', color: { rgb: '000000' } },
                        bottom: { style: 'thin', color: { rgb: '000000' } },
                        left: { style: 'thin', color: { rgb: '000000' } },
                        right: { style: 'thin', color: { rgb: '000000' } }
                    }
                },
                failedTotalStyle: {
                    font: { color: { rgb: 'FF0000' } }, alignment: { horizontal: 'center' },
                    border: {
                        top: { style: 'thin', color: { rgb: '000000' } },
                        bottom: { style: 'thin', color: { rgb: '000000' } },
                        left: { style: 'thin', color: { rgb: '000000' } },
                        right: { style: 'thin', color: { rgb: '000000' } }
                    }
                },
                distinctionTotalStyle: {
                    font: { color: { rgb: '008000' } }, alignment: { horizontal: 'center' },
                    border: {
                        top: { style: 'thin', color: { rgb: '000000' } },
                        bottom: { style: 'thin', color: { rgb: '000000' } },
                        left: { style: 'thin', color: { rgb: '000000' } },
                        right: { style: 'thin', color: { rgb: '000000' } }
                    }
                },
                srStyle: {
                    alignment: { horizontal: 'center' }, fill: { fgColor: { rgb: 'C9DAF8' } },
                    border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
                }
            };

            const range = xlsx.utils.decode_range(ws['!ref']);
            for (let col = range.s.c; col <= range.e.c; col++) {
                const cellRef = xlsx.utils.encode_cell({ r: 0, c: col });
                if (!ws[cellRef]) continue;
                ws[cellRef].s = styles.headerStyle;
            }

            for (let row = 1; row <= range.e.r; row++) {
                const remarksCellRef = xlsx.utils.encode_cell({ r: row, c: subjects.length + 4 });
                const resultCellRef = xlsx.utils.encode_cell({ r: row, c: 3 });
                const srCellRef = xlsx.utils.encode_cell({ r: row, c: 0 });
                const cellValue = String(ws[remarksCellRef]?.v || '');

                const remarks = ws[remarksCellRef]?.v || '';
                const result = ws[resultCellRef]?.v || '';

                if (ws[srCellRef]) {
                    ws[srCellRef].s = styles.srStyle;
                }

                ws[remarksCellRef].s = cellValue.includes('Failed') ? styles.failedStyle : styles.distinctionStyle;

                if (remarks !== '-') {
                    ws[remarksCellRef].s = result === 'Failed' ? styles.failedStyle : styles.distinctionStyle;
                }

                // Style Subjects columns
                for (let col = 4; col < subjects.length + 4; col++) { // Subjects start from column index 4
                    const cellRef = xlsx.utils.encode_cell({ r: row, c: col });
                    if (!ws[cellRef]) continue;

                    const cellValue = ws[cellRef]?.v || 0;

                    const studentIndex = row - 1;
                    const studentData = data[studentIndex];
                    const subjectKey = subjects[col - 4];

                    const totalMarks = studentData[`${subjectKey}_Total`] || 100;
                    if (cellValue >= calculateEightyFivePercent(totalMarks) && result !== 'Failed' && subjectKey !== 'IslPak') {
                        ws[cellRef].s = styles.distinctionBox;
                    } else if (studentData[`${subjectKey}_Result`] === 'Fail') {
                        ws[cellRef].s = styles.failedBox;
                    }
                }

                // Style Result column
                if (ws[resultCellRef]?.v === 'Failed') {
                    ws[resultCellRef].s = styles.failedStyle;
                }
            }

            // Now, style remaining cells in data rows with default style (left alignment)
            for (let row = 1; row <= range.e.r; row++) {
                for (let col = 0; col <= range.e.c; col++) {
                    const cellRef = xlsx.utils.encode_cell({ r: row, c: col });
                    if (!ws[cellRef]) continue;
                    // If the cell doesn't have a style, or doesn't have an alignment property, assign defaultStyle
                    if (!ws[cellRef].s) {
                        ws[cellRef].s = styles.defaultStyle;
                    } else if (!ws[cellRef].s.alignment) {
                        ws[cellRef].s.alignment = styles.defaultStyle.alignment;
                    }
                }
            }

            // Add header row for distinction and supply counts
            const headerRow = [
                { v: '', t: 's' },
                { v: '', t: 's' },
                { v: '', t: 's' },
                { v: '', t: 's', s: styles.headerStyle },
                ...subjects.map((subject) => ({ v: subject, t: 's', s: styles.headerStyle }))
            ];

            // Add separate row for total distinction counts
            const distinctionRow = [
                { v: '', t: 's' },
                { v: '', t: 's' },
                { v: '', t: 's' },
                { v: 'Distinction', t: 's', s: styles.headerStyle },
                ...subjects.map((subject) => ({
                    v: distinctionCounts[subject] !== undefined ? distinctionCounts[subject].toString() : "0", // Explicitly convert to string
                    t: 's', // Use 's' to explicitly make it a string
                    s: styles.distinctionTotalStyle
                }))
            ];

            // Add separate row for total supply counts
            const supplyRow = [
                { v: '', t: 's' },
                { v: '', t: 's' },
                { v: '', t: 's' },
                { v: 'Supply', t: 's', s: styles.headerStyle },
                ...subjects.map((subject) => ({
                    v: supplyCounts[subject] !== undefined ? supplyCounts[subject].toString() : "0", // Explicitly convert to string
                    t: 's', // Use 's' to explicitly make it a string
                    s: styles.failedTotalStyle
                }))
            ];

            // Append all four rows separately
            xlsx.utils.sheet_add_aoa(ws, [['']], { origin: -1 })
            xlsx.utils.sheet_add_aoa(ws, [headerRow], { origin: -1 });
            xlsx.utils.sheet_add_aoa(ws, [distinctionRow], { origin: -1 });
            xlsx.utils.sheet_add_aoa(ws, [supplyRow], { origin: -1 });

            // Auto-adjust column widths
            const columnWidths = [];
            for (let col = range.s.c; col <= range.e.c; col++) {
                let maxWidth = 0;
                for (let row = 0; row <= range.e.r + 4; row++) {
                    const cellRef = xlsx.utils.encode_cell({ r: row, c: col });
                    const cellValue = ws[cellRef]?.v?.toString() || '';
                    maxWidth = Math.max(maxWidth, cellValue.length);
                }
                columnWidths.push({ wch: maxWidth + 1 });
            }
            ws['!cols'] = columnWidths;

            xlsx.utils.book_append_sheet(wb, ws, 'Results');
            xlsx.writeFile(wb, OUTPUT_XLSX);
            console.log(`✅ Excel file "${OUTPUT_XLSX}" created successfully!`);
        });
}

convertCSVtoExcel();
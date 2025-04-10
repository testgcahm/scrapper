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
    const subjects = [];

    fs.createReadStream(INPUT_CSV)
        .pipe(csv({ headers: false }))
        .on('data', (row) => {
            const entries = Object.values(row);
            const rollNumber = entries[0];
            const studentName = entries[1];
            const result = entries[2];
            const subjectData = entries.slice(3);
            let rawResult = 0;
            let rawResultMatch = result.match(/(\d+)\/(\d+)/);
            if (rawResultMatch) rawResult = parseInt(rawResultMatch[1]);

            let studentRecord = {
                "Roll No.": rollNumber,
                "Name": studentName,
                "Result": result,
                "RawResult": rawResult
            };

            let distinctions = [];
            let supplySubjects = [];

            subjectData.forEach(subjectStr => {
                const match = subjectStr.match(/(.+?)\s*-\s*Th:\s*(\d+)\s*&\s*Pr:\s*(\d+)\s*&\s*(\d+)\/(\d+)\s*\((Pass|Fail)\)/);
                if (match) {
                    const subject = match[1].trim();
                    const theory = parseInt(match[2]);
                    const practical = parseInt(match[3]);
                    const obtained = parseInt(match[4]);
                    const total = parseInt(match[5]);
                    const status = match[6];
                    const mappedSubject = subjectMapping[subject] || subject;

                    if (!subjects.includes(mappedSubject)) {
                        subjects.push(mappedSubject);
                        distinctionCounts[mappedSubject] = 0;
                        supplyCounts[mappedSubject] = 0;
                    }

                    studentRecord[`${mappedSubject}_Th`] = theory;
                    studentRecord[`${mappedSubject}_Pr`] = practical;
                    studentRecord[`${mappedSubject}_Total`] = obtained;
                    studentRecord[`${mappedSubject}_Max`] = total;
                    studentRecord[`${mappedSubject}_Result`] = status;

                    if (obtained >= calculateEightyFivePercent(total) && result !== "Failed" && subject !== "Islamic Studies / Ethics & Pak Studies") {
                        distinctions.push(mappedSubject.toUpperCase());
                        distinctionCounts[mappedSubject]++;
                    }
                    if (status === 'Fail') {
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

            // Build final data as array of arrays
            const finalData = [];
            data.forEach((student, index) => {
                const rowData = [
                    index + 1,
                    student["Roll No."],
                    student["Name"],
                    student["Result"],
                ];
                subjects.forEach(subject => {
                    rowData.push(subject);
                    if (student[`${subject}_Th`] !== 0) rowData.push(student[`${subject}_Th`]);
                    if (student[`${subject}_Pr`] !== 0) rowData.push(student[`${subject}_Pr`]);
                    rowData.push(student[`${subject}_Total`]);
                });
                rowData.push(student["Remarks"]);
                finalData.push(rowData);
            });

            // Apply styles
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
                    font: { bold: true, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'left' }, fill: { fgColor: { rgb: '41ff36' } },
                    border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
                },
                failedBox: {
                    font: { bold: true, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'left' }, fill: { fgColor: { rgb: 'ff3636' } },
                    border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
                },
                theoryStyle: {
                    fill: { fgColor: { rgb: 'edffe7' } }, border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
                },
                practicalStyle: {
                    fill: { fgColor: { rgb: 'eaf5ff' } }, border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
                },
                subjectStyle: {
                    font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '7336ff' } }, border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
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
                srStyle: {
                    alignment: { horizontal: 'center' }, fill: { fgColor: { rgb: 'C9DAF8' } },
                    border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
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
            };

            // Combine all data
            const allData = [
                ["Sr#", "Roll No.", "Name", "Result", ...subjects.map(subject => {
                    if (subject === 'IslPak') {
                        return ["Subject", "Total"]; // For IslPak, we only need "Subject" and "Total"
                    }
                    return ["Subject", "Theory", "Practical", "Total"]; // For other subjects, include all 4 columns
                }).flat(), "Remarks"],
                ...finalData,
            ];
            
            // Create worksheet
            const wb = xlsx.utils.book_new();
            const ws = xlsx.utils.aoa_to_sheet(allData);

            // Apply header styles
            for (let col = 0; col < allData[0].length; col++) {
                const cell = xlsx.utils.encode_cell({ r: 0, c: col });
                ws[cell].s = styles.headerStyle;
            }

            // Apply data row styles
            for (let row = 1; row < allData.length; row++) {
                const rowData = allData[row];
                if (rowData.length === 0) continue;

                for (let col = 0; col < rowData.length; col++) {
                    const cell = xlsx.utils.encode_cell({ r: row, c: col });
                    if (!ws[cell]) continue;

                    // Default style
                    ws[cell].s = styles.defaultStyle;

                    // Sr# style
                    if (col === 0) ws[cell].s = styles.srStyle;

                    // Theory and Practical styles
                    if (col >= 4 && (col - 4) % 4 === 1) ws[cell].s = styles.theoryStyle;
                    if (col >= 4 && (col - 4) % 4 === 2) ws[cell].s = styles.practicalStyle;

                    // Apply distinction and failed box styles to subject totals
                    if (col >= 4 && (col - 4) % 4 === 0) {
                        const subjectIndex = Math.floor((col - 4) / 4);
                        const subject = subjects[subjectIndex];
                        const obtainedMarks = rowData[col + 3];
                        const maxMarks = data[row - 1][`${subject}_Max`];
                        const subjectResult = data[row - 1][`${subject}_Result`];

                        if (maxMarks !== undefined && obtainedMarks >= calculateEightyFivePercent(maxMarks) && rowData[3] !== 'Failed' && subject !== "IslPak") {
                            ws[cell].s = styles.distinctionBox;
                        } else if (subjectResult === 'Fail') {
                            ws[cell].s = styles.failedBox;
                        } else {
                            ws[cell].s = styles.subjectStyle;
                        }
                    }

                    // Result and Remarks styles
                    if (col === 3) {
                        if (rowData[col] === 'Failed') ws[cell].s = styles.failedStyle;
                    }
                    if (col === allData[0].length - 1) {
                        if (rowData[col] && rowData[col] !== '-') {
                            if (rowData[3] === 'Failed') ws[cell].s = styles.failedStyle;
                            else ws[cell].s = styles.distinctionStyle;
                        } else if (rowData[col] && rowData[col] === '-') {
                            ws[cell].s = styles.defaultStyle;
                        }
                    }

                }
            }

            // Apply counts row styles
            const countsStartRow = allData.length - 3;
            for (let row = countsStartRow; row < allData.length; row++) {
                const rowData = allData[row];
                if (rowData?.length === 0) continue;

                for (let col = 0; col < rowData?.length; col++) {
                    const cell = xlsx.utils.encode_cell({ r: row, c: col });
                    if (!ws[cell]) continue;

                    if (rowData[1] === 'Distinction') ws[cell].s = styles.distinctionTotalStyle;
                    if (rowData[1] === 'Supply') ws[cell].s = styles.failedTotalStyle;
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
            xlsx.utils.sheet_add_aoa(ws, [['']], { origin: -1 });
            xlsx.utils.sheet_add_aoa(ws, [headerRow], { origin: -1 });
            xlsx.utils.sheet_add_aoa(ws, [distinctionRow], { origin: -1 });
            xlsx.utils.sheet_add_aoa(ws, [supplyRow], { origin: -1 });

            // Auto-adjust column widths
            const columnWidths = [];
            for (let col = 0; col < allData[0].length; col++) {
                let maxWidth = 0;
                for (let row = 0; row < allData.length; row++) {
                    const cellValue = allData[row][col]?.toString() || '';
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
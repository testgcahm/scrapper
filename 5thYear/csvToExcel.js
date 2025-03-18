const fs = require('fs');
const xlsx = require('xlsx-js-style');
const csv = require('csv-parser');

const INPUT_CSV = 'results.csv';
const OUTPUT_XLSX = 'results.xlsx';

function calculateEightyFivePercent(number) {
    return number * 0.85;
}

async function convertCSVtoExcel() {
    const data = [];
    const distinctionCounts = { Medicine: 0, Surgery: 0, GynObs: 0, Paeds: 0 };
    const supplyCounts = { Medicine: 0, Surgery: 0, GynObs: 0, Paeds: 0 };

    fs.createReadStream(INPUT_CSV)
        .pipe(csv({
            headers: ['Roll No.', 'Names', 'Result', 'Medicine', 'MedicineTotal', 'MedicineResult', 'Surgery', 'SurgeryTotal', 'SurgeryResult', 'GynObs', 'GynObsTotal', 'GynObsResult', 'Paeds', 'PaedsTotal', 'PaedsResult']
        }))
        .on('data', (row) => {
            let match = row['Result'].match(/^(\d+)\/\d+$/);
            row['RawResult'] = match ? parseInt(match[1]) : 0;
            row['Result'] = match ? `${parseInt(match[1])}/1000` : 'Failed';

            let distinctions = [];
            if (parseInt(row['Medicine']) >= calculateEightyFivePercent(parseInt(row['MedicineTotal'])) && row['Result'] !== 'Failed') {
                distinctions.push('Medicine');
                distinctionCounts.Medicine++;
            }
            if (parseInt(row['Surgery']) >= calculateEightyFivePercent(parseInt(row['SurgeryTotal'])) && row['Result'] !== 'Failed') {
                distinctions.push('Surgery');
                distinctionCounts.Surgery++;
            }
            if (parseInt(row['GynObs']) >= calculateEightyFivePercent(parseInt(row['GynObsTotal'])) && row['Result'] !== 'Failed') {
                distinctions.push('GynObs');
                distinctionCounts.GynObs++;
            }
            if (parseInt(row['Paeds']) >= calculateEightyFivePercent(parseInt(row['PaedsTotal'])) && row['Result'] !== 'Failed') {
                distinctions.push('Paeds');
                distinctionCounts.Paeds++;
            }

            let supplySubjects = [];
            if (row['MedicineResult'].toLowerCase() === 'fail') {
                supplySubjects.push('Medicine');
                supplyCounts.Medicine++;
            }
            if (row['SurgeryResult'].toLowerCase() === 'fail') {
                supplySubjects.push('Surgery');
                supplyCounts.Surgery++;
            }
            if (row['GynObsResult'].toLowerCase() === 'fail') {
                supplySubjects.push('GynObs');
                supplyCounts.GynObs++;
            }
            if (row['PaedsResult'].toLowerCase() === 'fail') {
                supplySubjects.push('Paeds');
                supplyCounts.Paeds++;
            }

            // Merge Supply and Distinctions into Remarks
            let remarks = supplySubjects.length > 0 ? supplySubjects.join(', ') : distinctions.length > 0 ? distinctions.join(', ') : '-';
            row['Remarks'] = remarks;

            data.push(row);
        })
        .on('end', () => {
            data.sort((a, b) => {
                if (a.Result === 'Failed') return 1;
                if (b.Result === 'Failed') return -1;
                return b.RawResult - a.RawResult;
            });

            const finalData = data.map(({ MedicineResult, SurgeryResult, GynObsResult, PaedsResult, MedicineTotal, SurgeryTotal, GynObsTotal, PaedsTotal, RawResult, ...rest }, index) => ({
                "Sr#": index + 1,
                ...rest
            }));

            // Create Excel workbook and worksheet
            const wb = xlsx.utils.book_new();
            const ws = xlsx.utils.json_to_sheet(finalData);

            // Define styles
            const headerStyle = {
                font: { bold: true, color: { rgb: 'FFFFFF' } },
                fill: { fgColor: { rgb: '351C75' } },
                alignment: { horizontal: 'center' },
                border: {
                    top: { style: 'thin', color: { rgb: '000000' } },
                    bottom: { style: 'thin', color: { rgb: '000000' } },
                    left: { style: 'thin', color: { rgb: '000000' } },
                    right: { style: 'thin', color: { rgb: '000000' } }
                }
            };
            const failedStyle = {
                font: { color: { rgb: 'FF0000' } }, alignment: { horizontal: 'left' },
                border: {
                    top: { style: 'thin', color: { rgb: '000000' } },
                    bottom: { style: 'thin', color: { rgb: '000000' } },
                    left: { style: 'thin', color: { rgb: '000000' } },
                    right: { style: 'thin', color: { rgb: '000000' } }
                }
            };
            const distinctionStyle = {
                font: { color: { rgb: '008000' } }, alignment: { horizontal: 'left' },
                border: {
                    top: { style: 'thin', color: { rgb: '000000' } },
                    bottom: { style: 'thin', color: { rgb: '000000' } },
                    left: { style: 'thin', color: { rgb: '000000' } },
                    right: { style: 'thin', color: { rgb: '000000' } }
                }
            };
            const failedTotalStyle = {
                font: { color: { rgb: 'FF0000' } }, alignment: { horizontal: 'center' },
                border: {
                    top: { style: 'thin', color: { rgb: '000000' } },
                    bottom: { style: 'thin', color: { rgb: '000000' } },
                    left: { style: 'thin', color: { rgb: '000000' } },
                    right: { style: 'thin', color: { rgb: '000000' } }
                }
            };
            const distinctionTotalStyle = {
                font: { color: { rgb: '008000' } }, alignment: { horizontal: 'center' },
                border: {
                    top: { style: 'thin', color: { rgb: '000000' } },
                    bottom: { style: 'thin', color: { rgb: '000000' } },
                    left: { style: 'thin', color: { rgb: '000000' } },
                    right: { style: 'thin', color: { rgb: '000000' } }
                }
            };
            const defaultStyle = {
                alignment: { horizontal: 'left' },
                border: {
                    top: { style: 'thin', color: { rgb: '000000' } },
                    bottom: { style: 'thin', color: { rgb: '000000' } },
                    left: { style: 'thin', color: { rgb: '000000' } },
                    right: { style: 'thin', color: { rgb: '000000' } }
                }
            };
            const distinctionStyleBox = {
                font: { color: { rgb: '008000' } }, alignment: { horizontal: 'center' }, fill: { fgColor: { rgb: 'C4FFBD' } },
                border: {
                    top: { style: 'thin', color: { rgb: '000000' } },
                    bottom: { style: 'thin', color: { rgb: '000000' } },
                    left: { style: 'thin', color: { rgb: '000000' } },
                    right: { style: 'thin', color: { rgb: '000000' } }
                }
            };
            const failedStyleBox = {
                font: { color: { rgb: 'FF0000' } }, alignment: { horizontal: 'center' }, fill: { fgColor: { rgb: 'FFBDBD' } },
                border: {
                    top: { style: 'thin', color: { rgb: '000000' } },
                    bottom: { style: 'thin', color: { rgb: '000000' } },
                    left: { style: 'thin', color: { rgb: '000000' } },
                    right: { style: 'thin', color: { rgb: '000000' } }
                }
            };
            const SrStyleBox = {
                alignment: { horizontal: 'center' }, fill: { fgColor: { rgb: 'C9DAF8' } },
                border: {
                    top: { style: 'thin', color: { rgb: '000000' } },
                    bottom: { style: 'thin', color: { rgb: '000000' } },
                    left: { style: 'thin', color: { rgb: '000000' } },
                    right: { style: 'thin', color: { rgb: '000000' } }
                }
            };

            // Style headers
            const range = xlsx.utils.decode_range(ws['!ref']);
            for (let col = range.s.c; col <= range.e.c; col++) {
                const cellRef = xlsx.utils.encode_cell({ r: 0, c: col });
                if (!ws[cellRef]) continue;
                ws[cellRef].s = headerStyle;
            }

            // Style data rows
            for (let row = 1; row <= range.e.r; row++) {
                const remarksCellRef = xlsx.utils.encode_cell({ r: row, c: 8 });
                const resultCellRef = xlsx.utils.encode_cell({ r: row, c: 3 });
                const MedicineCellRef = xlsx.utils.encode_cell({ r: row, c: 4 });
                const SurgeryCellRef = xlsx.utils.encode_cell({ r: row, c: 5 });
                const GynObsCellRef = xlsx.utils.encode_cell({ r: row, c: 6 });
                const PaedsCellRef = xlsx.utils.encode_cell({ r: row, c: 7 });
                const srCellRef = xlsx.utils.encode_cell({ r: row, c: 0 });

                if (ws[srCellRef]) {
                    ws[srCellRef].s = SrStyleBox;
                }

                const remarks = ws[remarksCellRef]?.v || '';
                const result = ws[resultCellRef]?.v || '';
                const medicine = ws[MedicineCellRef]?.v || 0;
                const surgery = ws[SurgeryCellRef]?.v || 0;
                const gynObs = ws[GynObsCellRef]?.v || 0;
                const paeds = ws[PaedsCellRef]?.v || 0;

                if (remarks !== '-') {
                    ws[remarksCellRef].s = result === 'Failed' ? failedStyle : distinctionStyle;
                }
                if (medicine >= calculateEightyFivePercent(data[row - 1].MedicineTotal) && result !== 'Failed') {
                    ws[MedicineCellRef].s = distinctionStyleBox;
                }
                if (data[row - 1].MedicineResult === 'Fail') {
                    ws[MedicineCellRef].s = failedStyleBox;
                }
                if (surgery >= calculateEightyFivePercent(data[row - 1].SurgeryTotal) && result !== 'Failed') {
                    ws[SurgeryCellRef].s = distinctionStyleBox;
                }
                if (data[row - 1].SurgeryResult === 'Fail') {
                    ws[SurgeryCellRef].s = failedStyleBox;
                }
                if (gynObs >= calculateEightyFivePercent(data[row - 1].GynObsTotal) && result !== 'Failed') {
                    ws[GynObsCellRef].s = distinctionStyleBox;
                }
                if (data[row - 1].GynObsResult === 'Fail') {
                    ws[GynObsCellRef].s = failedStyleBox;
                }
                if (paeds >= calculateEightyFivePercent(data[row - 1].PaedsTotal) && result !== 'Failed') {
                ws[PaedsCellRef].s = distinctionStyleBox;
            }
            if (data[row - 1].PaedsResult === 'Fail') {
                ws[PaedsCellRef].s = failedStyleBox;
            }

            // Style Result column
            if (ws[resultCellRef]?.v === 'Failed') {
                ws[resultCellRef].s = failedStyle;
            }
        }

            // Now, style remaining cells in data rows with default style (left alignment)
            for (let row = 1; row <= range.e.r; row++) {
        for (let col = 0; col <= range.e.c; col++) {
            const cellRef = xlsx.utils.encode_cell({ r: row, c: col });
            if (!ws[cellRef]) continue;
            // If the cell doesn't have a style, or doesn't have an alignment property, assign defaultStyle
            if (!ws[cellRef].s) {
                ws[cellRef].s = defaultStyle;
            } else if (!ws[cellRef].s.alignment) {
                ws[cellRef].s.alignment = defaultStyle.alignment;
            }
        }
    }

    // Add header row for distinction and supply counts
    const headerRow = [
        { v: '', t: 's' },
        { v: '', t: 's' },
        { v: '', t: 's' },
        { v: '', t: 's', s: headerStyle },
        { v: 'Medicine', t: 's', s: headerStyle },
        { v: 'Surgery', t: 's', s: headerStyle },
        { v: 'GynObs', t: 's', s: headerStyle },
        { v: 'Paeds', t: 's', s: headerStyle }
    ];

    // Add separate row for total distinction counts
    const distinctionRow = [
        { v: '', t: 's' },
        { v: '', t: 's' },
        { v: '', t: 's' },
        { v: 'Distinction', t: 's', s: headerStyle },
        { v: distinctionCounts.Medicine, t: 'n', s: distinctionTotalStyle },
        { v: distinctionCounts.Surgery, t: 'n', s: distinctionTotalStyle },
        { v: distinctionCounts.GynObs, t: 'n', s: distinctionTotalStyle },
        { v: distinctionCounts.Paeds, t: 'n', s: distinctionTotalStyle }
    ];

    // Add separate row for total supply counts
    const supplyRow = [
        { v: '', t: 's' },
        { v: '', t: 's' },
        { v: '', t: 's' },
        { v: 'Supply', t: 's', s: headerStyle },
        { v: supplyCounts.Medicine, t: 'n', s: failedTotalStyle },
        { v: supplyCounts.Surgery, t: 'n', s: failedTotalStyle },
        { v: supplyCounts.GynObs, t: 'n', s: failedTotalStyle },
        { v: supplyCounts.Paeds, t: 'n', s: failedTotalStyle }
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
        for (let row = 0; row <= range.e.r + 2; row++) {
            const cellRef = xlsx.utils.encode_cell({ r: row, c: col });
            const cellValue = ws[cellRef]?.v?.toString() || '';
            maxWidth = Math.max(maxWidth, cellValue.length);
        }
        columnWidths.push({ wch: maxWidth + 1 });
    }
    ws['!cols'] = columnWidths;

    // Append worksheet to workbook and save
    xlsx.utils.book_append_sheet(wb, ws, 'Results');
    xlsx.writeFile(wb, OUTPUT_XLSX);
    console.log(`✅ Excel file "${OUTPUT_XLSX}" created successfully!`);
});
}

convertCSVtoExcel();

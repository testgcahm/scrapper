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
    const distinctionCounts = { Cmed: 0, Ent: 0, Eye: 0, Patho: 0 };
    const supplyCounts = { Cmed: 0, Ent: 0, Eye: 0, Patho: 0 };

    fs.createReadStream(INPUT_CSV)
        .pipe(csv({
            headers: ['Roll No.', 'Names', 'Result', 'Cmed', 'CmedResult', 'CmedTotal', 'Ent', 'EntResult', 'EntTotal', 'Eye', 'EyeResult', 'EyeTotal', 'Patho', 'PathoResult', 'PathoTotal']
        }))
        .on('data', (row) => {
            let match = row['Result'].match(/^(\d+)\/\d+$/);
            row['RawResult'] = match ? parseInt(match[1]) : 0;
            row['Result'] = match ? `${parseInt(match[1])}/1000` : 'Failed';

            let distinctions = [];
            if (parseInt(row['Ent']) >= calculateEightyFivePercent(parseInt(row['EntTotal'])) && row['Result'] !== 'Failed') {
                distinctions.push('ENT');
                distinctionCounts.Ent++;
            }
            if (parseInt(row['Eye']) >= calculateEightyFivePercent(parseInt(row['EyeTotal'])) && row['Result'] !== 'Failed') {
                distinctions.push('EYE');
                distinctionCounts.Eye++;
            }
            if (parseInt(row['Patho']) >= calculateEightyFivePercent(parseInt(row['PathoTotal'])) && row['Result'] !== 'Failed') {
                distinctions.push('PATHO');
                distinctionCounts.Patho++;
            }
            if (parseInt(row['Cmed']) >= calculateEightyFivePercent(parseInt(row['CmedTotal'])) && row['Result'] !== 'Failed') {
                distinctions.push('CMED');
                distinctionCounts.Cmed++;
            }

            let supplySubjects = [];
            if (row['CmedResult'].toLowerCase() === 'fail') {
                supplySubjects.push('CMED');
                supplyCounts.Cmed++;
            }
            if (row['EntResult'].toLowerCase() === 'fail') {
                supplySubjects.push('ENT');
                supplyCounts.Ent++;
            }
            if (row['EyeResult'].toLowerCase() === 'fail') {
                supplySubjects.push('EYE');
                supplyCounts.Eye++;
            }
            if (row['PathoResult'].toLowerCase() === 'fail') {
                supplySubjects.push('PATHO');
                supplyCounts.Patho++;
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

            const finalData = data.map(({ CmedResult, EntResult, EyeResult, PathoResult, CmedTotal, EntTotal, EyeTotal, PathoTotal, RawResult, ...rest }, index) => ({
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
                const cmedCellRef = xlsx.utils.encode_cell({ r: row, c: 4 });
                const entCellRef = xlsx.utils.encode_cell({ r: row, c: 5 });
                const eyeCellRef = xlsx.utils.encode_cell({ r: row, c: 6 });
                const pathoCellRef = xlsx.utils.encode_cell({ r: row, c: 7 });
                const srCellRef = xlsx.utils.encode_cell({ r: row, c: 0 });

                if (ws[srCellRef]) {
                    ws[srCellRef].s = SrStyleBox;
                }

                const remarks = ws[remarksCellRef]?.v || '';
                const result = ws[resultCellRef]?.v || '';
                const cmed = ws[cmedCellRef]?.v || 0;
                const ent = ws[entCellRef]?.v || 0;
                const eye = ws[eyeCellRef]?.v || 0;
                const patho = ws[pathoCellRef]?.v || 0;

                if (remarks !== '-') {
                    ws[remarksCellRef].s = result === 'Failed' ? failedStyle : distinctionStyle;
                }
                if (cmed >= calculateEightyFivePercent(data[row - 1].CmedTotal) && result !== 'Failed') {
                    ws[cmedCellRef].s = distinctionStyleBox;
                }
                if (data[row - 1].CmedResult === 'Fail') {
                    ws[cmedCellRef].s = failedStyleBox;
                }
                if (ent >= calculateEightyFivePercent(data[row - 1].EntTotal) && result !== 'Failed') {
                    ws[entCellRef].s = distinctionStyleBox;
                }
                if (data[row - 1].EntResult === 'Fail') {
                    ws[entCellRef].s = failedStyleBox;
                }
                if (eye >= calculateEightyFivePercent(data[row - 1].EyeTotal) && result !== 'Failed') {
                    ws[eyeCellRef].s = distinctionStyleBox;
                }
                if (data[row - 1].EyeResult === 'Fail') {
                    ws[eyeCellRef].s = failedStyleBox;
                }
                if (patho >= calculateEightyFivePercent(data[row - 1].PathoTotal) && result !== 'Failed') {
                    ws[pathoCellRef].s = distinctionStyleBox;
                }
                if (data[row - 1].PathoResult === 'Fail') {
                    ws[pathoCellRef].s = failedStyleBox;
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
                { v: 'Cmed', t: 's', s: headerStyle },
                { v: 'Ent', t: 's', s: headerStyle },
                { v: 'Eye', t: 's', s: headerStyle },
                { v: 'Patho', t: 's', s: headerStyle }
            ];

            // Add separate row for total distinction counts
            const distinctionRow = [
                { v: '', t: 's' },
                { v: '', t: 's' },
                { v: '', t: 's' },
                { v: 'Distinction', t: 's', s: headerStyle },
                { v: distinctionCounts.Cmed, t: 'n', s: distinctionTotalStyle },
                { v: distinctionCounts.Ent, t: 'n', s: distinctionTotalStyle },
                { v: distinctionCounts.Eye, t: 'n', s: distinctionTotalStyle },
                { v: distinctionCounts.Patho, t: 'n', s: distinctionTotalStyle }
            ];

            // Add separate row for total supply counts
            const supplyRow = [
                { v: '', t: 's' },
                { v: '', t: 's' },
                { v: '', t: 's' },
                { v: 'Supply', t: 's', s: headerStyle },
                { v: supplyCounts.Cmed, t: 'n', s: failedTotalStyle },
                { v: supplyCounts.Ent, t: 'n', s: failedTotalStyle },
                { v: supplyCounts.Eye, t: 'n', s: failedTotalStyle },
                { v: supplyCounts.Patho, t: 'n', s: failedTotalStyle }
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

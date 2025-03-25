const fs = require('fs');
const xlsx = require('xlsx-js-style');
const csv = require('csv-parser');

const INPUT_CSV = 'results.csv';
const OUTPUT_XLSX = 'results.xlsx';

async function convertCSVtoExcel() {
    const data = [];

    fs.createReadStream(INPUT_CSV)
        .pipe(csv({ headers: ['Roll No.', 'Names', 'Result'] }))
        .on('data', (row) => {
            let match = row['Result'].match(/^(\d+)\/(\d+)$/);
            row['RawResult'] = match ? parseInt(match[1]) : 0;
            row['Result'] = match ? `${parseInt(match[1])}/${parseInt(match[2])}` : 'Failed';
            data.push(row);
        })
        .on('end', () => {
            data.sort((a, b) => b.RawResult - a.RawResult);

            const finalData = data.map(({ RawResult, ...rest }, index) => ({
                "Sr#": index + 1,
                ...rest
            }));

            const styles = {
                headerStyle: {
                    font: { bold: true, color: { rgb: 'FFFFFF' } },
                    fill: { fgColor: { rgb: '351C75' } },
                    alignment: { horizontal: 'center' },
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
                srStyle: {
                    alignment: { horizontal: 'center' }, fill: { fgColor: { rgb: 'C9DAF8' } },
                    border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
                },
                failedStyle: {
                    font: { color: { rgb: 'FF0000' } },
                    border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
                }
            };

            const wb = xlsx.utils.book_new();
            const ws = xlsx.utils.json_to_sheet(finalData);

            const range = xlsx.utils.decode_range(ws['!ref']);
            for (let col = range.s.c; col <= range.e.c; col++) {
                const cellRef = xlsx.utils.encode_cell({ r: 0, c: col });
                if (!ws[cellRef]) continue;
                ws[cellRef].s = styles.headerStyle;
            }

            for (let row = 1; row <= range.e.r; row++) {
                const resultCellRef = xlsx.utils.encode_cell({ r: row, c: 3 });
                const srCellRef = xlsx.utils.encode_cell({ r: row, c: 0 });

                if (ws[srCellRef]) {
                    ws[srCellRef].s = styles.srStyle;
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

            const columnWidths = [];
            for (let col = range.s.c; col <= range.e.c; col++) {
                let maxWidth = 1; // Default minimum width
                for (let row = 0; row <= range.e.r; row++) {
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
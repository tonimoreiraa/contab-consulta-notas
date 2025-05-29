import puppeteer from "puppeteer";
import xlsx from 'xlsx'
import fs from 'fs'
import path from "path";

declare var document: any

async function main()
{
    if (!fs.existsSync('output')) {
        fs.mkdirSync('output');
        fs.mkdirSync('output/prints');
    }

    const reports = fs.readdirSync('relatorios').filter(file => file.endsWith('.xlsx'));

    for (const report of reports) {
        console.log('Relatório: ', report)
        const workbook = xlsx.readFile(`relatorios/${report}`);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const range = xlsx.utils.decode_range(worksheet['!ref']!);

        const notas: any = [];
        for (let row = range.s.r; row <= range.e.r; row++) {
            const cellAddress = { c: 3, r: row }; // coluna D = índice 3
            const cellRef = xlsx.utils.encode_cell(cellAddress);
            const cell = worksheet[cellRef];
            if (cell?.v?.length > 10 && !String(cell.v).startsWith('27')) {
                notas.push(cell.v)
            }
        }

        const browser = await puppeteer.launch({
            headless: false,
        })

        const table: any = []

        const page = await browser.newPage();
        await page.goto('https://contribuinte.sefaz.al.gov.br/cobrancadfe/#/consultar-valor-imposto-nfe', {
            waitUntil: 'networkidle0'
        })
        for (const nota of notas) {
            console.log('Nota: ', nota.trim())
            const input = await page.waitForSelector('#field_chaveNota')
            await input?.type(nota);
            const button = await page.waitForSelector('body > jhi-main > div.container-fluid > div > jhi-consultar-valor-imposto-nfe > div > form > div > div.input-group.mb-3 > div.input-group-append > button')
            await button?.click()
            await page.waitForNetworkIdle()

            // Extrai planilha
            // @ts-nocheck
            const result = await page.evaluate(() => {
                const table = document.querySelector('body > jhi-main > div.container-fluid > div > jhi-consultar-valor-imposto-nfe > div > div.table-responsive > div > table');
                if (!table) return [];

                const headers = Array.from(table.querySelectorAll('thead th')).map((th: any) => th.innerText.trim());
                const rows = Array.from(table.querySelectorAll('tbody tr'));

                return rows.map((row: any) => {
                    const cells = Array.from(row.querySelectorAll('td')) as any;
                    return headers.reduce((obj, header, index) => {
                        obj[header] = cells[index]?.innerText.trim() || '';
                        return obj;
                    }, {});
                });
            });

            table.push(...result.map(i => ({
                'Nota': nota,
                ...i,
            })))

            if (table.length > 0) {
                const ws = xlsx.utils.json_to_sheet(table);
                const wb = xlsx.utils.book_new();
                xlsx.utils.book_append_sheet(wb, ws, "Notas");
                const reportName = path.basename(report);
                xlsx.writeFile(wb, `output/${reportName}.xlsx`);
            }

            await page.reload()
            await page.screenshot({ path: `./output/prints/${nota}.png`, fullPage: true })
        }
    }

}

main()
import fs from 'fs-extra';
import archiver from 'archiver';
import { escape } from 'he';

function normalizeText(text: string): string {
    if (text == undefined || text == null)
        return text;
    let chars = text.split("");
    for (var i = 0; i < chars.length; i++) {
        if (chars[i].charCodeAt(0) == 8230) {
            chars[i] = "...";
        }
        if (chars[i].charCodeAt(0) == 8212) {
            chars[i] = "-";
        }
        if (chars[i].charCodeAt(0) == 8216 ||
            chars[i].charCodeAt(0) == 8217 ||
            chars[i].charCodeAt(0) == 8249 ||
            chars[i].charCodeAt(0) == 8250 ||
            chars[i].charCodeAt(0) == 8216 ||
            chars[i].charCodeAt(0) == 8217 ||
            chars[i].charCodeAt(0) == 8218 ||
            chars[i].charCodeAt(0) == 8219 ||
            false) {
            chars[i] = "'";
        }
        if (chars[i].charCodeAt(0) == 171 ||
            chars[i].charCodeAt(0) == 187 ||
            chars[i].charCodeAt(0) == 8220 ||
            chars[i].charCodeAt(0) == 8221 ||
            chars[i].charCodeAt(0) == 8222 ||
            chars[i].charCodeAt(0) == 8223 ||
            chars[i].charCodeAt(0) == 11842 ||
            chars[i].charCodeAt(0) == 65282 ||
            false) {
            chars[i] = '"';
        }
        if ((chars[i] == "'" || chars[i] == '"') && chars[i] == chars[i + 1]) {
            chars.splice(i, 1);
            i--;
            continue;
        }
    }
    return chars.join("");
}
function normalizeString(str: string): string {
    return normalizeText(str).normalize('NFD').replace(/[\u0300-\u036f]|[^0-9a-zA-Z!@#$%^&*()_+=\-[\]{}|;':",.<>?/\\ ]/g, '');
}

// this is obviously not accurate, but it is quick and dirty way to get this done
const charWidths: Record<string, number> = {
    '0': 6, '1': 6, '2': 6, '3': 6, '4': 6, '5': 6, '6': 6, '7': 6, '8': 6, '9': 6, ' ': 3, '!': 4, '\"': 4.9, '#': 6, '$': 6, '%': 10, '&': 9.333, '\'': 2.166, '(': 4, ')': 4, '*': 6, '+': 6.7666, ',': 3, '-': 4, '.': 3, '/': 3.3333, ':': 3.33333, ';': 3.33333, '<': 6.76666, '=': 6.76666, '>': 6.76666, '?': 5.3333, '@': 11.0500, 'A': 8.6666, 'B': 8, 'C': 8, 'D': 8.6666, 'E': 7.3333, 'F': 6.6666, 'G': 8.6666, 'H': 8.6666, 'I': 4, 'J': 4.6666, 'K': 8.6666, 'L': 7.3333, 'M': 10.6666, 'N': 8.6666, 'O': 8.6666, 'P': 6.6666, 'Q': 8.6666, 'R': 8, 'S': 6.6666, 'T': 7.3333, 'U': 8.6666, 'V': 8.6666, 'W': 11.33333, 'X': 8.6666, 'Y': 8.6666, 'Z': 7.3333, '[': 4, '\\': 3.33333, ']': 4, '^': 5.6333, '_': 6, '`': 4, 'a': 5.3333, 'b': 6, 'c': 5.3333, 'd': 6, 'e': 5.3333, 'f': 4, 'g': 6, 'h': 6, 'i': 3.33333, 'j': 3.33333, 'k': 6, 'l': 3.33333, 'm': 9.33333, 'n': 6, 'o': 6, 'p': 6, 'q': 6, 'r': 4, 's': 4.6666, 't': 3.33333, 'u': 6, 'v': 6, 'w': 8.6666, 'x': 6, 'y': 6, 'z': 5.3333, '{': 5.76666, '|': 2.4, '}': 5.76666, '~': 6.5,
}
function getCharWidth(chr: string): number {
    return charWidths[chr] ?? 12;
}

export async function toXLSX(data: string[][], outputFilePath: string) {
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
<sheetPr filterMode="false"><pageSetUpPr fitToPage="false"/></sheetPr><dimension ref="A1:Z1"/>
<sheetViews><sheetView showFormulas="false" showGridLines="true" showRowColHeaders="true" showZeros="true" rightToLeft="false" tabSelected="true" showOutlineSymbols="true" defaultGridColor="true" view="normal" topLeftCell="A1" colorId="64" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100" workbookViewId="0">
<selection pane="topLeft" activeCell="A1" activeCellId="0" sqref="A1"/>
</sheetView></sheetViews>\n`;

    xml += '<sheetFormatPr defaultRowHeight="0" autoSizeCol="true"/>\n';
    xml += '<cols>\n';

    let maxWidth: number[] = [];
    data = data.map(row => row.map(cell => ("" + (cell ?? ""))));
    for (let row of data) {
        let i = 0;
        for (let cell of row) {
            let lines = cell.split(/\r\n|\n|\r/);
            for (let line of lines) {
                line = normalizeString(line);
                let maxLineWidth = 0;
                for (let j = 0; j < line.length; j++) {
                    maxLineWidth += getCharWidth(line.charAt(j));
                }
                maxWidth[i] = maxWidth[i] ? Math.max(maxLineWidth, maxWidth[i]) : maxLineWidth;
            }
            i++;
        }
    }

    let i = 0;
    for (let width of maxWidth) {
        xml += '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + (width / 6 + 1) + '" customWidth="true"/>\n';
        i++;
    }
    xml += '</cols>\n';
    xml += '<sheetData>\n';

    // Write the data to the XML
    for (let row of data) {
        let maxHeight = 0;
        for (let cell of row) {
            let cellHeight = (cell.match(/\r\n|\n|\r/g) || []).length + 1;
            if (cellHeight > maxHeight) {
                maxHeight = cellHeight;
            }
        }
        xml += '<row ht="' + (maxHeight * 11.5 + 1.3) + '">\n';
        for (let i = 0; i < row.length; i++) {
            let cell = escape(row[i]);
            xml += '<c s="1" t="inlineStr">\n';
            xml += '<is>\n';
            xml += '<t xml:space="preserve">' + cell + '</t>\n';
            xml += '</is>\n';
            xml += '</c>\n';
        }
        xml += '</row>\n';
    }
    xml += '</sheetData>\n';

    // Close the XML
    xml += `
<printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>
<pageMargins left="0.7875" right="0.7875" top="1.05277777777778" bottom="1.05277777777778" header="0.7875" footer="0.7875"/><pageSetup paperSize="9" scale="100" fitToWidth="1" fitToHeight="1" pageOrder="downThenOver" orientation="portrait" blackAndWhite="false" draft="false" cellComments="none" firstPageNumber="1" useFirstPageNumber="true" horizontalDpi="300" verticalDpi="300" copies="1"/>
<headerFooter differentFirst="false" differentOddEven="false"><oddHeader>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;A</oddHeader><oddFooter>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12Page &amp;P</oddFooter></headerFooter>
</worksheet>`;

    const xmlFilePath = __dirname+'/static/xlsx/xl/worksheets/sheet1.xml';

    // Delete sheet1.xml file if it exists
    if (await fs.exists(xmlFilePath)) {
        await fs.unlink(xmlFilePath);
    }

    // Output the XML as the XLSX file
    await fs.writeFile(xmlFilePath, xml);

    // Create a zip archive
    const output = fs.createWriteStream(outputFilePath);
    const archive = archiver('zip', { zlib: { level: 9 } });

    let doneFn: Function= ()=>{};
    let errorFn: Function = ()=>{};
    let promiseDone = new Promise((_doneFn,_errorFn)=>{
        doneFn = _doneFn;
        errorFn = _errorFn;
    });

    output.on('close', function () {
        doneFn();
    });
    output.on('error', function () {
        errorFn();
    })

    archive.pipe(output);

    const addFolderToZip = (dir: string, zip: archiver.Archiver, zipdir: string = '') => {
        if (fs.existsSync(dir) && fs.lstatSync(dir).isDirectory()) {
            const files = fs.readdirSync(dir);
            files.forEach(file => {
                if (file !== '.' && file !== '..') {
                    if (fs.lstatSync(dir + file).isDirectory()) {
                        addFolderToZip(dir + file + '/', zip, zipdir + file + '/');
                    } else {
                        zip.append(fs.createReadStream(dir + file), { name: zipdir + file });
                    }
                }
            });
        }
    };

    addFolderToZip(__dirname+'/static/xlsx/', archive);

    archive.finalize();

    await promiseDone;
}

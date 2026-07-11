const XLSX = require('xlsx');
const fs = require('fs');

const xlsxPath = 'C:\\Users\\Lenovo\\Desktop\\蓝染纹样生成数据库.xlsx';
const outPath = 'D:\\WorkBuddy\\2026-06-18-16-15-27\\xlsx_data.json';

const workbook = XLSX.readFile(xlsxPath);
const result = {
    sheetNames: workbook.SheetNames,
    sheets: {}
};

for (const name of workbook.SheetNames) {
    const sheet = workbook.Sheets[name];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    result.sheets[name] = {
        rowCount: data.length,
        headers: data[0] || [],
        sampleRows: data.slice(0, 20).map(row =>
            row.map(cell => {
                if (cell === null || cell === undefined) return '';
                const s = String(cell);
                return s.length > 100 ? s.substring(0, 100) + '...' : s;
            })
        )
    };
}

// Also extract images using the ZIP structure
const AdmZip = require('adm-zip');
try {
    const zip = new AdmZip(xlsxPath);
    const entries = zip.getEntries();
    const imgEntries = entries.filter(e =>
        /\.(png|jpg|jpeg|gif|bmp|webp|svg)$/i.test(e.entryName) &&
        !e.isDirectory
    );

    result.images = imgEntries.map(e => ({
        name: e.entryName,
        size: e.header.size,
        sizeKB: (e.header.size / 1024).toFixed(1)
    }));

    // Extract images to a folder
    const imgDir = 'D:\\WorkBuddy\\2026-06-18-16-15-27\\extracted_images';
    if (!fs.existsSync(imgDir)) fs.mkdirSync(imgDir, { recursive: true });

    for (const img of imgEntries) {
        const fileName = img.entryName.replace(/^.*[\\/]/, '');
        const outFile = imgDir + '\\' + fileName;
        if (!fs.existsSync(outFile)) {
            fs.writeFileSync(outFile, img.getData());
        }
    }
    result.extractedTo = imgDir;
} catch (e) {
    result.imageError = e.message;
}

fs.writeFileSync(outPath, JSON.stringify(result, null, 2), 'utf8');
console.log(JSON.stringify(result, null, 2));

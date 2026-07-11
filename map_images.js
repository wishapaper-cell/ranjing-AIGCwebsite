const XLSX = require('xlsx');
const fs = require('fs');
const AdmZip = require('adm-zip');

const xlsxPath = 'C:\\Users\\Lenovo\\Desktop\\蓝染纹样生成数据库.xlsx';

const zip = new AdmZip(xlsxPath);
const entries = zip.getEntries();

// Extract relationship mapping files
const relFiles = entries.filter(e =>
    e.entryName.includes('_rels') &&
    (e.entryName.includes('cellimages') || e.entryName.includes('drawings'))
);

console.log("=== Relationship files ===");
for (const f of relFiles) {
    console.log(`\n--- ${f.entryName} ---`);
    console.log(f.getData().toString('utf8'));
}

// Also look for drawing files
const drawingFiles = entries.filter(e =>
    e.entryName.includes('drawings') && e.entryName.endsWith('.xml')
);
console.log("\n\n=== Drawing files ===");
for (const f of drawingFiles) {
    console.log(`\n--- ${f.entryName} ---`);
    const content = f.getData().toString('utf8');
    console.log(content.substring(0, 5000));
}

// Look for any file referencing the DISPIMG IDs
console.log("\n\n=== Files referencing DISPIMG IDs ===");
for (const entry of entries) {
    if (entry.isDirectory) continue;
    if (entry.entryName.match(/\.(xml|rels)$/)) {
        const content = entry.getData().toString('utf8');
        if (content.includes('B38F1D275B1B42FE8648AF12928B9562')) {
            console.log(`Found in: ${entry.entryName}`);
            console.log(content.substring(0, 3000));
            console.log('...');
        }
    }
}

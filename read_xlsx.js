const fs = require('fs');
const path = require('path');

const xlsxPath = process.argv[2] || 'C:\\Users\\Lenovo\\Desktop\\蓝染纹样生成数据库.xlsx';
const outPath = process.argv[3] || 'D:\\WorkBuddy\\2026-06-18-16-15-27\\xlsx_output.txt';

// Simple XLSX reader without external deps
const data = fs.readFileSync(xlsxPath);
const log = [];

function readStr(buf, off, len) {
    return buf.toString('utf8', off, off + len).replace(/\0/g, '').trim();
}

// XLSX is a ZIP file
// Look for xl/sharedStrings.xml and xl/worksheets/
function findInZip(buffer, filename) {
    const str = buffer.toString('binary');
    const idx = str.indexOf(filename);
    if (idx === -1) return null;
    // Find the file entry in central directory
    // This is a simplified approach
    return idx;
}

log.push("File size: " + (data.length / 1024 / 1024).toFixed(2) + " MB");
log.push("File header: " + data.slice(0, 4).toString('hex'));

// Check if it's a ZIP (PK header)
if (data[0] === 0x50 && data[1] === 0x4B) {
    log.push("Valid ZIP/XLSX file detected");

    // Look for sheet names in workbook.xml
    const wbIdx = data.indexOf('workbook.xml');
    if (wbIdx > 0) {
        // Find sheet names nearby
        const chunk = data.slice(Math.max(0, wbIdx - 100), wbIdx + 2000).toString('utf8');
        const sheetMatches = chunk.match(/name="([^"]+)"/g);
        if (sheetMatches) {
            log.push("\nSheet names:");
            sheetMatches.forEach(m => log.push("  " + m));
        }
    }

    // Look for shared strings
    const ssIdx = data.indexOf('sharedStrings.xml');
    if (ssIdx > 0) {
        const chunk = data.slice(ssIdx, Math.min(ssIdx + 50000, data.length)).toString('utf8');
        const siMatches = chunk.match(/<si>[\s\S]*?<\/si>/g);
        if (siMatches) {
            log.push("\nShared Strings (first 50):");
            siMatches.slice(0, 50).forEach((si, i) => {
                let text = si.replace(/<[^>]+>/g, '').trim();
                if (text.length > 100) text = text.substring(0, 100) + '...';
                if (text) log.push(`  [${i}] ${text}`);
            });
        }
    }

    // Try to find image files in the ZIP
    log.push("\n\nImage files in XLSX:");
    const imgExts = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.svg'];
    let imgCount = 0;
    for (const ext of imgExts) {
        let pos = 0;
        while (pos < data.length) {
            const idx = data.indexOf(ext, pos);
            if (idx === -1) break;
            // Check if it looks like a filename in ZIP
            const before = data.slice(Math.max(0, idx - 30), idx).toString('binary');
            const fnMatch = before.match(/([\w\/.-]+)$/);
            if (fnMatch) {
                imgCount++;
                if (imgCount <= 50) {
                    log.push(`  ${fnMatch[1]}${ext}`);
                }
            }
            pos = idx + 1;
        }
    }
    if (imgCount === 0) log.push("  (no images found in ZIP)");
    else if (imgCount > 50) log.push(`  ... and ${imgCount - 50} more`);
} else {
    log.push("Not a valid ZIP/XLSX file");
}

fs.writeFileSync(outPath, log.join('\n'), 'utf8');
console.log("Output written to: " + outPath);
console.log("Lines: " + log.length);

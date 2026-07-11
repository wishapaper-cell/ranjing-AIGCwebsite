const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');

const WORK_DIR = 'D:\\WorkBuddy\\2026-06-18-16-15-27';
const XLSX_PATH = 'C:\\Users\\Lenovo\\Desktop\\蓝染纹样生成数据库.xlsx';
const HTML_SRC = path.join(WORK_DIR, '染境-蓝染非遗AI纹样生成平台-v2.html');
const IMG_SRC = path.join(WORK_DIR, 'extracted_images');
const IMG_DST = path.join(WORK_DIR, 'images');

// Step 1: Extract full DISPIMG ID → image file mapping
console.log('=== Step 1: Extracting image mapping ===');
const zip = new AdmZip(XLSX_PATH);

// Parse cellimages.xml to get DISPIMG ID → rId
const cellImagesXml = zip.readAsText('xl/cellimages.xml');
const dispimgToRid = {};
const idPattern = /name="(ID_[A-F0-9]+)".*?r:embed="(rId\d+)"/gs;
let match;
while ((match = idPattern.exec(cellImagesXml)) !== null) {
    dispimgToRid[match[1]] = match[2];
}
console.log(`Found ${Object.keys(dispimgToRid).length} DISPIMG entries`);

// Parse cellimages.xml.rels to get rId → image file
const relsXml = zip.readAsText('xl/_rels/cellimages.xml.rels');
const ridToFile = {};
const relPattern = /Id="(rId\d+)".*?Target="([^"]+)"/g;
while ((match = relPattern.exec(relsXml)) !== null) {
    if (match[2] !== 'NULL') {
        ridToFile[match[1]] = match[2].replace('media/', '');
    }
}
console.log(`Found ${Object.keys(ridToFile).length} image relationships`);

// Combine: DISPIMG ID → image file
const idToImg = {};
for (const [id, rid] of Object.entries(dispimgToRid)) {
    if (ridToFile[rid]) {
        idToImg[id] = ridToFile[rid];
    }
}
console.log(`Mapped ${Object.keys(idToImg).length} DISPIMG IDs to image files`);

// Step 2: Read the XLSX sheet data with DISPIMG formulas
console.log('\n=== Step 2: Reading pattern data ===');
const workbook = XLSX.readFile(XLSX_PATH);
const ws = workbook.Sheets['纹样数据库'];
const rawData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

// Parse patterns with DISPIMG IDs
const patterns = [];
for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    if (!row[0] && !row[1]) continue; // skip empty rows

    const id = row[0];
    const imgFormula = String(row[1] || '');
    const contentTag = String(row[2] || '');
    const formTag = String(row[3] || '');
    const craft = String(row[4] || '');
    const color = String(row[5] || '');
    const keywords = String(row[6] || '');

    // Extract DISPIMG ID from formula
    const dispMatch = imgFormula.match(/ID_([A-F0-9]+)/);
    const dispId = dispMatch ? 'ID_' + dispMatch[1] : null;
    const imgFile = dispId ? idToImg[dispId] : null;

    if (id || imgFile) {
        // Determine region from the data
        let region = '云贵';
        if (contentTag.includes('苗族') || craft.includes('蜡染')) region = '湘西';
        else if (craft.includes('蓝印花布') || craft.includes('夹缬')) region = '江苏';
        else if (contentTag.includes('自贡') || craft.includes('蜀')) region = '四川';
        else if (craft.includes('现代')) region = '浙江';

        patterns.push({
            id: id,
            imgFile: imgFile,
            name: contentTag ? contentTag.split('、').slice(0, 2).join('·') : '蓝染纹样',
            contentTag: contentTag,
            formTag: formTag,
            craft: craft,
            color: color,
            keywords: keywords,
            region: region,
            desc: `${craft} · ${color || '靛蓝'} · ${contentTag ? contentTag.split('、')[0] : '传统纹样'}`
        });
    }
}
console.log(`Parsed ${patterns.length} pattern entries`);

// Show some examples
patterns.slice(0, 5).forEach(p => {
    console.log(`  [${p.id}] ${p.name} | ${p.craft} | ${p.color} | img=${p.imgFile}`);
});

// Step 3: Copy images to destination folder
console.log('\n=== Step 3: Copying images ===');
if (!fs.existsSync(IMG_DST)) fs.mkdirSync(IMG_DST, { recursive: true });

// Copy all pattern images (image6.jpeg and onwards, and any webp/png)
let copiedCount = 0;
const validPatterns = patterns.filter(p => p.imgFile);
for (const p of validPatterns) {
    const src = path.join(IMG_SRC, p.imgFile);
    const dst = path.join(IMG_DST, p.imgFile);
    if (fs.existsSync(src) && !fs.existsSync(dst)) {
        fs.copyFileSync(src, dst);
        copiedCount++;
    }
}
console.log(`Copied ${copiedCount} images to ${IMG_DST}`);

// Also copy the chart sheet images (image2.png, image3.png, image4.png, image5.png) for possible use
const extraImages = ['image1.jpeg', 'image2.png', 'image3.png', 'image4.png', 'image5.png'];
for (const img of extraImages) {
    const src = path.join(IMG_SRC, img);
    const dst = path.join(IMG_DST, img);
    if (fs.existsSync(src) && !fs.existsSync(dst)) {
        fs.copyFileSync(src, dst);
    }
}

// Step 4: Generate the new DB_PATTERNS array and image CSS classes
console.log('\n=== Step 4: Generating HTML updates ===');

// Generate DB_PATTERNS array as JavaScript
const dbEntries = validPatterns.map(p => {
    const safeName = p.name.replace(/['"\\]/g, '');
    const safeCraft = p.craft.replace(/['"\\]/g, '');
    const safeRegion = p.region.replace(/['"\\]/g, '');
    const safeDesc = p.desc.replace(/['"\\]/g, '');
    const safeColor = (p.color || '').replace(/['"\\]/g, '');
    const safeKeywords = (p.keywords || '').replace(/['"\\]/g, '');
    return `{ name:'${safeName}', craft:'${safeCraft}', region:'${safeRegion}', desc:'${safeDesc}', color:'${safeColor}', img:'images/${p.imgFile}', keywords:'${safeKeywords}' }`;
});

console.log(`Generated ${dbEntries.length} DB entries`);

// Step 5: Update the HTML file
console.log('\n=== Step 5: Updating HTML ===');
let html = fs.readFileSync(HTML_SRC, 'utf8');

// 5a: Replace DB_PATTERNS array
const oldDbStart = html.indexOf('const DB_PATTERNS = [');
const oldDbEnd = html.indexOf('];', oldDbStart) + 2;
if (oldDbStart > 0 && oldDbEnd > oldDbStart) {
    const newDb = 'const DB_PATTERNS = [\n  ' + dbEntries.join(',\n  ') + '\n];';
    html = html.slice(0, oldDbStart) + newDb + html.slice(oldDbEnd);
    console.log('Replaced DB_PATTERNS array');
}

// 5b: Update DB_GRADIENTS - no longer needed, but keep for fallback
// We'll add a DB_IMAGES array and update renderDatabase to use images

// 5c: Update renderDatabase function to use actual images
const oldRenderFn = `grid.innerHTML = page.map((p, i) => \`
    <div class="bg-white card-shadow overflow-hidden group cursor-pointer hover:-translate-y-1 transition fade-up" style="animation-delay:\${i*0.05}s">
      <div class="h-48 bg-gradient-to-br \${DB_GRADIENTS[i%DB_GRADIENTS.length]} \${DB_PATTERNS_CSS[i%DB_PATTERNS_CSS.length]} relative flex items-center justify-center">
        <div class="absolute inset-0 opacity-0 group-hover:opacity-100 transition bg-black/20 flex items-center justify-center">
          <span class="text-white text-sm font-medium bg-white/20 px-4 py-2">查看详情</span>
        </div>
      </div>
      <div class="p-4">
        <h3 class="font-serif text-[16px] font-medium text-[#312E81] mb-1">\${p.name}</h3>
        <p class="text-[12px] text-[#64748B]">\${p.desc}</p>
        <div class="flex gap-2 mt-2">
          <span class="px-2 py-0.5 bg-[#EEF2FF] text-[11px] text-[#6366F1] font-medium">\${p.craft}</span>
          <span class="px-2 py-0.5 bg-[#EEF2FF] text-[11px] text-[#64748B]">\${p.region}</span>
        </div>
      </div>
    </div>
  \`).join('');`;

const newRenderFn = `grid.innerHTML = page.map((p, i) => {
    const imgSrc = p.img || '';
    const hasImg = imgSrc && imgSrc !== 'undefined';
    return \`
    <div class="bg-white card-shadow overflow-hidden group cursor-pointer hover:-translate-y-1 transition fade-up" style="animation-delay:\${i*0.05}s">
      <div class="h-48 relative flex items-center justify-center overflow-hidden">
        \${hasImg ? \`<img src="\${imgSrc}" alt="\${p.name}" class="w-full h-full object-cover img-zoom" loading="lazy" />\` : \`<div class="w-full h-full bg-gradient-to-br \${DB_GRADIENTS[i%DB_GRADIENTS.length]} \${DB_PATTERNS_CSS[i%DB_PATTERNS_CSS.length]}"></div>\`}
        <div class="absolute inset-0 opacity-0 group-hover:opacity-100 transition bg-black/20 flex items-center justify-center">
          <span class="text-white text-sm font-medium bg-white/20 px-4 py-2">查看详情</span>
        </div>
      </div>
      <div class="p-4">
        <h3 class="font-serif text-[16px] font-medium text-[#312E81] mb-1">\${p.name}</h3>
        <p class="text-[12px] text-[#64748B]">\${p.desc}</p>
        <div class="flex gap-2 mt-2">
          <span class="px-2 py-0.5 bg-[#EEF2FF] text-[11px] text-[#6366F1] font-medium">\${p.craft}</span>
          <span class="px-2 py-0.5 bg-[#EEF2FF] text-[11px] text-[#64748B]">\${p.region}</span>
        </div>
      </div>
    </div>
  \`}).join('');`;

if (html.includes(oldRenderFn)) {
    html = html.replace(oldRenderFn, newRenderFn);
    console.log('Updated renderDatabase function');
} else {
    console.log('WARNING: Could not find old renderDatabase function - trying partial match');
    // Try to find and replace just the template part
}

// 5d: Replace 每日推荐 (daily recommendations) section with real images
// Card 1: Use a 蓝印花布 with 冰裂纹 pattern
const featuredImg1 = validPatterns.find(p => p.name.includes('冰裂') || p.keywords.includes('冰裂纹'));
const featuredImg2 = validPatterns.find(p => p.name.includes('螺旋') || p.keywords.includes('螺旋纹'));
const featuredImg3 = validPatterns.find(p => p.name.includes('水波') || p.keywords.includes('水波纹'));

// Replace the three daily recommendation cards' gradient divs with img tags
// Card 1
if (featuredImg1) {
    const oldCard1Bg = `<div class="h-60 bg-gradient-to-br from-[#1e2143] via-[#3a3e6b] to-[#6366F1] pattern-crackle relative flex items-center justify-center">`;
    const newCard1Bg = `<div class="h-60 relative flex items-center justify-center overflow-hidden"><img src="images/${featuredImg1.imgFile}" alt="${featuredImg1.name}" class="w-full h-full object-cover img-zoom" loading="lazy" />`;
    html = html.replace(oldCard1Bg, newCard1Bg);
    console.log(`Card 1: ${featuredImg1.name} -> ${featuredImg1.imgFile}`);
}

if (featuredImg2) {
    const oldCard2Bg = `<div class="h-60 bg-gradient-to-br from-[#1a2e4d] via-[#344a73] to-[#5a6fb5] pattern-spiral relative flex items-center justify-center">`;
    const newCard2Bg = `<div class="h-60 relative flex items-center justify-center overflow-hidden"><img src="images/${featuredImg2.imgFile}" alt="${featuredImg2.name}" class="w-full h-full object-cover img-zoom" loading="lazy" />`;
    html = html.replace(oldCard2Bg, newCard2Bg);
    console.log(`Card 2: ${featuredImg2.name} -> ${featuredImg2.imgFile}`);
}

if (featuredImg3) {
    const oldCard3Bg = `<div class="h-60 bg-gradient-to-br from-[#152347] via-[#2a3657] to-[#4a5694] pattern-wave relative flex items-center justify-center">`;
    const newCard3Bg = `<div class="h-60 relative flex items-center justify-center overflow-hidden"><img src="images/${featuredImg3.imgFile}" alt="${featuredImg3.name}" class="w-full h-full object-cover img-zoom" loading="lazy" />`;
    html = html.replace(oldCard3Bg, newCard3Bg);
    console.log(`Card 3: ${featuredImg3.name} -> ${featuredImg3.imgFile}`);
}

// 5e: Replace 灵感广场 (inspiration square) cards with real images
// Find the 8 inspiration cards and give them real images
const inspImages = validPatterns.slice(4, 12); // Use patterns 5-12
if (inspImages.length >= 8) {
    const inspCardPattern = /<div class="h-\[(\d+)px\] bg-gradient-to-br ([^"]+) pattern-(\w+)"><\/div>/g;
    let inspIdx = 0;
    html = html.replace(/<div class="h-\[(\d+)px\] bg-gradient-to-br [^"]*"><\/div>/g, (match) => {
        if (inspIdx < 8) {
            const img = inspImages[inspIdx];
            const newDiv = `<div class="h-[${inspIdx % 2 === 0 ? '200px' : '150px'}] relative overflow-hidden"><img src="images/${img.imgFile}" alt="${img.name}" class="w-full h-full object-cover img-zoom" loading="lazy" /></div>`;
            inspIdx++;
            return newDiv;
        }
        return match;
    });
    console.log(`Replaced ${inspIdx} inspiration cards with real images`);
}

// 5f: Replace AI generation result cards with real images
const genImages = validPatterns.slice(8, 12);
let genIdx = 0;
html = html.replace(/<div class="h-\[(\d+)px\] bg-gradient-to-br ([^"]+) pattern-(\w+) relative"><div class="absolute inset-0 opacity-0 group-hover:opacity-100 transition bg-black\/10 flex items-center justify-center"><button class="bg-white\/90 text-\[\#312E81\] px-4 py-2 text-sm font-medium">查看大图<\/button><\/div><\/div>/g, (match) => {
    if (genIdx < 4) {
        const img = genImages[genIdx];
        const h = genIdx % 2 === 0 ? '300px' : '260px';
        const newDiv = `<div class="h-[${h}] relative overflow-hidden"><img src="images/${img.imgFile}" alt="${img.name}" class="w-full h-full object-cover img-zoom" loading="lazy" /><div class="absolute inset-0 opacity-0 group-hover:opacity-100 transition bg-black/10 flex items-center justify-center"><button class="bg-white/90 text-[#312E81] px-4 py-2 text-sm font-medium">查看大图</button></div></div>`;
        genIdx++;
        return newDiv;
    }
    return match;
});
console.log(`Replaced ${genIdx} generation result cards`);

// 5g: Replace exhibition gallery images
const exhibitImages = validPatterns.slice(12, 16);
let exhIdx = 0;
// Exhibition featured gallery - large images with gradient overlays
html = html.replace(/<div class="h-\[(\d+)px\] bg-gradient-to-br ([^"]+) pattern-(\w+) relative flex items-center justify-center">/g, (match) => {
    if (exhIdx < 4) {
        const img = exhibitImages[exhIdx];
        const h = exhIdx < 2 ? '400px' : '300px';
        const newDiv = `<div class="h-[${h}] relative flex items-center justify-center overflow-hidden"><img src="images/${img.imgFile}" alt="${img.name}" class="w-full h-full object-cover img-zoom" loading="lazy" />`;
        exhIdx++;
        return newDiv;
    }
    return match;
});
console.log(`Replaced ${exhIdx} exhibition gallery images`);

// 5h: Replace exhibition story cards
const storyImages = validPatterns.slice(16, 19);
let storyIdx = 0;
html = html.replace(/<div class="h-48 bg-gradient-to-br [^"]+ pattern-\w+ mb-4 rounded flex items-center justify-center">/g, (match) => {
    if (storyIdx < 3) {
        const img = storyImages[storyIdx];
        const newDiv = `<div class="h-48 relative mb-4 rounded overflow-hidden flex items-center justify-center"><img src="images/${img.imgFile}" alt="${img.name}" class="w-full h-full object-cover img-zoom" loading="lazy" />`;
        storyIdx++;
        return newDiv;
    }
    return match;
});
console.log(`Replaced ${storyIdx} exhibition story cards`);

// 5i: Replace daily recommendation thumbnails in exhibition
const dailyExImg = validPatterns.slice(19, 21);
let dailyIdx = 0;
html = html.replace(/<div class="w-36 h-36 shrink-0 bg-gradient-to-br [^"]+ pattern-\w+ rounded flex items-center justify-center">/g, (match) => {
    if (dailyIdx < 2) {
        const img = dailyExImg[dailyIdx];
        const newDiv = `<div class="w-36 h-36 shrink-0 relative rounded overflow-hidden flex items-center justify-center"><img src="images/${img.imgFile}" alt="${img.name}" class="w-full h-full object-cover img-zoom" loading="lazy" />`;
        dailyIdx++;
        return newDiv;
    }
    return match;
});
console.log(`Replaced ${dailyIdx} daily exhibition thumbnails`);

// 5j: Replace knowledge section decorative backgrounds with real pattern images
const knowImages = validPatterns.slice(21, 24);
let knowIdx = 0;
html = html.replace(/<div class="bg-gradient-to-br from-\[\#1e2143\] via-\[\#3a3e6b\] to-\[\#6366F1\] pattern-crackle p-8 flex flex-col justify-center items-center text-center min-h-\[400px\]">/g, (match) => {
    if (knowIdx < 1) {
        const img = knowImages[0];
        return `<div class="relative p-8 flex flex-col justify-center items-center text-center min-h-[400px] overflow-hidden"><img src="images/${img.imgFile}" alt="${img.name}" class="absolute inset-0 w-full h-full object-cover" loading="lazy" /><div class="absolute inset-0 bg-gradient-to-br from-[#1e2143]/80 via-[#3a3e6b]/70 to-[#6366F1]/80"></div>`;
    }
    knowIdx++;
    return match;
});
console.log(`Replaced ${knowIdx} knowledge section backgrounds`);

// 5k: Also update the generatePatterns function to use real images
const oldGenFn = `html += \`<div class="bg-white card-shadow overflow-hidden group animate-[fadeIn_0.5s_ease-out_\${i*0.1}s]"><div class="h-[\${heights[i]}] bg-gradient-to-br \${gradients[i]} \${patterns[i]} relative"><div class="absolute inset-0 opacity-0 group-hover:opacity-100 transition bg-black/10 flex items-center justify-center"><button class="bg-white/90 text-[#312E81] px-4 py-2 text-sm font-medium">查看大图</button></div></div><div class="p-5"><h3 class="font-serif text-[16px] font-medium text-[#312E81] mb-3">\${patternName} \${variations[i]}</h3><div class="flex gap-3"><button class="text-[12px] text-[#64748B] hover:text-[#6366F1] transition">收藏</button><button class="text-[12px] text-[#64748B] hover:text-[#6366F1] transition">分享</button>\${hasLocal ? '<button class="text-[12px] text-[#6366F1] font-medium hover:opacity-80">局部变化</button>' : ''}</div></div></div>\`;`;

if (html.includes(oldGenFn)) {
    const newGenFn = `const genImgs = validPatterns.slice((i*3) % validPatterns.length, (i*3) % validPatterns.length + 1);
    const genImg = genImgs[0];
    html += \`<div class="bg-white card-shadow overflow-hidden group animate-[fadeIn_0.5s_ease-out_\${i*0.1}s]"><div class="h-[\${heights[i]}] relative overflow-hidden"><img src="\${genImg ? 'images/' + genImg.imgFile : ''}" alt="\${patternName}" class="w-full h-full object-cover img-zoom" loading="lazy" /><div class="absolute inset-0 opacity-0 group-hover:opacity-100 transition bg-black/10 flex items-center justify-center"><button class="bg-white/90 text-[#312E81] px-4 py-2 text-sm font-medium">查看大图</button></div></div><div class="p-5"><h3 class="font-serif text-[16px] font-medium text-[#312E81] mb-3">\${patternName} \${variations[i]}</h3><div class="flex gap-3"><button class="text-[12px] text-[#64748B] hover:text-[#6366F1] transition">收藏</button><button class="text-[12px] text-[#64748B] hover:text-[#6366F1] transition">分享</button>\${hasLocal ? '<button class="text-[12px] text-[#6366F1] font-medium hover:opacity-80">局部变化</button>' : ''}</div></div></div>\`;`;
    html = html.replace(oldGenFn, newGenFn);
    console.log('Updated generatePatterns function');
}

// 5l: Add validPatterns reference for generatePatterns to use
// Insert the reference at the beginning of the script section
html = html.replace('// ===== 数据库数据 =====', '// ===== 数据库数据 =====\nconst validPatterns = DB_PATTERNS.filter(p => p.img);');

// Write the updated HTML
fs.writeFileSync(HTML_SRC, html, 'utf8');
console.log('\n=== Done! ===');
console.log(`Updated HTML written to: ${HTML_SRC}`);
console.log(`Images folder: ${IMG_DST}`);
console.log(`Total pattern images available: ${validPatterns.length}`);

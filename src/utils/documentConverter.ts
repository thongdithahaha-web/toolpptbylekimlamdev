import JSZip from 'jszip';
import { jsPDF } from 'jspdf';
import * as docx from 'docx';
import pptxgen from 'pptxgenjs';
import * as XLSX from 'xlsx';
import { parsePPTXFile } from './pptxParser.ts';

const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType } = docx;

export interface ConversionResult {
  blob: Blob;
  filename: string;
  mimeType: string;
  itemCount: number;
  format: string;
}

// ============================================================================
// 1. PPTX CONVERSIONS
// ============================================================================

/** PPTX -> PDF */
export async function convertPptxToPdf(file: File): Promise<ConversionResult> {
  const { slides } = await parsePPTXFile(file);
  const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  const pageWidth = 297;
  const pageHeight = 210;

  for (let i = 0; i < slides.length; i++) {
    if (i > 0) pdf.addPage('a4', 'landscape');
    const slide = slides[i];

    pdf.setFillColor(248, 250, 252);
    pdf.rect(0, 0, pageWidth, pageHeight, 'F');

    pdf.setDrawColor(226, 232, 240);
    pdf.rect(10, 10, pageWidth - 20, pageHeight - 20);

    pdf.setFillColor(225, 29, 72);
    pdf.rect(10, 10, pageWidth - 20, 14, 'F');

    pdf.setTextColor(255, 255, 255);
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(11);
    pdf.text(`toolgopfilepptbylekimlam • Slide ${i + 1} / ${slides.length}`, 16, 19.5);

    pdf.setTextColor(15, 23, 42);
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(20);
    const safeTitle = slide.title || `Trang thuyết trình #${slide.slideNumber}`;
    pdf.text(safeTitle, 20, 42, { maxWidth: pageWidth - 40 });

    pdf.setDrawColor(244, 63, 94);
    pdf.setLineWidth(0.8);
    pdf.line(20, 48, 80, 48);

    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(13);
    pdf.setTextColor(51, 65, 85);

    let currentY = 62;
    if (slide.previewText && slide.previewText.length > 0) {
      for (const text of slide.previewText) {
        if (currentY > pageHeight - 30) break;
        pdf.setFillColor(225, 29, 72);
        pdf.circle(23, currentY - 1.5, 1.2, 'F');
        const splitText = pdf.splitTextToSize(text, pageWidth - 55);
        pdf.text(splitText, 28, currentY);
        currentY += splitText.length * 8 + 4;
      }
    }

    pdf.setFontSize(9);
    pdf.setTextColor(100, 116, 139);
    pdf.text(`Tài liệu: ${file.name}`, 20, pageHeight - 16);
    pdf.text(`Trang ${i + 1}`, pageWidth - 35, pageHeight - 16);
  }

  const blob = pdf.output('blob');
  const baseName = file.name.replace(/\.[^/.]+$/, '');
  return { blob, filename: `${baseName}.pdf`, mimeType: 'application/pdf', itemCount: slides.length, format: 'PDF' };
}

/** PPTX -> Handout PDF (3 slides / page with note lines) */
export async function convertPptxToHandoutPdf(file: File): Promise<ConversionResult> {
  const { slides } = await parsePPTXFile(file);
  const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const pageWidth = 210;
  const pageHeight = 297;
  const slidesPerPage = 3;

  for (let i = 0; i < slides.length; i++) {
    const pageIndex = Math.floor(i / slidesPerPage);
    const slotIndex = i % slidesPerPage;

    if (slotIndex === 0 && i > 0) pdf.addPage('a4', 'portrait');

    if (slotIndex === 0) {
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(12);
      pdf.setTextColor(225, 29, 72);
      pdf.text(`BÁO CÁO TÓM TẮT THUYẾT TRÌNH (HANDOUT)`, 20, 16);
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(9);
      pdf.setTextColor(100, 116, 139);
      pdf.text(`Tệp: ${file.name} • Trang ${pageIndex + 1}`, 20, 22);
      pdf.line(20, 25, pageWidth - 20, 25);
    }

    const slotY = 32 + slotIndex * 82;
    const slide = slides[i];

    // Left mini slide box (16:9)
    pdf.setFillColor(248, 250, 252);
    pdf.setDrawColor(203, 213, 225);
    pdf.rect(20, slotY, 80, 45, 'FD');

    pdf.setFillColor(225, 29, 72);
    pdf.rect(20, slotY, 80, 6, 'F');
    pdf.setTextColor(255, 255, 255);
    pdf.setFontSize(7);
    pdf.setFont('helvetica', 'bold');
    pdf.text(`Slide #${i + 1}`, 23, slotY + 4.5);

    pdf.setTextColor(15, 23, 42);
    pdf.setFontSize(9);
    const safeTitle = (slide.title || `Trang ${i + 1}`).slice(0, 35);
    pdf.text(safeTitle, 23, slotY + 12);

    pdf.setFontSize(7);
    pdf.setTextColor(71, 85, 105);
    if (slide.previewText && slide.previewText[0]) {
      const splitPreview = pdf.splitTextToSize(slide.previewText[0], 74);
      pdf.text(splitPreview.slice(0, 3), 23, slotY + 18);
    }

    // Right note lines
    pdf.setDrawColor(226, 232, 240);
    pdf.setFontSize(8);
    pdf.setTextColor(148, 163, 184);
    pdf.text('Ghi chú:', 110, slotY + 6);
    for (let line = 0; line < 5; line++) {
      const lineY = slotY + 14 + line * 7;
      pdf.line(110, lineY, pageWidth - 20, lineY);
    }
  }

  const blob = pdf.output('blob');
  const baseName = file.name.replace(/\.[^/.]+$/, '');
  return { blob, filename: `${baseName}_Handout.pdf`, mimeType: 'application/pdf', itemCount: slides.length, format: 'PDF Handout' };
}

/** PPTX -> Word (DOCX) */
export async function convertPptxToDocx(file: File): Promise<ConversionResult> {
  const { slides } = await parsePPTXFile(file);
  const baseName = file.name.replace(/\.[^/.]+$/, '');

  const children: docx.Paragraph[] = [
    new Paragraph({ text: `BÁO CÁO NỘI DUNG BÀI THUYẾT TRÌNH`, heading: HeadingLevel.TITLE, spacing: { after: 200 } }),
    new Paragraph({ children: [new TextRun({ text: `Tên tệp: `, bold: true }), new TextRun({ text: file.name })], spacing: { after: 100 } }),
    new Paragraph({ children: [new TextRun({ text: `Tổng số slide: `, bold: true }), new TextRun({ text: `${slides.length} trang` })], spacing: { after: 100 } }),
    new Paragraph({ children: [new TextRun({ text: `Công cụ xuất: `, bold: true }), new TextRun({ text: `toolgopfilepptbylekimlam` })], spacing: { after: 400 } }),
  ];

  for (let i = 0; i < slides.length; i++) {
    const s = slides[i];
    children.push(
      new Paragraph({ text: `Slide ${i + 1}: ${s.title || `Trang thuyết trình #${i + 1}`}`, heading: HeadingLevel.HEADING_1, spacing: { before: 250, after: 120 } })
    );
    if (s.previewText && s.previewText.length > 0) {
      for (const line of s.previewText) {
        children.push(new Paragraph({ text: `• ${line}`, spacing: { after: 60 } }));
      }
    }
  }

  const doc = new Document({ sections: [{ properties: {}, children }] });
  const blob = await Packer.toBlob(doc);
  return { blob, filename: `${baseName}_Bao_Cao.docx`, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', itemCount: slides.length, format: 'Word (DOCX)' };
}

/** PPTX -> TXT / Markdown */
export async function convertPptxToTxt(file: File, isMarkdown = false): Promise<ConversionResult> {
  const { slides } = await parsePPTXFile(file);
  const baseName = file.name.replace(/\.[^/.]+$/, '');

  let text = `# ${file.name}\nTổng số slide: ${slides.length}\nXuất bởi: toolgopfilepptbylekimlam • Lê Kim Lâm\n==================================================\n\n`;

  for (let i = 0; i < slides.length; i++) {
    const s = slides[i];
    text += `## Slide ${i + 1}: ${s.title || `Trang #${i + 1}`}\n`;
    if (s.previewText && s.previewText.length > 0) {
      for (const line of s.previewText) text += `- ${line}\n`;
    }
    text += `\n`;
  }

  const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
  const ext = isMarkdown ? 'md' : 'txt';
  return { blob, filename: `${baseName}.${ext}`, mimeType: 'text/plain', itemCount: slides.length, format: isMarkdown ? 'Markdown (MD)' : 'Văn bản (TXT)' };
}

/** PPTX -> Extract all original images as ZIP */
export async function convertPptxExtractMedia(file: File): Promise<ConversionResult> {
  const arrayBuffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);
  const outZip = new JSZip();

  const mediaFiles = Object.keys(zip.files).filter(k => k.startsWith('ppt/media/') && !zip.files[k].dir);
  if (mediaFiles.length === 0) throw new Error('Bài thuyết trình này không chứa tệp hình ảnh nào.');

  for (const mPath of mediaFiles) {
    const mName = mPath.replace('ppt/media/', '');
    const data = await zip.file(mPath)!.async('arraybuffer');
    outZip.file(mName, data);
  }

  const blob = await outZip.generateAsync({ type: 'blob' });
  const baseName = file.name.replace(/\.[^/.]+$/, '');
  return { blob, filename: `${baseName}_Hinh_Anh.zip`, mimeType: 'application/zip', itemCount: mediaFiles.length, format: 'Tệp nén hình ảnh (.zip)' };
}

/** PPTX -> Standalone HTML Slide Deck */
export async function convertPptxToHtml(file: File): Promise<ConversionResult> {
  const { slides } = await parsePPTXFile(file);
  const baseName = file.name.replace(/\.[^/.]+$/, '');
  const slidesJson = JSON.stringify(slides);

  const html = `<!DOCTYPE html>
<html lang="vi"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>${file.name}</title>
<style>
* { box-sizing: border-box; margin:0; padding:0; }
body { font-family: system-ui, -apple-system, sans-serif; background: #0f172a; color: #f8fafc; height: 100vh; display: flex; flex-direction: column; overflow: hidden; }
header { background: #1e293b; padding: 12px 24px; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #334155; }
main { flex: 1; display: flex; align-items: center; justify-content: center; padding: 24px; }
.card { background: white; color: #0f172a; width: 100%; max-width: 960px; aspect-ratio: 16/9; border-radius: 20px; box-shadow: 0 25px 50px -12px rgba(0,0,0,0.5); padding: 48px; display: flex; flex-direction: column; justify-content: space-between; }
.title { font-size: 28px; font-weight: 800; color: #0f172a; margin-bottom: 24px; }
.body { flex: 1; font-size: 18px; color: #334155; line-height: 1.8; }
.body li { margin-bottom: 12px; display: flex; gap: 10px; }
.body li::before { content: "•"; color: #e11d48; font-weight: bold; font-size: 24px; }
footer { background: #1e293b; padding: 16px 24px; display: flex; justify-content: space-between; align-items: center; }
.btn { background: #e11d48; color: white; border: none; padding: 10px 20px; border-radius: 10px; font-weight: 700; cursor: pointer; }
.btn:disabled { opacity: 0.3; cursor: not-allowed; }
</style></head>
<body>
<header><div>toolgopfilepptbylekimlam</div><div>${file.name}</div></header>
<main><div class="card"><div class="title" id="t"></div><div class="body" id="b"></div><div style="font-size:12px;color:#64748b;display:flex;justify-content:space-between;border-top:1px solid #e2e8f0;padding-top:12px;"><span>toolgopfilepptbylekimlam</span><span id="idx"></span></div></div></main>
<footer><button class="btn" id="p" onclick="nav(-1)">◀ Trước</button><span id="c"></span><button class="btn" id="n" onclick="nav(1)">Kế Tiếp ▶</button></footer>
<script>
const slides = ${slidesJson}; let cur = 0;
function render() {
  const s = slides[cur];
  document.getElementById('t').innerText = s.title || ('Slide ' + (cur + 1));
  document.getElementById('b').innerHTML = (s.previewText && s.previewText.length) ? '<ul>' + s.previewText.map(t => '<li><span>' + t + '</span></li>').join('') + '</ul>' : '<p style="color:#94a3b8;">(Nội dung đồ hoạ)</p>';
  document.getElementById('idx').innerText = (cur+1) + ' / ' + slides.length;
  document.getElementById('c').innerText = 'Slide ' + (cur+1) + ' / ' + slides.length;
  document.getElementById('p').disabled = cur === 0;
  document.getElementById('n').disabled = cur === slides.length - 1;
}
function nav(d) { if (cur+d >= 0 && cur+d < slides.length) { cur += d; render(); } }
window.addEventListener('keydown', e => { if (e.key === 'ArrowRight' || e.key === ' ') nav(1); if (e.key === 'ArrowLeft') nav(-1); });
render();
</script></body></html>`;

  const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
  return { blob, filename: `${baseName}_Trinh_Chieu.html`, mimeType: 'text/html', itemCount: slides.length, format: 'Trình chiếu Web (HTML)' };
}

/** PPTX -> JSON Data */
export async function convertPptxToJson(file: File): Promise<ConversionResult> {
  const { slides } = await parsePPTXFile(file);
  const data = {
    sourceFile: file.name,
    fileSize: file.size,
    exportedAt: new Date().toISOString(),
    tool: 'toolgopfilepptbylekimlam',
    author: 'Lê Kim Lâm',
    totalSlides: slides.length,
    slides: slides.map(s => ({ slideNumber: s.slideNumber, title: s.title, contents: s.previewText, hasImages: s.hasImages })),
  };
  const jsonStr = JSON.stringify(data, null, 2);
  const blob = new Blob([jsonStr], { type: 'application/json' });
  const baseName = file.name.replace(/\.[^/.]+$/, '');
  return { blob, filename: `${baseName}.json`, mimeType: 'application/json', itemCount: slides.length, format: 'Dữ liệu JSON' };
}

/** PPTX -> CSV Outline */
export async function convertPptxToCsv(file: File): Promise<ConversionResult> {
  const { slides } = await parsePPTXFile(file);
  let csv = 'Số trang,Tiêu đề slide,Nội dung chi tiết\n';
  for (const s of slides) {
    const t = `"${(s.title || '').replace(/"/g, '""')}"`;
    const c = `"${(s.previewText ? s.previewText.join(' | ') : '').replace(/"/g, '""')}"`;
    csv += `${s.slideNumber},${t},${c}\n`;
  }
  const blob = new Blob([new Uint8Array([0xef, 0xbb, 0xbf]), csv], { type: 'text/csv;charset=utf-8' });
  const baseName = file.name.replace(/\.[^/.]+$/, '');
  return { blob, filename: `${baseName}.csv`, mimeType: 'text/csv', itemCount: slides.length, format: 'Bảng tính CSV' };
}

// ============================================================================
// 2. WORD (DOCX) CONVERSIONS
// ============================================================================

/** Word (DOCX) -> PowerPoint (PPTX) */
export async function convertDocxToPptx(file: File): Promise<ConversionResult> {
  const arrayBuffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);
  const docXml = await zip.file('word/document.xml')?.async('text');
  if (!docXml) throw new Error('Tệp Word không hợp lệ.');

  const paragraphs: { text: string; isHeading: boolean }[] = [];
  for (const match of docXml.matchAll(/<w:p[^>]*>([\s\S]*?)<\/w:p>/g)) {
    const pContent = match[1];
    const isHeading = /<w:pStyle[^>]*w:val="Heading\d"/i.test(pContent) || /<w:b\/>/i.test(pContent);
    const textMatches = Array.from(pContent.matchAll(/<w:t[^>]*>(.*?)<\/w:t>/g)).map(m => m[1]);
    const fullText = textMatches.join('').trim();
    if (fullText.length > 0) paragraphs.push({ text: fullText, isHeading });
  }

  const ppt = new pptxgen();
  ppt.layout = 'LAYOUT_WIDE';

  const sCover = ppt.addSlide();
  sCover.background = { color: '1E293B' };
  sCover.addText(file.name.replace(/\.[^/.]+$/, ''), { x: 1, y: 2, w: 11, fontSize: 32, bold: true, color: 'FFFFFF' });
  sCover.addText('Tạo tự động từ Word • toolgopfilepptbylekimlam', { x: 1, y: 3.5, w: 11, fontSize: 18, color: '94A3B8' });

  let currentTitle = 'Nội dung';
  let currentBullets: string[] = [];
  let slideCount = 1;

  const flushSlide = () => {
    if (currentBullets.length > 0 || currentTitle !== 'Nội dung') {
      const s = ppt.addSlide();
      s.background = { color: 'F8FAFC' };
      s.addText(currentTitle, { x: 0.8, y: 0.8, w: 11.5, fontSize: 24, bold: true, color: '0F172A' });
      const bulletText = currentBullets.map(b => `• ${b}`).join('\n\n');
      s.addText(bulletText || '(Nội dung tài liệu)', { x: 0.8, y: 2.0, w: 11.5, h: 4.5, fontSize: 16, color: '334155' });
      slideCount++;
      currentBullets = [];
    }
  };

  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    if (p.isHeading && i > 0) {
      flushSlide();
      currentTitle = p.text;
    } else if (p.isHeading && i === 0) {
      currentTitle = p.text;
    } else {
      currentBullets.push(p.text);
      if (currentBullets.length >= 5) {
        flushSlide();
        currentTitle = `Tiếp tục (${slideCount + 1})`;
      }
    }
  }
  flushSlide();

  const pptxArrayBuffer = (await ppt.write({ outputType: 'arraybuffer' })) as ArrayBuffer;
  const baseName = file.name.replace(/\.[^/.]+$/, '');
  return { blob: new Blob([pptxArrayBuffer]), filename: `${baseName}_Slides.pptx`, mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', itemCount: slideCount, format: 'PowerPoint (PPTX)' };
}

/** Word (DOCX) -> PDF */
export async function convertDocxToPdf(file: File): Promise<ConversionResult> {
  const arrayBuffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);
  const docXml = await zip.file('word/document.xml')?.async('text');
  if (!docXml) throw new Error('Tệp Word không hợp lệ.');

  const paragraphs: { text: string; isHeading: boolean }[] = [];
  for (const match of docXml.matchAll(/<w:p[^>]*>([\s\S]*?)<\/w:p>/g)) {
    const pContent = match[1];
    const isHeading = /<w:pStyle[^>]*w:val="Heading\d"/i.test(pContent) || /<w:b\/>/i.test(pContent);
    const textMatches = Array.from(pContent.matchAll(/<w:t[^>]*>(.*?)<\/w:t>/g)).map(m => m[1]);
    const fullText = textMatches.join('').trim();
    if (fullText.length > 0) paragraphs.push({ text: fullText, isHeading });
  }

  const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const pageWidth = 210;
  const pageHeight = 297;
  let currentY = 25;

  pdf.setFont('helvetica', 'bold');
  pdf.setFontSize(18);
  pdf.setTextColor(15, 23, 42);
  pdf.text(file.name.replace(/\.[^/.]+$/, ''), 20, currentY);
  currentY += 12;

  pdf.setFontSize(9);
  pdf.setTextColor(100, 116, 139);
  pdf.text('toolgopfilepptbylekimlam • Lê Kim Lâm', 20, currentY);
  currentY += 10;
  pdf.setDrawColor(226, 232, 240);
  pdf.line(20, currentY, pageWidth - 20, currentY);
  currentY += 12;

  for (const p of paragraphs) {
    if (currentY > pageHeight - 25) {
      pdf.addPage();
      currentY = 25;
    }
    if (p.isHeading) {
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(14);
      pdf.setTextColor(225, 29, 72);
      currentY += 4;
      const lines = pdf.splitTextToSize(p.text, pageWidth - 40);
      pdf.text(lines, 20, currentY);
      currentY += lines.length * 7 + 4;
    } else {
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(11);
      pdf.setTextColor(51, 65, 85);
      const lines = pdf.splitTextToSize(p.text, pageWidth - 40);
      pdf.text(lines, 20, currentY);
      currentY += lines.length * 6 + 3;
    }
  }

  const blob = pdf.output('blob');
  const baseName = file.name.replace(/\.[^/.]+$/, '');
  return { blob, filename: `${baseName}.pdf`, mimeType: 'application/pdf', itemCount: paragraphs.length, format: 'PDF' };
}

/** Word (DOCX) -> Extract Images as ZIP */
export async function convertDocxExtractMedia(file: File): Promise<ConversionResult> {
  const arrayBuffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);
  const outZip = new JSZip();

  const mediaFiles = Object.keys(zip.files).filter(k => k.startsWith('word/media/') && !zip.files[k].dir);
  if (mediaFiles.length === 0) throw new Error('Tệp Word này không chứa tệp hình ảnh nào.');

  for (const mPath of mediaFiles) {
    const mName = mPath.replace('word/media/', '');
    const data = await zip.file(mPath)!.async('arraybuffer');
    outZip.file(mName, data);
  }

  const blob = await outZip.generateAsync({ type: 'blob' });
  const baseName = file.name.replace(/\.[^/.]+$/, '');
  return { blob, filename: `${baseName}_Word_Hinh_Anh.zip`, mimeType: 'application/zip', itemCount: mediaFiles.length, format: 'Tệp nén hình ảnh Word (.zip)' };
}

// ============================================================================
// 3. TEXT / MARKDOWN CONVERSIONS
// ============================================================================

/** TXT / MD -> PowerPoint (PPTX) */
export async function convertTxtToPptx(file: File): Promise<ConversionResult> {
  const textContent = await file.text();
  const lines = textContent.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
  if (lines.length === 0) throw new Error('Tệp văn bản trống.');

  const ppt = new pptxgen();
  ppt.layout = 'LAYOUT_WIDE';

  const sCover = ppt.addSlide();
  sCover.background = { color: '0F172A' };
  sCover.addText(file.name.replace(/\.[^/.]+$/, ''), { x: 1, y: 2, w: 11, fontSize: 32, bold: true, color: 'FFFFFF' });
  sCover.addText('Tạo từ văn bản • toolgopfilepptbylekimlam', { x: 1, y: 3.5, w: 11, fontSize: 18, color: '94A3B8' });

  let currentTitle = 'Nội dung';
  let currentBullets: string[] = [];
  let slideCount = 1;

  const flushSlide = () => {
    if (currentBullets.length > 0 || currentTitle !== 'Nội dung') {
      const s = ppt.addSlide();
      s.background = { color: 'FFFFFF' };
      s.addText(currentTitle, { x: 0.8, y: 0.8, w: 11.5, fontSize: 24, bold: true, color: '0F172A' });
      const bulletText = currentBullets.map(b => `• ${b.replace(/^[-*•#]+\s*/, '')}`).join('\n\n');
      s.addText(bulletText, { x: 0.8, y: 2.0, w: 11.5, h: 4.5, fontSize: 16, color: '334155' });
      slideCount++;
      currentBullets = [];
    }
  };

  for (const line of lines) {
    const isHeader = line.startsWith('#') || line.startsWith('===') || line.startsWith('---');
    if (isHeader) {
      flushSlide();
      currentTitle = line.replace(/^[#=\-\s]+/, '');
    } else {
      currentBullets.push(line);
      if (currentBullets.length >= 5) {
        flushSlide();
        currentTitle = `Tiếp tục (${slideCount + 1})`;
      }
    }
  }
  flushSlide();

  const pptxArrayBuffer = (await ppt.write({ outputType: 'arraybuffer' })) as ArrayBuffer;
  const baseName = file.name.replace(/\.[^/.]+$/, '');
  return { blob: new Blob([pptxArrayBuffer]), filename: `${baseName}_Slides.pptx`, mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', itemCount: slideCount, format: 'PowerPoint (PPTX)' };
}

/** TXT / MD -> Word (DOCX) */
export async function convertTxtToDocx(file: File): Promise<ConversionResult> {
  const textContent = await file.text();
  const lines = textContent.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
  const baseName = file.name.replace(/\.[^/.]+$/, '');

  const children: docx.Paragraph[] = [
    new Paragraph({ text: file.name.replace(/\.[^/.]+$/, ''), heading: HeadingLevel.TITLE, spacing: { after: 200 } }),
  ];

  for (const line of lines) {
    if (line.startsWith('#')) {
      children.push(new Paragraph({ text: line.replace(/^#+\s*/, ''), heading: HeadingLevel.HEADING_1, spacing: { before: 200, after: 100 } }));
    } else {
      children.push(new Paragraph({ text: line, spacing: { after: 80 } }));
    }
  }

  const doc = new Document({ sections: [{ properties: {}, children }] });
  const blob = await Packer.toBlob(doc);
  return { blob, filename: `${baseName}.docx`, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', itemCount: lines.length, format: 'Word (DOCX)' };
}

/** TXT / MD -> PDF */
export async function convertTxtToPdf(file: File): Promise<ConversionResult> {
  const textContent = await file.text();
  const lines = textContent.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
  const baseName = file.name.replace(/\.[^/.]+$/, '');

  const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  let currentY = 25;
  pdf.setFont('helvetica', 'bold');
  pdf.setFontSize(16);
  pdf.text(baseName, 20, currentY);
  currentY += 12;

  pdf.setFont('helvetica', 'normal');
  pdf.setFontSize(10);
  pdf.setTextColor(51, 65, 85);

  for (const line of lines) {
    if (currentY > 280) { pdf.addPage(); currentY = 20; }
    const splitLines = pdf.splitTextToSize(line, 170);
    pdf.text(splitLines, 20, currentY);
    currentY += splitLines.length * 6 + 2;
  }

  const blob = pdf.output('blob');
  return { blob, filename: `${baseName}.pdf`, mimeType: 'application/pdf', itemCount: lines.length, format: 'PDF' };
}

// ============================================================================
// 4. EXCEL (XLSX / CSV) CONVERSIONS
// ============================================================================

/** Excel / CSV -> PowerPoint (PPTX Table Slides) */
export async function convertExcelToPptx(file: File): Promise<ConversionResult> {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const ppt = new pptxgen();
  ppt.layout = 'LAYOUT_WIDE';

  let totalSlideCount = 0;
  for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json<string[]>(worksheet, { header: 1 });

    if (jsonData.length === 0) continue;

    // Create a slide for each sheet or table slice
    const s = ppt.addSlide();
    s.background = { color: 'F8FAFC' };
    s.addText(`Bảng tính: ${sheetName}`, { x: 0.8, y: 0.8, w: 11.5, fontSize: 22, bold: true, color: '0F172A' });

    // Limit to top 10 rows and 6 columns for clean display
    const rowsSlice: pptxgen.TableRow[] = jsonData.slice(0, 10).map((r) =>
      (r || []).slice(0, 6).map((val) => ({ text: String(val ?? '') }))
    );
    if (rowsSlice.length > 0) {
      s.addTable(rowsSlice, {
        x: 0.8,
        y: 1.8,
        w: 11.5,
        fontSize: 12,
        color: '334155',
        fill: { color: 'FFFFFF' },
        border: { pt: 1, color: 'CBD5E1' },
      });
    }
    totalSlideCount++;
  }

  if (totalSlideCount === 0) throw new Error('Không tìm thấy dữ liệu bảng trong tệp Excel.');

  const pptxArrayBuffer = (await ppt.write({ outputType: 'arraybuffer' })) as ArrayBuffer;
  const baseName = file.name.replace(/\.[^/.]+$/, '');
  return { blob: new Blob([pptxArrayBuffer]), filename: `${baseName}_Bang_Bieu.pptx`, mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', itemCount: totalSlideCount, format: 'PowerPoint (PPTX)' };
}

/** Excel / CSV -> PDF Table */
export async function convertExcelToPdf(file: File): Promise<ConversionResult> {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  let isFirst = true;

  for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json<string[]>(worksheet, { header: 1 });
    if (jsonData.length === 0) continue;

    if (!isFirst) pdf.addPage('a4', 'landscape');
    isFirst = false;

    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(16);
    pdf.setTextColor(15, 23, 42);
    pdf.text(`Báo cáo Bảng tính: ${sheetName}`, 20, 20);

    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(9);
    pdf.setTextColor(100, 116, 139);
    pdf.text(`Nguồn: ${file.name} • toolgopfilepptbylekimlam`, 20, 26);

    let currentY = 36;
    for (let r = 0; r < Math.min(25, jsonData.length); r++) {
      const row = jsonData[r] || [];
      const rowText = row.slice(0, 7).map(c => String(c ?? '')).join('  |  ');
      pdf.setFontSize(8);
      pdf.setTextColor(r === 0 ? 225 : 51, r === 0 ? 29 : 65, r === 0 ? 72 : 85);
      if (r === 0) pdf.setFont('helvetica', 'bold');
      else pdf.setFont('helvetica', 'normal');

      pdf.text(rowText.slice(0, 140), 20, currentY);
      currentY += 8;
    }
  }

  const blob = pdf.output('blob');
  const baseName = file.name.replace(/\.[^/.]+$/, '');
  return { blob, filename: `${baseName}.pdf`, mimeType: 'application/pdf', itemCount: workbook.SheetNames.length, format: 'PDF Bảng tính' };
}

// ============================================================================
// 5. IMAGES CONVERSIONS
// ============================================================================

/** Image(s) -> PowerPoint (PPTX Slides) */
export async function convertImagesToPptx(file: File): Promise<ConversionResult> {
  const ppt = new pptxgen();
  ppt.layout = 'LAYOUT_WIDE';

  // Read file as data URL
  const dataUrl = await new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });

  const s = ppt.addSlide();
  s.background = { color: '0F172A' };
  s.addImage({ data: dataUrl, x: 0.5, y: 0.5, w: 12.3, h: 6.5, sizing: { type: 'contain', w: 12.3, h: 6.5 } });

  const pptxArrayBuffer = (await ppt.write({ outputType: 'arraybuffer' })) as ArrayBuffer;
  const baseName = file.name.replace(/\.[^/.]+$/, '');
  return { blob: new Blob([pptxArrayBuffer]), filename: `${baseName}_Slide.pptx`, mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', itemCount: 1, format: 'PowerPoint (PPTX)' };
}

/** Image -> PDF */
export async function convertImagesToPdf(file: File): Promise<ConversionResult> {
  const dataUrl = await new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });

  const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  pdf.addImage(dataUrl, 'JPEG', 15, 15, 180, 240, undefined, 'FAST');

  const blob = pdf.output('blob');
  const baseName = file.name.replace(/\.[^/.]+$/, '');
  return { blob, filename: `${baseName}.pdf`, mimeType: 'application/pdf', itemCount: 1, format: 'PDF' };
}

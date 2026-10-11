import JSZip from 'jszip';
import { jsPDF } from 'jspdf';
import * as docx from 'docx';
import pptxgen from 'pptxgenjs';
import { parsePPTXFile } from './pptxParser.ts';

const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, BorderStyle } = docx;

export interface ConversionResult {
  blob: Blob;
  filename: string;
  mimeType: string;
  itemCount: number;
  format: string;
}

/**
 * 1. Convert PowerPoint (.pptx) to PDF (.pdf)
 */
export async function convertPptxToPdf(file: File): Promise<ConversionResult> {
  const { slides } = await parsePPTXFile(file);
  const pdf = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4', // 297mm x 210mm
  });

  const pageWidth = 297;
  const pageHeight = 210;

  for (let i = 0; i < slides.length; i++) {
    if (i > 0) {
      pdf.addPage('a4', 'landscape');
    }

    const slide = slides[i];

    // Background gradient/fill
    pdf.setFillColor(248, 250, 252); // slate-50
    pdf.rect(0, 0, pageWidth, pageHeight, 'F');

    // Slide border frame
    pdf.setDrawColor(226, 232, 240); // slate-200
    pdf.setLineWidth(1);
    pdf.rect(10, 10, pageWidth - 20, pageHeight - 20);

    // Top banner bar
    pdf.setFillColor(225, 29, 72); // rose-600
    pdf.rect(10, 10, pageWidth - 20, 14, 'F');

    // Top Title in banner
    pdf.setTextColor(255, 255, 255);
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(11);
    pdf.text(`toolgopfilepptbylekimlam • Slide ${i + 1} / ${slides.length}`, 16, 19.5);

    // Slide Heading
    pdf.setTextColor(15, 23, 42); // slate-900
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(20);

    const safeTitle = slide.title || `Trang thuyet trinh #${slide.slideNumber}`;
    pdf.text(safeTitle, 20, 42, { maxWidth: pageWidth - 40 });

    // Decorative separator line
    pdf.setDrawColor(244, 63, 94); // rose-500
    pdf.setLineWidth(0.8);
    pdf.line(20, 48, 80, 48);

    // Slide bullet points / content
    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(13);
    pdf.setTextColor(51, 65, 85); // slate-700

    let currentY = 62;
    if (slide.previewText && slide.previewText.length > 0) {
      for (const text of slide.previewText) {
        if (currentY > pageHeight - 30) break;
        // Bullet circle
        pdf.setFillColor(225, 29, 72);
        pdf.circle(23, currentY - 1.5, 1.2, 'F');
        // Bullet text
        const splitText = pdf.splitTextToSize(text, pageWidth - 55);
        pdf.text(splitText, 28, currentY);
        currentY += splitText.length * 8 + 4;
      }
    } else {
      pdf.setFontSize(11);
      pdf.setTextColor(148, 163, 184);
      pdf.text('(Noi dung do hoa hoac tieu de khong chua van ban phu)', 20, currentY);
    }

    // Bottom Footer
    pdf.setFontSize(9);
    pdf.setTextColor(100, 116, 139);
    pdf.text(`Tai lieu: ${file.name}`, 20, pageHeight - 16);
    pdf.text(`Trang ${i + 1}`, pageWidth - 35, pageHeight - 16);
  }

  const pdfBlob = pdf.output('blob');
  const baseName = file.name.replace(/\.[^/.]+$/, '');

  return {
    blob: pdfBlob,
    filename: `${baseName}.pdf`,
    mimeType: 'application/pdf',
    itemCount: slides.length,
    format: 'PDF',
  };
}

/**
 * 2. Convert PowerPoint (.pptx) to Word (.docx)
 */
export async function convertPptxToDocx(file: File): Promise<ConversionResult> {
  const { slides } = await parsePPTXFile(file);
  const baseName = file.name.replace(/\.[^/.]+$/, '');

  const children: docx.Paragraph[] = [];

  // Cover Page Title
  children.push(
    new Paragraph({
      text: `BÁO CÁO NỘI DUNG BÀI THUYẾT TRÌNH`,
      heading: HeadingLevel.TITLE,
      spacing: { after: 200 },
    }),
    new Paragraph({
      children: [
        new TextRun({ text: `Tên tệp gốc: `, bold: true }),
        new TextRun({ text: file.name }),
      ],
      spacing: { after: 100 },
    }),
    new Paragraph({
      children: [
        new TextRun({ text: `Tổng số slide: `, bold: true }),
        new TextRun({ text: `${slides.length} trang` }),
      ],
      spacing: { after: 100 },
    }),
    new Paragraph({
      children: [
        new TextRun({ text: `Công cụ chuyển đổi: `, bold: true }),
        new TextRun({ text: `toolgopfilepptbylekimlam • Bản quyền Lê Kim Lâm` }),
      ],
      spacing: { after: 400 },
    })
  );

  // Each slide as a section
  for (let i = 0; i < slides.length; i++) {
    const s = slides[i];
    children.push(
      new Paragraph({
        text: `Slide ${i + 1}: ${s.title || `Trang thuyết trình #${i + 1}`}`,
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 300, after: 150 },
      })
    );

    if (s.previewText && s.previewText.length > 0) {
      for (const line of s.previewText) {
        children.push(
          new Paragraph({
            text: `• ${line}`,
            spacing: { after: 80 },
          })
        );
      }
    } else {
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: `(Slide không chứa đoạn văn bản phụ hoặc là bố cục đồ hoạ)`,
              italics: true,
              color: '888888',
            }),
          ],
          spacing: { after: 150 },
        })
      );
    }
  }

  const doc = new Document({
    sections: [
      {
        properties: {},
        children,
      },
    ],
  });

  const docxBlob = await Packer.toBlob(doc);

  return {
    blob: docxBlob,
    filename: `${baseName}_Bao_Cao.docx`,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    itemCount: slides.length,
    format: 'Word (DOCX)',
  };
}

/**
 * 3. Convert Word (.docx) to PowerPoint (.pptx)
 */
export async function convertDocxToPptx(file: File): Promise<ConversionResult> {
  const arrayBuffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);
  const docXml = await zip.file('word/document.xml')?.async('text');

  if (!docXml) {
    throw new Error('Tệp Word không hợp lệ hoặc thiếu cấu trúc word/document.xml');
  }

  // Extract paragraphs: <w:p>...</w:p>
  const pRegex = /<w:p[^>]*>([\s\S]*?)<\/w:p>/g;
  const paragraphs: { text: string; isHeading: boolean }[] = [];

  for (const match of docXml.matchAll(pRegex)) {
    const pContent = match[1];
    // Check if paragraph is a heading: Heading1, Heading2, or bold
    const isHeading = /<w:pStyle[^>]*w:val="Heading\d"/i.test(pContent) || /<w:b\/>/i.test(pContent);

    // Extract text runs: <w:t>...</w:t>
    const textMatches = Array.from(pContent.matchAll(/<w:t[^>]*>(.*?)<\/w:t>/g)).map((m) => m[1]);
    const fullText = textMatches.join('').trim();

    if (fullText.length > 0) {
      paragraphs.push({ text: fullText, isHeading });
    }
  }

  if (paragraphs.length === 0) {
    throw new Error('Không tìm thấy nội dung văn bản trong tệp Word đã chọn.');
  }

  const ppt = new pptxgen();
  ppt.layout = 'LAYOUT_WIDE';

  // Title Slide
  const sCover = ppt.addSlide();
  sCover.background = { color: '1E293B' };
  sCover.addText(file.name.replace(/\.[^/.]+$/, ''), {
    x: 1,
    y: 2,
    w: 11,
    fontSize: 32,
    bold: true,
    color: 'FFFFFF',
  });
  sCover.addText('Tạo tự động từ tài liệu Word • toolgopfilepptbylekimlam', {
    x: 1,
    y: 3.5,
    w: 11,
    fontSize: 18,
    color: '94A3B8',
  });

  // Group into slides by headings or every 4-5 paragraphs
  let currentTitle = 'Nội dung';
  let currentBullets: string[] = [];
  let slideCount = 1;

  const flushSlide = () => {
    if (currentBullets.length > 0 || currentTitle !== 'Nội dung') {
      const s = ppt.addSlide();
      s.background = { color: 'F8FAFC' };
      s.addText(currentTitle, {
        x: 0.8,
        y: 0.8,
        w: 11.5,
        fontSize: 24,
        bold: true,
        color: '0F172A',
      });

      const bulletText = currentBullets.map((b) => `• ${b}`).join('\n\n');
      s.addText(bulletText || '(Chi tiết trong tài liệu Word)', {
        x: 0.8,
        y: 2.0,
        w: 11.5,
        h: 4.5,
        fontSize: 16,
        color: '334155',
      });
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
        currentTitle = `Tiếp theo: ${paragraphs[Math.max(0, i - 4)].text.slice(0, 30)}...`;
      }
    }
  }
  flushSlide();

  const pptxArrayBuffer = (await ppt.write({ outputType: 'arraybuffer' })) as ArrayBuffer;
  const baseName = file.name.replace(/\.[^/.]+$/, '');

  return {
    blob: new Blob([pptxArrayBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    }),
    filename: `${baseName}_Converted.pptx`,
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    itemCount: slideCount,
    format: 'PowerPoint (PPTX)',
  };
}

/**
 * 4. Convert Text / Markdown (.txt / .md) to PowerPoint (.pptx)
 */
export async function convertTxtToPptx(file: File): Promise<ConversionResult> {
  const textContent = await file.text();
  const lines = textContent.split(/\r?\n/).map((l) => l.trim()).filter((l) => l.length > 0);

  if (lines.length === 0) {
    throw new Error('Tệp văn bản trống.');
  }

  const ppt = new pptxgen();
  ppt.layout = 'LAYOUT_WIDE';

  // Title Slide
  const sCover = ppt.addSlide();
  sCover.background = { color: '0F172A' };
  sCover.addText(file.name.replace(/\.[^/.]+$/, ''), {
    x: 1,
    y: 2,
    w: 11,
    fontSize: 32,
    bold: true,
    color: 'FFFFFF',
  });
  sCover.addText('Chuyển đổi từ văn bản thuần • toolgopfilepptbylekimlam', {
    x: 1,
    y: 3.5,
    w: 11,
    fontSize: 18,
    color: '94A3B8',
  });

  let currentTitle = 'Nội dung';
  let currentBullets: string[] = [];
  let slideCount = 1;

  const flushSlide = () => {
    if (currentBullets.length > 0 || currentTitle !== 'Nội dung') {
      const s = ppt.addSlide();
      s.background = { color: 'FFFFFF' };
      s.addText(currentTitle, {
        x: 0.8,
        y: 0.8,
        w: 11.5,
        fontSize: 24,
        bold: true,
        color: '0F172A',
      });
      const bulletText = currentBullets.map((b) => `• ${b.replace(/^[-*•#]+\s*/, '')}`).join('\n\n');
      s.addText(bulletText, {
        x: 0.8,
        y: 2.0,
        w: 11.5,
        h: 4.5,
        fontSize: 16,
        color: '334155',
      });
      slideCount++;
      currentBullets = [];
    }
  };

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    // Check if line is title or header (# Title or starts with number 1. or capitalized)
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

  return {
    blob: new Blob([pptxArrayBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    }),
    filename: `${baseName}_Slides.pptx`,
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    itemCount: slideCount,
    format: 'PowerPoint (PPTX)',
  };
}

/**
 * 5. Convert PowerPoint (.pptx) to Text / Markdown (.txt)
 */
export async function convertPptxToTxt(file: File): Promise<ConversionResult> {
  const { slides } = await parsePPTXFile(file);
  const baseName = file.name.replace(/\.[^/.]+$/, '');

  let resultText = `# BÁO CÁO NỘI DUNG: ${file.name}\n`;
  resultText += `Tổng số slide: ${slides.length}\n`;
  resultText += `Công cụ xuất: toolgopfilepptbylekimlam (Bản quyền Lê Kim Lâm)\n`;
  resultText += `--------------------------------------------------\n\n`;

  for (let i = 0; i < slides.length; i++) {
    const s = slides[i];
    resultText += `## SLIDE ${i + 1}: ${s.title || `Trang #${i + 1}`}\n`;
    if (s.previewText && s.previewText.length > 0) {
      for (const line of s.previewText) {
        resultText += `- ${line}\n`;
      }
    } else {
      resultText += `(Không có văn bản phụ)\n`;
    }
    resultText += `\n`;
  }

  const blob = new Blob([resultText], { type: 'text/plain;charset=utf-8' });

  return {
    blob,
    filename: `${baseName}_Van_Ban.txt`,
    mimeType: 'text/plain',
    itemCount: slides.length,
    format: 'Văn bản (TXT)',
  };
}

/**
 * 6. Convert PowerPoint (.pptx) to Standalone Offline HTML Slide Deck (.html)
 */
export async function convertPptxToHtml(file: File): Promise<ConversionResult> {
  const { slides } = await parsePPTXFile(file);
  const baseName = file.name.replace(/\.[^/.]+$/, '');

  const slidesJson = JSON.stringify(slides);

  const htmlContent = `<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${file.name} - Trình Chiếu Web</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #0f172a; color: #f8fafc; height: 100vh; display: flex; flex-direction: column; overflow: hidden; }
    header { background: #1e293b; padding: 12px 24px; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #334155; }
    .brand { font-size: 13px; font-weight: 700; color: #fb7185; }
    .title { font-size: 14px; font-weight: 600; color: #94a3b8; }
    main { flex: 1; display: flex; align-items: center; justify-content: center; padding: 24px; position: relative; }
    .slide-card { background: #ffffff; color: #0f172a; width: 100%; max-width: 960px; aspect-ratio: 16/9; border-radius: 20px; box-shadow: 0 25px 50px -12px rgba(0,0,0,0.5); padding: 48px; display: flex; flex-direction: column; justify-content: space-between; position: relative; overflow: hidden; animation: fadeIn 0.25s ease; }
    @keyframes fadeIn { from { opacity: 0; transform: scale(0.98); } to { opacity: 1; transform: scale(1); } }
    .slide-title { font-size: 28px; font-weight: 800; color: #0f172a; margin-bottom: 24px; line-height: 1.3; }
    .slide-body { flex: 1; font-size: 18px; color: #334155; line-height: 1.8; }
    .slide-body ul { list-style: none; }
    .slide-body li { margin-bottom: 12px; display: flex; align-items: flex-start; gap: 10px; }
    .slide-body li::before { content: "•"; color: #e11d48; font-weight: bold; font-size: 24px; line-height: 1; }
    .slide-footer { display: flex; justify-content: space-between; align-items: center; border-top: 1px solid #e2e8f0; padding-top: 16px; font-size: 12px; color: #64748b; font-weight: 600; }
    footer { background: #1e293b; padding: 16px 24px; display: flex; justify-content: space-between; align-items: center; border-top: 1px solid #334155; }
    .btn { background: #e11d48; color: white; border: none; padding: 10px 20px; border-radius: 10px; font-weight: 700; cursor: pointer; transition: 0.15s; font-size: 14px; }
    .btn:hover { background: #be123c; }
    .btn:disabled { opacity: 0.3; cursor: not-allowed; }
    .counter { font-weight: 700; font-size: 14px; color: #cbd5e1; }
  </style>
</head>
<body>
  <header>
    <div class="brand">toolgopfilepptbylekimlam</div>
    <div class="title">${file.name}</div>
    <button class="btn" style="background:#334155; padding:6px 14px; font-size:12px;" onclick="toggleFullScreen()">Toàn màn hình</button>
  </header>
  <main>
    <div class="slide-card" id="slideCard">
      <div>
        <div class="slide-title" id="slideTitle">Tiêu đề slide</div>
        <div class="slide-body" id="slideBody"></div>
      </div>
      <div class="slide-footer">
        <span>Tạo bởi toolgopfilepptbylekimlam • Lê Kim Lâm</span>
        <span id="slideIndex">Slide 1 / ${slides.length}</span>
      </div>
    </div>
  </main>
  <footer>
    <button class="btn" id="prevBtn" onclick="prevSlide()">◀ Slide Trước</button>
    <div class="counter" id="footerCounter">Trang 1 / ${slides.length}</div>
    <button class="btn" id="nextBtn" onclick="nextSlide()">Slide Kế Tiếp ▶</button>
  </footer>

  <script>
    const slides = ${slidesJson};
    let currentIndex = 0;

    function renderSlide() {
      const s = slides[currentIndex];
      document.getElementById('slideTitle').innerText = s.title || ('Slide ' + (currentIndex + 1));
      
      const bodyEl = document.getElementById('slideBody');
      if (s.previewText && s.previewText.length > 0) {
        bodyEl.innerHTML = '<ul>' + s.previewText.map(t => '<li><span>' + t + '</span></li>').join('') + '</ul>';
      } else {
        bodyEl.innerHTML = '<p style="color:#94a3b8; font-style:italic;">(Nội dung đồ hoạ hoặc không có văn bản phụ)</p>';
      }

      document.getElementById('slideIndex').innerText = 'Slide ' + (currentIndex + 1) + ' / ' + slides.length;
      document.getElementById('footerCounter').innerText = 'Trang ' + (currentIndex + 1) + ' / ' + slides.length;

      document.getElementById('prevBtn').disabled = currentIndex === 0;
      document.getElementById('nextBtn').disabled = currentIndex === slides.length - 1;
    }

    function prevSlide() {
      if (currentIndex > 0) { currentIndex--; renderSlide(); }
    }
    function nextSlide() {
      if (currentIndex < slides.length - 1) { currentIndex++; renderSlide(); }
    }
    function toggleFullScreen() {
      if (!document.fullscreenElement) { document.documentElement.requestFullscreen(); }
      else { document.exitFullscreen(); }
    }

    window.addEventListener('keydown', (e) => {
      if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') nextSlide();
      if (e.key === 'ArrowLeft' || e.key === 'PageUp') prevSlide();
    });

    renderSlide();
  </script>
</body>
</html>`;

  const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });

  return {
    blob,
    filename: `${baseName}_Trinh_Chieu.html`,
    mimeType: 'text/html',
    itemCount: slides.length,
    format: 'Trình chiếu Web (HTML)',
  };
}

/**
 * 7. Convert Word (.docx) to PDF (.pdf)
 */
export async function convertDocxToPdf(file: File): Promise<ConversionResult> {
  const arrayBuffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);
  const docXml = await zip.file('word/document.xml')?.async('text');

  if (!docXml) {
    throw new Error('Tệp Word không hợp lệ.');
  }

  // Extract paragraphs
  const pRegex = /<w:p[^>]*>([\s\S]*?)<\/w:p>/g;
  const paragraphs: { text: string; isHeading: boolean }[] = [];

  for (const match of docXml.matchAll(pRegex)) {
    const pContent = match[1];
    const isHeading = /<w:pStyle[^>]*w:val="Heading\d"/i.test(pContent) || /<w:b\/>/i.test(pContent);
    const textMatches = Array.from(pContent.matchAll(/<w:t[^>]*>(.*?)<\/w:t>/g)).map((m) => m[1]);
    const fullText = textMatches.join('').trim();
    if (fullText.length > 0) {
      paragraphs.push({ text: fullText, isHeading });
    }
  }

  const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const pageWidth = 210;
  const pageHeight = 297;
  let currentY = 25;

  // Title
  pdf.setFont('helvetica', 'bold');
  pdf.setFontSize(18);
  pdf.setTextColor(15, 23, 42);
  pdf.text(file.name.replace(/\.[^/.]+$/, ''), 20, currentY);
  currentY += 12;

  pdf.setFontSize(9);
  pdf.setTextColor(100, 116, 139);
  pdf.text('Chuyen doi boi toolgopfilepptbylekimlam • Le Kim Lam', 20, currentY);
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
      pdf.setTextColor(225, 29, 72); // rose-600
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

  const pdfBlob = pdf.output('blob');
  const baseName = file.name.replace(/\.[^/.]+$/, '');

  return {
    blob: pdfBlob,
    filename: `${baseName}.pdf`,
    mimeType: 'application/pdf',
    itemCount: paragraphs.length,
    format: 'PDF',
  };
}

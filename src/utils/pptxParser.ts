import JSZip from 'jszip';
import { SlideInfo } from '../types.ts';

export async function parsePPTXFile(file: File): Promise<{ slideCount: number; slides: SlideInfo[] }> {
  // Check for legacy binary .ppt file
  const headerSlice = await file.slice(0, 8).arrayBuffer();
  const headerBytes = new Uint8Array(headerSlice);
  // OLE Compound Document signature for .ppt: D0 CF 11 E0 A1 B1 1A E1
  if (
    headerBytes[0] === 0xd0 &&
    headerBytes[1] === 0xcf &&
    headerBytes[2] === 0x11 &&
    headerBytes[3] === 0xe0
  ) {
    throw new Error(
      `Tệp "${file.name}" ở định dạng .ppt cũ (Office 97-2003). Vui lòng mở bằng PowerPoint hoặc Google Slides rồi "Lưu thành" .pptx để gộp.`
    );
  }

  let zip: JSZip;
  try {
    const arrayBuffer = await file.arrayBuffer();
    zip = await JSZip.loadAsync(arrayBuffer);
  } catch (err: any) {
    throw new Error(`Không thể đọc tệp "${file.name}". Vui lòng đảm bảo tệp là định dạng PowerPoint .pptx hợp lệ.`);
  }

  // Find all slide XML files
  const slideEntries = Object.keys(zip.files).filter((fileName) => {
    return /^ppt\/slides\/slide\d+\.xml$/i.test(fileName);
  });

  if (slideEntries.length === 0) {
    // Check if presentation.xml exists
    const presFile = zip.file('ppt/presentation.xml');
    if (!presFile) {
      throw new Error(`Tệp "${file.name}" không có cấu trúc PowerPoint hợp lệ.`);
    }
  }

  // Natural sort slide filenames: slide1.xml, slide2.xml, ..., slide10.xml
  slideEntries.sort((a, b) => {
    const numA = parseInt(a.match(/slide(\d+)\.xml/i)?.[1] || '0', 10);
    const numB = parseInt(b.match(/slide(\d+)\.xml/i)?.[1] || '0', 10);
    return numA - numB;
  });

  const slides: SlideInfo[] = [];

  for (let i = 0; i < slideEntries.length; i++) {
    const slidePath = slideEntries[i];
    const slideNumber = parseInt(slidePath.match(/slide(\d+)\.xml/i)?.[1] || `${i + 1}`, 10);

    let title = `Trang ${i + 1}`;
    const previewText: string[] = [];
    let hasImages = false;

    // Check relationships for images
    const relsPath = `ppt/slides/_rels/slide${slideNumber}.xml.rels`;
    const relsFile = zip.file(relsPath);
    if (relsFile) {
      try {
        const relsXml = await relsFile.async('text');
        if (relsXml.includes('/image') || relsXml.includes('relationships/image')) {
          hasImages = true;
        }
      } catch {
        // ignore
      }
    }

    // Parse slide XML
    const slideFile = zip.file(slidePath);
    if (slideFile) {
      try {
        const xml = await slideFile.async('text');

        // Extract text elements: <a:t>Content</a:t>
        const textMatches = Array.from(xml.matchAll(/<a:t[^>]*>(.*?)<\/a:t>/gs)).map((m) =>
          m[1].replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').trim()
        ).filter(t => t.length > 0);

        if (textMatches.length > 0) {
          title = textMatches[0].length > 60 ? textMatches[0].slice(0, 60) + '...' : textMatches[0];
          for (let k = 1; k < Math.min(textMatches.length, 6); k++) {
            if (textMatches[k] && textMatches[k] !== title) {
              previewText.push(textMatches[k]);
            }
          }
        }
      } catch (err) {
        console.warn(`Could not parse text for slide ${slideNumber}:`, err);
      }
    }

    slides.push({
      slideNumber: i + 1, // 1-indexed relative slide index in this presentation
      title: title || `Slide ${i + 1}`,
      previewText,
      hasImages,
    });
  }

  const slideCount = Math.max(slides.length, 1);

  return {
    slideCount,
    slides,
  };
}

export function formatBytes(bytes: number, decimals = 1): string {
  if (bytes === 0) return '0 B';
  const k = 1024;
  const dm = decimals < 0 ? 0 : decimals;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return `${parseFloat((bytes / Math.pow(k, i)).toFixed(dm))} ${sizes[i]}`;
}

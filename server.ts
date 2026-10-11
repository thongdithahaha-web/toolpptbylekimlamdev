import express from 'express';
import { createServer as createViteServer } from 'vite';
import multer from 'multer';
import path from 'path';
import fs from 'fs/promises';
import fsSync from 'fs';
import os from 'os';
import AutomizerModule from 'pptx-automizer';
import pptxgen from 'pptxgenjs';
import JSZip from 'jszip';

const Automizer = (AutomizerModule as any).Automizer || (AutomizerModule as any).default || AutomizerModule;

const app = express();
const PORT = process.env.PORT ? parseInt(process.env.PORT) : 3000;

// Multer for chunked uploads (max 15MB per single chunk)
const chunkUpload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 15 * 1024 * 1024,
    fieldSize: 10 * 1024 * 1024,
  },
});

// Multer fallback for direct small batch merge (max 25MB total)
const legacyUpload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 25 * 1024 * 1024,
    files: 50,
  },
});

app.use(express.json({ limit: '20mb' }));
app.use(express.urlencoded({ extended: true, limit: '20mb' }));

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Periodic session cleanup (removes sessions older than 30 mins)
setInterval(async () => {
  try {
    const tmpDir = os.tmpdir();
    const entries = await fs.readdir(tmpDir);
    const now = Date.now();
    for (const entry of entries) {
      if (entry.startsWith('pptx-session-') || entry.startsWith('pptx-merge-')) {
        const full = path.join(tmpDir, entry);
        const stats = await fs.stat(full).catch(() => null);
        if (stats && now - stats.mtimeMs > 30 * 60 * 1000) {
          await fs.rm(full, { recursive: true, force: true }).catch(() => {});
        }
      }
    }
  } catch (err) {
    console.warn('Cleanup error:', err);
  }
}, 5 * 60 * 1000);

// Endpoint to generate 3 sample PPTX files for instant testing
app.get('/api/samples', async (req, res) => {
  try {
    const samples: { name: string; base64: string; slideCount: number }[] = [];

    // Sample 1: Báo cáo Tổng quan (Overview)
    const p1 = new pptxgen();
    p1.layout = 'LAYOUT_WIDE';
    
    // Slide 1.1
    const s1_1 = p1.addSlide();
    s1_1.background = { color: '1E293B' }; // Dark Slate
    s1_1.addText('BÁO CÁO TỔNG QUAN DOANH NGHIỆP', {
      x: 1,
      y: 2,
      w: 11,
      fontSize: 32,
      bold: true,
      color: 'FFFFFF',
      fontFace: 'Arial',
    });
    s1_1.addText('Phần 1: Tình hình hoạt động & Chỉ số then chốt', {
      x: 1,
      y: 3.2,
      w: 11,
      fontSize: 20,
      color: '94A3B8',
      fontFace: 'Arial',
    });
    s1_1.addText('Tài liệu lưu hành nội bộ • 2026', {
      x: 1,
      y: 6.2,
      w: 11,
      fontSize: 14,
      color: '64748B',
      fontFace: 'Arial',
    });

    // Slide 1.2
    const s1_2 = p1.addSlide();
    s1_2.background = { color: 'F8FAFC' };
    s1_2.addText('KẾT QUẢ TĂNG TRƯỞNG NĂM 2025', {
      x: 0.8,
      y: 0.8,
      w: 11,
      fontSize: 24,
      bold: true,
      color: '0F172A',
      fontFace: 'Arial',
    });
    s1_2.addText('• Tăng trưởng doanh thu toàn công ty đạt +28% so với cùng kỳ', {
      x: 1,
      y: 2.0,
      w: 10,
      fontSize: 18,
      color: '334155',
      fontFace: 'Arial',
    });
    s1_2.addText('• Mở rộng thị phần sang 3 khu vực kinh tế trọng điểm mới', {
      x: 1,
      y: 2.8,
      w: 10,
      fontSize: 18,
      color: '334155',
      fontFace: 'Arial',
    });
    s1_2.addText('• Chỉ số hài lòng khách hàng (CSAT) đạt 94.6%', {
      x: 1,
      y: 3.6,
      w: 10,
      fontSize: 18,
      color: '334155',
      fontFace: 'Arial',
    });

    const buf1 = await p1.write({ outputType: 'nodebuffer' });
    samples.push({
      name: '01_Bao_Cao_Tong_Quan.pptx',
      base64: (buf1 as Buffer).toString('base64'),
      slideCount: 2,
    });

    // Sample 2: Kế hoạch Chiến lược (Strategy)
    const p2 = new pptxgen();
    p2.layout = 'LAYOUT_WIDE';
    
    // Slide 2.1
    const s2_1 = p2.addSlide();
    s2_1.background = { color: '0284C7' }; // Ocean Blue
    s2_1.addText('KẾ HOẠCH CHIẾN LƯỢC QUÝ MỚI', {
      x: 1,
      y: 2,
      w: 11,
      fontSize: 32,
      bold: true,
      color: 'FFFFFF',
      fontFace: 'Arial',
    });
    s2_1.addText('Phần 2: Mục tiêu trọng tâm & Kế hoạch ngân sách', {
      x: 1,
      y: 3.2,
      w: 11,
      fontSize: 20,
      color: 'E0F2FE',
      fontFace: 'Arial',
    });

    // Slide 2.2
    const s2_2 = p2.addSlide();
    s2_2.background = { color: 'F0F9FF' };
    s2_2.addText('3 TRỤ CỘT CHIẾN LƯỢC ƯU TIÊN', {
      x: 0.8,
      y: 0.8,
      w: 11,
      fontSize: 24,
      bold: true,
      color: '0369A1',
      fontFace: 'Arial',
    });
    s2_2.addText('1. Chuyển đổi số toàn diện quy trình vận hành và báo cáo tự động', {
      x: 1,
      y: 2.0,
      w: 10,
      fontSize: 18,
      color: '0F172A',
      fontFace: 'Arial',
    });
    s2_2.addText('2. Nâng cấp năng lực đội ngũ và hệ thống hạ tầng đám mây', {
      x: 1,
      y: 2.8,
      w: 10,
      fontSize: 18,
      color: '0F172A',
      fontFace: 'Arial',
    });
    s2_2.addText('3. Tối ưu hoá chi phí sản xuất, gia tăng biên lợi nhuận ròng', {
      x: 1,
      y: 3.6,
      w: 10,
      fontSize: 18,
      color: '0F172A',
      fontFace: 'Arial',
    });

    const buf2 = await p2.write({ outputType: 'nodebuffer' });
    samples.push({
      name: '02_Ke_Hoach_Chien_Luoc.pptx',
      base64: (buf2 as Buffer).toString('base64'),
      slideCount: 2,
    });

    // Sample 3: Kết luận & Kiến nghị (Conclusion)
    const p3 = new pptxgen();
    p3.layout = 'LAYOUT_WIDE';
    
    // Slide 3.1
    const s3_1 = p3.addSlide();
    s3_1.background = { color: '059669' }; // Emerald
    s3_1.addText('KẾT LUẬN & KIẾN NGHỊ THỰC HIỆN', {
      x: 1,
      y: 2,
      w: 11,
      fontSize: 32,
      bold: true,
      color: 'FFFFFF',
      fontFace: 'Arial',
    });
    s3_1.addText('Phần 3: Đề xuất phê duyệt & Lộ trình phân công', {
      x: 1,
      y: 3.2,
      w: 11,
      fontSize: 20,
      color: 'D1FAE5',
      fontFace: 'Arial',
    });

    // Slide 3.2
    const s3_2 = p3.addSlide();
    s3_2.background = { color: 'ECFDF5' };
    s3_2.addText('LỘ TRÌNH TRIỂN KHAI VÀ BÀN GIAO', {
      x: 0.8,
      y: 0.8,
      w: 11,
      fontSize: 24,
      bold: true,
      color: '047857',
      fontFace: 'Arial',
    });
    s3_2.addText('• Tuần 1 - 2: Hoàn thiện phê duyệt kinh phí và phân bổ nhân sự', {
      x: 1,
      y: 2.0,
      w: 10,
      fontSize: 18,
      color: '064E3B',
      fontFace: 'Arial',
    });
    s3_2.addText('• Tuần 3 - 6: Triển khai giai đoạn thử nghiệm (Pilot phase)', {
      x: 1,
      y: 2.8,
      w: 10,
      fontSize: 18,
      color: '064E3B',
      fontFace: 'Arial',
    });
    s3_2.addText('• Tuần 7+: Đánh giá kết quả và mở rộng trên toàn hệ thống', {
      x: 1,
      y: 3.6,
      w: 10,
      fontSize: 18,
      color: '064E3B',
      fontFace: 'Arial',
    });

    const buf3 = await p3.write({ outputType: 'nodebuffer' });
    samples.push({
      name: '03_Ket_Luan_Kien_Nghi.pptx',
      base64: (buf3 as Buffer).toString('base64'),
      slideCount: 2,
    });

    res.json({ success: true, samples });
  } catch (err: any) {
    console.error('Error generating sample files:', err);
    res.status(500).json({ error: 'Không thể tạo file mẫu: ' + err.message });
  }
});

// Endpoint: Upload chunk (each request is small ~2-5MB, preventing HTTP 413)
app.post('/api/upload-chunk', chunkUpload.single('chunk'), async (req, res) => {
  try {
    const { sessionId, fileIndex, chunkIndex, totalChunks, fileName } = req.body;
    if (!sessionId || fileIndex === undefined || chunkIndex === undefined || !req.file) {
      return res.status(400).json({ error: 'Thiếu thông tin phân mảnh tệp' });
    }

    const sessionDir = path.join(os.tmpdir(), `pptx-session-${sessionId}`);
    await fs.mkdir(sessionDir, { recursive: true });

    const partPath = path.join(sessionDir, `file_${fileIndex}.part`);
    await fs.appendFile(partPath, req.file.buffer);

    const isLastChunk = parseInt(chunkIndex) === parseInt(totalChunks) - 1;
    if (isLastChunk) {
      const finalFilePath = path.join(sessionDir, `input_${fileIndex}.pptx`);
      // If final file already existed (retry), remove it first
      if (fsSync.existsSync(finalFilePath)) {
        await fs.unlink(finalFilePath);
      }
      await fs.rename(partPath, finalFilePath);
      await fs.writeFile(
        path.join(sessionDir, `file_${fileIndex}.meta.json`),
        JSON.stringify({ fileName, fileIndex: parseInt(fileIndex) })
      );
    }

    res.json({ success: true, chunkIndex: parseInt(chunkIndex), isLastChunk });
  } catch (err: any) {
    console.error('Error in /api/upload-chunk:', err);
    res.status(500).json({ error: 'Lỗi tải phân mảnh: ' + err.message });
  }
});

// Endpoint: Merge session files in order
app.post('/api/merge-session', async (req, res) => {
  const { sessionId, order, slideSelections, outputName } = req.body;
  if (!sessionId) {
    return res.status(400).json({ error: 'Thiếu mã phiên làm việc (sessionId)' });
  }

  const sessionDir = path.join(os.tmpdir(), `pptx-session-${sessionId}`);
  if (!fsSync.existsSync(sessionDir)) {
    return res.status(404).json({ error: 'Phiên làm việc không tồn tại hoặc đã hết hạn.' });
  }

  try {
    const orderIndices: number[] = Array.isArray(order) ? order : [];
    const selections: Record<string, number[]> = slideSelections || {};

    let outputFilename = (outputName || 'Merged_Presentation.pptx').trim();
    if (!outputFilename.toLowerCase().endsWith('.pptx')) {
      outputFilename += '.pptx';
    }

    if (orderIndices.length === 0) {
      return res.status(400).json({ error: 'Danh sách thứ tự file không hợp lệ' });
    }

    // Verify all files exist and prepare ordered aliases
    const orderedFilePaths: string[] = [];
    for (let i = 0; i < orderIndices.length; i++) {
      const origIdx = orderIndices[i];
      const origFilePath = path.join(sessionDir, `input_${origIdx}.pptx`);
      if (!fsSync.existsSync(origFilePath)) {
        throw new Error(`Tệp số ${origIdx + 1} chưa được tải lên máy chủ đầy đủ.`);
      }

      const orderedAlias = `ordered_${i}.pptx`;
      const orderedFullPath = path.join(sessionDir, orderedAlias);
      await fs.copyFile(origFilePath, orderedFullPath);
      orderedFilePaths.push(orderedAlias);
    }

    const automizer = new Automizer({
      templateDir: sessionDir,
      outputDir: sessionDir,
      verbosity: 0,
    });

    const rootFilename = orderedFilePaths[0];
    const pres = automizer.loadRoot(rootFilename);

    // Merge subsequent presentations
    for (let i = 1; i < orderedFilePaths.length; i++) {
      const templateAlias = `tpl_${i}`;
      const templateFilename = orderedFilePaths[i];
      automizer.load(templateFilename, templateAlias);

      const originalIndex = orderIndices[i];
      const customSlideSelection = selections[originalIndex.toString()];

      let slidesToAdd: number[] = [];
      if (customSlideSelection && Array.isArray(customSlideSelection) && customSlideSelection.length > 0) {
        slidesToAdd = customSlideSelection;
      } else {
        slidesToAdd = await pres.getTemplate(templateAlias).getAllSlideNumbers();
      }

      for (const slideNum of slidesToAdd) {
        pres.addSlide(templateAlias, slideNum);
      }
    }

    const mergedTempFile = 'merged_output.pptx';
    await pres.write(mergedTempFile);

    const mergedFullPath = path.join(sessionDir, mergedTempFile);
    if (!fsSync.existsSync(mergedFullPath)) {
      throw new Error('Tệp PowerPoint kết quả không được tạo thành công.');
    }

    const stats = await fs.stat(mergedFullPath);

    // Count slides accurately from the merged file
    let totalSlides = 0;
    try {
      const zipBuf = await fs.readFile(mergedFullPath);
      const zip = await JSZip.loadAsync(zipBuf);
      totalSlides = Object.keys(zip.files).filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k)).length;
    } catch {
      totalSlides = orderIndices.length;
    }

    res.json({
      success: true,
      sessionId,
      fileSize: stats.size,
      totalSlides,
      filename: outputFilename,
      downloadUrl: `/api/download/${sessionId}?filename=${encodeURIComponent(outputFilename)}`,
    });
  } catch (err: any) {
    console.error('Error during merge-session:', err);
    res.status(500).json({
      error: 'Quá trình gộp slide gặp lỗi: ' + (err.message || 'Lỗi không xác định'),
    });
  }
});

// Endpoint: Download merged PPTX
app.get('/api/download/:sessionId', async (req, res) => {
  const { sessionId } = req.params;
  const filename = (req.query.filename as string) || 'Merged_Presentation.pptx';
  const sessionDir = path.join(os.tmpdir(), `pptx-session-${sessionId}`);
  const mergedFullPath = path.join(sessionDir, 'merged_output.pptx');

  if (!fsSync.existsSync(mergedFullPath)) {
    return res.status(404).send('Tệp kết quả không tồn tại hoặc đã hết hạn.');
  }

  const stat = await fs.stat(mergedFullPath);
  const encodedFilename = encodeURIComponent(filename);

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"; filename*=UTF-8''${encodedFilename}`);
  res.setHeader('Content-Length', stat.size);
  res.setHeader('Accept-Ranges', 'bytes');

  const stream = fsSync.createReadStream(mergedFullPath);
  stream.pipe(res);
});

// Endpoint: Chunked download for files exceeding 25MB (prevents Cloud Run response 32MB limit)
app.get('/api/download-chunk/:sessionId', async (req, res) => {
  const { sessionId } = req.params;
  const chunkIndex = parseInt((req.query.chunkIndex as string) || '0');
  const chunkSize = parseInt((req.query.chunkSize as string) || `${5 * 1024 * 1024}`); // 5MB

  const sessionDir = path.join(os.tmpdir(), `pptx-session-${sessionId}`);
  const mergedFullPath = path.join(sessionDir, 'merged_output.pptx');

  if (!fsSync.existsSync(mergedFullPath)) {
    return res.status(404).json({ error: 'Tệp không tồn tại' });
  }

  const stat = await fs.stat(mergedFullPath);
  const start = chunkIndex * chunkSize;
  const end = Math.min(start + chunkSize, stat.size);

  if (start >= stat.size) {
    return res.status(416).json({ error: 'Vượt quá dung lượng tệp' });
  }

  res.setHeader('Content-Type', 'application/octet-stream');
  res.setHeader('Content-Length', end - start);
  res.setHeader('X-Total-Size', stat.size);
  res.setHeader('X-Chunk-Index', chunkIndex);
  res.setHeader('X-Total-Chunks', Math.ceil(stat.size / chunkSize));

  const stream = fsSync.createReadStream(mergedFullPath, { start, end: end - 1 });
  stream.pipe(res);
});

// Endpoint: Clean up session
app.delete('/api/session/:sessionId', async (req, res) => {
  const { sessionId } = req.params;
  const sessionDir = path.join(os.tmpdir(), `pptx-session-${sessionId}`);
  try {
    if (fsSync.existsSync(sessionDir)) {
      await fs.rm(sessionDir, { recursive: true, force: true });
    }
    res.json({ success: true });
  } catch {
    res.json({ success: false });
  }
});

// Endpoint: Direct merge fallback for small batches
app.post('/api/merge', legacyUpload.array('files'), async (req, res) => {
  const files = (req.files as Express.Multer.File[]) || [];
  if (!files || files.length === 0) {
    return res.status(400).json({ error: 'Vui lòng tải lên ít nhất một file PowerPoint (.pptx)' });
  }

  let orderIndices: number[] = [];
  try {
    if (req.body.order) {
      orderIndices = JSON.parse(req.body.order);
    } else {
      orderIndices = files.map((_, i) => i);
    }
  } catch {
    orderIndices = files.map((_, i) => i);
  }

  let slideSelections: Record<string, number[]> = {};
  try {
    if (req.body.slideSelections) {
      slideSelections = JSON.parse(req.body.slideSelections);
    }
  } catch (e) {
    console.warn('Could not parse slideSelections:', e);
  }

  let outputFilename = (req.body.outputName || 'Merged_Presentation.pptx').trim();
  if (!outputFilename.toLowerCase().endsWith('.pptx')) {
    outputFilename += '.pptx';
  }

  const orderedFiles = orderIndices
    .filter((idx) => idx >= 0 && idx < files.length)
    .map((idx) => ({ originalIndex: idx, file: files[idx] }));

  if (orderedFiles.length === 0) {
    return res.status(400).json({ error: 'Danh sách thứ tự file không hợp lệ' });
  }

  const tempDir = path.join(os.tmpdir(), `pptx-merge-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`);
  await fs.mkdir(tempDir, { recursive: true });

  try {
    const tempFilePaths: string[] = [];
    for (let i = 0; i < orderedFiles.length; i++) {
      const { file } = orderedFiles[i];
      const safeFilename = `input_${i}.pptx`;
      const fullPath = path.join(tempDir, safeFilename);
      await fs.writeFile(fullPath, file.buffer);
      tempFilePaths.push(safeFilename);
    }

    const automizer = new Automizer({
      templateDir: tempDir,
      outputDir: tempDir,
      verbosity: 0,
    });

    const rootFilename = tempFilePaths[0];
    const pres = automizer.loadRoot(rootFilename);

    for (let i = 1; i < tempFilePaths.length; i++) {
      const templateAlias = `tpl_${i}`;
      const templateFilename = tempFilePaths[i];
      automizer.load(templateFilename, templateAlias);

      const origIdx = orderedFiles[i].originalIndex;
      const customSlideSelection = slideSelections[origIdx.toString()];

      let slidesToAdd: number[] = [];
      if (customSlideSelection && Array.isArray(customSlideSelection) && customSlideSelection.length > 0) {
        slidesToAdd = customSlideSelection;
      } else {
        slidesToAdd = await pres.getTemplate(templateAlias).getAllSlideNumbers();
      }

      for (const slideNum of slidesToAdd) {
        pres.addSlide(templateAlias, slideNum);
      }
    }

    const mergedTempFile = 'merged_output.pptx';
    await pres.write(mergedTempFile);

    const mergedFullPath = path.join(tempDir, mergedTempFile);
    if (!fsSync.existsSync(mergedFullPath)) {
      throw new Error('File kết quả không được tạo thành công.');
    }

    const mergedBuffer = await fs.readFile(mergedFullPath);
    const encodedFilename = encodeURIComponent(outputFilename);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${outputFilename}"; filename*=UTF-8''${encodedFilename}`);
    res.setHeader('Content-Length', mergedBuffer.length);
    res.send(mergedBuffer);
  } catch (err: any) {
    console.error('Error during PPTX merge:', err);
    res.status(500).json({
      error: 'Quá trình gộp slide gặp lỗi: ' + (err.message || 'Lỗi không xác định'),
    });
  } finally {
    try {
      await fs.rm(tempDir, { recursive: true, force: true });
    } catch (cleanupErr) {
      console.warn('Could not clean temp directory:', cleanupErr);
    }
  }
});

// Express error handler for multer/payload errors
app.use((err: any, req: express.Request, res: express.Response, next: express.NextFunction) => {
  if (err && (err.code === 'LIMIT_FILE_SIZE' || err.status === 413)) {
    return res.status(413).json({
      error: 'Dung lượng tệp vượt quá giới hạn truyền tải của một yêu cầu. Hệ thống đang tự động chuyển sang chế độ tải từng phần.',
    });
  }
  next(err);
});

async function startServer() {
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(path.resolve('dist')));
    app.get('*', (req, res) => {
      res.sendFile(path.resolve('dist', 'index.html'));
    });
  }

  if (process.env.VERCEL !== '1') {
    app.listen(PORT, '0.0.0.0', () => {
      console.log(`Server running at http://0.0.0.0:${PORT}`);
    });
  }
}

export default app;

if (process.env.VERCEL !== '1') {
  startServer().catch((err) => {
    console.error('Failed to start server:', err);
  });
}

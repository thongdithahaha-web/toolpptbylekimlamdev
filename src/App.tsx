import React, { useState, useEffect } from 'react';
import { Header } from './components/Header.tsx';
import { DropZone } from './components/DropZone.tsx';
import { FileItem } from './components/FileItem.tsx';
import { MergeToolbar } from './components/MergeToolbar.tsx';
import { SlideInspectorModal } from './components/SlideInspectorModal.tsx';
import { MergeModal } from './components/MergeModal.tsx';
import { MergeSuccessCard } from './components/MergeSuccessCard.tsx';
import { GuideSection } from './components/GuideSection.tsx';
import { PPTFileItem, MergeResult } from './types.ts';
import { parsePPTXFile } from './utils/pptxParser.ts';
import { AlertCircle, Presentation, CheckCircle, ArrowDownUp } from 'lucide-react';

export default function App() {
  const [files, setFiles] = useState<PPTFileItem[]>([]);
  const [outputName, setOutputName] = useState<string>('Gop_Slide_Tong_Hop.pptx');
  const [isMerging, setIsMerging] = useState(false);
  const [mergeStep, setMergeStep] = useState(1);
  const [uploadProgress, setUploadProgress] = useState(0);
  const [statusMessage, setStatusMessage] = useState<string>('');
  const [mergeError, setMergeError] = useState<string | null>(null);
  const [mergeResult, setMergeResult] = useState<MergeResult | null>(null);
  const [inspectingItem, setInspectingItem] = useState<PPTFileItem | null>(null);
  const [isLoadingSamples, setIsLoadingSamples] = useState(false);
  const [generalError, setGeneralError] = useState<string | null>(null);
  const [draggedIndex, setDraggedIndex] = useState<number | null>(null);

  // Set default output filename with formatted date
  useEffect(() => {
    const today = new Date();
    const dateStr = today.toISOString().slice(0, 10).replace(/-/g, '');
    setOutputName(`Gop_PowerPoint_${dateStr}.pptx`);
  }, []);

  // Handle newly selected/dropped files
  const handleFilesSelected = async (newFiles: File[]) => {
    setGeneralError(null);
    if (!newFiles || newFiles.length === 0) return;

    const initialItems: PPTFileItem[] = newFiles.map((file) => ({
      id: `${file.name}-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
      file,
      name: file.name,
      size: file.size,
      slideCount: 1,
      slides: [],
      selectedSlideNumbers: [1],
      isAnalyzing: true,
    }));

    setFiles((prev) => [...prev, ...initialItems]);

    // Parse each file client-side with JSZip to get slide counts & previews
    for (const item of initialItems) {
      try {
        const { slideCount, slides } = await parsePPTXFile(item.file);
        setFiles((prev) =>
          prev.map((f) =>
            f.id === item.id
              ? {
                  ...f,
                  slideCount,
                  slides,
                  selectedSlideNumbers: slides.map((s) => s.slideNumber),
                  isAnalyzing: false,
                }
              : f
          )
        );
      } catch (err: any) {
        console.warn(`Error parsing ${item.name}:`, err);
        setFiles((prev) =>
          prev.map((f) =>
            f.id === item.id
              ? {
                  ...f,
                  isAnalyzing: false,
                }
              : f
          )
        );
        setGeneralError(err.message || `Không thể đọc tệp ${item.name}`);
      }
    }
  };

  // Load sample demo files for immediate testing
  const handleLoadSamples = async () => {
    try {
      setIsLoadingSamples(true);
      setGeneralError(null);
      const res = await fetch('/api/samples');
      if (!res.ok) {
        throw new Error('Không thể tải các tệp mẫu từ máy chủ.');
      }
      const data = await res.json();
      if (!data.samples || !Array.isArray(data.samples)) {
        throw new Error('Dữ liệu tệp mẫu không đúng định dạng.');
      }

      const sampleFiles: File[] = data.samples.map((s: any) => {
        const binaryString = atob(s.base64);
        const bytes = new Uint8Array(binaryString.length);
        for (let i = 0; i < binaryString.length; i++) {
          bytes[i] = binaryString.charCodeAt(i);
        }
        return new File([bytes], s.name, {
          type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        });
      });

      await handleFilesSelected(sampleFiles);
    } catch (err: any) {
      console.error('Failed to load samples:', err);
      setGeneralError('Không thể tạo file mẫu: ' + err.message);
    } finally {
      setIsLoadingSamples(false);
    }
  };

  // Move operations
  const handleMoveUp = (index: number) => {
    if (index <= 0) return;
    setFiles((prev) => {
      const copy = [...prev];
      const temp = copy[index - 1];
      copy[index - 1] = copy[index];
      copy[index] = temp;
      return copy;
    });
  };

  const handleMoveDown = (index: number) => {
    if (index >= files.length - 1) return;
    setFiles((prev) => {
      const copy = [...prev];
      const temp = copy[index + 1];
      copy[index + 1] = copy[index];
      copy[index] = temp;
      return copy;
    });
  };

  const handleMoveToTop = (index: number) => {
    if (index <= 0) return;
    setFiles((prev) => {
      const copy = [...prev];
      const item = copy.splice(index, 1)[0];
      copy.unshift(item);
      return copy;
    });
  };

  const handleMoveToBottom = (index: number) => {
    if (index >= files.length - 1) return;
    setFiles((prev) => {
      const copy = [...prev];
      const item = copy.splice(index, 1)[0];
      copy.push(item);
      return copy;
    });
  };

  // Sorting helpers
  const handleSortAZ = () => {
    setFiles((prev) => [...prev].sort((a, b) => a.name.localeCompare(b.name, 'vi')));
  };

  const handleSortZA = () => {
    setFiles((prev) => [...prev].sort((a, b) => b.name.localeCompare(a.name, 'vi')));
  };

  const handleReverse = () => {
    setFiles((prev) => [...prev].reverse());
  };

  const handleClearAll = () => {
    setFiles([]);
    setGeneralError(null);
    setMergeResult(null);
  };

  const handleRemove = (id: string) => {
    setFiles((prev) => prev.filter((f) => f.id !== id));
  };

  // Drag and Drop reordering
  const handleDragStart = (e: React.DragEvent, index: number) => {
    setDraggedIndex(index);
    e.dataTransfer.effectAllowed = 'move';
  };

  const handleDragOver = (e: React.DragEvent, index: number) => {
    e.preventDefault();
    if (draggedIndex === null || draggedIndex === index) return;

    setFiles((prev) => {
      const copy = [...prev];
      const item = copy.splice(draggedIndex, 1)[0];
      copy.splice(index, 0, item);
      return copy;
    });
    setDraggedIndex(index);
  };

  const handleDragEnd = () => {
    setDraggedIndex(null);
  };

  // Slide Inspector
  const handleUpdateSelectedSlides = (id: string, selectedSlideNumbers: number[]) => {
    setFiles((prev) =>
      prev.map((f) => (f.id === id ? { ...f, selectedSlideNumbers } : f))
    );
    if (inspectingItem && inspectingItem.id === id) {
      setInspectingItem((prev) => (prev ? { ...prev, selectedSlideNumbers } : null));
    }
  };

  // Execution: Merge files using chunked upload to prevent HTTP 413
  const handleStartMerge = async () => {
    if (files.length < 2) {
      setGeneralError('Vui lòng thêm ít nhất 2 bài thuyết trình PowerPoint để gộp.');
      return;
    }

    setIsMerging(true);
    setMergeStep(1);
    setUploadProgress(0);
    setMergeError(null);
    setStatusMessage('Đang chuẩn bị tệp tin...');

    const sessionId = `sess_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
    const CHUNK_SIZE = 2 * 1024 * 1024; // 2MB chunk prevents HTTP 413 Payload Too Large on Cloud Run

    try {
      // Calculate total bytes across all files
      const totalBytes = files.reduce((acc, f) => acc + f.size, 0);
      let uploadedBytes = 0;

      // STEP 1: Upload files chunk by chunk
      for (let fileIdx = 0; fileIdx < files.length; fileIdx++) {
        const item = files[fileIdx];
        const file = item.file;
        const totalChunks = Math.max(1, Math.ceil(file.size / CHUNK_SIZE));

        for (let chunkIdx = 0; chunkIdx < totalChunks; chunkIdx++) {
          const start = chunkIdx * CHUNK_SIZE;
          const end = Math.min(start + CHUNK_SIZE, file.size);
          const chunkBlob = file.slice(start, end);

          const chunkFormData = new FormData();
          chunkFormData.append('sessionId', sessionId);
          chunkFormData.append('fileIndex', fileIdx.toString());
          chunkFormData.append('chunkIndex', chunkIdx.toString());
          chunkFormData.append('totalChunks', totalChunks.toString());
          chunkFormData.append('fileName', item.name);
          chunkFormData.append('chunk', chunkBlob, item.name);

          setStatusMessage(`Đang truyền tải "${item.name}" (phần ${chunkIdx + 1}/${totalChunks})...`);

          const chunkRes = await fetch('/api/upload-chunk', {
            method: 'POST',
            body: chunkFormData,
          });

          if (!chunkRes.ok) {
            let chunkErrMsg = 'Lỗi truyền tải tệp';
            try {
              const errJson = await chunkRes.json();
              chunkErrMsg = errJson.error || chunkErrMsg;
            } catch {
              chunkErrMsg = `Lỗi HTTP ${chunkRes.status}`;
            }
            throw new Error(`Tải tệp "${item.name}" thất bại: ${chunkErrMsg}`);
          }

          uploadedBytes += (end - start);
          setUploadProgress(Math.min(99, (uploadedBytes / Math.max(1, totalBytes)) * 100));
        }
      }

      setUploadProgress(100);

      // STEP 2: Analyze OpenXML structure
      setMergeStep(2);
      setStatusMessage('Đang phân tích cấu trúc slide và bố cục XML...');
      await new Promise((r) => setTimeout(r, 350));

      // STEP 3: Merge slides in specified order
      setMergeStep(3);
      setStatusMessage('Đang ghép các slide theo thứ tự đã chọn...');

      const orderIndices = files.map((_, i) => i);
      const slideSelections: Record<string, number[]> = {};
      files.forEach((item, index) => {
        if (item.selectedSlideNumbers && item.selectedSlideNumbers.length > 0) {
          slideSelections[index.toString()] = item.selectedSlideNumbers;
        }
      });

      const mergeRes = await fetch('/api/merge-session', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          sessionId,
          order: orderIndices,
          slideSelections,
          outputName: outputName.trim() || 'Merged_Presentation.pptx',
        }),
      });

      if (!mergeRes.ok) {
        let mergeErrMsg = 'Không thể hoàn tất ghép slide.';
        try {
          const errJson = await mergeRes.json();
          mergeErrMsg = errJson.error || mergeErrMsg;
        } catch {
          mergeErrMsg = `Lỗi máy chủ HTTP ${mergeRes.status}`;
        }
        throw new Error(mergeErrMsg);
      }

      const mergeData = await mergeRes.json();

      // STEP 4: Package and fetch result
      setMergeStep(4);
      setStatusMessage('Đang hoàn thiện và chuẩn bị tải về...');

      let finalBlob: Blob;
      const finalSize = mergeData.fileSize || 0;

      // Handle download safely
      if (finalSize > 25 * 1024 * 1024) {
        // Multi-chunk download for large files > 25MB
        setStatusMessage('Đang nhận dữ liệu bài thuyết trình lớn...');
        const dlChunkSize = 5 * 1024 * 1024;
        const totalDlChunks = Math.ceil(finalSize / dlChunkSize);
        const downloadedParts: Blob[] = [];

        for (let c = 0; c < totalDlChunks; c++) {
          const dlRes = await fetch(`/api/download-chunk/${sessionId}?chunkIndex=${c}&chunkSize=${dlChunkSize}`);
          if (!dlRes.ok) {
            throw new Error('Lỗi khi tải dữ liệu bài thuyết trình');
          }
          const part = await dlRes.blob();
          downloadedParts.push(part);
        }

        finalBlob = new Blob(downloadedParts, {
          type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        });
      } else {
        const directRes = await fetch(mergeData.downloadUrl);
        if (!directRes.ok) {
          throw new Error('Không thể tải tệp kết quả từ máy chủ');
        }
        finalBlob = await directRes.blob();
      }

      const blobUrl = URL.createObjectURL(finalBlob);

      setMergeResult({
        blob: finalBlob,
        url: blobUrl,
        filename: mergeData.filename || outputName.trim() || 'Merged_Presentation.pptx',
        totalSlides: mergeData.totalSlides || totalSelectedSlides,
        totalFiles: files.length,
        fileSize: finalBlob.size,
        createdAt: new Date(),
      });

      // Cleanup session on server
      fetch(`/api/session/${sessionId}`, { method: 'DELETE' }).catch(() => {});

      setIsMerging(false);
    } catch (err: any) {
      console.error('Merge error:', err);
      setMergeError(err.message || 'Quá trình gộp gặp sự cố ngoài ý muốn.');
    }
  };

  // Calculate cumulative slide positions
  let cumulativeSlideCount = 1;
  const slidePositions = files.map((item) => {
    const start = cumulativeSlideCount;
    const count = item.selectedSlideNumbers?.length || item.slideCount;
    cumulativeSlideCount += count;
    return start;
  });

  const totalSelectedSlides = files.reduce(
    (acc, f) => acc + (f.selectedSlideNumbers?.length || f.slideCount),
    0
  );
  const totalFileSize = files.reduce((acc, f) => acc + f.size, 0);

  return (
    <div className="min-h-screen bg-gradient-to-b from-slate-50 via-white to-slate-100/70 text-slate-900 font-sans antialiased flex flex-col">
      <Header onLoadSamples={handleLoadSamples} isLoadingSamples={isLoadingSamples} />

      <main className="flex-1 max-w-5xl w-full mx-auto px-4 sm:px-6 py-8 sm:py-12">
        {/* Hero title */}
        <div className="text-center max-w-2xl mx-auto mb-10">
          <div className="inline-flex items-center gap-2 px-3 py-1.5 rounded-full bg-rose-50 border border-rose-200 text-rose-700 text-xs font-bold mb-4 shadow-2xs">
            <Presentation className="w-4 h-4 text-rose-600" />
            <span>Công cụ ghép nối slide PowerPoint trực tuyến</span>
          </div>

          <h1 className="text-3xl sm:text-5xl font-extrabold tracking-tight text-slate-900 leading-[1.15]">
            Gộp nhiều file <span className="bg-gradient-to-r from-rose-600 via-orange-600 to-amber-600 bg-clip-text text-transparent">PPTX</span> thành 1 file theo thứ tự
          </h1>

          <p className="mt-4 text-sm sm:text-base text-slate-600 leading-relaxed">
            Tải lên nhiều bài thuyết trình PowerPoint, tự do kéo thả sắp xếp trình tự các trang, và tải về 1 tệp duy nhất với chất lượng đồ họa và bố cục nguyên vẹn.
          </p>
        </div>

        {/* Global error banner if any */}
        {generalError && (
          <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-2xl flex items-start justify-between gap-3 text-red-700 text-sm animate-in fade-in">
            <div className="flex items-start gap-2.5">
              <AlertCircle className="w-5 h-5 text-red-600 shrink-0 mt-0.5" />
              <span>{generalError}</span>
            </div>
            <button
              type="button"
              onClick={() => setGeneralError(null)}
              className="text-xs font-semibold underline hover:text-red-900 cursor-pointer"
            >
              Đóng
            </button>
          </div>
        )}

        {/* If merge result is available, show success card */}
        {mergeResult ? (
          <MergeSuccessCard
            result={mergeResult}
            onReset={handleClearAll}
            onBackToEdit={() => setMergeResult(null)}
          />
        ) : (
          <div className="space-y-6">
            {/* Upload Drop Zone */}
            <DropZone
              onFilesSelected={handleFilesSelected}
              onLoadSamples={handleLoadSamples}
              isLoadingSamples={isLoadingSamples}
              hasFiles={files.length > 0}
            />

            {/* If files uploaded, show list and reordering tools */}
            {files.length > 0 && (
              <div className="space-y-6 animate-in fade-in duration-300">
                {/* List Header with Order Hint */}
                <div className="flex items-center justify-between px-1">
                  <div className="flex items-center gap-2">
                    <ArrowDownUp className="w-4 h-4 text-rose-600" />
                    <h3 className="text-sm sm:text-base font-extrabold text-slate-900">
                      Danh sách file ({files.length} bài thuyết trình)
                    </h3>
                  </div>
                  <span className="text-xs text-slate-500 hidden sm:inline">
                    💡 Kéo biểu tượng <strong className="text-slate-700">⋮⋮</strong> hoặc bấm mũi tên ▲/▼ để đổi vị trí
                  </span>
                </div>

                {/* Ordered Files List */}
                <div className="space-y-3">
                  {files.map((item, index) => (
                    <FileItem
                      key={item.id}
                      item={item}
                      index={index}
                      totalItems={files.length}
                      startSlideNumber={slidePositions[index]}
                      onMoveUp={handleMoveUp}
                      onMoveDown={handleMoveDown}
                      onMoveToTop={handleMoveToTop}
                      onMoveToBottom={handleMoveToBottom}
                      onRemove={handleRemove}
                      onOpenInspector={(it) => setInspectingItem(it)}
                      onDragStart={handleDragStart}
                      onDragOver={handleDragOver}
                      onDragEnd={handleDragEnd}
                      isDragging={draggedIndex === index}
                    />
                  ))}
                </div>

                {/* Merge Toolbar and Controls */}
                <MergeToolbar
                  totalFiles={files.length}
                  totalSlides={totalSelectedSlides}
                  totalSize={totalFileSize}
                  outputName={outputName}
                  onChangeOutputName={setOutputName}
                  onSortAZ={handleSortAZ}
                  onSortZA={handleSortZA}
                  onReverse={handleReverse}
                  onClearAll={handleClearAll}
                  onStartMerge={handleStartMerge}
                  isMerging={isMerging}
                />
              </div>
            )}
          </div>
        )}

        {/* Modal: Slide Inspector */}
        {inspectingItem && (
          <SlideInspectorModal
            item={inspectingItem}
            onClose={() => setInspectingItem(null)}
            onUpdateSelectedSlides={handleUpdateSelectedSlides}
          />
        )}

        {/* Modal: Merge Progress & Loading */}
        <MergeModal
          isOpen={isMerging || !!mergeError}
          currentStep={mergeStep}
          uploadProgress={uploadProgress}
          statusMessage={statusMessage}
          error={mergeError}
          onRetry={handleStartMerge}
          onCancel={() => {
            setIsMerging(false);
            setMergeError(null);
          }}
        />

        {/* Guide & FAQ */}
        <GuideSection />
      </main>

      {/* Footer */}
      <footer className="mt-20 border-t border-slate-200 bg-white py-8 text-center text-xs text-slate-500">
        <div className="max-w-6xl mx-auto px-4 flex flex-col sm:flex-row items-center justify-between gap-4">
          <div className="flex items-center gap-2 font-medium">
            <span className="w-2 h-2 rounded-full bg-emerald-500 inline-block animate-pulse" />
            <span>PPTMerge • Hỗ trợ ghép nối Microsoft PowerPoint (.pptx)</span>
          </div>
          <p className="text-slate-400">
            Dữ liệu được xử lý riêng tư và an toàn trong phiên làm việc của bạn.
          </p>
        </div>
      </footer>
    </div>
  );
}

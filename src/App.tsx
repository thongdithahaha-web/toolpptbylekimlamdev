import React, { useState, useEffect } from 'react';
import { Header } from './components/Header.tsx';
import { DropZone } from './components/DropZone.tsx';
import { FileItem } from './components/FileItem.tsx';
import { MergeToolbar } from './components/MergeToolbar.tsx';
import { SlideInspectorModal } from './components/SlideInspectorModal.tsx';
import { MergeModal } from './components/MergeModal.tsx';
import { MergeSuccessCard } from './components/MergeSuccessCard.tsx';
import { GuideSection } from './components/GuideSection.tsx';
import { CopyrightModal } from './components/CopyrightModal.tsx';
import { ConverterTab } from './components/ConverterTab.tsx';
import { PPTFileItem, MergeResult } from './types.ts';
import { parsePPTXFile } from './utils/pptxParser.ts';
import { mergePPTXClientSide } from './utils/clientPptxMerger.ts';
import { generateClientSamples } from './utils/sampleGenerator.ts';
import { AlertCircle, Presentation, CheckCircle, ArrowDownUp, BadgeCheck, ShieldCheck, Layers, ArrowRightLeft } from 'lucide-react';

export default function App() {
  const [activeTab, setActiveTab] = useState<'merge' | 'convert'>('merge');
  const [files, setFiles] = useState<PPTFileItem[]>([]);
  const [outputName, setOutputName] = useState<string>('toolgopfilepptbylekimlam.pptx');
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
  const [isCopyrightOpen, setIsCopyrightOpen] = useState(false);

  // Set default output filename with formatted date
  useEffect(() => {
    const today = new Date();
    const dateStr = today.toISOString().slice(0, 10).replace(/-/g, '');
    setOutputName(`toolgopfilepptbylekimlam_${dateStr}.pptx`);
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
      
      // Try instant client-side generation first
      try {
        const clientSamples = await generateClientSamples();
        await handleFilesSelected(clientSamples);
        return;
      } catch (clientErr) {
        console.warn('Client sample generator fallback:', clientErr);
      }

      // Server fallback if needed
      const res = await fetch('/api/samples');
      if (!res.ok) {
        throw new Error('Không thể tải các tệp mẫu.');
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

  // Execution: Pure client-side merge engine (100% reliable, works on Vercel, zero HTTP 500/413)
  const handleStartMerge = async () => {
    if (files.length < 2) {
      setGeneralError('Vui lòng thêm ít nhất 2 bài thuyết trình PowerPoint để gộp.');
      return;
    }

    setIsMerging(true);
    setMergeStep(1);
    setUploadProgress(10);
    setMergeError(null);
    setStatusMessage('Đang khởi tạo trình gộp OpenXML...');

    try {
      // Execute fast, reliable in-browser OpenXML merge
      const result = await mergePPTXClientSide(files, (info) => {
        setMergeStep(info.step);
        setUploadProgress(info.percentage);
        setStatusMessage(info.message);
      });

      const blobUrl = URL.createObjectURL(result.blob);

      setMergeResult({
        blob: result.blob,
        url: blobUrl,
        filename: outputName.trim() || 'toolgopfilepptbylekimlam.pptx',
        totalSlides: result.totalSlides,
        totalFiles: files.length,
        fileSize: result.fileSize,
        createdAt: new Date(),
        auditItems: result.auditItems,
      });

      setIsMerging(false);
    } catch (err: any) {
      console.error('Merge error:', err);
      setMergeError(err.message || 'Quá trình gộp gặp sự cố ngoài ý muốn. Vui lòng kiểm tra lại tính hợp lệ của tệp.');
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
      <Header
        activeTab={activeTab}
        onChangeTab={setActiveTab}
        onLoadSamples={handleLoadSamples}
        isLoadingSamples={isLoadingSamples}
        onOpenCopyright={() => setIsCopyrightOpen(true)}
      />

      <main className="flex-1 max-w-5xl w-full mx-auto px-4 sm:px-6 py-8 sm:py-12">
        {/* Navigation Tab Selector Bar */}
        <div className="flex justify-center mb-8">
          <div className="inline-flex p-1.5 bg-slate-200/80 rounded-2xl border border-slate-300 shadow-xs max-w-full overflow-x-auto">
            <button
              type="button"
              onClick={() => setActiveTab('merge')}
              className={`flex items-center gap-2 px-5 py-2.5 rounded-xl font-bold text-xs sm:text-sm transition-all cursor-pointer shrink-0 ${
                activeTab === 'merge'
                  ? 'bg-white text-rose-600 shadow-sm'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Layers className="w-4 h-4" />
              <span>Gộp File PowerPoint (Không thiếu sót)</span>
            </button>
            <button
              type="button"
              onClick={() => setActiveTab('convert')}
              className={`flex items-center gap-2 px-5 py-2.5 rounded-xl font-bold text-xs sm:text-sm transition-all cursor-pointer shrink-0 ${
                activeTab === 'convert'
                  ? 'bg-white text-rose-600 shadow-sm'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <ArrowRightLeft className="w-4 h-4" />
              <span>Chuyển Đổi Định Dạng Đa Năng</span>
              <span className="text-[10px] bg-rose-500 text-white font-extrabold px-1.5 py-0.5 rounded-md uppercase">Mới</span>
            </button>
          </div>
        </div>

        {activeTab === 'convert' ? (
          <ConverterTab />
        ) : (
          <>
            {/* Hero title with Copyright Branding */}
            <div className="text-center max-w-3xl mx-auto mb-10">
              <div className="inline-flex items-center gap-2 px-3.5 py-1.5 rounded-full bg-rose-50 border border-rose-200 text-rose-700 text-xs font-bold mb-4 shadow-2xs flex-wrap justify-center">
                <Presentation className="w-4 h-4 text-rose-600" />
                <span className="font-mono">toolgopfilepptbylekimlam</span>
                <span className="text-slate-300">•</span>
                <button
                  type="button"
                  onClick={() => setIsCopyrightOpen(true)}
                  className="inline-flex items-center gap-1 text-emerald-700 hover:text-emerald-800 font-bold underline cursor-pointer"
                >
                  <BadgeCheck className="w-3.5 h-3.5 text-emerald-600" />
                  <span>Bản quyền © Lê Kim Lâm</span>
                </button>
              </div>

              <h1 className="text-3xl sm:text-5xl font-extrabold tracking-tight text-slate-900 leading-[1.15]">
                <span className="font-mono text-rose-600 block sm:inline">toolgopfilepptbylekimlam</span>
                <span className="block text-2xl sm:text-4xl text-slate-800 font-bold mt-2">
                  Gộp nhiều file <span className="bg-gradient-to-r from-rose-600 via-orange-600 to-amber-600 bg-clip-text text-transparent">PowerPoint</span> theo thứ tự
                </span>
              </h1>

              <p className="mt-4 text-sm sm:text-base text-slate-600 leading-relaxed max-w-2xl mx-auto">
                Công cụ trực tuyến bản quyền bởi <strong>Lê Kim Lâm</strong>: Cam kết <strong>không bao giờ thiếu sót file</strong>, giữ nguyên thứ tự ưu tiên 1 đến cuối, bảo toàn chất lượng đồ họa và bố cục.
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
                {/* List Header with Order & Priority Rule */}
                <div className="space-y-2">
                  <div className="flex flex-wrap items-center justify-between gap-2 px-1">
                    <div className="flex items-center gap-2">
                      <ArrowDownUp className="w-4 h-4 text-rose-600" />
                      <h3 className="text-sm sm:text-base font-extrabold text-slate-900">
                        Thứ tự ưu tiên ghép slide ({files.length} file)
                      </h3>
                    </div>
                    <span className="text-xs text-slate-500">
                      💡 Kéo <strong className="text-slate-700">⋮⋮</strong> hoặc bấm <strong className="text-slate-700">▲/▼</strong> để đổi vị trí ưu tiên
                    </span>
                  </div>

                  {/* Priority Flow Diagram */}
                  <div className="p-3 bg-gradient-to-r from-amber-50/70 via-rose-50/50 to-indigo-50/50 border border-amber-200/80 rounded-2xl">
                    <div className="flex items-center gap-2 mb-2 text-xs font-bold text-amber-900">
                      <span className="w-2 h-2 rounded-full bg-amber-500 animate-ping inline-block" />
                      <span>Trình tự ghép: File chọn đầu tiên (#1) là ƯU TIÊN CAO NHẤT (bắt đầu từ Slide 1 đến cuối)</span>
                    </div>

                    <div className="flex items-center gap-1.5 overflow-x-auto py-1 text-xs scrollbar-thin">
                      {files.map((f, i) => (
                        <React.Fragment key={f.id}>
                          <div
                            className={`flex items-center gap-1.5 px-2.5 py-1 rounded-lg shrink-0 border font-medium ${
                              i === 0
                                ? 'bg-amber-400 text-slate-950 font-bold border-amber-500 shadow-xs'
                                : 'bg-white text-slate-700 border-slate-200'
                            }`}
                          >
                            <span className="font-mono text-[11px] font-black">
                              {i === 0 ? '🌟 #1' : `#${i + 1}`}
                            </span>
                            <span className="truncate max-w-[120px]">{f.name}</span>
                          </div>
                          {i < files.length - 1 && (
                            <span className="text-slate-400 font-bold shrink-0">➔</span>
                          )}
                        </React.Fragment>
                      ))}
                    </div>
                  </div>
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
        </>
        )}

        {/* Guide & FAQ */}
        <GuideSection />
      </main>

      {/* Modal: Copyright Details */}
      <CopyrightModal
        isOpen={isCopyrightOpen}
        onClose={() => setIsCopyrightOpen(false)}
      />

      {/* Footer */}
      <footer className="mt-20 border-t border-slate-200 bg-white py-8 text-xs text-slate-500">
        <div className="max-w-6xl mx-auto px-4 sm:px-6 flex flex-col md:flex-row items-center justify-between gap-4">
          <div className="flex flex-col sm:flex-row items-center gap-2 sm:gap-3 text-center sm:text-left">
            <div className="flex items-center gap-2 font-bold text-slate-800 font-mono text-sm">
              <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 inline-block animate-pulse" />
              <span>toolgopfilepptbylekimlam</span>
            </div>
            <span className="text-slate-300 hidden sm:inline">•</span>
            <span className="text-slate-600">
              Bản quyền © 2026 <strong>Lê Kim Lâm</strong>. Mọi quyền được bảo lưu.
            </span>
          </div>

          <div className="flex items-center gap-3 flex-wrap justify-center">
            <a
              href="mailto:lekimlam2707@gmail.com"
              className="text-slate-600 hover:text-rose-600 transition-colors font-mono"
            >
              lekimlam2707@gmail.com
            </a>
            <span className="text-slate-300">•</span>
            <button
              type="button"
              onClick={() => setIsCopyrightOpen(true)}
              className="flex items-center gap-1 text-emerald-700 hover:text-emerald-800 font-semibold bg-emerald-50 px-2.5 py-1 rounded-lg border border-emerald-200 hover:bg-emerald-100 transition-all cursor-pointer"
            >
              <BadgeCheck className="w-3.5 h-3.5 text-emerald-600" />
              <span>Chứng nhận bản quyền</span>
            </button>
          </div>
        </div>
      </footer>
    </div>
  );
}

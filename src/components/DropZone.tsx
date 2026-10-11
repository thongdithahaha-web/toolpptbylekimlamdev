import React, { useRef, useState } from 'react';
import { UploadCloud, FileType, Sparkles, Plus, AlertCircle } from 'lucide-react';

interface DropZoneProps {
  onFilesSelected: (files: File[]) => void;
  onLoadSamples: () => void;
  isLoadingSamples: boolean;
  hasFiles: boolean;
}

export const DropZone: React.FC<DropZoneProps> = ({
  onFilesSelected,
  onLoadSamples,
  isLoadingSamples,
  hasFiles,
}) => {
  const [isDragOver, setIsDragOver] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);

    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const droppedFiles = Array.from(e.dataTransfer.files).filter((file) =>
        file.name.toLowerCase().endsWith('.pptx') || file.name.toLowerCase().endsWith('.ppt')
      );
      if (droppedFiles.length > 0) {
        onFilesSelected(droppedFiles);
      }
    }
  };

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      const selectedFiles = Array.from(e.target.files);
      onFilesSelected(selectedFiles);
      // Reset input value so same files can be re-selected if removed
      e.target.value = '';
    }
  };

  if (hasFiles) {
    // Compact Add Files button bar when files are already present
    return (
      <div className="flex flex-wrap items-center justify-between gap-3 p-4 bg-white rounded-2xl border border-slate-200/90 shadow-xs">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 rounded-xl bg-rose-50 border border-rose-100 flex items-center justify-center text-rose-600 font-bold">
            <FileType className="w-5 h-5" />
          </div>
          <div>
            <h4 className="text-sm font-bold text-slate-800">Thêm tệp PowerPoint khác</h4>
            <p className="text-xs text-slate-500">Chọn thêm các file thuyết trình để sắp xếp vào danh sách</p>
          </div>
        </div>

        <div className="flex items-center gap-2">
          <input
            ref={fileInputRef}
            type="file"
            multiple
            accept=".pptx,.ppt,application/vnd.openxmlformats-officedocument.presentationml.presentation"
            onChange={handleInputChange}
            className="hidden"
          />
          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            className="flex items-center gap-2 px-4 py-2 text-sm font-semibold text-rose-700 bg-rose-50 hover:bg-rose-100 border border-rose-200 rounded-xl transition-all shadow-xs cursor-pointer active:scale-95"
          >
            <Plus className="w-4 h-4" />
            <span>Thêm file .PPTX</span>
          </button>
        </div>
      </div>
    );
  }

  return (
    <div
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
      className={`relative rounded-3xl border-2 border-dashed transition-all duration-200 p-8 sm:p-12 text-center cursor-pointer ${
        isDragOver
          ? 'border-rose-500 bg-rose-50/50 scale-[1.01]'
          : 'border-slate-300 hover:border-rose-400 bg-white hover:bg-slate-50/50 shadow-sm'
      }`}
      onClick={() => fileInputRef.current?.click()}
    >
      <input
        ref={fileInputRef}
        type="file"
        multiple
        accept=".pptx,.ppt,application/vnd.openxmlformats-officedocument.presentationml.presentation"
        onChange={handleInputChange}
        className="hidden"
      />

      <div className="max-w-md mx-auto flex flex-col items-center">
        <div
          className={`w-20 h-20 rounded-2xl flex items-center justify-center mb-5 transition-transform duration-200 ${
            isDragOver
              ? 'bg-rose-500 text-white scale-110 shadow-lg shadow-rose-500/25'
              : 'bg-rose-100/80 text-rose-600'
          }`}
        >
          <UploadCloud className="w-10 h-10" />
        </div>

        <h3 className="text-xl sm:text-2xl font-bold text-slate-900 mb-2">
          Kéo & thả các file PowerPoint vào đây
        </h3>
        <p className="text-sm text-slate-600 mb-6 leading-relaxed">
          hoặc bấm để chọn nhiều tệp tin từ máy tính của bạn. Bạn sẽ có thể tự do thay đổi thứ tự các slide trước khi gộp.
        </p>

        <div className="flex flex-col sm:flex-row items-center gap-3 w-full sm:w-auto">
          <button
            type="button"
            className="w-full sm:w-auto flex items-center justify-center gap-2 px-6 py-3 bg-gradient-to-r from-rose-600 to-orange-600 hover:from-rose-700 hover:to-orange-700 text-white font-semibold text-sm rounded-xl shadow-md shadow-rose-600/20 transition-all cursor-pointer"
          >
            <FileType className="w-4 h-4" />
            <span>Chọn file PowerPoint (.pptx)</span>
          </button>

          <button
            type="button"
            onClick={(e) => {
              e.stopPropagation();
              onLoadSamples();
            }}
            disabled={isLoadingSamples}
            className="w-full sm:w-auto flex items-center justify-center gap-2 px-5 py-3 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold text-sm rounded-xl border border-slate-200 transition-all cursor-pointer disabled:opacity-50"
          >
            <Sparkles className="w-4 h-4 text-amber-500" />
            <span>{isLoadingSamples ? 'Đang tạo mẫu...' : 'Thử ngay với 3 file mẫu'}</span>
          </button>
        </div>

        <div className="mt-6 flex items-center gap-2 text-xs text-slate-500 bg-slate-100/70 px-3.5 py-1.5 rounded-full border border-slate-200">
          <AlertCircle className="w-3.5 h-3.5 text-slate-400" />
          <span>Hỗ trợ định dạng <strong>.pptx</strong> (PowerPoint 2007 trở lên). Giữ trọn vẹn định dạng.</span>
        </div>
      </div>
    </div>
  );
};

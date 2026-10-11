import React from 'react';
import {
  ArrowDownAZ,
  ArrowUpAZ,
  ArrowUpDown,
  Trash2,
  FileCheck,
  Sparkles,
  Edit3,
} from 'lucide-react';

interface MergeToolbarProps {
  totalFiles: number;
  totalSlides: number;
  totalSize: number;
  outputName: string;
  onChangeOutputName: (name: string) => void;
  onSortAZ: () => void;
  onSortZA: () => void;
  onReverse: () => void;
  onClearAll: () => void;
  onStartMerge: () => void;
  isMerging: boolean;
}

export const MergeToolbar: React.FC<MergeToolbarProps> = ({
  totalFiles,
  totalSlides,
  totalSize,
  outputName,
  onChangeOutputName,
  onSortAZ,
  onSortZA,
  onReverse,
  onClearAll,
  onStartMerge,
  isMerging,
}) => {
  return (
    <div className="bg-white rounded-3xl p-5 sm:p-6 border border-slate-200/90 shadow-sm space-y-5">
      {/* Top row: Reorder helpers & Clear */}
      <div className="flex flex-wrap items-center justify-between gap-3 pb-4 border-b border-slate-100">
        <div className="flex items-center gap-2">
          <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">
            Sắp xếp nhanh:
          </span>
          <button
            type="button"
            onClick={onSortAZ}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors cursor-pointer"
            title="Sắp xếp file theo tên từ A đến Z"
          >
            <ArrowDownAZ className="w-3.5 h-3.5 text-slate-600" />
            <span>Tên A → Z</span>
          </button>
          <button
            type="button"
            onClick={onSortZA}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors cursor-pointer"
            title="Sắp xếp file theo tên từ Z đến A"
          >
            <ArrowUpAZ className="w-3.5 h-3.5 text-slate-600" />
            <span>Tên Z → A</span>
          </button>
          <button
            type="button"
            onClick={onReverse}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors cursor-pointer"
            title="Đảo ngược hoàn toàn thứ tự hiện tại"
          >
            <ArrowUpDown className="w-3.5 h-3.5 text-slate-600" />
            <span>Đảo ngược</span>
          </button>
        </div>

        <button
          type="button"
          onClick={onClearAll}
          className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-red-600 hover:text-red-700 bg-red-50 hover:bg-red-100 rounded-lg transition-colors cursor-pointer"
          title="Xóa toàn bộ file và bắt đầu lại"
        >
          <Trash2 className="w-3.5 h-3.5" />
          <span>Xóa tất cả</span>
        </button>
      </div>

      {/* Middle row: Output filename customizer */}
      <div className="grid grid-cols-1 md:grid-cols-12 gap-4 items-center">
        <div className="md:col-span-7">
          <label className="block text-xs font-bold text-slate-700 mb-1.5 flex items-center gap-1.5">
            <Edit3 className="w-3.5 h-3.5 text-rose-600" />
            <span>Tên file PowerPoint sau khi gộp</span>
          </label>
          <div className="relative">
            <input
              type="text"
              value={outputName}
              onChange={(e) => onChangeOutputName(e.target.value)}
              placeholder="VD: Bao_Cao_Tong_Hop_2026.pptx"
              className="w-full px-4 py-2.5 text-sm font-medium bg-slate-50 border border-slate-200 rounded-xl focus:bg-white focus:outline-none focus:ring-2 focus:ring-rose-500/20 focus:border-rose-500 transition-all text-slate-900 pr-16"
            />
            <span className="absolute right-3.5 top-1/2 -translate-y-1/2 text-xs font-mono font-semibold text-slate-400 select-none">
              .pptx
            </span>
          </div>
        </div>

        {/* Summary stats */}
        <div className="md:col-span-5 flex items-center justify-between sm:justify-end gap-3 sm:gap-6 p-3 bg-slate-50 rounded-xl border border-slate-100">
          <div>
            <span className="text-[11px] text-slate-500 block uppercase font-bold">Số file</span>
            <span className="text-base sm:text-lg font-extrabold text-slate-900 font-mono">
              {totalFiles}
            </span>
          </div>
          <div className="h-8 w-px bg-slate-200" />
          <div>
            <span className="text-[11px] text-slate-500 block uppercase font-bold">Tổng slide</span>
            <span className="text-base sm:text-lg font-extrabold text-rose-600 font-mono">
              {totalSlides}
            </span>
          </div>
          <div className="h-8 w-px bg-slate-200" />
          <div>
            <span className="text-[11px] text-slate-500 block uppercase font-bold">Dung lượng</span>
            <span className="text-sm sm:text-base font-bold text-slate-700 font-mono">
              {(totalSize / (1024 * 1024)).toFixed(1)} MB
            </span>
          </div>
        </div>
      </div>

      {/* Bottom row: Big prominent Action CTA */}
      <div className="pt-2">
        <button
          type="button"
          onClick={onStartMerge}
          disabled={totalFiles < 2 || isMerging}
          className={`w-full py-4 px-6 rounded-2xl font-extrabold text-base sm:text-lg flex items-center justify-center gap-3 transition-all duration-200 shadow-lg cursor-pointer ${
            totalFiles < 2
              ? 'bg-slate-200 text-slate-400 cursor-not-allowed shadow-none'
              : 'bg-gradient-to-r from-rose-600 via-rose-500 to-orange-600 hover:from-rose-700 hover:to-orange-700 text-white shadow-rose-600/30 hover:shadow-rose-600/40 hover:scale-[1.008] active:scale-[0.99]'
          }`}
        >
          {totalFiles < 2 ? (
            <span>Cần tải lên ít nhất 2 file PPTX để tiến hành gộp</span>
          ) : (
            <>
              <FileCheck className="w-6 h-6 animate-bounce" />
              <span>GỘP {totalFiles} FILE THÀNH 1 FILE THEO THỨ TỰ</span>
              <Sparkles className="w-5 h-5 text-amber-300" />
            </>
          )}
        </button>
      </div>
    </div>
  );
};

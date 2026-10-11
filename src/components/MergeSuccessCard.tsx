import React from 'react';
import { Download, CheckCircle2, RotateCcw, Presentation, Sparkles, Share2 } from 'lucide-react';
import { MergeResult } from '../types.ts';
import { formatBytes } from '../utils/pptxParser.ts';

interface MergeSuccessCardProps {
  result: MergeResult;
  onReset: () => void;
  onBackToEdit: () => void;
}

export const MergeSuccessCard: React.FC<MergeSuccessCardProps> = ({
  result,
  onReset,
  onBackToEdit,
}) => {
  const handleDownload = () => {
    const a = document.createElement('a');
    a.href = result.url;
    a.download = result.filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  return (
    <div className="bg-white rounded-3xl p-6 sm:p-10 border border-slate-200/90 shadow-xl max-w-2xl mx-auto text-center animate-in zoom-in-95 duration-200">
      <div className="w-20 h-20 rounded-3xl bg-gradient-to-tr from-emerald-500 to-teal-400 text-white flex items-center justify-center mx-auto mb-6 shadow-lg shadow-emerald-500/25">
        <CheckCircle2 className="w-10 h-10" />
      </div>

      <div className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-emerald-50 text-emerald-700 text-xs font-bold border border-emerald-200 mb-3">
        <Sparkles className="w-3.5 h-3.5" />
        <span>Gộp file hoàn tất thành công!</span>
      </div>

      <h2 className="text-2xl sm:text-3xl font-extrabold text-slate-900 mb-2">
        File PowerPoint mới của bạn đã sẵn sàng
      </h2>
      <p className="text-sm text-slate-600 mb-8 max-w-lg mx-auto leading-relaxed">
        Các bài thuyết trình đã được hợp nhất chính xác theo đúng thứ tự bạn đã sắp xếp, giữ trọn vẹn bố cục và hình ảnh gốc.
      </p>

      {/* Info card */}
      <div className="p-5 bg-slate-50 rounded-2xl border border-slate-200 text-left mb-8 space-y-3">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-rose-500 text-white flex items-center justify-center shrink-0">
            <Presentation className="w-5 h-5" />
          </div>
          <div className="min-w-0 flex-1">
            <h4 className="text-sm font-bold text-slate-900 truncate" title={result.filename}>
              {result.filename}
            </h4>
            <p className="text-xs text-slate-500 font-mono mt-0.5">Định dạng Microsoft PowerPoint (.pptx)</p>
          </div>
        </div>

        <div className="pt-3 border-t border-slate-200/70 grid grid-cols-3 gap-2 text-center">
          <div className="p-2 bg-white rounded-xl border border-slate-100">
            <span className="text-[10px] text-slate-400 uppercase font-bold block">Tổng slide</span>
            <span className="text-base font-extrabold text-rose-600 font-mono">{result.totalSlides}</span>
          </div>
          <div className="p-2 bg-white rounded-xl border border-slate-100">
            <span className="text-[10px] text-slate-400 uppercase font-bold block">Từ số file</span>
            <span className="text-base font-extrabold text-slate-800 font-mono">{result.totalFiles}</span>
          </div>
          <div className="p-2 bg-white rounded-xl border border-slate-100">
            <span className="text-[10px] text-slate-400 uppercase font-bold block">Kích thước</span>
            <span className="text-base font-extrabold text-slate-800 font-mono">
              {formatBytes(result.fileSize)}
            </span>
          </div>
        </div>
      </div>

      {/* Action buttons */}
      <div className="space-y-3">
        <button
          type="button"
          onClick={handleDownload}
          className="w-full py-4 px-6 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 text-white font-extrabold text-base sm:text-lg rounded-2xl shadow-lg shadow-emerald-600/25 transition-all flex items-center justify-center gap-2 cursor-pointer active:scale-98"
        >
          <Download className="w-5 h-5" />
          <span>TẢI FILE POWERPOINT VỀ MÁY</span>
        </button>

        <div className="flex items-center justify-center gap-3 pt-2">
          <button
            type="button"
            onClick={onBackToEdit}
            className="px-4 py-2 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-xl transition-colors cursor-pointer"
          >
            Chỉnh sửa lại thứ tự
          </button>
          <button
            type="button"
            onClick={onReset}
            className="flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-xl transition-colors cursor-pointer"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            <span>Gộp nhóm file mới</span>
          </button>
        </div>
      </div>
    </div>
  );
};

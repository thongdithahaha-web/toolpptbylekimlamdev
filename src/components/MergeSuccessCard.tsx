import React, { useState } from 'react';
import { Download, CheckCircle2, RotateCcw, Presentation, Sparkles, ShieldCheck, ChevronDown, ChevronUp } from 'lucide-react';
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
  const [showAuditList, setShowAuditList] = useState(true);

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
      <p className="text-sm text-slate-600 mb-6 max-w-lg mx-auto leading-relaxed">
        Các bài thuyết trình đã được hợp nhất chính xác theo đúng thứ tự ưu tiên, bảo toàn 100% không thiếu sót bất kỳ tệp hay slide nào.
      </p>

      {/* Info card */}
      <div className="p-5 bg-slate-50 rounded-2xl border border-slate-200 text-left mb-6 space-y-3">
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
            <span className="text-[10px] text-slate-400 uppercase font-bold block">Số tệp đã gộp</span>
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

      {/* Integrity Audit Report: Guarantees 0 missed files */}
      {result.auditItems && result.auditItems.length > 0 && (
        <div className="mb-8 p-4 bg-emerald-50/60 rounded-2xl border border-emerald-200/90 text-left">
          <button
            type="button"
            onClick={() => setShowAuditList(!showAuditList)}
            className="w-full flex items-center justify-between text-xs font-bold text-emerald-900 cursor-pointer"
          >
            <div className="flex items-center gap-2">
              <ShieldCheck className="w-4 h-4 text-emerald-600" />
              <span>Kiểm định toàn vẹn: Đã lấy đủ 100% {result.auditItems.length} tệp tin (không thiếu sót)</span>
            </div>
            {showAuditList ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
          </button>

          {showAuditList && (
            <div className="mt-3 pt-3 border-t border-emerald-200/60 space-y-2">
              {result.auditItems.map((item, idx) => (
                <div
                  key={idx}
                  className="flex items-center justify-between text-xs p-2 bg-white/90 rounded-xl border border-emerald-100"
                >
                  <div className="flex items-center gap-2 truncate pr-2">
                    <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                    <span className="font-bold text-slate-800 font-mono text-[11px]">
                      #{item.fileIndex + 1}
                    </span>
                    <span className="truncate text-slate-700 font-medium max-w-[200px] sm:max-w-[280px]">
                      {item.fileName}
                    </span>
                  </div>
                  <div className="shrink-0 flex items-center gap-2 font-mono text-[11px]">
                    <span className="text-emerald-700 font-bold">
                      {item.includedSlidesCount}/{item.totalOriginalSlides} slide
                    </span>
                    <span className="text-slate-400">•</span>
                    <span className="text-indigo-600 font-semibold">
                      Trang #{item.slideRangeStart}→#{item.slideRangeEnd}
                    </span>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

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

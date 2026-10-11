import React from 'react';
import { Presentation, Sparkles, Layers, ShieldCheck, BadgeCheck } from 'lucide-react';

interface HeaderProps {
  onLoadSamples: () => void;
  isLoadingSamples: boolean;
  onOpenCopyright?: () => void;
}

export const Header: React.FC<HeaderProps> = ({ onLoadSamples, isLoadingSamples, onOpenCopyright }) => {
  return (
    <header className="border-b border-slate-200/80 bg-white/95 backdrop-blur-md sticky top-0 z-30 shadow-xs">
      <div className="max-w-6xl mx-auto px-4 sm:px-6 h-16 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-amber-500 via-rose-500 to-orange-600 flex items-center justify-center text-white shadow-md shadow-orange-500/20 shrink-0">
            <Presentation className="w-5 h-5" />
          </div>
          <div className="min-w-0">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="font-extrabold text-base sm:text-lg tracking-tight bg-gradient-to-r from-slate-900 via-rose-950 to-rose-700 bg-clip-text text-transparent font-mono">
                toolgopfilepptbylekimlam
              </span>
              <button
                type="button"
                onClick={onOpenCopyright}
                className="inline-flex items-center gap-1 text-[10px] font-bold tracking-wider px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200 hover:bg-emerald-100 transition-colors cursor-pointer"
                title="Bản quyền chính thức thuộc về Lê Kim Lâm"
              >
                <BadgeCheck className="w-3 h-3 text-emerald-600" />
                <span>© Lê Kim Lâm</span>
              </button>
            </div>
            <p className="text-xs text-slate-500 hidden sm:block">
              Công cụ gộp file PowerPoint trực tuyến theo thứ tự • Tác giả: Lê Kim Lâm
            </p>
          </div>
        </div>

        <div className="flex items-center gap-3">
          <div className="hidden lg:flex items-center gap-4 text-xs font-medium text-slate-600 mr-2">
            <span className="flex items-center gap-1.5 text-slate-600">
              <Layers className="w-4 h-4 text-rose-500" />
              Sắp xếp tự do
            </span>
            <span className="flex items-center gap-1.5 text-slate-600">
              <ShieldCheck className="w-4 h-4 text-emerald-600" />
              Bảo mật 100%
            </span>
          </div>

          <button
            type="button"
            onClick={onLoadSamples}
            disabled={isLoadingSamples}
            className="flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-semibold rounded-lg border border-rose-200 bg-rose-50/70 text-rose-700 hover:bg-rose-100 hover:border-rose-300 transition-all cursor-pointer shadow-xs disabled:opacity-50"
            title="Tải 3 file PowerPoint mẫu để thử nghiệm ghép ngay"
          >
            <Sparkles className="w-3.5 h-3.5 text-rose-600 animate-pulse" />
            <span>{isLoadingSamples ? 'Đang tạo mẫu...' : 'Thử với file mẫu'}</span>
          </button>
        </div>
      </div>
    </header>
  );
};

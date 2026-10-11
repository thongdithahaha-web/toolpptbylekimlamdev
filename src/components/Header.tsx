import React from 'react';
import { Presentation, Sparkles, Layers, ShieldCheck, BadgeCheck, ArrowRightLeft } from 'lucide-react';

interface HeaderProps {
  activeTab: 'merge' | 'convert';
  onChangeTab: (tab: 'merge' | 'convert') => void;
  onLoadSamples: () => void;
  isLoadingSamples: boolean;
  onOpenCopyright?: () => void;
}

export const Header: React.FC<HeaderProps> = ({
  activeTab,
  onChangeTab,
  onLoadSamples,
  isLoadingSamples,
  onOpenCopyright,
}) => {
  return (
    <header className="border-b border-slate-200/80 bg-white/95 backdrop-blur-md sticky top-0 z-30 shadow-xs">
      <div className="max-w-6xl mx-auto px-4 sm:px-6 h-16 flex items-center justify-between gap-4">
        <div className="flex items-center gap-3 min-w-0">
          <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-amber-500 via-rose-500 to-orange-600 flex items-center justify-center text-white shadow-md shadow-orange-500/20 shrink-0">
            <Presentation className="w-5 h-5" />
          </div>
          <div className="min-w-0">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="font-extrabold text-base sm:text-lg tracking-tight bg-gradient-to-r from-slate-900 via-rose-950 to-rose-700 bg-clip-text text-transparent font-mono truncate">
                toolgopfilepptbylekimlam
              </span>
              <button
                type="button"
                onClick={onOpenCopyright}
                className="inline-flex items-center gap-1 text-[10px] font-bold tracking-wider px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200 hover:bg-emerald-100 transition-colors cursor-pointer shrink-0"
                title="Bản quyền chính thức thuộc về Lê Kim Lâm"
              >
                <BadgeCheck className="w-3 h-3 text-emerald-600" />
                <span>© Lê Kim Lâm</span>
              </button>
            </div>
            <p className="text-xs text-slate-500 hidden md:block">
              Gộp slide PowerPoint không thiếu sót & Chuyển đổi định dạng đa năng
            </p>
          </div>
        </div>

        {/* Tab switch buttons */}
        <div className="flex items-center bg-slate-100/90 p-1 rounded-xl border border-slate-200 shrink-0">
          <button
            type="button"
            onClick={() => onChangeTab('merge')}
            className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold rounded-lg transition-all cursor-pointer ${
              activeTab === 'merge'
                ? 'bg-white text-rose-600 shadow-xs'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <Layers className="w-3.5 h-3.5" />
            <span>Gộp PPTX</span>
          </button>

          <button
            type="button"
            onClick={() => onChangeTab('convert')}
            className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold rounded-lg transition-all cursor-pointer ${
              activeTab === 'convert'
                ? 'bg-white text-rose-600 shadow-xs'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <ArrowRightLeft className="w-3.5 h-3.5" />
            <span>Chuyển Đổi File</span>
          </button>
        </div>

        {/* Actions */}
        <div className="hidden sm:flex items-center gap-3 shrink-0">
          {activeTab === 'merge' && (
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
          )}
        </div>
      </div>
    </header>
  );
};

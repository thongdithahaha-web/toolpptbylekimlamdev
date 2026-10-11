import React from 'react';
import { X, CheckSquare, Square, Layers, Image as ImageIcon, Check } from 'lucide-react';
import { PPTFileItem } from '../types.ts';
import { formatBytes } from '../utils/pptxParser.ts';

interface SlideInspectorModalProps {
  item: PPTFileItem | null;
  onClose: () => void;
  onUpdateSelectedSlides: (id: string, selectedSlideNumbers: number[]) => void;
}

export const SlideInspectorModal: React.FC<SlideInspectorModalProps> = ({
  item,
  onClose,
  onUpdateSelectedSlides,
}) => {
  if (!item) return null;

  const handleToggleSlide = (slideNum: number) => {
    let updated: number[];
    if (item.selectedSlideNumbers.includes(slideNum)) {
      if (item.selectedSlideNumbers.length === 1) {
        // Must have at least 1 slide selected
        return;
      }
      updated = item.selectedSlideNumbers.filter((n) => n !== slideNum);
    } else {
      updated = [...item.selectedSlideNumbers, slideNum].sort((a, b) => a - b);
    }
    onUpdateSelectedSlides(item.id, updated);
  };

  const handleSelectAll = () => {
    const all = item.slides.map((s) => s.slideNumber);
    onUpdateSelectedSlides(item.id, all);
  };

  const handleSelectOnlyFirst = () => {
    onUpdateSelectedSlides(item.id, [1]);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs animate-in fade-in duration-200">
      <div className="bg-white rounded-3xl shadow-2xl border border-slate-200 w-full max-w-2xl max-h-[85vh] flex flex-col overflow-hidden">
        {/* Header */}
        <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
          <div>
            <div className="flex items-center gap-2">
              <span className="text-xs font-bold text-rose-600 bg-rose-50 px-2 py-0.5 rounded-md border border-rose-100 uppercase tracking-wider">
                Chi tiết tệp trình chiếu
              </span>
              <span className="text-xs text-slate-500 font-mono">{formatBytes(item.size)}</span>
            </div>
            <h3 className="text-base sm:text-lg font-bold text-slate-900 truncate max-w-md mt-1" title={item.name}>
              {item.name}
            </h3>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-2 text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded-xl transition-colors cursor-pointer"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Toolbar */}
        <div className="px-6 py-3 bg-white border-b border-slate-100 flex items-center justify-between flex-wrap gap-2 text-xs">
          <div className="flex items-center gap-2 text-slate-600 font-medium">
            <Layers className="w-4 h-4 text-indigo-600" />
            <span>
              Đã chọn <strong>{item.selectedSlideNumbers.length}</strong> / {item.slideCount} slide
            </span>
          </div>

          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={handleSelectAll}
              className="flex items-center gap-1 px-2.5 py-1 text-slate-700 hover:text-slate-900 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors cursor-pointer"
            >
              <CheckSquare className="w-3.5 h-3.5 text-rose-600" />
              <span>Chọn tất cả</span>
            </button>
            <button
              type="button"
              onClick={handleSelectOnlyFirst}
              className="flex items-center gap-1 px-2.5 py-1 text-slate-700 hover:text-slate-900 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors cursor-pointer"
            >
              <Square className="w-3.5 h-3.5 text-slate-400" />
              <span>Chỉ slide đầu</span>
            </button>
          </div>
        </div>

        {/* Slide List */}
        <div className="p-6 overflow-y-auto space-y-3 flex-1 bg-slate-50/30">
          {item.slides.length === 0 ? (
            <div className="text-center py-8 text-slate-400 text-sm">
              Không tìm thấy thông tin cấu trúc slide.
            </div>
          ) : (
            item.slides.map((slide) => {
              const isSelected = item.selectedSlideNumbers.includes(slide.slideNumber);
              return (
                <div
                  key={slide.slideNumber}
                  onClick={() => handleToggleSlide(slide.slideNumber)}
                  className={`flex items-start gap-4 p-4 rounded-2xl border transition-all cursor-pointer ${
                    isSelected
                      ? 'bg-white border-rose-200 shadow-xs'
                      : 'bg-slate-100/60 border-slate-200 opacity-60 hover:opacity-90'
                  }`}
                >
                  {/* Checkbox indicator */}
                  <div
                    className={`mt-0.5 w-6 h-6 rounded-lg flex items-center justify-center shrink-0 transition-colors ${
                      isSelected
                        ? 'bg-rose-600 text-white'
                        : 'border-2 border-slate-300 bg-white text-transparent'
                    }`}
                  >
                    <Check className="w-4 h-4 stroke-[3]" />
                  </div>

                  {/* Slide number mockup */}
                  <div className="w-12 h-9 rounded-lg bg-slate-900 text-amber-300 font-extrabold text-xs flex flex-col items-center justify-center shrink-0">
                    <span className="text-[9px] text-slate-400 uppercase font-semibold leading-none">Slide</span>
                    <span className="leading-none mt-0.5">{slide.slideNumber}</span>
                  </div>

                  {/* Slide content preview */}
                  <div className="min-w-0 flex-1">
                    <div className="flex items-center gap-2">
                      <h5 className="text-sm font-bold text-slate-900 truncate">
                        {slide.title || `Trang chiếu #${slide.slideNumber}`}
                      </h5>
                      {slide.hasImages && (
                        <span className="inline-flex items-center gap-1 text-[10px] text-emerald-700 bg-emerald-50 px-1.5 py-0.5 rounded font-medium">
                          <ImageIcon className="w-3 h-3" />
                          Có hình ảnh
                        </span>
                      )}
                    </div>

                    {slide.previewText.length > 0 ? (
                      <ul className="mt-1.5 space-y-0.5">
                        {slide.previewText.slice(0, 3).map((text, idx) => (
                          <li key={idx} className="text-xs text-slate-500 truncate flex items-center gap-1.5">
                            <span className="w-1 h-1 rounded-full bg-slate-300 shrink-0" />
                            <span className="truncate">{text}</span>
                          </li>
                        ))}
                      </ul>
                    ) : (
                      <p className="text-xs text-slate-400 italic mt-1">Nội dung đồ hoạ hoặc tiêu đề trống</p>
                    )}
                  </div>
                </div>
              );
            })
          )}
        </div>

        {/* Footer */}
        <div className="px-6 py-4 border-t border-slate-100 flex items-center justify-between bg-white">
          <p className="text-xs text-slate-500">
            Mẹo: Bấm vào từng slide để bật/tắt khi gộp.
          </p>
          <button
            type="button"
            onClick={onClose}
            className="px-5 py-2 text-sm font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded-xl transition-colors cursor-pointer"
          >
            Xong
          </button>
        </div>
      </div>
    </div>
  );
};

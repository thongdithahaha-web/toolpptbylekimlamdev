import React from 'react';
import {
  GripVertical,
  ChevronUp,
  ChevronDown,
  Trash2,
  SlidersHorizontal,
  Presentation,
  ChevronsUp,
  ChevronsDown,
  Loader2,
} from 'lucide-react';
import { PPTFileItem } from '../types.ts';
import { formatBytes } from '../utils/pptxParser.ts';

interface FileItemProps {
  item: PPTFileItem;
  index: number;
  totalItems: number;
  startSlideNumber: number;
  onMoveUp: (index: number) => void;
  onMoveDown: (index: number) => void;
  onMoveToTop: (index: number) => void;
  onMoveToBottom: (index: number) => void;
  onRemove: (id: string) => void;
  onOpenInspector: (item: PPTFileItem) => void;
  onDragStart: (e: React.DragEvent, index: number) => void;
  onDragOver: (e: React.DragEvent, index: number) => void;
  onDragEnd: (e: React.DragEvent) => void;
  isDragging: boolean;
}

export const FileItem: React.FC<FileItemProps> = ({
  item,
  index,
  totalItems,
  startSlideNumber,
  onMoveUp,
  onMoveDown,
  onMoveToTop,
  onMoveToBottom,
  onRemove,
  onOpenInspector,
  onDragStart,
  onDragOver,
  onDragEnd,
  isDragging,
}) => {
  const activeSlideCount = item.selectedSlideNumbers.length;
  const endSlideNumber = startSlideNumber + activeSlideCount - 1;

  const isFirst = index === 0;
  const isLast = index === totalItems - 1;

  return (
    <div
      draggable
      onDragStart={(e) => onDragStart(e, index)}
      onDragOver={(e) => onDragOver(e, index)}
      onDragEnd={onDragEnd}
      className={`group relative flex flex-col sm:flex-row sm:items-center justify-between p-4 bg-white rounded-2xl border transition-all duration-150 ${
        isDragging
          ? 'opacity-40 border-dashed border-rose-400 bg-rose-50/30'
          : 'border-slate-200/90 hover:border-slate-300 hover:shadow-md'
      }`}
    >
      {/* Left section: drag handle + index + file info */}
      <div className="flex items-center gap-3 min-w-0 flex-1">
        {/* Drag handle */}
        <div
          className="text-slate-400 hover:text-slate-700 cursor-grab active:cursor-grabbing p-1 rounded-md hover:bg-slate-100 transition-colors"
          title="Kéo thả để đổi thứ tự"
        >
          <GripVertical className="w-5 h-5" />
        </div>

        {/* Order index badge */}
        <div className="flex flex-col items-center justify-center min-w-[36px] h-9 px-2 bg-slate-900 text-white rounded-xl font-bold text-xs shadow-xs">
          <span className="text-[10px] text-slate-400 uppercase leading-none font-semibold">Vị trí</span>
          <span className="text-sm font-extrabold text-amber-300 leading-none mt-0.5">#{index + 1}</span>
        </div>

        {/* File icon */}
        <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-rose-500 to-orange-500 flex items-center justify-center text-white shrink-0 shadow-xs">
          <Presentation className="w-5 h-5" />
        </div>

        {/* File meta */}
        <div className="min-w-0 flex-1">
          <div className="flex items-center gap-2 flex-wrap">
            <h4
              className="text-sm font-bold text-slate-900 truncate max-w-[280px] sm:max-w-[360px]"
              title={item.name}
            >
              {item.name}
            </h4>
            {item.isAnalyzing && (
              <span className="inline-flex items-center gap-1 text-[11px] text-amber-600 bg-amber-50 px-2 py-0.5 rounded-md font-medium">
                <Loader2 className="w-3 h-3 animate-spin" />
                Đang đọc...
              </span>
            )}
          </div>

          <div className="flex items-center gap-2 text-xs text-slate-500 mt-1 flex-wrap">
            <span className="font-mono text-slate-600 font-medium">{formatBytes(item.size)}</span>
            <span>•</span>
            <span className="inline-flex items-center gap-1 font-semibold text-rose-600 bg-rose-50 px-2 py-0.5 rounded-md border border-rose-100">
              {activeSlideCount === item.slideCount
                ? `${item.slideCount} slide`
                : `${activeSlideCount}/${item.slideCount} slide đã chọn`}
            </span>
            <span>•</span>
            <span className="inline-flex items-center gap-1 font-semibold text-indigo-700 bg-indigo-50 px-2 py-0.5 rounded-md border border-indigo-100 font-mono">
              Trang đích: #{startSlideNumber} → #{endSlideNumber}
            </span>
          </div>
        </div>
      </div>

      {/* Right section: Reorder controls and actions */}
      <div className="flex items-center justify-end gap-1.5 mt-3 sm:mt-0 pt-2 sm:pt-0 border-t sm:border-t-0 border-slate-100 shrink-0">
        {/* Slide customizer button */}
        <button
          type="button"
          onClick={() => onOpenInspector(item)}
          className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors cursor-pointer mr-1"
          title="Xem trước tiêu đề và chọn từng slide"
        >
          <SlidersHorizontal className="w-3.5 h-3.5 text-slate-500" />
          <span className="hidden sm:inline">Chi tiết slide</span>
        </button>

        {/* Up / Down buttons */}
        <div className="flex items-center bg-slate-100 rounded-lg p-0.5 border border-slate-200">
          <button
            type="button"
            onClick={() => onMoveToTop(index)}
            disabled={isFirst}
            className="p-1 rounded text-slate-600 hover:text-slate-900 hover:bg-white disabled:opacity-30 disabled:pointer-events-none transition-all cursor-pointer"
            title="Đưa lên đầu danh sách"
          >
            <ChevronsUp className="w-4 h-4" />
          </button>
          <button
            type="button"
            onClick={() => onMoveUp(index)}
            disabled={isFirst}
            className="p-1 rounded text-slate-600 hover:text-slate-900 hover:bg-white disabled:opacity-30 disabled:pointer-events-none transition-all cursor-pointer"
            title="Di chuyển lên 1 bậc"
          >
            <ChevronUp className="w-4 h-4" />
          </button>
          <button
            type="button"
            onClick={() => onMoveDown(index)}
            disabled={isLast}
            className="p-1 rounded text-slate-600 hover:text-slate-900 hover:bg-white disabled:opacity-30 disabled:pointer-events-none transition-all cursor-pointer"
            title="Di chuyển xuống 1 bậc"
          >
            <ChevronDown className="w-4 h-4" />
          </button>
          <button
            type="button"
            onClick={() => onMoveToBottom(index)}
            disabled={isLast}
            className="p-1 rounded text-slate-600 hover:text-slate-900 hover:bg-white disabled:opacity-30 disabled:pointer-events-none transition-all cursor-pointer"
            title="Đưa xuống cuối danh sách"
          >
            <ChevronsDown className="w-4 h-4" />
          </button>
        </div>

        {/* Remove button */}
        <button
          type="button"
          onClick={() => onRemove(item.id)}
          className="p-1.5 rounded-lg text-slate-400 hover:text-red-600 hover:bg-red-50 transition-colors cursor-pointer ml-1"
          title="Xoá file này khỏi danh sách gộp"
        >
          <Trash2 className="w-4 h-4" />
        </button>
      </div>
    </div>
  );
};

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
  Star,
  Flag,
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
          : isFirst
          ? 'border-amber-300 ring-2 ring-amber-400/20 shadow-xs bg-gradient-to-r from-amber-50/20 to-white'
          : 'border-slate-200/90 hover:border-slate-300 hover:shadow-md'
      }`}
    >
      {/* Left section: drag handle + index + file info */}
      <div className="flex items-center gap-3 min-w-0 flex-1">
        {/* Drag handle */}
        <div
          className="text-slate-400 hover:text-slate-700 cursor-grab active:cursor-grabbing p-1 rounded-md hover:bg-slate-100 transition-colors"
          title="Kéo thả để thay đổi vị trí ưu tiên"
        >
          <GripVertical className="w-5 h-5" />
        </div>

        {/* Order index badge */}
        <div
          className={`flex flex-col items-center justify-center min-w-[42px] h-10 px-2 rounded-xl font-bold text-xs shadow-xs transition-colors ${
            isFirst
              ? 'bg-amber-500 text-slate-950 font-black'
              : 'bg-slate-900 text-white'
          }`}
          title={isFirst ? 'File ưu tiên số 1 - Luôn nằm ở đầu bài thuyết trình' : `File ở vị trí #${index + 1}`}
        >
          <span className="text-[9px] uppercase leading-none font-extrabold opacity-80">
            {isFirst ? 'Ưu tiên' : 'Vị trí'}
          </span>
          <span className="text-base font-black leading-none mt-0.5">
            #{index + 1}
          </span>
        </div>

        {/* File icon */}
        <div
          className={`w-10 h-10 rounded-xl flex items-center justify-center text-white shrink-0 shadow-xs ${
            isFirst
              ? 'bg-gradient-to-br from-amber-500 to-rose-500'
              : 'bg-gradient-to-br from-rose-500 to-orange-500'
          }`}
        >
          <Presentation className="w-5 h-5" />
        </div>

        {/* File meta */}
        <div className="min-w-0 flex-1">
          <div className="flex items-center gap-2 flex-wrap">
            <h4
              className="text-sm font-bold text-slate-900 truncate max-w-[260px] sm:max-w-[340px]"
              title={item.name}
            >
              {item.name}
            </h4>

            {isFirst && (
              <span className="inline-flex items-center gap-1 text-[11px] font-extrabold text-amber-800 bg-amber-100/90 px-2 py-0.5 rounded-md border border-amber-300">
                <Star className="w-3 h-3 text-amber-600 fill-amber-500" />
                <span>Ưu tiên số 1 (Mở đầu)</span>
              </span>
            )}

            {isLast && totalItems > 1 && (
              <span className="inline-flex items-center gap-1 text-[10px] font-bold text-slate-600 bg-slate-100 px-2 py-0.5 rounded-md border border-slate-200">
                <Flag className="w-3 h-3 text-slate-500" />
                <span>Vị trí cuối</span>
              </span>
            )}

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
            <span className={`inline-flex items-center gap-1 font-bold px-2 py-0.5 rounded-md border font-mono ${
              isFirst
                ? 'text-amber-900 bg-amber-50 border-amber-200'
                : 'text-indigo-700 bg-indigo-50 border-indigo-100'
            }`}>
              Trang sau khi gộp: #{startSlideNumber} → #{endSlideNumber}
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
            title="Đưa lên vị trí ưu tiên số 1"
          >
            <ChevronsUp className="w-4 h-4" />
          </button>
          <button
            type="button"
            onClick={() => onMoveUp(index)}
            disabled={isFirst}
            className="p-1 rounded text-slate-600 hover:text-slate-900 hover:bg-white disabled:opacity-30 disabled:pointer-events-none transition-all cursor-pointer"
            title="Tăng mức ưu tiên lên 1 bậc"
          >
            <ChevronUp className="w-4 h-4" />
          </button>
          <button
            type="button"
            onClick={() => onMoveDown(index)}
            disabled={isLast}
            className="p-1 rounded text-slate-600 hover:text-slate-900 hover:bg-white disabled:opacity-30 disabled:pointer-events-none transition-all cursor-pointer"
            title="Giảm mức ưu tiên xuống 1 bậc"
          >
            <ChevronDown className="w-4 h-4" />
          </button>
          <button
            type="button"
            onClick={() => onMoveToBottom(index)}
            disabled={isLast}
            className="p-1 rounded text-slate-600 hover:text-slate-900 hover:bg-white disabled:opacity-30 disabled:pointer-events-none transition-all cursor-pointer"
            title="Đưa xuống vị trí cuối cùng"
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

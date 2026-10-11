import React from 'react';
import { Loader2, CheckCircle2, AlertCircle } from 'lucide-react';

interface MergeModalProps {
  isOpen: boolean;
  currentStep: number;
  uploadProgress?: number;
  statusMessage?: string;
  error: string | null;
  onRetry: () => void;
  onCancel: () => void;
}

export const MergeModal: React.FC<MergeModalProps> = ({
  isOpen,
  currentStep,
  uploadProgress = 0,
  statusMessage,
  error,
  onRetry,
  onCancel,
}) => {
  if (!isOpen) return null;

  const steps = [
    { title: 'Tải lên & phân mảnh dữ liệu', desc: 'Đang truyền tải an toàn các tệp trình chiếu (tránh lỗi 413)...' },
    { title: 'Phân tích cấu trúc OpenXML', desc: 'Đọc hệ thống slide, hình ảnh và bố cục...' },
    { title: 'Ghép slide theo thứ tự đã định', desc: 'Sắp xếp chuẩn xác từng trang thuyết trình...' },
    { title: 'Đóng gói hoàn tất file PPTX', desc: 'Tối ưu hoá và hoàn thiện tệp PowerPoint...' },
  ];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs animate-in fade-in duration-200">
      <div className="bg-white rounded-3xl shadow-2xl border border-slate-200 w-full max-w-lg p-6 sm:p-8 overflow-hidden">
        {error ? (
          <div className="text-center py-4">
            <div className="w-16 h-16 rounded-2xl bg-red-100 text-red-600 flex items-center justify-center mx-auto mb-4">
              <AlertCircle className="w-8 h-8" />
            </div>
            <h3 className="text-xl font-bold text-slate-900 mb-2">Đã xảy ra lỗi khi gộp</h3>
            <p className="text-sm text-red-600 bg-red-50 p-4 rounded-xl border border-red-200 mb-6 text-left break-words">
              {error}
            </p>
            <div className="flex items-center justify-center gap-3">
              <button
                type="button"
                onClick={onCancel}
                className="px-5 py-2.5 text-sm font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-xl transition-colors cursor-pointer"
              >
                Đóng
              </button>
              <button
                type="button"
                onClick={onRetry}
                className="px-5 py-2.5 text-sm font-semibold text-white bg-rose-600 hover:bg-rose-700 rounded-xl transition-colors cursor-pointer"
              >
                Thử lại
              </button>
            </div>
          </div>
        ) : (
          <div>
            <div className="flex items-center justify-between mb-4">
              <div>
                <h3 className="text-lg sm:text-xl font-bold text-slate-900">
                  Đang tiến hành gộp file PowerPoint
                </h3>
                <p className="text-xs text-slate-500 mt-0.5">
                  {statusMessage || 'Vui lòng chờ trong giây lát...'}
                </p>
              </div>
              <div className="w-10 h-10 rounded-xl bg-rose-50 text-rose-600 flex items-center justify-center">
                <Loader2 className="w-6 h-6 animate-spin text-rose-600" />
              </div>
            </div>

            {/* Upload Progress Bar if in step 1 */}
            {currentStep === 1 && (
              <div className="my-4 p-3 bg-slate-50 rounded-2xl border border-slate-200/80">
                <div className="flex items-center justify-between text-xs font-semibold text-slate-700 mb-1.5">
                  <span>Tiến trình tải lên</span>
                  <span className="font-mono text-rose-600">{Math.round(uploadProgress)}%</span>
                </div>
                <div className="w-full h-2.5 bg-slate-200 rounded-full overflow-hidden">
                  <div
                    className="h-full bg-gradient-to-r from-rose-500 to-orange-500 transition-all duration-200 rounded-full"
                    style={{ width: `${Math.max(5, uploadProgress)}%` }}
                  />
                </div>
              </div>
            )}

            {/* Step list */}
            <div className="space-y-3.5 my-5">
              {steps.map((step, idx) => {
                const stepNum = idx + 1;
                const isCompleted = currentStep > stepNum;
                const isActive = currentStep === stepNum;

                return (
                  <div
                    key={idx}
                    className={`flex items-start gap-3.5 p-3 rounded-xl transition-all ${
                      isActive
                        ? 'bg-rose-50/70 border border-rose-200'
                        : isCompleted
                        ? 'bg-slate-50 opacity-80'
                        : 'opacity-40'
                    }`}
                  >
                    <div className="mt-0.5 shrink-0">
                      {isCompleted ? (
                        <CheckCircle2 className="w-5 h-5 text-emerald-600" />
                      ) : isActive ? (
                        <Loader2 className="w-5 h-5 animate-spin text-rose-600" />
                      ) : (
                        <div className="w-5 h-5 rounded-full border-2 border-slate-300 flex items-center justify-center text-[10px] font-bold text-slate-400">
                          {stepNum}
                        </div>
                      )}
                    </div>
                    <div className="min-w-0 flex-1">
                      <h4
                        className={`text-sm font-bold ${
                          isActive ? 'text-rose-950' : 'text-slate-800'
                        }`}
                      >
                        {step.title}
                      </h4>
                      <p className="text-xs text-slate-500 truncate">{step.desc}</p>
                    </div>
                  </div>
                );
              })}
            </div>

            <p className="text-center text-xs text-slate-400 italic">
              Công nghệ tải phân mảnh bảo toàn dung lượng lớn và giữ trọn vẹn chất lượng slide.
            </p>
          </div>
        )}
      </div>
    </div>
  );
};

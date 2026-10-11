import React from 'react';
import { ShieldCheck, BadgeCheck, Mail, User, Calendar, X, FileText, CheckCircle2 } from 'lucide-react';

interface CopyrightModalProps {
  isOpen: boolean;
  onClose: () => void;
}

export const CopyrightModal: React.FC<CopyrightModalProps> = ({ isOpen, onClose }) => {
  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs animate-in fade-in duration-200">
      <div className="bg-white rounded-3xl shadow-2xl border border-slate-200 w-full max-w-lg p-6 sm:p-8 overflow-hidden animate-in zoom-in-95 duration-200">
        <div className="flex items-start justify-between pb-4 border-b border-slate-100">
          <div className="flex items-center gap-3">
            <div className="w-12 h-12 rounded-2xl bg-gradient-to-tr from-emerald-600 to-teal-500 text-white flex items-center justify-center shadow-md shadow-emerald-600/20">
              <ShieldCheck className="w-6 h-6" />
            </div>
            <div>
              <div className="flex items-center gap-1.5">
                <span className="text-xs font-bold text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-full border border-emerald-200">
                  Chứng nhận bản quyền
                </span>
              </div>
              <h3 className="text-lg font-extrabold text-slate-900 mt-0.5 font-mono">
                toolgopfilepptbylekimlam
              </h3>
            </div>
          </div>

          <button
            type="button"
            onClick={onClose}
            className="p-1.5 rounded-xl text-slate-400 hover:text-slate-700 hover:bg-slate-100 transition-colors cursor-pointer"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="my-6 space-y-4 text-xs sm:text-sm text-slate-600">
          <div className="p-4 rounded-2xl bg-slate-50 border border-slate-200/80 space-y-3">
            <div className="flex items-center justify-between">
              <span className="text-slate-500 font-medium flex items-center gap-2">
                <User className="w-4 h-4 text-rose-600" />
                Tác giả & Nhà phát triển:
              </span>
              <span className="font-bold text-slate-900 flex items-center gap-1">
                Lê Kim Lâm
                <BadgeCheck className="w-4 h-4 text-emerald-600" />
              </span>
            </div>

            <div className="flex items-center justify-between">
              <span className="text-slate-500 font-medium flex items-center gap-2">
                <Mail className="w-4 h-4 text-rose-600" />
                Email liên hệ:
              </span>
              <a
                href="mailto:lekimlam2707@gmail.com"
                className="font-bold text-rose-600 hover:underline font-mono"
              >
                lekimlam2707@gmail.com
              </a>
            </div>

            <div className="flex items-center justify-between">
              <span className="text-slate-500 font-medium flex items-center gap-2">
                <Calendar className="w-4 h-4 text-rose-600" />
                Năm phát hành bản quyền:
              </span>
              <span className="font-bold text-slate-900 font-mono">2026</span>
            </div>

            <div className="flex items-center justify-between">
              <span className="text-slate-500 font-medium flex items-center gap-2">
                <FileText className="w-4 h-4 text-rose-600" />
                Tình trạng bản quyền:
              </span>
              <span className="font-bold text-emerald-700">All Rights Reserved (Toàn quyền bảo lưu)</span>
            </div>
          </div>

          <div className="space-y-2 text-xs leading-relaxed text-slate-500 p-3 rounded-xl bg-emerald-50/60 border border-emerald-100">
            <p className="flex items-start gap-2 text-emerald-950 font-medium">
              <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
              <span>
                Trang web <strong>toolgopfilepptbylekimlam</strong> được xây dựng và sở hữu bởi <strong>Lê Kim Lâm</strong> nhằm cung cấp giải pháp gộp và xử lý slide trình chiếu PowerPoint an toàn, nhanh chóng và giữ nguyên bố cục.
              </span>
            </p>
            <p className="flex items-start gap-2 text-emerald-950 font-medium">
              <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
              <span>
                Toàn bộ dữ liệu của người dùng được mã hoá tạm thời và tự động xoá sau khi hoàn tất phiên làm việc, cam kết không lưu trữ hoặc chia sẻ tệp tin của người dùng.
              </span>
            </p>
          </div>
        </div>

        <div className="pt-2 text-center">
          <button
            type="button"
            onClick={onClose}
            className="w-full py-3 bg-slate-900 hover:bg-slate-800 text-white font-bold text-sm rounded-xl transition-all cursor-pointer"
          >
            Đã hiểu & Tiếp tục sử dụng
          </button>
        </div>
      </div>
    </div>
  );
};

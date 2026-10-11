import React from 'react';
import { Upload, ArrowDownUp, Download, HelpCircle, CheckCircle, ShieldCheck } from 'lucide-react';

export const GuideSection: React.FC = () => {
  return (
    <div className="mt-16 space-y-12">
      {/* 3 Step Workflow */}
      <div>
        <div className="text-center mb-8">
          <h3 className="text-xl sm:text-2xl font-extrabold text-slate-900">
            Cách thức gộp PowerPoint đơn giản trong 3 bước
          </h3>
          <p className="text-sm text-slate-500 mt-1">
            Không cần cài đặt phần mềm phức tạp, xử lý trực tiếp trên trình duyệt
          </p>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          <div className="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs relative">
            <div className="w-10 h-10 rounded-xl bg-rose-100 text-rose-600 flex items-center justify-center font-extrabold text-sm mb-4">
              1
            </div>
            <h4 className="font-bold text-slate-900 mb-2 flex items-center gap-2">
              <Upload className="w-4 h-4 text-rose-600" />
              Tải các file PPTX lên
            </h4>
            <p className="text-xs text-slate-600 leading-relaxed">
              Kéo thả hoặc chọn nhiều bài thuyết trình PowerPoint (.pptx). Hệ thống sẽ tự động đọc số slide và tiêu đề tệp.
            </p>
          </div>

          <div className="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs relative">
            <div className="w-10 h-10 rounded-xl bg-amber-100 text-amber-700 flex items-center justify-center font-extrabold text-sm mb-4">
              2
            </div>
            <h4 className="font-bold text-slate-900 mb-2 flex items-center gap-2">
              <ArrowDownUp className="w-4 h-4 text-amber-600" />
              Sắp xếp theo thứ tự mong muốn
            </h4>
            <p className="text-xs text-slate-600 leading-relaxed">
              Dùng chuột kéo thả hoặc bấm nút mũi tên Lên/Xuống để đặt vị trí. Bạn có thể chọn chi tiết từng slide cần ghép.
            </p>
          </div>

          <div className="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs relative">
            <div className="w-10 h-10 rounded-xl bg-emerald-100 text-emerald-700 flex items-center justify-center font-extrabold text-sm mb-4">
              3
            </div>
            <h4 className="font-bold text-slate-900 mb-2 flex items-center gap-2">
              <Download className="w-4 h-4 text-emerald-600" />
              Bấm Gộp & Tải về ngay
            </h4>
            <p className="text-xs text-slate-600 leading-relaxed">
              Hệ thống xử lý nối toàn bộ slide vào 1 file duy nhất với đầy đủ hình ảnh, bảng biểu và hiệu ứng ban đầu.
            </p>
          </div>
        </div>
      </div>

      {/* Highlights / FAQ */}
      <div className="bg-slate-900 text-white rounded-3xl p-6 sm:p-10 shadow-lg">
        <div className="flex items-center gap-3 mb-6">
          <div className="w-10 h-10 rounded-xl bg-rose-600 flex items-center justify-center text-white">
            <HelpCircle className="w-5 h-5" />
          </div>
          <div>
            <h3 className="text-lg sm:text-xl font-bold">Các câu hỏi thường gặp</h3>
            <p className="text-xs text-slate-400">Giải đáp về cách xử lý và định dạng tệp</p>
          </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 text-sm">
          <div className="space-y-1.5 p-4 rounded-xl bg-slate-800/60 border border-slate-700/50">
            <h5 className="font-bold text-amber-300 flex items-center gap-2">
              <CheckCircle className="w-4 h-4 text-emerald-400" />
              Thứ tự sau khi gộp được xác định như thế nào?
            </h5>
            <p className="text-xs text-slate-300 leading-relaxed">
              File ở vị trí #1 sẽ xuất hiện đầu tiên (Slide 1, 2...), tiếp theo là toàn bộ slide của File #2, File #3... Bạn có thể tự do thay đổi thứ tự trước khi bấm Gộp.
            </p>
          </div>

          <div className="space-y-1.5 p-4 rounded-xl bg-slate-800/60 border border-slate-700/50">
            <h5 className="font-bold text-amber-300 flex items-center gap-2">
              <CheckCircle className="w-4 h-4 text-emerald-400" />
              Định dạng hình ảnh, video và font chữ có bị mất không?
            </h5>
            <p className="text-xs text-slate-300 leading-relaxed">
              Không. Công cụ kế thừa cấu trúc XML gốc của PowerPoint, bảo toàn độ phân giải ảnh, bố cục chữ và màu nền của từng trang chiếu.
            </p>
          </div>

          <div className="space-y-1.5 p-4 rounded-xl bg-slate-800/60 border border-slate-700/50">
            <h5 className="font-bold text-amber-300 flex items-center gap-2">
              <CheckCircle className="w-4 h-4 text-emerald-400" />
              Tôi có thể chọn lọc chỉ lấy một vài slide trong file không?
            </h5>
            <p className="text-xs text-slate-300 leading-relaxed">
              Có. Bấm vào nút <strong>"Chi tiết slide"</strong> trên từng file để bật/tắt các slide bạn muốn hoặc không muốn đưa vào file gộp.
            </p>
          </div>

          <div className="space-y-1.5 p-4 rounded-xl bg-slate-800/60 border border-slate-700/50">
            <h5 className="font-bold text-amber-300 flex items-center gap-2">
              <ShieldCheck className="w-4 h-4 text-emerald-400" />
              Dữ liệu của tôi có được bảo mật không?
            </h5>
            <p className="text-xs text-slate-300 leading-relaxed">
              Các tệp tải lên chỉ được xử lý tạm thời trong phiên làm việc của bạn và được tự động xoá hoàn toàn khỏi bộ nhớ ngay khi quá trình gộp hoàn tất.
            </p>
          </div>
        </div>
      </div>
    </div>
  );
};

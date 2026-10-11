import React, { useState, useRef } from 'react';
import {
  FileText,
  Presentation,
  FileSpreadsheet,
  FileType,
  ArrowRight,
  Download,
  RotateCcw,
  Sparkles,
  CheckCircle2,
  AlertCircle,
  Loader2,
  UploadCloud,
  FileCode,
  ShieldCheck,
  Zap,
} from 'lucide-react';
import {
  convertPptxToPdf,
  convertPptxToDocx,
  convertDocxToPptx,
  convertTxtToPptx,
  convertPptxToTxt,
  convertPptxToHtml,
  convertDocxToPdf,
  ConversionResult,
} from '../utils/documentConverter.ts';
import { formatBytes } from '../utils/pptxParser.ts';

type ConvertType =
  | 'pptx-to-pdf'
  | 'pptx-to-docx'
  | 'docx-to-pptx'
  | 'txt-to-pptx'
  | 'pptx-to-txt'
  | 'pptx-to-html'
  | 'docx-to-pdf';

interface PresetOption {
  id: ConvertType;
  title: string;
  desc: string;
  fromExt: string;
  toExt: string;
  fromIcon: any;
  toIcon: any;
  badge: string;
  accept: string;
}

const PRESETS: PresetOption[] = [
  {
    id: 'pptx-to-pdf',
    title: 'PowerPoint sang PDF',
    desc: 'Chuyển bài thuyết trình thành tệp PDF trang ngang sắc nét, sẵn sàng in ấn.',
    fromExt: '.pptx',
    toExt: '.pdf',
    fromIcon: Presentation,
    toIcon: FileText,
    badge: 'Phổ biến nhất',
    accept: '.pptx,.ppt',
  },
  {
    id: 'pptx-to-docx',
    title: 'PowerPoint sang Word',
    desc: 'Trích xuất toàn bộ slide thành bài báo cáo tài liệu Word (.docx) chuẩn.',
    fromExt: '.pptx',
    toExt: '.docx',
    fromIcon: Presentation,
    toIcon: FileType,
    badge: 'Trích xuất báo cáo',
    accept: '.pptx,.ppt',
  },
  {
    id: 'docx-to-pptx',
    title: 'Word sang PowerPoint',
    desc: 'Tự động đọc các tiêu đề, đề mục Word và tạo slide thuyết trình tương ứng.',
    fromExt: '.docx',
    toExt: '.pptx',
    fromIcon: FileType,
    toIcon: Presentation,
    badge: 'Tạo slide tự động',
    accept: '.docx',
  },
  {
    id: 'txt-to-pptx',
    title: 'Văn bản (TXT/MD) sang PowerPoint',
    desc: 'Chuyển dàn ý ghi chú văn bản thuần thành bài thuyết trình PowerPoint.',
    fromExt: '.txt, .md',
    toExt: '.pptx',
    fromIcon: FileCode,
    toIcon: Presentation,
    badge: 'Soạn slide siêu tốc',
    accept: '.txt,.md,.text',
  },
  {
    id: 'pptx-to-html',
    title: 'PowerPoint sang Web Slide (HTML)',
    desc: 'Tạo bài thuyết trình Web độc lập mở được trên mọi điện thoại & máy tính.',
    fromExt: '.pptx',
    toExt: '.html',
    fromIcon: Presentation,
    toIcon: FileCode,
    badge: 'Trình chiếu di động',
    accept: '.pptx,.ppt',
  },
  {
    id: 'pptx-to-txt',
    title: 'PowerPoint sang Văn bản thuần',
    desc: 'Trích xuất toàn bộ chữ, ghi chú và dàn ý slide ra tệp .txt chỉnh sửa.',
    fromExt: '.pptx',
    toExt: '.txt',
    fromIcon: Presentation,
    toIcon: FileText,
    badge: 'Lấy nội dung',
    accept: '.pptx,.ppt',
  },
  {
    id: 'docx-to-pdf',
    title: 'Word sang PDF',
    desc: 'Chuyển đổi tệp Word (.docx) sang định dạng PDF để khóa nội dung.',
    fromExt: '.docx',
    toExt: '.pdf',
    fromIcon: FileType,
    toIcon: FileText,
    badge: 'Bảo mật in ấn',
    accept: '.docx',
  },
];

export const ConverterTab: React.FC = () => {
  const [selectedPreset, setSelectedPreset] = useState<ConvertType>('pptx-to-pdf');
  const [uploadedFile, setUploadedFile] = useState<File | null>(null);
  const [isConverting, setIsConverting] = useState(false);
  const [conversionResult, setConversionResult] = useState<ConversionResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const currentPreset = PRESETS.find((p) => p.id === selectedPreset) || PRESETS[0];

  const handleSelectPreset = (preset: PresetOption) => {
    setSelectedPreset(preset.id);
    setUploadedFile(null);
    setConversionResult(null);
    setError(null);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      setUploadedFile(file);
      setConversionResult(null);
      setError(null);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      const ext = file.name.split('.').pop()?.toLowerCase();

      // Auto-detect preset if file extension matches
      if (ext === 'docx') {
        setSelectedPreset('docx-to-pptx');
      } else if (ext === 'txt' || ext === 'md') {
        setSelectedPreset('txt-to-pptx');
      } else if (ext === 'pptx' || ext === 'ppt') {
        setSelectedPreset('pptx-to-pdf');
      }

      setUploadedFile(file);
      setConversionResult(null);
      setError(null);
    }
  };

  const handleExecuteConvert = async () => {
    if (!uploadedFile) {
      setError('Vui lòng chọn một tệp để bắt đầu chuyển đổi.');
      return;
    }

    setIsConverting(true);
    setError(null);

    try {
      let res: ConversionResult;

      switch (selectedPreset) {
        case 'pptx-to-pdf':
          res = await convertPptxToPdf(uploadedFile);
          break;
        case 'pptx-to-docx':
          res = await convertPptxToDocx(uploadedFile);
          break;
        case 'docx-to-pptx':
          res = await convertDocxToPptx(uploadedFile);
          break;
        case 'txt-to-pptx':
          res = await convertTxtToPptx(uploadedFile);
          break;
        case 'pptx-to-txt':
          res = await convertPptxToTxt(uploadedFile);
          break;
        case 'pptx-to-html':
          res = await convertPptxToHtml(uploadedFile);
          break;
        case 'docx-to-pdf':
          res = await convertDocxToPdf(uploadedFile);
          break;
        default:
          throw new Error('Chế độ chuyển đổi không hợp lệ.');
      }

      setConversionResult(res);
    } catch (err: any) {
      console.error('Conversion failed:', err);
      setError(err.message || 'Chuyển đổi tệp gặp sự cố ngoài ý muốn.');
    } finally {
      setIsConverting(false);
    }
  };

  const handleDownloadResult = () => {
    if (!conversionResult) return;
    const url = URL.createObjectURL(conversionResult.blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = conversionResult.filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  const handleReset = () => {
    setUploadedFile(null);
    setConversionResult(null);
    setError(null);
  };

  return (
    <div className="space-y-8 animate-in fade-in duration-200">
      {/* Intro Banner */}
      <div className="bg-gradient-to-r from-rose-900 via-slate-900 to-indigo-950 text-white rounded-3xl p-6 sm:p-8 shadow-md">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
          <div className="space-y-2 max-w-2xl">
            <div className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-rose-500/20 text-rose-300 text-xs font-bold border border-rose-500/30">
              <Zap className="w-3.5 h-3.5 text-amber-300" />
              <span>Chuyển đổi tức thì 100% trong trình duyệt</span>
            </div>
            <h2 className="text-2xl sm:text-3xl font-extrabold tracking-tight">
              Chuyển Đổi Định Dạng Đa Năng
            </h2>
            <p className="text-xs sm:text-sm text-slate-300 leading-relaxed">
              Hỗ trợ chuyển đổi qua lại giữa tất cả các định dạng tài liệu phổ biến nhất: 
              <strong> PowerPoint (.pptx)</strong>, <strong>Word (.docx)</strong>, <strong>PDF</strong>, <strong>Văn bản (.txt)</strong> và <strong>Trình chiếu Web (HTML)</strong> mà không gửi tệp lên máy chủ.
            </p>
          </div>

          <div className="flex items-center gap-2 p-3 bg-white/10 rounded-2xl border border-white/15 backdrop-blur-xs text-xs text-slate-200 shrink-0">
            <ShieldCheck className="w-5 h-5 text-emerald-400 shrink-0" />
            <div>
              <span className="font-bold block text-white">Bảo mật tuyệt đối</span>
              <span className="text-[11px] text-slate-300">Không gửi file lên server</span>
            </div>
          </div>
        </div>
      </div>

      {/* Preset Grid Selection */}
      <div className="space-y-3">
        <div className="flex items-center justify-between px-1">
          <h3 className="text-sm font-extrabold text-slate-900 uppercase tracking-wider">
            Chọn dịch vụ chuyển đổi:
          </h3>
          <span className="text-xs text-slate-500">7 định dạng chuyển đổi phổ biến</span>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
          {PRESETS.map((preset) => {
            const isSelected = selectedPreset === preset.id;
            const FromIcon = preset.fromIcon;
            const ToIcon = preset.toIcon;

            return (
              <button
                key={preset.id}
                type="button"
                onClick={() => handleSelectPreset(preset)}
                className={`p-4 rounded-2xl border text-left transition-all cursor-pointer relative flex flex-col justify-between ${
                  isSelected
                    ? 'bg-rose-50/70 border-rose-500 ring-2 ring-rose-500/20 shadow-sm'
                    : 'bg-white border-slate-200/90 hover:border-slate-300 hover:shadow-xs'
                }`}
              >
                <div>
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-[10px] font-extrabold uppercase px-2 py-0.5 rounded-md bg-slate-100 text-slate-700 font-mono">
                      {preset.badge}
                    </span>
                    <div className="flex items-center gap-1.5 text-xs font-bold text-slate-700">
                      <span className="text-rose-600 font-mono">{preset.fromExt}</span>
                      <ArrowRight className="w-3.5 h-3.5 text-slate-400" />
                      <span className="text-indigo-600 font-mono">{preset.toExt}</span>
                    </div>
                  </div>

                  <h4 className="font-bold text-sm text-slate-900 mb-1 flex items-center gap-2">
                    <FromIcon className="w-4 h-4 text-rose-600 shrink-0" />
                    <span>{preset.title}</span>
                  </h4>
                  <p className="text-xs text-slate-500 leading-relaxed">{preset.desc}</p>
                </div>
              </button>
            );
          })}
        </div>
      </div>

      {/* File Upload & Converter Action Box */}
      <div className="bg-white rounded-3xl p-6 sm:p-8 border border-slate-200/90 shadow-sm">
        {error && (
          <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-2xl flex items-start gap-3 text-red-700 text-sm">
            <AlertCircle className="w-5 h-5 text-red-600 shrink-0 mt-0.5" />
            <div className="flex-1">
              <span className="font-bold block">Không thể chuyển đổi:</span>
              <span>{error}</span>
            </div>
            <button
              type="button"
              onClick={() => setError(null)}
              className="text-xs font-semibold underline hover:text-red-900"
            >
              Đóng
            </button>
          </div>
        )}

        {conversionResult ? (
          /* Conversion Success State */
          <div className="text-center py-6 animate-in zoom-in-95 duration-150">
            <div className="w-16 h-16 rounded-2xl bg-gradient-to-tr from-emerald-500 to-teal-400 text-white flex items-center justify-center mx-auto mb-4 shadow-md shadow-emerald-500/20">
              <CheckCircle2 className="w-8 h-8" />
            </div>

            <div className="inline-flex items-center gap-1 px-3 py-1 rounded-full bg-emerald-50 text-emerald-700 text-xs font-bold border border-emerald-200 mb-2">
              <Sparkles className="w-3.5 h-3.5" />
              <span>Chuyển đổi hoàn tất thành công!</span>
            </div>

            <h3 className="text-xl sm:text-2xl font-extrabold text-slate-900 mb-2">
              Tệp {conversionResult.format} của bạn đã sẵn sàng
            </h3>
            <p className="text-xs sm:text-sm text-slate-600 mb-6 max-w-md mx-auto">
              Định dạng đã được tạo chuẩn xác từ tệp gốc với dung lượng tối ưu.
            </p>

            <div className="p-4 bg-slate-50 rounded-2xl border border-slate-200 max-w-md mx-auto mb-6 text-left flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-indigo-600 text-white flex items-center justify-center shrink-0">
                <FileText className="w-5 h-5" />
              </div>
              <div className="min-w-0 flex-1">
                <h5 className="text-sm font-bold text-slate-900 truncate" title={conversionResult.filename}>
                  {conversionResult.filename}
                </h5>
                <div className="flex items-center gap-2 text-xs text-slate-500 mt-0.5 font-mono">
                  <span>{formatBytes(conversionResult.blob.size)}</span>
                  <span>•</span>
                  <span>{conversionResult.itemCount} mục/trang</span>
                </div>
              </div>
            </div>

            <div className="flex flex-col sm:flex-row items-center justify-center gap-3 max-w-md mx-auto">
              <button
                type="button"
                onClick={handleDownloadResult}
                className="w-full sm:flex-1 py-3.5 px-6 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 text-white font-extrabold text-sm rounded-xl shadow-md shadow-emerald-600/20 transition-all flex items-center justify-center gap-2 cursor-pointer active:scale-98"
              >
                <Download className="w-4 h-4" />
                <span>TẢI VỀ MÁY TÍNH</span>
              </button>

              <button
                type="button"
                onClick={handleReset}
                className="w-full sm:w-auto py-3.5 px-5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold text-sm rounded-xl transition-colors flex items-center justify-center gap-1.5 cursor-pointer"
              >
                <RotateCcw className="w-4 h-4" />
                <span>Chuyển file khác</span>
              </button>
            </div>
          </div>
        ) : (
          /* Upload and Convert Trigger State */
          <div className="space-y-6">
            <input
              ref={fileInputRef}
              type="file"
              accept={currentPreset.accept}
              onChange={handleFileChange}
              className="hidden"
            />

            <div
              onDragOver={(e) => e.preventDefault()}
              onDrop={handleDrop}
              onClick={() => fileInputRef.current?.click()}
              className="border-2 border-dashed border-slate-300 hover:border-rose-400 bg-slate-50/50 hover:bg-white rounded-3xl p-8 sm:p-10 text-center transition-all cursor-pointer"
            >
              <div className="w-16 h-16 rounded-2xl bg-rose-100 text-rose-600 flex items-center justify-center mx-auto mb-4">
                <UploadCloud className="w-8 h-8" />
              </div>

              {uploadedFile ? (
                <div>
                  <span className="text-xs font-bold text-emerald-700 bg-emerald-50 px-2.5 py-1 rounded-full border border-emerald-200">
                    Đã chọn tệp thành công
                  </span>
                  <h4 className="text-base font-bold text-slate-900 mt-2 truncate max-w-md mx-auto">
                    {uploadedFile.name}
                  </h4>
                  <p className="text-xs text-slate-500 font-mono mt-1">
                    {formatBytes(uploadedFile.size)} • Bấm để chọn tệp khác
                  </p>
                </div>
              ) : (
                <div>
                  <h4 className="text-base sm:text-lg font-bold text-slate-900 mb-1">
                    Kéo thả tệp {currentPreset.fromExt} vào đây
                  </h4>
                  <p className="text-xs sm:text-sm text-slate-500 mb-4">
                    hoặc bấm để chọn tệp từ thiết bị của bạn
                  </p>
                  <button
                    type="button"
                    className="px-5 py-2.5 bg-slate-900 hover:bg-slate-800 text-white font-bold text-xs rounded-xl shadow-xs transition-all pointer-events-none"
                  >
                    Chọn tệp {currentPreset.fromExt}
                  </button>
                </div>
              )}
            </div>

            {/* Convert Trigger Button */}
            <div className="pt-2">
              <button
                type="button"
                onClick={handleExecuteConvert}
                disabled={!uploadedFile || isConverting}
                className={`w-full py-4 px-6 rounded-2xl font-extrabold text-base flex items-center justify-center gap-2.5 transition-all shadow-md cursor-pointer ${
                  !uploadedFile || isConverting
                    ? 'bg-slate-200 text-slate-400 cursor-not-allowed shadow-none'
                    : 'bg-gradient-to-r from-rose-600 to-orange-600 hover:from-rose-700 hover:to-orange-700 text-white shadow-rose-600/25 active:scale-99'
                }`}
              >
                {isConverting ? (
                  <>
                    <Loader2 className="w-5 h-5 animate-spin" />
                    <span>Đang tiến hành chuyển đổi định dạng...</span>
                  </>
                ) : (
                  <>
                    <span>TIẾN HÀNH CHUYỂN ĐỔI SANG {currentPreset.toExt.toUpperCase()}</span>
                    <Sparkles className="w-4 h-4 text-amber-300" />
                  </>
                )}
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

import React, { useState, useRef, useMemo } from 'react';
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
  Image as ImageIcon,
  FolderArchive,
  Search,
  Check,
} from 'lucide-react';
import {
  convertPptxToPdf,
  convertPptxToHandoutPdf,
  convertPptxToDocx,
  convertDocxToPptx,
  convertTxtToPptx,
  convertPptxToTxt,
  convertPptxToHtml,
  convertPptxExtractMedia,
  convertPptxToJson,
  convertPptxToCsv,
  convertDocxToPdf,
  convertDocxExtractMedia,
  convertTxtToDocx,
  convertTxtToPdf,
  convertExcelToPptx,
  convertExcelToPdf,
  convertImagesToPptx,
  convertImagesToPdf,
  ConversionResult,
} from '../utils/documentConverter.ts';
import { formatBytes } from '../utils/pptxParser.ts';

export type CategoryKey = 'all' | 'pptx' | 'docx' | 'excel' | 'text' | 'image';

export interface ConverterService {
  id: string;
  category: CategoryKey;
  title: string;
  desc: string;
  fromExt: string;
  toExt: string;
  fromIcon: any;
  toIcon: any;
  badge: string;
  accept: string;
  action: (file: File) => Promise<ConversionResult>;
}

const SERVICES: ConverterService[] = [
  // 1. PPTX Services
  {
    id: 'pptx-to-pdf',
    category: 'pptx',
    title: 'PowerPoint sang PDF chuẩn',
    desc: 'Xuất bài thuyết trình thành tệp PDF trang ngang sắc nét, chuẩn in ấn và nộp bài.',
    fromExt: '.pptx',
    toExt: '.pdf',
    fromIcon: Presentation,
    toIcon: FileText,
    badge: 'Phổ biến nhất',
    accept: '.pptx,.ppt',
    action: convertPptxToPdf,
  },
  {
    id: 'pptx-to-handout-pdf',
    category: 'pptx',
    title: 'PowerPoint sang PDF Handout (3 slide/trang)',
    desc: 'Tạo tài liệu tóm tắt dạng Handout phát tay với 3 slide và các dòng kẻ ghi chú mỗi trang.',
    fromExt: '.pptx',
    toExt: '.pdf',
    fromIcon: Presentation,
    toIcon: FileText,
    badge: 'In báo cáo phát tay',
    accept: '.pptx,.ppt',
    action: convertPptxToHandoutPdf,
  },
  {
    id: 'pptx-to-docx',
    category: 'pptx',
    title: 'PowerPoint sang Word (.docx)',
    desc: 'Trích xuất toàn bộ slide thành bài báo cáo tài liệu Word chuyên nghiệp có tiêu đề và mục lục.',
    fromExt: '.pptx',
    toExt: '.docx',
    fromIcon: Presentation,
    toIcon: FileType,
    badge: 'Trích xuất báo cáo',
    accept: '.pptx,.ppt',
    action: convertPptxToDocx,
  },
  {
    id: 'pptx-extract-media',
    category: 'pptx',
    title: 'Tách tất cả hình ảnh từ PowerPoint (.zip)',
    desc: 'Tách toàn bộ tệp hình ảnh, sơ đồ gốc có trong slide ra một tệp nén ZIP độ nét cao.',
    fromExt: '.pptx',
    toExt: '.zip',
    fromIcon: Presentation,
    toIcon: FolderArchive,
    badge: 'Lấy ảnh gốc',
    accept: '.pptx,.ppt',
    action: convertPptxExtractMedia,
  },
  {
    id: 'pptx-to-html',
    category: 'pptx',
    title: 'PowerPoint sang Web Slide (.html)',
    desc: 'Tạo bài thuyết trình Web độc lập chạy mượt trên mọi điện thoại & máy tính không cần Office.',
    fromExt: '.pptx',
    toExt: '.html',
    fromIcon: Presentation,
    toIcon: FileCode,
    badge: 'Trình chiếu di động',
    accept: '.pptx,.ppt',
    action: convertPptxToHtml,
  },
  {
    id: 'pptx-to-txt',
    category: 'pptx',
    title: 'PowerPoint sang Văn bản (.txt)',
    desc: 'Trích xuất toàn bộ chữ, ghi chú và dàn ý slide ra tệp văn bản thuần để biên tập.',
    fromExt: '.pptx',
    toExt: '.txt',
    fromIcon: Presentation,
    toIcon: FileText,
    badge: 'Lấy văn bản',
    accept: '.pptx,.ppt',
    action: (f) => convertPptxToTxt(f, false),
  },
  {
    id: 'pptx-to-md',
    category: 'pptx',
    title: 'PowerPoint sang Markdown (.md)',
    desc: 'Chuyển đổi bài thuyết trình sang cấu trúc Markdown cho lập trình viên & Notion.',
    fromExt: '.pptx',
    toExt: '.md',
    fromIcon: Presentation,
    toIcon: FileCode,
    badge: 'Dành cho Notion/Dev',
    accept: '.pptx,.ppt',
    action: (f) => convertPptxToTxt(f, true),
  },
  {
    id: 'pptx-to-json',
    category: 'pptx',
    title: 'PowerPoint sang Dữ liệu JSON',
    desc: 'Xuất toàn bộ cấu trúc slide thành tệp JSON có phân cấp dữ liệu chi tiết.',
    fromExt: '.pptx',
    toExt: '.json',
    fromIcon: Presentation,
    toIcon: FileCode,
    badge: 'Dữ liệu lập trình',
    accept: '.pptx,.ppt',
    action: convertPptxToJson,
  },
  {
    id: 'pptx-to-csv',
    category: 'pptx',
    title: 'PowerPoint sang Bảng tính CSV',
    desc: 'Tổng hợp danh sách các trang và nội dung vào bảng tính Excel / CSV.',
    fromExt: '.pptx',
    toExt: '.csv',
    fromIcon: Presentation,
    toIcon: FileSpreadsheet,
    badge: 'Bảng tính nội dung',
    accept: '.pptx,.ppt',
    action: convertPptxToCsv,
  },

  // 2. Word (DOCX) Services
  {
    id: 'docx-to-pptx',
    category: 'docx',
    title: 'Word (.docx) sang PowerPoint',
    desc: 'Tự động đọc các tiêu đề, đề mục Heading của Word để tạo các slide thuyết trình tương ứng.',
    fromExt: '.docx',
    toExt: '.pptx',
    fromIcon: FileType,
    toIcon: Presentation,
    badge: 'Tạo slide tự động',
    accept: '.docx',
    action: convertDocxToPptx,
  },
  {
    id: 'docx-to-pdf',
    category: 'docx',
    title: 'Word (.docx) sang PDF',
    desc: 'Chuyển đổi tài liệu Word sang PDF chuẩn để khóa chỉnh sửa và in ấn.',
    fromExt: '.docx',
    toExt: '.pdf',
    fromIcon: FileType,
    toIcon: FileText,
    badge: 'Khóa văn bản',
    accept: '.docx',
    action: convertDocxToPdf,
  },
  {
    id: 'docx-extract-media',
    category: 'docx',
    title: 'Tách tất cả hình ảnh từ Word (.zip)',
    desc: 'Trích xuất toàn bộ hình ảnh được chèn trong file Word thành một tệp ZIP gốc.',
    fromExt: '.docx',
    toExt: '.zip',
    fromIcon: FileType,
    toIcon: FolderArchive,
    badge: 'Lấy ảnh Word',
    accept: '.docx',
    action: convertDocxExtractMedia,
  },

  // 3. Text / Markdown Services
  {
    id: 'txt-to-pptx',
    category: 'text',
    title: 'Văn bản (TXT/MD) sang PowerPoint',
    desc: 'Chuyển dàn ý ghi chú văn bản thuần thành các trang slide thuyết trình hoàn chỉnh.',
    fromExt: '.txt, .md',
    toExt: '.pptx',
    fromIcon: FileCode,
    toIcon: Presentation,
    badge: 'Soạn slide siêu tốc',
    accept: '.txt,.md,.text',
    action: convertTxtToPptx,
  },
  {
    id: 'txt-to-docx',
    category: 'text',
    title: 'Văn bản (TXT/MD) sang Word (.docx)',
    desc: 'Định dạng dàn ý văn bản thành tệp tài liệu Word với các cấp tiêu đề Heading 1.',
    fromExt: '.txt, .md',
    toExt: '.docx',
    fromIcon: FileCode,
    toIcon: FileType,
    badge: 'Đóng gói văn bản',
    accept: '.txt,.md,.text',
    action: convertTxtToDocx,
  },
  {
    id: 'txt-to-pdf',
    category: 'text',
    title: 'Văn bản (TXT/MD) sang PDF',
    desc: 'Xuất các ghi chú văn bản thành tài liệu PDF có đánh số trang và canh lề đẹp mắt.',
    fromExt: '.txt, .md',
    toExt: '.pdf',
    fromIcon: FileCode,
    toIcon: FileText,
    badge: 'Văn bản sang PDF',
    accept: '.txt,.md,.text',
    action: convertTxtToPdf,
  },

  // 4. Excel / Spreadsheet Services
  {
    id: 'excel-to-pptx',
    category: 'excel',
    title: 'Excel (XLSX/CSV) sang PowerPoint',
    desc: 'Tự động đưa các bảng tính dữ liệu Excel vào các trang slide thuyết trình dạng bảng đẹp.',
    fromExt: '.xlsx, .csv',
    toExt: '.pptx',
    fromIcon: FileSpreadsheet,
    toIcon: Presentation,
    badge: 'Báo cáo số liệu',
    accept: '.xlsx,.xls,.csv',
    action: convertExcelToPptx,
  },
  {
    id: 'excel-to-pdf',
    category: 'excel',
    title: 'Excel (XLSX/CSV) sang PDF Bảng biểu',
    desc: 'Xuất toàn bộ bảng tính Excel thành các trang PDF nằm ngang chuẩn in báo cáo.',
    fromExt: '.xlsx, .csv',
    toExt: '.pdf',
    fromIcon: FileSpreadsheet,
    toIcon: FileText,
    badge: 'In bảng tính',
    accept: '.xlsx,.xls,.csv',
    action: convertExcelToPdf,
  },

  // 5. Image Services
  {
    id: 'image-to-pptx',
    category: 'image',
    title: 'Hình ảnh (PNG/JPG) sang PowerPoint',
    desc: 'Đưa tệp hình ảnh vào một trang slide thuyết trình PowerPoint toàn màn hình.',
    fromExt: '.png, .jpg',
    toExt: '.pptx',
    fromIcon: ImageIcon,
    toIcon: Presentation,
    badge: 'Chuyển ảnh sang slide',
    accept: 'image/png,image/jpeg,image/webp',
    action: convertImagesToPptx,
  },
  {
    id: 'image-to-pdf',
    category: 'image',
    title: 'Hình ảnh (PNG/JPG) sang PDF',
    desc: 'Đóng gói tệp hình ảnh thành một trang tài liệu PDF chuẩn khổ A4.',
    fromExt: '.png, .jpg',
    toExt: '.pdf',
    fromIcon: ImageIcon,
    toIcon: FileText,
    badge: 'Chuyển ảnh sang PDF',
    accept: 'image/png,image/jpeg,image/webp',
    action: convertImagesToPdf,
  },
];

export const ConverterTab: React.FC = () => {
  const [selectedServiceId, setSelectedServiceId] = useState<string>('pptx-to-pdf');
  const [activeCategory, setActiveCategory] = useState<CategoryKey>('all');
  const [searchQuery, setSearchQuery] = useState('');
  const [uploadedFile, setUploadedFile] = useState<File | null>(null);
  const [isConverting, setIsConverting] = useState(false);
  const [conversionResult, setConversionResult] = useState<ConversionResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const currentService =
    SERVICES.find((s) => s.id === selectedServiceId) || SERVICES[0];

  // Filter services by category and search keyword
  const filteredServices = useMemo(() => {
    return SERVICES.filter((s) => {
      const matchCategory =
        activeCategory === 'all' || s.category === activeCategory;
      const matchSearch =
        searchQuery.trim() === '' ||
        s.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
        s.desc.toLowerCase().includes(searchQuery.toLowerCase()) ||
        s.fromExt.toLowerCase().includes(searchQuery.toLowerCase()) ||
        s.toExt.toLowerCase().includes(searchQuery.toLowerCase()) ||
        s.badge.toLowerCase().includes(searchQuery.toLowerCase());
      return matchCategory && matchSearch;
    });
  }, [activeCategory, searchQuery]);

  const handleSelectService = (service: ConverterService) => {
    setSelectedServiceId(service.id);
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

      // Smart auto-detection of best service
      if (ext === 'docx') {
        setSelectedServiceId('docx-to-pptx');
        setActiveCategory('docx');
      } else if (ext === 'xlsx' || ext === 'xls' || ext === 'csv') {
        setSelectedServiceId('excel-to-pptx');
        setActiveCategory('excel');
      } else if (ext === 'txt' || ext === 'md') {
        setSelectedServiceId('txt-to-pptx');
        setActiveCategory('text');
      } else if (ext === 'png' || ext === 'jpg' || ext === 'jpeg' || ext === 'webp') {
        setSelectedServiceId('image-to-pptx');
        setActiveCategory('image');
      } else if (ext === 'pptx' || ext === 'ppt') {
        setSelectedServiceId('pptx-to-pdf');
        setActiveCategory('pptx');
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
      const res = await currentService.action(uploadedFile);
      setConversionResult(res);
    } catch (err: any) {
      console.error('Conversion failed:', err);
      setError(err.message || 'Chuyển đổi tệp gặp sự cố ngoài ý muốn. Vui lòng kiểm tra định dạng tệp.');
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
      <div className="bg-gradient-to-r from-rose-950 via-slate-900 to-indigo-950 text-white rounded-3xl p-6 sm:p-8 shadow-md">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
          <div className="space-y-2 max-w-2xl">
            <div className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-rose-500/20 text-rose-300 text-xs font-bold border border-rose-500/30">
              <Zap className="w-3.5 h-3.5 text-amber-300" />
              <span>20 Dịch vụ chuyển đổi siêu tốc • 100% trong trình duyệt</span>
            </div>
            <h2 className="text-2xl sm:text-3xl font-extrabold tracking-tight">
              Trung Tâm Chuyển Đổi Định Dạng Đa Năng
            </h2>
            <p className="text-xs sm:text-sm text-slate-300 leading-relaxed">
              Chuyển đổi qua lại giữa tất cả các loại tệp văn phòng phổ biến nhất: 
              <strong> PowerPoint (.pptx)</strong>, <strong>Word (.docx)</strong>, <strong>PDF</strong>, 
              <strong> Excel (.xlsx/.csv)</strong>, <strong>Văn bản (.txt/.md)</strong>, <strong>Hình ảnh</strong> và <strong>Trình chiếu Web (HTML)</strong> mà không cần tải lên máy chủ.
            </p>
          </div>

          <div className="flex items-center gap-2 p-3 bg-white/10 rounded-2xl border border-white/15 backdrop-blur-xs text-xs text-slate-200 shrink-0">
            <ShieldCheck className="w-5 h-5 text-emerald-400 shrink-0" />
            <div>
              <span className="font-bold block text-white">Bảo mật vĩnh viễn</span>
              <span className="text-[11px] text-slate-300">Không gửi file lên server</span>
            </div>
          </div>
        </div>
      </div>

      {/* Category Pills & Search Bar */}
      <div className="space-y-4">
        <div className="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3">
          {/* Category Tabs */}
          <div className="flex items-center gap-1.5 overflow-x-auto pb-1 scrollbar-thin">
            {[
              { id: 'all', label: 'Tất cả (20 dịch vụ)' },
              { id: 'pptx', label: 'PowerPoint (.pptx)' },
              { id: 'docx', label: 'Word (.docx)' },
              { id: 'excel', label: 'Excel (.xlsx/.csv)' },
              { id: 'text', label: 'Văn bản (.txt/.md)' },
              { id: 'image', label: 'Hình ảnh' },
            ].map((cat) => (
              <button
                key={cat.id}
                type="button"
                onClick={() => setActiveCategory(cat.id as CategoryKey)}
                className={`px-3.5 py-1.5 rounded-xl text-xs font-bold transition-all whitespace-nowrap cursor-pointer ${
                  activeCategory === cat.id
                    ? 'bg-rose-600 text-white shadow-xs'
                    : 'bg-slate-100 hover:bg-slate-200 text-slate-700'
                }`}
              >
                {cat.label}
              </button>
            ))}
          </div>

          {/* Quick Search */}
          <div className="relative w-full sm:w-64 shrink-0">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              placeholder="Tìm dịch vụ chuyển đổi..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full pl-9 pr-3 py-1.5 text-xs bg-white border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-rose-500/20 focus:border-rose-500"
            />
          </div>
        </div>

        {/* Services Grid (Showing 20 Services) */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
          {filteredServices.map((service) => {
            const isSelected = selectedServiceId === service.id;
            const FromIcon = service.fromIcon;
            const ToIcon = service.toIcon;

            return (
              <button
                key={service.id}
                type="button"
                onClick={() => handleSelectService(service)}
                className={`p-4 rounded-2xl border text-left transition-all cursor-pointer relative flex flex-col justify-between ${
                  isSelected
                    ? 'bg-rose-50/70 border-rose-500 ring-2 ring-rose-500/20 shadow-sm'
                    : 'bg-white border-slate-200/90 hover:border-slate-300 hover:shadow-xs'
                }`}
              >
                <div>
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-[10px] font-extrabold uppercase px-2 py-0.5 rounded-md bg-slate-100 text-slate-700 font-mono">
                      {service.badge}
                    </span>
                    <div className="flex items-center gap-1.5 text-xs font-bold text-slate-700">
                      <span className="text-rose-600 font-mono">{service.fromExt}</span>
                      <ArrowRight className="w-3 h-3 text-slate-400" />
                      <span className="text-indigo-600 font-mono">{service.toExt}</span>
                    </div>
                  </div>

                  <h4 className="font-bold text-sm text-slate-900 mb-1 flex items-center gap-2">
                    <FromIcon className="w-4 h-4 text-rose-600 shrink-0" />
                    <span>{service.title}</span>
                  </h4>
                  <p className="text-xs text-slate-500 leading-relaxed">{service.desc}</p>
                </div>

                {isSelected && (
                  <div className="mt-3 pt-2 border-t border-rose-200 flex items-center justify-between text-[11px] font-bold text-rose-700">
                    <span>Đang chọn dịch vụ này</span>
                    <Check className="w-3.5 h-3.5 text-rose-600" />
                  </div>
                )}
              </button>
            );
          })}
        </div>
      </div>

      {/* Upload Zone & Action Section */}
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
              className="text-xs font-semibold underline hover:text-red-900 cursor-pointer"
            >
              Đóng
            </button>
          </div>
        )}

        {conversionResult ? (
          /* Conversion Success Card */
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
              Dữ liệu được xử lý 100% trong trình duyệt với dung lượng tối ưu và chuẩn xác.
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
          /* Dropzone Upload Section */
          <div className="space-y-6">
            <input
              ref={fileInputRef}
              type="file"
              accept={currentService.accept}
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
                    Kéo thả tệp {currentService.fromExt} vào đây
                  </h4>
                  <p className="text-xs sm:text-sm text-slate-500 mb-4">
                    hoặc bấm để chọn tệp từ thiết bị của bạn
                  </p>
                  <button
                    type="button"
                    className="px-5 py-2.5 bg-slate-900 hover:bg-slate-800 text-white font-bold text-xs rounded-xl shadow-xs transition-all pointer-events-none"
                  >
                    Chọn tệp {currentService.fromExt}
                  </button>
                </div>
              )}
            </div>

            {/* Convert Execute Button */}
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
                    <span>TIẾN HÀNH CHUYỂN ĐỔI SANG {currentService.toExt.toUpperCase()}</span>
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

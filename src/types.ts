export interface SlideInfo {
  slideNumber: number;
  title: string;
  previewText: string[];
  hasImages: boolean;
}

export interface PPTFileItem {
  id: string;
  file: File;
  name: string;
  size: number;
  slideCount: number;
  slides: SlideInfo[];
  selectedSlideNumbers: number[]; // e.g. [1, 2, 3] or all
  isAnalyzing?: boolean;
}

export interface FileAuditItem {
  fileIndex: number;
  fileName: string;
  totalOriginalSlides: number;
  includedSlidesCount: number;
  slideRangeStart: number;
  slideRangeEnd: number;
  verified: boolean;
}

export interface MergeResult {
  blob: Blob;
  url: string;
  filename: string;
  totalSlides: number;
  totalFiles: number;
  fileSize: number;
  createdAt: Date;
  auditItems?: FileAuditItem[];
}

import pptxgen from 'pptxgenjs';

export async function generateClientSamples(): Promise<File[]> {
  const sampleFiles: File[] = [];

  // Sample 1: Báo cáo Tổng quan (Overview)
  const p1 = new pptxgen();
  p1.layout = 'LAYOUT_WIDE';
  
  const s1_1 = p1.addSlide();
  s1_1.background = { color: '1E293B' }; // Dark Slate
  s1_1.addText('BÁO CÁO TỔNG QUAN DOANH NGHIỆP', {
    x: 1,
    y: 2,
    w: 11,
    fontSize: 32,
    bold: true,
    color: 'FFFFFF',
    fontFace: 'Arial',
  });
  s1_1.addText('Phần 1: Tình hình hoạt động & Chỉ số then chốt', {
    x: 1,
    y: 3.2,
    w: 11,
    fontSize: 20,
    color: '94A3B8',
    fontFace: 'Arial',
  });
  s1_1.addText('toolgopfilepptbylekimlam • 2026', {
    x: 1,
    y: 6.2,
    w: 11,
    fontSize: 14,
    color: '64748B',
    fontFace: 'Arial',
  });

  const s1_2 = p1.addSlide();
  s1_2.background = { color: 'F8FAFC' };
  s1_2.addText('KẾT QUẢ TĂNG TRƯỞNG NĂM 2025', {
    x: 0.8,
    y: 0.8,
    w: 11,
    fontSize: 24,
    bold: true,
    color: '0F172A',
    fontFace: 'Arial',
  });
  s1_2.addText('• Tăng trưởng doanh thu toàn công ty đạt +28% so với cùng kỳ', {
    x: 1,
    y: 2.0,
    w: 10,
    fontSize: 18,
    color: '334155',
    fontFace: 'Arial',
  });
  s1_2.addText('• Mở rộng thị phần sang 3 khu vực kinh tế trọng điểm mới', {
    x: 1,
    y: 2.8,
    w: 10,
    fontSize: 18,
    color: '334155',
    fontFace: 'Arial',
  });
  s1_2.addText('• Chỉ số hài lòng khách hàng (CSAT) đạt 94.6%', {
    x: 1,
    y: 3.6,
    w: 10,
    fontSize: 18,
    color: '334155',
    fontFace: 'Arial',
  });

  const b1 = (await p1.write({ outputType: 'arraybuffer' })) as ArrayBuffer;
  sampleFiles.push(
    new File([b1], '01_Bao_Cao_Tong_Quan.pptx', {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    })
  );

  // Sample 2: Kế hoạch Chiến lược (Strategy)
  const p2 = new pptxgen();
  p2.layout = 'LAYOUT_WIDE';
  
  const s2_1 = p2.addSlide();
  s2_1.background = { color: '0284C7' }; // Ocean Blue
  s2_1.addText('KẾ HOẠCH CHIẾN LƯỢC QUÝ MỚI', {
    x: 1,
    y: 2,
    w: 11,
    fontSize: 32,
    bold: true,
    color: 'FFFFFF',
    fontFace: 'Arial',
  });
  s2_1.addText('Phần 2: Mục tiêu trọng tâm & Kế hoạch ngân sách', {
    x: 1,
    y: 3.2,
    w: 11,
    fontSize: 20,
    color: 'E0F2FE',
    fontFace: 'Arial',
  });

  const s2_2 = p2.addSlide();
  s2_2.background = { color: 'F0F9FF' };
  s2_2.addText('3 TRỤ CỘT CHIẾN LƯỢC ƯU TIÊN', {
    x: 0.8,
    y: 0.8,
    w: 11,
    fontSize: 24,
    bold: true,
    color: '0369A1',
    fontFace: 'Arial',
  });
  s2_2.addText('1. Chuyển đổi số toàn diện quy trình vận hành và báo cáo tự động', {
    x: 1,
    y: 2.0,
    w: 10,
    fontSize: 18,
    color: '0F172A',
    fontFace: 'Arial',
  });
  s2_2.addText('2. Nâng cấp năng lực đội ngũ và hệ thống hạ tầng đám mây', {
    x: 1,
    y: 2.8,
    w: 10,
    fontSize: 18,
    color: '0F172A',
    fontFace: 'Arial',
  });
  s2_2.addText('3. Tối ưu hoá chi phí sản xuất, gia tăng biên lợi nhuận ròng', {
    x: 1,
    y: 3.6,
    w: 10,
    fontSize: 18,
    color: '0F172A',
    fontFace: 'Arial',
  });

  const b2 = (await p2.write({ outputType: 'arraybuffer' })) as ArrayBuffer;
  sampleFiles.push(
    new File([b2], '02_Ke_Hoach_Chien_Luoc.pptx', {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    })
  );

  // Sample 3: Kết luận & Kiến nghị (Conclusion)
  const p3 = new pptxgen();
  p3.layout = 'LAYOUT_WIDE';
  
  const s3_1 = p3.addSlide();
  s3_1.background = { color: '059669' }; // Emerald
  s3_1.addText('KẾT LUẬN & KIẾN NGHỊ THỰC HIỆN', {
    x: 1,
    y: 2,
    w: 11,
    fontSize: 32,
    bold: true,
    color: 'FFFFFF',
    fontFace: 'Arial',
  });
  s3_1.addText('Phần 3: Đề xuất phê duyệt & Lộ trình phân công', {
    x: 1,
    y: 3.2,
    w: 11,
    fontSize: 20,
    color: 'D1FAE5',
    fontFace: 'Arial',
  });

  const s3_2 = p3.addSlide();
  s3_2.background = { color: 'ECFDF5' };
  s3_2.addText('LỘ TRÌNH TRIỂN KHAI VÀ BÀN GIAO', {
    x: 0.8,
    y: 0.8,
    w: 11,
    fontSize: 24,
    bold: true,
    color: '047857',
    fontFace: 'Arial',
  });
  s3_2.addText('• Tuần 1 - 2: Hoàn thiện phê duyệt kinh phí và phân bổ nhân sự', {
    x: 1,
    y: 2.0,
    w: 10,
    fontSize: 18,
    color: '064E3B',
    fontFace: 'Arial',
  });
  s3_2.addText('• Tuần 3 - 6: Triển khai giai đoạn thử nghiệm (Pilot phase)', {
    x: 1,
    y: 2.8,
    w: 10,
    fontSize: 18,
    color: '064E3B',
    fontFace: 'Arial',
  });
  s3_2.addText('• Tuần 7+: Đánh giá kết quả và mở rộng trên toàn hệ thống', {
    x: 1,
    y: 3.6,
    w: 10,
    fontSize: 18,
    color: '064E3B',
    fontFace: 'Arial',
  });

  const b3 = (await p3.write({ outputType: 'arraybuffer' })) as ArrayBuffer;
  sampleFiles.push(
    new File([b3], '03_Ket_Luan_Kien_Nghi.pptx', {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    })
  );

  return sampleFiles;
}

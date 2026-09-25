import * as XLSX from 'xlsx';
import { Race } from '../data/races';
import { CertificatePlacements, CertificateFieldPlacement } from '../types';
import { DEFAULT_NGHE_AN_PLACEMENTS, getSavedPlacements } from '../data/certificatePlacements';
import { exportRaceStaticApi } from './exportRaceStaticApi';

const SAMPLE_VALUES: Record<string, string> = {
  name: 'PHÙNG HỮU THANH',
  distance: '42K',
  bib: '90110',
  chipTime: '03:14:48',
  finishTime: '03:15:20',
  overallRank: '26',
  genderRank: '22',
  ageGroupRank: '8',
};

const FIELD_NOTES: Record<string, string> = {
  name: 'Tên VĐV in hoa, căn giữa phôi phía dưới logo & huy hiệu giải',
  distance: 'Cự ly hoàn thành (VD: 42K, 21K, 10K, 5K), căn giữa ô bên trái',
  bib: 'Số BIB thi đấu (VD: 90110), căn giữa ô bên phải ngang hàng cự ly',
  chipTime: 'Thời gian thực tế qua vạch (Chip Time), cỡ chữ to nhất ở giữa phôi',
  finishTime: 'Thời gian xuất phát đến đích (Gun Time), căn trái trong ô thời gian',
  overallRank: 'Thứ hạng toàn giải (Overall Rank), căn trái ô xếp hạng chung',
  genderRank: 'Thứ hạng theo giới tính (Gender Rank), căn trái ô giới tính',
  ageGroupRank: 'Thứ hạng lứa tuổi (Age Group Rank), căn trái ô lứa tuổi',
};

/**
 * Xuất toàn bộ cấu hình giải đấu VÀ cấu hình chỉnh phôi ra file Excel (.xlsx)
 */
export function exportRaceToExcel(
  race: Race,
  placements?: CertificatePlacements,
  alsoExportJson: boolean = true
) {
  const wb = XLSX.utils.book_new();

  const baseUrl = typeof window !== 'undefined' ? window.location.origin : '';
  const fullPageUrl = `${baseUrl}/${race.slug}`;

  // Đảm bảo luôn luôn có đầy đủ cấu hình chỉnh phôi, không bao giờ bị thiếu
  const activePlacements: CertificatePlacements =
    placements && Object.keys(placements).length > 0
      ? placements
      : race.placements && Object.keys(race.placements).length > 0
      ? race.placements
      : getSavedPlacements() || DEFAULT_NGHE_AN_PLACEMENTS;

  const nameP = activePlacements.name || DEFAULT_NGHE_AN_PLACEMENTS.name;
  const distP = activePlacements.distance || DEFAULT_NGHE_AN_PLACEMENTS.distance;
  const bibP = activePlacements.bib || DEFAULT_NGHE_AN_PLACEMENTS.bib;
  const chipP = activePlacements.chipTime || DEFAULT_NGHE_AN_PLACEMENTS.chipTime;
  const finishP = activePlacements.finishTime || DEFAULT_NGHE_AN_PLACEMENTS.finishTime;
  const overallP = activePlacements.overallRank || DEFAULT_NGHE_AN_PLACEMENTS.overallRank;
  const genderP = activePlacements.genderRank || DEFAULT_NGHE_AN_PLACEMENTS.genderRank;
  const ageP = activePlacements.ageGroupRank || DEFAULT_NGHE_AN_PLACEMENTS.ageGroupRank;

  // 1. Sheet: Cấu Hình Giải Đấu & Tổng Quan Chỉnh Phôi (ThongTinGiai)
  const raceInfoRows = [
    // 4 thành phần tạo giải bắt buộc
    { 'MỤC CẤU HÌNH': '1. Tên giải đấu', 'GIÁ TRỊ CÀI ĐẶT': race.name, 'GHI CHÚ': 'Tên chính thức hiển thị trên toàn hệ thống' },
    { 'MỤC CẤU HÌNH': '2. URL ĐỂ VÀO TRANG (Slug)', 'GIÁ TRỊ CÀI ĐẶT': `/${race.slug}`, 'GHI CHÚ': 'Đường dẫn định tuyến vào trang giải' },
    { 'MỤC CẤU HÌNH': 'Đường link truy cập đầy đủ', 'GIÁ TRỊ CÀI ĐẶT': fullPageUrl, 'GHI CHÚ': 'Gửi link này cho VĐV hoặc chia sẻ MXH' },
    { 'MỤC CẤU HÌNH': '3. Background Certificate', 'GIÁ TRỊ CÀI ĐẶT': race.defaultBgUrl, 'GHI CHÚ': 'Phôi chứng nhận chuẩn HD (tỷ lệ giống NA26: 1469x3508px)' },
    { 'MỤC CẤU HÌNH': '4. Add Script (Data riêng)', 'GIÁ TRỊ CÀI ĐẶT': race.appsScriptUrl || 'Chưa liên kết', 'GHI CHÚ': 'Google Apps Script tải kết quả VĐV riêng theo giải' },

    // CẤU HÌNH CHỈNH PHÔI (Bổ sung chi tiết theo yêu cầu người dùng)
    { 'MỤC CẤU HÌNH': '5. CẤU HÌNH CHỈNH PHÔI (Trạng thái)', 'GIÁ TRỊ CÀI ĐẶT': 'ĐÃ THIẾT LẬP HOÀN CHỈNH (8 trường toạ độ & kích thước)', 'GHI CHÚ': 'Xem chi tiết toạ độ từng trường ở Sheet "CauHinhChinhPhoi"' },
    { 'MỤC CẤU HÌNH': '   • Kích thước chuẩn phôi Certificate', 'GIÁ TRỊ CÀI ĐẶT': '1469 × 3508 px (Chuẩn in ấn A4 dọc HD / 300 DPI)', 'GHI CHÚ': 'Tỷ lệ khung hình 1 : 2.388' },
    { 'MỤC CẤU HÌNH': '   • Phông chữ áp dụng trên phôi', 'GIÁ TRỊ CÀI ĐẶT': 'Montserrat (Hỗ trợ tiếng Việt đầy đủ dấu) / Plus Jakarta Sans', 'GHI CHÚ': 'Font đồ hoạ vector sắc nét khi xuất ảnh' },
    { 'MỤC CẤU HÌNH': '   • Chế độ hiển thị Tên VĐV', 'GIÁ TRỊ CÀI ĐẶT': 'IN HOA (UPPERCASE)', 'GHI CHÚ': 'Tự động chuẩn hoá viết hoa tên VĐV' },
    { 'MỤC CẤU HÌNH': '   • Vị trí 1 - Tên VĐV (name)', 'GIÁ TRỊ CÀI ĐẶT': `X: ${nameP.x}% | Y: ${nameP.y}% | Size: ${nameP.fontSize}px | Màu: ${nameP.color} | Căn: ${nameP.align}`, 'GHI CHÚ': `Pixel: (${Math.round((nameP.x / 100) * 1469)}, ${Math.round((nameP.y / 100) * 3508)})` },
    { 'MỤC CẤU HÌNH': '   • Vị trí 2 - Cự ly (distance)', 'GIÁ TRỊ CÀI ĐẶT': `X: ${distP.x}% | Y: ${distP.y}% | Size: ${distP.fontSize}px | Màu: ${distP.color} | Căn: ${distP.align}`, 'GHI CHÚ': `Pixel: (${Math.round((distP.x / 100) * 1469)}, ${Math.round((distP.y / 100) * 3508)})` },
    { 'MỤC CẤU HÌNH': '   • Vị trí 3 - Số BIB (bib)', 'GIÁ TRỊ CÀI ĐẶT': `X: ${bibP.x}% | Y: ${bibP.y}% | Size: ${bibP.fontSize}px | Màu: ${bibP.color} | Căn: ${bibP.align}`, 'GHI CHÚ': `Pixel: (${Math.round((bibP.x / 100) * 1469)}, ${Math.round((bibP.y / 100) * 3508)})` },
    { 'MỤC CẤU HÌNH': '   • Vị trí 4 - Chip Time (chipTime)', 'GIÁ TRỊ CÀI ĐẶT': `X: ${chipP.x}% | Y: ${chipP.y}% | Size: ${chipP.fontSize}px | Màu: ${chipP.color} | Căn: ${chipP.align}`, 'GHI CHÚ': `Pixel: (${Math.round((chipP.x / 100) * 1469)}, ${Math.round((chipP.y / 100) * 3508)})` },
    { 'MỤC CẤU HÌNH': '   • Vị trí 5 - Gun Time (finishTime)', 'GIÁ TRỊ CÀI ĐẶT': `X: ${finishP.x}% | Y: ${finishP.y}% | Size: ${finishP.fontSize}px | Màu: ${finishP.color} | Căn: ${finishP.align}`, 'GHI CHÚ': `Pixel: (${Math.round((finishP.x / 100) * 1469)}, ${Math.round((finishP.y / 100) * 3508)})` },
    { 'MỤC CẤU HÌNH': '   • Vị trí 6 - Hạng chung cuộc (overallRank)', 'GIÁ TRỊ CÀI ĐẶT': `X: ${overallP.x}% | Y: ${overallP.y}% | Size: ${overallP.fontSize}px | Màu: ${overallP.color} | Căn: ${overallP.align}`, 'GHI CHÚ': `Pixel: (${Math.round((overallP.x / 100) * 1469)}, ${Math.round((overallP.y / 100) * 3508)})` },
    { 'MỤC CẤU HÌNH': '   • Vị trí 7 - Hạng giới tính (genderRank)', 'GIÁ TRỊ CÀI ĐẶT': `X: ${genderP.x}% | Y: ${genderP.y}% | Size: ${genderP.fontSize}px | Màu: ${genderP.color} | Căn: ${genderP.align}`, 'GHI CHÚ': `Pixel: (${Math.round((genderP.x / 100) * 1469)}, ${Math.round((genderP.y / 100) * 3508)})` },
    { 'MỤC CẤU HÌNH': '   • Vị trí 8 - Hạng lứa tuổi (ageGroupRank)', 'GIÁ TRỊ CÀI ĐẶT': `X: ${ageP.x}% | Y: ${ageP.y}% | Size: ${ageP.fontSize}px | Màu: ${ageP.color} | Căn: ${ageP.align}`, 'GHI CHÚ': `Pixel: (${Math.round((ageP.x / 100) * 1469)}, ${Math.round((ageP.y / 100) * 3508)})` },

    // Thông tin bổ sung
    { 'MỤC CẤU HÌNH': 'Mã nhận diện giải (Code)', 'GIÁ TRỊ CÀI ĐẶT': race.code || race.slug.toUpperCase(), 'GHI CHÚ': 'Mã viết tắt (VD: NA26, HL26)' },
    { 'MỤC CẤU HÌNH': 'Ngày thi đấu', 'GIÁ TRỊ CÀI ĐẶT': race.date || 'Chưa cập nhật', 'GHI CHÚ': 'Ngày diễn ra giải chạy' },
    { 'MỤC CẤU HÌNH': 'Địa điểm tổ chức', 'GIÁ TRỊ CÀI ĐẶT': race.locationFull || race.city || 'Việt Nam', 'GHI CHÚ': 'Tỉnh / Thành phố' },
    { 'MỤC CẤU HÌNH': 'Mô tả giải', 'GIÁ TRỊ CÀI ĐẶT': race.description || '', 'GHI CHÚ': 'Thông tin giới thiệu ngắn' },
    { 'MỤC CẤU HÌNH': 'Thời gian xuất file cấu hình', 'GIÁ TRỊ CÀI ĐẶT': new Date().toLocaleString('vi-VN'), 'GHI CHÚ': 'Thời điểm xuất file Excel từ Admin' },
  ];

  const wsRaceInfo = XLSX.utils.json_to_sheet(raceInfoRows);
  wsRaceInfo['!cols'] = [
    { wch: 38 },
    { wch: 80 },
    { wch: 55 },
  ];
  XLSX.utils.book_append_sheet(wb, wsRaceInfo, 'ThongTinGiai');

  // 2. Sheet: CẤU HÌNH CHỈNH PHÔI CHI TIẾT (CauHinhChinhPhoi)
  // Đảm bảo luôn luôn có sheet này với đầy đủ các trường
  const orderedFieldKeys = [
    'name',
    'distance',
    'bib',
    'chipTime',
    'finishTime',
    'overallRank',
    'genderRank',
    'ageGroupRank',
  ];

  const placementsRows = orderedFieldKeys.map((key, idx) => {
    const item: CertificateFieldPlacement = (activePlacements as any)[key] || (DEFAULT_NGHE_AN_PLACEMENTS as any)[key];
    const xPct = Number(item.x || 50);
    const yPct = Number(item.y || 50);
    const pixelX = Math.round((xPct / 100) * 1469);
    const pixelY = Math.round((yPct / 100) * 3508);

    return {
      'STT': idx + 1,
      'Mã trường (Field ID)': item.id || key,
      'Tên trường hiển thị': item.label || key,
      'Tọa độ X (%)': Number(xPct).toFixed(2),
      'Tọa độ Y (%)': Number(yPct).toFixed(2),
      'Tọa độ Pixel X (1469px)': pixelX,
      'Tọa độ Pixel Y (3508px)': pixelY,
      'Cỡ chữ (Font Size px)': item.fontSize,
      'Mã màu (Color HEX)': item.color || '#000000',
      'Độ đậm (Font Weight)': item.fontWeight || 700,
      'Căn lề (Align)': item.align || 'center',
      'Font chữ': 'Montserrat, sans-serif',
      'Chữ in hoa': key === 'name' ? 'Có (UPPERCASE)' : 'Không',
      'Dữ liệu mẫu trên phôi': SAMPLE_VALUES[key] || '',
      'Ghi chú & Hướng dẫn chỉnh': FIELD_NOTES[key] || 'Căn chỉnh theo ô trên phôi chứng nhận',
    };
  });

  const wsPlacements = XLSX.utils.json_to_sheet(placementsRows);
  wsPlacements['!cols'] = [
    { wch: 6 },
    { wch: 22 },
    { wch: 32 },
    { wch: 15 },
    { wch: 15 },
    { wch: 24 },
    { wch: 24 },
    { wch: 22 },
    { wch: 18 },
    { wch: 22 },
    { wch: 16 },
    { wch: 24 },
    { wch: 16 },
    { wch: 25 },
    { wch: 60 },
  ];
  XLSX.utils.book_append_sheet(wb, wsPlacements, 'CauHinhChinhPhoi');

  // 3. Sheet: THÔNG SỐ KỸ THUẬT PHÔI (ThongSoKyThuatPhoi)
  const techSpecsRows = [
    { 'THÔNG SỐ KỸ THUẬT': 'Chiều rộng Canvas (Width)', 'GIÁ TRỊ': '1469 px', 'Ý NGHĨA': 'Chiều ngang chuẩn ảnh phôi chứng nhận' },
    { 'THÔNG SỐ KỸ THUẬT': 'Chiều cao Canvas (Height)', 'GIÁ TRỊ': '3508 px', 'Ý NGHĨA': 'Chiều dọc chuẩn ảnh phôi chứng nhận (chuẩn A4)' },
    { 'THÔNG SỐ KỸ THUẬT': 'Tỷ lệ khung hình (Aspect Ratio)', 'GIÁ TRỊ': '1 : 2.388 (0.4188)', 'Ý NGHĨA': 'Tỷ lệ chuẩn in ấn dọc chuẩn quốc tế' },
    { 'THÔNG SỐ KỸ THUẬT': 'Độ phân giải (DPI)', 'GIÁ TRỊ': '300 DPI High Resolution', 'Ý NGHĨA': 'Đảm bảo in ấn và phóng to không bị vỡ hạt' },
    { 'THÔNG SỐ KỸ THUẬT': 'Tọa độ tâm phôi (Center Guide)', 'GIÁ TRỊ': 'X: 50.00%, Y: 50.00%', 'Ý NGHĨA': 'Tâm căn giữa khi thiết kế phôi' },
    { 'THÔNG SỐ KỸ THUẬT': 'Khoảng cách an toàn lề (Safe Margins)', 'GIÁ TRỊ': '5% (73px mỗi bên)', 'Ý NGHĨA': 'Tránh chữ bị sát mép giấy khi in' },
    { 'THÔNG SỐ KỸ THUẬT': 'Định dạng xuất ảnh chứng nhận', 'GIÁ TRỊ': 'PNG chất lượng cao (24-bit RGB)', 'Ý NGHĨA': 'Đầy đủ màu sắc và độ sắc nét cao nhất' },
    { 'THÔNG SỐ KỸ THUẬT': 'Quy tắc đặt tên file ảnh tải về', 'GIÁ TRỊ': 'Certificate_[BIB]_[TenVĐV].png', 'Ý NGHĨA': 'Giúp VĐV dễ nhận biết file chứng nhận của mình' },
  ];

  const wsTechSpecs = XLSX.utils.json_to_sheet(techSpecsRows);
  wsTechSpecs['!cols'] = [
    { wch: 38 },
    { wch: 35 },
    { wch: 55 },
  ];
  XLSX.utils.book_append_sheet(wb, wsTechSpecs, 'ThongSoKyThuatPhoi');

  // Tạo và tải file Excel về máy
  const cleanSlug = (race.slug || 'giai').replace(/[^a-zA-Z0-9_-]/g, '_');
  const nowStr = new Date().toISOString().slice(0, 10);
  const fileName = `Cau_Hinh_Chinh_Phoi_${cleanSlug}_${nowStr}.xlsx`;

  XLSX.writeFile(wb, fileName);

  // Tự động xuất file API tĩnh (.json) để ném vào thư mục public/races/ sinh giải mới
  if (alsoExportJson) {
    setTimeout(() => {
      try {
        exportRaceStaticApi(race, activePlacements);
      } catch (err) {
        console.warn('Lỗi xuất file API tĩnh kèm theo:', err);
      }
    }, 250);
  }
}

/**
 * Đọc file Excel cấu hình để khôi phục lại cấu hình giải đấu và toạ độ phôi
 */
export async function importRaceFromExcel(file: File): Promise<{
  raceInfo: Partial<Race>;
  placements: CertificatePlacements;
}> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });

        const raceInfo: Partial<Race> = {};
        const placements: CertificatePlacements = { ...DEFAULT_NGHE_AN_PLACEMENTS };

        // 1. Đọc sheet ThongTinGiai
        const raceInfoSheet = wb.Sheets['ThongTinGiai'] || wb.Sheets[wb.SheetNames[0]];
        if (raceInfoSheet) {
          const rows: any[] = XLSX.utils.sheet_to_json(raceInfoSheet);
          for (const row of rows) {
            const itemKey = String(row['MỤC CẤU HÌNH'] || row['MUC CAU HINH'] || '').toLowerCase();
            const val = String(row['GIÁ TRỊ CÀI ĐẶT'] || row['GIA TRI CAI DAT'] || '').trim();

            if (itemKey.includes('tên giải') || itemKey.includes('ten giai')) {
              raceInfo.name = val;
              raceInfo.shortName = val;
            } else if (itemKey.includes('slug') || itemKey.includes('url')) {
              raceInfo.slug = val.replace(/^\//, '');
            } else if (itemKey.includes('background') || itemKey.includes('phôi') || itemKey.includes('phoi')) {
              raceInfo.defaultBgUrl = val;
            } else if (itemKey.includes('script') || itemKey.includes('apps script')) {
              raceInfo.appsScriptUrl = val;
            } else if (itemKey.includes('mã') || itemKey.includes('code')) {
              raceInfo.code = val;
            } else if (itemKey.includes('ngày') || itemKey.includes('date')) {
              raceInfo.date = val;
            } else if (itemKey.includes('địa điểm') || itemKey.includes('location')) {
              raceInfo.locationFull = val;
            }
          }
        }

        // 2. Đọc sheet CauHinhChinhPhoi hoặc ToaDoChungNhan
        const placementsSheet =
          wb.Sheets['CauHinhChinhPhoi'] ||
          wb.Sheets['ToaDoChungNhan'] ||
          (wb.SheetNames.length > 1 ? wb.Sheets[wb.SheetNames[1]] : null);

        if (placementsSheet) {
          const rows: any[] = XLSX.utils.sheet_to_json(placementsSheet);
          for (const row of rows) {
            const fieldId =
              row['Mã trường (Field ID)'] ||
              row['Mã trường'] ||
              row['Ma truong'] ||
              row['id'];
            if (!fieldId) continue;

            const cleanId = String(fieldId).trim();
            const existing = placements[cleanId] || (DEFAULT_NGHE_AN_PLACEMENTS as any)[cleanId];

            const xVal = parseFloat(row['Tọa độ X (%)'] || row['X'] || existing?.x || 50);
            const yVal = parseFloat(row['Tọa độ Y (%)'] || row['Y'] || existing?.y || 50);
            const fontSizeVal = parseInt(row['Cỡ chữ (Font Size px)'] || row['Cỡ chữ'] || row['fontSize'] || existing?.fontSize || 50, 10);
            const colorVal = String(row['Mã màu (Color HEX)'] || row['Màu sắc'] || row['color'] || existing?.color || '#000000').trim();
            const alignVal = String(row['Căn lề (Align)'] || row['Căn lề'] || row['align'] || existing?.align || 'center').trim() as any;
            const weightVal = parseInt(row['Độ đậm (Font Weight)'] || row['Độ đậm'] || row['fontWeight'] || existing?.fontWeight || 700, 10) as any;

            placements[cleanId] = {
              id: cleanId,
              label: String(row['Tên trường hiển thị'] || existing?.label || cleanId),
              x: isNaN(xVal) ? existing?.x || 50 : xVal,
              y: isNaN(yVal) ? existing?.y || 50 : yVal,
              fontSize: isNaN(fontSizeVal) ? existing?.fontSize || 50 : fontSizeVal,
              color: colorVal || existing?.color || '#000000',
              align: ['left', 'center', 'right'].includes(alignVal) ? alignVal : existing?.align || 'center',
              fontWeight: [600, 700, 800, 900].includes(weightVal) ? weightVal : existing?.fontWeight || 700,
            };
          }
        }

        resolve({ raceInfo, placements });
      } catch (err) {
        reject(new Error('Không thể đọc file Excel. Vui lòng kiểm tra định dạng file .xlsx'));
      }
    };

    reader.onerror = () => reject(new Error('Lỗi khi đọc file'));
    reader.readAsArrayBuffer(file);
  });
}

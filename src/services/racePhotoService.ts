import { Race } from '../data/races';
import { DEMO_RACE_PHOTOS, DEMO_PHOTOS } from '../data/mockRunners';
import { getDirectGoogleDriveImageUrl } from './sheetService';

/**
 * Mã nguồn mẫu Google Apps Script để người dùng gắn vào Google Sheet ảnh thi đấu (Cột BIB & Cột IMG)
 */
export const GOOGLE_APPS_SCRIPT_PHOTOS_CODE = `/**
 * ==============================================================================
 * GOOGLE APPS SCRIPT: LẤY ẢNH THI ĐẤU THEO SỐ BIB (VNEXPRESS MARATHON)
 * ==============================================================================
 * Cấu trúc Google Sheet (2 cột):
 * - Cột A: BIB (hoặc 'Số BIB', 'bib') -> Số BIB của VĐV (Ví dụ: 90110, 61137, 52535...)
 * - Cột B: IMG (hoặc 'Link ảnh', 'Ảnh', 'Photo', 'image') -> Link ảnh (URL Google Drive hoặc link web trực tiếp)
 *
 * * Một số BIB có thể có nhiều dòng tương ứng với nhiều góc chụp khác nhau.
 *
 * CÁCH TRIỂN KHAI NHANH:
 * 1. Mở Google Sheet chứa 2 cột (BIB và IMG).
 * 2. Menu: Tiện ích mở rộng (Extensions) > Apps Script.
 * 3. Dán toàn bộ mã nguồn này vào.
 * 4. Bấm "Triển khai" (Deploy) > "Quản lý bản triển khai mới" (New deployment).
 * 5. Chọn loại: "Ứng dụng web" (Web App).
 *    - Mô tả: "API Lấy ảnh thi đấu Marathon"
 *    - Thực thi dưới dạng (Execute as): "Tôi" (Me)
 *    - Ai có quyền truy cập (Who has access): "Bất kỳ ai" (Anyone) -> [RẤT QUAN TRỌNG]
 * 6. Bấm "Triển khai" (Deploy) và cấp quyền truy cập nếu được hỏi.
 * 7. Copy URL Web App (kết thúc bằng /exec) để dán vào hệ thống!
 */

function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      return responseJSON({ success: false, message: "Sheet rỗng hoặc không có dữ liệu" });
    }
    
    // Tìm vị trí cột BIB và IMG từ dòng tiêu đề (Dòng 1)
    var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
    
    var bibColIdx = -1;
    var imgColIdx = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var h = headers[i];
      if (h === 'bib' || h.indexOf('số bib') !== -1 || h === 'so bib' || h.indexOf('bib') !== -1) {
        if (bibColIdx === -1) bibColIdx = i;
      } else if (h === 'img' || h.indexOf('ảnh') !== -1 || h.indexOf('anh') !== -1 || h.indexOf('photo') !== -1 || h.indexOf('link') !== -1 || h.indexOf('image') !== -1) {
        if (imgColIdx === -1) imgColIdx = i;
      }
    }
    
    // Mặc định: Cột 0 là BIB, Cột 1 là IMG
    if (bibColIdx === -1) bibColIdx = 0;
    if (imgColIdx === -1) imgColIdx = 1;
    
    var requestedBib = (e && e.parameter && e.parameter.bib) ? String(e.parameter.bib).trim().toLowerCase() : '';
    
    var photosByBib = {};
    var allPhotos = [];
    
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var rawBib = String(row[bibColIdx] || '').trim();
      var rawImg = String(row[imgColIdx] || '').trim();
      
      if (!rawBib || !rawImg) continue;
      
      // Tự động chuyển đổi link xem Google Drive sang link xem ảnh trực tiếp
      var directImgUrl = convertDriveLink(rawImg);
      
      var bibKey = rawBib.toLowerCase();
      if (!photosByBib[bibKey]) {
        photosByBib[bibKey] = [];
      }
      photosByBib[bibKey].push(directImgUrl);
      allPhotos.push({ bib: rawBib, img: directImgUrl });
    }
    
    // Trả về ảnh của đúng BIB nếu có truyền ?bib=...
    if (requestedBib) {
      var runnerPhotos = photosByBib[requestedBib] || [];
      return responseJSON({
        success: true,
        bib: requestedBib,
        count: runnerPhotos.length,
        photos: runnerPhotos
      });
    }
    
    // Trả về toàn bộ danh sách và từ điển gom theo BIB
    return responseJSON({
      success: true,
      total: allPhotos.length,
      photosByBib: photosByBib,
      data: allPhotos
    });
    
  } catch (err) {
    return responseJSON({ success: false, error: err.toString() });
  }
}

// Chuyển link Google Drive sang direct image stream
function convertDriveLink(url) {
  if (!url) return '';
  var match = url.match(/\\/file\\/d\\/([a-zA-Z0-9_-]+)/) || url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    return 'https://lh3.googleusercontent.com/d/' + match[1];
  }
  return url;
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
`;

// In-memory cache for fetched photos
const memoryPhotosCache: Record<string, string[]> = {};

/**
 * Lấy danh sách ảnh thi đấu của vận động viên theo số BIB
 */
export async function getRunnerRacePhotos(bib: string, race?: Race): Promise<string[]> {
  const cleanBib = (bib || '').trim().toLowerCase();
  if (!cleanBib) return [];

  const raceKey = race?.id || 'default';
  const cacheKey = `${raceKey}_${cleanBib}`;

  // 1. Kiểm tra cache trong bộ nhớ
  if (memoryPhotosCache[cacheKey] && memoryPhotosCache[cacheKey].length > 0) {
    return memoryPhotosCache[cacheKey];
  }

  // 2. Kiểm tra sessionStorage
  try {
    const sessionKey = `vm_race_photos_${cacheKey}`;
    const saved = sessionStorage.getItem(sessionKey);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed) && parsed.length > 0) {
        memoryPhotosCache[cacheKey] = parsed;
        return parsed;
      }
    }
  } catch {}

  // 3. Nếu có link Google Apps Script chuyên lấy ảnh cho giải này
  const photosUrl = race?.photosScriptUrl?.trim() || 'https://script.google.com/macros/s/AKfycbyUr1QYj9Eyp60HaDLhXJINbr8Yozt3TXMRlPHpJ7QWhpkK6D4D_ZGMhW5dUerljLT3/exec';
  if (photosUrl) {
    try {
      // Gọi qua backend proxy hoặc trực tiếp để bypass CORS
      const proxyEndpoint = `/api/race-photos?url=${encodeURIComponent(photosUrl)}&bib=${encodeURIComponent(cleanBib)}`;
      const resp = await fetch(proxyEndpoint);
      if (resp.ok) {
        const json = await resp.json();
        let extractedPhotos: string[] = [];

        if (Array.isArray(json.photos)) {
          extractedPhotos = json.photos;
        } else if (Array.isArray(json)) {
          extractedPhotos = json.map((item: any) => typeof item === 'string' ? item : item.img || item.url).filter(Boolean);
        } else if (json.photosByBib && json.photosByBib[cleanBib]) {
          extractedPhotos = json.photosByBib[cleanBib];
        }

        // Chuẩn hoá link Google Drive và tách các URL phân tách bằng dấu phẩy
        const cleanList: string[] = [];
        for (const raw of extractedPhotos) {
          if (!raw) continue;
          const parts = String(raw).split(/[\n,;]+/);
          for (const p of parts) {
            const trimmed = p.trim();
            if (trimmed.startsWith('http://') || trimmed.startsWith('https://')) {
              cleanList.push(getDirectGoogleDriveImageUrl(trimmed) || trimmed);
            }
          }
        }

        if (cleanList.length > 0) {
          memoryPhotosCache[cacheKey] = cleanList;
          try {
            sessionStorage.setItem(`vm_race_photos_${cacheKey}`, JSON.stringify(cleanList));
          } catch {}
          return cleanList;
        }
      }
    } catch (e) {
      console.warn('Lỗi khi tải ảnh thi đấu từ Apps Script:', e);
    }
  }

  // 4. Fallback vào demo race photos được cấu hình trong giải hoặc mockRunners
  const demoMap = race?.demoRacePhotos || DEMO_RACE_PHOTOS;
  if (demoMap && demoMap[cleanBib]) {
    const list = demoMap[cleanBib];
    memoryPhotosCache[cacheKey] = list;
    return list;
  }

  // Fallback 1 ảnh đơn nếu có
  const singleMap = race?.demoPhotos || DEMO_PHOTOS;
  if (singleMap && singleMap[cleanBib]) {
    const list = [singleMap[cleanBib]];
    memoryPhotosCache[cacheKey] = list;
    return list;
  }

  return [];
}

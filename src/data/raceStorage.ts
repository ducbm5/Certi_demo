import { Race, RACES, DEFAULT_RACE, ensureRaceRunners } from './races';
import { CertificatePlacements } from '../types';

const STORAGE_KEY_RACES = 'vm_custom_races_list_v1';

export interface CreateRacePayload {
  name: string;
  slug: string;
  code?: string;
  defaultBgUrl: string;
  appsScriptUrl?: string;
  photosScriptUrl?: string;
  city?: string;
  province?: string;
  date?: string;
  description?: string;
  placements?: CertificatePlacements;
}

/**
 * Lấy danh sách toàn bộ các giải đấu từ localStorage và static config
 */
export function getLocalRaces(): Race[] {
  if (typeof window === 'undefined') return RACES.map(ensureRaceRunners);
  try {
    const raw = localStorage.getItem(STORAGE_KEY_RACES);
    if (raw) {
      const parsed: Race[] = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed.map(ensureRaceRunners);
      }
    }
  } catch (err) {
    console.warn('Lỗi đọc danh sách giải từ localStorage:', err);
  }
  return RACES.map(ensureRaceRunners);
}

/**
 * Tải danh sách giải đấu từ backend API (/api/races)
 */
export async function fetchAllRaces(): Promise<Race[]> {
  try {
    const resp = await fetch('/api/races');
    if (resp.ok) {
      const data: Race[] = await resp.json();
      if (Array.isArray(data) && data.length > 0) {
        const fullRaces = data.map(ensureRaceRunners);
        // Đồng bộ vào localStorage
        if (typeof window !== 'undefined') {
          localStorage.setItem(STORAGE_KEY_RACES, JSON.stringify(fullRaces));
        }
        return fullRaces;
      }
    }
  } catch (err) {
    console.warn('Không thể tải danh sách giải từ API, dùng local fallback:', err);
  }
  return getLocalRaces();
}

/**
 * Phân giải giải đấu dựa trên đường dẫn URL (pathname, hash, query)
 */
export function resolveRaceFromPath(pathStr: string, raceList: Race[] = getLocalRaces()): Race {
  if (!pathStr) return ensureRaceRunners(raceList[0] || DEFAULT_RACE);

  const clean = pathStr.toLowerCase();

  // Kiểm tra query ?race=slug
  const queryMatch = clean.match(/[?&]race=([a-z0-9_-]+)/);
  if (queryMatch && queryMatch[1]) {
    const found = raceList.find((r) => r.slug.toLowerCase() === queryMatch[1] || r.id.toLowerCase() === queryMatch[1]);
    if (found) return ensureRaceRunners(found);
  }

  // Kiểm tra hash #slug
  const hashMatch = clean.match(/#([a-z0-9_-]+)/);
  if (hashMatch && hashMatch[1] && hashMatch[1] !== 'admin') {
    const found = raceList.find((r) => r.slug.toLowerCase() === hashMatch[1] || r.id.toLowerCase() === hashMatch[1]);
    if (found) return ensureRaceRunners(found);
  }

  // Kiểm tra pathname /ha-long-2026 hoặc /vm-ha-long-2026
  const pathname = clean.split('?')[0].split('#')[0].replace(/^\/+|\/+$/g, '');
  if (pathname && pathname !== 'admin') {
    const found = raceList.find((r) => r.slug.toLowerCase() === pathname || r.id.toLowerCase() === pathname);
    if (found) return ensureRaceRunners(found);
  }

  return ensureRaceRunners(raceList[0] || DEFAULT_RACE);
}

/**
 * Kiểm tra kết nối thử nghiệm đến Google Apps Script
 */
export async function testScriptConnection(scriptUrl: string): Promise<{ success: boolean; count?: number; message: string; sample?: any }> {
  if (!scriptUrl || !scriptUrl.trim()) {
    return { success: false, message: 'Vui lòng nhập đường link Google Apps Script' };
  }

  try {
    const cleanUrl = scriptUrl.trim();
    // Gọi thông qua proxy server để bypass CORS
    const proxyUrl = `/api/proxy-sheet?url=${encodeURIComponent(cleanUrl)}`;
    const resp = await fetch(proxyUrl);
    if (!resp.ok) {
      return { success: false, message: `Lỗi kết nối HTTP ${resp.status}: Máy chủ không phản hồi.` };
    }

    const text = await resp.text();
    const trimmed = text.trim();

    // Thử parse JSON
    if (trimmed.startsWith('{') || trimmed.startsWith('[')) {
      const data = JSON.parse(trimmed);
      const runnersList = Array.isArray(data) ? data : data.runners || data.data || [];
      const count = runnersList.length;
      return {
        success: true,
        count,
        message: `Kết nối thành công! Đã tìm thấy ${count} vận động viên trong dữ liệu giải.`,
        sample: runnersList[0],
      };
    }

    // Nếu trả về TSV / CSV
    const lines = trimmed.split(/\r?\n/).filter((l) => l.trim().length > 0);
    if (lines.length > 1) {
      return {
        success: true,
        count: lines.length - 1,
        message: `Kết nối thành công! Nhận được ${lines.length - 1} dòng dữ liệu từ Google Apps Script.`,
      };
    }

    return {
      success: false,
      message: 'Script phản hồi nhưng không chứa cấu trúc dữ liệu JSON/TSV hợp lệ.',
    };
  } catch (err: any) {
    return {
      success: false,
      message: err.message || 'Không thể kết nối đến đường dẫn Script đã cung cấp.',
    };
  }
}

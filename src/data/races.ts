import { Runner, CertificatePlacements } from '../types';
import { INITIAL_RUNNERS, DEMO_RUNNERS, DEMO_PHOTOS } from './mockRunners';

export interface Race {
  id: string;
  slug: string;
  code: string;
  name: string;
  shortName: string;
  city: string;
  province: string;
  locationFull: string;
  date: string;
  officialUrl: string;
  defaultLogoUrl: string;
  defaultBgUrl: string;
  accentColor: string;
  themeBadgeBg: string;
  themeDotBg: string;
  storageKeyPrefix: string;
  initialRunners?: Runner[];
  demoRunners?: Runner[];
  demoPhotos?: Record<string, string>;
  demoRacePhotos?: Record<string, string[]>; // Mảng nhiều ảnh thi đấu cho mỗi BIB
  appsScriptUrl?: string;
  photosScriptUrl?: string; // Link Google Apps Script Web App lấy dữ liệu ảnh từ Google Sheet (Cột BIB & IMG)
  description: string;
  placements?: CertificatePlacements;
}

// Clone runners for Nghệ An 2026 with consistent date and metadata
const CLONED_NGHE_AN_RUNNERS: Runner[] = INITIAL_RUNNERS.map((runner) => ({
  ...runner,
  date: '13/09/2026',
}));

const CLONED_NGHE_AN_DEMO_RUNNERS: Runner[] = DEMO_RUNNERS.map((runner) => ({
  ...runner,
  date: '13/09/2026',
}));

/**
 * Đảm bảo mọi giải đấu (kể cả tải động từ file .json hay API) luôn có danh sách vận động viên mặc định
 */
export function ensureRaceRunners(race: Race): Race {
  if (!race) return DEFAULT_RACE;
  const raceDate = race.date || '13/09/2026';
  const fallbackList = INITIAL_RUNNERS.map((r) => ({ ...r, date: raceDate }));
  return {
    ...race,
    initialRunners:
      race.initialRunners && Array.isArray(race.initialRunners) && race.initialRunners.length > 0
        ? race.initialRunners
        : fallbackList,
    demoRunners:
      race.demoRunners && Array.isArray(race.demoRunners) && race.demoRunners.length > 0
        ? race.demoRunners
        : fallbackList,
    demoPhotos:
      race.demoPhotos && Object.keys(race.demoPhotos).length > 0 ? race.demoPhotos : DEMO_PHOTOS,
  };
}

export const RACES: Race[] = [
  {
    id: 'nghe-an-2026',
    slug: 'nghe-an-2026',
    code: 'NA26',
    name: 'VnExpress Marathon Grand Tour Nghe An 2026',
    shortName: 'Grand Tour Nghe An 2026',
    city: 'TP. Vinh',
    province: 'Nghệ An',
    locationFull: 'TP. Vinh, Nghệ An',
    date: '13/09/2026',
    officialUrl: 'https://vm.vnexpress.net/nghe-an-2026',
    defaultLogoUrl: '/race_logo.png',
    defaultBgUrl: '/NA26.png',
    accentColor: '#0369a1', // sky-700
    themeBadgeBg: 'bg-sky-50 text-sky-700 border-sky-200/80',
    themeDotBg: 'bg-sky-600',
    storageKeyPrefix: 'vm_nghean',
    initialRunners: CLONED_NGHE_AN_RUNNERS,
    demoRunners: CLONED_NGHE_AN_DEMO_RUNNERS,
    demoPhotos: DEMO_PHOTOS,
    appsScriptUrl:
      'https://script.google.com/macros/s/AKfycbwwY2MgGaURMrB20UHGVvUZ3INSOrkd8jIQok1JpnDTWMzblecdDOdDTn7qtrbtPPzquw/exec?key=ducbm900966559155',
    photosScriptUrl:
      'https://script.google.com/macros/s/AKfycbyUr1QYj9Eyp60HaDLhXJINbr8Yozt3TXMRlPHpJ7QWhpkK6D4D_ZGMhW5dUerljLT3/exec',
    description: 'Tra cứu kết quả & Chứng nhận điện tử VnExpress Marathon Grand Tour Nghe An 2026',
  },
];

export const DEFAULT_RACE = RACES[0];

// Helper to resolve race from pathname or hash or query
export const getRaceFromPath = (_path: string): Race => {
  return RACES[0];
};

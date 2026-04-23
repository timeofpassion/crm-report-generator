export interface HospitalConfig {
  name: string;
  shortName: string;
  team: string;
  categoryKeywords: Record<string, string[]>;
}

export const HOSPITAL_CONFIGS: Record<string, HospitalConfig> = {
  mellow_sinsa: {
    name: '멜로우피부과 신사점',
    shortName: '멜로우',
    team: '열정의시간 마케팅팀',
    categoryKeywords: {
      '보톡스/필러': ['보톡스', '보툴리눔', '톡신', '필러', '쥬비덤', '레스틸렌', '엘란쎄', '테오시알',
        '이마주름', '미간주름', '눈가주름', '팔자주름', '입술', '코필러', '턱', '사각턱', '승모근', '종아리', '애교살'],
      '리프팅': ['실리프팅', '실매선', '써마지', '울세라', '울쎄라', '슈링크', '인모드', '포텐자',
        '올리지오', '하이푸', '리프팅', 'HIFU', '프로파운드', '더블로', '아이디얼'],
      '레이저': ['레이저', '피코', '피코슈어', 'IPL', '토닝', '프락셀', '스타워커', '클라리티',
        '루비', '엑시머', '어븀', 'CO2', '이산화탄소', '점빼기', '색소', '홍조', '혈관', '제모',
        '여드름레이저', '듀얼토닝', '스펙트라', '메디라이트'],
      '스킨케어': ['물광', '수분', '비타민', '엑소좀', '줄기세포', 'MTS', '더마펜', '피부관리',
        '각질', '모공', '화이트닝', '미백', '보습', '재생', '성장인자', '스킨부스터', '리쥬란', '필링'],
      '지방/체형': ['지방', '비만', '다이어트', '지방흡입', '지방분해', '지방용해', 'GLP', '삭센다',
        '위고비', '체형', '비만주사', '윤곽', '쿨스컬'],
      '수술': ['수술', '절개', '매몰', '쌍꺼풀', '눈매교정', '코성형', '안면윤곽', '지방이식',
        '이마거상', '안면리프트', '귀족'],
      '기타': ['상담', '재진', '처방', '약품', '원내처방'],
    },
  },
  new_hospital: {
    name: '신규 병원',
    shortName: '신규',
    team: '열정의시간 마케팅팀',
    categoryKeywords: {
      '보톡스/필러': [], '리프팅': [], '레이저': [], '스킨케어': [], '지방/체형': [], '수술': [], '기타': [],
    },
  },
};

export function categorize(orderName: string, keywords: Record<string, string[]>): string {
  const name = String(orderName || '');
  for (const [cat, kws] of Object.entries(keywords)) {
    if (kws.some(kw => name.includes(kw))) return cat;
  }
  return '기타';
}

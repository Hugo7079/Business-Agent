import PptxGenJS from 'pptxgenjs';
import { AnalysisResult } from '../types';

// ─── 文字工具 ────────────────────────────────────────────

/** 硬性截斷，超過就加省略號 */
const trunc = (s: string, max: number): string => {
  if (!s) return '—';
  const t = s.trim().replace(/[。！？!?.]+$/g, '').trim();
  return t.length <= max ? t : t.slice(0, max - 1) + '…';
};

/**
 * 主要的 bullet 拆分函式
 * 優先按 \n 拆（Gemini 新格式），否則按句號拆
 * 每條強制在 maxLen 字以內，最多 maxLines 條
 */
const toBullets = (text: string, maxLines = 5, maxLen = 22): string[] => {
  if (!text) return ['—'];

  // 優先用 \n 分割（Gemini 新 schema 輸出格式）
  const byNewline = text.split(/\n+/).map(s => s.trim()).filter(s => s.length > 1);
  if (byNewline.length >= 2) {
    return byNewline.slice(0, maxLines).map(s => trunc(s, maxLen));
  }

  // fallback：按句號拆，再按逗號取第一段
  const bySentence = text
    .replace(/([。！？!?；;])/g, '$1\n')
    .split('\n')
    .map(s => s.replace(/[。！？!?；;]+$/g, '').trim())
    .filter(s => s.length > 1);

  const results: string[] = [];
  for (const sentence of bySentence) {
    if (results.length >= maxLines) break;
    if (sentence.length <= maxLen) {
      results.push(sentence);
    } else {
      // 按逗號拆，取最短有意義的子句
      const parts = sentence.split(/[，,、]/).map(p => p.trim()).filter(p => p.length > 1);
      results.push(trunc(parts[0] || sentence, maxLen));
    }
  }
  return results.length ? results : [trunc(text, maxLen)];
};

/**
 * 單行標籤：優先取 \n 第一行，否則按逗號取第一段
 * 用在 strength / weakness / concern / mitigation / KPI 值等單欄位
 */
const toTag = (text: string, maxLen = 20): string => {
  if (!text) return '—';
  // 有換行就取第一行
  const firstLine = text.split(/\n/)[0].trim();
  const cleaned = firstLine.replace(/[。！？!?.]+$/g, '').trim();
  if (cleaned.length <= maxLen) return cleaned;
  // 按逗號拆取第一段
  const parts = cleaned.split(/[，,、]/).map(p => p.trim()).filter(p => p.length > 0);
  return trunc(parts[0] || cleaned, maxLen);
};

/**
 * 引言專用：取整段中最短有力的一句（≤ maxLen 字）
 */
const toQuote = (text: string, maxLen = 18): string => {
  if (!text) return '—';
  // 有換行先取第一行
  const firstLine = text.split(/\n/)[0].trim().replace(/[。！？!?.]+$/g, '');
  if (firstLine.length <= maxLen) return firstLine;
  // 按句號拆，取最短的
  const sentences = text
    .replace(/([。！？!?])/g, '$1\n')
    .split('\n')
    .map(s => s.replace(/[。！？!?]+$/g, '').trim())
    .filter(s => s.length >= 4);
  const shortest = sentences.sort((a, b) => a.length - b.length)[0] || firstLine;
  return trunc(shortest, maxLen);
};

const fmt = (n: number) => {
  const abs = Math.abs(n);
  if (abs >= 1_000_000) return `$${(n / 1_000_000).toFixed(1)}M`;
  if (abs >= 1_000) return `$${(n / 1_000).toFixed(0)}K`;
  return `$${n}`;
};

// ─── 產業主題偵測 ────────────────────────────────────────
interface Theme {
  name: string;
  bg: string;        // 主背景
  card: string;      // 卡片背景
  accent: string;    // 主強調色
  accent2: string;   // 副強調色
  accent3: string;   // 第三色
  headerBg: string;  // 表格 header
  divider: string;   // 分隔線
  tagline: string;   // 封面副標語風格詞
}

const THEMES: Record<string, Theme> = {
  tech: {
    name: '科技', bg: '0A0F1E', card: '111827', accent: '3B82F6', accent2: '06B6D4',
    accent3: '8B5CF6', headerBg: '1E3A5F', divider: '1E3A5F', tagline: 'TECHNOLOGY × INNOVATION',
  },
  health: {
    name: '醫療健康', bg: '071A14', card: '0D2418', accent: '10B981', accent2: '34D399',
    accent3: '06B6D4', headerBg: '0D3320', divider: '134D2E', tagline: 'HEALTH × WELLNESS',
  },
  food: {
    name: '餐飲食品', bg: '1A0C00', card: '2D1500', accent: 'F97316', accent2: 'FBBF24',
    accent3: 'EF4444', headerBg: '3D1F00', divider: '4D2800', tagline: 'FOOD × LIFESTYLE',
  },
  finance: {
    name: '金融財務', bg: '090E1A', card: '0F1829', accent: 'F59E0B', accent2: '10B981',
    accent3: '3B82F6', headerBg: '1A2540', divider: '1E2D4D', tagline: 'FINANCE × GROWTH',
  },
  retail: {
    name: '零售電商', bg: '0F0A1A', card: '1A1029', accent: 'A855F7', accent2: 'EC4899',
    accent3: 'F97316', headerBg: '2D1040', divider: '3D1555', tagline: 'RETAIL × EXPERIENCE',
  },
  edu: {
    name: '教育', bg: '0A1020', card: '0F1A2E', accent: '6366F1', accent2: '8B5CF6',
    accent3: '06B6D4', headerBg: '1A2040', divider: '1E2D55', tagline: 'EDUCATION × FUTURE',
  },
  default: {
    name: '商業', bg: '0F172A', card: '1E293B', accent: '3B82F6', accent2: '10B981',
    accent3: '7C3AED', headerBg: '1E3A5F', divider: '334155', tagline: 'BUSINESS × STRATEGY',
  },
};

const detectTheme = (idea: string, marketData: string): Theme => {
  const text = (idea + marketData).toLowerCase();
  if (/醫|健|藥|療|病|診|wellness|health|medical|biotech/.test(text)) return THEMES.health;
  if (/餐|食|飲|咖|料理|外送|food|restaurant|beverage|cafe/.test(text)) return THEMES.food;
  if (/金融|投資|貸|保險|fintech|finance|bank|crypto|blockchain/.test(text)) return THEMES.finance;
  if (/零售|電商|購物|fashion|retail|ecommerce|商城/.test(text)) return THEMES.retail;
  if (/教育|學習|課程|edu|learning|school|training|輔導/.test(text)) return THEMES.edu;
  if (/科技|ai|app|平台|軟體|saas|tech|software|雲端|cloud|iot/.test(text)) return THEMES.tech;
  return THEMES.default;
};

// ─── 共用繪製函式 ─────────────────────────────────────────
const addBg = (slide: PptxGenJS.Slide, t: Theme) => {
  slide.background = { fill: t.bg };
  // 頂部色條
  slide.addShape('rect', { x: 0, y: 0, w: '100%', h: 0.055, fill: { type: 'solid', color: t.accent } });
};

const addHeader = (slide: PptxGenJS.Slide, t: Theme, title: string) => {
  slide.addText(title, {
    x: 0.45, y: 0.12, w: 7.5, h: 0.52,
    fontSize: 20, bold: true, color: 'F8FAFC',
    fontFace: 'Arial',
  });
  // 底線
  slide.addShape('rect', { x: 0.45, y: 0.66, w: 0.9, h: 0.03, fill: { type: 'solid', color: t.accent } });
};

const addPageNum = (slide: PptxGenJS.Slide, n: number, total: number, t: Theme) => {
  slide.addText(`${n} / ${total}`, {
    x: 8.5, y: 7.0, w: 0.9, h: 0.28,
    fontSize: 8, color: t.divider, align: 'right',
  });
};

const addCard = (
  slide: PptxGenJS.Slide, t: Theme,
  x: number, y: number, w: number, h: number,
  borderColor?: string,
) => {
  slide.addShape('roundRect', {
    x, y, w, h, rectRadius: 0.12,
    fill: { type: 'solid', color: t.card },
    line: { color: borderColor || t.divider, width: 0.75 },
  });
};

const scoreColor = (s: number) =>
  s >= 80 ? '10B981' : s >= 60 ? '3B82F6' : s >= 40 ? 'FBBF24' : 'EF4444';

// ─── 安全邊界常數（16:9 = 10" × 7.5"）─────────────────────
const SLIDE_W = 10;
const SLIDE_H = 7.5;
const MARGIN = 0.45;                  // 左右 margin
const SAFE_T = 0.85;                  // 內容起始 Y（header 下方）
const SAFE_B = 6.8;                   // 內容底線（留空間給頁碼）
const CONTENT_H = SAFE_B - SAFE_T;   // = 5.95，可用內容高度

// ─── 主函式 ──────────────────────────────────────────────
export const generatePptx = async (result: AnalysisResult): Promise<void> => {
  const theme = detectTheme(
    result.executiveSummary,
    result.marketAnalysis.description,
  );

  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';
  pptx.author = 'OmniView AI';
  pptx.title = 'OmniView 商業提案分析報告';

  const TOTAL = 10;
  const T = theme;

  // ════════════════════════════════════════════════════
  // SLIDE 1 — COVER
  // ════════════════════════════════════════════════════
  {
    const slide = pptx.addSlide();
    slide.background = { fill: T.bg };

    // 幾何裝飾（主題色調）— 控制在畫布範圍內
    slide.addShape('rect', { x: 0, y: 0, w: '100%', h: 0.055, fill: { type: 'solid', color: T.accent } });
    slide.addShape('ellipse', { x: 6.8, y: -0.5, w: 3.0, h: 3.0, fill: { type: 'solid', color: T.card } });
    slide.addShape('ellipse', { x: -0.8, y: 4.8, w: 2.5, h: 2.5, fill: { type: 'solid', color: T.card } });

    // OmniView badge
    slide.addShape('roundRect', { x: 0.5, y: 0.4, w: 0.6, h: 0.6, rectRadius: 0.1, fill: { type: 'solid', color: T.accent } });
    slide.addText('OV', { x: 0.5, y: 0.4, w: 0.6, h: 0.6, fontSize: 14, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
    slide.addText('OmniView AI', { x: 1.25, y: 0.45, w: 3, h: 0.5, fontSize: 16, bold: true, color: 'F8FAFC' });

    // Tagline chip
    slide.addShape('roundRect', { x: 0.5, y: 1.3, w: 3.2, h: 0.32, rectRadius: 0.16, fill: { type: 'solid', color: T.accent + '28' }, line: { color: T.accent + '60', width: 0.75 } });
    slide.addText(T.tagline, { x: 0.5, y: 1.3, w: 3.2, h: 0.32, fontSize: 8.5, bold: true, color: T.accent, align: 'center', valign: 'middle', charSpacing: 1.5 });

    // 主標題
    slide.addText('商業提案\n分析報告', {
      x: 0.5, y: 1.85, w: 8.5, h: 1.8,
      fontSize: 42, bold: true, color: 'F8FAFC', lineSpacingMultiple: 1.1,
    });

    // 成功機率 badge
    const sc = result.successProbability;
    const sC = scoreColor(sc);
    slide.addShape('roundRect', { x: 0.5, y: 3.9, w: 2.4, h: 0.9, rectRadius: 0.14, fill: { type: 'solid', color: sC + '22' }, line: { color: sC, width: 1.5 } });
    slide.addText('AI 評估成功機率', { x: 0.5, y: 3.95, w: 2.4, h: 0.28, fontSize: 9, color: sC, align: 'center' });
    slide.addText(`${sc}%`, { x: 0.5, y: 4.22, w: 2.4, h: 0.5, fontSize: 26, bold: true, color: sC, align: 'center', valign: 'middle' });

    // 核心摘要（精簡版）— 用 toTag 抓重點而非截斷
    const summaryLine = toTag(result.executiveSummary, 80);
    slide.addText(summaryLine, {
      x: 0.5, y: 5.1, w: 9.0, h: 0.6,
      fontSize: 12, color: '94A3B8', lineSpacingMultiple: 1.4, italic: true,
    });

    // 底部 — 底邊 = 6.95 + 0.35 = 7.3，在 7.5 之內
    slide.addShape('rect', { x: 0, y: 6.95, w: '100%', h: 0.35, fill: { type: 'solid', color: T.card } });
    slide.addText('由 OmniView AI 360° 虛擬董事會自動生成', { x: 0.5, y: 6.98, w: 6, h: 0.28, fontSize: 9, color: '475569' });
    addPageNum(slide, 1, TOTAL, T);
  }

  // ════════════════════════════════════════════════════
  // SLIDE 2 — EXECUTIVE SUMMARY
  // ════════════════════════════════════════════════════
  {
    const slide = pptx.addSlide();
    addBg(slide, T);
    addHeader(slide, T, '執行摘要');

    // 左欄：摘要 bullets — 底邊 = 0.85 + 5.2 = 6.05
    const leftCardH = Math.min(5.2, CONTENT_H);
    addCard(slide, T, 0.45, SAFE_T, 5.4, leftCardH);
    slide.addText('核心觀點', { x: 0.65, y: 0.97, w: 5.0, h: 0.3, fontSize: 11, bold: true, color: T.accent });
    const bullets = toBullets(result.executiveSummary, 5, 50);
    const bRows = bullets.map(b => [
      { text: '▸  ', options: { color: T.accent, bold: true, fontSize: 12 } },
      { text: b, options: { color: 'CBD5E1', fontSize: 12 } },
    ]);
    bRows.forEach((row, i) => {
      const by = 1.35 + i * 0.82;
      if (by + 0.7 > SAFE_T + leftCardH - 0.1) return; // 超出卡片就跳過
      slide.addText(row, { x: 0.65, y: by, w: 5.0, h: 0.7, lineSpacingMultiple: 1.4, valign: 'middle' });
    });

    // 右欄：關鍵指標 — 3 個卡片均分，底邊 ≤ SAFE_B
    const kpis = [
      { label: '市場規模 (TAM)', val: trunc(result.marketAnalysis.size, 20), color: T.accent },
      { label: '成長率 (CAGR)', val: trunc(result.marketAnalysis.growthRate, 20), color: T.accent2 },
      { label: '損益平衡', val: trunc(result.breakEvenPoint, 20), color: T.accent3 },
    ];
    const kpiGap = 0.15;
    const kpiH = (CONTENT_H - kpiGap * 2) / 3; // = (5.95 - 0.3) / 3 ≈ 1.88
    kpis.forEach((kpi, i) => {
      const y = SAFE_T + i * (kpiH + kpiGap);
      addCard(slide, T, 6.05, y, 3.35, kpiH, kpi.color + '55');
      slide.addShape('rect', { x: 6.05, y, w: 0.04, h: kpiH, fill: { type: 'solid', color: kpi.color } });
      slide.addText(kpi.label, { x: 6.25, y: y + 0.15, w: 3.0, h: 0.28, fontSize: 10, color: '94A3B8' });
      slide.addText(kpi.val, { x: 6.25, y: y + 0.48, w: 3.0, h: kpiH - 0.65, fontSize: 16, bold: true, color: kpi.color, lineSpacingMultiple: 1.2 });
    });

    addPageNum(slide, 2, TOTAL, T);
  }

  // ════════════════════════════════════════════════════
  // SLIDE 3 — MARKET OPPORTUNITY
  // ════════════════════════════════════════════════════
  {
    const slide = pptx.addSlide();
    addBg(slide, T);
    addHeader(slide, T, '市場機會');

    // 3 個 KPI 橫排
    const kpis = [
      { label: 'TAM 市場規模', val: result.marketAnalysis.size, color: T.accent },
      { label: 'CAGR 年均成長', val: result.marketAnalysis.growthRate, color: T.accent2 },
      { label: '損益平衡點', val: result.breakEvenPoint, color: T.accent3 },
    ];
    const kpiCardW = 2.8;
    const kpiGap = 0.3;
    kpis.forEach((k, i) => {
      const x = 0.45 + i * (kpiCardW + kpiGap);
      addCard(slide, T, x, SAFE_T, kpiCardW, 1.3, k.color + '66');
      slide.addShape('rect', { x, y: SAFE_T, w: kpiCardW, h: 0.04, fill: { type: 'solid', color: k.color } });
      slide.addText(k.label, { x: x + 0.15, y: SAFE_T + 0.1, w: kpiCardW - 0.3, h: 0.28, fontSize: 10, color: '94A3B8' });
      slide.addText(trunc(k.val, 22), { x: x + 0.15, y: SAFE_T + 0.42, w: kpiCardW - 0.3, h: 0.65, fontSize: 17, bold: true, color: k.color, lineSpacingMultiple: 1.1 });
    });

    // 市場描述 bullets — 頂 = 0.85+1.3+0.2 = 2.35，底 = SAFE_B = 6.8，高 = 4.45
    const mCardTop = SAFE_T + 1.3 + 0.2;
    const mCardH = SAFE_B - mCardTop; // = 6.8 - 2.35 = 4.45
    addCard(slide, T, 0.45, mCardTop, 9.0, mCardH);
    slide.addText('市場洞察', { x: 0.65, y: mCardTop + 0.12, w: 8.6, h: 0.3, fontSize: 11, bold: true, color: T.accent });
    const mBullets = toBullets(result.marketAnalysis.description, 5, 65);
    mBullets.forEach((b, i) => {
      const by = mCardTop + 0.52 + i * 0.62;
      if (by + 0.55 > mCardTop + mCardH - 0.1) return; // 超出卡片跳過
      slide.addText([
        { text: '◆  ', options: { color: T.accent2, bold: true, fontSize: 11 } },
        { text: b, options: { color: 'CBD5E1', fontSize: 12 } },
      ], { x: 0.65, y: by, w: 8.6, h: 0.55, valign: 'middle' });
    });

    addPageNum(slide, 3, TOTAL, T);
  }

  // ════════════════════════════════════════════════════
  // SLIDE 4 — FINANCIAL PROJECTIONS
  // ════════════════════════════════════════════════════
  {
    const slide = pptx.addSlide();
    addBg(slide, T);
    addHeader(slide, T, '財務預測');

    // 表格：最多 4 筆資料（控制高度）
    const finData = result.financials.slice(0, 4);

    const fHeader: PptxGenJS.TableCell[] = ['年度', '營收', '成本', '淨利'].map(t => ({
      text: t,
      options: { bold: true, color: 'FFFFFF', fill: { color: T.headerBg }, fontSize: 12, align: 'center' as const },
    }));

    const fRows: PptxGenJS.TableCell[][] = [fHeader];
    finData.forEach((f, i) => {
      const bg = i % 2 === 0 ? T.card : T.bg;
      const profit = Number(f.profit) || 0;
      fRows.push([
        { text: trunc(f.year ?? '—', 12), options: { color: 'F8FAFC', fill: { color: bg }, fontSize: 12, align: 'center' as const, bold: true } },
        { text: fmt(Number(f.revenue) || 0), options: { color: T.accent, fill: { color: bg }, fontSize: 12, align: 'center' as const, bold: true } },
        { text: fmt(Number(f.costs) || 0), options: { color: '94A3B8', fill: { color: bg }, fontSize: 12, align: 'center' as const } },
        { text: fmt(profit), options: { color: profit >= 0 ? T.accent2 : 'EF4444', fill: { color: bg }, fontSize: 12, align: 'center' as const, bold: true } },
      ]);
    });

    const tableRowH = 0.45;
    const dataRowH = 0.48;
    const tableBottom = SAFE_T + tableRowH + finData.length * dataRowH;

    slide.addTable(fRows, {
      x: MARGIN, y: SAFE_T, w: 9.0,
      border: { type: 'solid', color: T.divider, pt: 0.5 },
      colW: [2.25, 2.25, 2.25, 2.25],
      rowH: [tableRowH, ...finData.map(() => dataRowH)],
    });

    // 視覺化長條（營收趨勢）
    const maxRev = Math.max(...result.financials.map(f => Number(f.revenue) || 1));
    const barCardTop = tableBottom + 0.25;
    const barCardH = SAFE_B - barCardTop; // 動態計算到安全底線

    if (barCardH > 0.8) { // 只在有足夠空間時才畫
      addCard(slide, T, MARGIN, barCardTop, 9.0, barCardH);
      slide.addText('營收趨勢', { x: 0.65, y: barCardTop + 0.1, w: 3, h: 0.25, fontSize: 10, bold: true, color: T.accent });
      slide.addText(`損益平衡：${trunc(result.breakEvenPoint, 28)}`, { x: 4.5, y: barCardTop + 0.1, w: 4.75, h: 0.25, fontSize: 10, color: T.accent2, align: 'right' });

      const barStartX = 1.5;
      const barMaxW = 7.0;
      const barH = 0.28;
      const barGap = 0.42;
      finData.forEach((f, i) => {
        const rev = Number(f.revenue) || 0;
        const ratio = maxRev > 0 ? rev / maxRev : 0;
        const bw = Math.max(ratio * barMaxW, 0.15);
        const by = barCardTop + 0.45 + i * barGap;
        if (by + barH > barCardTop + barCardH - 0.08) return;
        slide.addText(trunc(f.year ?? '—', 8), { x: 0.65, y: by, w: 0.75, h: barH, fontSize: 9, color: '94A3B8', valign: 'middle' });
        slide.addShape('roundRect', { x: barStartX, y: by + 0.02, w: bw, h: barH - 0.04, rectRadius: 0.05, fill: { type: 'solid', color: T.accent } });
        const labelX = (barStartX + bw + 0.1 + 1.0 > 9.25) ? barStartX + bw - 1.1 : barStartX + bw + 0.1;
        const labelColor = (barStartX + bw + 0.1 + 1.0 > 9.25) ? 'FFFFFF' : T.accent;
        slide.addText(fmt(rev), { x: labelX, y: by, w: 1.0, h: barH, fontSize: 9, color: labelColor, bold: true, valign: 'middle' });
      });
    }

    addPageNum(slide, 4, TOTAL, T);
  }

  // ════════════════════════════════════════════════════
  // SLIDE 5 — COMPETITIVE LANDSCAPE
  // ════════════════════════════════════════════════════
  {
    const slide = pptx.addSlide();
    addBg(slide, T);
    addHeader(slide, T, '競爭態勢分析');

    const comps = result.competitors.slice(0, 5);
    const headerRowH = 0.42;
    const maxDataH = SAFE_B - SAFE_T - headerRowH;
    const rowH = Math.min(0.7, maxDataH / comps.length);

    const cHeader: PptxGenJS.TableCell[] = [
      { text: '競爭對手', options: { bold: true, color: 'FFFFFF', fill: { color: T.headerBg }, fontSize: 12, align: 'center' as const } },
      { text: '優勢', options: { bold: true, color: 'FFFFFF', fill: { color: T.headerBg }, fontSize: 12, align: 'center' as const } },
      { text: '劣勢', options: { bold: true, color: 'FFFFFF', fill: { color: T.headerBg }, fontSize: 12, align: 'center' as const } },
    ];
    const cRows: PptxGenJS.TableCell[][] = [cHeader];
    comps.forEach((c, i) => {
      const bg = i % 2 === 0 ? T.card : T.bg;
      cRows.push([
        { text: toTag(c.name, 16), options: { bold: true, color: 'F8FAFC', fill: { color: bg }, fontSize: 11, align: 'center' as const } },
        { text: toTag(c.strength, 30), options: { color: T.accent2, fill: { color: bg }, fontSize: 10 } },
        { text: toTag(c.weakness, 30), options: { color: 'F87171', fill: { color: bg }, fontSize: 10 } },
      ]);
    });

    slide.addTable(cRows, {
      x: MARGIN, y: SAFE_T, w: 9.0,
      border: { type: 'solid', color: T.divider, pt: 0.5 },
      colW: [2.1, 3.45, 3.45],
      rowH: [headerRowH, ...comps.map(() => rowH)],
    });

    addPageNum(slide, 5, TOTAL, T);
  }

  // ════════════════════════════════════════════════════
  // SLIDE 6 — STRATEGIC ROADMAP
  // ════════════════════════════════════════════════════
  {
    const slide = pptx.addSlide();
    addBg(slide, T);
    addHeader(slide, T, '策略路線圖');

    const items = result.roadmap.slice(0, 4);
    const itemH = CONTENT_H / items.length;

    items.forEach((item, i) => {
      const y = SAFE_T + i * itemH;

      // 時間軸線
      if (i < items.length - 1) {
        const lineTop = y + 0.3;
        const lineH = Math.min(itemH, SAFE_B - lineTop);
        slide.addShape('rect', { x: 0.62, y: lineTop, w: 0.03, h: lineH, fill: { type: 'solid', color: T.divider } });
      }
      // 節點圓
      slide.addShape('ellipse', { x: 0.5, y: y + 0.12, w: 0.26, h: 0.26, fill: { type: 'solid', color: T.accent } });
      slide.addText(`${i + 1}`, { x: 0.5, y: y + 0.12, w: 0.26, h: 0.26, fontSize: 9, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });

      // Phase + timeframe
      slide.addText(toTag(item.phase, 18), { x: 0.95, y, w: 3.5, h: 0.32, fontSize: 13, bold: true, color: T.accent });
      slide.addShape('roundRect', { x: 4.6, y: y + 0.03, w: 1.5, h: 0.26, rectRadius: 0.13, fill: { type: 'solid', color: T.accent + '25' }, line: { color: T.accent + '60', width: 0.5 } });
      slide.addText(toTag(item.timeframe, 14), { x: 4.6, y: y + 0.03, w: 1.5, h: 0.26, fontSize: 9, color: T.accent, align: 'center', valign: 'middle' });

      // 內容卡
      const cardTop = y + 0.38;
      const cardH = Math.min(itemH - 0.43, SAFE_B - cardTop);
      if (cardH < 0.4) return;
      addCard(slide, T, 0.95, cardTop, 8.4, cardH);

      // 產品（左半）— 用 toTag 取關鍵詞
      slide.addText('產品', { x: 1.15, y: cardTop + 0.08, w: 0.7, h: 0.22, fontSize: 9, bold: true, color: T.accent2 });
      slide.addText(toTag(item.product, 38), { x: 1.15, y: cardTop + 0.3, w: 3.7, h: cardH - 0.38, fontSize: 10, color: 'CBD5E1', lineSpacingMultiple: 1.3, valign: 'top' });
      // 分隔
      slide.addShape('rect', { x: 5.1, y: cardTop + 0.08, w: 0.02, h: cardH - 0.16, fill: { type: 'solid', color: T.divider } });
      // 技術（右半）— 用 toTag 取關鍵詞
      slide.addText('技術', { x: 5.25, y: cardTop + 0.08, w: 0.7, h: 0.22, fontSize: 9, bold: true, color: T.accent3 });
      slide.addText(toTag(item.technology, 38), { x: 5.25, y: cardTop + 0.3, w: 3.9, h: cardH - 0.38, fontSize: 10, color: 'CBD5E1', lineSpacingMultiple: 1.3, valign: 'top' });
    });

    addPageNum(slide, 6, TOTAL, T);
  }

  // ════════════════════════════════════════════════════
  // SLIDE 7 — RISK ASSESSMENT
  // ════════════════════════════════════════════════════
  {
    const slide = pptx.addSlide();
    addBg(slide, T);
    addHeader(slide, T, '風險評估');

    const risks = result.risks.slice(0, 6);
    const cols = 2;
    const rows = Math.ceil(risks.length / cols);
    const gap = 0.12;
    const cardW = (9.0 - gap) / cols;
    const cardH = (CONTENT_H - gap * (rows - 1)) / rows;

    risks.forEach((r, i) => {
      const col = i % cols;
      const row = Math.floor(i / cols);
      const x = MARGIN + col * (cardW + gap);
      const y = SAFE_T + row * (cardH + gap);
      if (y + cardH > SAFE_B) return;
      const ic = r.impact === 'High' ? 'EF4444' : r.impact === 'Medium' ? 'FBBF24' : '10B981';
      const label = r.impact === 'High' ? '高' : r.impact === 'Medium' ? '中' : '低';

      addCard(slide, T, x, y, cardW, cardH, ic + '55');
      slide.addShape('rect', { x, y, w: 0.04, h: cardH, fill: { type: 'solid', color: ic } });
      slide.addShape('roundRect', { x: x + cardW - 0.95, y: y + 0.1, w: 0.82, h: 0.26, rectRadius: 0.13, fill: { type: 'solid', color: ic + '30' }, line: { color: ic, width: 0.5 } });
      slide.addText(label, { x: x + cardW - 0.95, y: y + 0.1, w: 0.82, h: 0.26, fontSize: 9, bold: true, color: ic, align: 'center', valign: 'middle' });
      // 風險標題 — 用 toTag 精簡
      slide.addText(toTag(r.risk, 24), { x: x + 0.18, y: y + 0.1, w: cardW - 1.2, h: 0.3, fontSize: 11, bold: true, color: 'F8FAFC' });
      // 因應策略 — 用 toTag 精簡
      slide.addText([
        { text: '因應：', options: { fontSize: 9, color: '64748B', bold: true } },
        { text: toTag(r.mitigation, 40), options: { fontSize: 9, color: 'CBD5E1' } },
      ], { x: x + 0.18, y: y + 0.45, w: cardW - 0.35, h: cardH - 0.55, lineSpacingMultiple: 1.35, valign: 'top' });
    });

    addPageNum(slide, 7, TOTAL, T);
  }

  // ════════════════════════════════════════════════════
  // SLIDE 8 — STAKEHOLDER BOARD
  // ════════════════════════════════════════════════════
  {
    const slide = pptx.addSlide();
    addBg(slide, T);
    addHeader(slide, T, 'AI 虛擬董事會');

    const personas = result.personaEvaluations.slice(0, 5);
    const cols = Math.min(personas.length, 3);
    const rows = Math.ceil(personas.length / cols);
    const gap = 0.12;
    const cardW = (9.0 - gap * (cols - 1)) / cols;
    const cardH = (CONTENT_H - gap * (rows - 1)) / rows;

    personas.forEach((p, i) => {
      const col = i % cols;
      const row = Math.floor(i / cols);
      const x = MARGIN + col * (cardW + gap);
      const y = SAFE_T + row * (cardH + gap);
      if (y + cardH > SAFE_B + 0.05) return;
      const sc = Number(p.score) || 0;
      const sC = scoreColor(sc);

      addCard(slide, T, x, y, cardW, cardH);
      slide.addShape('rect', { x, y, w: cardW, h: 0.04, fill: { type: 'solid', color: sC } });

      // 角色名 + 分數
      slide.addText(toTag(p.role, 12), { x: x + 0.15, y: y + 0.1, w: cardW - 1.1, h: 0.3, fontSize: 12, bold: true, color: 'F8FAFC' });
      slide.addText(`${sc}`, { x: x + cardW - 0.95, y: y + 0.08, w: 0.8, h: 0.32, fontSize: 18, bold: true, color: sC, align: 'right' });

      // 分數條
      const barW = (cardW - 0.3) * (sc / 100);
      slide.addShape('roundRect', { x: x + 0.15, y: y + 0.45, w: cardW - 0.3, h: 0.06, rectRadius: 0.03, fill: { type: 'solid', color: T.divider } });
      slide.addShape('roundRect', { x: x + 0.15, y: y + 0.45, w: Math.max(barW, 0.05), h: 0.06, rectRadius: 0.03, fill: { type: 'solid', color: sC } });

      // 引言 — 用 toQuote 提取最精華的一句
      const quoteH = Math.min(rows === 1 ? 0.85 : 0.55, cardH - 1.2);
      slide.addText(`"${toQuote(p.keyQuote, 28)}"`, {
        x: x + 0.15, y: y + 0.58, w: cardW - 0.3, h: quoteH,
        fontSize: 9.5, italic: true, color: '94A3B8', lineSpacingMultiple: 1.3, valign: 'top',
      });

      // 擔憂 — 用 toTag 精簡
      const concernY = y + 0.58 + quoteH + 0.05;
      const concernH = cardH - (concernY - y) - 0.08;
      if (concernH > 0.2) {
        slide.addText([
          { text: '! ', options: { fontSize: 9, color: 'FBBF24', bold: true } },
          { text: toTag(p.concern, 38), options: { fontSize: 9, color: 'CBD5E1' } },
        ], { x: x + 0.15, y: concernY, w: cardW - 0.3, h: concernH, lineSpacingMultiple: 1.3, valign: 'top' });
      }
    });

    addPageNum(slide, 8, TOTAL, T);
  }

  // ════════════════════════════════════════════════════
  // SLIDE 9 — FINAL VERDICTS
  // ════════════════════════════════════════════════════
  {
    const slide = pptx.addSlide();
    addBg(slide, T);
    addHeader(slide, T, '最終裁決');

    const verdicts = [
      { title: '激進觀點', text: result.finalVerdicts.aggressive, color: 'F97316' },
      { title: '平衡觀點', text: result.finalVerdicts.balanced, color: T.accent },
      { title: '保守觀點', text: result.finalVerdicts.conservative, color: T.accent3 },
    ];
    const gap = 0.18;
    const vCardW = (9.0 - gap * 2) / 3;
    const vCardH = CONTENT_H; // = 5.95，底邊 = 0.85 + 5.95 = 6.8

    verdicts.forEach((v, i) => {
      const x = MARGIN + i * (vCardW + gap);
      addCard(slide, T, x, SAFE_T, vCardW, vCardH, v.color + '55');
      slide.addShape('rect', { x, y: SAFE_T, w: vCardW, h: 0.04, fill: { type: 'solid', color: v.color } });
      slide.addText(v.title, { x: x + 0.15, y: SAFE_T + 0.1, w: vCardW - 0.3, h: 0.35, fontSize: 13, bold: true, color: v.color });
      slide.addShape('rect', { x: x + 0.15, y: SAFE_T + 0.48, w: 0.8, h: 0.03, fill: { type: 'solid', color: v.color } });
      // Content bullets
      const vBullets = toBullets(v.text, 6, 33);
      vBullets.forEach((b, bi) => {
        const bulletY = SAFE_T + 0.62 + bi * 0.65;
        if (bulletY + 0.58 > SAFE_T + vCardH - 0.08) return;
        slide.addText([
          { text: '▸ ', options: { color: v.color, bold: true, fontSize: 10 } },
          { text: b, options: { color: 'CBD5E1', fontSize: 10 } },
        ], { x: x + 0.15, y: bulletY, w: vCardW - 0.3, h: 0.58, lineSpacingMultiple: 1.3, valign: 'middle' });
      });
    });

    addPageNum(slide, 9, TOTAL, T);
  }

  // ════════════════════════════════════════════════════
  // SLIDE 10 — CALL TO ACTION
  // ════════════════════════════════════════════════════
  {
    const slide = pptx.addSlide();
    slide.background = { fill: T.bg };

    slide.addShape('rect', { x: 0, y: 0, w: '100%', h: 0.055, fill: { type: 'solid', color: T.accent } });
    slide.addShape('ellipse', { x: 3.0, y: 1.0, w: 4.0, h: 4.0, fill: { type: 'solid', color: T.card } });

    slide.addText('下一步行動', {
      x: 0.5, y: 1.5, w: 9.0, h: 0.9,
      fontSize: 40, bold: true, color: 'F8FAFC', align: 'center',
    });
    slide.addText(T.tagline, {
      x: 0.5, y: 2.45, w: 9.0, h: 0.38,
      fontSize: 11, color: T.accent, align: 'center', charSpacing: 2.5, bold: true,
    });

    // 成功機率卡 — 底邊 = 3.1 + 1.4 = 4.5
    const sc = result.successProbability;
    const sC = scoreColor(sc);
    addCard(slide, T, 3.6, 3.1, 2.8, 1.4, sC);
    slide.addText('AI 評估成功機率', { x: 3.6, y: 3.2, w: 2.8, h: 0.28, fontSize: 10, color: '94A3B8', align: 'center' });
    slide.addText(`${sc}%`, { x: 3.6, y: 3.5, w: 2.8, h: 0.8, fontSize: 34, bold: true, color: sC, align: 'center' });

    // 底部文字 — y = 6.2，底邊 = 6.5
    slide.addText('此報告由 OmniView AI 360° 虛擬董事會自動生成', {
      x: 1, y: 6.2, w: 8, h: 0.3, fontSize: 10, color: '475569', align: 'center',
    });
    addPageNum(slide, 10, TOTAL, T);
  }

  // ════════════════════════════════════════════════════
  // EXPORT
  // ════════════════════════════════════════════════════
  await pptx.writeFile({ fileName: 'OmniView_商業提案報告.pptx' });
};

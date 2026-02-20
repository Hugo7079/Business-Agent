import PptxGenJS from 'pptxgenjs';
import { AnalysisResult } from '../types';

// ─── 文字工具 ────────────────────────────────────────────
const trunc = (s: string, max: number) =>
  s && s.length > max ? s.slice(0, max - 1) + '…' : (s || '—');

/** 把長段落切成精簡 bullet，每條最多 maxLen 字，最多 maxLines 條 */
const toBullets = (text: string, maxLines = 3, maxLen = 40): string[] => {
  const sentences = text
    .replace(/([。！？\.!?])/g, '$1|')
    .split('|')
    .map(s => s.trim())
    .filter(s => s.length > 2);
  const bullets: string[] = [];
  for (const s of sentences) {
    if (bullets.length >= maxLines) break;
    bullets.push(trunc(s, maxLen));
  }
  return bullets.length ? bullets : [trunc(text, maxLen)];
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
    x: 8.9, y: 7.05, w: 0.9, h: 0.28,
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

    // 如果有生成圖，作為背景圖層 (加上黑色半透明遮罩)
    if (result.introImage) {
      slide.addImage({ data: result.introImage, x: 0, y: 0, w: '100%', h: '100%', sizing: { type: 'cover', w: 10, h: 5.625 } });
      slide.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { type: 'solid', color: '000000', transparency: 60 } });
    } else {
       // 無圖時的備用幾何裝飾
       slide.addShape('rect', { x: 0, y: 0, w: '100%', h: 0.055, fill: { type: 'solid', color: T.accent } });
       slide.addShape('ellipse', { x: 6.5, y: -1.2, w: 5.5, h: 5.5, fill: { type: 'solid', color: T.card } });
       slide.addShape('ellipse', { x: -1.5, y: 4.5, w: 4, h: 4, fill: { type: 'solid', color: T.card } });
       slide.addShape('ellipse', { x: 5.5, y: 4.5, w: 2.5, h: 2.5, fill: { type: 'solid', color: T.accent + '18' } });
    }

    // OmniView badge
    slide.addShape('roundRect', { x: 0.5, y: 0.4, w: 0.6, h: 0.6, rectRadius: 0.1, fill: { type: 'solid', color: T.accent } });
    slide.addText('OV', { x: 0.5, y: 0.4, w: 0.6, h: 0.6, fontSize: 14, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
    slide.addText('OmniView AI', { x: 1.25, y: 0.45, w: 3, h: 0.5, fontSize: 16, bold: true, color: 'F8FAFC' });

    // Tagline chip
    slide.addShape('roundRect', { x: 0.5, y: 1.3, w: 3.2, h: 0.32, rectRadius: 0.16, fill: { type: 'solid', color: T.accent + '28' }, line: { color: T.accent + '60', width: 0.75 } });
    slide.addText(T.tagline, { x: 0.5, y: 1.3, w: 3.2, h: 0.32, fontSize: 8.5, bold: true, color: T.accent, align: 'center', valign: 'middle', charSpacing: 1.5 });

    // 主標題
    slide.addText('商業提案\n分析報告', {
      x: 0.5, y: 1.85, w: 8.5, h: 2.0,
      fontSize: 44, bold: true, color: 'F8FAFC', lineSpacingMultiple: 1.1,
    });

    // 成功機率 badge
    const sc = result.successProbability;
    const sC = scoreColor(sc);
    slide.addShape('roundRect', { x: 0.5, y: 4.15, w: 2.4, h: 0.95, rectRadius: 0.14, fill: { type: 'solid', color: sC + '22' }, line: { color: sC, width: 1.5 } });
    slide.addText('AI 評估成功機率', { x: 0.5, y: 4.2, w: 2.4, h: 0.3, fontSize: 9, color: sC, align: 'center' });
    slide.addText(`${sc}%`, { x: 0.5, y: 4.5, w: 2.4, h: 0.55, fontSize: 28, bold: true, color: sC, align: 'center', valign: 'middle' });

    // 核心摘要（精簡版）
    const summaryLine = trunc(result.executiveSummary, 90);
    slide.addText(summaryLine, {
      x: 0.5, y: 5.4, w: 8.5, h: 0.7,
      fontSize: 13, color: '94A3B8', lineSpacingMultiple: 1.5, italic: true,
    });

    // 底部
    // slide.addShape('rect', { x: 0, y: 7.15, w: '100%', h: 0.35, fill: { type: 'solid', color: T.card } });
    // 如果有圖片 底部改為半透明黑
    slide.addShape('rect', { x: 0, y: 7.15, w: '100%', h: 0.35, fill: { type: 'solid', color: result.introImage ? '000000' : T.card, transparency: result.introImage ? 40 : 0 } });

    slide.addText('由 OmniView AI 360° 虛擬董事會自動生成', { x: 0.5, y: 7.18, w: 6, h: 0.28, fontSize: 9, color: 'CBD5E1' });
    addPageNum(slide, 1, TOTAL, T);
  }

  // ════════════════════════════════════════════════════
  // SLIDE 2 — EXECUTIVE SUMMARY
  // ════════════════════════════════════════════════════
  {
    const slide = pptx.addSlide();
    addBg(slide, T);
    addHeader(slide, T, '執行摘要');

    // 左欄：摘要 bullets
    addCard(slide, T, 0.45, 0.85, 5.6, 5.6);
    slide.addText('核心觀點', { x: 0.7, y: 0.97, w: 5.1, h: 0.32, fontSize: 14, bold: true, color: T.accent });
    const bullets = toBullets(result.executiveSummary, 3, 45);
    const bulletRows = bullets.map(b => [
      // 移除 Emoji，改用純文字符號
      { text: '•  ', options: { color: T.accent, bold: true, fontSize: 18 } },
      { text: b, options: { color: 'CBD5E1', fontSize: 16 } },
    ]);
    bulletRows.forEach((row, i) => {
      slide.addText(row, { x: 0.7, y: 1.6 + i * 1.3, w: 5.1, h: 1.0, lineSpacingMultiple: 1.4, valign: 'top' });
    });

    // 右欄：關鍵指標
    const kpis = [
      { label: '市場規模 (TAM)', val: trunc(result.marketAnalysis.size, 18), color: T.accent },
      { label: '成長率 (CAGR)', val: trunc(result.marketAnalysis.growthRate, 18), color: T.accent2 },
      { label: '損益平衡', val: trunc(result.breakEvenPoint, 18), color: T.accent3 },
    ];
    kpis.forEach((kpi, i) => {
      const y = 0.85 + i * 1.85;
      addCard(slide, T, 6.25, y, 3.15, 1.65, kpi.color + '55');
      slide.addShape('rect', { x: 6.25, y, w: 0.04, h: 1.65, fill: { type: 'solid', color: kpi.color } });
      slide.addText(kpi.label, { x: 6.45, y: y + 0.25, w: 2.8, h: 0.3, fontSize: 12, color: '94A3B8' });
      slide.addText(kpi.val, { x: 6.45, y: y + 0.65, w: 2.8, h: 0.8, fontSize: 20, bold: true, color: kpi.color, lineSpacingMultiple: 1.2, valign: 'top' });
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
    kpis.forEach((k, i) => {
      const x = 0.45 + i * 3.18;
      addCard(slide, T, x, 0.85, 2.95, 1.4, k.color + '66');
      slide.addShape('rect', { x, y: 0.85, w: 2.95, h: 0.04, fill: { type: 'solid', color: k.color } });
      slide.addText(k.label, { x: x + 0.18, y: 1.05, w: 2.6, h: 0.3, fontSize: 12, color: '94A3B8' });
      slide.addText(trunc(k.val, 18), { x: x + 0.18, y: 1.45, w: 2.6, h: 0.7, fontSize: 22, bold: true, color: k.color, lineSpacingMultiple: 1.1, valign: 'top' });
    });

    // 市場描述 bullets
    addCard(slide, T, 0.45, 2.45, 9.05, 4.0);
    slide.addText('市場洞察', { x: 0.72, y: 2.65, w: 8.5, h: 0.32, fontSize: 14, bold: true, color: T.accent });
    const mBullets = toBullets(result.marketAnalysis.description, 3, 55);
    mBullets.forEach((b, i) => {
      slide.addText([
        // 移除 Emoji
        { text: '•  ', options: { color: T.accent2, bold: true, fontSize: 16 } },
        { text: b, options: { color: 'CBD5E1', fontSize: 16 } },
      ], { x: 0.72, y: 3.2 + i * 1.0, w: 8.5, h: 0.8, lineSpacingMultiple: 1.4, valign: 'top' });
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

    // 表格
    const fHeader: PptxGenJS.TableCell[] = ['年度', '營收', '成本', '淨利'].map(t => ({
      text: t,
      options: { bold: true, color: 'FFFFFF', fill: { color: T.headerBg }, fontSize: 12, align: 'center' as const },
    }));

    const fRows: PptxGenJS.TableCell[][] = [fHeader];
    result.financials.slice(0, 5).forEach((f, i) => {
      const bg = i % 2 === 0 ? T.card : T.bg;
      const profit = Number(f.profit) || 0;
      fRows.push([
        { text: trunc(f.year ?? '—', 12), options: { color: 'F8FAFC', fill: { color: bg }, fontSize: 12, align: 'center' as const, bold: true } },
        { text: fmt(Number(f.revenue) || 0), options: { color: T.accent, fill: { color: bg }, fontSize: 12, align: 'center' as const, bold: true } },
        { text: fmt(Number(f.costs) || 0), options: { color: '94A3B8', fill: { color: bg }, fontSize: 12, align: 'center' as const } },
        { text: fmt(profit), options: { color: profit >= 0 ? T.accent2 : 'EF4444', fill: { color: bg }, fontSize: 12, align: 'center' as const, bold: true } },
      ]);
    });

    slide.addTable(fRows, {
      x: 0.45, y: 0.92, w: 9.05,
      border: { type: 'solid', color: T.divider, pt: 0.5 },
      colW: [2.26, 2.26, 2.26, 2.27],
      rowH: [0.45, ...result.financials.slice(0, 5).map(() => 0.52)],
    });

    // 視覺化長條（營收趨勢）
    const maxRev = Math.max(...result.financials.map(f => Number(f.revenue) || 1));
    const barY = 0.92 + 0.45 + result.financials.slice(0, 5).length * 0.52 + 0.35;

    addCard(slide, T, 0.45, barY, 9.05, 7.5 - barY - 0.2);
    slide.addText('營收趨勢', { x: 0.72, y: barY + 0.12, w: 3, h: 0.28, fontSize: 10, bold: true, color: T.accent });
    slide.addText(`損益平衡：${trunc(result.breakEvenPoint, 30)}`, { x: 4, y: barY + 0.12, w: 5.2, h: 0.28, fontSize: 10, color: T.accent2, align: 'right' });

    const barAreaW = 8.4;
    const barH = 0.32;
    result.financials.slice(0, 4).forEach((f, i) => {
      const rev = Number(f.revenue) || 0;
      const ratio = maxRev > 0 ? rev / maxRev : 0;
      const bw = Math.max(ratio * barAreaW, 0.15);
      const by = barY + 0.55 + i * 0.5;
      slide.addText(trunc(f.year ?? '—', 8), { x: 0.72, y: by, w: 0.7, h: barH, fontSize: 9, color: '94A3B8', valign: 'middle' });
      slide.addShape('roundRect', { x: 1.5, y: by + 0.02, w: bw, h: barH - 0.04, rectRadius: 0.05, fill: { type: 'solid', color: T.accent } });
      slide.addText(fmt(rev), { x: 1.6 + bw, y: by, w: 2, h: barH, fontSize: 9, color: T.accent, bold: true, valign: 'middle' });
    });

    addPageNum(slide, 4, TOTAL, T);
  }

  // ════════════════════════════════════════════════════
  // SLIDE 5 — COMPETITIVE LANDSCAPE
  // ════════════════════════════════════════════════════
  {
    const slide = pptx.addSlide();
    addBg(slide, T);
    addHeader(slide, T, '競爭態勢分析');

    const comps = result.competitors.slice(0, 3);
    const cHeader: PptxGenJS.TableCell[] = [
      // 移除 Emoji
      { text: '競爭對手', options: { bold: true, color: 'FFFFFF', fill: { color: T.headerBg }, fontSize: 14, align: 'center' as const } },
      { text: '優勢 (Pros)', options: { bold: true, color: 'FFFFFF', fill: { color: T.headerBg }, fontSize: 14, align: 'center' as const } },
      { text: '劣勢 (Cons)', options: { bold: true, color: 'FFFFFF', fill: { color: T.headerBg }, fontSize: 14, align: 'center' as const } },
    ];
    const cRows: PptxGenJS.TableCell[][] = [cHeader];
    comps.forEach((c, i) => {
      const bg = i % 2 === 0 ? T.card : T.bg;
      cRows.push([
        { text: trunc(c.name, 15), options: { bold: true, color: 'F8FAFC', fill: { color: bg }, fontSize: 14, align: 'center' as const, valign: 'middle' as const } },
        { text: trunc(c.strength, 30), options: { color: T.accent2, fill: { color: bg }, fontSize: 13, valign: 'middle' as const, lineSpacingMultiple: 1.4 } },
        { text: trunc(c.weakness, 30), options: { color: 'F87171', fill: { color: bg }, fontSize: 13, valign: 'middle' as const, lineSpacingMultiple: 1.4 } },
      ]);
    });

    const rowH = Math.min(1.6, (6.4 / (comps.length + 1)));
    slide.addTable(cRows, {
      x: 0.45, y: 0.88, w: 9.05,
      border: { type: 'solid', color: T.divider, pt: 0.5 },
      colW: [2.2, 3.42, 3.43],
      rowH: [0.5, ...comps.map(() => rowH)],
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

    const items = result.roadmap.slice(0, 3);
    const totalH = 6.2;
    const itemH = totalH / items.length;

    items.forEach((item, i) => {
      const y = 0.88 + i * itemH;

      // 時間軸線
      if (i < items.length - 1) {
        slide.addShape('rect', { x: 0.67, y: y + 0.3, w: 0.03, h: itemH, fill: { type: 'solid', color: T.divider } });
      }
      // 節點圓
      slide.addShape('ellipse', { x: 0.54, y: y + 0.14, w: 0.28, h: 0.28, fill: { type: 'solid', color: T.accent } });
      // 序號
      slide.addText(`${i + 1}`, { x: 0.54, y: y + 0.14, w: 0.28, h: 0.28, fontSize: 10, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });

      // Phase + timeframe
      slide.addText(`${trunc(item.phase, 15)}`, { x: 1.0, y, w: 3.5, h: 0.35, fontSize: 15, bold: true, color: T.accent });
      slide.addShape('roundRect', { x: 4.7, y: y + 0.03, w: 1.5, h: 0.28, rectRadius: 0.14, fill: { type: 'solid', color: T.accent + '25' }, line: { color: T.accent + '60', width: 0.5 } });
      slide.addText(trunc(item.timeframe, 12), { x: 4.7, y: y + 0.03, w: 1.5, h: 0.28, fontSize: 10, color: T.accent, align: 'center', valign: 'middle' });

      // 內容卡
      const cardH = itemH - 0.35;
      addCard(slide, T, 1.0, y + 0.35, 8.45, cardH - 0.05);

      const halfW = 4.0;
      // 產品
      slide.addText('產品', { x: 1.18, y: y + 0.45, w: 0.7, h: 0.24, fontSize: 12, bold: true, color: T.accent2 });
      slide.addText(trunc(item.product, 35), { x: 1.18, y: y + 0.75, w: halfW - 0.3, h: cardH - 0.45, fontSize: 13, color: 'CBD5E1', lineSpacingMultiple: 1.4, valign: 'top' });
      // 分隔
      slide.addShape('rect', { x: 5.2, y: y + 0.45, w: 0.02, h: cardH - 0.2, fill: { type: 'solid', color: T.divider } });
      // 技術
      slide.addText('技術', { x: 5.35, y: y + 0.45, w: 0.7, h: 0.24, fontSize: 12, bold: true, color: T.accent3 });
      slide.addText(trunc(item.technology, 35), { x: 5.35, y: y + 0.75, w: 3.8, h: cardH - 0.45, fontSize: 13, color: 'CBD5E1', lineSpacingMultiple: 1.4, valign: 'top' });
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

    const risks = result.risks.slice(0, 3);
    const cardW = 2.85;
    const cardH = 5.2;
    const gap = 0.25;

    risks.forEach((r, i) => {
      const x = 0.45 + i * (cardW + gap);
      const y = 1.0;
      const ic = r.impact === 'High' ? 'EF4444' : r.impact === 'Medium' ? 'FBBF24' : '10B981';
      const label = r.impact === 'High' ? '高' : r.impact === 'Medium' ? '中' : '低';

      addCard(slide, T, x, y, cardW, cardH, ic + '55');
      // 頂部色條
      slide.addShape('rect', { x, y, w: cardW, h: 0.06, fill: { type: 'solid', color: ic } });
      
      // 衝擊標籤
      slide.addShape('roundRect', { x: x + 0.2, y: y + 0.25, w: 0.82, h: 0.35, rectRadius: 0.17, fill: { type: 'solid', color: ic + '30' }, line: { color: ic, width: 0.5 } });
      slide.addText(label, { x: x + 0.2, y: y + 0.25, w: 0.82, h: 0.35, fontSize: 11, bold: true, color: ic, align: 'center', valign: 'middle' });
      
      // 風險標題
      slide.addText(trunc(r.risk, 25), { x: x + 0.2, y: y + 0.8, w: cardW - 0.4, h: 0.8, fontSize: 16, bold: true, color: 'F8FAFC', valign: 'top', lineSpacingMultiple: 1.3 });
      
      // 分隔線
      slide.addShape('rect', { x: x + 0.2, y: y + 1.7, w: cardW - 0.4, h: 0.02, fill: { type: 'solid', color: T.divider } });

      // 因應策略
      slide.addText([
        { text: '因應策略\n', options: { fontSize: 12, color: '64748B', bold: true } },
        { text: trunc(r.mitigation, 45), options: { fontSize: 13, color: 'CBD5E1' } },
      ], { x: x + 0.2, y: y + 1.9, w: cardW - 0.4, h: cardH - 2.1, lineSpacingMultiple: 1.5, valign: 'top' });
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

    const personas = result.personaEvaluations.slice(0, 3);
    const cardW = 2.85;
    const cardH = 5.2;
    const gap = 0.25;

    personas.forEach((p, i) => {
      const x = 0.45 + i * (cardW + gap);
      const y = 1.0;
      const sc = Number(p.score) || 0;
      const sC = scoreColor(sc);

      addCard(slide, T, x, y, cardW, cardH);
      slide.addShape('rect', { x, y, w: cardW, h: 0.06, fill: { type: 'solid', color: sC } });

      // 角色名 + 分數
      slide.addText(trunc(p.role, 15), { x: x + 0.2, y: y + 0.2, w: cardW - 1.2, h: 0.4, fontSize: 15, bold: true, color: 'F8FAFC' });
      slide.addText(`${sc}`, { x: x + cardW - 1.0, y: y + 0.2, w: 0.8, h: 0.4, fontSize: 24, bold: true, color: sC, align: 'right' });

      // 分數條
      const barW = (cardW - 0.4) * (sc / 100);
      slide.addShape('roundRect', { x: x + 0.2, y: y + 0.7, w: cardW - 0.4, h: 0.08, rectRadius: 0.04, fill: { type: 'solid', color: T.divider } });
      slide.addShape('roundRect', { x: x + 0.2, y: y + 0.7, w: Math.max(barW, 0.05), h: 0.08, rectRadius: 0.04, fill: { type: 'solid', color: sC } });

      // 引言
      slide.addText(`"${trunc(p.keyQuote, 35)}"`, {
        x: x + 0.2, y: y + 1.0, w: cardW - 0.4, h: 1.5,
        fontSize: 13, italic: true, color: '94A3B8', lineSpacingMultiple: 1.5, valign: 'top',
      });

      // 分隔線
      slide.addShape('rect', { x: x + 0.2, y: y + 2.6, w: cardW - 0.4, h: 0.02, fill: { type: 'solid', color: T.divider } });

      // 擔憂
      slide.addText([
        // 移除 Emoji
        { text: '核心擔憂\n', options: { fontSize: 12, color: 'FBBF24', bold: true } },
        { text: trunc(p.concern, 40), options: { fontSize: 13, color: 'CBD5E1' } },
      ], { x: x + 0.2, y: y + 2.8, w: cardW - 0.4, h: cardH - 3.0, lineSpacingMultiple: 1.5, valign: 'top' });
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
      // 移除 Emoji
      { title: '激進觀點 (Aggressive)', text: result.finalVerdicts.aggressive, color: 'F97316' },
      { title: '平衡觀點 (Balanced)', text: result.finalVerdicts.balanced, color: T.accent },
      { title: '保守觀點 (Conservative)', text: result.finalVerdicts.conservative, color: T.accent3 },
    ];
    const vCardW = 2.85;
    const vCardH = 5.2;
    const gap = 0.25;

    verdicts.forEach((v, i) => {
      const x = 0.45 + i * (vCardW + gap);
      const y = 1.0;
      addCard(slide, T, x, y, vCardW, vCardH, v.color + '55');
      // 頂色條
      slide.addShape('rect', { x, y, w: vCardW, h: 0.06, fill: { type: 'solid', color: v.color } });
      // Header
      slide.addText(v.title, { x: x + 0.2, y: y + 0.2, w: vCardW - 0.4, h: 0.5, fontSize: 16, bold: true, color: v.color });
      slide.addShape('rect', { x: x + 0.2, y: y + 0.8, w: 0.8, h: 0.03, fill: { type: 'solid', color: v.color } });
      // Content bullets
      const vBullets = toBullets(v.text, 3, 35);
      vBullets.forEach((b, bi) => {
        slide.addText([
          { text: '• ', options: { color: v.color, bold: true, fontSize: 16 } },
          { text: b, options: { color: 'CBD5E1', fontSize: 14 } },
        ], { x: x + 0.2, y: y + 1.1 + bi * 1.2, w: vCardW - 0.4, h: 1.1, lineSpacingMultiple: 1.5, valign: 'top' });
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

    // 裝飾
    slide.addShape('rect', { x: 0, y: 0, w: '100%', h: 0.055, fill: { type: 'solid', color: T.accent } });
    slide.addShape('ellipse', { x: 2.5, y: 0.8, w: 5, h: 5, fill: { type: 'solid', color: T.card } });
    slide.addShape('ellipse', { x: 3.5, y: 1.5, w: 3, h: 3, fill: { type: 'solid', color: T.accent + '10' } });

    // 主文案
    slide.addText('下一步行動', {
      x: 0.5, y: 1.7, w: 9.5, h: 1.0,
      fontSize: 42, bold: true, color: 'F8FAFC', align: 'center',
    });
    slide.addText(T.tagline, {
      x: 0.5, y: 2.75, w: 9.5, h: 0.42,
      fontSize: 11, color: T.accent, align: 'center', charSpacing: 2.5, bold: true,
    });

    // 成功機率卡
    const sc = result.successProbability;
    const sC = scoreColor(sc);
    addCard(slide, T, 3.6, 3.4, 2.8, 1.55, sC);
    slide.addText('AI 評估成功機率', { x: 3.6, y: 3.52, w: 2.8, h: 0.3, fontSize: 10, color: '94A3B8', align: 'center' });
    slide.addText(`${sc}%`, { x: 3.6, y: 3.85, w: 2.8, h: 0.85, fontSize: 36, bold: true, color: sC, align: 'center' });

    // 底部
    slide.addText('此報告由 OmniView AI 360° 虛擬董事會自動生成', {
      x: 1, y: 6.55, w: 8, h: 0.35, fontSize: 10, color: '475569', align: 'center',
    });
    addPageNum(slide, 10, TOTAL, T);
  }

  // ════════════════════════════════════════════════════
  // EXPORT
  // ════════════════════════════════════════════════════
  await pptx.writeFile({ fileName: 'OmniView_商業提案報告.pptx' });
};

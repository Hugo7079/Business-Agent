import PptxGenJS from 'pptxgenjs';
import { AnalysisResult } from '../types';

// ════════════════════════════════════════════════════
// 版面常數  16:9 = 10" × 7.5"
// ════════════════════════════════════════════════════
const W      = 10;      // 投影片寬
const H      = 7.5;     // 投影片高
const MX     = 0.45;    // 左右 margin
const CW     = W - MX * 2;  // 可用寬 = 9.1"
const BODY_T = 0.82;    // header 結束後內容起始 Y
const BODY_B = 6.75;    // 內容底線（頁碼上方留 0.55"）
const BODY_H = BODY_B - BODY_T;  // = 5.93"

// ════════════════════════════════════════════════════
// 文字工具
// ════════════════════════════════════════════════════

/** 硬截：超過 max 字就截斷加省略號，同時去句末標點 */
const cut = (s: string, max: number): string => {
  if (!s) return '—';
  const t = s.trim().replace(/\s+/g, ' ').replace(/[。！？!?.…]+$/g, '').trim();
  return t.length <= max ? t : t.slice(0, max - 1) + '…';
};

/**
 * 把長文切成 N 條精簡短標題
 * 邏輯：先按句號分句，每句再按逗號取「最前面有意義的子句」，強制截到 maxLen
 */
const toLines = (text: string, maxLines: number, maxLen: number): string[] => {
  if (!text) return ['—'];
  // 先按句號/分號斷句
  const sentences = text
    .replace(/([。！？!?；;])/g, '$1|||')
    .split('|||')
    .map(s => s.replace(/[。！？!?；;]+$/g, '').trim())
    .filter(s => s.length > 1);

  const out: string[] = [];
  for (const sen of sentences) {
    if (out.length >= maxLines) break;
    if (sen.length <= maxLen) {
      out.push(sen);
    } else {
      // 按逗號拆，逐段累加直到超過 maxLen
      const parts = sen.split(/[，,、]/).map(p => p.trim()).filter(p => p.length > 0);
      let line = '';
      for (const p of parts) {
        const candidate = line ? line + '、' + p : p;
        if (candidate.length > maxLen) break;
        line = candidate;
      }
      out.push(cut(line || parts[0] || sen, maxLen));
    }
  }
  return out.length ? out : [cut(text, maxLen)];
};

/** 單行標籤：從長段落取出最核心的短標籤 */
const tag = (text: string, maxLen: number): string => {
  if (!text) return '—';
  const t = text.trim().replace(/[。！？!?.]+$/g, '').trim();
  if (t.length <= maxLen) return t;
  // 先按句號取第一句
  const firstSen = t.split(/[。！？!?；;]/)[0].trim();
  if (firstSen.length <= maxLen) return firstSen;
  // 再按逗號取第一段
  const firstClause = firstSen.split(/[，,、]/)[0].trim();
  return cut(firstClause || firstSen, maxLen);
};

const fmt = (n: number) => {
  const abs = Math.abs(n);
  if (abs >= 1_000_000) return `$${(n / 1_000_000).toFixed(1)}M`;
  if (abs >= 1_000)     return `$${(n / 1_000).toFixed(0)}K`;
  return `$${n}`;
};

// ════════════════════════════════════════════════════
// 產業主題
// ════════════════════════════════════════════════════
interface Theme {
  bg: string; card: string; accent: string; accent2: string;
  accent3: string; headerBg: string; divider: string; tagline: string;
}
const THEMES: Record<string, Theme> = {
  tech:    { bg:'0A0F1E', card:'111827', accent:'3B82F6', accent2:'06B6D4', accent3:'8B5CF6', headerBg:'1E3A5F', divider:'1E3A5F', tagline:'TECHNOLOGY × INNOVATION' },
  health:  { bg:'071A14', card:'0D2418', accent:'10B981', accent2:'34D399', accent3:'06B6D4', headerBg:'0D3320', divider:'134D2E', tagline:'HEALTH × WELLNESS' },
  food:    { bg:'1A0C00', card:'2D1500', accent:'F97316', accent2:'FBBF24', accent3:'EF4444', headerBg:'3D1F00', divider:'4D2800', tagline:'FOOD × LIFESTYLE' },
  finance: { bg:'090E1A', card:'0F1829', accent:'F59E0B', accent2:'10B981', accent3:'3B82F6', headerBg:'1A2540', divider:'1E2D4D', tagline:'FINANCE × GROWTH' },
  retail:  { bg:'0F0A1A', card:'1A1029', accent:'A855F7', accent2:'EC4899', accent3:'F97316', headerBg:'2D1040', divider:'3D1555', tagline:'RETAIL × EXPERIENCE' },
  edu:     { bg:'0A1020', card:'0F1A2E', accent:'6366F1', accent2:'8B5CF6', accent3:'06B6D4', headerBg:'1A2040', divider:'1E2D55', tagline:'EDUCATION × FUTURE' },
  default: { bg:'0F172A', card:'1E293B', accent:'3B82F6', accent2:'10B981', accent3:'7C3AED', headerBg:'1E3A5F', divider:'334155', tagline:'BUSINESS × STRATEGY' },
};
const detectTheme = (a: string, b: string): Theme => {
  const t = (a + b).toLowerCase();
  if (/醫|健|藥|療|病|診|health|medical|biotech/.test(t))          return THEMES.health;
  if (/餐|食|飲|咖|料理|外送|food|restaurant|beverage/.test(t))   return THEMES.food;
  if (/金融|投資|貸|保險|fintech|finance|bank|crypto/.test(t))     return THEMES.finance;
  if (/零售|電商|購物|fashion|retail|ecommerce/.test(t))           return THEMES.retail;
  if (/教育|學習|課程|edu|learning|school|training/.test(t))       return THEMES.edu;
  if (/科技|ai|app|平台|軟體|saas|tech|software|雲端/.test(t))    return THEMES.tech;
  return THEMES.default;
};

// ════════════════════════════════════════════════════
// 共用元件
// ════════════════════════════════════════════════════
const addBg = (slide: PptxGenJS.Slide, T: Theme) => {
  slide.background = { fill: T.bg };
  slide.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ type:'solid', color:T.accent } });
};

const addHeader = (slide: PptxGenJS.Slide, T: Theme, title: string) => {
  slide.addText(title, { x:MX, y:0.1, w:8, h:0.48, fontSize:19, bold:true, color:'F8FAFC', fontFace:'Arial' });
  slide.addShape('rect', { x:MX, y:0.6, w:0.85, h:0.03, fill:{ type:'solid', color:T.accent } });
};

// 頁碼永遠在安全區內
const addPageNum = (slide: PptxGenJS.Slide, n: number, total: number, T: Theme) =>
  slide.addText(`${n} / ${total}`, { x:8.6, y:6.95, w:0.8, h:0.25, fontSize:7.5, color:T.divider, align:'right' });

const addCard = (slide: PptxGenJS.Slide, T: Theme, x:number, y:number, w:number, h:number, border?: string) =>
  slide.addShape('roundRect', { x, y, w, h, rectRadius:0.1,
    fill:{ type:'solid', color:T.card }, line:{ color: border || T.divider, width:0.75 } });

const scoreColor = (s: number) =>
  s >= 80 ? '10B981' : s >= 60 ? '3B82F6' : s >= 40 ? 'FBBF24' : 'EF4444';

// ════════════════════════════════════════════════════
// 主函式
// ════════════════════════════════════════════════════
export const generatePptx = async (result: AnalysisResult): Promise<void> => {
  const T = detectTheme(result.executiveSummary, result.marketAnalysis.description);
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';
  pptx.author = 'OmniView AI';
  pptx.title  = 'OmniView 商業提案分析報告';
  const TOTAL = 10;

  // ──────────────────────────────────────────────────
  // SLIDE 1 — COVER
  // 版面：H=7.5"，所有元素嚴格在 0 ~ 7.3" 以內
  // ──────────────────────────────────────────────────
  {
    const sl = pptx.addSlide();
    sl.background = { fill: T.bg };
    sl.addShape('rect',    { x:0,   y:0,    w:'100%', h:0.05, fill:{ type:'solid', color:T.accent } });
    sl.addShape('ellipse', { x:6.5, y:-0.8, w:4.0,   h:4.0,  fill:{ type:'solid', color:T.card } });
    sl.addShape('ellipse', { x:-1,  y:5.0,  w:3.0,   h:3.0,  fill:{ type:'solid', color:T.card } });

    // badge  y:0.3 ~ 0.9
    sl.addShape('roundRect', { x:0.5, y:0.3, w:0.55, h:0.55, rectRadius:0.08, fill:{ type:'solid', color:T.accent } });
    sl.addText('OV', { x:0.5, y:0.3, w:0.55, h:0.55, fontSize:13, bold:true, color:'FFFFFF', align:'center', valign:'middle' });
    sl.addText('OmniView AI', { x:1.15, y:0.35, w:3, h:0.45, fontSize:15, bold:true, color:'F8FAFC' });

    // tagline chip  y:1.05 ~ 1.37
    sl.addShape('roundRect', { x:0.5, y:1.05, w:3.0, h:0.32, rectRadius:0.16,
      fill:{ type:'solid', color:T.accent+'28' }, line:{ color:T.accent+'60', width:0.75 } });
    sl.addText(T.tagline, { x:0.5, y:1.05, w:3.0, h:0.32, fontSize:8, bold:true, color:T.accent, align:'center', valign:'middle', charSpacing:1.5 });

    // 主標題  y:1.55 ~ 3.25  (h=1.7, fontSize=40, 2行)
    sl.addText('商業提案\n分析報告', { x:0.5, y:1.55, w:8.5, h:1.7, fontSize:40, bold:true, color:'F8FAFC', lineSpacingMultiple:1.1 });

    // 成功機率  y:3.5 ~ 4.35
    const sc = result.successProbability;
    const sC = scoreColor(sc);
    sl.addShape('roundRect', { x:0.5, y:3.5, w:2.3, h:0.85, rectRadius:0.12,
      fill:{ type:'solid', color:sC+'22' }, line:{ color:sC, width:1.5 } });
    sl.addText('AI 評估成功機率', { x:0.5, y:3.56, w:2.3, h:0.26, fontSize:8.5, color:sC, align:'center' });
    sl.addText(`${sc}%`, { x:0.5, y:3.82, w:2.3, h:0.48, fontSize:26, bold:true, color:sC, align:'center', valign:'middle' });

    // 摘要一行  y:4.65 ~ 5.15  (h=0.5, fontSize=11, 1行)
    sl.addText(tag(result.executiveSummary, 55), { x:0.5, y:4.65, w:9.0, h:0.5, fontSize:11, color:'94A3B8', italic:true });

    // 底部欄  y:6.9 ~ 7.3
    sl.addShape('rect', { x:0, y:6.9, w:'100%', h:0.38, fill:{ type:'solid', color:T.card } });
    sl.addText('由 OmniView AI 360° 虛擬董事會自動生成', { x:0.5, y:6.93, w:6, h:0.28, fontSize:8.5, color:'475569' });
    addPageNum(sl, 1, TOTAL, T);
  }

  // ──────────────────────────────────────────────────
  // SLIDE 2 — EXECUTIVE SUMMARY
  // 左卡 x:0.45 w:5.35  右側3張KPI卡 x:5.95 w:3.65
  // 全部 y:BODY_T ~ BODY_B  (0.82 ~ 6.75  h=5.93)
  // ──────────────────────────────────────────────────
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '執行摘要');

    // 左卡
    const lW = 5.35, lH = BODY_H;   // 5.93"
    addCard(sl, T, MX, BODY_T, lW, lH);
    sl.addText('核心觀點', { x:0.65, y:BODY_T+0.12, w:lW-0.4, h:0.28, fontSize:10.5, bold:true, color:T.accent });

    // bullets: 最多 5 條，每條 h=0.82"，從 BODY_T+0.5 開始
    // 最大 Y = BODY_T+0.5 + 4*0.82 + 0.82 = BODY_T+4.1 = 4.92 < BODY_B ✓
    const bLines = toLines(result.executiveSummary, 5, 26);
    const bRowH = Math.min(0.82, (lH - 0.5) / bLines.length);
    bLines.forEach((b, i) => {
      const by = BODY_T + 0.5 + i * bRowH;
      sl.addText([
        { text:'▸  ', options:{ color:T.accent, bold:true, fontSize:11.5 } },
        { text:b,    options:{ color:'CBD5E1',  fontSize:11.5 } },
      ], { x:0.62, y:by, w:lW-0.3, h:bRowH, valign:'middle' });
    });

    // 右側 KPI × 3，均分 BODY_H，gap=0.12
    const kpis = [
      { label:'市場規模 (TAM)', val: tag(result.marketAnalysis.size,      18), color:T.accent  },
      { label:'成長率 (CAGR)',  val: tag(result.marketAnalysis.growthRate, 18), color:T.accent2 },
      { label:'損益平衡點',     val: tag(result.breakEvenPoint,            18), color:T.accent3 },
    ];
    const kGap = 0.12;
    const kH   = (BODY_H - kGap * 2) / 3;   // ≈ 1.90"
    const kX   = MX + lW + 0.1;
    const kW   = W - MX - kX;               // = 10 - 0.45 - 5.9 = 3.65"
    kpis.forEach((k, i) => {
      const ky = BODY_T + i * (kH + kGap);
      addCard(sl, T, kX, ky, kW, kH, k.color+'55');
      sl.addShape('rect', { x:kX, y:ky, w:0.04, h:kH, fill:{ type:'solid', color:k.color } });
      sl.addText(k.label, { x:kX+0.14, y:ky+0.14, w:kW-0.2, h:0.26, fontSize:9.5, color:'94A3B8' });
      // 值：fontSize=15，h 留 kH-0.55 確保不超出卡片
      sl.addText(k.val, { x:kX+0.14, y:ky+0.45, w:kW-0.2, h:kH-0.55, fontSize:15, bold:true, color:k.color, lineSpacingMultiple:1.15 });
    });

    addPageNum(sl, 2, TOTAL, T);
  }

  // ──────────────────────────────────────────────────
  // SLIDE 3 — MARKET OPPORTUNITY
  // KPI 3張橫排 h=1.25"，市場卡佔剩餘高度
  // ──────────────────────────────────────────────────
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '市場機會');

    const KPI_H = 1.2;
    const kpis = [
      { label:'TAM 市場規模', val: tag(result.marketAnalysis.size,      16), color:T.accent  },
      { label:'CAGR 年均成長', val: tag(result.marketAnalysis.growthRate,16), color:T.accent2 },
      { label:'損益平衡點',    val: tag(result.breakEvenPoint,           16), color:T.accent3 },
    ];
    const kW = (CW - 0.2) / 3;   // ≈ 2.97"
    kpis.forEach((k, i) => {
      const kx = MX + i * (kW + 0.1);
      addCard(sl, T, kx, BODY_T, kW, KPI_H, k.color+'66');
      sl.addShape('rect', { x:kx, y:BODY_T, w:kW, h:0.04, fill:{ type:'solid', color:k.color } });
      sl.addText(k.label, { x:kx+0.12, y:BODY_T+0.1, w:kW-0.24, h:0.26, fontSize:9.5, color:'94A3B8' });
      sl.addText(k.val,   { x:kx+0.12, y:BODY_T+0.42, w:kW-0.24, h:0.62, fontSize:16, bold:true, color:k.color });
    });

    // 市場卡：y = BODY_T+KPI_H+0.15，底 = BODY_B
    const mTop = BODY_T + KPI_H + 0.15;   // = 0.82+1.2+0.15 = 2.17
    const mH   = BODY_B - mTop;            // = 6.75-2.17 = 4.58"
    addCard(sl, T, MX, mTop, CW, mH);
    sl.addText('市場洞察', { x:MX+0.2, y:mTop+0.12, w:CW-0.4, h:0.28, fontSize:10.5, bold:true, color:T.accent });

    // bullets：最多 5 條，每條高度均分可用空間
    // 可用 = mH - 0.5（標題）= 4.08"，每條 ≤ 0.72"
    const mLines = toLines(result.marketAnalysis.description, 5, 38);
    const mRowH  = Math.min(0.72, (mH - 0.5) / mLines.length);
    mLines.forEach((b, i) => {
      const by = mTop + 0.5 + i * mRowH;
      sl.addText([
        { text:'◆  ', options:{ color:T.accent2, bold:true, fontSize:11 } },
        { text:b,    options:{ color:'CBD5E1',   fontSize:11 } },
      ], { x:MX+0.2, y:by, w:CW-0.4, h:mRowH, valign:'middle' });
    });

    addPageNum(sl, 3, TOTAL, T);
  }

  // ──────────────────────────────────────────────────
  // SLIDE 4 — FINANCIAL PROJECTIONS
  // 表格 + 長條圖，全部嚴格在 BODY_T ~ BODY_B
  // ──────────────────────────────────────────────────
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '財務預測');

    const finData = result.financials.slice(0, 4);
    const tHdrH  = 0.42;
    const tRowH  = 0.46;
    const tH     = tHdrH + finData.length * tRowH;  // max: 0.42+4*0.46=2.26"

    const fHdr: PptxGenJS.TableCell[] = ['年度','營收','成本','淨利'].map(t => ({
      text: t,
      options: { bold:true, color:'FFFFFF', fill:{ color:T.headerBg }, fontSize:11.5, align:'center' as const },
    }));
    const fRows: PptxGenJS.TableCell[][] = [fHdr];
    finData.forEach((f, i) => {
      const bg  = i%2===0 ? T.card : T.bg;
      const pft = Number(f.profit)||0;
      fRows.push([
        { text: cut(f.year??'—',10),           options:{ color:'F8FAFC',                      fill:{color:bg}, fontSize:11.5, align:'center' as const, bold:true } },
        { text: fmt(Number(f.revenue)||0),      options:{ color:T.accent,                      fill:{color:bg}, fontSize:11.5, align:'center' as const, bold:true } },
        { text: fmt(Number(f.costs)||0),        options:{ color:'94A3B8',                      fill:{color:bg}, fontSize:11.5, align:'center' as const } },
        { text: fmt(pft),                       options:{ color: pft>=0 ? T.accent2 : 'EF4444', fill:{color:bg}, fontSize:11.5, align:'center' as const, bold:true } },
      ]);
    });

    sl.addTable(fRows, {
      x:MX, y:BODY_T, w:CW,
      border:{ type:'solid', color:T.divider, pt:0.5 },
      colW:[2.275, 2.275, 2.275, 2.275],
      rowH:[tHdrH, ...finData.map(()=>tRowH)],
    });

    // 長條圖卡：y = BODY_T+tH+0.2，底 = BODY_B
    const barTop = BODY_T + tH + 0.2;
    const barH   = BODY_B - barTop;   // 動態高度，一定不超出

    if (barH >= 0.9) {
      addCard(sl, T, MX, barTop, CW, barH);
      sl.addText('營收趨勢', { x:0.65, y:barTop+0.1, w:3, h:0.24, fontSize:9.5, bold:true, color:T.accent });
      sl.addText(`損益平衡：${tag(result.breakEvenPoint, 22)}`, { x:4.5, y:barTop+0.1, w:4.75, h:0.24, fontSize:9.5, color:T.accent2, align:'right' });

      const maxRev  = Math.max(...finData.map(f=>Number(f.revenue)||0), 1);
      const barSX   = 1.5;
      const barMaxW = 6.8;  // barSX + barMaxW + label(0.9) = 1.5+6.8+0.9 = 9.2 < 9.55 ✓
      const rowH2   = Math.min(0.38, (barH - 0.42) / finData.length);
      finData.forEach((f, i) => {
        const rev = Number(f.revenue)||0;
        const bw  = Math.max(rev / maxRev * barMaxW, 0.12);
        const by  = barTop + 0.38 + i * rowH2;
        if (by + rowH2 > BODY_B - 0.05) return;
        sl.addText(cut(f.year??'—', 6), { x:0.55, y:by, w:0.85, h:rowH2, fontSize:8.5, color:'94A3B8', valign:'middle' });
        sl.addShape('roundRect', { x:barSX, y:by+0.04, w:bw, h:rowH2-0.08, rectRadius:0.04, fill:{ type:'solid', color:T.accent } });
        // label 貼在 bar 右側，若快碰右邊就改成白色放 bar 內
        const lx = barSX + bw + 0.08;
        if (lx + 0.85 <= MX + CW) {
          sl.addText(fmt(rev), { x:lx, y:by, w:0.85, h:rowH2, fontSize:8.5, color:T.accent, bold:true, valign:'middle' });
        } else {
          sl.addText(fmt(rev), { x:barSX+bw-0.95, y:by, w:0.85, h:rowH2, fontSize:8.5, color:'FFFFFF', bold:true, valign:'middle' });
        }
      });
    }
    addPageNum(sl, 4, TOTAL, T);
  }

  // ──────────────────────────────────────────────────
  // SLIDE 5 — COMPETITIVE LANDSCAPE
  // 表格，行高動態計算確保不超出 BODY_B
  // ──────────────────────────────────────────────────
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '競爭態勢分析');

    const comps   = result.competitors.slice(0, 5);
    const hdrH    = 0.40;
    // 動態行高：確保 hdrH + n*rowH ≤ BODY_H
    const rowH    = Math.min(0.68, (BODY_H - hdrH) / comps.length);

    const cHdr: PptxGenJS.TableCell[] = ['競爭對手','優勢','劣勢'].map(t => ({
      text: t,
      options: { bold:true, color:'FFFFFF', fill:{ color:T.headerBg }, fontSize:11.5, align:'center' as const },
    }));
    const cRows: PptxGenJS.TableCell[][] = [cHdr];
    comps.forEach((c, i) => {
      const bg = i%2===0 ? T.card : T.bg;
      // 每格字數硬截：按字型 10pt、行寬 3.2" 約可放 22 字/行，rowH=0.68" 約 1.5 行 → 取 22 字
      cRows.push([
        { text: tag(c.name,     14), options:{ bold:true, color:'F8FAFC',  fill:{color:bg}, fontSize:11, align:'center' as const } },
        { text: tag(c.strength, 22), options:{ color:T.accent2,             fill:{color:bg}, fontSize:10 } },
        { text: tag(c.weakness, 22), options:{ color:'F87171',              fill:{color:bg}, fontSize:10 } },
      ]);
    });

    sl.addTable(cRows, {
      x:MX, y:BODY_T, w:CW,
      border:{ type:'solid', color:T.divider, pt:0.5 },
      colW:[2.0, 3.55, 3.55],
      rowH:[hdrH, ...comps.map(()=>rowH)],
    });
    addPageNum(sl, 5, TOTAL, T);
  }

  // ──────────────────────────────────────────────────
  // SLIDE 6 — STRATEGIC ROADMAP
  // N 個 item 均分 BODY_H，每個 item 的所有元素都在自己的區塊內
  // ──────────────────────────────────────────────────
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '策略路線圖');

    const items  = result.roadmap.slice(0, 4);
    const iH     = BODY_H / items.length;  // 每 item 高度，4個時 = 1.48"

    items.forEach((item, i) => {
      const iY  = BODY_T + i * iH;
      const iB  = iY + iH;   // 此 item 底線

      // 時間軸線（只到下一個 item 頂，不超過 iB）
      if (i < items.length - 1)
        sl.addShape('rect', { x:0.61, y:iY+0.28, w:0.025, h:iH-0.28, fill:{ type:'solid', color:T.divider } });

      // 節點
      sl.addShape('ellipse', { x:0.49, y:iY+0.1, w:0.25, h:0.25, fill:{ type:'solid', color:T.accent } });
      sl.addText(`${i+1}`, { x:0.49, y:iY+0.1, w:0.25, h:0.25, fontSize:8.5, bold:true, color:'FFFFFF', align:'center', valign:'middle' });

      // Phase 標題  h=0.3"
      sl.addText(tag(item.phase, 16), { x:0.88, y:iY+0.05, w:3.3, h:0.3, fontSize:12.5, bold:true, color:T.accent });

      // Timeframe chip  右側
      sl.addShape('roundRect', { x:4.35, y:iY+0.07, w:1.45, h:0.25, rectRadius:0.12,
        fill:{ type:'solid', color:T.accent+'25' }, line:{ color:T.accent+'60', width:0.5 } });
      sl.addText(tag(item.timeframe, 12), { x:4.35, y:iY+0.07, w:1.45, h:0.25, fontSize:8.5, color:T.accent, align:'center', valign:'middle' });

      // 內容卡：y=iY+0.38，底=iB-0.05
      const cTop = iY + 0.38;
      const cH   = iB - cTop - 0.05;   // 嚴格不超過 item 底線
      if (cH < 0.35) return;
      addCard(sl, T, 0.88, cTop, CW - 0.45, cH);

      // 產品 / 技術 各佔一半，文字高 = cH - 0.32
      const txtH = cH - 0.32;
      const half = (CW - 0.45 - 0.1) / 2;  // 各欄寬
      // 左：產品
      sl.addText('產品', { x:1.06, y:cTop+0.06, w:0.6, h:0.2, fontSize:8.5, bold:true, color:T.accent2 });
      sl.addText(tag(item.product, 28), { x:1.06, y:cTop+0.27, w:half-0.2, h:txtH, fontSize:9.5, color:'CBD5E1', valign:'top' });
      // 分隔線
      sl.addShape('rect', { x: 0.88 + half + 0.05, y:cTop+0.06, w:0.02, h:cH-0.12, fill:{ type:'solid', color:T.divider } });
      // 右：技術
      const rx = 0.88 + half + 0.12;
      sl.addText('技術', { x:rx, y:cTop+0.06, w:0.6, h:0.2, fontSize:8.5, bold:true, color:T.accent3 });
      sl.addText(tag(item.technology, 28), { x:rx, y:cTop+0.27, w:half-0.2, h:txtH, fontSize:9.5, color:'CBD5E1', valign:'top' });
    });
    addPageNum(sl, 6, TOTAL, T);
  }

  // ──────────────────────────────────────────────────
  // SLIDE 7 — RISK ASSESSMENT
  // 2欄 × ceil(N/2) 列，每張卡嚴格均分 BODY_H
  // ──────────────────────────────────────────────────
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '風險評估');

    const risks = result.risks.slice(0, 6);
    const COLS  = 2;
    const nRows = Math.ceil(risks.length / COLS);
    const gap   = 0.1;
    const cW    = (CW - gap) / COLS;            // 每卡寬 ≈ 4.5"
    const cH    = (BODY_H - gap*(nRows-1)) / nRows;  // 每卡高

    risks.forEach((r, i) => {
      const col = i % COLS;
      const row = Math.floor(i / COLS);
      const cx  = MX + col * (cW + gap);
      const cy  = BODY_T + row * (cH + gap);
      // 最終底線驗證
      if (cy + cH > BODY_B + 0.02) return;

      const ic    = r.impact==='High' ? 'EF4444' : r.impact==='Medium' ? 'FBBF24' : '10B981';
      const label = r.impact==='High' ? '高影響' : r.impact==='Medium' ? '中影響' : '低影響';

      addCard(sl, T, cx, cy, cW, cH, ic+'55');
      sl.addShape('rect', { x:cx, y:cy, w:cW, h:0.04, fill:{ type:'solid', color:ic } });

      // 衝擊標籤  右上
      sl.addShape('roundRect', { x:cx+cW-0.95, y:cy+0.1, w:0.82, h:0.24, rectRadius:0.12,
        fill:{ type:'solid', color:ic+'30' }, line:{ color:ic, width:0.5 } });
      sl.addText(label, { x:cx+cW-0.95, y:cy+0.1, w:0.82, h:0.24, fontSize:8.5, bold:true, color:ic, align:'center', valign:'middle' });

      // 風險標題  h=0.28"，字數截到 18 字（11pt × 約18字 ≤ cW-1.1"）
      sl.addText(tag(r.risk, 18), { x:cx+0.16, y:cy+0.1, w:cW-1.15, h:0.28, fontSize:11, bold:true, color:'F8FAFC' });

      // 因應策略  剩餘高度 = cH - 0.48
      const mitH = cH - 0.48;
      if (mitH > 0.2) {
        sl.addText([
          { text:'因應：', options:{ fontSize:8.5, color:'64748B', bold:true } },
          { text: tag(r.mitigation, 30), options:{ fontSize:8.5, color:'CBD5E1' } },
        ], { x:cx+0.16, y:cy+0.44, w:cW-0.28, h:mitH, valign:'top' });
      }
    });
    addPageNum(sl, 7, TOTAL, T);
  }

  // ──────────────────────────────────────────────────
  // SLIDE 8 — STAKEHOLDER BOARD (最多5人：3+2配置)
  // ──────────────────────────────────────────────────
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, 'AI 虛擬董事會');

    const personas = result.personaEvaluations.slice(0, 5);
    const COLS  = Math.min(personas.length, 3);
    const nRows = Math.ceil(personas.length / COLS);
    const gap   = 0.1;
    const cW    = (CW - gap*(COLS-1)) / COLS;
    const cH    = (BODY_H - gap*(nRows-1)) / nRows;

    personas.forEach((p, i) => {
      const col = i % COLS;
      const row = Math.floor(i / COLS);
      const cx  = MX + col * (cW + gap);
      const cy  = BODY_T + row * (cH + gap);
      if (cy + cH > BODY_B + 0.02) return;

      const sc = Number(p.score)||0;
      const sC = scoreColor(sc);

      addCard(sl, T, cx, cy, cW, cH);
      sl.addShape('rect', { x:cx, y:cy, w:cW, h:0.04, fill:{ type:'solid', color:sC } });

      // 角色名  h=0.28"
      sl.addText(tag(p.role, 10), { x:cx+0.12, y:cy+0.1, w:cW-0.9, h:0.28, fontSize:11.5, bold:true, color:'F8FAFC' });
      // 分數
      sl.addText(`${sc}`, { x:cx+cW-0.88, y:cy+0.08, w:0.75, h:0.3, fontSize:17, bold:true, color:sC, align:'right' });

      // 分數條  y=cy+0.43
      const barFill = (cW-0.24)*(sc/100);
      sl.addShape('roundRect', { x:cx+0.12, y:cy+0.43, w:cW-0.24, h:0.055, rectRadius:0.028, fill:{ type:'solid', color:T.divider } });
      sl.addShape('roundRect', { x:cx+0.12, y:cy+0.43, w:Math.max(barFill,0.05), h:0.055, rectRadius:0.028, fill:{ type:'solid', color:sC } });

      // 引言  y=cy+0.56，h=0.5"（字 9pt，最多 2 行，截到 20 字）
      sl.addText(`"${tag(p.keyQuote, 20)}"`, { x:cx+0.12, y:cy+0.56, w:cW-0.24, h:0.5, fontSize:9, italic:true, color:'94A3B8', valign:'top' });

      // 擔憂  y=cy+1.1，h=剩餘
      const conY = cy + 1.1;
      const conH = (cy + cH) - conY - 0.08;
      if (conH > 0.18) {
        sl.addText([
          { text:'! ', options:{ fontSize:8.5, color:'FBBF24', bold:true } },
          { text: tag(p.concern, 24), options:{ fontSize:8.5, color:'CBD5E1' } },
        ], { x:cx+0.12, y:conY, w:cW-0.24, h:conH, valign:'top' });
      }
    });
    addPageNum(sl, 8, TOTAL, T);
  }

  // ──────────────────────────────────────────────────
  // SLIDE 9 — FINAL VERDICTS
  // 3欄等寬，每欄 bullet 高度動態計算
  // ──────────────────────────────────────────────────
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '最終裁決');

    const verdicts = [
      { title:'激進觀點', text: result.finalVerdicts.aggressive,  color:'F97316'  },
      { title:'平衡觀點', text: result.finalVerdicts.balanced,    color:T.accent  },
      { title:'保守觀點', text: result.finalVerdicts.conservative, color:T.accent3 },
    ];
    const gap  = 0.15;
    const vW   = (CW - gap*2) / 3;   // ≈ 2.93"
    const vH   = BODY_H;              // = 5.93"，底 = BODY_B ✓

    verdicts.forEach((v, i) => {
      const vx = MX + i * (vW + gap);
      addCard(sl, T, vx, BODY_T, vW, vH, v.color+'55');
      sl.addShape('rect', { x:vx, y:BODY_T, w:vW, h:0.04, fill:{ type:'solid', color:v.color } });
      // 標題  h=0.32"
      sl.addText(v.title, { x:vx+0.12, y:BODY_T+0.1, w:vW-0.24, h:0.32, fontSize:12.5, bold:true, color:v.color });
      sl.addShape('rect', { x:vx+0.12, y:BODY_T+0.45, w:0.75, h:0.025, fill:{ type:'solid', color:v.color } });

      // bullets：最多 6 條，每條高度均分剩餘空間
      // 剩餘 = vH - 0.55 = 5.38"，6條 = 0.90"/條
      const vLines = toLines(v.text, 6, 20);
      const bH     = Math.min(0.88, (vH - 0.55) / vLines.length);
      vLines.forEach((b, bi) => {
        const bY = BODY_T + 0.55 + bi * bH;
        // 確認不超出卡片底（= BODY_T + vH = BODY_B）
        if (bY + bH > BODY_B - 0.04) return;
        sl.addText([
          { text:'▸ ', options:{ color:v.color, bold:true, fontSize:10 } },
          { text:b,   options:{ color:'CBD5E1',             fontSize:10 } },
        ], { x:vx+0.12, y:bY, w:vW-0.24, h:bH, valign:'middle' });
      });
    });
    addPageNum(sl, 9, TOTAL, T);
  }

  // ──────────────────────────────────────────────────
  // SLIDE 10 — CALL TO ACTION
  // ──────────────────────────────────────────────────
  {
    const sl = pptx.addSlide();
    sl.background = { fill: T.bg };
    sl.addShape('rect',    { x:0,   y:0,   w:'100%', h:0.05, fill:{ type:'solid', color:T.accent } });
    sl.addShape('ellipse', { x:2.8, y:0.8, w:4.5,   h:4.5,  fill:{ type:'solid', color:T.card  } });

    // 主標題  y:1.5 h:0.9
    sl.addText('下一步行動', { x:0.5, y:1.5, w:9.0, h:0.9, fontSize:38, bold:true, color:'F8FAFC', align:'center' });
    // tagline  y:2.5 h:0.38
    sl.addText(T.tagline,   { x:0.5, y:2.5, w:9.0, h:0.38, fontSize:10.5, color:T.accent, align:'center', charSpacing:2.5, bold:true });

    // 機率卡  y:3.1 ~ 4.6  (h=1.5)
    const sc = result.successProbability;
    const sC = scoreColor(sc);
    addCard(sl, T, 3.5, 3.1, 3.0, 1.5, sC);
    sl.addText('AI 評估成功機率', { x:3.5, y:3.22, w:3.0, h:0.28, fontSize:9.5,  color:'94A3B8', align:'center' });
    sl.addText(`${sc}%`,          { x:3.5, y:3.54, w:3.0, h:0.85, fontSize:34,   bold:true, color:sC, align:'center' });

    // 底部  y:6.55 h:0.28
    sl.addText('此報告由 OmniView AI 360° 虛擬董事會自動生成', { x:1, y:6.55, w:8, h:0.28, fontSize:9.5, color:'475569', align:'center' });
    addPageNum(sl, 10, TOTAL, T);
  }

  await pptx.writeFile({ fileName: 'OmniView_商業提案報告.pptx' });
};

import PptxGenJS from 'pptxgenjs';
import { AnalysisResult } from '../types';

// 16:9 = 10" × 7.5"
const W      = 10;
const MX     = 0.45;
const CW     = W - MX * 2;   // 9.1"
const BODY_T = 0.82;
const BODY_B = 6.75;
const BODY_H = BODY_B - BODY_T;  // 5.93"

// ── 文字工具 ──────────────────────────────────────────────

/** 硬截，保證輸出 ≤ max 字 */
const cut = (s: string, max: number): string => {
  if (!s) return '—';
  const t = s.trim().replace(/\s+/g, ' ').replace(/[。！？!?.…、,，]+$/g, '').trim();
  return t.length <= max ? t : t.slice(0, max - 1) + '…';
};

/**
 * 從長段落萃取單行短標籤（保證 ≤ maxLen 字）
 * 句號取第一句 → 逗號取第一子句 → 硬截
 */
const tag = (text: string, maxLen: number): string => {
  if (!text) return '—';
  const t = text.trim().replace(/[。！？!?.]+$/g, '').trim();
  if (t.length <= maxLen) return t;
  const sen = t.split(/[。！？!?；;]/)[0].trim();
  if (sen.length <= maxLen) return sen;
  const clause = sen.split(/[，,、]/)[0].trim();
  return cut(clause || sen, maxLen);
};

/**
 * 把長文切成最多 maxLines 條，每條保證 ≤ maxLen 字
 * 句號分句 → 逗號取第一子句 → 硬截
 */
const toLines = (text: string, maxLines: number, maxLen: number): string[] => {
  if (!text) return ['—'];
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
      const parts = sen.split(/[，,、]/).map(p => p.trim()).filter(p => p.length > 0);
      let line = '';
      for (const p of parts) {
        const next = line ? line + '、' + p : p;
        if (next.length > maxLen) break;
        line = next;
      }
      out.push(cut(line || parts[0] || sen, maxLen));
    }
  }
  return out.length ? out : [cut(text, maxLen)];
};

const fmt = (n: number) => {
  const abs = Math.abs(n);
  if (abs >= 1_000_000) return `$${(n / 1_000_000).toFixed(1)}M`;
  if (abs >= 1_000)     return `$${(n / 1_000).toFixed(0)}K`;
  return `$${n}`;
};

// ── 產業主題 ──────────────────────────────────────────────

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
  if (/醫|健|藥|療|病|診|health|medical|biotech/.test(t))        return THEMES.health;
  if (/餐|食|飲|咖|料理|外送|food|restaurant|beverage/.test(t)) return THEMES.food;
  if (/金融|投資|貸|保險|fintech|finance|bank|crypto/.test(t))   return THEMES.finance;
  if (/零售|電商|購物|fashion|retail|ecommerce/.test(t))         return THEMES.retail;
  if (/教育|學習|課程|edu|learning|school|training/.test(t))     return THEMES.edu;
  if (/科技|ai|app|平台|軟體|saas|tech|software|雲端/.test(t))  return THEMES.tech;
  return THEMES.default;
};

// ── 共用元件 ──────────────────────────────────────────────

const addBg = (sl: PptxGenJS.Slide, T: Theme) => {
  sl.background = { fill: T.bg };
  sl.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ type:'solid', color:T.accent } });
};
const addHeader = (sl: PptxGenJS.Slide, T: Theme, title: string) => {
  sl.addText(title, { x:MX, y:0.1, w:8, h:0.48, fontSize:19, bold:true, color:'F8FAFC', fontFace:'Arial' });
  sl.addShape('rect', { x:MX, y:0.6, w:0.85, h:0.03, fill:{ type:'solid', color:T.accent } });
};
const addPageNum = (sl: PptxGenJS.Slide, n: number, total: number, T: Theme) =>
  sl.addText(`${n} / ${total}`, { x:8.6, y:6.95, w:0.8, h:0.25, fontSize:7.5, color:T.divider, align:'right' });
const addCard = (sl: PptxGenJS.Slide, T: Theme, x:number, y:number, w:number, h:number, border?: string) =>
  sl.addShape('roundRect', { x, y, w, h, rectRadius:0.1,
    fill:{ type:'solid', color:T.card }, line:{ color:border||T.divider, width:0.75 } });
const scoreColor = (s: number) =>
  s >= 80 ? '10B981' : s >= 60 ? '3B82F6' : s >= 40 ? 'FBBF24' : 'EF4444';

// ── 主函式 ────────────────────────────────────────────────

export const generatePptx = async (result: AnalysisResult): Promise<void> => {
  const T = detectTheme(result.executiveSummary, result.marketAnalysis.description);
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';
  pptx.author = 'OmniView AI';
  pptx.title  = 'OmniView 商業提案分析報告';
  const TOTAL = 10;

  // ════ SLIDE 1 COVER ════════════════════════════════
  {
    const sl = pptx.addSlide();
    sl.background = { fill: T.bg };
    sl.addShape('rect',    { x:0,   y:0,    w:'100%', h:0.05, fill:{ type:'solid', color:T.accent } });
    sl.addShape('ellipse', { x:6.5, y:-0.8, w:4.0, h:4.0, fill:{ type:'solid', color:T.card } });
    sl.addShape('ellipse', { x:-1,  y:5.0,  w:3.0, h:3.0, fill:{ type:'solid', color:T.card } });

    sl.addShape('roundRect', { x:0.5, y:0.3, w:0.55, h:0.55, rectRadius:0.08, fill:{ type:'solid', color:T.accent } });
    sl.addText('OV',          { x:0.5, y:0.3, w:0.55, h:0.55, fontSize:13, bold:true, color:'FFFFFF', align:'center', valign:'middle' });
    sl.addText('OmniView AI', { x:1.15, y:0.35, w:3, h:0.45, fontSize:15, bold:true, color:'F8FAFC' });

    sl.addShape('roundRect', { x:0.5, y:1.05, w:3.0, h:0.32, rectRadius:0.16,
      fill:{ type:'solid', color:T.accent+'28' }, line:{ color:T.accent+'60', width:0.75 } });
    sl.addText(T.tagline, { x:0.5, y:1.05, w:3.0, h:0.32, fontSize:8, bold:true, color:T.accent,
      align:'center', valign:'middle', charSpacing:1.5 });

    // 主標題：fontSize=40，2行，h=1.7"（40pt×2行×1.1spacing = 1.22" < 1.7 ✓）
    sl.addText('商業提案\n分析報告', { x:0.5, y:1.55, w:8.5, h:1.7, fontSize:40, bold:true, color:'F8FAFC', lineSpacingMultiple:1.1 });

    const sc = result.successProbability;
    const sC = scoreColor(sc);
    sl.addShape('roundRect', { x:0.5, y:3.5, w:2.3, h:0.85, rectRadius:0.12,
      fill:{ type:'solid', color:sC+'22' }, line:{ color:sC, width:1.5 } });
    sl.addText('AI 評估成功機率', { x:0.5, y:3.56, w:2.3, h:0.26, fontSize:8.5, color:sC, align:'center' });
    sl.addText(`${sc}%`, { x:0.5, y:3.82, w:2.3, h:0.48, fontSize:26, bold:true, color:sC, align:'center', valign:'middle' });

    // 摘要：h=0.38，fontSize=11，lineH=11/72*1.15*1.4=0.246" → 1行 → maxLen=50字 ✓
    sl.addText(tag(result.executiveSummary, 50), { x:0.5, y:4.55, w:9.0, h:0.38, fontSize:11, color:'94A3B8', italic:true });

    sl.addShape('rect', { x:0, y:6.9, w:'100%', h:0.38, fill:{ type:'solid', color:T.card } });
    sl.addText('由 OmniView AI 360° 虛擬董事會自動生成', { x:0.5, y:6.93, w:6, h:0.28, fontSize:8.5, color:'475569' });
    addPageNum(sl, 1, TOTAL, T);
  }

  // ════ SLIDE 2 EXECUTIVE SUMMARY ════════════════════
  // 左卡 w=5.35 h=BODY_H，右側3個KPI卡均分BODY_H
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '執行摘要');

    const lW = 5.35;
    addCard(sl, T, MX, BODY_T, lW, BODY_H);
    sl.addText('核心觀點', { x:0.65, y:BODY_T+0.12, w:lW-0.4, h:0.28, fontSize:10.5, bold:true, color:T.accent });

    // bullets：5條，bRowH=(BODY_H-0.5)/5=1.086"
    // fontSize=11.5，lineH=0.257"，bRowH可放4行，每行31字 → maxLen=26 ✓
    const bLines = toLines(result.executiveSummary, 5, 26);
    const bRowH  = (BODY_H - 0.5) / bLines.length;
    bLines.forEach((b, i) => {
      sl.addText([
        { text:'▸  ', options:{ color:T.accent, bold:true, fontSize:11.5 } },
        { text:b,    options:{ color:'CBD5E1',  fontSize:11.5 } },
      ], { x:0.62, y:BODY_T+0.5+i*bRowH, w:lW-0.3, h:bRowH, valign:'middle' });
    });

    // 右側KPI：kH=(BODY_H-0.24)/3=1.897"
    // val fontSize=14，h=0.36"（1行）→ lineH=14/72*1.15*1.4=0.314" → 放1行 → maxLen=11字 ✓
    const kGap = 0.12;
    const kH   = (BODY_H - kGap * 2) / 3;
    const kX   = MX + lW + 0.1;
    const kW   = W - MX - kX;
    const kpis = [
      { label:'市場規模 TAM', val: tag(result.marketAnalysis.size,      11), color:T.accent  },
      { label:'CAGR 年均成長', val: tag(result.marketAnalysis.growthRate,11), color:T.accent2 },
      { label:'損益平衡點',    val: tag(result.breakEvenPoint,           11), color:T.accent3 },
    ];
    kpis.forEach((k, i) => {
      const ky = BODY_T + i * (kH + kGap);
      addCard(sl, T, kX, ky, kW, kH, k.color+'55');
      sl.addShape('rect', { x:kX, y:ky, w:0.04, h:kH, fill:{ type:'solid', color:k.color } });
      sl.addText(k.label, { x:kX+0.14, y:ky+0.14, w:kW-0.2, h:0.26, fontSize:9, color:'94A3B8' });
      // h=0.36 只放1行，fontSize=14
      sl.addText(k.val,   { x:kX+0.14, y:ky+0.48, w:kW-0.2, h:0.36, fontSize:14, bold:true, color:k.color });
    });

    addPageNum(sl, 2, TOTAL, T);
  }

  // ════ SLIDE 3 MARKET OPPORTUNITY ═══════════════════
  // 上方3個KPI卡 h=1.2"，下方市場洞察卡至BODY_B
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '市場機會');

    const KPI_H = 1.2;
    const kW    = (CW - 0.2) / 3;  // 2.967"
    // KPI val：fontSize=14，h=0.36"，lineH=0.314" → 1行 → maxLen=floor(2.727/0.194)=14字
    const kpis = [
      { label:'TAM 市場規模', val: tag(result.marketAnalysis.size,      10), color:T.accent  },
      { label:'CAGR 年均成長', val: tag(result.marketAnalysis.growthRate,10), color:T.accent2 },
      { label:'損益平衡點',    val: tag(result.breakEvenPoint,           10), color:T.accent3 },
    ];
    kpis.forEach((k, i) => {
      const kx = MX + i * (kW + 0.1);
      addCard(sl, T, kx, BODY_T, kW, KPI_H, k.color+'66');
      sl.addShape('rect', { x:kx, y:BODY_T, w:kW, h:0.04, fill:{ type:'solid', color:k.color } });
      sl.addText(k.label, { x:kx+0.12, y:BODY_T+0.1,  w:kW-0.24, h:0.26, fontSize:9.5, color:'94A3B8' });
      sl.addText(k.val,   { x:kx+0.12, y:BODY_T+0.44, w:kW-0.24, h:0.36, fontSize:14, bold:true, color:k.color });
    });

    // 市場洞察卡：top=2.17，底=BODY_B ✓
    const mTop = BODY_T + KPI_H + 0.15;
    const mH   = BODY_B - mTop;  // 4.58"
    addCard(sl, T, MX, mTop, CW, mH);
    sl.addText('市場洞察', { x:MX+0.2, y:mTop+0.12, w:CW-0.4, h:0.28, fontSize:10.5, bold:true, color:T.accent });

    // bullets：5條，mRowH=(4.58-0.5)/5=0.816"
    // fontSize=11，lineH=0.246"，每條3行，每行56字 → maxLen=38 ✓
    const mLines = toLines(result.marketAnalysis.description, 5, 38);
    const mRowH  = (mH - 0.5) / mLines.length;
    mLines.forEach((b, i) => {
      sl.addText([
        { text:'◆  ', options:{ color:T.accent2, bold:true, fontSize:11 } },
        { text:b,    options:{ color:'CBD5E1',   fontSize:11 } },
      ], { x:MX+0.2, y:mTop+0.5+i*mRowH, w:CW-0.4, h:mRowH, valign:'middle' });
    });

    addPageNum(sl, 3, TOTAL, T);
  }

  // ════ SLIDE 4 FINANCIAL PROJECTIONS ════════════════
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '財務預測');

    const finData = result.financials.slice(0, 4);
    const tHdrH  = 0.42;
    const tRowH  = 0.46;
    const tH     = tHdrH + finData.length * tRowH;

    const fHdr: PptxGenJS.TableCell[] = ['年度','營收','成本','淨利'].map(t => ({
      text: t,
      options: { bold:true, color:'FFFFFF', fill:{ color:T.headerBg }, fontSize:11.5, align:'center' as const },
    }));
    const fRows: PptxGenJS.TableCell[][] = [fHdr];
    finData.forEach((f, i) => {
      const bg  = i%2===0 ? T.card : T.bg;
      const pft = Number(f.profit)||0;
      fRows.push([
        { text: cut(f.year??'—', 8),      options:{ color:'F8FAFC',                       fill:{color:bg}, fontSize:11.5, align:'center' as const, bold:true } },
        { text: fmt(Number(f.revenue)||0), options:{ color:T.accent,                        fill:{color:bg}, fontSize:11.5, align:'center' as const, bold:true } },
        { text: fmt(Number(f.costs)||0),   options:{ color:'94A3B8',                        fill:{color:bg}, fontSize:11.5, align:'center' as const } },
        { text: fmt(pft),                  options:{ color: pft>=0 ? T.accent2 : 'EF4444', fill:{color:bg}, fontSize:11.5, align:'center' as const, bold:true } },
      ]);
    });
    sl.addTable(fRows, {
      x:MX, y:BODY_T, w:CW,
      border:{ type:'solid', color:T.divider, pt:0.5 },
      colW:[2.275, 2.275, 2.275, 2.275],
      rowH:[tHdrH, ...finData.map(()=>tRowH)],
    });

    // 長條圖：top=BODY_T+tH+0.2，底=BODY_B ✓
    const barTop = BODY_T + tH + 0.2;
    const barH   = BODY_B - barTop;
    if (barH >= 0.9) {
      addCard(sl, T, MX, barTop, CW, barH);
      sl.addText('營收趨勢', { x:0.65, y:barTop+0.1, w:3, h:0.24, fontSize:9.5, bold:true, color:T.accent });
      sl.addText(`損益平衡：${tag(result.breakEvenPoint, 18)}`, { x:4.5, y:barTop+0.1, w:4.75, h:0.24, fontSize:9.5, color:T.accent2, align:'right' });

      const maxRev = Math.max(...finData.map(f=>Number(f.revenue)||0), 1);
      const barSX  = 1.5;
      const barMaxW= 6.8;
      const rowH2  = Math.min(0.36, (barH - 0.42) / finData.length);
      finData.forEach((f, i) => {
        const rev = Number(f.revenue)||0;
        const bw  = Math.max(rev/maxRev*barMaxW, 0.12);
        const by  = barTop + 0.38 + i * rowH2;
        if (by + rowH2 > BODY_B - 0.05) return;
        sl.addText(cut(f.year??'—', 6), { x:0.55, y:by, w:0.85, h:rowH2, fontSize:8.5, color:'94A3B8', valign:'middle' });
        sl.addShape('roundRect', { x:barSX, y:by+0.04, w:bw, h:rowH2-0.08, rectRadius:0.04, fill:{ type:'solid', color:T.accent } });
        const lx = barSX+bw+0.08;
        if (lx+0.85 <= MX+CW) {
          sl.addText(fmt(rev), { x:lx, y:by, w:0.85, h:rowH2, fontSize:8.5, color:T.accent, bold:true, valign:'middle' });
        } else {
          sl.addText(fmt(rev), { x:barSX+bw-0.95, y:by, w:0.85, h:rowH2, fontSize:8.5, color:'FFFFFF', bold:true, valign:'middle' });
        }
      });
    }
    addPageNum(sl, 4, TOTAL, T);
  }

  // ════ SLIDE 5 COMPETITIVE LANDSCAPE ════════════════
  // 表格行高動態確保總高=BODY_H，底=BODY_B ✓
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '競爭態勢分析');

    const comps = result.competitors.slice(0, 5);
    const hdrH  = 0.40;
    const rowH  = (BODY_H - hdrH) / comps.length;
    // fontSize=10，rowH動態，1行字寬=10/72"=0.139"，欄寬3.55" → 25字/行
    // rowH≈1.106"/條（5條時），lineH=0.224" → 4行 → 4×25=100字
    // 但只要1行有意義，maxLen=22 安全

    const cHdr: PptxGenJS.TableCell[] = ['競爭對手','優勢','劣勢'].map(t => ({
      text: t,
      options: { bold:true, color:'FFFFFF', fill:{ color:T.headerBg }, fontSize:11.5, align:'center' as const },
    }));
    const cRows: PptxGenJS.TableCell[][] = [cHdr];
    comps.forEach((c, i) => {
      const bg = i%2===0 ? T.card : T.bg;
      cRows.push([
        { text: tag(c.name,     12), options:{ bold:true, color:'F8FAFC', fill:{color:bg}, fontSize:11, align:'center' as const } },
        { text: tag(c.strength, 22), options:{ color:T.accent2,           fill:{color:bg}, fontSize:10 } },
        { text: tag(c.weakness, 22), options:{ color:'F87171',            fill:{color:bg}, fontSize:10 } },
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

  // ════ SLIDE 6 STRATEGIC ROADMAP ════════════════════
  // 4個item，iH=BODY_H/4=1.482"，每個item所有元素在[iY, iY+iH)內
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '策略路線圖');

    const items = result.roadmap.slice(0, 4);
    const iH    = BODY_H / items.length;

    items.forEach((item, i) => {
      const iY  = BODY_T + i * iH;
      const iB  = iY + iH;

      if (i < items.length-1)
        sl.addShape('rect', { x:0.61, y:iY+0.28, w:0.025, h:iH-0.28, fill:{ type:'solid', color:T.divider } });

      sl.addShape('ellipse', { x:0.49, y:iY+0.1, w:0.25, h:0.25, fill:{ type:'solid', color:T.accent } });
      sl.addText(`${i+1}`, { x:0.49, y:iY+0.1, w:0.25, h:0.25,
        fontSize:8.5, bold:true, color:'FFFFFF', align:'center', valign:'middle' });

      // Phase：h=0.3，fontSize=12.5，lineH=0.279" → 1行 → maxLen=16 ✓
      sl.addText(tag(item.phase, 16), { x:0.88, y:iY+0.05, w:3.3, h:0.3, fontSize:12.5, bold:true, color:T.accent });

      // Timeframe chip：h=0.25，fontSize=8.5，1行 → maxLen=10 ✓
      sl.addShape('roundRect', { x:4.35, y:iY+0.07, w:1.45, h:0.25, rectRadius:0.12,
        fill:{ type:'solid', color:T.accent+'25' }, line:{ color:T.accent+'60', width:0.5 } });
      sl.addText(tag(item.timeframe, 10), { x:4.35, y:iY+0.07, w:1.45, h:0.25,
        fontSize:8.5, color:T.accent, align:'center', valign:'middle' });

      // 內容卡：top=iY+0.38，底=iB-0.05，cH=iH-0.43=1.052"
      const cTop = iY + 0.38;
      const cH   = iB - cTop - 0.05;
      if (cH < 0.3) return;
      addCard(sl, T, 0.88, cTop, CW-0.45, cH);

      // 產品/技術：txtH=cH-0.3=0.752"，fontSize=9.5，lineH=0.213" → 3行
      // 欄寬half≈4.275"，3行×30字=90字 → maxLen=20 安全
      const txtH = cH - 0.3;
      const half = (CW - 0.45 - 0.1) / 2;
      sl.addText('產品', { x:1.06, y:cTop+0.06, w:0.6, h:0.2, fontSize:8.5, bold:true, color:T.accent2 });
      sl.addText(tag(item.product, 20), { x:1.06, y:cTop+0.27, w:half-0.2, h:txtH, fontSize:9.5, color:'CBD5E1', valign:'top' });
      sl.addShape('rect', { x:0.88+half+0.05, y:cTop+0.06, w:0.02, h:cH-0.12, fill:{ type:'solid', color:T.divider } });
      const rx = 0.88 + half + 0.12;
      sl.addText('技術', { x:rx, y:cTop+0.06, w:0.6, h:0.2, fontSize:8.5, bold:true, color:T.accent3 });
      sl.addText(tag(item.technology, 20), { x:rx, y:cTop+0.27, w:half-0.2, h:txtH, fontSize:9.5, color:'CBD5E1', valign:'top' });
    });
    addPageNum(sl, 6, TOTAL, T);
  }

  // ════ SLIDE 7 RISK ASSESSMENT ══════════════════════
  // 2欄×3列，cH均分BODY_H，所有元素在卡片內
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '風險評估');

    const risks = result.risks.slice(0, 6);
    const COLS  = 2;
    const nRows = Math.ceil(risks.length / COLS);
    const gap   = 0.1;
    const cW    = (CW - gap) / COLS;
    const cH    = (BODY_H - gap*(nRows-1)) / nRows;
    // 6個risk：nRows=3，cH=(5.93-0.2)/3=1.91"
    // 風險標題：h=0.28，fontSize=11，1行，欄寬3.5" → maxLen=18 ✓
    // 因應：h=cH-0.48=1.43"，fontSize=8.5，lineH=0.19" → 7行，欄寬4.14" → maxLen=26 ✓

    risks.forEach((r, i) => {
      const col = i%COLS, row = Math.floor(i/COLS);
      const cx  = MX + col*(cW+gap);
      const cy  = BODY_T + row*(cH+gap);
      if (cy+cH > BODY_B+0.01) return;

      const ic    = r.impact==='High' ? 'EF4444' : r.impact==='Medium' ? 'FBBF24' : '10B981';
      const label = r.impact==='High' ? '高' : r.impact==='Medium' ? '中' : '低';

      addCard(sl, T, cx, cy, cW, cH, ic+'55');
      sl.addShape('rect', { x:cx, y:cy, w:cW, h:0.04, fill:{ type:'solid', color:ic } });

      sl.addShape('roundRect', { x:cx+cW-0.78, y:cy+0.1, w:0.64, h:0.24, rectRadius:0.12,
        fill:{ type:'solid', color:ic+'30' }, line:{ color:ic, width:0.5 } });
      sl.addText(label, { x:cx+cW-0.78, y:cy+0.1, w:0.64, h:0.24,
        fontSize:8.5, bold:true, color:ic, align:'center', valign:'middle' });

      sl.addText(tag(r.risk, 18), { x:cx+0.14, y:cy+0.1, w:cW-0.98, h:0.28, fontSize:11, bold:true, color:'F8FAFC' });

      const mitH = cH - 0.48;
      if (mitH > 0.15) {
        sl.addText([
          { text:'因應：', options:{ fontSize:8.5, color:'64748B', bold:true } },
          { text: tag(r.mitigation, 26), options:{ fontSize:8.5, color:'CBD5E1' } },
        ], { x:cx+0.14, y:cy+0.44, w:cW-0.26, h:mitH, valign:'top' });
      }
    });
    addPageNum(sl, 7, TOTAL, T);
  }

  // ════ SLIDE 8 STAKEHOLDER BOARD ════════════════════
  // 5人=3+2，pCH均分BODY_H
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, 'AI 虛擬董事會');

    const personas = result.personaEvaluations.slice(0, 5);
    const COLS  = 3;
    const nRows = Math.ceil(personas.length / COLS);  // 2
    const gap   = 0.1;
    const cW    = (CW - gap*(COLS-1)) / COLS;  // 2.967"
    const cH    = (BODY_H - gap*(nRows-1)) / nRows;  // (5.93-0.1)/2=2.915"
    // 角色名：h=0.28，fontSize=11.5，1行，欄寬2.567" → maxLen=10 ✓
    // 引言：y=cy+0.54，h=0.42"，fontSize=9，lineH=0.201" → 2行，每行21字 → maxLen=18 ✓
    // 擔憂：y=cy+1.0，h=cH-1.08=1.835"，fontSize=8.5 → maxLen=20 ✓

    personas.forEach((p, i) => {
      const col = i%COLS, row = Math.floor(i/COLS);
      const cx  = MX + col*(cW+gap);
      const cy  = BODY_T + row*(cH+gap);
      if (cy+cH > BODY_B+0.01) return;

      const sc = Number(p.score)||0;
      const sC = scoreColor(sc);

      addCard(sl, T, cx, cy, cW, cH);
      sl.addShape('rect', { x:cx, y:cy, w:cW, h:0.04, fill:{ type:'solid', color:sC } });

      sl.addText(tag(p.role, 10), { x:cx+0.12, y:cy+0.1, w:cW-0.85, h:0.28, fontSize:11.5, bold:true, color:'F8FAFC' });
      sl.addText(`${sc}`, { x:cx+cW-0.82, y:cy+0.08, w:0.68, h:0.3, fontSize:17, bold:true, color:sC, align:'right' });

      const barFill = (cW-0.24)*(sc/100);
      sl.addShape('roundRect', { x:cx+0.12, y:cy+0.42, w:cW-0.24, h:0.055, rectRadius:0.028, fill:{ type:'solid', color:T.divider } });
      sl.addShape('roundRect', { x:cx+0.12, y:cy+0.42, w:Math.max(barFill,0.05), h:0.055, rectRadius:0.028, fill:{ type:'solid', color:sC } });

      sl.addText(`"${tag(p.keyQuote, 18)}"`, { x:cx+0.12, y:cy+0.54, w:cW-0.24, h:0.42,
        fontSize:9, italic:true, color:'94A3B8', valign:'top' });

      const conH = cH - 1.08;
      if (conH > 0.15) {
        sl.addText([
          { text:'! ', options:{ fontSize:8.5, color:'FBBF24', bold:true } },
          { text: tag(p.concern, 20), options:{ fontSize:8.5, color:'CBD5E1' } },
        ], { x:cx+0.12, y:cy+1.0, w:cW-0.24, h:conH, valign:'top' });
      }
    });
    addPageNum(sl, 8, TOTAL, T);
  }

  // ════ SLIDE 9 FINAL VERDICTS ═══════════════════════
  // 3欄，vH=BODY_H，底=BODY_B ✓
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '最終裁決');

    const verdicts = [
      { title:'激進觀點', text: result.finalVerdicts.aggressive,   color:'F97316'  },
      { title:'平衡觀點', text: result.finalVerdicts.balanced,     color:T.accent  },
      { title:'保守觀點', text: result.finalVerdicts.conservative, color:T.accent3 },
    ];
    const gap = 0.15;
    const vW  = (CW - gap*2) / 3;  // 2.933"
    const vH  = BODY_H;

    verdicts.forEach((v, i) => {
      const vx = MX + i*(vW+gap);
      addCard(sl, T, vx, BODY_T, vW, vH, v.color+'55');
      sl.addShape('rect', { x:vx, y:BODY_T, w:vW, h:0.04, fill:{ type:'solid', color:v.color } });
      sl.addText(v.title, { x:vx+0.12, y:BODY_T+0.1, w:vW-0.24, h:0.3, fontSize:12.5, bold:true, color:v.color });
      sl.addShape('rect', { x:vx+0.12, y:BODY_T+0.43, w:0.75, h:0.025, fill:{ type:'solid', color:v.color } });

      // bullets：6條，bH=(BODY_H-0.5)/6=0.905"
      // fontSize=10，lineH=0.225" → 每條4行 → maxLen=16 安全
      const vLines = toLines(v.text, 6, 16);
      const bH     = (vH - 0.5) / vLines.length;
      vLines.forEach((b, bi) => {
        const bY = BODY_T + 0.5 + bi*bH;
        if (bY+bH > BODY_B-0.02) return;
        sl.addText([
          { text:'▸ ', options:{ color:v.color, bold:true, fontSize:10 } },
          { text:b,   options:{ color:'CBD5E1',             fontSize:10 } },
        ], { x:vx+0.12, y:bY, w:vW-0.24, h:bH, valign:'middle' });
      });
    });
    addPageNum(sl, 9, TOTAL, T);
  }

  // ════ SLIDE 10 CALL TO ACTION ══════════════════════
  {
    const sl = pptx.addSlide();
    sl.background = { fill: T.bg };
    sl.addShape('rect',    { x:0,   y:0,   w:'100%', h:0.05, fill:{ type:'solid', color:T.accent } });
    sl.addShape('ellipse', { x:2.8, y:0.8, w:4.5, h:4.5, fill:{ type:'solid', color:T.card } });

    sl.addText('下一步行動', { x:0.5, y:1.5,  w:9.0, h:0.9,  fontSize:38, bold:true, color:'F8FAFC', align:'center' });
    sl.addText(T.tagline,   { x:0.5, y:2.5,  w:9.0, h:0.36, fontSize:10.5, color:T.accent, align:'center', charSpacing:2.5, bold:true });

    const sc = result.successProbability;
    const sC = scoreColor(sc);
    addCard(sl, T, 3.5, 3.1, 3.0, 1.5, sC);
    sl.addText('AI 評估成功機率', { x:3.5, y:3.22, w:3.0, h:0.28, fontSize:9.5, color:'94A3B8', align:'center' });
    sl.addText(`${sc}%`,          { x:3.5, y:3.54, w:3.0, h:0.85, fontSize:34, bold:true, color:sC, align:'center' });

    sl.addText('此報告由 OmniView AI 360° 虛擬董事會自動生成',
      { x:1, y:6.55, w:8, h:0.28, fontSize:9.5, color:'475569', align:'center' });
    addPageNum(sl, 10, TOTAL, T);
  }

  await pptx.writeFile({ fileName: 'OmniView_商業提案報告.pptx' });
};

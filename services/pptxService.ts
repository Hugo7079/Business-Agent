import PptxGenJS from 'pptxgenjs';
import { AnalysisResult } from '../types';

// ── 版面常數 (16:9 標準尺寸: 10" x 5.625") ─────────────────────────
const W      = 10;
const H      = 5.625;
const MX     = 0.5;            // 左右邊距

// ── 上下留白更充裕 ──
const BODY_T = 1.5;            // 內容區頂部（Header 佔 0~1.5"，留白充足）
const BODY_B = 4.95;           // 內容區底部（Footer 佔 4.95~5.625"，留白充足）
const BODY_H = BODY_B - BODY_T; // 3.45"

const CW     = W - MX * 2;    // 內容寬度 9"

// ── 字體大小常數 ──────────────────────────────────────────────
const FONT_TITLE    = 32;
const FONT_SUBTITLE = 18;
const FONT_SECTION  = 13;   // 章節小標（≤8字）
const FONT_BODY     = 9;    // 內文基準（>8字）↓ 從10降為9，全局減壓
const FONT_SMALL    = 8;    // 小標籤
const FONT_XSMALL   = 7;    // 極小（footer / 因應對策）

/** 依字數判斷：≤8 字視為標題，>8 字視為內文 */
const titleOrBody = (txt: string): number =>
  (txt?.length || 0) <= 8 ? FONT_SECTION : FONT_BODY;

/** 把行陣列合併成一個 string，供整塊文字框使用 */
const joinBullets = (lines: string[], bullet = '▸  ') =>
  lines.map(l => bullet + l).join('\n');

const fmt = (n: number) => {
  const abs = Math.abs(n);
  if (abs >= 1_000_000) return `$${(n / 1_000_000).toFixed(1)}M`;
  if (abs >= 1_000)     return `$${(n / 1_000).toFixed(0)}K`;
  return `$${n}`;
};

// ── 產業主題 ──────────────────────────────────────────────

interface Theme {
  bg: string; card: string; accent: string; accent2: string; accent3: string;
  headerBg: string; divider: string; tagline: string;
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
  sl.addShape('rect', { x:0, y:0, w:'100%', h:0.06, fill:{ type:'solid', color:T.accent } });
};

const addHeader = (sl: PptxGenJS.Slide, T: Theme, title: string) => {
  // Logo 小標 左上
  sl.addShape('roundRect', { x:MX, y:0.18, w:0.4, h:0.4, rectRadius:0.06, fill:{ type:'solid', color:T.accent } });
  sl.addText('OV', { x:MX, y:0.18, w:0.4, h:0.4, fontSize:12, bold:true, color:'FFFFFF', align:'center', valign:'middle' });
  // 頁面標題
  sl.addText(title, { x:MX+0.55, y:0.22, w:7.5, h:0.45, fontSize:22, bold:true, color:'F8FAFC', fontFace:'Arial', valign:'middle' });
  // 分隔線，緊貼在 BODY_T 上方
  sl.addShape('rect', { x:MX, y:BODY_T - 0.12, w:CW, h:0.02, fill:{ type:'solid', color:T.divider } });
  sl.addShape('rect', { x:MX, y:BODY_T - 0.12, w:0.9, h:0.02, fill:{ type:'solid', color:T.accent } });
};

const addFooter = (sl: PptxGenJS.Slide, T: Theme, n: number, total: number) => {
  // Footer 分隔線
  sl.addShape('rect', { x:MX, y:BODY_B + 0.1, w:CW, h:0.02, fill:{ type:'solid', color:T.divider } });
  sl.addText('由 OmniView AI 360° 虛擬董事會自動生成',
    { x:MX, y:BODY_B + 0.16, w:7, h:0.28, fontSize:FONT_XSMALL, color:'475569', valign:'middle' });
  sl.addText(`${n} / ${total}`,
    { x:W-1.5, y:BODY_B + 0.16, w:1.0, h:0.28, fontSize:FONT_XSMALL, color:'475569', align:'right', valign:'middle' });
};

const addCard = (sl: PptxGenJS.Slide, T: Theme, x:number, y:number, w:number, h:number, border?: string) =>
  sl.addShape('roundRect', { x, y, w, h, rectRadius:0.1,
    fill:{ type:'solid', color:T.card }, line:{ color:border||T.divider, width:0.75 } });

const scoreColor = (s: number) =>
  s >= 80 ? '10B981' : s >= 60 ? '3B82F6' : s >= 40 ? 'FBBF24' : 'EF4444';

// ── 內文自動縮排輔助：所有長文字框統一設定 autoFit + shrinkText ──
const bodyTextOpts = (
  x:number, y:number, w:number, h:number,
  extra: PptxGenJS.TextPropsOptions = {}
): PptxGenJS.TextPropsOptions => ({
  x, y, w, h,
  fontSize: FONT_BODY,
  color: 'CBD5E1',
  lineSpacingMultiple: 1.15,
  valign: 'top',
  shrinkText: true,
  ...extra,
});

// ── 主函式 ────────────────────────────────────────────────

export const generatePptx = async (result: AnalysisResult): Promise<void> => {
  const r = result;
  const T = detectTheme(r.executiveSummary, r.marketAnalysis.description);
  const pptx = new PptxGenJS();

  // ── 明確設定 16:9 (10 x 5.625 英寸) ──
  pptx.defineLayout({ name:'LAYOUT_16x9', width:10, height:5.625 });
  pptx.layout = 'LAYOUT_16x9';
  pptx.author = 'OmniView AI';
  pptx.title  = 'OmniView 商業提案分析報告';
  const TOTAL = 10;

  // ════ SLIDE 1 COVER ════════════════════════════════
  {
    const sl = pptx.addSlide();
    sl.background = { fill: T.bg };
    sl.addShape('rect',    { x:0,    y:0,    w:'100%', h:0.06, fill:{ type:'solid', color:T.accent } });
    sl.addShape('ellipse', { x:6.8,  y:-1.2, w:5.5, h:5.5, fill:{ type:'solid', color:T.card } });
    sl.addShape('ellipse', { x:-1.8, y:3.2,  w:4.5, h:4.5, fill:{ type:'solid', color:T.card } });

    // Logo
    sl.addShape('roundRect', { x:0.6, y:0.5, w:0.55, h:0.55, rectRadius:0.08, fill:{ type:'solid', color:T.accent } });
    sl.addText('OV',          { x:0.6, y:0.5, w:0.55, h:0.55, fontSize:17, bold:true, color:'FFFFFF', align:'center', valign:'middle' });
    sl.addText('OmniView AI', { x:1.25, y:0.55, w:3.5, h:0.45, fontSize:20, bold:true, color:'F8FAFC' });

    // Tagline badge
    sl.addShape('roundRect', { x:0.6, y:1.22, w:3.2, h:0.34, rectRadius:0.17,
      fill:{ type:'solid', color:T.accent+'28' }, line:{ color:T.accent+'60', width:0.75 } });
    sl.addText(T.tagline, { x:0.6, y:1.22, w:3.2, h:0.34, fontSize:10, bold:true, color:T.accent,
      align:'center', valign:'middle', charSpacing:1.5 });

    // 主標題
    sl.addText('商業提案\n分析報告', { x:0.6, y:1.75, w:8.5, h:1.7, fontSize:FONT_TITLE, bold:true, color:'F8FAFC', lineSpacingMultiple:1.15 });

    // 成功機率徽章
    const sc = r.successProbability;
    const sC = scoreColor(sc);
    sl.addShape('roundRect', { x:0.6, y:3.65, w:2.2, h:0.9, rectRadius:0.14,
      fill:{ type:'solid', color:sC+'22' }, line:{ color:sC, width:1.5 } });
    sl.addText('AI 評估成功機率', { x:0.6, y:3.7, w:2.2, h:0.22, fontSize:FONT_SMALL, color:sC, align:'center' });
    sl.addText(`${sc}%`, { x:0.6, y:3.92, w:2.2, h:0.55, fontSize:30, bold:true, color:sC, align:'center', valign:'middle' });

    // 摘要第一行（內文，shrinkText）
    const coverLine = (r.executiveSummary || '').split('\n')[0] || '';
    sl.addText(coverLine, {
      x:0.6, y:4.7, w:8.8, h:0.38,
      fontSize:FONT_BODY, color:'94A3B8', italic:true,
      shrinkText: true, valign:'middle'
    });

    // Footer
    sl.addShape('rect', { x:0, y:H-0.48, w:'100%', h:0.48, fill:{ type:'solid', color:T.card } });
    sl.addText('由 OmniView AI 360° 虛擬董事會自動生成',
      { x:0.5, y:H-0.36, w:6, h:0.28, fontSize:FONT_XSMALL, color:'475569' });
    sl.addText('1 / 10', { x:W-1.5, y:H-0.36, w:1.0, h:0.28, fontSize:FONT_XSMALL, color:'475569', align:'right' });
  }

  // ════ SLIDE 2 EXECUTIVE SUMMARY ════════════════════
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '執行摘要');

    const lW  = 5.5;
    const HDR = 0.38;

    addCard(sl, T, MX, BODY_T, lW, BODY_H);
    sl.addText('核心觀點', { x:MX+0.2, y:BODY_T+0.1, w:lW-0.4, h:0.28, fontSize:FONT_SECTION, bold:true, color:T.accent });
    sl.addShape('rect', { x:MX+0.2, y:BODY_T+HDR, w:lW-0.4, h:0.01, fill:{ type:'solid', color:T.divider } });

    // 核心觀點內文 — 空間充裕(5.5"寬 × 3.07"高)，lineSpacing 放寬
    sl.addText(
      joinBullets((r.executiveSummary || '').split('\n').filter(Boolean)),
      bodyTextOpts(MX+0.2, BODY_T+HDR+0.08, lW-0.4, BODY_H-HDR-0.18, { lineSpacingMultiple: 1.3 })
    );

    // 右側 KPI 卡
    const kGap = 0.12;
    const kH   = (BODY_H - kGap * 2) / 3;
    const kX   = MX + lW + 0.2;
    const kW   = CW - lW - 0.2;

    const kpis = [
      { label:'市場規模 TAM',  val: r.marketAnalysis.size,       color:T.accent  },
      { label:'CAGR 年均成長', val: r.marketAnalysis.growthRate, color:T.accent2 },
      { label:'損益平衡點',    val: r.breakEvenPoint,            color:T.accent3 },
    ];
    kpis.forEach((k, i) => {
      const ky = BODY_T + i * (kH + kGap);
      addCard(sl, T, kX, ky, kW, kH, k.color+'55');
      sl.addShape('rect', { x:kX, y:ky, w:0.04, h:kH, fill:{ type:'solid', color:k.color } });
      // label
      sl.addText(k.label, { x:kX+0.12, y:ky+0.08, w:kW-0.16, h:0.24, fontSize:FONT_SMALL, color:'94A3B8', shrinkText:true });
      // 數值 — 框高改為 kH-0.38 讓數字有更多空間，fontSize 降到 16 避免截字
      sl.addText(k.val, {
        x:kX+0.12, y:ky+0.34, w:kW-0.16, h:kH-0.38,
        fontSize:16, bold:true, color:k.color, valign:'middle', shrinkText:true
      });
    });

    addFooter(sl, T, 2, TOTAL);
  }

  // ════ SLIDE 3 MARKET OPPORTUNITY ═══════════════════
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '市場機會');

    const KPI_H = 0.9;
    const kGap  = 0.12;
    const kW    = (CW - kGap*2) / 3;
    const kpis = [
      { label:'TAM 市場規模',  val: r.marketAnalysis.size,       color:T.accent  },
      { label:'CAGR 年均成長', val: r.marketAnalysis.growthRate, color:T.accent2 },
      { label:'損益平衡點',    val: r.breakEvenPoint,            color:T.accent3 },
    ];
    kpis.forEach((k, i) => {
      const kx = MX + i * (kW + kGap);
      addCard(sl, T, kx, BODY_T, kW, KPI_H, k.color+'66');
      sl.addShape('rect', { x:kx, y:BODY_T, w:kW, h:0.04, fill:{ type:'solid', color:k.color } });
      sl.addText(k.label, { x:kx+0.12, y:BODY_T+0.06, w:kW-0.24, h:0.22, fontSize:FONT_SMALL, color:'94A3B8', shrinkText:true });
      // 數值降到 15 確保不超框
      sl.addText(k.val, { x:kx+0.12, y:BODY_T+0.3, w:kW-0.24, h:KPI_H-0.36, fontSize:15, bold:true, color:k.color, valign:'middle', shrinkText:true });
    });

    const mTop = BODY_T + KPI_H + 0.18;
    const mH   = BODY_B - mTop;          // ≈ 2.37" 很充裕
    addCard(sl, T, MX, mTop, CW, mH);
    sl.addText('市場洞察', { x:MX+0.2, y:mTop+0.1, w:CW-0.4, h:0.3, fontSize:FONT_SECTION, bold:true, color:T.accent });
    sl.addShape('rect', { x:MX+0.2, y:mTop+0.44, w:CW-0.4, h:0.01, fill:{ type:'solid', color:T.divider } });

    // 空間充裕，字體放大至 11、行距放寬
    sl.addText(
      joinBullets((r.marketAnalysis?.description || '').split('\n').filter(Boolean), '◆  '),
      bodyTextOpts(MX+0.2, mTop+0.52, CW-0.4, mH-0.62, { fontSize:11, lineSpacingMultiple:1.35 })
    );

    addFooter(sl, T, 3, TOTAL);
  }

  // ════ SLIDE 4 FINANCIAL PROJECTIONS ════════════════
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '財務預測');

    const finData = r.financials.slice(0, 4);
    const tHdrH  = 0.38;
    const tRowH  = 0.38;
    const tH     = tHdrH + finData.length * tRowH;

    const fHdr: PptxGenJS.TableCell[] = ['年度','營收','成本','淨利'].map(t => ({
      text: t,
      options: { bold:true, color:'FFFFFF', fill:{ color:T.headerBg }, fontSize:11, align:'center' as const },
    }));
    const fRows: PptxGenJS.TableCell[][] = [fHdr];
    finData.forEach((f, i) => {
      const bg  = i%2===0 ? T.card : T.bg;
      const pft = Number(f.profit)||0;
      fRows.push([
        { text: f.year ?? '—',             options:{ color:'F8FAFC',                        fill:{color:bg}, fontSize:11, align:'center' as const, bold:true } },
        { text: fmt(Number(f.revenue)||0), options:{ color:T.accent,                        fill:{color:bg}, fontSize:11, align:'center' as const, bold:true } },
        { text: fmt(Number(f.costs)||0),   options:{ color:'94A3B8',                        fill:{color:bg}, fontSize:11, align:'center' as const } },
        { text: fmt(pft),                  options:{ color: pft>=0 ? T.accent2 : 'EF4444', fill:{color:bg}, fontSize:11, align:'center' as const, bold:true } },
      ]);
    });
    sl.addTable(fRows, {
      x:MX, y:BODY_T, w:CW,
      border:{ type:'solid', color:T.divider, pt:0.5 },
      colW:[CW*0.25, CW*0.25, CW*0.25, CW*0.25],
      rowH:[tHdrH, ...finData.map(()=>tRowH)],
    });

    const barTop = BODY_T + tH + 0.25;
    const barH   = BODY_B - barTop;

    if (barH >= 1.0) {
      addCard(sl, T, MX, barTop, CW, barH);
      // 圖表區空間充裕，標題字放大
      sl.addText('營收趨勢', { x:MX+0.2, y:barTop+0.1, w:3, h:0.28, fontSize:14, bold:true, color:T.accent });
      sl.addText(`損益平衡：${r.breakEvenPoint}`,
        { x:MX+3.5, y:barTop+0.1, w:CW-3.7, h:0.28, fontSize:FONT_BODY, color:T.accent2, align:'right', shrinkText:true });

      const maxRev = Math.max(...finData.map(f=>Number(f.revenue)||0), 1);
      const rowH2  = (barH - 0.45) / finData.length;

      finData.forEach((f, i) => {
        const rev = Number(f.revenue)||0;
        const bw  = Math.max(rev/maxRev*(CW-2.5), 0.1);
        const by  = barTop + 0.4 + i * rowH2;
        sl.addText(f.year ?? '—', { x:MX+0.2, y:by, w:0.8, h:rowH2, fontSize:FONT_BODY, color:'94A3B8', valign:'middle' });
        sl.addShape('roundRect', { x:MX+1.2, y:by+(rowH2*0.2), w:bw, h:rowH2*0.6, rectRadius:0.04, fill:{ type:'solid', color:T.accent } });
        sl.addText(fmt(rev), { x:MX+1.2+bw+0.05, y:by, w:1.4, h:rowH2, fontSize:FONT_BODY, color:T.accent, bold:true, valign:'middle', shrinkText:true });
      });
    }

    addFooter(sl, T, 4, TOTAL);
  }

  // ════ SLIDE 5 COMPETITIVE LANDSCAPE ════════════════
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '競爭態勢分析');

    const comps = r.competitors.slice(0, 4);
    const hdrH  = 0.42;
    const rowH  = (BODY_H - hdrH) / (comps.length || 1);  // ≈ 0.76" per row — 很充裕

    const cHdr: PptxGenJS.TableCell[] = ['競爭對手','優勢','劣勢'].map(t => ({
      text: t,
      options: { bold:true, color:'FFFFFF', fill:{ color:T.headerBg }, fontSize:12, align:'center' as const },
    }));
    const cRows: PptxGenJS.TableCell[][] = [cHdr];
    comps.forEach((c, i) => {
      const bg = i%2===0 ? T.card : T.bg;
      // 表格行高充裕，字體從 FONT_SMALL(8) 放大到 10
      cRows.push([
        { text: c.name,     options:{ bold:true, color:'F8FAFC', fill:{color:bg}, fontSize:10, align:'center' as const, margin:0.08 } },
        { text: c.strength, options:{ color:T.accent2,           fill:{color:bg}, fontSize:10, margin:0.08 } },
        { text: c.weakness, options:{ color:'F87171',            fill:{color:bg}, fontSize:10, margin:0.08 } },
      ]);
    });
    sl.addTable(cRows, {
      x:MX, y:BODY_T, w:CW,
      border:{ type:'solid', color:T.divider, pt:0.5 },
      colW:[CW*0.2, CW*0.4, CW*0.4],
      rowH:[hdrH, ...comps.map(()=>rowH)],
    });

    addFooter(sl, T, 5, TOTAL);
  }

  // ════ SLIDE 6 STRATEGIC ROADMAP ════════════════════
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '策略路線圖');

    const items = r.roadmap.slice(0, 3);
    const iH    = BODY_H / items.length;   // ≈ 1.15" each

    items.forEach((item, i) => {
      const iY   = BODY_T + i * iH;
      const iBot = iY + iH - 0.08;

      if (i < items.length-1)
        sl.addShape('rect', { x:MX+0.31, y:iY+0.32, w:0.02, h:iH, fill:{ type:'solid', color:T.divider } });

      sl.addShape('ellipse', { x:MX+0.18, y:iY+0.08, w:0.28, h:0.28, fill:{ type:'solid', color:T.accent } });
      sl.addText(`${i+1}`, { x:MX+0.18, y:iY+0.08, w:0.28, h:0.28, fontSize:FONT_SMALL, bold:true, color:'FFFFFF', align:'center', valign:'middle' });

      // phase — 標題行高 0.28 確保不被截
      sl.addText(item.phase, {
        x:MX+0.6, y:iY+0.04, w:4.2, h:0.28,
        fontSize:titleOrBody(item.phase), bold:true, color:T.accent, shrinkText:true, valign:'middle'
      });

      sl.addShape('roundRect', { x:MX+5.1, y:iY+0.06, w:1.2, h:0.24, rectRadius:0.1,
        fill:{ type:'solid', color:T.accent+'25' }, line:{ color:T.accent+'60', width:0.5 } });
      sl.addText(item.timeframe, { x:MX+5.1, y:iY+0.06, w:1.2, h:0.24,
        fontSize:FONT_XSMALL, color:T.accent, align:'center', valign:'middle', shrinkText:true });

      // 內容卡：從 0.36 開始，給下方更多空間
      const cTop = iY + 0.36;
      const cH   = iBot - cTop;           // ≈ 0.71"
      addCard(sl, T, MX+0.6, cTop, CW-0.6, cH);

      const colW = (CW - 0.6 - 0.15) / 2;

      sl.addText('產品', { x:MX+0.72, y:cTop+0.04, w:colW-0.1, h:0.18, fontSize:FONT_XSMALL, bold:true, color:T.accent2 });
      // 內文框緊貼小標下方，maximise 可用高度
      sl.addText(item.product,
        bodyTextOpts(MX+0.72, cTop+0.22, colW-0.1, cH-0.26, { lineSpacingMultiple:1.1 })
      );

      sl.addShape('rect', { x:MX+0.6+colW+0.07, y:cTop+0.04, w:0.01, h:cH-0.08, fill:{ type:'solid', color:T.divider } });

      sl.addText('技術', { x:MX+0.6+colW+0.16, y:cTop+0.04, w:colW-0.1, h:0.18, fontSize:FONT_XSMALL, bold:true, color:T.accent3 });
      sl.addText(item.technology,
        bodyTextOpts(MX+0.6+colW+0.16, cTop+0.22, colW-0.1, cH-0.26, { lineSpacingMultiple:1.1 })
      );
    });

    addFooter(sl, T, 6, TOTAL);
  }

  // ════ SLIDE 7 RISK ASSESSMENT ══════════════════════
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '風險評估');

    const risks = r.risks.slice(0, 4);
    const COLS  = 2;
    const gap   = 0.16;
    const cW    = (CW - gap) / COLS;
    const cH    = (BODY_H - gap) / 2;     // ≈ 1.645" each

    risks.forEach((risk, i) => {
      const col = i % COLS;
      const row = Math.floor(i / COLS);
      const cx  = MX + col*(cW+gap);
      const cy  = BODY_T + row*(cH+gap);

      const ic    = risk.impact==='High' ? 'EF4444' : risk.impact==='Medium' ? 'FBBF24' : '10B981';
      const label = risk.impact==='High' ? '高風險' : risk.impact==='Medium' ? '中風險' : '低風險';

      addCard(sl, T, cx, cy, cW, cH, ic+'55');
      sl.addShape('rect', { x:cx, y:cy, w:cW, h:0.04, fill:{ type:'solid', color:ic } });

      // 風險標題 — h 加大到 0.32 避免多行被截
      sl.addText(risk.risk, {
        x:cx+0.12, y:cy+0.08, w:cW-0.84, h:0.32,
        fontSize:titleOrBody(risk.risk), bold:true, color:'F8FAFC', valign:'top', shrinkText:true
      });

      sl.addShape('roundRect', { x:cx+cW-0.72, y:cy+0.1, w:0.62, h:0.24, rectRadius:0.1,
        fill:{ type:'solid', color:ic+'30' }, line:{ color:ic, width:0.5 } });
      sl.addText(label, { x:cx+cW-0.72, y:cy+0.1, w:0.62, h:0.24,
        fontSize:FONT_XSMALL, bold:true, color:ic, align:'center', valign:'middle' });

      // 因應對策 — 起始 y 從 0.44 改為 0.42，h 撐滿剩餘空間
      sl.addText([
        { text:'因應對策：', options:{ fontSize:FONT_SMALL, color:'94A3B8', bold:true, breakLine:true } },
        { text: risk.mitigation, options:{ fontSize:FONT_SMALL, color:'CBD5E1' } },
      ], { x:cx+0.12, y:cy+0.44, w:cW-0.24, h:cH-0.52, valign:'top', shrinkText:true, lineSpacingMultiple:1.2 });
    });

    addFooter(sl, T, 7, TOTAL);
  }

  // ════ SLIDE 8 STAKEHOLDER BOARD ════════════════════
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, 'AI 虛擬董事會');

    const personas = r.personaEvaluations.slice(0, 3);
    const gap = 0.16;
    const cW  = (CW - gap * 2) / 3;      // ≈ 2.89"
    const cH  = BODY_H;                   // 3.45"

    personas.forEach((p, i) => {
      const cx = MX + i*(cW+gap);
      const cy = BODY_T;
      const sc = Number(p.score)||0;
      const sC = scoreColor(sc);

      addCard(sl, T, cx, cy, cW, cH);
      sl.addShape('rect', { x:cx, y:cy, w:cW, h:0.04, fill:{ type:'solid', color:sC } });

      // role — 空間有限，固定 FONT_SMALL 確保不超框
      sl.addText(p.role, {
        x:cx+0.1, y:cy+0.1, w:cW-0.2, h:0.28,
        fontSize:FONT_SMALL, bold:true, color:'F8FAFC', align:'center', shrinkText:true, valign:'middle'
      });
      // 分數大字
      sl.addText(`${sc}`, { x:cx+0.1, y:cy+0.4, w:cW-0.2, h:0.42, fontSize:26, bold:true, color:sC, align:'center', valign:'middle' });

      // keyQuote — h 從 0.55 加大到 0.7，並縮小 fontSize 到 FONT_XSMALL
      sl.addText(`"${p.keyQuote}"`, {
        x:cx+0.1, y:cy+0.84, w:cW-0.2, h:0.7,
        fontSize:FONT_XSMALL, italic:true, color:'94A3B8', align:'center', valign:'top', shrinkText:true
      });

      sl.addShape('rect', { x:cx+0.15, y:cy+1.6, w:cW-0.3, h:0.01, fill:{ type:'solid', color:T.divider } });

      // 顧慮點 — 起始 y 從 1.44 調整到 1.66，h 撐到底
      sl.addText([
        { text:'顧慮點：\n', options:{ fontSize:FONT_SMALL, color:'FBBF24', bold:true } },
        { text: p.concern,   options:{ fontSize:FONT_SMALL, color:'CBD5E1' } },
      ], { x:cx+0.1, y:cy+1.66, w:cW-0.2, h:cH-1.74, valign:'top', shrinkText:true, lineSpacingMultiple:1.2 });
    });

    addFooter(sl, T, 8, TOTAL);
  }

  // ════ SLIDE 9 FINAL VERDICT ════════════════════════
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '最終裁決');

    const verdicts = [
      { title:'激進觀點', text: r.finalVerdicts?.aggressive || '',   color:'F97316'  },
      { title:'平衡觀點', text: r.finalVerdicts?.balanced || '',     color:T.accent  },
      { title:'保守觀點', text: r.finalVerdicts?.conservative || '', color:T.accent3 },
    ];
    const gap   = 0.16;
    const vW    = (CW - gap*2) / 3;      // ≈ 2.89"
    const vH    = BODY_H;
    const HDR_H = 0.48;                  // 縮小 header 區，讓內文得到更多高度

    verdicts.forEach((v, i) => {
      const vx     = MX + i*(vW+gap);
      const availH = vH - HDR_H - 0.06;  // ≈ 2.91"

      addCard(sl, T, vx, BODY_T, vW, vH, v.color+'55');
      sl.addShape('rect', { x:vx, y:BODY_T, w:vW, h:0.04, fill:{ type:'solid', color:v.color } });
      sl.addText(v.title, { x:vx+0.14, y:BODY_T+0.08, w:vW-0.28, h:0.28, fontSize:FONT_SECTION, bold:true, color:v.color });
      sl.addShape('rect', { x:vx+0.14, y:BODY_T+0.4, w:0.7, h:0.02, fill:{ type:'solid', color:v.color } });

      // 內文框高度更充裕，字可稍微放大到 10
      sl.addText(
        joinBullets((v.text || '').split('\n').filter(Boolean)),
        bodyTextOpts(vx+0.14, BODY_T+HDR_H, vW-0.28, availH, {
          fontSize: 10, lineSpacingMultiple: 1.2, margin: 0.02
        })
      );
    });

    addFooter(sl, T, 9, TOTAL);
  }

  // ════ SLIDE 10 CALL TO ACTION ══════════════════════
  {
    const sl = pptx.addSlide();
    sl.background = { fill: T.bg };
    sl.addShape('rect',    { x:0,   y:0,   w:'100%', h:0.06, fill:{ type:'solid', color:T.accent } });
    sl.addShape('ellipse', { x:2.75, y:0.7, w:4.7, h:4.7, fill:{ type:'solid', color:T.card } });

    sl.addText('下一步行動', { x:0.5, y:1.1, w:9.0, h:1.0, fontSize:FONT_TITLE, bold:true, color:'F8FAFC', align:'center', valign:'middle' });
    sl.addText(T.tagline,   { x:0.5, y:2.25, w:9.0, h:0.4, fontSize:FONT_SUBTITLE, color:T.accent, align:'center', charSpacing:2.5, bold:true });

    const sc = r.successProbability;
    const sC = scoreColor(sc);
    addCard(sl, T, 3.3, 2.85, 3.4, 1.7, sC);
    sl.addText('AI 評估成功機率', { x:3.3, y:2.96, w:3.4, h:0.26, fontSize:FONT_BODY, color:'94A3B8', align:'center' });
    // 成功機率數字空間充裕，放大 fontSize
    sl.addText(`${sc}%`, { x:3.3, y:3.26, w:3.4, h:1.0, fontSize:36, bold:true, color:sC, align:'center', valign:'middle' });

    // Footer
    sl.addShape('rect', { x:0, y:H-0.48, w:'100%', h:0.48, fill:{ type:'solid', color:T.card } });
    sl.addText('此報告由 OmniView AI 360° 虛擬董事會自動生成',
      { x:0.5, y:H-0.36, w:8, h:0.28, fontSize:FONT_XSMALL, color:'475569', align:'center' });
    sl.addText('10 / 10', { x:W-1.5, y:H-0.36, w:1.0, h:0.28, fontSize:FONT_XSMALL, color:'475569', align:'right' });
  }

  await pptx.writeFile({ fileName: 'OmniView_商業提案報告.pptx' });
};

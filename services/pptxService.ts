import PptxGenJS from 'pptxgenjs';
import { AnalysisResult } from '../types';

// ── 版面常數 (16:9 = 10" × 7.5") ─────────────────────────
const W      = 10;
const MX     = 0.85;          // 左右邊距
const CW     = W - MX * 2;   // 內容寬度 = 9.1"
const BODY_T = 1.5;          // 內容區頂部
const BODY_B = 7.1;           // 內容區底部（安全線，留 0.4" 給頁碼）
const BODY_H = BODY_B - BODY_T; // = 5.6"

// ── 字體大小常數 ──────────────────────────────────────────
const FONT_TITLE = 44;
const FONT_SUBTITLE = 24;
const FONT_BODY = 18;

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
  sl.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ type:'solid', color:T.accent } });
};
const addHeader = (sl: PptxGenJS.Slide, T: Theme, title: string) => {
  sl.addText(title, { x:MX, y:0.1, w:8, h:0.48, fontSize:FONT_SUBTITLE, bold:true, color:'F8FAFC', fontFace:'Arial' });
  sl.addShape('rect', { x:MX, y:0.6, w:0.85, h:0.03, fill:{ type:'solid', color:T.accent } });
};
const addPageNum = (sl: PptxGenJS.Slide, n: number, total: number, T: Theme) =>
  sl.addText(`${n} / ${total}`, { x:8.6, y:7.22, w:0.8, h:0.2, fontSize:FONT_BODY, color:T.divider, align:'right' });
const addCard = (sl: PptxGenJS.Slide, T: Theme, x:number, y:number, w:number, h:number, border?: string) =>
  sl.addShape('roundRect', { x, y, w, h, rectRadius:0.1,
    fill:{ type:'solid', color:T.card }, line:{ color:border||T.divider, width:0.75 } });
const scoreColor = (s: number) =>
  s >= 80 ? '10B981' : s >= 60 ? '3B82F6' : s >= 40 ? 'FBBF24' : 'EF4444';

// ── 主函式 ────────────────────────────────────────────────

export const generatePptx = async (result: AnalysisResult): Promise<void> => {
  const r = result;
  const T = detectTheme(r.executiveSummary, r.marketAnalysis.description);
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
    sl.addText('OV',          { x:0.5,  y:0.3,  w:0.55, h:0.55, fontSize:FONT_BODY, bold:true, color:'FFFFFF', align:'center', valign:'middle' });
    sl.addText('OmniView AI', { x:1.15, y:0.35, w:3,    h:0.45, fontSize:FONT_SUBTITLE, bold:true, color:'F8FAFC' });

    sl.addShape('roundRect', { x:0.5, y:1.05, w:3.0, h:0.32, rectRadius:0.16,
      fill:{ type:'solid', color:T.accent+'28' }, line:{ color:T.accent+'60', width:0.75 } });
    sl.addText(T.tagline, { x:0.5, y:1.05, w:3.0, h:0.32, fontSize:FONT_BODY, bold:true, color:T.accent,
      align:'center', valign:'middle', charSpacing:1.5 });

    sl.addText('商業提案\n分析報告', { x:0.5, y:1.55, w:8.5, h:1.7, fontSize:FONT_TITLE, bold:true, color:'F8FAFC', lineSpacingMultiple:1.1 });

    const sc = r.successProbability;
    const sC = scoreColor(sc);
    sl.addShape('roundRect', { x:0.5, y:3.5, w:2.3, h:0.85, rectRadius:0.12,
      fill:{ type:'solid', color:sC+'22' }, line:{ color:sC, width:1.5 } });
    sl.addText('AI 評估成功機率', { x:0.5, y:3.56, w:2.3, h:0.26, fontSize:FONT_BODY, color:sC, align:'center' });
    sl.addText(`${sc}%`, { x:0.5, y:3.82, w:2.3, h:0.48, fontSize:FONT_TITLE, bold:true, color:sC, align:'center', valign:'middle' });

    // 封面摘要：只取第一條
    const coverLine = r.executiveSummary.split('\n')[0] || '';
    sl.addText(coverLine, { x:0.5, y:4.55, w:9.0, h:0.38, fontSize:FONT_BODY, color:'94A3B8', italic:true, margin:0.02, fit: 'shrink' });

    sl.addShape('rect', { x:0, y:6.9, w:'100%', h:0.38, fill:{ type:'solid', color:T.card } });
    sl.addText('由 OmniView AI 360° 虛擬董事會自動生成',
      { x:0.5, y:6.93, w:6, h:0.28, fontSize:FONT_BODY, color:'475569' });
    addPageNum(sl, 1, TOTAL, T);
  }

  // ════ SLIDE 2 EXECUTIVE SUMMARY ════════════════════
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '執行摘要');

    // 左：核心觀點卡
    const lW   = 5.35;
    const HDR  = 0.48;   // 卡片內 header 高度
    const PAD  = 0.12;
    addCard(sl, T, MX, BODY_T, lW, BODY_H);
    sl.addText('核心觀點', { x:MX+0.2, y:BODY_T+PAD, w:lW-0.4, h:HDR-PAD, fontSize:FONT_SUBTITLE, bold:true, color:T.accent });

    const bAvailH = BODY_H - HDR - 0.08;
    sl.addText(joinBullets(r.executiveSummary.split('\n').filter(Boolean)), {
      x:MX+0.2, y:BODY_T+HDR, w:lW-0.35, h:bAvailH,
      fontSize:FONT_BODY, color:'CBD5E1', lineSpacingMultiple:1.2, valign:'top', margin:0.02,
      fit: 'shrink'
    });

    // 右：KPI 卡 ×3
    const kGap = 0.12;
    const kH   = (BODY_H - kGap * 2) / 3;
    const kX   = MX + lW + 0.1;
    const kW   = W - MX - kX;
    const kpis = [
      { label:'市場規模 TAM', val: r.marketAnalysis.size,       color:T.accent  },
      { label:'CAGR 年均成長', val: r.marketAnalysis.growthRate, color:T.accent2 },
      { label:'損益平衡點',    val: r.breakEvenPoint,            color:T.accent3 },
    ];
    kpis.forEach((k, i) => {
      const ky = BODY_T + i * (kH + kGap);
      addCard(sl, T, kX, ky, kW, kH, k.color+'55');
      sl.addShape('rect', { x:kX, y:ky, w:0.04, h:kH, fill:{ type:'solid', color:k.color } });
      sl.addText(k.label, { x:kX+0.14, y:ky+0.12, w:kW-0.2, h:0.26, fontSize:FONT_BODY,  color:'94A3B8' });
      sl.addText(k.val,   { x:kX+0.14, y:ky+0.44, w:kW-0.2, h:kH-0.54, fontSize:FONT_SUBTITLE, bold:true, color:k.color, valign:'top', fit: 'shrink' });
    });

    addPageNum(sl, 2, TOTAL, T);
  }

  // ════ SLIDE 3 MARKET OPPORTUNITY ═══════════════════
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '市場機會');

    const KPI_H = 1.2;
    const kW    = (CW - 0.2) / 3;
    const kpis = [
      { label:'TAM 市場規模', val: r.marketAnalysis.size,       color:T.accent  },
      { label:'CAGR 年均成長', val: r.marketAnalysis.growthRate, color:T.accent2 },
      { label:'損益平衡點',    val: r.breakEvenPoint,            color:T.accent3 },
    ];
    kpis.forEach((k, i) => {
      const kx = MX + i * (kW + 0.1);
      addCard(sl, T, kx, BODY_T, kW, KPI_H, k.color+'66');
      sl.addShape('rect', { x:kx, y:BODY_T, w:kW, h:0.04, fill:{ type:'solid', color:k.color } });
      sl.addText(k.label, { x:kx+0.12, y:BODY_T+0.1,  w:kW-0.24, h:0.28, fontSize:FONT_BODY, color:'94A3B8' });
      sl.addText(k.val,   { x:kx+0.12, y:BODY_T+0.44, w:kW-0.24, h:KPI_H-0.54, fontSize:FONT_SUBTITLE, bold:true, color:k.color, valign:'top', fit: 'shrink' });
    });

    const mTop   = BODY_T + KPI_H + 0.15;
    const mH     = BODY_B - mTop;
    const mHDR   = 0.44;
    const mAvail = mH - mHDR - 0.08;
    addCard(sl, T, MX, mTop, CW, mH);
    sl.addText('市場洞察', { x:MX+0.2, y:mTop+0.1, w:CW-0.4, h:0.28, fontSize:FONT_SUBTITLE, bold:true, color:T.accent });

    sl.addText(joinBullets(r.marketAnalysis.description.split('\n').filter(Boolean), '◆  '), {
      x:MX+0.2, y:mTop+mHDR, w:CW-0.4, h:mAvail,
      fontSize:FONT_BODY, color:'CBD5E1', lineSpacingMultiple:1.2, valign:'top', margin:0.02,
      fit: 'shrink'
    });

    addPageNum(sl, 3, TOTAL, T);
  }

  // ════ SLIDE 4 FINANCIAL PROJECTIONS ════════════════
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '財務預測');

    const finData = r.financials.slice(0, 4);
    const tHdrH  = 0.42;
    const tRowH  = 0.46;
    const tH     = tHdrH + finData.length * tRowH;

    const fHdr: PptxGenJS.TableCell[] = ['年度','營收','成本','淨利'].map(t => ({
      text: t,
      options: { bold:true, color:'FFFFFF', fill:{ color:T.headerBg }, fontSize:FONT_BODY, align:'center' as const },
    }));
    const fRows: PptxGenJS.TableCell[][] = [fHdr];
    finData.forEach((f, i) => {
      const bg  = i%2===0 ? T.card : T.bg;
      const pft = Number(f.profit)||0;
      fRows.push([
        { text: f.year ?? '—',              options:{ color:'F8FAFC',                        fill:{color:bg}, fontSize:FONT_BODY, align:'center' as const, bold:true } },
        { text: fmt(Number(f.revenue)||0),  options:{ color:T.accent,                        fill:{color:bg}, fontSize:FONT_BODY, align:'center' as const, bold:true } },
        { text: fmt(Number(f.costs)||0),    options:{ color:'94A3B8',                        fill:{color:bg}, fontSize:FONT_BODY, align:'center' as const } },
        { text: fmt(pft),                   options:{ color: pft>=0 ? T.accent2 : 'EF4444', fill:{color:bg}, fontSize:FONT_BODY, align:'center' as const, bold:true } },
      ]);
    });
    sl.addTable(fRows, {
      x:MX, y:BODY_T, w:CW,
      border:{ type:'solid', color:T.divider, pt:0.5 },
      colW:[2.275, 2.275, 2.275, 2.275],
      rowH:[tHdrH, ...finData.map(()=>tRowH)],
    });

    const barTop = BODY_T + tH + 0.2;
    const barH   = BODY_B - barTop;
    if (barH >= 0.9) {
      addCard(sl, T, MX, barTop, CW, barH);
      sl.addText('營收趨勢', { x:MX+0.2, y:barTop+0.1, w:3, h:0.24, fontSize:FONT_SUBTITLE, bold:true, color:T.accent });
      sl.addText(`損益平衡：${r.breakEvenPoint}`,
        { x:4.5, y:barTop+0.1, w:4.75, h:0.24, fontSize:FONT_BODY, color:T.accent2, align:'right' });

      const maxRev  = Math.max(...finData.map(f=>Number(f.revenue)||0), 1);
      const barSX   = 1.5;
      const barMaxW = 6.8;
      const rowH2   = Math.min(0.36, (barH - 0.42) / finData.length);
      finData.forEach((f, i) => {
        const rev = Number(f.revenue)||0;
        const bw  = Math.max(rev/maxRev*barMaxW, 0.12);
        const by  = barTop + 0.38 + i * rowH2;
        if (by + rowH2 > BODY_B - 0.05) return;
        sl.addText(f.year ?? '—', { x:0.55, y:by, w:0.85, h:rowH2, fontSize:FONT_BODY, color:'94A3B8', valign:'middle' });
        sl.addShape('roundRect', { x:barSX, y:by+0.04, w:bw, h:rowH2-0.08, rectRadius:0.04, fill:{ type:'solid', color:T.accent } });
        const lx = barSX+bw+0.08;
        if (lx+0.85 <= MX+CW) {
          sl.addText(fmt(rev), { x:lx, y:by, w:0.85, h:rowH2, fontSize:FONT_BODY, color:T.accent, bold:true, valign:'middle' });
        } else {
          sl.addText(fmt(rev), { x:barSX+bw-0.95, y:by, w:0.85, h:rowH2, fontSize:FONT_BODY, color:'FFFFFF', bold:true, valign:'middle' });
        }
      });
    }
    addPageNum(sl, 4, TOTAL, T);
  }

  // ════ SLIDE 5 COMPETITIVE LANDSCAPE ════════════════
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '競爭態勢分析');

    const comps  = r.competitors.slice(0, 5);
    const hdrH   = 0.40;
    // 每列固定高度，讓 table 撐到 BODY_B
    const rowH   = (BODY_B - BODY_T - hdrH) / comps.length;

    const cHdr: PptxGenJS.TableCell[] = ['競爭對手','優勢','劣勢'].map(t => ({
      text: t,
      options: { bold:true, color:'FFFFFF', fill:{ color:T.headerBg }, fontSize:FONT_BODY, align:'center' as const },
    }));
    const cRows: PptxGenJS.TableCell[][] = [cHdr];
    
    comps.forEach((c, i) => {
      const bg = i%2===0 ? T.card : T.bg;
      cRows.push([
        { text: c.name,    options:{ bold:true, color:'F8FAFC', fill:{color:bg}, fontSize:FONT_BODY, align:'center' as const, margin:0.02 } },
        { text: c.strength,  options:{ color:T.accent2,           fill:{color:bg}, fontSize:FONT_BODY, margin:0.02 } },
        { text: c.weakness,   options:{ color:'F87171',            fill:{color:bg}, fontSize:FONT_BODY, margin:0.02 } },
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
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '策略路線圖');

    const items  = r.roadmap.slice(0, 4);
    const iH     = BODY_H / items.length;   // 每個 roadmap item 的高度份額

    items.forEach((item, i) => {
      const iY   = BODY_T + i * iH;
      const iBot = Math.min(iY + iH, BODY_B);
      const safeH = iBot - iY;

      // 連接線
      if (i < items.length-1)
        sl.addShape('rect', { x:0.61, y:iY+0.3, w:0.025, h:safeH-0.3,
          fill:{ type:'solid', color:T.divider } });

      // 序號圓點
      sl.addShape('ellipse', { x:0.49, y:iY+0.1, w:0.25, h:0.25, fill:{ type:'solid', color:T.accent } });
      sl.addText(`${i+1}`, { x:0.49, y:iY+0.1, w:0.25, h:0.25,
        fontSize:FONT_BODY, bold:true, color:'FFFFFF', align:'center', valign:'middle' });

      // 標題 + 時間標籤（一行高 0.32，y 限制在 iBot 以內）
      const titleY = iY + 0.05;
      if (titleY + 0.32 <= iBot) {
        sl.addText(item.phase, { x:0.88, y:titleY, w:3.3, h:0.3, fontSize:FONT_SUBTITLE, bold:true, color:T.accent });
        sl.addShape('roundRect', { x:4.35, y:titleY+0.02, w:1.45, h:0.25, rectRadius:0.12,
          fill:{ type:'solid', color:T.accent+'25' }, line:{ color:T.accent+'60', width:0.5 } });
        sl.addText(item.timeframe, { x:4.35, y:titleY+0.02, w:1.45, h:0.25,
          fontSize:FONT_BODY, color:T.accent, align:'center', valign:'middle' });
      }

      // 內容卡
      const cTop = iY + 0.38;
      const cH   = iBot - cTop - 0.04;
      if (cH < 0.28) return;
      addCard(sl, T, 0.88, cTop, CW-0.45, cH);

      const labelH  = 0.22;
      const textH   = cH - labelH - 0.1;
      const half    = (CW - 0.45 - 0.1) / 2;
      if (textH <= 0) return;

      sl.addText('產品', { x:1.06, y:cTop+0.06, w:half-0.2, h:labelH, fontSize:FONT_BODY, bold:true, color:T.accent2, margin:0.02 });
      sl.addText(item.product, { x:1.06, y:cTop+0.06+labelH, w:half-0.2, h:textH,
        fontSize:FONT_BODY, color:'CBD5E1', valign:'top', lineSpacingMultiple:1.2, margin:0.02, fit: 'shrink' });

      sl.addShape('rect', { x:0.88+half+0.05, y:cTop+0.06, w:0.02, h:cH-0.12,
        fill:{ type:'solid', color:T.divider } });

      const rx = 0.88 + half + 0.12;
      sl.addText('技術', { x:rx, y:cTop+0.06, w:half-0.2, h:labelH, fontSize:FONT_BODY, bold:true, color:T.accent3, margin:0.02 });
      sl.addText(item.technology, { x:rx, y:cTop+0.06+labelH, w:half-0.2, h:textH,
        fontSize:FONT_BODY, color:'CBD5E1', valign:'top', lineSpacingMultiple:1.2, margin:0.02, fit: 'shrink' });
    });
    addPageNum(sl, 6, TOTAL, T);
  }

  // ════ SLIDE 7 RISK ASSESSMENT ══════════════════════
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '風險評估');

    const risks = r.risks.slice(0, 6);
    const COLS  = 2;
    const nRows = Math.ceil(risks.length / COLS);
    const gap   = 0.1;
    const cW    = (CW - gap) / COLS;
    const cH    = (BODY_H - gap*(nRows-1)) / nRows;

    risks.forEach((risk, i) => {
      const col = i%COLS, row = Math.floor(i/COLS);
      const cx  = MX + col*(cW+gap);
      const cy  = BODY_T + row*(cH+gap);
      if (cy+cH > BODY_B+0.02) return;

      const ic    = risk.impact==='High' ? 'EF4444' : risk.impact==='Medium' ? 'FBBF24' : '10B981';
      const label = risk.impact==='High' ? '高' : risk.impact==='Medium' ? '中' : '低';

      addCard(sl, T, cx, cy, cW, cH, ic+'55');
      sl.addShape('rect', { x:cx, y:cy, w:cW, h:0.04, fill:{ type:'solid', color:ic } });

      sl.addShape('roundRect', { x:cx+cW-0.78, y:cy+0.1, w:0.64, h:0.24, rectRadius:0.12,
        fill:{ type:'solid', color:ic+'30' }, line:{ color:ic, width:0.5 } });
      sl.addText(label, { x:cx+cW-0.78, y:cy+0.1, w:0.64, h:0.24,
        fontSize:FONT_BODY, bold:true, color:ic, align:'center', valign:'middle' });

      sl.addText(risk.risk, { x:cx+0.14, y:cy+0.1, w:cW-0.98, h:0.28, fontSize:FONT_SUBTITLE, bold:true, color:'F8FAFC', fit: 'shrink' });

      // 因應措施文字
      const mitTop = cy + 0.42;
      const mitH   = (cy + cH) - mitTop - 0.06;
      if (mitH > 0.14) {
        sl.addText([
          { text:'因應：', options:{ fontSize:FONT_BODY, color:'64748B', bold:true } },
          { text: risk.mitigation.replace(/\n/g, ' '), options:{ fontSize:FONT_BODY, color:'CBD5E1' } },
        ], { x:cx+0.14, y:mitTop, w:cW-0.26, h:mitH, valign:'top', margin:0.02, fit: 'shrink' });
      }
    });
    addPageNum(sl, 7, TOTAL, T);
  }

  // ════ SLIDE 8 STAKEHOLDER BOARD ════════════════════
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, 'AI 虛擬董事會');

    const personas = r.personaEvaluations.slice(0, 5);
    const COLS  = 3;
    const nRows = Math.ceil(personas.length / COLS);
    const gap   = 0.1;
    const cW    = (CW - gap*(COLS-1)) / COLS;
    const cH    = (BODY_H - gap*(nRows-1)) / nRows;

    personas.forEach((p, i) => {
      const col = i%COLS, row = Math.floor(i/COLS);
      const cx  = MX + col*(cW+gap);
      const cy  = BODY_T + row*(cH+gap);
      if (cy+cH > BODY_B+0.02) return;

      const sc = Number(p.score)||0;
      const sC = scoreColor(sc);

      addCard(sl, T, cx, cy, cW, cH);
      sl.addShape('rect', { x:cx, y:cy, w:cW, h:0.04, fill:{ type:'solid', color:sC } });

      sl.addText(p.role, { x:cx+0.12, y:cy+0.1, w:cW-0.85, h:0.28, fontSize:FONT_SUBTITLE, bold:true, color:'F8FAFC', fit: 'shrink' });
      sl.addText(`${sc}`, { x:cx+cW-0.82, y:cy+0.08, w:0.68, h:0.3, fontSize:FONT_SUBTITLE, bold:true, color:sC, align:'right' });

      const barFill = (cW-0.24)*(sc/100);
      sl.addShape('roundRect', { x:cx+0.12, y:cy+0.42, w:cW-0.24, h:0.055, rectRadius:0.028,
        fill:{ type:'solid', color:T.divider } });
      sl.addShape('roundRect', { x:cx+0.12, y:cy+0.42, w:Math.max(barFill,0.05), h:0.055, rectRadius:0.028,
        fill:{ type:'solid', color:sC } });

      // 金句：固定高 0.44
      const quoteH = 0.44;
      const quoteY = cy + 0.54;
      sl.addText(`"${p.keyQuote.replace(/\n/g, ' ')}"`, { x:cx+0.12, y:quoteY, w:cW-0.24, h:quoteH,
        fontSize:FONT_BODY, italic:true, color:'94A3B8', valign:'top', margin:0.02, fit: 'shrink' });

      // concern：剩餘空間
      const conTop = quoteY + quoteH + 0.04;
      const conH   = (cy+cH) - conTop - 0.06;
      if (conH > 0.12) {
        sl.addText([
          { text:'! ', options:{ fontSize:FONT_BODY, color:'FBBF24', bold:true } },
          { text: p.concern.replace(/\n/g, ' '), options:{ fontSize:FONT_BODY, color:'CBD5E1' } },
        ], { x:cx+0.12, y:conTop, w:cW-0.24, h:conH, valign:'top', margin:0.02, fit: 'shrink' });
      }
    });
    addPageNum(sl, 8, TOTAL, T);
  }

  // ════ SLIDE 9 FINAL VERDICTS ═══════════════════════
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '最終裁決');

    const verdicts = [
      { title:'激進觀點', text: r.finalVerdicts.aggressive,   color:'F97316'  },
      { title:'平衡觀點', text: r.finalVerdicts.balanced,     color:T.accent  },
      { title:'保守觀點', text: r.finalVerdicts.conservative, color:T.accent3 },
    ];
    const gap = 0.15;
    const vW  = (CW - gap*2) / 3;
    const vH  = BODY_B - BODY_T;

    verdicts.forEach((v, i) => {
      const vx      = MX + i*(vW+gap);
      const HDR_H   = 0.52;   // 標題區高度
      const availH  = vH - HDR_H - 0.08;

      addCard(sl, T, vx, BODY_T, vW, vH, v.color+'55');
      sl.addShape('rect', { x:vx, y:BODY_T, w:vW, h:0.04, fill:{ type:'solid', color:v.color } });
      sl.addText(v.title, { x:vx+0.12, y:BODY_T+0.1, w:vW-0.24, h:0.3, fontSize:FONT_SUBTITLE, bold:true, color:v.color });
      sl.addShape('rect', { x:vx+0.12, y:BODY_T+0.43, w:0.75, h:0.025, fill:{ type:'solid', color:v.color } });

      sl.addText(joinBullets(v.text.split('\n').filter(Boolean)), {
        x:vx+0.12, y:BODY_T+HDR_H, w:vW-0.24, h:availH,
        fontSize:FONT_BODY, color:'CBD5E1', lineSpacingMultiple:1.2, valign:'top', margin:0.02, fit: 'shrink'
      });
    });
    addPageNum(sl, 9, TOTAL, T);
  }

  // ════ SLIDE 10 CALL TO ACTION ══════════════════════
  {
    const sl = pptx.addSlide();
    sl.background = { fill: T.bg };
    sl.addShape('rect',    { x:0,   y:0,   w:'100%', h:0.05, fill:{ type:'solid', color:T.accent } });
    sl.addShape('ellipse', { x:2.8, y:0.8, w:4.5,   h:4.5,  fill:{ type:'solid', color:T.card } });

    sl.addText('下一步行動', { x:0.5, y:1.5, w:9.0, h:0.9,  fontSize:FONT_TITLE, bold:true, color:'F8FAFC', align:'center' });
    sl.addText(T.tagline,   { x:0.5, y:2.5, w:9.0, h:0.36, fontSize:FONT_SUBTITLE, color:T.accent, align:'center', charSpacing:2.5, bold:true });

    const sc = r.successProbability;
    const sC = scoreColor(sc);
    addCard(sl, T, 3.5, 3.1, 3.0, 1.5, sC);
    sl.addText('AI 評估成功機率', { x:3.5, y:3.22, w:3.0, h:0.28, fontSize:FONT_BODY, color:'94A3B8', align:'center' });
    sl.addText(`${sc}%`,          { x:3.5, y:3.54, w:3.0, h:0.85, fontSize:FONT_TITLE, bold:true, color:sC, align:'center' });

    sl.addText('此報告由 OmniView AI 360° 虛擬董事會自動生成',
      { x:1, y:6.55, w:8, h:0.28, fontSize:FONT_BODY, color:'475569', align:'center' });
    addPageNum(sl, 10, TOTAL, T);
  }

  // ════ SLIDE 11 CONTINUE TO ITERATE ═════════════════
  {
    const sl = pptx.addSlide();
    addBg(sl, T); addHeader(sl, T, '持續迭代');

    const cW = CW;
    const cH = BODY_H;

    addCard(sl, T, MX, BODY_T, cW, cH);
    sl.addText('持續迭代建議', { x:MX+0.2, y:BODY_T+0.1, w:cW-0.4, h:0.28, fontSize:FONT_SUBTITLE, bold:true, color:T.accent });

    sl.addText(joinBullets(r.continueToIterate.split('\n').filter(Boolean), '◆  '), {
      x:MX+0.2, y:BODY_T+0.44, w:cW-0.4, h:cH-0.52,
      fontSize:FONT_BODY, color:'CBD5E1', lineSpacingMultiple:1.2, valign:'top', margin:0.02,
      fit: 'shrink'
    });

    addPageNum(sl, 11, TOTAL + 1, T);
  }

  await pptx.writeFile({ fileName: 'OmniView_商業提案報告.pptx' });
};

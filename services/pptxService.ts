import PptxGenJS from 'pptxgenjs';
import { AnalysisResult } from '../types';

const DARK_BG = '0F172A';
const CARD_BG = '1E293B';
const ACCENT_BLUE = '3B82F6';
const ACCENT_GREEN = '10B981';
const ACCENT_PURPLE = '7C3AED';
const ACCENT_RED = 'EF4444';
const ACCENT_YELLOW = 'FBBF24';
const TEXT_WHITE = 'F8FAFC';
const TEXT_MUTED = '94A3B8';
const TEXT_DIM = '64748B';

const addGradientBg = (slide: PptxGenJS.Slide, color1 = DARK_BG, color2 = '1E3A5F') => {
  slide.background = { fill: DARK_BG };
  // top accent bar
  slide.addShape('rect', { x: 0, y: 0, w: '100%', h: 0.06, fill: { type: 'solid', color: ACCENT_BLUE } });
};

const addSlideNumber = (slide: PptxGenJS.Slide, num: number, total: number) => {
  slide.addText(`${num} / ${total}`, {
    x: 8.8, y: 7.0, w: 1.2, h: 0.35,
    fontSize: 9, color: TEXT_DIM, align: 'right',
  });
};

const addSectionTitle = (slide: PptxGenJS.Slide, icon: string, title: string) => {
  slide.addText(`${icon}  ${title}`, {
    x: 0.5, y: 0.25, w: 9, h: 0.55,
    fontSize: 22, bold: true, color: TEXT_WHITE,
  });
  slide.addShape('rect', { x: 0.5, y: 0.82, w: 1.2, h: 0.04, fill: { type: 'solid', color: ACCENT_BLUE } });
};

export const generatePptx = async (result: AnalysisResult): Promise<void> => {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';
  pptx.author = 'OmniView AI';
  pptx.title = 'OmniView 商業提案分析報告';

  const TOTAL = 10;
  let slideNum = 0;

  // ========== SLIDE 1: COVER ==========
  {
    slideNum++;
    const slide = pptx.addSlide();
    slide.background = { fill: DARK_BG };

    // Decorative shapes
    slide.addShape('rect', { x: 0, y: 0, w: '100%', h: 0.08, fill: { type: 'solid', color: ACCENT_BLUE } });
    slide.addShape('ellipse', { x: 7, y: -1, w: 5, h: 5, fill: { type: 'solid', color: '1E3A5F' } });
    slide.addShape('ellipse', { x: -1, y: 4, w: 4, h: 4, fill: { type: 'solid', color: '1A1F3A' } });

    // Logo
    slide.addShape('roundRect', { x: 0.6, y: 0.5, w: 0.65, h: 0.65, rectRadius: 0.12, fill: { type: 'solid', color: ACCENT_BLUE } });
    slide.addText('OV', { x: 0.6, y: 0.5, w: 0.65, h: 0.65, fontSize: 16, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
    slide.addText('OmniView AI', { x: 1.4, y: 0.55, w: 3, h: 0.5, fontSize: 18, bold: true, color: TEXT_WHITE });

    // Title
    slide.addText('商業提案分析報告', {
      x: 0.6, y: 2.2, w: 8.5, h: 1.0,
      fontSize: 40, bold: true, color: TEXT_WHITE,
    });

    // Score badge
    const scoreColor = result.successProbability >= 80 ? ACCENT_GREEN : result.successProbability >= 60 ? ACCENT_BLUE : result.successProbability >= 40 ? ACCENT_YELLOW : ACCENT_RED;
    slide.addShape('roundRect', { x: 0.6, y: 3.4, w: 2.2, h: 1.0, rectRadius: 0.15, fill: { type: 'solid', color: scoreColor } });
    slide.addText(`成功機率 ${result.successProbability}%`, {
      x: 0.6, y: 3.4, w: 2.2, h: 1.0,
      fontSize: 20, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle',
    });

    // Summary
    slide.addText(result.executiveSummary.slice(0, 120) + '...', {
      x: 0.6, y: 4.7, w: 8.5, h: 1.0,
      fontSize: 14, color: TEXT_MUTED, lineSpacingMultiple: 1.5,
    });

    slide.addText('由 OmniView AI 360° 董事會生成', {
      x: 0.6, y: 6.7, w: 5, h: 0.4, fontSize: 10, color: TEXT_DIM,
    });
    addSlideNumber(slide, slideNum, TOTAL);
  }

  // ========== SLIDE 2: EXECUTIVE SUMMARY ==========
  {
    slideNum++;
    const slide = pptx.addSlide();
    addGradientBg(slide);
    addSectionTitle(slide, '📋', '執行摘要');

    slide.addShape('roundRect', {
      x: 0.5, y: 1.1, w: 9, h: 4.5, rectRadius: 0.2,
      fill: { type: 'solid', color: CARD_BG },
      line: { color: '334155', width: 1 },
    });
    slide.addText(result.executiveSummary, {
      x: 0.8, y: 1.3, w: 8.4, h: 4.1,
      fontSize: 15, color: 'CBD5E1', lineSpacingMultiple: 1.8,
      valign: 'top',
    });
    addSlideNumber(slide, slideNum, TOTAL);
  }

  // ========== SLIDE 3: MARKET OPPORTUNITY ==========
  {
    slideNum++;
    const slide = pptx.addSlide();
    addGradientBg(slide);
    addSectionTitle(slide, '📈', '市場機會');

    // KPI cards
    const kpis = [
      { label: '市場規模 (TAM)', value: result.marketAnalysis.size, color: ACCENT_BLUE },
      { label: '年均成長率 (CAGR)', value: result.marketAnalysis.growthRate, color: ACCENT_GREEN },
      { label: '損益平衡', value: result.breakEvenPoint, color: ACCENT_PURPLE },
    ];
    kpis.forEach((kpi, i) => {
      const x = 0.5 + i * 3.15;
      slide.addShape('roundRect', { x, y: 1.2, w: 2.9, h: 1.6, rectRadius: 0.15, fill: { type: 'solid', color: CARD_BG }, line: { color: '334155', width: 1 } });
      slide.addText(kpi.label, { x: x + 0.2, y: 1.35, w: 2.5, h: 0.4, fontSize: 11, color: TEXT_MUTED });
      slide.addText(kpi.value, { x: x + 0.2, y: 1.8, w: 2.5, h: 0.6, fontSize: 22, bold: true, color: kpi.color });
    });

    // Description
    slide.addShape('roundRect', { x: 0.5, y: 3.1, w: 9, h: 3.5, rectRadius: 0.15, fill: { type: 'solid', color: CARD_BG }, line: { color: '334155', width: 1 } });
    slide.addText(result.marketAnalysis.description, {
      x: 0.8, y: 3.3, w: 8.4, h: 3.1,
      fontSize: 13, color: 'CBD5E1', lineSpacingMultiple: 1.7, valign: 'top',
    });
    addSlideNumber(slide, slideNum, TOTAL);
  }

  // ========== SLIDE 4: FINANCIAL PROJECTIONS ==========
  {
    slideNum++;
    const slide = pptx.addSlide();
    addGradientBg(slide);
    addSectionTitle(slide, '💰', '財務預測');

    // Table
    const headerRow: PptxGenJS.TableCell[] = [
      { text: '年度', options: { bold: true, color: 'FFFFFF', fill: { color: '1E3A5F' }, fontSize: 13, align: 'center' } },
      { text: '營收', options: { bold: true, color: 'FFFFFF', fill: { color: '1E3A5F' }, fontSize: 13, align: 'center' } },
      { text: '成本', options: { bold: true, color: 'FFFFFF', fill: { color: '1E3A5F' }, fontSize: 13, align: 'center' } },
      { text: '淨利', options: { bold: true, color: 'FFFFFF', fill: { color: '1E3A5F' }, fontSize: 13, align: 'center' } },
    ];

    const rows: PptxGenJS.TableCell[][] = [headerRow];
    result.financials.forEach((f, i) => {
      const bg = i % 2 === 0 ? CARD_BG : '162032';
      const revenue = Number(f.revenue) || 0;
      const costs = Number(f.costs) || 0;
      const profit = Number(f.profit) || 0;
      const profitColor = profit >= 0 ? ACCENT_GREEN : ACCENT_RED;
      rows.push([
        { text: f.year ?? '—', options: { color: TEXT_WHITE, fill: { color: bg }, fontSize: 12, align: 'center' } },
        { text: `$${revenue.toLocaleString()}`, options: { color: ACCENT_BLUE, fill: { color: bg }, fontSize: 12, align: 'center', bold: true } },
        { text: `$${costs.toLocaleString()}`, options: { color: TEXT_MUTED, fill: { color: bg }, fontSize: 12, align: 'center' } },
        { text: `$${profit.toLocaleString()}`, options: { color: profitColor, fill: { color: bg }, fontSize: 12, align: 'center', bold: true } },
      ]);
    });

    slide.addTable(rows, {
      x: 0.5, y: 1.2, w: 9,
      border: { type: 'solid', color: '334155', pt: 0.5 },
      colW: [2.25, 2.25, 2.25, 2.25],
      rowH: [0.5, ...result.financials.map(() => 0.45)],
    });

    // Break-even callout
    const tableBottom = 1.2 + 0.5 + result.financials.length * 0.45 + 0.3;
    slide.addShape('roundRect', { x: 0.5, y: tableBottom, w: 9, h: 0.8, rectRadius: 0.12, fill: { type: 'solid', color: CARD_BG }, line: { color: '334155', width: 1 } });
    slide.addText(`📍 損益平衡點：${result.breakEvenPoint}`, {
      x: 0.8, y: tableBottom, w: 8.4, h: 0.8,
      fontSize: 16, bold: true, color: ACCENT_GREEN, valign: 'middle',
    });
    addSlideNumber(slide, slideNum, TOTAL);
  }

  // ========== SLIDE 5: COMPETITIVE LANDSCAPE ==========
  {
    slideNum++;
    const slide = pptx.addSlide();
    addGradientBg(slide);
    addSectionTitle(slide, '⚔️', '競爭態勢分析');

    const compHeader: PptxGenJS.TableCell[] = [
      { text: '競爭對手', options: { bold: true, color: 'FFFFFF', fill: { color: '1E3A5F' }, fontSize: 13, align: 'center' } },
      { text: '優勢', options: { bold: true, color: 'FFFFFF', fill: { color: '1E3A5F' }, fontSize: 13, align: 'center' } },
      { text: '劣勢', options: { bold: true, color: 'FFFFFF', fill: { color: '1E3A5F' }, fontSize: 13, align: 'center' } },
    ];
    const compRows: PptxGenJS.TableCell[][] = [compHeader];
    result.competitors.forEach((c, i) => {
      const bg = i % 2 === 0 ? CARD_BG : '162032';
      compRows.push([
        { text: c.name, options: { bold: true, color: TEXT_WHITE, fill: { color: bg }, fontSize: 12, align: 'left', margin: [0, 0, 0, 8] } },
        { text: c.strength, options: { color: ACCENT_GREEN, fill: { color: bg }, fontSize: 11 } },
        { text: c.weakness, options: { color: ACCENT_RED, fill: { color: bg }, fontSize: 11 } },
      ]);
    });

    slide.addTable(compRows, {
      x: 0.5, y: 1.2, w: 9,
      border: { type: 'solid', color: '334155', pt: 0.5 },
      colW: [2.5, 3.25, 3.25],
      rowH: [0.5, ...result.competitors.map(() => 0.7)],
      autoPage: true,
    });
    addSlideNumber(slide, slideNum, TOTAL);
  }

  // ========== SLIDE 6: STRATEGIC ROADMAP ==========
  {
    slideNum++;
    const slide = pptx.addSlide();
    addGradientBg(slide);
    addSectionTitle(slide, '🗺️', '策略路線圖');

    result.roadmap.forEach((item, i) => {
      const y = 1.2 + i * 1.35;
      if (y > 6.5) return; // safety

      // Timeline dot + line
      slide.addShape('ellipse', { x: 0.7, y: y + 0.15, w: 0.25, h: 0.25, fill: { type: 'solid', color: ACCENT_PURPLE } });
      if (i < result.roadmap.length - 1) {
        slide.addShape('rect', { x: 0.8, y: y + 0.4, w: 0.05, h: 1.0, fill: { type: 'solid', color: '334155' } });
      }

      // Phase label
      slide.addText(`${item.timeframe}  —  ${item.phase}`, {
        x: 1.2, y, w: 4, h: 0.35,
        fontSize: 13, bold: true, color: ACCENT_PURPLE,
      });

      // Content card
      slide.addShape('roundRect', { x: 1.2, y: y + 0.38, w: 8, h: 0.8, rectRadius: 0.1, fill: { type: 'solid', color: CARD_BG }, line: { color: '334155', width: 1 } });
      slide.addText([
        { text: '產品：', options: { fontSize: 11, color: TEXT_DIM, bold: true } },
        { text: item.product + '    ', options: { fontSize: 11, color: 'CBD5E1' } },
        { text: '技術：', options: { fontSize: 11, color: TEXT_DIM, bold: true } },
        { text: item.technology, options: { fontSize: 11, color: 'CBD5E1' } },
      ], { x: 1.4, y: y + 0.42, w: 7.6, h: 0.7, lineSpacingMultiple: 1.4, valign: 'middle' });
    });
    addSlideNumber(slide, slideNum, TOTAL);
  }

  // ========== SLIDE 7: RISK MATRIX ==========
  {
    slideNum++;
    const slide = pptx.addSlide();
    addGradientBg(slide);
    addSectionTitle(slide, '🛡️', '風險評估');

    const riskHeader: PptxGenJS.TableCell[] = [
      { text: '風險', options: { bold: true, color: 'FFFFFF', fill: { color: '1E3A5F' }, fontSize: 13 } },
      { text: '衝擊程度', options: { bold: true, color: 'FFFFFF', fill: { color: '1E3A5F' }, fontSize: 13, align: 'center' } },
      { text: '因應策略', options: { bold: true, color: 'FFFFFF', fill: { color: '1E3A5F' }, fontSize: 13 } },
    ];
    const riskRows: PptxGenJS.TableCell[][] = [riskHeader];
    result.risks.forEach((r, i) => {
      const bg = i % 2 === 0 ? CARD_BG : '162032';
      const impactColor = r.impact === 'High' ? ACCENT_RED : r.impact === 'Medium' ? ACCENT_YELLOW : ACCENT_GREEN;
      const impactLabel = r.impact === 'High' ? '🔴 高' : r.impact === 'Medium' ? '🟡 中' : '🟢 低';
      riskRows.push([
        { text: r.risk, options: { color: TEXT_WHITE, fill: { color: bg }, fontSize: 11, bold: true } },
        { text: impactLabel, options: { color: impactColor, fill: { color: bg }, fontSize: 12, bold: true, align: 'center' } },
        { text: r.mitigation, options: { color: 'CBD5E1', fill: { color: bg }, fontSize: 11 } },
      ]);
    });

    slide.addTable(riskRows, {
      x: 0.5, y: 1.2, w: 9,
      border: { type: 'solid', color: '334155', pt: 0.5 },
      colW: [2.5, 1.5, 5],
      rowH: [0.5, ...result.risks.map(() => 0.65)],
      autoPage: true,
    });
    addSlideNumber(slide, slideNum, TOTAL);
  }

  // ========== SLIDE 8: STAKEHOLDER PERSPECTIVES ==========
  {
    slideNum++;
    const slide = pptx.addSlide();
    addGradientBg(slide);
    addSectionTitle(slide, '👥', 'AI 虛擬董事會評估');

    const personas = result.personaEvaluations.slice(0, 6); // max 6
    const cols = personas.length <= 3 ? personas.length : 3;
    const cardW = (9 - (cols - 1) * 0.2) / cols;

    personas.forEach((p, i) => {
      const col = i % cols;
      const row = Math.floor(i / cols);
      const x = 0.5 + col * (cardW + 0.2);
      const y = 1.2 + row * 2.8;

      slide.addShape('roundRect', { x, y, w: cardW, h: 2.5, rectRadius: 0.15, fill: { type: 'solid', color: CARD_BG }, line: { color: '334155', width: 1 } });

      // Score
      const sc = p.score;
      const sColor = sc >= 80 ? ACCENT_GREEN : sc >= 60 ? ACCENT_BLUE : sc >= 40 ? ACCENT_YELLOW : ACCENT_RED;
      slide.addText(`${sc}`, { x: x + cardW - 0.8, y: y + 0.1, w: 0.7, h: 0.4, fontSize: 18, bold: true, color: sColor, align: 'right' });

      // Name
      slide.addText(p.role, { x: x + 0.15, y: y + 0.1, w: cardW - 1, h: 0.35, fontSize: 13, bold: true, color: TEXT_WHITE });

      // Quote
      slide.addText(`"${p.keyQuote}"`, { x: x + 0.15, y: y + 0.5, w: cardW - 0.3, h: 0.7, fontSize: 10, italic: true, color: TEXT_MUTED, lineSpacingMultiple: 1.4, valign: 'top' });

      // Concern
      slide.addText([
        { text: '擔憂：', options: { fontSize: 9, color: ACCENT_RED, bold: true } },
        { text: p.concern, options: { fontSize: 9, color: TEXT_MUTED } },
      ], { x: x + 0.15, y: y + 1.3, w: cardW - 0.3, h: 1.0, lineSpacingMultiple: 1.3, valign: 'top' });
    });
    addSlideNumber(slide, slideNum, TOTAL);
  }

  // ========== SLIDE 9: FINAL VERDICTS ==========
  {
    slideNum++;
    const slide = pptx.addSlide();
    addGradientBg(slide);
    addSectionTitle(slide, '⚖️', '最終裁決');

    const verdicts = [
      { title: '🔥 激進觀點', text: result.finalVerdicts.aggressive, color: 'F97316', bgColor: '2D1B0E' },
      { title: '⚖️ 平衡觀點', text: result.finalVerdicts.balanced, color: ACCENT_BLUE, bgColor: '0E1B2D' },
      { title: '🛡️ 保守觀點', text: result.finalVerdicts.conservative, color: ACCENT_PURPLE, bgColor: '1B0E2D' },
    ];

    verdicts.forEach((v, i) => {
      const x = 0.5 + i * 3.15;
      slide.addShape('roundRect', { x, y: 1.2, w: 2.9, h: 5.2, rectRadius: 0.15, fill: { type: 'solid', color: v.bgColor }, line: { color: v.color + '40', width: 1 } });
      slide.addText(v.title, { x: x + 0.15, y: 1.35, w: 2.6, h: 0.45, fontSize: 15, bold: true, color: v.color });
      slide.addShape('rect', { x: x + 0.15, y: 1.85, w: 1, h: 0.03, fill: { type: 'solid', color: v.color } });
      slide.addText(v.text, { x: x + 0.15, y: 2.0, w: 2.6, h: 4.2, fontSize: 11, color: 'CBD5E1', lineSpacingMultiple: 1.6, valign: 'top' });
    });
    addSlideNumber(slide, slideNum, TOTAL);
  }

  // ========== SLIDE 10: CALL TO ACTION ==========
  {
    slideNum++;
    const slide = pptx.addSlide();
    slide.background = { fill: DARK_BG };

    // Decorative
    slide.addShape('ellipse', { x: 3, y: 1.5, w: 4, h: 4, fill: { type: 'solid', color: '1E3A5F' } });

    slide.addText('下一步行動', {
      x: 0.5, y: 2.0, w: 9, h: 1,
      fontSize: 38, bold: true, color: TEXT_WHITE, align: 'center',
    });

    slide.addText('準備好將這個商業提案推進到下一階段了嗎？', {
      x: 1, y: 3.2, w: 8, h: 0.7,
      fontSize: 18, color: TEXT_MUTED, align: 'center',
    });

    // Score reminder
    const scoreColor = result.successProbability >= 80 ? ACCENT_GREEN : result.successProbability >= 60 ? ACCENT_BLUE : result.successProbability >= 40 ? ACCENT_YELLOW : ACCENT_RED;
    slide.addShape('roundRect', { x: 3.5, y: 4.2, w: 3, h: 1.2, rectRadius: 0.2, fill: { type: 'solid', color: CARD_BG }, line: { color: scoreColor, width: 2 } });
    slide.addText([
      { text: 'AI 評估成功機率\n', options: { fontSize: 12, color: TEXT_MUTED } },
      { text: `${result.successProbability}%`, options: { fontSize: 32, bold: true, color: scoreColor } },
    ], { x: 3.5, y: 4.2, w: 3, h: 1.2, align: 'center', valign: 'middle' });

    slide.addText('此報告由 OmniView AI 360° 董事會自動生成', {
      x: 1, y: 6.5, w: 8, h: 0.4, fontSize: 10, color: TEXT_DIM, align: 'center',
    });
    addSlideNumber(slide, slideNum, TOTAL);
  }

  // ========== EXPORT ==========
  await pptx.writeFile({ fileName: 'OmniView_商業提案報告.pptx' });
};

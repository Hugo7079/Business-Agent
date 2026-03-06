import React, { useState, useRef, useEffect } from 'react';
import {
  DeepScanPhase, Phase1DiagnosisResult, Phase2Message,
  Phase3DeepReport, BusinessInput, BcgItem, SwotAnalysis,
  SteepAnalysis, PortersFiveForces, PriorityMapItem, MarketSizeAnalysis,
} from '../types';
import { runPhase1DeepScan, runPhase2Reply, runPhase3DeepReport } from '../services/geminiService';
import {
  Sparkles, Mic, StopCircle, Loader2, AlertTriangle, ChevronRight,
  MessageSquare, Send, Zap, ArrowRight, CheckCircle2, BarChart3,
  Target, TrendingUp, ShieldAlert, Layers, Map, Compass, Star,
  GitBranch, Flag, Clock, X,
} from 'lucide-react';
import {
  RadarChart, Radar, PolarGrid, PolarAngleAxis, ResponsiveContainer,
  Tooltip, Legend, ScatterChart, Scatter, XAxis, YAxis, CartesianGrid,
  ZAxis, Cell, ReferenceLine,
} from 'recharts';

// DeepScanFlow 現在作為嵌入式區塊，不接受 onExit prop
interface Props {
  initialData?: BusinessInput; // 可從 AnalysisDashboard 傳入現有分析資料
  initialReport?: Phase3DeepReport | null; // 從歷史紀錄載入的報告
}

// ─── helpers ────────────────────────────────────────────────────────────────

const priorityColor = (level: 'HIGH' | 'MEDIUM' | 'LOW') =>
  level === 'HIGH' ? '#ef4444' : level === 'MEDIUM' ? '#fbbf24' : '#34d399';

const impactBg = (impact: 'HIGH' | 'MEDIUM' | 'LOW') =>
  impact === 'HIGH' ? 'rgba(239,68,68,0.1)' : impact === 'MEDIUM' ? 'rgba(251,191,36,0.1)' : 'rgba(52,211,153,0.1)';

const bcgColor = (cat: BcgItem['category']) => {
  switch (cat) {
    case 'STAR':          return { bg: 'rgba(251,191,36,0.12)', border: '#fbbf24', label: '⭐ 明星' };
    case 'CASH_COW':      return { bg: 'rgba(52,211,153,0.12)', border: '#34d399', label: '🐄 金牛' };
    case 'QUESTION_MARK': return { bg: 'rgba(96,165,250,0.12)', border: '#60a5fa', label: '❓ 問號' };
    case 'DOG':           return { bg: 'rgba(148,163,184,0.12)', border: '#94a3b8', label: '🐕 瘦狗' };
  }
};

const levelToNum = (level: 'HIGH' | 'MEDIUM' | 'LOW') =>
  level === 'HIGH' ? 3 : level === 'MEDIUM' ? 2 : 1;

const blobToBase64 = (blob: Blob): Promise<string> =>
  new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onloadend = () => res((reader.result as string).split(',')[1]);
    reader.onerror = rej;
    reader.readAsDataURL(blob);
  });

// ─── Chart: SWOT 象限圖 ──────────────────────────────────────────────────────

const SwotChart: React.FC<{ swot: SwotAnalysis }> = ({ swot }) => {
  const quadrants = [
    { key: 'strengths'     as keyof SwotAnalysis, label: '優勢 Strengths',     short: 'S', color: '#34d399', bg: 'rgba(52,211,153,0.08)',  border: 'rgba(52,211,153,0.25)',  pos: 'top-left'     },
    { key: 'weaknesses'    as keyof SwotAnalysis, label: '劣勢 Weaknesses',    short: 'W', color: '#f87171', bg: 'rgba(248,113,113,0.08)', border: 'rgba(248,113,113,0.25)', pos: 'top-right'    },
    { key: 'opportunities' as keyof SwotAnalysis, label: '機會 Opportunities', short: 'O', color: '#60a5fa', bg: 'rgba(96,165,250,0.08)',  border: 'rgba(96,165,250,0.25)',  pos: 'bottom-left'  },
    { key: 'threats'       as keyof SwotAnalysis, label: '威脅 Threats',       short: 'T', color: '#fbbf24', bg: 'rgba(251,191,36,0.08)',  border: 'rgba(251,191,36,0.25)',  pos: 'bottom-right' },
  ];
  return (
    <div className="swot-chart-grid">
      {/* centre axis labels */}
      <div className="swot-axis-h">
        <span style={{ color: '#34d399' }}>內部因素</span>
        <span style={{ color: '#60a5fa' }}>外部因素</span>
      </div>
      <div className="swot-axis-v">
        <span style={{ color: '#94a3b8' }}>正面</span>
        <span style={{ color: '#94a3b8' }}>負面</span>
      </div>
      {quadrants.map(q => (
        <div
          key={q.key}
          className={`swot-quadrant swot-quadrant-${q.pos.replace('-', '_')}`}
          style={{ background: q.bg, border: `1px solid ${q.border}` }}
        >
          <div className="swot-q-header">
            <span className="swot-q-letter" style={{ background: q.color + '22', color: q.color, border: `1px solid ${q.color}55` }}>{q.short}</span>
            <span className="swot-q-label" style={{ color: q.color }}>{q.label}</span>
          </div>
          <ul className="swot-q-list">
            {swot[q.key].map((item, i) => (
              <li key={i} className="swot-q-item">
                <span className="swot-q-dot" style={{ background: q.color }} />
                <span>{item}</span>
              </li>
            ))}
          </ul>
        </div>
      ))}
    </div>
  );
};

// ─── Chart: Porter's 五力雷達圖 ──────────────────────────────────────────────────────

const porterLevelFromText = (text: string): number => {
  const t = text.toLowerCase();
  if (t.includes('高') || t.includes('strong') || t.includes('intense') || t.includes('significant')) return 85;
  if (t.includes('中') || t.includes('moderate') || t.includes('medium')) return 55;
  if (t.includes('低') || t.includes('weak') || t.includes('low') || t.includes('minor')) return 25;
  return 50;
};

const PortersRadarChart: React.FC<{ forces: PortersFiveForces }> = ({ forces }) => {
  const data = [
    { force: '現有競爭', score: porterLevelFromText(forces.competitiveRivalry),        fullMark: 100 },
    { force: '新進入者', score: porterLevelFromText(forces.threatOfNewEntrants),        fullMark: 100 },
    { force: '替代品',   score: porterLevelFromText(forces.threatOfSubstitutes),        fullMark: 100 },
    { force: '買方議價', score: porterLevelFromText(forces.bargainingPowerOfBuyers),    fullMark: 100 },
    { force: '供應商',   score: porterLevelFromText(forces.bargainingPowerOfSuppliers), fullMark: 100 },
    { force: '互補者',   score: porterLevelFromText(forces.complementors),              fullMark: 100 },
  ];
  return (
    <div className="porter-chart-wrap">
      <div style={{ height: 280 }}>
        <ResponsiveContainer width="100%" height="100%">
          <RadarChart data={data} margin={{ top: 10, right: 30, bottom: 10, left: 30 }}>
            <PolarGrid stroke="#334155" />
            <PolarAngleAxis dataKey="force" tick={{ fill: '#94a3b8', fontSize: 12 }} />
            <Radar name="威脅強度" dataKey="score" stroke="#f87171" fill="#f87171" fillOpacity={0.25} />
            <Tooltip
              contentStyle={{ background: '#1e293b', border: '1px solid #334155', color: '#f8fafc', borderRadius: 8 }}
              formatter={(v: number) => [`${v}`, '威脅強度']}
            />
          </RadarChart>
        </ResponsiveContainer>
      </div>
      <div className="porter-legend">
        {[
          { label: '現有競爭', value: forces.competitiveRivalry },
          { label: '新進入者威脅', value: forces.threatOfNewEntrants },
          { label: '替代品威脁', value: forces.threatOfSubstitutes },
          { label: '買方議價', value: forces.bargainingPowerOfBuyers },
          { label: '供應商議價', value: forces.bargainingPowerOfSuppliers },
          { label: '互補者（第六力）', value: forces.complementors },
        ].map(f => (
          <div key={f.label} className="porter-legend-item">
            <span className="porter-legend-label">{f.label}</span>
            <span className="porter-legend-text">{f.value}</span>
          </div>
        ))}
      </div>
    </div>
  );
};

// ─── Chart: 優先級矩陣散點圖 ──────────────────────────────────────────────────

const IMPACT_LABELS: Record<number, string> = { 1: 'LOW', 2: 'MEDIUM', 3: 'HIGH' };
const COST_LABELS:   Record<number, string> = { 1: 'LOW', 2: 'MEDIUM', 3: 'HIGH' };

const PriorityMatrixChart: React.FC<{ items: PriorityMapItem[] }> = ({ items }) => {
  const scatterData = items.map(item => ({
    x: levelToNum(item.cost),
    y: levelToNum(item.impact),
    z: 200,
    name: item.action,
    impact: item.impact,
    cost: item.cost,
    priority: item.priority,
  }));

  const quadrantColor = (x: number, y: number) => {
    if (y === 3 && x <= 2) return '#34d399'; // high impact low/med cost → quick win
    if (y === 3 && x === 3) return '#fbbf24'; // high impact high cost → strategic
    if (y <= 2 && x <= 2) return '#60a5fa';  // low impact low cost → fill-in
    return '#94a3b8';                          // low impact high cost → avoid
  };

  const CustomDot = (props: any) => {
    const { cx, cy, payload } = props;
    const color = quadrantColor(payload.x, payload.y);
    return (
      <g>
        <circle cx={cx} cy={cy} r={18} fill={color} fillOpacity={0.2} stroke={color} strokeWidth={2} />
        <text x={cx} y={cy} textAnchor="middle" dominantBaseline="central" fill={color} fontSize={11} fontWeight={700}>
          {payload.priority}
        </text>
      </g>
    );
  };

  const CustomTooltip = ({ active, payload }: any) => {
    if (!active || !payload?.length) return null;
    const d = payload[0].payload;
    return (
      <div style={{ background: '#1e293b', border: '1px solid #334155', borderRadius: 8, padding: '10px 14px', maxWidth: 220 }}>
        <div style={{ fontWeight: 700, color: '#e2e8f0', marginBottom: 6, fontSize: 13 }}>#{d.priority} {d.name}</div>
        <div style={{ fontSize: 12, color: '#94a3b8' }}>影響力：<span style={{ color: priorityColor(d.impact) }}>{d.impact}</span></div>
        <div style={{ fontSize: 12, color: '#94a3b8' }}>成本：<span style={{ color: priorityColor(d.cost) }}>{d.cost}</span></div>
      </div>
    );
  };

  return (
    <div className="priority-matrix-wrap">
      <div className="priority-matrix-labels-y">
        <span>↑ 高影響力</span>
        <span>低影響力 ↓</span>
      </div>
      <div style={{ flex: 1 }}>
        <div style={{ height: 280 }}>
          <ResponsiveContainer width="100%" height="100%">
            <ScatterChart margin={{ top: 20, right: 20, bottom: 20, left: 20 }}>
              <CartesianGrid stroke="#1e293b" strokeDasharray="4 4" />
              <XAxis
                type="number" dataKey="x" domain={[0.5, 3.5]} ticks={[1, 2, 3]}
                tickFormatter={v => COST_LABELS[v] ?? v}
                tick={{ fill: '#64748b', fontSize: 11 }}
                label={{ value: '成本 →', position: 'insideBottomRight', offset: -5, fill: '#64748b', fontSize: 12 }}
              />
              <YAxis
                type="number" dataKey="y" domain={[0.5, 3.5]} ticks={[1, 2, 3]}
                tickFormatter={v => IMPACT_LABELS[v] ?? v}
                tick={{ fill: '#64748b', fontSize: 11 }}
              />
              <ZAxis type="number" dataKey="z" range={[200, 200]} />
              <ReferenceLine x={2} stroke="#334155" strokeDasharray="3 3" />
              <ReferenceLine y={2} stroke="#334155" strokeDasharray="3 3" />
              <Tooltip content={<CustomTooltip />} />
              <Scatter data={scatterData} shape={<CustomDot />} />
            </ScatterChart>
          </ResponsiveContainer>
        </div>
        <div style={{ height: 2 }} />
        <div className="priority-matrix-legend">
          {[
            { color: '#34d399', label: '快速勝利（高影響 / 低成本）' },
            { color: '#fbbf24', label: '策略投資（高影響 / 高成本）' },
            { color: '#60a5fa', label: '填補空缺（低影響 / 低成本）' },
            { color: '#94a3b8', label: '考慮避免（低影響 / 高成本）' },
          ].map(l => (
            <div key={l.label} className="priority-matrix-legend-item">
              <span className="priority-matrix-dot" style={{ background: l.color }} />
              <span>{l.label}</span>
            </div>
          ))}
        </div>
      </div>
      <div className="priority-matrix-labels-x">← 低成本 &nbsp;&nbsp;&nbsp; 高成本 →</div>
    </div>
  );
};

// ─── Chart: TAM/SAM/SOM 金字塔圖 ─────────────────────────────────────────────

const TamPyramidChart: React.FC<{ marketSize: MarketSizeAnalysis }> = ({ marketSize }) => {
  const layers = [
    { key: 'tam' as const, label: 'TAM', sub: '總體潛在市場', value: marketSize.tam, color: '#60a5fa', width: '100%' },
    { key: 'sam' as const, label: 'SAM', sub: '可服務市場',   value: marketSize.sam, color: '#34d399', width: '72%'  },
    { key: 'som' as const, label: 'SOM', sub: '可獲得份額',   value: marketSize.som, color: '#a78bfa', width: '44%'  },
  ];
  return (
    <div className="tam-pyramid-wrap">
      <div className="tam-pyramid-bars">
        {layers.map((l, i) => (
          <div key={l.key} className="tam-pyramid-row" style={{ animationDelay: `${i * 120}ms` }}>
            <div className="tam-pyramid-bar-outer">
              <div
                className="tam-pyramid-bar-inner"
                style={{ width: l.width, background: `linear-gradient(90deg, ${l.color}33, ${l.color}18)`, border: `1px solid ${l.color}55` }}
              >
                <span className="tam-pyramid-bar-badge" style={{ background: l.color + '22', color: l.color, border: `1px solid ${l.color}55` }}>
                  {l.label}
                </span>
                <span className="tam-pyramid-bar-value">{l.value}</span>
              </div>
            </div>
          </div>
        ))}
      </div>
      <div className="tam-pyramid-legend">
        {layers.map(l => (
          <div key={l.key} className="tam-pyramid-legend-item">
            <span className="tam-pyramid-legend-dot" style={{ background: l.color }} />
            <div>
              <span className="tam-pyramid-legend-label" style={{ color: l.color }}>{l.label}</span>
              <span className="tam-pyramid-legend-sub"> · {l.sub}</span>
            </div>
          </div>
        ))}
      </div>
      <p className="ds-muted-text" style={{ marginTop: 16, fontSize: '0.8rem' }}>{marketSize.description}</p>
    </div>
  );
};

// ─── Chart: STEEP 雷達圖 ──────────────────────────────────────────────────────

const steepLevelFromText = (text: string): number => {
  const words = text.toLowerCase().split(/\s+/);
  let score = 50;
  const highWords = ['重大', '顯著', '快速', '強烈', '關鍵', 'significant', 'major', 'rapid', 'strong', 'critical', 'high'];
  const lowWords  = ['輕微', '緩慢', '穩定', '低', '有限', 'minor', 'slow', 'stable', 'low', 'limited'];
  if (highWords.some(w => words.some(wd => wd.includes(w)))) score = 80;
  else if (lowWords.some(w => words.some(wd => wd.includes(w)))) score = 30;
  return score;
};

const SteepRadarChart: React.FC<{ steep: SteepAnalysis }> = ({ steep }) => {
  const data = [
    { factor: '社會 S', score: steepLevelFromText(steep.social),        fullMark: 100 },
    { factor: '技術 T', score: steepLevelFromText(steep.technological),  fullMark: 100 },
    { factor: '經濟 E', score: steepLevelFromText(steep.economic),       fullMark: 100 },
    { factor: '環境 E', score: steepLevelFromText(steep.environmental),  fullMark: 100 },
    { factor: '政治 P', score: steepLevelFromText(steep.political),      fullMark: 100 },
  ];
  const details = [
    { label: '社會', text: steep.social },
    { label: '技術', text: steep.technological },
    { label: '經濟', text: steep.economic },
    { label: '環境', text: steep.environmental },
    { label: '政治', text: steep.political },
  ];
  return (
    <div className="steep-chart-wrap">
      <div style={{ height: 270 }}>
        <ResponsiveContainer width="100%" height="100%">
          <RadarChart data={data} margin={{ top: 10, right: 30, bottom: 10, left: 30 }}>
            <PolarGrid stroke="#334155" />
            <PolarAngleAxis dataKey="factor" tick={{ fill: '#94a3b8', fontSize: 12 }} />
            <Radar name="影響強度" dataKey="score" stroke="#818cf8" fill="#818cf8" fillOpacity={0.3} />
            <Tooltip
              contentStyle={{ background: '#1e293b', border: '1px solid #334155', color: '#f8fafc', borderRadius: 8 }}
              formatter={(v: number) => [`${v}`, '影響強度']}
            />
          </RadarChart>
        </ResponsiveContainer>
      </div>
      <div className="steep-detail-grid">
        {details.map(d => (
          <div key={d.label} className="steep-detail-cell">
            <div className="steep-detail-header">
              <span className="steep-detail-label">{d.label}</span>
            </div>
            <p className="ds-muted-text" style={{ fontSize: '0.78rem' }}>{d.text}</p>
          </div>
        ))}
      </div>
    </div>
  );
};

// ─── Phase 1 Card ────────────────────────────────────────────────────────────

const Phase1Card: React.FC<{
  result: Phase1DiagnosisResult;
  onContinue: (firstQ: string) => void;
  onSkipToPhase3: () => void;
}> = ({ result, onContinue, onSkipToPhase3 }) => {
  const firstQuestion = result.infoGaps.find(g => g.priority === 'HIGH')?.question || result.infoGaps[0]?.question || '';
  return (
    <div className="ds-phase1-card">
      <div className="ds-phase-badge"><Sparkles size={14} /> Phase 1 · 初步掃描完成</div>
      <div className="ds-insight-box">
        <div className="ds-insight-label">核心洞察</div>
        <div className="ds-insight-text">"{result.coreInsight}"</div>
        <div className="ds-insight-score">
          初步成功率：
          <span style={{ color: result.successRateEstimate >= 60 ? '#34d399' : result.successRateEstimate >= 40 ? '#fbbf24' : '#ef4444' }}>
            {result.successRateEstimate}%
          </span>
        </div>
      </div>
      {result.riskWarnings.length > 0 && (
        <div className="ds-section">
          <div className="ds-section-title"><AlertTriangle size={16} style={{ color: '#fbbf24' }} /> 風險警示</div>
          <div className="ds-risk-list">
            {result.riskWarnings.map((w, i) => (
              <div className="ds-risk-item" key={i}><span className="ds-risk-dot" />{w}</div>
            ))}
          </div>
        </div>
      )}
      {result.infoGaps.length > 0 && (
        <div className="ds-section">
          <div className="ds-section-title"><MessageSquare size={16} style={{ color: '#60a5fa' }} /> 資訊缺口</div>
          <div className="ds-gap-list">
            {result.infoGaps.map((g, i) => (
              <div className="ds-gap-item" key={i}>
                <span className={`ds-gap-priority ds-gap-priority-${g.priority.toLowerCase()}`}>{g.priority}</span>
                <div>
                  <div className="ds-gap-field">{g.field}</div>
                  <div className="ds-gap-question">{g.question}</div>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
      <div className="ds-phase1-actions">
        <button className="ds-btn-primary" onClick={() => onContinue(firstQuestion)}>
          <MessageSquare size={16} /> 開始互動補全 (Phase 2)<ChevronRight size={16} />
        </button>
        <button className="ds-btn-ghost" onClick={onSkipToPhase3}>
          <Zap size={16} /> 直接生成完整報告
        </button>
      </div>
    </div>
  );
};

// ─── Phase 2 Chat ────────────────────────────────────────────────────────────

const Phase2Chat: React.FC<{
  messages: Phase2Message[];
  completionRate: number;
  isReady: boolean;
  isLoading: boolean;
  onSend: (text: string, audio?: string) => void;
  onGenerateReport: () => void;
}> = ({ messages, completionRate, isReady, isLoading, onSend, onGenerateReport }) => {
  const [input, setInput] = useState('');
  const [isRecording, setIsRecording] = useState(false);
  const mediaRecorderRef = useRef<MediaRecorder | null>(null);
  const audioChunksRef = useRef<Blob[]>([]);
  const bottomRef = useRef<HTMLDivElement>(null);

  useEffect(() => { bottomRef.current?.scrollIntoView({ behavior: 'smooth' }); }, [messages]);

  const handleSend = () => {
    if (!input.trim() || isLoading) return;
    onSend(input.trim()); setInput('');
  };
  const startRecording = async () => {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      const mr = new MediaRecorder(stream);
      mediaRecorderRef.current = mr; audioChunksRef.current = [];
      mr.ondataavailable = e => { if (e.data.size > 0) audioChunksRef.current.push(e.data); };
      mr.start(); setIsRecording(true);
    } catch { alert('無法存取麥克風，請檢查權限設定。'); }
  };
  const stopRecording = () => {
    if (mediaRecorderRef.current && isRecording) {
      mediaRecorderRef.current.stop(); setIsRecording(false);
      mediaRecorderRef.current.stream.getTracks().forEach(t => t.stop());
      mediaRecorderRef.current.onstop = async () => {
        const b64 = await blobToBase64(new Blob(audioChunksRef.current, { type: 'audio/webm' }));
        onSend(input.trim(), b64); setInput('');
      };
    }
  };
  const barColor = completionRate >= 80 ? '#34d399' : completionRate >= 50 ? '#60a5fa' : '#fbbf24';

  return (
    <div className="ds-chat-container">
      <div className="ds-progress-bar-wrap">
        <div className="ds-progress-bar-label">
          <span>資訊飽滿度</span>
          <span style={{ color: barColor, fontWeight: 700 }}>{completionRate}%</span>
        </div>
        <div className="ds-progress-track">
          <div className="ds-progress-fill" style={{ width: `${completionRate}%`, background: barColor, transition: 'width 0.5s ease' }} />
        </div>
        {isReady && <div className="ds-ready-hint"><CheckCircle2 size={14} />資訊已充足，可以生成完整報告！</div>}
      </div>
      <div className="ds-chat-messages">
        {messages.map((msg, i) => (
          <div key={i} className={`ds-msg ds-msg-${msg.role}`}>
            {msg.role === 'assistant' && <div className="ds-msg-avatar"><Sparkles size={14} /></div>}
            <div className={`ds-msg-bubble ds-msg-bubble-${msg.role}`}>{msg.content}</div>
          </div>
        ))}
        {isLoading && (
          <div className="ds-msg ds-msg-assistant">
            <div className="ds-msg-avatar"><Sparkles size={14} /></div>
            <div className="ds-msg-bubble ds-msg-bubble-assistant ds-thinking"><span /><span /><span /></div>
          </div>
        )}
        <div ref={bottomRef} />
      </div>
      <div className="ds-chat-input-area">
        {isReady && (
          <button className="ds-generate-btn" onClick={onGenerateReport}>
            <Zap size={16} /> 生成完整深度報告 (Phase 3)<ArrowRight size={16} />
          </button>
        )}
        <div className="ds-input-row">
          <textarea
            className="ds-chat-input"
            placeholder="輸入你的回答..."
            value={input}
            onChange={e => setInput(e.target.value)}
            onKeyDown={e => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleSend(); } }}
            disabled={isLoading || isRecording}
            rows={2}
          />
          <div className="ds-input-btns">
            {!isRecording
              ? <button className="ds-mic-btn" onClick={startRecording} disabled={isLoading}><Mic size={18} /></button>
              : <button className="ds-mic-btn ds-mic-stop" onClick={stopRecording}><StopCircle size={22} /></button>}
            <button className="ds-send-btn" onClick={handleSend} disabled={!input.trim() || isLoading}>
              <Send size={18} />
            </button>
          </div>
        </div>
        <div className="ds-input-hint">按 Enter 傳送 · Shift+Enter 換行 · 或點選麥克風語音輸入</div>
      </div>
    </div>
  );
};

// ─── Phase 3 Report ──────────────────────────────────────────────────────────

const Phase3Report: React.FC<{ report: Phase3DeepReport; onReset: () => void }> = ({ report, onReset }) => {
  const [activeTab, setActiveTab] = useState<'core' | 'market' | 'strategy' | 'action'>('core');
  const radarData = report.radarDimensions.map(d => ({ dimension: d.dimension, 自身: d.selfScore, 競爭對手: d.competitorAvgScore }));
  const fanSections = report.fanStrategy.split('\n---\n');
  const fanLabels = ['財務面', '客戶面', '內部流程面', '學習與成長面'];

  return (
    <div className="ds-report">
      <div className="ds-report-header">
        <div>
          <div className="ds-phase-badge" style={{ marginBottom: 8 }}><Star size={14} /> Phase 3 · Master Plan 完整深度報告</div>
          <h2 className="ds-report-title">全維度商業戰略分析</h2>
        </div>
        <button className="ds-btn-ghost" onClick={onReset}><X size={15} /> 重新掃描</button>
      </div>

      {/* 關鍵建議 */}
      <div className="ds-key-rec">
        <div className="ds-key-rec-title"><Target size={18} style={{ color: '#60a5fa' }} /> 關鍵建議</div>
        <div className="ds-key-rec-content">
          {report.keyRecommendation.split('\n').map((line, i) => (
            <div key={i} className="ds-bullet-line">
              <ChevronRight size={14} style={{ color: '#60a5fa', flexShrink: 0, marginTop: 3 }} />
              <span>{line}</span>
            </div>
          ))}
        </div>
      </div>

      {report.continueToIterate && (
        <div className="ds-key-rec" style={{ marginTop: 16, borderColor: 'rgba(167, 139, 250, 0.3)', background: 'rgba(167, 139, 250, 0.05)' }}>
          <div className="ds-key-rec-title" style={{ color: '#a78bfa' }}><Zap size={18} style={{ color: '#a78bfa' }} /> 持續迭代建議</div>
          <div className="ds-key-rec-content">
            {report.continueToIterate.split('\n').map((line, i) => (
              <div key={i} className="ds-bullet-line">
                <ChevronRight size={14} style={{ color: '#a78bfa', flexShrink: 0, marginTop: 3 }} />
                <span>{line}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Tab 導覽 */}
      <div className="ds-tab-nav">
        {([
          { key: 'core',     label: '商業核心架構',   icon: <Layers size={14} /> },
          { key: 'market',   label: '市場與競爭環境', icon: <BarChart3 size={14} /> },
          { key: 'strategy', label: '戰略定位',       icon: <Compass size={14} /> },
          { key: 'action',   label: '執行與行動',     icon: <Map size={14} /> },
        ] as { key: typeof activeTab; label: string; icon: React.ReactNode }[]).map(tab => (
          <button key={tab.key} className={`ds-tab ${activeTab === tab.key ? 'ds-tab-active' : ''}`} onClick={() => setActiveTab(tab.key)}>
            {tab.icon} {tab.label}
          </button>
        ))}
      </div>

      {/* ── 商業核心架構 ── */}
      {activeTab === 'core' && (
        <div className="ds-section-content">
          <div className="ds-card">
            <div className="ds-card-title"><Layers size={18} style={{ color: '#818cf8' }} /> 商業模式畫布 (BMC)</div>
            <div className="ds-bmc-grid">
              {([
                { key: 'keyPartners',           label: '關鍵合作夥伴', col: '1/2', row: '1/3' },
                { key: 'keyActivities',          label: '關鍵活動',     col: '2/3', row: '1/2' },
                { key: 'valuePropositions',      label: '價值主張',     col: '3/4', row: '1/3' },
                { key: 'customerRelationships',  label: '客戶關係',     col: '4/5', row: '1/2' },
                { key: 'customerSegments',       label: '客戶區隔',     col: '5/6', row: '1/3' },
                { key: 'keyResources',           label: '關鍵資源',     col: '2/3', row: '2/3' },
                { key: 'channels',               label: '通路',         col: '4/5', row: '2/3' },
                { key: 'costStructure',          label: '成本結構',     col: '1/4', row: '3/4' },
                { key: 'revenueStreams',          label: '收益來源',     col: '4/6', row: '3/4' },
              ] as { key: keyof typeof report.businessModelCanvas; label: string; col: string; row: string }[]).map(cell => (
                <div key={cell.key} className="ds-bmc-cell" style={{ gridColumn: cell.col, gridRow: cell.row }}>
                  <div className="ds-bmc-cell-label">{cell.label}</div>
                  <div className="ds-bmc-cell-content">
                    {report.businessModelCanvas[cell.key].split('\n').map((line, i) => <div key={i} className="ds-bmc-line">• {line}</div>)}
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* ▶ SWOT 象限圖 */}
          <div className="ds-card">
            <div className="ds-card-title"><ShieldAlert size={18} style={{ color: '#f87171' }} /> SWOT 分析</div>
            <SwotChart swot={report.swot} />
          </div>
        </div>
      )}

      {/* ── 市場與競爭環境 ── */}
      {activeTab === 'market' && (
        <div className="ds-section-content">
          {/* ▶ TAM/SAM/SOM 金字塔 */}
          <div className="ds-card">
            <div className="ds-card-title"><TrendingUp size={18} style={{ color: '#34d399' }} /> 市場規模評估 TAM / SAM / SOM</div>
            <TamPyramidChart marketSize={report.marketSize} />
          </div>

          {/* ▶ STEEP 雷達圖 */}
          <div className="ds-card">
            <div className="ds-card-title"><BarChart3 size={18} style={{ color: '#818cf8' }} /> STEEP 分析</div>
            <SteepRadarChart steep={report.steep} />
          </div>

          <div className="ds-card">
            <div className="ds-card-title"><Target size={18} style={{ color: '#34d399' }} /> 4C 分析</div>
            <div className="ds-fourC-grid">
              {([
                { key: 'consumer',      label: '消費者 Consumer',   color: '#60a5fa' },
                { key: 'cost',          label: '成本 Cost',          color: '#f87171' },
                { key: 'convenience',   label: '便利 Convenience',   color: '#34d399' },
                { key: 'communication', label: '溝通 Communication', color: '#fbbf24' },
              ] as { key: keyof typeof report.fourC; label: string; color: string }[]).map(c => (
                <div key={c.key} className="ds-fourC-cell" style={{ borderLeftColor: c.color }}>
                  <div className="ds-fourC-label" style={{ color: c.color }}>{c.label}</div>
                  <p className="ds-muted-text">{report.fourC[c.key]}</p>
                </div>
              ))}
            </div>
          </div>

          {/* ▶ Porter's 五力雷達圖 */}
          <div className="ds-card">
            <div className="ds-card-title"><GitBranch size={18} style={{ color: '#f87171' }} /> Porter's 五力 + 第六力</div>
            <PortersRadarChart forces={report.portersFiveForces} />
          </div>
        </div>
      )}

      {/* ── 戰略定位 ── */}
      {activeTab === 'strategy' && (
        <div className="ds-section-content">
          <div className="ds-card">
            <div className="ds-card-title"><Compass size={18} style={{ color: '#60a5fa' }} /> 市場定位雷達圖</div>
            <div style={{ height: 320 }}>
              <ResponsiveContainer width="100%" height="100%">
                <RadarChart data={radarData}>
                  <PolarGrid stroke="#334155" />
                  <PolarAngleAxis dataKey="dimension" tick={{ fill: '#94a3b8', fontSize: 12 }} />
                  <Radar name="自身" dataKey="自身" stroke="#3b82f6" fill="#3b82f6" fillOpacity={0.3} />
                  <Radar name="競爭對手" dataKey="競爭對手" stroke="#ef4444" fill="#ef4444" fillOpacity={0.2} />
                  <Legend wrapperStyle={{ color: '#94a3b8' }} />
                  <Tooltip contentStyle={{ background: '#1e293b', border: '1px solid #334155', color: '#f8fafc' }} />
                </RadarChart>
              </ResponsiveContainer>
            </div>
          </div>
          <div className="ds-card">
            <div className="ds-card-title"><Star size={18} style={{ color: '#fbbf24' }} /> BCG 矩陣</div>
            <div className="ds-bcg-grid">
              {report.bcgMatrix.map((item, i) => {
                const cfg = bcgColor(item.category);
                return (
                  <div key={i} className="ds-bcg-cell" style={{ background: cfg.bg, borderColor: cfg.border + '55' }}>
                    <div className="ds-bcg-cat" style={{ color: cfg.border }}>{cfg.label}</div>
                    <div className="ds-bcg-name">{item.name}</div>
                    <p className="ds-muted-text" style={{ fontSize: '0.8rem' }}>{item.description}</p>
                  </div>
                );
              })}
            </div>
          </div>
          <div className="ds-card">
            <div className="ds-card-title"><Compass size={18} style={{ color: '#a78bfa' }} /> SPAN 戰略定位分析</div>
            <div className="ds-span-grid">
              {([
                { key: 'scope',       label: '範疇 Scope',       color: '#60a5fa' },
                { key: 'positioning', label: '定位 Positioning',  color: '#34d399' },
                { key: 'advantage',   label: '優勢 Advantage',    color: '#fbbf24' },
                { key: 'network',     label: '網絡 Network',      color: '#a78bfa' },
              ] as { key: keyof typeof report.spanAnalysis; label: string; color: string }[]).map(s => (
                <div key={s.key} className="ds-span-cell" style={{ borderTopColor: s.color }}>
                  <div className="ds-span-label" style={{ color: s.color }}>{s.label}</div>
                  <p className="ds-muted-text">{report.spanAnalysis[s.key]}</p>
                </div>
              ))}
            </div>
          </div>
          <div className="ds-card">
            <div className="ds-card-title"><Flag size={18} style={{ color: '#34d399' }} /> FAN 策略地圖</div>
            <div className="ds-fan-grid">
              {fanSections.map((section, i) => (
                <div key={i} className="ds-fan-cell">
                  <div className="ds-fan-label">{fanLabels[i] || `面向 ${i + 1}`}</div>
                  {section.split('\n').filter(Boolean).map((line, j) => (
                    <div key={j} className="ds-bullet-line">
                      <ChevronRight size={13} style={{ color: '#475569', flexShrink: 0, marginTop: 3 }} />
                      <span className="ds-muted-text">{line}</span>
                    </div>
                  ))}
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* ── 執行與行動 ── */}
      {activeTab === 'action' && (
        <div className="ds-section-content">
          {/* ▶ 優先級矩陣散點圖 */}
          <div className="ds-card">
            <div className="ds-card-title"><Target size={18} style={{ color: '#f87171' }} /> 優先級矩陣</div>
            <PriorityMatrixChart items={report.priorityMap} />
          </div>
          <div className="ds-card">
            <div className="ds-card-title"><Clock size={18} style={{ color: '#60a5fa' }} /> 行動地圖（未來 3-6 個月）</div>
            <div className="ds-action-timeline">
              {report.actionMap.map((milestone, i) => (
                <div key={i} className="ds-action-item">
                  <div className="ds-action-dot" />
                  <div className="ds-action-content">
                    <div className="ds-action-month">{milestone.month}</div>
                    <div className="ds-action-milestone">{milestone.milestone}</div>
                    <div className="ds-action-steps">
                      {milestone.keyActions.map((action, j) => (
                        <div key={j} className="ds-action-step">
                          <ChevronRight size={13} style={{ color: '#60a5fa', flexShrink: 0 }} />{action}
                        </div>
                      ))}
                    </div>
                    <div className="ds-action-metric">
                      <CheckCircle2 size={13} />{milestone.successMetric}
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

// ─── Main Component (Embedded) ───────────────────────────────────────────────

const DeepScanFlow: React.FC<Props> = ({ initialData, initialReport }) => {
  const [phase, setPhase] = useState<DeepScanPhase>('INPUT');
  const [inputText, setInputText] = useState('');
  const [isRecording, setIsRecording] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const [phase1Result, setPhase1Result] = useState<Phase1DiagnosisResult | null>(null);
  const [messages, setMessages] = useState<Phase2Message[]>([]);
  const [filledData, setFilledData] = useState<BusinessInput | null>(initialData || null);
  const [completionRate, setCompletionRate] = useState(0);
  const [isReady, setIsReady] = useState(false);
  const [phase3Report, setPhase3Report] = useState<Phase3DeepReport | null>(null);

  const mediaRecorderRef = useRef<MediaRecorder | null>(null);
  const audioChunksRef = useRef<Blob[]>([]);

  // 若傳入 initialReport，直接顯示報告
  useEffect(() => {
    if (initialReport) {
      setPhase3Report(initialReport);
      setPhase('PHASE3');
      setFilledData(initialData || null);
    }
  }, [initialReport, initialData]);

  // 若傳入 initialData，預填 inputText
  useEffect(() => {
    if (initialData?.idea) setInputText(initialData.idea);
  }, [initialData]);

  const handlePhase1 = async (audioBase64?: string) => {
    if (!inputText.trim() && !audioBase64) return;
    setIsLoading(true); setError(null);
    try {
      const result = await runPhase1DeepScan(inputText, audioBase64);
      setPhase1Result(result);
      setFilledData(result.extractedData);
      setPhase('PHASE1');
    } catch (err: any) { setError(err.message || '掃描失敗，請重試。'); }
    finally { setIsLoading(false); }
  };

  const handleStartPhase2 = (firstQuestion: string) => {
    setMessages([{
      role: 'assistant',
      content: firstQuestion
        ? `我已完成初步掃描。為了讓分析更精準，我需要了解更多細節。\n\n**${firstQuestion}**`
        : '初步掃描完成！請告訴我更多關於你的提案細節，以便進行更深入的分析。',
      timestamp: Date.now(),
    }]);
    setPhase('PHASE2');
  };

  const handlePhase2Send = async (text: string, audioBase64?: string) => {
    if (!filledData) return;
    const userMsg: Phase2Message = { role: 'user', content: text || '（語音輸入）', timestamp: Date.now() };
    setMessages(prev => [...prev, userMsg]);
    setIsLoading(true);
    try {
      const currentMessages = messages;
      const history = [...currentMessages, userMsg].map(m => ({ role: m.role, content: m.content }));
      const resp = await runPhase2Reply(text, filledData, history, audioBase64);
      setFilledData(resp.updatedData);
      setCompletionRate(resp.completionRate);
      setIsReady(resp.isReady);
      let aiContent = resp.dynamicFeedback;
      if (resp.nextQuestions.length > 0) {
        aiContent += '\n\n' + resp.nextQuestions.map((q, i) => `**問題${i + 1}：${q}**`).join('\n');
      } else if (resp.isReady) {
        aiContent += '\n\n資訊已充足，可以點擊下方按鈕生成完整的深度分析報告！';
      }
      setMessages(prev => [...prev, { role: 'assistant', content: aiContent, timestamp: Date.now() }]);
    } catch (err: any) { setError(err.message || 'AI 回覆失敗，請重試。'); }
    finally { setIsLoading(false); }
  };

  const handleGeneratePhase3 = async () => {
    if (!filledData) return;
    setPhase('PHASE3_LOADING'); setError(null);
    try {
      const report = await runPhase3DeepReport(filledData);
      setPhase3Report(report); setPhase('PHASE3');
    } catch (err: any) {
      setError(err.message || '報告生成失敗，請重試。');
      setPhase(phase1Result ? 'PHASE2' : 'INPUT');
    }
  };

  const handleReset = () => {
    setPhase('INPUT'); setPhase1Result(null); setMessages([]);
    setFilledData(initialData || null); setCompletionRate(0);
    setIsReady(false); setPhase3Report(null);
    setInputText(initialData?.idea || '');
  };

  const startRecording = async () => {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      const mr = new MediaRecorder(stream);
      mediaRecorderRef.current = mr; audioChunksRef.current = [];
      mr.ondataavailable = e => { if (e.data.size > 0) audioChunksRef.current.push(e.data); };
      mr.start(); setIsRecording(true);
    } catch { alert('無法存取麥克風，請檢查權限設定。'); }
  };
  const stopRecording = () => {
    if (mediaRecorderRef.current && isRecording) {
      mediaRecorderRef.current.stop(); setIsRecording(false);
      mediaRecorderRef.current.stream.getTracks().forEach(t => t.stop());
      mediaRecorderRef.current.onstop = async () => {
        const b64 = await blobToBase64(new Blob(audioChunksRef.current, { type: 'audio/webm' }));
        await handlePhase1(b64);
      };
    }
  };

  // 進度步驟
  const stepOrder: DeepScanPhase[] = ['INPUT', 'PHASE1', 'PHASE2', 'PHASE3'];
  const currentStepIndex = phase === 'PHASE3_LOADING' ? 3 : stepOrder.indexOf(phase);
  const stepLabels = ['輸入', 'Phase 1', 'Phase 2', 'Master Plan'];

  return (
    <div className="ds-embed">
      {/* 進度列 */}
      <div className="ds-embed-progress">
        {stepLabels.map((label, i) => {
          const isActive = i === currentStepIndex;
          const isDone = i < currentStepIndex;
          return (
            <React.Fragment key={i}>
              <div className={`ds-step ${isActive ? 'ds-step-active' : isDone ? 'ds-step-done' : ''}`}>
                {isDone ? <CheckCircle2 size={14} /> : <span>{i + 1}</span>}
                <span className="ds-step-label">{label}</span>
              </div>
              {i < 3 && <div className={`ds-step-line ${isDone ? 'ds-step-line-done' : ''}`} />}
            </React.Fragment>
          );
        })}
      </div>

      {/* 錯誤提示 */}
      {error && (
        <div className="ds-error">
          <AlertTriangle size={16} />
          <span style={{ whiteSpace: 'pre-line', flex: 1 }}>{error}</span>
          <button onClick={() => setError(null)}><X size={14} /></button>
        </div>
      )}

      {/* INPUT */}
      {phase === 'INPUT' && (
        <div className="ds-embed-input">
          <p className="ds-embed-hint">用任何方式描述你的構想（或直接使用上方表單的資料），AI 會進行初步掃描，再互動補全，最後產出完整 Master Plan 報告。</p>
          <div className="textarea-wrap" style={{ marginBottom: 12 }}>
            <textarea
              className="quick-textarea"
              style={{ minHeight: 120 }}
              placeholder={isRecording ? '🎙 正在聆聽...' : '描述你的創業構想，越詳細越好...'}
              value={inputText}
              onChange={e => setInputText(e.target.value)}
              disabled={isRecording || isLoading}
            />
            {isRecording && <div className="recording-badge"><div className="recording-dot" /> 錄音中</div>}
          </div>
          <div className="ds-embed-actions">
            <div className="mic-section" style={{ flex: 'none', flexDirection: 'row', gap: 8 }}>
              {!isRecording
                ? <button onClick={startRecording} disabled={isLoading} className="mic-btn" title="語音輸入"><Mic size={20} /></button>
                : <button onClick={stopRecording} className="mic-btn-stop"><StopCircle size={26} /></button>}
              <span className="mic-hint" style={{ fontSize: '0.75rem' }}>{isRecording ? '點擊停止' : '語音輸入'}</span>
            </div>
            <button
              className="ds-scan-btn"
              onClick={() => handlePhase1()}
              disabled={!inputText.trim() || isLoading || isRecording}
            >
              {isLoading
                ? <><Loader2 size={16} className="spin-icon" /> 掃描中...</>
                : <><Sparkles size={16} /> 開始深度掃描</>}
            </button>
          </div>
        </div>
      )}

      {/* PHASE 1 */}
      {phase === 'PHASE1' && phase1Result && (
        <Phase1Card result={phase1Result} onContinue={handleStartPhase2} onSkipToPhase3={handleGeneratePhase3} />
      )}

      {/* PHASE 2 */}
      {phase === 'PHASE2' && (
        <Phase2Chat
          messages={messages}
          completionRate={completionRate}
          isReady={isReady}
          isLoading={isLoading}
          onSend={handlePhase2Send}
          onGenerateReport={handleGeneratePhase3}
        />
      )}

      {/* PHASE 3 LOADING */}
      {phase === 'PHASE3_LOADING' && (
        <div style={{ padding: '3rem 0', textAlign: 'center' }}>
          <div className="spinner" style={{ margin: '0 auto 1.5rem' }}>
            <div className="spin-ring spin-ring-1" />
            <div className="spin-ring spin-ring-2" />
            <div className="spin-ring spin-ring-3" />
          </div>
          <h3 className="loading-title" style={{ fontSize: '1.25rem' }}>正在生成 Master Plan</h3>
          <p className="loading-sub" style={{ fontSize: '0.875rem' }}>BMC · SWOT · STEEP · 4C · Porter's 五力 · BCG · 雷達圖 · 行動地圖...</p>
        </div>
      )}

      {/* PHASE 3 */}
      {phase === 'PHASE3' && phase3Report && (
        <Phase3Report report={phase3Report} onReset={handleReset} />
      )}
    </div>
  );
};

export default DeepScanFlow;

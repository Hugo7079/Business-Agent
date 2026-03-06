import React, { useState } from 'react';
import { AnalysisResult, Phase3DeepReport, BcgItem, BusinessInput } from '../types';
import {
  Sparkles, AlertTriangle, ChevronRight, Zap, ArrowRight, CheckCircle2,
  BarChart3, Target, TrendingUp, ShieldAlert, Layers, Map, Compass, Star,
  GitBranch, Flag, Clock, Users, Swords, MessageSquare, DollarSign, Route,
  Brain, Trophy, Loader2, RefreshCw, X, Download, Truck, Leaf, Smartphone, Gavel, ArrowDown
} from 'lucide-react';
import {
  RadarChart, Radar, PolarGrid, PolarAngleAxis, ResponsiveContainer,
  Tooltip, Legend, LineChart, Line, CartesianGrid, XAxis, YAxis
} from 'recharts';
import { generatePptx } from '../services/pptxService';
import { summarizeForSlides, generateThemeImage, generateProposalTitle } from '../services/geminiService';

// ─── helpers ──────────────────────────────────────────────────────────────────

const priorityColor = (level: 'HIGH' | 'MEDIUM' | 'LOW') =>
  level === 'HIGH' ? '#ef4444' : level === 'MEDIUM' ? '#fbbf24' : '#34d399';

const impactBg = (impact: 'HIGH' | 'MEDIUM' | 'LOW') =>
  impact === 'HIGH' ? 'rgba(239,68,68,0.08)' : impact === 'MEDIUM' ? 'rgba(251,191,36,0.08)' : 'rgba(52,211,153,0.08)';

const bcgColor = (cat: BcgItem['category']) => {
  switch (cat) {
    case 'STAR':          return { bg: 'rgba(251,191,36,0.12)', border: '#fbbf24', label: '明星', icon: <Star size={16} /> };
    case 'CASH_COW':      return { bg: 'rgba(52,211,153,0.12)', border: '#34d399', label: '金牛', icon: <DollarSign size={16} /> };
    case 'QUESTION_MARK': return { bg: 'rgba(96,165,250,0.12)', border: '#60a5fa', label: '問號', icon: <MessageSquare size={16} /> };
    case 'DOG':           return { bg: 'rgba(148,163,184,0.12)', border: '#94a3b8', label: '瘦狗', icon: <X size={16} /> };
  }
};

const scoreColorFunc = (s: number) => s >= 65 ? '#34d399' : s >= 40 ? '#fbbf24' : '#ef4444';

// ─── Phase 3 Report (Main Dashboard) ──────────────────────────────────────────────────────────

interface DashboardProps {
  result: AnalysisResult;             // Full Result
  phase3Report: Phase3DeepReport | null; // Deep Report
  sourceInput?: BusinessInput;
  onReset: () => void;
  mode?: string; 
}

const AnalysisDashboard: React.FC<DashboardProps> = ({ result, phase3Report, onReset }) => {
  const [activeTab, setActiveTab] = useState<'core' | 'market' | 'strategy' | 'action' | 'full'>('core');
  const [subTab, setSubTab] = useState<'finance' | 'personas' | 'competitors' | 'verdicts'>('finance');
  
  // PPT 產生狀態
  const [isGeneratingPpt, setIsGeneratingPpt] = useState(false);
  const [pptError, setPptError] = useState<string | null>(null);
  const [pptStage, setPptStage] = useState<string>('');

  // 若沒有 phase3Report，則只顯示 Full Result (這通常發生在舊歷史紀錄)
  React.useEffect(() => {
    if (!phase3Report && result) {
      setActiveTab('full');
    }
  }, [phase3Report, result]);

  if (!phase3Report && !result) return <div>無資料</div>;

  // Safe fallback
  const report = phase3Report || {
    keyRecommendation: '資料缺失',
    businessModelCanvas: { keyPartners: '', keyActivities: '', valuePropositions: '', customerRelationships: '', customerSegments: '', keyResources: '', channels: '', costStructure: '', revenueStreams: '' },
    swot: { strengths: [], weaknesses: [], opportunities: [], threats: [] },
    marketSize: { tam: '', sam: '', som: '', description: '' },
    steep: { social: '', technological: '', economic: '', environmental: '', political: '' },
    fourC: { consumer: '', cost: '', convenience: '', communication: '' },
    portersFiveForces: { competitiveRivalry: '', threatOfNewEntrants: '', threatOfSubstitutes: '', bargainingPowerOfBuyers: '', bargainingPowerOfSuppliers: '', complementors: '' },
    radarDimensions: [],
    bcgMatrix: [],
    spanAnalysis: { scope: '', positioning: '', advantage: '', network: '' },
    fanStrategy: '',
    priorityMap: [],
    actionMap: [],
    continueToIterate: ''
  } as unknown as Phase3DeepReport;

  const radarData = report.radarDimensions.map(d => ({ dimension: d.dimension, 自身: d.selfScore, 競爭對手: d.competitorAvgScore }));
  
  // Update FAN parsing to handle loose "---" separators seen in user data
  const fanSections = report.fanStrategy 
    ? report.fanStrategy.split(/\s*---\s*|\n---\n/).filter(s => s.trim().length > 0) 
    : [];
    
  const fanLabels = ['財務面', '客戶面', '內部流程面', '學習與成長面'];

  // PPT 下載功能
  const handleGeneratePpt = async () => {
    setIsGeneratingPpt(true);
    setPptError(null);
    setPptStage('精簡內文中...');
    try {
      const summarized = await summarizeForSlides(result);
      setPptStage('生成提案標題...');
      const title = await generateProposalTitle(summarized.executiveSummary);
      setPptStage('生成主題插圖...');
      const themeImage = await generateThemeImage(summarized.executiveSummary, summarized.marketAnalysis.description);
      setPptStage('組裝簡報中...');
      await generatePptx(summarized, themeImage, title);
    } catch (err: any) {
      setPptError(err.message || '產生失敗，請重試');
    } finally {
      setIsGeneratingPpt(false);
      setPptStage('');
    }
  };

  return (
    <div className="cf-report-wrapper">
      <div className="ds-report">
        {/* 標題列 + 重新開始按鈕 */}
        <div className="ds-report-header">
          <div>
            <div className="ds-phase-badge" style={{ marginBottom: 8 }}>
              <Star size={14} /> Master Plan · 完整深度報告 (歷史紀錄)
            </div>
            <h2 className="ds-report-title">全維度商業戰略分析</h2>
          </div>
          <div style={{ display: 'flex', gap: 8 }}>
            <button
              className={`cf-reset-btn ${isGeneratingPpt ? 'disabled' : ''}`}
              onClick={handleGeneratePpt}
              disabled={isGeneratingPpt}
              title="下載 PPT"
              style={{ background: '#3b82f6', color: '#fff', border: 'none' }}
            >
              {isGeneratingPpt ? <Loader2 size={14} className="spin-icon" /> : <Download size={14} />} 
              {isGeneratingPpt ? ' 處理中...' : ' 下載簡報'}
            </button>
            <button className="cf-reset-btn" onClick={onReset}>
              <RefreshCw size={14} /> 回首頁
            </button>
          </div>
        </div>

        {pptError && <div className="ds-error" style={{marginBottom: 16}}>{pptError}</div>}

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

        {/* Tab 導覽 */}
        <div className="ds-tab-nav">
          {([
            { key: 'core',     label: '商業核心架構',   icon: <Layers size={14} /> },
            { key: 'market',   label: '市場與競爭環境', icon: <BarChart3 size={14} /> },
            { key: 'strategy', label: '戰略定位',       icon: <Compass size={14} /> },
            { key: 'action',   label: '執行與行動',     icon: <Map size={14} /> },
            { key: 'full',     label: '深度分析', icon: <Trophy size={14} /> },
          ] as { key: typeof activeTab; label: string; icon: React.ReactNode }[]).map(tab => (
            <button key={tab.key} className={`ds-tab ${activeTab === tab.key ? 'ds-tab-active' : ''}`} onClick={() => setActiveTab(tab.key)}>
              {tab.icon} {tab.label}
            </button>
          ))}
        </div>

        {activeTab === 'core' && (
          <div className="ds-section-content">
            <div className="ds-card">
              <div className="ds-card-title"><Layers size={18} style={{ color: '#818cf8' }} /> 商業模式畫布 (BMC)</div>
              <div className="ds-bmc-grid">
                {([
                  { key: 'keyPartners',          label: '關鍵合作夥伴', col: '1/2', row: '1/3' },
                  { key: 'keyActivities',         label: '關鍵活動',     col: '2/3', row: '1/2' },
                  { key: 'valuePropositions',     label: '價值主張',     col: '3/4', row: '1/3' },
                  { key: 'customerRelationships', label: '客戶關係',     col: '4/5', row: '1/2' },
                  { key: 'customerSegments',      label: '客戶區隔',     col: '5/6', row: '1/3' },
                  { key: 'keyResources',          label: '關鍵資源',     col: '2/3', row: '2/3' },
                  { key: 'channels',              label: '通路',         col: '4/5', row: '2/3' },
                  { key: 'costStructure',         label: '成本結構',     col: '1/4', row: '3/4' },
                  { key: 'revenueStreams',         label: '收益來源',     col: '4/6', row: '3/4' },
                ] as { key: keyof typeof report.businessModelCanvas; label: string; col: string; row: string }[]).map(cell => (
                  <div key={cell.key} className="ds-bmc-cell" style={{ gridColumn: cell.col, gridRow: cell.row }}>
                    <div className="ds-bmc-cell-label">{cell.label}</div>
                    <div className="ds-bmc-cell-content">
                      {report.businessModelCanvas[cell.key] && report.businessModelCanvas[cell.key].split('\n').map((line, i) => <div key={i} className="ds-bmc-line">• {line}</div>)}
                    </div>
                  </div>
                ))}
              </div>
            </div>
            <div className="ds-card">
              <div className="ds-card-title"><ShieldAlert size={18} style={{ color: '#f87171' }} /> SWOT 分析</div>
              <div className="ds-swot-grid">
                {([
                  { key: 'strengths',     label: '優勢 S', color: '#34d399', bg: 'rgba(52,211,153,0.08)' },
                  { key: 'weaknesses',    label: '劣勢 W', color: '#f87171', bg: 'rgba(248,113,113,0.08)' },
                  { key: 'opportunities', label: '機會 O', color: '#60a5fa', bg: 'rgba(96,165,250,0.08)' },
                  { key: 'threats',       label: '威脅 T', color: '#fbbf24', bg: 'rgba(251,191,36,0.08)' },
                ] as { key: keyof typeof report.swot; label: string; color: string; bg: string }[]).map(q => (
                  <div key={q.key} className="ds-swot-cell" style={{ background: q.bg, borderColor: q.color + '33' }}>
                    <div className="ds-swot-label" style={{ color: q.color }}>{q.label}</div>
                    {report.swot[q.key] && report.swot[q.key].map((item, i) => <div key={i} className="ds-swot-item">• {item}</div>)}
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {activeTab === 'market' && (
          <div className="ds-section-content">
            <div className="ds-card">
              <div className="ds-card-title"><TrendingUp size={18} style={{ color: '#34d399' }} /> 市場規模評估 TAM / SAM / SOM</div>
              <div className="ds-tam-grid">
                {([
                  { key: 'tam', label: 'TAM', sub: '總體潛在市場', color: '#60a5fa' },
                  { key: 'sam', label: 'SAM', sub: '可服務有效市場', color: '#34d399' },
                  { key: 'som', label: 'SOM', sub: '可獲得市場份額', color: '#a78bfa' },
                ] as { key: 'tam'|'sam'|'som'; label: string; sub: string; color: string }[]).map(m => (
                  <div key={m.key} className="ds-tam-cell" style={{ borderColor: m.color + '44' }}>
                    <div className="ds-tam-label" style={{ color: m.color }}>{m.label}</div>
                    <div className="ds-tam-sub">{m.sub}</div>
                    <div className="ds-tam-value">{report.marketSize[m.key]}</div>
                  </div>
                ))}
              </div>
              <p className="ds-muted-text" style={{ marginTop: 12 }}>{report.marketSize.description}</p>
            </div>
            <div className="ds-card">
              <div className="ds-card-title"><BarChart3 size={18} style={{ color: '#818cf8' }} /> STEEP 分析</div>
              <div className="ds-steep-grid">
                {([
                  { key: 'social',        label: '社會 S', icon: <Users size={16} /> },
                  { key: 'technological', label: '技術 T', icon: <Smartphone size={16} /> },
                  { key: 'economic',      label: '經濟 E', icon: <TrendingUp size={16} /> },
                  { key: 'environmental', label: '環境 E', icon: <Leaf size={16} /> },
                  { key: 'political',     label: '政治 P', icon: <Gavel size={16} /> },
                ] as { key: keyof typeof report.steep; label: string; icon: React.ReactNode }[]).map(s => (
                  <div key={s.key} className="ds-steep-cell">
                    <div className="ds-steep-label" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                      {s.icon} {s.label}
                    </div>
                    {report.steep[s.key] && report.steep[s.key].split('\n').map((line, i) => <div key={i} className="ds-steep-item">• {line}</div>)}
                  </div>
                ))}
              </div>
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
            <div className="ds-card">
              <div className="ds-card-title"><GitBranch size={18} style={{ color: '#f87171' }} /> Porter's 五力 + 第六力</div>
              <div className="ds-porter-grid">
                {([
                  { key: 'competitiveRivalry',        label: '現有競爭者',      icon: <Swords size={16} /> },
                  { key: 'threatOfNewEntrants',       label: '新進入者威脅',    icon: <ArrowRight size={16} /> },
                  { key: 'threatOfSubstitutes',       label: '替代品威脅',      icon: <RefreshCw size={16} /> },
                  { key: 'bargainingPowerOfBuyers',   label: '買方議價能力',    icon: <Users size={16} /> },
                  { key: 'bargainingPowerOfSuppliers', label: '供應商議價能力',  icon: <Truck size={16} /> },
                  { key: 'complementors',             label: '互補者（第六力）', icon: <Sparkles size={16} /> },
                ] as { key: keyof typeof report.portersFiveForces; label: string; icon: React.ReactNode }[]).map(f => (
                  <div key={f.key} className="ds-porter-cell">
                    <div className="ds-porter-label" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                      {f.icon} {f.label}
                    </div>
                    <p className="ds-muted-text">{report.portersFiveForces[f.key]}</p>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

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
                {['STAR', 'CASH_COW', 'QUESTION_MARK', 'DOG'].map((catKey) => {
                  const cat = catKey as BcgItem['category'];
                  const item = report.bcgMatrix.find(i => i.category === cat);
                  const cfg = bcgColor(cat);
                  
                  return (
                    <div key={cat} className="ds-bcg-cell" style={{ 
                      background: cfg.bg, 
                      borderColor: cfg.border + '55',
                      position: 'relative',
                      overflow: 'hidden',
                      opacity: item ? 1 : 0.7
                    }}>
                      <div className="ds-bcg-cat" style={{ color: cfg.border, display: 'flex', alignItems: 'center', gap: 6 }}>
                        {cfg.icon} {cfg.label}
                      </div>
                      <div className="ds-bcg-name">{item?.name || '暫無項目'}</div>
                      <p className="ds-muted-text" style={{ fontSize: '0.8rem' }}>
                        {item?.description || '本象限目前無相關產品或服務'}
                      </p>
                      {/* Decoration Icon */}
                      <div style={{
                        position: 'absolute', right: -10, bottom: -10,
                        opacity: 0.1, transform: 'rotate(-20deg)',
                        color: cfg.border
                      }}>
                        {React.cloneElement(cfg.icon as React.ReactElement, { size: 64 })}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
            <div className="ds-card">
              <div className="ds-card-title"><Compass size={18} style={{ color: '#a78bfa' }} /> SPAN 戰略定位分析</div>
              <div className="ds-span-grid">
                {([
                  { key: 'scope',       label: '範疇 Scope',      color: '#60a5fa' },
                  { key: 'positioning', label: '定位 Positioning', color: '#34d399' },
                  { key: 'advantage',   label: '優勢 Advantage',   color: '#fbbf24' },
                  { key: 'network',     label: '網絡 Network',     color: '#a78bfa' },
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
              <div className="ds-fan-container" style={{ display: 'flex', flexDirection: 'column', gap: '0' }}>
                {fanSections.map((section, i) => (
                  <React.Fragment key={i}>
                    <div className="ds-fan-box" style={{ 
                      background: '#1e293b', 
                      border: '1px solid #334155', 
                      borderRadius: '8px', 
                      padding: '16px',
                      position: 'relative'
                    }}>
                      <div className="ds-fan-label" style={{ 
                        marginBottom: '8px', fontWeight: 600, color: '#f8fafc',
                        display: 'flex', alignItems: 'center', gap: '8px'
                      }}>
                        <div style={{ width: 6, height: 6, borderRadius: '50%', background: '#34d399' }} />
                        {fanLabels[i] || `面向 ${i + 1}`}
                      </div>
                      {section.replace(/^(財務面|客戶面|內部流程面|學習與成長面)[:：]?\s*/, '').split(/[\\n]+/).filter(Boolean).map((line, j) => (
                        <div key={j} className="ds-bullet-line" style={{ paddingLeft: '14px' }}>
                          <ChevronRight size={13} style={{ color: '#64748b', flexShrink: 0, marginTop: 3 }} />
                          <span className="ds-muted-text" style={{ fontSize: '0.9rem' }}>{line}</span>
                        </div>
                      ))}
                    </div>
                    {i < fanSections.length - 1 && (
                      <div style={{ display: 'flex', justifyContent: 'center', padding: '4px 0' }}>
                        <ArrowDown size={20} style={{ color: '#475569' }} />
                      </div>
                    )}
                  </React.Fragment>
                ))}
              </div>
            </div>
          </div>
        )}

        {activeTab === 'action' && (
          <div className="ds-section-content">
            <div className="ds-card">
              <div className="ds-card-title"><Target size={18} style={{ color: '#f87171' }} /> 優先級地圖</div>
              <div className="ds-priority-table">
                <div className="ds-priority-header">
                  <span>#</span><span>行動項目</span><span>影響力</span><span>成本</span>
                </div>
                {[...report.priorityMap].sort((a, b) => a.priority - b.priority).map((item, i) => (
                  <div key={i} className="ds-priority-row" style={{ background: impactBg(item.impact) }}>
                    <span className="ds-priority-num">{item.priority}</span>
                    <span className="ds-priority-action">{item.action}</span>
                    <span className="ds-priority-badge" style={{ color: priorityColor(item.impact) }}>{item.impact}</span>
                    <span className="ds-priority-badge" style={{ color: priorityColor(item.cost) }}>{item.cost}</span>
                  </div>
                ))}
              </div>
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
            {/* 新增 Continue to Iterate 區塊 */}
            <div className="ds-card">
              <div className="ds-card-title"><RefreshCw size={18} style={{ color: '#34d399' }} /> 持續迭代建議</div>
              <div className="ds-iterate-content" style={{ marginTop: 12 }}>
                {(report.continueToIterate || result.continueToIterate) ? (report.continueToIterate || result.continueToIterate).split('\n').map((line, i) => (
                  <div key={i} className="ds-bullet-line" style={{ marginBottom: 6 }}>
                    <ChevronRight size={14} style={{ color: '#34d399', flexShrink: 0, marginTop: 4 }} />
                    <span style={{ color: '#e2e8f0', lineHeight: 1.6 }}>{line}</span>
                  </div>
                )) : <div className="ds-muted-text">暫無迭代建議</div>}
              </div>
            </div>
          </div>
        )}

        {activeTab === 'full' && (
          <div className="cf-full-result">
            <div className="cf-final-score-banner">
              <div className="cf-final-score-left">
                <Trophy size={20} style={{ color: '#fbbf24' }} />
                <div>
                  <div className="cf-final-score-label">深度財務與戰略評估</div>
                  <div className="cf-final-score-sub">基於完整商業分析模型</div>
                </div>
              </div>
              <div className="cf-final-score-value" style={{ color: scoreColorFunc(result.successProbability) }}>
                {result.successProbability}%
              </div>
            </div>

            <div className="cf-full-tab-nav">
              {([
                { key: 'finance',     label: '財務預測',    icon: <DollarSign size={13} /> },
                { key: 'personas',    label: 'AI 虛擬董事會', icon: <Brain size={13} /> },
                { key: 'competitors', label: '競爭態勢',    icon: <Swords size={13} /> },
                { key: 'verdicts',    label: '三種觀點',    icon: <MessageSquare size={13} /> },
              ] as { key: typeof subTab; label: string; icon: React.ReactNode }[]).map(t => (
                <button key={t.key} className={`cf-full-tab ${subTab === t.key ? 'cf-full-tab-active' : ''}`} onClick={() => setSubTab(t.key)}>
                  {t.icon} {t.label}
                </button>
              ))}
            </div>

            {subTab === 'finance' && (
              <div className="cf-tab-content">
                <div className="cf-tab-card">
                  <div className="cf-tab-card-title" style={{ marginBottom: '24px' }}><DollarSign size={16} style={{ color: '#34d399' }} /> 3 年財務預測</div>
                  <div style={{ height: 220 }}>
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={result.financials} margin={{ top: 8, right: 16, left: 0, bottom: 0 }}>
                        <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                        <XAxis dataKey="year" tick={{ fill: '#64748b', fontSize: 11 }} axisLine={false} />
                        <YAxis tick={{ fill: '#64748b', fontSize: 10 }} axisLine={false} tickLine={false}
                          tickFormatter={(v: number) => v >= 1000 ? `${(v / 1000).toFixed(0)}K` : String(Number(v)) } />
                        <Tooltip contentStyle={{ background: '#1e293b', border: '1px solid #334155', borderRadius: 8, color: '#f8fafc', fontSize: 12 }}
                          formatter={(v: number) => [`NT$ ${v.toLocaleString()}`, '']} />
                        <Legend wrapperStyle={{ color: '#94a3b8', fontSize: 12 }} />
                        <Line type="monotone" dataKey="revenue" name="營收" stroke="#34d399" strokeWidth={2} dot={{ fill: '#34d399', r: 3 }} />
                        <Line type="monotone" dataKey="profit"  name="利潤" stroke="#60a5fa" strokeWidth={2} dot={{ fill: '#60a5fa', r: 3 }} />
                        <Line type="monotone" dataKey="costs"   name="成本" stroke="#f87171" strokeWidth={2} dot={{ fill: '#f87171', r: 3 }} strokeDasharray="4 2" />
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                  <div className="cf-break-even">
                    <CheckCircle2 size={14} style={{ color: '#34d399' }} />
                    <span>預估損益兩平點：<strong style={{ color: '#34d399' }}>{result.breakEvenPoint}</strong></span>
                  </div>
                </div>
              </div>
            )}

            {subTab === 'personas' && (
              <div className="cf-tab-content">
                <div className="cf-tab-card">
                  <div className="cf-tab-card-title" style={{ marginBottom: '24px' }}><Brain size={16} style={{ color: '#a78bfa' }} /> AI 虛擬董事會評估</div>
                  <div className="cf-personas-grid" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(300px, 1fr))', gap: '24px' }}>
                    {result.personaEvaluations.map((p, i) => (
                      <div key={i} className="cf-persona-card" style={{ 
                        background: 'rgba(30, 41, 59, 0.6)', 
                        border: '1px solid rgba(148, 163, 184, 0.2)', 
                        borderRadius: '16px', 
                        padding: '20px',
                        display: 'flex',
                        flexDirection: 'column',
                        gap: '16px',
                        boxShadow: '0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06)'
                      }}>
                        <div className="cf-persona-top" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid rgba(255,255,255,0.1)', paddingBottom: '16px' }}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                            <div className="cf-persona-icon" style={{ 
                              width: '48px', height: '48px', borderRadius: '12px', 
                              background: 'rgba(56, 189, 248, 0.15)', color: '#38bdf8', 
                              display: 'flex', alignItems: 'center', justifyContent: 'center' 
                            }}>
                              <Users size={24} />
                            </div>
                            <div>
                                <div className="cf-persona-role" style={{ color: '#f1f5f9', fontWeight: 600, fontSize: '1.05rem' }}>{p.role}</div>
                                <div style={{ fontSize: '0.8rem', color: '#94a3b8' }}>董事會成員</div>
                            </div>
                          </div>
                          <div className="cf-persona-score" style={{ 
                            fontSize: '1.5rem', fontWeight: 800, 
                            color: scoreColorFunc(p.score),
                            background: 'rgba(15, 23, 42, 0.5)',
                            padding: '4px 12px',
                            borderRadius: '8px'
                          }}>{p.score}</div>
                        </div>
                        
                        <div className="cf-persona-quote" style={{ 
                          fontStyle: 'italic', color: '#cbd5e1', fontSize: '0.95rem', 
                          background: 'rgba(51, 65, 85, 0.3)', padding: '12px', borderRadius: '8px',
                          borderLeft: '4px solid #38bdf8'
                        }}>
                          「{p.keyQuote}」
                        </div>
                        
                        <div className="cf-persona-perspective" style={{ fontSize: '1rem', color: '#cbd5e1', lineHeight: 1.6 }}>
                          {p.perspective}
                        </div>
                        
                        <div className="cf-persona-concern" style={{ 
                          marginTop: 'auto', paddingTop: '16px', 
                          borderTop: '1px solid rgba(255,255,255,0.1)', display: 'flex', gap: '10px', alignItems: 'start'
                        }}>
                          <ShieldAlert size={16} style={{ color: '#ef4444', flexShrink: 0, marginTop: 3 }} />
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                            <span style={{ 
                              fontSize: '0.75rem', fontWeight: 700,
                              color: '#ef4444', textTransform: 'uppercase'
                            }}>潛在顧慮</span>
                            <span style={{ fontSize: '0.9rem', color: '#94a3b8', lineHeight: 1.5 }}>{p.concern}</span>
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            )}

            {subTab === 'competitors' && (
              <div className="cf-tab-content">
                <div className="cf-tab-card">
                  <div className="cf-tab-card-title" style={{ marginBottom: '24px' }}><Swords size={16} style={{ color: '#f87171' }} /> 競爭態勢深度分析</div>
                  <div className="cf-competitors-grid" style={{ 
                    display: 'grid', gap: '0', 
                    border: '1px solid rgba(148, 163, 184, 0.2)',
                    borderRadius: '8px', overflow: 'hidden'
                  }}>
                     {/* Header */}
                     <div style={{ 
                       display: 'grid', gridTemplateColumns: '20% 40% 40%', 
                       background: 'rgba(30, 41, 59, 0.9)', 
                       borderBottom: '1px solid rgba(148, 163, 184, 0.4)',
                       fontSize: '0.9rem', color: '#94a3b8', fontWeight: 700
                     }}>
                       <div style={{ padding: '12px 16px', borderRight: '1px solid rgba(148, 163, 184, 0.2)' }}>競爭對手</div>
                       <div style={{ padding: '12px 16px', borderRight: '1px solid rgba(148, 163, 184, 0.2)' }}>核心優勢</div>
                       <div style={{ padding: '12px 16px' }}>劣勢與破綻</div>
                     </div>
                    {result.competitors.map((c, i) => (
                      <div key={i} className="cf-competitor-item" style={{ 
                        display: 'grid', gridTemplateColumns: '20% 40% 40%', 
                        borderBottom: i === result.competitors.length - 1 ? 'none' : '1px solid rgba(148, 163, 184, 0.4)',
                        alignItems: 'stretch', fontSize: '0.95rem',
                        background: i % 2 === 0 ? 'rgba(30, 41, 59, 0.3)' : 'transparent'
                      }}>
                        <div className="cf-competitor-name" style={{ 
                          color: '#f8fafc', fontWeight: 600, padding: '16px', 
                          borderRight: '1px solid rgba(148, 163, 184, 0.2)',
                          display: 'flex', alignItems: 'center'
                        }}>{c.name}</div>
                        <div className="cf-competitor-str" style={{ 
                          padding: '16px', color: '#34d399', lineHeight: 1.5,
                          borderRight: '1px solid rgba(148, 163, 184, 0.2)' 
                        }}>{c.strength}</div>
                        <div className="cf-competitor-weak" style={{ 
                          padding: '16px', color: '#f87171', lineHeight: 1.5
                        }}>{c.weakness}</div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            )}

            {subTab === 'verdicts' && (
              <div className="cf-tab-content">
                <div className="cf-tab-card">
                  <div className="cf-tab-card-title" style={{ marginBottom: '24px' }}><MessageSquare size={16} style={{ color: '#60a5fa' }} /> 三種戰略觀點</div>
                  <div className="cf-verdicts-grid" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(300px, 1fr))', gap: '24px' }}>
                    
                    {/* 激進派 */}
                    <div className="cf-verdict" style={{ 
                      background: 'linear-gradient(145deg, rgba(239, 68, 68, 0.05) 0%, rgba(239, 68, 68, 0.1) 100%)', 
                      border: '1px solid rgba(239, 68, 68, 0.3)', 
                      borderRadius: '16px', padding: '24px',
                      boxShadow: '0 4px 6px -1px rgba(239, 68, 68, 0.05)'
                    }}>
                      <h4 style={{ color: '#ef4444', marginBottom: '16px', display: 'flex', alignItems: 'center', gap: '10px', fontSize: '1.1rem', fontWeight: 700 }}>
                        <div style={{ padding: 8, background: 'rgba(239,68,68,0.2)', borderRadius: '8px' }}><Zap size={20} /></div>
                        激進派觀點
                      </h4>
                      <p style={{ color: '#fecaca', lineHeight: 1.8, fontSize: '1rem', whiteSpace: 'pre-wrap' }}>{result.finalVerdicts.aggressive}</p>
                    </div>

                    {/* 均衡派 */}
                    <div className="cf-verdict" style={{ 
                      background: 'linear-gradient(145deg, rgba(59, 130, 246, 0.05) 0%, rgba(59, 130, 246, 0.1) 100%)', 
                      border: '1px solid rgba(59, 130, 246, 0.3)', 
                      borderRadius: '16px', padding: '24px',
                      boxShadow: '0 4px 6px -1px rgba(59, 130, 246, 0.05)'
                    }}>
                      <h4 style={{ color: '#3b82f6', marginBottom: '16px', display: 'flex', alignItems: 'center', gap: '10px', fontSize: '1.1rem', fontWeight: 700 }}>
                        <div style={{ padding: 8, background: 'rgba(59,130,246,0.2)', borderRadius: '8px' }}><Target size={20} /></div>
                        均衡派觀點
                      </h4>
                      <p style={{ color: '#bfdbfe', lineHeight: 1.8, fontSize: '1rem', whiteSpace: 'pre-wrap' }}>{result.finalVerdicts.balanced}</p>
                    </div>

                    {/* 保守派 */}
                    <div className="cf-verdict" style={{ 
                      background: 'linear-gradient(145deg, rgba(16, 185, 129, 0.05) 0%, rgba(16, 185, 129, 0.1) 100%)', 
                      border: '1px solid rgba(16, 185, 129, 0.3)', 
                      borderRadius: '16px', padding: '24px',
                      boxShadow: '0 4px 6px -1px rgba(16, 185, 129, 0.05)'
                    }}>
                      <h4 style={{ color: '#10b981', marginBottom: '16px', display: 'flex', alignItems: 'center', gap: '10px', fontSize: '1.1rem', fontWeight: 700 }}>
                        <div style={{ padding: 8, background: 'rgba(16,185,129,0.2)', borderRadius: '8px' }}><ShieldAlert size={20} /></div>
                        保守派觀點
                      </h4>
                      <p style={{ color: '#bbf7d0', lineHeight: 1.8, fontSize: '1rem', whiteSpace: 'pre-wrap' }}>{result.finalVerdicts.conservative}</p>
                    </div>

                  </div>
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};

export default AnalysisDashboard;
import React, { useState } from 'react';
import { AnalysisResult, PersonaEvaluation, FinancialYear, RoadmapItem } from '../types';
import { TrendingUp, Users, ShieldAlert, Target, DollarSign, BarChart3, AlertTriangle, Briefcase, User, Factory, Sword, FileDown, Loader2, FileText } from 'lucide-react';
import { AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import { generatePptx } from '../services/pptxService';
import { summarizeForSlides, generateThemeImage, generateProposalTitle } from '../services/geminiService';

interface Props { result: AnalysisResult; onReset: () => void; mode?: string; }

const getScoreClass = (score: number) => {
  if (score >= 80) return 'score-green';
  if (score >= 60) return 'score-blue';
  if (score >= 40) return 'score-yellow';
  return 'score-red';
};

const AnalysisDashboard: React.FC<Props> = ({ result, onReset, mode }) => {
  const [isGeneratingPpt, setIsGeneratingPpt] = useState(false);
  const [pptError, setPptError] = useState<string | null>(null);
  const [pptStage, setPptStage] = useState<string>('');
  const [proposalTitle, setProposalTitle] = useState<string>('');

  // 強制將 financials 每個欄位轉成純 number，防止 AI 回傳字串導致 Recharts 無法渲染
  const chartData = (result.financials || []).map(f => ({
    year: f.year ?? '',
    revenue: Number(f.revenue) || 0,
    costs: Number(f.costs) || 0,
    profit: Number(f.profit) || 0,
  }));

  const hasFinancialData = chartData.length > 0 && chartData.some(f => f.revenue !== 0 || f.costs !== 0 || f.profit !== 0);

  // 將大數字格式化為 K / M，讓 Y 軸刻度可讀
  const formatYAxis = (value: number) => {
    if (Math.abs(value) >= 1_000_000) return `${(value / 1_000_000).toFixed(1)}M`;
    if (Math.abs(value) >= 1_000) return `${(value / 1_000).toFixed(0)}K`;
    return String(value);
  };

  const handleGeneratePpt = async () => {
    setIsGeneratingPpt(true);
    setPptError(null);
    setPptStage('精簡內文中...');
    try {
      // 第一步：用 LLM 精簡化所有投影片內文
      const summarized = await summarizeForSlides(result);

      setPptStage('生成提案標題...');
      const title = await generateProposalTitle(summarized.executiveSummary);
      setProposalTitle(title);

      // 第二步：根據提案實際內容用 Imagen 生成主題插圖
      setPptStage('生成主題插圖...');
      const themeImage = await generateThemeImage(
        summarized.executiveSummary,
        summarized.marketAnalysis.description
      );

      // 第三步：產生 PPTX（傳入精簡化後的內容與 AI 圖片）
      setPptStage('組裝簡報中...');
      await generatePptx(summarized, themeImage, title);
    } catch (err: any) {
      setPptError(err.message || '產生失敗，請重試');
    } finally {
      setIsGeneratingPpt(false);
      setPptStage('');
    }
  };

  // PDF 功能已移除，因前端無法產生真正 PDF

  return (
    <div className="dashboard">
      <div className="dash-header">
        <div>
          <h1>分析報告</h1>
          <p>全方位 360° 評估</p>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: '1rem', flexWrap: 'wrap' }}>
          <div>
            <div className="score-label">成功機率</div>
            <div className={`score-value ${getScoreClass(result.successProbability)}`}>{result.successProbability}%</div>
          </div>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
            <button className="reset-btn" onClick={onReset}>重新分析</button>
            <button
              className={`ppt-btn${isGeneratingPpt ? ' ppt-btn-loading' : ''}`}
              onClick={handleGeneratePpt}
              disabled={isGeneratingPpt}
              title="下載 PowerPoint 簡報"
            >
              {isGeneratingPpt
                ? <><Loader2 size={15} className="spin-icon" /> PPT 產生中...</>
                : <><FileDown size={15} /> 下載 .pptx</>}
            </button>
          </div>
        </div>
      </div>

      {pptError && (
        <div className="ppt-error-bar">⚠ {pptError}</div>
      )}

      {(isGeneratingPpt) && pptStage && (
        <div className="ppt-progress-bar">
          <div className="ppt-progress-inner">
            <Loader2 size={18} className="spin-icon" />
            <span>{pptStage}</span>
          </div>
        </div>
      )}

      <div className="card section-gap">
        <div className="card-title"><Target size={20} style={{color:'#60a5fa'}} /> 執行摘要</div>
        <p className="summary-text">{result.executiveSummary}</p>
      </div>

      <div className="metrics-grid">
        <div className="metric-card">
          <div className="metric-label"><TrendingUp size={16} /> 市場分析</div>
          <div className="metric-big">{result.marketAnalysis.size}</div>
          <div className="metric-sub-green">總體潛在市場 (TAM)</div>
          <div className="metric-big" style={{marginTop:'1rem'}}>{result.marketAnalysis.growthRate}</div>
          <div className="metric-sub-blue">年均複合成長率 (CAGR)</div>
          <div className="metric-divider" />
          <p style={{fontSize:'0.75rem', color:'#94a3b8'}}>{result.marketAnalysis.description}</p>
        </div>

        <div className="metric-card">
          <div className="metric-label"><ShieldAlert size={16} /> 主要風險</div>
          {result.risks.slice(0, 3).map((risk, i) => (
            <div className="risk-item" key={i}>
              <AlertTriangle size={16} className={risk.impact === 'High' ? 'icon-red' : 'icon-yellow'} style={{color: risk.impact === 'High' ? '#ef4444' : '#fbbf24', flexShrink:0, marginTop:2}} />
              <div>
                <div className="risk-title">{risk.risk}</div>
                <div className="risk-mit">{risk.mitigation}</div>
              </div>
            </div>
          ))}
        </div>

        <div className="metric-card">
          <div className="metric-label"><DollarSign size={16} /> 財務展望</div>
          <div className="metric-big">損益平衡點</div>
          <div className="metric-sub-blue">{result.breakEvenPoint}</div>
          <div className="metric-divider" />
          <div style={{fontSize:'0.875rem', color:'#94a3b8', marginBottom:4}}>預估第三年營收</div>
          <div className="metric-big">${(result.financials[2]?.revenue || 0).toLocaleString()}</div>
        </div>
      </div>

      <div className="charts-grid">
        <div className="chart-card">
          <div className="chart-title"><BarChart3 size={20} style={{color:'#34d399'}} /> 財務預測</div>
          <div className="chart-wrap">
            {hasFinancialData ? (
              <ResponsiveContainer width="100%" height="100%">
                <AreaChart data={chartData}>
                  <defs>
                    <linearGradient id="colorRev" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor="#3b82f6" stopOpacity={0.8}/>
                      <stop offset="95%" stopColor="#3b82f6" stopOpacity={0}/>
                    </linearGradient>
                    <linearGradient id="colorProf" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor="#10b981" stopOpacity={0.8}/>
                      <stop offset="95%" stopColor="#10b981" stopOpacity={0}/>
                    </linearGradient>
                  </defs>
                  <CartesianGrid strokeDasharray="3 3" stroke="#334155" />
                  <XAxis dataKey="year" stroke="#94a3b8" />
                  <YAxis stroke="#94a3b8" tickFormatter={formatYAxis} />
                  <Tooltip contentStyle={{ backgroundColor: '#1e293b', borderColor: '#334155', color: '#f8fafc' }} />
                  <Legend />
                  <Area type="monotone" dataKey="revenue" stroke="#3b82f6" fillOpacity={1} fill="url(#colorRev)" name="營收" />
                  <Area type="monotone" dataKey="profit" stroke="#10b981" fillOpacity={1} fill="url(#colorProf)" name="淨利" />
                  <Area type="monotone" dataKey="costs" stroke="#ef4444" fill="transparent" strokeDasharray="5 5" name="成本" />
                </AreaChart>
              </ResponsiveContainer>
            ) : (
              <div className="empty-state">⚠ AI 未提供足夠資訊以產生財務預測圖表</div>
            )}
          </div>
        </div>

        <div className="chart-card">
          <div className="chart-title"><TrendingUp size={20} style={{color:'#818cf8'}} /> 策略路線圖</div>
          <div className="roadmap-list">
            {result.roadmap && result.roadmap.length > 0 ? result.roadmap.map((item, idx) => (
              <div className="roadmap-item" key={idx}>
                <div className="roadmap-dot" />
                <div className="roadmap-phase">{item.timeframe} - {item.phase}</div>
                <div className="roadmap-box">
                  <div className="roadmap-cols">
                    <div style={{flex:1}}>
                      <span className="roadmap-col-label">產品 (Product)</span>
                      <span className="roadmap-col-val">{item.product || '—'}</span>
                    </div>
                    <div style={{flex:1}}>
                      <span className="roadmap-col-label">技術 (Technology)</span>
                      <span className="roadmap-col-val">{item.technology || '—'}</span>
                    </div>
                  </div>
                </div>
              </div>
            )) : (
              <div className="empty-state">⚠ AI 未提供策略路線圖資料</div>
            )}
          </div>
        </div>
      </div>

      <div className="personas-section">
        <h3><Users size={24} style={{color:'#22d3ee'}} /> AI 虛擬董事會</h3>
        <div className="personas-grid">
          {result.personaEvaluations.map((persona, idx) => <PersonaCard key={idx} persona={persona} />)}
        </div>
      </div>

      <div className="competitors-card section-gap">
        <div className="card-title"><Sword size={20} style={{color:'#f87171'}} /> 競爭態勢分析</div>
        <div className="competitors-grid">
          {result.competitors.map((comp, idx) => (
            <div className="competitor-item" key={idx}>
              <div className="competitor-name">{comp.name}</div>
              <div className="comp-strength">優勢：</div>
              <div className="comp-text">{comp.strength}</div>
              <div className="comp-weakness">劣勢：</div>
              <div className="comp-text">{comp.weakness}</div>
            </div>
          ))}
        </div>
      </div>

      <div className="verdicts-grid">
        <div className="verdict-aggressive"><h4>激進觀點</h4><p className="verdict-text">{result.finalVerdicts.aggressive}</p></div>
        <div className="verdict-balanced"><h4>平衡觀點</h4><p className="verdict-text">{result.finalVerdicts.balanced}</p></div>
        <div className="verdict-conservative"><h4>保守觀點</h4><p className="verdict-text">{result.finalVerdicts.conservative}</p></div>
      </div>

      <div className="card section-gap" style={{ marginTop: '2rem' }}>
        <div className="card-title"><TrendingUp size={20} style={{color:'#10b981'}} /> 持續迭代建議</div>
        <p className="summary-text" style={{ whiteSpace: 'pre-line' }}>{result.continueToIterate}</p>
      </div>

      {/* PPT CTA Banner at bottom */}
      <div className="ppt-cta-banner">
        <div className="ppt-cta-left">
          <FileDown size={28} style={{color:'#3b82f6'}} />
          <div>
            <div className="ppt-cta-title">下載簡報</div>
            <div className="ppt-cta-sub">
              <strong style={{color:'#60a5fa'}}>.pptx</strong> — 深色風格 10 頁投資人簡報，適合桌機開啟與演示。<br/>
              {/* PDF 已移除 */}
            </div>
          </div>
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
          <button
            className={`ppt-btn ppt-btn-large${isGeneratingPpt ? ' ppt-btn-loading' : ''}`}
            onClick={handleGeneratePpt}
            disabled={isGeneratingPpt}
          >
            {isGeneratingPpt
              ? <><Loader2 size={18} className="spin-icon" /> 產生中...</>
              : <><FileDown size={18} /> 下載 .pptx 簡報</>}
          </button>
          {/* PDF 按鈕已移除 */}
        </div>
      </div>
    </div>
  );
};

const PersonaCard: React.FC<{ persona: PersonaEvaluation }> = ({ persona }) => {
  const getIcon = (icon: string) => {
    switch (icon.toLowerCase()) {
      case 'money': case 'investor': return <DollarSign size={20} style={{color:'#34d399'}} />;
      case 'user': case 'consumer': case 'customer': return <User size={20} style={{color:'#60a5fa'}} />;
      case 'shield': return <ShieldAlert size={20} style={{color:'#fb923c'}} />;
      case 'briefcase': case 'employee': return <Briefcase size={20} style={{color:'#c084fc'}} />;
      case 'factory': case 'supplier': return <Factory size={20} style={{color:'#fbbf24'}} />;
      default: return <Sword size={20} style={{color:'#f87171'}} />;
    }
  };
  return (
    <div className="persona-card">
      <div className="persona-top">
        <div className="persona-identity">
          <div className="persona-icon-box">{getIcon(persona.icon || persona.role)}</div>
          <div>
            <div className="persona-name">{persona.role}</div>
            <div className="persona-role-sub">董事會成員</div>
          </div>
        </div>
        <div className={`score-value ${getScoreClass(persona.score)}`} style={{fontSize:'1.25rem'}}>{persona.score}</div>
      </div>
      <div className="persona-quote">"{persona.keyQuote}"</div>
      <p className="persona-perspective">{persona.perspective}</p>
      <div className="persona-footer">
        <div className="persona-concern-label">主要擔憂</div>
        <div className="persona-concern">{persona.concern}</div>
      </div>
    </div>
  );
};

export default AnalysisDashboard;
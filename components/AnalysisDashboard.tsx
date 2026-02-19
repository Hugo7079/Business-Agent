import React, { useState } from 'react';
import { AnalysisResult, PersonaEvaluation, FinancialYear, RoadmapItem } from '../types';
import { TrendingUp, Users, ShieldAlert, Target, DollarSign, BarChart3, AlertTriangle, Briefcase, User, Factory, Sword, FileDown, Loader2 } from 'lucide-react';
import { AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import { generatePptNotebook } from '../services/geminiService';

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

  const handleGeneratePpt = async () => {
    setIsGeneratingPpt(true);
    setPptError(null);
    try {
      const pythonCode = await generatePptNotebook(result);

      // Build a .ipynb notebook with 2 cells: install deps + the generated code
      const notebook = {
        nbformat: 4,
        nbformat_minor: 5,
        metadata: { kernelspec: { display_name: 'Python 3', language: 'python', name: 'python3' }, language_info: { name: 'python', version: '3.10.0' } },
        cells: [
          {
            cell_type: 'markdown',
            metadata: {},
            source: ['# OmniView AI — 投資提案 PPT 產生器\n', '執行下方程式碼即可在同目錄產生 `OmniView_投資提案.pptx`']
          },
          {
            cell_type: 'code',
            metadata: {},
            execution_count: null,
            outputs: [],
            source: ['# 安裝必要套件（第一次執行時需要）\n', 'import subprocess, sys\n', 'subprocess.check_call([sys.executable, "-m", "pip", "install", "python-pptx", "matplotlib", "-q"])']
          },
          {
            cell_type: 'code',
            metadata: {},
            execution_count: null,
            outputs: [],
            source: pythonCode.split('\n').map((line, i, arr) => i < arr.length - 1 ? line + '\n' : line)
          }
        ]
      };

      const blob = new Blob([JSON.stringify(notebook, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'OmniView_PPT產生器.ipynb';
      a.click();
      URL.revokeObjectURL(url);
    } catch (err: any) {
      setPptError(err.message || '產生失敗，請重試');
    } finally {
      setIsGeneratingPpt(false);
    }
  };

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
              title="用 AI 產生投資人簡報 PPT"
            >
              {isGeneratingPpt
                ? <><Loader2 size={15} className="spin-icon" /> AI 撰寫中...</>
                : <><FileDown size={15} /> 產生 PPT 簡報</>}
            </button>
          </div>
        </div>
      </div>

      {pptError && (
        <div className="ppt-error-bar">⚠ {pptError}</div>
      )}

      {isGeneratingPpt && (
        <div className="ppt-progress-bar">
          <div className="ppt-progress-inner">
            <Loader2 size={18} className="spin-icon" />
            <span>AI 正在根據分析結果撰寫完整的 Python 簡報程式碼，請稍候（約 30-60 秒）...</span>
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
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={result.financials}>
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
                <YAxis stroke="#94a3b8" />
                <Tooltip contentStyle={{ backgroundColor: '#1e293b', borderColor: '#334155', color: '#f8fafc' }} />
                <Legend />
                <Area type="monotone" dataKey="revenue" stroke="#3b82f6" fillOpacity={1} fill="url(#colorRev)" name="營收" />
                <Area type="monotone" dataKey="profit" stroke="#10b981" fillOpacity={1} fill="url(#colorProf)" name="淨利" />
                <Area type="monotone" dataKey="costs" stroke="#ef4444" fill="transparent" strokeDasharray="5 5" name="成本" />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="chart-card">
          <div className="chart-title"><TrendingUp size={20} style={{color:'#818cf8'}} /> 策略路線圖</div>
          <div className="roadmap-list">
            {result.roadmap.map((item, idx) => (
              <div className="roadmap-item" key={idx}>
                <div className="roadmap-dot" />
                <div className="roadmap-phase">{item.timeframe} - {item.phase}</div>
                <div className="roadmap-box">
                  <div className="roadmap-cols">
                    <div style={{flex:1}}>
                      <span className="roadmap-col-label">產品 (Product)</span>
                      <span className="roadmap-col-val">{item.product}</span>
                    </div>
                    <div style={{flex:1}}>
                      <span className="roadmap-col-label">技術 (Technology)</span>
                      <span className="roadmap-col-val">{item.technology}</span>
                    </div>
                  </div>
                </div>
              </div>
            ))}
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

      {/* PPT CTA Banner at bottom */}
      <div className="ppt-cta-banner">
        <div className="ppt-cta-left">
          <FileDown size={28} style={{color:'#3b82f6'}} />
          <div>
            <div className="ppt-cta-title">下載 AI 投資人簡報</div>
            <div className="ppt-cta-sub">
              AI 將為每張投影片生成專屬背景圖與插圖（使用 Imagen 3），並結合 matplotlib 數據圖表與格式化表格，產出完整 10 頁的高質感簡報。下載 .ipynb 後在 Jupyter 執行即可。
            </div>
          </div>
        </div>
        <button
          className={`ppt-btn ppt-btn-large${isGeneratingPpt ? ' ppt-btn-loading' : ''}`}
          onClick={handleGeneratePpt}
          disabled={isGeneratingPpt}
        >
          {isGeneratingPpt
            ? <><Loader2 size={18} className="spin-icon" /> AI 撰寫中...</>
            : <><FileDown size={18} /> 產生並下載 .ipynb</>}
        </button>
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
import React, { useState, useRef, useEffect } from 'react';
import {
  InitialScanResult, Phase2Message, Phase2Response,
  BusinessInput, Phase3DeepReport, AnalysisResult,
} from '../types';
import { runPhase2Reply, runPhase3DeepReport } from '../services/geminiService';
import {
  Sparkles, Mic, StopCircle, AlertTriangle,
  Send, Zap, ArrowRight, CheckCircle2, BarChart3,
  TrendingUp, X, Users, Swords, MessageSquare
} from 'lucide-react';
import {
  ResponsiveContainer,
  Tooltip, BarChart, Bar, XAxis, YAxis, Cell
} from 'recharts';

// // ─── helpers ──────────────────────────────────────────────────────────────────

// const priorityColor = (level: 'HIGH' | 'MEDIUM' | 'LOW') =>
//   level === 'HIGH' ? '#ef4444' : level === 'MEDIUM' ? '#fbbf24' : '#34d399';

// const impactBg = (impact: 'HIGH' | 'MEDIUM' | 'LOW') =>
//   impact === 'HIGH' ? 'rgba(239,68,68,0.08)' : impact === 'MEDIUM' ? 'rgba(251,191,36,0.08)' : 'rgba(52,211,153,0.08)';

// const bcgColor = (cat: BcgItem['category']) => {
//   switch (cat) {
//     case 'STAR':          return { bg: 'rgba(251,191,36,0.12)', border: '#fbbf24', label: '⭐ 明星' };
//     case 'CASH_COW':      return { bg: 'rgba(52,211,153,0.12)', border: '#34d399', label: '🐄 金牛' };
//     case 'QUESTION_MARK': return { bg: 'rgba(96,165,250,0.12)', border: '#60a5fa', label: '❓ 問號' };
//     case 'DOG':           return { bg: 'rgba(148,163,184,0.12)', border: '#94a3b8', label: '🐕 瘦狗' };
//   }
// };

const threatToScore = (threat: string): number => {
  const t = threat.toLowerCase();
  if (t.includes('極高') || t.includes('very high')) return 95;
  if (t.includes('高') || t.includes('high')) return 78;
  if (t.includes('中') || t.includes('medium') || t.includes('moderate')) return 55;
  if (t.includes('低') || t.includes('low')) return 30;
  return 60;
};

const threatBarColor = (score: number) =>
  score >= 80 ? '#ef4444' : score >= 60 ? '#fbbf24' : '#34d399';

const blobToBase64 = (blob: Blob): Promise<string> =>
  new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onloadend = () => res((reader.result as string).split(',')[1]);
    reader.onerror = rej;
    reader.readAsDataURL(blob);
  });

// ─── Initial Scan Result Panel ────────────────────────────────────────────────

const InitialScanPanel: React.FC<{ scan: InitialScanResult }> = ({ scan }) => {
  const scoreColor =
    scan.successProbability >= 65 ? '#34d399' :
    scan.successProbability >= 40 ? '#fbbf24' : '#ef4444';

  const competitorChartData = scan.competitors.map(c => ({
    name: c.name.length > 10 ? c.name.slice(0, 10) + '…' : c.name,
    fullName: c.name,
    threat: c.threat,
    score: threatToScore(c.threat),
  }));

  return (
    <div className="cf-scan-panel">
      <div className="cf-scan-header">
        <div className="cf-score-ring" style={{ '--score-color': scoreColor } as React.CSSProperties}>
          <svg viewBox="0 0 80 80" className="cf-score-svg">
            <circle cx="40" cy="40" r="34" className="cf-score-track" />
            <circle cx="40" cy="40" r="34" className="cf-score-fill"
              style={{
                stroke: scoreColor,
                strokeDasharray: `${2 * Math.PI * 34}`,
                strokeDashoffset: `${2 * Math.PI * 34 * (1 - scan.successProbability / 100)}`,
              }}
            />
          </svg>
          <div className="cf-score-label">
            <span className="cf-score-num" style={{ color: scoreColor }}>{scan.successProbability}%</span>
            <span className="cf-score-sub">初步成功率</span>
          </div>
        </div>
        <div className="cf-understanding">
          <div className="cf-understanding-title"><Sparkles size={14} /> OmniView 的理解</div>
          <p className="cf-understanding-text">{scan.coreUnderstanding}</p>
        </div>
      </div>
      <div className="cf-market-row">
        <div className="cf-market-cell">
          <TrendingUp size={14} className="cf-market-icon" style={{ color: '#60a5fa' }} />
          <span className="cf-market-label">市場規模</span>
          <span className="cf-market-value">{scan.marketSnapshot.size}</span>
        </div>
        <div className="cf-market-cell">
          <BarChart3 size={14} className="cf-market-icon" style={{ color: '#34d399' }} />
          <span className="cf-market-label">成長率</span>
          <span className="cf-market-value">{scan.marketSnapshot.growthRate}</span>
        </div>
        <div className="cf-market-insight">
          <AlertTriangle size={13} style={{ color: '#fbbf24', flexShrink: 0, marginTop: 2 }} />
          <span>{scan.marketSnapshot.insight}</span>
        </div>
      </div>
      <div className="cf-section">
        <div className="cf-section-title"><Swords size={14} /> 競爭威脅評分</div>
        <div style={{ height: 160 }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={competitorChartData} layout="vertical" margin={{ left: 0, right: 24, top: 4, bottom: 4 }}>
              <XAxis type="number" domain={[0, 100]} tick={{ fill: '#64748b', fontSize: 10 }} axisLine={false} tickLine={false} />
              <YAxis type="category" dataKey="name" tick={{ fill: '#94a3b8', fontSize: 11 }} width={80} axisLine={false} tickLine={false} />
              <Tooltip
                contentStyle={{ background: '#1e293b', border: '1px solid #334155', borderRadius: 8, color: '#f8fafc', fontSize: 12 }}
                itemStyle={{ color: '#f8fafc' }}
                cursor={{ fill: 'rgba(255, 255, 255, 0.05)' }}
                formatter={(_: any, __: any, props: any) => [props.payload.threat, '威脅程度']}
                labelFormatter={(label: string, payload: any[]) => payload?.[0]?.payload?.fullName || label}
              />
              <Bar dataKey="score" radius={[0, 4, 4, 0]} barSize={14} background={{ fill: 'rgba(255, 255, 255, 0.03)' }}>
                {competitorChartData.map((entry, i) => <Cell key={i} fill={threatBarColor(entry.score)} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
      <div className="cf-section">
        <div className="cf-section-title"><Users size={14} /> 需要補充的關鍵資訊</div>
        <div className="cf-gaps">
          {scan.infoGaps.map((g, i) => (
            <div className="cf-gap-item" key={i}>
              <span className={`cf-gap-tag cf-gap-tag-${g.priority.toLowerCase()}`}>{g.priority}</span>
              <div>
                <div className="cf-gap-field">{g.field}</div>
                <div className="cf-gap-q">{g.question}</div>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

// ─── Chat Messages ─────────────────────────────────────────────────────────────

const ChatMessages: React.FC<{
  messages: Phase2Message[];
  isLoading: boolean;
  bottomRef: React.RefObject<HTMLDivElement>;
}> = ({ messages, isLoading, bottomRef }) => (
  <div className="cf-messages">
    {messages.map((msg, i) => (
      <div key={i} className={`cf-msg cf-msg-${msg.role}`}>
        {msg.role === 'assistant' && <div className="cf-msg-avatar"><Sparkles size={13} /></div>}
        <div className={`cf-msg-bubble cf-msg-bubble-${msg.role}`}>
          {msg.content.split('\n').map((line, j) => {
            const bold = line.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
            return <p key={j} dangerouslySetInnerHTML={{ __html: bold }} style={{ margin: '2px 0' }} />;
          })}
        </div>
      </div>
    ))}
    {isLoading && (
      <div className="cf-msg cf-msg-assistant">
        <div className="cf-msg-avatar"><Sparkles size={13} /></div>
        <div className="cf-msg-bubble cf-msg-bubble-assistant cf-thinking">
          <span /><span /><span />
        </div>
      </div>
    )}
    <div ref={bottomRef} />
  </div>
);

// // ─── Full Result Tabs (財務預測 / 路線圖 / 人物誌 / 競爭態勢 / 三種觀點) ──────────

// const FullResultTabs: React.FC<{ result: AnalysisResult }> = ({ result }) => {
//   const [tab, setTab] = useState<'finance' | 'roadmap' | 'personas' | 'competitors' | 'verdicts'>('finance');

//   const scoreColor = (s: number) => s >= 65 ? '#34d399' : s >= 40 ? '#fbbf24' : '#ef4444';

//   return (
//     <div className="cf-full-result">
//       {/* 最終成功率 Banner */}
//       <div className="cf-final-score-banner">
//         <div className="cf-final-score-left">
//           <Trophy size={20} style={{ color: '#fbbf24' }} />
//           <div>
//             <div className="cf-final-score-label">補充資訊後的最終成功率</div>
//             <div className="cf-final-score-sub">基於完整商業分析模型評估</div>
//           </div>
//         </div>
//         <div className="cf-final-score-value" style={{ color: scoreColor(result.successProbability) }}>
//           {result.successProbability}%
//         </div>
//       </div>

//       {/* Tab 導覽 */}
//       <div className="cf-full-tab-nav">
//         {([
//           { key: 'finance',     label: '財務預測',    icon: <DollarSign size={13} /> },
//           { key: 'roadmap',     label: '策略路線圖',  icon: <Route size={13} /> },
//           { key: 'personas',    label: 'AI 虛擬董事會', icon: <Brain size={13} /> },
//           { key: 'competitors', label: '競爭態勢',    icon: <Swords size={13} /> },
//           { key: 'verdicts',    label: '三種觀點',    icon: <MessageSquare size={13} /> },
//         ] as { key: typeof tab; label: string; icon: React.ReactNode }[]).map(t => (
//           <button key={t.key} className={`cf-full-tab ${tab === t.key ? 'cf-full-tab-active' : ''}`} onClick={() => setTab(t.key)}>
//             {t.icon} {t.label}
//           </button>
//         ))}
//       </div>

//       {/* 財務預測 */}
//       {tab === 'finance' && (
//         <div className="cf-tab-content">
//           <div className="cf-tab-card">
//             <div className="cf-tab-card-title" style={{ marginBottom: '24px' }}><DollarSign size={16} style={{ color: '#34d399' }} /> 3 年財務預測</div>
//             <div style={{ height: 220 }}>
//               <ResponsiveContainer width="100%" height="100%">
//                 <LineChart data={result.financials} margin={{ top: 8, right: 16, left: 0, bottom: 0 }}>
//                   <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
//                   <XAxis dataKey="year" tick={{ fill: '#64748b', fontSize: 11 }} axisLine={false} />
//                   <YAxis tick={{ fill: '#64748b', fontSize: 10 }} axisLine={false} tickLine={false}
//                     tickFormatter={(v: number) => v >= 1000 ? `${(v / 1000).toFixed(0)}K` : String(v)} />
//                   <Tooltip contentStyle={{ background: '#1e293b', border: '1px solid #334155', borderRadius: 8, color: '#f8fafc', fontSize: 12 }}
//                     formatter={(v: number) => [`NT$ ${v.toLocaleString()}`, '']} />
//                   <Legend wrapperStyle={{ color: '#94a3b8', fontSize: 12 }} />
//                   <Line type="monotone" dataKey="revenue" name="營收" stroke="#34d399" strokeWidth={2} dot={{ fill: '#34d399', r: 3 }} />
//                   <Line type="monotone" dataKey="profit"  name="利潤" stroke="#60a5fa" strokeWidth={2} dot={{ fill: '#60a5fa', r: 3 }} />
//                   <Line type="monotone" dataKey="costs"   name="成本" stroke="#f87171" strokeWidth={2} dot={{ fill: '#f87171', r: 3 }} strokeDasharray="4 2" />
//                 </LineChart>
//               </ResponsiveContainer>
//             </div>
//             <div className="cf-break-even">
//               <CheckCircle2 size={14} style={{ color: '#34d399' }} />
//               <span>預估損益兩平點：<strong style={{ color: '#34d399' }}>{result.breakEvenPoint}</strong></span>
//             </div>
//           </div>
//         </div>
//       )}

//       {/* 策略路線圖 */}
//       {tab === 'roadmap' && (
//         <div className="cf-tab-content">
//           <div className="cf-tab-card">
//             <div className="cf-tab-card-title"><Route size={16} style={{ color: '#818cf8' }} /> 策略發展路線圖</div>
//             <div className="cf-roadmap-list">
//               {result.roadmap.map((item, i) => (
//                 <div key={i} className="cf-roadmap-item">
//                   <div className="cf-roadmap-dot" />
//                   <div className="cf-roadmap-body">
//                     <div className="cf-roadmap-phase">{item.phase}</div>
//                     <div className="cf-roadmap-box">
//                       <div className="cf-roadmap-row">
//                         <span className="cf-roadmap-col-label">時程</span>
//                         <span className="cf-roadmap-col-val">{item.timeframe}</span>
//                       </div>
//                       <div className="cf-roadmap-row">
//                         <span className="cf-roadmap-col-label">技術重點</span>
//                         <span className="cf-roadmap-col-val">{item.technology}</span>
//                       </div>
//                       <div className="cf-roadmap-row">
//                         <span className="cf-roadmap-col-label">產品里程碑</span>
//                         <span className="cf-roadmap-col-val">{item.product}</span>
//                       </div>
//                     </div>
//                   </div>
//                 </div>
//               ))}
//             </div>
//           </div>
//         </div>
//       )}

//       {/* AI 虛擬董事會 */}
//       {tab === 'personas' && (
//         <div className="cf-tab-content">
//           <div className="cf-tab-card">
//             <div className="cf-tab-card-title" style={{ marginBottom: '24px' }}><Brain size={16} style={{ color: '#a78bfa' }} /> AI 虛擬董事會評估</div>
//             <div className="cf-personas-grid">
//               {result.personaEvaluations.map((p, i) => (
//                 <div key={i} className="cf-persona-card">
//                   <div className="cf-persona-top">
//                     <div className="cf-persona-icon">{p.icon}</div>
//                     <div className="cf-persona-info">
//                       <div className="cf-persona-role">{p.role}</div>
//                       <div className="cf-persona-score" style={{ color: scoreColor(p.score) }}>{p.score}%</div>
//                     </div>
//                   </div>
//                   <div className="cf-persona-quote">「{p.keyQuote}」</div>
//                   <div className="cf-persona-perspective">{p.perspective}</div>
//                   <div className="cf-persona-concern">
//                     <span className="cf-persona-concern-label">主要顧慮</span>
//                     <span>{p.concern}</span>
//                   </div>
//                 </div>
//               ))}
//             </div>
//           </div>
//         </div>
//       )}

//       {/* 競爭態勢 */}
//       {tab === 'competitors' && (
//         <div className="cf-tab-content">
//           <div className="cf-tab-card">
//             <div className="cf-tab-card-title" style={{ marginBottom: '24px' }}><Swords size={16} style={{ color: '#f87171' }} /> 競爭態勢深度分析</div>
//             <div className="cf-competitors-grid">
//               {result.competitors.map((c, i) => (
//                 <div key={i} className="cf-competitor-item">
//                   <div className="cf-competitor-name">{c.name}</div>
//                   <div className="cf-competitor-row">
//                     <span className="cf-competitor-label" style={{ color: '#34d399' }}>強項</span>
//                     <span className="cf-competitor-text">{c.strength}</span>
//                   </div>
//                   <div className="cf-competitor-row">
//                     <span className="cf-competitor-label" style={{ color: '#f87171' }}>弱點</span>
//                     <span className="cf-competitor-text">{c.weakness}</span>
//                   </div>
//                 </div>
//               ))}
//             </div>
//           </div>
//         </div>
//       )}

//       {/* 三種觀點 */}
//       {tab === 'verdicts' && (
//         <div className="cf-tab-content">
//           <div className="cf-tab-card">
//             <div className="cf-tab-card-title" style={{ marginBottom: '24px' }}><MessageSquare size={16} style={{ color: '#60a5fa' }} /> 三種戰略觀點</div>
//             <div className="cf-verdicts-grid">
//               <div className="cf-verdict cf-verdict-aggressive">
//                 <h4>激進派觀點</h4>
//                 <p>{result.finalVerdicts.aggressive}</p>
//               </div>
//               <div className="cf-verdict cf-verdict-balanced">
//                 <h4>均衡派觀點</h4>
//                 <p>{result.finalVerdicts.balanced}</p>
//               </div>
//               <div className="cf-verdict cf-verdict-conservative">
//                 <h4>保守派觀點</h4>
//                 <p>{result.finalVerdicts.conservative}</p>
//               </div>
//             </div>
//             {result.continueToIterate && (
//               <div className="cf-iterate-note">
//                 <ChevronRight size={14} style={{ color: '#60a5fa', flexShrink: 0 }} />
//                 <span>{result.continueToIterate}</span>
//               </div>
//             )}
//           </div>
//         </div>
//       )}
//     </div>
//   );
// };

// ─── Main ChatFlow Component ───────────────────────────────────────────────────

interface ChatFlowProps {
  initialScan: InitialScanResult;
  fullResult: AnalysisResult | null;
  onComplete: (report: Phase3DeepReport, finalData: BusinessInput) => void;
  onReset: () => void;
}

const ChatFlow: React.FC<ChatFlowProps> = ({ initialScan, fullResult, onComplete, onReset }) => {
  const [messages, setMessages] = useState<Phase2Message[]>([]);
  const [filledData, setFilledData] = useState<BusinessInput>(initialScan.extractedData);
  const [completionRate, setCompletionRate] = useState(30);
  const [isReady, setIsReady] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [isGenerating, setIsGenerating] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [input, setInput] = useState('');
  const [isRecording, setIsRecording] = useState(false);
  const [phase3Report, setPhase3Report] = useState<Phase3DeepReport | null>(null);
  const [mobileTab, setMobileTab] = useState<'analysis' | 'chat'>('chat');

  const mediaRecorderRef = useRef<MediaRecorder | null>(null);
  const audioChunksRef = useRef<Blob[]>([]);
  const bottomRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const highPriorityGap = initialScan.infoGaps.find(g => g.priority === 'HIGH') || initialScan.infoGaps[0];
    const firstQ = highPriorityGap?.question || '請告訴我更多關於你的提案細節。';
    const otherQs = initialScan.infoGaps
      .filter(g => g !== highPriorityGap).slice(0, 2)
      .map((g, i) => `**問題${i + 2}：${g.question}**`);
    let content = `初步掃描完成！根據你的提案，我需要進一步了解幾個關鍵點：\n\n**問題1：${firstQ}**`;
    if (otherQs.length > 0) content += '\n\n' + otherQs.join('\n');
    setMessages([{ role: 'assistant', content, timestamp: Date.now() }]);
  }, []);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, isLoading]);

  const handleSend = async (text: string, audioBase64?: string) => {
    if ((!text.trim() && !audioBase64) || isLoading) return;
    const userMsg: Phase2Message = { role: 'user', content: text || '（語音輸入）', timestamp: Date.now() };
    setMessages(prev => [...prev, userMsg]);
    setInput('');
    setIsLoading(true);
    setError(null);
    try {
      const history = [...messages, userMsg].map(m => ({ role: m.role, content: m.content }));
      const resp: Phase2Response = await runPhase2Reply(text, filledData, history, audioBase64);
      setFilledData(resp.updatedData);
      setCompletionRate(resp.completionRate);
      setIsReady(resp.isReady);
      let aiContent = resp.dynamicFeedback;
      if (resp.nextQuestions.length > 0) {
        aiContent += '\n\n' + resp.nextQuestions.map((q, i) => `**問題${i + 1}：${q}**`).join('\n');
      } else if (resp.isReady) {
        aiContent += '\n\n資訊已充足！可以點擊下方按鈕生成完整的深度分析報告。';
      }
      setMessages(prev => [...prev, { role: 'assistant', content: aiContent, timestamp: Date.now() }]);
    } catch (err: any) {
      setError(err.message || 'AI 回覆失敗，請重試。');
    } finally {
      setIsLoading(false);
    }
  };

  const handleGenerateReport = async () => {
    setIsGenerating(true);
    setError(null);
    try {
      const report = await runPhase3DeepReport(filledData);
      setPhase3Report(report);
      onComplete(report, filledData);
    } catch (err: any) {
      setError(err.message || '報告生成失敗，請重試。');
    } finally {
      setIsGenerating(false);
    }
  };

  const startRecording = async () => {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      const mr = new MediaRecorder(stream);
      mediaRecorderRef.current = mr;
      audioChunksRef.current = [];
      mr.ondataavailable = e => { if (e.data.size > 0) audioChunksRef.current.push(e.data); };
      mr.start();
      setIsRecording(true);
    } catch { alert('無法存取麥克風，請檢查權限設定。'); }
  };

  const stopRecording = () => {
    if (mediaRecorderRef.current && isRecording) {
      mediaRecorderRef.current.stop();
      setIsRecording(false);
      mediaRecorderRef.current.stream.getTracks().forEach(t => t.stop());
      mediaRecorderRef.current.onstop = async () => {
        const b64 = await blobToBase64(new Blob(audioChunksRef.current, { type: 'audio/webm' }));
        await handleSend(input.trim(), b64);
      };
    }
  };

  const barColor = completionRate >= 80 ? '#34d399' : completionRate >= 50 ? '#60a5fa' : '#fbbf24';

  // 正在生成報告
  if (isGenerating) {
    return (
      <div className="cf-root cf-generating">
        <div className="spinner" style={{ margin: '0 auto 1.5rem' }}>
          <div className="spin-ring spin-ring-1" />
          <div className="spin-ring spin-ring-2" />
          <div className="spin-ring spin-ring-3" />
        </div>
        <h3 className="loading-title" style={{ fontSize: '1.25rem' }}>正在生成完整 Master Plan</h3>
        <p className="loading-sub" style={{ fontSize: '0.875rem' }}>
          BMC · SWOT · STEEP · 4C · Porter's 五力 · BCG · 雷達圖 · 行動地圖...
        </p>
      </div>
    );
  }

  return (
    <div className="cf-root">
      <div className="cf-mobile-tabs">
        <button className={`cf-mobile-tab ${mobileTab === 'analysis' ? 'cf-mobile-tab-active' : ''}`} onClick={() => setMobileTab('analysis')}>
          <BarChart3 size={15} /> 掃描分析
        </button>
        <button className={`cf-mobile-tab ${mobileTab === 'chat' ? 'cf-mobile-tab-active' : ''}`} onClick={() => setMobileTab('chat')}>
          <MessageSquare size={15} /> AI 對話
          {messages.filter(m => m.role === 'assistant').length > 0 && (
            <span className="cf-mobile-tab-badge">{messages.filter(m => m.role === 'assistant').length}</span>
          )}
        </button>
      </div>

      <div className={`cf-left ${mobileTab === 'analysis' ? 'cf-mobile-visible' : 'cf-mobile-hidden'}`}>
        <InitialScanPanel scan={initialScan} />
      </div>

      <div className={`cf-right ${mobileTab === 'chat' ? 'cf-mobile-visible' : 'cf-mobile-hidden'}`}>
        <div className="cf-progress-wrap">
          <div className="cf-progress-label">
            <span>資訊飽滿度</span>
            <span style={{ color: barColor, fontWeight: 700 }}>{completionRate}%</span>
          </div>
          <div className="cf-progress-track">
            <div className="cf-progress-fill" style={{ width: `${completionRate}%`, background: barColor, transition: 'width 0.5s ease' }} />
          </div>
          {isReady && (
            <div className="cf-ready-hint"><CheckCircle2 size={13} /> 資訊已充足，可以生成完整報告！</div>
          )}
        </div>

        {error && (
          <div className="ds-error">
            <AlertTriangle size={14} />
            <span style={{ whiteSpace: 'pre-line', flex: 1 }}>{error}</span>
            <button onClick={() => setError(null)}><X size={13} /></button>
          </div>
        )}

        <ChatMessages messages={messages} isLoading={isLoading} bottomRef={bottomRef} />

        {isReady ? (
          <button className="cf-generate-btn" onClick={handleGenerateReport} disabled={isGenerating}>
            <Zap size={16} /> 生成完整深度報告 <ArrowRight size={16} />
          </button>
        ) : completionRate >= 20 ? (
          <button className="cf-generate-btn-warn" onClick={handleGenerateReport} disabled={isGenerating} title={`目前資訊飽滿度 ${completionRate}%，報告品質可能受限`}>
            <AlertTriangle size={15} /> 資訊不足，仍要生成報告（{completionRate}%）
          </button>
        ) : null}

        <div className="cf-input-area">
          <textarea
            className="cf-input"
            placeholder={isRecording ? '正在聆聽...' : '回覆 AI 的問題，或補充任何細節...'}
            value={input}
            onChange={e => setInput(e.target.value)}
            onKeyDown={e => {
              if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleSend(input.trim()); }
            }}
            disabled={isLoading || isRecording}
            rows={3}
          />
          <div className="cf-input-footer">
            <span className="cf-input-saturation" style={{ color: barColor }}>飽滿度 {completionRate}%</span>
            <div className="cf-input-btns">
              {!isRecording
                ? <button className="cf-mic-btn" onClick={startRecording} disabled={isLoading} title="語音輸入"><Mic size={17} /></button>
                : <button className="cf-mic-btn cf-mic-stop" onClick={stopRecording}><StopCircle size={20} /></button>}
              <button className="cf-send-btn" onClick={() => handleSend(input.trim())} disabled={!input.trim() || isLoading}>
                <Send size={17} />
              </button>
            </div>
          </div>
          <div className="cf-input-hint">Enter 傳送 · Shift+Enter 換行</div>
        </div>
      </div>
    </div>
  );
};

export default ChatFlow;

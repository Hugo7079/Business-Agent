import React, { useState, useEffect } from 'react';
import { AppState, AnalysisMode, BusinessInput, AnalysisResult, HistoryRecord, InitialScanResult, Phase3DeepReport } from './types';
import InputForm from './components/InputForm';
import AnalysisDashboard from './components/AnalysisDashboard';
import HistoryPanel from './components/HistoryPanel';
import ChatFlow from './components/ChatFlow';
import { analyzeBusiness, runInitialScan } from './services/geminiService';
import { AlertCircle, Settings, Save, X, Clock } from 'lucide-react';

const HISTORY_KEY = 'omniview_history';
const MAX_HISTORY = 20;

const loadHistory = (): HistoryRecord[] => {
  try { return JSON.parse(localStorage.getItem(HISTORY_KEY) || '[]'); }
  catch { return []; }
};
const saveHistoryToStorage = (records: HistoryRecord[]) => {
  localStorage.setItem(HISTORY_KEY, JSON.stringify(records));
};

const ApiKeyModal = ({ isOpen, onClose }: { isOpen: boolean; onClose: () => void }) => {
  const [key, setKey] = useState(localStorage.getItem('GEMINI_API_KEY') || '');
  const handleSave = () => { localStorage.setItem('GEMINI_API_KEY', key); onClose(); window.location.reload(); };
  if (!isOpen) return null;
  return (
    <div className="modal-overlay">
      <div className="modal-box">
        <div className="modal-header">
          <h3 className="modal-title"><Settings size={20} /> 設定 API Key</h3>
          <button className="modal-close" onClick={onClose}><X size={24} /></button>
        </div>
        <p className="modal-desc">請輸入您的 Google Gemini API Key 以啟用 AI 分析功能。您的 Key 僅會儲存在本地瀏覽器中，不會上傳至任何伺服器。</p>
        <label className="modal-label">Gemini API Key</label>
        <input type="password" className="modal-input" value={key} onChange={(e) => setKey(e.target.value)} placeholder="AIzaSy..." />
        <button className="modal-save-btn" onClick={handleSave}><Save size={20} /> 儲存並重整</button>
        <p className="modal-link">還沒有 Key？ <a href="https://aistudio.google.com/app/apikey" target="_blank" rel="noreferrer">前往 Google AI Studio 申請</a></p>
      </div>
    </div>
  );
};

const App: React.FC = () => {
  const [appState, setAppState] = useState<AppState>('IDLE');
  const [result, setResult] = useState<AnalysisResult | null>(null);
  const [phase3Report, setPhase3Report] = useState<Phase3DeepReport | null>(null);
  const [initialScan, setInitialScan] = useState<InitialScanResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [isHistoryOpen, setIsHistoryOpen] = useState(false);
  const [history, setHistory] = useState<HistoryRecord[]>(loadHistory);
  const [currentInput, setCurrentInput] = useState<BusinessInput | null>(null);

  useEffect(() => { saveHistoryToStorage(history); }, [history]);

  // 使用者提交提案 → 觸發初步掃描
  const handleSubmitProposal = async (text: string, audioBase64?: string) => {
    setAppState('SCANNING');
    setError(null);
    try {
      const scan = await runInitialScan(text, audioBase64);
      setCurrentInput(scan.extractedData);
      setInitialScan(scan);
      setAppState('CHATTING');
    } catch (err: any) {
      setError(err.message || '初步掃描失敗，請重試。');
      setAppState('ERROR');
    }
  };

  // ChatFlow 完成後：同時儲存報告到歷史，並觸發完整分析
  const handleChatComplete = async (report: Phase3DeepReport, finalData: BusinessInput) => {
    setCurrentInput(finalData);
    setPhase3Report(report);
    setAppState('COMPLETE'); // 確保跳轉到結果頁面

    // 背景同步呼叫 analyzeBusiness 取得財務預測、路線圖、人物誌、競爭態勢、三種觀點
    try {
      const fullResult = await analyzeBusiness(AnalysisMode.BUSINESS, finalData);
      setResult(fullResult);

      // 儲存到歷史紀錄
      const record: HistoryRecord = {
        id: Date.now().toString(),
        createdAt: Date.now(),
        title: finalData.idea.slice(0, 60) || '未命名提案',
        input: finalData,
        result: fullResult,
        phase3Report: report,
        finalSuccessProbability: fullResult.successProbability,
      };
      setHistory(prev => [record, ...prev].slice(0, MAX_HISTORY));
    } catch {
      // analyzeBusiness 失敗不影響主報告顯示，僅無法儲存完整歷史
      const record: HistoryRecord = {
        id: Date.now().toString(),
        createdAt: Date.now(),
        title: finalData.idea.slice(0, 60) || '未命名提案',
        input: finalData,
        result: {
          successProbability: 0,
          executiveSummary: '',
          marketAnalysis: { size: '', growthRate: '', description: '' },
          competitors: [],
          roadmap: [],
          financials: [],
          breakEvenPoint: '',
          risks: [],
          personaEvaluations: [],
          teamAnalysis: '',
          finalVerdicts: { aggressive: '', balanced: '', conservative: '' },
          continueToIterate: '',
        },
        phase3Report: report,
      };
      setHistory(prev => [record, ...prev].slice(0, MAX_HISTORY));
    }
  };

  const handleLoadHistory = (record: HistoryRecord) => {
    setResult(record.result);
    setCurrentInput(record.input);
    setPhase3Report(record.phase3Report ?? null);
    setAppState('COMPLETE');
  };

  const handleDeleteHistory = (id: string) => {
    setHistory(prev => prev.filter(r => r.id !== id));
  };

  const handleClearHistory = () => { setHistory([]); };

  const handleReset = () => {
    setResult(null);
    setInitialScan(null);
    setPhase3Report(null);
    setError(null);
    setCurrentInput(null);
    setAppState('IDLE');
  };

  return (
    <div className="app-bg">
      <ApiKeyModal isOpen={isSettingsOpen} onClose={() => setIsSettingsOpen(false)} />
      <HistoryPanel
        isOpen={isHistoryOpen}
        onClose={() => setIsHistoryOpen(false)}
        records={history}
        onLoad={handleLoadHistory}
        onDelete={handleDeleteHistory}
        onClearAll={handleClearHistory}
      />

      <div className="bg-deco">
        <div className="bg-deco-top" />
        <div className="bg-deco-bottom" />
      </div>

      <div className="relative-z10">
        <header className="header">
          <div className="header-inner">
            <div className="logo">
              <div className="logo-icon">OV</div>
              <span className="logo-text">OmniView AI</span>
            </div>
            <div className="header-right">
              <button className="history-header-btn" onClick={() => setIsHistoryOpen(true)} title="歷史記錄">
                <Clock size={18} />
                {history.length > 0 && <span className="history-badge">{history.length}</span>}
              </button>
              <button className="settings-btn" onClick={() => setIsSettingsOpen(true)} title="設定 API Key">
                <Settings size={20} />
              </button>
            </div>
          </div>
        </header>

        <main>
          {/* IDLE：輸入提案 */}
          {appState === 'IDLE' && (
            <InputForm onSubmitProposal={handleSubmitProposal} isLoading={false} />
          )}

          {/* SCANNING：初步掃描 loading */}
          {appState === 'SCANNING' && (
            <div className="loading-overlay" style={{ position: 'relative', padding: '6rem 0' }}>
              <div className="spinner">
                <div className="spin-ring spin-ring-1" />
                <div className="spin-ring spin-ring-2" />
                <div className="spin-ring spin-ring-3" />
              </div>
              <h2 className="loading-title">OmniView 正在掃描你的提案</h2>
              <p className="loading-sub">分析市場、識別競爭對手、評估成功機率...</p>
            </div>
          )}

          {/* CHATTING：初步結果 + 互動補充 */}
          {appState === 'CHATTING' && initialScan && (
            <ChatFlow
              initialScan={initialScan}
              fullResult={result}
              onComplete={handleChatComplete}
              onReset={handleReset}
            />
          )}

          {/* ERROR */}
          {appState === 'ERROR' && (
            <div className="error-box">
              <div className="error-icon"><AlertCircle size={48} /></div>
              <h3 className="error-title">分析失敗</h3>
              <p className="error-msg" style={{ whiteSpace: 'pre-line' }}>{error}</p>
              <div style={{ display: 'flex', gap: '12px', justifyContent: 'center', flexWrap: 'wrap' }}>
                <button className="error-retry-btn" onClick={handleReset}>重試</button>
                <button className="error-retry-btn" style={{ background: '#1e40af' }} onClick={() => setIsSettingsOpen(true)}>設定 API Key</button>
              </div>
            </div>
          )}

          {/* COMPLETE：從對話完成或歷史載入的完整報告 */}
          {appState === 'COMPLETE' && (phase3Report || result) && (
            <AnalysisDashboard
              result={result || {
                successProbability: phase3Report?.radarDimensions[0]?.selfScore || 0,
                executiveSummary: phase3Report?.keyRecommendation || '',
                marketAnalysis: { size: phase3Report?.marketSize.tam || '', growthRate: '', description: phase3Report?.marketSize.description || '' },
                competitors: [],
                roadmap: [],
                financials: [],
                breakEvenPoint: '',
                risks: [],
                personaEvaluations: [],
                teamAnalysis: '',
                finalVerdicts: { aggressive: '', balanced: '', conservative: '' },
                continueToIterate: phase3Report?.continueToIterate || '',
              } as AnalysisResult}
              onReset={handleReset}
              sourceInput={currentInput ?? undefined}
              phase3Report={phase3Report}
            />
          )}
        </main>

        <footer>© {new Date().getFullYear()} OmniView AI. Powered by Google Gemini.</footer>
      </div>
    </div>
  );
};

export default App;
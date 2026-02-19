import React, { useState, useEffect } from 'react';
import { AppState, AnalysisMode, BusinessInput, AnalysisResult, HistoryRecord } from './types';
import InputForm from './components/InputForm';
import AnalysisDashboard from './components/AnalysisDashboard';
import HistoryPanel from './components/HistoryPanel';
import { analyzeBusiness } from './services/geminiService';
import { AlertCircle, Settings, Save, X, Clock } from 'lucide-react';

const HISTORY_KEY = 'omniview_history';
const MAX_HISTORY = 20;

const loadHistory = (): HistoryRecord[] => {
  try { return JSON.parse(localStorage.getItem(HISTORY_KEY) || '[]'); }
  catch { return []; }
};
const saveHistory = (records: HistoryRecord[]) => {
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
  const [error, setError] = useState<string | null>(null);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [isHistoryOpen, setIsHistoryOpen] = useState(false);
  const [history, setHistory] = useState<HistoryRecord[]>(loadHistory);
  const [currentInput, setCurrentInput] = useState<BusinessInput | null>(null);

  useEffect(() => { saveHistory(history); }, [history]);

  const handleAnalyze = async (mode: AnalysisMode, data: BusinessInput) => {
    setAppState('ANALYZING');
    setError(null);
    setCurrentInput(data);
    try {
      const analysisResult = await analyzeBusiness(mode, data);
      setResult(analysisResult);
      setAppState('COMPLETE');

      // 自動儲存到歷史記錄
      const record: HistoryRecord = {
        id: Date.now().toString(),
        createdAt: Date.now(),
        title: data.idea.slice(0, 40) || '未命名提案',
        input: data,
        result: analysisResult,
      };
      setHistory(prev => [record, ...prev].slice(0, MAX_HISTORY));
    } catch (err: any) {
      setError(err.message || '模擬過程中發生意外錯誤。');
      setAppState('ERROR');
    }
  };

  const handleLoadHistory = (record: HistoryRecord) => {
    setResult(record.result);
    setCurrentInput(record.input);
    setAppState('COMPLETE');
  };

  const handleDeleteHistory = (id: string) => {
    setHistory(prev => prev.filter(r => r.id !== id));
  };

  const handleClearHistory = () => {
    setHistory([]);
  };

  const handleReset = () => { setResult(null); setError(null); setCurrentInput(null); setAppState('IDLE'); };

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
          {appState === 'IDLE' && <InputForm onAnalyze={handleAnalyze} isLoading={false} />}

          {appState === 'ANALYZING' && (
            <>
              <InputForm onAnalyze={handleAnalyze} isLoading={true} />
              <div className="loading-overlay">
                <div className="spinner">
                  <div className="spin-ring spin-ring-1" />
                  <div className="spin-ring spin-ring-2" />
                  <div className="spin-ring spin-ring-3" />
                </div>
                <h2 className="loading-title">OmniView 董事會正在開會</h2>
                <p className="loading-sub">正在模擬投資人、競爭對手和客戶的觀點，運算市場數據並評估風險...</p>
              </div>
            </>
          )}

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

          {appState === 'COMPLETE' && result && (
            <AnalysisDashboard result={result} onReset={handleReset} />
          )}
        </main>

        <footer>© {new Date().getFullYear()} OmniView AI. Powered by Google Gemini.</footer>
      </div>
    </div>
  );
};

export default App;
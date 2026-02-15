import React, { useState } from 'react';
import { AppState, AnalysisMode, GeneralInput, CorporateInput, AnalysisResult } from './types';
import InputForm from './components/InputForm';
import AnalysisDashboard from './components/AnalysisDashboard';
import { analyzeBusiness } from './services/geminiService';
import { AlertCircle, Settings, Save, X } from 'lucide-react';

const ApiKeyModal = ({ isOpen, onClose }: { isOpen: boolean; onClose: () => void }) => {
  const [key, setKey] = useState(localStorage.getItem('GEMINI_API_KEY') || '');

  const handleSave = () => {
    localStorage.setItem('GEMINI_API_KEY', key);
    onClose();
    // 重新整理頁面以確保服務讀取到新的 key (非必要，但較保險)
    window.location.reload(); 
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
      <div className="bg-slate-800 border border-slate-700 rounded-2xl p-6 w-full max-w-md shadow-2xl">
        <div className="flex justify-between items-center mb-4">
          <h3 className="text-xl font-bold text-white flex items-center">
            <Settings className="w-5 h-5 mr-2" /> 設定 API Key
          </h3>
          <button onClick={onClose} className="text-slate-400 hover:text-white">
            <X className="w-6 h-6" />
          </button>
        </div>
        
        <p className="text-slate-400 text-sm mb-4">
          請輸入您的 Google Gemini API Key 以啟用 AI 分析功能。
          您的 Key 僅會儲存在本地瀏覽器中，不會上傳至任何伺服器。
        </p>

        <div className="space-y-4">
          <div>
            <label className="block text-sm font-medium text-slate-300 mb-2">Gemini API Key</label>
            <input 
              type="password" 
              value={key}
              onChange={(e) => setKey(e.target.value)}
              placeholder="AIzaSy..."
              className="w-full bg-slate-900 border border-slate-700 rounded-lg p-3 text-white placeholder-slate-600 focus:ring-2 focus:ring-blue-500 focus:outline-none"
            />
          </div>
          
          <button 
            onClick={handleSave}
            className="w-full bg-blue-600 hover:bg-blue-500 text-white font-bold py-3 rounded-xl flex items-center justify-center transition-colors"
          >
            <Save className="w-5 h-5 mr-2" />
            儲存並重整
          </button>
          
          <div className="text-xs text-center text-slate-500 mt-4">
            還沒有 Key？ <a href="https://aistudio.google.com/app/apikey" target="_blank" rel="noreferrer" className="text-blue-400 hover:underline">前往 Google AI Studio 申請</a>
          </div>
        </div>
      </div>
    </div>
  );
};

const App: React.FC = () => {
  const [appState, setAppState] = useState<AppState>('IDLE');
  const [result, setResult] = useState<AnalysisResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);

  const handleAnalyze = async (mode: AnalysisMode, data: GeneralInput | CorporateInput) => {
    setAppState('ANALYZING');
    setError(null);
    try {
      const analysisResult = await analyzeBusiness(mode, data);
      setResult(analysisResult);
      setAppState('COMPLETE');
    } catch (err: any) {
      console.error(err);
      setError(err.message || "模擬過程中發生意外錯誤。");
      setAppState('ERROR');
    }
  };

  const handleReset = () => {
    setResult(null);
    setError(null);
    setAppState('IDLE');
  };

  return (
    <div className="min-h-screen bg-[#0f172a] text-slate-200 selection:bg-blue-500/30">
      <ApiKeyModal isOpen={isSettingsOpen} onClose={() => setIsSettingsOpen(false)} />
      
      {/* Background decoration */}
      <div className="fixed inset-0 pointer-events-none overflow-hidden">
        <div className="absolute top-0 left-0 w-full h-[500px] bg-blue-600/10 blur-[120px] rounded-full opacity-40 transform -translate-y-1/2"></div>
        <div className="absolute bottom-0 right-0 w-full h-[500px] bg-purple-600/10 blur-[120px] rounded-full opacity-30 transform translate-y-1/3"></div>
      </div>

      <div className="relative z-10">
        <header className="border-b border-slate-800 bg-slate-900/50 backdrop-blur-md sticky top-0 z-50">
          <div className="max-w-7xl mx-auto px-4 md:px-8 h-16 flex items-center justify-between">
            <div className="flex items-center space-x-2">
              <div className="w-8 h-8 bg-gradient-to-br from-blue-500 to-indigo-600 rounded-lg flex items-center justify-center font-bold text-white shadow-lg shadow-blue-500/20">
                OV
              </div>
              <span className="font-bold text-lg tracking-tight text-white">OmniView AI</span>
            </div>
            
            <div className="flex items-center space-x-4">
              <div className="text-sm text-slate-500 hidden md:block">
                Business Intelligence Agent v1.0
              </div>
              <button 
                onClick={() => setIsSettingsOpen(true)}
                className="p-2 text-slate-400 hover:text-white hover:bg-slate-800 rounded-lg transition-colors"
                title="設定 API Key"
              >
                <Settings className="w-5 h-5" />
              </button>
            </div>
          </div>
        </header>

        <main className="py-8">
          {appState === 'IDLE' && (
             <div className="animate-in fade-in slide-in-from-bottom-4 duration-500">
               <InputForm onAnalyze={handleAnalyze} isLoading={false} />
             </div>
          )}

          {appState === 'ANALYZING' && (
            <div className="flex flex-col items-center justify-center min-h-[60vh] animate-in fade-in duration-500">
              <InputForm onAnalyze={handleAnalyze} isLoading={true} />
              <div className="fixed inset-0 bg-slate-900/80 backdrop-blur-sm z-50 flex flex-col items-center justify-center p-4">
                 <div className="relative w-24 h-24 mb-8">
                    <div className="absolute inset-0 border-t-4 border-blue-500 rounded-full animate-spin"></div>
                    <div className="absolute inset-2 border-r-4 border-purple-500 rounded-full animate-spin reverse"></div>
                    <div className="absolute inset-4 border-b-4 border-emerald-500 rounded-full animate-spin"></div>
                 </div>
                 <h2 className="text-2xl font-bold text-white mb-2">OmniView 董事會正在開會</h2>
                 <p className="text-slate-400 text-center max-w-md animate-pulse">
                   正在模擬投資人、競爭對手和客戶的觀點。正在運算市場數據並評估風險...
                 </p>
              </div>
            </div>
          )}

          {appState === 'ERROR' && (
            <div className="max-w-2xl mx-auto mt-20 p-8 bg-red-900/20 border border-red-500/50 rounded-2xl text-center">
              <AlertCircle className="w-12 h-12 text-red-500 mx-auto mb-4" />
              <h3 className="text-xl font-bold text-white mb-2">分析失敗</h3>
              <p className="text-red-200 mb-6">{error}</p>
              <button 
                onClick={handleReset}
                className="px-6 py-2 bg-red-600 hover:bg-red-500 text-white rounded-lg font-semibold transition-colors"
              >
                重試
              </button>
            </div>
          )}

          {appState === 'COMPLETE' && result && (
            <AnalysisDashboard result={result} onReset={handleReset} />
          )}
        </main>

        <footer className="py-8 text-center text-slate-600 text-sm">
          <p>© {new Date().getFullYear()} OmniView AI. Powered by Google Gemini.</p>
        </footer>
      </div>
    </div>
  );
};

export default App;
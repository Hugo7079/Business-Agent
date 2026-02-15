import React, { useState, useRef } from 'react';
import { AnalysisMode, GeneralInput, CorporateInput } from '../types';
import { Briefcase, Building2, Rocket, ArrowRight, Loader2, Mic, StopCircle, Sparkles, PenTool } from 'lucide-react';
import { parseBusinessIdea } from '../services/geminiService';

interface Props {
  onAnalyze: (mode: AnalysisMode, data: GeneralInput | CorporateInput) => void;
  isLoading: boolean;
}

const InputForm: React.FC<Props> = ({ onAnalyze, isLoading }) => {
  const [mode, setMode] = useState<AnalysisMode>(AnalysisMode.STARTUP);
  const [inputMethod, setInputMethod] = useState<'quick' | 'manual'>('quick');
  const [isProcessingInput, setIsProcessingInput] = useState(false);
  const [feedback, setFeedback] = useState<string | null>(null);

  // Quick Input State
  const [quickText, setQuickText] = useState("");
  const [isRecording, setIsRecording] = useState(false);
  const mediaRecorderRef = useRef<MediaRecorder | null>(null);
  const audioChunksRef = useRef<Blob[]>([]);

  // State for General Startup
  const [startupData, setStartupData] = useState<GeneralInput>({
    marketData: '',
    productDetails: '',
    literature: '',
    painPoints: '',
    techInnovation: '',
  });

  // State for Corporate
  const [corpData, setCorpData] = useState<CorporateInput>({
    brandTraits: '',
    targetConsumer: '',
    channels: '',
    proposalGoal: '',
    financialReport: '',
  });

  // Audio Logic
  const startRecording = async () => {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      const mediaRecorder = new MediaRecorder(stream);
      mediaRecorderRef.current = mediaRecorder;
      audioChunksRef.current = [];

      mediaRecorder.ondataavailable = (event) => {
        if (event.data.size > 0) {
          audioChunksRef.current.push(event.data);
        }
      };

      mediaRecorder.start();
      setIsRecording(true);
    } catch (err) {
      console.error("Error accessing microphone:", err);
      alert("無法存取麥克風，請檢查您的權限設定。");
    }
  };

  const stopRecording = () => {
    if (mediaRecorderRef.current && isRecording) {
      mediaRecorderRef.current.stop();
      setIsRecording(false);
      // Stop all tracks
      mediaRecorderRef.current.stream.getTracks().forEach(track => track.stop());
      
      // Handle the blob after stop
      mediaRecorderRef.current.onstop = async () => {
        const audioBlob = new Blob(audioChunksRef.current, { type: 'audio/webm' });
        await handleProcessInput(audioBlob);
      };
    }
  };

  const blobToBase64 = (blob: Blob): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => {
        const base64String = reader.result as string;
        // Remove the data URL prefix (e.g., "data:audio/webm;base64,")
        const base64Data = base64String.split(',')[1];
        resolve(base64Data);
      };
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  };

  const handleProcessInput = async (audioBlob?: Blob) => {
    if (!quickText.trim() && !audioBlob) return;

    setIsProcessingInput(true);
    setFeedback(null);
    
    try {
      let audioBase64: string | undefined = undefined;
      if (audioBlob) {
        audioBase64 = await blobToBase64(audioBlob);
      }

      const result = await parseBusinessIdea(mode, quickText, audioBase64);
      
      if (mode === AnalysisMode.STARTUP) {
        setStartupData(result.data as GeneralInput);
      } else {
        setCorpData(result.data as CorporateInput);
      }
      
      setFeedback(result.feedback);
      setInputMethod('manual'); // Switch to manual view to show filled data
    } catch (error) {
      console.error(error);
      alert("輸入處理失敗。請重試或手動填寫表格。");
    } finally {
      setIsProcessingInput(false);
    }
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (mode === AnalysisMode.STARTUP) {
      onAnalyze(mode, startupData);
    } else {
      onAnalyze(mode, corpData);
    }
  };

  const handleStartupChange = (key: keyof GeneralInput, value: string) => {
    setStartupData(prev => ({ ...prev, [key]: value }));
  };

  const handleCorpChange = (key: keyof CorporateInput, value: string) => {
    setCorpData(prev => ({ ...prev, [key]: value }));
  };

  return (
    <div className="w-full max-w-4xl mx-auto p-6">
      <div className="mb-8 text-center">
        <h1 className="text-4xl font-bold bg-clip-text text-transparent bg-gradient-to-r from-cyan-400 to-blue-600 mb-4">
          OmniView AI 全方位商業評估
        </h1>
        <p className="text-slate-400">
          您的 360° AI 董事會。
        </p>
      </div>

      {/* Mode Selection */}
      <div className="flex justify-center mb-6">
        <div className="bg-slate-800 p-1 rounded-xl flex shadow-lg border border-slate-700">
          <button
            onClick={() => setMode(AnalysisMode.STARTUP)}
            className={`flex items-center px-6 py-3 rounded-lg text-sm font-semibold transition-all ${
              mode === AnalysisMode.STARTUP
                ? 'bg-blue-600 text-white shadow-md'
                : 'text-slate-400 hover:text-white hover:bg-slate-700'
            }`}
          >
            <Rocket className="w-4 h-4 mr-2" />
            新創事業 / 創業
          </button>
          <button
            onClick={() => setMode(AnalysisMode.CORPORATE)}
            className={`flex items-center px-6 py-3 rounded-lg text-sm font-semibold transition-all ${
              mode === AnalysisMode.CORPORATE
                ? 'bg-purple-600 text-white shadow-md'
                : 'text-slate-400 hover:text-white hover:bg-slate-700'
            }`}
          >
            <Building2 className="w-4 h-4 mr-2" />
            企業品牌 / 提案
          </button>
        </div>
      </div>

      {/* Input Method Tabs */}
      <div className="flex space-x-2 mb-4 justify-center">
        <button 
           onClick={() => setInputMethod('quick')}
           className={`px-4 py-2 rounded-t-lg text-sm font-medium transition-colors ${
             inputMethod === 'quick' ? 'bg-slate-800/50 text-blue-400 border-t border-x border-slate-700' : 'text-slate-500 hover:text-slate-300'
           }`}
        >
          <Sparkles className="w-4 h-4 inline mr-2" />
          AI 快速輔助 (語音/文字)
        </button>
        <button 
           onClick={() => setInputMethod('manual')}
           className={`px-4 py-2 rounded-t-lg text-sm font-medium transition-colors ${
             inputMethod === 'manual' ? 'bg-slate-800/50 text-blue-400 border-t border-x border-slate-700' : 'text-slate-500 hover:text-slate-300'
           }`}
        >
          <PenTool className="w-4 h-4 inline mr-2" />
          詳細表單
        </button>
      </div>

      <div className="bg-slate-800/50 border border-slate-700 rounded-2xl p-8 shadow-xl backdrop-blur-sm min-h-[500px]">
        
        {/* Quick Input View */}
        {inputMethod === 'quick' && (
          <div className="space-y-8 animate-in fade-in slide-in-from-bottom-2 duration-300">
             <div className="text-center space-y-2">
               <h2 className="text-xl font-bold text-white">描述您的想法</h2>
               <p className="text-slate-400 text-sm max-w-lg mx-auto">
                 請自由輸入或說話。我們的 AI 將聆聽、分析您的輸入，並自動將其整理為專業的商業模式格式。
               </p>
             </div>

             <div className="relative">
                <textarea
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl p-6 text-slate-200 placeholder-slate-600 focus:ring-2 focus:ring-blue-500 focus:border-transparent transition-all resize-none h-48 text-lg"
                  placeholder={isRecording ? "正在聆聽..." : "例如：我想建立一個針對有機狗糧的訂閱服務，利用 AI 客製化飲食計畫..."}
                  value={quickText}
                  onChange={(e) => setQuickText(e.target.value)}
                  disabled={isRecording || isProcessingInput}
                />
                
                {/* Recording Pulse Animation */}
                {isRecording && (
                  <div className="absolute top-4 right-4 flex items-center space-x-2 bg-red-500/10 text-red-500 px-3 py-1 rounded-full text-xs font-bold animate-pulse border border-red-500/20">
                    <div className="w-2 h-2 bg-red-500 rounded-full"></div>
                    <span>錄音中</span>
                  </div>
                )}
             </div>

             <div className="flex flex-col items-center space-y-4">
                <div className="flex items-center space-x-4">
                  {!isRecording ? (
                    <button
                      onClick={startRecording}
                      disabled={isProcessingInput}
                      className="flex items-center justify-center w-16 h-16 rounded-full bg-slate-700 hover:bg-red-500/20 hover:text-red-500 text-slate-300 border border-slate-600 transition-all group"
                      title="開始錄音"
                    >
                      <Mic className="w-6 h-6 group-hover:scale-110 transition-transform" />
                    </button>
                  ) : (
                    <button
                      onClick={stopRecording}
                      className="flex items-center justify-center w-16 h-16 rounded-full bg-red-500 text-white hover:bg-red-600 shadow-lg shadow-red-500/30 transition-all animate-pulse"
                      title="停止並處理"
                    >
                      <StopCircle className="w-8 h-8" />
                    </button>
                  )}
                </div>
                <div className="text-xs text-slate-500 font-medium">
                  {isRecording ? "點擊停止說話" : "點擊麥克風開始說話"}
                </div>
             </div>

             <div className="pt-6 border-t border-slate-700/50 flex justify-end">
                <button
                  onClick={() => handleProcessInput()}
                  disabled={isProcessingInput || (!quickText && !isRecording)}
                  className={`
                    flex items-center px-6 py-3 rounded-xl text-md font-bold shadow-lg transition-all
                    ${isProcessingInput || (!quickText && !isRecording)
                      ? 'bg-slate-700 text-slate-500 cursor-not-allowed' 
                      : 'bg-blue-600 hover:bg-blue-500 text-white shadow-blue-500/20'
                    }
                  `}
                >
                  {isProcessingInput ? (
                    <>
                      <Loader2 className="w-5 h-5 mr-2 animate-spin" />
                      正在整理資料...
                    </>
                  ) : (
                    <>
                      <Sparkles className="w-5 h-5 mr-2" />
                      處理並自動填寫
                    </>
                  )}
                </button>
             </div>
          </div>
        )}

        {/* Manual Form View */}
        {inputMethod === 'manual' && (
          <form onSubmit={handleSubmit} className="animate-in fade-in slide-in-from-bottom-2 duration-300">
            {feedback && (
              <div className="mb-6 p-4 bg-blue-500/10 border border-blue-500/20 rounded-lg text-sm text-blue-200 flex items-start">
                <Sparkles className="w-5 h-5 mr-3 flex-shrink-0 mt-0.5 text-blue-400" />
                <div>
                  <span className="font-bold block mb-1">AI 助理：</span>
                  {feedback}
                </div>
              </div>
            )}

            {mode === AnalysisMode.STARTUP ? (
              <div className="space-y-6">
                <h2 className="text-xl text-blue-400 font-semibold mb-4 flex items-center">
                  <Rocket className="w-5 h-5 mr-2" /> 新創事業參數
                </h2>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <InputArea 
                    label="市場資料 & 趨勢 (Market Data)" 
                    placeholder="描述目前市場規模、CAGR、趨勢..." 
                    value={startupData.marketData}
                    onChange={(v) => handleStartupChange('marketData', v)}
                  />
                  <InputArea 
                    label="產品細節 (Product Details)" 
                    placeholder="產品是什麼？核心功能為何？" 
                    value={startupData.productDetails}
                    onChange={(v) => handleStartupChange('productDetails', v)}
                  />
                  <InputArea 
                    label="目前痛點 / 市場缺口 (Pain Points)" 
                    placeholder="為什麼市場現在需要這個？" 
                    value={startupData.painPoints}
                    onChange={(v) => handleStartupChange('painPoints', v)}
                  />
                  <InputArea 
                    label="技術創新 / 獨特賣點 (USP)" 
                    placeholder="這項技術有什麼獨特之處？" 
                    value={startupData.techInnovation}
                    onChange={(v) => handleStartupChange('techInnovation', v)}
                  />
                  <InputArea 
                    label="支持文獻 / 數據 (Literature/Data)" 
                    placeholder="任何關鍵報告、論文或現有數據？" 
                    value={startupData.literature}
                    onChange={(v) => handleStartupChange('literature', v)}
                    fullWidth
                  />
                </div>
              </div>
            ) : (
              <div className="space-y-6">
                <h2 className="text-xl text-purple-400 font-semibold mb-4 flex items-center">
                  <Building2 className="w-5 h-5 mr-2" /> 品牌提案參數
                </h2>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                   <InputArea 
                    label="品牌特性 (Brand Traits)" 
                    placeholder="核心價值、定位、歷史..." 
                    value={corpData.brandTraits}
                    onChange={(v) => handleCorpChange('brandTraits', v)}
                  />
                  <InputArea 
                    label="目標消費者 (Target Consumer)" 
                    placeholder="人口統計、心理特徵、行為..." 
                    value={corpData.targetConsumer}
                    onChange={(v) => handleCorpChange('targetConsumer', v)}
                  />
                  <InputArea 
                    label="分銷通路 (Distribution Channels)" 
                    placeholder="零售、DTC、B2B 合作夥伴..." 
                    value={corpData.channels}
                    onChange={(v) => handleCorpChange('channels', v)}
                  />
                  <InputArea 
                    label="提案目標 & 預期效果 (Goal & Effect)" 
                    placeholder="推出新產品線？品牌重塑？擴張？" 
                    value={corpData.proposalGoal}
                    onChange={(v) => handleCorpChange('proposalGoal', v)}
                  />
                  <InputArea 
                    label="財務背景 (Financial Context)" 
                    placeholder="目前預算、營收、損益狀況..." 
                    value={corpData.financialReport}
                    onChange={(v) => handleCorpChange('financialReport', v)}
                    fullWidth
                  />
                </div>
              </div>
            )}

            <div className="mt-8 flex justify-end">
              <button
                type="submit"
                disabled={isLoading}
                className={`
                  flex items-center px-8 py-4 rounded-xl text-lg font-bold shadow-lg transition-all
                  ${isLoading 
                    ? 'bg-slate-700 text-slate-400 cursor-not-allowed' 
                    : 'bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 text-white hover:shadow-blue-500/25 ring-2 ring-transparent hover:ring-blue-400'
                  }
                `}
              >
                {isLoading ? (
                  <>
                    <Loader2 className="w-6 h-6 mr-2 animate-spin" />
                    正在模擬董事會會議...
                  </>
                ) : (
                  <>
                    開始 AI 分析
                    <ArrowRight className="w-6 h-6 ml-2" />
                  </>
                )}
              </button>
            </div>
          </form>
        )}
      </div>
    </div>
  );
};

const InputArea = ({ label, placeholder, value, onChange, fullWidth = false }: { 
  label: string; 
  placeholder: string; 
  value: string; 
  onChange: (val: string) => void;
  fullWidth?: boolean;
}) => (
  <div className={fullWidth ? 'col-span-1 md:col-span-2' : ''}>
    <label className="block text-sm font-medium text-slate-300 mb-2">{label}</label>
    <textarea
      className="w-full bg-slate-900 border border-slate-700 rounded-lg p-3 text-slate-200 placeholder-slate-600 focus:ring-2 focus:ring-blue-500 focus:border-transparent transition-all resize-none h-32 text-sm"
      placeholder={placeholder}
      value={value}
      onChange={(e) => onChange(e.target.value)}
      required
    />
  </div>
);

export default InputForm;
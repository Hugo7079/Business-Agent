import React, { useState, useRef } from 'react';
import { AnalysisMode, GeneralInput, CorporateInput } from '../types';
import { Building2, Rocket, ArrowRight, Loader2, Mic, StopCircle, Sparkles, PenTool } from 'lucide-react';
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
  const [quickText, setQuickText] = useState("");
  const [isRecording, setIsRecording] = useState(false);
  const mediaRecorderRef = useRef<MediaRecorder | null>(null);
  const audioChunksRef = useRef<Blob[]>([]);

  const [startupData, setStartupData] = useState<GeneralInput>({ marketData: '', productDetails: '', literature: '', painPoints: '', techInnovation: '' });
  const [corpData, setCorpData] = useState<CorporateInput>({ brandTraits: '', targetConsumer: '', channels: '', proposalGoal: '', financialReport: '' });

  const startRecording = async () => {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      const mediaRecorder = new MediaRecorder(stream);
      mediaRecorderRef.current = mediaRecorder;
      audioChunksRef.current = [];
      mediaRecorder.ondataavailable = (e) => { if (e.data.size > 0) audioChunksRef.current.push(e.data); };
      mediaRecorder.start();
      setIsRecording(true);
    } catch { alert("無法存取麥克風，請檢查您的權限設定。"); }
  };

  const stopRecording = () => {
    if (mediaRecorderRef.current && isRecording) {
      mediaRecorderRef.current.stop();
      setIsRecording(false);
      mediaRecorderRef.current.stream.getTracks().forEach(t => t.stop());
      mediaRecorderRef.current.onstop = async () => {
        const blob = new Blob(audioChunksRef.current, { type: 'audio/webm' });
        await handleProcessInput(blob);
      };
    }
  };

  const blobToBase64 = (blob: Blob): Promise<string> => new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve((reader.result as string).split(',')[1]);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });

  const handleProcessInput = async (audioBlob?: Blob) => {
    if (!quickText.trim() && !audioBlob) return;
    setIsProcessingInput(true); setFeedback(null);
    try {
      const audioBase64 = audioBlob ? await blobToBase64(audioBlob) : undefined;
      const result = await parseBusinessIdea(mode, quickText, audioBase64);
      if (mode === AnalysisMode.STARTUP) setStartupData(result.data as GeneralInput);
      else setCorpData(result.data as CorporateInput);
      setFeedback(result.feedback);
      setInputMethod('manual');
    } catch { alert("輸入處理失敗。請重試或手動填寫表格。"); }
    finally { setIsProcessingInput(false); }
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    onAnalyze(mode, mode === AnalysisMode.STARTUP ? startupData : corpData);
  };

  return (
    <div className="form-wrapper">
      <div className="form-heading">
        <h1 className="form-title">OmniView AI 全方位商業評估</h1>
        <p className="form-subtitle">您的 360° AI 董事會。</p>
      </div>

      <div className="mode-toggle">
        <div className="mode-toggle-inner">
          <button onClick={() => setMode(AnalysisMode.STARTUP)} className={`mode-btn ${mode === AnalysisMode.STARTUP ? 'active-startup' : ''}`}>
            <Rocket size={16} /> 新創事業 / 創業
          </button>
          <button onClick={() => setMode(AnalysisMode.CORPORATE)} className={`mode-btn ${mode === AnalysisMode.CORPORATE ? 'active-corporate' : ''}`}>
            <Building2 size={16} /> 企業品牌 / 提案
          </button>
        </div>
      </div>

      <div className="tab-bar">
        <button onClick={() => setInputMethod('quick')} className={`tab-btn ${inputMethod === 'quick' ? 'active' : ''}`}>
          <Sparkles size={16} /> AI 快速輔助 (語音/文字)
        </button>
        <button onClick={() => setInputMethod('manual')} className={`tab-btn ${inputMethod === 'manual' ? 'active' : ''}`}>
          <PenTool size={16} /> 詳細表單
        </button>
      </div>

      <div className="form-card">
        {inputMethod === 'quick' && (
          <div className="quick-input-section">
            <div className="quick-heading">
              <h2>描述您的想法</h2>
              <p>請自由輸入或說話。我們的 AI 將聆聽、分析您的輸入，並自動將其整理為專業的商業模式格式。</p>
            </div>
            <div className="textarea-wrap">
              <textarea
                className="quick-textarea"
                placeholder={isRecording ? "正在聆聽..." : "例如：我想建立一個針對有機狗糧的訂閱服務，利用 AI 客製化飲食計畫..."}
                value={quickText}
                onChange={(e) => setQuickText(e.target.value)}
                disabled={isRecording || isProcessingInput}
              />
              {isRecording && (
                <div className="recording-badge">
                  <div className="recording-dot" /> 錄音中
                </div>
              )}
            </div>
            <div className="mic-section">
              {!isRecording ? (
                <button onClick={startRecording} disabled={isProcessingInput} className="mic-btn" title="開始錄音">
                  <Mic size={24} />
                </button>
              ) : (
                <button onClick={stopRecording} className="mic-btn-stop" title="停止並處理">
                  <StopCircle size={32} />
                </button>
              )}
              <span className="mic-hint">{isRecording ? "點擊停止說話" : "點擊麥克風開始說話"}</span>
            </div>
            <div className="quick-footer">
              <button
                onClick={() => handleProcessInput()}
                disabled={isProcessingInput || (!quickText && !isRecording)}
                className={`process-btn ${isProcessingInput || (!quickText && !isRecording) ? 'disabled' : 'enabled'}`}
              >
                {isProcessingInput ? <><Loader2 size={20} /> 正在整理資料...</> : <><Sparkles size={20} /> 處理並自動填寫</>}
              </button>
            </div>
          </div>
        )}

        {inputMethod === 'manual' && (
          <form onSubmit={handleSubmit}>
            {feedback && (
              <div className="feedback-box">
                <Sparkles size={20} />
                <div><strong>AI 助理：</strong>{feedback}</div>
              </div>
            )}
            {mode === AnalysisMode.STARTUP ? (
              <>
                <div className="section-title-startup"><Rocket size={18} /> 新創事業參數</div>
                <div className="fields-grid">
                  <InputArea label="市場資料 & 趨勢 (Market Data)" placeholder="描述目前市場規模、CAGR、趨勢..." value={startupData.marketData} onChange={v => setStartupData(p => ({...p, marketData: v}))} />
                  <InputArea label="產品細節 (Product Details)" placeholder="產品是什麼？核心功能為何？" value={startupData.productDetails} onChange={v => setStartupData(p => ({...p, productDetails: v}))} />
                  <InputArea label="目前痛點 / 市場缺口 (Pain Points)" placeholder="為什麼市場現在需要這個？" value={startupData.painPoints} onChange={v => setStartupData(p => ({...p, painPoints: v}))} />
                  <InputArea label="技術創新 / 獨特賣點 (USP)" placeholder="這項技術有什麼獨特之處？" value={startupData.techInnovation} onChange={v => setStartupData(p => ({...p, techInnovation: v}))} />
                  <InputArea label="支持文獻 / 數據 (Literature/Data)" placeholder="任何關鍵報告、論文或現有數據？" value={startupData.literature} onChange={v => setStartupData(p => ({...p, literature: v}))} fullWidth />
                </div>
              </>
            ) : (
              <>
                <div className="section-title-corp"><Building2 size={18} /> 品牌提案參數</div>
                <div className="fields-grid">
                  <InputArea label="品牌特性 (Brand Traits)" placeholder="核心價值、定位、歷史..." value={corpData.brandTraits} onChange={v => setCorpData(p => ({...p, brandTraits: v}))} />
                  <InputArea label="目標消費者 (Target Consumer)" placeholder="人口統計、心理特徵、行為..." value={corpData.targetConsumer} onChange={v => setCorpData(p => ({...p, targetConsumer: v}))} />
                  <InputArea label="分銷通路 (Distribution Channels)" placeholder="零售、DTC、B2B 合作夥伴..." value={corpData.channels} onChange={v => setCorpData(p => ({...p, channels: v}))} />
                  <InputArea label="提案目標 & 預期效果 (Goal & Effect)" placeholder="推出新產品線？品牌重塑？擴張？" value={corpData.proposalGoal} onChange={v => setCorpData(p => ({...p, proposalGoal: v}))} />
                  <InputArea label="財務背景 (Financial Context)" placeholder="目前預算、營收、損益狀況..." value={corpData.financialReport} onChange={v => setCorpData(p => ({...p, financialReport: v}))} fullWidth />
                </div>
              </>
            )}
            <div className="submit-row">
              <button type="submit" disabled={isLoading} className="submit-btn">
                {isLoading ? <><Loader2 size={22} /> 正在模擬董事會會議...</> : <>開始 AI 分析 <ArrowRight size={22} /></>}
              </button>
            </div>
          </form>
        )}
      </div>
    </div>
  );
};

const InputArea = ({ label, placeholder, value, onChange, fullWidth = false }: {
  label: string; placeholder: string; value: string;
  onChange: (val: string) => void; fullWidth?: boolean;
}) => (
  <div className={`field-wrap${fullWidth ? ' full' : ''}`}>
    <label>{label}</label>
    <textarea className="field-textarea" placeholder={placeholder} value={value} onChange={(e) => onChange(e.target.value)} required />
  </div>
);

export default InputForm;
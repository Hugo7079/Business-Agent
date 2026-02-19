import React, { useState, useRef } from 'react';
import { AnalysisMode, BusinessInput } from '../types';
import { ArrowRight, Loader2, Mic, StopCircle, Sparkles, PenTool } from 'lucide-react';
import { parseBusinessIdea } from '../services/geminiService';

interface Props {
  onAnalyze: (mode: AnalysisMode, data: BusinessInput) => void;
  isLoading: boolean;
}

const InputForm: React.FC<Props> = ({ onAnalyze, isLoading }) => {
  const [inputMethod, setInputMethod] = useState<'quick' | 'manual'>('quick');
  const [isProcessingInput, setIsProcessingInput] = useState(false);
  const [feedback, setFeedback] = useState<string | null>(null);
  const [quickText, setQuickText] = useState('');
  const [isRecording, setIsRecording] = useState(false);
  const mediaRecorderRef = useRef<MediaRecorder | null>(null);
  const audioChunksRef = useRef<Blob[]>([]);

  const [formData, setFormData] = useState<BusinessInput>({
    idea: '',
    marketData: '',
    productDetails: '',
    painPoints: '',
    targetConsumer: '',
    financialContext: '',
  });

  const startRecording = async () => {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      const mediaRecorder = new MediaRecorder(stream);
      mediaRecorderRef.current = mediaRecorder;
      audioChunksRef.current = [];
      mediaRecorder.ondataavailable = (e) => { if (e.data.size > 0) audioChunksRef.current.push(e.data); };
      mediaRecorder.start();
      setIsRecording(true);
    } catch { alert('無法存取麥克風，請檢查您的權限設定。'); }
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
    setIsProcessingInput(true);
    setFeedback(null);
    try {
      const audioBase64 = audioBlob ? await blobToBase64(audioBlob) : undefined;
      const result = await parseBusinessIdea(AnalysisMode.BUSINESS, quickText, audioBase64);
      setFormData(result.data as BusinessInput);
      setFeedback(result.feedback);
      setInputMethod('manual');
    } catch (err: any) {
      alert(err.message || '輸入處理失敗，請重試。');
    } finally {
      setIsProcessingInput(false);
    }
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    onAnalyze(AnalysisMode.BUSINESS, formData);
  };

  const fields: { key: keyof BusinessInput; label: string; placeholder: string }[] = [
    { key: 'idea',            label: '核心想法',          placeholder: '用一兩句話描述您的商業提案核心...' },
    { key: 'marketData',      label: '市場資料與趨勢',     placeholder: '市場規模、CAGR、主要趨勢...' },
    { key: 'productDetails',  label: '產品或服務細節',     placeholder: '您的產品是什麼？核心功能或服務內容...' },
    { key: 'painPoints',      label: '市場痛點',           placeholder: '目前市場缺少什麼？您解決了什麼問題...' },
    { key: 'targetConsumer',  label: '目標客群',           placeholder: '誰會使用您的產品？人口統計、行為特徵...' },
    { key: 'financialContext', label: '財務背景',          placeholder: '預算、預期營收、現有資金狀況...' },
  ];

  return (
    <div className="form-wrapper">
      <div className="form-heading">
        <h1 className="form-title">OmniView AI<br/>商業提案評估</h1>
        <p className="form-subtitle">您的 360° AI 董事會，提交任何商業想法，獲得全方位專業分析。</p>
      </div>

      <div className="tab-bar">
        <button onClick={() => setInputMethod('quick')} className={`tab-btn ${inputMethod === 'quick' ? 'active' : ''}`}>
          <Sparkles size={16} /> AI 快速輸入
        </button>
        <button onClick={() => setInputMethod('manual')} className={`tab-btn ${inputMethod === 'manual' ? 'active' : ''}`}>
          <PenTool size={16} /> 詳細表單
        </button>
      </div>

      <div className="form-card">
        {inputMethod === 'quick' && (
          <div className="quick-input-section">
            <div className="quick-heading">
              <h2>描述您的商業想法</h2>
              <p>無論是新創、品牌提案、產品計畫，直接說或打出來，AI 會自動整理成完整的分析表單。</p>
            </div>
            <div className="textarea-wrap">
              <textarea
                className="quick-textarea"
                placeholder={isRecording ? '正在聆聽...' : '例如：我想推出一個 AI 客製化寵物飲食訂閱服務，目標是有機飼主市場，預計第一年營收 300 萬...'}
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
              <span className="mic-hint">{isRecording ? '點擊停止說話' : '點擊麥克風開始說話'}</span>
            </div>
            <div className="quick-footer">
              <button
                onClick={() => handleProcessInput()}
                disabled={isProcessingInput || (!quickText.trim() && !isRecording)}
                className={`process-btn ${isProcessingInput || (!quickText.trim() && !isRecording) ? 'disabled' : 'enabled'}`}
              >
                {isProcessingInput
                  ? <><Loader2 size={20} /> 正在整理資料...</>
                  : <><Sparkles size={20} /> AI 自動填寫表單</>}
              </button>
            </div>
          </div>
        )}

        {inputMethod === 'manual' && (
          <form onSubmit={handleSubmit}>
            {feedback && (
              <div className="feedback-box">
                <Sparkles size={20} style={{flexShrink:0}} />
                <div><strong>AI 助理：</strong>{feedback}</div>
              </div>
            )}
            <div className="fields-grid">
              {fields.map(({ key, label, placeholder }) => (
                <div key={key} className={`field-wrap ${key === 'idea' ? 'full' : ''}`}>
                  <label>{label}</label>
                  <textarea
                    className="field-textarea"
                    placeholder={placeholder}
                    value={formData[key]}
                    onChange={(e) => setFormData(p => ({ ...p, [key]: e.target.value }))}
                    required
                  />
                </div>
              ))}
            </div>
            <div className="submit-row">
              <button type="submit" disabled={isLoading} className="submit-btn">
                {isLoading
                  ? <><Loader2 size={22} /> 董事會分析中...</>
                  : <>開始 AI 分析 <ArrowRight size={22} /></>}
              </button>
            </div>
          </form>
        )}
      </div>
    </div>
  );
};

export default InputForm;
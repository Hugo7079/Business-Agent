import React, { useState, useRef } from 'react';
import { ArrowRight, Loader2, Mic, StopCircle, Sparkles, Paperclip, X, FileText, Image, File } from 'lucide-react';
import { parseFileInput } from '../services/geminiService';

interface Props {
  onSubmitProposal: (text: string, audioBase64?: string) => void;
  isLoading: boolean;
}

type InputTab = 'quick' | 'file';

const ACCEPTED_TYPES = [
  'application/pdf',
  'image/png', 'image/jpeg', 'image/webp', 'image/gif',
  'text/plain',
];
const ACCEPTED_EXTENSIONS = '.pdf,.png,.jpg,.jpeg,.webp,.gif,.txt';

const fileTypeIcon = (mime: string) => {
  if (mime === 'application/pdf') return <FileText size={18} style={{ color: '#f87171' }} />;
  if (mime.startsWith('image/')) return <Image size={18} style={{ color: '#34d399' }} />;
  return <File size={18} style={{ color: '#94a3b8' }} />;
};

const blobToBase64 = (blob: Blob): Promise<string> => new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.onloadend = () => resolve((reader.result as string).split(',')[1]);
  reader.onerror = reject;
  reader.readAsDataURL(blob);
});

const InputForm: React.FC<Props> = ({ onSubmitProposal, isLoading }) => {
  const [inputMethod, setInputMethod] = useState<InputTab>('quick');
  const [isProcessingFile, setIsProcessingFile] = useState(false);
  const [quickText, setQuickText] = useState('');
  const [isRecording, setIsRecording] = useState(false);
  const mediaRecorderRef = useRef<MediaRecorder | null>(null);
  const audioChunksRef = useRef<Blob[]>([]);

  // File upload state
  const [uploadedFile, setUploadedFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // ── Recording ──
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
        const b64 = await blobToBase64(new Blob(audioChunksRef.current, { type: 'audio/webm' }));
        onSubmitProposal('', b64);
      };
    }
  };

  // ── Quick submit ──
  const handleQuickSubmit = () => {
    if (!quickText.trim()) return;
    onSubmitProposal(quickText.trim());
  };

  // ── File handling ──
  const handleFileSelect = (file: File) => {
    if (!ACCEPTED_TYPES.includes(file.type)) {
      alert('不支援的檔案格式。請上傳 PDF、圖片（PNG/JPG/WEBP/GIF）或純文字（TXT）檔案。');
      return;
    }
    if (file.size > 20 * 1024 * 1024) {
      alert('檔案大小不能超過 20MB。');
      return;
    }
    setUploadedFile(file);
  };

  const handleFileDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) handleFileSelect(file);
  };

  const handleFileInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) handleFileSelect(file);
    e.target.value = '';
  };

  const handleProcessFile = async () => {
    if (!uploadedFile) return;
    setIsProcessingFile(true);
    try {
      const b64 = await blobToBase64(uploadedFile);
      const result = await parseFileInput(uploadedFile.name, uploadedFile.type, b64);
      // Use the parsed idea text as the proposal text
      const text = (result.data as any).idea || JSON.stringify(result.data);
      onSubmitProposal(text);
    } catch (err: any) {
      alert(err.message || '檔案解析失敗，請重試。');
    } finally {
      setIsProcessingFile(false);
    }
  };

  const busy = isLoading || isProcessingFile;

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
        <button onClick={() => setInputMethod('file')} className={`tab-btn ${inputMethod === 'file' ? 'active' : ''}`}>
          <Paperclip size={16} /> 上傳檔案
        </button>
      </div>

      <div className="form-card">

        {/* ── Quick Input ── */}
        {inputMethod === 'quick' && (
          <div className="quick-input-section">
            <div className="quick-heading">
              <h2>描述您的商業想法</h2>
              <p>無論是新創、品牌提案、產品計畫，直接說或打出來，AI 會立即掃描並開始深度對話分析。</p>
            </div>
            <div className="textarea-wrap">
              <textarea
                className="quick-textarea"
                placeholder={isRecording ? '正在聆聽...' : '例如：我想推出一個 AI 客製化寵物飲食訂閱服務，目標是有機飼主市場，預計第一年營收 300 萬...'}
                value={quickText}
                onChange={(e) => setQuickText(e.target.value)}
                disabled={isRecording || busy}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleQuickSubmit(); }
                }}
              />
              {isRecording && (
                <div className="recording-badge">
                  <div className="recording-dot" /> 錄音中
                </div>
              )}
            </div>
            <div className="mic-section">
              {!isRecording ? (
                <button onClick={startRecording} disabled={busy} className="mic-btn" title="開始錄音">
                  <Mic size={24} />
                </button>
              ) : (
                <button onClick={stopRecording} className="mic-btn-stop" title="停止並分析">
                  <StopCircle size={32} />
                </button>
              )}
              <span className="mic-hint">{isRecording ? '點擊停止說話' : '點擊麥克風開始說話'}</span>
            </div>
            <div className="quick-footer">
              <button
                onClick={handleQuickSubmit}
                disabled={busy || (!quickText.trim() && !isRecording)}
                className={`process-btn ${busy || (!quickText.trim() && !isRecording) ? 'disabled' : 'enabled'}`}
              >
                {busy
                  ? <><Loader2 size={20} className="spin-icon" /> 掃描中...</>
                  : <><Sparkles size={20} /> 開始 AI 掃描 <ArrowRight size={18} /></>}
              </button>
            </div>
          </div>
        )}

        {/* ── File Upload ── */}
        {inputMethod === 'file' && (
          <div className="file-upload-section">
            <div className="quick-heading">
              <h2>上傳提案文件</h2>
              <p>上傳 PDF 商業計畫書、截圖、或純文字檔案，AI 會自動提取關鍵資訊並進入深度對話分析。</p>
            </div>

            <div
              className={`file-drop-zone ${isDragging ? 'file-drop-zone-active' : ''} ${uploadedFile ? 'file-drop-zone-filled' : ''}`}
              onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={handleFileDrop}
              onClick={() => !uploadedFile && fileInputRef.current?.click()}
            >
              <input
                ref={fileInputRef}
                type="file"
                accept={ACCEPTED_EXTENSIONS}
                onChange={handleFileInputChange}
                style={{ display: 'none' }}
              />
              {!uploadedFile ? (
                <div className="file-drop-empty">
                  <Paperclip size={36} style={{ color: '#475569', marginBottom: 12 }} />
                  <div className="file-drop-title">拖放檔案至此，或點擊選擇</div>
                  <div className="file-drop-hint">支援格式：PDF · PNG · JPG · WEBP · TXT　最大 20MB</div>
                </div>
              ) : (
                <div className="file-drop-filled">
                  {fileTypeIcon(uploadedFile.type)}
                  <div className="file-info">
                    <div className="file-name">{uploadedFile.name}</div>
                    <div className="file-size">{(uploadedFile.size / 1024).toFixed(0)} KB · {uploadedFile.type}</div>
                  </div>
                  <button
                    className="file-remove-btn"
                    onClick={(e) => { e.stopPropagation(); setUploadedFile(null); }}
                    title="移除檔案"
                  >
                    <X size={16} />
                  </button>
                </div>
              )}
            </div>

            <div className="file-format-list">
              <div className="file-format-item"><FileText size={14} style={{ color: '#f87171' }} /> PDF 商業計畫書、提案書</div>
              <div className="file-format-item"><Image size={14} style={{ color: '#34d399' }} /> 截圖、簡報圖片（PNG/JPG/WEBP）</div>
              <div className="file-format-item"><File size={14} style={{ color: '#94a3b8' }} /> 純文字檔案（TXT）</div>
            </div>

            <div className="quick-footer">
              <button
                onClick={handleProcessFile}
                disabled={!uploadedFile || busy}
                className={`process-btn ${!uploadedFile || busy ? 'disabled' : 'enabled'}`}
              >
                {isProcessingFile
                  ? <><Loader2 size={20} className="spin-icon" /> AI 解析檔案中...</>
                  : <><Sparkles size={20} /> AI 解析並開始分析 <ArrowRight size={18} /></>}
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default InputForm;
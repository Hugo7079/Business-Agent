import React, { useState } from 'react';
import { HistoryRecord } from '../types';
import { Clock, Trash2, ChevronRight, X, AlertTriangle } from 'lucide-react';

interface Props {
  isOpen: boolean;
  onClose: () => void;
  records: HistoryRecord[];
  onLoad: (record: HistoryRecord) => void;
  onDelete: (id: string) => void;
  onClearAll: () => void;
}

const HistoryPanel: React.FC<Props> = ({ isOpen, onClose, records, onLoad, onDelete, onClearAll }) => {
  const [confirmClear, setConfirmClear] = useState(false);

  const formatDate = (ts: number) => {
    const d = new Date(ts);
    return `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
  };

  const getScoreColor = (s: number) => s >= 80 ? '#34d399' : s >= 60 ? '#60a5fa' : s >= 40 ? '#fbbf24' : '#f87171';

  return (
    <>
      {/* 背景遮罩 */}
      {isOpen && <div className="history-backdrop" onClick={onClose} />}

      {/* 側邊面板 */}
      <div className={`history-panel${isOpen ? ' open' : ''}`}>
        <div className="history-panel-header">
          <div className="history-panel-title">
            <Clock size={18} style={{ color: '#60a5fa' }} />
            歷史提案記錄
          </div>
          <button className="history-close-btn" onClick={onClose}><X size={20} /></button>
        </div>

        {records.length === 0 ? (
          <div className="history-empty">
            <Clock size={40} style={{ color: '#334155', marginBottom: '12px' }} />
            <p>尚無歷史記錄</p>
            <span>完成分析後會自動儲存在這裡</span>
          </div>
        ) : (
          <>
            <div className="history-list">
              {records.map(r => (
                <div key={r.id} className="history-item">
                  <button className="history-item-main" onClick={() => { onLoad(r); onClose(); }}>
                    <div className="history-item-top">
                      <span className="history-item-title">{r.title}</span>
                      <span className="history-item-score" style={{ color: getScoreColor(r.result.successProbability) }}>
                        {r.result.successProbability}%
                      </span>
                    </div>
                    <div className="history-item-meta">
                      <Clock size={11} />
                      {formatDate(r.createdAt)}
                      <ChevronRight size={13} style={{ marginLeft: 'auto' }} />
                    </div>
                  </button>
                  <button className="history-delete-btn" onClick={() => onDelete(r.id)} title="刪除此記錄">
                    <Trash2 size={14} />
                  </button>
                </div>
              ))}
            </div>

            <div className="history-panel-footer">
              {confirmClear ? (
                <div className="history-confirm-clear">
                  <AlertTriangle size={14} style={{ color: '#fbbf24' }} />
                  <span>確定要清除所有記錄？</span>
                  <button className="history-confirm-yes" onClick={() => { onClearAll(); setConfirmClear(false); }}>確定</button>
                  <button className="history-confirm-no" onClick={() => setConfirmClear(false)}>取消</button>
                </div>
              ) : (
                <button className="history-clear-btn" onClick={() => setConfirmClear(true)}>
                  <Trash2 size={13} /> 清除所有記錄
                </button>
              )}
            </div>
          </>
        )}
      </div>
    </>
  );
};

export default HistoryPanel;

import { AnalysisResult } from '../types';
import { summarizeForSlides, generateThemeImage } from './geminiService';
import { buildPptxInstance } from './pptxService';

/**
 * 直接把已建好的 PPTX 輸出成 blob 下載。
 * 副檔名固定 .pptx（純前端無法真正轉 PDF binary）。
 */
export const generatePdf = async (
  result: AnalysisResult,
  onProgress?: (stage: string) => void,
): Promise<void> => {
  onProgress?.('精簡投影片內文...');
  const summarized = await summarizeForSlides(result);

  onProgress?.('生成主題插圖...');
  const themeImg = await generateThemeImage(
    summarized.executiveSummary,
    summarized.marketAnalysis.description,
  );

  onProgress?.('組裝投影片...');
  const pptx = buildPptxInstance(summarized, themeImg);

  onProgress?.('輸出檔案...');
  const blob = await pptx.write({ outputType: 'blob' }) as Blob;

  // 建立隱藏連結觸發下載
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'OmniView_商業提案報告.pptx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
};

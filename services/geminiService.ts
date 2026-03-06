import { GoogleGenAI, Type, Schema, PersonGeneration } from "@google/genai";
import { AnalysisMode, BusinessInput, AnalysisResult, ParsedInputResponse, Phase1DiagnosisResult, Phase2Response, Phase3DeepReport, InitialScanResult } from "../types";

const analysisSchema: Schema = {
  type: Type.OBJECT,
  properties: {
    successProbability: { type: Type.NUMBER, description: '0到100之間的整數，例如 65 代表 65%，絕對不可以用小數如 0.65' },
    executiveSummary: { type: Type.STRING, description: '5到7個重點，每點用換行符\\n分隔，每點為完整的分析句子，充分說明理由與依據' },
    marketAnalysis: {
      type: Type.OBJECT,
      properties: {
        size: { type: Type.STRING, description: '市場規模，必須極度精簡，最多20個中文字，格式如：全球 XXX 市場規模約 $XXX（XXXX年）' },
        growthRate: { type: Type.STRING, description: '成長率，必須極度精簡，最多20個中文字，格式如：CAGR XX%（XXXX–XXXX年）' },
        description: { type: Type.STRING, description: '市場洞察，用3到5個重點，每點用換行符\\n分隔，每點為完整分析句子' },
      },
      required: ['size', 'growthRate', 'description'],
    },
    competitors: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          name: { type: Type.STRING, description: '競爭對手名稱' },
          strength: { type: Type.STRING, description: '核心優勢，完整說明其競爭優勢的原因與影響' },
          weakness: { type: Type.STRING, description: '核心弱點，完整說明其弱點的原因與影響' },
        },
        required: ['name', 'strength', 'weakness'],
      },
    },
    roadmap: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          phase: { type: Type.STRING, description: '階段名稱，最多10個中文字' },
          timeframe: { type: Type.STRING, description: '時間範圍，最多8個中文字，如：第1-6個月' },
          technology: { type: Type.STRING, description: '核心技術，極度精簡，最多2-3行，每行15字以內，只列最關鍵技術棧' },
          product: { type: Type.STRING, description: '產品里程碑，極度精簡，最多2-3行，每行15字以內，只列最關鍵交付物' },
        },
        required: ['phase', 'timeframe', 'technology', 'product'],
      },
    },
    financials: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          year: { type: Type.STRING, description: '年度，例如「第1年」' },
          revenue: { type: Type.NUMBER },
          profit: { type: Type.NUMBER },
          costs: { type: Type.NUMBER },
        },
        required: ['year', 'revenue', 'profit', 'costs'],
      },
    },
    breakEvenPoint: { type: Type.STRING, description: '損益平衡，必須極度精簡，最多24個中文字，格式如：預計第X年達成，月營收需達 $XXX' },
    risks: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          risk: { type: Type.STRING, description: '風險名稱，最多12個中文字，簡短有力' },
          impact: { type: Type.STRING, enum: ["High", "Medium", "Low"] },
          mitigation: { type: Type.STRING, description: '因應對策，精簡說明，最多40個中文字' },
        },
        required: ['risk', 'impact', 'mitigation'],
      },
    },
    personaEvaluations: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          role: { type: Type.STRING, description: '角色名稱，最多8個中文字' },
          icon: { type: Type.STRING, description: '請回傳與角色相關的 Lucide icon 名稱 (例如 "Briefcase", "Calculator", "Scale" 等)，不要回傳 emoji' },
          perspective: { type: Type.STRING, description: '此角色對提案的完整看法，用3到4個重點，每點用換行符\\n分隔，每點為完整分析句子，充分表達該角色的立場與理由' },
          score: { type: Type.NUMBER, description: '0到100之間的整數評分' },
          keyQuote: { type: Type.STRING, description: '此角色最具代表性的一句話，最多20個中文字，像口號或金句，直接表達核心立場' },
          concern: { type: Type.STRING, description: '最主要的擔憂，精簡說明，最多50個中文字' },
        },
        required: ['role', 'icon', 'perspective', 'score', 'keyQuote', 'concern'],
      },
    },
    teamAnalysis: { type: Type.STRING, description: '團隊能力分析，完整評估現有團隊與所需人才缺口' },
    finalVerdicts: {
      type: Type.OBJECT,
      properties: {
        aggressive: { type: Type.STRING, description: '激進觀點，用4到6個重點，每點用換行符\\n分隔，每點為完整的分析句子，充分闡述激進策略的論點、機會與行動建議' },
        balanced: { type: Type.STRING, description: '平衡觀點，用4到6個重點，每點用換行符\\n分隔，每點為完整的分析句子，充分闡述平衡策略的論點、取捨與執行建議' },
        conservative: { type: Type.STRING, description: '保守觀點，用4到6個重點，每點用換行符\\n分隔，每點為完整的分析句子，充分闡述保守策略的論點、風險控制與謹慎行動方案' },
      },
      required: ['aggressive', 'balanced', 'conservative'],
    },
    continueToIterate: { type: Type.STRING, description: '持續迭代建議，用4到6個重點，每點用換行符\\n分隔，每點為完整的分析句子，充分闡述下一步的具體行動與優化方向' },
  },
  required: [
    "successProbability", "executiveSummary", "marketAnalysis", "competitors",
    "roadmap", "financials", "breakEvenPoint", "risks",
    "personaEvaluations", "teamAnalysis", "finalVerdicts", "continueToIterate"
  ],
};

const businessParserSchema: Schema = {
  type: Type.OBJECT,
  properties: {
    data: {
      type: Type.OBJECT,
      properties: {
        idea: { type: Type.STRING },
        marketData: { type: Type.STRING },
        productDetails: { type: Type.STRING },
        painPoints: { type: Type.STRING },
        targetConsumer: { type: Type.STRING },
        financialContext: { type: Type.STRING },
      },
      required: ["idea", "marketData", "productDetails", "painPoints", "targetConsumer", "financialContext"]
    },
    feedback: { type: Type.STRING }
  },
  required: ["data", "feedback"]
};

// 統一的 API 配額錯誤偵測
const handleApiError = (error: any): never => {
  const msg: string = error?.message || String(error);
  const status = error?.status || error?.statusCode || '';
  
  // 配額/速率限制错误（429）
  if (
    status === 429 ||
    msg.includes('429') ||
    (msg.includes('RESOURCE_EXHAUSTED') && msg.includes('quota'))
  ) {
    throw new Error(
      'API 配額已用盡\n\n您的 Gemini API Key 已達到免費方案的用量上限。\n\n解決方式：\n• 等待配額重置（通常每分鐘或每天重置）\n• 前往 Google AI Studio 升級至付費方案\n• 使用不同的 API Key'
    );
  }
  
  // API Key 验证错误
  if (
    status === 401 ||
    msg.includes('API_KEY_INVALID') ||
    msg.includes('UNAUTHENTICATED') ||
    msg.includes('API key not valid')
  ) {
    throw new Error('API Key 無效\n\n請點擊右上角設定按鈕，重新輸入正確的 Gemini API Key。');
  }
  
  // 模型不存在或不可用
  if (
    status === 404 ||
    msg.includes('NOT_FOUND') ||
    msg.includes('Model not found') ||
    msg.includes('does not exist')
  ) {
    throw new Error('AI 模型暫時無法使用\n\n目前指定的 Gemini 模型不可用，可能是 Google 更新了模型版本。\n請稍後再試，或聯繫開發者更新模型設定。');
  }
  
  // 权限不足（可能是免费 Key 不支持此模型）
  if (
    msg.includes('PERMISSION_DENIED') ||
    msg.includes('permission') ||
    msg.includes('not permitted') ||
    msg.includes('not enabled')
  ) {
    throw new Error('API 權限不足\n\n您的 API Key 可能無法存取此功能。\n\n解決方式：\n• 前往 Google AI Studio 檢查 API 設定\n• 升級至付費方案以存取 Pro 模型\n• 確認已啟用所有必要的 API');
  }
  
  // 网络超时或连接错误
  if (
    msg.includes('DEADLINE_EXCEEDED') ||
    msg.includes('timeout') ||
    msg.includes('net::ERR') ||
    msg.includes('Failed to fetch')
  ) {
    throw new Error('網路連線超時\n\n請檢查您的網際網路連接，然後重試。\n\n如果問題持續，可能是 Google API 伺服器暫時不可用。');
  }
  
  // 通用错误：显示实际错误信息便于调试
  throw new Error(
    `API 請求失敗\n\n錯誤訊息：${msg}\n\n如果問題持續，請聯繫開發者並提供上述錯誤訊息。`
  );
};

const getApiKey = (): string => {
  const key = localStorage.getItem('GEMINI_API_KEY') || process.env.API_KEY || '';
  if (!key) throw new Error('找不到 API Key。請點擊右上角設定按鈕輸入您的 Google Gemini API Key。');
  return key;
};

export const parseBusinessIdea = async (
  _mode: AnalysisMode,
  textInput: string,
  audioBase64?: string
): Promise<ParsedInputResponse> => {
  const apiKey = getApiKey();
  const ai = new GoogleGenAI({ apiKey });

  const prompt = `你是一位專業的商業分析助理。
使用者將描述一個商業想法或提案（文字或語音）。

請將其整理成以下結構化欄位：
- idea: 核心想法一句話摘要
- marketData: 市場規模、趨勢、CAGR 等
- productDetails: 產品或服務的具體細節
- painPoints: 解決的市場痛點
- targetConsumer: 目標客群描述
- financialContext: 預算、營收目標、財務背景

若使用者未提供某欄位，請根據上下文合理推斷並填入。
在 feedback 欄位說明你提取了什麼、推斷了什麼。
【重要】所有輸出內容（字串）中，絕對不可以使用 Emoji 表情符號（如 💰, 🚀, 👍, 📊 等）。
所有輸出必須使用繁體中文。

用戶輸入：${textInput || '（處理語音輸入中）'}`;

  const contents: any[] = [];
  if (audioBase64) {
    contents.push({ inlineData: { mimeType: 'audio/webm', data: audioBase64 } });
  }
  contents.push({ text: prompt });

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: [{ role: 'user', parts: contents }],
      config: { responseMimeType: 'application/json', responseSchema: businessParserSchema },
    });
    return JSON.parse(response.text!) as ParsedInputResponse;
  } catch (error) {
    handleApiError(error);
  }
};

export const analyzeBusiness = async (
  _mode: AnalysisMode,
  data: BusinessInput
): Promise<AnalysisResult> => {
  const apiKey = getApiKey();
  const ai = new GoogleGenAI({ apiKey });

  const prompt = `你是 OmniView，一個進階 AI 商業模擬系統。
你的任務是模擬一場「董事會會議」，由不同 AI 角色對以下商業提案進行全方位 360° 嚴格評估。

【絕對規定 — 無論提案品質高低，都必須完整填寫以下所有欄位，不得留空或省略】
【重要禁令】所有輸出內容（字串）中，絕對不可以使用 Emoji 表情符號（如 💰, 🚀, 👍, 📊 等）。

必填數量要求：
- competitors：至少 3 個競爭對手（若市場上存在替代品或間接競爭者也算）
- roadmap：至少 3 個階段（即使提案很差，也要描述「如果要做」的路線）
- financials：必須有 3 筆資料，分別代表第 1、2、3 年的悲觀預測數字（若提案差，數字可以很低甚至是負數，但不能不填）
- risks：至少 4 個風險項目
- personaEvaluations：必須包含以下 5 個角色，每個都要有完整的 perspective、keyQuote、concern：
  1. 風險投資人（icon: "investor"）
  2. 競爭對手（icon: "sword"）
  3. 目標消費者（icon: "consumer"）
  4. 供應商（icon: "supplier"）
  5. 員工（icon: "employee"）

評估範圍：
1. 市場規模與成長性（即使市場小，也要給出估計數字）
2. 競爭對手分析（列出直接與間接競爭者）
3. 產品與技術路線圖（描述理論上的執行步驟）
4. 團隊能力需求
5. 財務預測（第 1-3 年的營收/成本/淨利，低分提案給悲觀數字即可）
6. 風險因素與因應策略
7. 三種最終裁決（激進/平衡/保守，每個至少 80 字的具體分析）

successProbability 的評分依據：市場可行性 30% + 競爭優勢 25% + 執行可行性 25% + 財務合理性 20%

請給出具體數字，避免空泛描述。所有輸出必須使用繁體中文。

商業提案資料：
- 核心想法：${data.idea}
- 市場資料：${data.marketData}
- 產品細節：${data.productDetails}
- 市場痛點：${data.painPoints}
- 目標消費者：${data.targetConsumer}
- 財務背景：${data.financialContext}`;

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      config: {
        responseMimeType: 'application/json',
        responseSchema: analysisSchema,
      },
    });
    return JSON.parse(response.text!) as AnalysisResult;
  } catch (error) {
    handleApiError(error);
  }
};

const slidesSummarySchema: Schema = {
  type: Type.OBJECT,
  properties: {
    executiveSummary: {
      type: Type.STRING,
      description: '5到7個重點，每點用\\n分隔，每點為完整的分析句子，極度精簡，避免冗長',
    },
    marketSize: {
      type: Type.STRING,
      description: 'TAM市場規模，絕對不能超過24個中文字（含數字與符號），格式如：全球 XXX 市場約 $XXX（XXXX年）',
    },
    marketGrowthRate: {
      type: Type.STRING,
      description: 'CAGR年均成長率，絕對不能超過24個中文字（含數字與符號），格式如：CAGR XX%（XXXX–XXXX年）',
    },
    marketDescription: {
      type: Type.STRING,
      description: '3到5個市場洞察重點，每點用\\n分隔，每點為完整的分析句子，極度精簡，避免冗長',
    },
    competitors: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          name:     { type: Type.STRING, description: '競爭對手名稱' },
          strength: { type: Type.STRING, description: '核心優勢，極度精簡，最多30個中文字' },
          weakness: { type: Type.STRING, description: '核心弱點，極度精簡，最多30個中文字' },
        },
        required: ['name', 'strength', 'weakness'],
      },
    },
    roadmap: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          phase:      { type: Type.STRING, description: '階段名稱，最多10個中文字' },
          timeframe:  { type: Type.STRING, description: '時間範圍，最多8個中文字' },
          technology: { type: Type.STRING, description: '核心技術，絕對不超過3行，每行絕對不超過12個中文字，只寫技術名稱' },
          product:    { type: Type.STRING, description: '產品里程碑，絕對不超過3行，每行絕對不超過12個中文字，只寫交付物名稱' },
        },
        required: ['phase', 'timeframe', 'technology', 'product'],
      },
    },
    breakEvenPoint: {
      type: Type.STRING,
      description: '損益平衡點，絕對不能超過24個中文字（含數字與符號），只保留時間點與關鍵數字，格式如：第X年達成，月營收需達 $XXX',
    },
    risks: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          risk:       { type: Type.STRING, description: '風險名稱，最多12個中文字' },
          impact:     { type: Type.STRING, enum: ['High', 'Medium', 'Low'] },
          mitigation: { type: Type.STRING, description: '因應對策，最多40個中文字' },
        },
        required: ['risk', 'impact', 'mitigation'],
      },
    },
    personaEvaluations: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          role:        { type: Type.STRING, description: '角色名稱，最多8個中文字' },
          icon:        { type: Type.STRING },
          score:       { type: Type.NUMBER },
          keyQuote:    { type: Type.STRING, description: '最具代表性的一句話，最多20個中文字' },
          concern:     { type: Type.STRING, description: '最主要擔憂，最多50個中文字' },
          perspective: { type: Type.STRING, description: '3到4個重點，每點用\\n分隔，每點為完整分析句子，每點最多25字' },
        },
        required: ['role', 'icon', 'score', 'keyQuote', 'concern', 'perspective'],
      },
    },
    finalVerdicts: {
      type: Type.OBJECT,
      properties: {
        aggressive:   { type: Type.STRING, description: '4到6個重點，每點用\\n分隔，每點最多25字，極度精簡' },
        balanced:     { type: Type.STRING, description: '4到6個重點，每點用\\n分隔，每點最多25字，極度精簡' },
        conservative: { type: Type.STRING, description: '4到6個重點，每點用\\n分隔，每點最多25字，極度精簡' },
      },
      required: ['aggressive', 'balanced', 'conservative'],
    },
    continueToIterate: { type: Type.STRING, description: '4到6個重點，每點用\\n分隔，每點最多25字，極度精簡' },
  },
  required: [
    'executiveSummary', 'marketSize', 'marketGrowthRate', 'marketDescription',
    'competitors', 'roadmap', 'breakEvenPoint', 'risks',
    'personaEvaluations', 'finalVerdicts', 'continueToIterate'
  ],
};

export const summarizeForSlides = async (result: AnalysisResult): Promise<AnalysisResult> => {
  const apiKey = getApiKey();
  const ai = new GoogleGenAI({ apiKey });

  const prompt = `你是一位專業的投影片簡報設計師，擅長將長篇報告壓縮成投影片可用的超短文字。

以下是一份商業分析報告（JSON格式），請按照嚴格規則精簡每個欄位：

【絕對字數上限，違反即為錯誤】
- marketSize（TAM市場規模）：絕對不超過 24 字，只保留數字與單位
- marketGrowthRate（CAGR年均成長）：絕對不超過 24 字，只保留數字與年份
- breakEvenPoint（損益平衡點）：絕對不超過 24 字，只保留時間點與關鍵數字
- roadmap 每個階段的 technology：絕對不超過 3 行，每行不超過 12 字，只寫技術名稱
- roadmap 每個階段的 product：絕對不超過 3 行，每行不超過 12 字，只寫交付物名稱
- 其他重點條列：每點不超過 25 字

【通用規則】
- 去除所有「由於」「因此」「透過」「藉由」等連接詞
- 去除所有說明性語句，只保留核心事實與數字
- 數字、專有名詞、品牌名稱必須保留
- 所有輸出必須使用繁體中文
- 多個重點之間用 \\n 分隔
- 【絕對禁止】所有輸出字串中不得包含任何 Emoji 表情符號

原始資料：
${JSON.stringify(result, null, 2)}`;

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      config: {
        responseMimeType: 'application/json',
        responseSchema: slidesSummarySchema,
      },
    });

    const summary = JSON.parse(response.text!);

    // 合併回原始 result，只覆蓋文字欄位，數字與財務資料維持原樣
    return {
      ...result,
      executiveSummary: summary.executiveSummary,
      marketAnalysis: {
        ...result.marketAnalysis,
        size:        summary.marketSize,
        growthRate:  summary.marketGrowthRate,
        description: summary.marketDescription,
      },
      competitors: summary.competitors,
      roadmap: summary.roadmap,
      breakEvenPoint: summary.breakEvenPoint,
      risks: summary.risks,
      personaEvaluations: result.personaEvaluations.map((p, i) => ({
        ...p,
        ...(summary.personaEvaluations[i] ?? {}),
      })),
      finalVerdicts: summary.finalVerdicts,
      continueToIterate: summary.continueToIterate,
    };
  } catch (error) {
    handleApiError(error);
  }
};

export const generateThemeImage = async (
  summary: string,
  marketDesc: string,
): Promise<string | null> => {
  const apiKey = getApiKey();
  const ai = new GoogleGenAI({ apiKey });

  // 用 LLM 產生圖像主題短語
  let themeKeyword = 'modern business strategy, professional team, innovation';
  try {
    const kwResp = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: [{
        role: 'user',
        parts: [{ text:
          `Based on this business proposal summary, write ONE short English phrase (max 12 words) describing the core industry/theme for an image generation prompt. Only output the phrase, nothing else.\n\nSummary: ${summary}\nMarket: ${marketDesc}`
        }]
      }],
    });
    const kw = kwResp.text?.trim().replace(/['"]/g, '');
    if (kw && kw.length > 3) themeKeyword = kw;
  } catch (e) {
    console.warn('[Imagen] keyword generation failed, using default:', e);
  }

  const imagePrompt =
    `A cinematic, ultra-detailed dark-themed illustration for a business pitch deck. ` +
    `Theme: ${themeKeyword}. ` +
    `Style: dark background, neon accent lighting, professional and futuristic, ` +
    `no text, no people faces, abstract concept art. ` +
    `Mood: innovative, inspiring, premium investor presentation.`;

  // 強制使用 imagen-3.0-generate-001
  const modelName = 'imagen-3.0-fast-generate-001';
  try {
    const response = await ai.models.generateImages({
      model: modelName,
      prompt: imagePrompt,
      config: {
        numberOfImages: 1,
        aspectRatio: '16:9',
        personGeneration: PersonGeneration.DONT_ALLOW,
      },
    });
    const b64 = response.generatedImages?.[0]?.image?.imageBytes;
    if (!b64) {
      console.warn('[Imagen] No image in response from', modelName, JSON.stringify(response).slice(0,300));
      return null;
    }
    return `data:image/png;base64,${b64}`;
  } catch (e: any) {
    console.warn(`[Imagen] generateImages failed for ${modelName}:`, e);
    return null;
  }
};

// 新增：根據執行摘要生成一個簡短有力的提案標題
export const generateProposalTitle = async (executiveSummary: string): Promise<string> => {
  const apiKey = getApiKey();
  const ai = new GoogleGenAI({ apiKey });
  const prompt = `請根據下列執行摘要為此商業提案取一個簡短有力的標題，最多10個中文字，只回傳標題本身。
【重要】絕對不可以使用任何 Emoji 表情符號。
執行摘要：${executiveSummary}`;
  try {
    const resp = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      config: { responseMimeType: 'text/plain' },
    });
    return resp.text?.trim() || '商業提案';
  } catch (error) {
    handleApiError(error);
  }
};

// =====================================================
// 檔案輸入解析 (File Input Parser)
// =====================================================

export const parseFileInput = async (
  fileName: string,
  mimeType: string,
  fileBase64: string
): Promise<ParsedInputResponse> => {
  const apiKey = getApiKey();
  const ai = new GoogleGenAI({ apiKey });

  const isTextFile = mimeType === 'text/plain';

  const prompt = `你是一位專業的商業分析助理。
使用者上傳了一份檔案（${fileName}），內容可能是商業計畫書、提案文件、簡報或相關資料。

請從這份檔案中提取商業相關資訊，整理成以下結構化欄位：
- idea: 核心想法一句話摘要
- marketData: 市場規模、趨勢、CAGR 等
- productDetails: 產品或服務的具體細節
- painPoints: 解決的市場痛點
- targetConsumer: 目標客群描述
- financialContext: 預算、營收目標、財務背景

若某欄位未在文件中提及，請根據上下文合理推斷並填入。
在 feedback 欄位說明你從文件中提取了什麼，以及有哪些是推斷補充的。
【重要】所有輸出內容（字串）中，絕對不可以使用 Emoji 表情符號。
所有輸出必須使用繁體中文。`;

  const parts: any[] = [];

  if (isTextFile) {
    // 文字檔：decode base64 再附加文字
    const decoded = decodeURIComponent(escape(atob(fileBase64)));
    parts.push({ text: prompt + '\n\n檔案內容：\n' + decoded.slice(0, 8000) });
  } else {
    // PDF / 圖片：以 inlineData 傳入
    parts.push({ inlineData: { mimeType, data: fileBase64 } });
    parts.push({ text: prompt });
  }

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: [{ role: 'user', parts }],
      config: { responseMimeType: 'application/json', responseSchema: businessParserSchema },
    });
    return JSON.parse(response.text!) as ParsedInputResponse;
  } catch (error) {
    handleApiError(error);
  }
};

// =====================================================
// Initial Scan: 輕量初步掃描（使用者提案後立刻觸發）
// =====================================================

const initialScanSchema: Schema = {
  type: Type.OBJECT,
  properties: {
    successProbability: { type: Type.NUMBER, description: '初步成功率估計，0到100之間的整數' },
    coreUnderstanding: { type: Type.STRING, description: '對提案的核心理解，2-3句話，說明你理解到的是什麼提案、解決什麼問題、目標族群是誰' },
    marketSnapshot: {
      type: Type.OBJECT,
      properties: {
        size:       { type: Type.STRING, description: '市場規模，一行，含數字，如：台灣夜市餐飲市場年產值約 800 億台幣' },
        growthRate: { type: Type.STRING, description: '成長率，一行，如：年均成長約 3-5%' },
        insight:    { type: Type.STRING, description: '最關鍵的一句市場洞察，說明這個市場為什麼值得進入或需要警惕' },
      },
      required: ['size', 'growthRate', 'insight'],
    },
    competitors: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          name:  { type: Type.STRING, description: '競爭對手名稱（直接或間接競爭者）' },
          threat: { type: Type.STRING, description: '一句話說明對本提案的威脅程度與原因，最多30字' },
        },
        required: ['name', 'threat'],
      },
      description: '2-4個主要競爭對手',
    },
    infoGaps: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          field:    { type: Type.STRING, description: '缺口欄位名稱，如：資金來源、首批客戶、差異化優勢' },
          question: { type: Type.STRING, description: '具體追問問題，讓使用者知道 AI 需要哪些資訊才能給出更準確的判斷' },
          priority: { type: Type.STRING, enum: ['HIGH', 'MEDIUM', 'LOW'] },
        },
        required: ['field', 'question', 'priority'],
      },
      description: '3-5個最關鍵的資訊缺口，按優先級排序',
    },
    extractedData: {
      type: Type.OBJECT,
      properties: {
        idea:             { type: Type.STRING },
        marketData:       { type: Type.STRING },
        productDetails:   { type: Type.STRING },
        painPoints:       { type: Type.STRING },
        targetConsumer:   { type: Type.STRING },
        financialContext: { type: Type.STRING },
      },
      required: ['idea', 'marketData', 'productDetails', 'painPoints', 'targetConsumer', 'financialContext'],
    },
  },
  required: ['successProbability', 'coreUnderstanding', 'marketSnapshot', 'competitors', 'infoGaps', 'extractedData'],
};

export const runInitialScan = async (
  textInput: string,
  audioBase64?: string
): Promise<InitialScanResult> => {
  const apiKey = getApiKey();
  const ai = new GoogleGenAI({ apiKey });

  const prompt = `你是 OmniView，一個頂尖的 AI 商業顧問。使用者剛剛提交了他的商業提案，你的任務是立刻給出「初步診斷」。

【重要原則】
- 這不是完整分析，只是第一印象的快速評估
- 給出初步成功率（誠實但有建設性）
- 說明你理解到的提案核心（讓使用者確認你有看懂）
- 給出市場快照（規模、成長、一句洞察）
- 列出主要競爭對手（2-4個，說明威脅）
- 列出最需要補充的資訊缺口（3-5個），並提出具體追問問題
- 從輸入中提取並結構化已有資訊到 extractedData

說話風格：專業、直接、像頂尖創投顧問在做盡職調查的第一輪評估。
所有輸出必須使用繁體中文。

使用者提案：${textInput || '（處理語音輸入中）'}`;

  const contents: any[] = [];
  if (audioBase64) contents.push({ inlineData: { mimeType: 'audio/webm', data: audioBase64 } });
  contents.push({ text: prompt });

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: [{ role: 'user', parts: contents }],
      config: { responseMimeType: 'application/json', responseSchema: initialScanSchema },
    });
    const result = JSON.parse(response.text!) as InitialScanResult;
    if (result.successProbability < 1) result.successProbability = Math.round(result.successProbability * 100);
    result.successProbability = Math.min(100, Math.max(0, Math.round(result.successProbability)));
    return result;
  } catch (error) {
    handleApiError(error);
  }
};

// =====================================================
// Phase 1: 初步掃描 (Deep Scan)
// =====================================================

const phase1Schema: Schema = {
  type: Type.OBJECT,
  properties: {
    coreInsight: { type: Type.STRING, description: '最吸引人的一句話核心描述，點出最大亮點' },
    successRateEstimate: { type: Type.NUMBER, description: '初步成功率估計，0到100之間的整數' },
    riskWarnings: {
      type: Type.ARRAY,
      items: { type: Type.STRING, description: '具體風險警示，直接說實話，每條最多40字' },
    },
    infoGaps: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          field: { type: Type.STRING, description: '缺口欄位名稱，如：資金來源、首批客戶' },
          question: { type: Type.STRING, description: '對應的追問問題，具體且有洞察力' },
          priority: { type: Type.STRING, enum: ['HIGH', 'MEDIUM', 'LOW'] },
        },
        required: ['field', 'question', 'priority'],
      },
    },
    extractedData: {
      type: Type.OBJECT,
      properties: {
        idea: { type: Type.STRING },
        marketData: { type: Type.STRING },
        productDetails: { type: Type.STRING },
        painPoints: { type: Type.STRING },
        targetConsumer: { type: Type.STRING },
        financialContext: { type: Type.STRING },
      },
      required: ['idea', 'marketData', 'productDetails', 'painPoints', 'targetConsumer', 'financialContext'],
    },
  },
  required: ['coreInsight', 'successRateEstimate', 'riskWarnings', 'infoGaps', 'extractedData'],
};

export const runPhase1DeepScan = async (
  textInput: string,
  audioBase64?: string
): Promise<Phase1DiagnosisResult> => {
  const apiKey = getApiKey();
  const ai = new GoogleGenAI({ apiKey });

  const prompt = `你是 OmniView 深度掃描系統（Phase 1）。你的任務是對使用者的創業構想進行初步診斷。

請完成以下三件事：

1. **提取核心**：從雜亂的資訊中提取產品、市場、受眾、痛點，整理成 extractedData。
2. **產出診斷報告**：
   - coreInsight：目前提案最吸引人的一句話核心描述
   - successRateEstimate：根據現有資訊給出初步成功率（0-100整數）
   - riskWarnings：針對市場飽和度、競爭對手、邏輯漏洞給予「實話」警示，每條具體有力
3. **找出資訊缺口**：明確列出使用者沒提到的關鍵點（如資金來源、首批客戶、特有資源），按優先級排序，並為每個缺口準備一個精準的追問問題。

說話風格：專業、直接、不客套，像一位頂尖創投顧問在做盡職調查。
所有輸出必須使用繁體中文。

使用者輸入：${textInput || '（處理語音輸入中）'}`;

  const contents: any[] = [];
  if (audioBase64) {
    contents.push({ inlineData: { mimeType: 'audio/webm', data: audioBase64 } });
  }
  contents.push({ text: prompt });

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: [{ role: 'user', parts: contents }],
      config: { responseMimeType: 'application/json', responseSchema: phase1Schema },
    });
    const result = JSON.parse(response.text!) as Phase1DiagnosisResult;
    // 修正 successRateEstimate
    if (result.successRateEstimate < 1) result.successRateEstimate = Math.round(result.successRateEstimate * 100);
    result.successRateEstimate = Math.min(100, Math.max(0, Math.round(result.successRateEstimate)));
    return result;
  } catch (error) {
    handleApiError(error);
  }
};

// =====================================================
// Phase 2: 互動式補全 (Collaborative Inquiry)
// =====================================================

const phase2ResponseSchema: Schema = {
  type: Type.OBJECT,
  properties: {
    dynamicFeedback: { type: Type.STRING, description: '對使用者回答的即時動態回饋，說明這個資訊如何改變評估，最多80字，具體有洞察力' },
    updatedData: {
      type: Type.OBJECT,
      properties: {
        idea: { type: Type.STRING },
        marketData: { type: Type.STRING },
        productDetails: { type: Type.STRING },
        painPoints: { type: Type.STRING },
        targetConsumer: { type: Type.STRING },
        financialContext: { type: Type.STRING },
      },
      required: ['idea', 'marketData', 'productDetails', 'painPoints', 'targetConsumer', 'financialContext'],
    },
    completionRate: { type: Type.NUMBER, description: '目前資訊飽滿度，0到100之間的整數' },
    nextQuestions: {
      type: Type.ARRAY,
      items: { type: Type.STRING, description: '下一輪追問問題，精準且有洞察力' },
      description: '接下來要問的1到2個最關鍵問題。若 isReady 為 true 則可為空陣列。',
    },
    isReady: { type: Type.BOOLEAN, description: '資訊是否已達80%飽滿度，可進入Phase 3生成完整報告' },
  },
  required: ['dynamicFeedback', 'updatedData', 'completionRate', 'nextQuestions', 'isReady'],
};

export const runPhase2Reply = async (
  userReply: string,
  currentData: Phase1DiagnosisResult['extractedData'],
  conversationHistory: Array<{ role: string; content: string }>,
  audioBase64?: string
): Promise<Phase2Response> => {
  const apiKey = getApiKey();
  const ai = new GoogleGenAI({ apiKey });

  const historyText = conversationHistory
    .map(m => `[${m.role === 'assistant' ? 'OmniView' : '使用者'}] ${m.content}`)
    .join('\n');

  const prompt = `你是 OmniView 深度掃描系統（Phase 2：互動式補全）。

目前已整理的商業提案資料：
${JSON.stringify(currentData, null, 2)}

對話歷史：
${historyText}

使用者最新回覆：${userReply || '（處理語音輸入中）'}

你的任務：
1. **dynamicFeedback**：針對使用者的回答給予即時動態回饋，說明這個資訊如何改變你對提案的評估（例如：「有了這份通路名單，你的 Go-to-Market 變得非常清晰」）。直接、有見地、不超過80字。
2. **updatedData**：將使用者回答的資訊整合進已有資料，更新對應欄位。
3. **completionRate**：重新評估資訊飽滿度（0-100整數），考慮：核心想法完整性、市場資料充分性、財務背景清晰度、競爭優勢獨特性、執行可行性。
4. **nextQuestions**：若尚未達到80%，提出1-2個最關鍵的追問問題。若已達80%或使用者要求生成報告，則設為空陣列。
5. **isReady**：若 completionRate >= 80 或使用者明確要求生成報告，設為 true。

說話風格：專業、直接，像頂尖創投顧問。所有輸出使用繁體中文。`;

  const contents: any[] = [];
  if (audioBase64) {
    contents.push({ inlineData: { mimeType: 'audio/webm', data: audioBase64 } });
  }
  contents.push({ text: prompt });

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: [{ role: 'user', parts: contents }],
      config: { responseMimeType: 'application/json', responseSchema: phase2ResponseSchema },
    });
    const result = JSON.parse(response.text!) as Phase2Response;
    if (result.completionRate < 1) result.completionRate = Math.round(result.completionRate * 100);
    result.completionRate = Math.min(100, Math.max(0, Math.round(result.completionRate)));
    return result;
  } catch (error) {
    handleApiError(error);
  }
};

// =====================================================
// Phase 3: 全維度深度分析報告 (The Master Plan)
// =====================================================

const phase3Schema: Schema = {
  type: Type.OBJECT,
  properties: {
    businessModelCanvas: {
      type: Type.OBJECT,
      properties: {
        keyPartners:           { type: Type.STRING, description: '關鍵合作夥伴，列出3-5個，每個用\\n分隔' },
        keyActivities:         { type: Type.STRING, description: '關鍵活動，列出3-5個，每個用\\n分隔' },
        keyResources:          { type: Type.STRING, description: '關鍵資源，列出3-5個，每個用\\n分隔' },
        valuePropositions:     { type: Type.STRING, description: '價值主張，列出2-4個核心價值，每個用\\n分隔' },
        customerRelationships: { type: Type.STRING, description: '客戶關係，描述如何建立與維繫' },
        channels:              { type: Type.STRING, description: '通路，列出主要通路方式，每個用\\n分隔' },
        customerSegments:      { type: Type.STRING, description: '客戶區隔，描述目標客群細分' },
        costStructure:         { type: Type.STRING, description: '成本結構，列出主要成本項目，每個用\\n分隔' },
        revenueStreams:         { type: Type.STRING, description: '收益來源，列出主要收益模式，每個用\\n分隔' },
      },
      required: ['keyPartners','keyActivities','keyResources','valuePropositions','customerRelationships','channels','customerSegments','costStructure','revenueStreams'],
    },
    swot: {
      type: Type.OBJECT,
      properties: {
        strengths:     { type: Type.ARRAY, items: { type: Type.STRING }, description: '3-5個內部優勢' },
        weaknesses:    { type: Type.ARRAY, items: { type: Type.STRING }, description: '3-5個內部劣勢' },
        opportunities: { type: Type.ARRAY, items: { type: Type.STRING }, description: '3-5個外部機會' },
        threats:       { type: Type.ARRAY, items: { type: Type.STRING }, description: '3-5個外部威脅' },
      },
      required: ['strengths','weaknesses','opportunities','threats'],
    },
    marketSize: {
      type: Type.OBJECT,
      properties: {
        tam: { type: Type.STRING, description: '總體可服務市場 TAM，含數字與說明' },
        sam: { type: Type.STRING, description: '可服務的有效市場 SAM，含數字與說明' },
        som: { type: Type.STRING, description: '可獲得的市場份額 SOM，含數字與說明' },
        description: { type: Type.STRING, description: '市場機會整體說明，2-3句話' },
      },
      required: ['tam','sam','som','description'],
    },
    steep: {
      type: Type.OBJECT,
      properties: {
        social:        { type: Type.STRING, description: '社會因素分析，2-3個重點用\\n分隔' },
        technological: { type: Type.STRING, description: '技術因素分析，2-3個重點用\\n分隔' },
        economic:      { type: Type.STRING, description: '經濟因素分析，2-3個重點用\\n分隔' },
        environmental: { type: Type.STRING, description: '環境因素分析，2-3個重點用\\n分隔' },
        political:     { type: Type.STRING, description: '政治因素分析，2-3個重點用\\n分隔' },
      },
      required: ['social','technological','economic','environmental','political'],
    },
    fourC: {
      type: Type.OBJECT,
      properties: {
        consumer:      { type: Type.STRING, description: '消費者需求與洞察' },
        cost:          { type: Type.STRING, description: '消費者總成本（不只是價格）' },
        convenience:   { type: Type.STRING, description: '購買與使用便利性' },
        communication: { type: Type.STRING, description: '溝通與互動策略' },
      },
      required: ['consumer','cost','convenience','communication'],
    },
    portersFiveForces: {
      type: Type.OBJECT,
      properties: {
        competitiveRivalry:          { type: Type.STRING, description: '現有競爭者威脅，含強度評估（高/中/低）與分析' },
        threatOfNewEntrants:         { type: Type.STRING, description: '新進入者威脅，含強度評估與分析' },
        threatOfSubstitutes:         { type: Type.STRING, description: '替代品威脅，含強度評估與分析' },
        bargainingPowerOfBuyers:     { type: Type.STRING, description: '買方議價能力，含強度評估與分析' },
        bargainingPowerOfSuppliers:  { type: Type.STRING, description: '供應商議價能力，含強度評估與分析' },
        complementors:               { type: Type.STRING, description: '第六力：互補者分析，誰能讓你的產品更有價值' },
      },
      required: ['competitiveRivalry','threatOfNewEntrants','threatOfSubstitutes','bargainingPowerOfBuyers','bargainingPowerOfSuppliers','complementors'],
    },
    radarDimensions: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          dimension:          { type: Type.STRING, description: '競爭維度名稱，如：品牌力、技術、價格' },
          selfScore:          { type: Type.NUMBER, description: '自身評分 0-100' },
          competitorAvgScore: { type: Type.NUMBER, description: '競爭對手平均分 0-100' },
        },
        required: ['dimension','selfScore','competitorAvgScore'],
      },
      description: '5-7個競爭維度，用於雷達圖',
    },
    bcgMatrix: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          name:        { type: Type.STRING, description: '產品/服務名稱' },
          category:    { type: Type.STRING, enum: ['STAR','CASH_COW','QUESTION_MARK','DOG'] },
          description: { type: Type.STRING, description: '為何歸類於此，以及建議策略' },
        },
        required: ['name','category','description'],
      },
    },
    spanAnalysis: {
      type: Type.OBJECT,
      properties: {
        scope:       { type: Type.STRING, description: '業務範圍與邊界定義' },
        positioning: { type: Type.STRING, description: '市場定位策略' },
        advantage:   { type: Type.STRING, description: '核心競爭優勢來源' },
        network:     { type: Type.STRING, description: '網絡效應與生態系建構' },
      },
      required: ['scope','positioning','advantage','network'],
    },
    fanStrategy: {
      type: Type.STRING,
      description: 'FAN策略地圖：財務面、客戶面、內部流程面、學習與成長面，各面向2-3個重點用\\n---\\n分隔',
    },
    priorityMap: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          action:   { type: Type.STRING, description: '具體行動項目，最多25字' },
          impact:   { type: Type.STRING, enum: ['HIGH','MEDIUM','LOW'] },
          cost:     { type: Type.STRING, enum: ['HIGH','MEDIUM','LOW'] },
          priority: { type: Type.NUMBER, description: '優先級數字，1=最優先' },
        },
        required: ['action','impact','cost','priority'],
      },
      description: '6-10個行動項目，按優先級排列（低成本高影響優先）',
    },
    actionMap: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          month:         { type: Type.STRING, description: '時間區間，如：第1-2個月' },
          milestone:     { type: Type.STRING, description: '關鍵里程碑名稱，最多20字' },
          keyActions:    { type: Type.ARRAY, items: { type: Type.STRING }, description: '2-4個具體行動步驟' },
          successMetric: { type: Type.STRING, description: '成功衡量指標，具體可量化' },
        },
        required: ['month','milestone','keyActions','successMetric'],
      },
      description: '未來3-6個月的行動計畫，4-6個里程碑',
    },
    keyRecommendation: {
      type: Type.STRING,
      description: '最終關鍵建議，3-5個重點用\\n分隔，整合所有分析給出最重要的行動指引',
    },
    continueToIterate: {
      type: Type.STRING,
      description: '持續迭代建議，用4到6個重點，每點用換行符\\n分隔，每點為完整的分析句子，充分闡述下一步的具體行動與優化方向',
    },
  },
  required: [
    'businessModelCanvas','swot','marketSize','steep','fourC','portersFiveForces',
    'radarDimensions','bcgMatrix','spanAnalysis','fanStrategy',
    'priorityMap','actionMap','keyRecommendation','continueToIterate',
  ],
};

export const runPhase3DeepReport = async (
  data: Phase1DiagnosisResult['extractedData']
): Promise<Phase3DeepReport> => {
  const apiKey = getApiKey();
  const ai = new GoogleGenAI({ apiKey });

  const prompt = `你是 OmniView，一個頂尖的 AI 商業戰略分析系統。
請對以下商業提案進行全維度深度分析，產出專業的「Master Plan」報告。

每個分析框架都必須有實質內容，不能泛泛而談，要給出具體數字、名稱、時間點。

分析框架包含：
1. 商業模式畫布 (Business Model Canvas) - 完整 9 個要素
2. SWOT 分析 - 基於提案的真實優劣勢機威
3. 市場規模 TAM/SAM/SOM - 給出估計數字
4. STEEP 分析 - 宏觀環境五個面向
5. 4C 分析 - 以消費者為中心的市場框架
6. Porter's 五力 + 第六力（互補者）
7. 競爭定位雷達圖維度（5-7個維度，給出自身與競爭對手評分）
8. BCG 矩陣 - 產品/服務組合分析
9. SPAN 戰略定位分析
10. FAN 策略地圖（財務、客戶、內部流程、學習成長）
11. 優先級地圖 - 6-10個行動，標明影響力與成本
12. 行動地圖 - 未來3-6個月關鍵里程碑與成功指標
13. 持續迭代建議 - 針對下一步優化與軸轉的具體建議

所有輸出必須使用繁體中文。

商業提案資料：
- 核心想法：${data.idea}
- 市場資料：${data.marketData}
- 產品細節：${data.productDetails}
- 市場痛點：${data.painPoints}
- 目標消費者：${data.targetConsumer}
- 財務背景：${data.financialContext}`;

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      config: { responseMimeType: 'application/json', responseSchema: phase3Schema },
    });
    return JSON.parse(response.text!) as Phase3DeepReport;
  } catch (error) {
    handleApiError(error);
  }
};
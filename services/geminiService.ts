import { GoogleGenAI, Type, Schema, PersonGeneration } from "@google/genai";
import { AnalysisMode, BusinessInput, AnalysisResult, ParsedInputResponse } from "../types";

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
          icon: { type: Type.STRING },
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
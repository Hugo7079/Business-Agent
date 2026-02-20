import { GoogleGenAI, Type, Schema } from "@google/genai";
import { AnalysisMode, BusinessInput, AnalysisResult, ParsedInputResponse } from "../types";

const analysisSchema: Schema = {
  type: Type.OBJECT,
  properties: {
    successProbability: { type: Type.NUMBER, description: '0到100之間的整數，例如 65 代表 65%，絕對不可以用小數如 0.65' },
    executiveSummary: { type: Type.STRING },
    marketAnalysis: {
      type: Type.OBJECT,
      properties: {
        size: { type: Type.STRING },
        growthRate: { type: Type.STRING },
        description: { type: Type.STRING },
      },
      required: ['size', 'growthRate', 'description'],
    },
    competitors: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          name: { type: Type.STRING },
          strength: { type: Type.STRING },
          weakness: { type: Type.STRING },
        },
        required: ['name', 'strength', 'weakness'],
      },
    },
    roadmap: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          phase: { type: Type.STRING },
          timeframe: { type: Type.STRING },
          technology: { type: Type.STRING },
          product: { type: Type.STRING },
        },
        required: ['phase', 'timeframe', 'technology', 'product'],
      },
    },
    financials: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          year: { type: Type.STRING },
          revenue: { type: Type.NUMBER },
          profit: { type: Type.NUMBER },
          costs: { type: Type.NUMBER },
        },
        required: ['year', 'revenue', 'profit', 'costs'],
      },
    },
    breakEvenPoint: { type: Type.STRING },
    risks: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          risk: { type: Type.STRING },
          impact: { type: Type.STRING, enum: ["High", "Medium", "Low"] },
          mitigation: { type: Type.STRING },
        },
        required: ['risk', 'impact', 'mitigation'],
      },
    },
    personaEvaluations: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          role: { type: Type.STRING },
          icon: { type: Type.STRING },
          perspective: { type: Type.STRING, description: '此角色對提案的詳細看法，至少 50 字' },
          score: { type: Type.NUMBER, description: '0到100之間的整數評分' },
          keyQuote: { type: Type.STRING, description: '此角色最具代表性的一句話引言' },
          concern: { type: Type.STRING, description: '此角色最主要的擔憂，必填，至少 20 字' },
        },
        required: ['role', 'icon', 'perspective', 'score', 'keyQuote', 'concern'],
      },
    },
    teamAnalysis: { type: Type.STRING },
    finalVerdicts: {
      type: Type.OBJECT,
      properties: {
        aggressive: { type: Type.STRING },
        balanced: { type: Type.STRING },
        conservative: { type: Type.STRING },
      },
      required: ['aggressive', 'balanced', 'conservative'],
    },
  },
  required: [
    "successProbability", "executiveSummary", "marketAnalysis", "competitors",
    "roadmap", "financials", "breakEvenPoint", "risks",
    "personaEvaluations", "teamAnalysis", "finalVerdicts"
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
  if (
    msg.includes('429') ||
    msg.includes('quota') ||
    msg.includes('RESOURCE_EXHAUSTED') ||
    msg.includes('rate limit') ||
    msg.includes('Too Many Requests')
  ) {
    throw new Error(
      '⚠️ API 配額已用盡\n\n您的 Gemini API Key 已達到免費方案的用量上限。\n\n解決方式：\n• 等待配額重置（通常每分鐘或每天重置）\n• 前往 Google AI Studio 升級至付費方案\n• 使用不同的 API Key'
    );
  }
  if (msg.includes('API_KEY_INVALID') || msg.includes('401')) {
    throw new Error('❌ API Key 無效\n\n請點擊右上角設定按鈕，重新輸入正確的 Gemini API Key。');
  }
  if (msg.includes('not found') || msg.includes('NOT_FOUND') || msg.includes('404')) {
    throw new Error('❌ AI 模型暫時無法使用\n\n目前指定的 Gemini 模型不可用，可能是 Google 更新了模型版本。\n請稍後再試，或聯繫開發者更新模型設定。');
  }
  throw error;
};

export const generateBusinessImage = async (prompt: string): Promise<string | null> => {
  const apiKey = getApiKey();
  const ai = new GoogleGenAI({ apiKey });

  try {
    const response = await ai.models.generateContent({
      model: 'imagen-3.0-generate-002',
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      config: { 
        responseMimeType: 'image/jpeg' 
      }
    });
    
    // Check if we got an image back. The structure depends on the specific client version responses
    // For @google/genai with Imagen, it typically returns the binary data or base64.
    // However, the standard generateContent might return text if the model isn't set up for image gen properly in this SDK version
    // or if the response format differs.
    // Let's assume standard handling: usually for image models it returns inline data.
    
    // Note: If the specific SDK version/model combo differs, we might need a fallback.
    // But let's try the standard pattern for the new SDK.
    const imagePart = response.candidates?.[0]?.content?.parts?.[0];
    
    if (imagePart && 'inlineData' in imagePart) {
      return `data:${imagePart.inlineData.mimeType};base64,${imagePart.inlineData.data}`;
    }
    
    // Fallback: Check if payload is directly in a different structure or try a REST call if SDK fails.
    // But for now let's hope the SDK works. If it returns null, the UI just won't show an image.
    return null; 
  } catch (e) {
    console.warn("Image generation failed:", e);
    return null;
  }
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
      model: 'gemini-2.0-flash',
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

  // Generate an image prompt based on the idea
  let base64Image: string | null = null;
  try {
     const imagePrompt = `High quality, photorealistic commercial concept art for a business proposal about: ${data.idea}. 
     Professional, cinematic lighting, corporate futuristic style, no text, 8k resolution.`;
     base64Image = await generateBusinessImage(imagePrompt);
  } catch (e) {
    console.warn("Failed to generate intro image", e);
  }

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
- 目標客群：${data.targetConsumer}
- 財務背景：${data.financialContext}`;

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-pro',
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      config: {
        responseMimeType: 'application/json',
        responseSchema: analysisSchema,
        thinkingConfig: { thinkingBudget: 4096 },
      },
    });
    return JSON.parse(response.text!) as AnalysisResult;
  } catch (error) {
    handleApiError(error);
  }
};

export const generatePptNotebook = async (result: AnalysisResult): Promise<string> => {
  const apiKey = getApiKey();
  const ai = new GoogleGenAI({ apiKey });
  const dataJson = JSON.stringify(result, null, 2);

  const prompt = `
You are a world-class Python developer and presentation designer. 
Generate a COMPLETE, RUNNABLE Python script that creates a stunning investor-grade PowerPoint (.pptx).

The script will be run in a Jupyter Notebook environment. ALL text content MUST be in Traditional Chinese (繁體中文).

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ARCHITECTURE — THREE TYPES OF VISUALS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

TYPE A — AI-GENERATED IMAGES (use google-generativeai with imagen-3.0-generate-002):
  Use this for: slide backgrounds, hero illustrations, conceptual images.
  Method:
    import google.generativeai as genai
    from google.generativeai.types import GenerateImagesConfig
    import base64, io
    client = genai.Client(api_key=os.environ.get('GEMINI_API_KEY',''))
    response = client.models.generate_images(
        model='imagen-3.0-generate-002',
        prompt='...',
        config=GenerateImagesConfig(number_of_images=1, aspect_ratio="16:9", output_mime_type='image/png')
    )
    img_bytes = base64.b64decode(response.generated_images[0].image.image_bytes)
  Wrap every Imagen call in try/except — if it fails, fall back to a matplotlib gradient background.

TYPE B — MATPLOTLIB CHARTS (only for data/numbers):
  Financial bar charts, risk scatter, persona radar, roadmap gantt.
  Dark bg (#1e293b), white text, dpi=150, save to BytesIO.
  plt.rcParams['font.family'] = ['Arial Unicode MS', 'PingFang TC', 'Heiti TC', 'sans-serif']

TYPE C — PYTHON-PPTX TABLES (only for structured comparison):
  Competitive landscape, roadmap milestones.
  Dark header (#1e3a5f), alternating rows (#1e293b / #0f172a), white text.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SLIDE STRUCTURE (10 slides):
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SLIDE 1 — COVER: TYPE A bg (dark space particles) + TYPE A hero (industry-matched illustration) + title/tagline/score overlay
SLIDE 2 — EXECUTIVE SUMMARY: TYPE A bg (dark blue waves) + bullet points with ◆ + semi-transparent text card
SLIDE 3 — MARKET OPPORTUNITY: TYPE A bg (night cityscape) + TYPE B horizontal bar chart + KPI numbers overlay
SLIDE 4 — FINANCIAL PROJECTIONS: solid #0f172a + TYPE B grouped bar chart + TYPE C mini data table
SLIDE 5 — COMPETITIVE LANDSCAPE: TYPE A bg (dark chess strategy) + TYPE C full-width styled table
SLIDE 6 — STRATEGIC ROADMAP: TYPE A bg (glowing timeline nodes) + TYPE B Gantt chart
SLIDE 7 — RISK MATRIX: solid #0f172a + TYPE B 2x2 scatter quadrant with labels
SLIDE 8 — STAKEHOLDER PERSPECTIVES: TYPE A bg (network silhouettes) + TYPE B radar/spider chart
SLIDE 9 — FINAL VERDICT: TYPE A bg (dramatic light rays) + 3-column text (Aggressive/Balanced/Conservative)
SLIDE 10 — CALL TO ACTION: TYPE A hero image (industry-specific inspiring) + centered action text

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RULES:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. API key: os.environ.get('GEMINI_API_KEY', '')
2. Always add semi-transparent dark rect behind text placed over images
3. Slide size: 13.33 x 7.5 inches (widescreen 16:9)
4. Font: Title=44pt, Subtitle=26pt, Body=18pt, Caption=13pt, all white
5. Save as: OmniView_投資提案.pptx
6. Print progress for each slide and final success message

Analysis data:
\`\`\`json
${dataJson}
\`\`\`

Return ONLY raw Python code. No markdown fences. No explanation. Start with import statements.
`;

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-pro',
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      config: { thinkingConfig: { thinkingBudget: 16000 } },
    });
    let code = response.text?.trim() || '';
    code = code.replace(/^```python\n?/m, '').replace(/^```\n?/m, '').replace(/\n?```$/m, '');
    return code;
  } catch (error) {
    handleApiError(error);
  }
};
import { GoogleGenAI, Type, Schema } from "@google/genai";
import { AnalysisMode, GeneralInput, CorporateInput, AnalysisResult, ParsedInputResponse } from "../types";

// Define the response schema strictly to ensure UI can render it.
const analysisSchema: Schema = {
  type: Type.OBJECT,
  properties: {
    successProbability: { type: Type.NUMBER, description: "Overall success probability percentage (0-100)" },
    executiveSummary: { type: Type.STRING, description: "A high-level summary of the entire analysis." },
    marketAnalysis: {
      type: Type.OBJECT,
      properties: {
        size: { type: Type.STRING },
        growthRate: { type: Type.STRING },
        description: { type: Type.STRING },
      },
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
      },
    },
    roadmap: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          phase: { type: Type.STRING, description: "Phase 1, Phase 2, etc." },
          timeframe: { type: Type.STRING, description: "e.g., Q1 2024" },
          technology: { type: Type.STRING, description: "Tech milestones" },
          product: { type: Type.STRING, description: "Product milestones" },
        },
      },
    },
    financials: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          year: { type: Type.STRING, description: "Year 1, Year 2..." },
          revenue: { type: Type.NUMBER },
          profit: { type: Type.NUMBER },
          costs: { type: Type.NUMBER },
        },
      },
    },
    breakEvenPoint: { type: Type.STRING, description: "Estimated time to break even." },
    risks: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          risk: { type: Type.STRING },
          impact: { type: Type.STRING, enum: ["High", "Medium", "Low"] },
          mitigation: { type: Type.STRING },
        },
      },
    },
    personaEvaluations: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          role: { type: Type.STRING, description: "e.g., Investor, Competitor, Customer" },
          icon: { type: Type.STRING, description: "Just a keyword: 'money', 'user', 'shield', 'briefcase', 'hammer'" },
          perspective: { type: Type.STRING, description: "Their specific point of view." },
          score: { type: Type.NUMBER, description: "Score 0-100 from their perspective." },
          keyQuote: { type: Type.STRING, description: "A fictional quote representing their stance." },
          concern: { type: Type.STRING, description: "Their biggest worry." },
        },
      },
    },
    teamAnalysis: { type: Type.STRING, description: "Analysis of required team capabilities." },
    finalVerdicts: {
      type: Type.OBJECT,
      properties: {
        aggressive: { type: Type.STRING, description: "Opinion from an aggressive growth perspective." },
        balanced: { type: Type.STRING, description: "Opinion from a balanced/realistic perspective." },
        conservative: { type: Type.STRING, description: "Opinion from a risk-averse perspective." },
      },
    },
  },
  required: [
    "successProbability",
    "executiveSummary",
    "marketAnalysis",
    "competitors",
    "roadmap",
    "financials",
    "breakEvenPoint",
    "risks",
    "personaEvaluations",
    "teamAnalysis",
    "finalVerdicts"
  ],
};

const startupParserSchema: Schema = {
  type: Type.OBJECT,
  properties: {
    data: {
      type: Type.OBJECT,
      properties: {
        marketData: { type: Type.STRING },
        productDetails: { type: Type.STRING },
        literature: { type: Type.STRING },
        painPoints: { type: Type.STRING },
        techInnovation: { type: Type.STRING },
      },
      required: ["marketData", "productDetails", "literature", "painPoints", "techInnovation"]
    },
    feedback: { type: Type.STRING, description: "Explanation of what was extracted, inferred, or is missing. Output in Traditional Chinese." }
  },
  required: ["data", "feedback"]
};

const corporateParserSchema: Schema = {
  type: Type.OBJECT,
  properties: {
    data: {
      type: Type.OBJECT,
      properties: {
        brandTraits: { type: Type.STRING },
        targetConsumer: { type: Type.STRING },
        channels: { type: Type.STRING },
        proposalGoal: { type: Type.STRING },
        financialReport: { type: Type.STRING },
      },
      required: ["brandTraits", "targetConsumer", "channels", "proposalGoal", "financialReport"]
    },
    feedback: { type: Type.STRING, description: "Explanation of what was extracted, inferred, or is missing. Output in Traditional Chinese." }
  },
  required: ["data", "feedback"]
};

export const parseBusinessIdea = async (
  mode: AnalysisMode,
  textInput: string,
  audioBase64?: string
): Promise<ParsedInputResponse> => {
  const apiKey = localStorage.getItem('GEMINI_API_KEY') || process.env.API_KEY;
  
  if (!apiKey) {
    throw new Error("找不到 API Key。請點擊右上角設定按鈕輸入您的 Google Gemini API Key。");
  }
  const ai = new GoogleGenAI({ apiKey });

  const isStartup = mode === AnalysisMode.STARTUP;
  const targetSchema = isStartup ? startupParserSchema : corporateParserSchema;
  const fields = isStartup 
    ? "市場資料 (Market Data), 產品細節 (Product Details), 文獻/參考資料 (Literature/References), 市場痛點 (Pain Points), 技術創新點 (Tech Innovation)"
    : "品牌特性 (Brand Traits), 目標消費者 (Target Consumer), 主要通路 (Channels), 提案目的 (Proposal Goal), 財務報告 (Financial Context)";

  const prompt = `
    You are an intelligent business analyst assistant. 
    The user will provide a rough description of a business idea or proposal (via text or audio).
    
    Your task is to:
    1. Organize this information into the structured fields required for a Business Model Canvas analysis: ${fields}.
    2. If the user does not provide specific details for a field, you must **INFER reasonable assumptions** based on the context to ensure the form is filled. Do not leave fields empty if possible.
    3. In the 'feedback' field, clearly summarize what you extracted and specifically mention what you had to infer or assume so the user knows.
    
    **IMPORTANT**: The 'feedback' and all extracted string fields MUST be in Traditional Chinese (繁體中文).
    
    Mode: ${isStartup ? "New Venture / Startup" : "Corporate Brand Proposal"}
  `;

  const contents = [];
  
  if (audioBase64) {
    contents.push({
      inlineData: {
        mimeType: "audio/webm", // Common for browser MediaRecorder
        data: audioBase64
      }
    });
  }
  
  contents.push({
    text: prompt + `\n\nUser Input: ${textInput || "(Processing Audio Input)"}`
  });

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-3-flash-preview', 
      contents: [{ role: 'user', parts: contents }],
      config: {
        responseMimeType: "application/json",
        responseSchema: targetSchema,
      },
    });

    const text = response.text;
    if (!text) throw new Error("Failed to parse input");
    
    return JSON.parse(text) as ParsedInputResponse;
  } catch (error) {
    console.error("Input Parsing Error:", error);
    throw error;
  }
};

export const analyzeBusiness = async (
  mode: AnalysisMode,
  data: GeneralInput | CorporateInput
): Promise<AnalysisResult> => {
  const apiKey = localStorage.getItem('GEMINI_API_KEY') || process.env.API_KEY;
  if (!apiKey) {
    throw new Error("找不到 API Key。請點擊右上角設定按鈕輸入您的 Google Gemini API Key。");
  }

  const ai = new GoogleGenAI({ apiKey });

  let systemPrompt = `
    You are OmniView, an advanced AI Business Simulator. 
    Your goal is to simulate a "Board of Directors" meeting composed of different AI personas (Investor, Competitor, Supplier, Employee, Consumer) to strictly evaluate a business opportunity.
    
    You must provide a 360-degree assessment covering:
    1. Market Size & Growth
    2. Competitor Analysis
    3. Tech & Product Roadmap
    4. Team Capabilities
    5. Resource & Time Allocation
    6. Financial Projections (P&L) & ROI
    7. Risk Factors
    8. Three distinct final verdicts: Aggressive, Balanced, Conservative.

    Be critical, realistic, and data-driven where possible based on your general knowledge.
    Avoid vague platitudes. Give specific numbers for financials (estimates) and concrete risks.

    **IMPORTANT**: ALL OUTPUT MUST BE IN TRADITIONAL CHINESE (繁體中文).
  `;

  let userPrompt = "";

  if (mode === AnalysisMode.STARTUP) {
    const input = data as GeneralInput;
    userPrompt = `
      請評估這個新創想法 (EVALUATE THIS STARTUP IDEA):
      - 市場資料 (Market Data Context): ${input.marketData}
      - 產品細節 (Product Details): ${input.productDetails}
      - 相關文獻/資料 (Related Literature/Data): ${input.literature}
      - 市場痛點 (Market Pain Points): ${input.painPoints}
      - 技術創新點 (Tech Innovation): ${input.techInnovation}
    `;
  } else {
    const input = data as CorporateInput;
    userPrompt = `
      請評估這個企業提案 (EVALUATE THIS CORPORATE PROPOSAL):
      - 品牌特性 (Brand Traits): ${input.brandTraits}
      - 目標消費者 (Target Consumer): ${input.targetConsumer}
      - 主要通路 (Main Channels): ${input.channels}
      - 提案目的 (Proposal Goal): ${input.proposalGoal}
      - 財務報告 (Financial Context): ${input.financialReport}
    `;
  }

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-3-pro-preview', // Using Pro for complex reasoning and role simulation
      contents: [
        { role: 'user', parts: [{ text: systemPrompt + "\n\n" + userPrompt }] }
      ],
      config: {
        responseMimeType: "application/json",
        responseSchema: analysisSchema,
        thinkingConfig: { thinkingBudget: 4096 } // Allow detailed thinking for financial models
      },
    });

    const text = response.text;
    if (!text) throw new Error("No response generated");
    
    return JSON.parse(text) as AnalysisResult;
  } catch (error) {
    console.error("Gemini Analysis Error:", error);
    throw error;
  }
};

export const generatePptNotebook = async (result: AnalysisResult): Promise<string> => {
  const apiKey = localStorage.getItem('GEMINI_API_KEY') || process.env.API_KEY;
  if (!apiKey) throw new Error("找不到 API Key。");
  const ai = new GoogleGenAI({ apiKey });

  const dataJson = JSON.stringify(result, null, 2);

  const prompt = `
You are a world-class Python developer and presentation designer. 
Generate a COMPLETE, RUNNABLE Python script that creates a stunning investor-grade PowerPoint (.pptx).

The script will be run in a Jupyter Notebook environment. ALL text content MUST be in Traditional Chinese (繁體中文).

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ARCHITECTURE — TWO TYPES OF VISUALS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

TYPE A — AI-GENERATED IMAGES (use google-generativeai with imagen-3.0-generate-002):
  Use this for: slide backgrounds, hero illustrations, conceptual images.
  Method:
    import google.generativeai as genai
    import base64, io
    genai.configure(api_key=YOUR_API_KEY)
    client = genai.Client()
    response = client.models.generate_images(
        model='imagen-3.0-generate-002',
        prompt='...',
        config=GenerateImagesConfig(number_of_images=1, aspect_ratio="16:9", output_mime_type='image/png')
    )
    img_bytes = base64.b64decode(response.generated_images[0].image.image_bytes)
    # Then add to slide as background or picture

TYPE B — MATPLOTLIB CHARTS (use only for data/numbers):
  Use this for: financial bar charts, risk scatter plots, persona radar charts, roadmap gantt.
  Always use: dark bg (#1e293b), white text, dpi=150, BytesIO to embed.
  Font: plt.rcParams['font.family'] = ['Arial Unicode MS', 'PingFang TC', 'Heiti TC', 'sans-serif']

TYPE C — PYTHON-PPTX TABLES (use only for structured comparison data):
  Use this for: competitive landscape table, roadmap milestone table.
  Style with dark header (#1e3a5f), alternating rows (#1e293b / #0f172a), white text.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SLIDE STRUCTURE (10 slides):
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

SLIDE 1 — COVER
  - TYPE A background: "cinematic dark navy abstract technology background, deep space particles, subtle blue glow, 16:9, ultra HD, no text"
  - TYPE A hero image: A photorealistic illustration matching the business domain (infer from the analysis data what industry it is)
  - Large title text (white, bold), tagline, success probability as a giant colored number overlay

SLIDE 2 — EXECUTIVE SUMMARY  
  - TYPE A background: "dark abstract gradient background, deep blue and indigo tones, smooth waves, 16:9, professional"
  - Key bullet points as large readable text with ◆ prefix
  - Semi-transparent dark overlay card for text readability

SLIDE 3 — MARKET OPPORTUNITY
  - TYPE A background: "dark market cityscape aerial view at night, blue lights, 16:9, cinematic"
  - TYPE B chart: Horizontal bar chart showing market size breakdown with gradient bars
  - KPI numbers (TAM, CAGR) as large overlay text

SLIDE 4 — FINANCIAL PROJECTIONS
  - Solid dark background (#0f172a) — no AI image needed, chart is the hero
  - TYPE B chart: Grouped bar chart (Revenue/Profit/Cost by year), full width, high quality
  - TYPE C mini table: Year / Revenue / Profit / Cost below the chart

SLIDE 5 — COMPETITIVE LANDSCAPE
  - TYPE A background: "abstract dark chess board strategy game concept, blue tones, 16:9"
  - TYPE C table: Competitor / Strength / Weakness — styled, full width

SLIDE 6 — STRATEGIC ROADMAP
  - TYPE A background: "dark futuristic timeline path glowing blue nodes connected, 16:9"
  - TYPE B chart: Horizontal Gantt-style chart using matplotlib barh with phase labels

SLIDE 7 — RISK MATRIX
  - Solid dark background (#0f172a)
  - TYPE B chart: 2×2 quadrant scatter (Impact vs Likelihood), colored by severity, with risk labels

SLIDE 8 — STAKEHOLDER PERSPECTIVES
  - TYPE A background: "dark abstract network of connected people silhouettes blue glow 16:9"
  - TYPE B chart: Spider/radar chart — each persona plotted as a polygon

SLIDE 9 — FINAL VERDICT
  - TYPE A background: "dark abstract dramatic light rays from above, navy blue, 16:9"
  - 3 text boxes side by side: Aggressive (orange) / Balanced (blue) / Conservative (purple), each with colored border and icon

SLIDE 10 — CALL TO ACTION
  - TYPE A background: Generate a bold, inspiring image relevant to this specific business domain
  - Large white centered text: "立即採取行動" with sub-points as next steps
  - Contact/branding footer

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
IMPORTANT RULES:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. The API key is read from environment variable: os.environ.get('GEMINI_API_KEY', '')
2. For Imagen calls, wrap in try/except — if image generation fails, fall back to a matplotlib-generated gradient background
3. All text on slides MUST be readable: always add a semi-transparent dark rectangle (RGBColor) behind text if placing over image
4. Slide size: 13.33 x 7.5 inches (widescreen 16:9)
5. Font sizes: Title=40-48pt, Subtitle=24-28pt, Body=16-20pt, Caption=12-14pt
6. Save as: OmniView_投資提案.pptx
7. At the end, print total slides generated and file path

Here is the analysis data:
\`\`\`json
${dataJson}
\`\`\`

Return ONLY the complete raw Python code. No markdown fences. No explanation. Start directly with import statements.
`;

  const response = await ai.models.generateContent({
    model: 'gemini-2.5-pro-preview-05-06',
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    config: { thinkingConfig: { thinkingBudget: 16000 } }
  });

  let code = response.text?.trim() || '';
  // Strip any accidental markdown fences
  code = code.replace(/^```python\n?/m, '').replace(/^```\n?/m, '').replace(/\n?```$/m, '');
  return code;
};
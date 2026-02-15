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
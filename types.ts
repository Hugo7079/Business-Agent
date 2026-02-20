export enum AnalysisMode {
  BUSINESS = 'BUSINESS',
}

export interface BusinessInput {
  idea: string;
  marketData: string;
  productDetails: string;
  painPoints: string;
  targetConsumer: string;
  financialContext: string;
}

export interface ParsedInputResponse {
  data: BusinessInput;
  feedback: string;
}

export interface PersonaEvaluation {
  role: string;
  icon: string; // Icon name reference
  perspective: string;
  score: number; // 0-100
  keyQuote: string;
  concern: string;
}

export interface FinancialYear {
  year: string;
  revenue: number;
  profit: number;
  costs: number;
}

export interface RoadmapItem {
  phase: string;
  timeframe: string;
  technology: string;
  product: string;
}

export interface AnalysisResult {
  successProbability: number;
  executiveSummary: string;
  marketAnalysis: {
    size: string;
    growthRate: string;
    description: string;
  };
  competitors: Array<{
    name: string;
    strength: string;
    weakness: string;
  }>;
  roadmap: RoadmapItem[];
  financials: FinancialYear[];
  breakEvenPoint: string;
  risks: Array<{
    risk: string;
    impact: 'High' | 'Medium' | 'Low';
    mitigation: string;
  }>;
  personaEvaluations: PersonaEvaluation[];
  teamAnalysis: string;
  finalVerdicts: {
    aggressive: string;
    balanced: string;
    conservative: string;
  };
  introImage?: string; // Base64 image
}

export interface HistoryRecord {
  id: string;
  createdAt: number; // timestamp
  title: string;     // 從 idea 欄位取前 40 字
  input: BusinessInput;
  result: AnalysisResult;
}

export type AppState = 'IDLE' | 'ANALYZING' | 'COMPLETE' | 'ERROR';
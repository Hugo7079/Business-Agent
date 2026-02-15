export enum AnalysisMode {
  STARTUP = 'STARTUP',
  CORPORATE = 'CORPORATE',
}

export interface GeneralInput {
  marketData: string;
  productDetails: string;
  literature: string; // References/Data
  painPoints: string;
  techInnovation: string;
}

export interface CorporateInput {
  brandTraits: string;
  targetConsumer: string;
  channels: string;
  proposalGoal: string;
  financialReport: string;
}

export interface ParsedInputResponse {
  data: GeneralInput | CorporateInput;
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
}

export type AppState = 'IDLE' | 'ANALYZING' | 'COMPLETE' | 'ERROR';
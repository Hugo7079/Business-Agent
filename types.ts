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
  icon: string;
  perspective: string;
  score: number;
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
  continueToIterate: string;
}

export interface HistoryRecord {
  id: string;
  createdAt: number;
  title: string;
  input: BusinessInput;
  result: AnalysisResult;
  phase3Report?: Phase3DeepReport;
  finalSuccessProbability?: number;
}

/** App 整體流程狀態
 *  IDLE       → 首頁輸入
 *  SCANNING   → 呼叫 initialScan，loading
 *  CHATTING   → 初步掃描完畢，顯示輕量結果 + 互動補充對話
 *  ANALYZING  → 使用者確認完成，呼叫完整分析 + Phase3，loading
 *  COMPLETE   → 完整報告顯示
 *  ERROR      → 任何步驟失敗
 */
export type AppState = 'IDLE' | 'SCANNING' | 'CHATTING' | 'ANALYZING' | 'COMPLETE' | 'ERROR';

/** 初步掃描結果（輕量版，使用者第一次輸入後立刻看到） */
export interface InitialScanResult {
  successProbability: number;
  coreUnderstanding: string;
  marketSnapshot: {
    size: string;
    growthRate: string;
    insight: string;
  };
  competitors: Array<{
    name: string;
    threat: string;
  }>;
  infoGaps: Array<{
    field: string;
    question: string;
    priority: 'HIGH' | 'MEDIUM' | 'LOW';
  }>;
  extractedData: BusinessInput;
}

/** Phase 2 單輪對話訊息 */
export interface Phase2Message {
  role: 'assistant' | 'user';
  content: string;
  timestamp: number;
}

/** Phase 2 AI 回應結構 */
export interface Phase2Response {
  dynamicFeedback: string;
  updatedData: BusinessInput;
  completionRate: number;
  nextQuestions: string[];
  isReady: boolean;
}

// === Phase 3 深度分析報告型別 ===

export interface BusinessModelCanvas {
  keyPartners: string;
  keyActivities: string;
  keyResources: string;
  valuePropositions: string;
  customerRelationships: string;
  channels: string;
  customerSegments: string;
  costStructure: string;
  revenueStreams: string;
}

export interface SwotAnalysis {
  strengths: string[];
  weaknesses: string[];
  opportunities: string[];
  threats: string[];
}

export interface MarketSizeAnalysis {
  tam: string;
  sam: string;
  som: string;
  description: string;
}

export interface SteepAnalysis {
  social: string;
  technological: string;
  economic: string;
  environmental: string;
  political: string;
}

export interface FourCAnalysis {
  consumer: string;
  cost: string;
  convenience: string;
  communication: string;
}

export interface PortersFiveForces {
  competitiveRivalry: string;
  threatOfNewEntrants: string;
  threatOfSubstitutes: string;
  bargainingPowerOfBuyers: string;
  bargainingPowerOfSuppliers: string;
  complementors: string;
}

export interface RadarDimension {
  dimension: string;
  selfScore: number;
  competitorAvgScore: number;
}

export interface BcgItem {
  name: string;
  category: 'STAR' | 'CASH_COW' | 'QUESTION_MARK' | 'DOG';
  description: string;
}

export interface SpanAnalysis {
  scope: string;
  positioning: string;
  advantage: string;
  network: string;
}

export interface PriorityMapItem {
  action: string;
  impact: 'HIGH' | 'MEDIUM' | 'LOW';
  cost: 'HIGH' | 'MEDIUM' | 'LOW';
  priority: number;
}

export interface ActionMilestone {
  month: string;
  milestone: string;
  keyActions: string[];
  successMetric: string;
}

/** Phase 3 完整深度報告 */
export interface Phase3DeepReport {
  businessModelCanvas: BusinessModelCanvas;
  swot: SwotAnalysis;
  marketSize: MarketSizeAnalysis;
  steep: SteepAnalysis;
  fourC: FourCAnalysis;
  portersFiveForces: PortersFiveForces;
  radarDimensions: RadarDimension[];
  bcgMatrix: BcgItem[];
  spanAnalysis: SpanAnalysis;
  fanStrategy: string;
  priorityMap: PriorityMapItem[];
  actionMap: ActionMilestone[];
  keyRecommendation: string;
  continueToIterate: string;
}

// 保留舊型別供 DeepScanFlow 內部使用
export interface Phase1DiagnosisResult {
  coreInsight: string;
  successRateEstimate: number;
  riskWarnings: string[];
  infoGaps: Array<{
    field: string;
    question: string;
    priority: 'HIGH' | 'MEDIUM' | 'LOW';
  }>;
  extractedData: BusinessInput;
}

export interface Phase2State {
  messages: Phase2Message[];
  currentQuestions: string[];
  filledData: BusinessInput;
  completionRate: number;
  dynamicFeedback: string;
  isReady: boolean;
}

export type DeepScanPhase = 'INPUT' | 'PHASE1' | 'PHASE2' | 'PHASE3_LOADING' | 'PHASE3';

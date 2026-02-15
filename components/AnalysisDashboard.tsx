import React from 'react';
import { AnalysisResult, PersonaEvaluation, FinancialYear, RoadmapItem } from '../types';
import { 
  TrendingUp, Users, ShieldAlert, Target, 
  DollarSign, BarChart3, ChevronRight, CheckCircle2, 
  AlertTriangle, Briefcase, User, Factory, Store,
  Sword
} from 'lucide-react';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, LineChart, Line, AreaChart, Area
} from 'recharts';

interface Props {
  result: AnalysisResult;
  onReset: () => void;
}

const AnalysisDashboard: React.FC<Props> = ({ result, onReset }) => {
  return (
    <div className="w-full max-w-7xl mx-auto p-4 md:p-8 space-y-8 animate-in fade-in duration-700">
      
      {/* Header Section */}
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center border-b border-slate-700 pb-6">
        <div>
          <h1 className="text-3xl font-bold text-white mb-2">分析報告</h1>
          <p className="text-slate-400">全方位 360° 評估</p>
        </div>
        <div className="mt-4 md:mt-0 flex items-center space-x-4">
          <div className="text-right">
             <div className="text-sm text-slate-400 uppercase tracking-wider font-semibold">成功機率</div>
             <div className={`text-4xl font-black ${getScoreColor(result.successProbability)}`}>
               {result.successProbability}%
             </div>
          </div>
          <button 
            onClick={onReset}
            className="px-4 py-2 text-sm font-medium text-slate-300 hover:text-white bg-slate-800 hover:bg-slate-700 rounded-lg transition-colors border border-slate-700"
          >
            重新分析
          </button>
        </div>
      </div>

      {/* Executive Summary */}
      <div className="bg-slate-800/50 border border-slate-700 rounded-2xl p-6 shadow-sm">
        <h2 className="text-xl font-semibold text-white mb-3 flex items-center">
          <Target className="w-5 h-5 mr-2 text-blue-400" /> 執行摘要
        </h2>
        <p className="text-slate-300 leading-relaxed text-lg">
          {result.executiveSummary}
        </p>
      </div>

      {/* Key Metrics Grid */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        <div className="bg-slate-900 border border-slate-800 p-6 rounded-xl">
          <h3 className="text-sm font-medium text-slate-500 uppercase mb-4 flex items-center">
            <TrendingUp className="w-4 h-4 mr-2" /> 市場分析
          </h3>
          <div className="space-y-4">
            <div>
              <div className="text-2xl font-bold text-white">{result.marketAnalysis.size}</div>
              <div className="text-sm text-emerald-400">總體潛在市場 (TAM)</div>
            </div>
            <div>
              <div className="text-xl font-bold text-white">{result.marketAnalysis.growthRate}</div>
              <div className="text-sm text-blue-400">年均複合成長率 (CAGR)</div>
            </div>
            <p className="text-xs text-slate-400 mt-2 border-t border-slate-800 pt-2">
              {result.marketAnalysis.description}
            </p>
          </div>
        </div>

        <div className="bg-slate-900 border border-slate-800 p-6 rounded-xl">
          <h3 className="text-sm font-medium text-slate-500 uppercase mb-4 flex items-center">
            <ShieldAlert className="w-4 h-4 mr-2" /> 主要風險
          </h3>
          <ul className="space-y-3">
            {result.risks.slice(0, 3).map((risk, i) => (
              <li key={i} className="flex items-start">
                <AlertTriangle className={`w-4 h-4 mr-2 mt-0.5 flex-shrink-0 ${
                  risk.impact === 'High' ? 'text-red-500' : 'text-yellow-500'
                }`} />
                <div>
                  <div className="text-sm text-white font-medium">{risk.risk}</div>
                  <div className="text-xs text-slate-500 mt-0.5">{risk.mitigation}</div>
                </div>
              </li>
            ))}
          </ul>
        </div>

        <div className="bg-slate-900 border border-slate-800 p-6 rounded-xl">
          <h3 className="text-sm font-medium text-slate-500 uppercase mb-4 flex items-center">
            <DollarSign className="w-4 h-4 mr-2" /> 財務展望
          </h3>
          <div className="space-y-4">
             <div>
               <div className="text-lg font-bold text-white">損益平衡點</div>
               <div className="text-sm text-indigo-400">{result.breakEvenPoint}</div>
             </div>
             <div className="pt-2 border-t border-slate-800">
               <div className="text-sm text-slate-400 mb-1">預估第三年營收</div>
               <div className="text-xl font-bold text-white">
                 ${(result.financials[2]?.revenue || 0).toLocaleString()}
               </div>
             </div>
          </div>
        </div>
      </div>

      {/* Main Charts Section */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
        {/* Financial Chart */}
        <div className="bg-slate-800/40 border border-slate-700 p-6 rounded-2xl">
          <h3 className="text-lg font-semibold text-white mb-6 flex items-center">
            <BarChart3 className="w-5 h-5 mr-2 text-emerald-400" /> 財務預測
          </h3>
          <div className="h-[300px] w-full">
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={result.financials}>
                <defs>
                  <linearGradient id="colorRev" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#3b82f6" stopOpacity={0.8}/>
                    <stop offset="95%" stopColor="#3b82f6" stopOpacity={0}/>
                  </linearGradient>
                  <linearGradient id="colorProf" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#10b981" stopOpacity={0.8}/>
                    <stop offset="95%" stopColor="#10b981" stopOpacity={0}/>
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" stroke="#334155" />
                <XAxis dataKey="year" stroke="#94a3b8" />
                <YAxis stroke="#94a3b8" />
                <Tooltip 
                  contentStyle={{ backgroundColor: '#1e293b', borderColor: '#334155', color: '#f8fafc' }}
                  itemStyle={{ color: '#f8fafc' }}
                />
                <Legend />
                <Area type="monotone" dataKey="revenue" stroke="#3b82f6" fillOpacity={1} fill="url(#colorRev)" name="營收 (Revenue)" />
                <Area type="monotone" dataKey="profit" stroke="#10b981" fillOpacity={1} fill="url(#colorProf)" name="淨利 (Net Profit)" />
                <Area type="monotone" dataKey="costs" stroke="#ef4444" fill="transparent" strokeDasharray="5 5" name="成本 (Costs)" />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* Roadmap */}
        <div className="bg-slate-800/40 border border-slate-700 p-6 rounded-2xl flex flex-col">
          <h3 className="text-lg font-semibold text-white mb-6 flex items-center">
            <TrendingUp className="w-5 h-5 mr-2 text-indigo-400" /> 策略路線圖
          </h3>
          <div className="space-y-6 overflow-y-auto pr-2 custom-scrollbar flex-grow">
            {result.roadmap.map((item, idx) => (
              <div key={idx} className="relative pl-6 border-l-2 border-slate-700 last:border-0">
                <div className="absolute -left-[9px] top-0 w-4 h-4 rounded-full bg-indigo-500 border-4 border-slate-900"></div>
                <div className="mb-1 text-xs font-bold uppercase tracking-wider text-indigo-400">{item.timeframe} - {item.phase}</div>
                <div className="bg-slate-900 p-3 rounded-lg border border-slate-800">
                  <div className="flex flex-col md:flex-row gap-2 md:gap-4">
                    <div className="flex-1">
                      <span className="text-xs text-slate-500 block mb-0.5">產品 (Product)</span>
                      <span className="text-sm text-slate-200">{item.product}</span>
                    </div>
                    <div className="flex-1">
                      <span className="text-xs text-slate-500 block mb-0.5">技術 (Technology)</span>
                      <span className="text-sm text-slate-200">{item.technology}</span>
                    </div>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* AI Persona Board */}
      <div>
        <h3 className="text-2xl font-bold text-white mb-6 flex items-center">
          <Users className="w-6 h-6 mr-3 text-cyan-400" /> AI 虛擬董事會
        </h3>
        <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-6">
          {result.personaEvaluations.map((persona, idx) => (
             <PersonaCard key={idx} persona={persona} />
          ))}
        </div>
      </div>

      {/* Competitor Analysis */}
      <div className="bg-slate-800/30 border border-slate-700 rounded-2xl p-6">
        <h3 className="text-lg font-semibold text-white mb-4 flex items-center">
          <Sword className="w-5 h-5 mr-2 text-red-400" /> 競爭態勢分析
        </h3>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          {result.competitors.map((comp, idx) => (
            <div key={idx} className="bg-slate-900 p-4 rounded-xl border border-slate-800">
              <div className="font-bold text-white mb-2 text-lg">{comp.name}</div>
              <div className="space-y-2 text-sm">
                <div>
                  <span className="text-green-500 font-medium">優勢 (Strength):</span>
                  <p className="text-slate-400 leading-tight">{comp.strength}</p>
                </div>
                <div>
                  <span className="text-red-500 font-medium">劣勢 (Weakness):</span>
                  <p className="text-slate-400 leading-tight">{comp.weakness}</p>
                </div>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* Final Verdicts */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        <VerdictCard title="激進觀點" content={result.finalVerdicts.aggressive} type="aggressive" />
        <VerdictCard title="平衡觀點" content={result.finalVerdicts.balanced} type="balanced" />
        <VerdictCard title="保守觀點" content={result.finalVerdicts.conservative} type="conservative" />
      </div>

    </div>
  );
};

const getScoreColor = (score: number) => {
  if (score >= 80) return "text-emerald-400";
  if (score >= 60) return "text-blue-400";
  if (score >= 40) return "text-yellow-400";
  return "text-red-400";
};

// Sub-components

interface PersonaCardProps {
  persona: PersonaEvaluation;
}

const PersonaCard: React.FC<PersonaCardProps> = ({ persona }) => {
  const getIcon = (iconStr: string) => {
    switch (iconStr.toLowerCase()) {
      case 'money': case 'investor': return <DollarSign className="w-5 h-5 text-emerald-400" />;
      case 'user': case 'consumer': case 'customer': return <User className="w-5 h-5 text-blue-400" />;
      case 'shield': return <ShieldAlert className="w-5 h-5 text-orange-400" />;
      case 'briefcase': case 'employee': return <Briefcase className="w-5 h-5 text-purple-400" />;
      case 'factory': case 'supplier': return <Factory className="w-5 h-5 text-amber-400" />;
      case 'hammer': case 'competitor': return <Sword className="w-5 h-5 text-red-400" />;
      default: return <Users className="w-5 h-5 text-slate-400" />;
    }
  };

  return (
    <div className="bg-slate-800 border border-slate-700 rounded-xl p-5 hover:border-blue-500/50 transition-colors">
      <div className="flex justify-between items-start mb-4">
        <div className="flex items-center space-x-3">
          <div className="p-2 bg-slate-900 rounded-lg border border-slate-700">
            {getIcon(persona.icon || persona.role)}
          </div>
          <div>
            <div className="font-bold text-slate-200">{persona.role}</div>
            <div className="text-xs text-slate-500">董事會成員</div>
          </div>
        </div>
        <div className={`text-xl font-bold ${getScoreColor(persona.score)}`}>
          {persona.score}
        </div>
      </div>
      
      <div className="mb-4">
        <div className="text-sm italic text-slate-400 mb-2 border-l-2 border-slate-600 pl-3">
          "{persona.keyQuote}"
        </div>
        <p className="text-sm text-slate-300 leading-relaxed">
          {persona.perspective}
        </p>
      </div>

      <div className="pt-3 border-t border-slate-700">
        <div className="text-xs text-red-400 font-semibold uppercase mb-1">主要擔憂</div>
        <div className="text-sm text-slate-400">{persona.concern}</div>
      </div>
    </div>
  );
};

interface VerdictCardProps {
  title: string;
  content: string;
  type: 'aggressive' | 'balanced' | 'conservative';
}

const VerdictCard: React.FC<VerdictCardProps> = ({ title, content, type }) => {
  const colors = {
    aggressive: 'bg-gradient-to-br from-orange-500/10 to-red-500/10 border-orange-500/30 text-orange-400',
    balanced: 'bg-gradient-to-br from-blue-500/10 to-cyan-500/10 border-blue-500/30 text-blue-400',
    conservative: 'bg-gradient-to-br from-purple-500/10 to-indigo-500/10 border-purple-500/30 text-purple-400',
  };

  return (
    <div className={`rounded-xl p-6 border ${colors[type]}`}>
      <h4 className="font-bold text-lg mb-3 capitalize">{title}</h4>
      <p className="text-slate-300 text-sm leading-relaxed">{content}</p>
    </div>
  );
};

export default AnalysisDashboard;
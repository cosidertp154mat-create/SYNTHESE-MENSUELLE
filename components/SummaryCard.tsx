
import React from 'react';
import { SynthesisResult } from '../types';

interface SummaryCardProps {
  result: SynthesisResult;
}

const SummaryCard: React.FC<SummaryCardProps> = ({ result }) => {
  return (
    <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 flex flex-col items-center justify-center text-center">
      <div className={`w-24 h-24 rounded-full flex items-center justify-center mb-4 border-4 ${result.color} bg-opacity-10`}>
        <span className={`text-3xl font-bold ${result.color.replace('border-', 'text-')}`}>
          {result.percentage.toFixed(0)}%
        </span>
      </div>
      <h3 className="text-xl font-bold text-slate-800 mb-1">{result.grade}</h3>
      <p className="text-slate-500 text-sm">
        Score Total: <span className="font-semibold text-slate-700">{result.totalScore}</span> / {result.totalMax}
      </p>
    </div>
  );
};

export default SummaryCard;

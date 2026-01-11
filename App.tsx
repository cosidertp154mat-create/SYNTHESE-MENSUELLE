
import React, { useState, useMemo, useEffect } from 'react';
import { Machine, GlobalSettings, DailyLog, DayType, DataPoint } from './types';
import { 
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, 
  Legend, ComposedChart, Bar
} from 'recharts';
import { generateAiSynthesis } from './services/geminiService';

const App: React.FC = () => {
  // --- STATES ---
  const [darkMode, setDarkMode] = useState(() => localStorage.getItem('theme') === 'dark');
  const [isAiLoading, setIsAiLoading] = useState(false);
  const [aiReport, setAiReport] = useState<string | null>(null);
  const [isAddModalOpen, setIsAddModalOpen] = useState(false);
  
  // New machine form state
  const [newMachine, setNewMachine] = useState({
    name: '',
    code: '',
    brand: '',
    type: ''
  });

  useEffect(() => {
    if (darkMode) {
      document.documentElement.classList.add('dark');
      localStorage.setItem('theme', 'dark');
    } else {
      document.documentElement.classList.remove('dark');
      localStorage.setItem('theme', 'light');
    }
  }, [darkMode]);

  const getInitialPeriod = () => {
    const now = new Date();
    let month = now.getDate() < 21 ? now.getMonth() - 1 : now.getMonth();
    let year = now.getFullYear();
    if (month < 0) { month = 11; year--; }
    return { month, year };
  };

  const [period, setPeriod] = useState(getInitialPeriod());
  
  const [machines, setMachines] = useState<Machine[]>(() => {
    const saved = localStorage.getItem('fleet_db');
    return saved ? JSON.parse(saved) : [
      { id: '1', name: 'Pelle Hydraulique', code: 'PH-01', brand: 'CAT', type: '320D', prevCounter: 5200, currentCounter: 5350, breakdownDays: [], worksOnWeekends: false }
    ];
  });

  useEffect(() => {
    localStorage.setItem('fleet_db', JSON.stringify(machines));
  }, [machines]);

  const [settings, setSettings] = useState<GlobalSettings>({
    intemperieDays: [],
    standardHoursPerDay: 8
  });
  const [activeTab, setActiveTab] = useState<'FLEET' | 'CALENDAR'>('FLEET');
  const [selectedMachineId, setSelectedMachineId] = useState<string | null>(machines[0]?.id || null);

  // --- CALCULS ---
  const daysInPeriod = useMemo(() => {
    const dates: string[] = [];
    const curr = new Date(period.year, period.month, 21);
    const end = new Date(period.year, period.month + 1, 20);
    while (curr <= end) {
      dates.push(new Date(curr).toISOString().split('T')[0]);
      curr.setDate(curr.getDate() + 1);
    }
    return dates;
  }, [period]);

  const isWeekend = (dateStr: string) => {
    const day = new Date(dateStr).getDay();
    return day === 5 || day === 6;
  };

  const distributeHours = (totalToDistribute: number, daysCount: number) => {
    if (daysCount <= 0) return Array(daysCount).fill(0);
    const base = Math.floor(totalToDistribute / daysCount);
    const remainder = totalToDistribute % daysCount;
    const distributed = new Array(daysCount).fill(base);
    for (let i = 0; i < remainder; i++) { distributed[i] += 1; }
    return distributed;
  };

  const calculateDailyLogs = (machine: Machine): DailyLog[] => {
    const totalActualHours = machine.currentCounter - machine.prevCounter;
    let logs: DailyLog[] = daysInPeriod.map(date => {
      const isWE = isWeekend(date);
      const isIntemperie = settings.intemperieDays.includes(date);
      const isPanne = machine.breakdownDays.includes(date);
      let type: DayType = 'WORK';
      if (isWE && !machine.worksOnWeekends) type = 'WEEKEND';
      else if (isIntemperie) type = 'INTEMPERIE';
      else if (isPanne) type = 'PANNE';
      return { date, type, workHours: 0, overtimeHours: 0, idleHours: 0, breakdownHours: type === 'PANNE' ? 8 : 0 };
    });
    const workDaysIndexes = logs.reduce((acc, l, idx) => l.type === 'WORK' ? [...acc, idx] : acc, [] as number[]);
    const breakdownTotal = logs.reduce((sum, l) => sum + l.breakdownHours, 0);
    let remainingToDistribute = totalActualHours - breakdownTotal;
    if (workDaysIndexes.length > 0) {
      const potentialWorkHours = workDaysIndexes.length * 8;
      if (remainingToDistribute >= potentialWorkHours) {
        const distribution = distributeHours(remainingToDistribute - potentialWorkHours, workDaysIndexes.length);
        workDaysIndexes.forEach((idx, i) => { logs[idx].workHours = 8; logs[idx].overtimeHours = distribution[i]; });
      } else {
        const distribution = distributeHours(Math.max(0, remainingToDistribute), workDaysIndexes.length);
        workDaysIndexes.forEach((idx, i) => { logs[idx].workHours = distribution[i]; logs[idx].idleHours = 8 - distribution[i]; });
      }
    }
    return logs;
  };

  const selectedMachine = machines.find(m => m.id === selectedMachineId);
  const logs = selectedMachine ? calculateDailyLogs(selectedMachine) : [];
  const totalService = selectedMachine ? selectedMachine.currentCounter - selectedMachine.prevCounter : 0;
  const totalPanne = logs.reduce((sum, l) => sum + l.breakdownHours, 0);
  const totalChomage = logs.reduce((sum, l) => sum + l.idleHours, 0);

  const totalPotential = useMemo(() => {
    if (!selectedMachine) return 0;
    return daysInPeriod.filter(d => {
      const isWE = isWeekend(d);
      const isInt = settings.intemperieDays.includes(d);
      if (isInt) return false;
      if (isWE && !selectedMachine.worksOnWeekends) return false;
      return true;
    }).length * 8;
  }, [daysInPeriod, selectedMachine, settings.intemperieDays]);

  // --- ACTIONS ---
  const handleExcelImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = (window as any).XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = (window as any).XLSX.utils.sheet_to_json(ws);
        const newMachines = data.map((item: any) => ({
          id: Date.now().toString() + Math.random(),
          name: item.Engin || item.Name || "Engin Importé",
          code: item.Code || '',
          brand: item.Marque || item.Brand || '',
          type: item.Type || '',
          prevCounter: parseFloat(item.Ancien || 0),
          currentCounter: parseFloat(item.Nouveau || 0),
          breakdownDays: [],
          worksOnWeekends: false
        }));
        setMachines(prev => [...prev, ...newMachines]);
      } catch (err) { alert("Erreur d'importation"); }
    };
    reader.readAsBinaryString(file);
    e.target.value = '';
  };

  const handleExportExcel = () => {
    if (!selectedMachine) return;
    const data = logs.map(l => ({ Date: l.date, Status: l.type, "H. Travail": l.workHours, "H. Supp": l.overtimeHours, "Total": l.workHours + l.overtimeHours, "Chômage": l.idleHours, "Panne": l.breakdownHours }));
    const ws = (window as any).XLSX.utils.json_to_sheet(data);
    const wb = (window as any).XLSX.utils.book_new();
    (window as any).XLSX.utils.book_append_sheet(wb, ws, "Pointage");
    (window as any).XLSX.writeFile(wb, `Pointage_${selectedMachine.name}.xlsx`);
  };

  const handleAiSynthesis = async () => {
    if (!selectedMachine) return;
    setIsAiLoading(true);
    const points: DataPoint[] = logs.map(l => ({ date: l.date, label: `Jour ${new Date(l.date).getDate()}`, category: l.type, value: l.workHours + l.overtimeHours, maxValue: 8 }));
    const label = `${new Date(period.year, period.month, 21).toLocaleDateString('fr-FR', {month: 'long'})}`;
    const result = await generateAiSynthesis(points, label);
    setAiReport(result || "Erreur");
    setIsAiLoading(false);
  };

  const handleCreateMachine = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newMachine.name) return;
    const m: Machine = {
      id: Date.now().toString(),
      name: newMachine.name,
      code: newMachine.code,
      brand: newMachine.brand,
      type: newMachine.type,
      prevCounter: 0,
      currentCounter: 0,
      breakdownDays: [],
      worksOnWeekends: false
    };
    setMachines([...machines, m]);
    setNewMachine({ name: '', code: '', brand: '', type: '' });
    setIsAddModalOpen(false);
    setSelectedMachineId(m.id);
  };

  const deleteMachine = (id: string) => {
    if (confirm("Supprimer cet engin ?")) {
      const filtered = machines.filter(m => m.id !== id);
      setMachines(filtered);
      if (selectedMachineId === id) setSelectedMachineId(filtered[0]?.id || null);
    }
  };

  const toggleGlobalIntemperie = (date: string) => {
    setSettings(prev => ({
      ...prev,
      intemperieDays: prev.intemperieDays.includes(date) ? prev.intemperieDays.filter(d => d !== date) : [...prev.intemperieDays, date]
    }));
  };

  const toggleMachinePanne = (date: string) => {
    if (!selectedMachineId) return;
    setMachines(prev => prev.map(m => m.id === selectedMachineId ? { ...m, breakdownDays: m.breakdownDays.includes(date) ? m.breakdownDays.filter(d => d !== date) : [...m.breakdownDays, date] } : m));
  };

  return (
    <div className="min-h-screen bg-slate-50 dark:bg-slate-950 transition-colors">
      {/* ADD MACHINE MODAL */}
      {isAddModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-300">
          <div className="bg-white dark:bg-slate-900 w-full max-w-md rounded-[2.5rem] shadow-2xl border border-slate-200 dark:border-slate-800 p-8 transform animate-in zoom-in-95 duration-300">
            <div className="flex justify-between items-center mb-6">
              <h2 className="text-xl font-black tracking-tighter uppercase dark:text-white">Nouvel Engin</h2>
              <button onClick={() => setIsAddModalOpen(false)} className="text-slate-400 hover:text-rose-500 transition-colors"><i className="fas fa-times-circle text-xl"></i></button>
            </div>
            <form onSubmit={handleCreateMachine} className="space-y-4">
              <div>
                <label className="text-[10px] font-black text-slate-400 uppercase ml-1">Nom / Désignation *</label>
                <input required className="w-full bg-slate-50 dark:bg-slate-800 rounded-2xl p-4 font-bold text-sm outline-none border-2 border-transparent focus:border-indigo-500 transition-all dark:text-white" value={newMachine.name} onChange={e => setNewMachine({...newMachine, name: e.target.value})} placeholder="ex: Pelle Hydraulique"/>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-slate-400 uppercase ml-1">Code Parc</label>
                  <input className="w-full bg-slate-50 dark:bg-slate-800 rounded-2xl p-4 font-bold text-sm outline-none border-2 border-transparent focus:border-indigo-500 transition-all dark:text-white" value={newMachine.code} onChange={e => setNewMachine({...newMachine, code: e.target.value})} placeholder="ex: PH-001"/>
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-400 uppercase ml-1">Marque</label>
                  <input className="w-full bg-slate-50 dark:bg-slate-800 rounded-2xl p-4 font-bold text-sm outline-none border-2 border-transparent focus:border-indigo-500 transition-all dark:text-white" value={newMachine.brand} onChange={e => setNewMachine({...newMachine, brand: e.target.value})} placeholder="ex: CAT, Komatsu"/>
                </div>
              </div>
              <div>
                <label className="text-[10px] font-black text-slate-400 uppercase ml-1">Type / Modèle</label>
                <input className="w-full bg-slate-50 dark:bg-slate-800 rounded-2xl p-4 font-bold text-sm outline-none border-2 border-transparent focus:border-indigo-500 transition-all dark:text-white" value={newMachine.type} onChange={e => setNewMachine({...newMachine, type: e.target.value})} placeholder="ex: 320D, D8R"/>
              </div>
              <button type="submit" className="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-black py-4 rounded-2xl shadow-lg shadow-indigo-500/20 transition-all mt-4 uppercase tracking-widest text-xs">
                Ajouter à la flotte
              </button>
            </form>
          </div>
        </div>
      )}

      <header className="bg-slate-900 dark:bg-slate-950 text-white p-6 shadow-xl sticky top-0 z-40 border-b border-slate-800">
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row justify-between items-center gap-4">
          <div className="flex items-center space-x-4">
            <div className="bg-indigo-600 p-3 rounded-2xl shadow-lg shadow-indigo-500/20">
              <i className="fas fa-truck-monster text-2xl"></i>
            </div>
            <div className="flex flex-col">
              <h1 className="text-xl font-black tracking-tighter uppercase">Pointage <span className="text-indigo-400">Pro</span></h1>
              <div className="flex items-center bg-slate-800 rounded-lg px-2 py-0.5 mt-1">
                <select className="bg-transparent text-[10px] font-black uppercase text-slate-400 outline-none cursor-pointer" value={period.month} onChange={(e) => setPeriod(p => ({ ...p, month: parseInt(e.target.value) }))}>
                  {Array.from({length: 12}).map((_, i) => (
                    <option key={i} value={i} className="bg-slate-900">{new Date(0, i).toLocaleDateString('fr-FR', {month: 'long'})}</option>
                  ))}
                </select>
                <span className="mx-1 text-slate-600">/</span>
                <select className="bg-transparent text-[10px] font-black uppercase text-slate-400 outline-none cursor-pointer" value={period.year} onChange={(e) => setPeriod(p => ({ ...p, year: parseInt(e.target.value) }))}>
                  {[2023, 2024, 2025, 2026].map(y => <option key={y} value={y} className="bg-slate-900">{y}</option>)}
                </select>
              </div>
            </div>
          </div>
          <div className="flex items-center space-x-3">
            <div className="flex bg-slate-800 dark:bg-slate-900 rounded-2xl p-1 border border-slate-700">
              <button onClick={() => setActiveTab('FLEET')} className={`px-4 py-2 rounded-xl text-xs font-black transition-all ${activeTab === 'FLEET' ? 'bg-indigo-600 text-white shadow-lg' : 'text-slate-400 hover:text-white'}`}>TABLEAU DE BORD</button>
              <button onClick={() => setActiveTab('CALENDAR')} className={`px-4 py-2 rounded-xl text-xs font-black transition-all ${activeTab === 'CALENDAR' ? 'bg-indigo-600 text-white shadow-lg' : 'text-slate-400 hover:text-white'}`}>CALENDRIER</button>
            </div>
            <button onClick={() => setDarkMode(!darkMode)} className="w-10 h-10 flex items-center justify-center rounded-2xl bg-slate-800 border border-slate-700 hover:bg-slate-700 transition-colors">
              <i className={`fas ${darkMode ? 'fa-sun text-amber-400' : 'fa-moon text-indigo-400'}`}></i>
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto p-6 grid grid-cols-1 lg:grid-cols-12 gap-8">
        <div className="lg:col-span-3 space-y-6">
          <div className="bg-white dark:bg-slate-900 rounded-[2.5rem] p-6 shadow-sm border border-slate-200 dark:border-slate-800">
            <div className="flex justify-between items-center mb-6">
              <h2 className="font-black text-slate-800 dark:text-slate-200 text-xs uppercase tracking-tighter">Votre Flotte ({machines.length})</h2>
              <div className="flex space-x-2">
                 <label className="cursor-pointer text-emerald-500 hover:scale-110 transition-transform">
                   <i className="fas fa-file-excel text-lg"></i>
                   <input type="file" hidden accept=".xlsx, .xls" onChange={handleExcelImport} />
                 </label>
                 <button onClick={() => setIsAddModalOpen(true)} className="text-indigo-600 dark:text-indigo-400 hover:scale-110 transition-transform">
                   <i className="fas fa-plus-circle text-lg"></i>
                 </button>
              </div>
            </div>
            <div className="space-y-3 max-h-[500px] overflow-y-auto custom-scrollbar pr-1">
              {machines.map(m => (
                <div key={m.id} className="relative group">
                  <button onClick={() => setSelectedMachineId(m.id)} className={`w-full text-left p-4 rounded-3xl transition-all border-2 relative overflow-hidden ${selectedMachineId === m.id ? 'border-indigo-500 bg-indigo-50 dark:bg-indigo-950/30' : 'border-transparent bg-slate-50 dark:bg-slate-800/50 hover:bg-slate-100 dark:hover:bg-slate-800'}`}>
                    {m.code && <span className="absolute top-0 right-0 bg-slate-200 dark:bg-slate-700 text-slate-500 dark:text-slate-400 text-[7px] font-black px-2 py-0.5 rounded-bl-lg uppercase">{m.code}</span>}
                    <p className={`font-black text-sm truncate pr-8 ${selectedMachineId === m.id ? 'text-indigo-600 dark:text-indigo-400' : 'text-slate-700 dark:text-slate-300'}`}>{m.name}</p>
                    <p className="text-[10px] text-slate-400 font-bold uppercase mt-1">{m.brand} {m.type} • Σ {m.currentCounter - m.prevCounter}h</p>
                  </button>
                </div>
              ))}
            </div>
          </div>

          {selectedMachine && (
            <div className="bg-white dark:bg-slate-900 rounded-[2.5rem] p-6 shadow-sm border border-slate-200 dark:border-slate-800 space-y-6">
              <div>
                <h3 className="text-xs font-black text-slate-400 dark:text-slate-500 uppercase mb-4 tracking-widest">Configuration</h3>
                <button 
                  onClick={() => setMachines(machines.map(m => m.id === selectedMachineId ? { ...m, worksOnWeekends: !m.worksOnWeekends } : m))}
                  className={`w-full flex items-center justify-between p-3 rounded-2xl border-2 transition-all ${selectedMachine.worksOnWeekends ? 'bg-amber-50 dark:bg-amber-950/20 border-amber-500/30 text-amber-700 dark:text-amber-400' : 'bg-slate-50 dark:bg-slate-800/30 border-slate-100 dark:border-slate-800 text-slate-400'}`}
                >
                  <span className="text-[10px] font-black uppercase">Travail 7j/7</span>
                  <div className={`w-8 h-4 rounded-full relative transition-colors ${selectedMachine.worksOnWeekends ? 'bg-amber-500' : 'bg-slate-300 dark:bg-slate-700'}`}>
                    <div className={`absolute top-0.5 w-3 h-3 bg-white rounded-full transition-all ${selectedMachine.worksOnWeekends ? 'right-0.5' : 'left-0.5'}`}></div>
                  </div>
                </button>
              </div>
              <div className="space-y-4">
                <h3 className="text-xs font-black text-slate-400 dark:text-slate-500 uppercase tracking-widest">Compteurs</h3>
                <input type="number" className="w-full bg-slate-50 dark:bg-slate-800 rounded-2xl p-4 font-black text-sm outline-none dark:text-white" value={selectedMachine.prevCounter} onChange={(e) => setMachines(machines.map(m => m.id === selectedMachineId ? {...m, prevCounter: Number(e.target.value)} : m))} placeholder="Ancien Compteur"/>
                <input type="number" className="w-full bg-slate-50 dark:bg-slate-800 rounded-2xl p-4 font-black text-sm outline-none text-indigo-600 dark:text-indigo-400" value={selectedMachine.currentCounter} onChange={(e) => setMachines(machines.map(m => m.id === selectedMachineId ? {...m, currentCounter: Number(e.target.value)} : m))} placeholder="Nouveau Compteur"/>
              </div>
              <button onClick={() => deleteMachine(selectedMachine.id)} className="w-full bg-rose-50 dark:bg-rose-950/20 text-rose-600 dark:text-rose-400 py-3 rounded-2xl font-black text-[10px] uppercase tracking-widest hover:bg-rose-100 transition-all"><i className="fas fa-trash-alt mr-2"></i>Supprimer</button>
            </div>
          )}
        </div>

        <div className="lg:col-span-9 space-y-8">
          {activeTab === 'FLEET' ? (
            <>
              <div className="flex flex-col md:flex-row justify-between items-center gap-6">
                <div className="grid grid-cols-2 md:grid-cols-4 gap-4 flex-1 w-full">
                  <div className="bg-white dark:bg-slate-900 p-5 rounded-[2rem] border border-slate-200 dark:border-slate-800 shadow-sm"><p className="text-[10px] font-black text-slate-400 uppercase mb-1">Potentiel</p><p className="text-xl font-black dark:text-white">{totalPotential}h</p></div>
                  <div className="bg-white dark:bg-slate-900 p-5 rounded-[2rem] border border-slate-200 dark:border-slate-800 shadow-sm"><p className="text-[10px] font-black text-slate-400 uppercase mb-1">Service</p><p className="text-xl font-black text-indigo-600 dark:text-indigo-400">{totalService}h</p></div>
                  <div className="bg-white dark:bg-slate-900 p-5 rounded-[2rem] border border-slate-200 dark:border-slate-800 shadow-sm"><p className="text-[10px] font-black text-slate-400 uppercase mb-1">Panne</p><p className="text-xl font-black text-rose-500">{totalPanne}h</p></div>
                  <div className="bg-white dark:bg-slate-900 p-5 rounded-[2rem] border border-slate-200 dark:border-slate-800 shadow-sm"><p className="text-[10px] font-black text-slate-400 uppercase mb-1">Chômage</p><p className="text-xl font-black text-amber-500">{totalChomage}h</p></div>
                </div>
                <div className="flex space-x-2">
                  <button onClick={handleExportExcel} className="bg-emerald-500 text-white px-5 py-4 rounded-3xl font-black text-xs shadow-lg shadow-emerald-500/20"><i className="fas fa-file-excel mr-2"></i>EXCEL</button>
                  <button onClick={handleAiSynthesis} disabled={isAiLoading} className="bg-indigo-600 text-white px-5 py-4 rounded-3xl font-black text-xs shadow-lg shadow-indigo-500/20 flex items-center">{isAiLoading ? <i className="fas fa-circle-notch fa-spin"></i> : <i className="fas fa-brain mr-2"></i>}SYNTHÈSE IA</button>
                </div>
              </div>

              {aiReport && (
                <div className="bg-indigo-50 dark:bg-indigo-900/20 p-8 rounded-[3rem] border border-indigo-100 dark:border-indigo-800 relative animate-in slide-in-from-top duration-500">
                  <button onClick={() => setAiReport(null)} className="absolute top-6 right-6 text-indigo-400"><i className="fas fa-times-circle"></i></button>
                  <h3 className="text-lg font-black text-indigo-700 dark:text-indigo-400 mb-4 uppercase tracking-tighter">Analyse Stratégique Gemini</h3>
                  <div className="prose dark:prose-invert max-w-none text-slate-700 dark:text-slate-300 whitespace-pre-wrap text-sm leading-relaxed">{aiReport}</div>
                </div>
              )}

              <div className="bg-white dark:bg-slate-900 p-8 rounded-[3rem] border border-slate-200 dark:border-slate-800 shadow-sm">
                <h3 className="text-lg font-black mb-8 tracking-tighter dark:text-white uppercase">Performance Quotidienne</h3>
                <div className="h-72 w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <ComposedChart data={logs}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke={darkMode ? "#1e293b" : "#f1f5f9"} />
                      <XAxis dataKey="date" tickFormatter={(str) => new Date(str).getDate().toString()} axisLine={false} tickLine={false} tick={{fill: '#94a3b8', fontSize: 10}} />
                      <YAxis axisLine={false} tickLine={false} tick={{fill: '#94a3b8', fontSize: 10}} />
                      <Tooltip contentStyle={{borderRadius: '20px', border: 'none', backgroundColor: darkMode ? '#0f172a' : '#fff', boxShadow: '0 20px 40px -10px rgba(0,0,0,0.2)'}} />
                      <Bar dataKey="workHours" stackId="a" fill="#6366f1" radius={[0, 0, 0, 0]} barSize={12} />
                      <Bar dataKey="overtimeHours" stackId="a" fill="#f59e0b" radius={[4, 4, 0, 0]} barSize={12} />
                    </ComposedChart>
                  </ResponsiveContainer>
                </div>
              </div>

              <div className="bg-white dark:bg-slate-900 rounded-[3rem] border border-slate-200 dark:border-slate-800 overflow-hidden shadow-sm">
                <div className="overflow-x-auto">
                  <table className="w-full text-left">
                    <thead className="bg-slate-50 dark:bg-slate-800/50 border-b border-slate-100 dark:border-slate-800">
                      <tr>
                        <th className="p-6 text-[10px] font-black text-slate-400 uppercase">Jour</th>
                        <th className="p-6 text-[10px] font-black text-slate-400 uppercase text-center">T. Travail</th>
                        <th className="p-6 text-[10px] font-black text-slate-400 uppercase text-center bg-indigo-50/50 dark:bg-indigo-900/10">TOTAL</th>
                        <th className="p-6 text-[10px] font-black text-slate-400 uppercase text-center">Chômage</th>
                        <th className="p-6 text-[10px] font-black text-slate-400 uppercase text-center">Panne</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-50 dark:divide-slate-800">
                      {logs.map(log => (
                        <tr key={log.date} className={isWeekend(log.date) ? 'bg-slate-50/30 dark:bg-slate-800/20' : ''}>
                          <td className="p-6 text-xs font-black dark:text-slate-200">
                            <span className="text-[10px] text-slate-400 uppercase block">{new Date(log.date).toLocaleDateString('fr-FR', {weekday: 'short'})}</span>
                            {new Date(log.date).getDate()}
                          </td>
                          <td className="p-6 text-center font-bold text-xs dark:text-slate-300">{log.workHours || '-'}</td>
                          <td className="p-6 text-center font-black text-indigo-600 dark:text-indigo-400 bg-indigo-50/30 dark:bg-indigo-900/5">{log.workHours + log.overtimeHours || '-'}</td>
                          <td className="p-6 text-center font-bold text-xs text-slate-300 dark:text-slate-600">{log.idleHours || '-'}</td>
                          <td className="p-6 text-center font-bold text-xs text-rose-500">{log.breakdownHours || '-'}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </>
          ) : (
            <div className="bg-white dark:bg-slate-900 p-10 rounded-[3rem] border border-slate-200 dark:border-slate-800 shadow-sm">
              <h3 className="text-2xl font-black mb-10 tracking-tighter uppercase dark:text-white">Calendrier des Événements</h3>
              <div className="grid grid-cols-7 gap-4">
                {['Dim', 'Lun', 'Mar', 'Mer', 'Jeu', 'Ven', 'Sam'].map(d => <div key={d} className="text-center text-[10px] font-black text-slate-400 uppercase">{d}</div>)}
                {daysInPeriod.map(date => {
                  const isInt = settings.intemperieDays.includes(date);
                  const isP = selectedMachine?.breakdownDays.includes(date);
                  return (
                    <div key={date} className="relative group">
                      <button onClick={() => !isWeekend(date) && toggleGlobalIntemperie(date)} className={`w-full aspect-square rounded-3xl flex flex-col items-center justify-center transition-all border-2 ${isInt ? 'bg-indigo-600 border-indigo-400 text-white' : isWeekend(date) && !selectedMachine?.worksOnWeekends ? 'bg-slate-50 dark:bg-slate-800/30 border-transparent opacity-30 text-slate-400 cursor-not-allowed' : 'bg-white dark:bg-slate-800 border-slate-100 dark:border-slate-800 hover:border-indigo-400 text-slate-700 dark:text-slate-300'}`}>
                        <span className="text-sm font-black">{new Date(date).getDate()}</span>
                        {isInt && <i className="fas fa-cloud-showers-heavy text-[10px] mt-1"></i>}
                      </button>
                      {(!isWeekend(date) || selectedMachine?.worksOnWeekends) && !isInt && (
                        <button onClick={(e) => { e.stopPropagation(); toggleMachinePanne(date); }} className={`absolute -top-1 -right-1 w-7 h-7 rounded-full border-4 border-white dark:border-slate-900 flex items-center justify-center transition-all z-10 shadow-lg ${isP ? 'bg-rose-500 text-white scale-110' : 'bg-slate-200 dark:bg-slate-700 text-slate-400 scale-90 opacity-0 group-hover:opacity-100'}`}><i className="fas fa-wrench text-[10px]"></i></button>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </div>
      </main>
    </div>
  );
};

export default App;


export type DayType = 'WORK' | 'WEEKEND' | 'INTEMPERIE' | 'PANNE';

export interface Machine {
  id: string;
  name: string;
  code?: string;
  brand?: string;
  type?: string;
  prevCounter: number;
  currentCounter: number;
  breakdownDays: string[]; // Dates au format YYYY-MM-DD
  worksOnWeekends: boolean;
}

export interface GlobalSettings {
  intemperieDays: string[];
  standardHoursPerDay: number;
}

export interface DailyLog {
  date: string;
  type: DayType;
  workHours: number;
  overtimeHours: number;
  idleHours: number; // Chômage
  breakdownHours: number;
}

export interface DataPoint {
  date: string;
  label: string;
  category: string;
  value: number;
  maxValue: number;
}

export interface SynthesisResult {
  color: string;
  percentage: number;
  grade: string;
  totalScore: number;
  totalMax: number;
}

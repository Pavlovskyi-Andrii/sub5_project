import React from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { useTranslation } from 'react-i18next';

interface RunningModels {
  valid: boolean;
  reason?: string;
  cs?: number;
  dPrime?: number;
  cs_pace?: number;
  pl_params?: {
    A: number; // S (velocità base in m/s)
    k: number; // E-1 (esponente endurance - 1)
  };
  lt1_range?: {
    min: number;
    max: number;
    estimate: number;
  };
  vo2max_pace?: number;
}

interface RunningSustainableTableProps {
  models: RunningModels;
}

interface SustainablePerformance {
  distance: string;
  meters: number;
  time: number;
  pace: number;
  model: string;
}

const RunningSustainableTable: React.FC<RunningSustainableTableProps> = ({ models }) => {
  const { t } = useTranslation();
  const formatTime = (totalSeconds: number): string => {
    if (!isFinite(totalSeconds) || totalSeconds <= 0) return "--:--";
    
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = Math.round(totalSeconds % 60);
    
    if (hours > 0) {
      return `${hours}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
    } else {
      return `${minutes}:${seconds.toString().padStart(2, '0')}`;
    }
  };

  const formatPace = (secondsPerKm: number): string => {
    if (!isFinite(secondsPerKm) || secondsPerKm <= 0) return "--:--";
    const minutes = Math.floor(secondsPerKm / 60);
    const seconds = Math.round(secondsPerKm % 60);
    return `${minutes}:${seconds.toString().padStart(2, '0')}`;
  };

  if (!models.valid || !models.cs || !models.dPrime || !models.pl_params) {
    return (
      <Card className="bg-slate-800 border-slate-700">
        <CardHeader>
          <CardTitle className="text-red-400 text-xl">{t('running.sustainable.error_title')}</CardTitle>
          <p className="text-slate-400">{t('running.sustainable.error_message')}</p>
        </CardHeader>
      </Card>
    );
  }

  const calculateTimeForDistance = (distanceMeters: number): { time: number; model: string } => {
    const CS = models.cs!;
    const Dprime = models.dPrime!;
    
    // First, estimate the time using CS-D' model to determine which model to use
    let estimatedTime = distanceMeters / CS;
    if (Dprime > 0) {
      // Solve quadratic equation: distance = CS*t + D'
      // Rearrange: CS*t + D' - distance = 0
      // t = (distance - D') / CS
      estimatedTime = (distanceMeters - Dprime) / CS;
      if (estimatedTime < 0) {
        // If negative, distance is too short for steady state
        estimatedTime = distanceMeters / CS;
      }
    }

    if (estimatedTime <= 1020) { // ≤ 17 minutes: use CS-D' model
      // For CS-D' model: distance = CS * time + D'
      // Solving for time: time = (distance - D') / CS
      let time = (distanceMeters - Dprime) / CS;
      if (time < 0) {
        // If negative, distance is within D' capacity
        time = distanceMeters / CS;
      }
      return { time, model: 'CS-D\'' };
    } else {
      // > 17 minutes: use Power Law model
      // distance = S * t^E
      // t = (distance / S)^(1/E)
      const S = models.pl_params!.A;
      const E_minus_1 = models.pl_params!.k;
      const E = E_minus_1 + 1;
      
      const time = Math.pow(distanceMeters / S, 1 / E);
      return { time, model: 'Power Law' };
    }
  };

  const distances = [
    { label: '1500m', meters: 1500 },
    { label: '2000m', meters: 2000 },
    { label: '3000m', meters: 3000 },
    { label: '5000m', meters: 5000 },
    { label: '10km', meters: 10000 },
    { label: '15km', meters: 15000 },
    { label: t('running.sustainable.half_marathon'), meters: 21097.5 },
    { label: '30km', meters: 30000 },
    { label: t('running.sustainable.marathon'), meters: 42195 },
    { label: '50km', meters: 50000 }
  ];

  const sustainablePerformances: SustainablePerformance[] = distances.map(distance => {
    const { time, model } = calculateTimeForDistance(distance.meters);
    const pace = (time / distance.meters) * 1000; // seconds per km
    
    return {
      distance: distance.label,
      meters: distance.meters,
      time,
      pace,
      model
    };
  });

  return (
    <Card className="bg-slate-800 border-slate-700">
      <CardHeader>
        <CardTitle className="text-lime-400 text-xl">{t('running.sustainable.title')}</CardTitle>
        <p className="text-slate-400">
          {t('running.sustainable.subtitle')}
        </p>
      </CardHeader>
      <CardContent>
        {/* Desktop Table - Hidden on Mobile */}
        <div className="desktop-only">
          <Table>
            <TableHeader>
              <TableRow className="bg-slate-700 hover:bg-slate-700">
                <TableHead className="text-white font-bold text-center min-w-[120px]">{t('running.sustainable.distance')}</TableHead>
                <TableHead className="text-white font-bold text-center min-w-[100px]">{t('running.sustainable.predicted_time')}</TableHead>
                <TableHead className="text-white font-bold text-center min-w-[120px]">{t('running.sustainable.average_pace')}</TableHead>
                <TableHead className="text-white font-bold text-center min-w-[100px]">{t('running.sustainable.model')}</TableHead>
              </TableRow>
            </TableHeader>
            <TableBody>
              {sustainablePerformances.map((performance, index) => (
                <TableRow key={index} className="bg-slate-100 hover:bg-slate-200">
                  <TableCell className="text-center font-medium text-slate-800">
                    {performance.distance}
                  </TableCell>
                  <TableCell className="text-center font-medium text-slate-800">
                    {formatTime(performance.time)}
                  </TableCell>
                  <TableCell className="text-center font-medium text-slate-800">
                    {formatPace(performance.pace)} /km
                  </TableCell>
                  <TableCell className="text-center">
                    <span className={`px-2 py-1 rounded text-xs font-semibold ${
                      performance.model === 'CS-D\'' ? 'bg-orange-500 text-white' :
                      'bg-blue-500 text-white'
                    }`}>
                      {performance.model}
                    </span>
                  </TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </div>

        {/* Mobile Cards - Visible only on Mobile */}
        <div className="mobile-only space-y-3">
          {sustainablePerformances.map((performance, index) => (
            <div key={index} className="bg-slate-100 rounded-lg p-4">
              <div className="flex items-center justify-between mb-3">
                <div className="text-slate-800 font-semibold text-lg">
                  {performance.distance}
                </div>
                <span className={`px-2 py-1 rounded text-xs font-semibold ${
                  performance.model === 'CS-D\'' ? 'bg-orange-500 text-white' :
                  'bg-blue-500 text-white'
                }`}>
                  {performance.model}
                </span>
              </div>
              <div className="space-y-2 text-sm text-slate-800">
                <div>
                  <span className="font-medium">{t('running.sustainable.predicted_time')}:</span> {formatTime(performance.time)}
                </div>
                <div>
                  <span className="font-medium">{t('running.sustainable.average_pace')}:</span> {formatPace(performance.pace)} /km
                </div>
              </div>
            </div>
          ))}
        </div>
        
        {/* Model Legend */}
        <div className="mt-6 space-y-3">
          <h4 className="text-lime-400 font-semibold text-lg mb-3">{t('running.sustainable.models_legend')}</h4>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-sm">
            <div className="flex items-center space-x-2">
              <span className="px-2 py-1 rounded text-xs font-semibold bg-orange-500 text-white">CS-D'</span>
              <span className="text-slate-300">≤ 17 minuti</span>
            </div>
            <div className="flex items-center space-x-2">
              <span className="px-2 py-1 rounded text-xs font-semibold bg-blue-500 text-white">Power Law</span>
              <span className="text-slate-300">&gt; 17 minuti</span>
            </div>
          </div>
        </div>
      </CardContent>
    </Card>
  );
};

export default RunningSustainableTable;
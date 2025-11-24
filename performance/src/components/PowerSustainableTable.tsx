import React from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { useTranslation } from 'react-i18next';

interface PowerModels {
  valid: boolean;
  reason?: string;
  cp?: number;
  wPrime?: number;
  pMax?: number;
  apr_params?: {
    po3min: number;
    apr: number;
    k: number;
  };
  pl_params?: {
    S: number;
    E: number;
  };
}

interface PowerSustainableTableProps {
  models: PowerModels;
  bodyWeight: number;
}

interface SustainablePower {
  duration: string;
  absolutePower: number;
  relativePower: number;
  model: string;
}

const PowerSustainableTable: React.FC<PowerSustainableTableProps> = ({ models, bodyWeight }) => {
  const { t } = useTranslation();
  if (!models.valid || !models.cp || !models.wPrime || !models.pMax || !models.apr_params || !models.pl_params) {
    return (
      <Card className="bg-slate-800 border-slate-700">
        <CardHeader>
          <CardTitle className="text-red-400 text-xl">{t('sustainable.title')}</CardTitle>
          <p className="text-slate-400">{t('sustainable.unavailable')}</p>
        </CardHeader>
      </Card>
    );
  }

  const calculatePowerForDuration = (seconds: number): { power: number; model: string } => {
    if (seconds < 180) {
      // APR model for < 3 minutes
      if (seconds === 5) {
        return { power: models.pMax!, model: 'APR' };
      } else {
        const power = models.apr_params!.po3min + models.apr_params!.apr * Math.exp(-models.apr_params!.k * seconds);
        return { power, model: 'APR' };
      }
    } else if (seconds <= 960) {
      // CP-W' model for 3-16 minutes
      const power = models.cp! + models.wPrime! / seconds;
      return { power, model: 'CP-W\'' };
    } else {
      // Power Law model for > 16 minutes
      const power = models.pl_params!.S * Math.pow(seconds, models.pl_params!.E - 1);
      return { power, model: 'Power Law' };
    }
  };

  const durations = [
    { label: '5s', seconds: 5 },
    { label: '60s', seconds: 60 },
    { label: '3m', seconds: 180 },
    { label: '5m', seconds: 300 },
    { label: '10m', seconds: 600 },
    { label: '20m', seconds: 1200 },
    { label: '30m', seconds: 1800 },
    { label: '45m', seconds: 2700 },
    { label: '60m', seconds: 3600 }
  ];

  const sustainablePowers: SustainablePower[] = durations.map(duration => {
    const { power, model } = calculatePowerForDuration(duration.seconds);
    return {
      duration: duration.label,
      absolutePower: power,
      relativePower: power / bodyWeight,
      model
    };
  });

  return (
    <Card className="bg-slate-800 border-slate-700">
      <CardHeader>
        <CardTitle className="text-lime-400 text-xl">{t('sustainable.title')}</CardTitle>
        <p className="text-slate-400">
          {t('sustainable.description')}
        </p>
      </CardHeader>
      <CardContent>
        {/* Desktop Table - Hidden on Mobile */}
        <div className="desktop-only">
          <Table>
            <TableHeader>
              <TableRow className="bg-slate-700 hover:bg-slate-700">
                <TableHead className="text-white font-bold text-center w-20">{t('sustainable.duration')}</TableHead>
                <TableHead className="text-white font-bold text-center min-w-[120px]">{t('sustainable.absolute_power')}</TableHead>
                <TableHead className="text-white font-bold text-center min-w-[120px]">{t('sustainable.relative_power')}</TableHead>
                <TableHead className="text-white font-bold text-center min-w-[100px]">{t('sustainable.model')}</TableHead>
              </TableRow>
            </TableHeader>
            <TableBody>
              {sustainablePowers.map((power, index) => (
                <TableRow key={index} className="bg-slate-100 hover:bg-slate-200">
                  <TableCell className="text-center font-medium text-slate-800">
                    {power.duration}
                  </TableCell>
                  <TableCell className="text-center font-medium text-slate-800">
                    {Math.round(power.absolutePower)} W
                  </TableCell>
                  <TableCell className="text-center font-medium text-slate-800">
                    {power.relativePower.toFixed(1)} W/kg
                  </TableCell>
                  <TableCell className="text-center">
                    <span className={`px-2 py-1 rounded text-xs font-semibold ${
                      power.model === 'APR' ? 'bg-purple-500 text-white' :
                      power.model === 'CP-W\'' ? 'bg-orange-500 text-white' :
                      'bg-blue-500 text-white'
                    }`}>
                      {power.model}
                    </span>
                  </TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </div>

        {/* Mobile Cards - Visible only on Mobile */}
        <div className="mobile-only space-y-3">
          {sustainablePowers.map((power, index) => (
            <div key={index} className="bg-slate-100 rounded-lg p-4">
              <div className="flex items-center justify-between mb-3">
                <div className="text-slate-800 font-semibold text-lg">
                  {power.duration}
                </div>
                <span className={`px-2 py-1 rounded text-xs font-semibold ${
                  power.model === 'APR' ? 'bg-purple-500 text-white' :
                  power.model === 'CP-W\'' ? 'bg-orange-500 text-white' :
                  'bg-blue-500 text-white'
                }`}>
                  {power.model}
                </span>
              </div>
              <div className="grid grid-cols-2 gap-4 text-sm text-slate-800">
                <div>
                  <span className="font-medium">Potenza Assoluta:</span> {Math.round(power.absolutePower)} W
                </div>
                <div>
                  <span className="font-medium">Potenza Relativa:</span> {power.relativePower.toFixed(1)} W/kg
                </div>
              </div>
            </div>
          ))}
        </div>
        
        {/* Model Legend */}
        <div className="mt-6 space-y-3">
          <h4 className="text-lime-400 font-semibold text-lg mb-3">{t('sustainable.models_used')}</h4>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-sm">
            <div className="flex items-center space-x-2">
              <span className="px-2 py-1 rounded text-xs font-semibold bg-purple-500 text-white">APR</span>
              <span className="text-slate-300">{t('sustainable.apr_range')}</span>
            </div>
            <div className="flex items-center space-x-2">
              <span className="px-2 py-1 rounded text-xs font-semibold bg-orange-500 text-white">CP-W'</span>
              <span className="text-slate-300">{t('sustainable.cpw_range')}</span>
            </div>
            <div className="flex items-center space-x-2">
              <span className="px-2 py-1 rounded text-xs font-semibold bg-blue-500 text-white">Power Law</span>
              <span className="text-slate-300">{t('sustainable.powerlaw_range')}</span>
            </div>
          </div>
        </div>
      </CardContent>
    </Card>
  );
};

export default PowerSustainableTable;
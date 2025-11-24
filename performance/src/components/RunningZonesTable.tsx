
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
    A: number; // Ora S (velocità base in m/s)
    k: number; // Ora E-1 (esponente endurance - 1)
  };
  lt1_range?: {
    min: number;
    max: number;
    estimate: number;
  };
  vo2max_pace?: number;
}

interface RunningZonesTableProps {
  models: RunningModels;
}

const RunningZonesTable: React.FC<RunningZonesTableProps> = ({ models }) => {
  const { t } = useTranslation();
  
  const formatPace = (secondsPerKm: number) => {
    if (!isFinite(secondsPerKm) || secondsPerKm <= 0) return "--:--";
    const minutes = Math.floor(secondsPerKm / 60);
    const seconds = Math.round(secondsPerKm % 60);
    return `${minutes}:${seconds.toString().padStart(2, '0')}`;
  };

  if (!models.valid || !models.cs || !models.dPrime || !models.vo2max_pace || !models.cs_pace || !models.pl_params || !models.lt1_range) {
    return (
      <Card className="bg-slate-800 border-slate-700">
        <CardHeader>
          <CardTitle className="text-red-400 text-xl">{t('running.zones.error_title')}</CardTitle>
          <p className="text-slate-400">{t('running.zones.error_message')}</p>
        </CardHeader>
      </Card>
    );
  }

  const CS = models.cs;
  const Dprime = models.dPrime;
  const { lt1_range } = models;
  
  // Use the new LT1 values from models.lt1_range
  const lt1_estimate_pace = lt1_range.estimate;
  const lt1_estimate_speed = 3600 / lt1_estimate_pace; // km/h

  // Funzione per calcolare velocità sostenibile per un dato tempo usando la formula corretta
  function velocityAtTime(tSec: number) {
    if (tSec <= 180) {
      // Per tempi brevi usa CS-D' model
      const dist = CS * tSec + Dprime;
      return (dist / tSec) * 3.6; // convertito in km/h
    } else {
      // Per tempi lunghi usa power law corretto
      if (models.pl_params) {
        const S = models.pl_params.A;     // velocità base
        const E_minus_1 = models.pl_params.k; // E-1
        
        // Formula corretta: v = S * t^(E-1)
        const velocity_ms = S * Math.pow(tSec, E_minus_1);
        return velocity_ms * 3.6; // convertito in km/h
      }
      
      // Fallback: usa una approssimazione basata su CS
      return (CS * 3.6) * Math.pow(180 / tSec, 0.06);
    }
  }

  // Calcola velocità per diversi tempi sostenibili
  const v60min = velocityAtTime(3600);  // 60 min
  const v25min = velocityAtTime(1500);  // 25 min  
  const v3min = velocityAtTime(180);    // 3 min

  // Converte in passo (s/km)
  const p_60min = 3600 / v60min;
  const p_25min = 3600 / v25min;
  const p_3min = 3600 / v3min;

  const zones = [
    {
      zone: 'Z1',
      name: t('running.zones.zones.z1'),
      color: 'bg-green-500',
      textColor: 'text-white',
      speedRange: `≤ ${(lt1_estimate_speed * 0.86).toFixed(2)}`, // < 86% LT1
      paceRange: `≥ ${formatPace(lt1_estimate_pace / 0.86)}`, // > passo di 86% LT1
      rpe: '1-2',
      domain: t('running.zones.domains.moderate'),
      domainColor: 'text-green-500'
    },
    {
      zone: 'Z2',
      name: t('running.zones.zones.z2'),
      color: 'bg-green-400',
      textColor: 'text-green-900',
      speedRange: `${(lt1_estimate_speed * 0.86).toFixed(2)} – ${lt1_estimate_speed.toFixed(2)}`, // 86% LT1 - LT1
      paceRange: `${formatPace(lt1_estimate_pace)} – ${formatPace(lt1_estimate_pace / 0.86)}`,
      rpe: '3-4',
      domain: t('running.zones.domains.moderate'),
      domainColor: 'text-green-500'
    },
    {
      zone: 'Z3',
      name: t('running.zones.zones.z3'),
      color: 'bg-yellow-500',
      textColor: 'text-yellow-900',
      speedRange: `${lt1_estimate_speed.toFixed(2)} – ${v60min.toFixed(2)}`, // LT1 - 60min
      paceRange: `${formatPace(p_60min)} – ${formatPace(lt1_estimate_pace)}`,
      rpe: '5-6',
      domain: t('running.zones.domains.heavy'),
      domainColor: 'text-yellow-500'
    },
    {
      zone: 'Z4',
      name: t('running.zones.zones.z4'),
      color: 'bg-orange-500',
      textColor: 'text-white',
      speedRange: `${v60min.toFixed(2)} – ${v25min.toFixed(2)}`, // 60min - 25min
      paceRange: `${formatPace(p_25min)} – ${formatPace(p_60min)}`,
      rpe: '7',
      domain: t('running.zones.domains.heavy'),
      domainColor: 'text-yellow-500'
    },
    {
      zone: 'Z5',
      name: t('running.zones.zones.z5'),
      color: 'bg-red-500',
      textColor: 'text-white',
      speedRange: `${v25min.toFixed(2)} – ${v3min.toFixed(2)}`, // 25min - 3min
      paceRange: `${formatPace(p_3min)} – ${formatPace(p_25min)}`,
      rpe: '>8',
      domain: t('running.zones.domains.severe'),
      domainColor: 'text-red-500'
    },
    {
      zone: 'Z6',
      name: t('running.zones.zones.z6'),
      color: 'bg-purple-600',
      textColor: 'text-white',
      speedRange: `> ${v3min.toFixed(2)}`, // > 3min
      paceRange: `< ${formatPace(p_3min)}`,
      rpe: '>8',
      domain: t('running.zones.domains.extreme'),
      domainColor: 'text-purple-500'
    }
  ];

  return (
    <Card className="bg-slate-800 border-slate-700">
      <CardHeader>
        <CardTitle className="text-lime-400 text-xl">{t('running.zones.title')}</CardTitle>
        <p className="text-slate-400">
          {t('running.zones.subtitle')}
        </p>
      </CardHeader>
      <CardContent>
        {/* Desktop Table - Hidden on Mobile */}
        <div className="desktop-only">
          <Table>
            <TableHeader>
              <TableRow className="bg-slate-700 hover:bg-slate-700">
                <TableHead className="text-white font-bold text-center w-16">{t('running.zones.zone')}</TableHead>
                <TableHead className="text-white font-bold text-center min-w-[120px]">{t('running.zones.name')}</TableHead>
                <TableHead className="text-white font-bold text-center min-w-[140px]">{t('running.zones.speed')}</TableHead>
                <TableHead className="text-white font-bold text-center min-w-[140px]">{t('running.zones.pace')}</TableHead>
                <TableHead className="text-white font-bold text-center w-20">{t('running.zones.rpe')}</TableHead>
                <TableHead className="text-white font-bold text-center min-w-[100px]">{t('running.zones.domain')}</TableHead>
              </TableRow>
            </TableHeader>
            <TableBody>
              {zones.map((zone) => (
                <TableRow key={zone.zone} className="bg-slate-100 hover:bg-slate-200">
                  <TableCell className="text-center">
                    <span className={`${zone.color} ${zone.textColor} px-2 py-1 rounded font-bold text-sm`}>
                      {zone.zone}
                    </span>
                  </TableCell>
                  <TableCell className="text-center">
                    <span className={`${zone.color} ${zone.textColor} px-3 py-1 rounded font-bold`}>
                      {zone.name}
                    </span>
                  </TableCell>
                  <TableCell className="text-center font-medium text-slate-800">
                    {zone.speedRange}
                  </TableCell>
                  <TableCell className="text-center font-medium text-slate-800">
                    {zone.paceRange}
                  </TableCell>
                  <TableCell className="text-center font-medium text-slate-800">
                    {zone.rpe}
                  </TableCell>
                  <TableCell className="text-center">
                    <span className={`${zone.domainColor} font-semibold`}>
                      {zone.domain}
                    </span>
                  </TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </div>

        {/* Mobile Cards - Visible only on Mobile */}
        <div className="mobile-only space-y-3">
          {zones.map((zone) => (
            <div key={zone.zone} className="bg-slate-100 rounded-lg p-4">
              <div className="flex items-center justify-between mb-3">
                <div className="flex items-center gap-3">
                  <span className={`${zone.color} ${zone.textColor} px-2 py-1 rounded font-bold text-sm`}>
                    {zone.zone}
                  </span>
                  <span className={`${zone.color} ${zone.textColor} px-3 py-1 rounded font-bold`}>
                    {zone.name}
                  </span>
                </div>
                <span className={`${zone.domainColor} font-semibold text-sm`}>
                  {zone.domain}
                </span>
              </div>
                <div className="space-y-2 text-sm text-slate-800">
                <div>
                  <span className="font-medium">{t('running.zones.speed').replace(' (km/h)', '')}:</span> {zone.speedRange} km/h
                </div>
                <div>
                  <span className="font-medium">{t('running.zones.pace').replace(' (min/km)', '')}:</span> {zone.paceRange}
                </div>
                <div>
                  <span className="font-medium">{t('running.zones.rpe')}:</span> {zone.rpe}
                </div>
              </div>
            </div>
          ))}
        </div>
        
        {/* Legend */}
        <div className="mt-6 space-y-3">
          <h4 className="text-lime-400 font-semibold text-lg mb-3">{t('running.zones.domains_legend')}</h4>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-sm mobile-grid-to-stack">
            <div className="space-y-2">
              <div className="flex items-start space-x-2">
                <span className="text-green-500 font-semibold min-w-[80px]">{t('running.zones.domains.moderate')}:</span>
                <span className="text-slate-300">{t('running.zones.domain_descriptions.moderate')}</span>
              </div>
              <div className="flex items-start space-x-2">
                <span className="text-yellow-500 font-semibold min-w-[80px]">{t('running.zones.domains.heavy')}:</span>
                <span className="text-slate-300">{t('running.zones.domain_descriptions.heavy')}</span>
              </div>
            </div>
            <div className="space-y-2">
              <div className="flex items-start space-x-2">
                <span className="text-red-500 font-semibold min-w-[80px]">{t('running.zones.domains.severe')}:</span>
                <span className="text-slate-300">{t('running.zones.domain_descriptions.severe')}</span>
              </div>
              <div className="flex items-start space-x-2">
                <span className="text-purple-500 font-semibold min-w-[80px]">{t('running.zones.domains.extreme')}:</span>
                <span className="text-slate-300">{t('running.zones.domain_descriptions.extreme')}</span>
              </div>
            </div>
          </div>
          
          {/* Empirical note for LT1 */}
          <div className="mt-4 pt-3 border-t border-slate-600">
            <p className="text-xs text-slate-400">
              {t('running.model_parameters.empirical_note')}
            </p>
          </div>
        </div>
      </CardContent>
    </Card>
  );
};

export default RunningZonesTable;

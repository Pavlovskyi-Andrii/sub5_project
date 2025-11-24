
import React from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { useTranslation } from 'react-i18next';

interface Models {
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
  lt1_range?: {
    min: number;
    max: number;
    estimate: number;
  };
}

interface PowerZonesTableProps {
  models: Models;
}

const PowerZonesTable: React.FC<PowerZonesTableProps> = ({ models }) => {
  const { t } = useTranslation();
  
  if (!models.valid || !models.cp || !models.lt1_range || !models.apr_params) {
    return (
      <Card className="bg-slate-800 border-slate-700">
        <CardHeader>
          <CardTitle className="text-red-400 text-xl">{t('cycling.zones.title')}</CardTitle>
          <p className="text-slate-400">{t('cycling.zones.unavailable')}</p>
        </CardHeader>
      </Card>
    );
  }

  const lt1 = models.lt1_range.estimate;
  const cp = models.cp;
  const p3min = models.apr_params.po3min;

  const zones = [
    {
      zone: 'Z1',
      name: t('cycling.zones.zone_names.easy'),
      color: 'bg-green-500',
      textColor: 'text-white',
      powerRange: `< ${Math.round(lt1 * 0.8)}`,
      rpe: '1-2',
      domain: t('common.domains.moderate'),
      domainColor: 'text-green-500'
    },
    {
      zone: 'Z2',
      name: t('cycling.zones.zone_names.endurance'),
      color: 'bg-green-400',
      textColor: 'text-green-900',
      powerRange: `${Math.round(lt1 * 0.8)} - ${Math.round(lt1)}*`,
      rpe: '3-4',
      domain: t('common.domains.moderate'),
      domainColor: 'text-green-500'
    },
    {
      zone: 'Z3',
      name: t('cycling.zones.zone_names.med_sst'),
      color: 'bg-yellow-500',
      textColor: 'text-yellow-900',
      powerRange: `${Math.round(lt1)} - ${Math.round(cp * 0.92)}`,
      rpe: '5-6',
      domain: t('common.domains.heavy'),
      domainColor: 'text-yellow-500'
    },
    {
      zone: 'Z4',
      name: t('cycling.zones.zone_names.cp'),
      color: 'bg-orange-500',
      textColor: 'text-white',
      powerRange: `${Math.round(cp * 0.92)} - ${Math.round(cp * 1.02)}`,
      rpe: '7',
      domain: t('common.domains.heavy'),
      domainColor: 'text-yellow-500'
    },
    {
      zone: 'Z5',
      name: t('cycling.zones.zone_names.hiit'),
      color: 'bg-red-500',
      textColor: 'text-white',
      powerRange: `${Math.round(cp * 1.02)} - ${Math.round(p3min)}`,
      rpe: '>8',
      domain: t('common.domains.severe'),
      domainColor: 'text-red-500'
    },
    {
      zone: 'Z6',
      name: t('cycling.zones.zone_names.sit'),
      color: 'bg-purple-600',
      textColor: 'text-white',
      powerRange: `> ${Math.round(p3min)}`,
      rpe: '>8',
      domain: t('common.domains.extreme'),
      domainColor: 'text-purple-500'
    }
  ];

  return (
    <Card className="bg-slate-800 border-slate-700">
      <CardHeader>
        <CardTitle className="text-lime-400 text-xl">{t('cycling.zones.title')}</CardTitle>
        <p className="text-slate-400">
          {t('cycling.zones.subtitle')}
        </p>
      </CardHeader>
      <CardContent>
        {/* Desktop Table - Hidden on Mobile */}
        <div className="desktop-only">
          <Table>
            <TableHeader>
              <TableRow className="bg-slate-700 hover:bg-slate-700">
                <TableHead className="text-white font-bold text-center w-16">{t('cycling.zones.table_headers.zone')}</TableHead>
                <TableHead className="text-white font-bold text-center min-w-[120px]">{t('cycling.zones.table_headers.name')}</TableHead>
                <TableHead className="text-white font-bold text-center min-w-[140px]">{t('cycling.zones.table_headers.power')}</TableHead>
                <TableHead className="text-white font-bold text-center w-20">{t('cycling.zones.table_headers.rpe')}</TableHead>
                <TableHead className="text-white font-bold text-center min-w-[100px]">{t('cycling.zones.table_headers.domain')}</TableHead>
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
                    {zone.powerRange}
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
              <div className="grid grid-cols-2 gap-4 text-sm text-slate-800">
                <div>
                  <span className="font-medium">{t('cycling.zones.mobile_labels.power')}</span> {zone.powerRange} W
                </div>
                <div>
                  <span className="font-medium">{t('cycling.zones.mobile_labels.rpe')}</span> {zone.rpe}
                </div>
              </div>
            </div>
          ))}
        </div>
        
        {/* Legend */}
        <div className="mt-6 space-y-3">
          <h4 className="text-lime-400 font-semibold text-lg mb-3">{t('cycling.zones.intensity_domains.title')}</h4>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-sm mobile-grid-to-stack">
            <div className="space-y-2">
              <div className="flex items-start space-x-2">
                <span className="text-green-500 font-semibold min-w-[80px]">{t('common.domains.moderate')}:</span>
                <span className="text-slate-300">{t('cycling.zones.intensity_domains.descriptions.moderate')}</span>
              </div>
              <div className="flex items-start space-x-2">
                <span className="text-yellow-500 font-semibold min-w-[80px]">{t('common.domains.heavy')}:</span>
                <span className="text-slate-300">{t('cycling.zones.intensity_domains.descriptions.heavy')}</span>
              </div>
            </div>
            <div className="space-y-2">
              <div className="flex items-start space-x-2">
                <span className="text-red-500 font-semibold min-w-[80px]">{t('common.domains.severe')}:</span>
                <span className="text-slate-300">{t('cycling.zones.intensity_domains.descriptions.severe')}</span>
              </div>
              <div className="flex items-start space-x-2">
                <span className="text-purple-500 font-semibold min-w-[80px]">{t('common.domains.extreme')}:</span>
                <span className="text-slate-300">{t('cycling.zones.intensity_domains.descriptions.extreme')}</span>
              </div>
            </div>
          </div>
        </div>
      </CardContent>
    </Card>
  );
};

export default PowerZonesTable;

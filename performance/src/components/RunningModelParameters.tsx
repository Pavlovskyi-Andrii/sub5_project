
import React from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { PrecisionModels, formatForDisplay } from '../utils/precisionCalculations';
import { useTranslation } from 'react-i18next';

interface RunningModelParametersProps {
  models: PrecisionModels;
}

const RunningModelParameters: React.FC<RunningModelParametersProps> = ({ models }) => {
  const { t } = useTranslation();
  
  if (!models.valid) {
    return (
      <Card className="bg-slate-800 border-slate-700">
        <CardHeader>
          <CardTitle className="text-red-400">{t('running.model_parameters.error_title')}</CardTitle>
        </CardHeader>
        <CardContent>
          <p className="text-slate-400">{t('running.model_parameters.error_message')}</p>
        </CardContent>
      </Card>
    );
  }

  // Parametri S ed E ora sono direttamente A e k+1 - con alta precisione
  const S = models.pl_params?.A || 0;      // S = velocità base (m/s) - full precision
  const E = models.pl_params ? models.pl_params.k + 1 : 0; // E = esponente endurance - full precision

  return (
    <Card className="bg-slate-800 border-slate-700">
      <CardHeader>
        <CardTitle className="text-lime-400">{t('running.model_parameters.title')}</CardTitle>
      </CardHeader>
      <CardContent>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mobile-grid-to-stack">
          {/* CS-D' Model */}
          <div className="bg-slate-900 p-4 rounded-lg border border-red-500/30">
            <h4 className="font-semibold text-red-400 mb-3 flex items-center">
              <div className="w-3 h-3 bg-red-500 rounded mr-2"></div>
              {t('running.model_parameters.hyperbolic_model')}
            </h4>
            <div className="space-y-2 text-sm">
              <div className="flex justify-between">
                <span className="text-slate-400">{t('running.model_parameters.cs')}:</span>
                <span className="text-white font-medium">{models.cs.toFixed(3)} m/s</span>
              </div>
              <div className="flex justify-between">
                <span className="text-slate-400">{t('running.model_parameters.cs_pace')}:</span>
                <span className="text-white font-medium">{formatForDisplay.pace(models.cs_pace)}/km</span>
              </div>
              <div className="flex justify-between">
                <span className="text-slate-400">{t('running.model_parameters.d_prime')}:</span>
                <span className="text-white font-medium">{models.dPrime.toFixed(1)} m</span>
              </div>
            </div>
          </div>

          {/* Power Law Model */}
          <div className="bg-slate-900 p-4 rounded-lg border border-yellow-500/30">
            <h4 className="font-semibold text-yellow-400 mb-3 flex items-center">
              <div className="w-3 h-3 bg-yellow-500 rounded mr-2"></div>
              {t('running.model_parameters.power_law_model')}
            </h4>
            <div className="space-y-2 text-sm">
              <div className="flex justify-between">
                <span className="text-slate-400">{t('running.model_parameters.speed_parameter')}:</span>
                <span className="text-white font-medium">
                  {S.toFixed(4)} m/s
                </span>
              </div>
              <div className="flex justify-between">
                <span className="text-slate-400">{t('running.model_parameters.endurance_parameter')}:</span>
                <span className="text-white font-medium">
                  {E.toFixed(6)}
                </span>
              </div>
              <p className="text-xs text-slate-500 mt-2">
                {t('running.model_parameters.formula_note')}
              </p>
            </div>
          </div>
        </div>

        {/* Thresholds Summary */}
        <div className="mt-6 bg-slate-900 p-4 rounded-lg">
          <h4 className="font-semibold text-lime-400 mb-3">{t('running.model_parameters.thresholds_title')}</h4>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-sm mobile-grid-to-stack">
            <div className="text-center">
              <div className="text-slate-400">{t('running.model_parameters.map_3min')}</div>
              <div className="text-white font-semibold text-lg">{formatForDisplay.pace(models.vo2max_pace)}/km</div>
            </div>
            <div className="text-center">
              <div className="text-slate-400">{t('running.model_parameters.critical_speed')}</div>
              <div className="text-white font-semibold text-lg">{formatForDisplay.pace(models.cs_pace)}/km</div>
            </div>
            <div className="text-center">
              <div className="text-slate-400">{t('running.model_parameters.lt1_estimated')}</div>
              <div className="text-white font-semibold text-lg">{formatForDisplay.pace(models.lt1_range.estimate)}/km</div>
              <div className="text-slate-500 text-xs">({formatForDisplay.pace(models.lt1_range.min)} - {formatForDisplay.pace(models.lt1_range.max)})</div>
            </div>
          </div>
          
          {/* Empirical note */}
          <div className="mt-4 pt-3 border-t border-slate-600">
            <p className="text-xs text-slate-400">
              {t('running.model_parameters.empirical_note')}
            </p>
          </div>
        </div>

        {/* Debug Info - only show in development */}
        {process.env.NODE_ENV === 'development' && (
          <div className="mt-4 p-3 bg-slate-950 rounded border border-slate-600">
            <h5 className="text-xs text-slate-400 mb-2">{t('running.model_parameters.debug_title')}</h5>
            <div className="text-xs text-slate-500 space-y-1">
              <div>CS: {models.cs} m/s</div>
              <div>D': {models.dPrime} m</div>
              <div>S: {S} m/s</div>
              <div>E: {E}</div>
              <div>{t('running.model_parameters.lt1_formula')}: {E > 0.90 ? 'E > 0.90 (CS × 0.84)' : 'E ≤ 0.90 (CS × 0.80)'}</div>
            </div>
          </div>
        )}
      </CardContent>
    </Card>
  );
};

export default RunningModelParameters;

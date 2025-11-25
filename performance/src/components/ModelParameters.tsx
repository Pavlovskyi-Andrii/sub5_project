
import React from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Tooltip, TooltipContent, TooltipProvider, TooltipTrigger } from '@/components/ui/tooltip';
import { Info } from 'lucide-react';
import { useTranslation } from 'react-i18next';
import { PowerPrecisionModels } from '../utils/powerPrecisionCalculations';

interface ModelParametersProps {
  models: PowerPrecisionModels;
  bodyWeight: number;
}

const ModelParameters: React.FC<ModelParametersProps> = ({ models, bodyWeight }) => {
  const { t } = useTranslation();
  
  if (!models.valid) {
    return (
      <Card className="bg-slate-800 border-slate-700">
        <CardHeader>
          <CardTitle className="text-red-400">{t('cycling.model_parameters.title')}</CardTitle>
        </CardHeader>
        <CardContent>
          <p className="text-slate-400">{t('common.invalid_data')}</p>
        </CardContent>
      </Card>
    );
  }

  // Calculate P50min using Power Law: P = S × t^(E-1)
  const p50min = models.pl_params.S * Math.pow(3000, models.pl_params.E - 1); // 50 min = 3000 seconds
  
  // Calculate VO2max: P5min = CP + (W'/300), then VO2max = 7.44 × (P5min ÷ Weight) + 27.51
  const p5min = models.cp + (models.wPrime / 300);
  const vo2max = 7.44 * (p5min / bodyWeight) + 27.51;

  // Calculate LT1 range
  const lt1Min = Math.round(models.lt1_range.min);
  const lt1Max = Math.round(models.lt1_range.max);

  return (
    <TooltipProvider>
      <Card className="bg-slate-800 border-slate-700">
        <CardHeader>
          <CardTitle className="text-lime-400">{t('cycling.model_parameters.title')}</CardTitle>
        </CardHeader>
        <CardContent>
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 mobile-grid-to-stack">
            {/* CP Model */}
            <div className="bg-slate-900 p-4 rounded-lg border border-red-500/30">
              <h4 className="font-semibold text-red-400 mb-3 flex items-center">
                <div className="w-3 h-3 bg-red-500 rounded mr-2"></div>
                {t('cycling.model_parameters.hyperbolic_model')}
              </h4>
              <div className="space-y-2 text-sm">
                <div className="flex justify-between">
                  <span className="text-slate-400">CP:</span>
                  <span className="text-white font-medium">{Math.round(models.cp)} W</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-400">W':</span>
                  <span className="text-white font-medium">{Math.round(models.wPrime)} J</span>
                </div>
                <p className="text-xs text-slate-500 mt-3 pt-3 border-t border-slate-700">
                  {t('cycling.model_parameters.hyperbolic_model_description')}
                </p>
              </div>
            </div>

            {/* APR Model */}
            <div className="bg-slate-900 p-4 rounded-lg border border-purple-500/30">
              <h4 className="font-semibold text-purple-400 mb-3 flex items-center">
                <div className="w-3 h-3 bg-purple-500 rounded mr-2"></div>
                {t('cycling.model_parameters.apr_model')}
              </h4>
              <div className="space-y-2 text-sm">
                <div className="flex justify-between">
                  <span className="text-slate-400">P3min:</span>
                  <span className="text-white font-medium">{Math.round(models.apr_params.po3min)} W</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-400">APR:</span>
                  <span className="text-white font-medium">{Math.round(models.apr_params.apr)} W</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-400">k:</span>
                  <span className="text-white font-medium">{models.apr_params.k.toFixed(3)}</span>
                </div>
                <p className="text-xs text-slate-500 mt-3 pt-3 border-t border-slate-700">
                  {t('cycling.model_parameters.apr_model_description')}
                </p>
              </div>
            </div>

            {/* Power Law Model */}
            <div className="bg-slate-900 p-4 rounded-lg border border-yellow-500/30">
              <h4 className="font-semibold text-yellow-400 mb-3 flex items-center">
                <div className="w-3 h-3 bg-yellow-500 rounded mr-2"></div>
                {t('cycling.model_parameters.power_law_model')}
              </h4>
              <div className="space-y-2 text-sm">
                <div className="flex justify-between">
                  <span className="text-slate-400">{t('cycling.model_parameters.s_parameter')}:</span>
                  <span className="text-white font-medium">{Math.round(models.pl_params.S)} W</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-400">{t('cycling.model_parameters.e_parameter')}:</span>
                  <span className="text-white font-medium">{models.pl_params.E.toFixed(6)}</span>
                </div>
                <p className="text-xs text-slate-500 mt-2">
                  {t('cycling.model_parameters.formula_note')}
                </p>
                <p className="text-xs text-slate-500 mt-3 pt-3 border-t border-slate-700">
                  {t('cycling.model_parameters.power_law_model_description')}
                </p>
              </div>
            </div>
          </div>

          {/* Redesigned Thresholds Summary */}
          <div className="mt-6 bg-slate-900 p-6 rounded-lg">
            <h4 className="font-semibold text-lime-400 mb-6 text-lg">{t('cycling.model_parameters.physiological_thresholds')}</h4>
            
            {/* Horizontal Layout with Vertical Text/Value Structure */}
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-6 mobile-grid-to-stack">
              
              {/* Pmax */}
              <div className="bg-slate-800 p-4 rounded text-center">
                <div className="flex justify-center items-center gap-2 mb-2">
                  <span className="text-slate-300 text-sm">Pmax:</span>
                  <Tooltip>
                    <TooltipTrigger asChild>
                      <Info className="w-4 h-4 text-slate-400 cursor-help hover:text-lime-400 transition-colors" />
                    </TooltipTrigger>
                    <TooltipContent className="max-w-xs bg-slate-800 border-slate-600 text-slate-200">
                      <div>
                        <h6 className="font-semibold text-lime-400 mb-1">{t('cycling.model_parameters.neuromuscular_power')}</h6>
                        <p className="text-sm">{t('cycling.model_parameters.pmax_tooltip')}</p>
                      </div>
                    </TooltipContent>
                  </Tooltip>
                </div>
                <div className="font-bold text-white text-lg">{models.pMax} W</div>
              </div>

              {/* MAP */}
              <div className="bg-slate-800 p-4 rounded text-center">
                <div className="flex justify-center items-center gap-2 mb-2">
                  <span className="text-slate-300 text-sm">MAP:</span>
                  <Tooltip>
                    <TooltipTrigger asChild>
                      <Info className="w-4 h-4 text-slate-400 cursor-help hover:text-lime-400 transition-colors" />
                    </TooltipTrigger>
                    <TooltipContent className="max-w-xs bg-slate-800 border-slate-600 text-slate-200">
                      <div>
                        <h6 className="font-semibold text-lime-400 mb-1">{t('cycling.model_parameters.aerobic_power')}</h6>
                        <p className="text-sm">{t('cycling.model_parameters.map_tooltip')}</p>
                      </div>
                    </TooltipContent>
                  </Tooltip>
                </div>
                <div className="font-bold text-white text-lg">{Math.round(models.apr_params.po3min)} W</div>
              </div>

              {/* MMSS */}
              <div className="bg-slate-800 p-4 rounded text-center">
                <div className="flex justify-center items-center gap-2 mb-2">
                  <span className="text-slate-300 text-sm">MMSS:</span>
                  <Tooltip>
                    <TooltipTrigger asChild>
                      <Info className="w-4 h-4 text-slate-400 cursor-help hover:text-lime-400 transition-colors" />
                    </TooltipTrigger>
                    <TooltipContent className="max-w-sm bg-slate-800 border-slate-600 text-slate-200">
                      <div>
                        <h6 className="font-semibold text-lime-400 mb-1">{t('cycling.model_parameters.mmss_boundary')}</h6>
                        <p className="text-sm">{t('cycling.model_parameters.mmss_tooltip', { cp: Math.round(models.cp) })}</p>
                      </div>
                    </TooltipContent>
                  </Tooltip>
                </div>
                <div className="font-bold text-lime-400 text-lg">
                  {Math.round(p50min)} - {Math.round(models.cp)} W
                </div>
              </div>

              {/* LT1 */}
              <div className="bg-slate-800 p-4 rounded text-center">
                <div className="flex justify-center items-center gap-2 mb-2">
                  <span className="text-slate-300 text-sm">LT1:</span>
                  <Tooltip>
                    <TooltipTrigger asChild>
                      <Info className="w-4 h-4 text-slate-400 cursor-help hover:text-lime-400 transition-colors" />
                    </TooltipTrigger>
                    <TooltipContent className="max-w-sm bg-slate-800 border-slate-600 text-slate-200">
                      <div>
                        <h6 className="font-semibold text-lime-400 mb-1">{t('cycling.model_parameters.lt1_boundary')}</h6>
                        <p className="text-sm">{t('cycling.model_parameters.lt1_tooltip')}</p>
                        <p className="text-xs text-slate-400 mt-2">{t('cycling.model_parameters.lt1_note')}</p>
                      </div>
                    </TooltipContent>
                  </Tooltip>
                </div>
                <div className="font-bold text-white text-lg">{lt1Min} - {lt1Max} W</div>
              </div>

              {/* VO2max */}
              <div className="bg-slate-800 p-4 rounded text-center">
                <div className="flex justify-center items-center gap-2 mb-2">
                  <span className="text-slate-300 text-sm">VO2max:</span>
                  <Tooltip>
                    <TooltipTrigger asChild>
                      <Info className="w-4 h-4 text-slate-400 cursor-help hover:text-lime-400 transition-colors" />
                    </TooltipTrigger>
                    <TooltipContent className="max-w-xs bg-slate-800 border-slate-600 text-slate-200">
                      <div>
                        <h6 className="font-semibold text-lime-400 mb-1">{t('cycling.model_parameters.vo2max_consumption')}</h6>
                        <p className="text-sm">{t('cycling.model_parameters.vo2max_tooltip')}</p>
                        <p className="text-xs text-slate-400 mt-2">{t('cycling.model_parameters.vo2max_note')}</p>
                      </div>
                    </TooltipContent>
                  </Tooltip>
                </div>
                <div className="font-bold text-lime-400 text-lg">{vo2max.toFixed(1)} ml/kg/min</div>
              </div>

            </div>
          </div>

          {/* Debug Info - only show in development */}
          {process.env.NODE_ENV === 'development' && (
            <div className="mt-4 p-3 bg-slate-950 rounded border border-slate-600">
              <h5 className="text-xs text-slate-400 mb-2">Debug Info (Alta Precisione)</h5>
              <div className="text-xs text-slate-500 space-y-1">
                <div>CP: {models.cp} W</div>
                <div>W': {models.wPrime} J</div>
                <div>S: {models.pl_params.S} W</div>
                <div>E: {models.pl_params.E}</div>
                <div>P50min: {p50min} W</div>
                <div>P5min: {p5min} W</div>
                <div>VO2max: {vo2max.toFixed(2)} ml/kg/min</div>
              </div>
            </div>
          )}
        </CardContent>
      </Card>
    </TooltipProvider>
  );
};

export default ModelParameters;

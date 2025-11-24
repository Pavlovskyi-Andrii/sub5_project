import React from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Tooltip, TooltipContent, TooltipProvider, TooltipTrigger } from '@/components/ui/tooltip';
import { Info } from 'lucide-react';
import { PrecisionModels } from '../utils/precisionCalculations';
import { calculateRunningAthleteProfile } from '../utils/runningAthleteCalculations';
import SpiderChart from './SpiderChart';
import { getProfileTheme } from '../utils/profileThemes';
import { getRarityLevel } from '../utils/raritySystem';
import { useTranslation } from 'react-i18next';

interface RunningPhysiologicalProfileProps {
  models: PrecisionModels;
  bodyWeight: number;
  gender: 'male' | 'female';
}

const RunningPhysiologicalProfile: React.FC<RunningPhysiologicalProfileProps> = ({ models, bodyWeight, gender }) => {
  const { t } = useTranslation();
  
  // Return early if models are not valid
  if (!models.valid) {
    return (
      <Card className="bg-slate-800 border-slate-700">
        <CardHeader>
          <CardTitle className="text-red-400 text-xl">
            {t('running.physiological_profile.error_title')}
          </CardTitle>
          <p className="text-slate-400">
            {t('running.physiological_profile.error_message')}
          </p>
        </CardHeader>
      </Card>
    );
  }

  // Calculate athlete profile with scores for spider chart
  const athleteProfile = calculateRunningAthleteProfile(models, gender, bodyWeight);
  const speedScore1500 = athleteProfile.stats['1500m']; // Speed/VO2 score
  const enduranceScore10k = athleteProfile.stats['10km']; // Endurance score
  
  // Get theme and rarity level for SpiderChart
  const theme = getProfileTheme(athleteProfile.profileType);
  const rarityLevel = getRarityLevel(athleteProfile.overallRating);
  
  // TASK 1: Indice di Resistenza alla Fatica (usando parametro k dal Power Law)
  const k = models.pl_params.k; // k = E - 1, dove E è l'esponente di endurance
  const fatigueIndex = (1 - Math.pow(2, k)) * 100;
  
  // Color and profile based on fatigue index
  const getFatigueColor = (index: number) => {
    if (index < 7.5) return '#4ADE80'; // Verde
    if (index <= 9.5) return '#FACC15'; // Giallo
    return '#F87171'; // Rosso
  };
  
  const getFatigueProfile = (index: number) => {
    if (index < 7.5) return t('running.physiological_profile.profiles.resistant');
    if (index <= 9.5) return t('running.physiological_profile.profiles.mixed');
    return t('running.physiological_profile.profiles.fast');
  };

  const fatigueColor = getFatigueColor(fatigueIndex);
  const fatigueProfile = getFatigueProfile(fatigueIndex);

  // TASK 2: Frazione di Utilizzo (usando velocità invece di potenza)
  const velocityAt5min = models.cs + (models.dPrime / 300); // Velocità a 5 min dal modello CS-D'
  const criticalSpeed = models.cs;
  const utilizationFraction = (criticalSpeed / velocityAt5min) * 100;
  
  // Delta_Profilo = punteggio_1500m - punteggio_10km (velocità vs endurance)
  const deltaProfilo = speedScore1500 - enduranceScore10k;
  
  // Analisi fisiologica personalizzata (senza consigli allenamento)
  const getPhysiologicalAnalysis = (delta: number) => {
    if (delta < -3) {
      return t('running.physiological_profile.endurance_oriented', { percentage: utilizationFraction.toFixed(1) });
    } else {
      return t('running.physiological_profile.speed_oriented', { percentage: utilizationFraction.toFixed(1) });
    }
  };

  const physiologicalAnalysis = getPhysiologicalAnalysis(deltaProfilo);

  return (
    <TooltipProvider>
      <Card className="bg-slate-800 border-slate-700">
        <CardHeader>
          <CardTitle className="text-lime-400 text-xl">
            {t('running.physiological_profile.title')}
          </CardTitle>
          <p className="text-slate-400">
            {t('running.physiological_profile.subtitle')}
          </p>
        </CardHeader>
        <CardContent>
          {/* Responsive Layout Container */}
          <div className="flex flex-col md:flex-row gap-8">
            
            {/* Left Column (Desktop) / Top Section (Mobile): Spider Chart */}
            <div className="flex-none md:flex-[4] flex justify-center items-center">
              <SpiderChart 
                stats={athleteProfile.stats}
                theme={theme}
                rarityLevel={rarityLevel}
                isRunning={true}
              />
            </div>
            
            {/* Right Column (Desktop) / Bottom Section (Mobile): Text Content */}
            <div className="flex-1 md:flex-[6] space-y-8">
              
              {/* SEZIONE 1: Indice di Resistenza alla Fatica - Layout Gerarchico */}
              <div className="space-y-4">
                <div className="flex items-center gap-2">
                  <h3 className="text-lg font-semibold text-slate-200">
                    {t('running.physiological_profile.fatigue_resistance')}
                  </h3>
                  <Tooltip>
                    <TooltipTrigger>
                      <Info className="h-4 w-4 text-slate-400 hover:text-slate-300" />
                    </TooltipTrigger>
                    <TooltipContent className="max-w-sm">
                      <div className="space-y-2">
                        <h4 className="font-semibold text-sm">{t('running.physiological_profile.fatigue_tooltip.title')}</h4>
                        <p className="text-xs">
                          {t('running.physiological_profile.fatigue_tooltip.description')}
                        </p>
                        <div className="space-y-1 text-xs">
                          <p className="font-medium">{t('running.physiological_profile.fatigue_tooltip.classification')}</p>
                          <p>{t('running.physiological_profile.fatigue_tooltip.resistant')}</p>
                          <p>{t('running.physiological_profile.fatigue_tooltip.mixed')}</p>
                          <p>{t('running.physiological_profile.fatigue_tooltip.fast')}</p>
                        </div>
                        <p className="text-xs italic text-slate-400">
                          {t('running.physiological_profile.fatigue_tooltip.note')}
                        </p>
                      </div>
                    </TooltipContent>
                  </Tooltip>
                </div>
                
                {/* Profilo come elemento principale */}
                <div className="text-center">
                  <div 
                    className="text-2xl font-bold mb-2" 
                    style={{ color: fatigueColor }}
                  >
                    {fatigueProfile}
                  </div>
                  <div className="text-slate-400 text-sm">
                    {t('running.physiological_profile.fatigue_decay', { percentage: fatigueIndex.toFixed(1) })}
                  </div>
                </div>
              </div>

              {/* Linea separatrice sottile */}
              <div className="border-t border-slate-600"></div>

              {/* SEZIONE 2: Analisi Fisiologica */}
              <div className="bg-slate-700 border border-slate-600 rounded-lg p-6 space-y-4">
                <div className="flex items-center gap-3 mb-4">
                  <Info className="h-6 w-6 text-lime-400" />
                  <h4 className="text-lg font-semibold text-slate-200">
                    {t('running.physiological_profile.phenotype_analysis')}
                  </h4>
                </div>
                
                {/* Diagnosi fisiologica */}
                <p className="text-slate-300 leading-relaxed">
                  {physiologicalAnalysis}
                </p>
                
                {/* Informazioni aggiuntive sulla frazione di utilizzo */}
                <div className="mt-4 p-4 bg-slate-800 rounded-lg">
                  <div className="text-sm text-slate-400 mb-2">
                    <strong>{t('running.physiological_profile.utilization_fraction', { percentage: utilizationFraction.toFixed(1) })}</strong>
                  </div>
                  <div className="text-xs text-slate-500">
                    {t('running.physiological_profile.utilization_description')}
                  </div>
                </div>
              </div>
              
            </div>
          </div>
        </CardContent>
      </Card>
    </TooltipProvider>
  );
};

export default RunningPhysiologicalProfile;
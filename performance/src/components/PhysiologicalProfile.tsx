import React from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Tooltip, TooltipContent, TooltipProvider, TooltipTrigger } from '@/components/ui/tooltip';
import { Info, Lightbulb, CheckCircle, Award, Shield, Zap, Activity, Target, Dumbbell } from 'lucide-react';
import { PowerPrecisionModels } from '../utils/powerPrecisionCalculations';
import { calculateAthleteProfile } from '../utils/athleteCardCalculations';
import SpiderChart from './SpiderChart';
import { getProfileTheme } from '../utils/profileThemes';
import { getRarityLevel } from '../utils/raritySystem';
import { useTranslation } from 'react-i18next';

interface PhysiologicalProfileProps {
  models: PowerPrecisionModels;
  bodyWeight: number;
  gender: 'male' | 'female';
}

const PhysiologicalProfile: React.FC<PhysiologicalProfileProps> = ({ models, bodyWeight, gender }) => {
  const { t } = useTranslation();
  
  // Return early if models are not valid
  if (!models.valid) {
    return (
      <Card className="bg-slate-800 border-slate-700">
        <CardHeader>
          <CardTitle className="text-red-400 text-xl">
            {t('cycling.physiological_profile.error_title')}
          </CardTitle>
          <p className="text-slate-400">
            {t('cycling.physiological_profile.error_message')}
          </p>
        </CardHeader>
      </Card>
    );
  }

  // Calculate athlete profile with scores for spider chart
  const athleteProfile = calculateAthleteProfile(models, bodyWeight, gender);
  const vo2Score = athleteProfile.stats.VO2; // 5 min score
  const staScore = athleteProfile.stats.STA; // 60 min score
  
  // Get theme and rarity level for SpiderChart
  const theme = getProfileTheme(athleteProfile.profileType);
  const rarityLevel = getRarityLevel(athleteProfile.overallRating);
  
  // TASK 1: Indice di Resistenza alla Fatica
  const E = models.pl_params.E;
  const fatigueIndex = (1 - Math.pow(2, E - 1)) * 100;
  
  // Color and profile based on fatigue index
  const getFatigueColor = (index: number) => {
    if (index < 7.5) return '#4ADE80'; // Verde
    if (index <= 9.5) return '#FACC15'; // Giallo
    return '#F87171'; // Rosso
  };
  
  const getFatigueProfile = (index: number) => {
    if (index < 7.5) return t('cycling.physiological_profile.profiles.resistant');
    if (index <= 9.5) return t('cycling.physiological_profile.profiles.mixed');
    return t('cycling.physiological_profile.profiles.fast');
  };

  const fatigueColor = getFatigueColor(fatigueIndex);
  const fatigueProfile = getFatigueProfile(fatigueIndex);

  // TASK 2: Frazione di Utilizzo e Consigli
  const powerAt5min = models.cp + (models.wPrime / 300); // Potenza a 5 min dal modello CP-W'
  const criticalPower = models.cp;
  const utilizationFraction = (criticalPower / powerAt5min) * 100;
  
  // Delta_Profilo = punteggio_STA - punteggio_VO2
  const deltaProfilo = staScore - vo2Score;
  
  // Consigli personalizzati basati su Delta_Profilo
  const getTrainingAdvice = (delta: number) => {
    if (delta < -5) {
      return {
        physiologicalProfile: t('cycling.physiological_profile.speed_oriented', { percentage: utilizationFraction.toFixed(1) }),
        trainingAdvice: t('cycling.physiological_profile.hiit_description_short')
      };
    } else {
      return {
        physiologicalProfile: t('cycling.physiological_profile.endurance_oriented', { percentage: utilizationFraction.toFixed(1) }),
        trainingAdvice: t('cycling.physiological_profile.hiit_description_long')
      };
    }
  };

  const advice = getTrainingAdvice(deltaProfilo);

  // Get profile icon based on type
  const getProfileIcon = (profileType: string) => {
    switch (profileType.toUpperCase()) {
      case 'SPRINTER': return Zap;
      case 'CLIMBER': return Activity;
      case 'TIME TRIALIST': return Target;
      case 'PUNCHEUR': return Dumbbell;
      default: return Award;
    }
  };

  const ProfileIcon = getProfileIcon(athleteProfile.profileType);

  return (
    <TooltipProvider>
      <Card className="bg-gradient-to-br from-slate-800 to-slate-900 border-slate-700 shadow-2xl">
        <CardHeader>
          <CardTitle className="text-lime-400 text-xl">
            {t('cycling.physiological_profile.title')}
          </CardTitle>
          <p className="text-slate-400">
            {t('cycling.physiological_profile.subtitle')}
          </p>
        </CardHeader>
        <CardContent>
          {/* Responsive Layout Container */}
          <div className="flex flex-col md:flex-row gap-8">
            
            {/* Left Column (Desktop) / Top Section (Mobile): Spider Chart with Category */}
            <div className="flex-none md:flex-[5] flex flex-col justify-center items-center space-y-4">
              {/* Category Title Above Chart */}
              <div className="flex items-center gap-3">
                <div className="p-2 rounded-lg bg-slate-700/60">
                  <ProfileIcon className="h-6 w-6" style={{ color: theme.primaryColor }} />
                </div>
                <h2 className="text-xl font-bold text-slate-200">
                  {(() => {
                    const profileKey = athleteProfile.profileType.toLowerCase().replace(/[\s-]/g, '_');
                    return t(`athlete_card.profiles.cycling.${profileKey}`, athleteProfile.profileType);
                  })()}
                </h2>
              </div>
              
              {/* Clean Spider Chart */}
              <SpiderChart 
                stats={athleteProfile.stats}
                theme={theme}
                rarityLevel={rarityLevel}
                isRunning={false}
                profileType={athleteProfile.profileType}
              />
            </div>
            
            {/* Right Column (Desktop) / Bottom Section (Mobile): Text Content */}
            <div className="flex-1 md:flex-[5] space-y-5">

              {/* SEZIONE 1: Resistenza alla Fatica */}
              <div className="bg-slate-800/60 border border-slate-700 rounded-xl p-5 animate-fade-in hover:bg-slate-800/80 transition-all duration-300"
                   style={{ borderLeftColor: fatigueColor, borderLeftWidth: '3px' }}>
                <div className="flex items-center gap-3 mb-4">
                  <div className="p-2 rounded-lg bg-slate-700/60">
                    {fatigueIndex < 7.5 ? (
                      <Shield className="h-5 w-5" style={{ color: fatigueColor }} />
                    ) : fatigueIndex <= 9.5 ? (
                      <Activity className="h-5 w-5" style={{ color: fatigueColor }} />
                    ) : (
                      <Zap className="h-5 w-5" style={{ color: fatigueColor }} />
                    )}
                  </div>
                  <h3 className="text-lg font-semibold text-slate-200">
                    {t('cycling.physiological_profile.fatigue_resistance')}
                  </h3>
                  <Tooltip>
                    <TooltipTrigger>
                      <Info className="h-4 w-4 text-slate-400 hover:text-slate-300" />
                    </TooltipTrigger>
                    <TooltipContent className="max-w-sm">
                      <div className="space-y-2">
                        <h4 className="font-semibold text-sm">{t('cycling.physiological_profile.fatigue_tooltip.title')}</h4>
                        <p className="text-xs">
                          {t('cycling.physiological_profile.fatigue_tooltip.description')}
                        </p>
                        <div className="space-y-1 text-xs">
                          <p className="font-medium">{t('cycling.physiological_profile.fatigue_tooltip.classification')}</p>
                          <p>{t('cycling.physiological_profile.fatigue_tooltip.resistant')}</p>
                          <p>{t('cycling.physiological_profile.fatigue_tooltip.mixed')}</p>
                          <p>{t('cycling.physiological_profile.fatigue_tooltip.fast')}</p>
                        </div>
                        <p className="text-xs italic text-slate-400">
                          {t('cycling.physiological_profile.fatigue_tooltip.note')}
                        </p>
                      </div>
                    </TooltipContent>
                  </Tooltip>
                </div>
                 
                <div className="space-y-3">
                  <div className="text-center">
                    <div 
                      className="text-xl font-bold mb-1" 
                      style={{ color: fatigueColor }}
                    >
                      {fatigueProfile}
                    </div>
                    <div className="text-slate-400 text-sm">
                      {t('cycling.physiological_profile.fatigue_decay', { percentage: fatigueIndex.toFixed(1) })}
                    </div>
                  </div>
                  
                  {/* Progress Bar */}
                  <div className="relative">
                    <div className="h-2 bg-slate-700 rounded-full overflow-hidden">
                      <div 
                        className="h-full rounded-full animate-progress-fill transition-all duration-1000 ease-out"
                        style={{ 
                          backgroundColor: fatigueColor,
                          width: `${Math.min(fatigueIndex * 8, 100)}%`
                        }}
                      />
                    </div>
                  </div>
                </div>
              </div>

              {/* SEZIONE 2: Analisi VO2max/Soglia */}
              <div className="bg-slate-800/60 border border-slate-700 rounded-xl p-5 animate-fade-in delay-150 hover:bg-slate-800/80 transition-all duration-300"
                   style={{ borderLeftColor: theme.primaryColor, borderLeftWidth: '3px' }}>
                <div className="flex items-center gap-3 mb-4">
                  <div className="p-2 rounded-lg bg-slate-700/60">
                    <Lightbulb className="h-5 w-5" style={{ color: theme.primaryColor }} />
                  </div>
                  <h3 className="text-lg font-semibold text-slate-200">
                    {t('cycling.physiological_profile.vo2_threshold_analysis')}
                  </h3>
                </div>
                
                <p className="text-slate-300 leading-relaxed">
                  {advice.physiologicalProfile}
                </p>
              </div>

              {/* SEZIONE 3: Allenamenti Raccomandati */}
              <div className="bg-slate-800/60 border border-slate-700 rounded-xl p-5 animate-fade-in delay-300 hover:bg-slate-800/80 transition-all duration-300"
                   style={{ borderLeftColor: '#4ADE80', borderLeftWidth: '3px' }}>
                <div className="flex items-center gap-3 mb-4">
                  <div className="p-2 rounded-lg bg-slate-700/60">
                    <Dumbbell className="h-5 w-5 text-emerald-400" />
                  </div>
                  <h3 className="text-lg font-semibold text-slate-200">
                    {t('cycling.physiological_profile.recommended_training')}
                  </h3>
                </div>
                
                <div className="space-y-4">
                  <p className="text-slate-300 leading-relaxed">
                    {advice.trainingAdvice}
                  </p>
                  
                  {/* Nota sui punti deboli */}
                   <div className="bg-orange-500/10 border border-orange-400/30 rounded-lg p-4">
                    <div className="flex items-center gap-3 mb-2">
                      <Target className="h-4 w-4 text-orange-400" />
                      <span className="text-orange-400 font-semibold">
                        Nota Importante
                      </span>
                    </div>
                    <p className="text-slate-300 leading-relaxed">
                      <span className="font-bold">Focus sui punti deboli:</span> I tuoi allenamenti dovrebbero concentrarsi sui punti dove il tuo profilo mostra margini di miglioramento per ottenere i risultati migliori.
                    </p>
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

export default PhysiologicalProfile;
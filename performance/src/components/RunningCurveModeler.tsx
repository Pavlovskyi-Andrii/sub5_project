
import React, { useState, useEffect, useCallback } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Slider } from '@/components/ui/slider';
import { Switch } from '@/components/ui/switch';
import { Label } from '@/components/ui/label';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { useTranslation } from 'react-i18next';
import RunningChart from './RunningChart';
import RunningModelParameters from './RunningModelParameters';
import RunningZonesTable from './RunningZonesTable';
import RunningPhysiologicalProfile from './RunningPhysiologicalProfile';
import PerformanceCalculator from './PerformanceCalculator';
import RunningSustainableTable from './RunningSustainableTable';
import PdfLeadForm from './PdfLeadForm';
import ConsultationModal from './ConsultationModal';
import RunningAthleteCardGenerator from './RunningAthleteCardGenerator';
import { calculateHighPrecisionModels, PrecisionModels, formatForDisplay } from '../utils/precisionCalculations';
import { PowerPrecisionModels } from '../utils/powerPrecisionCalculations';

const RunningCurveModeler = () => {
  const { t } = useTranslation();
  const [severeDistance, setSevereDistance] = useState(1200);
  const [severeMinutes, setSevereMinutes] = useState(3);
  const [severeSeconds, setSevereSeconds] = useState(30);
  const [thresholdDistance, setThresholdDistance] = useState(4000);
  const [thresholdMinutes, setThresholdMinutes] = useState(15);
  const [thresholdSeconds, setThresholdSeconds] = useState(0);
  const [enableLongTest, setEnableLongTest] = useState(false);
  const [longDistance, setLongDistance] = useState('10');
  const [longHours, setLongHours] = useState(0);
  const [longMinutes, setLongMinutes] = useState(40);
  const [longSeconds, setLongSeconds] = useState(0);
  const [showConsultationModal, setShowConsultationModal] = useState(false);
  const [bodyWeight, setBodyWeight] = useState(70);
  const [gender, setGender] = useState<'male' | 'female'>('male');

  const [currentModels, setCurrentModels] = useState<PrecisionModels>({ 
    valid: false, 
    cs: 0, dPrime: 0, cs_pace: 0, 
    pl_params: { A: 0, k: 0 }, 
    lt1_range: { min: 0, max: 0, estimate: 0 }, 
    vo2max_pace: 0
  });

  const calculateModels = useCallback((): PrecisionModels => {
    const severeTime = severeMinutes * 60 + severeSeconds;
    const thresholdTime = thresholdMinutes * 60 + thresholdSeconds;
    const longTime = longHours * 3600 + longMinutes * 60 + longSeconds;
    
    return calculateHighPrecisionModels(
      severeDistance,
      severeTime,
      thresholdDistance,
      thresholdTime,
      enableLongTest,
      longDistance,
      longTime
    );
  }, [
    severeDistance, severeMinutes, severeSeconds,
    thresholdDistance, thresholdMinutes, thresholdSeconds,
    enableLongTest, longDistance, longHours, longMinutes, longSeconds
  ]);

  useEffect(() => {
    setCurrentModels(calculateModels());
  }, [calculateModels]);

  // Convert PrecisionModels to PowerPrecisionModels format for ConsultationModal
  const modelsForConsultation: PowerPrecisionModels = {
    valid: currentModels.valid,
    reason: currentModels.reason,
    cp: currentModels.cs, // Use CS as CP equivalent
    wPrime: currentModels.dPrime, // Use D' as W' equivalent
    pMax: 0, // Not applicable for running
    cp_power: currentModels.cs, // Use CS
    apr_params: { po3min: 0, apr: 0, k: currentModels.pl_params.k },
    pl_params: { S: currentModels.pl_params.A, E: currentModels.pl_params.k + 1 },
    lt1_range: currentModels.lt1_range,
    vo2max_power: currentModels.vo2max_pace // Use pace as power equivalent
  };

  return (
    <div className="min-h-screen bg-slate-900 text-white p-4">
      <div className="max-w-7xl mx-auto space-y-6">
        {/* Header Principale con CTA */}
        <div className="text-center mb-8">
          <h1 className="text-4xl font-bold text-lime-400 mb-3">
            {t('app_title')}
          </h1>
          <p className="text-slate-400 text-lg">
            {t('running.lab_subtitle')}
          </p>
          
          {/* Blocco CTA Ottimizzato */}
          <div className="mt-6 flex flex-col items-center gap-3">
            {/* Riga dei pulsanti perfettamente allineati */}
            <div className="flex flex-col sm:flex-row items-center gap-4 mobile-cta-buttons">
              <PdfLeadForm models={currentModels} isRunning={true} weight={bodyWeight} gender={gender} />
              <ConsultationModal 
                models={modelsForConsultation}
                isOpen={showConsultationModal}
                onOpenChange={setShowConsultationModal}
                isRunning={true}
                weight={bodyWeight}
                gender={gender}
              />
              <RunningAthleteCardGenerator models={currentModels} />
            </div>
            
            {/* Testo di supporto centrato sotto entrambi i pulsanti */}
            <p className="text-slate-400 text-sm max-w-md text-center">
              {t('running.cta_description')}
            </p>
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mobile-stack">
          {/* Controls Column */}
          <div className="lg:col-span-1 space-y-4 order-2 lg:order-1">
            {/* Severe Domain */}
            <Card className="bg-slate-800 border-slate-700">
              <CardHeader className="pb-3">
                <CardTitle className="text-lime-400 text-lg">{t('running.controls.severe_test_title')}</CardTitle>
                <p className="text-slate-400 text-sm">{t('running.controls.severe_test_description')}</p>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="space-y-3">
                  <div className="flex justify-between items-center">
                    <Label className="text-slate-200 font-medium">{t('running.controls.distance')}</Label>
                    <span className="font-semibold bg-slate-700 text-lime-400 px-3 py-1 rounded">
                      {severeDistance} m
                    </span>
                  </div>
                  <Slider
                    value={[severeDistance]}
                    onValueChange={(value) => setSevereDistance(value[0])}
                    min={800}
                    max={2000}
                    step={50}
                    className="w-full"
                  />
                </div>
                <div className="grid grid-cols-2 gap-3">
                  <div className="space-y-2">
                    <Label className="text-slate-200 font-medium">{t('running.controls.minutes')}</Label>
                    <Slider
                      value={[severeMinutes]}
                      onValueChange={(value) => setSevereMinutes(value[0])}
                      min={2}
                      max={5}
                      step={1}
                      className="w-full"
                    />
                    <span className="text-lime-400 text-sm">{severeMinutes}m</span>
                  </div>
                  <div className="space-y-2">
                    <Label className="text-slate-200 font-medium">{t('running.controls.seconds')}</Label>
                    <Slider
                      value={[severeSeconds]}
                      onValueChange={(value) => setSevereSeconds(value[0])}
                      min={0}
                      max={59}
                      step={1}
                      className="w-full"
                    />
                    <span className="text-lime-400 text-sm">{severeSeconds}s</span>
                  </div>
                </div>
                <div className="text-center">
                  <span className="font-semibold bg-slate-700 text-lime-400 px-3 py-1 rounded">
                    {t('running.controls.total_time')}: {formatForDisplay.time(severeMinutes * 60 + severeSeconds)}
                  </span>
                </div>
              </CardContent>
            </Card>

            {/* Threshold Domain */}
            <Card className="bg-slate-800 border-slate-700">
              <CardHeader className="pb-3">
                <CardTitle className="text-lime-400 text-lg">Sforzo Media Durata (5-16 min)</CardTitle>
                <p className="text-slate-400 text-sm">Test per stima Critical Speed. Durata ideale: 12-16 min</p>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="space-y-3">
                  <div className="flex justify-between items-center">
                    <Label className="text-slate-200 font-medium">Distanza</Label>
                    <span className="font-semibold bg-slate-700 text-lime-400 px-3 py-1 rounded">
                      {thresholdDistance} m
                    </span>
                  </div>
                  <Slider
                    value={[thresholdDistance]}
                    onValueChange={(value) => setThresholdDistance(value[0])}
                    min={3000}
                    max={5000}
                    step={100}
                    className="w-full"
                  />
                </div>
                <div className="grid grid-cols-2 gap-3">
                  <div className="space-y-2">
                    <Label className="text-slate-200 font-medium">Minuti</Label>
                    <Slider
                      value={[thresholdMinutes]}
                      onValueChange={(value) => setThresholdMinutes(value[0])}
                      min={10}
                      max={20}
                      step={1}
                      className="w-full"
                    />
                    <span className="text-lime-400 text-sm">{thresholdMinutes}m</span>
                  </div>
                  <div className="space-y-2">
                    <Label className="text-slate-200 font-medium">Secondi</Label>
                    <Slider
                      value={[thresholdSeconds]}
                      onValueChange={(value) => setThresholdSeconds(value[0])}
                      min={0}
                      max={59}
                      step={1}
                      className="w-full"
                    />
                    <span className="text-lime-400 text-xs">{thresholdSeconds}s</span>
                  </div>
                </div>
                <div className="text-center">
                  <span className="font-semibold bg-slate-700 text-lime-400 px-3 py-1 rounded">
                    Tempo totale: {formatForDisplay.time(thresholdMinutes * 60 + thresholdSeconds)}
                  </span>
                </div>
              </CardContent>
            </Card>

            {/* Long Test */}
            <Card className="bg-slate-800 border-slate-700">
              <CardHeader className="pb-3">
                <div className="flex items-center space-x-3">
                  <Switch
                    checked={enableLongTest}
                    onCheckedChange={setEnableLongTest}
                  />
                  <CardTitle className="text-lime-400 text-lg">Sforzo Lunga Durata (&gt;16 min)</CardTitle>
                </div>
                <p className="text-slate-400 text-sm">Migliora la stima del Power Law per l'endurance.</p>
              </CardHeader>
              <CardContent className="space-y-4">
                <fieldset disabled={!enableLongTest} className="space-y-4 disabled:opacity-50">
                  <div className="space-y-3">
                    <Label className="text-slate-200 font-medium">Distanza</Label>
                    <Select value={longDistance} onValueChange={setLongDistance} disabled={!enableLongTest}>
                      <SelectTrigger className="bg-slate-700 border-slate-600">
                        <SelectValue />
                      </SelectTrigger>
                      <SelectContent className="bg-slate-700 border-slate-600">
                        <SelectItem value="10">10 km</SelectItem>
                        <SelectItem value="15">15 km</SelectItem>
                        <SelectItem value="21.097">21 km (Mezza Maratona)</SelectItem>
                        <SelectItem value="42.195">42 km (Maratona)</SelectItem>
                      </SelectContent>
                    </Select>
                  </div>
                  <div className="grid grid-cols-3 gap-2">
                    <div className="space-y-2">
                      <Label className="text-slate-200 font-medium text-sm">Ore</Label>
                      <Slider
                        value={[longHours]}
                        onValueChange={(value) => setLongHours(value[0])}
                        min={0}
                        max={5}
                        step={1}
                        className="w-full"
                        disabled={!enableLongTest}
                      />
                      <span className="text-lime-400 text-xs">{longHours}h</span>
                    </div>
                    <div className="space-y-2">
                      <Label className="text-slate-200 font-medium text-sm">Min</Label>
                      <Slider
                        value={[longMinutes]}
                        onValueChange={(value) => setLongMinutes(value[0])}
                        min={0}
                        max={59}
                        step={1}
                        className="w-full"
                        disabled={!enableLongTest}
                      />
                      <span className="text-lime-400 text-xs">{longMinutes}m</span>
                    </div>
                    <div className="space-y-2">
                      <Label className="text-slate-200 font-medium text-sm">Sec</Label>
                      <Slider
                        value={[longSeconds]}
                        onValueChange={(value) => setLongSeconds(value[0])}
                        min={0}
                        max={59}
                        step={1}
                        className="w-full"
                        disabled={!enableLongTest}
                      />
                      <span className="text-lime-400 text-xs">{longSeconds}s</span>
                    </div>
                  </div>
                  <div className="text-center">
                    <span className="font-semibold bg-slate-700 text-lime-400 px-3 py-1 rounded text-sm">
                      Tempo: {formatForDisplay.time(longHours * 3600 + longMinutes * 60 + longSeconds)}
                    </span>
                  </div>
                </fieldset>
              </CardContent>
            </Card>

            {/* Athlete Data */}
            <Card className="bg-slate-800 border-slate-700">
              <CardHeader className="pb-3">
                <CardTitle className="text-lime-400 text-lg">Dati Atleta</CardTitle>
                <p className="text-slate-400 text-sm">Per calcolo del profilo fisiologico</p>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="space-y-3">
                  <Label className="text-slate-200 font-medium">Genere</Label>
                  <Select value={gender} onValueChange={(value: 'male' | 'female') => setGender(value)}>
                    <SelectTrigger className="bg-slate-700 border-slate-600">
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent className="bg-slate-700 border-slate-600">
                      <SelectItem value="male">Uomo</SelectItem>
                      <SelectItem value="female">Donna</SelectItem>
                    </SelectContent>
                  </Select>
                </div>
                <div className="space-y-3">
                  <div className="flex justify-between items-center">
                    <Label className="text-slate-200 font-medium">Peso Corporeo</Label>
                    <span className="font-semibold bg-slate-700 text-lime-400 px-3 py-1 rounded">
                      {bodyWeight} kg
                    </span>
                  </div>
                  <Slider
                    value={[bodyWeight]}
                    onValueChange={(value) => setBodyWeight(value[0])}
                    min={40}
                    max={120}
                    step={1}
                    className="w-full"
                  />
                </div>
              </CardContent>
            </Card>
          </div>

          {/* Chart and Parameters Column */}
          <div className="lg:col-span-2 space-y-6 order-1 lg:order-2">
            {/* Model Parameters */}
            <RunningModelParameters models={currentModels} />
            
            {/* Chart */}
            <Card className="bg-slate-800 border-slate-700">
              <CardHeader>
                <CardTitle className={`text-xl ${currentModels.valid ? 'text-lime-400' : 'text-red-400'}`}>
                  {currentModels.valid ? 'Profilo Velocità-Durata' : 'Dati Non Validi'}
                </CardTitle>
                <p className="text-slate-400">
                  {currentModels.valid 
                    ? 'Curve passo-distanza stimate dai modelli fisiologici.' 
                    : currentModels.reason
                  }
                </p>
              </CardHeader>
              <CardContent>
                <div className="h-[400px] lg:h-[600px]" id="running-chart-container">
                  <RunningChart models={currentModels} />
                </div>
              </CardContent>
            </Card>
          </div>
        </div>

        {/* Running Zones Table - Full Width */}
        <div className="mt-8" id="running-zones-container">
          <RunningZonesTable models={currentModels} />
        </div>

        {/* Running Physiological Profile - Full Width */}
        <div className="mt-8" id="running-physiological-profile">
          <RunningPhysiologicalProfile 
            models={currentModels} 
            bodyWeight={bodyWeight}
            gender={gender}
          />
        </div>

        {/* Performance Calculator - Full Width */}
        <div className="mt-8">
          <PerformanceCalculator models={currentModels} />
        </div>
        
        {/* Sustainable Performances Table - Full Width */}
        <div className="mt-8">
          <RunningSustainableTable models={currentModels} />
        </div>
      </div>
    </div>
  );
};

export default RunningCurveModeler;

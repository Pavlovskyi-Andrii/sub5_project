import React, { useState, useEffect, useCallback } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Slider } from '@/components/ui/slider';
import { Switch } from '@/components/ui/switch';
import { Label } from '@/components/ui/label';
import { useTranslation } from 'react-i18next';
import PowerChart from './PowerChart';
import ModelParameters from './ModelParameters';
import PowerZonesTable from './PowerZonesTable';
import PowerPerformanceCalculator from './PowerPerformanceCalculator';
import PhysiologicalProfile from './PhysiologicalProfile';
import PowerSustainableTable from './PowerSustainableTable';
import PdfLeadForm from './PdfLeadForm';
import AthleteCardGenerator from './AthleteCardGenerator';
import ConsultationModal from './ConsultationModal';
import { calculateHighPrecisionPowerModels, PowerPrecisionModels, formatForDisplay } from '../utils/powerPrecisionCalculations';

const PowerCurveModeler = () => {
  const { t } = useTranslation();
  
  // State for all input values
  const [pmaxPower, setPmaxPower] = useState(1110);
  const [severePower, setSeverePower] = useState(397);
  const [severeTime, setSevereTime] = useState(240);
  const [thresholdPower, setThresholdPower] = useState(348);
  const [thresholdTime, setThresholdTime] = useState(902);
  const [shortPower, setShortPower] = useState(572);
  const [shortTime, setShortTime] = useState(60);
  const [longPower, setLongPower] = useState(297);
  const [longTime, setLongTime] = useState(3600);
  const [enableShortTest, setEnableShortTest] = useState(false);
  const [enableLongTest, setEnableLongTest] = useState(false);
  const [bodyWeight, setBodyWeight] = useState(70);
  const [gender, setGender] = useState<'male' | 'female'>('male');
  const [showConsultationModal, setShowConsultationModal] = useState(false);

  const [currentModels, setCurrentModels] = useState<PowerPrecisionModels>({ 
    valid: false, 
    cp: 0, wPrime: 0, pMax: 0, cp_power: 0,
    apr_params: { po3min: 0, apr: 0, k: 0 },
    pl_params: { S: 0, E: 0 },
    lt1_range: { min: 0, max: 0, estimate: 0 },
    vo2max_power: 0
  });

  const calculateModels = useCallback((): PowerPrecisionModels => {
    return calculateHighPrecisionPowerModels(
      pmaxPower,
      severePower,
      severeTime,
      thresholdPower,
      thresholdTime,
      enableShortTest,
      shortPower,
      shortTime,
      enableLongTest,
      longPower,
      longTime
    );
  }, [
    pmaxPower, severePower, severeTime, thresholdPower, thresholdTime,
    shortPower, shortTime, longPower, longTime, enableShortTest, enableLongTest
  ]);

  useEffect(() => {
    setCurrentModels(calculateModels());
  }, [calculateModels]);

  return (
    <div className="min-h-screen bg-slate-900 text-white p-4">
      <div className="max-w-7xl mx-auto space-y-6">
        <div className="text-center mb-8">
          <h1 className="text-4xl font-bold text-lime-400 mb-2">
            {t('app_title')}
          </h1>
          <p className="text-slate-400 text-lg">
            {t('cycling.lab_subtitle')}
          </p>
          
          {/* Blocco CTA Ottimizzato */}
          <div className="mt-6 flex flex-col items-center gap-3">
            {/* Riga dei pulsanti perfettamente allineati */}
            <div className="flex flex-col sm:flex-row items-center gap-4 mobile-cta-buttons">
              <PdfLeadForm models={currentModels} weight={bodyWeight} gender={gender} />
              <AthleteCardGenerator 
                models={currentModels} 
                bodyWeight={bodyWeight}
                gender={gender}
              />
              <ConsultationModal 
                models={currentModels}
                isOpen={showConsultationModal}
                onOpenChange={setShowConsultationModal}
                weight={bodyWeight}
                gender={gender}
              />
            </div>
            
            {/* Testo di supporto centrato sotto entrambi i pulsanti */}
            <p className="text-slate-400 text-sm max-w-md text-center">
              {t('cycling.cta_description')}
            </p>
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mobile-stack">
          {/* Controls Column */}
          <div className="lg:col-span-1 space-y-4 order-2 lg:order-1">
            {/* Body Weight and Gender Input */}
            <Card className="bg-slate-800 border-slate-700">
              <CardHeader className="pb-3">
                <CardTitle className="text-lime-400 text-lg">{t('common.athlete_params')}</CardTitle>
                <p className="text-slate-400 text-sm">Peso corporeo e sesso necessari per i calcoli.</p>
              </CardHeader>
              <CardContent>
                <div className="space-y-4">
                  {/* Gender Selection */}
                  <div className="space-y-3">
                    <Label className="text-slate-200 font-medium">{t('common.gender')}</Label>
                    <div className="flex space-x-4">
                      <label className="flex items-center space-x-2 cursor-pointer">
                        <input
                          type="radio"
                          name="gender"
                          value="male"
                          checked={gender === 'male'}
                          onChange={(e) => setGender(e.target.value as 'male' | 'female')}
                          className="w-4 h-4 text-lime-400 bg-slate-700 border-slate-600 focus:ring-lime-400 focus:ring-2"
                        />
                        <span className="text-slate-200">{t('common.male')}</span>
                      </label>
                      <label className="flex items-center space-x-2 cursor-pointer">
                        <input
                          type="radio"
                          name="gender"
                          value="female"
                          checked={gender === 'female'}
                          onChange={(e) => setGender(e.target.value as 'male' | 'female')}
                          className="w-4 h-4 text-lime-400 bg-slate-700 border-slate-600 focus:ring-lime-400 focus:ring-2"
                        />
                        <span className="text-slate-200">{t('common.female')}</span>
                      </label>
                    </div>
                  </div>
                  
                  {/* Body Weight */}
                  <div className="space-y-3">
                    <div className="flex justify-between items-center">
                      <Label className="text-slate-200 font-medium">{t('common.weight')}</Label>
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
                </div>
              </CardContent>
            </Card>

            {/* Pmax */}
            <Card className="bg-slate-800 border-slate-700">
              <CardHeader className="pb-3">
                <CardTitle className="text-lime-400 text-lg">{t('cycling.controls.pmax_title')}</CardTitle>
                <p className="text-slate-400 text-sm">{t('cycling.controls.pmax_description')}</p>
              </CardHeader>
              <CardContent>
                <div className="space-y-3">
                  <div className="flex justify-between items-center">
                    <Label className="text-slate-200 font-medium">{t('cycling.controls.power')}</Label>
                    <span className="font-semibold bg-slate-700 text-lime-400 px-3 py-1 rounded">
                      {pmaxPower} W
                    </span>
                  </div>
                  <Slider
                    value={[pmaxPower]}
                    onValueChange={(value) => setPmaxPower(value[0])}
                    min={300}
                    max={2500}
                    step={5}
                    className="w-full"
                  />
                </div>
              </CardContent>
            </Card>

            {/* Test Sforzo Breve */}
            <Card className="bg-slate-800 border-slate-700">
              <CardHeader className="pb-3">
                <CardTitle className="text-lime-400 text-lg">{t('cycling.controls.severe_test_title')}</CardTitle>
                <p className="text-slate-400 text-sm">{t('cycling.controls.severe_test_description')}</p>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="space-y-3">
                  <div className="flex justify-between items-center">
                    <Label className="text-slate-200 font-medium">{t('cycling.controls.power')}</Label>
                    <span className="font-semibold bg-slate-700 text-lime-400 px-3 py-1 rounded">
                      {severePower} W
                    </span>
                  </div>
                  <Slider
                    value={[severePower]}
                    onValueChange={(value) => setSeverePower(value[0])}
                    min={150}
                    max={800}
                    step={1}
                    className="w-full"
                  />
                </div>
                <div className="space-y-3">
                  <div className="flex justify-between items-center">
                    <Label className="text-slate-200 font-medium">{t('cycling.controls.time')}</Label>
                    <span className="font-semibold bg-slate-700 text-lime-400 px-3 py-1 rounded">
                      {formatForDisplay.time(severeTime)}
                    </span>
                  </div>
                  <Slider
                    value={[severeTime]}
                    onValueChange={(value) => setSevereTime(value[0])}
                    min={180}
                    max={300}
                    step={1}
                    className="w-full"
                  />
                </div>
              </CardContent>
            </Card>

            {/* Test Sforzo Medio */}
            <Card className="bg-slate-800 border-slate-700">
              <CardHeader className="pb-3">
                <CardTitle className="text-lime-400 text-lg">{t('cycling.controls.threshold_test_title')}</CardTitle>
                <p className="text-slate-400 text-sm">{t('cycling.controls.threshold_test_description')}</p>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="space-y-3">
                  <div className="flex justify-between items-center">
                    <Label className="text-slate-200 font-medium">{t('cycling.controls.power')}</Label>
                    <span className="font-semibold bg-slate-700 text-lime-400 px-3 py-1 rounded">
                      {thresholdPower} W
                    </span>
                  </div>
                  <Slider
                    value={[thresholdPower]}
                    onValueChange={(value) => setThresholdPower(value[0])}
                    min={100}
                    max={600}
                    step={1}
                    className="w-full"
                  />
                </div>
                <div className="space-y-3">
                  <div className="flex justify-between items-center">
                    <Label className="text-slate-200 font-medium">{t('cycling.controls.time')}</Label>
                    <span className="font-semibold bg-slate-700 text-lime-400 px-3 py-1 rounded">
                      {formatForDisplay.time(thresholdTime)}
                    </span>
                  </div>
                  <Slider
                    value={[thresholdTime]}
                    onValueChange={(value) => setThresholdTime(value[0])}
                    min={720}
                    max={1200}
                    step={1}
                    className="w-full"
                  />
                </div>
              </CardContent>
            </Card>

            {/* Test Sprint Lungo */}
            <Card className="bg-slate-800 border-slate-700">
              <CardHeader className="pb-3">
                <div className="flex items-center space-x-3">
                  <Switch
                    checked={enableShortTest}
                    onCheckedChange={setEnableShortTest}
                  />
                  <CardTitle className="text-lime-400 text-lg">{t('cycling.controls.short_test_title')}</CardTitle>
                </div>
                <p className="text-slate-400 text-sm">{t('cycling.controls.short_test_description')}</p>
              </CardHeader>
              <CardContent className="space-y-4">
                <fieldset disabled={!enableShortTest} className="space-y-4 disabled:opacity-50">
                  <div className="space-y-3">
                    <div className="flex justify-between items-center">
                      <Label className="text-slate-200 font-medium">{t('cycling.controls.power')}</Label>
                      <span className="font-semibold bg-slate-700 text-lime-400 px-3 py-1 rounded">
                        {shortPower} W
                      </span>
                    </div>
                    <Slider
                      value={[shortPower]}
                      onValueChange={(value) => setShortPower(value[0])}
                      min={200}
                      max={1500}
                      step={1}
                      className="w-full"
                    />
                  </div>
                  <div className="space-y-3">
                    <div className="flex justify-between items-center">
                      <Label className="text-slate-200 font-medium">{t('cycling.controls.time')}</Label>
                      <span className="font-semibold bg-slate-700 text-lime-400 px-3 py-1 rounded">
                        {formatForDisplay.time(shortTime)}
                      </span>
                    </div>
                    <Slider
                      value={[shortTime]}
                      onValueChange={(value) => setShortTime(value[0])}
                      min={15}
                      max={90}
                      step={1}
                      className="w-full"
                    />
                  </div>
                </fieldset>
              </CardContent>
            </Card>

            {/* Test Lungo */}
            <Card className="bg-slate-800 border-slate-700">
              <CardHeader className="pb-3">
                <div className="flex items-center space-x-3">
                  <Switch
                    checked={enableLongTest}
                    onCheckedChange={setEnableLongTest}
                  />
                  <CardTitle className="text-lime-400 text-lg">{t('cycling.controls.long_test_title')}</CardTitle>
                </div>
                <p className="text-slate-400 text-sm">{t('cycling.controls.long_test_description')}</p>
              </CardHeader>
              <CardContent className="space-y-4">
                <fieldset disabled={!enableLongTest} className="space-y-4 disabled:opacity-50">
                  <div className="space-y-3">
                    <div className="flex justify-between items-center">
                      <Label className="text-slate-200 font-medium">{t('cycling.controls.power')}</Label>
                      <span className="font-semibold bg-slate-700 text-lime-400 px-3 py-1 rounded">
                        {longPower} W
                      </span>
                    </div>
                    <Slider
                      value={[longPower]}
                      onValueChange={(value) => setLongPower(value[0])}
                      min={80}
                      max={500}
                      step={1}
                      className="w-full"
                    />
                  </div>
                  <div className="space-y-3">
                    <div className="flex justify-between items-center">
                      <Label className="text-slate-200 font-medium">{t('cycling.controls.time')}</Label>
                      <span className="font-semibold bg-slate-700 text-lime-400 px-3 py-1 rounded">
                        {formatForDisplay.time(longTime)}
                      </span>
                    </div>
                    <Slider
                      value={[longTime]}
                      onValueChange={(value) => setLongTime(value[0])}
                      min={1200}
                      max={21600}
                      step={60}
                      className="w-full"
                    />
                  </div>
                </fieldset>
              </CardContent>
            </Card>
          </div>

          {/* Chart and Parameters Column */}
          <div className="lg:col-span-2 space-y-6 order-1 lg:order-2">
            {/* Model Parameters */}
            <ModelParameters models={currentModels} bodyWeight={bodyWeight} />
            
            {/* Chart */}
            <Card className="bg-slate-800 border-slate-700">
              <CardHeader>
                <CardTitle className={`text-xl ${currentModels.valid ? 'text-lime-400' : 'text-red-400'}`}>
                  {currentModels.valid ? t('cycling.chart.physiological_profile') : t('cycling.chart.invalid_data')}
                </CardTitle>
                <p className="text-slate-400">
                  {currentModels.valid 
                    ? t('cycling.chart.valid_description') 
                    : currentModels.reason
                  }
                </p>
              </CardHeader>
              <CardContent>
                <div className="h-[400px] lg:h-[600px]">
                  <PowerChart models={currentModels} />
                </div>
              </CardContent>
            </Card>
          </div>
        </div>

        {/* Power Zones Table - Full Width */}
        <div className="mt-8">
          <PowerZonesTable models={currentModels} />
        </div>

        {/* Physiological Profile - Full Width */}
        <div className="mt-8" id="physiological-profile">
          <PhysiologicalProfile models={currentModels} bodyWeight={bodyWeight} gender={gender} />
        </div>

        {/* Power Performance Calculator - Full Width */}
        <div className="mt-8">
          <PowerPerformanceCalculator models={currentModels} />
        </div>
        
        {/* Sustainable Powers Table - Full Width */}
        <div className="mt-8">
          <PowerSustainableTable models={currentModels} bodyWeight={bodyWeight} />
        </div>
      </div>
    </div>
  );
};

export default PowerCurveModeler;


import React, { useState, useEffect } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { PowerPrecisionModels, formatForDisplay } from '../utils/powerPrecisionCalculations';
import { useTranslation } from 'react-i18next';

interface PowerPerformanceCalculatorProps {
  models: PowerPrecisionModels;
}

interface DualPrediction {
  aprPower: number;
  cpwPower: number;
  difference: number;
  differencePercent: number;
}

interface DualTimePrediction {
  aprTime: number;
  cpwTime: number;
  difference: number;
  differencePercent: number;
}

const PowerPerformanceCalculator: React.FC<PowerPerformanceCalculatorProps> = ({ models }) => {
  const { t } = useTranslation();
  // Time to Power calculator
  const [inputHours, setInputHours] = useState(0);
  const [inputMinutes, setInputMinutes] = useState(20);
  const [inputSeconds, setInputSeconds] = useState(0);
  const [calculatedPower, setCalculatedPower] = useState<number | null>(null);
  const [usedModel, setUsedModel] = useState<string>('');
  const [dualPrediction, setDualPrediction] = useState<DualPrediction | null>(null);
  
  // Power to Time calculator
  const [inputPower, setInputPower] = useState(300);
  const [calculatedTime, setCalculatedTime] = useState<number | null>(null);
  const [usedModelTime, setUsedModelTime] = useState<string>('');
  const [dualTimePrediction, setDualTimePrediction] = useState<DualTimePrediction | null>(null);

  // APR model power calculation - USING EXACT CHART FORMULAS
  const calculateAPRPower = (timeSeconds: number): number => {
    const { pMax, apr_params } = models;
    
    // EXACT same formula as PowerChart.tsx
    if (timeSeconds === 1) {
      return pMax;
    } else {
      return apr_params.po3min + apr_params.apr * Math.exp(-apr_params.k * timeSeconds);
    }
  };

  // CP-W' model power calculation
  const calculateCPWPower = (timeSeconds: number): number => {
    return models.cp + models.wPrime / timeSeconds;
  };

  // APR solver - same as before but using correct formula
  const solveAPRTime = (targetPower: number): number => {
    const { pMax, apr_params } = models;
    
    console.log(`=== APR SOLVER ===`);
    console.log(`Target power: ${targetPower}W`);
    console.log(`Model params: Pmax=${pMax}, po3min=${apr_params.po3min.toFixed(1)}, APR=${apr_params.apr.toFixed(1)}, k=${apr_params.k}`);
    
    // Check bounds
    const maxPowerAPR = pMax; // At 1 second
    const minPowerAPR = calculateAPRPower(180); // At 3 minutes
    
    console.log(`APR power range: ${maxPowerAPR.toFixed(1)}W (1s) to ${minPowerAPR.toFixed(1)}W (180s)`);
    
    if (targetPower > maxPowerAPR || targetPower < minPowerAPR) {
      console.log(`Target power ${targetPower}W outside APR range`);
      return NaN;
    }
    
    // Special case for Pmax
    if (Math.abs(targetPower - pMax) < 0.1) {
      return 1;
    }
    
    // Bisection method for po3min + apr * exp(-k*t) = targetPower
    let minTime = 1;
    let maxTime = 180;
    const tolerance = 0.01;
    const powerTolerance = 0.1;
    
    let iterations = 0;
    const maxIterations = 100;
    
    while (maxTime - minTime > tolerance && iterations < maxIterations) {
      const midTime = (minTime + maxTime) / 2;
      const midPower = calculateAPRPower(midTime);
      
      if (Math.abs(midPower - targetPower) < powerTolerance) {
        console.log(`APR solver converged: ${midTime.toFixed(2)}s for ${targetPower}W`);
        return midTime;
      }
      
      if (midPower > targetPower) {
        minTime = midTime;
      } else {
        maxTime = midTime;
      }
      
      iterations++;
    }
    
    const finalTime = (minTime + maxTime) / 2;
    console.log(`APR solver result: ${finalTime.toFixed(2)}s`);
    return finalTime;
  };

  const calculatePowerFromTime = () => {
    if (!models.valid) return;
    
    const totalSeconds = inputHours * 3600 + inputMinutes * 60 + inputSeconds;
    if (totalSeconds <= 0) return;
    
    console.log(`\n=== POWER FROM TIME CALCULATION ===`);
    console.log(`Input: ${totalSeconds}s (${formatForDisplay.time(totalSeconds)})`);
    
    // Check if we're in the transition zone (2-3 minutes)
    const isTransitionZone = totalSeconds >= 120 && totalSeconds <= 180;
    
    if (isTransitionZone) {
      console.log(`*** TRANSITION ZONE (2-3min) - Calculating both models ***`);
      
      // Calculate both APR and CP-W' predictions
      const aprPower = calculateAPRPower(totalSeconds);
      const cpwPower = calculateCPWPower(totalSeconds);
      const difference = Math.abs(aprPower - cpwPower);
      const differencePercent = (difference / Math.max(aprPower, cpwPower)) * 100;
      
      console.log(`APR Model: ${aprPower.toFixed(1)}W`);
      console.log(`CP-W' Model: ${cpwPower.toFixed(1)}W`);
      console.log(`Difference: ${difference.toFixed(1)}W (${differencePercent.toFixed(1)}%)`);
      
      setDualPrediction({
        aprPower,
        cpwPower,
        difference,
        differencePercent
      });
      
      // Use the lower prediction as the main result (more conservative)
      const primaryPower = Math.min(aprPower, cpwPower);
      const primaryModel = aprPower < cpwPower ? 'APR' : 'CP-W\'';
      
      setCalculatedPower(primaryPower);
      setUsedModel(`${primaryModel} (zona transizione)`);
    } else {
      // Normal single model logic
      setDualPrediction(null);
      
      let power: number;
      let modelUsed: string;
      
      if (totalSeconds <= 180) { // ≤ 3 minutes: use APR model
        power = calculateAPRPower(totalSeconds);
        modelUsed = 'APR';
        console.log(`Using APR model: ${power.toFixed(1)}W`);
      } else if (totalSeconds <= 960) { // 3-16 minutes: use CP-W' model
        power = calculateCPWPower(totalSeconds);
        modelUsed = 'CP-W\'';
        console.log(`Using CP-W' model: CP=${models.cp.toFixed(1)} + W'/t=${(models.wPrime/totalSeconds).toFixed(1)} = ${power.toFixed(1)}W`);
      } else { // > 16 minutes: use Power Law model
        power = models.pl_params.S * Math.pow(totalSeconds, models.pl_params.E - 1);
        modelUsed = 'Power Law';
        console.log(`Using Power Law model: S=${models.pl_params.S.toFixed(1)} * t^(E-1) = ${power.toFixed(1)}W`);
      }
      
      console.log(`Final result: ${power.toFixed(1)}W using ${modelUsed} model`);
      
      setCalculatedPower(power);
      setUsedModel(modelUsed);
    }
  };

  const calculateTimeFromPower = () => {
    if (!models.valid || inputPower <= 0) return;
    
    console.log(`\n=== TIME FROM POWER CALCULATION ===`);
    console.log(`Target power: ${inputPower}W`);
    
    let time: number = NaN;
    let modelUsed: string = '';
    let dualResult: DualTimePrediction | null = null;
    
    // Try APR model first (for high powers, short durations)
    console.log(`\n--- Trying APR model ---`);
    const aprTime = solveAPRTime(inputPower);
    
    // Try CP-W' model
    console.log(`\n--- Trying CP-W' model ---`);
    let cpwTime = NaN;
    if (inputPower > models.cp) {
      cpwTime = models.wPrime / (inputPower - models.cp);
      console.log(`CP-W' calculation: t = W'/(P-CP) = ${models.wPrime.toFixed(0)}/(${inputPower}-${models.cp.toFixed(1)}) = ${cpwTime.toFixed(2)}s`);
    }
    
    // Check if both models give valid results in the transition zone
    const aprInTransition = !isNaN(aprTime) && aprTime >= 120 && aprTime <= 180;
    const cpwInTransition = !isNaN(cpwTime) && cpwTime >= 120 && cpwTime <= 180;
    
    if (aprInTransition && cpwInTransition) {
      console.log(`*** BOTH MODELS VALID IN TRANSITION ZONE ***`);
      
      const difference = Math.abs(aprTime - cpwTime);
      const differencePercent = (difference / Math.max(aprTime, cpwTime)) * 100;
      
      dualResult = {
        aprTime,
        cpwTime,
        difference,
        differencePercent
      };
      
      // Use the shorter time as primary (more conservative for performance prediction)
      time = Math.min(aprTime, cpwTime);
      modelUsed = aprTime < cpwTime ? 'APR (zona transizione)' : 'CP-W\' (zona transizione)';
      
      console.log(`APR time: ${aprTime.toFixed(2)}s`);
      console.log(`CP-W' time: ${cpwTime.toFixed(2)}s`);
      console.log(`Using: ${time.toFixed(2)}s (${modelUsed})`);
    } else {
      // Standard single model selection
      if (!isNaN(aprTime) && aprTime >= 1 && aprTime <= 180) {
        time = aprTime;
        modelUsed = 'APR';
        console.log(`APR solution: ${time.toFixed(2)}s`);
      } else if (!isNaN(cpwTime) && cpwTime > 180 && cpwTime <= 960) {
        time = cpwTime;
        modelUsed = 'CP-W\'';
        console.log(`CP-W' solution: ${time.toFixed(2)}s`);
      } else if (!isNaN(cpwTime) && cpwTime <= 180) {
        time = cpwTime;
        modelUsed = 'CP-W\' (fallback)';
      } else {
        // Use Power Law model
        console.log(`\n--- Using Power Law model ---`);
        const exponent = 1 / (models.pl_params.E - 1);
        time = Math.pow(inputPower / models.pl_params.S, exponent);
        modelUsed = 'Power Law';
        console.log(`Power Law calculation: t = (P/S)^(1/(E-1)) = ${time.toFixed(2)}s`);
      }
    }
    
    // Validate result
    if (isNaN(time) || time <= 0) {
      console.log(`Invalid result: ${time}`);
      time = NaN;
      modelUsed = 'No solution';
      dualResult = null;
    } else {
      console.log(`Final result: ${formatForDisplay.time(time)} using ${modelUsed} model`);
    }
    
    setCalculatedTime(time);
    setUsedModelTime(modelUsed);
    setDualTimePrediction(dualResult);
  };

  // Auto-calculate when inputs change
  useEffect(() => {
    if (models.valid) {
      calculatePowerFromTime();
    }
  }, [inputHours, inputMinutes, inputSeconds, models]);

  useEffect(() => {
    if (models.valid) {
      calculateTimeFromPower();
    }
  }, [inputPower, models]);

  if (!models.valid) {
    return (
      <Card className="bg-slate-800 border-slate-700">
        <CardHeader>
          <CardTitle className="text-red-400">{t('calculators.performance_calculator')}</CardTitle>
        </CardHeader>
        <CardContent>
          <p className="text-slate-400">{t('calculators.invalid_data')}</p>
        </CardContent>
      </Card>
    );
  }

  return (
    <Card className="bg-slate-800 border-slate-700">
      <CardHeader>
        <CardTitle className="text-lime-400">{t('calculators.performance_calculator')}</CardTitle>
        <p className="text-slate-400 text-sm">
          {t('calculators.explanation')}
          <br />
          <span className="text-yellow-400">🔀 {t('calculators.transition_explanation')}</span>
        </p>
      </CardHeader>
      <CardContent>
        <Tabs defaultValue="time-to-power" className="w-full">
          <TabsList className="grid w-full grid-cols-2 bg-slate-700">
            <TabsTrigger value="time-to-power" className="data-[state=active]:bg-lime-400 data-[state=active]:text-slate-900">
              {t('calculators.duration_to_power')}
            </TabsTrigger>
            <TabsTrigger value="power-to-time" className="data-[state=active]:bg-lime-400 data-[state=active]:text-slate-900">
              {t('calculators.power_to_duration')}
            </TabsTrigger>
          </TabsList>
          
          <TabsContent value="time-to-power" className="space-y-4">
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4 items-end">
              <div className="space-y-2">
                <Label className="text-slate-200">{t('running.controls.hours')}</Label>
                <Input
                  type="number"
                  min="0"
                  max="10"
                  value={inputHours}
                  onChange={(e) => setInputHours(Number(e.target.value))}
                  className="bg-slate-700 border-slate-600 text-white"
                />
              </div>
              <div className="space-y-2">
                <Label className="text-slate-200">{t('running.controls.minutes')}</Label>
                <Input
                  type="number"
                  min="0"
                  max="59"
                  value={inputMinutes}
                  onChange={(e) => setInputMinutes(Number(e.target.value))}
                  className="bg-slate-700 border-slate-600 text-white"
                />
              </div>
              <div className="space-y-2">
                <Label className="text-slate-200">{t('running.controls.seconds')}</Label>
                <Input
                  type="number"
                  min="0"
                  max="59"
                  value={inputSeconds}
                  onChange={(e) => setInputSeconds(Number(e.target.value))}
                  className="bg-slate-700 border-slate-600 text-white"
                />
              </div>
              <div className="space-y-2">
                <Label className="text-slate-200">{t('calculators.sustainable_power')}</Label>
                <div className="h-10 bg-slate-900 rounded border border-slate-600 flex items-center px-3">
                  <span className="text-lime-400 font-semibold">
                    {calculatedPower ? `${Math.round(calculatedPower)} W` : '--- W'}
                  </span>
                </div>
              </div>
            </div>
            
            {/* Dual prediction display for transition zone */}
            {dualPrediction && (
              <div className="mt-4 p-4 bg-gradient-to-r from-purple-900/30 to-red-900/30 rounded-lg border border-yellow-500/50">
                <h4 className="text-yellow-400 font-semibold mb-2 flex items-center">
                  🔀 Zona di Transizione (2-3min) - Confronto Modelli
                </h4>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-sm">
                  <div className="text-center">
                    <div className="text-purple-300 font-medium">APR Model</div>
                    <div className="text-white font-bold text-lg">{Math.round(dualPrediction.aprPower)} W</div>
                  </div>
                  <div className="text-center">
                    <div className="text-red-300 font-medium">CP-W' Model</div>
                    <div className="text-white font-bold text-lg">{Math.round(dualPrediction.cpwPower)} W</div>
                  </div>
                  <div className="text-center">
                    <div className="text-yellow-300 font-medium">Differenza</div>
                    <div className="text-white font-bold text-lg">
                      {Math.round(dualPrediction.difference)} W ({dualPrediction.differencePercent.toFixed(1)}%)
                    </div>
                  </div>
                </div>
                <p className="text-xs text-slate-400 mt-2 text-center">
                  Entrambi i modelli sono applicabili in questa zona. Risultato principale: valore più basso (conservativo).
                </p>
              </div>
            )}
            
            {calculatedPower && (
              <div className="mt-4 p-3 bg-slate-900 rounded border border-slate-600">
                <p className="text-xs text-slate-400">
                  <strong>Tempo totale:</strong> {formatForDisplay.time(inputHours * 3600 + inputMinutes * 60 + inputSeconds)} •{' '}
                  <strong>Modello utilizzato:</strong> {usedModel}
                </p>
                <div className="mt-2 text-xs text-slate-500">
                  <strong>Selezione automatica:</strong> APR (≤3min) • CP-W' (3-16min) • Power Law (&gt;16min) • Zona transizione (2-3min)
                </div>
              </div>
            )}
          </TabsContent>
          
          <TabsContent value="power-to-time" className="space-y-4">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 items-end">
              <div className="space-y-2">
                <Label className="text-slate-200">Potenza Target (W)</Label>
                <Input
                  type="number"
                  min="50"
                  max="2000"
                  step="5"
                  value={inputPower}
                  onChange={(e) => setInputPower(Number(e.target.value))}
                  className="bg-slate-700 border-slate-600 text-white"
                />
              </div>
              <div className="space-y-2">
                <Label className="text-slate-200">Durata Sostenibile</Label>
                <div className="h-10 bg-slate-900 rounded border border-slate-600 flex items-center px-3">
                  <span className="text-lime-400 font-semibold">
                    {calculatedTime && !isNaN(calculatedTime) ? formatForDisplay.time(calculatedTime) : '--- '}
                  </span>
                </div>
              </div>
            </div>
            
            {/* Dual time prediction display for transition zone */}
            {dualTimePrediction && (
              <div className="mt-4 p-4 bg-gradient-to-r from-purple-900/30 to-red-900/30 rounded-lg border border-yellow-500/50">
                <h4 className="text-yellow-400 font-semibold mb-2 flex items-center">
                  🔀 Zona di Transizione (2-3min) - Confronto Modelli
                </h4>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-sm">
                  <div className="text-center">
                    <div className="text-purple-300 font-medium">APR Model</div>
                    <div className="text-white font-bold text-lg">{formatForDisplay.time(dualTimePrediction.aprTime)}</div>
                  </div>
                  <div className="text-center">
                    <div className="text-red-300 font-medium">CP-W' Model</div>
                    <div className="text-white font-bold text-lg">{formatForDisplay.time(dualTimePrediction.cpwTime)}</div>
                  </div>
                  <div className="text-center">
                    <div className="text-yellow-300 font-medium">Differenza</div>
                    <div className="text-white font-bold text-lg">
                      {formatForDisplay.time(dualTimePrediction.difference)} ({dualTimePrediction.differencePercent.toFixed(1)}%)
                    </div>
                  </div>
                </div>
                <p className="text-xs text-slate-400 mt-2 text-center">
                  Entrambi i modelli predicono durate nella zona 2-3min. Risultato principale: durata più breve (conservativo).
                </p>
              </div>
            )}
            
            {calculatedTime && !isNaN(calculatedTime) && (
              <div className="mt-4 p-3 bg-slate-900 rounded border border-slate-600">
                <p className="text-xs text-slate-400">
                  <strong>Potenza:</strong> {inputPower} W •{' '}
                  <strong>Modello utilizzato:</strong> {usedModelTime}
                </p>
                <div className="mt-2 text-xs text-slate-500">
                  <strong>Logica di selezione:</strong> APR prima per alte potenze, poi CP-W', infine Power Law per basse potenze
                </div>
              </div>
            )}
          </TabsContent>
        </Tabs>
      </CardContent>
    </Card>
  );
};

export default PowerPerformanceCalculator;


import React, { useState } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { PrecisionModels, formatForDisplay } from '../utils/precisionCalculations';

interface PerformanceCalculatorProps {
  models: PrecisionModels;
}

const PerformanceCalculator: React.FC<PerformanceCalculatorProps> = ({ models }) => {
  const [distanceInput, setDistanceInput] = useState('');
  const [distanceUnit, setDistanceUnit] = useState('m');
  const [timeHours, setTimeHours] = useState('0');
  const [timeMinutes, setTimeMinutes] = useState('');
  const [timeSeconds, setTimeSeconds] = useState('');
  const [result, setResult] = useState<{
    predictedTime?: string;
    predictedDistance?: string;
    pace?: string;
  } | null>(null);

  const calculateFromDistance = () => {
    if (!models.valid || !models.cs || !models.dPrime || !models.pl_params) {
      return;
    }

    const distance = parseFloat(distanceInput);
    if (!distance || distance <= 0) return;

    const distanceInMeters = distanceUnit === 'km' ? distance * 1000 : distance;

    console.log('=== CALCOLO DA DISTANZA ===');
    console.log('Distanza:', distanceInMeters, 'm');

    // Try CS-D' model first - using high precision values
    const timeFromCS = (distanceInMeters - models.dPrime) / models.cs;
    
    let predictedTime: number;
    let modelUsed: string;

    if (timeFromCS > 0 && timeFromCS < 1020) { // < 17 minutes
      predictedTime = timeFromCS;
      modelUsed = "CS-D'";
      console.log('Usando CS-D\': tempo =', predictedTime, 's');
    } else {
      // Use Power Law model: v = S * t^(E-1)  =>  t = (D/S)^(1/E)
      const S = models.pl_params.A;     // velocità base
      const E_minus_1 = models.pl_params.k; // E-1
      const E = E_minus_1 + 1;          // E
      
      // Formula corretta: t = (D/S)^(1/E)
      predictedTime = Math.pow(distanceInMeters / S, 1 / E);
      modelUsed = "Power Law";
      
      console.log('Usando Power Law:');
      console.log('  S =', S, 'm/s');
      console.log('  E =', E);
      console.log('  t = (', distanceInMeters, '/', S, ')^(1/', E, ') =', predictedTime, 's');
    }

    const pacePerKm = (predictedTime * 1000) / distanceInMeters;
    console.log('Passo finale:', formatForDisplay.pace(pacePerKm), '/km');

    setResult({
      predictedTime: `${formatForDisplay.time(predictedTime)} (${modelUsed})`,
      pace: formatForDisplay.pace(pacePerKm)
    });
  };

  const calculateFromTime = () => {
    if (!models.valid || !models.cs || !models.dPrime || !models.pl_params) {
      return;
    }

    const hours = parseInt(timeHours) || 0;
    const minutes = parseInt(timeMinutes) || 0;
    const seconds = parseInt(timeSeconds) || 0;
    const totalSeconds = hours * 3600 + minutes * 60 + seconds;

    if (totalSeconds <= 0) return;

    console.log('=== CALCOLO DA TEMPO ===');
    console.log('Tempo:', totalSeconds, 's');

    let predictedDistance: number;
    let modelUsed: string;

    if (totalSeconds < 1020) { // < 17 minutes, use CS-D'
      predictedDistance = models.cs * totalSeconds + models.dPrime;
      modelUsed = "CS-D'";
      console.log('Usando CS-D\': distanza =', predictedDistance, 'm');
    } else {
      // Use Power Law model: v = S * t^(E-1)  =>  D = v * t = S * t^E
      const S = models.pl_params.A;     // velocità base
      const E_minus_1 = models.pl_params.k; // E-1
      const E = E_minus_1 + 1;          // E
      
      // Formula corretta: D = S * t^E
      predictedDistance = S * Math.pow(totalSeconds, E);
      modelUsed = "Power Law";
      
      console.log('Usando Power Law:');
      console.log('  S =', S, 'm/s');
      console.log('  E =', E);
      console.log('  D = S * t^E =', S, '*', totalSeconds, '^', E, '=', predictedDistance, 'm');
    }

    const pacePerKm = (totalSeconds * 1000) / predictedDistance;
    const distanceInKm = predictedDistance / 1000;

    console.log('Distanza finale:', distanceInKm, 'km');
    console.log('Passo finale:', formatForDisplay.pace(pacePerKm), '/km');

    setResult({
      predictedDistance: `${formatForDisplay.distance(predictedDistance)} (${modelUsed})`,
      pace: formatForDisplay.pace(pacePerKm)
    });
  };

  if (!models.valid) {
    return (
      <Card className="bg-slate-800 border-slate-700">
        <CardHeader>
          <CardTitle className="text-red-400">Calcolatore Performance</CardTitle>
        </CardHeader>
        <CardContent>
          <p className="text-slate-400">Modelli non validi. Inserire dati corretti per utilizzare il calcolatore.</p>
        </CardContent>
      </Card>
    );
  }

  return (
    <Card className="bg-slate-800 border-slate-700">
      <CardHeader>
        <CardTitle className="text-lime-400">Calcolatore Performance</CardTitle>
        <p className="text-slate-400">
          Predici tempo o distanza usando i modelli CS-D' (&lt;17min) e Power Law (&gt;17min)
        </p>
      </CardHeader>
      <CardContent>
        <Tabs defaultValue="distance" className="w-full">
          <TabsList className="grid w-full grid-cols-2 bg-slate-700">
            <TabsTrigger value="distance" className="data-[state=active]:bg-lime-600">
              Da Distanza a Tempo
            </TabsTrigger>
            <TabsTrigger value="time" className="data-[state=active]:bg-lime-600">
              Da Tempo a Distanza
            </TabsTrigger>
          </TabsList>
          
          <TabsContent value="distance" className="space-y-4">
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <div className="space-y-2">
                <Label className="text-slate-200">Distanza</Label>
                <Input
                  type="number"
                  placeholder="Inserisci distanza"
                  value={distanceInput}
                  onChange={(e) => setDistanceInput(e.target.value)}
                  className="bg-slate-700 border-slate-600"
                />
              </div>
              <div className="space-y-2">
                <Label className="text-slate-200">Unità</Label>
                <Select value={distanceUnit} onValueChange={setDistanceUnit}>
                  <SelectTrigger className="bg-slate-700 border-slate-600">
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent className="bg-slate-700 border-slate-600">
                    <SelectItem value="m">metri</SelectItem>
                    <SelectItem value="km">chilometri</SelectItem>
                  </SelectContent>
                </Select>
              </div>
              <div className="flex items-end">
                <Button 
                  onClick={calculateFromDistance}
                  className="bg-lime-600 hover:bg-lime-700 w-full"
                >
                  Calcola Tempo
                </Button>
              </div>
            </div>
          </TabsContent>
          
          <TabsContent value="time" className="space-y-4">
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
              <div className="space-y-2">
                <Label className="text-slate-200">Ore</Label>
                <Input
                  type="number"
                  placeholder="0"
                  value={timeHours}
                  onChange={(e) => setTimeHours(e.target.value)}
                  className="bg-slate-700 border-slate-600"
                />
              </div>
              <div className="space-y-2">
                <Label className="text-slate-200">Minuti</Label>
                <Input
                  type="number"
                  placeholder="Minuti"
                  value={timeMinutes}
                  onChange={(e) => setTimeMinutes(e.target.value)}
                  className="bg-slate-700 border-slate-600"
                />
              </div>
              <div className="space-y-2">
                <Label className="text-slate-200">Secondi</Label>
                <Input
                  type="number"
                  placeholder="Secondi"
                  value={timeSeconds}
                  onChange={(e) => setTimeSeconds(e.target.value)}
                  className="bg-slate-700 border-slate-600"
                />
              </div>
              <div className="flex items-end">
                <Button 
                  onClick={calculateFromTime}
                  className="bg-lime-600 hover:bg-lime-700 w-full"
                >
                  Calcola Distanza
                </Button>
              </div>
            </div>
          </TabsContent>
        </Tabs>

        {/* Results */}
        {result && (
          <div className="mt-6 p-4 bg-slate-900 rounded-lg border border-lime-500/30">
            <h4 className="text-lime-400 font-semibold mb-3">Risultati Predizione</h4>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {result.predictedTime && (
                <div className="text-center">
                  <div className="text-slate-400 text-sm">Tempo Previsto</div>
                  <div className="text-white font-bold text-lg">{result.predictedTime}</div>
                </div>
              )}
              {result.predictedDistance && (
                <div className="text-center">
                  <div className="text-slate-400 text-sm">Distanza Prevista</div>
                  <div className="text-white font-bold text-lg">{result.predictedDistance}</div>
                </div>
              )}
              {result.pace && (
                <div className="text-center">
                  <div className="text-slate-400 text-sm">Passo Medio</div>
                  <div className="text-white font-bold text-lg">{result.pace}/km</div>
                </div>
              )}
            </div>
          </div>
        )}
      </CardContent>
    </Card>
  );
};

export default PerformanceCalculator;

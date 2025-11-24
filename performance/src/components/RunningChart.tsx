import React, { useEffect, useRef } from 'react';

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

interface RunningChartProps {
  models: RunningModels;
}

const RunningChart: React.FC<RunningChartProps> = ({ models }) => {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const chartRef = useRef<any>(null);

  const formatPace = (secondsPerKm: number) => {
    if (!isFinite(secondsPerKm) || secondsPerKm <= 0) return "--:--";
    const minutes = Math.floor(secondsPerKm / 60);
    const seconds = Math.round(secondsPerKm % 60);
    return `${minutes}:${seconds.toString().padStart(2, '0')}`;
  };

  const formatTime = (totalSeconds: number, brief = false) => {
    if (!isFinite(totalSeconds) || totalSeconds <= 0) return "--:--";
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = Math.round(totalSeconds % 60);
    
    if (brief) {
      if (hours > 0) {
        return `${hours}h`;
      } else if (minutes > 0) {
        return `${minutes}m`;
      } else {
        return `${seconds}s`;
      }
    }
    
    if (hours > 0) {
      return `${hours}h ${minutes}m ${seconds}s`;
    } else if (minutes > 0) {
      return `${minutes}m ${seconds}s`;
    } else {
      return `${seconds}s`;
    }
  };

  const formatDistance = (meters: number) => {
    if (meters >= 1000) {
      return `${(meters / 1000).toFixed(1)}km`;
    }
    return `${Math.round(meters)}m`;
  };

  const calculateRacePace = (distance: number, models: RunningModels) => {
    if (!models.valid || !models.pl_params) return null;
    
    const S = models.pl_params.A;     // velocità base
    const E_minus_1 = models.pl_params.k; // E-1
    const E = E_minus_1 + 1;          // E
    
    // Formula corretta: t = (D/S)^(1/E)
    const time = Math.pow(distance / S, 1 / E);
    const pace = (time * 1000) / distance; // s/km
    
    return pace;
  };

  const calculateCSModel = (time: number, models: RunningModels) => {
    if (!models.valid || !models.cs || !models.dPrime) return null;
    const distance = models.cs * time + models.dPrime;
    if (distance <= 0) return null;
    const pace = (time * 1000) / distance;
    return { pace, distance };
  };

  const calculatePowerLaw = (time: number, models: RunningModels) => {
    if (!models.valid || !models.pl_params) return null;
    
    const S = models.pl_params.A;     // velocità base
    const E_minus_1 = models.pl_params.k; // E-1
    
    // Formula corretta: v = S * t^(E-1)
    const velocity = S * Math.pow(time, E_minus_1); // m/s
    const distance = velocity * time; // m
    const pace = 1000 / velocity; // s/km
    
    return { pace, distance };
  };

  const generateCurveData = (models: RunningModels) => {
    if (!models || !models.valid) return { csModel: [], powerLaw: [] };
    
    const csModelData: any[] = [];
    const powerLawData: any[] = [];
    
    // CS-D' Model: da 2min30sec (150s) a 60min (3600s)
    for (let time = 150; time <= 3600; time += 10) {
      const result = calculateCSModel(time, models);
      if (result) {
        csModelData.push({ 
          x: time, 
          y: result.pace,
          distance: result.distance,
          model: 'CS-D\''
        });
      }
    }
    
    // Power Law Model: da 3min (180s) a 6h (21600s)
    for (let time = 180; time <= 21600; time += 60) {
      const result = calculatePowerLaw(time, models);
      if (result) {
        powerLawData.push({ 
          x: time, 
          y: result.pace,
          distance: result.distance,
          model: 'Power Law'
        });
      }
    }
    
    return { csModel: csModelData, powerLaw: powerLawData };
  };

  useEffect(() => {
    const loadChart = async () => {
      const { Chart, registerables } = await import('chart.js');
      const annotationPlugin = await import('chartjs-plugin-annotation');
      
      Chart.register(...registerables, annotationPlugin.default);

      if (!canvasRef.current) return;
      
      if (chartRef.current) {
        chartRef.current.destroy();
      }

      const ctx = canvasRef.current.getContext('2d');
      if (!ctx) return;

      const curveData = generateCurveData(models);
      
      let minPace = 120; // 2:00/km (veloce) - sarà il valore più basso dell'asse
      let maxPace = 360; // 6:00/km (lento) - sarà il valore più alto dell'asse
      
      if (models.valid) {
        const allPoints = [...curveData.csModel, ...curveData.powerLaw];
        if (allPoints.length > 0) {
          const paces = allPoints.map(point => point.y);
          const actualMin = Math.min(...paces);
          const actualMax = Math.max(...paces);
          const range = actualMax - actualMin;
          
          minPace = Math.max(120, actualMin - range * 0.1);
          
          // Assicurarsi che il limite inferiore di LT1 sia sempre visibile
          if (models.lt1_range && models.lt1_range.max) {
            maxPace = Math.max(actualMax + range * 0.1, models.lt1_range.max * 1.05);
          } else {
            maxPace = actualMax + range * 0.1;
          }
          
          maxPace = Math.min(600, maxPace);
        }
      }

      const annotations: any = {};
      
      if (models.valid && models.cs_pace && models.vo2max_pace && models.lt1_range) {
        const { cs_pace, vo2max_pace, lt1_range } = models;
        
        // Intensity domains using the new LT1 values
        annotations.extremeBackground = {
          type: 'box',
          xMin: 150,
          xMax: 21600,
          yMin: minPace,
          yMax: vo2max_pace,
          backgroundColor: 'rgba(147, 51, 234, 0.1)',
          borderWidth: 0,
          z: -1
        };
        
        annotations.extremeLabel = {
          type: 'label',
          xValue: 2000,
          yValue: (minPace + vo2max_pace) / 2,
          backgroundColor: 'rgba(147, 51, 234, 0.8)',
          color: 'white',
          content: 'Extreme Exercise Intensity Domain',
          font: { size: 11, weight: 'bold' },
          padding: 6,
          borderRadius: 4
        };
        
        annotations.severeBackground = {
          type: 'box',
          xMin: 150,
          xMax: 21600,
          yMin: vo2max_pace,
          yMax: cs_pace,
          backgroundColor: 'rgba(239, 68, 68, 0.1)',
          borderWidth: 0,
          z: -1
        };
        
        annotations.severeLabel = {
          type: 'label',
          xValue: 2000,
          yValue: (vo2max_pace + cs_pace) / 2,
          backgroundColor: 'rgba(239, 68, 68, 0.8)',
          color: 'white',
          content: 'Severe Exercise Intensity Domain',
          font: { size: 11, weight: 'bold' },
          padding: 6,
          borderRadius: 4
        };
        
        const cs_range_min = cs_pace * (1 / 1.00); // 100% CS
        const cs_range_max = cs_pace * (1 / 0.96); // 96% CS
        
        annotations.csRangeBackground = {
          type: 'box',
          xMin: 150,
          xMax: 21600,
          yMin: cs_range_min,
          yMax: cs_range_max,
          backgroundColor: 'rgba(255, 165, 0, 0.2)',
          borderColor: 'rgba(255, 165, 0, 0.6)',
          borderWidth: 1,
          borderDash: [3, 3],
          z: 0
        };
        
        annotations.heavyBackground = {
          type: 'box',
          xMin: 150,
          xMax: 21600,
          yMin: cs_pace,
          yMax: lt1_range.estimate,
          backgroundColor: 'rgba(255, 193, 7, 0.1)',
          borderWidth: 0,
          z: -1
        };
        
        annotations.heavyLabel = {
          type: 'label',
          xValue: 2000,
          yValue: (cs_pace + lt1_range.estimate) / 2,
          backgroundColor: 'rgba(255, 193, 7, 0.8)',
          color: 'black',
          content: 'Heavy Exercise Intensity Domain',
          font: { size: 11, weight: 'bold' },
          padding: 6,
          borderRadius: 4
        };
        
        annotations.moderateBackground = {
          type: 'box',
          xMin: 150,
          xMax: 21600,
          yMin: lt1_range.estimate,
          yMax: maxPace,
          backgroundColor: 'rgba(34, 197, 94, 0.1)',
          borderWidth: 0,
          z: -1
        };
        
        annotations.moderateLabel = {
          type: 'label',
          xValue: 2000,
          yValue: (lt1_range.estimate + maxPace) / 2,
          backgroundColor: 'rgba(34, 197, 94, 0.8)',
          color: 'white',
          content: 'Moderate Exercise Intensity Domain',
          font: { size: 11, weight: 'bold' },
          padding: 6,
          borderRadius: 4
        };
        
        // Threshold lines using new LT1 values
        const pace1500m = calculateRacePace(1500, models);
        const pace5000m = calculateRacePace(5000, models);
        const pace10km = calculateRacePace(10000, models);
        const paceHalfMarathon = calculateRacePace(21097, models);
        const paceMarathon = calculateRacePace(42195, models);
        
        annotations.vo2maxLine = {
          type: 'line',
          yMin: vo2max_pace,
          yMax: vo2max_pace,
          borderColor: 'rgba(147, 51, 234, 0.8)',
          borderWidth: 2,
          borderDash: [5, 5],
          label: {
            display: true,
            content: `VO2max: ${formatPace(vo2max_pace)}/km`,
            position: 'start',
            backgroundColor: 'rgba(147, 51, 234, 0.8)',
            color: 'white',
            padding: 4,
            font: { size: 10 }
          }
        };
        
        annotations.csLine = {
          type: 'line',
          yMin: cs_pace,
          yMax: cs_pace,
          borderColor: 'rgba(239, 68, 68, 0.8)',
          borderWidth: 2,
          borderDash: [5, 5],
          label: {
            display: true,
            content: `CS: ${formatPace(cs_pace)}/km`,
            position: 'start',
            backgroundColor: 'rgba(239, 68, 68, 0.8)',
            color: 'white',
            padding: 4,
            font: { size: 10 }
          }
        };
        
        annotations.lt1Range = {
          type: 'box',
          xMin: 150,
          xMax: 21600,
          yMin: lt1_range.min,
          yMax: lt1_range.max,
          backgroundColor: 'rgba(34, 197, 94, 0.2)',
          borderColor: 'rgba(34, 197, 94, 0.6)',
          borderWidth: 1,
          borderDash: [3, 3],
          z: 0
        };
        
        annotations.lt1Line = {
          type: 'line',
          yMin: lt1_range.estimate,
          yMax: lt1_range.estimate,
          borderColor: 'rgba(34, 197, 94, 0.8)',
          borderWidth: 2,
          borderDash: [10, 5],
          label: {
            display: true,
            content: `LT1: ${formatPace(lt1_range.estimate)}/km`,
            position: 'start',
            backgroundColor: 'rgba(34, 197, 94, 0.8)',
            color: 'white',
            padding: 4,
            font: { size: 10 }
          }
        };
        
        // Race pace annotations
        if (pace1500m) {
          annotations.pace1500Box = {
            type: 'label',
            xValue: 12000,
            yValue: pace1500m,
            backgroundColor: 'rgba(255, 99, 132, 0.8)',
            color: 'white',
            content: `1500m: ${formatPace(pace1500m)}/km`,
            font: { size: 10, weight: 'bold' },
            padding: 6,
            borderRadius: 4
          };
        }
        
        if (pace5000m) {
          annotations.pace5000Box = {
            type: 'label',
            xValue: 12000,
            yValue: pace5000m,
            backgroundColor: 'rgba(54, 162, 235, 0.8)',
            color: 'white',
            content: `5000m: ${formatPace(pace5000m)}/km`,
            font: { size: 10, weight: 'bold' },
            padding: 6,
            borderRadius: 4
          };
        }
        
        if (pace10km) {
          annotations.pace10kmBox = {
            type: 'label',
            xValue: 12000,
            yValue: pace10km,
            backgroundColor: 'rgba(255, 206, 86, 0.8)',
            color: 'black',
            content: `10km: ${formatPace(pace10km)}/km`,
            font: { size: 10, weight: 'bold' },
            padding: 6,
            borderRadius: 4
          };
        }
        
        if (paceHalfMarathon) {
          annotations.paceHalfMarathonBox = {
            type: 'label',
            xValue: 12000,
            yValue: paceHalfMarathon,
            backgroundColor: 'rgba(75, 192, 192, 0.8)',
            color: 'white',
            content: `Mezza: ${formatPace(paceHalfMarathon)}/km`,
            font: { size: 10, weight: 'bold' },
            padding: 6,
            borderRadius: 4
          };
        }
        
        if (paceMarathon) {
          annotations.paceMarathonBox = {
            type: 'label',
            xValue: 12000,
            yValue: paceMarathon,
            backgroundColor: 'rgba(153, 102, 255, 0.8)',
            color: 'white',
            content: `Maratona: ${formatPace(paceMarathon)}/km`,
            font: { size: 10, weight: 'bold' },
            padding: 6,
            borderRadius: 4
          };
        }
      }

      chartRef.current = new Chart(ctx, {
        type: 'line',
        data: {
          datasets: [
            {
              label: 'CS-D\' Model',
              data: curveData.csModel,
              borderColor: '#EF5350', // Red
              backgroundColor: 'rgba(239, 83, 80, 0.1)',
              borderWidth: 3,
              tension: 0.1,
              pointRadius: 0,
              pointHoverRadius: 4,
              fill: false
            },
            {
              label: 'Power Law Model',
              data: curveData.powerLaw,
              borderColor: '#FFCA28', // Yellow
              backgroundColor: 'rgba(255, 202, 40, 0.1)',
              borderWidth: 3,
              tension: 0.1,
              pointRadius: 0,
              pointHoverRadius: 4,
              fill: false
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          interaction: {
            intersect: false,
            mode: 'x'
          },
          plugins: {
            legend: {
              display: true,
              position: 'top',
              labels: {
                color: '#EAEBEE',
                usePointStyle: true,
                padding: 20,
                font: {
                  size: 12
                }
              }
            },
            tooltip: {
              backgroundColor: 'rgba(28, 31, 36, 0.95)',
              titleColor: '#EAEBEE',
              bodyColor: '#EAEBEE',
              borderColor: '#49505B',
              borderWidth: 1,
              cornerRadius: 8,
              displayColors: true,
              callbacks: {
                title: (context) => {
                  const time = context[0]?.parsed?.x;
                  return time ? `Tempo: ${formatTime(time)}` : '';
                },
                afterBody: (context) => {
                  const time = context[0]?.parsed?.x;
                  if (models.valid && time) {
                    const results = [];
                    
                    // Show applicable models for this time
                    if (time <= 1020) {
                      const csResult = calculateCSModel(time, models);
                      if (csResult) {
                        results.push(`CS-D': ${formatPace(csResult.pace)}/km, ${formatDistance(csResult.distance)}`);
                      }
                    }
                    
                    if (time >= 840) {
                      const plResult = calculatePowerLaw(time, models);
                      if (plResult) {
                        results.push(`Power Law: ${formatPace(plResult.pace)}/km, ${formatDistance(plResult.distance)}`);
                      }
                    }
                    
                    return results;
                  }
                  return [];
                }
              }
            },
            annotation: {
              annotations: annotations
            }
          },
          scales: {
            x: {
              type: 'logarithmic',
              min: 150,
              max: 21600,
              title: {
                display: true,
                text: 'Tempo',
                color: '#EAEBEE',
                font: { size: 14, weight: 'bold' }
              },
              ticks: {
                color: '#EAEBEE',
                callback: function(value: any) {
                  const majorTicks = [150, 300, 600, 900, 1200, 1800, 3600, 7200, 21600];
                  return majorTicks.includes(Number(value)) ? formatTime(value, true) : '';
                },
                maxTicksLimit: 15,
                font: { size: 11 }
              },
              grid: {
                color: '#313641'
              }
            },
            y: {
              type: 'linear',
              reverse: true, // Passo più veloce in alto
              min: minPace,
              max: maxPace,
              title: {
                display: true,
                text: 'Passo (min/km)',
                color: '#EAEBEE',
                font: { size: 14, weight: 'bold' }
              },
              ticks: {
                color: '#EAEBEE',
                callback: function(value: any) {
                  return formatPace(Number(value));
                },
                font: { size: 11 }
              },
              grid: {
                color: '#313641'
              }
            }
          }
        }
      });
    };

    loadChart();

    return () => {
      if (chartRef.current) {
        chartRef.current.destroy();
      }
    };
  }, [models]);

  return (
    <div className="w-full h-full">
      <canvas ref={canvasRef} />
    </div>
  );
};

export default RunningChart;

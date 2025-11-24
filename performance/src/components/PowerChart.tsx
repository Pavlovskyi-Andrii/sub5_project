
import React, { useEffect, useRef } from 'react';
import { PowerPrecisionModels } from '../utils/powerPrecisionCalculations';

interface PowerChartProps {
  models: PowerPrecisionModels;
}

const PowerChart: React.FC<PowerChartProps> = ({ models }) => {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const chartRef = useRef<any>(null);

  const formatTime = (seconds: number, brief = false) => {
    const sec = Math.round(Number(seconds));
    if (sec < 60) return `${sec}s`;
    if (sec < 3600) {
      const m = Math.floor(sec / 60);
      const s = sec % 60;
      return brief ? `${m}m` : `${m}m ${s}s`;
    }
    const h = Math.floor(sec / 3600);
    const m = Math.round((sec % 3600) / 60);
    return brief ? `${h}h` : `${h}h ${m}m`;
  };

  const generateCurveData = (models: PowerPrecisionModels) => {
    if (!models || !models.valid) return { apr: [], cp_w: [], powerLaw: [] };
    
    const { cp, wPrime, apr_params, pl_params, pMax } = models;
    const timePoints = [
      ...Array.from({length: 180}, (_, i) => i + 1),
      ...Array.from({length: 34}, (_, i) => 180 + (i + 1) * 30),
      ...Array.from({length: 108}, (_, i) => 1200 + (i + 1) * 180),
    ].filter((v, i, a) => a.indexOf(v) === i);

    const aprData: any[] = [], cpwData: any[] = [], powerLawData: any[] = [];
    
    timePoints.forEach(t => {
      if (t <= 180) {
        aprData.push({ 
          x: t, 
          y: (t === 1) ? pMax : apr_params.po3min + apr_params.apr * Math.exp(-apr_params.k * t) 
        });
      }
      if (t >= 120 && t <= 3600) { // Esteso fino a 60min
        cpwData.push({ x: t, y: cp + wPrime / t });
      }
      if (t >= 180) { // Esteso fino a 3min
        powerLawData.push({ x: t, y: pl_params.S * Math.pow(t, pl_params.E - 1) });
      }
    });
    
    return { apr: aprData, cp_w: cpwData, powerLaw: powerLawData };
  };

  useEffect(() => {
    const loadChart = async () => {
      // Dynamically import Chart.js to avoid SSR issues
      const { Chart, registerables } = await import('chart.js');
      const annotationPlugin = await import('chartjs-plugin-annotation');
      
      Chart.register(...registerables, annotationPlugin.default);

      if (!canvasRef.current) return;
      
      // Destroy existing chart
      if (chartRef.current) {
        chartRef.current.destroy();
      }

      const ctx = canvasRef.current.getContext('2d');
      if (!ctx) return;

      const curveData = generateCurveData(models);
      
      // Create intensity domain annotations
      const annotations: any = {};
      
      if (models.valid) {
        const { cp, pMax, apr_params, lt1_range, pl_params } = models;
        const po3min = apr_params.po3min;
        
        // Calculate CP range (from 50min power law estimate to CP)
        const cp50minEstimate = pl_params.S * Math.pow(3000, pl_params.E - 1); // 50min = 3000s
        
        // Intensity domains (horizontal zones) - corrected positioning
        annotations.extremeDomain = {
          type: 'box',
          xMin: 1,
          xMax: 21600,
          yMin: po3min,
          yMax: pMax! * 1.2,
          backgroundColor: 'rgba(147, 51, 234, 0.15)', // Purple
          borderWidth: 0,
          label: {
            display: true,
            content: 'Extreme Exercise Intensity Domain',
            position: 'center',
            color: 'rgba(147, 51, 234, 0.8)',
            font: { size: 12 }
          }
        };
        
        annotations.severeDomain = {
          type: 'box',
          xMin: 1,
          xMax: 21600,
          yMin: cp!,
          yMax: po3min,
          backgroundColor: 'rgba(239, 68, 68, 0.15)', // Red
          borderWidth: 0,
          label: {
            display: true,
            content: 'Severe Exercise Intensity Domain',
            position: 'center',
            color: 'rgba(239, 68, 68, 0.8)',
            font: { size: 12 }
          }
        };
        
        annotations.heavyDomain = {
          type: 'box',
          xMin: 1,
          xMax: 21600,
          yMin: lt1_range!.estimate,
          yMax: cp!,
          backgroundColor: 'rgba(234, 179, 8, 0.15)', // Yellow
          borderWidth: 0,
          label: {
            display: true,
            content: 'Heavy Exercise Intensity Domain',
            position: 'center',
            color: 'rgba(234, 179, 8, 0.8)',
            font: { size: 12 }
          }
        };
        
        annotations.moderateDomain = {
          type: 'box',
          xMin: 1,
          xMax: 21600,
          yMin: Math.max(50, lt1_range!.estimate * 0.5),
          yMax: lt1_range!.estimate,
          backgroundColor: 'rgba(34, 197, 94, 0.15)', // Green
          borderWidth: 0,
          label: {
            display: true,
            content: 'Moderate Exercise Intensity Domain',
            position: 'center',
            color: 'rgba(34, 197, 94, 0.8)',
            font: { size: 12 }
          }
        };
        
        // Threshold lines
        annotations.pmaxLine = {
          type: 'line',
          yMin: pMax,
          yMax: pMax,
          borderColor: 'rgba(255, 255, 255, 0.8)',
          borderWidth: 2,
          borderDash: [5, 5],
          label: {
            display: true,
            content: `Pmax: ${pMax}W`,
            position: 'end',
            backgroundColor: 'rgba(0, 0, 0, 0.8)',
            color: 'white',
            padding: 4
          }
        };
        
        annotations.p3minLine = {
          type: 'line',
          yMin: po3min,
          yMax: po3min,
          borderColor: 'rgba(192, 114, 224, 0.8)',
          borderWidth: 2,
          borderDash: [5, 5],
          label: {
            display: true,
            content: `P3min: ${Math.round(po3min)}W`,
            position: 'end',
            backgroundColor: 'rgba(0, 0, 0, 0.8)',
            color: 'white',
            padding: 4
          }
        };
        
        annotations.cpLine = {
          type: 'line',
          yMin: cp,
          yMax: cp,
          borderColor: 'rgba(239, 83, 80, 0.8)',
          borderWidth: 2,
          borderDash: [5, 5],
          label: {
            display: true,
            content: `CP: ${Math.round(cp!)}W`,
            position: 'end',
            backgroundColor: 'rgba(0, 0, 0, 0.8)',
            color: 'white',
            padding: 4
          }
        };
        
        annotations.lt1Line = {
          type: 'line',
          yMin: lt1_range!.estimate,
          yMax: lt1_range!.estimate,
          borderColor: 'rgba(76, 175, 80, 0.8)',
          borderWidth: 2,
          borderDash: [10, 5],
          label: {
            display: true,
            content: `LT1 (stima): ${Math.round(lt1_range!.estimate)}W`,
            position: 'end',
            backgroundColor: 'rgba(0, 0, 0, 0.8)',
            color: 'white',
            padding: 4
          }
        };
        
        // LT1 range box (green)
        annotations.lt1RangeBox = {
          type: 'box',
          xMin: 1,
          xMax: 21600,
          yMin: lt1_range!.min,
          yMax: lt1_range!.max,
          backgroundColor: 'rgba(76, 175, 80, 0.2)',
          borderColor: 'rgba(76, 175, 80, 0.6)',
          borderWidth: 1,
          borderDash: [3, 3]
        };
        
        // CP range box (orange) - from 50min power law estimate to CP
        annotations.cpRangeBox = {
          type: 'box',
          xMin: 1,
          xMax: 21600,
          yMin: cp50minEstimate,
          yMax: cp!,
          backgroundColor: 'rgba(251, 146, 60, 0.2)',
          borderColor: 'rgba(251, 146, 60, 0.6)',
          borderWidth: 1,
          borderDash: [3, 3]
        };
      }

      chartRef.current = new Chart(ctx, {
        type: 'line',
        data: {
          datasets: [
            {
              label: 'APR Model',
              data: curveData.apr,
              borderColor: '#C072E0', // Purple
              backgroundColor: 'rgba(192, 114, 224, 0.1)',
              borderWidth: 3,
              tension: 0.1,
              pointRadius: 0,
              pointHoverRadius: 4,
              fill: false
            },
            {
              label: 'CP-W\' Model',
              data: curveData.cp_w,
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
              displayColors: false,
              callbacks: {
                label: () => null,
                title: (context) => {
                  const time = context[0]?.parsed?.x;
                  return time ? `Tempo: ${formatTime(time)}` : '';
                },
                afterBody: (context) => {
                  const time = context[0]?.parsed?.x;
                  if (models.valid && time) {
                    const { cp, wPrime, apr_params, pl_params, pMax } = models;
                    const results = [];
                    
                    // Show all applicable models for this time
                    if (time < 180) {
                      const aprPower = (time === 1) ? pMax : apr_params.po3min + apr_params.apr * Math.exp(-apr_params.k * time);
                      results.push(`APR Model: ${Math.round(aprPower)}W`);
                    }
                    
                    if (time >= 120 && time <= 3600) {
                      const cpPower = cp + wPrime / time;
                      results.push(`CP-W' Model: ${Math.round(cpPower)}W`);
                    }
                    
                    if (time >= 180) {
                      const plPower = pl_params.S * Math.pow(time, pl_params.E - 1);
                      results.push(`Power Law Model: ${Math.round(plPower)}W`);
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
              min: 1,
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
                  const majorTicks = [1, 5, 20, 60, 180, 300, 600, 1200, 3600, 7200, 21600];
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
              type: 'logarithmic',
              min: models.valid ? Math.max(50, models.cp * 0.5) : 50,
              max: models.valid ? Math.ceil((models.pMax * 1.05) / 100) * 100 : 1000,
              title: {
                display: true,
                text: 'Potenza (W)',
                color: '#EAEBEE',
                font: { size: 14, weight: 'bold' }
              },
              ticks: {
                color: '#EAEBEE',
                callback: function(value: any) {
                  return Math.round(Number(value));
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
    <div id="power-chart-container" className="w-full h-full">
      <canvas ref={canvasRef} />
    </div>
  );
};

export default PowerChart;

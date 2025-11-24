
import React from 'react';
import { AthleteStats } from '@/utils/athleteCardCalculations';
import { RunningAthleteStats } from '@/utils/runningAthleteCalculations';

interface StatsSectionProps {
  stats: AthleteStats | RunningAthleteStats;
}

const StatsSection: React.FC<StatsSectionProps> = ({ stats }) => {
  const getStatName = (key: string) => {
    switch (key) {
      case 'NM': return 'NEURO';
      case 'LACT': return 'ANAEROBICO';
      case 'VO2': return 'VO2MAX';
      case 'THR': return 'SOGLIA';
      case 'STA': return 'STAMINA';
      case '1500m': return '1500M';
      case '5000m': return '5000M';
      case '10km': return '10KM';
      case 'Mezza': return 'MEZZA';
      case 'Mara': return 'MARA';
      default: return key;
    }
  };

  const getStatDuration = (key: string) => {
    switch (key) {
      case 'NM': return '5s';
      case 'LACT': return '1m';
      case 'VO2': return '5m';
      case 'THR': return '20m';
      case 'STA': return '60m';
      case '1500m': return '3-4min';
      case '5000m': return '12-16min';
      case '10km': return '26-35min';
      case 'Mezza': return '1h 5-15min';
      case 'Mara': return '2h 5-35min';
      default: return '';
    }
  };

  return (
    <div className="space-y-4">
      {Object.entries(stats).map(([key, value]) => (
        <div key={key} className="flex justify-between items-center">
          <div>
            <div 
              className="font-normal uppercase tracking-wider"
              style={{ 
                fontSize: '20px',
                color: '#BDBDBD',
                fontFamily: 'Inter, sans-serif',
                marginBottom: '2px'
              }}
            >
              {getStatName(key)}
            </div>
            <div 
              className="font-normal"
              style={{ 
                fontSize: '14px',
                color: '#757575',
                fontFamily: 'Inter, sans-serif'
              }}
            >
              {getStatDuration(key)}
            </div>
          </div>
          <div 
            className="font-bold"
            style={{ 
              fontSize: '36px',
              color: '#FFFFFF',
              fontFamily: 'Inter, sans-serif'
            }}
          >
            {value}
          </div>
        </div>
      ))}
    </div>
  );
};

export default StatsSection;

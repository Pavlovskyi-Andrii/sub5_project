import React from 'react';
import InstagramStoryTemplate from './InstagramStoryTemplate';
import { RunningAthleteProfile } from '@/utils/runningAthleteCalculations';
import { Country } from '@/utils/countries';
import avatarRunnerMaleBrown from '@/assets/avatar-runner-male-brown.png';

const JakobTest: React.FC = () => {
  // Jakob Ingebrigtsen test data - GOAT level runner
  const testProfile: RunningAthleteProfile = {
    stats: {
      '1500m': 93,
      '5000m': 95,
      '10km': 91,  
      'Mezza': 88,
      'Mara': 87
    },
    overallRating: 96, // GOAT level (95+)
    profileType: 'MEZZOFONDISTA',
    rarityTier: 'HERO' as any // This will map to GOAT in the new system
  };

  const norwayCountry: Country = {
    name: 'Norvegia',
    code: 'NO',
    flag: '🇳🇴'
  };

  return (
    <div className="flex justify-center items-center min-h-screen bg-gray-900 p-4">
      <div style={{ transform: 'scale(0.3)', transformOrigin: 'center' }}>
        <InstagramStoryTemplate
          athleteName="Jakob Ingebrigtsen"
          country={norwayCountry}
          profileImage={avatarRunnerMaleBrown}
          profile={testProfile}
          isRunning={true}
        />
      </div>
    </div>
  );
};

export default JakobTest;
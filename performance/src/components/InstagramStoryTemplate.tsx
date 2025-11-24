
import React from 'react';
import { AthleteProfile } from '@/utils/athleteCardCalculations';
import { RunningAthleteProfile } from '@/utils/runningAthleteCalculations';
import { Country } from '@/utils/countries';
import { getProfileTheme } from '@/utils/profileThemes';
import { getRarityLevel, getAvatarBorderStyles } from '@/utils/raritySystem';
import RarityEffects from './RarityEffects';
import ScoreDisplay from './ScoreDisplay';
import SpiderChart from './SpiderChart';

interface InstagramStoryTemplateProps {
  athleteName: string;
  country: Country;
  profileImage: string;
  profile: AthleteProfile | RunningAthleteProfile;
  isRunning?: boolean;
}

const InstagramStoryTemplate: React.FC<InstagramStoryTemplateProps> = ({
  athleteName,
  country,
  profileImage,
  profile,
  isRunning = false
}) => {
  const theme = getProfileTheme(profile.profileType, isRunning);
  const rarityLevel = getRarityLevel(profile.overallRating);
  const avatarBorderStyles = getAvatarBorderStyles(rarityLevel, theme);

  return (
    <div 
      className="relative"
      style={{ 
        width: '1080px', 
        height: '1920px',
        background: 'linear-gradient(180deg, #0A0A0A 0%, #004d40 100%)'
      }}
    >
      {/* Effetti di rarità */}
      <RarityEffects rarityLevel={rarityLevel} theme={theme} />
      
      {/* Topographic pattern overlay */}
      <div 
        className="absolute inset-0"
        style={{
          opacity: 0.03,
          backgroundImage: `
            repeating-linear-gradient(
              45deg,
              transparent,
              transparent 60px,
              #333333 60px,
              #333333 61px
            ),
            repeating-linear-gradient(
              -45deg,
              transparent,
              transparent 60px,
              #333333 60px,
              #333333 61px
            )
          `
        }}
      />

      {/* Two Column Layout */}
      <div className="absolute inset-0 flex" style={{ padding: '120px 80px 80px 80px' }}>
        
        {/* Left Column - 45% width */}
        <div className="flex flex-col" style={{ width: '45%', paddingRight: '40px' }}>
          
          {/* Avatar positioned at 1/3 from top */}
          <div className="flex flex-col items-start" style={{ marginTop: '400px' }}>
            <div 
              className="rounded-full border-4 overflow-hidden"
              style={{
                width: '320px',
                height: '320px',
                ...avatarBorderStyles
              }}
            >
              <img 
                src={profileImage} 
                alt={athleteName}
                className="w-full h-full object-cover"
              />
            </div>
            
            {/* Athlete Data - aligned to avatar left edge */}
            <div style={{ marginTop: '40px', width: '320px' }}>
              <div 
                className="font-black leading-tight"
                style={{ 
                  fontSize: '52px',
                  color: '#FFFFFF',
                  fontFamily: 'Inter, sans-serif',
                  marginBottom: '4px',
                  textAlign: 'left'
                }}
              >
                {athleteName.split(' ')[0]}
              </div>
              <div 
                className="font-black leading-tight"
                style={{ 
                  fontSize: '52px',
                  color: '#FFFFFF',
                  fontFamily: 'Inter, sans-serif',
                  marginBottom: '12px',
                  textAlign: 'left'
                }}
              >
                {athleteName.split(' ').slice(1).join(' ')}
              </div>
              
              <div className="flex items-center" style={{ gap: '12px' }}>
                <img 
                  src={`https://flagcdn.com/w40/${country.code.toLowerCase()}.png`}
                  alt={country.name}
                  className="rounded shadow-md"
                  style={{ width: '32px', height: '24px' }}
                />
                <span 
                  className="font-normal"
                  style={{ 
                    fontSize: '28px',
                    color: '#BDBDBD',
                    fontFamily: 'Inter, sans-serif'
                  }}
                >
                  {country.name}
                </span>
              </div>
            </div>
          </div>
        </div>

        {/* Right Column - 55% width */}
        <div className="flex flex-col" style={{ width: '55%', paddingTop: '400px' }}>
          
          {/* Main Score with significant spacing */}
          <ScoreDisplay
            score={profile.overallRating}
            profileType={profile.profileType}
            rarityLevel={rarityLevel}
            theme={theme}
            isRunning={isRunning}
          />

          {/* Divider Line */}
          <div 
            style={{
              width: '100%',
              height: '1px',
              background: 'rgba(189, 189, 189, 0.3)',
              marginBottom: '40px'
            }}
          />

          {/* Spider Chart - positioned significantly lower */}
          <SpiderChart 
            stats={profile.stats}
            theme={theme}
            rarityLevel={rarityLevel}
            isRunning={isRunning}
          />
        </div>
      </div>

      {/* Bottom Info Box */}
      <div 
        className="absolute bottom-0 left-0 right-0 mx-20 mb-20 rounded-2xl text-center"
        style={{
          backgroundColor: 'rgba(0, 61, 51, 0.8)',
          backdropFilter: 'blur(10px)',
          padding: '32px 40px'
        }}
      >
        <div 
          className="font-bold mb-2"
          style={{ 
            fontSize: '32px',
            color: '#FFFFFF',
            fontFamily: 'Inter, sans-serif'
          }}
        >
          Qual è il tuo profilo?
        </div>
        <div 
          className="font-normal mb-4"
          style={{ 
            fontSize: '28px',
            color: theme.primaryColor,
            fontFamily: 'Inter, sans-serif'
          }}
        >
          Scoprilo su miccoach.it
        </div>
        <div 
          className="font-normal"
          style={{ 
            fontSize: '20px',
            color: '#BDBDBD',
            fontFamily: 'Inter, sans-serif'
          }}
        >
          #MicCoachLab • @MicCoach
        </div>
      </div>

      {/* CSS per animazioni delle particelle */}
      <style>{`
        @keyframes float {
          0%, 100% { transform: translateY(0px) translateX(0px); }
          25% { transform: translateY(-12px) translateX(6px); }
          50% { transform: translateY(-8px) translateX(-4px); }
          75% { transform: translateY(-18px) translateX(3px); }
        }
        
        @keyframes pulse {
          0%, 100% { opacity: 0.3; transform: scale(1); }
          50% { opacity: 0.6; transform: scale(1.1); }
        }
      `}</style>
    </div>
  );
};

export default InstagramStoryTemplate;

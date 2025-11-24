
import React from 'react';
import { getScoreStyles, RarityLevel } from '@/utils/raritySystem';
import { ProfileTheme } from '@/utils/profileThemes';
import { useTranslation } from 'react-i18next';

interface ScoreDisplayProps {
  score: number;
  profileType: string;
  rarityLevel: RarityLevel;
  theme: ProfileTheme;
  isRunning?: boolean;
}

const ScoreDisplay: React.FC<ScoreDisplayProps> = ({
  score,
  profileType,
  rarityLevel,
  theme,
  isRunning = false
}) => {
  const { t } = useTranslation();
  const scoreStyles = getScoreStyles(rarityLevel, theme);

  const getProfileDisplayName = () => {
    // Create profile key based on original profileType
    const profileKey = profileType.toLowerCase()
      .replace(/\s+/g, '_')
      .replace(/-/g, '_');
    
    // Try to get translation with appropriate prefix
    const keyPrefix = isRunning ? 'common.profiles.runner_' : 'common.profiles.';
    const translationKey = `${keyPrefix}${profileKey}`;
    
    // Return translated profile name or fallback to original
    return t(translationKey, profileType);
  };

  return (
    <div className="mb-12 relative">
      <div className="relative inline-block">
        {/* MAIN SCORE - Golden color for LEGGENDA+ levels */}
        <div 
          className="font-black leading-none"
          style={{ 
            fontSize: '240px',
            fontFamily: 'Impact, Arial Black, sans-serif',
            marginBottom: '40px',
            ...scoreStyles
          }}
        >
          {score}
        </div>
      </div>
      
      {/* PROFILE TYPE - Using category accent color */}
      <div 
        className="font-bold"
        style={{ 
          fontSize: '48px',
          color: theme.primaryColor,
          fontFamily: 'Inter, sans-serif',
          letterSpacing: '2px',
          textTransform: 'uppercase'
        }}
      >
        {getProfileDisplayName()}
      </div>
    </div>
  );
};

export default ScoreDisplay;

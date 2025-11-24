
import React from 'react';
import { RarityLevel, shouldShowParticles } from '@/utils/raritySystem';
import { ProfileTheme } from '@/utils/profileThemes';

interface RarityEffectsProps {
  rarityLevel: RarityLevel;
  theme: ProfileTheme;
}

const RarityEffects: React.FC<RarityEffectsProps> = ({ rarityLevel, theme }) => {
  const showParticles = shouldShowParticles(rarityLevel);

  return (
    <>
      {/* Particelle dorate fluttuanti per GOAT */}
      {showParticles && (
        <div 
          className="absolute inset-0 pointer-events-none overflow-hidden" 
          style={{ zIndex: 2 }}
        >
          {/* Particelle di luce dorata */}
          {[...Array(15)].map((_, i) => (
            <div
              key={i}
              className="absolute w-2 h-2 rounded-full opacity-70"
              style={{
                background: `radial-gradient(circle, #FFD700, #FFD700CC)`,
                left: `${10 + (i * 6)}%`,
                top: `${20 + (i % 5) * 12}%`,
                animation: `float ${3 + (i % 3)}s ease-in-out infinite`,
                animationDelay: `${i * 0.4}s`,
                filter: 'blur(0.5px)'
              }}
            />
          ))}
          
          {/* Raggi di luce dorata dal punteggio */}
          <div 
            className="absolute opacity-40"
            style={{
              right: '10%',
              top: '15%',
              width: '400px',
              height: '400px',
              background: `radial-gradient(circle, #FFD70060 0%, transparent 70%)`,
              borderRadius: '50%',
              animation: 'pulse 5s ease-in-out infinite',
              zIndex: 1
            }}
          />
        </div>
      )}
    </>
  );
};

export default RarityEffects;

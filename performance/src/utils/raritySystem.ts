
import { ProfileTheme } from './profileThemes';

export type RarityLevel = 'BASE' | 'ELITE' | 'LEGGENDA' | 'GOAT';

// Determina il livello di rarità
export const getRarityLevel = (overallRating: number): RarityLevel => {
  if (overallRating >= 95) return 'GOAT';
  if (overallRating >= 90) return 'LEGGENDA';
  if (overallRating >= 85) return 'ELITE';
  return 'BASE';
};

// Stili per il numero del punteggio - ORO per LEGGENDA+
export const getScoreStyles = (rarityLevel: RarityLevel, theme: ProfileTheme) => {
  switch (rarityLevel) {
    case 'GOAT':
    case 'LEGGENDA':
      return {
        color: '#FFD700', // ORO VIBRANTE per livelli alti
        fontWeight: 900,
        position: 'relative' as const,
        zIndex: 20,
        opacity: 1
      };
    case 'ELITE':
    case 'BASE':
    default:
      return {
        color: '#FFFFFF', // BIANCO per livelli standard
        fontWeight: 900,
        position: 'relative' as const,
        zIndex: 20,
        opacity: 1
      };
  }
};

// Stili per il bordo dell'avatar - Colore categoria per ELITE+
export const getAvatarBorderStyles = (rarityLevel: RarityLevel, theme: ProfileTheme) => {
  if (rarityLevel === 'BASE') {
    return {
      borderColor: theme.primaryColor,
      boxShadow: `0 0 120px ${theme.primaryColor}60, 0 0 60px ${theme.primaryColor}40`,
    };
  }
  
  // ELITE+ levels get enhanced border with category accent color
  return {
    borderColor: theme.primaryColor,
    boxShadow: `0 0 120px ${theme.primaryColor}80, 0 0 60px ${theme.primaryColor}60`,
    filter: `drop-shadow(0 0 40px ${theme.primaryColor}70)`
  };
};

// Particelle per livello GOAT
export const shouldShowParticles = (rarityLevel: RarityLevel): boolean => {
  return rarityLevel === 'GOAT';
};

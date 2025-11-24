// Gaming-style color palettes and themes for athlete cards

export interface ProfileTheme {
  primaryColor: string;
  secondaryColor: string;
  accentColor: string;
  gradient: string;
  glowColor: string;
  particleColors: readonly string[];
  patternOpacity: number;
}

export const CYCLING_THEMES = {
  SPRINTER: {
    primaryColor: '#DC2626', // Red-600
    secondaryColor: '#F59E0B', // Amber-500
    accentColor: '#FEF3C7', // Amber-100
    gradient: 'radial-gradient(circle at 30% 20%, rgba(220, 38, 38, 0.2) 0%, transparent 50%), radial-gradient(circle at 70% 80%, rgba(245, 158, 11, 0.2) 0%, transparent 50%), linear-gradient(135deg, #0f0f0f 0%, #1a0e0e 50%, #0f0f0f 100%)',
    glowColor: '#DC2626',
    particleColors: ['#DC2626', '#F59E0B', '#EF4444'],
    patternOpacity: 0.08
  },
  CLIMBER: {
    primaryColor: '#059669', // Emerald-600
    secondaryColor: '#0EA5E9', // Sky-500
    accentColor: '#A7F3D0', // Emerald-200
    gradient: 'radial-gradient(circle at 20% 30%, rgba(5, 150, 105, 0.15) 0%, transparent 50%), radial-gradient(circle at 80% 70%, rgba(14, 165, 233, 0.15) 0%, transparent 50%), linear-gradient(135deg, #0f0f0f 0%, #0a1a1a 50%, #0f0f0f 100%)',
    glowColor: '#059669',
    particleColors: ['#059669', '#0EA5E9', '#10B981'],
    patternOpacity: 0.1
  },
  'TIME TRIALIST': {
    primaryColor: '#2563EB', // Blue-600
    secondaryColor: '#6B7280', // Gray-500
    accentColor: '#DBEAFE', // Blue-100
    gradient: 'radial-gradient(circle at 40% 30%, rgba(37, 99, 235, 0.2) 0%, transparent 50%), radial-gradient(circle at 60% 70%, rgba(107, 114, 128, 0.1) 0%, transparent 50%), linear-gradient(135deg, #0f0f0f 0%, #0f1419 50%, #0f0f0f 100%)',
    glowColor: '#2563EB',
    particleColors: ['#2563EB', '#3B82F6', '#60A5FA'],
    patternOpacity: 0.06
  },
  PUNCHEUR: {
    primaryColor: '#EA580C', // Orange-600
    secondaryColor: '#FACC15', // Yellow-400
    accentColor: '#FED7AA', // Orange-200
    gradient: 'radial-gradient(circle at 25% 25%, rgba(234, 88, 12, 0.2) 0%, transparent 50%), radial-gradient(circle at 75% 75%, rgba(250, 204, 21, 0.15) 0%, transparent 50%), linear-gradient(135deg, #0f0f0f 0%, #1a1209 50%, #0f0f0f 100%)',
    glowColor: '#EA580C',
    particleColors: ['#EA580C', '#FACC15', '#FB923C'],
    patternOpacity: 0.09
  },
  'ALL-ROUNDER': {
    primaryColor: '#7C3AED', // Violet-600
    secondaryColor: '#EC4899', // Pink-500
    accentColor: '#DDD6FE', // Violet-200
    gradient: 'radial-gradient(circle at 20% 20%, rgba(124, 58, 237, 0.15) 0%, transparent 40%), radial-gradient(circle at 50% 50%, rgba(236, 72, 153, 0.1) 0%, transparent 40%), radial-gradient(circle at 80% 80%, rgba(5, 150, 105, 0.1) 0%, transparent 40%), linear-gradient(135deg, #0f0f0f 0%, #1a0f1a 50%, #0f0f0f 100%)',
    glowColor: '#7C3AED',
    particleColors: ['#7C3AED', '#EC4899', '#059669', '#F59E0B'],
    patternOpacity: 0.07
  }
} as const;

export const RUNNING_THEMES = {
  SPRINTER: {
    primaryColor: '#DC2626', // Red-600
    secondaryColor: '#F59E0B', // Amber-500  
    accentColor: '#FEF3C7', // Amber-100
    gradient: 'radial-gradient(circle at 30% 20%, rgba(220, 38, 38, 0.2) 0%, transparent 50%), radial-gradient(circle at 70% 80%, rgba(245, 158, 11, 0.2) 0%, transparent 50%), linear-gradient(135deg, #0f0f0f 0%, #1a0e0e 50%, #0f0f0f 100%)',
    glowColor: '#DC2626',
    particleColors: ['#DC2626', '#F59E0B', '#EF4444'],
    patternOpacity: 0.08
  },
  MEZZOFONDISTA: {
    primaryColor: '#4DB6AC', // Verde Acqua (richiesto per Jakob)
    secondaryColor: '#0284C7', // Sky-600
    accentColor: '#B2DFDB', // Teal-100
    gradient: 'radial-gradient(circle at 25% 30%, rgba(77, 182, 172, 0.18) 0%, transparent 50%), radial-gradient(circle at 75% 70%, rgba(2, 132, 199, 0.15) 0%, transparent 50%), linear-gradient(135deg, #0f0f0f 0%, #0a1717 50%, #0f0f0f 100%)',
    glowColor: '#4DB6AC',
    particleColors: ['#4DB6AC', '#0284C7', '#26A69A'],
    patternOpacity: 0.09
  },
  ENDURANCE: {
    primaryColor: '#0F766E', // Teal-700
    secondaryColor: '#0284C7', // Sky-600
    accentColor: '#99F6E4', // Teal-200
    gradient: 'radial-gradient(circle at 25% 30%, rgba(15, 118, 110, 0.18) 0%, transparent 50%), radial-gradient(circle at 75% 70%, rgba(2, 132, 199, 0.15) 0%, transparent 50%), linear-gradient(135deg, #0f0f0f 0%, #0a1717 50%, #0f0f0f 100%)',
    glowColor: '#0F766E',
    particleColors: ['#0F766E', '#0284C7', '#14B8A6'],
    patternOpacity: 0.09
  },
  BALANCED: {
    primaryColor: '#7C3AED', // Violet-600
    secondaryColor: '#EC4899', // Pink-500
    accentColor: '#DDD6FE', // Violet-200
    gradient: 'radial-gradient(circle at 20% 20%, rgba(124, 58, 237, 0.15) 0%, transparent 40%), radial-gradient(circle at 50% 50%, rgba(236, 72, 153, 0.1) 0%, transparent 40%), radial-gradient(circle at 80% 80%, rgba(15, 118, 110, 0.1) 0%, transparent 40%), linear-gradient(135deg, #0f0f0f 0%, #1a0f1a 50%, #0f0f0f 100%)',
    glowColor: '#7C3AED',
    particleColors: ['#7C3AED', '#EC4899', '#0F766E', '#F59E0B'],
    patternOpacity: 0.07
  }
} as const;

export const getProfileTheme = (profileType: string, isRunning: boolean = false): ProfileTheme => {
  if (isRunning) {
    // Map running profiles to running themes
    switch (profileType) {
      case 'SPRINTER':
      case 'MEZZOFONDISTA VELOCE':
        return RUNNING_THEMES.SPRINTER;
      case 'MEZZOFONDISTA':
        return RUNNING_THEMES.MEZZOFONDISTA; // Verde Acqua theme for Jakob test
      case 'FONDISTA':
      case 'SPECIALISTA MEZZA':
      case 'MARATONETA':
        return RUNNING_THEMES.ENDURANCE;
      default:
        return RUNNING_THEMES.BALANCED;
    }
  }
  
  return CYCLING_THEMES[profileType as keyof typeof CYCLING_THEMES] || CYCLING_THEMES['ALL-ROUNDER'];
};
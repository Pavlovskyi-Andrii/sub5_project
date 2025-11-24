import { PrecisionModels } from './precisionCalculations';

// Record del mondo come benchmark differenziati per sesso (in secondi)
export const MALE_BENCHMARKS_RUNNING = {
  '1500m': 206,   // Hicham El Guerrouj 3:26:00
  '5000m': 769,   // Joshua Cheptegei 12:49
  '10km': 1584,   // Rhonex Kipruto 26:24
  'Mezza': 3451,  // Jacob Kiplimo 57:31
  'Mara': 7235    // Kelvin Kiptum 2:00:35
} as const;

export const FEMALE_BENCHMARKS_RUNNING = {
  '1500m': 229,   // Faith Kipyegon 3:49
  '5000m': 853,   // Letesenbet Gidey 14:13
  '10km': 1801,   // Beatrice Chebet 30:01
  'Mezza': 3916,  // Letesenbet Gidey 1:05:16
  'Mara': 8221    // Tigst Assefa 2:17:01
} as const;

// Soglie di peso per la classificazione profilo
export const WEIGHT_THRESHOLDS = {
  male: 70,    // kg - confine per maschi
  female: 53   // kg - confine per femmine
};

export type Gender = 'male' | 'female';

export interface RunningAthleteStats {
  '1500m': number;
  '5000m': number;
  '10km': number;
  'Mezza': number;
  'Mara': number;
}

export type RarityTier = 'ALIEN' | 'HERO' | 'PRO' | 'ELITE' | 'STANDARD';

export interface RunningAthleteProfile {
  stats: RunningAthleteStats;
  overallRating: number;
  profileType: 'MEZZOFONDISTA VELOCE' | 'MEZZOFONDISTA' | 'FONDISTA' | 'SPECIALISTA MEZZA' | 'MARATONETA' | 'ATLETA COMPLETO';
  rarityTier: RarityTier;
}

// Metadati per le statistiche del running
export const RUNNING_STAT_METADATA = {
  '1500m': { label: '1500m', duration: '3-4min' },
  '5000m': { label: '5000m', duration: '12-16min' },
  '10km': { label: '10km', duration: '26-35min' },
  'Mezza': { label: 'Mezza', duration: '1h 5-15min' },
  'Mara': { label: 'Mara', duration: '2h 5-35min' }
} as const;

/**
 * Calcola i tempi previsti dell'atleta per le 5 distanze usando la logica ibrida
 */
export const calculateRunningAthleteTimes = (models: PrecisionModels): RunningAthleteStats => {
  const { cs, dPrime, pl_params } = models;
  const { A, k } = pl_params;
  
  console.log('=== RUNNING ATHLETE TIMES CALCULATION ===');
  console.log('CS:', cs, 'm/s, D\':', dPrime, 'm');
  console.log('Power Law params: A=', A, ', k=', k);

  // Calcola i tempi per ogni distanza
  const calculateTime = (distance: number): number => {
    // Stima iniziale del tempo usando il modello CS-D'
    const estimatedTime = (distance - dPrime) / cs;
    
    console.log(`Distance ${distance}m: Estimated time from CS-D': ${estimatedTime}s`);
    
    // Se il tempo stimato è <= 16 minuti (960s), usa CS-D'
    if (estimatedTime <= 960) {
      const finalTime = Math.max(estimatedTime, 0);
      console.log(`Using CS-D' model: ${finalTime}s`);
      return finalTime;
    } else {
      // Altrimenti usa Power Law: V = A * t^k => t = (V/A)^(1/k)
      // dove V = distance/time, quindi distance/time = A * t^k
      // Risolviamo per t: distance = A * t^(k+1) => t = (distance/A)^(1/(k+1))
      const velocity = distance / estimatedTime;
      const powerLawTime = Math.pow(distance / A, 1 / (k + 1));
      console.log(`Using Power Law model: ${powerLawTime}s`);
      return powerLawTime;
    }
  };

  return {
    '1500m': calculateTime(1500),
    '5000m': calculateTime(5000),
    '10km': calculateTime(10000),
    'Mezza': calculateTime(21097.5), // Mezza maratona
    'Mara': calculateTime(42195)     // Maratona
  };
};

/**
 * Calcola i punteggi per ogni statistica basandosi sui record del mondo
 */
export const calculateRunningAthleteScores = (
  athleteTimes: RunningAthleteStats, 
  gender: Gender
): RunningAthleteStats => {
  const benchmarks = gender === 'male' ? MALE_BENCHMARKS_RUNNING : FEMALE_BENCHMARKS_RUNNING;
  
  console.log('=== RUNNING SCORES CALCULATION ===');
  console.log('Benchmarks:', benchmarks);
  console.log('Athlete times:', athleteTimes);
  
  const scores = {
    '1500m': Math.round((benchmarks['1500m'] / athleteTimes['1500m']) * 95),
    '5000m': Math.round((benchmarks['5000m'] / athleteTimes['5000m']) * 95),
    '10km': Math.round((benchmarks['10km'] / athleteTimes['10km']) * 95),
    'Mezza': Math.round((benchmarks['Mezza'] / athleteTimes['Mezza']) * 95),
    'Mara': Math.round((benchmarks['Mara'] / athleteTimes['Mara']) * 95)
  };
  
  console.log('Calculated scores:', scores);
  return scores;
};

/**
 * Calcola il rating complessivo (media dei 5 punteggi)
 */
export const calculateRunningOverallRating = (scores: RunningAthleteStats): number => {
  const values = Object.values(scores);
  const sum = values.reduce((acc, val) => acc + val, 0);
  return Math.round(sum / values.length);
};

/**
 * Determina il tier di rarità basato sul rating complessivo
 */
export const determineRunningRarityTier = (overallRating: number): RarityTier => {
  if (overallRating >= 100) return 'ALIEN';
  if (overallRating >= 95) return 'HERO';
  if (overallRating >= 90) return 'PRO';
  if (overallRating >= 85) return 'ELITE';
  return 'STANDARD';
};

/**
 * Funzione helper per mappare la specialità di picco al profilo atleta (running)
 */
const getRunningProfileFromSpecialty = (
  peakSpecialty: string
): 'MEZZOFONDISTA VELOCE' | 'MEZZOFONDISTA' | 'FONDISTA' | 'SPECIALISTA MEZZA' | 'MARATONETA' | 'ATLETA COMPLETO' => {
  switch (peakSpecialty) {
    case '1500m':
      return 'MEZZOFONDISTA VELOCE';
    case '5000m':
      return 'MEZZOFONDISTA';
    case '10km':
      return 'FONDISTA';
    case 'Mezza':
      return 'SPECIALISTA MEZZA';
    case 'Mara':
      return 'MARATONETA';
    default:
      return 'ATLETA COMPLETO';
  }
};

/**
 * Determina il profilo dell'atleta runner basandosi sulla logica Z-Score
 */
export const determineRunningAthleteProfile = (
  scores: RunningAthleteStats,
  overallRating: number,
  gender: Gender = 'male',
  weight: number = 70
): 'MEZZOFONDISTA VELOCE' | 'MEZZOFONDISTA' | 'FONDISTA' | 'SPECIALISTA MEZZA' | 'MARATONETA' | 'ATLETA COMPLETO' => {
  const values = Object.values(scores);
  const specialties = ['1500m', '5000m', '10km', 'Mezza', 'Mara'];
  
  // STEP 1: Calcoli Statistici
  const mean = values.reduce((acc, val) => acc + val, 0) / values.length;
  const variance = values.reduce((acc, val) => acc + Math.pow(val - mean, 2), 0) / values.length;
  const standardDeviation = Math.sqrt(variance);
  
  // STEP 2: Calcolo degli Z-Score
  const zScores = values.map(score => (score - mean) / standardDeviation);
  
  // STEP 3: Identificare il Picco Relativo
  const maxZScore = Math.max(...zScores);
  const maxZScoreIndex = zScores.indexOf(maxZScore);
  const peakSpecialty = specialties[maxZScoreIndex];
  
  // STEP 4: Logica di Classificazione Finale
  const sogliaSpecializzazione = 1.0;
  
  if (maxZScore > sogliaSpecializzazione) {
    // L'atleta è uno SPECIALISTA
    return getRunningProfileFromSpecialty(peakSpecialty);
  } else {
    // L'atleta è ATLETA COMPLETO
    return 'ATLETA COMPLETO';
  }
};

/**
 * Calcola il profilo completo dell'atleta runner
 * CORRETTO: Passa i parametri gender e weight alla funzione di determinazione profilo
 */
export const calculateRunningAthleteProfile = (
  models: PrecisionModels,
  gender: Gender,
  weight: number = gender === 'male' ? 70 : 53
): RunningAthleteProfile => {
  console.log('=== RUNNING ATHLETE PROFILE CALCULATION START ===');
  
  const athleteTimes = calculateRunningAthleteTimes(models);
  const scores = calculateRunningAthleteScores(athleteTimes, gender);
  const overallRating = calculateRunningOverallRating(scores);
  const profileType = determineRunningAthleteProfile(scores, overallRating, gender, weight);
  const rarityTier = determineRunningRarityTier(overallRating);
  
  console.log('=== RUNNING ATHLETE PROFILE CALCULATION END ===');
  console.log('Final profile:', { stats: scores, overallRating, profileType, rarityTier });
  
  return {
    stats: scores,
    overallRating,
    profileType,
    rarityTier
  };
};

// Utility per formattare i nomi dei profili
export const formatRunningProfileType = (profileType: string): string => {
  const translations = {
    'MEZZOFONDISTA VELOCE': 'Mezzofondista Veloce',
    'MEZZOFONDISTA': 'Mezzofondista',
    'FONDISTA': 'Fondista',
    'SPECIALISTA MEZZA': 'Specialista Mezza',
    'MARATONETA': 'Maratoneta',
    'ATLETA COMPLETO': 'Atleta Completo'
  };
  return translations[profileType as keyof typeof translations] || profileType;
};

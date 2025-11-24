import { PowerPrecisionModels } from './powerPrecisionCalculations';

// Benchmark differenziati per sesso (W/kg)
export const MALE_BENCHMARKS = {
  NM: 24.65,    // 5 secondi - Neuromuscular Power
  LACT: 11.33,  // 1 minuto - Lactate Power
  VO2: 7.65,    // 5 minuti - VO2max Power
  THR: 6.59,    // 20 minuti - Threshold Power
  STA: 5.76     // 60 minuti - Stamina Power
} as const;

export const FEMALE_BENCHMARKS = {
  NM: 17.2,     // 5 secondi - Neuromuscular Power
  LACT: 9.4,    // 1 minuto - Lactate Power
  VO2: 6.5,     // 5 minuti - VO2max Power
  THR: 5.5,     // 20 minuti - Threshold Power
  STA: 4.9      // 60 minuti - Stamina Power
} as const;

export type Gender = 'male' | 'female';

export interface AthleteStats {
  NM: number;
  LACT: number;
  VO2: number;
  THR: number;
  STA: number;
}

export type RarityTier = 'ALIEN' | 'HERO' | 'PRO' | 'ELITE' | 'STANDARD';

export interface AthleteProfile {
  stats: AthleteStats;
  overallRating: number;
  profileType: 'SPRINTER' | 'PUNCHEUR' | 'CLIMBER' | 'TIME TRIALIST' | 'ALL-ROUNDER';
  rarityTier: RarityTier;
}

// Metadati per le statistiche (per la visualizzazione)
export const STAT_METADATA = {
  NM: { label: 'Neuromuscolare', duration: '5s' },
  LACT: { label: 'Anaerobico', duration: '1m' },
  VO2: { label: 'VO2max', duration: '5m' },
  THR: { label: 'Soglia', duration: '20m' },
  STA: { label: 'Stamina', duration: '60m' }
} as const;

/**
 * Calcola i W/kg dell'atleta per le 5 durate usando i modelli fisiologici corretti
 */
export const calculateAthleteWattPerKg = (
  models: PowerPrecisionModels,
  bodyWeight: number
): AthleteStats => {
  const { cp, wPrime, apr_params, pl_params } = models;
  const { po3min, apr, k } = apr_params;
  const { S, E } = pl_params;
  
  // 5 secondi (NM) - Usa APR model: P = Po3min + APR * exp(-k*t)
  const nm5s = (po3min + apr * Math.exp(-k * 5)) / bodyWeight;
  
  // 1 minuto (LACT) - Usa APR model: P = Po3min + APR * exp(-k*t)  
  const lact1m = (po3min + apr * Math.exp(-k * 60)) / bodyWeight;
  
  // 5 minuti (VO2) - Usa modello CP/W': P = CP + W'/t
  const vo25m = (cp + wPrime / 300) / bodyWeight;
  
  // 20 minuti (THR) - Usa Power Law: P = S * t^(E-1)
  const thr20m = (S * Math.pow(1200, E - 1)) / bodyWeight;
  
  // 60 minuti (STA) - Usa Power Law: P = S * t^(E-1)
  const sta60m = (S * Math.pow(3600, E - 1)) / bodyWeight;
  
  return {
    NM: nm5s,
    LACT: lact1m,
    VO2: vo25m,
    THR: thr20m,
    STA: sta60m
  };
};

/**
 * Calcola i punteggi per ogni statistica basandosi sui benchmark differenziati per sesso
 */
export const calculateAthleteScores = (
  athleteWattPerKg: AthleteStats, 
  gender: Gender
): AthleteStats => {
  const benchmarks = gender === 'male' ? MALE_BENCHMARKS : FEMALE_BENCHMARKS;
  
  return {
    NM: Math.round((athleteWattPerKg.NM / benchmarks.NM) * 95),
    LACT: Math.round((athleteWattPerKg.LACT / benchmarks.LACT) * 95),
    VO2: Math.round((athleteWattPerKg.VO2 / benchmarks.VO2) * 95),
    THR: Math.round((athleteWattPerKg.THR / benchmarks.THR) * 95),
    STA: Math.round((athleteWattPerKg.STA / benchmarks.STA) * 95)
  };
};

/**
 * Calcola il rating complessivo (media dei 5 punteggi)
 */
export const calculateOverallRating = (scores: AthleteStats): number => {
  const values = Object.values(scores);
  const sum = values.reduce((acc, val) => acc + val, 0);
  return Math.round(sum / values.length);
};

/**
 * Determina il tier di rarità basato sul rating complessivo
 */
export const determineRarityTier = (overallRating: number): RarityTier => {
  if (overallRating >= 100) return 'ALIEN';
  if (overallRating >= 95) return 'HERO';
  if (overallRating >= 90) return 'PRO';
  if (overallRating >= 85) return 'ELITE';
  return 'STANDARD';
};

/**
 * Funzione helper per mappare la specialità di picco al profilo atleta (ciclismo)
 */
const getProfileFromSpecialty = (
  peakSpecialty: string,
  bodyWeight: number
): 'SPRINTER' | 'PUNCHEUR' | 'CLIMBER' | 'TIME TRIALIST' | 'ALL-ROUNDER' => {
  switch (peakSpecialty) {
    case 'NM':
      return 'SPRINTER';
    case 'LACT':
    case 'VO2':
      return 'PUNCHEUR';
    case 'THR':
    case 'STA':
      return (bodyWeight <= 70) ? 'CLIMBER' : 'TIME TRIALIST';
    default:
      return 'ALL-ROUNDER';
  }
};

/**
 * Determina il profilo dell'atleta basandosi sulla logica Z-Score
 */
export const determineAthleteProfile = (
  scores: AthleteStats, 
  overallRating: number,
  bodyWeight: number
): 'SPRINTER' | 'PUNCHEUR' | 'CLIMBER' | 'TIME TRIALIST' | 'ALL-ROUNDER' => {
  const { NM, LACT, VO2, THR, STA } = scores;
  const values = [NM, LACT, VO2, THR, STA];
  const specialties = ['NM', 'LACT', 'VO2', 'THR', 'STA'];
  
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
    return getProfileFromSpecialty(peakSpecialty, bodyWeight);
  } else {
    // L'atleta è ALL-ROUNDER
    return 'ALL-ROUNDER';
  }
};

/**
 * Calcola il profilo completo dell'atleta
 */
export const calculateAthleteProfile = (
  models: PowerPrecisionModels,
  bodyWeight: number,
  gender: Gender
): AthleteProfile => {
  const wattPerKg = calculateAthleteWattPerKg(models, bodyWeight);
  const scores = calculateAthleteScores(wattPerKg, gender);
  const overallRating = calculateOverallRating(scores);
  const profileType = determineAthleteProfile(scores, overallRating, bodyWeight);
  const rarityTier = determineRarityTier(overallRating);
  
  return {
    stats: scores,
    overallRating,
    profileType,
    rarityTier
  };
};

// Utility per formattare i nomi dei profili
export const formatProfileType = (profileType: string): string => {
  const translations = {
    'SPRINTER': 'Sprinter',
    'PUNCHEUR': 'Puncheur',
    'CLIMBER': 'Scalatore', 
    'TIME TRIALIST': 'Cronoman',
    'ALL-ROUNDER': 'All-Rounder'
  };
  return translations[profileType as keyof typeof translations] || profileType;
};
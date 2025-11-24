
// Utility functions for high-precision calculations
export interface PrecisionModels {
  valid: boolean;
  reason?: string;
  cs: number; // Critical Speed in m/s (high precision)
  dPrime: number; // D' in meters (high precision)
  cs_pace: number; // CS pace in seconds per km (high precision)
  pl_params: {
    A: number; // S (velocità base in m/s) - high precision
    k: number; // E-1 (esponente endurance - 1) - high precision
  };
  lt1_range: {
    min: number;
    max: number;
    estimate: number;
  };
  vo2max_pace: number; // 3min sustainable pace (high precision)
}

export const calculateHighPrecisionModels = (
  severeDistance: number,
  severeTime: number,
  thresholdDistance: number,
  thresholdTime: number,
  enableLongTest: boolean,
  longDistance: string,
  longTime: number
): PrecisionModels => {
  console.log('=== HIGH PRECISION CALCULATION START ===');
  
  if (severeTime <= 0 || thresholdTime <= 0) {
    return { 
      valid: false, 
      reason: "I tempi devono essere maggiori di zero.",
      cs: 0, dPrime: 0, cs_pace: 0, 
      pl_params: { A: 0, k: 0 }, 
      lt1_range: { min: 0, max: 0, estimate: 0 }, 
      vo2max_pace: 0
    };
  }

  if (severeTime >= thresholdTime) {
    return { 
      valid: false, 
      reason: "Il tempo dello sforzo breve deve essere minore del tempo medio.",
      cs: 0, dPrime: 0, cs_pace: 0, 
      pl_params: { A: 0, k: 0 }, 
      lt1_range: { min: 0, max: 0, estimate: 0 }, 
      vo2max_pace: 0
    };
  }

  if (severeDistance >= thresholdDistance) {
    return { 
      valid: false, 
      reason: "La distanza dello sforzo breve deve essere minore della distanza media.",
      cs: 0, dPrime: 0, cs_pace: 0, 
      pl_params: { A: 0, k: 0 }, 
      lt1_range: { min: 0, max: 0, estimate: 0 }, 
      vo2max_pace: 0
    };
  }

  // Calculate CS-D' model parameters with maximum precision
  const cs = (thresholdDistance - severeDistance) / (thresholdTime - severeTime);
  const dPrime = severeDistance - cs * severeTime;

  console.log('CS-D\' High Precision:');
  console.log('  CS:', cs, 'm/s');
  console.log('  D\':', dPrime, 'm');

  if (cs <= 0 || dPrime <= 0) {
    return { 
      valid: false, 
      reason: "Calcolo invalido: CS o D' negativi. Verificare i dati inseriti.",
      cs: 0, dPrime: 0, cs_pace: 0, 
      pl_params: { A: 0, k: 0 }, 
      lt1_range: { min: 0, max: 0, estimate: 0 }, 
      vo2max_pace: 0
    };
  }

  const cs_pace = 1000 / cs;

  // Calculate VO2max pace (3min sustainable pace) with high precision
  const vo2max_distance = cs * 180 + dPrime;
  const vo2max_pace = (180 * 1000) / vo2max_distance;

  // Generate high-precision points for Power Law regression
  console.log('=== POWER LAW HIGH PRECISION REGRESSION ===');
  
  let powerLawPoints = [];
  
  // Generate points from CS-D' model for 5, 10, 15 minutes with max precision
  const durations = [300, 600, 900]; // 5, 10, 15 minutes
  durations.forEach(duration => {
    const distance = cs * duration + dPrime;
    const velocity = distance / duration;
    powerLawPoints.push({ time: duration, velocity: velocity });
    console.log(`CS-D\' Point ${duration/60}min: T=${duration}s, V=${velocity}m/s (full precision)`);
  });

  // Add long test if enabled
  if (enableLongTest) {
    const longDistanceM = parseFloat(longDistance) * 1000;
    
    if (longTime > 1200) { // must be over 20 minutes
      const longVelocity = longDistanceM / longTime;
      powerLawPoints.push({ time: longTime, velocity: longVelocity });
      console.log(`Long Point: T=${longTime}s, V=${longVelocity}m/s (full precision)`);
    }
  }

  if (powerLawPoints.length < 2) {
    return { 
      valid: false, 
      reason: "Servono almeno due punti per calcolare il modello Power Law.",
      cs: 0, dPrime: 0, cs_pace: 0, 
      pl_params: { A: 0, k: 0 }, 
      lt1_range: { min: 0, max: 0, estimate: 0 }, 
      vo2max_pace: 0
    };
  }

  // High-precision regression: ln(tempo) vs ln(velocità)
  let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
  powerLawPoints.forEach(point => {
    const x = Math.log(point.time);
    const y = Math.log(point.velocity);
    sumX += x;
    sumY += y;
    sumXY += x * y;
    sumX2 += x * x;
    console.log(`Regression Point: ln(${point.time}) = ${x}, ln(${point.velocity}) = ${y} (full precision)`);
  });

  const n = powerLawPoints.length;
  const denominator = n * sumX2 - sumX * sumX;
  
  if (Math.abs(denominator) < 1e-15) { // Higher precision threshold
    return { 
      valid: false, 
      reason: "Impossibile calcolare Power Law.",
      cs: 0, dPrime: 0, cs_pace: 0, 
      pl_params: { A: 0, k: 0 }, 
      lt1_range: { min: 0, max: 0, estimate: 0 }, 
      vo2max_pace: 0
    };
  }

  const slope = (n * sumXY - sumX * sumY) / denominator;      // E-1 (full precision)
  const intercept = (sumY - slope * sumX) / n;                // ln(S) (full precision)
  
  // High precision parameters
  const S = Math.exp(intercept);  // S = velocità base (m/s) - full precision
  const E = slope + 1;            // E = esponente endurance - full precision
  
  console.log('High Precision Regression Results:');
  console.log('  Slope (E-1):', slope);
  console.log('  Intercept ln(S):', intercept);
  console.log('  S (speed base):', S, 'm/s');
  console.log('  E (endurance):', E);
  
  const A = S;        // A is S (speed base) - full precision
  const k = E - 1;    // k is E-1 (exponent - 1) - full precision
  
  console.log('Final High Precision Parameters:');
  console.log('  A (S):', A, 'm/s');
  console.log('  k (E-1):', k);

  // Calculate LT1 range based on E parameter with high precision
  console.log('=== LT1 CALCULATION (EMPIRICAL) ===');
  console.log('E parameter:', E);
  
  let lt1_upper_speed, lt1_lower_speed;
  
  if (E > 0.90) {
    // For higher endurance athletes (E > 0.90): use 84% of CS
    lt1_upper_speed = cs * 0.84;
    lt1_lower_speed = (cs * 0.84) * 0.96;
    console.log('Using E > 0.90 formula: CS × 0.84 and (CS × 0.84) × 0.96');
  } else {
    // For lower endurance athletes (E ≤ 0.90): use 80% of CS
    lt1_upper_speed = cs * 0.80;
    lt1_lower_speed = (cs * 0.80) * 0.96;
    console.log('Using E ≤ 0.90 formula: CS × 0.80 and (CS × 0.80) × 0.96');
  }
  
  // Convert speeds to paces (s/km) - high precision
  const lt1_min_pace = 1000 / lt1_upper_speed; // Upper speed = faster pace (lower number)
  const lt1_max_pace = 1000 / lt1_lower_speed; // Lower speed = slower pace (higher number)
  const lt1_estimate_pace = (lt1_min_pace + lt1_max_pace) / 2; // Center of range
  
  console.log('LT1 Speed Range (high precision):');
  console.log('  Upper speed (84%/80% CS):', lt1_upper_speed, 'm/s');
  console.log('  Lower speed (96% of upper):', lt1_lower_speed, 'm/s');
  console.log('LT1 Pace Range (high precision):');
  console.log('  Min pace (fastest):', lt1_min_pace, 's/km');
  console.log('  Max pace (slowest):', lt1_max_pace, 's/km');
  console.log('  Estimate (center):', lt1_estimate_pace, 's/km');

  const lt1_range = {
    min: lt1_min_pace,
    max: lt1_max_pace,
    estimate: lt1_estimate_pace
  };

  console.log('=== HIGH PRECISION CALCULATION END ===');

  return {
    valid: true,
    cs,
    dPrime,
    cs_pace,
    pl_params: { A, k },
    lt1_range,
    vo2max_pace
  };
};

// Formatting utilities that round only for display
export const formatForDisplay = {
  pace: (secondsPerKm: number): string => {
    if (!isFinite(secondsPerKm) || secondsPerKm <= 0) return "--:--";
    const minutes = Math.floor(secondsPerKm / 60);
    const seconds = Math.round(secondsPerKm % 60);
    return `${minutes}:${seconds.toString().padStart(2, '0')}`;
  },

  time: (totalSeconds: number): string => {
    if (!isFinite(totalSeconds) || totalSeconds <= 0) return "--:--";
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = Math.round(totalSeconds % 60);
    
    if (hours > 0) {
      return `${hours}h ${minutes}m ${seconds}s`;
    } else if (minutes > 0) {
      return `${minutes}m ${seconds}s`;
    } else {
      return `${seconds}s`;
    }
  },

  distance: (meters: number, unit: 'km' | 'm' = 'km'): string => {
    if (unit === 'km') {
      return `${(meters / 1000).toFixed(2)} km`;
    }
    return `${Math.round(meters)} m`;
  },

  power: (value: number, decimals: number = 0): string => {
    return value.toFixed(decimals);
  }
};

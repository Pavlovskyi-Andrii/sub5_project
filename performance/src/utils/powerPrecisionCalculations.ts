
// High precision calculations for cycling power models
// Similar to precisionCalculations.ts but for power/cycling data

export interface PowerPrecisionModels {
  valid: boolean;
  reason?: string;
  cp: number;
  wPrime: number;
  pMax: number;
  cp_power: number; // CP power for display
  apr_params: {
    po3min: number;
    apr: number;
    k: number;
  };
  pl_params: {
    S: number; // Power Law S parameter (high precision)
    E: number; // Power Law E parameter (high precision)
  };
  lt1_range: {
    min: number;
    max: number;
    estimate: number;
  };
  vo2max_power: number; // 3-minute power estimate
}

export const formatForDisplay = {
  power: (watts: number): string => `${Math.round(watts)} W`,
  time: (seconds: number): string => {
    const sec = Math.round(seconds);
    if (sec < 60) return `${sec}s`;
    if (sec < 3600) {
      const m = Math.floor(sec / 60);
      const s = sec % 60;
      return s > 0 ? `${m}m ${s}s` : `${m}m`;
    }
    const h = Math.floor(sec / 3600);
    const m = Math.round((sec % 3600) / 60);
    return m > 0 ? `${h}h ${m}m` : `${h}h`;
  }
};

export const calculateHighPrecisionPowerModels = (
  pmaxPower: number,
  severePower: number,
  severeTime: number,
  thresholdPower: number,
  thresholdTime: number,
  enableShortTest: boolean,
  shortPower: number,
  shortTime: number,
  enableLongTest: boolean,
  longPower: number,
  longTime: number
): PowerPrecisionModels => {
  
  console.log('=== HIGH PRECISION POWER CALCULATIONS ===');
  console.log('Input data:', {
    pmaxPower, severePower, severeTime, thresholdPower, thresholdTime,
    enableShortTest, shortPower, shortTime, enableLongTest, longPower, longTime
  });

  // Validation checks
  if (severePower <= thresholdPower) {
    return { 
      valid: false, 
      reason: "Incoerenza: P(sforzo breve) deve essere > P(sforzo medio).",
      cp: 0, wPrime: 0, pMax: pmaxPower, cp_power: 0,
      apr_params: { po3min: 0, apr: 0, k: 0 },
      pl_params: { S: 0, E: 0 },
      lt1_range: { min: 0, max: 0, estimate: 0 },
      vo2max_power: 0
    };
  }
  
  if (severeTime >= thresholdTime) {
    return { 
      valid: false, 
      reason: "Incoerenza: T(sforzo breve) deve essere < T(sforzo medio).",
      cp: 0, wPrime: 0, pMax: pmaxPower, cp_power: 0,
      apr_params: { po3min: 0, apr: 0, k: 0 },
      pl_params: { S: 0, E: 0 },
      lt1_range: { min: 0, max: 0, estimate: 0 },
      vo2max_power: 0
    };
  }

  // High precision CP-W' calculation
  const wPrime = (severePower - thresholdPower) * severeTime * thresholdTime / (thresholdTime - severeTime);
  const cp = severePower - (wPrime / severeTime);

  console.log('CP-W\' Model (high precision):', { cp, wPrime });

  if (cp <= 0 || wPrime <= 0) {
    return { 
      valid: false, 
      reason: "Calcolo invalido: CP o W' negativi.",
      cp: 0, wPrime: 0, pMax: pmaxPower, cp_power: 0,
      apr_params: { po3min: 0, apr: 0, k: 0 },
      pl_params: { S: 0, E: 0 },
      lt1_range: { min: 0, max: 0, estimate: 0 },
      vo2max_power: 0
    };
  }

  // APR Model calculations
  const po3min = cp + wPrime / 180; // 3 minutes = 180 seconds
  if (pmaxPower <= po3min) {
    return { 
      valid: false, 
      reason: "Incoerenza: Pmax deve essere > P stimata a 3 min.",
      cp: 0, wPrime: 0, pMax: pmaxPower, cp_power: 0,
      apr_params: { po3min: 0, apr: 0, k: 0 },
      pl_params: { S: 0, E: 0 },
      lt1_range: { min: 0, max: 0, estimate: 0 },
      vo2max_power: 0
    };
  }

  const apr = pmaxPower - po3min;
  let k_apr = 0.026; // default value
  
  // Optimize k if short test is enabled
  if (enableShortTest) {
    if (shortPower > po3min && apr > 0) {
      const ratio = (shortPower - po3min) / apr;
      if (ratio > 0 && ratio < 1) {
        k_apr = Math.max(0.015, Math.min(-(1 / shortTime) * Math.log(ratio), 0.05));
      }
    }
  }

  console.log('APR Model (high precision):', { po3min, apr, k_apr });

  // Power Law Model - HIGH PRECISION calculation
  const pl_points = [
    { t: severeTime, P: severePower }, 
    { t: thresholdTime, P: thresholdPower }
  ];
  
  if (enableLongTest) {
    if (longPower < thresholdPower && longTime > thresholdTime) {
      pl_points.push({ t: longTime, P: longPower });
    }
  }

  console.log('Power Law points:', pl_points);

  // Linear regression on log-transformed data (HIGH PRECISION)
  let sx = 0, sy = 0, sxy = 0, sx2 = 0, sy2 = 0;
  
  pl_points.forEach(pt => {
    const x = Math.log(pt.t);  // ln(time)
    const y = Math.log(pt.P);  // ln(power)
    sx += x;
    sy += y;
    sxy += x * y;
    sx2 += x * x;
    sy2 += y * y;
  });
  
  const n = pl_points.length;
  const denominator = (n * sx2 - sx * sx);
  
  if (Math.abs(denominator) < 1e-12) {
    return { 
      valid: false, 
      reason: "Impossibile calcolare Power Law: punti collineari.",
      cp: 0, wPrime: 0, pMax: pmaxPower, cp_power: 0,
      apr_params: { po3min: 0, apr: 0, k: 0 },
      pl_params: { S: 0, E: 0 },
      lt1_range: { min: 0, max: 0, estimate: 0 },
      vo2max_power: 0
    };
  }
  
  // High precision slope and intercept
  const slope = (n * sxy - sx * sy) / denominator;
  const intercept = (sy - slope * sx) / n;
  
  // Convert back to S and E (HIGH PRECISION)
  const S = Math.exp(intercept);  // S parameter (no rounding)
  const E = slope + 1;            // E parameter (no rounding)
  
  console.log('Power Law Model (high precision):', { 
    slope, intercept, S, E,
    'S (rounded for display)': Math.round(S),
    'E (rounded for display)': Number(E.toFixed(6))
  });

  // LT1 calculation (current method - 72-80% of 30min power)
  const p30min = cp + wPrime / 1800; // 30 min = 1800 seconds
  const lt1_range = {
    min: p30min * 0.72,
    max: p30min * 0.80,
    estimate: p30min * 0.75
  };

  console.log('LT1 Range:', lt1_range);

  const result: PowerPrecisionModels = {
    valid: true,
    cp,
    wPrime,
    pMax: pmaxPower,
    cp_power: cp, // for display
    apr_params: { po3min, apr, k: k_apr },
    pl_params: { S, E }, // High precision values
    lt1_range,
    vo2max_power: po3min // 3-minute power
  };

  console.log('=== FINAL RESULT (HIGH PRECISION) ===', result);
  
  return result;
};

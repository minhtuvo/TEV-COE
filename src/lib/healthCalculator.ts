export interface DNOAssetParams {
  age: number;
  normalExpectedLife: number;
  dutyFactor: number;
  locationFactor: number;
  healthScoreFactor: number;
  reliabilityFactor: number;
  healthScoreCap: number;
  healthScoreCollar: number;
  reliabilityCollar: number;
}

export interface SwitchgearParams extends DNOAssetParams {
  observedFactor: number;
  measuredFactor: number;
}

export interface TransformerParams {
  mainTransformer: DNOAssetParams;
  tapchanger: DNOAssetParams;
}

export interface MotorPhaseData {
  R: number | null;
  Y: number | null;
  B: number | null;
}

export interface MotorDiagnosticTests {
  tanDelta?: MotorPhaseData;
  tipUp?: MotorPhaseData;
  pd?: MotorPhaseData;
  ir?: MotorPhaseData;
  pi?: MotorPhaseData;
  dd?: MotorPhaseData;
  elcid?: MotorPhaseData;
  ratedKV: number;
}

// --- PART 1: SWITCHGEAR & TRANSFORMERS (DNO Methodology) ---

export function calculateDNOHealthScore(params: DNOAssetParams): number {
  const {
    age,
    normalExpectedLife,
    dutyFactor,
    locationFactor,
    healthScoreFactor,
    reliabilityFactor,
    healthScoreCap,
    healthScoreCollar,
    reliabilityCollar
  } = params;

  const expectedLife = normalExpectedLife / (dutyFactor * locationFactor);
  const beta1 = Math.log(5.5 / 0.5) / expectedLife;
  
  let initialHealthScore = 0.5 * Math.exp(beta1 * age);
  initialHealthScore = Math.min(initialHealthScore, 5.5);

  let currentHealthScore = initialHealthScore * healthScoreFactor * reliabilityFactor;
  
  // Apply Cap and Collar
  currentHealthScore = Math.min(currentHealthScore, healthScoreCap, 10);
  const collar = Math.max(healthScoreCollar, reliabilityCollar);
  currentHealthScore = Math.max(currentHealthScore, collar);

  return currentHealthScore;
}

export function getHealthIndexBanding(score: number): string {
  if (score < 4) return 'HI1';
  if (score < 5.5) return 'HI2';
  if (score < 6.5) return 'HI3';
  if (score < 8) return 'HI4';
  return 'HI5';
}

export function calculateSwitchgearHealth(params: SwitchgearParams): { score: number, banding: string } {
  const a = Math.max(params.observedFactor, params.measuredFactor);
  const b = Math.min(params.observedFactor, params.measuredFactor);
  
  let healthScoreFactor = 1;
  if (a > 1 && b > 1) {
    healthScoreFactor = a + ((b - 1) / 1.5);
  } else if (a > 1 && b <= 1) {
    healthScoreFactor = a;
  } else if (a <= 1 && b <= 1) {
    healthScoreFactor = b + ((a - 1) / 1.5);
  }

  const dnoParams: DNOAssetParams = {
    ...params,
    healthScoreFactor
  };

  const score = calculateDNOHealthScore(dnoParams);
  return { score, banding: getHealthIndexBanding(score) };
}

export function calculateTransformerHealth(params: TransformerParams): { score: number, banding: string } {
  const mainScore = calculateDNOHealthScore(params.mainTransformer);
  const tapchangerScore = calculateDNOHealthScore(params.tapchanger);
  
  const finalScore = Math.max(mainScore, tapchangerScore);
  return { score: finalScore, banding: getHealthIndexBanding(finalScore) };
}

// --- PART 2: HIGH VOLTAGE (HV) MOTORS (Weighted Scoring Method) ---

function getWorstValue(data: MotorPhaseData | undefined, type: 'max' | 'min'): number | null {
  if (!data) return null;
  const values = [data.R, data.Y, data.B].filter((v): v is number => v !== null && v !== undefined);
  if (values.length === 0) return null;
  return type === 'max' ? Math.max(...values) : Math.min(...values);
}

export function calculateMotorHealth(tests: MotorDiagnosticTests): { absoluteHI: number, hiPercentage: number, banding: string } {
  let totalScoreWeight = 0;
  let totalWeight = 0;

  // Test 1: Tan-delta (MAX)
  const tanDelta = getWorstValue(tests.tanDelta, 'max');
  if (tanDelta !== null) {
    const w = 2;
    let s = 1;
    if (tanDelta < 0.02) s = 10;
    else if (tanDelta <= 0.04) s = 7;
    else if (tanDelta <= 0.07) s = 5;
    else s = 1;
    
    totalScoreWeight += s * w;
    totalWeight += w;
  }

  // Test 2: Tip-up (MAX)
  const tipUp = getWorstValue(tests.tipUp, 'max');
  if (tipUp !== null) {
    const w = 2;
    let s = 1;
    if (tipUp < 0.002) s = 10;
    else if (tipUp <= 0.004) s = 7;
    else if (tipUp <= 0.006) s = 5;
    else s = 1;
    
    totalScoreWeight += s * w;
    totalWeight += w;
  }

  // Test 3: PD in pC (MAX)
  const pd = getWorstValue(tests.pd, 'max');
  if (pd !== null) {
    const w = 3;
    let s = 1;
    if (pd < 5000) s = 10;
    else if (pd <= 10000) s = 7;
    else if (pd <= 15000) s = 5;
    else s = 1;
    
    totalScoreWeight += s * w;
    totalWeight += w;
  }

  // Test 4: IR in G-ohm (MIN)
  const ir = getWorstValue(tests.ir, 'min');
  if (ir !== null) {
    const w = 1;
    let s = 1;
    const threshold = 1 + (tests.ratedKV / 1000);
    if (ir > 50) s = 10;
    else if (ir >= 10) s = 7;
    else if (ir >= 1) s = 5;
    else s = 1;
    
    if (ir < threshold) s = 1;
    
    totalScoreWeight += s * w;
    totalWeight += w;
  }

  // Test 5: PI (MIN)
  const pi = getWorstValue(tests.pi, 'min');
  if (pi !== null) {
    const w = 1;
    let s = 1;
    if (pi > 2) s = 10;
    else if (pi >= 1.5) s = 7;
    else if (pi >= 1) s = 5;
    else s = 1;
    
    totalScoreWeight += s * w;
    totalWeight += w;
  }

  // Test 6: DD (MAX)
  const dd = getWorstValue(tests.dd, 'max');
  if (dd !== null) {
    const w = 1;
    let s = 1;
    if (dd < 1) s = 10;
    else if (dd <= 4) s = 7;
    else if (dd <= 8) s = 5;
    else s = 1;
    
    totalScoreWeight += s * w;
    totalWeight += w;
  }

  // Test 7: ELCID in mA (MAX)
  const elcid = getWorstValue(tests.elcid, 'max');
  if (elcid !== null) {
    const w = 1;
    let s = 1;
    if (elcid < 100) s = 10;
    else if (elcid <= 200) s = 7;
    else if (elcid <= 300) s = 5;
    else s = 1;
    
    totalScoreWeight += s * w;
    totalWeight += w;
  }

  if (totalWeight === 0) return { absoluteHI: 0, hiPercentage: 0, banding: 'N/A' };

  const absoluteHI = totalScoreWeight / totalWeight;
  const hiPercentage = (totalScoreWeight / (10 * totalWeight)) * 100;
  
  // Map percentage to banding (similar to DNO 1-10 scale, 100% = 10)
  const equivalentScore = (hiPercentage / 100) * 10;
  // Note: For motors, higher is better (10 is healthy). For DNO, lower is better (0.5 is healthy).
  // We'll invert the banding logic for motors to match DNO's HI1 (good) to HI5 (bad)
  let banding = 'HI1';
  if (hiPercentage < 40) banding = 'HI5';
  else if (hiPercentage < 55) banding = 'HI4';
  else if (hiPercentage < 65) banding = 'HI3';
  else if (hiPercentage < 80) banding = 'HI2';

  return { absoluteHI, hiPercentage, banding };
}

// --- DUMMY DATA TEST ---
export function runDummyTests() {
  const sampleSwitchgear: SwitchgearParams = {
    age: 20,
    normalExpectedLife: 40,
    dutyFactor: 1.0,
    locationFactor: 1.0,
    reliabilityFactor: 1.0,
    healthScoreCap: 10,
    healthScoreCollar: 0.5,
    reliabilityCollar: 0.5,
    observedFactor: 1.2,
    measuredFactor: 1.5,
    healthScoreFactor: 1 // Will be recalculated
  };

  const sampleTransformer: TransformerParams = {
    mainTransformer: {
      age: 30,
      normalExpectedLife: 50,
      dutyFactor: 1.2,
      locationFactor: 1.0,
      healthScoreFactor: 1.5,
      reliabilityFactor: 1.0,
      healthScoreCap: 10,
      healthScoreCollar: 0.5,
      reliabilityCollar: 0.5
    },
    tapchanger: {
      age: 30,
      normalExpectedLife: 40,
      dutyFactor: 1.0,
      locationFactor: 1.0,
      healthScoreFactor: 1.2,
      reliabilityFactor: 1.0,
      healthScoreCap: 10,
      healthScoreCollar: 0.5,
      reliabilityCollar: 0.5
    }
  };

  const sampleMotor: MotorDiagnosticTests = {
    ratedKV: 6600,
    tanDelta: { R: 0.01, Y: 0.03, B: 0.015 }, // Max: 0.03 -> Score 7
    tipUp: { R: 0.001, Y: 0.001, B: 0.003 }, // Max: 0.003 -> Score 7
    pd: { R: 4000, Y: 4500, B: 12000 }, // Max: 12000 -> Score 7 (approx 10000)
    ir: { R: 60, Y: 55, B: 45 }, // Min: 45 -> Score 7
    pi: { R: 2.5, Y: 2.1, B: 1.9 }, // Min: 1.9 -> Score 5 (since < 2 but not 1, fallback to 5 or 1 based on logic)
    dd: { R: 0.5, Y: 0.8, B: 1.2 }, // Max: 1.2 -> Score 7
    elcid: { R: 50, Y: 60, B: 40 } // Max: 60 -> Score 10
  };

  console.log("--- Switchgear Test ---");
  console.log(calculateSwitchgearHealth(sampleSwitchgear));

  console.log("--- Transformer Test ---");
  console.log(calculateTransformerHealth(sampleTransformer));

  console.log("--- Motor Test ---");
  console.log(calculateMotorHealth(sampleMotor));
}

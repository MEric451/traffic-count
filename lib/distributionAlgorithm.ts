// Vehicle class names matching Excel file
export const VEHICLE_CLASSES = [
  'Motorcycle/Tuktuk',
  'Cars(<=5seats)',
  'Large Cars,4WD,SUVs',
  'Pickups/Vans',
  'Minibus/Matatus(<=14seat',
  'Small bus(14-33seat)',
  'Large bus(>33seat)',
  'Light Trucks(2axles)',
  'Medium Trucks(2axle,Double rear wheel)',
  'Heavy Trucks(3, 4 axles)',
  'Artics/draw-bar trucks',
  'Others (e.g tractors)'
];

// Traffic pattern definitions (percentage ranges for each vehicle class)
export type TrafficPattern = 'Heavy' | 'Balanced' | 'Truck-Heavy' | 'learned';

interface PatternRanges {
  min: number;
  max: number;
}

const PATTERNS: Record<TrafficPattern, PatternRanges[]> = {
  'Heavy': [
    { min: 0.15, max: 0.20 }, // Motorcycle/Tuktuk
    { min: 0.25, max: 0.30 }, // Cars
    { min: 0.08, max: 0.12 }, // Large Cars
    { min: 0.08, max: 0.12 }, // Pickups/Vans
    { min: 0.15, max: 0.20 }, // Matatus
    { min: 0.03, max: 0.05 }, // Small Bus
    { min: 0.02, max: 0.04 }, // Large Bus
    { min: 0.03, max: 0.05 }, // Light Trucks
    { min: 0.02, max: 0.04 }, // Med Trucks
    { min: 0.01, max: 0.03 }, // Heavy Trucks
    { min: 0.00, max: 0.01 }, // Artics
    { min: 0.01, max: 0.02 }  // Others
  ],
  'Balanced': [
    { min: 0.10, max: 0.15 }, // Motorcycle/Tuktuk
    { min: 0.20, max: 0.25 }, // Cars
    { min: 0.10, max: 0.12 }, // Large Cars
    { min: 0.08, max: 0.10 }, // Pickups/Vans
    { min: 0.12, max: 0.15 }, // Matatus
    { min: 0.05, max: 0.07 }, // Small Bus
    { min: 0.05, max: 0.07 }, // Large Bus
    { min: 0.05, max: 0.08 }, // Light Trucks
    { min: 0.05, max: 0.08 }, // Med Trucks
    { min: 0.03, max: 0.06 }, // Heavy Trucks
    { min: 0.01, max: 0.03 }, // Artics
    { min: 0.02, max: 0.04 }  // Others
  ],
  'Truck-Heavy': [
    { min: 0.05, max: 0.08 }, // Motorcycle/Tuktuk
    { min: 0.15, max: 0.20 }, // Cars
    { min: 0.05, max: 0.08 }, // Large Cars
    { min: 0.05, max: 0.08 }, // Pickups/Vans
    { min: 0.08, max: 0.12 }, // Matatus
    { min: 0.03, max: 0.05 }, // Small Bus
    { min: 0.03, max: 0.05 }, // Large Bus
    { min: 0.10, max: 0.15 }, // Light Trucks
    { min: 0.10, max: 0.15 }, // Med Trucks
    { min: 0.10, max: 0.15 }, // Heavy Trucks
    { min: 0.05, max: 0.10 }, // Artics
    { min: 0.02, max: 0.05 }  // Others
  ]
};

/**
 * Scale learned pattern to exact total
 * @param learnedPattern - Original pattern from Excel
 * @param targetTotal - New total to scale to
 * @param enabledClasses - Which classes are enabled
 * @returns Scaled distribution that sums exactly to targetTotal
 */
function scalePatternToTotal(
  learnedPattern: number[],
  targetTotal: number,
  enabledClasses: boolean[]
): number[] {
  const originalTotal = learnedPattern.reduce((a, b) => a + b, 0);
  
  if (originalTotal === 0) {
    return new Array(12).fill(0);
  }

  // Calculate exact scaling factor
  const scaleFactor = targetTotal / originalTotal;
  
  // Scale each value and handle disabled classes
  const scaled = learnedPattern.map((value, index) => {
    if (!enabledClasses[index]) return 0;
    return Math.round(value * scaleFactor);
  });

  // Ensure exact total by adjusting largest enabled value
  const currentTotal = scaled.reduce((a, b) => a + b, 0);
  const difference = targetTotal - currentTotal;
  
  if (difference !== 0) {
    // Find the largest enabled value to adjust
    let maxIndex = -1;
    let maxValue = -1;
    
    for (let i = 0; i < scaled.length; i++) {
      if (enabledClasses[i] && scaled[i] > maxValue) {
        maxValue = scaled[i];
        maxIndex = i;
      }
    }
    
    if (maxIndex !== -1) {
      scaled[maxIndex] = Math.max(0, scaled[maxIndex] + difference);
    }
  }

  return scaled;
}

/**
 * Generate realistic traffic distribution using exact learned patterns from Excel
 * @param total - Total number of vehicles
 * @param pattern - Traffic pattern type
 * @param enabledClasses - Array of booleans indicating which classes are enabled
 * @param learnedPattern - Exact pattern from Excel analysis
 * @returns Array of 12 numbers that sum exactly to total
 */
export function generateDistribution(
  total: number,
  pattern: TrafficPattern | 'learned',
  enabledClasses: boolean[],
  learnedPattern?: number[]
): number[] {
  if (total <= 0) {
    return new Array(12).fill(0);
  }

  // Use exact learned pattern if available
  if (pattern === 'learned' && learnedPattern && learnedPattern.length === 12) {
    return scalePatternToTotal(learnedPattern, total, enabledClasses);
  }

  // Fallback to predefined patterns
  const patternRanges = PATTERNS[pattern as TrafficPattern] || PATTERNS['Balanced'];
  const distribution: number[] = new Array(12).fill(0);

  // Step 1: Generate initial values based on pattern
  let sum = 0;
  for (let i = 0; i < 12; i++) {
    if (!enabledClasses[i]) {
      distribution[i] = 0;
      continue;
    }

    const range = patternRanges[i];
    // Random percentage within range
    const percentage = range.min + Math.random() * (range.max - range.min);
    let value = Math.floor(total * percentage);

    // Avoid too many round numbers (ending in 0 or 5)
    if (value > 10 && (value % 10 === 0 || value % 10 === 5)) {
      const adjustment = Math.random() > 0.5 ? 
        Math.floor(Math.random() * 3) + 1 : 
        -(Math.floor(Math.random() * 3) + 1);
      value += adjustment;
    }

    distribution[i] = Math.max(0, value);
    sum += distribution[i];
  }

  // Step 2: Adjust to match exact total
  const difference = total - sum;
  
  if (difference !== 0) {
    // Find enabled classes with non-zero values
    const adjustableIndices = distribution
      .map((val, idx) => ({ val, idx }))
      .filter(item => enabledClasses[item.idx] && item.val > 0)
      .sort((a, b) => b.val - a.val); // Sort by value descending

    if (adjustableIndices.length > 0) {
      // Distribute difference across adjustable classes
      const perClass = Math.floor(difference / adjustableIndices.length);
      let remainder = difference % adjustableIndices.length;

      for (let i = 0; i < adjustableIndices.length; i++) {
        const idx = adjustableIndices[i].idx;
        let adjustment = perClass;
        
        if (remainder !== 0) {
          adjustment += remainder > 0 ? 1 : -1;
          remainder += remainder > 0 ? -1 : 1;
        }

        distribution[idx] = Math.max(0, distribution[idx] + adjustment);
      }
    }
  }

  // Step 3: Final verification and correction
  const finalSum = distribution.reduce((a, b) => a + b, 0);
  if (finalSum !== total) {
    // Force exact match by adjusting the largest enabled value
    const largestIdx = distribution
      .map((val, idx) => ({ val, idx }))
      .filter(item => enabledClasses[item.idx])
      .sort((a, b) => b.val - a.val)[0]?.idx;

    if (largestIdx !== undefined) {
      distribution[largestIdx] += (total - finalSum);
    }
  }

  return distribution;
}

/**
 * Validate distribution array
 * @param distribution - Array of vehicle counts
 * @param expectedCount - Expected number of values (12, 16, or 24)
 * @param expectedTotal - Expected sum (optional)
 * @returns Validation result
 */
export function validateDistribution(
  distribution: number[],
  expectedCount: number,
  expectedTotal?: number
): { valid: boolean; error?: string; sum: number } {
  // Check if array exists
  if (!Array.isArray(distribution)) {
    return { valid: false, error: 'Invalid format', sum: 0 };
  }

  // Check count
  if (distribution.length !== expectedCount) {
    return {
      valid: false,
      error: `Expected ${expectedCount} values, got ${distribution.length}`,
      sum: 0
    };
  }

  // Check if all are numbers
  if (!distribution.every(val => typeof val === 'number' && !isNaN(val))) {
    return { valid: false, error: 'All values must be numbers', sum: 0 };
  }

  // Check if all are non-negative
  if (!distribution.every(val => val >= 0)) {
    return { valid: false, error: 'All values must be non-negative', sum: 0 };
  }

  // Calculate sum
  const sum = distribution.reduce((a, b) => a + b, 0);

  // Check sum if expected total provided
  if (expectedTotal !== undefined && sum !== expectedTotal) {
    return {
      valid: false,
      error: `Sum is ${sum} but target is ${expectedTotal} (difference: ${sum - expectedTotal > 0 ? '+' : ''}${sum - expectedTotal})`,
      sum
    };
  }

  return { valid: true, sum };
}

/**
 * Parse distribution from various input formats
 * @param input - String input (can be "[1,2,3]" or "1,2,3" or "1 2 3")
 * @returns Array of numbers or null if invalid
 */
export function parseDistributionInput(input: string): number[] | null {
  try {
    // Remove brackets and extra spaces
    const cleaned = input.trim().replace(/[\[\]]/g, '');
    
    // Split by comma or space
    const parts = cleaned.split(/[,\s]+/).filter(p => p.length > 0);
    
    // Convert to numbers
    const numbers = parts.map(p => {
      const num = parseFloat(p);
      if (isNaN(num)) throw new Error('Invalid number');
      return Math.floor(num); // Ensure integers
    });

    return numbers;
  } catch {
    return null;
  }
}

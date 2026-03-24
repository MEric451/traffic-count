import { type ClassValue, clsx } from "clsx";
import { twMerge } from "tailwind-merge";

// Utility function for merging Tailwind classes
export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

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
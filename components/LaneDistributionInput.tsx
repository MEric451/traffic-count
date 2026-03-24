'use client';

import { useState } from 'react';
import { Input } from '@/components/ui/input';
import { parseDistributionInput, validateDistribution, VEHICLE_CLASSES } from '@/lib/utils';

interface LaneDistributionInputProps {
  laneName: string;
  onDistributionChange: (distribution: number[]) => void;
  expectedCount: number;
}

export default function LaneDistributionInput({ 
  laneName, 
  onDistributionChange,
  expectedCount
}: LaneDistributionInputProps) {
  const [distribution, setDistribution] = useState<number[]>([]);
  const [manualInput, setManualInput] = useState<string>('');
  const [validation, setValidation] = useState<{ valid: boolean; error?: string; sum: number }>({ valid: true, sum: 0 });

  const handleManualChange = (value: string) => {
    setManualInput(value);
    const parsed = parseDistributionInput(value);
    
    if (parsed) {
      setDistribution(parsed);
      onDistributionChange(parsed);
      
      const calculatedSum = parsed.reduce((sum, val) => sum + val, 0);
      const validation = validateDistribution(parsed, expectedCount, calculatedSum);
      setValidation(validation);
    } else {
      setValidation({ valid: false, error: 'Invalid format', sum: 0 });
    }
  };

  const handleValueChange = (index: number, value: string) => {
    const newDist = [...distribution];
    const numValue = value === '' ? 0 : parseInt(value) || 0;
    newDist[index] = numValue;
    setDistribution(newDist);
    onDistributionChange(newDist);
    
    const updatedArray = `[${newDist.join(', ')}]`;
    setManualInput(updatedArray);
    
    const calculatedSum = newDist.reduce((sum, val) => sum + val, 0);
    const validation = validateDistribution(newDist, expectedCount, calculatedSum);
    setValidation(validation);
  };

  return (
    <div className="p-4 border rounded-lg bg-white">
      <h3 className="text-lg font-semibold text-gray-800 mb-4">{laneName}</h3>
        {/* Total Vehicles Input */}
        {/* <div>
          <label className="text-sm font-medium mb-2 block text-gray-700">Target Total Vehicles</label>
          <Input
            type="number"
            placeholder="e.g., 11428"
            value={total}
            onChange={(e) => !isAutoCalculated && setTotal(e.target.value)}
            disabled={isAutoCalculated}
            className={`border-2 border-gray-300 bg-white text-gray-700 placeholder:text-gray-500 ${
              isAutoCalculated ? 'bg-gray-100 cursor-not-allowed' : ''
            }`}
          />
        </div> */}

        {/* Manual Entry */}
        <div className="space-y-2">
          <label className="text-sm font-medium block text-gray-700 mt-4">Paste Distribution Array </label>
          <textarea
            className="w-full min-h-[100px] p-3 border-2 border-gray-300 bg-white rounded-md text-sm font-mono text-black placeholder:text-gray-500"
            placeholder="[1421, 3098, 1368, 895, 1720, 809, 176, 150, 566, 438, 0, 0]"
            value={manualInput}
            onChange={(e) => handleManualChange(e.target.value)}
          />
        </div>

        {/* Validation. */}
        {manualInput && (
          <div className={`text-sm p-2 rounded ${validation.valid ? 'bg-green-50 text-green-700' : 'bg-red-50 text-red-700'} mb-2`}>
            {validation.valid ? (
              <span>✓ Valid: {expectedCount} values, Sum: {validation.sum.toLocaleString()}</span>
            ) : (
              <span>✗ {validation.error}</span>
            )}
          </div>
        )}

        {/* Generated Distribution Display */}
        {distribution.length > 0 && validation.valid && (
          <div className="space-y-2 border rounded-lg p-4 bg-gray-50">
            <div className="flex justify-between items-center mb-2">
              <label className="text-sm font-medium text-black">Distribution (Editable)</label>
              <span className="text-sm font-medium text-green-600">
                Sum: {validation.sum.toLocaleString()} ✓
              </span>
            </div>

            <div className="grid grid-cols-1 gap-2 max-h-64 overflow-y-auto">
              {VEHICLE_CLASSES.map((className, index) => (
                <div key={index} className="flex items-center gap-2">
                  <span className="text-xs flex-1 truncate text-black">{className}:</span>
                  <Input
                    type="number"
                    value={distribution[index] === 0 ? '' : distribution[index]}
                    onChange={(e) => handleValueChange(index, e.target.value)}
                    className="w-24 h-8 text-sm border-2 border-black bg-white text-black"
                    placeholder="0"
                  />
                  <span className="text-xs text-gray-500 w-16 text-right">
                    {distribution[index] > 0 && validation.sum > 0
                      ? `${((distribution[index] / validation.sum) * 100).toFixed(1)}%`
                      : '0%'}
                  </span>
                </div>
              ))}
            </div>
          </div>
        )}
    </div>
  );
}

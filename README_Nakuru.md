# Nakuru Traffic Count Modifier (14-Hour Format)

## Overview
This script is specifically designed for the **Nakuru Area Day 1 Thursday Counts, Safari Center Site Counts.xlsx** file, which uses a 14-hour format from 7-8AM to 8-9PM, plus a DAY totals sheet.

## Key Differences from 16-Hour Version

### 1. **Time Range**
- **16-hour version**: 6-7AM to 11-12AM (16 hours) + DAY sheet
- **14-hour version**: 7-8AM to 8-9PM (14 hours) + DAY sheet

### 2. **Direction Names**
- **16-hour version**: 'Kamandura Bound' and 'Ruaka Bound'
- **14-hour version**: 'Nakuru Bound' and 'Nairobi Bound'

### 3. **Traffic Intensity Weights**
The 14-hour version uses optimized weights for the 7AM-9PM period:
```python
weights = {
    '7-8AM': 0.08,   # Morning rush start
    '8-9AM': 0.09,   # Peak morning rush
    '9-10AM': 0.075, # Post morning rush
    '10-11AM': 0.07, # Mid-morning
    '11-12PM': 0.07, # Late morning
    '12-1PM': 0.075, # Lunch hour
    '1-2PM': 0.075,  # Post lunch
    '2-3PM': 0.075,  # Afternoon
    '3-4PM': 0.075,  # Late afternoon
    '4-5PM': 0.085,  # Evening rush start
    '5-6PM': 0.09,   # Peak evening rush
    '6-7PM': 0.09,   # Evening rush continues
    '7-8PM': 0.07,   # Evening wind down
    '8-9PM': 0.05    # Night start
}
```

## How to Use

### 1. **Setup**
- Ensure the Nakuru Excel file is in the same directory as the script
- Install required dependency: `pip install openpyxl`

### 2. **Configure Target Totals**
Edit the `totals` dictionary in the script:
```python
totals = {
    'Nakuru Bound':  [450, 890, 320, 280, 150, 45, 12, 95, 180, 75, 15, 8],
    'Nairobi Bound': [380, 720, 290, 240, 130, 25, 8, 40, 220, 85, 20, 10]
}
```

Each array contains 12 values representing vehicle classes 1-12.

### 3. **Set File Paths**
```python
input_file = "Nakuru Area Day 1 Thursday Counts, Safari Center Site Counts.xlsx"
output_file = "Nakuru Area Modified.xlsx"
```

### 4. **Run the Script**
```bash
python force_exact_totals_nakuru_final.py
```

## Output
The script will:
1. Load the original Excel file
2. Distribute your target totals across 14 hourly sheets
3. Apply realistic traffic variation patterns
4. Save the modified file
5. Verify that totals match exactly

## File Structure Expected
```
Nakuru Excel File:
├── 7-8AM (hourly data)
├── 8-9AM (hourly data)
├── 9-10AM (hourly data)
├── 10-11AM (hourly data)
├── 11-12PM (hourly data)
├── 12-1PM (hourly data)
├── 1-2PM (hourly data)
├── 2-3PM (hourly data)
├── 3-4PM (hourly data)
├── 4-5PM (hourly data)
├── 5-6PM (hourly data)
├── 6-7PM (hourly data)
├── 7-8PM (hourly data)
├── 8-9PM (hourly data)
└── DAY (totals - not modified)
```

## Features
- **Exact Total Matching**: Guarantees your target totals are met precisely
- **Realistic Distribution**: Uses traffic intensity weights for natural patterns
- **Format Preservation**: Maintains original Excel formatting and formulas
- **Comprehensive Verification**: Shows detailed before/after comparison
- **Error Handling**: Validates input and provides clear error messages

## Example Output
```
NAKURU BOUND:
  Vehicle Class  1:  450 ->  450 [MATCH]
  Vehicle Class  2:  890 ->  890 [MATCH]
  ...
  DIRECTION TOTAL:     2520 -> 2520 [PERFECT MATCH!]

NAIROBI BOUND:
  Vehicle Class  1:  380 ->  380 [MATCH]
  Vehicle Class  2:  720 ->  720 [MATCH]
  ...
  DIRECTION TOTAL:     2168 -> 2168 [PERFECT MATCH!]
```

## Notes
- The script automatically detects direction rows by scanning column A
- Traffic variation (±20% hourly, ±30% per row) creates realistic patterns
- The DAY sheet is preserved but not modified
- All original formatting and formulas are maintained
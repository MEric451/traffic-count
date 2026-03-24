import { NextRequest, NextResponse } from 'next/server';
import { writeFile, unlink } from 'fs/promises';
import { exec } from 'child_process';
import { promisify } from 'util';
import path from 'path';

const execAsync = promisify(exec);

export async function POST(request: NextRequest) {
  let tempFilePath = '';

  try {
    const formData = await request.formData();
    const file = formData.get('file') as File;

    if (!file) {
      return NextResponse.json(
        { success: false, error: 'No file uploaded' },
        { status: 400 }
      );
    }

    // Save file temporarily
    const bytes = await file.arrayBuffer();
    const buffer = Buffer.from(bytes);
    tempFilePath = path.join(process.cwd(), `temp_analyze_${Date.now()}.xlsx`);
    
    await writeFile(tempFilePath, buffer);

    // Call Python script
    const pythonScript = path.join(process.cwd(), 'api', 'python', 'analyze_pattern.py');
    const command = `python "${pythonScript}" "${tempFilePath}"`;

    const { stdout, stderr } = await execAsync(command);

    // Clean up temp file
    try {
      await unlink(tempFilePath);
    } catch {
      console.error('Error cleaning up temp file');
    }

    if (stderr && !stderr.includes('DeprecationWarning')) {
      console.error('Python stderr:', stderr);
    }

    // Parse Python output
    const result = JSON.parse(stdout);

    if (!result.success) {
      return NextResponse.json(
        { success: false, error: result.error },
        { status: 400 }
      );
    }

    return NextResponse.json({
      success: true,
      patterns: result.patterns,
      grandTotal: result.grand_total,
      laneCount: result.lane_count
    });

  } catch (error) {
    console.error('Error analyzing pattern:', error);

    // Clean up temp file on error
    try {
      if (tempFilePath) await unlink(tempFilePath);
    } catch {
      console.error('Error cleaning up after error');
    }

    return NextResponse.json(
      { 
        success: false, 
        error: error instanceof Error ? error.message : 'Unknown error occurred' 
      },
      { status: 500 }
    );
  }
}
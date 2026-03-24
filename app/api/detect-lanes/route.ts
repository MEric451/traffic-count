import { NextRequest, NextResponse } from 'next/server';
import { writeFile } from 'fs/promises';
import { exec } from 'child_process';
import { promisify } from 'util';
import path from 'path';

const execAsync = promisify(exec);

export async function POST(request: NextRequest) {
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
    const tempFilePath = path.join(process.cwd(), `temp_detect_${Date.now()}.xlsx`);
    
    await writeFile(tempFilePath, buffer);

    // Call Python script
    const pythonScript = path.join(process.cwd(), 'api', 'python', 'detect_lanes.py');
    const command = `python "${pythonScript}" "${tempFilePath}"`;

    const { stdout, stderr } = await execAsync(command);

    // Clean up temp file
    try {
      const { unlink } = await import('fs/promises');
      await unlink(tempFilePath);
    } catch (cleanupError) {
      console.error('Error cleaning up temp file:', cleanupError);
    }

    if (stderr) {
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
      lanes: result.lanes,
      sheetUsed: result.sheet_used
    });

  } catch (error) {
    console.error('Error detecting lanes:', error);
    return NextResponse.json(
      { 
        success: false, 
        error: error instanceof Error ? error.message : 'Unknown error occurred' 
      },
      { status: 500 }
    );
  }
}

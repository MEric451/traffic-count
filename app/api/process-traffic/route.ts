import { NextRequest, NextResponse } from 'next/server';
import { writeFile, readFile, unlink } from 'fs/promises';
import { exec } from 'child_process';
import { promisify } from 'util';
import path from 'path';

const execAsync = promisify(exec);

export async function POST(request: NextRequest) {
  let tempInputPath = '';
  let tempOutputPath = '';

  try {
    const formData = await request.formData();
    
    // Extract form data
    const file = formData.get('file') as File;
    const scriptType = formData.get('scriptType') as string; // '12hour', '16hour', '24hour'
    const outputFilename = formData.get('outputFilename') as string;
    const totalsJson = formData.get('totals') as string;

    // Validate inputs
    if (!file) {
      return NextResponse.json(
        { success: false, error: 'No file uploaded' },
        { status: 400 }
      );
    }

    if (!scriptType || !['12hour', '16hour', '24hour'].includes(scriptType)) {
      return NextResponse.json(
        { success: false, error: 'Invalid script type' },
        { status: 400 }
      );
    }

    if (!totalsJson) {
      return NextResponse.json(
        { success: false, error: 'No totals provided' },
        { status: 400 }
      );
    }

    // Parse totals
    let totals;
    try {
      totals = JSON.parse(totalsJson);
    } catch {
      return NextResponse.json(
        { success: false, error: 'Invalid totals format' },
        { status: 400 }
      );
    }

    // Save input file temporarily
    const bytes = await file.arrayBuffer();
    const buffer = Buffer.from(bytes);
    const timestamp = Date.now();
    tempInputPath = path.join(process.cwd(), `temp_input_${timestamp}.xlsx`);
    tempOutputPath = path.join(process.cwd(), `temp_output_${timestamp}.xlsx`);
    
    await writeFile(tempInputPath, buffer);

    // Determine which Python script to use
    const scriptMap = {
      '12hour': 'force_exact_totals_12hour.py',
      '16hour': 'force_exact_totals_16hour.py',
      '24hour': 'force_exact_totals_24hour.py'
    };

    // Create a temporary Python wrapper to pass totals
    const wrapperScript = path.join(process.cwd(), `temp_wrapper_${timestamp}.py`);
    const wrapperContent = `
import sys
import json
sys.path.insert(0, '${process.cwd().replace(/\\/g, '\\\\')}')

from ${scriptMap[scriptType as keyof typeof scriptMap].replace('.py', '')} import ${scriptType === '24hour' ? 'force_exact_totals_24hour' : 'force_exact_totals'}

totals = json.loads('''${JSON.stringify(totals)}''')
input_file = '''${tempInputPath.replace(/\\/g, '\\\\')}'''
output_file = '''${tempOutputPath.replace(/\\/g, '\\\\')}'''

${scriptType === '24hour' ? 'force_exact_totals_24hour' : 'force_exact_totals'}(input_file, output_file, totals)
`;

    await writeFile(wrapperScript, wrapperContent);

    // Execute Python script
    const command = `python "${wrapperScript}"`;
    console.log('Executing:', command);

    const { stdout, stderr } = await execAsync(command, {
      maxBuffer: 10 * 1024 * 1024 // 10MB buffer for logs
    });

    // Log Python output
    if (stdout) {
      console.log('Python output:', stdout);
    }

    // Clean up wrapper script
    try {
      await unlink(wrapperScript);
    } catch {
      console.error('Error cleaning wrapper script');
    }

    if (stderr && !stderr.includes('DeprecationWarning')) {
      console.error('Python stderr:', stderr);
    }

    // Read processed file
    const processedBuffer = await readFile(tempOutputPath);

    // Clean up temp files
    try {
      await unlink(tempInputPath);
      await unlink(tempOutputPath);
    } catch (cleanupError) {
      console.error('Error cleaning up temp files:', cleanupError);
    }

    // Return processed file
    return new NextResponse(processedBuffer, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${outputFilename || 'processed.xlsx'}"`,
      },
    });

  } catch (error) {
    console.error('Error processing traffic:', error);

    // Clean up temp files on error
    try {
      if (tempInputPath) await unlink(tempInputPath);
      if (tempOutputPath) await unlink(tempOutputPath);
    } catch (cleanupError) {
      console.error('Error cleaning up after error:', cleanupError);
    }

    return NextResponse.json(
      { 
        success: false, 
        error: error instanceof Error ? error.message : 'Unknown error occurred',
        details: error instanceof Error ? error.stack : undefined
      },
      { status: 500 }
    );
  }
}

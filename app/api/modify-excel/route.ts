import { NextRequest, NextResponse } from 'next/server';
import { exec } from 'child_process';
import { promisify } from 'util';
import { writeFile, readFile, unlink } from 'fs/promises';
import path from 'path';

const execAsync = promisify(exec);

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get('file') as File;
    const percentage = formData.get('percentage') as string || '13';
    const operation = formData.get('operation') as string || 'increase';
    
    if (!file) {
      return NextResponse.json({ error: 'No file provided' }, { status: 400 });
    }

    const bytes = await file.arrayBuffer();
    const buffer = Buffer.from(bytes);
    
    let modifiedFile: Buffer;
    let logOutput: string;
    
    // Use different processing based on environment
    if (process.env.VERCEL_URL) {
      // Production: Use Vercel Python function
      const pythonResponse = await fetch(`${process.env.VERCEL_URL}/api/python/process-excel`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          file: buffer.toString('base64'),
          percentage: percentage,
          operation: operation
        })
      });

      if (!pythonResponse.ok) {
        const error = await pythonResponse.json();
        throw new Error(error.error || 'Python processing failed');
      }

      const result = await pythonResponse.json();
      modifiedFile = Buffer.from(result.file, 'base64');
      logOutput = result.log;
    } else {
      // Development: Use local Python script
      const inputPath = path.join(process.cwd(), `temp_input_${Date.now()}.xlsx`);
      const outputPath = path.join(process.cwd(), `temp_output_${Date.now()}.xlsx`);
      
      await writeFile(inputPath, buffer);
      
      const { stdout } = await execAsync(`python modify_excel.py "${inputPath}" "${outputPath}" ${percentage} ${operation}`);
      
      modifiedFile = await readFile(outputPath);
      logOutput = stdout;
      
      await unlink(inputPath);
      await unlink(outputPath);
    }
    
    return new NextResponse(new Uint8Array(modifiedFile), {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="Modified_Traffic_Counts.xlsx"',
        'X-Process-Log': Buffer.from(logOutput).toString('base64'),
      },
    });
  } catch (error) {
    console.error(error);
    return NextResponse.json({ error: 'Processing failed' }, { status: 500 });
  }
}

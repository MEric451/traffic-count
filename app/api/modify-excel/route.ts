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
    
    const inputPath = path.join(process.cwd(), `temp_input_${Date.now()}.xlsx`);
    const outputPath = path.join(process.cwd(), `temp_output_${Date.now()}.xlsx`);
    
    await writeFile(inputPath, buffer);
    
    const { stdout } = await execAsync(`python modify_excel.py "${inputPath}" "${outputPath}" ${percentage} ${operation}`);
    
    const modifiedFile = await readFile(outputPath);
    
    await unlink(inputPath);
    await unlink(outputPath);
    
    return new NextResponse(modifiedFile, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="Modified_Traffic_Counts.xlsx"',
        'X-Process-Log': Buffer.from(stdout).toString('base64'),
      },
    });
  } catch (error) {
    console.error(error);
    return NextResponse.json({ error: 'Processing failed' }, { status: 500 });
  }
}

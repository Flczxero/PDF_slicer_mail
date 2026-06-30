import { NextRequest, NextResponse } from 'next/server';
import * as xlsx from 'xlsx';

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const dataFile = formData.get('dataFile') as File | null;

    if (!dataFile) {
      return NextResponse.json({ error: 'Missing data file' }, { status: 400 });
    }

    const dataBuffer = Buffer.from(await dataFile.arrayBuffer());
    const workbook = xlsx.read(dataBuffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // Get headers and data
    const json = xlsx.utils.sheet_to_json<any>(worksheet, { header: 1 });
    let columns: string[] = [];
    if (json.length > 0) {
      columns = json[0] as string[];
    }
    
    // Also parse as array of objects for the frontend table
    const jsonData = xlsx.utils.sheet_to_json<any>(worksheet);

    return NextResponse.json({ columns, data: jsonData });
  } catch (error: any) {
    console.error('Error extracting columns:', error);
    return NextResponse.json({ error: error.message || 'Internal server error' }, { status: 500 });
  }
}

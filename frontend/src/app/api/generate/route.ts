import { NextRequest, NextResponse } from 'next/server';
import { parseExcelFiles } from '@/lib/excelParser';
import { DataAnalyzer } from '@/lib/dataAnalyzer';
import { generatePPT } from '@/lib/pptGenerator';
import { HOSPITAL_CONFIGS } from '@/lib/hospitalConfig';

export const runtime = 'nodejs';
export const maxDuration = 60;

export async function POST(request: NextRequest) {
  try {
    const formData   = await request.formData();
    const files      = formData.getAll('files') as File[];
    const hospitalKey  = (formData.get('hospitalKey')   as string) || 'mellow_sinsa';
    const year         = parseInt((formData.get('year')  as string) || '2026');
    const month        = parseInt((formData.get('month') as string) || '1');
    const hospitalName = (formData.get('hospitalName')   as string) || '';
    const teamName     = (formData.get('teamName')       as string) || '열정의시간 마케팅팀';

    if (!files.length) {
      return NextResponse.json({ error: '파일을 업로드하세요.' }, { status: 400 });
    }

    const cfg = HOSPITAL_CONFIGS[hospitalKey] || HOSPITAL_CONFIGS.mellow_sinsa;

    const parsedData = await parseExcelFiles(files);
    const analyzer   = new DataAnalyzer(cfg);
    const analysis   = analyzer.analyze(parsedData);

    const rc = {
      hospitalName: hospitalName || cfg.name,
      year, month, teamName,
    };

    const pptBuffer = await generatePPT(analysis, rc, cfg);
    const fname = encodeURIComponent(`${rc.hospitalName}_${year}년${month}월_매출분석보고서.pptx`);

    return new NextResponse(pptBuffer as unknown as BodyInit, {
      headers: {
        'Content-Type':        'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': `attachment; filename*=UTF-8''${fname}`,
      },
    });
  } catch (err) {
    console.error(err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

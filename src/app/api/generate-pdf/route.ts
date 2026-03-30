import { NextRequest, NextResponse } from "next/server";
import { HOLIDAYS_2026 } from "@/lib/holidays";
import path from "path";
import fs from "fs";
import os from "os";
import { execSync } from "child_process";

const ExcelJS = require("exceljs");

const SOFFICE_PATH =
  process.platform === "darwin"
    ? "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    : "soffice";

interface TimesheetRequest {
  employeeName: string;
  managerName: string;
  client: string;
  month: number;
  year: number;
  defaultHours: number;
  includeHolidays: boolean;
  dayOverrides?: Record<string, { regular: number; overtime: number }>;
}

interface WeekInfo {
  weekNumber: number;
  days: {
    date: Date;
    dayName: string;
    isWeekend: boolean;
    isHoliday: boolean;
    isOutsideMonth: boolean;
  }[];
  weekEnding: Date;
}

function utcNoon(d: Date): Date {
  return new Date(
    Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), 12, 0, 0)
  );
}

function getWeeksForMonth(year: number, month: number): WeekInfo[] {
  const holidaySet = new Set(HOLIDAYS_2026.map((h) => h.date));
  const dayNames = [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
  ];

  const firstDay = new Date(Date.UTC(year, month - 1, 1));
  const firstMonday = new Date(firstDay);
  firstMonday.setUTCDate(
    firstMonday.getUTCDate() - ((firstMonday.getUTCDay() + 6) % 7)
  );
  const lastDay = new Date(Date.UTC(year, month, 0));
  const lastSunday = new Date(lastDay);
  lastSunday.setUTCDate(lastSunday.getUTCDate() + ((7 - lastSunday.getUTCDay()) % 7));

  const weeks: WeekInfo[] = [];
  let monday = new Date(firstMonday);

  while (monday <= lastSunday) {
    const days = [];
    for (let i = 0; i < 7; i++) {
      const d = new Date(monday);
      d.setUTCDate(d.getUTCDate() + i);
      const iso = `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;
      days.push({
        date: d,
        dayName: dayNames[d.getUTCDay()],
        isWeekend: d.getUTCDay() === 0 || d.getUTCDay() === 6,
        isHoliday: holidaySet.has(iso),
        isOutsideMonth:
          d.getUTCMonth() !== month - 1 || d.getUTCFullYear() !== year,
      });
    }
    const hasWorkingDayInMonth = days.some(
      (d) => !d.isOutsideMonth && !d.isWeekend
    );
    if (hasWorkingDayInMonth) {
      weeks.push({
        weekNumber: weeks.length + 1,
        days,
        weekEnding: days[6].date,
      });
    }
    monday.setUTCDate(monday.getUTCDate() + 7);
  }
  return weeks;
}

function fmtDateUS(d: Date): string {
  return `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${d.getUTCFullYear()}`;
}

function computeHours(
  day: WeekInfo["days"][0],
  params: TimesheetRequest
): { regular: number; overtime: number } {
  if (day.isOutsideMonth || day.isWeekend) return { regular: 0, overtime: 0 };
  if (params.includeHolidays && day.isHoliday)
    return { regular: 0, overtime: 0 };
  const iso = `${day.date.getUTCFullYear()}-${String(day.date.getUTCMonth() + 1).padStart(2, "0")}-${String(day.date.getUTCDate()).padStart(2, "0")}`;
  const override = params.dayOverrides?.[iso];
  if (override)
    return { regular: override.regular, overtime: override.overtime };
  return { regular: params.defaultHours, overtime: 0 };
}

const MONTH_SHORT = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
];

const CENTURY_GOTHIC_FONT = {
  name: "Century Gothic",
  size: 9,
  color: { argb: "FF333333" },
};

async function fillTemplate(
  week: WeekInfo,
  params: TimesheetRequest
): Promise<Buffer> {
  const templatePath = path.join(process.cwd(), "public", "template.xlsx");
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(templatePath);

  const ws = wb.worksheets[0];

  while (wb.worksheets.length > 1) {
    wb.removeWorksheet(wb.worksheets[wb.worksheets.length - 1].id);
  }

  const setCenturyFont = (address: string, size?: number) => {
    ws.getCell(address).font = {
      ...CENTURY_GOTHIC_FONT,
      ...(size ? { size } : {}),
    };
  };

  ws.getCell("D4").value = params.employeeName;
  setCenturyFont("D4", 10);
  ws.getCell("D5").value = params.managerName;
  setCenturyFont("D5", 10);
  ws.getCell("D6").value = params.client;
  setCenturyFont("D6");
  ws.getCell("D7").value = utcNoon(week.weekEnding);
  setCenturyFont("D7");

  const dayRows = [10, 11, 12, 13, 14, 15, 16];
  let totalRegular = 0;
  let totalOvertime = 0;

  for (let i = 0; i < 7; i++) {
    const r = dayRows[i];
    const day = week.days[i];
    const { regular, overtime } = computeHours(day, params);
    const total = regular + overtime;

    ws.getCell(r, 3).value = utcNoon(day.date);
    ws.getCell(r, 5).value = regular > 0 ? regular : null;
    ws.getCell(r, 6).value = overtime > 0 ? overtime : null;
    ws.getCell(r, 7).value = total;
    setCenturyFont(`C${r}`);
    setCenturyFont(`E${r}`);
    setCenturyFont(`F${r}`);
    setCenturyFont(`G${r}`);

    totalRegular += regular;
    totalOvertime += overtime;
  }

  ws.getCell("E17").value = totalRegular;
  ws.getCell("F17").value = totalOvertime;
  ws.getCell("G17").value = totalRegular + totalOvertime;
  setCenturyFont("E17");
  setCenturyFont("F17");
  setCenturyFont("G17");

  ws.getCell("B21").value = params.employeeName;
  setCenturyFont("B21", 10);
  ws.getCell("F21").value = `Date: ${fmtDateUS(week.weekEnding)}`;
  setCenturyFont("F21", 10);

  ws.name = `Weekly Time Sheet_${MONTH_SHORT[params.month - 1]}_${week.weekNumber}`;

  const buffer = await wb.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

function convertToPDF(xlsxBuffer: Buffer, filename: string): Buffer {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "timesheet-"));
  const xlsxPath = path.join(tmpDir, `${filename}.xlsx`);
  fs.writeFileSync(xlsxPath, xlsxBuffer);
  const pdfFilter =
    'pdf:calc_pdf_Export:{"EmbedStandardFonts":{"type":"boolean","value":"true"}}';

  try {
    execSync(
      `"${SOFFICE_PATH}" --headless --convert-to '${pdfFilter}' --outdir "${tmpDir}" "${xlsxPath}"`,
      { timeout: 30000, stdio: "pipe" }
    );

    const pdfPath = path.join(tmpDir, `${filename}.pdf`);
    const pdfBuffer = fs.readFileSync(pdfPath);
    return pdfBuffer;
  } finally {
    try {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    } catch (_) {}
  }
}

export async function POST(request: NextRequest) {
  try {
    const params: TimesheetRequest = await request.json();
    const weeks = getWeeksForMonth(params.year, params.month);

    if (weeks.length === 0) {
      return NextResponse.json(
        { error: "No weeks found for the given month" },
        { status: 400 }
      );
    }

    const employeeSlug = params.employeeName
      .toLowerCase()
      .replace(/\s+/g, "");
    const monthSlug = MONTH_SHORT[params.month - 1].toLowerCase();

    const results: {
      filename: string;
      pdfData: string;
      xlsxData: string;
    }[] = [];

    for (const week of weeks) {
      const baseName = `timesheets_${employeeSlug}_${monthSlug}_week${week.weekNumber}`;
      const xlsxBuffer = await fillTemplate(week, params);
      const pdfBuffer = convertToPDF(xlsxBuffer, baseName);

      results.push({
        filename: baseName,
        pdfData: pdfBuffer.toString("base64"),
        xlsxData: xlsxBuffer.toString("base64"),
      });
    }

    return NextResponse.json({
      success: true,
      count: results.length,
      timesheets: results,
    });
  } catch (error: any) {
    console.error("Generation error:", error);
    return NextResponse.json(
      { error: error.message || "Failed to generate timesheets" },
      { status: 500 }
    );
  }
}

"use client";

import { useEffect, useMemo, useState } from "react";
import { HOLIDAYS_2026 } from "@/lib/holidays";

const MONTHS = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];

interface TimesheetResult {
  filename: string;
  pdfData: string;
  xlsxData: string;
}

export default function Home() {
  const [employeeName, setEmployeeName] = useState("");
  const [managerName, setManagerName] = useState("Murali Rudra Raju");
  const [client, setClient] = useState("Omnicell Inc");
  const [month, setMonth] = useState(new Date().getMonth() + 1);
  const [year, setYear] = useState(2026);
  const [defaultHours, setDefaultHours] = useState(8);
  const [includeHolidays, setIncludeHolidays] = useState(true);
  const [leaveDates, setLeaveDates] = useState<string[]>([]);
  const [loading, setLoading] = useState(false);
  const [results, setResults] = useState<TimesheetResult[]>([]);
  const [error, setError] = useState("");

  const holidaysInMonth = useMemo(
    () =>
      HOLIDAYS_2026.filter((h) => {
        const d = new Date(h.date + "T00:00:00");
        return d.getMonth() + 1 === month && d.getFullYear() === year;
      }),
    [month, year]
  );
  const holidaySet = useMemo(
    () => new Set(holidaysInMonth.map((h) => h.date)),
    [holidaysInMonth]
  );
  const daysInMonth = new Date(year, month, 0).getDate();
  const firstDayIndex = new Date(year, month - 1, 1).getDay();
  const calendarCells = Array.from(
    { length: firstDayIndex + daysInMonth },
    (_, idx) => {
      if (idx < firstDayIndex) return null;
      const dayNumber = idx - firstDayIndex + 1;
      const iso = `${year}-${String(month).padStart(2, "0")}-${String(dayNumber).padStart(2, "0")}`;
      const weekday = new Date(year, month - 1, dayNumber).getDay();
      return {
        iso,
        dayNumber,
        isWeekend: weekday === 0 || weekday === 6,
        isHoliday: holidaySet.has(iso),
      };
    }
  );

  function toggleLeaveDate(iso: string) {
    setLeaveDates((prev) =>
      prev.includes(iso) ? prev.filter((d) => d !== iso) : [...prev, iso]
    );
  }

  useEffect(() => {
    setLeaveDates([]);
  }, [month, year]);

  useEffect(() => {
    setLeaveDates((prev) =>
      {
        const next = prev.filter((iso) => {
          const d = new Date(iso + "T00:00:00");
          const isWeekend = d.getDay() === 0 || d.getDay() === 6;
          if (isWeekend) return false;
          if (includeHolidays && holidaySet.has(iso)) return false;
          return d.getMonth() + 1 === month && d.getFullYear() === year;
        });
        if (
          next.length === prev.length &&
          next.every((date, i) => date === prev[i])
        ) {
          return prev;
        }
        return next;
      }
    );
  }, [includeHolidays, month, year, holidaySet]);

  async function handleGenerate() {
    if (!employeeName.trim()) {
      setError("Please enter the employee name");
      return;
    }
    setError("");
    setLoading(true);
    setResults([]);

    try {
      const dayOverrides = leaveDates.reduce<Record<string, { regular: number; overtime: number }>>(
        (acc, iso) => {
          acc[iso] = { regular: 0, overtime: 0 };
          return acc;
        },
        {}
      );
      const res = await fetch("/api/generate-pdf", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          employeeName: employeeName.trim(),
          managerName: managerName.trim(),
          client: client.trim(),
          month,
          year,
          defaultHours,
          includeHolidays,
          dayOverrides,
        }),
      });

      const data = await res.json();
      if (!res.ok) {
        setError(data.error || "Failed to generate timesheets");
        return;
      }
      setResults(data.timesheets);
    } catch (err: any) {
      setError(err.message || "Something went wrong");
    } finally {
      setLoading(false);
    }
  }

  function downloadFile(base64: string, filename: string, mimeType: string) {
    const byteChars = atob(base64);
    const byteNums = new Array(byteChars.length);
    for (let i = 0; i < byteChars.length; i++) byteNums[i] = byteChars.charCodeAt(i);
    const blob = new Blob([new Uint8Array(byteNums)], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  }

  function downloadAllPDFs() {
    results.forEach((r, i) => {
      setTimeout(() => downloadFile(r.pdfData, `${r.filename}.pdf`, "application/pdf"), i * 300);
    });
  }

  function downloadAllXLSX() {
    results.forEach((r, i) => {
      setTimeout(
        () => downloadFile(r.xlsxData, `${r.filename}.xlsx`, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        i * 300
      );
    });
  }

  return (
    <div className="max-w-2xl mx-auto py-10 px-4">
      {/* Header */}
      <div className="text-center mb-8">
        <div className="flex items-center justify-center gap-2 mb-2">
          <div className="w-8 h-8 bg-teal-600 rounded-full flex items-center justify-center text-white font-bold text-sm">C</div>
          <h1 className="text-2xl font-bold text-gray-800">CodeVrt Timesheet Generator</h1>
        </div>
        <p className="text-gray-500 text-sm">Generate weekly timesheet PDFs &amp; Excel files for any month</p>
      </div>

      {/* Form */}
      <div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6 space-y-5">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">Employee Name *</label>
          <input
            type="text"
            value={employeeName}
            onChange={(e) => setEmployeeName(e.target.value)}
            placeholder="e.g. Repala Manideep"
            className="w-full border border-gray-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500 focus:border-transparent"
          />
        </div>

        <div className="grid grid-cols-2 gap-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Manager&apos;s Name</label>
            <input type="text" value={managerName} onChange={(e) => setManagerName(e.target.value)}
              className="w-full border border-gray-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500 focus:border-transparent" />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Client</label>
            <input type="text" value={client} onChange={(e) => setClient(e.target.value)}
              className="w-full border border-gray-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500 focus:border-transparent" />
          </div>
        </div>

        <div className="grid grid-cols-3 gap-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Month</label>
            <select value={month} onChange={(e) => setMonth(Number(e.target.value))}
              className="w-full border border-gray-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500 focus:border-transparent">
              {MONTHS.map((m, i) => <option key={i} value={i + 1}>{m}</option>)}
            </select>
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Year</label>
            <input type="number" value={year} onChange={(e) => setYear(Number(e.target.value))} min={2024} max={2030}
              className="w-full border border-gray-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500 focus:border-transparent" />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Daily Hours</label>
            <input type="number" value={defaultHours} onChange={(e) => setDefaultHours(Number(e.target.value))} min={0} max={24} step={0.5}
              className="w-full border border-gray-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500 focus:border-transparent" />
          </div>
        </div>

        <div className="flex items-center gap-3">
          <label className="relative inline-flex items-center cursor-pointer">
            <input type="checkbox" checked={includeHolidays} onChange={(e) => setIncludeHolidays(e.target.checked)} className="sr-only peer" />
            <div className="w-9 h-5 bg-gray-200 peer-focus:outline-none peer-focus:ring-2 peer-focus:ring-teal-300 rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-teal-600"></div>
          </label>
          <span className="text-sm text-gray-700">Account for holidays (0 hours on holidays)</span>
        </div>

        {includeHolidays && holidaysInMonth.length > 0 && (
          <div className="bg-amber-50 border border-amber-200 rounded-md p-3">
            <p className="text-xs font-medium text-amber-800 mb-1">Holidays in {MONTHS[month - 1]} {year}:</p>
            {holidaysInMonth.map((h) => (
              <p key={h.date} className="text-xs text-amber-700">
                {h.name} — {new Date(h.date + "T00:00:00").toLocaleDateString("en-US", { weekday: "long", month: "short", day: "numeric" })}
              </p>
            ))}
          </div>
        )}

        <div className="border border-gray-200 rounded-md p-3">
          <div className="flex items-center justify-between mb-2">
            <p className="text-sm font-medium text-gray-700">Leave Days (0 hours)</p>
            <button
              type="button"
              onClick={() => setLeaveDates([])}
              className="text-xs text-gray-500 hover:text-gray-700"
            >
              Clear
            </button>
          </div>
          <p className="text-xs text-gray-500 mb-3">
            Selected: {leaveDates.length} day{leaveDates.length === 1 ? "" : "s"}
          </p>
          <div className="grid grid-cols-7 gap-2 text-center text-xs font-medium text-gray-500 mb-2">
            <div>Sun</div>
            <div>Mon</div>
            <div>Tue</div>
            <div>Wed</div>
            <div>Thu</div>
            <div>Fri</div>
            <div>Sat</div>
          </div>
          <div className="grid grid-cols-7 gap-2">
            {calendarCells.map((cell, idx) => {
              if (!cell) {
                return <div key={`empty-${idx}`} className="h-10" />;
              }
              const selected = leaveDates.includes(cell.iso);
              const disabled = cell.isWeekend || (includeHolidays && cell.isHoliday);
              const baseClass =
                "h-10 rounded-md border text-sm font-medium transition-colors";
              const stateClass = disabled
                ? "border-gray-200 bg-gray-100 text-gray-400 cursor-not-allowed"
                : selected
                  ? "border-teal-600 bg-teal-600 text-white"
                  : "border-gray-300 bg-white text-gray-700 hover:border-teal-500 hover:text-teal-700";
              return (
                <button
                  key={cell.iso}
                  type="button"
                  disabled={disabled}
                  onClick={() => toggleLeaveDate(cell.iso)}
                  className={`${baseClass} ${stateClass}`}
                  title={
                    cell.isWeekend
                      ? "Weekend"
                      : cell.isHoliday
                        ? "Holiday"
                        : selected
                          ? "Leave selected"
                          : "Select as leave"
                  }
                >
                  {cell.dayNumber}
                </button>
              );
            })}
          </div>
        </div>

        {error && (
          <div className="bg-red-50 border border-red-200 rounded-md p-3 text-sm text-red-700">{error}</div>
        )}

        <button onClick={handleGenerate} disabled={loading}
          className="w-full bg-teal-600 hover:bg-teal-700 disabled:bg-teal-400 text-white font-medium py-2.5 px-4 rounded-md text-sm transition-colors">
          {loading ? "Generating..." : "Generate Timesheets"}
        </button>
      </div>

      {/* Results */}
      {results.length > 0 && (
        <div className="mt-6 bg-white rounded-lg shadow-sm border border-gray-200 p-6">
          <div className="flex items-center justify-between mb-4">
            <h2 className="text-lg font-semibold text-gray-800">
              Generated Timesheets ({results.length} weeks)
            </h2>
            <div className="flex gap-3">
              <button onClick={downloadAllPDFs} className="text-sm text-teal-600 hover:text-teal-700 font-medium">
                All PDFs
              </button>
              <button onClick={downloadAllXLSX} className="text-sm text-blue-600 hover:text-blue-700 font-medium">
                All Excel
              </button>
            </div>
          </div>
          <div className="space-y-2">
            {results.map((r) => (
              <div key={r.filename} className="flex items-center justify-between bg-gray-50 rounded-md px-4 py-3">
                <div className="flex items-center gap-3">
                  <svg className="w-5 h-5 text-red-500 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                    <path d="M4 18h12a2 2 0 002-2V6l-4-4H4a2 2 0 00-2 2v12a2 2 0 002 2zm8-14l4 4h-4V4z" />
                  </svg>
                  <span className="text-sm text-gray-700">{r.filename}</span>
                </div>
                <div className="flex gap-3">
                  <button onClick={() => downloadFile(r.pdfData, `${r.filename}.pdf`, "application/pdf")}
                    className="text-sm text-teal-600 hover:text-teal-700 font-medium">PDF</button>
                  <button onClick={() => downloadFile(r.xlsxData, `${r.filename}.xlsx`, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
                    className="text-sm text-blue-600 hover:text-blue-700 font-medium">Excel</button>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      <p className="text-center text-xs text-gray-400 mt-8">Wintech Services India Pvt Ltd</p>
    </div>
  );
}

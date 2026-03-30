// 2026 Holiday List - Wintech Services India Pvt Ltd
export interface Holiday {
  name: string;
  date: string; // YYYY-MM-DD
}

export const HOLIDAYS_2026: Holiday[] = [
  { name: "New Year Day", date: "2026-01-01" },
  { name: "Bhogi", date: "2026-01-13" },
  { name: "Makara Sankranthi", date: "2026-01-14" },
  { name: "Republic Day", date: "2026-01-26" },
  { name: "Holi", date: "2026-03-04" },
  { name: "Ugadi", date: "2026-03-19" },
  { name: "Ramzan Eid", date: "2026-03-20" },
  { name: "Independence Day", date: "2026-08-15" },
  { name: "Vinayaka Chaturthi", date: "2026-09-14" },
  { name: "Dussehra (Vijayadashami)", date: "2026-10-20" },
  { name: "Diwali (Deepavali)", date: "2026-11-08" },
  { name: "Christmas Day", date: "2026-12-25" },
];

export function isHoliday(dateStr: string): Holiday | undefined {
  return HOLIDAYS_2026.find((h) => h.date === dateStr);
}

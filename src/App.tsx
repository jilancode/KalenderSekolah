/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';

// Types
export type CategoryKey = 'holiday' | 'school' | 'teacher' | 'important';

export interface KaldikEvent {
  id: string;
  title: string;
  category: CategoryKey;
  startDate: string;
  endDate: string;
  time?: string;
  location?: string;
  pic?: string;
  description?: string;
  isNonEffective?: boolean;
  isDemo?: boolean;
}

export interface AppSettings {
  schoolName: string;
  schoolNPSN: string;
  principalName: string;
  principalNIP: string;
  address: string;
  academicYear: string;
  sem1Start: string;
  sem1End: string;
  sem2Start: string;
  sem2End: string;
  cityName: string;
  developerRole: string;
  developerName: string;
  developerNIP: string;
  signDate: string;
  logoLeft: string;
  logoRight: string;
}

export interface AppData {
  settings: AppSettings;
  events: KaldikEvent[];
  hebOverrides: Record<string, number>;
}

// Storage Key
const STORAGE_KEY = 'KALDIK_APP_DATA_V1';

// Base64 Clean SVGs for instant canvas rendering (prevents html2canvas hanging)
const DEFAULT_LOGO_LEFT = "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAxMDAgMTIwIiB3aWR0aD0iODAiIGhlaWdodD0iOTYiPjxwYXRoIGQ9Ik01MCA1IEw5MCAyNSBWNjUgQzkwIDkwIDUwIDExNSA1MCAxMTUgQzUwIDExNSAxMCA5MCAxMCA2NSBWMjUgWiIgZmlsbD0iIzFlM2E4YSIgc3Ryb2tlPSIjZmJmYjI0IiBzdHJva2Utd2lkdGg9IjQiLz48cG9seWdvbiBwb2ludHM9IjUwLDI1IDU3LDQwIDczLDQwIDYwLDUwIDY1LDY2IDUwLDU1IDM1LDY2IDQwLDUwIDI3LDQwIDQzLDQwIiBmaWxsPSIjZmJmYjI0Ii8+PHBhdGggZD0iTTI1IDgwIEM0MCA5NSA2MCA5NSA3NSA4MCIgc3Ryb2tlPSIjZmZmZmZmIiBzdHJva2Utd2lkdGg9IjMiIGZpbGw9Im5vbmUiLz48L3N2Zz4=";
const DEFAULT_LOGO_RIGHT = "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAxMDAgMTAwIiB3aWR0aD0iODAiIGhlaWdodD0iODAiPjxjaXJjbGUgY3g9IjUwIiBjeT0iNTAiIHI9IjQ1IiBmaWxsPSIjMjU2M2ViIiBzdHJva2U9IiMxZDRlZDgiIHN0cm9rZS13aWR0aD0iMyIvPjxwYXRoIGQ9Ik0yNSA2MCBRNTAgNDUgNzUgNjAgVjcwIFE1MCA1NSAyNSA3MCBaIiBmaWxsPSIjZmZmZmZmIi8+PHBhdGggZD0iTTUwIDM1IEw3NSA0OCBMNTAgNjAgTDI1IDQ4IFoiIGZpbGw9IiNmYmJmMjQiLz48Y2lyY2xlIGN4PSI1MCIgY3k9IjI4IiByPSI2IiBmaWxsPSIjZmJmYjI0Ii8+PC9zdmc+";

// Category Definitions
export const CATEGORIES: Record<CategoryKey, {
  id: CategoryKey;
  name: string;
  color: string;
  bgClass: string;
  textClass: string;
  badgeClass: string;
  dotColor: string;
}> = {
  holiday: {
    id: 'holiday',
    name: 'Libur Nasional',
    color: 'rose',
    bgClass: 'bg-rose-500',
    textClass: 'text-rose-700',
    badgeClass: 'bg-rose-100 text-rose-800 border-rose-300',
    dotColor: '#f43f5e'
  },
  school: {
    id: 'school',
    name: 'Agenda Sekolah',
    color: 'sky',
    bgClass: 'bg-sky-500',
    textClass: 'text-sky-700',
    badgeClass: 'bg-sky-100 text-sky-800 border-sky-300',
    dotColor: '#0ea5e9'
  },
  teacher: {
    id: 'teacher',
    name: 'Agenda Guru',
    color: 'purple',
    bgClass: 'bg-purple-500',
    textClass: 'text-purple-700',
    badgeClass: 'bg-purple-100 text-purple-800 border-purple-300',
    dotColor: '#a855f7'
  },
  important: {
    id: 'important',
    name: 'Peristiwa Penting',
    color: 'amber',
    bgClass: 'bg-amber-500',
    textClass: 'text-amber-700',
    badgeClass: 'bg-amber-100 text-amber-800 border-amber-300',
    dotColor: '#f59e0b'
  }
};

// Default Initial Settings
const DEFAULT_SETTINGS: AppSettings = {
  schoolName: "SD NEGERI HARAPAN BANGSA 01",
  schoolNPSN: "20109988",
  principalName: "Dr. H. Ahmad Dahlan, M.Pd.",
  principalNIP: "19720512 199803 1 004",
  address: "Jl. Pendidikan No. 45, Kecamatan Kota, Kabupaten Bandung",
  academicYear: "2024/2025",
  sem1Start: "2024-07-15",
  sem1End: "2024-12-20",
  sem2Start: "2025-01-06",
  sem2End: "2025-06-20",
  cityName: "Bandung",
  developerRole: "Tim Pengembang Kurikulum",
  developerName: "Panitia Kalender Pendidikan",
  developerNIP: "",
  signDate: "",
  logoLeft: DEFAULT_LOGO_LEFT,
  logoRight: DEFAULT_LOGO_RIGHT
};

// Comprehensive Indonesian National Holidays Database & Generator by Year
export interface HolidayDefinition {
  title: string;
  startDate: string;
  endDate: string;
  description: string;
}

// Fixed-date national holidays
export function getFixedHolidays(year: number): HolidayDefinition[] {
  return [
    { title: `Tahun Baru ${year} Masehi`, startDate: `${year}-01-01`, endDate: `${year}-01-01`, description: 'Tahun Baru Masehi' },
    { title: 'Hari Buruh Internasional', startDate: `${year}-05-01`, endDate: `${year}-05-01`, description: 'Hari Buruh Internasional (May Day)' },
    { title: 'Hari Lahir Pancasila', startDate: `${year}-06-01`, endDate: `${year}-06-01`, description: 'Hari Lahir Pancasila' },
    { title: `HUT Proklamasi Kemerdekaan RI ke-${year - 1945}`, startDate: `${year}-08-17`, endDate: `${year}-08-17`, description: 'Hari Kemerdekaan Republik Indonesia' },
    { title: 'Hari Raya Natal', startDate: `${year}-12-25`, endDate: `${year}-12-25`, description: 'Hari Raya Natal' },
    { title: 'Cuti Bersama Hari Raya Natal', startDate: `${year}-12-26`, endDate: `${year}-12-26`, description: 'Cuti Bersama Natal' },
  ];
}

// Variable Islamic, Lunar, Saka, and Christian Holidays based on Indonesian Government Decrees (SKB 3 Menteri)
export const VARIABLE_HOLIDAYS_BY_YEAR: Record<number, HolidayDefinition[]> = {
  2023: [
    { title: 'Tahun Baru Imlek 2574 Kongzili', startDate: '2023-01-22', endDate: '2023-01-22', description: 'Tahun Baru Imlek' },
    { title: 'Isra Mikraj Nabi Muhammad SAW', startDate: '2023-02-18', endDate: '2023-02-18', description: 'Hari Besar Keagamaan Islam' },
    { title: 'Hari Suci Nyepi (Tahun Baru Saka 1945)', startDate: '2023-03-22', endDate: '2023-03-22', description: 'Hari Suci Nyepi' },
    { title: 'Wafat Isa Almasih', startDate: '2023-04-07', endDate: '2023-04-07', description: 'Wafat Yesus Kristus' },
    { title: 'Hari Raya Idul Fitri 1444 H', startDate: '2023-04-22', endDate: '2023-04-23', description: 'Libur Nasional Idul Fitri 1444 H' },
    { title: 'Kenaikan Isa Almasih', startDate: '2023-05-18', endDate: '2023-05-18', description: 'Kenaikan Yesus Kristus' },
    { title: 'Hari Raya Waisak 2567 BE', startDate: '2023-06-04', endDate: '2023-06-04', description: 'Hari Raya Waisak' },
    { title: 'Hari Raya Idul Adha 1444 H', startDate: '2023-06-29', endDate: '2023-06-29', description: 'Hari Raya Kurban Idul Adha' },
    { title: 'Tahun Baru Islam 1445 H', startDate: '2023-07-19', endDate: '2023-07-19', description: '1 Muharram 1445 H' },
    { title: 'Maulid Nabi Muhammad SAW', startDate: '2023-09-28', endDate: '2023-09-28', description: 'Peringatan Maulid Nabi SAW' },
  ],
  2024: [
    { title: 'Isra Mikraj Nabi Muhammad SAW', startDate: '2024-02-08', endDate: '2024-02-08', description: 'Hari Besar Keagamaan Islam' },
    { title: 'Tahun Baru Imlek 2575 Kongzili', startDate: '2024-02-10', endDate: '2024-02-10', description: 'Tahun Baru Imlek 2575' },
    { title: 'Hari Suci Nyepi (Tahun Baru Saka 1946)', startDate: '2024-03-11', endDate: '2024-03-11', description: 'Hari Suci Nyepi' },
    { title: 'Wafat Isa Almasih', startDate: '2024-03-29', endDate: '2024-03-29', description: 'Wafat Yesus Kristus (Jumat Agung)' },
    { title: 'Hari Paskah', startDate: '2024-03-31', endDate: '2024-03-31', description: 'Kebangkitan Yesus Kristus' },
    { title: 'Hari Raya Idul Fitri 1445 H', startDate: '2024-04-10', endDate: '2024-04-11', description: 'Libur Nasional Idul Fitri 1445 H' },
    { title: 'Kenaikan Isa Almasih', startDate: '2024-05-09', endDate: '2024-05-09', description: 'Kenaikan Yesus Kristus' },
    { title: 'Hari Raya Waisak 2568 BE', startDate: '2024-05-23', endDate: '2024-05-23', description: 'Hari Raya Tri Suci Waisak' },
    { title: 'Hari Raya Idul Adha 1445 H', startDate: '2024-06-17', endDate: '2024-06-17', description: 'Hari Raya Idul Adha 1445 H' },
    { title: 'Tahun Baru Islam 1446 H', startDate: '2024-07-07', endDate: '2024-07-07', description: '1 Muharram 1446 H' },
    { title: 'Maulid Nabi Muhammad SAW', startDate: '2024-09-16', endDate: '2024-09-16', description: 'Peringatan Maulid Nabi SAW' },
  ],
  2025: [
    { title: 'Isra Mikraj Nabi Muhammad SAW', startDate: '2025-01-27', endDate: '2025-01-27', description: 'Hari Besar Keagamaan Islam' },
    { title: 'Tahun Baru Imlek 2576 Kongzili', startDate: '2025-01-29', endDate: '2025-01-29', description: 'Tahun Baru Imlek 2576' },
    { title: 'Hari Suci Nyepi (Tahun Baru Saka 1947)', startDate: '2025-03-29', endDate: '2025-03-29', description: 'Hari Suci Nyepi Saka 1947' },
    { title: 'Hari Raya Idul Fitri 1446 H', startDate: '2025-03-31', endDate: '2025-04-01', description: 'Libur Nasional Idul Fitri 1446 H' },
    { title: 'Wafat Yesus Kristus', startDate: '2025-04-18', endDate: '2025-04-18', description: 'Jumat Agung Wafat Yesus Kristus' },
    { title: 'Hari Raya Waisak 2569 BE', startDate: '2025-05-12', endDate: '2025-05-12', description: 'Hari Raya Tri Suci Waisak' },
    { title: 'Kenaikan Yesus Kristus', startDate: '2025-05-29', endDate: '2025-05-29', description: 'Kenaikan Yesus Kristus' },
    { title: 'Hari Raya Idul Adha 1446 H', startDate: '2025-06-07', endDate: '2025-06-07', description: 'Hari Raya Kurban Idul Adha 1446 H' },
    { title: 'Tahun Baru Islam 1447 H', startDate: '2025-06-27', endDate: '2025-06-27', description: '1 Muharram 1447 H' },
    { title: 'Maulid Nabi Muhammad SAW', startDate: '2025-09-05', endDate: '2025-09-05', description: 'Peringatan Maulid Nabi Muhammad SAW' },
  ],
  2026: [
    { title: 'Isra Mikraj Nabi Muhammad SAW', startDate: '2026-01-16', endDate: '2026-01-16', description: 'Hari Besar Keagamaan Islam' },
    { title: 'Tahun Baru Imlek 2577 Kongzili', startDate: '2026-02-17', endDate: '2026-02-17', description: 'Tahun Baru Imlek 2577' },
    { title: 'Hari Suci Nyepi (Tahun Baru Saka 1948)', startDate: '2026-03-19', endDate: '2026-03-19', description: 'Hari Suci Nyepi Saka 1948' },
    { title: 'Hari Raya Idul Fitri 1447 H', startDate: '2026-03-20', endDate: '2026-03-21', description: 'Libur Nasional Idul Fitri 1447 H' },
    { title: 'Wafat Yesus Kristus', startDate: '2026-04-03', endDate: '2026-04-03', description: 'Jumat Agung Wafat Yesus Kristus' },
    { title: 'Kenaikan Yesus Kristus', startDate: '2026-05-14', endDate: '2026-05-14', description: 'Kenaikan Yesus Kristus' },
    { title: 'Hari Raya Idul Adha 1447 H', startDate: '2026-05-27', endDate: '2026-05-27', description: 'Hari Raya Kurban Idul Adha 1447 H' },
    { title: 'Hari Raya Waisak 2570 BE', startDate: '2026-05-31', endDate: '2026-05-31', description: 'Hari Raya Tri Suci Waisak' },
    { title: 'Tahun Baru Islam 1448 H', startDate: '2026-06-16', endDate: '2026-06-16', description: '1 Muharram 1448 H' },
    { title: 'Maulid Nabi Muhammad SAW', startDate: '2026-08-25', endDate: '2026-08-25', description: 'Peringatan Maulid Nabi Muhammad SAW' },
  ],
  2027: [
    { title: 'Isra Mikraj Nabi Muhammad SAW', startDate: '2027-01-06', endDate: '2027-01-06', description: 'Hari Besar Keagamaan Islam' },
    { title: 'Tahun Baru Imlek 2578 Kongzili', startDate: '2027-02-06', endDate: '2027-02-06', description: 'Tahun Baru Imlek 2578' },
    { title: 'Hari Raya Idul Fitri 1448 H', startDate: '2027-03-09', endDate: '2027-03-10', description: 'Libur Nasional Idul Fitri 1448 H' },
    { title: 'Wafat Yesus Kristus', startDate: '2027-03-26', endDate: '2027-03-26', description: 'Jumat Agung Wafat Yesus Kristus' },
    { title: 'Hari Suci Nyepi (Tahun Baru Saka 1949)', startDate: '2027-04-07', endDate: '2027-04-07', description: 'Hari Suci Nyepi Saka 1949' },
    { title: 'Kenaikan Yesus Kristus', startDate: '2027-05-06', endDate: '2027-05-06', description: 'Kenaikan Yesus Kristus' },
    { title: 'Hari Raya Idul Adha 1448 H', startDate: '2027-05-16', endDate: '2027-05-16', description: 'Hari Raya Idul Adha 1448 H' },
    { title: 'Hari Raya Waisak 2571 BE', startDate: '2027-05-20', endDate: '2027-05-20', description: 'Hari Raya Tri Suci Waisak' },
    { title: 'Tahun Baru Islam 1449 H', startDate: '2027-06-06', endDate: '2027-06-06', description: '1 Muharram 1449 H' },
    { title: 'Maulid Nabi Muhammad SAW', startDate: '2027-08-15', endDate: '2027-08-15', description: 'Peringatan Maulid Nabi Muhammad SAW' },
  ],
  2028: [
    { title: 'Isra Mikraj Nabi Muhammad SAW', startDate: '2028-01-25', endDate: '2028-01-25', description: 'Hari Besar Keagamaan Islam' },
    { title: 'Tahun Baru Imlek 2579 Kongzili', startDate: '2028-01-26', endDate: '2028-01-26', description: 'Tahun Baru Imlek 2579' },
    { title: 'Hari Raya Idul Fitri 1449 H', startDate: '2028-02-27', endDate: '2028-02-28', description: 'Libur Nasional Idul Fitri 1449 H' },
    { title: 'Hari Suci Nyepi (Tahun Baru Saka 1950)', startDate: '2028-03-26', endDate: '2028-03-26', description: 'Hari Suci Nyepi Saka 1950' },
    { title: 'Wafat Yesus Kristus', startDate: '2028-04-14', endDate: '2028-04-14', description: 'Jumat Agung Wafat Yesus Kristus' },
    { title: 'Hari Raya Idul Adha 1449 H', startDate: '2028-05-05', endDate: '2028-05-05', description: 'Hari Raya Idul Adha 1449 H' },
    { title: 'Hari Raya Waisak 2572 BE', startDate: '2028-05-09', endDate: '2028-05-09', description: 'Hari Raya Tri Suci Waisak' },
    { title: 'Tahun Baru Islam 1450 H', startDate: '2028-05-25', endDate: '2028-05-25', description: '1 Muharram 1450 H' },
    { title: 'Kenaikan Yesus Kristus', startDate: '2028-05-25', endDate: '2028-05-25', description: 'Kenaikan Yesus Kristus' },
    { title: 'Maulid Nabi Muhammad SAW', startDate: '2028-08-03', endDate: '2028-08-03', description: 'Peringatan Maulid Nabi Muhammad SAW' },
  ]
};

// Generates accurate national holidays for an academic year (e.g. 2024/2025)
export function getHolidaysForAcademicYear(
  academicYearStr: string,
  sem1Start?: string,
  sem2End?: string
): KaldikEvent[] {
  const match = academicYearStr.match(/(\d{4})[^\d]+(\d{4})/);
  let startYear = 2024;
  let endYear = 2025;

  if (match) {
    startYear = parseInt(match[1], 10);
    endYear = parseInt(match[2], 10);
  } else {
    const singleYear = parseInt(academicYearStr.replace(/\D/g, ''), 10);
    if (!isNaN(singleYear) && singleYear >= 2000 && singleYear <= 2100) {
      startYear = singleYear;
      endYear = singleYear + 1;
    }
  }

  // Academic year default range
  const rangeStart = sem1Start || `${startYear}-07-01`;
  const rangeEnd = sem2End || `${endYear}-06-30`;

  const allCandidates: HolidayDefinition[] = [];

  // Year 1 (July - December)
  allCandidates.push(...getFixedHolidays(startYear));
  if (VARIABLE_HOLIDAYS_BY_YEAR[startYear]) {
    allCandidates.push(...VARIABLE_HOLIDAYS_BY_YEAR[startYear]);
  }

  // Year 2 (January - June)
  allCandidates.push(...getFixedHolidays(endYear));
  if (VARIABLE_HOLIDAYS_BY_YEAR[endYear]) {
    allCandidates.push(...VARIABLE_HOLIDAYS_BY_YEAR[endYear]);
  }

  // Filter within academic year boundaries
  const filtered = allCandidates.filter(h => {
    return h.startDate >= rangeStart && h.startDate <= rangeEnd;
  });

  // Sort chronologically
  filtered.sort((a, b) => a.startDate.localeCompare(b.startDate));

  return filtered.map((h, idx) => ({
    id: `ev-hol-${startYear}-${endYear}-${idx + 1}`,
    title: h.title,
    category: 'holiday',
    startDate: h.startDate,
    endDate: h.endDate,
    description: h.description,
    isDemo: false
  }));
}

// Robust Excel date parsing helper
export function parseExcelDate(val: any): string {
  if (!val) return '';
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  if (typeof val === 'number') {
    // Excel serial date code
    const d = new Date(Math.round((val - 25569) * 86400 * 1000));
    if (!isNaN(d.getTime())) {
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      return `${y}-${m}-${day}`;
    }
  }
  const str = String(val).trim();
  if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(str)) {
    const [y, m, d] = str.split('-');
    return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }
  const dmyMatch = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
  if (dmyMatch) {
    const [, d, m, y] = dmyMatch;
    return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    const y = parsed.getFullYear();
    const m = String(parsed.getMonth() + 1).padStart(2, '0');
    const d = String(parsed.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return '';
}

// Flexible header field finder
export function getFieldVal(row: Record<string, any>, possibleKeys: string[]): string {
  const rowKeys = Object.keys(row);
  for (const targetKey of possibleKeys) {
    const foundKey = rowKeys.find(k =>
      k.toLowerCase().replace(/[^a-z0-9]/g, '').includes(targetKey.toLowerCase().replace(/[^a-z0-9]/g, ''))
    );
    if (foundKey && row[foundKey] !== undefined && row[foundKey] !== null) {
      return String(row[foundKey]).trim();
    }
  }
  return '';
}

// Initial Demo Events
const DEMO_EVENTS: KaldikEvent[] = [
  // Holidays for 2024/2025 automatically fetched
  ...getHolidaysForAcademicYear('2024/2025', '2024-07-15', '2025-06-20'),

  // Agenda Sekolah
  { id: 'ev-14', title: 'Masa Pengenalan Lingkungan Sekolah (MPLS)', category: 'school', startDate: '2024-07-15', endDate: '2024-07-17', time: '07:30 - 11:30', location: 'Lapangan Sekolah', pic: 'Wakasek Kesiswaan', description: 'Pengenalan lingkungan bagi siswa baru SD', isNonEffective: true, isDemo: true },
  { id: 'ev-15', title: 'Asesmen Nasional Berbasis Komputer (ANBK) SD', category: 'school', startDate: '2024-09-23', endDate: '2024-09-27', time: '08:00 - 12:00', location: 'Lab Komputer', pic: 'Tim IT & ANBK', description: 'Pelaksanaan ANBK Kelas 5', isNonEffective: false, isDemo: true },
  { id: 'ev-16', title: 'Sumatif Akhir Semester 1 (SAS 1)', category: 'school', startDate: '2024-12-02', endDate: '2024-12-07', time: '07:30 - 11:00', location: 'Ruang Kelas 1-6', pic: 'Panitia Asesmen', description: 'Penilaian hasil belajar Semester 1', isNonEffective: false, isDemo: true },
  { id: 'ev-17', title: 'Pembagian Rapor Semester 1', category: 'school', startDate: '2024-12-19', endDate: '2024-12-19', time: '08:00 - 11:30', location: 'Kelas Masing-masing', pic: 'Wali Kelas', description: 'Penyerahan Laporan Hasil Belajar ke Orang Tua', isNonEffective: false, isDemo: true },
  { id: 'ev-18', title: 'Libur Semester 1 (Ganjil)', category: 'school', startDate: '2024-12-23', endDate: '2025-01-04', location: '-', pic: 'Sekolah', description: 'Libur pembelajaran akhir semester 1', isNonEffective: true, isDemo: true },
  { id: 'ev-19', title: 'Sumatif Akhir Tahun (SAT) Semester 2', category: 'school', startDate: '2025-06-09', endDate: '2025-06-14', time: '07:30 - 11:00', location: 'Ruang Kelas 1-6', pic: 'Panitia Asesmen', description: 'Ujian Kenaikan Kelas', isNonEffective: false, isDemo: true },
  { id: 'ev-20', title: 'Pembagian Rapor & Kenaikan Kelas', category: 'school', startDate: '2025-06-20', endDate: '2025-06-20', time: '08:00 - 11:30', location: 'Aula Utama', pic: 'Wali Kelas & Kepala Sekolah', description: 'Penyerahan Rapor Genap & Wisuda Kelas 6', isNonEffective: false, isDemo: true },
  { id: 'ev-21', title: 'Libur Akhir Tahun Pelajaran', category: 'school', startDate: '2025-06-23', endDate: '2025-07-12', location: '-', pic: 'Sekolah', description: 'Libur akhir tahun pelajaran 2024/2025', isNonEffective: true, isDemo: true },

  // Agenda Guru
  { id: 'ev-22', title: 'Rapat Kerja Awal Tahun Pelajaran', category: 'teacher', startDate: '2024-07-12', endDate: '2024-07-12', time: '08:00 - 15:00', location: 'Ruang Guru', pic: 'Kepala Sekolah', description: 'Penyusunan Kurikulum Satuan Pendidikan (KSP) & Modul Ajar', isDemo: true },
  { id: 'ev-23', title: 'Pelatihan Kombel & PMM Berdiferensiasi', category: 'teacher', startDate: '2024-08-14', endDate: '2024-08-14', time: '13:00 - 15:30', location: 'Ruang Multimedia', pic: 'Ketua Komunitas Belajar', description: 'Berbagi praktik baik pembelajaran berdiferensiasi', isDemo: true },
  { id: 'ev-24', title: 'Workshop Penyusunan Soal HOTS SD', category: 'teacher', startDate: '2024-10-10', endDate: '2024-10-10', time: '08:30 - 14:00', location: 'Aula Sekolah', pic: 'Narasumber Dinas Pendidikan', description: 'Pelatihan pembuatan asesmen berbasis HOTS', isDemo: true },
  { id: 'ev-25', title: 'Evaluasi Pembelajaran & Progam Sem 2', category: 'teacher', startDate: '2025-01-10', endDate: '2025-01-10', time: '13:00 - 15:30', location: 'Ruang Guru', pic: 'Kepala Sekolah', description: 'Rapat koordinasi dan evaluasi perangkat ajar', isDemo: true },

  // Peristiwa Penting
  { id: 'ev-26', title: 'Hari Anak Nasional', category: 'important', startDate: '2024-07-23', endDate: '2024-07-23', description: 'Peringatan Hari Anak Nasional di Sekolah', isDemo: true },
  { id: 'ev-27', title: 'Hari Pramuka', category: 'important', startDate: '2024-08-14', endDate: '2024-08-14', description: 'Upacara Peringatan Hari Pramuka', isDemo: true },
  { id: 'ev-28', title: 'Hari Batik Nasional', category: 'important', startDate: '2024-10-02', endDate: '2024-10-02', description: 'Seluruh warga sekolah mengenakan seragam batik', isDemo: true },
  { id: 'ev-29', title: 'Hari Guru Nasional (HGN) & HUT PGRI', category: 'important', startDate: '2024-11-25', endDate: '2024-11-25', description: 'Upacara peringatan HGN dan apresiasi guru', isDemo: true },
  { id: 'ev-30', title: 'Hari Pendidikan Nasional (Hardiknas)', category: 'important', startDate: '2025-05-02', endDate: '2025-05-02', description: 'Upacara Bendera Hardiknas', isDemo: true }
];

// Indonesian Date Helpers
export function formatDateIndo(dateStr: string): string {
  if (!dateStr) return '-';
  const parts = dateStr.split('-');
  if (parts.length < 3) return dateStr;
  const [y, m, d] = parts;
  return `${d}/${m}/${y}`;
}

export function formatLongDateIndo(dateStr: string): string {
  if (!dateStr) return '-';
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return dateStr;
  const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  return `${days[d.getDay()]}, ${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;
}

export function getMonthNameIndo(monthIdx: number): string {
  const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  return months[monthIdx] || '';
}

export function getLocalDateString(d: Date = new Date()): string {
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

export function isDateInRange(targetDateStr: string, startStr: string, endStr: string): boolean {
  return targetDateStr >= startStr && targetDateStr <= endStr;
}

export function getDatesInRange(startStr: string, endStr: string): string[] {
  const dates: string[] = [];
  const curr = new Date(startStr);
  const end = new Date(endStr);
  while (curr <= end) {
    const y = curr.getFullYear();
    const m = String(curr.getMonth() + 1).padStart(2, '0');
    const d = String(curr.getDate()).padStart(2, '0');
    dates.push(`${y}-${m}-${d}`);
    curr.setDate(curr.getDate() + 1);
  }
  return dates;
}

export default function App() {
  // Master State
  const [appData, setAppData] = useState<AppData>(() => {
    try {
      const stored = localStorage.getItem(STORAGE_KEY);
      if (stored) {
        const parsed = JSON.parse(stored);
        let logoLeft = parsed.settings?.logoLeft;
        let logoRight = parsed.settings?.logoRight;
        if (!logoLeft || logoLeft.includes('utf8') || logoLeft.includes('<svg')) {
          logoLeft = DEFAULT_LOGO_LEFT;
        }
        if (!logoRight || logoRight.includes('utf8') || logoRight.includes('<svg')) {
          logoRight = DEFAULT_LOGO_RIGHT;
        }

        return {
          settings: { ...DEFAULT_SETTINGS, ...parsed.settings, logoLeft, logoRight },
          events: parsed.events || [...DEMO_EVENTS],
          hebOverrides: parsed.hebOverrides || {}
        };
      }
    } catch (e) {
      console.error('Failed to load stored Kaldik data:', e);
    }
    return {
      settings: { ...DEFAULT_SETTINGS },
      events: [...DEMO_EVENTS],
      hebOverrides: {}
    };
  });

  // UI Navigation State
  const [activeTab, setActiveTab] = useState<string>('dashboard');
  const [activeSemester, setActiveSemester] = useState<1 | 2>(1);

  // Derived academic year boundaries strictly following appData.settings
  const tpYears = useMemo(() => {
    let startYear = 2024;
    let endYear = 2025;

    const match = appData.settings.academicYear.match(/(\d{4})[^\d]+(\d{4})/);
    if (match) {
      startYear = parseInt(match[1], 10);
      endYear = parseInt(match[2], 10);
    } else {
      const single = appData.settings.academicYear.match(/(\d{4})/);
      if (single) {
        startYear = parseInt(single[1], 10);
        endYear = startYear + 1;
      }
    }

    if (appData.settings.sem1Start) {
      const s1 = new Date(appData.settings.sem1Start);
      if (!isNaN(s1.getFullYear()) && s1.getFullYear() >= 2000 && s1.getFullYear() <= 2100) {
        startYear = s1.getFullYear();
      }
    }

    if (appData.settings.sem2End) {
      const s2 = new Date(appData.settings.sem2End);
      if (!isNaN(s2.getFullYear()) && s2.getFullYear() >= 2000 && s2.getFullYear() <= 2100) {
        endYear = s2.getFullYear();
      }
    } else if (appData.settings.sem2Start) {
      const s2 = new Date(appData.settings.sem2Start);
      if (!isNaN(s2.getFullYear()) && s2.getFullYear() >= 2000 && s2.getFullYear() <= 2100) {
        endYear = s2.getFullYear();
      }
    }

    if (endYear <= startYear) {
      endYear = startYear + 1;
    }

    return { startYear, endYear };
  }, [appData.settings.academicYear, appData.settings.sem1Start, appData.settings.sem2Start, appData.settings.sem2End]);

  // Calendar date initialized to the start year of configured TP
  const [currentCalendarDate, setCurrentCalendarDate] = useState<Date>(() => {
    const s1 = new Date(appData.settings.sem1Start);
    if (!isNaN(s1.getFullYear()) && s1.getFullYear() >= 2000) {
      return new Date(s1.getFullYear(), s1.getMonth() || 6, 1);
    }
    const match = appData.settings.academicYear.match(/(\d{4})/);
    const yr = match ? parseInt(match[1], 10) : 2024;
    return new Date(yr, 6, 1);
  });

  // Ensure currentCalendarDate stays aligned when academic year settings change
  useEffect(() => {
    const curY = currentCalendarDate.getFullYear();
    const curM = currentCalendarDate.getMonth();
    const inTP =
      (curY === tpYears.startYear && curM >= 6) ||
      (curY === tpYears.endYear && curM <= 5);

    if (!inTP) {
      if (activeSemester === 1) {
        setCurrentCalendarDate(new Date(tpYears.startYear, 6, 1));
      } else {
        setCurrentCalendarDate(new Date(tpYears.endYear, 0, 1));
      }
    }
  }, [tpYears.startYear, tpYears.endYear, activeSemester]);

  // 12 Months of the configured Academic Year
  const tpMonths = useMemo(() => [
    { year: tpYears.startYear, month: 6, label: `Juli ${tpYears.startYear}` },
    { year: tpYears.startYear, month: 7, label: `Agustus ${tpYears.startYear}` },
    { year: tpYears.startYear, month: 8, label: `September ${tpYears.startYear}` },
    { year: tpYears.startYear, month: 9, label: `Oktober ${tpYears.startYear}` },
    { year: tpYears.startYear, month: 10, label: `November ${tpYears.startYear}` },
    { year: tpYears.startYear, month: 11, label: `Desember ${tpYears.startYear}` },
    { year: tpYears.endYear, month: 0, label: `Januari ${tpYears.endYear}` },
    { year: tpYears.endYear, month: 1, label: `Februari ${tpYears.endYear}` },
    { year: tpYears.endYear, month: 2, label: `Maret ${tpYears.endYear}` },
    { year: tpYears.endYear, month: 3, label: `April ${tpYears.endYear}` },
    { year: tpYears.endYear, month: 4, label: `Mei ${tpYears.endYear}` },
    { year: tpYears.endYear, month: 5, label: `Juni ${tpYears.endYear}` },
  ], [tpYears.startYear, tpYears.endYear]);

  const [calendarViewMode, setCalendarViewMode] = useState<'month' | 'semester'>('month');
  const [mobileMenuOpen, setMobileMenuOpen] = useState<boolean>(false);

  // Rekap Filters
  const [rekapFilterCat, setRekapFilterCat] = useState<string>('all');
  const [rekapFilterSem, setRekapFilterSem] = useState<string>('all');
  const [rekapFilterMonth, setRekapFilterMonth] = useState<string>('all');
  const [rekapSearchInput, setRekapSearchInput] = useState<string>('');

  // Modals State
  const [agendaModalOpen, setAgendaModalOpen] = useState<boolean>(false);
  const [editingAgendaId, setEditingAgendaId] = useState<string | null>(null);
  const [agendaForm, setAgendaForm] = useState<{
    id: string;
    category: CategoryKey;
    title: string;
    startDate: string;
    endDate: string;
    time: string;
    location: string;
    pic: string;
    description: string;
    isNonEffective: boolean;
  }>({
    id: '',
    category: 'holiday',
    title: '',
    startDate: getLocalDateString(),
    endDate: getLocalDateString(),
    time: '',
    location: '',
    pic: '',
    description: '',
    isNonEffective: false
  });

  const [dateDetailModalOpen, setDateDetailModalOpen] = useState<boolean>(false);
  const [selectedDateStr, setSelectedDateStr] = useState<string>('');

  const [printModalOpen, setPrintModalOpen] = useState<boolean>(false);
  const [isGeneratingPDF, setIsGeneratingPDF] = useState<boolean>(false);
  const [printType, setPrintType] = useState<'cal-sem1' | 'cal-sem2' | 'cal-annual' | 'rekap' | 'heb'>('cal-sem1');
  const [printPaperSize, setPrintPaperSize] = useState<'a4' | 'f4'>('a4');
  const [printOrientation, setPrintOrientation] = useState<'landscape' | 'portrait'>('landscape');

  // Confirmation Modal
  const [confirmModal, setConfirmModal] = useState<{
    open: boolean;
    title: string;
    message: string;
    onConfirm: () => void;
  }>({
    open: false,
    title: '',
    message: '',
    onConfirm: () => {}
  });

  // Toast Notifications
  const [toasts, setToasts] = useState<Array<{ id: number; msg: string; type: 'success' | 'error' | 'info' }>>([]);
  const toastCountRef = useRef(0);

  const showToast = (msg: string, type: 'success' | 'error' | 'info' = 'info') => {
    const id = ++toastCountRef.current;
    setToasts(prev => [...prev, { id, msg, type }]);
    setTimeout(() => {
      setToasts(prev => prev.filter(t => t.id !== id));
    }, 3500);
  };

  // Sync to localStorage
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(appData));
    } catch (e) {
      console.error('Failed to save to localStorage:', e);
    }
  }, [appData]);

  // Tab Titles Map
  const tabTitles: Record<string, [string, string]> = {
    'dashboard': ['Dashboard', 'Ringkasan Kalender Pendidikan & Agenda Sekolah'],
    'kalender': ['Kalender Master', 'Tampilan Kalender Pendidikan Interaktif & Matriks Semester'],
    'holidays': ['Hari Libur Nasional', 'Daftar Libur Resmi & Cuti Bersama Nasional'],
    'school-agenda': ['Agenda Sekolah', 'Pengelolaan Kegiatan Akademik, Ujian & Libur Semester'],
    'teacher-agenda': ['Agenda Guru', 'Jadwal Rapat, Kombel, Pelatihan & Pengembangan Guru'],
    'events': ['Peristiwa Penting', 'Hari Besar Nasional & Momentum Penting Sekolah'],
    'rekap': ['Rekap Agenda', 'Laporan & Rekapitulasi Seluruh Kegiatan Terfilter'],
    'heb': ['Hari Efektif Belajar (HEB)', 'Kalkulasi Otomatis Hari Belajar Efektif Per Bulan & Semester'],
    'settings': ['Pengaturan Aplikasi', 'Konfigurasi Profile Sekolah & Manajemen Data']
  };

  // Category counts
  const categoryCounts = useMemo(() => {
    const counts = { holiday: 0, school: 0, teacher: 0, important: 0 };
    appData.events.forEach(ev => {
      if (counts[ev.category] !== undefined) counts[ev.category]++;
    });
    return counts;
  }, [appData.events]);

  // HEB calculations
  const calculateMonthHEB = (year: number, month: number) => {
    const { sem1Start, sem1End, sem2Start, sem2End } = appData.settings;
    const totalDaysInMonth = new Date(year, month + 1, 0).getDate();
    let workDaysCount = 0;
    let holidayCount = 0;
    let nonEffectiveSchoolDays = 0;

    for (let day = 1; day <= totalDaysInMonth; day++) {
      const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
      const inSem1 = isDateInRange(dateStr, sem1Start, sem1End);
      const inSem2 = isDateInRange(dateStr, sem2Start, sem2End);

      if (!inSem1 && !inSem2) continue;

      const dayOfWeek = new Date(year, month, day).getDay();
      const isWorkDay = dayOfWeek >= 1 && dayOfWeek <= 5;

      if (isWorkDay) {
        workDaysCount++;

        const isHoliday = appData.events.some(ev => ev.category === 'holiday' && isDateInRange(dateStr, ev.startDate, ev.endDate));
        if (isHoliday) {
          holidayCount++;
          continue;
        }

        const isNonEff = appData.events.some(ev => ev.category === 'school' && ev.isNonEffective && isDateInRange(dateStr, ev.startDate, ev.endDate));
        if (isNonEff) {
          nonEffectiveSchoolDays++;
        }
      }
    }

    const computedHEB = Math.max(0, workDaysCount - holidayCount - nonEffectiveSchoolDays);
    const key = `${year}-${String(month + 1).padStart(2, '0')}`;
    const finalHEB = appData.hebOverrides[key] !== undefined ? appData.hebOverrides[key] : computedHEB;

    return {
      monthKey: key,
      year,
      month,
      workDaysCount,
      holidayCount,
      nonEffectiveSchoolDays,
      computedHEB,
      finalHEB,
      isOverridden: appData.hebOverrides[key] !== undefined
    };
  };

  const calculateSemesterHEB = (semesterNum: 1 | 2) => {
    const months = semesterNum === 1
      ? [6, 7, 8, 9, 10, 11].map(m => ({ year: tpYears.startYear, month: m }))
      : [0, 1, 2, 3, 4, 5].map(m => ({ year: tpYears.endYear, month: m }));

    let totalHEB = 0;
    const monthlyBreakdown = months.map(m => {
      const res = calculateMonthHEB(m.year, m.month);
      totalHEB += res.finalHEB;
      return res;
    });

    return { semesterNum, totalHEB, monthlyBreakdown };
  };

  const sem1HEB = useMemo(() => calculateSemesterHEB(1), [appData, tpYears]);
  const sem2HEB = useMemo(() => calculateSemesterHEB(2), [appData, tpYears]);
  const activeSemesterHEB = activeSemester === 1 ? sem1HEB : sem2HEB;

  // Navigate calendar month
  const navigateMonth = (delta: number) => {
    setCurrentCalendarDate(prev => new Date(prev.getFullYear(), prev.getMonth() + delta, 1));
  };

  const goToCurrentMonth = () => {
    const now = new Date();
    setCurrentCalendarDate(new Date(now.getFullYear(), now.getMonth(), 1));
  };

  // Open Agenda Modal
  const openAddAgendaModal = (cat: CategoryKey = 'holiday', defaultDate?: string) => {
    const d = defaultDate || getLocalDateString();
    setEditingAgendaId(null);
    setAgendaForm({
      id: '',
      category: cat,
      title: '',
      startDate: d,
      endDate: d,
      time: '',
      location: '',
      pic: '',
      description: '',
      isNonEffective: false
    });
    setAgendaModalOpen(true);
  };

  const openEditAgendaModal = (ev: KaldikEvent) => {
    setEditingAgendaId(ev.id);
    setAgendaForm({
      id: ev.id,
      category: ev.category,
      title: ev.title,
      startDate: ev.startDate,
      endDate: ev.endDate,
      time: ev.time || '',
      location: ev.location || '',
      pic: ev.pic || '',
      description: ev.description || '',
      isNonEffective: !!ev.isNonEffective
    });
    setAgendaModalOpen(true);
  };

  const handleSaveAgenda = (e: React.FormEvent) => {
    e.preventDefault();
    if (!agendaForm.title.trim()) {
      showToast('Nama agenda tidak boleh kosong', 'error');
      return;
    }
    if (agendaForm.startDate > agendaForm.endDate) {
      showToast('Tanggal mulai tidak boleh melebihi tanggal selesai', 'error');
      return;
    }

    const newAgenda: KaldikEvent = {
      id: editingAgendaId || 'ev-' + Date.now(),
      category: agendaForm.category,
      title: agendaForm.title.trim(),
      startDate: agendaForm.startDate,
      endDate: agendaForm.endDate,
      time: agendaForm.time.trim(),
      location: agendaForm.location.trim(),
      pic: agendaForm.pic.trim(),
      description: agendaForm.description.trim(),
      isNonEffective: agendaForm.category === 'school' ? agendaForm.isNonEffective : false,
      isDemo: false
    };

    setAppData(prev => {
      let updatedEvents = [...prev.events];
      if (editingAgendaId) {
        updatedEvents = updatedEvents.map(item => item.id === editingAgendaId ? newAgenda : item);
      } else {
        updatedEvents.push(newAgenda);
      }
      return { ...prev, events: updatedEvents };
    });

    setAgendaModalOpen(false);
    showToast(editingAgendaId ? 'Agenda berhasil diperbarui!' : 'Agenda berhasil ditambahkan!', 'success');
  };

  const handleDeleteAgenda = (id: string) => {
    const ev = appData.events.find(e => e.id === id);
    if (!ev) return;
    setConfirmModal({
      open: true,
      title: 'Konfirmasi Hapus Agenda',
      message: `Apakah Anda yakin ingin menghapus agenda "${ev.title}"?`,
      onConfirm: () => {
        setAppData(prev => ({
          ...prev,
          events: prev.events.filter(e => e.id !== id)
        }));
        showToast('Agenda berhasil dihapus', 'info');
        setConfirmModal(c => ({ ...c, open: false }));
      }
    });
  };

  // Reset / Empty all data in a specific category (Agenda Sekolah, Agenda Guru, Peristiwa Penting)
  const handleResetCategoryData = (catKey: CategoryKey) => {
    const count = appData.events.filter(e => e.category === catKey).length;
    const categoryName = CATEGORIES[catKey]?.name || catKey;

    if (count === 0) {
      showToast(`Data ${categoryName} sudah kosong.`, 'info');
      return;
    }

    setConfirmModal({
      open: true,
      title: `Kosongkan Seluruh Data ${categoryName}`,
      message: `PERINGATAN: Apakah Anda yakin ingin mereset dan menghapus SEMUA kegiatan pada "${categoryName}" (${count} data kegiatan)? Tindakan ini akan mengosongkan kategori ini dan tidak dapat dibatalkan.`,
      onConfirm: () => {
        setAppData(prev => ({
          ...prev,
          events: prev.events.filter(e => e.category !== catKey)
        }));
        showToast(`Seluruh data ${categoryName} (${count} kegiatan) berhasil dikosongkan.`, 'success');
        setConfirmModal(c => ({ ...c, open: false }));
      }
    });
  };

  // Image upload handler for logos
  const handleLogoUpload = (e: React.ChangeEvent<HTMLInputElement>, side: 'left' | 'right') => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (file.size > 2 * 1024 * 1024) {
      showToast('Ukuran gambar maksimal 2MB', 'error');
      return;
    }

    const reader = new FileReader();
    reader.onload = (event) => {
      const result = event.target?.result as string;
      setAppData(prev => ({
        ...prev,
        settings: {
          ...prev.settings,
          [side === 'left' ? 'logoLeft' : 'logoRight']: result
        }
      }));
      showToast(`Logo ${side === 'left' ? 'Kabupaten (Kiri)' : 'Sekolah (Kanan)'} berhasil diperbarui`, 'success');
    };
    reader.readAsDataURL(file);
  };

  const removeLogo = (side: 'left' | 'right') => {
    setAppData(prev => ({
      ...prev,
      settings: {
        ...prev.settings,
        [side === 'left' ? 'logoLeft' : 'logoRight']: ''
      }
    }));
    showToast(`Logo ${side === 'left' ? 'Kiri' : 'Kanan'} telah dihapus`, 'info');
  };

  // HEB override
  const updateHEBOverride = (monthKey: string, val: string) => {
    const num = parseInt(val, 10);
    if (!isNaN(num) && num >= 0) {
      setAppData(prev => ({
        ...prev,
        hebOverrides: {
          ...prev.hebOverrides,
          [monthKey]: num
        }
      }));
      showToast(`Penyesuaian HEB ${monthKey} disimpan`, 'success');
    }
  };

  // Synchronize national holidays according to active academic year
  const syncNationalHolidays = (showToastNotice = true) => {
    const newHolidays = getHolidaysForAcademicYear(
      appData.settings.academicYear,
      appData.settings.sem1Start,
      appData.settings.sem2End
    );

    setAppData(prev => {
      const nonHolidayEvents = prev.events.filter(e => e.category !== 'holiday');
      return {
        ...prev,
        events: [...nonHolidayEvents, ...newHolidays]
      };
    });

    if (showToastNotice) {
      showToast(`Libur Nasional berhasil disesuaikan untuk TP ${appData.settings.academicYear} (${newHolidays.length} hari libur resmi)`, 'success');
    }
  };

  // Download Category Excel Template
  const downloadCategoryExcelTemplate = (catKey: CategoryKey) => {
    let title = '';
    let headers: string[] = [];
    let sampleRows: (string | number)[][] = [];
    let colWidths: { wch: number }[] = [];

    if (catKey === 'school') {
      title = 'TEMPLATE IMPORT AGENDA & KEGIATAN SEKOLAH';
      headers = [
        'Nama Kegiatan',
        'Tanggal Mulai',
        'Tanggal Selesai',
        'Waktu',
        'Tempat',
        'PIC / Penanggung Jawab',
        'Status Hari Efektif',
        'Keterangan'
      ];
      sampleRows = [
        [
          'Masa Pengenalan Lingkungan Sekolah (MPLS)',
          '2024-07-15',
          '2024-07-17',
          '07:30 - 11:30',
          'Lapangan Sekolah',
          'Wakasek Kesiswaan',
          'Tidak Efektif',
          'Kegiatan orientasi lingkungan siswa baru'
        ],
        [
          'Asesmen Sumatif Akhir Semester 1 (SAS)',
          '2024-12-02',
          '2024-12-07',
          '07:30 - 11:00',
          'Ruang Kelas 1-6',
          'Panitia Asesmen',
          'Efektif',
          'Penilaian hasil belajar semester ganjil'
        ],
        [
          'Pembagian Rapor Semester 1',
          '2024-12-19',
          '2024-12-19',
          '08:00 - 11:30',
          'Kelas Masing-masing',
          'Wali Kelas',
          'Efektif',
          'Penyerahan buku laporan hasil belajar'
        ]
      ];
      colWidths = [{ wch: 35 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 22 }, { wch: 25 }, { wch: 20 }, { wch: 40 }];
    } else if (catKey === 'teacher') {
      title = 'TEMPLATE IMPORT AGENDA GURU & TENAGA KEPENDIDIKAN';
      headers = [
        'Nama Kegiatan',
        'Tanggal Mulai',
        'Tanggal Selesai',
        'Waktu',
        'Tempat',
        'Peserta / PIC',
        'Keterangan'
      ];
      sampleRows = [
        [
          'Rapat Pleno Kerja Awal Tahun Pelajaran',
          '2024-07-12',
          '2024-07-12',
          '08:00 - 14:00',
          'Ruang Guru',
          'Kepala Sekolah & Seluruh Guru',
          'Penyusunan KSP dan pembagian beban mengajar'
        ],
        [
          'Workshop Komunitas Belajar (Kombel) PMM',
          '2024-08-14',
          '2024-08-14',
          '13:00 - 15:30',
          'Ruang Multimedia',
          'Ketua Kombel & Guru Kelas',
          'Pelatihan pembelajaran berdiferensiasi'
        ],
        [
          'KKG Mini: Bedah Kisi-kisi Asesmen HOTS',
          '2024-10-10',
          '2024-10-10',
          '09:00 - 12:00',
          'Aula Sekolah',
          'Guru Kelas 1 - 6',
          'Penyusunan instrumen soal evaluasi'
        ]
      ];
      colWidths = [{ wch: 35 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 22 }, { wch: 30 }, { wch: 40 }];
    } else if (catKey === 'important') {
      title = 'TEMPLATE IMPORT PERISTIWA PENTING & PERINGATAN';
      headers = [
        'Nama Peristiwa / Peringatan',
        'Tanggal Mulai',
        'Tanggal Selesai',
        'Keterangan'
      ];
      sampleRows = [
        [
          'Hari Anak Nasional',
          '2024-07-23',
          '2024-07-23',
          'Peringatan kreasi bakat siswa di sekolah'
        ],
        [
          'Hari Pramuka ke-63',
          '2024-08-14',
          '2024-08-14',
          'Upacara bendera seragam pramuka lengkap'
        ],
        [
          'Hari Batik Nasional',
          '2024-10-02',
          '2024-10-02',
          'Seluruh warga sekolah mengenakan batik nusantara'
        ],
        [
          'Hari Guru Nasional (HGN) & HUT PGRI',
          '2024-11-25',
          '2024-11-25',
          'Upacara bendera dan apresiasi dedikasi guru'
        ]
      ];
      colWidths = [{ wch: 35 }, { wch: 15 }, { wch: 15 }, { wch: 45 }];
    } else {
      title = 'TEMPLATE IMPORT HARI LIBUR NASIONAL';
      headers = ['Nama Hari Libur', 'Tanggal Mulai', 'Tanggal Selesai', 'Keterangan'];
      sampleRows = [['Tahun Baru Masehi', '2025-01-01', '2025-01-01', 'Libur Resmi Tahun Baru']];
      colWidths = [{ wch: 35 }, { wch: 15 }, { wch: 15 }, { wch: 35 }];
    }

    const aoaData: any[][] = [
      [title],
      [`Tahun Pelajaran: ${appData.settings.academicYear}`],
      ['PETUNJUK: Masukkan data kegiatan mulai dari baris ke-5. Format tanggal: YYYY-MM-DD (contoh: 2024-08-17) atau DD/MM/YYYY.'],
      [],
      headers,
      ...sampleRows
    ];

    const ws = XLSX.utils.aoa_to_sheet(aoaData);
    ws['!cols'] = colWidths;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");

    const fileName = `Template_${CATEGORIES[catKey].name.replace(/\s+/g, '_')}.xlsx`;
    XLSX.writeFile(wb, fileName);
    showToast(`Template Excel untuk ${CATEGORIES[catKey].name} berhasil diunduh!`, 'success');
  };

  // Upload and Import Category Excel File
  const handleCategoryExcelUpload = async (
    e: React.ChangeEvent<HTMLInputElement>,
    catKey: CategoryKey
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      showToast('Membaca dan memproses file Excel...', 'info');
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, { type: 'array', cellDates: true });
      const sheetName = wb.SheetNames[0];
      const sheet = wb.Sheets[sheetName];

      const rawRows = XLSX.utils.sheet_to_json<Record<string, any>>(sheet, { defval: '', raw: false });

      if (rawRows.length === 0) {
        showToast('File Excel kosong atau tidak memiliki baris data yang valid', 'error');
        return;
      }

      const importedEvents: KaldikEvent[] = [];

      rawRows.forEach((row, index) => {
        // Find title
        const title = getFieldVal(row, ['kegiatan', 'agenda', 'peristiwa', 'peringatan', 'nama', 'judul', 'title', 'event']);
        if (!title || title.startsWith('PETUNJUK') || title.startsWith('TEMPLATE')) return;

        // Find start date
        const startRaw = getFieldVal(row, ['mulai', 'start', 'tanggal', 'tgl', 'date']);
        const startDate = parseExcelDate(startRaw) || getLocalDateString();

        // Find end date
        const endRaw = getFieldVal(row, ['selesai', 'end', 'akhir', 'sampai']);
        let endDate = parseExcelDate(endRaw);
        if (!endDate || endDate < startDate) {
          endDate = startDate;
        }

        // Optional fields
        const time = getFieldVal(row, ['waktu', 'jam', 'time', 'pukul']);
        const location = getFieldVal(row, ['tempat', 'lokasi', 'ruang', 'location']);
        const pic = getFieldVal(row, ['pic', 'penanggung', 'peserta', 'guru', 'panitia']);
        const description = getFieldVal(row, ['keterangan', 'catatan', 'deskripsi', 'desc']);
        
        const effectiveRaw = getFieldVal(row, ['efektif', 'status', 'heb']).toLowerCase();
        const isNonEffective = effectiveRaw.includes('tidak') || effectiveRaw.includes('non') || effectiveRaw.includes('libur');

        importedEvents.push({
          id: `ev-imp-${catKey}-${Date.now()}-${index + 1}`,
          category: catKey,
          title,
          startDate,
          endDate,
          time: time || undefined,
          location: location || undefined,
          pic: pic || undefined,
          description: description || undefined,
          isNonEffective: catKey === 'school' ? isNonEffective : false,
          isDemo: false
        });
      });

      if (importedEvents.length === 0) {
        showToast('Tidak ada kegiatan yang berhasil dibaca. Pastikan kolom sesuai template!', 'error');
        return;
      }

      setAppData(prev => ({
        ...prev,
        events: [...prev.events, ...importedEvents]
      }));

      showToast(`Berhasil mengimpor ${importedEvents.length} kegiatan ke ${CATEGORIES[catKey].name}!`, 'success');
    } catch (err) {
      console.error('Error importing Excel:', err);
      showToast('Gagal membaca file Excel. Pastikan file berformat .xlsx atau .xls yang benar.', 'error');
    } finally {
      e.target.value = '';
    }
  };

  // Excel Export
  const exportRekapExcel = () => {
    const rows: (string | number)[][] = [
      ['REKAPITULASI KALENDER PENDIDIKAN & AGENDA SEKOLAH'],
      [appData.settings.schoolName],
      [`Tahun Pelajaran ${appData.settings.academicYear}`],
      [''],
      ['No', 'Kategori', 'Tanggal Mulai', 'Tanggal Selesai', 'Waktu', 'Nama Kegiatan / Agenda', 'Tempat / Lokasi', 'PIC / Penanggung Jawab', 'Keterangan']
    ];

    const sorted = [...appData.events].sort((a, b) => a.startDate.localeCompare(b.startDate));
    sorted.forEach((ev, idx) => {
      const catName = CATEGORIES[ev.category]?.name || ev.category;
      rows.push([
        idx + 1,
        catName,
        formatDateIndo(ev.startDate),
        formatDateIndo(ev.endDate),
        ev.time || '-',
        ev.title,
        ev.location || '-',
        ev.pic || '-',
        ev.description || '-'
      ]);
    });

    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rekap Agenda");
    XLSX.writeFile(wb, `Rekap_Agenda_${appData.settings.academicYear.replace('/', '-')}.xlsx`);
    showToast('Rekapitulasi Excel berhasil diunduh', 'success');
  };

  // Backup & Restore
  const backupJSON = () => {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(appData, null, 2));
    const a = document.createElement('a');
    a.href = dataStr;
    a.download = `KALDIK_BACKUP_${appData.settings.academicYear.replace('/', '-')}_${getLocalDateString()}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    showToast('File backup JSON berhasil diunduh', 'success');
  };

  const restoreJSON = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const parsed = JSON.parse(evt.target?.result as string);
        if (parsed.settings && parsed.events) {
          setAppData({
            settings: { ...DEFAULT_SETTINGS, ...parsed.settings },
            events: parsed.events,
            hebOverrides: parsed.hebOverrides || {}
          });
          showToast('Data berhasil dipulihkan dari backup!', 'success');
          setActiveTab('dashboard');
        } else {
          showToast('Format file JSON backup tidak valid', 'error');
        }
      } catch (err) {
        showToast('Gagal membaca file JSON', 'error');
      }
    };
    reader.readAsText(file);
  };

  const confirmResetDemo = () => {
    setConfirmModal({
      open: true,
      title: 'Reset Data ke Standar Demo',
      message: 'Apakah Anda yakin ingin mengembalikan seluruh data pengaturan dan agenda ke kondisi awal demo?',
      onConfirm: () => {
        setAppData({
          settings: { ...DEFAULT_SETTINGS },
          events: [...DEMO_EVENTS],
          hebOverrides: {}
        });
        showToast('Data berhasil di-reset ke standar demo', 'info');
        setConfirmModal(c => ({ ...c, open: false }));
      }
    });
  };

  // Fast and robust PDF Download handler using bundled jsPDF + html2canvas
  const downloadPDFDocument = async () => {
    const element = document.getElementById('print-area');
    if (!element) {
      showToast('Area dokumen tidak ditemukan', 'error');
      return;
    }

    setIsGeneratingPDF(true);
    showToast('Sedang memproses berkas PDF...', 'info');

    try {
      // Capture canvas with optimized settings (no hanging on external assets)
      const canvas = await html2canvas(element, {
        scale: 1.5,
        useCORS: true,
        allowTaint: true,
        logging: false,
        backgroundColor: '#ffffff',
        imageTimeout: 4000,
      });

      const imgData = canvas.toDataURL('image/jpeg', 0.95);
      const isLandscape = printOrientation === 'landscape';
      const isF4 = printPaperSize === 'f4';

      // Width and Height in mm
      // A4: 210 x 297 mm, F4: 215 x 330 mm
      const paperWidth = isF4 ? (isLandscape ? 330 : 215) : (isLandscape ? 297 : 210);
      const paperHeight = isF4 ? (isLandscape ? 215 : 330) : (isLandscape ? 210 : 297);

      const pdf = new jsPDF({
        orientation: isLandscape ? 'landscape' : 'portrait',
        unit: 'mm',
        format: isF4 ? [215, 330] : 'a4'
      });

      const margin = 8;
      const printableWidth = paperWidth - (margin * 2);
      const printableHeight = paperHeight - (margin * 2);
      const imgHeight = (canvas.height * printableWidth) / canvas.width;

      if (imgHeight <= printableHeight) {
        // Fits on single page
        pdf.addImage(imgData, 'JPEG', margin, margin, printableWidth, imgHeight);
      } else {
        // Multi-page slicing
        let heightLeft = imgHeight;
        let position = margin;
        let page = 1;

        while (heightLeft > 0) {
          if (page > 1) {
            pdf.addPage();
          }
          pdf.addImage(imgData, 'JPEG', margin, position, printableWidth, imgHeight);
          heightLeft -= printableHeight;
          position -= printableHeight;
          page++;
        }
      }

      pdf.save(`KALDIK_${appData.settings.academicYear.replace('/', '-')}_${printType}.pdf`);
      showToast('PDF berhasil diunduh!', 'success');
    } catch (err) {
      console.error('PDF Generation Error:', err);
      showToast('Gagal memproses via canvas, membuka dialog cetak browser...', 'error');
      window.print();
    } finally {
      setIsGeneratingPDF(false);
    }
  };

  // Filtered Rekap List
  const filteredRekapEvents = useMemo(() => {
    let list = [...appData.events];

    if (rekapFilterCat !== 'all') {
      list = list.filter(e => e.category === rekapFilterCat);
    }

    if (rekapFilterSem !== 'all') {
      const bounds = rekapFilterSem === '1'
        ? [appData.settings.sem1Start, appData.settings.sem1End]
        : [appData.settings.sem2Start, appData.settings.sem2End];
      list = list.filter(e => e.startDate >= bounds[0] && e.startDate <= bounds[1]);
    }

    if (rekapFilterMonth !== 'all') {
      const mTarget = parseInt(rekapFilterMonth, 10);
      list = list.filter(e => new Date(e.startDate).getMonth() === mTarget);
    }

    if (rekapSearchInput.trim()) {
      const kw = rekapSearchInput.toLowerCase().trim();
      list = list.filter(e =>
        e.title.toLowerCase().includes(kw) ||
        (e.description && e.description.toLowerCase().includes(kw)) ||
        (e.pic && e.pic.toLowerCase().includes(kw)) ||
        (e.location && e.location.toLowerCase().includes(kw))
      );
    }

    return list.sort((a, b) => a.startDate.localeCompare(b.startDate));
  }, [appData.events, appData.settings, rekapFilterCat, rekapFilterSem, rekapFilterMonth, rekapSearchInput]);

  // Upcoming Agendas
  const upcomingAgendas = useMemo(() => {
    const todayStr = getLocalDateString();
    return appData.events
      .filter(ev => ev.endDate >= todayStr)
      .sort((a, b) => a.startDate.localeCompare(b.startDate))
      .slice(0, 5);
  }, [appData.events]);

  // Today activities
  const todayStr = getLocalDateString();
  const todayEvents = useMemo(() => {
    return appData.events.filter(ev => isDateInRange(todayStr, ev.startDate, ev.endDate));
  }, [appData.events, todayStr]);

  // Calendar Grid Generator Helper
  const renderCalendarMonthGrid = (targetDate: Date) => {
    const year = targetDate.getFullYear();
    const month = targetDate.getMonth();
    const firstDay = new Date(year, month, 1).getDay();
    const startOffset = (firstDay + 6) % 7; // Monday = 0
    const totalDays = new Date(year, month + 1, 0).getDate();

    const cells: React.ReactNode[] = [];

    // Empty offset cells
    for (let i = 0; i < startOffset; i++) {
      cells.push(<div key={`empty-${i}`} className="day-cell border-r border-b border-slate-200 bg-slate-50/50 p-1"></div>);
    }

    // Days cells
    for (let day = 1; day <= totalDays; day++) {
      const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
      const dayOfWeek = new Date(year, month, day).getDay();
      const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
      const isToday = dateStr === todayStr;

      const dayEvents = appData.events.filter(ev => isDateInRange(dateStr, ev.startDate, ev.endDate));
      const hasHoliday = dayEvents.some(e => e.category === 'holiday');

      let cellBgClass = 'bg-white hover:bg-slate-50';
      if (hasHoliday || isWeekend) cellBgClass = 'bg-rose-50/40 hover:bg-rose-50/70';
      if (isToday) cellBgClass = 'bg-indigo-50/80 hover:bg-indigo-100/80 ring-2 ring-indigo-500 ring-inset';

      cells.push(
        <div
          key={dateStr}
          onClick={() => {
            setSelectedDateStr(dateStr);
            setDateDetailModalOpen(true);
          }}
          className={`day-cell border-r border-b border-slate-200 p-1.5 cursor-pointer transition flex flex-col justify-between ${cellBgClass}`}
        >
          <div className="flex items-center justify-between mb-1">
            <span className={`text-xs font-bold ${
              isToday
                ? 'bg-indigo-600 text-white w-5 h-5 rounded-full flex items-center justify-center'
                : (isWeekend || hasHoliday ? 'text-rose-600' : 'text-slate-800')
            }`}>
              {day}
            </span>
            {dayEvents.length > 0 && (
              <span className="text-[10px] font-mono font-semibold px-1 rounded bg-slate-200 text-slate-700">
                {dayEvents.length}
              </span>
            )}
          </div>

          <div className="space-y-1 overflow-hidden">
            {dayEvents.slice(0, 2).map(ev => {
              const cat = CATEGORIES[ev.category] || CATEGORIES.school;
              return (
                <div
                  key={ev.id}
                  className={`text-[10px] leading-tight font-medium px-1.5 py-0.5 rounded truncate border ${cat.badgeClass}`}
                  title={ev.title}
                >
                  {ev.title}
                </div>
              );
            })}
            {dayEvents.length > 2 && (
              <div className="text-[9px] text-slate-500 font-semibold px-1">
                +{dayEvents.length - 2} lagi...
              </div>
            )}
          </div>
        </div>
      );
    }

    // Trailing cells
    const totalCells = startOffset + totalDays;
    const trailingCells = (7 - (totalCells % 7)) % 7;
    for (let i = 0; i < trailingCells; i++) {
      cells.push(<div key={`trail-${i}`} className="day-cell border-r border-b border-slate-200 bg-slate-50/50 p-1"></div>);
    }

    return (
      <div>
        <div className="grid grid-cols-7 border-t border-l border-slate-200 text-center font-semibold text-xs bg-slate-100 text-slate-700">
          <div className="py-2 border-r border-b border-slate-200">Sen</div>
          <div className="py-2 border-r border-b border-slate-200">Sel</div>
          <div className="py-2 border-r border-b border-slate-200">Rab</div>
          <div className="py-2 border-r border-b border-slate-200">Kam</div>
          <div className="py-2 border-r border-b border-slate-200">Jum</div>
          <div className="py-2 border-r border-b border-slate-200 text-rose-600">Sab</div>
          <div className="py-2 border-r border-b border-slate-200 text-rose-600">Min</div>
        </div>
        <div className="grid grid-cols-7 border-l border-slate-200">
          {cells}
        </div>
      </div>
    );
  };

  // Mini Calendar for Semester Matrix
  const renderMiniCalendar = (year: number, month: number) => {
    const firstDay = new Date(year, month, 1).getDay();
    const startOffset = (firstDay + 6) % 7;
    const totalDays = new Date(year, month + 1, 0).getDate();

    const cells: React.ReactNode[] = [];
    for (let i = 0; i < startOffset; i++) {
      cells.push(<div key={`empty-mini-${i}`}></div>);
    }

    for (let day = 1; day <= totalDays; day++) {
      const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
      const dayOfWeek = new Date(year, month, day).getDay();
      const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
      const dayEvents = appData.events.filter(ev => isDateInRange(dateStr, ev.startDate, ev.endDate));
      const hasHoliday = dayEvents.some(e => e.category === 'holiday');
      const isToday = dateStr === todayStr;

      let bgClass = 'text-slate-700';
      if (hasHoliday || isWeekend) bgClass = 'text-rose-600 font-semibold bg-rose-50 rounded';
      if (dayEvents.length > 0 && !hasHoliday) bgClass = 'bg-indigo-100 text-indigo-900 font-bold rounded';
      if (isToday) bgClass = 'bg-indigo-600 text-white font-bold rounded-full';

      cells.push(
        <div
          key={dateStr}
          onClick={() => {
            setSelectedDateStr(dateStr);
            setDateDetailModalOpen(true);
          }}
          className={`py-1 cursor-pointer hover:opacity-80 transition text-center text-xs ${bgClass}`}
          title={dayEvents.map(e => e.title).join(', ')}
        >
          {day}
        </div>
      );
    }

    return (
      <div className="border border-slate-200 rounded-xl p-3 bg-white shadow-sm">
        <h4 className="font-bold text-sm text-slate-800 text-center mb-3 bg-slate-100 py-1 rounded">
          {getMonthNameIndo(month)} {year}
        </h4>
        <div className="grid grid-cols-7 text-center font-semibold text-[10px] text-slate-500 mb-1">
          <div>S</div><div>S</div><div>R</div><div>K</div><div>J</div><div className="text-rose-500">S</div><div className="text-rose-500">M</div>
        </div>
        <div className="grid grid-cols-7 text-center text-xs gap-1">
          {cells}
        </div>
      </div>
    );
  };

  // Helper for Category Tables
  const renderCategoryCRUDTable = (catKey: CategoryKey) => {
    const list = appData.events
      .filter(e => e.category === catKey)
      .sort((a, b) => a.startDate.localeCompare(b.startDate));

    return (
      <div className="bg-white rounded-2xl border border-slate-200 p-4 sm:p-6 shadow-sm">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-6">
          <div>
            <div className="flex items-center gap-2">
              <span className={`w-3 h-3 rounded-full ${CATEGORIES[catKey].bgClass}`}></span>
              <h3 className="text-lg font-bold text-slate-800">
                {catKey === 'holiday' && 'Hari Libur Nasional & Cuti Bersama'}
                {catKey === 'school' && 'Agenda & Kegiatan Sekolah'}
                {catKey === 'teacher' && 'Agenda Guru & Tenaga Kependidikan'}
                {catKey === 'important' && 'Peristiwa Penting & Peringatan'}
              </h3>
              {catKey === 'holiday' && (
                <span className="text-[11px] bg-rose-100 text-rose-800 font-bold px-2 py-0.5 rounded-full border border-rose-200">
                  TP {appData.settings.academicYear}
                </span>
              )}
            </div>
            <p className="text-xs text-slate-500 mt-1">
              {catKey === 'holiday' && 'Disesuaikan otomatis dengan libur resmi SKB 3 Menteri sesuai tahun pelajaran yang digunakan'}
              {catKey === 'school' && 'Kegiatan akademik, ujian, pembagian rapor, dan libur semester (tersedia import Excel)'}
              {catKey === 'teacher' && 'Rapat dinas, workshop, KKG/MGMP, dan pelatihan guru (tersedia import Excel)'}
              {catKey === 'important' && 'Hari Besar Nasional, Peringatan Pendidikan & momentum penting (tersedia import Excel)'}
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            {catKey === 'holiday' ? (
              <button
                onClick={() => syncNationalHolidays(true)}
                className="bg-rose-700 hover:bg-rose-800 text-white px-3.5 py-2 rounded-lg text-xs font-semibold flex items-center gap-2 shadow-sm transition cursor-pointer"
                title="Sesuaikan otomatis hari libur nasional sesuai Tahun Pelajaran aktif"
              >
                <i className="fa-solid fa-arrows-rotate"></i>
                <span>Sinkronkan Libur Otomatis ({appData.settings.academicYear})</span>
              </button>
            ) : (
              <>
                {/* Download Template Excel */}
                <button
                  onClick={() => downloadCategoryExcelTemplate(catKey)}
                  className="bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition cursor-pointer shadow-sm"
                  title={`Unduh template file Excel untuk ${CATEGORIES[catKey].name}`}
                >
                  <i className="fa-solid fa-file-arrow-down text-emerald-600"></i>
                  <span>Unduh Template Excel</span>
                </button>

                {/* Upload Excel */}
                <label
                  className="bg-emerald-600 hover:bg-emerald-700 text-white px-3.5 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition cursor-pointer shadow-sm"
                  title={`Unggah file Excel untuk menambah kegiatan ${CATEGORIES[catKey].name}`}
                >
                  <i className="fa-solid fa-file-arrow-up"></i>
                  <span>Upload Excel</span>
                  <input
                    type="file"
                    accept=".xlsx, .xls"
                    onChange={e => handleCategoryExcelUpload(e, catKey)}
                    className="hidden"
                  />
                </label>

                {/* Reset Data (Kosongkan Data) */}
                <button
                  type="button"
                  onClick={() => handleResetCategoryData(catKey)}
                  className="bg-white hover:bg-rose-50 text-rose-600 hover:text-rose-700 border border-rose-200 hover:border-rose-300 px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition cursor-pointer shadow-sm"
                  title={`Kosongkan / hapus seluruh kegiatan ${CATEGORIES[catKey].name}`}
                >
                  <i className="fa-solid fa-trash-can"></i>
                  <span>Reset Data</span>
                </button>
              </>
            )}

            {/* Manual Add Button */}
            <button
              onClick={() => openAddAgendaModal(catKey)}
              className={`text-white px-4 py-2 rounded-lg text-xs font-semibold flex items-center gap-2 shadow-sm transition cursor-pointer ${
                catKey === 'holiday' ? 'bg-rose-600 hover:bg-rose-700' :
                catKey === 'school' ? 'bg-sky-600 hover:bg-sky-700' :
                catKey === 'teacher' ? 'bg-purple-600 hover:bg-purple-700' :
                'bg-amber-600 hover:bg-amber-700'
              }`}
            >
              <i className="fa-solid fa-plus"></i> Tambah {CATEGORIES[catKey].name}
            </button>
          </div>
        </div>

        <div className="overflow-x-auto border border-slate-200 rounded-xl">
          <table className="w-full text-left border-collapse text-sm">
            <thead className="bg-slate-50 text-slate-700 text-xs uppercase font-semibold border-b border-slate-200">
              <tr>
                <th className="py-3 px-4">Tanggal / Waktu</th>
                <th className="py-3 px-4">Nama Kegiatan</th>
                {catKey === 'holiday' && <th className="py-3 px-4">Keterangan</th>}
                {catKey === 'holiday' && <th className="py-3 px-4 text-center">Durasi</th>}
                {catKey === 'school' && <th className="py-3 px-4">Tempat / PIC</th>}
                {catKey === 'school' && <th className="py-3 px-4">Status Efektif</th>}
                {catKey === 'teacher' && <th className="py-3 px-4">Peserta / PIC</th>}
                {catKey === 'teacher' && <th className="py-3 px-4">Tempat</th>}
                {catKey === 'important' && <th className="py-3 px-4">Keterangan</th>}
                <th className="py-3 px-4 text-center">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200 text-slate-700">
              {list.length === 0 ? (
                <tr>
                  <td colSpan={6} className="py-8 text-center text-xs text-slate-400 italic">
                    Belum ada data kegiatan pada kategori ini.
                  </td>
                </tr>
              ) : (
                list.map(ev => {
                  const dateRange = ev.startDate === ev.endDate ? formatDateIndo(ev.startDate) : `${formatDateIndo(ev.startDate)} s.d. ${formatDateIndo(ev.endDate)}`;
                  const durationDays = getDatesInRange(ev.startDate, ev.endDate).length;

                  return (
                    <tr key={ev.id} className="hover:bg-slate-50 transition text-xs">
                      <td className="py-3 px-4">
                        <div className={`font-semibold ${CATEGORIES[catKey].textClass}`}>{dateRange}</div>
                        {ev.time && <div className="text-[10px] text-slate-400">{ev.time}</div>}
                      </td>
                      <td className="py-3 px-4 font-bold text-slate-800">
                        {ev.title}
                        {ev.isDemo && (
                          <span className="text-[9px] bg-amber-100 text-amber-800 px-1.5 py-0.5 rounded font-mono ml-1.5">
                            DEMO
                          </span>
                        )}
                      </td>

                      {catKey === 'holiday' && (
                        <>
                          <td className="py-3 px-4 text-slate-500">{ev.description || '-'}</td>
                          <td className="py-3 px-4 text-center font-mono font-medium">{durationDays} Hari</td>
                        </>
                      )}

                      {catKey === 'school' && (
                        <>
                          <td className="py-3 px-4">
                            <div className="text-slate-700">{ev.location || '-'}</div>
                            <div className="text-[10px] text-slate-400">PIC: {ev.pic || '-'}</div>
                          </td>
                          <td className="py-3 px-4">
                            {ev.isNonEffective ? (
                              <span className="px-2 py-0.5 rounded-full bg-rose-100 text-rose-700 font-semibold text-[10px]">
                                Tidak Efektif
                              </span>
                            ) : (
                              <span className="px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-700 font-semibold text-[10px]">
                                Hari Efektif
                              </span>
                            )}
                          </td>
                        </>
                      )}

                      {catKey === 'teacher' && (
                        <>
                          <td className="py-3 px-4 text-slate-600">{ev.pic || 'Seluruh Guru'}</td>
                          <td className="py-3 px-4 text-slate-600">{ev.location || '-'}</td>
                        </>
                      )}

                      {catKey === 'important' && (
                        <td className="py-3 px-4 text-slate-500">{ev.description || '-'}</td>
                      )}

                      <td className="py-3 px-4 text-center">
                        <div className="flex items-center justify-center gap-2">
                          <button
                            onClick={() => openEditAgendaModal(ev)}
                            className="text-slate-600 hover:text-indigo-600 p-1 transition"
                            title="Edit"
                          >
                            <i className="fa-solid fa-pen-to-square"></i>
                          </button>
                          <button
                            onClick={() => handleDeleteAgenda(ev.id)}
                            className="text-slate-600 hover:text-rose-600 p-1 transition"
                            title="Hapus"
                          >
                            <i className="fa-solid fa-trash"></i>
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>
    );
  };

  // Selected Date events in modal
  const selectedDateEvents = useMemo(() => {
    if (!selectedDateStr) return [];
    return appData.events.filter(ev => isDateInRange(selectedDateStr, ev.startDate, ev.endDate));
  }, [appData.events, selectedDateStr]);

  return (
    <div className="min-h-screen flex flex-col md:flex-row bg-slate-100 text-slate-800 font-sans antialiased">
      {/* ==================== SIDEBAR NAV ==================== */}
      <aside className="w-full md:w-64 bg-slate-900 text-white flex-shrink-0 flex flex-col justify-between transition-all duration-300 z-30">
        <div>
          {/* App Header */}
          <div className="p-4 border-b border-slate-800 flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-lg bg-indigo-600 flex items-center justify-center font-bold text-xl text-white shadow-lg shadow-indigo-500/30">
                <i className="fa-solid fa-calendar-days text-lg"></i>
              </div>
              <div>
                <h1 className="font-bold text-sm tracking-wide leading-tight text-white">KALDIK SEKOLAH</h1>
                <p className="text-xs text-slate-400">Sekolah Dasar (SD)</p>
              </div>
            </div>
            <button
              onClick={() => setMobileMenuOpen(prev => !prev)}
              className="md:hidden text-slate-400 hover:text-white p-1"
            >
              <i className="fa-solid fa-bars text-xl"></i>
            </button>
          </div>

          {/* Navigation Links */}
          <nav className={`p-3 space-y-1 ${mobileMenuOpen ? 'block' : 'hidden'} md:block`}>
            <button
              onClick={() => { setActiveTab('dashboard'); setMobileMenuOpen(false); }}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${
                activeTab === 'dashboard' ? 'bg-indigo-600 text-white' : 'text-slate-300 hover:bg-slate-800 hover:text-white'
              }`}
            >
              <i className="fa-solid fa-chart-pie w-5 text-center"></i>
              <span>Dashboard</span>
            </button>

            <button
              onClick={() => { setActiveTab('kalender'); setMobileMenuOpen(false); }}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${
                activeTab === 'kalender' ? 'bg-indigo-600 text-white' : 'text-slate-300 hover:bg-slate-800 hover:text-white'
              }`}
            >
              <i className="fa-solid fa-calendar-days w-5 text-center"></i>
              <span>Kalender Master</span>
            </button>

            <div className="pt-3 pb-1 px-3 text-[11px] font-semibold uppercase tracking-wider text-slate-400">
              Kategori Agenda
            </div>

            <button
              onClick={() => { setActiveTab('holidays'); setMobileMenuOpen(false); }}
              className={`w-full flex items-center justify-between px-3 py-2 rounded-lg text-sm font-medium transition-colors ${
                activeTab === 'holidays' ? 'bg-indigo-600 text-white' : 'text-slate-300 hover:bg-slate-800 hover:text-white'
              }`}
            >
              <div className="flex items-center gap-3">
                <span className="w-2.5 h-2.5 rounded-full bg-rose-500"></span>
                <span>Libur Nasional</span>
              </div>
              <span className="bg-rose-950 text-rose-300 text-xs px-2 py-0.5 rounded-full font-mono">
                {categoryCounts.holiday}
              </span>
            </button>

            <button
              onClick={() => { setActiveTab('school-agenda'); setMobileMenuOpen(false); }}
              className={`w-full flex items-center justify-between px-3 py-2 rounded-lg text-sm font-medium transition-colors ${
                activeTab === 'school-agenda' ? 'bg-indigo-600 text-white' : 'text-slate-300 hover:bg-slate-800 hover:text-white'
              }`}
            >
              <div className="flex items-center gap-3">
                <span className="w-2.5 h-2.5 rounded-full bg-sky-500"></span>
                <span>Agenda Sekolah</span>
              </div>
              <span className="bg-sky-950 text-sky-300 text-xs px-2 py-0.5 rounded-full font-mono">
                {categoryCounts.school}
              </span>
            </button>

            <button
              onClick={() => { setActiveTab('teacher-agenda'); setMobileMenuOpen(false); }}
              className={`w-full flex items-center justify-between px-3 py-2 rounded-lg text-sm font-medium transition-colors ${
                activeTab === 'teacher-agenda' ? 'bg-indigo-600 text-white' : 'text-slate-300 hover:bg-slate-800 hover:text-white'
              }`}
            >
              <div className="flex items-center gap-3">
                <span className="w-2.5 h-2.5 rounded-full bg-purple-500"></span>
                <span>Agenda Guru</span>
              </div>
              <span className="bg-purple-950 text-purple-300 text-xs px-2 py-0.5 rounded-full font-mono">
                {categoryCounts.teacher}
              </span>
            </button>

            <button
              onClick={() => { setActiveTab('events'); setMobileMenuOpen(false); }}
              className={`w-full flex items-center justify-between px-3 py-2 rounded-lg text-sm font-medium transition-colors ${
                activeTab === 'events' ? 'bg-indigo-600 text-white' : 'text-slate-300 hover:bg-slate-800 hover:text-white'
              }`}
            >
              <div className="flex items-center gap-3">
                <span className="w-2.5 h-2.5 rounded-full bg-amber-500"></span>
                <span>Peristiwa Penting</span>
              </div>
              <span className="bg-amber-950 text-amber-300 text-xs px-2 py-0.5 rounded-full font-mono">
                {categoryCounts.important}
              </span>
            </button>

            <div className="pt-3 pb-1 px-3 text-[11px] font-semibold uppercase tracking-wider text-slate-400">
              Laporan & Analisis
            </div>

            <button
              onClick={() => { setActiveTab('rekap'); setMobileMenuOpen(false); }}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${
                activeTab === 'rekap' ? 'bg-indigo-600 text-white' : 'text-slate-300 hover:bg-slate-800 hover:text-white'
              }`}
            >
              <i className="fa-solid fa-file-contract w-5 text-center"></i>
              <span>Rekap Agenda</span>
            </button>

            <button
              onClick={() => { setActiveTab('heb'); setMobileMenuOpen(false); }}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${
                activeTab === 'heb' ? 'bg-indigo-600 text-white' : 'text-slate-300 hover:bg-slate-800 hover:text-white'
              }`}
            >
              <i className="fa-solid fa-calculator w-5 text-center"></i>
              <span>Hari Efektif (HEB)</span>
            </button>

            <button
              onClick={() => { setActiveTab('settings'); setMobileMenuOpen(false); }}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${
                activeTab === 'settings' ? 'bg-indigo-600 text-white' : 'text-slate-300 hover:bg-slate-800 hover:text-white'
              }`}
            >
              <i className="fa-solid fa-sliders w-5 text-center"></i>
              <span>Pengaturan</span>
            </button>
          </nav>
        </div>

        {/* Active Academic Year Box */}
        <div className="p-4 border-t border-slate-800 hidden md:block">
          <div className="bg-slate-800/80 rounded-lg p-3 text-xs border border-slate-700/50">
            <div className="text-slate-400 flex items-center justify-between mb-1">
              <span>Tahun Pelajaran</span>
              <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
            </div>
            <div className="font-bold text-white text-sm">{appData.settings.academicYear}</div>
            <div className="text-slate-400 text-[11px] truncate mt-1">{appData.settings.schoolName}</div>
          </div>
        </div>
      </aside>

      {/* ==================== MAIN CONTENT ==================== */}
      <main className="flex-1 flex flex-col min-w-0 overflow-y-auto">
        {/* Topbar */}
        <header className="bg-white border-b border-slate-200 sticky top-0 z-20 shadow-sm px-4 py-3 flex items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <button
              onClick={() => setMobileMenuOpen(prev => !prev)}
              className="md:hidden text-slate-600 hover:text-slate-900 p-1"
            >
              <i className="fa-solid fa-bars text-xl"></i>
            </button>
            <div>
              <h2 className="text-lg font-bold text-slate-800 leading-tight">
                {tabTitles[activeTab]?.[0] || 'Dashboard'}
              </h2>
              <p className="text-xs text-slate-500 hidden sm:block">
                {tabTitles[activeTab]?.[1] || 'Ringkasan Kalender Pendidikan & Agenda Sekolah'}
              </p>
            </div>
          </div>

          {/* Controls: Active Semester & Quick Actions */}
          <div className="flex items-center gap-2 sm:gap-3">
            {/* Semester Switcher */}
            <div className="bg-slate-100 p-1 rounded-lg flex items-center border border-slate-200 text-xs font-medium">
              <button
                onClick={() => {
                  setActiveSemester(1);
                  setCurrentCalendarDate(new Date(tpYears.startYear, 6, 1));
                  showToast(`Semester 1 (Ganjil ${tpYears.startYear}) diaktifkan`, 'info');
                }}
                className={`px-3 py-1.5 rounded-md transition-all font-semibold cursor-pointer ${
                  activeSemester === 1 ? 'bg-white text-indigo-600 shadow-sm' : 'text-slate-600 hover:text-slate-900'
                }`}
                title={`Semester 1 (${tpYears.startYear})`}
              >
                Sem 1 ({tpYears.startYear})
              </button>
              <button
                onClick={() => {
                  setActiveSemester(2);
                  setCurrentCalendarDate(new Date(tpYears.endYear, 0, 1));
                  showToast(`Semester 2 (Genap ${tpYears.endYear}) diaktifkan`, 'info');
                }}
                className={`px-3 py-1.5 rounded-md transition-all font-semibold cursor-pointer ${
                  activeSemester === 2 ? 'bg-white text-indigo-600 shadow-sm' : 'text-slate-600 hover:text-slate-900'
                }`}
                title={`Semester 2 (${tpYears.endYear})`}
              >
                Sem 2 ({tpYears.endYear})
              </button>
            </div>

            {/* Add Event Quick Button */}
            <button
              onClick={() => openAddAgendaModal('school')}
              className="bg-indigo-600 hover:bg-indigo-700 text-white px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5 shadow-sm transition"
            >
              <i className="fa-solid fa-plus"></i>
              <span className="hidden sm:inline">Tambah Agenda</span>
            </button>

            {/* Print Dialog Trigger */}
            <button
              onClick={() => setPrintModalOpen(true)}
              className="bg-slate-800 hover:bg-slate-900 text-white px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5 shadow-sm transition"
            >
              <i className="fa-solid fa-print"></i>
              <span className="hidden sm:inline">Cetak / PDF</span>
            </button>
          </div>
        </header>

        {/* TAB CONTENTS */}
        <div className="p-4 sm:p-6 space-y-6 flex-1">
          {/* 1. DASHBOARD TAB */}
          {activeTab === 'dashboard' && (
            <div className="space-y-6">
              {/* Banner */}
              <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 text-white rounded-2xl p-5 sm:p-6 shadow-md border border-slate-800 relative overflow-hidden">
                <div className="relative z-10 flex flex-col md:flex-row md:items-center justify-between gap-4">
                  <div className="flex items-center gap-4">
                    {/* Logo Kiri */}
                    {appData.settings.logoLeft && (
                      <div className="w-14 h-14 sm:w-16 sm:h-16 flex-shrink-0 bg-white/10 backdrop-blur-md rounded-xl p-1.5 flex items-center justify-center border border-white/20">
                        <img
                          src={appData.settings.logoLeft}
                          alt="Logo Kabupaten"
                          className="max-w-full max-h-full object-contain"
                        />
                      </div>
                    )}
                    <div className="space-y-1">
                      <div className="inline-flex items-center gap-2 px-2.5 py-1 rounded-full bg-indigo-500/20 text-indigo-300 text-xs border border-indigo-500/30 font-medium mb-1">
                        <i className="fa-solid fa-graduation-cap"></i>
                        <span>NPSN: {appData.settings.schoolNPSN}</span>
                      </div>
                      <h2 className="text-xl sm:text-2xl font-extrabold text-white">
                        {appData.settings.schoolName}
                      </h2>
                      <p className="text-xs sm:text-sm text-slate-300">{appData.settings.address}</p>
                    </div>
                  </div>

                  <div className="flex items-center gap-4 self-start md:self-auto">
                    {/* Logo Kanan */}
                    {appData.settings.logoRight && (
                      <div className="w-14 h-14 sm:w-16 sm:h-16 flex-shrink-0 bg-white/10 backdrop-blur-md rounded-xl p-1.5 flex items-center justify-center border border-white/20">
                        <img
                          src={appData.settings.logoRight}
                          alt="Logo Sekolah"
                          className="max-w-full max-h-full object-contain"
                        />
                      </div>
                    )}
                    <div className="bg-white/10 backdrop-blur-md rounded-xl p-3 border border-white/10 flex items-center gap-4">
                      <div className="text-center px-2">
                        <div className="text-[10px] uppercase text-indigo-200 font-semibold tracking-wider">
                          Tahun Pelajaran
                        </div>
                        <div className="text-lg font-bold text-white">{appData.settings.academicYear}</div>
                      </div>
                      <div className="w-px h-8 bg-white/20"></div>
                      <div className="text-center px-2">
                        <div className="text-[10px] uppercase text-indigo-200 font-semibold tracking-wider">
                          Semester Aktif
                        </div>
                        <div className="text-lg font-bold text-emerald-400">
                          {activeSemester === 1 ? 'Semester 1 (Ganjil)' : 'Semester 2 (Genap)'}
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              {/* Statistics Cards */}
              <div className="grid grid-cols-2 lg:grid-cols-6 gap-3 sm:gap-4">
                <div className="bg-white rounded-xl p-4 border border-slate-200 shadow-sm flex flex-col justify-between">
                  <div className="flex items-center justify-between text-slate-500 mb-2">
                    <span className="text-xs font-medium">Total Agenda</span>
                    <div className="w-8 h-8 rounded-lg bg-slate-100 flex items-center justify-center text-slate-600">
                      <i className="fa-solid fa-list-check text-sm"></i>
                    </div>
                  </div>
                  <div className="text-2xl font-bold text-slate-800">{appData.events.length}</div>
                  <div className="text-[11px] text-slate-500 mt-1">Keseluruhan entri</div>
                </div>

                <div className="bg-white rounded-xl p-4 border border-slate-200 shadow-sm flex flex-col justify-between">
                  <div className="flex items-center justify-between text-rose-600 mb-2">
                    <span className="text-xs font-medium">Libur Nasional</span>
                    <div className="w-8 h-8 rounded-lg bg-rose-50 flex items-center justify-center text-rose-600">
                      <i className="fa-solid fa-calendar-xmark text-sm"></i>
                    </div>
                  </div>
                  <div className="text-2xl font-bold text-rose-600">{categoryCounts.holiday}</div>
                  <div className="text-[11px] text-slate-500 mt-1">Hari libur resmi</div>
                </div>

                <div className="bg-white rounded-xl p-4 border border-slate-200 shadow-sm flex flex-col justify-between">
                  <div className="flex items-center justify-between text-sky-600 mb-2">
                    <span className="text-xs font-medium">Agenda Sekolah</span>
                    <div className="w-8 h-8 rounded-lg bg-sky-50 flex items-center justify-center text-sky-600">
                      <i className="fa-solid fa-school text-sm"></i>
                    </div>
                  </div>
                  <div className="text-2xl font-bold text-sky-600">{categoryCounts.school}</div>
                  <div className="text-[11px] text-slate-500 mt-1">Kegiatan institusi</div>
                </div>

                <div className="bg-white rounded-xl p-4 border border-slate-200 shadow-sm flex flex-col justify-between">
                  <div className="flex items-center justify-between text-purple-600 mb-2">
                    <span className="text-xs font-medium">Agenda Guru</span>
                    <div className="w-8 h-8 rounded-lg bg-purple-50 flex items-center justify-center text-purple-600">
                      <i className="fa-solid fa-chalkboard-user text-sm"></i>
                    </div>
                  </div>
                  <div className="text-2xl font-bold text-purple-600">{categoryCounts.teacher}</div>
                  <div className="text-[11px] text-slate-500 mt-1">Kombel & Pelatihan</div>
                </div>

                <div className="bg-white rounded-xl p-4 border border-slate-200 shadow-sm flex flex-col justify-between">
                  <div className="flex items-center justify-between text-amber-600 mb-2">
                    <span className="text-xs font-medium">Peristiwa Penting</span>
                    <div className="w-8 h-8 rounded-lg bg-amber-50 flex items-center justify-center text-amber-600">
                      <i className="fa-solid fa-star text-sm"></i>
                    </div>
                  </div>
                  <div className="text-2xl font-bold text-amber-600">{categoryCounts.important}</div>
                  <div className="text-[11px] text-slate-500 mt-1">Hari besar & momen</div>
                </div>

                <div className="bg-white rounded-xl p-4 border border-emerald-200 bg-emerald-50/30 shadow-sm flex flex-col justify-between">
                  <div className="flex items-center justify-between text-emerald-700 mb-2">
                    <span className="text-xs font-semibold">Total HEB Active</span>
                    <div className="w-8 h-8 rounded-lg bg-emerald-100 flex items-center justify-center text-emerald-700">
                      <i className="fa-solid fa-calculator text-sm"></i>
                    </div>
                  </div>
                  <div className="text-2xl font-extrabold text-emerald-700">{activeSemesterHEB.totalHEB}</div>
                  <div className="text-[11px] text-emerald-600 font-medium mt-1">Hari belajar efektif</div>
                </div>
              </div>

              {/* Calendar & Upcoming Grid */}
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                <div className="lg:col-span-2 bg-white rounded-2xl border border-slate-200 shadow-sm p-4 sm:p-5 flex flex-col">
                  {/* Calendar Navigation Header */}
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4 pb-3 border-b border-slate-100">
                    <div className="flex items-center gap-3">
                      <h3 className="text-lg font-bold text-slate-800">
                        {getMonthNameIndo(currentCalendarDate.getMonth())} {currentCalendarDate.getFullYear()}
                      </h3>
                      <span className={`text-xs px-2.5 py-0.5 rounded-full font-medium ${
                        currentCalendarDate.getMonth() >= 6
                          ? 'bg-indigo-50 text-indigo-700 border border-indigo-200'
                          : 'bg-emerald-50 text-emerald-700 border border-emerald-200'
                      }`}>
                        {currentCalendarDate.getMonth() >= 6 ? `Sem 1 (${tpYears.startYear})` : `Sem 2 (${tpYears.endYear})`}
                      </span>
                      <span className="text-xs px-2 py-0.5 rounded-full bg-slate-100 text-slate-600 font-semibold hidden md:inline">
                        TP {appData.settings.academicYear}
                      </span>
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => navigateMonth(-1)}
                        className="p-2 rounded-lg border border-slate-200 hover:bg-slate-50 text-slate-600 transition text-xs cursor-pointer"
                        title="Bulan Sebelumnya"
                      >
                        <i className="fa-solid fa-chevron-left"></i>
                      </button>
                      <button
                        onClick={() => {
                          setCurrentCalendarDate(new Date(tpYears.startYear, 6, 1));
                          setActiveSemester(1);
                        }}
                        className="px-2.5 py-1.5 rounded-lg border border-slate-200 hover:bg-slate-50 text-slate-700 font-semibold text-xs transition cursor-pointer"
                        title={`Lompat ke Awal Tahun Pelajaran (Juli ${tpYears.startYear})`}
                      >
                        Awal TP
                      </button>
                      <button
                        onClick={goToCurrentMonth}
                        className="px-3 py-1.5 rounded-lg border border-slate-200 hover:bg-slate-50 text-slate-700 font-semibold text-xs transition cursor-pointer"
                        title="Ke Bulan Hari Ini"
                      >
                        Hari Ini
                      </button>
                      <button
                        onClick={() => navigateMonth(1)}
                        className="p-2 rounded-lg border border-slate-200 hover:bg-slate-50 text-slate-600 transition text-xs cursor-pointer"
                        title="Bulan Berikutnya"
                      >
                        <i className="fa-solid fa-chevron-right"></i>
                      </button>
                    </div>
                  </div>

                  {/* Category Color Legend */}
                  <div className="flex flex-wrap items-center gap-3 mb-4 text-xs">
                    <div className="flex items-center gap-1.5">
                      <span className="w-3 h-3 rounded-full bg-rose-500 inline-block"></span>
                      <span className="text-slate-600">Libur Nasional</span>
                    </div>
                    <div className="flex items-center gap-1.5">
                      <span className="w-3 h-3 rounded-full bg-sky-500 inline-block"></span>
                      <span className="text-slate-600">Agenda Sekolah</span>
                    </div>
                    <div className="flex items-center gap-1.5">
                      <span className="w-3 h-3 rounded-full bg-purple-500 inline-block"></span>
                      <span className="text-slate-600">Agenda Guru</span>
                    </div>
                    <div className="flex items-center gap-1.5">
                      <span className="w-3 h-3 rounded-full bg-amber-500 inline-block"></span>
                      <span className="text-slate-600">Peristiwa Penting</span>
                    </div>
                  </div>

                  {/* Calendar Grid Component */}
                  {renderCalendarMonthGrid(currentCalendarDate)}
                </div>

                {/* Right Side Column: Today & Upcoming */}
                <div className="space-y-6">
                  {/* Today Card */}
                  <div className="bg-gradient-to-br from-indigo-50 to-slate-50 rounded-2xl border border-indigo-100 p-4 sm:p-5 shadow-sm">
                    <div className="flex items-center justify-between mb-3">
                      <div className="text-xs font-bold uppercase tracking-wider text-indigo-700">Tanggal Hari Ini</div>
                      <span className="text-xs bg-indigo-100 text-indigo-700 px-2 py-0.5 rounded font-mono font-medium">
                        {formatDateIndo(todayStr)}
                      </span>
                    </div>
                    <div className="text-2xl font-extrabold text-slate-900">
                      {['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'][new Date().getDay()]}
                    </div>
                    <div className="text-sm text-slate-500 font-medium">
                      {new Date().getDate()} {getMonthNameIndo(new Date().getMonth())} {new Date().getFullYear()}
                    </div>

                    <div className="mt-4 pt-3 border-t border-indigo-100/80">
                      <div className="text-xs font-semibold text-slate-700 mb-2">Kegiatan Hari Ini:</div>
                      {todayEvents.length === 0 ? (
                        <p className="text-xs text-slate-400 italic">Tidak ada kegiatan terjadwal hari ini.</p>
                      ) : (
                        <div className="space-y-2">
                          {todayEvents.map(ev => {
                            const cat = CATEGORIES[ev.category] || CATEGORIES.school;
                            return (
                              <div
                                key={ev.id}
                                className={`p-2 rounded-lg border text-xs ${cat.badgeClass} flex items-center justify-between`}
                              >
                                <div>
                                  <div className="font-bold">{ev.title}</div>
                                  {ev.time && (
                                    <div className="text-[10px] opacity-80">
                                      <i className="fa-regular fa-clock mr-1"></i>{ev.time}
                                    </div>
                                  )}
                                </div>
                                <span className="text-[10px] uppercase font-bold px-1.5 py-0.5 rounded bg-white/70">
                                  {cat.name}
                                </span>
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  </div>

                  {/* Upcoming Agendas */}
                  <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-4 sm:p-5">
                    <div className="flex items-center justify-between mb-4">
                      <h4 className="font-bold text-slate-800 text-sm">Agenda Mendatang</h4>
                      <button
                        onClick={() => setActiveTab('rekap')}
                        className="text-indigo-600 hover:text-indigo-800 text-xs font-semibold"
                      >
                        Lihat Semua
                      </button>
                    </div>

                    {upcomingAgendas.length === 0 ? (
                      <p className="text-xs text-slate-400 italic">Tidak ada agenda mendatang.</p>
                    ) : (
                      <div className="space-y-3">
                        {upcomingAgendas.map(ev => {
                          const cat = CATEGORIES[ev.category] || CATEGORIES.school;
                          return (
                            <div key={ev.id} className="p-2.5 rounded-xl border border-slate-100 bg-slate-50 flex items-start gap-3">
                              <div className={`w-2 h-2 rounded-full ${cat.bgClass} mt-1.5 flex-shrink-0`}></div>
                              <div className="flex-1 min-w-0">
                                <div className="text-xs font-bold text-slate-800 truncate">{ev.title}</div>
                                <div className="text-[11px] text-slate-500 flex items-center gap-2 mt-0.5">
                                  <span>
                                    <i className="fa-regular fa-calendar text-[10px] mr-1"></i>
                                    {formatDateIndo(ev.startDate)} {ev.startDate !== ev.endDate ? `- ${formatDateIndo(ev.endDate)}` : ''}
                                  </span>
                                </div>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* 2. KALENDER MASTER TAB */}
          {activeTab === 'kalender' && (
            <div className="bg-white rounded-2xl border border-slate-200 p-4 sm:p-6 shadow-sm">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
                <div>
                  <div className="flex items-center gap-2.5">
                    <h3 className="text-lg font-bold text-slate-800">Kalender Master Pendidikan</h3>
                    <span className="text-xs bg-indigo-100 text-indigo-800 font-bold px-2.5 py-0.5 rounded-full border border-indigo-200">
                      Tahun Pelajaran {appData.settings.academicYear}
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mt-1">
                    Sinkron otomatis dengan Tahun Pelajaran aktif ({tpYears.startYear}/{tpYears.endYear}) sesuai pengaturan sistem
                  </p>
                </div>
                <div className="flex items-center gap-2 bg-slate-100 p-1 rounded-lg text-xs font-medium border border-slate-200">
                  <button
                    onClick={() => setCalendarViewMode('month')}
                    className={`px-3 py-1.5 rounded-md transition font-semibold cursor-pointer ${
                      calendarViewMode === 'month' ? 'bg-white text-indigo-600 shadow-sm' : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Tampilan Bulanan
                  </button>
                  <button
                    onClick={() => setCalendarViewMode('semester')}
                    className={`px-3 py-1.5 rounded-md transition font-semibold cursor-pointer ${
                      calendarViewMode === 'semester' ? 'bg-white text-indigo-600 shadow-sm' : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Matriks Semester (6 Bulan)
                  </button>
                </div>
              </div>

              {calendarViewMode === 'month' ? (
                <div>
                  {/* Calendar Navigation & Academic Year Selector Bar */}
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 mb-4 bg-slate-50 p-3 sm:p-4 rounded-xl border border-slate-200">
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => navigateMonth(-1)}
                        className="px-3 py-1.5 rounded-lg bg-white border border-slate-200 text-slate-700 hover:bg-slate-100 text-xs font-semibold cursor-pointer shadow-sm transition"
                        title="Bulan Sebelumnya"
                      >
                        <i className="fa-solid fa-chevron-left mr-1"></i> Bulan Lalu
                      </button>
                      <button
                        onClick={() => navigateMonth(1)}
                        className="px-3 py-1.5 rounded-lg bg-white border border-slate-200 text-slate-700 hover:bg-slate-100 text-xs font-semibold cursor-pointer shadow-sm transition"
                        title="Bulan Berikutnya"
                      >
                        Bulan Depan <i className="fa-solid fa-chevron-right ml-1"></i>
                      </button>
                    </div>

                    <div className="text-center">
                      <div className="flex items-center justify-center gap-2">
                        <h3 className="text-base sm:text-lg font-bold text-slate-800">
                          {getMonthNameIndo(currentCalendarDate.getMonth())} {currentCalendarDate.getFullYear()}
                        </h3>
                        <span className={`text-[11px] font-bold px-2.5 py-0.5 rounded-full border ${
                          currentCalendarDate.getMonth() >= 6
                            ? 'bg-indigo-100 text-indigo-800 border-indigo-200'
                            : 'bg-emerald-100 text-emerald-800 border-emerald-200'
                        }`}>
                          {currentCalendarDate.getMonth() >= 6 ? `Sem 1 (Ganjil ${tpYears.startYear})` : `Sem 2 (Genap ${tpYears.endYear})`}
                        </span>
                      </div>
                      <div className="text-[11px] text-slate-500 font-medium mt-0.5">
                        Tahun Pelajaran {appData.settings.academicYear}
                      </div>
                    </div>

                    <div className="flex flex-wrap items-center gap-2">
                      {/* Direct 12-Month Selector of Configured TP */}
                      <select
                        value={`${currentCalendarDate.getFullYear()}-${currentCalendarDate.getMonth()}`}
                        onChange={e => {
                          const [y, m] = e.target.value.split('-').map(Number);
                          setCurrentCalendarDate(new Date(y, m, 1));
                          if (m >= 6 && activeSemester !== 1) setActiveSemester(1);
                          if (m < 6 && activeSemester !== 2) setActiveSemester(2);
                        }}
                        className="text-xs bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 font-semibold text-slate-700 focus:ring-2 focus:ring-indigo-500 cursor-pointer shadow-sm"
                        title="Pilih langsung bulan dalam Tahun Pelajaran ini"
                      >
                        <optgroup label={`Semester 1 (${tpYears.startYear})`}>
                          {tpMonths.filter(m => m.month >= 6).map(m => (
                            <option key={`${m.year}-${m.month}`} value={`${m.year}-${m.month}`}>
                              {m.label} (Sem 1)
                            </option>
                          ))}
                        </optgroup>
                        <optgroup label={`Semester 2 (${tpYears.endYear})`}>
                          {tpMonths.filter(m => m.month < 6).map(m => (
                            <option key={`${m.year}-${m.month}`} value={`${m.year}-${m.month}`}>
                              {m.label} (Sem 2)
                            </option>
                          ))}
                        </optgroup>
                      </select>

                      <button
                        onClick={() => {
                          setCurrentCalendarDate(new Date(tpYears.startYear, 6, 1));
                          setActiveSemester(1);
                        }}
                        className="px-2.5 py-1.5 rounded-lg bg-white border border-slate-200 text-slate-700 hover:bg-slate-100 text-xs font-semibold cursor-pointer shadow-sm transition"
                        title={`Lompat ke Awal Tahun Pelajaran (Juli ${tpYears.startYear})`}
                      >
                        Awal TP
                      </button>
                      <button
                        onClick={goToCurrentMonth}
                        className="px-2.5 py-1.5 rounded-lg bg-indigo-50 border border-indigo-200 text-indigo-700 hover:bg-indigo-100 text-xs font-semibold cursor-pointer shadow-sm transition"
                        title="Ke Bulan Hari Ini"
                      >
                        Hari Ini
                      </button>
                    </div>
                  </div>

                  {/* Category Color Legend */}
                  <div className="flex flex-wrap items-center gap-3 mb-4 text-xs bg-slate-50/70 p-2.5 rounded-lg border border-slate-100">
                    <span className="font-semibold text-slate-500 text-[11px] mr-1">Keterangan Warna:</span>
                    <div className="flex items-center gap-1.5">
                      <span className="w-3 h-3 rounded-full bg-rose-500 inline-block"></span>
                      <span className="text-slate-600">Libur Nasional</span>
                    </div>
                    <div className="flex items-center gap-1.5">
                      <span className="w-3 h-3 rounded-full bg-sky-500 inline-block"></span>
                      <span className="text-slate-600">Agenda Sekolah</span>
                    </div>
                    <div className="flex items-center gap-1.5">
                      <span className="w-3 h-3 rounded-full bg-purple-500 inline-block"></span>
                      <span className="text-slate-600">Agenda Guru</span>
                    </div>
                    <div className="flex items-center gap-1.5">
                      <span className="w-3 h-3 rounded-full bg-amber-500 inline-block"></span>
                      <span className="text-slate-600">Peristiwa Penting</span>
                    </div>
                  </div>

                  {renderCalendarMonthGrid(currentCalendarDate)}
                </div>
              ) : (
                <div>
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-5 bg-indigo-50 p-3 sm:p-4 rounded-xl border border-indigo-100">
                    <div>
                      <div className="font-bold text-slate-800 text-sm">
                        Matriks 6 Bulan: {activeSemester === 1 ? `Semester 1 (Ganjil): Juli - Desember ${tpYears.startYear}` : `Semester 2 (Genap): Januari - Juni ${tpYears.endYear}`}
                      </div>
                      <p className="text-xs text-indigo-700 font-medium mt-0.5">
                        Tahun Pelajaran {appData.settings.academicYear} | Sesuai Pengaturan Sistem
                      </p>
                    </div>
                    <div className="flex items-center gap-2 bg-white p-1 rounded-lg border border-indigo-200 text-xs">
                      <button
                        onClick={() => {
                          setActiveSemester(1);
                          setCurrentCalendarDate(new Date(tpYears.startYear, 6, 1));
                        }}
                        className={`px-3 py-1 rounded-md font-semibold transition cursor-pointer ${
                          activeSemester === 1 ? 'bg-indigo-600 text-white shadow-sm' : 'text-slate-600 hover:text-slate-900'
                        }`}
                      >
                        Sem 1 ({tpYears.startYear})
                      </button>
                      <button
                        onClick={() => {
                          setActiveSemester(2);
                          setCurrentCalendarDate(new Date(tpYears.endYear, 0, 1));
                        }}
                        className={`px-3 py-1 rounded-md font-semibold transition cursor-pointer ${
                          activeSemester === 2 ? 'bg-indigo-600 text-white shadow-sm' : 'text-slate-600 hover:text-slate-900'
                        }`}
                      >
                        Sem 2 ({tpYears.endYear})
                      </button>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                    {activeSemester === 1
                      ? [6, 7, 8, 9, 10, 11].map(m => (
                          <div key={m}>
                            {renderMiniCalendar(tpYears.startYear, m)}
                          </div>
                        ))
                      : [0, 1, 2, 3, 4, 5].map(m => (
                          <div key={m}>
                            {renderMiniCalendar(tpYears.endYear, m)}
                          </div>
                        ))}
                  </div>
                </div>
              )}
            </div>
          )}

          {/* 3. LIBUR NASIONAL TAB */}
          {activeTab === 'holidays' && renderCategoryCRUDTable('holiday')}

          {/* 4. AGENDA SEKOLAH TAB */}
          {activeTab === 'school-agenda' && renderCategoryCRUDTable('school')}

          {/* 5. AGENDA GURU TAB */}
          {activeTab === 'teacher-agenda' && renderCategoryCRUDTable('teacher')}

          {/* 6. PERISTIWA PENTING TAB */}
          {activeTab === 'events' && renderCategoryCRUDTable('important')}

          {/* 7. REKAP AGENDA TAB */}
          {activeTab === 'rekap' && (
            <div className="bg-white rounded-2xl border border-slate-200 p-4 sm:p-6 shadow-sm">
              <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-6">
                <div>
                  <h3 className="text-lg font-bold text-slate-800">Rekapitulasi Agenda & Kegiatan</h3>
                  <p className="text-xs text-slate-500">Filter, analisis, dan ekspor laporan kegiatan akademik lengkap</p>
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  <button
                    onClick={exportRekapExcel}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white px-3.5 py-2 rounded-lg text-xs font-semibold flex items-center gap-2 shadow-sm transition"
                  >
                    <i className="fa-solid fa-file-excel"></i> Ekspor Excel
                  </button>
                  <button
                    onClick={() => {
                      setPrintType('rekap');
                      setPrintModalOpen(true);
                    }}
                    className="bg-slate-800 hover:bg-slate-900 text-white px-3.5 py-2 rounded-lg text-xs font-semibold flex items-center gap-2 shadow-sm transition"
                  >
                    <i className="fa-solid fa-print"></i> Cetak Rekap
                  </button>
                </div>
              </div>

              {/* Filters Bar */}
              <div className="bg-slate-50 rounded-xl p-4 border border-slate-200 mb-6 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">Kategori</label>
                  <select
                    value={rekapFilterCat}
                    onChange={e => setRekapFilterCat(e.target.value)}
                    className="w-full text-xs rounded-lg border-slate-300 bg-white p-2 border focus:ring-2 focus:ring-indigo-500"
                  >
                    <option value="all">-- Semua Kategori --</option>
                    <option value="holiday">Hari Libur Nasional</option>
                    <option value="school">Agenda Sekolah</option>
                    <option value="teacher">Agenda Guru</option>
                    <option value="important">Peristiwa Penting</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">Semester</label>
                  <select
                    value={rekapFilterSem}
                    onChange={e => setRekapFilterSem(e.target.value)}
                    className="w-full text-xs rounded-lg border-slate-300 bg-white p-2 border focus:ring-2 focus:ring-indigo-500"
                  >
                    <option value="all">-- Semua Semester --</option>
                    <option value="1">Semester 1 (Ganjil)</option>
                    <option value="2">Semester 2 (Genap)</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">Bulan Spesifik</label>
                  <select
                    value={rekapFilterMonth}
                    onChange={e => setRekapFilterMonth(e.target.value)}
                    className="w-full text-xs rounded-lg border-slate-300 bg-white p-2 border focus:ring-2 focus:ring-indigo-500"
                  >
                    <option value="all">-- Semua Bulan --</option>
                    <option value="0">Januari</option>
                    <option value="1">Februari</option>
                    <option value="2">Maret</option>
                    <option value="3">April</option>
                    <option value="4">Mei</option>
                    <option value="5">Juni</option>
                    <option value="6">Juli</option>
                    <option value="7">Agustus</option>
                    <option value="8">September</option>
                    <option value="9">Oktober</option>
                    <option value="10">November</option>
                    <option value="11">Desember</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">Cari Kata Kunci</label>
                  <div className="relative">
                    <input
                      type="text"
                      value={rekapSearchInput}
                      onChange={e => setRekapSearchInput(e.target.value)}
                      placeholder="Nama kegiatan / PIC..."
                      className="w-full text-xs rounded-lg border-slate-300 bg-white p-2 pl-8 border focus:ring-2 focus:ring-indigo-500"
                    />
                    <i className="fa-solid fa-magnifying-glass absolute left-2.5 top-2.5 text-slate-400 text-xs"></i>
                  </div>
                </div>
              </div>

              {/* Summary Bar */}
              <div className="flex items-center justify-between text-xs text-slate-600 bg-indigo-50/50 p-3 rounded-lg border border-indigo-100 mb-4">
                <div>
                  Menampilkan <span className="font-bold text-indigo-700">{filteredRekapEvents.length}</span> item agenda sesuai filter
                </div>
                <button
                  onClick={() => {
                    setRekapFilterCat('all');
                    setRekapFilterSem('all');
                    setRekapFilterMonth('all');
                    setRekapSearchInput('');
                  }}
                  className="text-indigo-600 hover:text-indigo-800 font-semibold flex items-center gap-1"
                >
                  <i className="fa-solid fa-rotate-left"></i> Reset Filter
                </button>
              </div>

              {/* Data Table */}
              <div className="overflow-x-auto border border-slate-200 rounded-xl">
                <table className="w-full text-left border-collapse text-sm">
                  <thead className="bg-slate-50 text-slate-700 text-xs uppercase font-semibold border-b border-slate-200">
                    <tr>
                      <th className="py-3 px-4">No</th>
                      <th className="py-3 px-4">Kategori</th>
                      <th className="py-3 px-4">Tanggal / Waktu</th>
                      <th className="py-3 px-4">Nama Kegiatan</th>
                      <th className="py-3 px-4">Tempat / PIC</th>
                      <th className="py-3 px-4">Keterangan</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200 text-slate-700">
                    {filteredRekapEvents.length === 0 ? (
                      <tr>
                        <td colSpan={6} className="py-8 text-center text-xs text-slate-400 italic">
                          Tidak ada data agenda yang memenuhi kriteria filter.
                        </td>
                      </tr>
                    ) : (
                      filteredRekapEvents.map((ev, idx) => {
                        const cat = CATEGORIES[ev.category] || CATEGORIES.school;
                        const dateRange = ev.startDate === ev.endDate ? formatDateIndo(ev.startDate) : `${formatDateIndo(ev.startDate)} s.d. ${formatDateIndo(ev.endDate)}`;
                        return (
                          <tr key={ev.id} className="hover:bg-slate-50 transition text-xs">
                            <td className="py-3 px-4 font-mono text-slate-400">{idx + 1}</td>
                            <td className="py-3 px-4">
                              <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${cat.badgeClass}`}>
                                {cat.name}
                              </span>
                            </td>
                            <td className="py-3 px-4">
                              <div className="font-bold text-slate-800">{dateRange}</div>
                              {ev.time && <div className="text-[10px] text-slate-400">{ev.time}</div>}
                            </td>
                            <td className="py-3 px-4 font-bold text-slate-800">{ev.title}</td>
                            <td className="py-3 px-4">
                              <div>{ev.location || '-'}</div>
                              {ev.pic && <div className="text-[10px] text-slate-400">PIC: {ev.pic}</div>}
                            </td>
                            <td className="py-3 px-4 text-slate-600">{ev.description || '-'}</td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* 8. HARI EFEKTIF BELAJAR (HEB) TAB */}
          {activeTab === 'heb' && (
            <div className="bg-white rounded-2xl border border-slate-200 p-4 sm:p-6 shadow-sm">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
                <div>
                  <h3 className="text-lg font-bold text-slate-800">Perhitungan Hari Efektif Belajar (HEB)</h3>
                  <p className="text-xs text-slate-500">
                    Kalkulasi otomatis hari belajar (Senin-Jumat) dikurangi libur nasional & libur sekolah
                  </p>
                </div>
                <button
                  onClick={() => {
                    setPrintType('heb');
                    setPrintModalOpen(true);
                  }}
                  className="bg-slate-800 hover:bg-slate-900 text-white px-3.5 py-2 rounded-lg text-xs font-semibold flex items-center gap-2 shadow-sm transition"
                >
                  <i className="fa-solid fa-print"></i> Cetak Analisis HEB
                </button>
              </div>

              {/* Formula Info Card */}
              <div className="bg-blue-50/70 border border-blue-200 rounded-xl p-4 text-xs text-slate-700 space-y-1.5 mb-6">
                <div className="font-bold text-blue-900 flex items-center gap-2">
                  <i className="fa-solid fa-circle-info text-blue-600"></i> Ketentuan & Formula Perhitungan HEB:
                </div>
                <ul className="list-disc list-inside space-y-1 text-slate-600 pl-1">
                  <li>Hari Kerja Standar: <strong>Senin - Jumat</strong> (Sabtu & Minggu tidak dihitung sebagai hari efektif).</li>
                  <li>Hari Libur Nasional & Cuti Bersama yang jatuh pada hari kerja otomatis memotong HEB.</li>
                  <li>Agenda Sekolah yang ditandai sebagai <em>"Hari Tidak Efektif"</em> (misal Libur Semester, MPLS) memotong HEB.</li>
                  <li>Jumlah HEB dapat disesuaikan manual (override) jika terdapat dispensasi atau penyesuaian edaran dinas.</li>
                </ul>
              </div>

              {/* HEB Breakdown Matrix Tables */}
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                {/* Semester 1 Matrix */}
                <div className="border border-slate-200 rounded-xl overflow-hidden shadow-sm">
                  <div className="bg-slate-800 text-white px-4 py-3 font-bold text-sm flex justify-between items-center">
                    <span>Semester 1 (Ganjil)</span>
                    <span className="bg-indigo-500 text-white text-xs px-2.5 py-0.5 rounded-full">
                      {sem1HEB.totalHEB} Hari HEB
                    </span>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse text-xs">
                      <thead className="bg-slate-100 text-slate-700 uppercase font-semibold border-b border-slate-200">
                        <tr>
                          <th className="py-2.5 px-3">Bulan</th>
                          <th className="py-2.5 px-2 text-center">Hari Kerja</th>
                          <th className="py-2.5 px-2 text-center">Libur Nas.</th>
                          <th className="py-2.5 px-2 text-center">Non-Efektif</th>
                          <th className="py-2.5 px-3 text-center bg-indigo-50 text-indigo-900 font-bold">Total HEB</th>
                          <th className="py-2.5 px-2 text-center">Penyesuaian</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200 text-slate-700">
                        {sem1HEB.monthlyBreakdown.map(item => (
                          <tr key={item.monthKey} className="hover:bg-slate-50 transition">
                            <td className="py-2.5 px-3 font-semibold text-slate-800">
                              {getMonthNameIndo(item.month)} {item.year}
                            </td>
                            <td className="py-2.5 px-2 text-center text-slate-600 font-mono">{item.workDaysCount}</td>
                            <td className="py-2.5 px-2 text-center text-rose-600 font-mono font-medium">{item.holidayCount}</td>
                            <td className="py-2.5 px-2 text-center text-amber-600 font-mono font-medium">{item.nonEffectiveSchoolDays}</td>
                            <td className="py-2.5 px-3 text-center bg-indigo-50 font-bold text-indigo-900 font-mono text-sm">{item.finalHEB}</td>
                            <td className="py-2.5 px-2 text-center">
                              <input
                                type="number"
                                min="0"
                                max="31"
                                value={item.finalHEB}
                                onChange={e => updateHEBOverride(item.monthKey, e.target.value)}
                                className="w-14 text-center p-1 rounded border border-slate-300 focus:ring-1 focus:ring-indigo-500 text-xs font-mono"
                              />
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* Semester 2 Matrix */}
                <div className="border border-slate-200 rounded-xl overflow-hidden shadow-sm">
                  <div className="bg-slate-800 text-white px-4 py-3 font-bold text-sm flex justify-between items-center">
                    <span>Semester 2 (Genap)</span>
                    <span className="bg-emerald-500 text-white text-xs px-2.5 py-0.5 rounded-full">
                      {sem2HEB.totalHEB} Hari HEB
                    </span>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse text-xs">
                      <thead className="bg-slate-100 text-slate-700 uppercase font-semibold border-b border-slate-200">
                        <tr>
                          <th className="py-2.5 px-3">Bulan</th>
                          <th className="py-2.5 px-2 text-center">Hari Kerja</th>
                          <th className="py-2.5 px-2 text-center">Libur Nas.</th>
                          <th className="py-2.5 px-2 text-center">Non-Efektif</th>
                          <th className="py-2.5 px-3 text-center bg-emerald-50 text-emerald-900 font-bold">Total HEB</th>
                          <th className="py-2.5 px-2 text-center">Penyesuaian</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200 text-slate-700">
                        {sem2HEB.monthlyBreakdown.map(item => (
                          <tr key={item.monthKey} className="hover:bg-slate-50 transition">
                            <td className="py-2.5 px-3 font-semibold text-slate-800">
                              {getMonthNameIndo(item.month)} {item.year}
                            </td>
                            <td className="py-2.5 px-2 text-center text-slate-600 font-mono">{item.workDaysCount}</td>
                            <td className="py-2.5 px-2 text-center text-rose-600 font-mono font-medium">{item.holidayCount}</td>
                            <td className="py-2.5 px-2 text-center text-amber-600 font-mono font-medium">{item.nonEffectiveSchoolDays}</td>
                            <td className="py-2.5 px-3 text-center bg-emerald-50 font-bold text-emerald-900 font-mono text-sm">{item.finalHEB}</td>
                            <td className="py-2.5 px-2 text-center">
                              <input
                                type="number"
                                min="0"
                                max="31"
                                value={item.finalHEB}
                                onChange={e => updateHEBOverride(item.monthKey, e.target.value)}
                                className="w-14 text-center p-1 rounded border border-slate-300 focus:ring-1 focus:ring-indigo-500 text-xs font-mono"
                              />
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>

              {/* Total Annual Summary */}
              <div className="bg-gradient-to-r from-slate-900 to-indigo-950 rounded-xl p-5 text-white flex flex-col sm:flex-row items-center justify-between gap-4 shadow-md mt-6">
                <div>
                  <div className="text-xs text-indigo-300 uppercase font-semibold">
                    Total Hari Efektif Belajar (HEB) 1 Tahun Pelajaran
                  </div>
                  <div className="text-slate-300 text-xs mt-0.5">
                    Semester 1 ({sem1HEB.totalHEB} Hari) + Semester 2 ({sem2HEB.totalHEB} Hari)
                  </div>
                </div>
                <div className="text-3xl font-extrabold text-emerald-400 font-mono">
                  {sem1HEB.totalHEB + sem2HEB.totalHEB} <span className="text-base font-sans font-normal text-slate-300">Hari</span>
                </div>
              </div>
            </div>
          )}

          {/* 9. PENGATURAN TAB */}
          {activeTab === 'settings' && (
            <div className="bg-white rounded-2xl border border-slate-200 p-4 sm:p-6 shadow-sm">
              <h3 className="text-lg font-bold text-slate-800 mb-1">Pengaturan Aplikasi & Identitas Sekolah</h3>
              <p className="text-xs text-slate-500 mb-6">Konfigurasi profile sekolah, rentang semester, dan manajemen cadangan data</p>

              <form
                onSubmit={e => {
                  e.preventDefault();
                  // Automatically synchronize national holidays for the selected academic year
                  const updatedHolidays = getHolidaysForAcademicYear(
                    appData.settings.academicYear,
                    appData.settings.sem1Start,
                    appData.settings.sem2End
                  );
                  setAppData(prev => {
                    const otherEvents = prev.events.filter(ev => ev.category !== 'holiday');
                    return {
                      ...prev,
                      events: [...otherEvents, ...updatedHolidays]
                    };
                  });
                  // Sync calendar view to the newly configured academic year
                  const match = appData.settings.academicYear.match(/(\d{4})/);
                  const startYr = match ? parseInt(match[1], 10) : (new Date(appData.settings.sem1Start).getFullYear() || 2024);
                  setCurrentCalendarDate(new Date(startYr, 6, 1));
                  showToast(`Pengaturan disimpan! Kalender Master & Libur Nasional disesuaikan untuk TP ${appData.settings.academicYear}!`, 'success');
                }}
                className="space-y-6 max-w-4xl"
              >
                {/* School Profile Section */}
                <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-4">
                  <h4 className="font-bold text-slate-800 text-sm flex items-center gap-2">
                    <i className="fa-solid fa-school text-indigo-600"></i> Identitas Sekolah & Logo Header Kop
                  </h4>

                  {/* Logo Upload Fields */}
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 p-3 bg-white rounded-xl border border-slate-200">
                    {/* Logo Pemda (Kiri) */}
                    <div className="space-y-2">
                      <label className="block font-semibold text-slate-700 text-xs">
                        <i className="fa-solid fa-image text-indigo-600 mr-1"></i> Logo Kabupaten / Pemda (Kiri Kop)
                      </label>
                      <div className="flex items-center gap-3">
                        <div className="w-16 h-16 rounded-lg border border-slate-300 bg-slate-50 p-1 flex items-center justify-center overflow-hidden flex-shrink-0">
                          {appData.settings.logoLeft ? (
                            <img src={appData.settings.logoLeft} alt="Logo Kiri" className="max-w-full max-h-full object-contain" />
                          ) : (
                            <span className="text-[10px] text-slate-400 text-center">Belum ada logo</span>
                          )}
                        </div>
                        <div className="space-y-1.5 flex-1">
                          <label className="inline-block bg-indigo-50 hover:bg-indigo-100 text-indigo-700 text-xs font-semibold px-3 py-1.5 rounded-lg border border-indigo-200 cursor-pointer transition">
                            <i className="fa-solid fa-upload mr-1"></i> Pilih Logo Kiri
                            <input
                              type="file"
                              accept="image/*"
                              onChange={e => handleLogoUpload(e, 'left')}
                              className="hidden"
                            />
                          </label>
                          {appData.settings.logoLeft && (
                            <button
                              type="button"
                              onClick={() => removeLogo('left')}
                              className="block text-[11px] text-rose-600 hover:underline"
                            >
                              Hapus Logo Kiri
                            </button>
                          )}
                        </div>
                      </div>
                    </div>

                    {/* Logo Sekolah (Kanan) */}
                    <div className="space-y-2">
                      <label className="block font-semibold text-slate-700 text-xs">
                        <i className="fa-solid fa-image text-indigo-600 mr-1"></i> Logo Sekolah / Pendidikan (Kanan Kop)
                      </label>
                      <div className="flex items-center gap-3">
                        <div className="w-16 h-16 rounded-lg border border-slate-300 bg-slate-50 p-1 flex items-center justify-center overflow-hidden flex-shrink-0">
                          {appData.settings.logoRight ? (
                            <img src={appData.settings.logoRight} alt="Logo Kanan" className="max-w-full max-h-full object-contain" />
                          ) : (
                            <span className="text-[10px] text-slate-400 text-center">Belum ada logo</span>
                          )}
                        </div>
                        <div className="space-y-1.5 flex-1">
                          <label className="inline-block bg-indigo-50 hover:bg-indigo-100 text-indigo-700 text-xs font-semibold px-3 py-1.5 rounded-lg border border-indigo-200 cursor-pointer transition">
                            <i className="fa-solid fa-upload mr-1"></i> Pilih Logo Kanan
                            <input
                              type="file"
                              accept="image/*"
                              onChange={e => handleLogoUpload(e, 'right')}
                              className="hidden"
                            />
                          </label>
                          {appData.settings.logoRight && (
                            <button
                              type="button"
                              onClick={() => removeLogo('right')}
                              className="block text-[11px] text-rose-600 hover:underline"
                            >
                              Hapus Logo Kanan
                            </button>
                          )}
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Form fields */}
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs">
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">Nama Sekolah</label>
                      <input
                        type="text"
                        required
                        value={appData.settings.schoolName}
                        onChange={e => setAppData(prev => ({
                          ...prev,
                          settings: { ...prev.settings, schoolName: e.target.value }
                        }))}
                        className="w-full p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500 bg-white"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">NPSN / Identitas</label>
                      <input
                        type="text"
                        required
                        value={appData.settings.schoolNPSN}
                        onChange={e => setAppData(prev => ({
                          ...prev,
                          settings: { ...prev.settings, schoolNPSN: e.target.value }
                        }))}
                        className="w-full p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500 bg-white"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">Nama Kepala Sekolah</label>
                      <input
                        type="text"
                        required
                        value={appData.settings.principalName}
                        onChange={e => setAppData(prev => ({
                          ...prev,
                          settings: { ...prev.settings, principalName: e.target.value }
                        }))}
                        className="w-full p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500 bg-white"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">NIP Kepala Sekolah</label>
                      <input
                        type="text"
                        value={appData.settings.principalNIP}
                        onChange={e => setAppData(prev => ({
                          ...prev,
                          settings: { ...prev.settings, principalNIP: e.target.value }
                        }))}
                        className="w-full p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500 bg-white"
                      />
                    </div>
                    <div className="sm:col-span-2">
                      <label className="block font-semibold text-slate-700 mb-1">Alamat Sekolah</label>
                      <input
                        type="text"
                        required
                        value={appData.settings.address}
                        onChange={e => setAppData(prev => ({
                          ...prev,
                          settings: { ...prev.settings, address: e.target.value }
                        }))}
                        className="w-full p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500 bg-white"
                      />
                    </div>
                  </div>
                </div>

                {/* Academic Year & Semester Bounds */}
                <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-4">
                  <h4 className="font-bold text-slate-800 text-sm flex items-center gap-2">
                    <i className="fa-solid fa-calendar-range text-indigo-600"></i> Tahun Pelajaran & Rentang Semester
                  </h4>

                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 text-xs">
                    <div>
                      <div className="flex items-center justify-between mb-1">
                        <label className="block font-semibold text-slate-700">Tahun Pelajaran</label>
                        <button
                          type="button"
                          onClick={() => syncNationalHolidays(true)}
                          className="text-[10px] text-indigo-600 hover:text-indigo-800 font-bold flex items-center gap-1 cursor-pointer"
                          title="Sinkronkan Libur Nasional untuk Tahun Pelajaran ini"
                        >
                          <i className="fa-solid fa-arrows-rotate"></i> Sinkron Libur
                        </button>
                      </div>
                      <input
                        type="text"
                        required
                        value={appData.settings.academicYear}
                        onChange={e => setAppData(prev => ({
                          ...prev,
                          settings: { ...prev.settings, academicYear: e.target.value }
                        }))}
                        placeholder="Contoh: 2024/2025"
                        className="w-full p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500 bg-white"
                      />
                      <p className="text-[10px] text-slate-400 mt-1">Kalender Master dan Libur Nasional otomatis disesuaikan sesuai tahun ini</p>
                      
                      {/* Quick Academic Year Presets */}
                      <div className="flex flex-wrap items-center gap-1.5 mt-2">
                        <span className="text-[11px] text-slate-500 font-medium">Pilihan Cepat TP:</span>
                        {[
                          { label: '2024/2025', y1: 2024, y2: 2025 },
                          { label: '2025/2026', y1: 2025, y2: 2026 },
                          { label: '2026/2027', y1: 2026, y2: 2027 },
                          { label: '2027/2028', y1: 2027, y2: 2028 },
                        ].map(preset => (
                          <button
                            key={preset.label}
                            type="button"
                            onClick={() => {
                              setAppData(prev => ({
                                ...prev,
                                settings: {
                                  ...prev.settings,
                                  academicYear: preset.label,
                                  sem1Start: `${preset.y1}-07-15`,
                                  sem1End: `${preset.y1}-12-20`,
                                  sem2Start: `${preset.y2}-01-05`,
                                  sem2End: `${preset.y2}-06-20`,
                                }
                              }));
                              setCurrentCalendarDate(new Date(preset.y1, 6, 1));
                              showToast(`Tahun Pelajaran diset ke ${preset.label}`, 'info');
                            }}
                            className={`px-2 py-0.5 rounded text-[11px] font-semibold border transition cursor-pointer ${
                              appData.settings.academicYear === preset.label
                                ? 'bg-indigo-600 text-white border-indigo-600 shadow-xs'
                                : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-100'
                            }`}
                          >
                            {preset.label}
                          </button>
                        ))}
                      </div>
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">Mulai Semester 1</label>
                      <input
                        type="date"
                        required
                        value={appData.settings.sem1Start}
                        onChange={e => setAppData(prev => ({
                          ...prev,
                          settings: { ...prev.settings, sem1Start: e.target.value }
                        }))}
                        className="w-full p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500 bg-white"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">Akhir Semester 1</label>
                      <input
                        type="date"
                        required
                        value={appData.settings.sem1End}
                        onChange={e => setAppData(prev => ({
                          ...prev,
                          settings: { ...prev.settings, sem1End: e.target.value }
                        }))}
                        className="w-full p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500 bg-white"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">Mulai Semester 2</label>
                      <input
                        type="date"
                        required
                        value={appData.settings.sem2Start}
                        onChange={e => setAppData(prev => ({
                          ...prev,
                          settings: { ...prev.settings, sem2Start: e.target.value }
                        }))}
                        className="w-full p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500 bg-white"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">Akhir Semester 2</label>
                      <input
                        type="date"
                        required
                        value={appData.settings.sem2End}
                        onChange={e => setAppData(prev => ({
                          ...prev,
                          settings: { ...prev.settings, sem2End: e.target.value }
                        }))}
                        className="w-full p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500 bg-white"
                      />
                    </div>
                  </div>
                </div>

                {/* Developer Team & Signatures */}
                <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-4">
                  <h4 className="font-bold text-slate-800 text-sm flex items-center gap-2">
                    <i className="fa-solid fa-user-pen text-indigo-600"></i> Tim Pengembang & Tanda Tangan Dokumen
                  </h4>

                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 text-xs">
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">Kota / Tempat Penetapan</label>
                      <input
                        type="text"
                        required
                        value={appData.settings.cityName}
                        onChange={e => setAppData(prev => ({
                          ...prev,
                          settings: { ...prev.settings, cityName: e.target.value }
                        }))}
                        placeholder="Contoh: Bandung"
                        className="w-full p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500 bg-white"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">Jabatan / Tim Pengembang</label>
                      <input
                        type="text"
                        required
                        value={appData.settings.developerRole}
                        onChange={e => setAppData(prev => ({
                          ...prev,
                          settings: { ...prev.settings, developerRole: e.target.value }
                        }))}
                        placeholder="Contoh: Tim Pengembang Kurikulum"
                        className="w-full p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500 bg-white"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">Nama Pengembang / Ketua Panitia</label>
                      <input
                        type="text"
                        required
                        value={appData.settings.developerName}
                        onChange={e => setAppData(prev => ({
                          ...prev,
                          settings: { ...prev.settings, developerName: e.target.value }
                        }))}
                        placeholder="Contoh: Panitia Kalender Pendidikan"
                        className="w-full p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500 bg-white"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">NIP / NUPTK Pengembang (Opsional)</label>
                      <input
                        type="text"
                        value={appData.settings.developerNIP}
                        onChange={e => setAppData(prev => ({
                          ...prev,
                          settings: { ...prev.settings, developerNIP: e.target.value }
                        }))}
                        placeholder="Contoh: 19820315 200801 1 003"
                        className="w-full p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500 bg-white"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">Tanggal Penetapan Dokumen (Opsional)</label>
                      <input
                        type="date"
                        value={appData.settings.signDate}
                        onChange={e => setAppData(prev => ({
                          ...prev,
                          settings: { ...prev.settings, signDate: e.target.value }
                        }))}
                        className="w-full p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500 bg-white"
                      />
                      <p className="text-[10px] text-slate-400 mt-1">*Kosongkan untuk tanggal dinamis hari ini</p>
                    </div>
                  </div>
                </div>

                <div className="flex justify-end">
                  <button
                    type="submit"
                    className="bg-indigo-600 hover:bg-indigo-700 text-white px-6 py-2.5 rounded-lg text-xs font-bold shadow-md transition flex items-center gap-2 cursor-pointer"
                  >
                    <i className="fa-solid fa-floppy-disk"></i> Simpan Pengaturan
                  </button>
                </div>
              </form>

              <hr className="my-8 border-slate-200" />

              {/* Backup, Restore & Reset Section */}
              <div className="space-y-4 max-w-4xl">
                <h4 className="font-bold text-slate-800 text-sm flex items-center gap-2">
                  <i className="fa-solid fa-database text-indigo-600"></i> Pengelolaan Data Backup & Restore
                </h4>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                  {/* Export JSON */}
                  <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 flex flex-col justify-between">
                    <div>
                      <div className="font-bold text-xs text-slate-800">Cadangkan Data (JSON)</div>
                      <p className="text-[11px] text-slate-500 mt-1">Unduh seluruh file data kalender & agenda untuk disimpan aman.</p>
                    </div>
                    <button
                      onClick={backupJSON}
                      className="mt-4 bg-slate-800 hover:bg-slate-900 text-white text-xs font-semibold py-2 px-3 rounded-lg transition flex items-center justify-center gap-2"
                    >
                      <i className="fa-solid fa-download"></i> Unduh Backup JSON
                    </button>
                  </div>

                  {/* Restore JSON */}
                  <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 flex flex-col justify-between">
                    <div>
                      <div className="font-bold text-xs text-slate-800">Pulihkan Data (JSON)</div>
                      <p className="text-[11px] text-slate-500 mt-1">Unggah file backup JSON untuk mengembalikan data sebelumnya.</p>
                    </div>
                    <label className="mt-4 bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold py-2 px-3 rounded-lg transition cursor-pointer flex items-center justify-center gap-2 text-center">
                      <i className="fa-solid fa-upload"></i> Unggah File Backup
                      <input
                        type="file"
                        onChange={restoreJSON}
                        accept=".json"
                        className="hidden"
                      />
                    </label>
                  </div>

                  {/* Reset Demo Data */}
                  <div className="bg-rose-50/60 p-4 rounded-xl border border-rose-200 flex flex-col justify-between">
                    <div>
                      <div className="font-bold text-xs text-rose-900">Reset Data Demo</div>
                      <p className="text-[11px] text-rose-700 mt-1">Kembalikan data ke kondisi standar demo awal sekolah.</p>
                    </div>
                    <button
                      onClick={confirmResetDemo}
                      className="mt-4 bg-rose-600 hover:bg-rose-700 text-white text-xs font-semibold py-2 px-3 rounded-lg transition flex items-center justify-center gap-2"
                    >
                      <i className="fa-solid fa-rotate"></i> Reset Data
                    </button>
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>
      </main>

      {/* ==================== MODALS ==================== */}

      {/* 1. AGENDA MODAL (ADD / EDIT) */}
      {agendaModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-slate-200 shadow-2xl w-full max-w-lg overflow-hidden animate-in fade-in zoom-in-95 duration-150">
            <div className="bg-slate-900 text-white px-5 py-4 flex items-center justify-between">
              <h3 className="font-bold text-sm flex items-center gap-2">
                <i className="fa-solid fa-calendar-plus text-indigo-400"></i>
                {editingAgendaId ? 'Edit Agenda Kegiatan' : 'Tambah Agenda Kegiatan Baru'}
              </h3>
              <button
                onClick={() => setAgendaModalOpen(false)}
                className="text-slate-400 hover:text-white transition"
              >
                <i className="fa-solid fa-xmark text-lg"></i>
              </button>
            </div>

            <form onSubmit={handleSaveAgenda} className="p-5 space-y-4 text-xs">
              <div>
                <label className="block font-semibold text-slate-700 mb-1">Kategori Agenda *</label>
                <select
                  value={agendaForm.category}
                  onChange={e => setAgendaForm(f => ({ ...f, category: e.target.value as CategoryKey }))}
                  required
                  className="w-full p-2.5 rounded-lg border border-slate-300 bg-white font-medium focus:ring-2 focus:ring-indigo-500"
                >
                  <option value="holiday">Hari Libur Nasional (Merah)</option>
                  <option value="school">Agenda Sekolah (Biru)</option>
                  <option value="teacher">Agenda Guru (Ungu)</option>
                  <option value="important">Peristiwa Penting (Oranye)</option>
                </select>
              </div>

              <div>
                <label className="block font-semibold text-slate-700 mb-1">Nama Kegiatan / Agenda *</label>
                <input
                  type="text"
                  required
                  value={agendaForm.title}
                  onChange={e => setAgendaForm(f => ({ ...f, title: e.target.value }))}
                  placeholder="Contoh: Rapat Evaluasi Pembelajaran"
                  className="w-full p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Tanggal Mulai *</label>
                  <input
                    type="date"
                    required
                    value={agendaForm.startDate}
                    onChange={e => setAgendaForm(f => ({ ...f, startDate: e.target.value }))}
                    className="w-full p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Tanggal Selesai *</label>
                  <input
                    type="date"
                    required
                    value={agendaForm.endDate}
                    onChange={e => setAgendaForm(f => ({ ...f, endDate: e.target.value }))}
                    className="w-full p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500"
                  />
                </div>
              </div>

              {(agendaForm.category === 'school' || agendaForm.category === 'teacher') && (
                <div className="space-y-3">
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">Waktu (Jam)</label>
                      <input
                        type="text"
                        value={agendaForm.time}
                        onChange={e => setAgendaForm(f => ({ ...f, time: e.target.value }))}
                        placeholder="08:00 - 12:00 WIB"
                        className="w-full p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">Tempat / Lokasi</label>
                      <input
                        type="text"
                        value={agendaForm.location}
                        onChange={e => setAgendaForm(f => ({ ...f, location: e.target.value }))}
                        placeholder="Ruang Rapat / Aula"
                        className="w-full p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500"
                      />
                    </div>
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">Penanggung Jawab / Peserta</label>
                    <input
                      type="text"
                      value={agendaForm.pic}
                      onChange={e => setAgendaForm(f => ({ ...f, pic: e.target.value }))}
                      placeholder="Nama Guru / Panitia"
                      className="w-full p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                </div>
              )}

              {agendaForm.category === 'school' && (
                <div className="bg-amber-50 p-3 rounded-lg border border-amber-200">
                  <label className="flex items-center gap-2 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={agendaForm.isNonEffective}
                      onChange={e => setAgendaForm(f => ({ ...f, isNonEffective: e.target.checked }))}
                      className="rounded text-indigo-600 focus:ring-indigo-500 w-4 h-4"
                    />
                    <span className="font-semibold text-amber-900 text-xs">
                      Tandai sebagai Hari Tidak Efektif Belajar (Mengurangi HEB)
                    </span>
                  </label>
                </div>
              )}

              <div>
                <label className="block font-semibold text-slate-700 mb-1">Keterangan / Catatan Tambahan</label>
                <textarea
                  rows={2}
                  value={agendaForm.description}
                  onChange={e => setAgendaForm(f => ({ ...f, description: e.target.value }))}
                  placeholder="Catatan atau detail kegiatan..."
                  className="w-full p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500"
                />
              </div>

              <div className="pt-3 flex items-center justify-end gap-2 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setAgendaModalOpen(false)}
                  className="px-4 py-2 rounded-lg border border-slate-300 text-slate-700 font-semibold hover:bg-slate-100 transition"
                >
                  Batal
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white font-bold transition shadow-sm cursor-pointer"
                >
                  Simpan Data
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* 2. DATE DETAIL MODAL */}
      {dateDetailModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-slate-200 shadow-2xl w-full max-w-md overflow-hidden">
            <div className="bg-indigo-950 text-white px-5 py-4 flex items-center justify-between">
              <div>
                <h3 className="font-bold text-base text-white">Detail Agenda Tanggal</h3>
                <p className="text-xs text-indigo-200">{formatLongDateIndo(selectedDateStr)}</p>
              </div>
              <button
                onClick={() => setDateDetailModalOpen(false)}
                className="text-slate-300 hover:text-white"
              >
                <i className="fa-solid fa-xmark text-lg"></i>
              </button>
            </div>

            <div className="p-5 space-y-4">
              <div className="space-y-3 max-h-72 overflow-y-auto pr-1">
                {selectedDateEvents.length === 0 ? (
                  <p className="text-xs text-slate-400 italic text-center py-4">
                    Tidak ada agenda pada tanggal ini.
                  </p>
                ) : (
                  selectedDateEvents.map(ev => {
                    const cat = CATEGORIES[ev.category] || CATEGORIES.school;
                    return (
                      <div key={ev.id} className={`p-3 rounded-xl border text-xs ${cat.badgeClass} space-y-1`}>
                        <div className="flex items-center justify-between">
                          <span className="font-bold text-sm">{ev.title}</span>
                          <div className="flex items-center gap-1">
                            <button
                              onClick={() => {
                                setDateDetailModalOpen(false);
                                openEditAgendaModal(ev);
                              }}
                              className="text-slate-600 hover:text-indigo-600 p-1"
                              title="Edit"
                            >
                              <i className="fa-solid fa-pen"></i>
                            </button>
                            <button
                              onClick={() => {
                                setDateDetailModalOpen(false);
                                handleDeleteAgenda(ev.id);
                              }}
                              className="text-slate-600 hover:text-rose-600 p-1"
                              title="Hapus"
                            >
                              <i className="fa-solid fa-trash"></i>
                            </button>
                          </div>
                        </div>
                        {ev.time && (
                          <div className="text-[11px]">
                            <i className="fa-regular fa-clock mr-1"></i>{ev.time}
                          </div>
                        )}
                        {ev.location && (
                          <div className="text-[11px]">
                            <i className="fa-solid fa-location-dot mr-1"></i>{ev.location}
                          </div>
                        )}
                        {ev.description && (
                          <div className="text-[11px] opacity-80 pt-1 border-t border-black/10">
                            {ev.description}
                          </div>
                        )}
                      </div>
                    );
                  })
                )}
              </div>

              <div className="pt-3 border-t border-slate-100 flex items-center justify-between">
                <button
                  onClick={() => {
                    setDateDetailModalOpen(false);
                    openAddAgendaModal('school', selectedDateStr);
                  }}
                  className="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg text-xs font-bold w-full transition shadow-sm flex items-center justify-center gap-2 cursor-pointer"
                >
                  <i className="fa-solid fa-plus"></i> Tambah Kegiatan Pada Tanggal Ini
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* 3. PRINT & PDF MODAL */}
      {printModalOpen && (
        <div className="print-modal-overlay fixed inset-0 bg-slate-900/80 backdrop-blur-md z-50 flex items-center justify-center p-2 sm:p-4">
          <div className="print-modal-wrapper bg-white rounded-2xl border border-slate-200 shadow-2xl w-full max-w-5xl h-[92vh] flex flex-col overflow-hidden">
            {/* Header Bar */}
            <div className="print-header-controls bg-slate-900 text-white px-5 py-3.5 flex flex-wrap items-center justify-between gap-3 flex-shrink-0">
              <div className="flex items-center gap-3">
                <i className="fa-solid fa-print text-indigo-400 text-lg"></i>
                <div>
                  <h3 className="font-bold text-sm text-white">Pratinjau Cetak / Ekspor PDF</h3>
                  <p className="text-[11px] text-slate-400">Pengaturan tata letak dokumen resmi</p>
                </div>
              </div>

              {/* Controls */}
              <div className="flex items-center gap-2 text-xs">
                <select
                  value={printType}
                  onChange={e => setPrintType(e.target.value as any)}
                  className="bg-slate-800 text-white border border-slate-700 rounded-lg p-1.5 focus:outline-none"
                >
                  <option value="cal-sem1">Kalender Semester 1 (Ganjil)</option>
                  <option value="cal-sem2">Kalender Semester 2 (Genap)</option>
                  <option value="cal-annual">Kalender 1 Tahun Pelajaran</option>
                  <option value="rekap">Laporan Rekapitulasi Agenda</option>
                  <option value="heb">Laporan Hari Efektif Belajar</option>
                </select>

                <select
                  value={printPaperSize}
                  onChange={e => setPrintPaperSize(e.target.value as any)}
                  className="bg-slate-800 text-white border border-slate-700 rounded-lg p-1.5 focus:outline-none"
                >
                  <option value="a4">Kertas A4</option>
                  <option value="f4">Kertas F4 / Folio</option>
                </select>

                <select
                  value={printOrientation}
                  onChange={e => setPrintOrientation(e.target.value as any)}
                  className="bg-slate-800 text-white border border-slate-700 rounded-lg p-1.5 focus:outline-none"
                >
                  <option value="landscape">Landscape (Mendatar)</option>
                  <option value="portrait">Portrait (Tegak)</option>
                </select>
              </div>

              <div className="flex items-center gap-2">
                <button
                  onClick={downloadPDFDocument}
                  disabled={isGeneratingPDF}
                  className={`px-3.5 py-1.5 rounded-lg text-xs font-bold transition flex items-center gap-1.5 cursor-pointer ${
                    isGeneratingPDF
                      ? 'bg-emerald-800 text-emerald-200 cursor-not-allowed opacity-80'
                      : 'bg-emerald-600 hover:bg-emerald-700 text-white'
                  }`}
                  title="Unduh langsung sebagai berkas PDF (.pdf)"
                >
                  {isGeneratingPDF ? (
                    <>
                      <i className="fa-solid fa-circle-notch fa-spin"></i>
                      <span>Mengunduh...</span>
                    </>
                  ) : (
                    <>
                      <i className="fa-solid fa-file-pdf"></i>
                      <span>Unduh PDF</span>
                    </>
                  )}
                </button>
                <button
                  onClick={() => window.print()}
                  className="bg-indigo-600 hover:bg-indigo-700 text-white px-3.5 py-1.5 rounded-lg text-xs font-bold transition flex items-center gap-1.5 cursor-pointer"
                  title="Cetak langsung atau simpan sebagai PDF via dialog browser"
                >
                  <i className="fa-solid fa-print"></i>
                  <span>Cetak</span>
                </button>
                <button
                  onClick={() => setPrintModalOpen(false)}
                  className="text-slate-400 hover:text-white p-1"
                >
                  <i className="fa-solid fa-xmark text-xl"></i>
                </button>
              </div>
            </div>

            {/* Document Preview */}
            <div className="print-modal-body flex-1 bg-slate-200 p-4 overflow-y-auto flex justify-center">
              <div
                id="print-area"
                style={{ maxWidth: printOrientation === 'landscape' ? '1050px' : '800px' }}
                className="bg-white shadow-xl p-8 rounded-sm text-slate-900 w-full min-h-[700px] transition-all my-auto"
              >
                {/* Kop Surat Header */}
                <div className="border-b-4 border-double border-slate-900 pb-3 mb-5 flex items-center justify-between gap-4">
                  {appData.settings.logoLeft ? (
                    <div className="w-20 h-20 flex-shrink-0 flex items-center justify-center">
                      <img src={appData.settings.logoLeft} alt="Logo Kiri" className="max-w-full max-h-full object-contain" />
                    </div>
                  ) : (
                    <div className="w-20 h-20 flex-shrink-0"></div>
                  )}

                  <div className="text-center flex-1">
                    <h2 className="font-extrabold text-xl text-slate-900 tracking-wide uppercase leading-tight">
                      {appData.settings.schoolName}
                    </h2>
                    <p className="text-xs text-slate-700 mt-1">
                      {appData.settings.address} | NPSN: {appData.settings.schoolNPSN}
                    </p>
                    <div className="mt-2 text-xs font-bold text-slate-800 uppercase tracking-wider bg-slate-100 py-1 px-3 rounded border border-slate-300 inline-block">
                      KALENDER PENDIDIKAN SEKOLAH DASAR (SD) - TAHUN PELAJARAN {appData.settings.academicYear}
                    </div>
                  </div>

                  {appData.settings.logoRight ? (
                    <div className="w-20 h-20 flex-shrink-0 flex items-center justify-center">
                      <img src={appData.settings.logoRight} alt="Logo Kanan" className="max-w-full max-h-full object-contain" />
                    </div>
                  ) : (
                    <div className="w-20 h-20 flex-shrink-0"></div>
                  )}
                </div>

                {/* Print Content Body */}
                {(printType === 'cal-sem1' || printType === 'cal-sem2' || printType === 'cal-annual') && (
                  <div>
                    {(printType === 'cal-sem1' || printType === 'cal-annual') && (
                      <div className="mb-6">
                        <div className="font-bold text-sm text-slate-900 border-b border-slate-300 pb-1 mb-3">
                          TAMPILAN KALENDER SEMESTER 1 (GANJIL)
                        </div>
                        <div className="grid grid-cols-3 gap-3 mb-4">
                          {[6, 7, 8, 9, 10, 11].map(m => {
                            const y = tpYears.startYear;
                            return (
                              <div key={m} className="border border-slate-300 rounded p-1.5 text-[10px]">
                                <div className="font-bold text-center bg-slate-200 py-0.5 rounded text-slate-900 mb-1">
                                  {getMonthNameIndo(m)} {y}
                                </div>
                                <div className="grid grid-cols-7 text-center font-bold text-[9px] border-b border-slate-200 mb-0.5">
                                  <span>S</span><span>S</span><span>R</span><span>K</span><span>J</span><span className="text-rose-600">S</span><span className="text-rose-600">M</span>
                                </div>
                                <div className="grid grid-cols-7 text-center leading-tight gap-0.5">
                                  {Array.from({ length: (new Date(y, m, 1).getDay() + 6) % 7 }).map((_, i) => (
                                    <div key={`emp-${i}`}></div>
                                  ))}
                                  {Array.from({ length: new Date(y, m + 1, 0).getDate() }).map((_, i) => {
                                    const day = i + 1;
                                    const dateStr = `${y}-${String(m + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                                    const dayOfWeek = new Date(y, m, day).getDay();
                                    const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
                                    const dayEvents = appData.events.filter(ev => isDateInRange(dateStr, ev.startDate, ev.endDate));
                                    const hasHoliday = dayEvents.some(e => e.category === 'holiday');

                                    let bgStyle = '';
                                    if (hasHoliday || isWeekend) bgStyle = 'bg-rose-100 text-rose-800 font-bold';
                                    else if (dayEvents.length > 0) bgStyle = 'bg-sky-200 text-sky-900 font-bold';

                                    return (
                                      <div key={day} className={`py-0.5 rounded ${bgStyle}`}>
                                        {day}
                                      </div>
                                    );
                                  })}
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    )}

                    {printType === 'cal-annual' && <div className="page-break my-6"></div>}

                    {(printType === 'cal-sem2' || printType === 'cal-annual') && (
                      <div className="mb-6">
                        <div className="font-bold text-sm text-slate-900 border-b border-slate-300 pb-1 mb-3">
                          TAMPILAN KALENDER SEMESTER 2 (GENAP)
                        </div>
                        <div className="grid grid-cols-3 gap-3 mb-4">
                          {[0, 1, 2, 3, 4, 5].map(m => {
                            const y = tpYears.endYear;
                            return (
                              <div key={m} className="border border-slate-300 rounded p-1.5 text-[10px]">
                                <div className="font-bold text-center bg-slate-200 py-0.5 rounded text-slate-900 mb-1">
                                  {getMonthNameIndo(m)} {y}
                                </div>
                                <div className="grid grid-cols-7 text-center font-bold text-[9px] border-b border-slate-200 mb-0.5">
                                  <span>S</span><span>S</span><span>R</span><span>K</span><span>J</span><span className="text-rose-600">S</span><span className="text-rose-600">M</span>
                                </div>
                                <div className="grid grid-cols-7 text-center leading-tight gap-0.5">
                                  {Array.from({ length: (new Date(y, m, 1).getDay() + 6) % 7 }).map((_, i) => (
                                    <div key={`emp2-${i}`}></div>
                                  ))}
                                  {Array.from({ length: new Date(y, m + 1, 0).getDate() }).map((_, i) => {
                                    const day = i + 1;
                                    const dateStr = `${y}-${String(m + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                                    const dayOfWeek = new Date(y, m, day).getDay();
                                    const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
                                    const dayEvents = appData.events.filter(ev => isDateInRange(dateStr, ev.startDate, ev.endDate));
                                    const hasHoliday = dayEvents.some(e => e.category === 'holiday');

                                    let bgStyle = '';
                                    if (hasHoliday || isWeekend) bgStyle = 'bg-rose-100 text-rose-800 font-bold';
                                    else if (dayEvents.length > 0) bgStyle = 'bg-sky-200 text-sky-900 font-bold';

                                    return (
                                      <div key={day} className={`py-0.5 rounded ${bgStyle}`}>
                                        {day}
                                      </div>
                                    );
                                  })}
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    )}
                  </div>
                )}

                {printType === 'rekap' && (
                  <div>
                    <h3 className="font-bold text-sm text-slate-900 border-b border-slate-300 pb-1 mb-3">
                      REKAPITULASI AGENDA SEKOLAH
                    </h3>
                    <table className="w-full text-left border-collapse text-xs mb-4">
                      <thead>
                        <tr className="bg-slate-200 text-slate-900 uppercase font-bold border border-slate-300">
                          <th className="p-2 border border-slate-300">No</th>
                          <th className="p-2 border border-slate-300">Kategori</th>
                          <th className="p-2 border border-slate-300">Tanggal</th>
                          <th className="p-2 border border-slate-300">Nama Agenda</th>
                          <th className="p-2 border border-slate-300">Lokasi / PIC</th>
                        </tr>
                      </thead>
                      <tbody>
                        {appData.events.map((ev, idx) => {
                          const cat = CATEGORIES[ev.category] || CATEGORIES.school;
                          const dateRange = ev.startDate === ev.endDate ? formatDateIndo(ev.startDate) : `${formatDateIndo(ev.startDate)} - ${formatDateIndo(ev.endDate)}`;
                          return (
                            <tr key={ev.id} className="border border-slate-300">
                              <td className="p-2 border border-slate-300 text-center font-mono">{idx + 1}</td>
                              <td className="p-2 border border-slate-300 font-bold">{cat.name}</td>
                              <td className="p-2 border border-slate-300">{dateRange}</td>
                              <td className="p-2 border border-slate-300 font-semibold">{ev.title}</td>
                              <td className="p-2 border border-slate-300">{ev.location || '-'} / {ev.pic || '-'}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}

                {printType === 'heb' && (
                  <div>
                    <h3 className="font-bold text-sm text-slate-900 border-b border-slate-300 pb-1 mb-3">
                      LAPORAN ANALISIS HARI EFEKTIF BELAJAR (HEB)
                    </h3>
                    <div className="grid grid-cols-2 gap-4 mb-4">
                      <div>
                        <h4 className="font-bold text-xs bg-slate-200 p-1.5 text-center">Semester 1 (Ganjil)</h4>
                        <table className="w-full text-left border-collapse text-xs border border-slate-300">
                          <thead>
                            <tr className="bg-slate-100 font-bold border-b border-slate-300">
                              <th className="p-1.5">Bulan</th>
                              <th className="p-1.5 text-center">Hari Kerja</th>
                              <th className="p-1.5 text-center">Libur</th>
                              <th className="p-1.5 text-center font-bold bg-slate-200">HEB</th>
                            </tr>
                          </thead>
                          <tbody>
                            {sem1HEB.monthlyBreakdown.map(m => (
                              <tr key={m.monthKey} className="border-b border-slate-200">
                                <td className="p-1.5 font-medium">{getMonthNameIndo(m.month)} {m.year}</td>
                                <td className="p-1.5 text-center font-mono">{m.workDaysCount}</td>
                                <td className="p-1.5 text-center font-mono">{m.holidayCount + m.nonEffectiveSchoolDays}</td>
                                <td className="p-1.5 text-center font-bold font-mono bg-slate-50">{m.finalHEB}</td>
                              </tr>
                            ))}
                            <tr className="font-bold bg-slate-200">
                              <td className="p-1.5">Total HEB Sem 1</td>
                              <td colSpan={3} className="p-1.5 text-center font-mono text-sm">{sem1HEB.totalHEB} Hari</td>
                            </tr>
                          </tbody>
                        </table>
                      </div>

                      <div>
                        <h4 className="font-bold text-xs bg-slate-200 p-1.5 text-center">Semester 2 (Genap)</h4>
                        <table className="w-full text-left border-collapse text-xs border border-slate-300">
                          <thead>
                            <tr className="bg-slate-100 font-bold border-b border-slate-300">
                              <th className="p-1.5">Bulan</th>
                              <th className="p-1.5 text-center">Hari Kerja</th>
                              <th className="p-1.5 text-center">Libur</th>
                              <th className="p-1.5 text-center font-bold bg-slate-200">HEB</th>
                            </tr>
                          </thead>
                          <tbody>
                            {sem2HEB.monthlyBreakdown.map(m => (
                              <tr key={m.monthKey} className="border-b border-slate-200">
                                <td className="p-1.5 font-medium">{getMonthNameIndo(m.month)} {m.year}</td>
                                <td className="p-1.5 text-center font-mono">{m.workDaysCount}</td>
                                <td className="p-1.5 text-center font-mono">{m.holidayCount + m.nonEffectiveSchoolDays}</td>
                                <td className="p-1.5 text-center font-bold font-mono bg-slate-50">{m.finalHEB}</td>
                              </tr>
                            ))}
                            <tr className="font-bold bg-slate-200">
                              <td className="p-1.5">Total HEB Sem 2</td>
                              <td colSpan={3} className="p-1.5 text-center font-mono text-sm">{sem2HEB.totalHEB} Hari</td>
                            </tr>
                          </tbody>
                        </table>
                      </div>
                    </div>

                    <div className="bg-slate-900 text-white p-3 rounded flex justify-between items-center text-xs font-bold mb-4">
                      <span>TOTAL HARI EFEKTIF BELAJAR (HEB) 1 TAHUN PELAJARAN:</span>
                      <span className="text-base text-emerald-400 font-mono">
                        {sem1HEB.totalHEB + sem2HEB.totalHEB} HARI
                      </span>
                    </div>
                  </div>
                )}

                {/* Signature Block */}
                <div className="mt-10 pt-4 flex items-end justify-between text-xs text-slate-800">
                  <div>
                    <p>Mengetahui,</p>
                    <p className="font-bold">Kepala {appData.settings.schoolName}</p>
                    <div className="h-16"></div>
                    <p className="font-bold underline">{appData.settings.principalName}</p>
                    <p className="text-[11px] text-slate-600">NIP. {appData.settings.principalNIP || '-'}</p>
                  </div>
                  <div className="text-right">
                    <p>
                      Ditetapkan di {appData.settings.cityName || 'Bandung'},{' '}
                      {appData.settings.signDate
                        ? formatLongDateIndo(appData.settings.signDate)
                        : formatLongDateIndo(getLocalDateString())}
                    </p>
                    <p className="font-bold">{appData.settings.developerRole || 'Tim Pengembang Kurikulum'}</p>
                    <div className="h-16"></div>
                    <p className="font-bold underline">{appData.settings.developerName || 'Panitia Kalender Pendidikan'}</p>
                    <p className="text-[11px] text-slate-600">
                      {appData.settings.developerNIP ? `NIP. ${appData.settings.developerNIP}` : ''}
                    </p>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* 4. CONFIRM MODAL */}
      {confirmModal.open && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-slate-200 shadow-2xl w-full max-w-sm overflow-hidden p-5 text-center">
            <div className="w-12 h-12 rounded-full bg-rose-100 text-rose-600 flex items-center justify-center mx-auto mb-3 text-xl">
              <i className="fa-solid fa-triangle-exclamation"></i>
            </div>
            <h3 className="font-bold text-slate-800 text-base mb-1">{confirmModal.title}</h3>
            <p className="text-xs text-slate-500 mb-5">{confirmModal.message}</p>
            <div className="flex items-center justify-center gap-3">
              <button
                onClick={() => setConfirmModal(c => ({ ...c, open: false }))}
                className="w-1/2 py-2 rounded-lg border border-slate-300 text-slate-700 text-xs font-semibold hover:bg-slate-100 transition"
              >
                Batal
              </button>
              <button
                onClick={confirmModal.onConfirm}
                className="w-1/2 py-2 rounded-lg bg-rose-600 hover:bg-rose-700 text-white text-xs font-bold shadow-sm transition"
              >
                Ya, Lanjutkan
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 5. TOAST CONTAINER */}
      <div className="fixed bottom-5 right-5 z-50 flex flex-col gap-2 pointer-events-none">
        {toasts.map(t => (
          <div
            key={t.id}
            className={`px-4 py-2.5 rounded-xl shadow-xl text-xs font-semibold flex items-center gap-2 border border-white/20 transition-all pointer-events-auto ${
              t.type === 'success'
                ? 'bg-emerald-600 text-white'
                : t.type === 'error'
                ? 'bg-rose-600 text-white'
                : 'bg-slate-800 text-white'
            }`}
          >
            <i
              className={`fa-solid ${
                t.type === 'success'
                  ? 'fa-circle-check'
                  : t.type === 'error'
                  ? 'fa-circle-exclamation'
                  : 'fa-circle-info'
              }`}
            ></i>
            <span>{t.msg}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

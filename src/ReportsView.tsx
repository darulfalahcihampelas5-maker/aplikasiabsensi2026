import React, { useState, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx-js-style';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { 
  FileText,
  Filter,
  CheckCircle2,
  FileSpreadsheet,
  FileWarning,
  Loader2,
  FileDown,
  ClipboardList,
  UserCheck,
  ArrowRight,
  Calendar
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { format, parseISO } from 'date-fns';
import { id } from 'date-fns/locale';
import { kopSuratBase64 } from './kop-surat-b64';
import HomeroomReportView from './HomeroomReportView';
import { compareClasses, compareStudentsByClass } from './classSortUtils';
import type { Firestore } from 'firebase/firestore';
import type { Auth } from 'firebase/auth';

type Status = 'Hadir' | 'Sakit' | 'Izin' | 'Alpa' | 'Dispen' | '';

interface Student {
  id: string;
  nisn: string;
  name: string;
  class: string;
  userId?: string;
}

interface AttendanceSession {
  id: string;
  date: string;
  className: string;
  meetingNumber: number;
  records: Record<string, Status>;
  userId?: string;
}

interface ReportsViewProps {
  classList: string[];
  students: Student[];
  attendanceSessions: AttendanceSession[];
  profileData: {
    namaGuruMapel: string;
    namaKepalaSekolah: string;
    nipGuruMapel: string;
    nipKepalaSekolah: string;
    namaKurikulum?: string;
    nipKurikulum?: string;
    jabatanKurikulum?: string;
    namaKesiswaan?: string;
    nipKesiswaan?: string;
    jabatanKesiswaan?: string;
    namaGuruWali?: string;
    nipGuruWali?: string;
    jabatanGuruWali?: string;
    namaBK?: string;
    nipBK?: string;
    jabatanBK?: string;
    namaHumas?: string;
    nipHumas?: string;
    jabatanHumas?: string;
    semester: string;
    tahunPelajaran: string;
    mataPelajaran: string;
    role?: string;
    waliKelasClass?: string;
  };
  activeDb?: Firestore;
  activeAuth?: Auth;
  trackOp?: (type: 'read' | 'write', count?: number) => void;
  showToast?: (message: string, type: 'success' | 'info' | 'error') => void;
  initialFrame?: 'presensi' | 'wali_kelas';
  onNavigateToProfile?: () => void;
}

export type SignerRoleType = 'kepala_sekolah' | 'kurikulum' | 'kesiswaan' | 'humas' | 'guru_wali' | 'guru_bk' | 'none';

export default function ReportsView({ 
  classList, 
  students, 
  attendanceSessions, 
  profileData, 
  activeDb,
  activeAuth,
  trackOp,
  showToast,
  initialFrame = 'presensi',
  onNavigateToProfile
}: ReportsViewProps) {
  const isWaliKelas = profileData?.role === 'Wali Kelas';
  const effectiveClassList = useMemo(() => {
    if (isWaliKelas && profileData?.waliKelasClass) {
      return [profileData.waliKelasClass];
    }
    return classList;
  }, [isWaliKelas, profileData?.waliKelasClass, classList]);

  const [activeReportFrame, setActiveReportFrame] = useState<'presensi' | 'wali_kelas'>(initialFrame);
  const [selectedClass, setSelectedClass] = useState<string>(() => {
    if (isWaliKelas && profileData?.waliKelasClass) {
      return profileData.waliKelasClass;
    }
    return '';
  });

  useEffect(() => {
    if (isWaliKelas && profileData?.waliKelasClass) {
      setSelectedClass(profileData.waliKelasClass);
    }
  }, [isWaliKelas, profileData?.waliKelasClass]);
  const [reportType, setReportType] = useState<'daily' | 'monthly' | 'summary' | 'custom'>('summary');
  const [selectedMonth, setSelectedMonth] = useState<string>(format(new Date(), 'yyyy-MM'));
  const [selectedDailyDate, setSelectedDailyDate] = useState<string>(format(new Date(), 'yyyy-MM-dd'));
  const [startDate, setStartDate] = useState<string>(format(new Date(), 'yyyy-MM-01'));
  const [endDate, setEndDate] = useState<string>(format(new Date(), 'yyyy-MM-dd'));
  const [showExportSuccess, setShowExportSuccess] = useState<'excel' | 'pdf' | 'none' | 'no_data'>('none');
  const [showPdfOptionsModal, setShowPdfOptionsModal] = useState(false);
  const [includePdfStats, setIncludePdfStats] = useState(true);
  const [includePdfTopStudents, setIncludePdfTopStudents] = useState(true);
  const [leftSignerRole, setLeftSignerRole] = useState<SignerRoleType>('kepala_sekolah');
  const [midSignerRole, setMidSignerRole] = useState<SignerRoleType>('none');
  const [isExporting, setIsExporting] = useState(false);

  const getSignerDetails = (role: SignerRoleType) => {
    switch (role) {
      case 'kurikulum':
        return {
          title: profileData?.jabatanKurikulum || 'Wakasek Kurikulum',
          name: profileData?.namaKurikulum || '(________________________)',
          nip: profileData?.nipKurikulum ? `NIP. ${profileData.nipKurikulum}` : '',
          rawNip: profileData?.nipKurikulum || '',
          enabled: true
        };
      case 'kesiswaan':
        return {
          title: profileData?.jabatanKesiswaan || 'Wakasek Kesiswaan',
          name: profileData?.namaKesiswaan || '(________________________)',
          nip: profileData?.nipKesiswaan ? `NIP. ${profileData.nipKesiswaan}` : '',
          rawNip: profileData?.nipKesiswaan || '',
          enabled: true
        };
      case 'humas':
        return {
          title: profileData?.jabatanHumas || 'Wakasek Humas',
          name: profileData?.namaHumas || '(________________________)',
          nip: profileData?.nipHumas ? `NIP. ${profileData.nipHumas}` : '',
          rawNip: profileData?.nipHumas || '',
          enabled: true
        };
      case 'guru_wali':
        return {
          title: profileData?.jabatanGuruWali || 'Guru Wali',
          name: profileData?.namaGuruWali || '(________________________)',
          nip: profileData?.nipGuruWali ? `NIP. ${profileData.nipGuruWali}` : '',
          rawNip: profileData?.nipGuruWali || '',
          enabled: true
        };
      case 'guru_bk':
        return {
          title: profileData?.jabatanBK || 'Guru BK',
          name: profileData?.namaBK || '(________________________)',
          nip: profileData?.nipBK ? `NIP. ${profileData.nipBK}` : '',
          rawNip: profileData?.nipBK || '',
          enabled: true
        };
      case 'none':
        return {
          title: '',
          name: '',
          nip: '',
          rawNip: '',
          enabled: false
        };
      case 'kepala_sekolah':
      default:
        return {
          title: 'Kepala Sekolah',
          name: profileData?.namaKepalaSekolah || '(________________________)',
          nip: profileData?.nipKepalaSekolah ? `NIP. ${profileData.nipKepalaSekolah}` : '',
          rawNip: profileData?.nipKepalaSekolah || '',
          enabled: true
        };
    }
  };

  const getLeftSignerDetails = () => getSignerDetails(leftSignerRole);
  const getMidSignerDetails = () => getSignerDetails(midSignerRole);

  // Logic for computing report based on selectedClass, ReportType, and selectedMonth
  const classStudents = useMemo(() => {
    if (!selectedClass) return [];
    if (selectedClass === 'all') {
      return [...students].sort(compareStudentsByClass);
    }
    return students.filter(s => s.class === selectedClass).sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }));
  }, [students, selectedClass]);

  const classSessions = useMemo(() => {
    if (!selectedClass) return [];
    if (selectedClass === 'all') {
      return [...attendanceSessions].sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
    }
    return attendanceSessions.filter(s => s.className === selectedClass).sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
  }, [attendanceSessions, selectedClass]);

  const monthlySessions = useMemo(() => {
    return classSessions.filter(session => session.date.startsWith(selectedMonth));
  }, [classSessions, selectedMonth]);

  const customRangeSessions = useMemo(() => {
    if (!startDate && !endDate) return classSessions;
    return classSessions.filter(session => {
      if (startDate && endDate) {
        const start = startDate <= endDate ? startDate : endDate;
        const end = startDate <= endDate ? endDate : startDate;
        return session.date >= start && session.date <= end;
      }
      if (startDate) return session.date >= startDate;
      if (endDate) return session.date <= endDate;
      return true;
    });
  }, [classSessions, startDate, endDate]);

  const availableDailyDates = useMemo(() => {
    return Array.from(new Set(classSessions.map(s => s.date))).sort().reverse();
  }, [classSessions]);

  const classStats = useMemo(() => {
    if (!selectedClass || classStudents.length === 0) return null;

    let totalHadir = 0;
    let totalSakit = 0;
    let totalIzin = 0;
    let totalAlpa = 0;
    let totalDispen = 0;

    classStudents.forEach(student => {
      let hadir = 0, sakit = 0, izin = 0, alpa = 0, dispen = 0;

      if (reportType === 'summary') {
        const studentSessions = classSessions.filter(s => selectedClass === 'all' ? s.className === student.class : true);
        studentSessions.forEach(session => {
          const status = session.records[student.id];
          if (status === 'Hadir') hadir++;
          if (status === 'Sakit') sakit++;
          if (status === 'Izin') izin++;
          if (status === 'Alpa') alpa++;
          if (status === 'Dispen') dispen++;
        });
      } else if (reportType === 'monthly') {
        const studentMonthlySessions = monthlySessions.filter(s => selectedClass === 'all' ? s.className === student.class : true);
        studentMonthlySessions.forEach(session => {
          const status = session.records[student.id];
          if (status === 'Hadir') hadir++;
          if (status === 'Sakit') sakit++;
          if (status === 'Izin') izin++;
          if (status === 'Alpa') alpa++;
          if (status === 'Dispen') dispen++;
        });
      } else if (reportType === 'custom') {
        const studentRangeSessions = customRangeSessions.filter(s => selectedClass === 'all' ? s.className === student.class : true);
        studentRangeSessions.forEach(session => {
          const status = session.records[student.id];
          if (status === 'Hadir') hadir++;
          if (status === 'Sakit') sakit++;
          if (status === 'Izin') izin++;
          if (status === 'Alpa') alpa++;
          if (status === 'Dispen') dispen++;
        });
      } else {
        // daily
        const studentDailySessions = classSessions.filter(s => 
          s.date === selectedDailyDate && (selectedClass === 'all' ? s.className === student.class : true)
        );
        const targetSession = studentDailySessions[studentDailySessions.length - 1];
        if (targetSession) {
          const status = targetSession.records[student.id];
          if (status === 'Hadir') hadir++;
          if (status === 'Sakit') sakit++;
          if (status === 'Izin') izin++;
          if (status === 'Alpa') alpa++;
          if (status === 'Dispen') dispen++;
        }
      }

      totalHadir += hadir;
      totalSakit += sakit;
      totalIzin += izin;
      totalAlpa += alpa;
      totalDispen += dispen;
    });

    const grandTotal = totalHadir + totalSakit + totalIzin + totalAlpa + totalDispen;
    const effectiveHadir = totalHadir + totalDispen;

    return {
      hadir: totalHadir,
      sakit: totalSakit,
      izin: totalIzin,
      alpa: totalAlpa,
      dispen: totalDispen,
      grandTotal,
      persentaseHadir: grandTotal > 0 ? (effectiveHadir / grandTotal) * 100 : 0,
      persentaseSakit: grandTotal > 0 ? (totalSakit / grandTotal) * 100 : 0,
      persentaseIzin: grandTotal > 0 ? (totalIzin / grandTotal) * 100 : 0,
      persentaseAlpa: grandTotal > 0 ? (totalAlpa / grandTotal) * 100 : 0,
      persentaseDispen: grandTotal > 0 ? (totalDispen / grandTotal) * 100 : 0,
    };
  }, [selectedClass, reportType, classStudents, classSessions, monthlySessions, customRangeSessions, selectedDailyDate]);

  React.useEffect(() => {
    if (reportType === 'daily' && availableDailyDates.length > 0 && !availableDailyDates.includes(selectedDailyDate)) {
      setSelectedDailyDate(availableDailyDates[0]);
    }
  }, [reportType, availableDailyDates, selectedDailyDate]);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const autoFitColumns = (data: any[]) => {
    if (data.length === 0) return [];
    const keys = Object.keys(data[0]);
    return keys.map(key => {
      let max = key.length;
      data.forEach(row => {
        const val = row[key];
        if (val !== undefined && val !== null) {
          max = Math.max(max, val.toString().length);
        }
      });
      return { wch: max + 2 };
    });
  };

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const applyStyles = (worksheet: any) => {
    if (!worksheet['!ref']) return;
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cell_address = {c: C, r: R};
        const cell_ref = XLSX.utils.encode_cell(cell_address);
        if (!worksheet[cell_ref]) continue;

        const cell = worksheet[cell_ref];
        if (!cell.s) cell.s = {};
        
        const headerCell = worksheet[XLSX.utils.encode_cell({c: C, r: 0})];
        const headerText = headerCell ? headerCell.v : '';

        // Determine alignment
        let hAlign = 'center'; 
        if (headerText === 'Nama Lengkap Siswa' || headerText === 'Nama Siswa') {
          hAlign = 'left';
        }

        cell.s.alignment = { horizontal: hAlign, vertical: 'center' };
        
        // Header row styling
        if (R === 0) {
          cell.s.font = { bold: true };
          cell.s.alignment = { horizontal: 'center', vertical: 'center' };
        }
      }
    }
  };

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const appendSignatures = (worksheet: any, data: any[]) => {
    if (data.length === 0) return;
    const keys = Object.keys(data[0]);
    const numCols = keys.length;
    
    const leftSigner = getLeftSignerDetails();
    const midSigner = getMidSignerDetails();
    const is3Signers = leftSigner.enabled && midSigner.enabled;
    
    let leftCol = 1;
    let midCol = Math.floor(numCols / 2);
    let rightCol = numCols > 3 ? numCols - 2 : numCols - 1;
    
    if (is3Signers) {
      if (numCols >= 6) {
        leftCol = 1;
        midCol = Math.floor(numCols / 2);
        rightCol = numCols - 2;
      } else if (numCols >= 3) {
        leftCol = 0;
        midCol = 1;
        rightCol = 2;
      } else {
        leftCol = 0;
        midCol = 1;
        rightCol = 2;
      }
    } else {
      if (rightCol <= leftCol) rightCol = leftCol + 1;
    }
    
    const createRow = () => new Array(Math.max(numCols, rightCol + 1)).fill('');
    
    const row1 = createRow();
    if (leftSigner.enabled) {
      row1[leftCol] = 'Mengetahui,';
    }
    if (midSigner.enabled) {
      row1[midCol] = 'Mengetahui,';
    }
    row1[rightCol] = `Cililin, ${format(new Date(), 'dd MMMM yyyy', {locale: id})}`;
    
    const row2 = createRow();
    if (leftSigner.enabled) {
      row2[leftCol] = leftSigner.title;
    }
    if (midSigner.enabled) {
      row2[midCol] = midSigner.title;
    }
    row2[rightCol] = profileData?.role === 'Wali Kelas'
      ? (profileData?.waliKelasClass ? `Wali Kelas ${profileData.waliKelasClass}` : 'Wali Kelas')
      : ((profileData?.mataPelajaran && profileData.mataPelajaran.trim()) 
          ? `Guru Mata Pelajaran ${profileData.mataPelajaran.trim()}` 
          : 'Guru Mata Pelajaran');
    
    const row3 = createRow();
    const row4 = createRow();
    const row5 = createRow();
    
    const row6 = createRow();
    if (leftSigner.enabled) {
      row6[leftCol] = leftSigner.name;
    }
    if (midSigner.enabled) {
      row6[midCol] = midSigner.name;
    }
    row6[rightCol] = profileData.namaGuruMapel || '(________________________)';

    const row7 = createRow();
    if (leftSigner.enabled) {
      row7[leftCol] = leftSigner.nip;
    }
    if (midSigner.enabled) {
      row7[midCol] = midSigner.nip;
    }
    row7[rightCol] = profileData.nipGuruMapel ? `NIP. ${profileData.nipGuruMapel}` : '';

    XLSX.utils.sheet_add_aoa(worksheet, [[], [], row1, row2, row3, row4, row5, row6, row7], { origin: -1 });

    // Center alignment for signatures
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    for (let r = range.e.r - 8; r <= range.e.r; r++) {
      for (let c = 0; c <= range.e.c; c++) {
        const cell_ref = XLSX.utils.encode_cell({c, r});
        if (worksheet[cell_ref]) {
           if (!worksheet[cell_ref].s) worksheet[cell_ref].s = {};
           if (c === leftCol || c === rightCol || (is3Signers && c === midCol)) {
             worksheet[cell_ref].s.alignment = { horizontal: 'center', vertical: 'center' };
           }
        }
      }
    }
  };

  const getExportData = () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let data: any[] = [];
    let fileName = '';
    let sheetName = '';

    if (reportType === 'summary') {
      data = classStudents.map((student, i) => {
        const studentSessions = classSessions.filter(s => s.className === student.class);
        let hadir = 0, sakit = 0, izin = 0, alpa = 0, dispen = 0;

        studentSessions.forEach(session => {
          const status = session.records[student.id];
          if (status === 'Hadir') hadir++;
          if (status === 'Sakit') sakit++;
          if (status === 'Izin') izin++;
          if (status === 'Alpa') alpa++;
          if (status === 'Dispen') dispen++;
        });

        const totalRecorded = hadir + sakit + izin + alpa + dispen;
        const presentPercentage = totalRecorded > 0 ? Math.round(((hadir + dispen) / totalRecorded) * 100) : 0;

        const row: Record<string, string | number | boolean> = {
          'No': i + 1,
          'NIS': student.nisn,
          'Nama Lengkap Siswa': student.name
        };
        
        if (selectedClass === 'all') {
          row['Kelas'] = student.class;
        }

        row['Hadir'] = hadir;
        row['Sakit'] = sakit;
        row['Izin'] = izin;
        row['Alpa'] = alpa;
        row['Dispen'] = dispen;
        row['Presentase %'] = presentPercentage;

        return row;
      });
      fileName = `Rekap_Total_${selectedClass === 'all' ? 'Semua_Kelas' : selectedClass}`;
      sheetName = 'Rekap Total';
    } else if (reportType === 'monthly') {
      data = classStudents.map((student, i) => {
        const row: Record<string, string | number | boolean> = {
          'No': i + 1,
          'NIS': student.nisn,
          'Nama Lengkap Siswa': student.name
        };
        
        if (selectedClass === 'all') {
          row['Kelas'] = student.class;
        }

        let hadir = 0, sakit = 0, izin = 0, alpa = 0, dispen = 0;
        const studentMonthlySessions = monthlySessions.filter(s => s.className === student.class);

        studentMonthlySessions.forEach(session => {
          const status = session.records[student.id];
          if (status === 'Hadir') hadir++;
          if (status === 'Sakit') sakit++;
          if (status === 'Izin') izin++;
          if (status === 'Alpa') alpa++;
          if (status === 'Dispen') dispen++;
        });

        row['Hadir'] = hadir;
        row['Sakit'] = sakit;
        row['Izin'] = izin;
        row['Alpa'] = alpa;
        row['Dispen'] = dispen;
        
        const totalRecordedMonthly = hadir + sakit + izin + alpa + dispen;
        const presentPercentage = totalRecordedMonthly > 0 ? Math.round(((hadir + dispen) / totalRecordedMonthly) * 100) : 0;
        row['Presentase %'] = presentPercentage;

        return row;
      });
      fileName = `Rekap_Bulan_${selectedMonth}_${selectedClass === 'all' ? 'Semua_Kelas' : selectedClass}`;
      sheetName = `Bulan_${selectedMonth}`;
    } else if (reportType === 'custom') {
      data = classStudents.map((student, i) => {
        const row: Record<string, string | number | boolean> = {
          'No': i + 1,
          'NIS': student.nisn,
          'Nama Lengkap Siswa': student.name
        };
        
        if (selectedClass === 'all') {
          row['Kelas'] = student.class;
        }

        let hadir = 0, sakit = 0, izin = 0, alpa = 0, dispen = 0;
        const studentRangeSessions = customRangeSessions.filter(s => s.className === student.class);

        studentRangeSessions.forEach(session => {
          const status = session.records[student.id];
          if (status === 'Hadir') hadir++;
          if (status === 'Sakit') sakit++;
          if (status === 'Izin') izin++;
          if (status === 'Alpa') alpa++;
          if (status === 'Dispen') dispen++;
        });

        row['Hadir'] = hadir;
        row['Sakit'] = sakit;
        row['Izin'] = izin;
        row['Alpa'] = alpa;
        row['Dispen'] = dispen;
        
        const totalRecordedRange = hadir + sakit + izin + alpa + dispen;
        const presentPercentage = totalRecordedRange > 0 ? Math.round(((hadir + dispen) / totalRecordedRange) * 100) : 0;
        row['Presentase %'] = presentPercentage;

        return row;
      });
      fileName = `Rekap_${startDate}_sd_${endDate}_${selectedClass === 'all' ? 'Semua_Kelas' : selectedClass}`;
      sheetName = `Rekap_Rentang`.substring(0, 31);
    } else {
      data = classStudents.map((student, i) => {
        const studentDailySessions = classSessions.filter(
          s => s.date === selectedDailyDate && (selectedClass === 'all' ? s.className === student.class : true)
        );
        const targetSession = studentDailySessions[studentDailySessions.length - 1];

        const studentAllSessions = classSessions.filter(
          s => selectedClass === 'all' ? s.className === student.class : true
        );
        let hadir = 0, sakit = 0, izin = 0, alpa = 0, dispen = 0;
        studentAllSessions.forEach(session => {
          const st = session.records[student.id];
          if (st === 'Hadir') hadir++;
          if (st === 'Sakit') sakit++;
          if (st === 'Izin') izin++;
          if (st === 'Alpa') alpa++;
          if (st === 'Dispen') dispen++;
        });
        const totalRecordedDaily = hadir + sakit + izin + alpa + dispen;
        const presentPercentage = totalRecordedDaily > 0 ? Math.round(((hadir + dispen) / totalRecordedDaily) * 100) : 0;

        const row: Record<string, string | number | boolean> = {
          'No': i + 1,
          'NIS': student.nisn,
          'Nama Lengkap Siswa': student.name
        };

        if (selectedClass === 'all') {
          row['Kelas'] = student.class;
        }

        row['Status'] = targetSession ? (targetSession.records[student.id] || '-') : '-';
        row['Presentase %'] = presentPercentage;

        return row;
      });
      fileName = `Rekap_Harian_${selectedDailyDate}_${selectedClass === 'all' ? 'Semua_Kelas' : selectedClass}`;
      sheetName = `Harian_${selectedDailyDate}`;
    }
    
    return { data, fileName, sheetName };
  };

  const exportToExcel = () => {
    if (!selectedClass) return;
    const { data, fileName, sheetName } = getExportData();
    if (!data || data.length === 0) {
      setShowExportSuccess('no_data');
      return;
    }

    setIsExporting(true);
    setTimeout(() => {
      try {
        const worksheet = XLSX.utils.json_to_sheet(data);
        applyStyles(worksheet);
        worksheet['!cols'] = autoFitColumns(data);
        appendSignatures(worksheet, data);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        XLSX.writeFile(workbook, `${fileName}.xlsx`);
        setShowExportSuccess('excel');
      } catch (err) {
        console.error(err);
      } finally {
        setIsExporting(false);
      }
    }, 800);
  };

  const exportToPdf = () => {
    if (!selectedClass) return;
    const { data, fileName } = getExportData();
    if (!data || data.length === 0) {
      setShowExportSuccess('no_data');
      return;
    }

    setIsExporting(true);
    setTimeout(() => {
      try {
        const isLandscape = data[0] && Object.keys(data[0]).length > 10;
        const doc = new jsPDF({ orientation: isLandscape ? 'landscape' : 'portrait' });
        const pageWidth = doc.internal.pageSize.getWidth();
        const pageHeight = doc.internal.pageSize.getHeight();
        
        // PDF Header - Kop Surat Formal dengan Gambar Presisi & Simetris
        // Posisi tepat 2 spasi dari atas kertas (2 x 12pt line spacing ≈ 8.5 mm)
        const imgX = 14;
        const imgY = 8.5; // Tepat 2 spasi dari tepi atas kertas
        const imgWidth = pageWidth - 28;
        const imgHeight = isLandscape ? 52 : (imgWidth * (341 / 1450));
        
        try {
          doc.addImage(kopSuratBase64, 'PNG', imgX, imgY, imgWidth, imgHeight);
        } catch (e) {
          console.error("Failed to add custom header image", e);
        }
        
        const startY = imgY + imgHeight + 8; // margin top (2 spasi) + image + spacing

        doc.setFontSize(14);
        doc.setFont("helvetica", "bold");
        doc.text('LAPORAN REKAPITULASI KEHADIRAN SISWA', pageWidth / 2, startY, { align: 'center' });
        
        doc.setFontSize(11);
        doc.setFont("helvetica", "normal");
        // Two-Column Symmetrical Header Layout (Row 1, Row 2, Row 3)
        const isWaliKelas = profileData.role === 'Wali Kelas';
        const labelLeftX = 14;
        const valLeftX = 48;

        const labelRightX = pageWidth - 85;
        const valRightX = pageWidth - 50;

        const row1Y = startY + 10;
        const row2Y = startY + 16;
        const row3Y = startY + 22;

        // Row 1: Periode (Left) | Kelas (Right)
        if (reportType === 'monthly') {
          doc.text(`Periode Bulan`, labelLeftX, row1Y);
          let displayMonth = selectedMonth;
          try {
            if (selectedMonth) {
              const [year, month] = selectedMonth.split('-');
              const dateObj = new Date(parseInt(year), parseInt(month) - 1, 1);
              displayMonth = format(dateObj, 'MMMM yyyy', { locale: id });
            }
          } catch (e) {
            console.error("Error formatting month:", e);
          }
          doc.text(`: ${displayMonth}`, valLeftX, row1Y);
        } else if (reportType === 'custom') {
          doc.text(`Periode Tanggal`, labelLeftX, row1Y);
          let startText = startDate || '-';
          let endText = endDate || '-';
          try {
            if (startDate) startText = format(parseISO(startDate), 'dd MMMM yyyy', { locale: id });
            if (endDate) endText = format(parseISO(endDate), 'dd MMMM yyyy', { locale: id });
          } catch (e) {
            console.error("Error formatting custom dates:", e);
          }
          doc.text(`: ${startText} s/d ${endText}`, valLeftX, row1Y);
        } else if (reportType === 'daily') {
          doc.text(`Periode Tanggal`, labelLeftX, row1Y);
          let displayDate = selectedDailyDate || '-';
          try {
            if (selectedDailyDate) {
              displayDate = format(parseISO(selectedDailyDate), 'dd MMMM yyyy', { locale: id });
            }
          } catch (e) {
            console.error("Error formatting date:", e);
          }
          doc.text(`: ${displayDate}`, valLeftX, row1Y);
        } else {
          doc.text(`Periode`, labelLeftX, row1Y);
          let summaryDateText = 'Semua Sesi (Total)';
          if (classSessions.length > 0) {
            try {
              const startDate = format(parseISO(classSessions[0].date), 'dd MMMM yyyy', {locale: id});
              const endDate = format(parseISO(classSessions[classSessions.length - 1].date), 'dd MMMM yyyy', {locale: id});
              summaryDateText = `${startDate} - ${endDate}`;
            } catch (e) {
              console.error("Error formatting summary dates:", e);
            }
          }
          doc.text(`: ${summaryDateText}`, valLeftX, row1Y);
        }

        doc.text(`Kelas`, labelRightX, row1Y);
        doc.text(`: ${selectedClass === 'all' ? 'Semua Kelas' : selectedClass}`, valRightX, row1Y);

        // Row 2: Tahun Pelajaran (Left) | Semester (Right)
        doc.text(`Tahun Pelajaran`, labelLeftX, row2Y);
        doc.text(`: ${profileData.tahunPelajaran || '-'}`, valLeftX, row2Y);

        doc.text(`Semester`, labelRightX, row2Y);
        doc.text(`: ${profileData.semester || '-'}`, valRightX, row2Y);

        // Row 3: Mata Pelajaran / Total Pertemuan
        const totalPertemuan = reportType === 'summary'
          ? (selectedClass === 'all' ? classSessions.length : classSessions.filter(s => s.className === selectedClass).length)
          : reportType === 'monthly'
          ? (selectedClass === 'all' ? monthlySessions.length : monthlySessions.filter(s => s.className === selectedClass).length)
          : reportType === 'custom'
          ? (selectedClass === 'all' ? customRangeSessions.length : customRangeSessions.filter(s => s.className === selectedClass).length)
          : (selectedClass === 'all' ? classSessions.filter(s => s.date === selectedDailyDate).length : classSessions.filter(s => s.date === selectedDailyDate && s.className === selectedClass).length);

        let tableStartY = row3Y + 8;
        if (!isWaliKelas) {
          doc.text(`Mata Pelajaran`, labelLeftX, row3Y);
          doc.text(`: ${profileData.mataPelajaran || '-'}`, valLeftX, row3Y);

          doc.text(`Total Pertemuan`, labelRightX, row3Y);
          doc.text(`: ${totalPertemuan} Pertemuan`, valRightX, row3Y);
          tableStartY = row3Y + 8;
        } else {
          doc.text(`Total Pertemuan`, labelLeftX, row3Y);
          doc.text(`: ${totalPertemuan} Pertemuan`, valLeftX, row3Y);
          tableStartY = row3Y + 8;
        }
        
        const headers = Object.keys(data[0]);
        const body = data.map(row => Object.values(row).map(val => val !== undefined && val !== null ? val.toString() : ''));

        const autoTableOptions = {
          startY: tableStartY,
          head: [headers],
          body: body,
          theme: 'grid' as const,
          headStyles: { fillColor: [230, 230, 230], textColor: 0, fontStyle: 'bold', halign: 'center', lineWidth: 0.2, lineColor: [0, 0, 0] },
          styles: { fontSize: 8, cellPadding: 2, lineWidth: 0.2, lineColor: [0, 0, 0], textColor: 0 },
          columnStyles: {
            /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
            2: { halign: 'left' } as any // Assuming index 2 is 'Nama Lengkap Siswa' always
          },
          /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
          didParseCell: function(cellData: any) {
            if (cellData.section === 'body' && cellData.column.index !== 2) {
              cellData.cell.styles.halign = 'center';
            }
          }
        };

        // Ensure we call autoTable safely across all bundler/Vite environments
        if (typeof autoTable === 'function') {
          /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
          autoTable(doc, autoTableOptions as any);
        /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
        } else if (typeof (doc as any).autoTable === 'function') {
          /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
          (doc as any).autoTable(autoTableOptions as any);
        } else {
          throw new Error("jsPDF AutoTable plugin is not loaded correctly. Please check integration.");
        }

        /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
        const finalY = (doc as any).lastAutoTable?.finalY || 40;
        
        let statsY = finalY + 12;
        if (classStats && includePdfStats) {
          if (statsY + 25 + 40 > pageHeight) {
            doc.addPage();
            statsY = 20;
          }

          doc.setFontSize(10);
          doc.setFont("helvetica", "bold");
          doc.text('STATISTIK AKUMULASI KELAS', 14, statsY);

          const statsHeaders = ['Kehadiran (Hadir)', 'Sakit', 'Izin', 'Alpa', 'Dispen', 'Total Sesi Presensi'];
          const statsBody = [
            [
              `${classStats.hadir} kali (${classStats.persentaseHadir.toFixed(2)}%)`,
              `${classStats.sakit} kali (${classStats.persentaseSakit.toFixed(2)}%)`,
              `${classStats.izin} kali (${classStats.persentaseIzin.toFixed(2)}%)`,
              `${classStats.alpa} kali (${classStats.persentaseAlpa.toFixed(2)}%)`,
              `${classStats.dispen} kali (${classStats.persentaseDispen.toFixed(2)}%)`,
              `${classStats.grandTotal} kali`
            ]
          ];

          const statsTableOptions = {
            startY: statsY + 4,
            head: [statsHeaders],
            body: statsBody,
            theme: 'grid' as const,
            headStyles: { fillColor: [230, 230, 230], textColor: 0, fontStyle: 'bold', halign: 'center', lineWidth: 0.2, lineColor: [0, 0, 0] },
            styles: { fontSize: 8, cellPadding: 3, lineWidth: 0.2, lineColor: [0, 0, 0], textColor: 0, halign: 'center' }
          };

          if (typeof autoTable === 'function') {
            /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
            autoTable(doc, statsTableOptions as any);
          /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
          } else if (typeof (doc as any).autoTable === 'function') {
            /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
            (doc as any).autoTable(statsTableOptions as any);
          }

          /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
          statsY = (doc as any).lastAutoTable?.finalY || (statsY + 20);
        }

        // Section: Daftar Siswa Terbanyak (Alpa, Sakit, Izin)
        let topStudentsY = statsY + 8;
        if (includePdfTopStudents) {
          const studentStats = classStudents.map(student => {
          let sakit = 0, izin = 0, alpa = 0;
          if (reportType === 'summary') {
            const studentSessions = classSessions.filter(s => selectedClass === 'all' ? s.className === student.class : true);
            studentSessions.forEach(session => {
              const status = session.records[student.id];
              if (status === 'Sakit') sakit++;
              if (status === 'Izin') izin++;
              if (status === 'Alpa') alpa++;
            });
          } else if (reportType === 'monthly') {
            const studentMonthlySessions = monthlySessions.filter(s => selectedClass === 'all' ? s.className === student.class : true);
            studentMonthlySessions.forEach(session => {
              const status = session.records[student.id];
              if (status === 'Sakit') sakit++;
              if (status === 'Izin') izin++;
              if (status === 'Alpa') alpa++;
            });
          } else if (reportType === 'custom') {
            const studentRangeSessions = customRangeSessions.filter(s => selectedClass === 'all' ? s.className === student.class : true);
            studentRangeSessions.forEach(session => {
              const status = session.records[student.id];
              if (status === 'Sakit') sakit++;
              if (status === 'Izin') izin++;
              if (status === 'Alpa') alpa++;
            });
          } else {
            // daily
            const studentDailySessions = classSessions.filter(s => 
              s.date === selectedDailyDate && (selectedClass === 'all' ? s.className === student.class : true)
            );
            const targetSession = studentDailySessions[studentDailySessions.length - 1];
            if (targetSession) {
              const status = targetSession.records[student.id];
              if (status === 'Sakit') sakit++;
              if (status === 'Izin') izin++;
              if (status === 'Alpa') alpa++;
            }
          }
          return { student, sakit, izin, alpa };
        });

        const combinedStudentStats = studentStats
          .map(s => ({
            name: s.student.name,
            className: s.student.class,
            sakit: s.sakit,
            izin: s.izin,
            alpa: s.alpa,
            totalKeterangan: s.sakit + s.izin + s.alpa
          }))
          .filter(s => s.totalKeterangan > 0)
          .sort((a, b) => b.totalKeterangan - a.totalKeterangan);

        const topTableBody: string[][] = [];

        if (combinedStudentStats.length > 0) {
          combinedStudentStats.forEach(s => {
            const categoryDetail = `Sakit: ${s.sakit}, Izin: ${s.izin}, Alpa: ${s.alpa}`;
            topTableBody.push([s.name, s.className, categoryDetail, `${s.totalKeterangan} Kali`]);
          });
        } else {
          topTableBody.push(['Nihil (Tidak ada siswa dengan keterangan)', '-', '-', '0 Kali']);
        }

        if (topStudentsY + 30 + 40 > pageHeight) {
          doc.addPage();
          topStudentsY = 20;
        }

        doc.setFontSize(10);
        doc.setFont("helvetica", "bold");
        doc.text('DAFTAR SISWA TERBANYAK (ALPA, SAKIT, IZIN)', 14, topStudentsY);

        const topTableOptions = {
          startY: topStudentsY + 4,
          head: [['Nama Lengkap Siswa', 'Kelas', 'Kategori Presensi', 'Jumlah Keterangan']],
          body: topTableBody,
          theme: 'grid' as const,
          headStyles: { fillColor: [230, 230, 230], textColor: 0, fontStyle: 'bold', halign: 'center', lineWidth: 0.2, lineColor: [0, 0, 0] },
          styles: { fontSize: 8, cellPadding: 2.5, lineWidth: 0.2, lineColor: [0, 0, 0], textColor: 0 },
          columnStyles: {
            /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
            0: { halign: 'left' } as any,
            /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
            1: { halign: 'center' } as any,
            /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
            2: { halign: 'left' } as any,
            /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
            3: { halign: 'center', fontStyle: 'bold' } as any
          }
        };

        if (typeof autoTable === 'function') {
          /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
          autoTable(doc, topTableOptions as any);
        /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
        } else if (typeof (doc as any).autoTable === 'function') {
          /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
          (doc as any).autoTable(topTableOptions as any);
        }

        /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
        const lastTopY = (doc as any).lastAutoTable?.finalY || (topStudentsY + 25);
        statsY = lastTopY;
        }

        let sigY = statsY + 15;
        if (sigY + 40 > pageHeight) {
           doc.addPage();
           sigY = 20;
        }

        doc.setFontSize(10);
        
        // Signatures
        let dateStr = '';
        try {
          dateStr = `Cililin, ${format(new Date(), 'dd MMMM yyyy', { locale: id })}`;
        } catch {
          // Safe manual fallback formatting if format or locale throws
          const d = new Date();
          const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
          dateStr = `Cililin, ${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;
        }

        const leftSigner = getLeftSignerDetails();
        const midSigner = getMidSignerDetails();
        const is3Signers = leftSigner.enabled && midSigner.enabled;

        const teacherName = profileData?.namaGuruMapel || '(________________________)';
        const teacherNIP = profileData?.nipGuruMapel ? `NIP. ${profileData.nipGuruMapel}` : '';

        const teacherLabel = profileData?.role === 'Wali Kelas'
          ? (profileData?.waliKelasClass ? `Wali Kelas ${profileData.waliKelasClass}` : 'Wali Kelas')
          : ((profileData?.mataPelajaran && profileData.mataPelajaran.trim())
              ? `Guru Mata Pelajaran ${profileData.mataPelajaran.trim()}`
              : 'Guru Mata Pelajaran');

        doc.setFont("helvetica", "normal");
        
        if (is3Signers) {
          // Format 3 Tanda Tangan: Kiri (Mengetahui 1), Tengah (Mengetahui 2), Kanan (Wali Kelas / Guru Mapel)
          const leftX = 14;
          const midX = (pageWidth / 2) - 25;
          const rightX = pageWidth - 65;

          // Kiri
          doc.text('Mengetahui,', leftX, sigY);
          doc.text(leftSigner.title, leftX, sigY + 5);
          doc.text(leftSigner.name, leftX, sigY + 25);
          if (leftSigner.nip) doc.text(leftSigner.nip, leftX, sigY + 30);

          // Tengah
          doc.text('Mengetahui,', midX, sigY);
          doc.text(midSigner.title, midX, sigY + 5);
          doc.text(midSigner.name, midX, sigY + 25);
          if (midSigner.nip) doc.text(midSigner.nip, midX, sigY + 30);

          // Kanan
          doc.text(dateStr, rightX, sigY);
          doc.text(teacherLabel, rightX, sigY + 5);
          doc.text(teacherName, rightX, sigY + 25);
          if (teacherNIP) doc.text(teacherNIP, rightX, sigY + 30);
        } else {
          // Format 1 atau 2 Tanda Tangan
          const activeLeft = leftSigner.enabled ? leftSigner : (midSigner.enabled ? midSigner : null);
          const rightX = pageWidth - 80;

          if (activeLeft) {
            doc.text('Mengetahui,', 20, sigY);
            doc.text(activeLeft.title, 20, sigY + 5);
            doc.text(activeLeft.name, 20, sigY + 25);
            if (activeLeft.nip) doc.text(activeLeft.nip, 20, sigY + 30);
          }

          doc.text(dateStr, rightX, sigY);
          doc.text(teacherLabel, rightX, sigY + 5);
          doc.text(teacherName, rightX, sigY + 25);
          if (teacherNIP) doc.text(teacherNIP, rightX, sigY + 30);
        }

        doc.save(`${fileName}.pdf`);
        setShowExportSuccess('pdf');
      } catch (error) {
        console.error('Error generating PDF:', error);
      } finally {
        setIsExporting(false);
      }
    }, 800);
  };

  return (
    <div className="p-4 sm:p-6 lg:p-8 space-y-6">
      {/* Switcher Header */}
      <div className="bg-slate-100 p-2 rounded-2xl border-2 border-slate-200 shadow-sm flex flex-col sm:flex-row items-stretch gap-2.5">
        <button
          onClick={() => setActiveReportFrame('presensi')}
          className={`flex-1 flex items-center justify-center gap-2.5 py-3.5 px-5 rounded-xl font-black text-sm transition-all cursor-pointer ${
            activeReportFrame === 'presensi'
              ? 'bg-[#8dc63f] text-white shadow-md shadow-[#8dc63f]/30'
              : 'text-slate-600 hover:text-slate-900 hover:bg-white'
          }`}
        >
          <FileSpreadsheet className="w-5 h-5" />
          <span>Laporan Presensi Siswa</span>
        </button>
        <button
          onClick={() => setActiveReportFrame('wali_kelas')}
          className={`flex-1 flex items-center justify-center gap-2.5 py-3.5 px-5 rounded-xl font-black text-sm transition-all cursor-pointer ${
            activeReportFrame === 'wali_kelas'
              ? 'bg-amber-500 text-white shadow-md shadow-amber-500/30'
              : 'text-slate-600 hover:text-slate-900 hover:bg-white'
          }`}
        >
          <ClipboardList className="w-5 h-5" />
          <span>Laporan Wali Kelas</span>
          {profileData?.role !== 'Wali Kelas' && (
            <span className={`text-[10px] px-2 py-0.5 rounded-full font-black uppercase tracking-wider ${
              activeReportFrame === 'wali_kelas' ? 'bg-amber-700 text-amber-100' : 'bg-amber-100 text-amber-800'
            }`}>
              Khusus
            </span>
          )}
        </button>
      </div>

      {activeReportFrame === 'wali_kelas' ? (
        <div className="space-y-6">
          {profileData?.role !== 'Wali Kelas' && (
            <div className="bg-amber-50 border-2 border-amber-200 rounded-2xl p-5 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 shadow-xs">
              <div className="flex items-start gap-3">
                <div className="p-2.5 bg-amber-100 text-amber-700 rounded-xl mt-0.5">
                  <ClipboardList className="w-5 h-5" />
                </div>
                <div>
                  <h4 className="font-bold text-amber-900 text-sm">Informasi Peran Pengguna</h4>
                  <p className="text-xs text-amber-700 mt-0.5 leading-relaxed">
                    Peran Anda saat ini adalah <b>{profileData?.role || 'Guru Mapel'}</b>. Laporan ini dirancang khusus untuk Wali Kelas dalam membina dan menindaklanjuti siswa. Anda dapat mengubah peran menjadi <b>Wali Kelas</b> secara mandiri di menu <b>Profil Pengguna</b>.
                  </p>
                </div>
              </div>
              {onNavigateToProfile && (
                <button
                  onClick={onNavigateToProfile}
                  className="px-4 py-2.5 bg-amber-500 hover:bg-amber-600 text-white text-xs font-bold rounded-xl transition-all shadow-sm flex items-center gap-1.5 whitespace-nowrap self-stretch sm:self-auto justify-center cursor-pointer active:scale-95"
                >
                  <UserCheck className="w-4 h-4" /> Ubah Peran di Profil <ArrowRight className="w-3.5 h-3.5" />
                </button>
              )}
            </div>
          )}

          <HomeroomReportView
            classList={effectiveClassList}
            students={students}
            attendanceSessions={attendanceSessions}
            activeDb={activeDb}
            activeAuth={activeAuth}
            trackOp={trackOp}
            showToast={showToast}
            profileData={profileData}
          />
        </div>
      ) : (
        /* Frame 1: Laporan Presensi Siswa */
        <>
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div>
              <h2 className="text-2xl sm:text-3xl font-black text-slate-800 tracking-tight">Laporan Presensi Siswa</h2>
              <p className="text-slate-600 mt-1">Unduh &amp; cetak rekapitulasi kehadiran siswa</p>
            </div>
            
            <div className="flex flex-wrap w-full md:w-auto gap-3">
              {selectedClass && (
                <>
                  <button 
                    onClick={exportToExcel}
                    disabled={isExporting}
                    className="flex-1 sm:flex-none bg-emerald-600 hover:bg-emerald-700 text-white px-6 py-3 rounded-2xl font-bold flex items-center justify-center gap-2.5 shadow-md hover:shadow-lg transition-all active:scale-95 disabled:opacity-50 disabled:active:scale-100 cursor-pointer"
                  >
                    {isExporting ? <Loader2 className="w-5 h-5 animate-spin" /> : <FileSpreadsheet className="w-5 h-5" />} 
                    Export Excel
                  </button>
                  <button 
                    onClick={() => setShowPdfOptionsModal(true)}
                    disabled={isExporting}
                    className="flex-1 sm:flex-none bg-rose-600 hover:bg-rose-700 text-white px-6 py-3 rounded-2xl font-bold flex items-center justify-center gap-2.5 shadow-md hover:shadow-lg transition-all active:scale-95 disabled:opacity-50 disabled:active:scale-100 cursor-pointer"
                  >
                    {isExporting ? <Loader2 className="w-5 h-5 animate-spin" /> : <FileDown className="w-5 h-5" />} 
                    Export PDF
                  </button>
                </>
              )}
            </div>
          </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        <div className="md:col-span-1 space-y-6">
          <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
            <h3 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">
              <Filter className="w-5 h-5 text-emerald-700" /> Filter Laporan
            </h3>
            
            <div className="space-y-4">
              <div>
                <div className="flex items-center justify-between mb-2">
                  <label className="block text-xs font-bold text-slate-600 uppercase tracking-wider">Pilih Kelas</label>
                  {isWaliKelas && profileData?.waliKelasClass && (
                    <span className="text-[10px] font-bold text-emerald-800 bg-emerald-50 px-2 py-0.5 rounded-md border border-emerald-200">
                      Kelas Binaan
                    </span>
                  )}
                </div>
                <select 
                  className="w-full p-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:ring-2 focus:ring-emerald-100 focus:border-emerald-600 transition-colors disabled:bg-slate-100 disabled:text-slate-700 disabled:cursor-not-allowed"
                  value={selectedClass}
                  onChange={(e) => setSelectedClass(e.target.value)}
                  disabled={isWaliKelas && !!profileData?.waliKelasClass}
                >
                  {!isWaliKelas && <option value="">-- Pilih Kelas --</option>}
                  {!isWaliKelas && <option value="all">Semua Kelas</option>}
                  {effectiveClassList.slice().sort(compareClasses).map(c => <option key={c} value={c}>{c}</option>)}
                </select>
                {isWaliKelas && profileData?.waliKelasClass && (
                  <p className="text-[11px] text-emerald-800 font-bold mt-1.5 flex items-center gap-1.5">
                    <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                    Hanya menampilkan kelas binaan Anda ({profileData.waliKelasClass})
                  </p>
                )}
              </div>

              {selectedClass && (
                <div>
                  <label className="block text-xs font-bold text-slate-600 uppercase tracking-wider mb-2">Jenis Laporan</label>
                  <select 
                    className="w-full p-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:ring-2 focus:ring-emerald-100 focus:border-emerald-600 transition-colors font-medium text-slate-700"
                    value={reportType}
                    onChange={(e) => setReportType(e.target.value as 'daily' | 'monthly' | 'custom' | 'summary')}
                  >
                    <option value="summary">Rekap Total Keseluruhan</option>
                    <option value="monthly">Rekap Bulanan</option>
                    <option value="custom">Rentang Tanggal (Kustom)</option>
                    <option value="daily">Harian (Berdasarkan Tanggal)</option>
                  </select>
                </div>
              )}

              {selectedClass && reportType === 'custom' && (
                <div className="p-3.5 bg-slate-50 border-2 border-slate-300 rounded-xl space-y-3">
                  <div>
                    <label className="flex items-center gap-1.5 text-xs font-bold text-slate-600 uppercase tracking-wider mb-1.5">
                      <Calendar className="w-3.5 h-3.5 text-emerald-600" />
                      Tanggal Awal
                    </label>
                    <input 
                      type="date" 
                      className="w-full p-2.5 bg-white border border-slate-300 rounded-lg focus:ring-2 focus:ring-emerald-100 focus:border-emerald-600 transition-colors text-sm font-medium text-slate-700"
                      value={startDate}
                      onChange={(e) => setStartDate(e.target.value)}
                    />
                  </div>
                  <div>
                    <label className="flex items-center gap-1.5 text-xs font-bold text-slate-600 uppercase tracking-wider mb-1.5">
                      <Calendar className="w-3.5 h-3.5 text-emerald-600" />
                      Tanggal Akhir
                    </label>
                    <input 
                      type="date" 
                      className="w-full p-2.5 bg-white border border-slate-300 rounded-lg focus:ring-2 focus:ring-emerald-100 focus:border-emerald-600 transition-colors text-sm font-medium text-slate-700"
                      value={endDate}
                      onChange={(e) => setEndDate(e.target.value)}
                    />
                  </div>
                </div>
              )}

              {selectedClass && reportType === 'daily' && (
                <div>
                  <label className="block text-xs font-bold text-slate-600 uppercase tracking-wider mb-2">Pilih Tanggal</label>
                  {availableDailyDates.length > 0 ? (
                    <select 
                      className="w-full p-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:ring-2 focus:ring-emerald-100 focus:border-emerald-600 transition-colors"
                      value={selectedDailyDate}
                      onChange={(e) => setSelectedDailyDate(e.target.value)}
                    >
                      {availableDailyDates.map(d => (
                        <option key={d} value={d}>{format(parseISO(d), 'dd MMMM yyyy', {locale: id})}</option>
                      ))}
                    </select>
                  ) : (
                    <div className="w-full p-3 bg-slate-100 text-slate-600 border-2 border-slate-300 rounded-xl text-sm italic">
                      Belum ada sesi absensi
                    </div>
                  )}
                </div>
              )}

              {selectedClass && reportType === 'monthly' && (
                <div>
                  <label className="block text-xs font-bold text-slate-600 uppercase tracking-wider mb-2">Pilih Bulan</label>
                  <input 
                    type="month" 
                    className="w-full p-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:ring-2 focus:ring-emerald-100 focus:border-emerald-600 transition-colors"
                    value={selectedMonth}
                    onChange={(e) => setSelectedMonth(e.target.value)}
                  />
                </div>
              )}

              {/* Selector Penandatangan Laporan (Mengetahui 1 & Mengetahui 2) */}
              {selectedClass && (
                <div className="pt-3 border-t border-slate-200/80 space-y-3">
                  <div className="flex items-center justify-between">
                    <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider">
                      Opsi Tanda Tangan
                    </label>
                    {onNavigateToProfile && (
                      <button 
                        type="button" 
                        onClick={onNavigateToProfile} 
                        className="text-[10px] font-bold text-emerald-700 hover:text-emerald-800 underline"
                      >
                        Edit Profil
                      </button>
                    )}
                  </div>

                  {/* Penandatangan Kiri */}
                  <div>
                    <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                      1. Ttd Kiri (Mengetahui)
                    </label>
                    <select 
                      className="w-full p-2.5 bg-slate-50 border-2 border-slate-300 rounded-xl focus:ring-2 focus:ring-emerald-100 focus:border-emerald-600 transition-colors text-xs font-bold text-slate-800"
                      value={leftSignerRole}
                      onChange={(e) => setLeftSignerRole(e.target.value as SignerRoleType)}
                    >
                      <option value="kepala_sekolah">🏫 Kepala Sekolah {profileData?.namaKepalaSekolah ? `(${profileData.namaKepalaSekolah})` : ''}</option>
                      <option value="kurikulum">📚 Pihak Kurikulum {profileData?.namaKurikulum ? `(${profileData.namaKurikulum})` : ''}</option>
                      <option value="kesiswaan">👥 Pihak Kesiswaan {profileData?.namaKesiswaan ? `(${profileData.namaKesiswaan})` : ''}</option>
                      <option value="humas">📢 Pihak Humas {profileData?.namaHumas ? `(${profileData.namaHumas})` : ''}</option>
                      <option value="guru_bk">🤝 Guru BK {profileData?.namaBK ? `(${profileData.namaBK})` : ''}</option>
                      <option value="guru_wali">🎓 Guru Wali {profileData?.namaGuruWali ? `(${profileData.namaGuruWali})` : ''}</option>
                      <option value="none">➖ Tanpa Tanda Tangan Kiri</option>
                    </select>
                  </div>

                  {/* Penandatangan Tengah (Bila Perlu 3 Tanda Tangan) */}
                  <div>
                    <label className="block text-[11px] font-semibold text-slate-600 mb-1 flex items-center justify-between">
                      <span>2. Ttd Tengah (Mengetahui 2)</span>
                      <span className="text-[9px] font-normal text-slate-400">Bila perlu 3 TTD</span>
                    </label>
                    <select 
                      className="w-full p-2.5 bg-slate-50 border-2 border-slate-300 rounded-xl focus:ring-2 focus:ring-emerald-100 focus:border-emerald-600 transition-colors text-xs font-bold text-slate-800"
                      value={midSignerRole}
                      onChange={(e) => setMidSignerRole(e.target.value as SignerRoleType)}
                    >
                      <option value="none">➖ Tanpa Ttd Tengah (2 TTD)</option>
                      <option value="kurikulum">📚 Pihak Kurikulum {profileData?.namaKurikulum ? `(${profileData.namaKurikulum})` : ''}</option>
                      <option value="kesiswaan">👥 Pihak Kesiswaan {profileData?.namaKesiswaan ? `(${profileData.namaKesiswaan})` : ''}</option>
                      <option value="humas">📢 Pihak Humas {profileData?.namaHumas ? `(${profileData.namaHumas})` : ''}</option>
                      <option value="guru_bk">🤝 Guru BK {profileData?.namaBK ? `(${profileData.namaBK})` : ''}</option>
                      <option value="guru_wali">🎓 Guru Wali {profileData?.namaGuruWali ? `(${profileData.namaGuruWali})` : ''}</option>
                      <option value="kepala_sekolah">🏫 Kepala Sekolah {profileData?.namaKepalaSekolah ? `(${profileData.namaKepalaSekolah})` : ''}</option>
                    </select>
                  </div>

                  {/* Ringkasan Penandatangan */}
                  <div className="p-2.5 bg-emerald-50/70 border border-emerald-200/80 rounded-xl text-[10.5px] space-y-1">
                    <div className="flex items-center justify-between text-slate-500 font-medium">
                      <span>Format:</span>
                      <span className="font-bold text-emerald-800">
                        {leftSignerRole !== 'none' && midSignerRole !== 'none' ? '3 Tanda Tangan (Kiri, Tengah, Kanan)' : (leftSignerRole !== 'none' || midSignerRole !== 'none') ? '2 Tanda Tangan (Kiri & Kanan)' : '1 Tanda Tangan (Kanan)'}
                      </span>
                    </div>
                    {leftSignerRole !== 'none' && (
                      <div className="text-slate-700 truncate">
                        <span className="font-bold text-emerald-900">Kiri:</span> {getLeftSignerDetails().title} ({getLeftSignerDetails().name})
                      </div>
                    )}
                    {midSignerRole !== 'none' && (
                      <div className="text-slate-700 truncate">
                        <span className="font-bold text-emerald-900">Tengah:</span> {getMidSignerDetails().title} ({getMidSignerDetails().name})
                      </div>
                    )}
                    <div className="text-slate-700 truncate">
                      <span className="font-bold text-emerald-900">Kanan:</span> {profileData?.role === 'Wali Kelas' ? 'Wali Kelas' : 'Guru Mapel'} ({profileData?.namaGuruMapel || 'Belum diisi'})
                    </div>
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>

        <div className="md:col-span-2">
          {!selectedClass ? (
            <div className="bg-slate-50/50 rounded-2xl border border-slate-100 h-full min-h-[300px] flex flex-col items-center justify-center p-8 text-center space-y-4">
              <div className="p-4 bg-white rounded-full shadow-sm">
                <FileText className="w-8 h-8 text-slate-300" />
              </div>
              <div>
                <h3 className="text-lg font-bold text-slate-700">Pilih Kelas Terlebih Dahulu</h3>
                <p className="text-slate-600 mt-1">Silakan pilih kelas melalui filter di samping untuk melihat preview laporan.</p>
              </div>
            </div>
          ) : (
            <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
              <div className="p-6 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                <div>
                  <h3 className="text-lg font-bold text-slate-800">
                    Preview Laporan {reportType === 'summary' ? 'Total' : reportType === 'monthly' ? 'Bulanan' : reportType === 'custom' ? 'Rentang Tanggal' : 'Harian'}
                  </h3>
                  <p className="text-slate-600 text-sm mt-1">
                    Kelas: <span className="font-bold text-slate-700">{selectedClass}</span> 
                    {reportType === 'monthly' && ` | Bulan: ${selectedMonth}`}
                    {reportType === 'custom' && startDate && endDate && ` | Rentang: ${format(parseISO(startDate), 'dd MMM yyyy', {locale: id})} - ${format(parseISO(endDate), 'dd MMM yyyy', {locale: id})}`}
                    {reportType === 'daily' && selectedDailyDate && availableDailyDates.length > 0 && ` | Tanggal: ${format(parseISO(selectedDailyDate), 'dd MMMM yyyy', {locale: id})}`}
                    {reportType === 'summary' && classSessions.length > 0 && ` | Tanggal: ${format(parseISO(classSessions[0].date), 'dd MMM yyyy', {locale: id})} - ${format(parseISO(classSessions[classSessions.length - 1].date), 'dd MMM yyyy', {locale: id})}`}
                  </p>
                </div>
                <div className="flex bg-slate-50 rounded-lg p-1 border border-slate-100">
                  <span className="px-3 py-1 text-xs font-bold text-slate-600">Total: {classStudents.length} Siswa</span>
                </div>
              </div>

              {/* Statistik Kelas Section */}
              {classStats && (
                <div className="p-6 bg-slate-50 border-b border-slate-100">
                  <h4 className="text-xs font-extrabold text-slate-500 uppercase tracking-wider mb-3">Statistik Akumulasi Kelas</h4>
                  <div className="grid grid-cols-2 sm:grid-cols-5 gap-4">
                    
                    {/* Hadir */}
                    <div className="bg-white p-4 rounded-xl border border-slate-100 shadow-xs flex flex-col justify-between">
                      <span className="text-[10px] font-bold text-emerald-600 uppercase tracking-wider">Kehadiran (Hadir)</span>
                      <div className="flex items-baseline gap-1.5 mt-1">
                        <span className="text-xl font-black text-emerald-700">{classStats.hadir}</span>
                        <span className="text-xs font-bold text-slate-400">kali</span>
                      </div>
                      <div className="mt-2 w-full bg-slate-100 h-1.5 rounded-full overflow-hidden">
                        <div className="bg-emerald-500 h-full rounded-full" style={{ width: `${classStats.persentaseHadir}%` }}></div>
                      </div>
                      <span className="text-[10px] font-black text-slate-500 mt-1.5">{classStats.persentaseHadir.toFixed(2)}% dari total</span>
                    </div>

                    {/* Sakit */}
                    <div className="bg-white p-4 rounded-xl border border-slate-100 shadow-xs flex flex-col justify-between">
                      <span className="text-[10px] font-bold text-amber-600 uppercase tracking-wider">Sakit</span>
                      <div className="flex items-baseline gap-1.5 mt-1">
                        <span className="text-xl font-black text-amber-700">{classStats.sakit}</span>
                        <span className="text-xs font-bold text-slate-400">kali</span>
                      </div>
                      <div className="mt-2 w-full bg-slate-100 h-1.5 rounded-full overflow-hidden">
                        <div className="bg-amber-500 h-full rounded-full" style={{ width: `${classStats.persentaseSakit}%` }}></div>
                      </div>
                      <span className="text-[10px] font-black text-slate-500 mt-1.5">{classStats.persentaseSakit.toFixed(2)}% dari total</span>
                    </div>

                    {/* Izin */}
                    <div className="bg-white p-4 rounded-xl border border-slate-100 shadow-xs flex flex-col justify-between">
                      <span className="text-[10px] font-bold text-blue-600 uppercase tracking-wider">Izin</span>
                      <div className="flex items-baseline gap-1.5 mt-1">
                        <span className="text-xl font-black text-blue-700">{classStats.izin}</span>
                        <span className="text-xs font-bold text-slate-400">kali</span>
                      </div>
                      <div className="mt-2 w-full bg-slate-100 h-1.5 rounded-full overflow-hidden">
                        <div className="bg-blue-500 h-full rounded-full" style={{ width: `${classStats.persentaseIzin}%` }}></div>
                      </div>
                      <span className="text-[10px] font-black text-slate-500 mt-1.5">{classStats.persentaseIzin.toFixed(2)}% dari total</span>
                    </div>

                    {/* Alpa */}
                    <div className="bg-white p-4 rounded-xl border border-slate-100 shadow-xs flex flex-col justify-between col-span-1">
                      <span className="text-[10px] font-bold text-rose-600 uppercase tracking-wider">Alpa</span>
                      <div className="flex items-baseline gap-1.5 mt-1">
                        <span className="text-xl font-black text-rose-700">{classStats.alpa}</span>
                        <span className="text-xs font-bold text-slate-400">kali</span>
                      </div>
                      <div className="mt-2 w-full bg-slate-100 h-1.5 rounded-full overflow-hidden">
                        <div className="bg-rose-500 h-full rounded-full" style={{ width: `${classStats.persentaseAlpa}%` }}></div>
                      </div>
                      <span className="text-[10px] font-black text-slate-500 mt-1.5">{classStats.persentaseAlpa.toFixed(2)}% dari total</span>
                    </div>

                    {/* Dispen */}
                    <div className="bg-white p-4 rounded-xl border border-slate-100 shadow-xs flex flex-col justify-between col-span-2 sm:col-span-1">
                      <span className="text-[10px] font-bold text-indigo-600 uppercase tracking-wider">Dispen</span>
                      <div className="flex items-baseline gap-1.5 mt-1">
                        <span className="text-xl font-black text-indigo-700">{classStats.dispen}</span>
                        <span className="text-xs font-bold text-slate-400">kali</span>
                      </div>
                      <div className="mt-2 w-full bg-slate-100 h-1.5 rounded-full overflow-hidden">
                        <div className="bg-indigo-500 h-full rounded-full" style={{ width: `${classStats.persentaseDispen}%` }}></div>
                      </div>
                      <span className="text-[10px] font-black text-slate-500 mt-1.5">{classStats.persentaseDispen.toFixed(2)}% dari total</span>
                    </div>

                  </div>
                </div>
              )}

              <div className="overflow-x-auto max-h-[500px] scrollbar-thin">
                <table className="w-full min-w-[600px] text-sm text-left">
                  <thead className="bg-slate-50 sticky top-0 z-10 shadow-[0_1px_0_0_rgba(226,232,240,1)]">
                    <tr>
                      <th className="p-4 font-bold text-slate-600 border-b border-slate-200">No</th>
                      <th className="p-4 font-bold text-slate-600 border-b border-slate-200">Nama Lengkap Siswa</th>
                      {selectedClass === 'all' && (
                        <th className="p-4 font-bold text-slate-600 border-b border-slate-200">Kelas</th>
                      )}
                      {(reportType === 'summary' || reportType === 'custom' || (reportType === 'monthly' && selectedClass === 'all')) && (
                        <>
                          <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200"><span className="text-emerald-700">Hadir</span></th>
                          <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200">Sakit</th>
                          <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200">Izin</th>
                          <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200"><span className="text-rose-600">Alpa</span></th>
                          <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200">Dispen</th>
                          <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200">Presentase %</th>
                        </>
                      )}
                      {(reportType === 'monthly' && selectedClass !== 'all') && (
                        <>
                           <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200">Hadir</th>
                           <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200">Sakit</th>
                           <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200">Izin</th>
                           <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200"><span className="text-rose-600">Alpa</span></th>
                           <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200">Dispen</th>
                           <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200">Presentase %</th>
                        </>
                      )}
                      {reportType === 'daily' && (
                         <>
                           <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200">Status Kehadiran</th>
                           <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200">Presentase %</th>
                         </>
                      )}
                    </tr>
                  </thead>
                  <tbody>
                    {classStudents.map((student, i) => {
                      let hadir = 0, sakit = 0, izin = 0, alpa = 0, dispen = 0;
                      let recentStatus = '-';

                      if (reportType === 'summary') {
                        const studentSessions = classSessions.filter(s => selectedClass === 'all' ? s.className === student.class : true);
                        studentSessions.forEach(session => {
                          const status = session.records[student.id];
                          if (status === 'Hadir') hadir++;
                          if (status === 'Sakit') sakit++;
                          if (status === 'Izin') izin++;
                          if (status === 'Alpa') alpa++;
                          if (status === 'Dispen') dispen++;
                        });
                      } else if (reportType === 'monthly') {
                        const studentMonthlySessions = monthlySessions.filter(s => selectedClass === 'all' ? s.className === student.class : true);
                        studentMonthlySessions.forEach(session => {
                          const status = session.records[student.id];
                          if (status === 'Hadir') hadir++;
                          if (status === 'Sakit') sakit++;
                          if (status === 'Izin') izin++;
                          if (status === 'Alpa') alpa++;
                          if (status === 'Dispen') dispen++;
                        });
                      } else if (reportType === 'custom') {
                        const studentRangeSessions = customRangeSessions.filter(s => selectedClass === 'all' ? s.className === student.class : true);
                        studentRangeSessions.forEach(session => {
                          const status = session.records[student.id];
                          if (status === 'Hadir') hadir++;
                          if (status === 'Sakit') sakit++;
                          if (status === 'Izin') izin++;
                          if (status === 'Alpa') alpa++;
                          if (status === 'Dispen') dispen++;
                        });
                      } else {
                        const studentDailySessions = classSessions.filter(s => 
                          s.date === selectedDailyDate && (selectedClass === 'all' ? s.className === student.class : true)
                        );
                        const targetSession = studentDailySessions[studentDailySessions.length - 1];
                        if (targetSession) {
                          recentStatus = targetSession.records[student.id] || '-';
                        }

                        const studentAllSessions = classSessions.filter(s => selectedClass === 'all' ? s.className === student.class : true);
                        studentAllSessions.forEach(session => {
                          const st = session.records[student.id];
                          if (st === 'Hadir') hadir++;
                          if (st === 'Sakit') sakit++;
                          if (st === 'Izin') izin++;
                          if (st === 'Alpa') alpa++;
                          if (st === 'Dispen') dispen++;
                        });
                      }

                      const totalRecorded = hadir + sakit + izin + alpa + dispen;
                      const persentase = totalRecorded > 0 ? ((hadir + dispen) / totalRecorded) * 100 : (dispen > 0 ? 100 : 0);

                      return (
                        <tr key={student.id} className="border-b last:border-b-0 border-slate-100 hover:bg-slate-50/50 transition-colors">
                          <td className="p-4 text-slate-600 font-medium">{i + 1}</td>
                          <td className="p-4">
                            <div className="font-bold text-slate-800">{student.name}</div>
                            <div className="text-[10px] text-slate-500 font-mono mt-0.5">{student.nisn}</div>
                          </td>
                          {selectedClass === 'all' && (
                            <td className="p-4 text-slate-600 font-bold">{student.class}</td>
                          )}
                          {(reportType === 'summary' || reportType === 'custom' || (reportType === 'monthly' && selectedClass === 'all')) && (
                            <>
                              <td className="p-4 text-center font-bold text-emerald-700">{hadir}</td>
                              <td className="p-4 text-center font-medium text-slate-600">{sakit}</td>
                              <td className="p-4 text-center font-medium text-slate-600">{izin}</td>
                              <td className="p-4 text-center font-bold text-rose-500">{alpa}</td>
                              <td className="p-4 text-center font-medium text-slate-600">{dispen}</td>
                              <td className="p-4 text-center">
                                <span className={`px-2.5 py-1 rounded-full text-xs font-bold ${persentase >= 80 ? 'bg-emerald-100 text-lime-700' : persentase >= 50 ? 'bg-amber-100 text-amber-700' : 'bg-rose-100 text-rose-700'}`}>
                                  {persentase.toFixed(2)}%
                                </span>
                              </td>
                            </>
                          )}
                          {(reportType === 'monthly' && selectedClass !== 'all') && (
                             <>
                              <td className="p-4 text-center font-bold text-emerald-700">{hadir}</td>
                              <td className="p-4 text-center font-medium text-slate-600">{sakit}</td>
                              <td className="p-4 text-center font-medium text-slate-600">{izin}</td>
                              <td className="p-4 text-center font-bold text-rose-500">{alpa}</td>
                              <td className="p-4 text-center font-medium text-slate-600">{dispen}</td>
                              <td className="p-4 text-center">
                                <span className={`px-2.5 py-1 rounded-full text-xs font-bold ${persentase >= 80 ? 'bg-emerald-100 text-lime-700' : persentase >= 50 ? 'bg-amber-100 text-amber-700' : 'bg-rose-100 text-rose-700'}`}>
                                  {persentase.toFixed(2)}%
                                </span>
                              </td>
                             </>
                          )}
                          {reportType === 'daily' && (
                             <>
                               <td className="p-4 text-center">
                                 <span className={`px-2.5 py-1 rounded-full text-xs font-bold inline-block
                                   ${recentStatus === 'Hadir' ? 'bg-emerald-100 text-emerald-600' : 
                                     recentStatus === 'Sakit' || recentStatus === 'Izin' || recentStatus === 'Dispen' ? 'bg-amber-100 text-amber-600' : 
                                     recentStatus === 'Alpa' ? 'bg-rose-100 text-rose-600' : 'bg-slate-100 text-slate-500'}
                                 `}>
                                   {recentStatus || '-'}
                                 </span>
                               </td>
                               <td className="p-4 text-center">
                                 <span className={`px-2.5 py-1 rounded-full text-xs font-bold ${persentase >= 80 ? 'bg-emerald-100 text-lime-700' : persentase >= 50 ? 'bg-amber-100 text-amber-700' : 'bg-rose-100 text-rose-700'}`}>
                                   {persentase.toFixed(2)}%
                                 </span>
                               </td>
                             </>
                          )}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      </div>
      </>
      )}

      <AnimatePresence>
        {isExporting && (
           <motion.div 
             initial={{ opacity: 0 }}
             animate={{ opacity: 1 }}
             exit={{ opacity: 0 }}
             className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 backdrop-blur-md p-4"
           >
             <motion.div
               initial={{ scale: 0.9, opacity: 0 }}
               animate={{ scale: 1, opacity: 1 }}
               exit={{ scale: 0.9, opacity: 0 }}
               className="bg-white rounded-[2.5rem] p-10 max-w-sm w-full shadow-[0_25px_60px_-15px_rgba(0,0,0,0.2)] text-center border border-slate-100"
             >
                <div className="relative mx-auto w-24 h-24 mb-6">
                  <motion.div 
                    animate={{ rotate: 360 }}
                    transition={{ duration: 2, repeat: Infinity, ease: "linear" }}
                    className="absolute inset-0 rounded-full border-4 border-emerald-100 border-t-emerald-500"
                  />
                  <div className="absolute inset-0 flex items-center justify-center">
                    <FileDown className="w-10 h-10 text-emerald-500" />
                  </div>
                </div>
                <h3 className="text-xl font-black text-slate-800 mb-2">Menyiapkan Laporan</h3>
                <p className="text-slate-600 text-sm leading-relaxed">
                  Mohon tunggu sebentar, kami sedang memproses data absensi menjadi format yang elegan untuk Anda.
                </p>
             </motion.div>
           </motion.div>
        )}

        {showExportSuccess !== 'none' && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-xl p-4"
          >
            <motion.div
              initial={{ y: 50, scale: 0.9, opacity: 0 }}
              animate={{ y: 0, scale: 1, opacity: 1 }}
              exit={{ y: 50, scale: 0.9, opacity: 0 }}
              transition={{ type: "spring", damping: 25, stiffness: 300 }}
              className="bg-white rounded-[3rem] p-10 max-w-sm w-full shadow-[0_40px_100px_-20px_rgba(0,0,0,0.3)] text-center relative overflow-hidden group"
            >
               {/* Decorative background elements */}
               <div className={`absolute top-0 left-0 w-full h-2 bg-gradient-to-r ${showExportSuccess === 'no_data' ? 'from-amber-400 to-rose-400' : 'from-emerald-400 to-sky-400'}`}></div>
               <div className={`absolute -top-12 -right-12 w-32 h-32 ${showExportSuccess === 'no_data' ? 'bg-amber-50' : 'bg-emerald-50'} rounded-full opacity-50 group-hover:scale-110 transition-transform duration-700`}></div>

               <div className="relative z-10">
                  <div className={`mx-auto w-24 h-24 ${showExportSuccess === 'no_data' ? 'bg-amber-100' : 'bg-emerald-100'} rounded-[2rem] flex items-center justify-center mb-8 rotate-3 group-hover:rotate-6 transition-transform`}>
                     {showExportSuccess === 'no_data' ? (
                       <FileWarning className="w-12 h-12 text-amber-600" />
                     ) : (
                       <CheckCircle2 className="w-12 h-12 text-emerald-600" />
                     )}
                  </div>
                  
                  <h3 className="text-2xl font-black text-slate-900 mb-3 tracking-tight">
                    {showExportSuccess === 'no_data' ? 'Data Kosong' : 'Ekspor Berhasil!'}
                  </h3>
                  <p className="text-slate-600 text-sm mb-10 leading-relaxed font-medium">
                    {showExportSuccess === 'no_data' 
                      ? 'Maaf, kami tidak menemukan data absensi untuk kriteria yang Anda pilih. Silakan pastikan kelas dan rentang waktu sudah benar.' 
                      : `Dokumen ${showExportSuccess === 'excel' ? 'Excel (.xlsx)' : 'PDF (.pdf)'} telah berhasil dibuat dan siap untuk Anda bagikan atau simpan.`}
                  </p>

                  <button 
                    onClick={() => setShowExportSuccess('none')}
                    className={`w-full ${showExportSuccess === 'no_data' ? 'bg-amber-600 hover:bg-amber-700' : 'bg-slate-900 hover:bg-slate-800'} text-white py-4 rounded-2xl font-black text-sm hover:scale-[1.02] active:scale-[0.98] transition-all shadow-xl ${showExportSuccess === 'no_data' ? 'shadow-amber-200' : 'shadow-slate-200'}`}
                  >
                    {showExportSuccess === 'no_data' ? 'Coba Lagi' : 'Selesai'}
                  </button>
               </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* PDF Export Options Modal */}
      <AnimatePresence>
        {showPdfOptionsModal && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-[60] flex items-center justify-center p-4"
          >
            <div className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm" onClick={() => setShowPdfOptionsModal(false)}></div>
            <motion.div 
              initial={{ scale: 0.9, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.9, opacity: 0, y: 20 }}
              className="bg-white w-full max-w-md rounded-[2rem] shadow-2xl relative z-10 overflow-hidden"
            >
              <div className="p-8">
                <div className="flex items-center gap-4 mb-6">
                  <div className="w-12 h-12 bg-rose-100 text-rose-600 rounded-2xl flex items-center justify-center rotate-3">
                    <FileDown className="w-6 h-6" />
                  </div>
                  <div>
                    <h3 className="text-xl font-black text-slate-800">Opsi Ekspor PDF</h3>
                    <p className="text-slate-500 text-sm font-medium">Pilih komponen tabel tambahan</p>
                  </div>
                </div>

                <div className="space-y-4 mb-6">
                  {/* Selector Penandatangan di Modal */}
                  <div className="p-4 bg-slate-50 border-2 border-slate-200/80 rounded-2xl space-y-3">
                    <div>
                      <label className="block text-xs font-bold text-slate-800 uppercase tracking-wider mb-1">
                        1. Penandatangan Kiri (Mengetahui)
                      </label>
                      <select 
                        value={leftSignerRole}
                        onChange={(e) => setLeftSignerRole(e.target.value as SignerRoleType)}
                        className="w-full p-2.5 bg-white border-2 border-slate-300 rounded-xl text-xs font-bold text-slate-800 focus:ring-2 focus:ring-rose-500/20 focus:border-rose-500 transition-colors"
                      >
                        <option value="kepala_sekolah">🏫 Kepala Sekolah ({profileData?.namaKepalaSekolah || 'Belum diisi'})</option>
                        <option value="kurikulum">📚 Pihak Kurikulum ({profileData?.namaKurikulum || 'Belum diisi'})</option>
                        <option value="kesiswaan">👥 Pihak Kesiswaan ({profileData?.namaKesiswaan || 'Belum diisi'})</option>
                        <option value="humas">📢 Pihak Humas ({profileData?.namaHumas || 'Belum diisi'})</option>
                        <option value="guru_bk">🤝 Guru BK ({profileData?.namaBK || 'Belum diisi'})</option>
                        <option value="guru_wali">🎓 Guru Wali ({profileData?.namaGuruWali || 'Belum diisi'})</option>
                        <option value="none">➖ Tanpa Tanda Tangan Kiri</option>
                      </select>
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-slate-800 uppercase tracking-wider mb-1 flex items-center justify-between">
                        <span>2. Penandatangan Tengah</span>
                        <span className="text-[10px] text-slate-400 font-normal">Bila perlu 3 TTD</span>
                      </label>
                      <select 
                        value={midSignerRole}
                        onChange={(e) => setMidSignerRole(e.target.value as SignerRoleType)}
                        className="w-full p-2.5 bg-white border-2 border-slate-300 rounded-xl text-xs font-bold text-slate-800 focus:ring-2 focus:ring-rose-500/20 focus:border-rose-500 transition-colors"
                      >
                        <option value="none">➖ Tanpa Ttd Tengah (2 TTD)</option>
                        <option value="kurikulum">📚 Pihak Kurikulum ({profileData?.namaKurikulum || 'Belum diisi'})</option>
                        <option value="kesiswaan">👥 Pihak Kesiswaan ({profileData?.namaKesiswaan || 'Belum diisi'})</option>
                        <option value="humas">📢 Pihak Humas ({profileData?.namaHumas || 'Belum diisi'})</option>
                        <option value="guru_bk">🤝 Guru BK ({profileData?.namaBK || 'Belum diisi'})</option>
                        <option value="guru_wali">🎓 Guru Wali ({profileData?.namaGuruWali || 'Belum diisi'})</option>
                        <option value="kepala_sekolah">🏫 Kepala Sekolah ({profileData?.namaKepalaSekolah || 'Belum diisi'})</option>
                      </select>
                    </div>

                    <div className="pt-2 border-t border-slate-200 text-[10.5px] text-slate-600">
                      <span className="font-bold text-rose-700">Tanda Tangan Kanan (Otomatis):</span> {profileData?.role === 'Wali Kelas' ? 'Wali Kelas' : 'Guru Mapel'} ({profileData?.namaGuruMapel || 'Belum diisi'})
                    </div>
                  </div>

                  <label className="flex items-start gap-3 p-4 rounded-xl border-2 border-slate-100 hover:border-slate-200 cursor-pointer transition-colors">
                    <input 
                      type="checkbox" 
                      checked={includePdfStats}
                      onChange={(e) => setIncludePdfStats(e.target.checked)}
                      className="mt-1 w-5 h-5 rounded border-slate-300 text-rose-600 focus:ring-rose-600"
                    />
                    <div>
                      <div className="font-bold text-slate-800">Statistik Akumulasi Kelas</div>
                      <div className="text-sm text-slate-500">Persentase kehadiran, sakit, izin, dll.</div>
                    </div>
                  </label>
                  <label className="flex items-start gap-3 p-4 rounded-xl border-2 border-slate-100 hover:border-slate-200 cursor-pointer transition-colors">
                    <input 
                      type="checkbox" 
                      checked={includePdfTopStudents}
                      onChange={(e) => setIncludePdfTopStudents(e.target.checked)}
                      className="mt-1 w-5 h-5 rounded border-slate-300 text-rose-600 focus:ring-rose-600"
                    />
                    <div>
                      <div className="font-bold text-slate-800">Daftar Siswa Terbanyak</div>
                      <div className="text-sm text-slate-500">Siswa dengan Alpa, Sakit, dan Izin tertinggi.</div>
                    </div>
                  </label>
                </div>

                <div className="flex gap-3">
                  <button 
                    onClick={() => setShowPdfOptionsModal(false)}
                    className="flex-1 bg-slate-100 hover:bg-slate-200 text-slate-700 py-3.5 rounded-xl font-bold transition-colors active:scale-95"
                  >
                    Batal
                  </button>
                  <button 
                    onClick={() => {
                      setShowPdfOptionsModal(false);
                      exportToPdf();
                    }}
                    className="flex-1 bg-rose-600 hover:bg-rose-700 text-white py-3.5 rounded-xl font-bold flex items-center justify-center gap-2 transition-colors shadow-lg shadow-rose-200 active:scale-95"
                  >
                    <CheckCircle2 className="w-5 h-5" />
                    Ekspor
                  </button>
                </div>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

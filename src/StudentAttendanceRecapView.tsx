import React, { useState, useMemo, useEffect, useCallback } from 'react';
import { 
  Calendar, 
  User as UserIcon, 
  Building2, 
  FileSpreadsheet, 
  Search, 
  ArrowLeft, 
  Loader2, 
  FileDown, 
  Info,
  ZoomIn,
  ZoomOut,
  RotateCcw
} from 'lucide-react';
import { format, parseISO } from 'date-fns';
import { id } from 'date-fns/locale';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as XLSX from 'xlsx';
import { kopSuratBase64 } from './kop-surat-b64';
import { compareClasses } from './classSortUtils';

export interface StudentRecapItem {
  id: string;
  nisn?: string;
  name: string;
  class: string;
  userId?: string;
}

export interface AttendanceSessionRecapItem {
  id: string;
  date: string;
  className: string;
  meetingNumber?: number;
  records: Record<string, string>;
  userId?: string;
}

export type RecapSignerRoleType = 'kepala_sekolah' | 'kurikulum' | 'kesiswaan' | 'humas' | 'guru_bk' | 'guru_bk_2' | 'guru_wali' | 'none';

interface StudentAttendanceRecapViewProps {
  students: StudentRecapItem[];
  attendanceSessions: AttendanceSessionRecapItem[];
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
    namaBK2?: string;
    nipBK2?: string;
    namaHumas?: string;
    nipHumas?: string;
    jabatanHumas?: string;
    semester: string;
    tahunPelajaran: string;
    mataPelajaran: string;
    role?: string;
    waliKelasClass?: string;
  };
  defaultClass?: string;
  showToast?: (message: string, type: 'success' | 'info' | 'error') => void;
  onBackToClassReport?: () => void;
}

export default function StudentAttendanceRecapView({
  students,
  attendanceSessions,
  profileData,
  defaultClass,
  showToast,
  onBackToClassReport
}: StudentAttendanceRecapViewProps) {
  const isWaliKelas = profileData?.role === 'Wali Kelas';
  const rawHomeroom = (profileData?.waliKelasClass || defaultClass || '').trim();
  const homeroomClass = rawHomeroom === 'rekap_siswa' ? '' : rawHomeroom;

  // Class selection (defaults to homeroom class)
  const availableClasses = useMemo(() => {
    const clsSet = new Set<string>();
    if (homeroomClass && homeroomClass !== 'rekap_siswa') clsSet.add(homeroomClass);
    students.forEach(s => {
      if (s.class && s.class !== 'rekap_siswa') clsSet.add(s.class);
    });
    attendanceSessions.forEach(s => {
      if (s.className && s.className !== 'rekap_siswa') clsSet.add(s.className);
    });
    return Array.from(clsSet).sort(compareClasses);
  }, [homeroomClass, students, attendanceSessions]);

  const [selectedClass, setSelectedClass] = useState<string>(() => {
    if (homeroomClass && homeroomClass !== 'rekap_siswa') return homeroomClass;
    return availableClasses[0] || '';
  });

  useEffect(() => {
    if (!selectedClass && availableClasses.length > 0) {
      setSelectedClass(homeroomClass || availableClasses[0]);
    }
  }, [homeroomClass, selectedClass, availableClasses]);

  // Students in selected class
  const classStudents = useMemo(() => {
    if (!selectedClass) return [];
    return students
      .filter(s => s.class === selectedClass)
      .sort((a, b) => a.name.localeCompare(b.name, 'id'));
  }, [students, selectedClass]);

  // Selected student
  const [selectedStudentId, setSelectedStudentId] = useState<string>('');
  const [studentSearch, setStudentSearch] = useState<string>('');

  useEffect(() => {
    if (classStudents.length > 0) {
      if (!selectedStudentId || !classStudents.some(s => s.id === selectedStudentId)) {
        setSelectedStudentId(classStudents[0].id);
      }
    } else {
      setSelectedStudentId('');
    }
  }, [classStudents, selectedStudentId]);

  const selectedStudent = useMemo(() => {
    return classStudents.find(s => s.id === selectedStudentId) || classStudents[0] || null;
  }, [classStudents, selectedStudentId]);

  // Period filter
  const [periodType, setPeriodType] = useState<'all' | 'monthly' | 'custom'>('all');
  const [selectedMonth, setSelectedMonth] = useState<string>(() => {
    return new Date().toISOString().substring(0, 7);
  });
  const [startDate, setStartDate] = useState<string>(() => {
    const now = new Date();
    const ym = now.toISOString().substring(0, 7);
    return `${ym}-01`;
  });
  const [endDate, setEndDate] = useState<string>(() => {
    return new Date().toISOString().substring(0, 10);
  });

  // Signer options
  const [leftSignerRole, setLeftSignerRole] = useState<RecapSignerRoleType>('kepala_sekolah');
  const [midSignerRole, setMidSignerRole] = useState<RecapSignerRoleType>('none');
  const [rightSignerType, setRightSignerType] = useState<'auto' | 'wali_kelas' | 'guru_mapel'>('auto');

  // Export loading state
  const [isExporting, setIsExporting] = useState<boolean>(false);

  // Zoom control state for A4 preview
  const [zoomScale, setZoomScale] = useState<number>(1);
  const handleZoomIn = () => setZoomScale(prev => Math.min(1.6, Number((prev + 0.1).toFixed(1))));
  const handleZoomOut = () => setZoomScale(prev => Math.max(0.5, Number((prev - 0.1).toFixed(1))));
  const handleResetZoom = () => setZoomScale(1);

  // Filtered attendance sessions for this student's class and selected period
  const studentSessions = useMemo(() => {
    if (!selectedClass || !selectedStudent) return [];

    let sessions = attendanceSessions.filter(s => s.className === selectedClass);

    if (periodType === 'monthly') {
      sessions = sessions.filter(s => s.date && s.date.startsWith(selectedMonth));
    } else if (periodType === 'custom') {
      if (startDate) sessions = sessions.filter(s => s.date >= startDate);
      if (endDate) sessions = sessions.filter(s => s.date <= endDate);
    }

    // Sort chronologically ascending
    return sessions.sort((a, b) => a.date.localeCompare(b.date));
  }, [selectedClass, selectedStudent, attendanceSessions, periodType, selectedMonth, startDate, endDate]);

  // Individual attendance rows
  const recapRows = useMemo(() => {
    if (!selectedStudent) return [];

    return studentSessions.map((session, idx) => {
      let formattedDate = session.date;
      try {
        if (session.date) {
          const parsed = parseISO(session.date);
          if (!isNaN(parsed.getTime())) {
            formattedDate = format(parsed, 'EEEE, dd MMMM yyyy', { locale: id });
          }
        }
      } catch {
        formattedDate = session.date;
      }

      const status = session.records?.[selectedStudent.id] || 'Belum Diisi';

      return {
        no: idx + 1,
        rawDate: session.date,
        formattedDate,
        meetingNumber: session.meetingNumber,
        status
      };
    });
  }, [selectedStudent, studentSessions]);

  // Statistics
  const recapStats = useMemo(() => {
    let hadir = 0;
    let sakit = 0;
    let izin = 0;
    let alpa = 0;
    let dispen = 0;

    recapRows.forEach(r => {
      if (r.status === 'Hadir') hadir++;
      else if (r.status === 'Sakit') sakit++;
      else if (r.status === 'Izin') izin++;
      else if (r.status === 'Alpa') alpa++;
      else if (r.status === 'Dispen') dispen++;
    });

    const totalSessions = recapRows.length;
    const recordedCount = hadir + sakit + izin + alpa + dispen;
    const effectiveHadir = hadir + dispen;
    const persentase = totalSessions > 0 
      ? Math.round((effectiveHadir / totalSessions) * 100) 
      : 0;

    return {
      totalSessions,
      recordedCount,
      hadir,
      sakit,
      izin,
      alpa,
      dispen,
      persentase
    };
  }, [recapRows]);

  // Text for Periode (Hanya Tanggal, Bulan dan Tahun tanpa teks Semester Ganjil / Genap)
  const periodeText = useMemo(() => {
    if (periodType === 'monthly') {
      try {
        const [y, m] = selectedMonth.split('-');
        const yearNum = parseInt(y, 10);
        const monthNum = parseInt(m, 10);
        const lastDay = new Date(yearNum, monthNum, 0).getDate();
        const startObj = new Date(yearNum, monthNum - 1, 1);
        const endObj = new Date(yearNum, monthNum - 1, lastDay);
        const startText = format(startObj, 'dd MMMM yyyy', { locale: id });
        const endText = format(endObj, 'dd MMMM yyyy', { locale: id });
        return `${startText} s/d ${endText}`;
      } catch {
        return selectedMonth;
      }
    } else if (periodType === 'custom') {
      try {
        const startText = startDate ? format(parseISO(startDate), 'dd MMMM yyyy', { locale: id }) : '-';
        const endText = endDate ? format(parseISO(endDate), 'dd MMMM yyyy', { locale: id }) : '-';
        if (startText === '-' && endText === '-') return '-';
        return startText === endText ? startText : `${startText} s/d ${endText}`;
      } catch {
        return `${startDate} s/d ${endDate}`;
      }
    } else {
      // Semester / Semua data: Ambil tanggal pertemuan pertama s/d terakhir
      if (studentSessions.length > 0) {
        try {
          const first = format(parseISO(studentSessions[0].date), 'dd MMMM yyyy', { locale: id });
          const last = format(parseISO(studentSessions[studentSessions.length - 1].date), 'dd MMMM yyyy', { locale: id });
          return first === last ? first : `${first} s/d ${last}`;
        } catch {
          return '-';
        }
      }
      return '-';
    }
  }, [periodType, selectedMonth, startDate, endDate, studentSessions]);

  // Signer Details Helper
  const getSignerDetails = useCallback((role: RecapSignerRoleType) => {
    switch (role) {
      case 'kurikulum':
        return {
          title: profileData?.jabatanKurikulum || 'Wakasek Kurikulum',
          name: profileData?.namaKurikulum || '(________________________)',
          nip: profileData?.nipKurikulum ? `NIP. ${profileData.nipKurikulum}` : '',
          enabled: true
        };
      case 'kesiswaan':
        return {
          title: profileData?.jabatanKesiswaan || 'Wakasek Kesiswaan',
          name: profileData?.namaKesiswaan || '(________________________)',
          nip: profileData?.nipKesiswaan ? `NIP. ${profileData.nipKesiswaan}` : '',
          enabled: true
        };
      case 'humas':
        return {
          title: profileData?.jabatanHumas || 'Wakasek Humas',
          name: profileData?.namaHumas || '(________________________)',
          nip: profileData?.nipHumas ? `NIP. ${profileData.nipHumas}` : '',
          enabled: true
        };
      case 'guru_wali':
        return {
          title: profileData?.jabatanGuruWali || 'Guru Wali',
          name: profileData?.namaGuruWali || '(________________________)',
          nip: profileData?.nipGuruWali ? `NIP. ${profileData.nipGuruWali}` : '',
          enabled: true
        };
      case 'guru_bk':
        return {
          title: profileData?.jabatanBK || 'Guru BK',
          name: profileData?.namaBK || '(________________________)',
          nip: profileData?.nipBK ? `NIP. ${profileData.nipBK}` : '',
          enabled: true
        };
      case 'guru_bk_2':
        return {
          title: 'Guru BK 2',
          name: profileData?.namaBK2 || '(________________________)',
          nip: profileData?.nipBK2 ? `NIP. ${profileData.nipBK2}` : '',
          enabled: true
        };
      case 'none':
        return { title: '', name: '', nip: '', enabled: false };
      case 'kepala_sekolah':
      default:
        return {
          title: 'Kepala Sekolah',
          name: profileData?.namaKepalaSekolah || '(________________________)',
          nip: profileData?.nipKepalaSekolah ? `NIP. ${profileData.nipKepalaSekolah}` : '',
          enabled: true
        };
    }
  }, [profileData]);

  const leftSigner = useMemo(() => getSignerDetails(leftSignerRole), [leftSignerRole, getSignerDetails]);
  const midSigner = useMemo(() => getSignerDetails(midSignerRole), [midSignerRole, getSignerDetails]);
  const is3Signers = midSigner.enabled && midSignerRole !== 'none';

  // Penentuan Tanda Tangan Kanan: Sesuai peran Wali Kelas atau Guru Mapel
  const resolvedRightRole = useMemo<'wali_kelas' | 'guru_mapel'>(() => {
    if (rightSignerType === 'wali_kelas') return 'wali_kelas';
    if (rightSignerType === 'guru_mapel') return 'guru_mapel';
    return isWaliKelas ? 'wali_kelas' : 'guru_mapel';
  }, [rightSignerType, isWaliKelas]);

  const rightSignerTitle = useMemo(() => {
    if (resolvedRightRole === 'wali_kelas') {
      const cls = selectedStudent?.class || selectedClass || profileData?.waliKelasClass || '';
      return `Wali Kelas ${cls}`.trim();
    }
    const mapel = (profileData?.mataPelajaran || '').trim();
    return mapel ? `Guru Mata Pelajaran ${mapel}` : 'Guru Mata Pelajaran';
  }, [resolvedRightRole, selectedStudent?.class, selectedClass, profileData?.waliKelasClass, profileData?.mataPelajaran]);

  // Nama Guru & NIP tanda tangan: Diambil langsung dari Nama Guru & NIP Guru yang ditulis di profil
  const teacherName = useMemo(() => {
    return (profileData?.namaGuruMapel || '').trim() || (profileData?.namaGuruWali || '').trim() || '(________________________)';
  }, [profileData?.namaGuruMapel, profileData?.namaGuruWali]);

  const teacherNIP = useMemo(() => {
    const nip = (profileData?.nipGuruMapel || '').trim() || (profileData?.nipGuruWali || '').trim();
    return nip ? `NIP. ${nip}` : '';
  }, [profileData?.nipGuruMapel, profileData?.nipGuruWali]);

  // Generate jsPDF Document
  const generateStudentPDF = (): { doc: jsPDF; fileName: string } => {
    if (!selectedStudent) {
      throw new Error('Siswa belum dipilih');
    }

    const doc = new jsPDF({ orientation: 'portrait', format: 'a4' });
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();

    // 1. Kop Surat
    const imgX = 14;
    const imgY = 8;
    const imgWidth = pageWidth - 28;
    const imgHeight = imgWidth * (341 / 1450);

    try {
      doc.addImage(kopSuratBase64, 'PNG', imgX, imgY, imgWidth, imgHeight);
    } catch (e) {
      console.error('Failed to add Kop Surat in student recap PDF', e);
    }

    // 2. Judul Laporan
    const titleY = imgY + imgHeight + 6;
    doc.setFontSize(13);
    doc.setFont('helvetica', 'bold');
    doc.text('REKAPITULASI KEHADIRAN SISWA', pageWidth / 2, titleY, { align: 'center' });

    // 3. Metadata Siswa & Periode (Dibawah Kop Surat)
    // Sebelah Kiri: Nama Lengkap Siswa, NIS, Kelas, Periode
    // Sebelah Kanan: Semester, Tahun Pelajaran
    doc.setFontSize(9.5);
    doc.setFont('helvetica', 'normal');

    // Posisi X kolom kiri yang presisi dan simetris
    const leftColX = 14;
    const leftColonX = 54;
    const leftValX = 57;

    // Posisi X kolom kanan yang presisi dan simetris
    const rightColX = 122;
    const rightColonX = 156;
    const rightValX = 159;

    const row1Y = titleY + 7;
    const row2Y = titleY + 13;
    const row3Y = titleY + 19;
    const row4Y = titleY + 25;

    // Row 1: Nama Lengkap Siswa | Semester
    doc.text('Nama Lengkap Siswa', leftColX, row1Y);
    doc.text(':', leftColonX, row1Y);
    doc.setFont('helvetica', 'bold');
    doc.text(selectedStudent.name, leftValX, row1Y);
    doc.setFont('helvetica', 'normal');

    doc.text('Semester', rightColX, row1Y);
    doc.text(':', rightColonX, row1Y);
    doc.setFont('helvetica', 'bold');
    doc.text(profileData?.semester || '-', rightValX, row1Y);
    doc.setFont('helvetica', 'normal');

    // Row 2: NIS | Tahun Pelajaran
    doc.text('NIS', leftColX, row2Y);
    doc.text(':', leftColonX, row2Y);
    doc.setFont('helvetica', 'bold');
    doc.text(selectedStudent.nisn || '-', leftValX, row2Y);
    doc.setFont('helvetica', 'normal');

    doc.text('Tahun Pelajaran', rightColX, row2Y);
    doc.text(':', rightColonX, row2Y);
    doc.setFont('helvetica', 'bold');
    doc.text(profileData?.tahunPelajaran || '-', rightValX, row2Y);
    doc.setFont('helvetica', 'normal');

    // Row 3: Kelas
    doc.text('Kelas', leftColX, row3Y);
    doc.text(':', leftColonX, row3Y);
    doc.setFont('helvetica', 'bold');
    doc.text(selectedStudent.class || selectedClass, leftValX, row3Y);
    doc.setFont('helvetica', 'normal');

    // Row 4: Periode
    doc.text('Periode', leftColX, row4Y);
    doc.text(':', leftColonX, row4Y);
    doc.setFont('helvetica', 'bold');
    doc.text(periodeText, leftValX, row4Y);
    doc.setFont('helvetica', 'normal');

    // 4. Tabel Presensi Persiswa:
    // Kolom: No | Tanggal Presensi | Keterangan | Persentase (di-merge sesuai jumlah baris per halaman)
    const tableStartY = row4Y + 6;
    const headers = [
      'No', 
      'Tanggal Presensi', 
      'Keterangan',
      'Persentase'
    ];

    // Pembagian rowSpan presisi agar tidak loncat ke halaman baru sebelum halaman 1 penuh:
    // Pada A4 portrait, halaman 1 mampu menampung hingga 26 baris sebelum batas margin bawah.
    // Halaman berikutnya mampu menampung hingga 38 baris.
    const PAGE_1_CAPACITY = 26;
    const PAGE_N_CAPACITY = 38;
    const spanBlocks: { startIndex: number; spanLength: number }[] = [];
    if (recapRows.length > 0) {
      if (recapRows.length <= PAGE_1_CAPACITY) {
        spanBlocks.push({ startIndex: 0, spanLength: recapRows.length });
      } else {
        spanBlocks.push({ startIndex: 0, spanLength: PAGE_1_CAPACITY });
        let remaining = recapRows.length - PAGE_1_CAPACITY;
        let currStart = PAGE_1_CAPACITY;
        while (remaining > 0) {
          const spanLen = Math.min(remaining, PAGE_N_CAPACITY);
          spanBlocks.push({ startIndex: currStart, spanLength: spanLen });
          currStart += spanLen;
          remaining -= spanLen;
        }
      }
    }

    const body = recapRows.map((r, idx) => {
      /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
      const rowArr: any[] = [
        r.no.toString(),
        r.formattedDate,
        r.status
      ];
      const block = spanBlocks.find(b => b.startIndex === idx);
      if (block) {
        rowArr.push({
          content: `${recapStats.persentase}%`,
          rowSpan: block.spanLength,
          styles: { 
            halign: 'center', 
            valign: 'middle', 
            fontStyle: 'bold', 
            fontSize: 10,
            fillColor: [250, 250, 250],
            textColor: [0, 0, 0],
            lineWidth: 0.2,
            lineColor: [0, 0, 0]
          }
        });
      }
      return rowArr;
    });

    const autoTableOptions = {
      startY: tableStartY,
      head: [headers],
      body: body,
      theme: 'grid' as const,
      rowPageBreak: 'auto' as const,
      pageBreak: 'auto' as const,
      margin: { left: 14, right: 14, top: 18, bottom: 14 },
      headStyles: {
        fillColor: [240, 240, 240],
        textColor: 0,
        fontStyle: 'bold',
        halign: 'center',
        lineWidth: 0.2,
        lineColor: [0, 0, 0],
        fontSize: 9
      },
      styles: {
        fontSize: 8.5,
        cellPadding: 2.2,
        lineWidth: 0.2,
        lineColor: [0, 0, 0],
        textColor: 0,
        valign: 'middle' as const
      },
      columnStyles: {
        0: { halign: 'center', cellWidth: 14 },
        1: { halign: 'left', cellWidth: 96 },
        2: { halign: 'center', cellWidth: 42 },
        3: { halign: 'center', valign: 'middle', cellWidth: 30 }
      },
      /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
      didParseCell: function(cellData: any) {
        if (cellData.section === 'body') {
          if (cellData.column.index === 0 || cellData.column.index === 2 || cellData.column.index === 3) {
            cellData.cell.styles.halign = 'center';
            cellData.cell.styles.valign = 'middle';
          }
          if (cellData.column.index === 2) {
            // Sesuai Request: warna text Keterangan hitam
            cellData.cell.styles.textColor = 0;
            cellData.cell.styles.fontStyle = 'bold';
          }
          if (cellData.column.index === 3) {
            cellData.cell.styles.textColor = 0;
            cellData.cell.styles.fontStyle = 'bold';
            cellData.cell.styles.halign = 'center';
            cellData.cell.styles.valign = 'middle';
          }
        }
      }
    };

    if (typeof autoTable === 'function') {
      /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
      autoTable(doc, autoTableOptions as any);
    /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
    } else if (typeof (doc as any).autoTable === 'function') {
      /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
      (doc as any).autoTable(autoTableOptions as any);
    }

    /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
    const tableFinalY = (doc as any).lastAutoTable?.finalY || 100;

    // 5. Tabel Rekapitulasi Statistik Kehadiran
    // Pastikan jika masih muat di halaman yang sama, tabel statistik tidak loncat ke halaman baru
    let statsY = tableFinalY + 5;
    const STATS_TABLE_HEIGHT = 16;
    if (statsY + STATS_TABLE_HEIGHT > pageHeight - 14) {
      doc.addPage();
      statsY = 20;
    }

    const statsHeaders = ['Hadir (H)', 'Sakit (S)', 'Izin (I)', 'Alpa (A)', 'Dispen (D)', 'Total Pertemuan', 'Persentase %'];
    const statsBody = [[
      `${recapStats.hadir} hari`,
      `${recapStats.sakit} hari`,
      `${recapStats.izin} hari`,
      `${recapStats.alpa} hari`,
      `${recapStats.dispen} hari`,
      `${recapStats.totalSessions} pertemuan`,
      `${recapStats.persentase}%`
    ]];

    const statsTableOptions = {
      startY: statsY,
      head: [statsHeaders],
      body: statsBody,
      theme: 'grid' as const,
      headStyles: {
        fillColor: [240, 240, 240],
        textColor: 0,
        fontStyle: 'bold',
        halign: 'center',
        lineWidth: 0.2,
        lineColor: [0, 0, 0],
        fontSize: 8.5
      },
      styles: {
        fontSize: 8.5,
        cellPadding: 2.2,
        lineWidth: 0.2,
        lineColor: [0, 0, 0],
        textColor: 0,
        halign: 'center'
      }
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
    const afterStatsY = (doc as any).lastAutoTable?.finalY || (statsY + 14);

    // 6. Tanda Tangan Formal (Tepat di bawah tabel)
    const SIG_HEIGHT = 42;
    let sigY = afterStatsY + 6;
    if (sigY + SIG_HEIGHT > pageHeight - 14) {
      doc.addPage();
      sigY = 20;
    }

    const leftSigner = getSignerDetails(leftSignerRole);
    const midSigner = getSignerDetails(midSignerRole);
    const is3Signers = midSigner.enabled && midSignerRole !== 'none';

    doc.setFontSize(9.5);
    doc.setFont('helvetica', 'normal');

    // Left signer
    if (leftSigner.enabled) {
      const leftX = 14;
      doc.text('Mengetahui,', leftX, sigY);
      doc.text(leftSigner.title, leftX, sigY + 5);
      doc.setFont('helvetica', 'bold');
      doc.text(leftSigner.name, leftX, sigY + 28);
      doc.setFont('helvetica', 'normal');
      if (leftSigner.nip) {
        doc.text(leftSigner.nip, leftX, sigY + 33);
      }
    }

    // Mid signer (if 3 signers enabled)
    if (is3Signers) {
      const midX = (pageWidth / 2) - 20;
      doc.text('Mengetahui,', midX, sigY);
      doc.text(midSigner.title, midX, sigY + 5);
      doc.setFont('helvetica', 'bold');
      doc.text(midSigner.name, midX, sigY + 28);
      doc.setFont('helvetica', 'normal');
      if (midSigner.nip) {
        doc.text(midSigner.nip, midX, sigY + 33);
      }
    }

    // Right signer (Wali Kelas / Guru Mapel) - Diambil dari Nama Guru di profil
    const rightX = pageWidth - 65;
    const dateFormattedNow = format(new Date(), 'dd MMMM yyyy', { locale: id });
    doc.text(`Cililin, ${dateFormattedNow}`, rightX, sigY);
    doc.text(rightSignerTitle, rightX, sigY + 5);
    doc.setFont('helvetica', 'bold');
    doc.text(teacherName, rightX, sigY + 28);
    doc.setFont('helvetica', 'normal');
    if (teacherNIP) {
      doc.text(teacherNIP, rightX, sigY + 33);
    }

    const safeName = selectedStudent.name.replace(/\s+/g, '_');
    const fileName = `Rekap_Presensi_${safeName}_${selectedStudent.class || selectedClass}.pdf`;

    return { doc, fileName };
  };

  // Export to PDF directly
  const handleExportPDF = () => {
    try {
      setIsExporting(true);
      const { doc, fileName } = generateStudentPDF();
      doc.save(fileName);
      showToast?.(`Dokumen ${fileName} berhasil diunduh.`, 'success');
    } catch (e) {
      console.error(e);
      showToast?.('Gagal mengekspor PDF rekap persiswa.', 'error');
    } finally {
      setIsExporting(false);
    }
  };

  // Export to Excel
  const handleExportExcel = () => {
    if (!selectedStudent) return;
    try {
      setIsExporting(true);
      const headerRows = [
        ['LAPORAN REKAPITULASI KEHADIRAN SISWA'],
        ['SMA NEGERI 1 CILILIN'],
        [],
        ['Nama Lengkap Siswa', ':', selectedStudent.name, '', 'Semester', ':', profileData?.semester || '-'],
        ['NIS', ':', selectedStudent.nisn || '-', '', 'Tahun Pelajaran', ':', profileData?.tahunPelajaran || '-'],
        ['Kelas', ':', selectedStudent.class || selectedClass],
        ['Periode', ':', periodeText],
        [],
        ['No', 'Tanggal Presensi', 'Keterangan', 'Persentase']
      ];

      const dataRows = recapRows.map((r, idx) => [
        r.no,
        r.formattedDate,
        r.status,
        idx === 0 ? `${recapStats.persentase}%` : ''
      ]);

      const statsRows = [
        [],
        ['REKAPITULASI KEHADIRAN'],
        ['Hadir', recapStats.hadir],
        ['Sakit', recapStats.sakit],
        ['Izin', recapStats.izin],
        ['Alpa', recapStats.alpa],
        ['Dispen', recapStats.dispen],
        ['Total Pertemuan', recapStats.totalSessions],
        ['Persentase Kehadiran', `${recapStats.persentase}%`]
      ];

      const dateFormattedNow = format(new Date(), 'dd MMMM yyyy', { locale: id });
      const sigRows = [
        [],
        [],
        [
          leftSigner.enabled ? 'Mengetahui,' : '',
          midSigner.enabled ? 'Mengetahui,' : '',
          '',
          `Cililin, ${dateFormattedNow}`
        ],
        [
          leftSigner.enabled ? leftSigner.title : '',
          midSigner.enabled ? midSigner.title : '',
          '',
          rightSignerTitle
        ],
        [],
        [],
        [],
        [
          leftSigner.enabled ? leftSigner.name : '',
          midSigner.enabled ? midSigner.name : '',
          '',
          teacherName
        ],
        [
          leftSigner.enabled ? (leftSigner.nip || '') : '',
          midSigner.enabled ? (midSigner.nip || '') : '',
          '',
          teacherNIP || ''
        ]
      ];

      const wsData = [...headerRows, ...dataRows, ...statsRows, ...sigRows];
      const worksheet = XLSX.utils.aoa_to_sheet(wsData);

      const startDataRow = 9;
      const merges = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 3 } },
        { s: { r: 1, c: 0 }, e: { r: 1, c: 3 } }
      ];
      if (recapRows.length > 0) {
        merges.push({
          s: { r: startDataRow, c: 3 },
          e: { r: startDataRow + recapRows.length - 1, c: 3 }
        });
      }
      worksheet['!merges'] = merges;

      worksheet['!cols'] = [
        { wch: 6 },
        { wch: 34 },
        { wch: 18 },
        { wch: 16 }
      ];

      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Rekap Presensi Siswa');

      const safeName = selectedStudent.name.replace(/\s+/g, '_');
      const fileName = `Rekap_Presensi_${safeName}_${selectedStudent.class || selectedClass}.xlsx`;
      XLSX.writeFile(workbook, fileName);

      showToast?.(`File Excel ${fileName} berhasil diekspor.`, 'success');
    } catch (e) {
      console.error(e);
      showToast?.('Gagal mengekspor file Excel.', 'error');
    } finally {
      setIsExporting(false);
    }
  };

  // Filter student dropdown by search keyword
  const filteredStudents = useMemo(() => {
    if (!studentSearch.trim()) return classStudents;
    const q = studentSearch.toLowerCase();
    return classStudents.filter(s => 
      s.name.toLowerCase().includes(q) || 
      (s.nisn && s.nisn.toLowerCase().includes(q))
    );
  }, [classStudents, studentSearch]);

  return (
    <div className="space-y-6">
      {/* Top Banner & Navigation */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 bg-white p-5 rounded-2xl border-2 border-slate-300 shadow-xs">
        <div>
          <div className="flex items-center gap-2 mb-1">
            {onBackToClassReport && (
              <button
                type="button"
                onClick={onBackToClassReport}
                className="p-1.5 hover:bg-slate-100 rounded-lg text-slate-600 transition-colors cursor-pointer mr-1"
                title="Kembali ke Rekap Kelas"
              >
                <ArrowLeft className="w-4 h-4" />
              </button>
            )}
            <span className="text-[10px] font-black uppercase tracking-wider px-2 py-0.5 rounded-md bg-amber-100 text-amber-800 border border-amber-200">
              Role Wali Kelas
            </span>
            <span className="text-xs font-bold text-slate-500">
              Kelas {selectedClass}
            </span>
            {isWaliKelas && profileData?.waliKelasClass && onBackToClassReport && (
              <button
                type="button"
                onClick={onBackToClassReport}
                className="ml-1.5 inline-flex items-center gap-1.5 text-xs font-bold text-emerald-800 bg-emerald-50 hover:bg-emerald-100 px-2.5 py-1 rounded-lg border border-emerald-300 transition-colors cursor-pointer shadow-2xs"
                title={`Kembali ke Laporan Rekap Kelas ${profileData.waliKelasClass}`}
              >
                <Building2 className="w-3.5 h-3.5 text-emerald-600" />
                <span>Kembali ke Kelas {profileData.waliKelasClass}</span>
              </button>
            )}
          </div>
          <h2 className="text-xl sm:text-2xl font-black text-slate-800 tracking-tight">
            Rekap Presensi Per Siswa
          </h2>
          <p className="text-xs text-slate-600 mt-0.5">
            Format resmi individual untuk cetak rekap kehadiran per siswa lengkap dengan Kop Surat &amp; Tanda Tangan.
          </p>
        </div>

        {/* Tombol Export Excel & Download PDF */}
        <div className="grid grid-cols-2 sm:flex sm:items-center gap-2.5 sm:gap-3 w-full sm:w-auto shrink-0">
          <button
            type="button"
            onClick={handleExportExcel}
            disabled={!selectedStudent || isExporting}
            className="h-10 px-4 sm:px-5 bg-emerald-600 hover:bg-emerald-700 active:bg-emerald-800 text-white rounded-xl font-bold text-xs sm:text-sm inline-flex items-center justify-center gap-2 shadow-xs hover:shadow-md border border-emerald-500/30 transition-all duration-150 active:scale-[0.98] disabled:opacity-50 disabled:cursor-not-allowed disabled:active:scale-100 disabled:shadow-none cursor-pointer"
            title="Export Rekap Kehadiran Siswa ke Excel (.xlsx)"
          >
            <FileSpreadsheet className="w-4 h-4 shrink-0 text-white" />
            <span className="whitespace-nowrap">Export Excel</span>
          </button>

          <button
            type="button"
            onClick={handleExportPDF}
            disabled={!selectedStudent || isExporting}
            className="h-10 px-4 sm:px-5 bg-rose-600 hover:bg-rose-700 active:bg-rose-800 text-white rounded-xl font-bold text-xs sm:text-sm inline-flex items-center justify-center gap-2 shadow-xs hover:shadow-md border border-rose-500/30 transition-all duration-150 active:scale-[0.98] disabled:opacity-50 disabled:cursor-not-allowed disabled:active:scale-100 disabled:shadow-none cursor-pointer"
            title="Download Dokumen Rekap Kehadiran Siswa Format PDF"
          >
            {isExporting ? (
              <Loader2 className="w-4 h-4 shrink-0 animate-spin text-white" />
            ) : (
              <FileDown className="w-4 h-4 shrink-0 text-white" />
            )}
            <span className="whitespace-nowrap">Download PDF</span>
          </button>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left Column: Filter Panel */}
        <div className="lg:col-span-1 space-y-5">
          <div className="bg-white p-5 rounded-2xl border-2 border-slate-300 shadow-xs space-y-4">
            <h3 className="text-sm font-black text-slate-800 uppercase tracking-wider pb-2 border-b border-slate-200 flex items-center gap-2">
              <UserIcon className="w-4 h-4 text-emerald-600" />
              Pilih Siswa &amp; Periode
            </h3>

            {/* 1. Pilih Kelas */}
            <div>
              <div className="flex items-center justify-between mb-1.5">
                <label className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                  <Building2 className="w-3.5 h-3.5 text-emerald-600" />
                  Kelas
                </label>
                {isWaliKelas && profileData?.waliKelasClass && (
                  <span className="text-[10px] font-bold text-emerald-800 bg-emerald-50 px-2 py-0.5 rounded-md border border-emerald-200">
                    Wali Kelas {profileData.waliKelasClass}
                  </span>
                )}
              </div>

              {/* Tombol Kelas yang diwali-kelasi & Rekap Persiswa agar tidak hilang */}
              {isWaliKelas && profileData?.waliKelasClass && (
                <div className="grid grid-cols-2 gap-1.5 mb-2 bg-slate-100 p-1 rounded-xl">
                  <button
                    type="button"
                    onClick={onBackToClassReport}
                    className="py-2 px-2 text-xs font-bold rounded-lg transition-all cursor-pointer text-slate-700 hover:bg-white hover:text-slate-900 flex items-center justify-center gap-1.5 border border-transparent hover:border-slate-200 shadow-none hover:shadow-xs"
                    title={`Pindah ke Laporan Rekap Kelas ${profileData.waliKelasClass}`}
                  >
                    <Building2 className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                    <span>Kelas {profileData.waliKelasClass}</span>
                  </button>
                  <button
                    type="button"
                    className="py-2 px-2 text-xs font-bold rounded-lg bg-emerald-600 text-white shadow-xs flex items-center justify-center gap-1.5 cursor-default"
                  >
                    <UserIcon className="w-3.5 h-3.5 shrink-0 text-white" />
                    <span>Rekap Persiswa</span>
                  </button>
                </div>
              )}

              <select
                value={selectedClass}
                onChange={(e) => setSelectedClass(e.target.value)}
                disabled={isWaliKelas && !!profileData?.waliKelasClass}
                className="w-full p-2.5 bg-slate-50 border-2 border-slate-300 rounded-xl font-bold text-xs text-slate-800 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-100 disabled:bg-slate-100 disabled:cursor-not-allowed"
              >
                {availableClasses.map(c => (
                  <option key={c} value={c}>{c}</option>
                ))}
              </select>
            </div>

            {/* 2. Pilih Siswa */}
            <div>
              <label className="block text-xs font-bold text-slate-700 mb-1.5">
                Pilih Siswa ({classStudents.length} Terdaftar)
              </label>

              {classStudents.length > 8 && (
                <div className="relative mb-2">
                  <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                  <input
                    type="text"
                    value={studentSearch}
                    onChange={(e) => setStudentSearch(e.target.value)}
                    placeholder="Cari nama atau NIS..."
                    className="w-full pl-8 pr-3 py-1.5 bg-slate-50 border border-slate-300 rounded-lg text-xs font-medium text-slate-700 focus:outline-none focus:border-emerald-600"
                  />
                </div>
              )}

              <select
                value={selectedStudentId}
                onChange={(e) => setSelectedStudentId(e.target.value)}
                className="w-full p-2.5 bg-slate-50 border-2 border-slate-300 rounded-xl font-bold text-xs text-slate-800 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-100"
              >
                {filteredStudents.length === 0 ? (
                  <option value="" disabled>-- Tidak ada siswa ditemukan --</option>
                ) : (
                  filteredStudents.map((s, idx) => (
                    <option key={s.id} value={s.id}>
                      {idx + 1}. {s.name} {s.nisn ? `(NIS: ${s.nisn})` : ''}
                    </option>
                  ))
                )}
              </select>
            </div>

            {/* 3. Periode Presensi */}
            <div>
              <label className="block text-xs font-bold text-slate-700 mb-1.5 flex items-center gap-1.5">
                <Calendar className="w-3.5 h-3.5 text-emerald-600" />
                Periode Rekapitulasi
              </label>
              <select
                value={periodType}
                onChange={(e) => setPeriodType(e.target.value as 'all' | 'monthly' | 'custom')}
                className="w-full p-2.5 bg-slate-50 border-2 border-slate-300 rounded-xl font-bold text-xs text-slate-800 focus:outline-none focus:border-emerald-600"
              >
                <option value="all">Semua Sesi / Semester ({profileData?.semester || 'Ganjil'})</option>
                <option value="monthly">Rekap Bulanan</option>
                <option value="custom">Rentang Tanggal (Kustom)</option>
              </select>

              {periodType === 'monthly' && (
                <div className="mt-2.5">
                  <label className="block text-[11px] font-bold text-slate-600 mb-1">Pilih Bulan</label>
                  <input
                    type="month"
                    value={selectedMonth}
                    onChange={(e) => setSelectedMonth(e.target.value)}
                    className="w-full p-2 bg-slate-50 border border-slate-300 rounded-lg text-xs font-bold text-slate-800 focus:outline-none focus:border-emerald-600"
                  />
                </div>
              )}

              {periodType === 'custom' && (
                <div className="mt-2.5 grid grid-cols-2 gap-2">
                  <div>
                    <label className="block text-[10px] font-bold text-slate-600 mb-1">Tanggal Awal</label>
                    <input
                      type="date"
                      value={startDate}
                      onChange={(e) => setStartDate(e.target.value)}
                      className="w-full p-2 bg-slate-50 border border-slate-300 rounded-lg text-xs font-bold text-slate-800 focus:outline-none focus:border-emerald-600"
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-bold text-slate-600 mb-1">Tanggal Akhir</label>
                    <input
                      type="date"
                      value={endDate}
                      onChange={(e) => setEndDate(e.target.value)}
                      className="w-full p-2 bg-slate-50 border border-slate-300 rounded-lg text-xs font-bold text-slate-800 focus:outline-none focus:border-emerald-600"
                    />
                  </div>
                </div>
              )}
            </div>

            {/* 4. Pengaturan Tanda Tangan */}
            <div className="pt-3 border-t border-slate-200 space-y-3">
              <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider">
                Opsi Penandatangan Dokumen
              </label>

              <div>
                <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                  1. Ttd Kiri (Mengetahui)
                </label>
                <select
                  value={leftSignerRole}
                  onChange={(e) => setLeftSignerRole(e.target.value as RecapSignerRoleType)}
                  className="w-full p-2 bg-slate-50 border border-slate-300 rounded-xl text-xs font-bold text-slate-800 focus:outline-none focus:border-emerald-600"
                >
                  <option value="kepala_sekolah">🏫 Kepala Sekolah {profileData?.namaKepalaSekolah ? `(${profileData.namaKepalaSekolah})` : ''}</option>
                  <option value="kurikulum">📚 Pihak Kurikulum {profileData?.namaKurikulum ? `(${profileData.namaKurikulum})` : ''}</option>
                  <option value="kesiswaan">👥 Pihak Kesiswaan {profileData?.namaKesiswaan ? `(${profileData.namaKesiswaan})` : ''}</option>
                  <option value="humas">📢 Pihak Humas {profileData?.namaHumas ? `(${profileData.namaHumas})` : ''}</option>
                  <option value="guru_bk">🤝 Guru BK {profileData?.namaBK ? `(${profileData.namaBK})` : ''}</option>
                  {profileData?.namaBK2 && (
                    <option value="guru_bk_2">🤝 Guru BK 2 ({profileData.namaBK2})</option>
                  )}
                  <option value="guru_wali">🎓 Guru Wali {profileData?.namaGuruWali ? `(${profileData.namaGuruWali})` : ''}</option>
                  <option value="none">➖ Tanpa Tanda Tangan Kiri</option>
                </select>
              </div>

              <div>
                <label className="block text-[11px] font-semibold text-slate-600 mb-1 flex items-center justify-between">
                  <span>2. Ttd Tengah (Bila Perlu)</span>
                  <span className="text-[9px] text-slate-400">3 TTD</span>
                </label>
                <select
                  value={midSignerRole}
                  onChange={(e) => setMidSignerRole(e.target.value as RecapSignerRoleType)}
                  className="w-full p-2 bg-slate-50 border border-slate-300 rounded-xl text-xs font-bold text-slate-800 focus:outline-none focus:border-emerald-600"
                >
                  <option value="none">➖ Tanpa Ttd Tengah (2 TTD)</option>
                  <option value="kepala_sekolah">🏫 Kepala Sekolah {profileData?.namaKepalaSekolah ? `(${profileData.namaKepalaSekolah})` : ''}</option>
                  <option value="kurikulum">📚 Pihak Kurikulum {profileData?.namaKurikulum ? `(${profileData.namaKurikulum})` : ''}</option>
                  <option value="kesiswaan">👥 Pihak Kesiswaan {profileData?.namaKesiswaan ? `(${profileData.namaKesiswaan})` : ''}</option>
                  <option value="humas">📢 Pihak Humas {profileData?.namaHumas ? `(${profileData.namaHumas})` : ''}</option>
                  <option value="guru_bk">🤝 Guru BK {profileData?.namaBK ? `(${profileData.namaBK})` : ''}</option>
                  {profileData?.namaBK2 && (
                    <option value="guru_bk_2">🤝 Guru BK 2 ({profileData.namaBK2})</option>
                  )}
                  <option value="guru_wali">🎓 Guru Wali {profileData?.namaGuruWali ? `(${profileData.namaGuruWali})` : ''}</option>
                </select>
              </div>

              <div>
                <label className="block text-[11px] font-semibold text-slate-600 mb-1 flex items-center justify-between">
                  <span>3. Ttd Kanan (Wali Kelas / Guru Mapel)</span>
                  <span className="text-[9px] text-emerald-700 bg-emerald-100 font-bold px-1.5 py-0.5 rounded">Profil</span>
                </label>
                <select
                  value={rightSignerType}
                  onChange={(e) => setRightSignerType(e.target.value as 'auto' | 'wali_kelas' | 'guru_mapel')}
                  className="w-full p-2 bg-slate-50 border border-slate-300 rounded-xl text-xs font-bold text-slate-800 focus:outline-none focus:border-emerald-600 mb-2"
                >
                  <option value="auto">
                    ⚙️ Otomatis Sesuai Profil ({isWaliKelas ? 'Wali Kelas' : 'Guru Mapel'})
                  </option>
                  <option value="wali_kelas">
                    🏫 Wali Kelas ({selectedStudent?.class || selectedClass})
                  </option>
                  <option value="guru_mapel">
                    📖 Guru Mata Pelajaran {profileData?.mataPelajaran ? `(${profileData.mataPelajaran})` : ''}
                  </option>
                </select>
              </div>

              <div className="p-2.5 bg-slate-50 rounded-xl border border-slate-200 text-[11px] text-slate-600 space-y-1">
                <div className="font-bold text-emerald-800 flex items-center justify-between">
                  <span>Penandatangan: {rightSignerTitle}</span>
                  <span className="text-[9px] bg-emerald-100 text-emerald-800 px-1.5 py-0.5 rounded font-bold">
                    Nama Guru Profil
                  </span>
                </div>
                <div className="font-bold text-slate-800">
                  {teacherName}
                </div>
                <div className="text-[10px] text-slate-500 font-mono">
                  {teacherNIP || 'NIP. -'}
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Right Column: Preview Rekapitulasi Persiswa */}
        <div className="lg:col-span-2 space-y-5">
          {!selectedStudent ? (
            <div className="bg-white p-12 rounded-2xl border-2 border-slate-300 text-center space-y-3">
              <Info className="w-8 h-8 text-slate-400 mx-auto" />
              <h4 className="font-bold text-slate-700 text-sm">Tidak Ada Siswa Terpilih</h4>
              <p className="text-xs text-slate-500">Silakan pilih kelas atau siswa pada filter di samping.</p>
            </div>
          ) : (
            /* Viewport Frame Pratinjau PDF dengan Scrolling Kiri-Kanan, Scrolling Atas-Bawah & Kontrol Zoom */
            <div className="bg-slate-200/90 rounded-2xl border-2 border-slate-300 shadow-inner flex flex-col max-h-[85vh] overflow-hidden">
              {/* Header Toolbar Kontrol Zoom (+ & -) */}
              <div className="bg-slate-800 text-white px-3 sm:px-4 py-2 flex items-center justify-between z-10 border-b border-slate-700/80 shadow-xs shrink-0">
                <div className="flex items-center gap-2">
                  <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></span>
                  <span className="text-xs font-bold text-slate-200">
                    Pratinjau A4 Portrait
                  </span>
                </div>

                {/* Tombol Zoom Out (-), Persentase, dan Zoom In (+) */}
                <div className="flex items-center gap-1 sm:gap-1.5 bg-slate-900/90 px-2 py-1 rounded-xl border border-slate-700 shadow-inner">
                  <button
                    type="button"
                    onClick={handleZoomOut}
                    disabled={zoomScale <= 0.5}
                    className="w-7 h-7 rounded-lg bg-slate-800 hover:bg-slate-700 active:bg-slate-600 disabled:opacity-40 disabled:hover:bg-slate-800 text-white flex items-center justify-center font-black text-sm transition-all cursor-pointer shadow-xs"
                    title="Zoom Out (-)"
                  >
                    <ZoomOut className="w-3.5 h-3.5" />
                  </button>

                  <button
                    type="button"
                    onClick={handleResetZoom}
                    className="px-2 py-0.5 text-xs font-mono font-bold text-slate-200 hover:text-white hover:bg-slate-800 rounded transition-colors cursor-pointer"
                    title="Reset Ukuran (100%)"
                  >
                    {Math.round(zoomScale * 100)}%
                  </button>

                  <button
                    type="button"
                    onClick={handleZoomIn}
                    disabled={zoomScale >= 1.6}
                    className="w-7 h-7 rounded-lg bg-slate-800 hover:bg-slate-700 active:bg-slate-600 disabled:opacity-40 disabled:hover:bg-slate-800 text-white flex items-center justify-center font-black text-sm transition-all cursor-pointer shadow-xs"
                    title="Zoom In (+)"
                  >
                    <ZoomIn className="w-3.5 h-3.5" />
                  </button>

                  {zoomScale !== 1 && (
                    <button
                      type="button"
                      onClick={handleResetZoom}
                      className="p-1 rounded-lg hover:bg-slate-800 text-slate-400 hover:text-white transition-colors cursor-pointer ml-0.5"
                      title="Reset ke 100%"
                    >
                      <RotateCcw className="w-3 h-3" />
                    </button>
                  )}
                </div>
              </div>

              {/* Area Canvas Dokumen dengan Scrolling Kiri-Kanan & Atas-Bawah */}
              <div className="flex-1 overflow-x-auto overflow-y-auto p-3 sm:p-6 overscroll-contain">
                <div
                  className="w-full flex justify-center py-2 transition-all duration-150"
                  style={{
                    minWidth: `${Math.round(794 * zoomScale)}px`,
                    minHeight: `${Math.round(1123 * zoomScale)}px`,
                  }}
                >
                  {/* LEMBAR DOKUMEN CETAK RESMI (Standar A4 Portrait: 210mm x 297mm ≈ 794px x 1123px, Margin 14mm ≈ 53px, Margin Atas 8mm ≈ 30px) */}
                  <div
                    style={{
                      transform: `scale(${zoomScale})`,
                      transformOrigin: 'top center',
                    }}
                    className="w-[794px] min-w-[794px] min-h-[1123px] bg-white px-[53px] pt-[30px] pb-[48px] rounded-xl shadow-2xl border border-slate-300 text-slate-900 font-sans shrink-0 transition-transform duration-150 print:shadow-none print:border-none print:p-0"
                  >
                    {/* 1. Kop Surat Resmi Sekolah (Proporsi & ukuran presisi sama seperti dokumen PDF yang diunduh: 182mm x 42.8mm) */}
                    <div className="w-full text-center">
                      <img
                        src={kopSuratBase64}
                        alt="Kop Surat Resmi SMA Negeri 1 Cililin"
                        className="w-full block mx-auto object-contain"
                        style={{
                          width: '100%',
                          maxWidth: '688px',
                          height: 'auto',
                          aspectRatio: '1450 / 341',
                        }}
                      />
                    </div>

                    {/* 2. Judul Laporan (Garis bawah dihilangkan sesuai permintaan) */}
                    <div className="text-center mt-5">
                      <h3 className="text-[15px] font-black tracking-wider uppercase text-black leading-tight">
                        REKAPITULASI KEHADIRAN SISWA
                      </h3>
                    </div>

                    {/* 3. Metadata Siswa & Periode di Bawah Kop Surat */}
                    {/* Sebelah Kiri: Nama Lengkap Siswa, NIS, Kelas, Periode | Sebelah Kanan: Semester, Tahun Pelajaran */}
                    <div className="flex justify-between items-start gap-6 text-[12.5px] mt-4">
                      {/* Sebelah Kiri (Titik Dua Simetris & Presisi) */}
                      <div className="space-y-1">
                        <div className="flex items-center">
                          <span className="w-[144px] font-semibold text-slate-700 shrink-0">Nama Lengkap Siswa</span>
                          <span className="w-[16px] text-center font-bold text-black shrink-0">:</span>
                          <span className="font-bold text-black flex-1">{selectedStudent.name}</span>
                        </div>
                        <div className="flex items-center">
                          <span className="w-[144px] font-semibold text-slate-700 shrink-0">NIS</span>
                          <span className="w-[16px] text-center font-bold text-black shrink-0">:</span>
                          <span className="font-bold text-black font-mono flex-1">{selectedStudent.nisn || '-'}</span>
                        </div>
                        <div className="flex items-center">
                          <span className="w-[144px] font-semibold text-slate-700 shrink-0">Kelas</span>
                          <span className="w-[16px] text-center font-bold text-black shrink-0">:</span>
                          <span className="font-bold text-black flex-1">{selectedStudent.class || selectedClass}</span>
                        </div>
                        <div className="flex items-center">
                          <span className="w-[144px] font-semibold text-slate-700 shrink-0">Periode</span>
                          <span className="w-[16px] text-center font-bold text-black shrink-0">:</span>
                          <span className="font-bold text-black flex-1">{periodeText}</span>
                        </div>
                      </div>

                      {/* Sebelah Kanan (Semester & Tahun Pelajaran - Titik Dua Simetris & Presisi) */}
                      <div className="space-y-1 min-w-[240px] text-left">
                        <div className="flex items-center">
                          <span className="w-[125px] font-semibold text-slate-700 shrink-0">Semester</span>
                          <span className="w-[16px] text-center font-bold text-black shrink-0">:</span>
                          <span className="font-bold text-black flex-1">{profileData?.semester || '-'}</span>
                        </div>
                        <div className="flex items-center">
                          <span className="w-[125px] font-semibold text-slate-700 shrink-0">Tahun Pelajaran</span>
                          <span className="w-[16px] text-center font-bold text-black shrink-0">:</span>
                          <span className="font-bold text-black flex-1">{profileData?.tahunPelajaran || '-'}</span>
                        </div>
                      </div>
                    </div>

                    {/* 4. Tabel Presensi Persiswa */}
                    {/* Kolom: No | Tanggal Presensi | Keterangan (warna hitam) | Persentase (di-merge rapih presisi dan simetris) */}
                    <div className="overflow-x-auto mt-4">
                      <table className="w-full text-xs sm:text-[13px] border-collapse border border-black table-fixed">
                        <colgroup>
                          <col style={{ width: '7.7%' }} />
                          <col style={{ width: '52.7%' }} />
                          <col style={{ width: '23.1%' }} />
                          <col style={{ width: '16.5%' }} />
                        </colgroup>
                        <thead>
                          <tr className="bg-slate-100 text-black border-b border-black text-center font-bold">
                            <th className="border border-black py-2 px-2 text-center">No</th>
                            <th className="border border-black py-2 px-3 text-left">Tanggal Presensi</th>
                            <th className="border border-black py-2 px-3 text-center">Keterangan</th>
                            <th className="border border-black py-2 px-3 text-center">Persentase</th>
                          </tr>
                        </thead>
                    <tbody>
                      {recapRows.length === 0 ? (
                        <tr>
                          <td colSpan={4} className="border border-black py-6 text-center text-slate-500 italic">
                            Belum ada catatan presensi pada periode yang dipilih.
                          </td>
                        </tr>
                      ) : (
                        recapRows.map((row, idx) => (
                          <tr key={row.no} className="hover:bg-slate-50/60">
                            <td className="border border-black py-1.5 px-2 text-center font-medium">{row.no}</td>
                            <td className="border border-black py-1.5 px-3 text-black font-medium">
                              {row.formattedDate}
                              {row.meetingNumber && (
                                <span className="text-[11px] text-slate-500 ml-1.5 font-normal">
                                  (Pertemuan {row.meetingNumber})
                                </span>
                              )}
                            </td>
                            {/* Keterangan: warna teks hitam */}
                            <td className="border border-black py-1.5 px-3 text-center font-bold text-black">
                              {row.status}
                            </td>
                            {idx === 0 && (
                              <td
                                rowSpan={recapRows.length}
                                className="border border-black p-2 text-center align-middle font-bold text-base bg-slate-50/70 text-black select-none"
                                style={{ verticalAlign: 'middle', textAlign: 'center' }}
                              >
                                <div className="flex flex-col items-center justify-center h-full min-h-[48px] text-center leading-normal">
                                  <span className="text-base sm:text-lg font-black text-black tracking-wide">
                                    {recapStats.persentase}%
                                  </span>
                                </div>
                              </td>
                            )}
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>

                {/* 5. Tabel Rekapitulasi Statistik Kehadiran */}
                <div className="overflow-x-auto pt-1">
                  <table className="w-full text-xs sm:text-[12px] border-collapse border border-black text-center">
                    <thead>
                      <tr className="bg-slate-100 font-bold border-b border-black text-black">
                        <th className="border border-black py-1.5 px-2">Hadir (H)</th>
                        <th className="border border-black py-1.5 px-2">Sakit (S)</th>
                        <th className="border border-black py-1.5 px-2">Izin (I)</th>
                        <th className="border border-black py-1.5 px-2">Alpa (A)</th>
                        <th className="border border-black py-1.5 px-2">Dispen (D)</th>
                        <th className="border border-black py-1.5 px-2">Total Pertemuan</th>
                        <th className="border border-black py-1.5 px-2">Persentase %</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr className="font-bold text-black">
                        <td className="border border-black py-2 px-2">{recapStats.hadir} hari</td>
                        <td className="border border-black py-2 px-2">{recapStats.sakit} hari</td>
                        <td className="border border-black py-2 px-2">{recapStats.izin} hari</td>
                        <td className="border border-black py-2 px-2">{recapStats.alpa} hari</td>
                        <td className="border border-black py-2 px-2">{recapStats.dispen} hari</td>
                        <td className="border border-black py-2 px-2">{recapStats.totalSessions} pertemuan</td>
                        <td className="border border-black py-2 px-2">{recapStats.persentase}%</td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                {/* 6. Tanda Tangan Formal (Tepat di bawah tabel presensi & rekapitulasi) */}
                <div className="pt-5 sm:pt-6 flex justify-between items-start text-xs sm:text-sm gap-6">
                  {/* Left Signer */}
                  {leftSigner.enabled ? (
                    <div className="space-y-1 text-center min-w-[200px]">
                      <p className="text-black">Mengetahui,</p>
                      <p className="font-medium text-black">{leftSigner.title}</p>
                      <div className="h-16"></div>
                      <p className="font-bold underline text-black">{leftSigner.name}</p>
                      {leftSigner.nip && <p className="text-[11px] text-slate-700">{leftSigner.nip}</p>}
                    </div>
                  ) : null}

                  {/* Mid Signer (if enabled) */}
                  {is3Signers ? (
                    <div className="space-y-1 text-center min-w-[200px]">
                      <p className="text-black">Mengetahui,</p>
                      <p className="font-medium text-black">{midSigner.title}</p>
                      <div className="h-16"></div>
                      <p className="font-bold underline text-black">{midSigner.name}</p>
                      {midSigner.nip && <p className="text-[11px] text-slate-700">{midSigner.nip}</p>}
                    </div>
                  ) : null}

                  {/* Right Signer (Wali Kelas / Guru Mapel) - Posisikan pasti di sebelah kanan */}
                  <div className="space-y-1 text-center min-w-[220px] ml-auto">
                    <p className="text-black">Cililin, {format(new Date(), 'dd MMMM yyyy', { locale: id })}</p>
                    <p className="font-medium text-black">{rightSignerTitle}</p>
                    <div className="h-16"></div>
                    <p className="font-bold underline text-black">
                      {teacherName}
                    </p>
                    <p className="text-[11px] text-slate-700">
                      {teacherNIP || 'NIP. -'}
                    </p>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}
        </div>
      </div>
    </div>
  );
}

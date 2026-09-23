import React, { useState, useMemo, useEffect, useCallback, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Building2,
  User as UserIcon,
  Clock,
  Activity,
  Save,
  Pencil,
  FileText,
  Trash2,
  Printer,
  Calendar,
  Users,
  CheckCircle2,
  PhoneCall,
  UserCheck,
  ShieldAlert,
  Layers,
  X,
  ChevronDown,
  Search,
  Check,
  Plus,
  Eye
} from 'lucide-react';
import { collection, query, where, getDocs, doc, setDoc, deleteDoc, serverTimestamp, Firestore } from 'firebase/firestore';
import { Auth } from 'firebase/auth';
import { handleFirestoreError, OperationType } from './firebase';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { kopSuratBase64 } from './kop-surat-b64';

interface Student {
  id: string;
  name: string;
  class: string;
}

interface AttendanceSession {
  id: string;
  date: string;
  className: string;
  records: Record<string, string>;
}

interface FollowupReport {
  id: string;
  studentId: string;
  className: string;
  terlambat: string;
  permasalahan: string;
  tindakLanjut: string;
  status: string;
  month: string; // YYYY-MM
  date?: string; // YYYY-MM-DD
}

export interface HomeroomMonthlyRecap {
  id: string;
  className: string;
  month: string; // YYYY-MM
  startDate?: string;
  endDate?: string;
  pembinaanMurid: string[] | number;
  kontakOrangTua: string[] | number;
  koordinasiBK: string[] | number;
  koordinasiGuruWali: string[] | number;
  koordinasiGuruMapel: string[] | number;
  kasusBelumSelesai: string[] | number;
  pembinaanMuridNames?: string[];
  kontakOrangTuaNames?: string[];
  koordinasiBKNames?: string[];
  koordinasiGuruWaliNames?: string[];
  koordinasiGuruMapelNames?: string[];
  kasusBelumSelesaiNames?: string[];
  keterangan?: string;
  userId?: string;
  updatedAt?: unknown;
}

interface HomeroomReportViewProps {
  classList: string[];
  students: Student[];
  attendanceSessions: AttendanceSession[];
  activeDb?: Firestore;
  activeAuth?: Auth;
  trackOp?: (type: 'read' | 'write', count?: number) => void;
  showToast?: (message: string, type: 'success' | 'info' | 'error') => void;
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
    jumlahSiswaLakiLaki?: number | string;
    jumlahSiswaPerempuan?: number | string;
  };
}

export type HomeroomSignerRoleType = 'kepala_sekolah' | 'kurikulum' | 'kesiswaan' | 'humas' | 'guru_wali' | 'guru_bk' | 'none';

const parseStudentIds = (val: string[] | number | undefined | null): string[] => {
  if (Array.isArray(val)) return val;
  return [];
};

const getCategoryCount = (val: string[] | number | undefined | null): number => {
  if (Array.isArray(val)) return val.length;
  if (typeof val === 'number') return val;
  return 0;
};

interface StudentMultiSelectDropdownProps {
  label: string;
  icon: React.ReactNode;
  themeColor: 'indigo' | 'emerald' | 'amber' | 'sky' | 'violet' | 'rose';
  description: string;
  students: Student[];
  selectedIds: string[];
  onChange: (ids: string[]) => void;
  placeholder?: string;
}

function StudentMultiSelectDropdown({
  label,
  icon,
  themeColor,
  description,
  students,
  selectedIds,
  onChange,
  placeholder = 'Pilih nama siswa...'
}: StudentMultiSelectDropdownProps) {
  const [isOpen, setIsOpen] = useState(false);
  const [search, setSearch] = useState('');
  const dropdownRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    }
    if (isOpen) {
      document.addEventListener('mousedown', handleClickOutside);
    }
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [isOpen]);

  const filteredStudents = useMemo(() => {
    if (!search.trim()) return students;
    const q = search.toLowerCase();
    return students.filter(s => s.name.toLowerCase().includes(q) || (s.id && s.id.toLowerCase().includes(q)));
  }, [students, search]);

  const selectedStudents = useMemo(() => {
    return selectedIds.map(id => students.find(s => s.id === id || s.name === id) || { id, name: id, class: '' });
  }, [selectedIds, students]);

  const colorStyles = {
    indigo: {
      border: 'border-indigo-200 focus-within:border-indigo-500',
      badge: 'bg-indigo-100 text-indigo-800 border-indigo-200',
      chip: 'bg-indigo-50 text-indigo-800 border-indigo-200 hover:bg-indigo-100',
      activeItem: 'bg-indigo-50/80 text-indigo-900 font-bold',
      checkbox: 'text-indigo-600 focus:ring-indigo-500',
      btn: 'bg-indigo-600 hover:bg-indigo-700 text-white'
    },
    emerald: {
      border: 'border-emerald-200 focus-within:border-emerald-500',
      badge: 'bg-emerald-100 text-emerald-800 border-emerald-200',
      chip: 'bg-emerald-50 text-emerald-800 border-emerald-200 hover:bg-emerald-100',
      activeItem: 'bg-emerald-50/80 text-emerald-900 font-bold',
      checkbox: 'text-emerald-600 focus:ring-emerald-500',
      btn: 'bg-emerald-600 hover:bg-emerald-700 text-white'
    },
    amber: {
      border: 'border-amber-200 focus-within:border-amber-500',
      badge: 'bg-amber-100 text-amber-800 border-amber-200',
      chip: 'bg-amber-50 text-amber-800 border-amber-200 hover:bg-amber-100',
      activeItem: 'bg-amber-50/80 text-amber-900 font-bold',
      checkbox: 'text-amber-600 focus:ring-amber-500',
      btn: 'bg-amber-600 hover:bg-amber-700 text-white'
    },
    sky: {
      border: 'border-sky-200 focus-within:border-sky-500',
      badge: 'bg-sky-100 text-sky-800 border-sky-200',
      chip: 'bg-sky-50 text-sky-800 border-sky-200 hover:bg-sky-100',
      activeItem: 'bg-sky-50/80 text-sky-900 font-bold',
      checkbox: 'text-sky-600 focus:ring-sky-500',
      btn: 'bg-sky-600 hover:bg-sky-700 text-white'
    },
    violet: {
      border: 'border-violet-200 focus-within:border-violet-500',
      badge: 'bg-violet-100 text-violet-800 border-violet-200',
      chip: 'bg-violet-50 text-violet-800 border-violet-200 hover:bg-violet-100',
      activeItem: 'bg-violet-50/80 text-violet-900 font-bold',
      checkbox: 'text-violet-600 focus:ring-violet-500',
      btn: 'bg-violet-600 hover:bg-violet-700 text-white'
    },
    rose: {
      border: 'border-rose-200 focus-within:border-rose-500',
      badge: 'bg-rose-100 text-rose-800 border-rose-200',
      chip: 'bg-rose-50 text-rose-800 border-rose-200 hover:bg-rose-100',
      activeItem: 'bg-rose-50/80 text-rose-900 font-bold',
      checkbox: 'text-rose-600 focus:ring-rose-500',
      btn: 'bg-rose-600 hover:bg-rose-700 text-white'
    }
  }[themeColor];

  const toggleStudent = (id: string) => {
    if (selectedIds.includes(id)) {
      onChange(selectedIds.filter(i => i !== id));
    } else {
      onChange([...selectedIds, id]);
    }
  };

  const removeStudent = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    onChange(selectedIds.filter(i => i !== id));
  };

  const selectAll = () => {
    const allFilteredIds = filteredStudents.map(s => s.id);
    const merged = Array.from(new Set([...selectedIds, ...allFilteredIds]));
    onChange(merged);
  };

  const clearAll = () => {
    onChange([]);
  };

  return (
    <div className="space-y-1.5 relative" ref={dropdownRef}>
      <label className="text-xs font-bold text-slate-700 flex items-center justify-between">
        <span className="flex items-center gap-1.5">
          {icon}
          {label}
        </span>
        <span className={`text-[10px] font-bold px-2 py-0.5 rounded-full border ${colorStyles.badge}`}>
          {selectedIds.length} Siswa
        </span>
      </label>

      {/* Main trigger button / box */}
      <div 
        onClick={() => setIsOpen(!isOpen)}
        className={`w-full min-h-[46px] p-2 bg-white border-2 rounded-xl text-slate-800 cursor-pointer transition-all flex flex-wrap items-center justify-between gap-1.5 shadow-2xs ${
          isOpen ? 'ring-2 ring-indigo-500/20 ' + colorStyles.border : 'border-slate-200 hover:border-slate-300'
        }`}
      >
        <div className="flex flex-wrap items-center gap-1 flex-1 min-w-0">
          {selectedStudents.length === 0 ? (
            <span className="text-xs text-slate-400 font-normal px-1 select-none flex items-center gap-1.5">
              <Plus className="w-3.5 h-3.5 text-slate-400" />
              {placeholder}
            </span>
          ) : (
            selectedStudents.map(s => (
              <span 
                key={s.id}
                className={`inline-flex items-center gap-1 text-[11px] font-semibold px-2 py-0.5 rounded-lg border ${colorStyles.chip} transition-all`}
              >
                <span className="max-w-[130px] truncate">{s.name}</span>
                <button
                  type="button"
                  onClick={(e) => removeStudent(s.id, e)}
                  className="p-0.5 rounded-full hover:bg-black/10 transition-colors"
                  title="Hapus siswa ini"
                >
                  <X className="w-3 h-3" />
                </button>
              </span>
            ))
          )}
        </div>
        <div className="flex items-center pl-1 text-slate-400 shrink-0">
          <ChevronDown className={`w-4 h-4 transition-transform duration-200 ${isOpen ? 'rotate-180 text-slate-600' : ''}`} />
        </div>
      </div>

      {/* Helper description */}
      <p className="text-[10px] text-slate-400 leading-tight truncate" title={description}>
        {description}
      </p>

      {/* Dropdown Menu */}
      <AnimatePresence>
        {isOpen && (
          <motion.div
            initial={{ opacity: 0, y: 5, scale: 0.98 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, y: 5, scale: 0.98 }}
            transition={{ duration: 0.15 }}
            className="absolute left-0 right-0 top-full mt-1.5 z-50 bg-white rounded-2xl shadow-xl border border-slate-200 overflow-hidden flex flex-col max-h-[300px]"
          >
            {/* Search header & Quick Actions */}
            <div className="p-2.5 bg-slate-50 border-b border-slate-100 space-y-2 shrink-0">
              <div className="relative">
                <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                <input 
                  type="text"
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  placeholder="Ketik untuk mencari nama siswa..."
                  className="w-full pl-8 pr-3 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-medium text-slate-700 outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500"
                  autoFocus
                  onClick={(e) => e.stopPropagation()}
                />
              </div>
              <div className="flex items-center justify-between text-[10px] font-bold px-0.5">
                <button
                  type="button"
                  onClick={(e) => { e.stopPropagation(); selectAll(); }}
                  className="text-indigo-600 hover:text-indigo-800 transition-colors"
                >
                  + Pilih Semua ({filteredStudents.length})
                </button>
                {selectedIds.length > 0 && (
                  <button
                    type="button"
                    onClick={(e) => { e.stopPropagation(); clearAll(); }}
                    className="text-rose-600 hover:text-rose-800 transition-colors"
                  >
                    Kosongkan Pilihan
                  </button>
                )}
              </div>
            </div>

            {/* Students Checklist */}
            <div className="overflow-y-auto p-1.5 divide-y divide-slate-50 flex-1 scrollbar-thin">
              {filteredStudents.length === 0 ? (
                <div className="py-6 text-center text-xs text-slate-400 font-medium">
                  Tidak ada nama siswa yang cocok
                </div>
              ) : (
                filteredStudents.map(student => {
                  const isSelected = selectedIds.includes(student.id);
                  return (
                    <div
                      key={student.id}
                      onClick={() => toggleStudent(student.id)}
                      className={`flex items-center justify-between p-2 rounded-xl cursor-pointer text-xs transition-colors ${
                        isSelected ? colorStyles.activeItem : 'hover:bg-slate-50 text-slate-700'
                      }`}
                    >
                      <div className="flex items-center gap-2.5 min-w-0 pr-2">
                        <div className={`w-4 h-4 rounded border flex items-center justify-center transition-colors shrink-0 ${
                          isSelected 
                            ? 'bg-indigo-600 border-indigo-600 text-white' 
                            : 'border-slate-300 bg-white'
                        }`}>
                          {isSelected && <Check className="w-3 h-3 stroke-[3]" />}
                        </div>
                        <span className="font-medium truncate">{student.name}</span>
                      </div>
                      {isSelected && (
                        <span className="text-[9px] font-bold uppercase tracking-wider text-indigo-600 shrink-0">
                          Terpilih
                        </span>
                      )}
                    </div>
                  );
                })
              )}
            </div>

            {/* Footer */}
            <div className="p-2 bg-slate-50 border-t border-slate-100 flex items-center justify-between shrink-0">
              <span className="text-[11px] font-bold text-slate-600">
                {selectedIds.length} dari {students.length} siswa dipilih
              </span>
              <button
                type="button"
                onClick={() => setIsOpen(false)}
                className={`px-3 py-1 rounded-lg text-xs font-bold ${colorStyles.btn} transition-colors shadow-2xs`}
              >
                Selesai
              </button>
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

export default function HomeroomReportView({
  classList,
  students,
  attendanceSessions,
  activeDb,
  activeAuth,
  trackOp,
  showToast,
  profileData
}: HomeroomReportViewProps) {
  const [selectedClass, setSelectedClass] = useState<string>(() => profileData?.waliKelasClass || '');
  const [selectedDate, setSelectedDate] = useState<string>(new Date().toISOString().substring(0, 10));
  const [selectedStudentId, setSelectedStudentId] = useState<string>('');
  const [terlambat, setTerlambat] = useState<string>('Tidak Pernah');
  const [permasalahan, setPermasalahan] = useState<string>('');
  const [tindakLanjut, setTindakLanjut] = useState<string>('');
  const [status, setStatus] = useState<string>('Normal');
  const [isEditing, setIsEditing] = useState<boolean>(true);
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [existingDocId, setExistingDocId] = useState<string>('');

  const [selectedMonth, setSelectedMonth] = useState<string>(new Date().toISOString().substring(0, 7));
  const currentMonth = selectedMonth;
  const [leftSignerRole, setLeftSignerRole] = useState<HomeroomSignerRoleType>('kepala_sekolah');
  const [midSignerRole, setMidSignerRole] = useState<HomeroomSignerRoleType>('none');

  // Preview Modal States
  const [showPreviewModal, setShowPreviewModal] = useState<boolean>(false);
  const [previewPdfUrl, setPreviewPdfUrl] = useState<string | null>(null);
  const [previewFileName, setPreviewFileName] = useState<string>('');
  const [isPreviewLoading, setIsPreviewLoading] = useState<boolean>(false);

  const getSignerInfo = (role: HomeroomSignerRoleType) => {
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
      case 'none':
        return {
          title: '',
          name: '',
          nip: '',
          enabled: false
        };
      case 'kepala_sekolah':
      default:
        return {
          title: 'Kepala Sekolah',
          name: profileData?.namaKepalaSekolah || '(________________________)',
          nip: profileData?.nipKepalaSekolah ? `NIP. ${profileData.nipKepalaSekolah}` : '',
          enabled: true
        };
    }
  };

  const getLeftSignerInfo = () => getSignerInfo(leftSignerRole);
  const getMidSignerInfo = () => getSignerInfo(midSignerRole);

  // Manual date range (Tanggal Awal & Tanggal Akhir)
  const [startDate, setStartDate] = useState<string>(() => {
    const ym = new Date().toISOString().substring(0, 7);
    return `${ym}-01`;
  });
  const [endDate, setEndDate] = useState<string>(() => {
    const now = new Date();
    const y = now.getFullYear();
    const m = now.getMonth() + 1;
    const lastDay = new Date(y, m, 0).getDate();
    return `${y}-${String(m).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;
  });

  const printScope: 'all_months' | 'specific_month' = 'specific_month';
  const [exportMonth, setExportMonth] = useState<string>(currentMonth);

  // Sync date range when selectedMonth changes
  useEffect(() => {
    setExportMonth(selectedMonth);
    const [yStr, mStr] = selectedMonth.split('-');
    const y = parseInt(yStr, 10) || new Date().getFullYear();
    const m = parseInt(mStr, 10) || (new Date().getMonth() + 1);
    const lastDay = new Date(y, m, 0).getDate();
    setStartDate(`${selectedMonth}-01`);
    setEndDate(`${selectedMonth}-${String(lastDay).padStart(2, '0')}`);
  }, [selectedMonth]);

  const MONTH_OPTIONS = useMemo(() => [
    { value: '01', label: 'Januari' },
    { value: '02', label: 'Februari' },
    { value: '03', label: 'Maret' },
    { value: '04', label: 'April' },
    { value: '05', label: 'Mei' },
    { value: '06', label: 'Juni' },
    { value: '07', label: 'Juli' },
    { value: '08', label: 'Agustus' },
    { value: '09', label: 'September' },
    { value: '10', label: 'Oktober' },
    { value: '11', label: 'November' },
    { value: '12', label: 'Desember' }
  ], []);

  const availableYears = useMemo(() => {
    const yearsSet = new Set<string>();
    const currentY = new Date().getFullYear();
    for (let y = currentY - 2; y <= currentY + 3; y++) {
      yearsSet.add(String(y));
    }
    attendanceSessions.forEach(s => {
      if (s.date && s.date.length >= 4) {
        yearsSet.add(s.date.substring(0, 4));
      }
    });
    return Array.from(yearsSet).sort().reverse();
  }, [attendanceSessions]);

  // Compute class accumulation stats based on manual date range
  const classAccumulationStats = useMemo(() => {
    if (!selectedClass) return { totalSiswa: 0, totalSesi: 0, hadir: 0, sakit: 0, izin: 0, alpa: 0, dispen: 0, hadirPct: 0, sakitPct: 0, izinPct: 0, alpaPct: 0, dispenPct: 0, persentase: 0 };
    const classStudents = students.filter(s => s.class === selectedClass);
    const classSessions = attendanceSessions.filter(s => {
      if (s.className !== selectedClass) return false;
      if (printScope === 'specific_month') {
        if (startDate && s.date < startDate) return false;
        if (endDate && s.date > endDate) return false;
        return true;
      }
      return true;
    });

    let hadir = 0, sakit = 0, izin = 0, alpa = 0, dispen = 0;
    classSessions.forEach(session => {
      Object.values(session.records || {}).forEach(status => {
        if (status === 'Hadir') hadir++;
        else if (status === 'Sakit') sakit++;
        else if (status === 'Izin') izin++;
        else if (status === 'Alpa') alpa++;
        else if (status === 'Dispen') dispen++;
      });
    });

    const totalRecords = hadir + sakit + izin + alpa + dispen;
    const effectiveHadir = hadir + dispen;
    const hadirPct = totalRecords > 0 ? (effectiveHadir / totalRecords) * 100 : 0;
    const sakitPct = totalRecords > 0 ? (sakit / totalRecords) * 100 : 0;
    const izinPct = totalRecords > 0 ? (izin / totalRecords) * 100 : 0;
    const alpaPct = totalRecords > 0 ? (alpa / totalRecords) * 100 : 0;
    const dispenPct = totalRecords > 0 ? (dispen / totalRecords) * 100 : 0;
    const persentase = hadirPct;

    return {
      totalSiswa: classStudents.length,
      totalSesi: classSessions.length,
      hadir,
      sakit,
      izin,
      alpa,
      dispen,
      hadirPct,
      sakitPct,
      izinPct,
      alpaPct,
      dispenPct,
      persentase
    };
  }, [selectedClass, students, attendanceSessions, printScope, startDate, endDate]);

  // Compute priority students based on chosen date range
  const priorityStudents = useMemo(() => {
    if (!selectedClass) return [];
    
    const classStudents = students.filter(s => s.class === selectedClass);
    const classSessions = attendanceSessions.filter(s => {
      if (s.className !== selectedClass) return false;
      if (startDate && s.date < startDate) return false;
      if (endDate && s.date > endDate) return false;
      return true;
    });
    
    const studentStats = classStudents.map(student => {
      let sakit = 0, izin = 0, alpa = 0;
      classSessions.forEach(session => {
        const stat = session.records[student.id];
        if (stat === 'Sakit') sakit++;
        if (stat === 'Izin') izin++;
        if (stat === 'Alpa') alpa++;
      });
      return {
        student,
        total: sakit + izin + alpa,
        details: `(Alpa: ${alpa}, Sakit: ${sakit}, Izin: ${izin})`
      };
    });

    return studentStats.sort((a, b) => {
      // Prioritaskan siswa dengan ketidakhadiran tertinggi, lalu alfabetis untuk seluruh siswa kelas
      if (b.total !== a.total) {
        return b.total - a.total;
      }
      return a.student.name.localeCompare(b.student.name, 'id');
    });
  }, [selectedClass, students, attendanceSessions, startDate, endDate]);

  const [savedReports, setSavedReports] = useState<(FollowupReport & { studentName?: string })[]>([]);
  const [isLoadingReports, setIsLoadingReports] = useState(false);
  const [printSelection, setPrintSelection] = useState<string>('all');
  const [deleteConfirmId, setDeleteConfirmId] = useState<string | null>(null);

  // --- REKAP TINDAK LANJUT BULAN INI STATE & CRUD ---
  const [monthlyRecaps, setMonthlyRecaps] = useState<HomeroomMonthlyRecap[]>([]);
  const [isLoadingRecaps, setIsLoadingRecaps] = useState(false);
  const [isSavingRecap, setIsSavingRecap] = useState(false);
  const [recapEditId, setRecapEditId] = useState<string | null>(null);
  const [deleteRecapConfirmId, setDeleteRecapConfirmId] = useState<string | null>(null);

  const [recapForm, setRecapForm] = useState<{
    pembinaanMurid: string[];
    kontakOrangTua: string[];
    koordinasiBK: string[];
    koordinasiGuruWali: string[];
    koordinasiGuruMapel: string[];
    kasusBelumSelesai: string[];
    keterangan: string;
  }>({
    pembinaanMurid: [],
    kontakOrangTua: [],
    koordinasiBK: [],
    koordinasiGuruWali: [],
    koordinasiGuruMapel: [],
    kasusBelumSelesai: [],
    keterangan: ''
  });

  const loadSavedReports = useCallback(async () => {
    if (!selectedClass) {
      setSavedReports([]);
      return;
    }
    
    setIsLoadingReports(true);
    try {
      // 1. Try local cache first
      const localKey = `kaguci_student_followups_${selectedClass}_${currentMonth}`;
      const cached = localStorage.getItem(localKey);
      if (cached) {
        try {
          const parsed = JSON.parse(cached);
          if (Array.isArray(parsed)) {
            setSavedReports(parsed);
          }
        } catch {
          // ignore parse error
        }
      }

      // 2. Fetch from Firestore if user is authenticated
      if (activeAuth?.currentUser && activeDb) {
        const q = query(
          collection(activeDb, 'student_followups'),
          where('className', '==', selectedClass),
          where('month', '==', currentMonth),
          where('userId', '==', activeAuth.currentUser.uid)
        );
        trackOp?.('read', 1);
        const querySnapshot = await getDocs(q);
        
        const reports: (FollowupReport & { studentName?: string })[] = [];
        querySnapshot.forEach((doc) => {
          const data = doc.data() as FollowupReport;
          const student = students.find(s => s.id === data.studentId);
          reports.push({
            ...data,
            id: doc.id,
            studentName: student?.name || 'Siswa Tidak Ditemukan'
          });
        });
        
        reports.sort((a, b) => (a.studentName || '').localeCompare(b.studentName || ''));
        setSavedReports(reports);
        localStorage.setItem(localKey, JSON.stringify(reports));
      }
    } catch (error) {
      console.error("Error loading saved reports:", error);
      handleFirestoreError(error, OperationType.GET, 'student_followups');
    } finally {
      setIsLoadingReports(false);
    }
  }, [selectedClass, currentMonth, activeAuth?.currentUser, activeDb, trackOp, students]);

  // Load Monthly Recaps from Firestore & LocalStorage
  const loadMonthlyRecaps = useCallback(async () => {
    if (!selectedClass) {
      setMonthlyRecaps([]);
      return;
    }

    setIsLoadingRecaps(true);
    try {
      // 1. Try local storage first for instant load
      const localKey = `kaguci_homeroom_recaps_${selectedClass}_${currentMonth}`;
      const cached = localStorage.getItem(localKey);
      if (cached) {
        try {
          const parsed = JSON.parse(cached);
          if (Array.isArray(parsed)) {
            setMonthlyRecaps(parsed);
          }
        } catch {
          // ignore parse error
        }
      }

      // 2. Fetch from Firestore if available
      if (activeAuth?.currentUser && activeDb) {
        const q = query(
          collection(activeDb, 'homeroom_monthly_recaps'),
          where('className', '==', selectedClass),
          where('month', '==', currentMonth),
          where('userId', '==', activeAuth.currentUser.uid)
        );
        trackOp?.('read', 1);
        const snapshot = await getDocs(q);
        const list: HomeroomMonthlyRecap[] = [];
        snapshot.forEach(docSnap => {
          list.push({
            id: docSnap.id,
            ...(docSnap.data() as Omit<HomeroomMonthlyRecap, 'id'>)
          });
        });

        if (list.length > 0) {
          setMonthlyRecaps(list);
          localStorage.setItem(localKey, JSON.stringify(list));
        }
      }
    } catch (err) {
      console.error("Error loading monthly recaps:", err);
      handleFirestoreError(err, OperationType.GET, 'homeroom_monthly_recaps');
    } finally {
      setIsLoadingRecaps(false);
    }
  }, [selectedClass, currentMonth, activeAuth?.currentUser, activeDb, trackOp]);

  useEffect(() => {
    loadSavedReports();
    loadMonthlyRecaps();
  }, [loadSavedReports, loadMonthlyRecaps]);

  // Handle class change
  useEffect(() => {
    setSelectedStudentId('');
    resetForm();
    resetRecapForm();
  }, [selectedClass]);

  // Load existing individual report data when student is selected
  useEffect(() => {
    async function loadStudentReport() {
      if (!selectedStudentId || !activeAuth?.currentUser || !activeDb) {
        resetForm();
        return;
      }

      try {
        const q = query(
          collection(activeDb, 'student_followups'),
          where('studentId', '==', selectedStudentId),
          where('month', '==', currentMonth),
          where('userId', '==', activeAuth.currentUser.uid)
        );
        trackOp?.('read', 1);
        const querySnapshot = await getDocs(q);

        if (!querySnapshot.empty) {
          const docs = querySnapshot.docs.map(d => ({ id: d.id, ...d.data() } as FollowupReport));
          const match = docs.find(d => d.date === selectedDate) || docs[0];
          const docData = match;
          
          setExistingDocId(match.id);
          setTerlambat(docData.terlambat || 'Tidak Pernah');
          setPermasalahan(docData.permasalahan || '');
          setTindakLanjut(docData.tindakLanjut || '');
          setStatus(docData.status || 'Normal');
          setIsEditing(false);
        } else {
          resetForm();
        }
      } catch (error) {
        console.error("Error loading report:", error);
      }
    }
    loadStudentReport();
  }, [selectedStudentId, selectedDate, activeAuth?.currentUser, currentMonth, activeDb, trackOp]);

  const resetForm = () => {
    setExistingDocId('');
    setTerlambat('Tidak Pernah');
    setPermasalahan('');
    setTindakLanjut('');
    setStatus('Normal');
    setIsEditing(true);
  };

  const resetRecapForm = () => {
    setRecapEditId(null);
    setRecapForm({
      pembinaanMurid: [],
      kontakOrangTua: [],
      koordinasiBK: [],
      koordinasiGuruWali: [],
      koordinasiGuruMapel: [],
      kasusBelumSelesai: [],
      keterangan: ''
    });
  };

  const handleSave = async () => {
    if (!selectedClass || !selectedStudentId) {
      showToast?.('Pilih kelas dan siswa terlebih dahulu', 'error');
      return;
    }
    if (!permasalahan.trim() || !tindakLanjut.trim()) {
      showToast?.('Isi permasalahan utama dan tindak lanjut', 'error');
      return;
    }
    if (!activeAuth?.currentUser || !activeDb) return;

    setIsSaving(true);
    try {
      const docId = existingDocId || `${activeAuth.currentUser.uid}_${selectedStudentId}_${selectedDate}`;
      const docRef = doc(activeDb, 'student_followups', docId);
      
      const payload = {
        studentId: selectedStudentId,
        className: selectedClass,
        terlambat,
        permasalahan,
        tindakLanjut,
        status,
        month: currentMonth,
        date: selectedDate,
        userId: activeAuth.currentUser.uid,
        updatedAt: serverTimestamp()
      };

      await setDoc(docRef, payload, { merge: true });
      trackOp?.('write', 1);
      
      setExistingDocId(docId);
      setIsEditing(false);
      showToast?.('Laporan tindak lanjut berhasil disimpan!', 'success');
      loadSavedReports();
    } catch (error) {
      console.error("Error saving report:", error);
      showToast?.('Gagal menyimpan laporan', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleDelete = async (reportId: string) => {
    if (!activeDb) return;
    try {
      await deleteDoc(doc(activeDb, 'student_followups', reportId));
      trackOp?.('write', 1);
      showToast?.('Laporan berhasil dihapus', 'success');
      setDeleteConfirmId(null);
      loadSavedReports();
      
      if (existingDocId === reportId) {
        setSelectedStudentId('');
        resetForm();
      }
    } catch (error) {
      console.error("Error deleting report:", error);
      showToast?.('Gagal menghapus laporan', 'error');
    }
  };

  // --- SAVE REKAP TINDAK LANJUT BULAN INI ---
  const handleSaveRecap = async () => {
    if (!selectedClass) {
      showToast?.('Pilih kelas terlebih dahulu', 'error');
      return;
    }

    const classStudents = students.filter(s => s.class === selectedClass);
    const getNames = (ids: string[]) => ids.map(id => classStudents.find(s => s.id === id)?.name || id);

    setIsSavingRecap(true);
    try {
      const id = recapEditId || `recap_${selectedClass}_${currentMonth}_${Date.now()}`;
      const recapItem: HomeroomMonthlyRecap = {
        id,
        className: selectedClass,
        month: currentMonth,
        startDate,
        endDate,
        pembinaanMurid: recapForm.pembinaanMurid,
        kontakOrangTua: recapForm.kontakOrangTua,
        koordinasiBK: recapForm.koordinasiBK,
        koordinasiGuruWali: recapForm.koordinasiGuruWali,
        koordinasiGuruMapel: recapForm.koordinasiGuruMapel,
        kasusBelumSelesai: recapForm.kasusBelumSelesai,
        pembinaanMuridNames: getNames(recapForm.pembinaanMurid),
        kontakOrangTuaNames: getNames(recapForm.kontakOrangTua),
        koordinasiBKNames: getNames(recapForm.koordinasiBK),
        koordinasiGuruWaliNames: getNames(recapForm.koordinasiGuruWali),
        koordinasiGuruMapelNames: getNames(recapForm.koordinasiGuruMapel),
        kasusBelumSelesaiNames: getNames(recapForm.kasusBelumSelesai),
        keterangan: recapForm.keterangan || '',
        userId: activeAuth?.currentUser?.uid || 'anonymous',
        updatedAt: new Date().toISOString()
      };

      // Save to state & localStorage
      setMonthlyRecaps(prev => {
        const filtered = prev.filter(r => r.id !== id);
        const updated = [recapItem, ...filtered];
        localStorage.setItem(`kaguci_homeroom_recaps_${selectedClass}_${currentMonth}`, JSON.stringify(updated));
        return updated;
      });

      // Save to Firestore if available
      if (activeAuth?.currentUser && activeDb) {
        await setDoc(doc(activeDb, 'homeroom_monthly_recaps', id), {
          ...recapItem,
          serverUpdatedAt: serverTimestamp()
        }, { merge: true });
        trackOp?.('write', 1);
      }

      showToast?.(recapEditId ? 'Rekap tindak lanjut berhasil diperbarui!' : 'Rekap tindak lanjut berhasil disimpan!', 'success');
      resetRecapForm();
    } catch (err) {
      console.error("Error saving recap:", err);
      showToast?.('Gagal menyimpan rekap tindak lanjut', 'error');
    } finally {
      setIsSavingRecap(false);
    }
  };

  const handleEditRecap = (recap: HomeroomMonthlyRecap) => {
    setRecapEditId(recap.id);
    setRecapForm({
      pembinaanMurid: parseStudentIds(recap.pembinaanMurid),
      kontakOrangTua: parseStudentIds(recap.kontakOrangTua),
      koordinasiBK: parseStudentIds(recap.koordinasiBK),
      koordinasiGuruWali: parseStudentIds(recap.koordinasiGuruWali),
      koordinasiGuruMapel: parseStudentIds(recap.koordinasiGuruMapel),
      kasusBelumSelesai: parseStudentIds(recap.kasusBelumSelesai),
      keterangan: recap.keterangan || ''
    });
    // Smooth scroll to form
    const elem = document.getElementById('rekap-form-anchor');
    if (elem) {
      elem.scrollIntoView({ behavior: 'smooth' });
    }
  };

  const handleDeleteRecap = async (recapId: string) => {
    try {
      setMonthlyRecaps(prev => {
        const updated = prev.filter(r => r.id !== recapId);
        localStorage.setItem(`kaguci_homeroom_recaps_${selectedClass}_${currentMonth}`, JSON.stringify(updated));
        return updated;
      });

      if (activeAuth?.currentUser && activeDb) {
        await deleteDoc(doc(activeDb, 'homeroom_monthly_recaps', recapId));
        trackOp?.('write', 1);
      }

      showToast?.('Rekap tindak lanjut berhasil dihapus', 'success');
      setDeleteRecapConfirmId(null);
      if (recapEditId === recapId) {
        resetRecapForm();
      }
    } catch (err) {
      console.error("Error deleting recap:", err);
      showToast?.('Gagal menghapus rekap', 'error');
    }
  };

  const [isExporting, setIsExporting] = useState(false);

  // Format date to Indonesian string helper
  const formatIndoDate = (dateStr: string) => {
    if (!dateStr) return '-';
    try {
      const d = new Date(dateStr);
      const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
      return `${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;
    } catch {
      return dateStr;
    }
  };

  const generateHomeroomPDFDocument = (): { doc: jsPDF; fileName: string } => {
    const doc = new jsPDF({ orientation: 'landscape', format: 'legal' });
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    
    // Header - Kop Surat Standar (Presisi, Simetris & Terkunci Ratio 1450:341 agar tidak meleyot)
    const imgHeight = 32; // Standar tinggi kop surat dinas/sekolah (3.2 cm)
    const imgRatio = 1450 / 341;
    const imgWidth = Math.round((imgHeight * imgRatio) * 100) / 100; // 136.07 mm
    const imgX = (pageWidth - imgWidth) / 2; // Presisi & simetris di tengah
    const imgY = 8;
    
    try {
      doc.addImage(kopSuratBase64, 'PNG', imgX, imgY, imgWidth, imgHeight);
    } catch (e) {
      console.error("Failed to add custom header image", e);
    }
    
    const startY = imgY + imgHeight + 6;
    doc.setFontSize(13);
    doc.setFont("helvetica", "bold");
    doc.text('LAPORAN TINDAK LANJUT WALI KELAS', pageWidth / 2, startY, { align: 'center' });
    
    doc.setFontSize(9.5);
    doc.setFont("helvetica", "normal");
    
    let displayPeriod = '';
    if (startDate && endDate) {
      displayPeriod = `${formatIndoDate(startDate)} s.d ${formatIndoDate(endDate)}`;
    } else if (printScope === 'all_months') {
      displayPeriod = 'Seluruh Periode';
    } else {
      const [yearStr, monthStr] = exportMonth.split('-');
      const y = parseInt(yearStr, 10);
      const m = parseInt(monthStr, 10);
      const lastDay = new Date(y, m, 0).getDate();
      const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
      const monthName = months[m - 1];
      displayPeriod = `1 ${monthName} ${y} s.d ${lastDay} ${monthName} ${y}`;
    }

    const leftLabelX = 14;
    const leftColonX = 52;
    const leftValX = 55;

    // Metadata Left Column
    doc.text(`Kelas`, leftLabelX, startY + 7);
    doc.text(`:`, leftColonX, startY + 7);
    doc.text(`${selectedClass || profileData?.waliKelasClass || '-'}`, leftValX, startY + 7);

    doc.text(`Nama Wali Kelas`, leftLabelX, startY + 12);
    doc.text(`:`, leftColonX, startY + 12);
    doc.text(`${profileData?.namaGuruMapel || '-'}`, leftValX, startY + 12);

    doc.text(`Periode Pelaporan`, leftLabelX, startY + 17);
    doc.text(`:`, leftColonX, startY + 17);
    doc.text(`${displayPeriod}`, leftValX, startY + 17);
    
    // Metadata Right Column (Jumlah Siswa L & P DI ATAS Semester)
    const rightLabelX = pageWidth - 90;
    const rightColonX = pageWidth - 55;
    const rightValX = pageWidth - 52;

    const jmlL = profileData?.jumlahSiswaLakiLaki || '0';
    const jmlP = profileData?.jumlahSiswaPerempuan || '0';

    doc.text(`Laki-laki`, rightLabelX, startY + 7);
    doc.text(`:`, rightColonX, startY + 7);
    doc.text(`${jmlL} Siswa`, rightValX, startY + 7);

    doc.text(`Perempuan`, rightLabelX, startY + 12);
    doc.text(`:`, rightColonX, startY + 12);
    doc.text(`${jmlP} Siswa`, rightValX, startY + 12);

    doc.text(`Semester`, rightLabelX, startY + 17);
    doc.text(`:`, rightColonX, startY + 17);
    doc.text(`${profileData?.semester || '-'}`, rightValX, startY + 17);

    doc.text(`Tahun Pelajaran`, rightLabelX, startY + 22);
    doc.text(`:`, rightColonX, startY + 22);
    doc.text(`${profileData?.tahunPelajaran || '-'}`, rightValX, startY + 22);

    // --- TABLE 1: STATISTIK AKUMULASI KELAS ---
    doc.setFontSize(10.5);
    doc.setFont("helvetica", "bold");
    doc.text(`I. STATISTIK AKUMULASI KEHADIRAN KELAS ${selectedClass || ''}`.trim(), 14, startY + 29);

    const statsHeaders = ['Total Siswa', 'Total Pertemuan', 'Hadir', 'Sakit', 'Izin', 'Alpa', 'Dispen', 'Persentase Kehadiran'];
    const statsBody = [[
      classAccumulationStats.totalSiswa.toString(),
      classAccumulationStats.totalSesi.toString(),
      `${classAccumulationStats.hadir} (${classAccumulationStats.hadirPct.toFixed(1)}%)`,
      `${classAccumulationStats.sakit} (${classAccumulationStats.sakitPct.toFixed(1)}%)`,
      `${classAccumulationStats.izin} (${classAccumulationStats.izinPct.toFixed(1)}%)`,
      `${classAccumulationStats.alpa} (${classAccumulationStats.alpaPct.toFixed(1)}%)`,
      `${classAccumulationStats.dispen} (${classAccumulationStats.dispenPct.toFixed(1)}%)`,
      `${classAccumulationStats.persentase.toFixed(1)}%`
    ]];

    const statsTableOptions = {
      startY: startY + 33,
      head: [statsHeaders],
      body: statsBody,
      theme: 'grid' as const,
      headStyles: { fillColor: [235, 235, 235], textColor: 0, fontStyle: 'bold', halign: 'center' as const, lineWidth: 0.2, lineColor: [0, 0, 0] },
      styles: { fontSize: 8, cellPadding: 2.5, lineWidth: 0.2, lineColor: [0, 0, 0], textColor: 0, halign: 'center' as const }
    };

    if (typeof autoTable === 'function') {
      autoTable(doc, statsTableOptions as unknown as Parameters<typeof autoTable>[1]);
    } else if (typeof (doc as unknown as { autoTable: (opt: unknown) => void }).autoTable === 'function') {
      (doc as unknown as { autoTable: (opt: unknown) => void }).autoTable(statsTableOptions);
    }

    const docWithTable = doc as unknown as { lastAutoTable?: { finalY: number } };
    const statsFinalY = docWithTable.lastAutoTable?.finalY || (startY + 48);

    // --- TABLE 2: MURID PRIORITAS TINDAK LANJUT ---
    doc.setFontSize(10.5);
    doc.setFont("helvetica", "bold");
    doc.text('II. MURID PRIORITAS TINDAK LANJUT', 14, statsFinalY + 8);
    
    const reportsToPrint = printSelection === 'all' 
      ? savedReports 
      : savedReports.filter(r => r.id === printSelection);
      
    const headers = ['Tanggal', 'Nama Siswa', 'Terlambat', 'Ketidakhadiran (A/S/I)', 'Permasalahan', 'Tindak Lanjut', 'Status'];
    const body = (reportsToPrint.length > 0 ? reportsToPrint : []).map(r => {
      let dateStr = '-';
      if (r.date) {
         const d = new Date(r.date);
         const months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Ags', 'Sep', 'Okt', 'Nov', 'Des'];
         dateStr = `${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;
      }
      
      let absenInfo = '-';
      const studentStat = priorityStudents.find(s => s.student.id === r.studentId);
      if (studentStat) {
        absenInfo = studentStat.details.replace(/[()]/g, '');
      }

      return [
        dateStr,
        r.studentName || '-',
        r.terlambat || '-',
        absenInfo,
        r.permasalahan || '-',
        r.tindakLanjut || '-',
        r.status || '-'
      ];
    });

    // Fallback row if empty
    const finalTable2Body = body.length > 0 ? body : [
      ['-', 'Tidak ada catatan siswa khusus pada periode ini', '-', '-', '-', '-', 'Normal']
    ];
    
    const autoTableOptions = {
      startY: statsFinalY + 12,
      head: [headers],
      body: finalTable2Body,
      theme: 'grid' as const,
      headStyles: { fillColor: [235, 235, 235], textColor: 0, fontStyle: 'bold', halign: 'center' as const, lineWidth: 0.2, lineColor: [0, 0, 0] },
      styles: { fontSize: 7.5, cellPadding: 2, lineWidth: 0.2, lineColor: [0, 0, 0], textColor: 0 },
      columnStyles: {
        0: { halign: 'center' as const, cellWidth: 22 },
        1: { halign: 'left' as const, cellWidth: 42 },
        2: { halign: 'center' as const, cellWidth: 22 },
        3: { halign: 'center' as const, cellWidth: 35 },
        4: { halign: 'left' as const },
        5: { halign: 'left' as const },
        6: { halign: 'center' as const, cellWidth: 25 }
      }
    };
    
    if (typeof autoTable === 'function') {
      autoTable(doc, autoTableOptions as unknown as Parameters<typeof autoTable>[1]);
    } else if (typeof (doc as unknown as { autoTable: (opt: unknown) => void }).autoTable === 'function') {
      (doc as unknown as { autoTable: (opt: unknown) => void }).autoTable(autoTableOptions);
    }
    
    let finalY2 = docWithTable.lastAutoTable?.finalY || (statsFinalY + 35);

    // --- TABLE 3: REKAP TINDAK LANJUT BULAN INI ---
    if (finalY2 + 55 > pageHeight - 35) {
      doc.addPage();
      finalY2 = 15;
    }

    doc.setFontSize(10.5);
    doc.setFont("helvetica", "bold");
    doc.text('III. REKAPITULASI TINDAK LANJUT WALI KELAS BULAN INI', 14, finalY2 + 8);

    const activeRecap = monthlyRecaps[0] || {
      pembinaanMurid: [],
      kontakOrangTua: [],
      koordinasiBK: [],
      koordinasiGuruWali: [],
      koordinasiGuruMapel: [],
      kasusBelumSelesai: [],
      keterangan: '-'
    };

    const classStudents = students.filter(s => s.class === selectedClass);
    const getNamesList = (idsOrNum: string[] | number | undefined, savedNames?: string[]) => {
      if (Array.isArray(idsOrNum)) {
        return idsOrNum.map((id, idx) => {
          const s = classStudents.find(st => st.id === id || st.name === id);
          if (s) return s.name;
          if (savedNames && savedNames[idx]) return savedNames[idx];
          return id;
        });
      }
      if (savedNames && savedNames.length > 0) return savedNames;
      return [];
    };

    const pembinaanNames = getNamesList(activeRecap.pembinaanMurid, activeRecap.pembinaanMuridNames);
    const kontakOrtuNames = getNamesList(activeRecap.kontakOrangTua, activeRecap.kontakOrangTuaNames);
    const bkNames = getNamesList(activeRecap.koordinasiBK, activeRecap.koordinasiBKNames);
    const guruWaliNames = getNamesList(activeRecap.koordinasiGuruWali, activeRecap.koordinasiGuruWaliNames);
    const guruMapelNames = getNamesList(activeRecap.koordinasiGuruMapel, activeRecap.koordinasiGuruMapelNames);
    const kasusBelumNames = getNamesList(activeRecap.kasusBelumSelesai, activeRecap.kasusBelumSelesaiNames);

    const formatDescWithNames = (names: string[], defaultDesc: string) => {
      if (names.length > 0) {
        return `Siswa: ${names.join(', ')} (${defaultDesc})`;
      }
      return defaultDesc;
    };

    const allSelectedNames = Array.from(new Set([
      ...pembinaanNames,
      ...kontakOrtuNames,
      ...bkNames,
      ...guruWaliNames,
      ...guruMapelNames,
      ...kasusBelumNames
    ]));

    const countPembinaan = getCategoryCount(activeRecap.pembinaanMurid);
    const countKontakOrtu = getCategoryCount(activeRecap.kontakOrangTua);
    const countBK = getCategoryCount(activeRecap.koordinasiBK);
    const countGuruWali = getCategoryCount(activeRecap.koordinasiGuruWali);
    const countGuruMapel = getCategoryCount(activeRecap.koordinasiGuruMapel);
    const countKasusBelum = getCategoryCount(activeRecap.kasusBelumSelesai);

    const uniqueMuridCount = allSelectedNames.length > 0 
      ? allSelectedNames.length 
      : (countPembinaan + countKontakOrtu + countBK + countGuruWali + countGuruMapel + countKasusBelum);

    const recapHeaders = ['No', 'Bentuk / Kategori Tindak Lanjut', 'Jumlah Murid', 'Deskripsi / Siswa Terpilih'];
    const recapBody = [
      ['1', 'Pembinaan Murid (Langsung oleh Wali Kelas)', `${countPembinaan} Murid`, formatDescWithNames(pembinaanNames, 'Bimbingan dan konseling personal terhadap siswa di kelas')],
      ['2', 'Kontak Orang Tua / Wali Murid', `${countKontakOrtu} Murid`, formatDescWithNames(kontakOrtuNames, 'Komunikasi telepon/pesan, surat pemanggilan, atau kunjungan rumah')],
      ['3', 'Koordinasi Bimbingan Konseling (BK)', `${countBK} Murid`, formatDescWithNames(bkNames, 'Rujukan penanganan dan kolaborasi bersama guru BK sekolah')],
      ['4', 'Koordinasi Guru Wali', `${countGuruWali} Murid`, formatDescWithNames(guruWaliNames, 'Koordinasi antarsejawat dan pembagian pola penanganan siswa')],
      ['5', 'Koordinasi Guru Mapel', `${countGuruMapel} Murid`, formatDescWithNames(guruMapelNames, 'Pembahasan progres kehadiran, tugas, dan capaian akademik mata pelajaran')],
      ['6', 'Kasus Belum Selesai (Dalam Pemantauan)', `${countKasusBelum} Murid`, formatDescWithNames(kasusBelumNames, 'Siswa yang masih memerlukan pemantauan dan evaluasi tindak lanjut lanjutan')],
      ['', 'TOTAL AKUMULASI TINDAK LANJUT', `${uniqueMuridCount} Murid`, allSelectedNames.length > 0 ? `Daftar Siswa: ${allSelectedNames.join(', ')} | ${activeRecap.keterangan || 'Akumulasi rekapitulasi pembinaan wali kelas'}` : (activeRecap.keterangan || 'Akumulasi rekapitulasi pembinaan wali kelas')]
    ];

    const recapTableOptions = {
      startY: finalY2 + 12,
      head: [recapHeaders],
      body: recapBody,
      theme: 'grid' as const,
      headStyles: { fillColor: [235, 235, 235], textColor: 0, fontStyle: 'bold', halign: 'center' as const, lineWidth: 0.2, lineColor: [0, 0, 0] },
      styles: { fontSize: 7.5, cellPadding: 2, lineWidth: 0.2, lineColor: [0, 0, 0], textColor: 0 },
      columnStyles: {
        0: { halign: 'center' as const, cellWidth: 14 },
        1: { halign: 'left' as const, cellWidth: 80 },
        2: { halign: 'center' as const, cellWidth: 45 },
        3: { halign: 'left' as const }
      }
    };

    if (typeof autoTable === 'function') {
      autoTable(doc, recapTableOptions as unknown as Parameters<typeof autoTable>[1]);
    } else if (typeof (doc as unknown as { autoTable: (opt: unknown) => void }).autoTable === 'function') {
      (doc as unknown as { autoTable: (opt: unknown) => void }).autoTable(recapTableOptions);
    }

    const finalY3 = docWithTable.lastAutoTable?.finalY || (finalY2 + 45);

    // --- SIGNATURE SECTION ---
    let sigY = finalY3 + 12;
    if (sigY + 38 > pageHeight) {
       doc.addPage();
       sigY = 20;
    }
    
    const d = new Date();
    const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    const dateStrNow = `Cililin, ${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;
    
    doc.setFontSize(9.5);
    doc.setFont("helvetica", "normal");
    
    const leftSig = getLeftSignerInfo();
    const midSig = getMidSignerInfo();
    const is3Signers = leftSig.enabled && midSig.enabled;
    
    if (is3Signers) {
      // Format 3 Tanda Tangan: Kiri (Mengetahui 1), Tengah (Mengetahui 2), Kanan (Wali Kelas)
      const leftX = 14;
      const midX = (pageWidth / 2) - 25;
      const rightX = pageWidth - 65;
      
      // Kiri
      doc.text('Mengetahui,', leftX, sigY);
      doc.text(leftSig.title, leftX, sigY + 5);
      doc.setFont("helvetica", "bold");
      doc.text(leftSig.name, leftX, sigY + 23);
      doc.setFont("helvetica", "normal");
      if (leftSig.nip) doc.text(leftSig.nip, leftX, sigY + 28);

      // Tengah
      doc.text('Mengetahui,', midX, sigY);
      doc.text(midSig.title, midX, sigY + 5);
      doc.setFont("helvetica", "bold");
      doc.text(midSig.name, midX, sigY + 23);
      doc.setFont("helvetica", "normal");
      if (midSig.nip) doc.text(midSig.nip, midX, sigY + 28);

      // Kanan (Wali Kelas)
      doc.text(dateStrNow, rightX, sigY);
      doc.text(`Wali Kelas ${selectedClass || profileData?.waliKelasClass || ''}`, rightX, sigY + 5);
      doc.setFont("helvetica", "bold");
      doc.text(profileData?.namaGuruMapel || '(________________________)', rightX, sigY + 23);
      doc.setFont("helvetica", "normal");
      if (profileData?.nipGuruMapel) doc.text(`NIP. ${profileData.nipGuruMapel}`, rightX, sigY + 28);
    } else {
      // Format 1 atau 2 Tanda Tangan
      const activeLeft = leftSig.enabled ? leftSig : (midSig.enabled ? midSig : null);
      const rightSigX = pageWidth - 80;

      if (activeLeft) {
        doc.text('Mengetahui,', 20, sigY);
        doc.text(activeLeft.title, 20, sigY + 5);
        doc.setFont("helvetica", "bold");
        doc.text(activeLeft.name, 20, sigY + 23);
        doc.setFont("helvetica", "normal");
        if (activeLeft.nip) {
          doc.text(activeLeft.nip, 20, sigY + 28);
        }
      }
      
      doc.text(dateStrNow, rightSigX, sigY);
      doc.text(`Wali Kelas ${selectedClass || profileData?.waliKelasClass || ''}`, rightSigX, sigY + 5);
      doc.setFont("helvetica", "bold");
      doc.text(profileData?.namaGuruMapel || '(________________________)', rightSigX, sigY + 23);
      doc.setFont("helvetica", "normal");
      if (profileData?.nipGuruMapel) {
        doc.text(`NIP. ${profileData.nipGuruMapel}`, rightSigX, sigY + 28);
      }
    }
    
    return {
      doc,
      fileName: `Laporan_Wali_Kelas_${selectedClass || 'Semua'}_${currentMonth}.pdf`
    };
  };

  const handleExportPDF = async () => {
    setIsExporting(true);
    
    setTimeout(() => {
      try {
        const { doc, fileName } = generateHomeroomPDFDocument();
        doc.save(fileName);
        showToast?.('Laporan Wali Kelas berhasil diunduh.', 'success');
      } catch (err) {
        console.error("Gagal mengekspor PDF laporan:", err);
        showToast?.('Gagal mengekspor PDF laporan.', 'error');
      } finally {
        setIsExporting(false);
      }
    }, 200);
  };

  const handlePreviewPDF = () => {
    setIsPreviewLoading(true);
    setTimeout(() => {
      try {
        const { doc, fileName } = generateHomeroomPDFDocument();
        const blob = doc.output('blob');
        const url = URL.createObjectURL(blob);
        setPreviewPdfUrl(url);
        setPreviewFileName(fileName);
        setShowPreviewModal(true);
      } catch (err) {
        console.error("Gagal membuat preview PDF:", err);
        showToast?.('Gagal memuat pratinjau dokumen PDF.', 'error');
      } finally {
        setIsPreviewLoading(false);
      }
    }, 200);
  };

  const handleClosePreview = () => {
    if (previewPdfUrl) {
      URL.revokeObjectURL(previewPdfUrl);
    }
    setPreviewPdfUrl(null);
    setShowPreviewModal(false);
  };

  return (
    <div className="max-w-5xl mx-auto space-y-6 pb-12">
      {/* CARD 1: FORM PENCATATAN TINDAK LANJUT SISWA */}
      <div className="bg-white p-6 sm:p-8 rounded-[2rem] shadow-sm border border-slate-100 relative overflow-hidden print:hidden">
        <div className="absolute top-0 right-0 w-64 h-64 bg-amber-50 rounded-full blur-3xl opacity-50 -mr-32 -mt-32"></div>
        
        <div className="relative z-10 flex flex-col sm:flex-row sm:items-center justify-between gap-6 mb-6">
          <div>
            <h2 className="text-2xl sm:text-3xl font-black text-slate-800 tracking-tight flex items-center gap-3">
              <div className="w-12 h-12 bg-amber-100 text-amber-600 rounded-2xl flex items-center justify-center rotate-3">
                <FileText className="w-6 h-6" />
              </div>
              Laporan Tindak Lanjut Wali Kelas
            </h2>
            <p className="text-slate-500 mt-1.5 font-medium text-sm">
              Kelola pencatatan kasus siswa dan rekapitulasi penanganan berkala wali kelas.
            </p>
          </div>

          {/* Profile gender summary tag */}
          {profileData?.role === 'Wali Kelas' && (
            <div className="flex flex-wrap items-center gap-2 bg-slate-50 border border-slate-200 px-4 py-2 rounded-2xl text-xs font-bold text-slate-700">
              <Users className="w-4 h-4 text-emerald-600" />
              <span>Siswa: <strong className="text-emerald-700">{profileData?.jumlahSiswaLakiLaki || 0} L</strong> / <strong className="text-emerald-700">{profileData?.jumlahSiswaPerempuan || 0} P</strong></span>
              <span className="text-slate-300">|</span>
              <span>Total: <strong className="text-slate-900">{Number(profileData?.jumlahSiswaLakiLaki || 0) + Number(profileData?.jumlahSiswaPerempuan || 0)}</strong></span>
            </div>
          )}
        </div>

        {/* CONTROLS GRID: Periode, Tanggal Manual, Kelas, Siswa */}
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 relative z-10 mb-6">
          {/* Pilih Bulan & Tahun */}
          <div className="space-y-1.5">
            <label className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
              <Calendar className="w-3.5 h-3.5 text-amber-500" />
              Bulan Acuan
            </label>
            <div className="grid grid-cols-2 gap-1.5">
              <select
                value={selectedMonth.split('-')[1] || '01'}
                onChange={(e) => {
                  const yr = selectedMonth.split('-')[0] || String(new Date().getFullYear());
                  const val = `${yr}-${e.target.value}`;
                  setSelectedMonth(val);
                  if (!selectedDate.startsWith(val)) {
                    setSelectedDate(`${val}-01`);
                  }
                }}
                className="w-full px-2.5 py-3 bg-slate-50 border-2 border-slate-200 rounded-xl text-slate-700 font-bold focus:ring-2 focus:ring-amber-500/20 focus:border-amber-500 transition-all text-xs"
              >
                {MONTH_OPTIONS.map(m => (
                  <option key={m.value} value={m.value}>{m.label}</option>
                ))}
              </select>
              <select
                value={selectedMonth.split('-')[0] || String(new Date().getFullYear())}
                onChange={(e) => {
                  const mo = selectedMonth.split('-')[1] || String(new Date().getMonth() + 1).padStart(2, '0');
                  const val = `${e.target.value}-${mo}`;
                  setSelectedMonth(val);
                  if (!selectedDate.startsWith(val)) {
                    setSelectedDate(`${val}-01`);
                  }
                }}
                className="w-full px-2.5 py-3 bg-slate-50 border-2 border-slate-200 rounded-xl text-slate-700 font-bold focus:ring-2 focus:ring-amber-500/20 focus:border-amber-500 transition-all text-xs"
              >
                {availableYears.map(yr => (
                  <option key={yr} value={yr}>{yr}</option>
                ))}
              </select>
            </div>
          </div>

          {/* Manual Date Range: Tanggal Awal & Tanggal Akhir (REQUESTED ITEM 1) */}
          <div className="space-y-1.5">
            <label className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
              <Calendar className="w-3.5 h-3.5 text-amber-500" />
              Periode Pelaporan (Awal - Akhir)
            </label>
            <div className="grid grid-cols-2 gap-1.5">
              <input 
                type="date"
                value={startDate}
                onChange={(e) => setStartDate(e.target.value)}
                title="Tanggal Awal Pelaporan"
                className="w-full px-2 py-2.5 bg-slate-50 border-2 border-slate-200 rounded-xl text-slate-700 font-bold focus:ring-2 focus:ring-amber-500/20 focus:border-amber-500 transition-all text-xs"
              />
              <input 
                type="date"
                value={endDate}
                onChange={(e) => setEndDate(e.target.value)}
                title="Tanggal Akhir Pelaporan"
                className="w-full px-2 py-2.5 bg-slate-50 border-2 border-slate-200 rounded-xl text-slate-700 font-bold focus:ring-2 focus:ring-amber-500/20 focus:border-amber-500 transition-all text-xs"
              />
            </div>
          </div>

          {/* Pilih Kelas */}
          <div className="space-y-1.5">
            <label className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
              <Building2 className="w-3.5 h-3.5 text-amber-500" />
              Pilih Kelas
            </label>
            <select 
              value={selectedClass} 
              onChange={(e) => setSelectedClass(e.target.value)}
              className="w-full px-3 py-2.5 bg-slate-50 border-2 border-slate-200 rounded-xl text-slate-700 font-bold focus:ring-2 focus:ring-amber-500/20 focus:border-amber-500 transition-all text-xs"
            >
              <option value="" disabled>-- Pilih Kelas --</option>
              {classList.map(cls => (
                <option key={cls} value={cls}>{cls}</option>
              ))}
            </select>
          </div>

          {/* Nama Lengkap Siswa */}
          <div className="space-y-1.5">
            <label className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
              <UserIcon className="w-3.5 h-3.5 text-amber-500" />
              Siswa Prioritas Tindak Lanjut
            </label>
            <select 
              value={selectedStudentId} 
              onChange={(e) => setSelectedStudentId(e.target.value)}
              disabled={!selectedClass}
              className="w-full px-3 py-2.5 bg-slate-50 border-2 border-slate-200 rounded-xl text-slate-700 font-bold focus:ring-2 focus:ring-amber-500/20 focus:border-amber-500 transition-all text-xs disabled:opacity-50"
            >
              <option value="">-- Pilih Siswa (Total: {priorityStudents.length} Siswa Terdaftar) --</option>
              {priorityStudents.map(stat => (
                <option key={stat.student.id} value={stat.student.id}>
                  {stat.student.name} {stat.details}
                </option>
              ))}
            </select>
          </div>
        </div>

        {/* Tanggal Kasus / Pertemuan untuk Siswa yang Dipilih */}
        {selectedStudentId && (
          <div className="mb-6 p-4 bg-amber-50/60 border border-amber-200/80 rounded-2xl flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3">
            <div className="flex items-center gap-2">
              <Clock className="w-4 h-4 text-amber-600" />
              <span className="text-xs font-bold text-slate-800">Tanggal Kejadian / Pembinaan:</span>
              <input 
                type="date"
                value={selectedDate}
                onChange={(e) => setSelectedDate(e.target.value)}
                className="px-3 py-1.5 bg-white border border-amber-300 rounded-xl text-xs font-bold text-slate-700 focus:outline-none focus:ring-2 focus:ring-amber-500"
              />
            </div>
            <span className="text-[11px] font-semibold text-amber-800">
              Menampilkan data untuk siswa terpilih pada tanggal kejadian di atas
            </span>
          </div>
        )}

        <AnimatePresence mode="wait">
          {selectedStudentId && (
            <motion.div 
              initial={{ opacity: 0, height: 0 }}
              animate={{ opacity: 1, height: 'auto' }}
              exit={{ opacity: 0, height: 0 }}
              className="space-y-6 relative z-10"
            >
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                {/* Inputan Terlambat */}
                <div className="space-y-2 md:col-span-1">
                  <label className="text-sm font-bold text-slate-700 flex items-center gap-2">
                    <Clock className="w-4 h-4 text-slate-400" />
                    Intensitas Terlambat
                  </label>
                  <select 
                    value={terlambat} 
                    onChange={(e) => setTerlambat(e.target.value)}
                    disabled={!isEditing}
                    className="w-full px-4 py-3 bg-white border-2 border-slate-200 rounded-2xl text-slate-700 font-bold focus:ring-4 focus:ring-amber-500/20 focus:border-amber-500 transition-all disabled:bg-slate-50 text-sm"
                  >
                    <option value="Tidak Pernah">Tidak Pernah</option>
                    <option value="Pernah">Pernah</option>
                    <option value="Cukup Sering">Cukup Sering</option>
                    <option value="Sering">Sering</option>
                    <option value="Sangat Sering">Sangat Sering</option>
                  </select>
                </div>

                {/* Status */}
                <div className="space-y-2 md:col-span-1">
                  <label className="text-sm font-bold text-slate-700 flex items-center gap-2">
                    <Activity className="w-4 h-4 text-slate-400" />
                    Status Prioritas
                  </label>
                  <select 
                    value={status} 
                    onChange={(e) => setStatus(e.target.value)}
                    disabled={!isEditing}
                    className="w-full px-4 py-3 bg-white border-2 border-slate-200 rounded-2xl text-slate-700 font-bold focus:ring-4 focus:ring-amber-500/20 focus:border-amber-500 transition-all disabled:bg-slate-50 text-sm"
                  >
                    <option value="Normal">Normal</option>
                    <option value="Perlu Perhatian">Perlu Perhatian</option>
                    <option value="Sangat Mendesak">Sangat Mendesak</option>
                  </select>
                </div>

                {/* Permasalahan */}
                <div className="space-y-2 md:col-span-2">
                  <label className="text-sm font-bold text-slate-700 flex items-center gap-2">
                    <ShieldAlert className="w-4 h-4 text-slate-400" />
                    Permasalahan Utama Siswa
                  </label>
                  <textarea 
                    value={permasalahan} 
                    onChange={(e) => setPermasalahan(e.target.value)}
                    disabled={!isEditing}
                    rows={2}
                    placeholder="Deskripsikan kendala kehadiran, kedisiplinan, atau akademik siswa..."
                    className="w-full p-4 bg-white border-2 border-slate-200 rounded-2xl text-slate-700 font-medium focus:ring-4 focus:ring-amber-500/20 focus:border-amber-500 transition-all disabled:bg-slate-50 text-sm"
                  />
                </div>

                {/* Tindak Lanjut */}
                <div className="space-y-2 md:col-span-2">
                  <label className="text-sm font-bold text-slate-700 flex items-center gap-2">
                    <CheckCircle2 className="w-4 h-4 text-slate-400" />
                    Rencana Tindak Lanjut Wali Kelas
                  </label>
                  <textarea 
                    value={tindakLanjut} 
                    onChange={(e) => setTindakLanjut(e.target.value)}
                    disabled={!isEditing}
                    rows={2}
                    placeholder="Contoh: Pemanggilan orang tua, koordinasi dengan BK/guru mapel, bimbingan berkala..."
                    className="w-full p-4 bg-white border-2 border-slate-200 rounded-2xl text-slate-700 font-medium focus:ring-4 focus:ring-amber-500/20 focus:border-amber-500 transition-all disabled:bg-slate-50 text-sm"
                  />
                </div>
              </div>

              {/* Action Buttons */}
              <div className="flex flex-wrap items-center justify-end gap-3 pt-4 border-t border-slate-100">
                {!isEditing ? (
                  <button
                    onClick={() => setIsEditing(true)}
                    className="px-6 py-2.5 bg-amber-500 hover:bg-amber-600 text-white font-bold rounded-xl flex items-center gap-2 transition-all text-sm shadow-sm"
                  >
                    <Pencil className="w-4 h-4" />
                    Edit Laporan Siswa Ini
                  </button>
                ) : (
                  <>
                    {existingDocId && (
                      <button
                        onClick={() => setIsEditing(false)}
                        className="px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl transition-all text-sm"
                      >
                        Batal
                      </button>
                    )}
                    <button
                      onClick={handleSave}
                      disabled={isSaving}
                      className="px-6 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-xl flex items-center gap-2 transition-all text-sm shadow-sm disabled:opacity-50"
                    >
                      <Save className="w-4 h-4" />
                      {isSaving ? 'Menyimpan...' : 'Simpan Laporan Siswa'}
                    </button>
                  </>
                )}
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </div>

      {/* CARD 2: STATISTIK AKUMULASI KELAS (Ringkasan Kehadiran) */}
      {selectedClass && (
        <div className="bg-white p-6 sm:p-8 rounded-[2rem] shadow-sm border border-slate-100">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-4">
            <div>
              <h3 className="text-lg font-black text-slate-800 tracking-tight flex items-center gap-2">
                <Activity className="w-5 h-5 text-emerald-600" />
                Statistik Akumulasi Kehadiran Kelas {selectedClass}
              </h3>
              <p className="text-xs text-slate-500 font-semibold mt-0.5">
                Periode: {formatIndoDate(startDate)} s.d {formatIndoDate(endDate)} ({classAccumulationStats.totalSesi} Sesi Pertemuan)
              </p>
            </div>
            
            <div className="flex items-center gap-2 bg-slate-50 border border-slate-200 px-3.5 py-1.5 rounded-xl text-xs font-bold text-slate-700">
              <span>Total Siswa: {classAccumulationStats.totalSiswa}</span>
            </div>
          </div>

          <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
            <div className="bg-emerald-50/60 p-3.5 rounded-2xl border border-emerald-100 text-center">
              <div className="text-[11px] font-bold text-emerald-600 uppercase tracking-wider mb-0.5">Hadir</div>
              <div className="text-lg font-black text-emerald-700">
                {classAccumulationStats.hadir}
                <span className="text-[10px] font-bold block text-emerald-600">({classAccumulationStats.hadirPct.toFixed(1)}%)</span>
              </div>
            </div>
            <div className="bg-amber-50/60 p-3.5 rounded-2xl border border-amber-100 text-center">
              <div className="text-[11px] font-bold text-amber-600 uppercase tracking-wider mb-0.5">Sakit</div>
              <div className="text-lg font-black text-amber-700">
                {classAccumulationStats.sakit}
                <span className="text-[10px] font-bold block text-amber-600">({classAccumulationStats.sakitPct.toFixed(1)}%)</span>
              </div>
            </div>
            <div className="bg-sky-50/60 p-3.5 rounded-2xl border border-sky-100 text-center">
              <div className="text-[11px] font-bold text-sky-600 uppercase tracking-wider mb-0.5">Izin</div>
              <div className="text-lg font-black text-sky-700">
                {classAccumulationStats.izin}
                <span className="text-[10px] font-bold block text-sky-600">({classAccumulationStats.izinPct.toFixed(1)}%)</span>
              </div>
            </div>
            <div className="bg-rose-50/60 p-3.5 rounded-2xl border border-rose-100 text-center">
              <div className="text-[11px] font-bold text-rose-600 uppercase tracking-wider mb-0.5">Alpa</div>
              <div className="text-lg font-black text-rose-700">
                {classAccumulationStats.alpa}
                <span className="text-[10px] font-bold block text-rose-600">({classAccumulationStats.alpaPct.toFixed(1)}%)</span>
              </div>
            </div>
            <div className="bg-indigo-50/60 p-3.5 rounded-2xl border border-indigo-100 text-center">
              <div className="text-[11px] font-bold text-indigo-600 uppercase tracking-wider mb-0.5">Dispen</div>
              <div className="text-lg font-black text-indigo-700">
                {classAccumulationStats.dispen}
                <span className="text-[10px] font-bold block text-indigo-600">({classAccumulationStats.dispenPct.toFixed(1)}%)</span>
              </div>
            </div>
            <div className="bg-slate-50 p-3.5 rounded-2xl border border-slate-200 text-center">
              <div className="text-[11px] font-bold text-slate-500 uppercase tracking-wider mb-0.5">Kehadiran</div>
              <div className="text-lg font-black text-slate-800">{classAccumulationStats.persentase.toFixed(1)}%</div>
            </div>
          </div>
        </div>
      )}

      {/* CARD 3: REKAP TINDAK LANJUT BULAN INI (REQUESTED ITEM 3) */}
      {selectedClass && (
        <div id="rekap-form-anchor" className="bg-white p-6 sm:p-8 rounded-[2rem] shadow-sm border border-slate-100 space-y-6">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <div>
              <h3 className="text-xl font-black text-slate-800 tracking-tight flex items-center gap-2">
                <Layers className="w-5 h-5 text-indigo-600" />
                Rekap Tindak Lanjut Bulan Ini
              </h3>
              <p className="text-xs text-slate-500 font-medium mt-1">
                Input akumulasi bentuk penanganan & koordinasi wali kelas untuk periode {currentMonth} ({selectedClass}).
              </p>
            </div>

            {recapEditId && (
              <span className="inline-flex items-center gap-1.5 px-3 py-1 bg-amber-100 text-amber-800 rounded-full text-xs font-bold animate-pulse">
                <Pencil className="w-3.5 h-3.5" /> Mode Edit Rekap
              </span>
            )}
          </div>

          {/* Form Inputan 6 Kategori Rekap Tindak Lanjut via StudentMultiSelectDropdown */}
          {(() => {
            const classStudentsList = students.filter(s => s.class === selectedClass);
            const totalCasesInForm = 
              recapForm.pembinaanMurid.length + 
              recapForm.kontakOrangTua.length + 
              recapForm.koordinasiBK.length + 
              recapForm.koordinasiGuruWali.length + 
              recapForm.koordinasiGuruMapel.length + 
              recapForm.kasusBelumSelesai.length;

            const allUniqueStudentIdsInForm = Array.from(new Set([
              ...recapForm.pembinaanMurid,
              ...recapForm.kontakOrangTua,
              ...recapForm.koordinasiBK,
              ...recapForm.koordinasiGuruWali,
              ...recapForm.koordinasiGuruMapel,
              ...recapForm.kasusBelumSelesai
            ]));

            return (
              <>
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4 bg-slate-50/70 p-5 rounded-2xl border border-slate-200/80">
                  {/* 1. Pembinaan Murid */}
                  <StudentMultiSelectDropdown 
                    label="Pembinaan Murid"
                    icon={<UserCheck className="w-3.5 h-3.5 text-indigo-600" />}
                    themeColor="indigo"
                    description="Bimbingan dan konseling personal oleh wali kelas"
                    students={classStudentsList}
                    selectedIds={recapForm.pembinaanMurid}
                    onChange={(ids) => setRecapForm(f => ({ ...f, pembinaanMurid: ids }))}
                    placeholder="Pilih nama siswa dibina..."
                  />

                  {/* 2. Kontak Orang Tua */}
                  <StudentMultiSelectDropdown 
                    label="Kontak Orang Tua"
                    icon={<PhoneCall className="w-3.5 h-3.5 text-emerald-600" />}
                    themeColor="emerald"
                    description="Komunikasi telepon/pesan, surat pemanggilan, atau kunjungan"
                    students={classStudentsList}
                    selectedIds={recapForm.kontakOrangTua}
                    onChange={(ids) => setRecapForm(f => ({ ...f, kontakOrangTua: ids }))}
                    placeholder="Pilih siswa ortu dihubungi..."
                  />

                  {/* 3. Koordinasi BK */}
                  <StudentMultiSelectDropdown 
                    label="Koordinasi BK"
                    icon={<ShieldAlert className="w-3.5 h-3.5 text-amber-600" />}
                    themeColor="amber"
                    description="Rujukan penanganan dan kolaborasi bersama guru BK"
                    students={classStudentsList}
                    selectedIds={recapForm.koordinasiBK}
                    onChange={(ids) => setRecapForm(f => ({ ...f, koordinasiBK: ids }))}
                    placeholder="Pilih rujukan siswa ke BK..."
                  />

                  {/* 4. Koordinasi Guru Wali */}
                  <StudentMultiSelectDropdown 
                    label="Koordinasi Guru Wali"
                    icon={<Users className="w-3.5 h-3.5 text-sky-600" />}
                    themeColor="sky"
                    description="Koordinasi antarsejawat wali kelas & pembagian tindak lanjut"
                    students={classStudentsList}
                    selectedIds={recapForm.koordinasiGuruWali}
                    onChange={(ids) => setRecapForm(f => ({ ...f, koordinasiGuruWali: ids }))}
                    placeholder="Pilih siswa koordinasi wali..."
                  />

                  {/* 5. Koordinasi Guru Mapel */}
                  <StudentMultiSelectDropdown 
                    label="Koordinasi Guru Mapel"
                    icon={<FileText className="w-3.5 h-3.5 text-violet-600" />}
                    themeColor="violet"
                    description="Pembahasan progres nilai, tugas, & kehadiran mapel"
                    students={classStudentsList}
                    selectedIds={recapForm.koordinasiGuruMapel}
                    onChange={(ids) => setRecapForm(f => ({ ...f, koordinasiGuruMapel: ids }))}
                    placeholder="Pilih koordinasi guru mapel..."
                  />

                  {/* 6. Kasus Belum Selesai */}
                  <StudentMultiSelectDropdown 
                    label="Kasus Belum Selesai"
                    icon={<Clock className="w-3.5 h-3.5 text-rose-600" />}
                    themeColor="rose"
                    description="Siswa dalam pemantauan berkelanjutan & evaluasi rutin"
                    students={classStudentsList}
                    selectedIds={recapForm.kasusBelumSelesai}
                    onChange={(ids) => setRecapForm(f => ({ ...f, kasusBelumSelesai: ids }))}
                    placeholder="Pilih siswa dalam pantauan..."
                  />

                  {/* Keterangan Tambahan */}
                  <div className="space-y-1.5 sm:col-span-2 lg:col-span-3">
                    <label className="text-xs font-bold text-slate-700">Keterangan / Catatan Tambahan (Opsional)</label>
                    <input 
                      type="text"
                      value={recapForm.keterangan}
                      onChange={(e) => setRecapForm(f => ({ ...f, keterangan: e.target.value }))}
                      placeholder="Contoh: Pembinaan intensif pasca UTS, koordinasi berkala dengan guru BK..."
                      className="w-full px-3.5 py-2.5 bg-white border-2 border-slate-200 rounded-xl text-slate-700 font-medium focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 transition-all text-xs"
                    />
                  </div>
                </div>

                {/* Action Buttons for Recap Form */}
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pt-2">
                  <div className="text-xs font-bold text-slate-600 bg-slate-50 sm:bg-transparent p-3 sm:p-0 rounded-xl border sm:border-0 border-slate-200/60 flex flex-wrap items-center gap-2">
                    <span>Total Penanganan:</span>
                    <span className="text-indigo-700 font-black text-sm bg-indigo-50 px-2.5 py-1 rounded-lg border border-indigo-100">
                      {totalCasesInForm} Kasus
                    </span>
                    <span className="text-emerald-700 font-black text-sm bg-emerald-50 px-2.5 py-1 rounded-lg border border-emerald-100">
                      {allUniqueStudentIdsInForm.length} Murid Terlibat
                    </span>
                  </div>

                  <div className="flex items-center gap-3">
                    {recapEditId ? (
                      <button
                        type="button"
                        onClick={resetRecapForm}
                        className="h-10 sm:h-11 px-5 sm:px-6 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl flex items-center justify-center gap-2 transition-all shadow-sm active:scale-95 border border-slate-200 min-w-[140px] flex-1 sm:flex-initial"
                      >
                        <X className="w-4 h-4 text-slate-500" />
                        Batal Edit
                      </button>
                    ) : (
                      <button
                        type="button"
                        onClick={() => {
                          if (monthlyRecaps.length > 0) {
                            handleEditRecap(monthlyRecaps[0]);
                          } else {
                            showToast?.('Belum ada data rekap yang tersimpan untuk diedit', 'info');
                          }
                        }}
                        disabled={monthlyRecaps.length === 0}
                        className="h-10 sm:h-11 px-5 sm:px-6 bg-amber-500 hover:bg-amber-600 text-white text-xs font-bold rounded-xl flex items-center justify-center gap-2 transition-all shadow-sm active:scale-95 disabled:opacity-40 disabled:cursor-not-allowed min-w-[140px] flex-1 sm:flex-initial"
                      >
                        <Pencil className="w-4 h-4" />
                        Edit Rekap
                      </button>
                    )}
                    <button
                      type="button"
                      onClick={handleSaveRecap}
                      disabled={isSavingRecap}
                      className="h-10 sm:h-11 px-5 sm:px-6 bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold rounded-xl flex items-center justify-center gap-2 transition-all shadow-sm active:scale-95 disabled:opacity-50 min-w-[140px] flex-1 sm:flex-initial"
                    >
                      <Save className="w-4 h-4" />
                      {isSavingRecap ? 'Menyimpan...' : recapEditId ? 'Perbarui Rekap' : 'Simpan Rekap'}
                    </button>
                  </div>
                </div>
              </>
            );
          })()}

          {/* TABEL OUTPUT ISIAN REKAP TINDAK LANJUT */}
          <div className="space-y-3 pt-4 border-t border-slate-100">
            <div className="flex items-center justify-between">
              <h4 className="text-sm font-bold text-slate-700">Daftar Rekap Tindak Lanjut ({selectedClass})</h4>
              <span className="text-xs font-semibold text-slate-500">Tersimpan: {monthlyRecaps.length} entri</span>
            </div>

            {isLoadingRecaps ? (
              <div className="py-8 flex justify-center">
                <div className="w-6 h-6 border-3 border-indigo-500/30 border-t-indigo-500 rounded-full animate-spin"></div>
              </div>
            ) : monthlyRecaps.length === 0 ? (
              <div className="p-6 text-center border-2 border-dashed border-slate-200 rounded-2xl bg-slate-50/50">
                <p className="text-xs font-bold text-slate-500">Belum ada rekap yang disimpan untuk kelas ini.</p>
              </div>
            ) : (
              <div className="overflow-x-auto rounded-2xl border border-slate-200">
                <table className="w-full text-left border-collapse min-w-[800px]">
                  <thead>
                    <tr className="bg-slate-100 text-slate-700 text-xs uppercase tracking-wider">
                      <th className="px-3 py-3 font-bold border-b border-slate-200 text-center w-10">No</th>
                      <th className="px-3 py-3 font-bold border-b border-slate-200">Periode / Bulan</th>
                      <th className="px-3 py-3 font-bold border-b border-slate-200 text-center">Pembinaan</th>
                      <th className="px-3 py-3 font-bold border-b border-slate-200 text-center">Kontak Ortu</th>
                      <th className="px-3 py-3 font-bold border-b border-slate-200 text-center">BK</th>
                      <th className="px-3 py-3 font-bold border-b border-slate-200 text-center">Guru Wali</th>
                      <th className="px-3 py-3 font-bold border-b border-slate-200 text-center">Guru Mapel</th>
                      <th className="px-3 py-3 font-bold border-b border-slate-200 text-center">Belum Selesai</th>
                      <th className="px-3 py-3 font-bold border-b border-slate-200 text-center">Total Murid</th>
                      <th className="px-3 py-3 font-bold border-b border-slate-200 text-center w-24">Aksi</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 bg-white text-xs">
                    {monthlyRecaps.map((recap, index) => {
                      const cPembinaan = getCategoryCount(recap.pembinaanMurid);
                      const cKontakOrtu = getCategoryCount(recap.kontakOrangTua);
                      const cBK = getCategoryCount(recap.koordinasiBK);
                      const cGuruWali = getCategoryCount(recap.koordinasiGuruWali);
                      const cGuruMapel = getCategoryCount(recap.koordinasiGuruMapel);
                      const cKasusBelum = getCategoryCount(recap.kasusBelumSelesai);
                      const totalKasus = cPembinaan + cKontakOrtu + cBK + cGuruWali + cGuruMapel + cKasusBelum;

                      const getNamesFromRecap = (idsOrNum: string[] | number | undefined, saved?: string[]) => {
                        if (saved && saved.length > 0) return saved;
                        if (Array.isArray(idsOrNum)) {
                          return idsOrNum.map(id => students.find(s => s.id === id || s.name === id)?.name || id);
                        }
                        return [];
                      };

                      const pembinaanNames = getNamesFromRecap(recap.pembinaanMurid, recap.pembinaanMuridNames);
                      const kontakOrtuNames = getNamesFromRecap(recap.kontakOrangTua, recap.kontakOrangTuaNames);
                      const bkNames = getNamesFromRecap(recap.koordinasiBK, recap.koordinasiBKNames);
                      const guruWaliNames = getNamesFromRecap(recap.koordinasiGuruWali, recap.koordinasiGuruWaliNames);
                      const guruMapelNames = getNamesFromRecap(recap.koordinasiGuruMapel, recap.koordinasiGuruMapelNames);
                      const kasusBelumNames = getNamesFromRecap(recap.kasusBelumSelesai, recap.kasusBelumSelesaiNames);

                      const allNamesInRow = Array.from(new Set([
                        ...pembinaanNames,
                        ...kontakOrtuNames,
                        ...bkNames,
                        ...guruWaliNames,
                        ...guruMapelNames,
                        ...kasusBelumNames
                      ]));
                      const uniqueMuridsInRow = allNamesInRow.length > 0 ? allNamesInRow.length : totalKasus;

                      return (
                        <tr key={recap.id} className="hover:bg-slate-50/80 transition-colors">
                          <td className="px-3 py-3 text-center font-bold text-slate-500">{index + 1}</td>
                          <td className="px-3 py-3 font-bold text-slate-800">
                            <div>{recap.month}</div>
                            {recap.startDate && recap.endDate && (
                              <div className="text-[10px] text-slate-500 font-normal">{recap.startDate} s.d {recap.endDate}</div>
                            )}
                            {recap.keterangan && (
                              <div className="text-[10px] text-slate-400 font-normal italic mt-0.5 truncate max-w-[150px]" title={recap.keterangan}>
                                {recap.keterangan}
                              </div>
                            )}
                          </td>
                          <td className="px-3 py-3 text-center bg-indigo-50/30">
                            <span className="font-bold text-indigo-700">{cPembinaan}</span>
                            {pembinaanNames.length > 0 && (
                              <div className="text-[9px] text-indigo-600 truncate max-w-[100px] mx-auto" title={pembinaanNames.join(', ')}>
                                {pembinaanNames.join(', ')}
                              </div>
                            )}
                          </td>
                          <td className="px-3 py-3 text-center bg-emerald-50/30">
                            <span className="font-bold text-emerald-700">{cKontakOrtu}</span>
                            {kontakOrtuNames.length > 0 && (
                              <div className="text-[9px] text-emerald-600 truncate max-w-[100px] mx-auto" title={kontakOrtuNames.join(', ')}>
                                {kontakOrtuNames.join(', ')}
                              </div>
                            )}
                          </td>
                          <td className="px-3 py-3 text-center bg-amber-50/30">
                            <span className="font-bold text-amber-700">{cBK}</span>
                            {bkNames.length > 0 && (
                              <div className="text-[9px] text-amber-600 truncate max-w-[100px] mx-auto" title={bkNames.join(', ')}>
                                {bkNames.join(', ')}
                              </div>
                            )}
                          </td>
                          <td className="px-3 py-3 text-center bg-sky-50/30">
                            <span className="font-bold text-sky-700">{cGuruWali}</span>
                            {guruWaliNames.length > 0 && (
                              <div className="text-[9px] text-sky-600 truncate max-w-[100px] mx-auto" title={guruWaliNames.join(', ')}>
                                {guruWaliNames.join(', ')}
                              </div>
                            )}
                          </td>
                          <td className="px-3 py-3 text-center bg-violet-50/30">
                            <span className="font-bold text-violet-700">{cGuruMapel}</span>
                            {guruMapelNames.length > 0 && (
                              <div className="text-[9px] text-violet-600 truncate max-w-[100px] mx-auto" title={guruMapelNames.join(', ')}>
                                {guruMapelNames.join(', ')}
                              </div>
                            )}
                          </td>
                          <td className="px-3 py-3 text-center bg-rose-50/30">
                            <span className="font-bold text-rose-700">{cKasusBelum}</span>
                            {kasusBelumNames.length > 0 && (
                              <div className="text-[9px] text-rose-600 truncate max-w-[100px] mx-auto" title={kasusBelumNames.join(', ')}>
                                {kasusBelumNames.join(', ')}
                              </div>
                            )}
                          </td>
                          <td className="px-3 py-3 text-center font-black text-slate-900 bg-slate-100/50">
                            {uniqueMuridsInRow} Murid
                          </td>
                          <td className="px-3 py-3 text-center">
                            {deleteRecapConfirmId === recap.id ? (
                              <div className="flex items-center justify-center gap-1">
                                <button
                                  type="button"
                                  onClick={() => handleDeleteRecap(recap.id)}
                                  className="h-7 px-2 bg-rose-600 hover:bg-rose-700 text-white text-[10px] font-bold rounded-lg transition-all shadow-xs"
                                >
                                  Ya
                                </button>
                                <button
                                  type="button"
                                  onClick={() => setDeleteRecapConfirmId(null)}
                                  className="h-7 px-2 bg-slate-200 hover:bg-slate-300 text-slate-700 text-[10px] font-bold rounded-lg transition-all"
                                >
                                  Batal
                                </button>
                              </div>
                            ) : (
                              <div className="flex items-center justify-center gap-1.5">
                                <button
                                  type="button"
                                  onClick={() => handleEditRecap(recap)}
                                  className="h-7 px-2.5 bg-amber-50 hover:bg-amber-100 text-amber-700 border border-amber-200 rounded-lg flex items-center gap-1 text-[11px] font-bold transition-all hover:border-amber-300 active:scale-95 shadow-2xs"
                                  title="Edit Rekap Ini"
                                >
                                  <Pencil className="w-3 h-3 text-amber-600" />
                                  Edit
                                </button>
                                <button
                                  type="button"
                                  onClick={() => setDeleteRecapConfirmId(recap.id)}
                                  className="h-7 px-2.5 bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 rounded-lg flex items-center gap-1 text-[11px] font-bold transition-all hover:border-rose-300 active:scale-95 shadow-2xs"
                                  title="Hapus Rekap Ini"
                                >
                                  <Trash2 className="w-3 h-3 text-rose-600" />
                                  Hapus
                                </button>
                              </div>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>
      )}

      {/* CARD 4: DAFTAR LAPORAN PER SISWA TERSIMPAN & CETAK PDF */}
      {selectedClass && (
        <div className="bg-white p-6 sm:p-8 rounded-[2rem] shadow-sm border border-slate-100">
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 mb-6">
            <div>
              <h3 className="text-xl font-black text-slate-800 tracking-tight flex items-center gap-2">
                <FileText className="w-5 h-5 text-amber-500" />
                Data Laporan Siswa Tersimpan
              </h3>
              <p className="text-xs text-slate-500 font-medium mt-1">
                Bulan {currentMonth} • Kelas {selectedClass} ({savedReports.length} Laporan Tercatat)
              </p>
            </div>
          </div>

          {isLoadingReports ? (
            <div className="py-12 flex justify-center">
              <div className="w-8 h-8 border-4 border-amber-500/30 border-t-amber-500 rounded-full animate-spin"></div>
            </div>
          ) : savedReports.length === 0 ? (
            <div className="py-10 text-center border-2 border-dashed border-slate-200 rounded-2xl bg-slate-50">
              <p className="text-slate-500 font-bold text-sm">Belum ada laporan tindak lanjut persiswa yang disimpan untuk bulan ini.</p>
              <p className="text-slate-400 text-xs mt-1">Gunakan formulir di atas untuk mencatat tindak lanjut siswa prioritas.</p>
            </div>
          ) : (
            <>
              <div className="overflow-x-auto rounded-2xl border border-slate-200">
                <table className="w-full text-left border-collapse min-w-[800px]">
                  <thead>
                    <tr className="bg-slate-100 text-slate-600 text-xs uppercase tracking-wider">
                      <th className="px-4 py-3 font-bold border-b border-slate-200">Tanggal</th>
                      <th className="px-4 py-3 font-bold border-b border-slate-200">Nama Siswa</th>
                      <th className="px-4 py-3 font-bold border-b border-slate-200">Terlambat</th>
                      <th className="px-4 py-3 font-bold border-b border-slate-200">Ketidakhadiran (A/S/I)</th>
                      <th className="px-4 py-3 font-bold border-b border-slate-200">Permasalahan</th>
                      <th className="px-4 py-3 font-bold border-b border-slate-200">Tindak Lanjut</th>
                      <th className="px-4 py-3 font-bold border-b border-slate-200">Status</th>
                      <th className="px-4 py-3 font-bold border-b border-slate-200 print:hidden text-center w-24">Aksi</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 bg-white">
                    {savedReports.map((report) => {
                      let absenInfo = '-';
                      const studentStat = priorityStudents.find(s => s.student.id === report.studentId);
                      if (studentStat) {
                        absenInfo = studentStat.details.replace(/[()]/g, '');
                      }
                      
                      return (
                      <tr key={report.id} className={`hover:bg-slate-50 transition-colors group ${printSelection !== 'all' && printSelection !== report.id ? 'print:hidden' : ''}`}>
                        <td className="px-4 py-3 text-xs text-slate-700 font-semibold whitespace-nowrap">
                          {report.date ? new Date(report.date).toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' }) : '-'}
                        </td>
                        <td className="px-4 py-3 text-xs font-bold text-slate-800 whitespace-nowrap">
                          {report.studentName}
                        </td>
                        <td className="px-4 py-3 text-xs text-slate-600 font-medium whitespace-nowrap">
                          {report.terlambat}
                        </td>
                        <td className="px-4 py-3 text-xs text-slate-600 font-medium whitespace-nowrap">
                          {absenInfo}
                        </td>
                        <td className="px-4 py-3 text-xs text-slate-600 font-medium min-w-[180px]">
                          {report.permasalahan}
                        </td>
                        <td className="px-4 py-3 text-xs text-slate-600 font-medium min-w-[180px]">
                          {report.tindakLanjut}
                        </td>
                        <td className="px-4 py-3">
                          <span className={`px-2.5 py-1 text-[10px] font-bold rounded-lg whitespace-nowrap ${
                            report.status === 'Normal' ? 'bg-emerald-100 text-emerald-700' :
                            report.status === 'Perlu Perhatian' ? 'bg-amber-100 text-amber-700' :
                            'bg-rose-100 text-rose-700'
                          }`}>
                            {report.status}
                          </span>
                        </td>
                        <td className="px-4 py-3 print:hidden">
                          {deleteConfirmId === report.id ? (
                            <div className="flex items-center justify-center gap-1.5 animate-in fade-in zoom-in duration-200">
                              <button
                                onClick={() => handleDelete(report.id!)}
                                className="px-2.5 py-1 bg-rose-500 hover:bg-rose-600 text-white text-[10px] font-bold rounded-lg transition-colors shadow-sm"
                              >
                                Yakin
                              </button>
                              <button
                                onClick={() => setDeleteConfirmId(null)}
                                className="px-2.5 py-1 bg-slate-200 hover:bg-slate-300 text-slate-700 text-[10px] font-bold rounded-lg transition-colors"
                              >
                                Batal
                              </button>
                            </div>
                          ) : (
                            <div className="flex items-center justify-center gap-1.5">
                              <button
                                onClick={() => {
                                  if (report.date) setSelectedDate(report.date);
                                  if (report.className) setSelectedClass(report.className);
                                  setSelectedStudentId(report.studentId);
                                  setExistingDocId(report.id || '');
                                  setTerlambat(report.terlambat || 'Tidak Pernah');
                                  setPermasalahan(report.permasalahan || '');
                                  setTindakLanjut(report.tindakLanjut || '');
                                  setStatus(report.status || 'Normal');
                                  setIsEditing(true);
                                  window.scrollTo({ top: 0, behavior: 'smooth' });
                                }}
                                className="p-1.5 bg-slate-50 border border-slate-200 hover:border-amber-300 hover:bg-amber-50 text-slate-600 hover:text-amber-600 rounded-lg transition-colors shadow-sm"
                                title="Edit Laporan"
                              >
                                <Pencil className="w-3.5 h-3.5" />
                              </button>
                              <button
                                onClick={() => setDeleteConfirmId(report.id!)}
                                className="p-1.5 bg-slate-50 border border-slate-200 hover:border-rose-300 hover:bg-rose-50 text-slate-600 hover:text-rose-600 rounded-lg transition-colors shadow-sm"
                                title="Hapus Laporan"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          )}
                        </td>
                      </tr>
                    )})}
                  </tbody>
                </table>
              </div>
            </>
          )}

          {/* Opsi & Tombol Cetak / Preview di Bawah Tabel */}
          <div className="mt-8 p-6 bg-gradient-to-b from-slate-50 to-slate-100/90 border-2 border-slate-200/90 rounded-2xl print:hidden shadow-sm space-y-6">
            {/* Header Konfigurasi */}
            <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between pb-4 border-b border-slate-200 gap-2">
              <div>
                <h4 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                  <FileText className="w-4 h-4 text-amber-600" />
                  Konfigurasi & Cetak Dokumen Laporan Wali Kelas
                </h4>
                <p className="text-xs text-slate-500 mt-0.5">
                  Atur cakupan siswa dan tanda tangan dokumen, tinjau pratinjau (preview), lalu cetak atau unduh dokumen PDF resmi.
                </p>
              </div>
              <span className="text-[11px] font-bold px-2.5 py-1 bg-amber-100 text-amber-800 rounded-lg border border-amber-200 whitespace-nowrap">
                Standar Legal Landscape
              </span>
            </div>

            {/* Grid 3 Kolom Simetris & Presisi untuk Dropdown */}
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              {/* Kolom 1: Filter Cakupan Siswa */}
              <div className="space-y-1.5">
                <label className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                  <Users className="w-3.5 h-3.5 text-slate-500" />
                  Filter Cakupan Cetak Siswa
                </label>
                <div className="relative">
                  <select
                    value={printSelection}
                    onChange={(e) => setPrintSelection(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border-2 border-slate-200 rounded-xl text-xs font-bold text-slate-700 focus:ring-2 focus:ring-amber-500/20 focus:border-amber-500 transition-all shadow-sm"
                  >
                    <option value="all">Semua Siswa Tercatat ({savedReports.length} Catatan)</option>
                    {savedReports.length > 0 && (
                      <optgroup label="Cetak Per Siswa Tertentu">
                        {savedReports.map(r => (
                          <option key={r.id} value={r.id!}>{r.studentName}</option>
                        ))}
                      </optgroup>
                    )}
                  </select>
                </div>
              </div>

              {/* Kolom 2: Tanda Tangan Kiri */}
              <div className="space-y-1.5">
                <label className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                  <UserCheck className="w-3.5 h-3.5 text-slate-500" />
                  1. Tanda Tangan Kiri (Mengetahui 1)
                </label>
                <div className="relative">
                  <select
                    value={leftSignerRole}
                    onChange={(e) => setLeftSignerRole(e.target.value as HomeroomSignerRoleType)}
                    className="w-full px-3.5 py-2.5 bg-white border-2 border-slate-200 rounded-xl text-xs font-bold text-slate-700 focus:ring-2 focus:ring-amber-500/20 focus:border-amber-500 transition-all shadow-sm"
                  >
                    <option value="kepala_sekolah">🏫 Kepala Sekolah ({profileData?.namaKepalaSekolah || 'Belum diisi'})</option>
                    <option value="kurikulum">📚 Pihak Kurikulum ({profileData?.namaKurikulum || 'Belum diisi'})</option>
                    <option value="kesiswaan">👥 Pihak Kesiswaan ({profileData?.namaKesiswaan || 'Belum diisi'})</option>
                    <option value="humas">📢 Pihak Humas ({profileData?.namaHumas || 'Belum diisi'})</option>
                    <option value="guru_bk">🤝 Guru BK ({profileData?.namaBK || 'Belum diisi'})</option>
                    <option value="guru_wali">🎓 Guru Wali ({profileData?.namaGuruWali || 'Belum diisi'})</option>
                    <option value="none">➖ Tanpa Ttd Kiri</option>
                  </select>
                </div>
              </div>

              {/* Kolom 3: Tanda Tangan Tengah */}
              <div className="space-y-1.5">
                <label className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                  <UserCheck className="w-3.5 h-3.5 text-slate-500" />
                  2. Tanda Tangan Tengah (Mengetahui 2)
                </label>
                <div className="relative">
                  <select
                    value={midSignerRole}
                    onChange={(e) => setMidSignerRole(e.target.value as HomeroomSignerRoleType)}
                    className="w-full px-3.5 py-2.5 bg-white border-2 border-slate-200 rounded-xl text-xs font-bold text-slate-700 focus:ring-2 focus:ring-amber-500/20 focus:border-amber-500 transition-all shadow-sm"
                  >
                    <option value="none">➖ Tanpa Ttd Tengah (Format 2 TTD)</option>
                    <option value="kurikulum">📚 Pihak Kurikulum ({profileData?.namaKurikulum || 'Belum diisi'})</option>
                    <option value="kesiswaan">👥 Pihak Kesiswaan ({profileData?.namaKesiswaan || 'Belum diisi'})</option>
                    <option value="humas">📢 Pihak Humas ({profileData?.namaHumas || 'Belum diisi'})</option>
                    <option value="guru_bk">🤝 Guru BK ({profileData?.namaBK || 'Belum diisi'})</option>
                    <option value="guru_wali">🎓 Guru Wali ({profileData?.namaGuruWali || 'Belum diisi'})</option>
                    <option value="kepala_sekolah">🏫 Kepala Sekolah ({profileData?.namaKepalaSekolah || 'Belum diisi'})</option>
                  </select>
                </div>
              </div>
            </div>

            {/* Tombol Aksi Simetris & Presisi: Preview & Cetak */}
            <div className="pt-4 border-t border-slate-200 flex flex-col sm:flex-row items-center justify-center gap-4">
              {/* Tombol Preview Sebelum Cetak */}
              <button
                type="button"
                onClick={handlePreviewPDF}
                disabled={isPreviewLoading || isExporting}
                className="w-full sm:w-1/2 max-w-sm px-6 py-3 bg-white hover:bg-amber-50 text-amber-700 border-2 border-amber-300 hover:border-amber-400 rounded-xl font-bold flex items-center justify-center gap-2.5 transition-all active:scale-[0.98] shadow-sm hover:shadow text-xs sm:text-sm disabled:opacity-50"
              >
                {isPreviewLoading ? (
                  <div className="w-4 h-4 border-2 border-amber-600/30 border-t-amber-600 rounded-full animate-spin"></div>
                ) : (
                  <Eye className="w-4 h-4 text-amber-600" />
                )}
                <span>{isPreviewLoading ? 'Menyiapkan Pratinjau...' : 'Pratinjau / Preview Dokumen'}</span>
              </button>

              {/* Tombol Cetak / Unduh PDF */}
              <button
                type="button"
                onClick={handleExportPDF}
                disabled={isExporting || isPreviewLoading}
                className="w-full sm:w-1/2 max-w-sm px-6 py-3 bg-slate-900 hover:bg-black text-white rounded-xl font-bold flex items-center justify-center gap-2.5 transition-all active:scale-[0.98] shadow-md hover:shadow-lg text-xs sm:text-sm disabled:opacity-50"
              >
                {isExporting ? (
                  <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin"></div>
                ) : (
                  <Printer className="w-4 h-4 text-amber-400" />
                )}
                <span>{isExporting ? 'Mengekspor PDF Lengkap...' : 'Cetak / Unduh Laporan Wali Kelas (PDF)'}</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal Preview PDF Laporan Wali Kelas */}
      <AnimatePresence>
        {showPreviewModal && previewPdfUrl && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4 bg-slate-900/70 backdrop-blur-sm print:hidden">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white w-full max-w-5xl rounded-2xl shadow-2xl border border-slate-200 overflow-hidden flex flex-col max-h-[92vh]"
            >
              {/* Modal Header */}
              <div className="px-5 py-4 bg-gradient-to-r from-amber-500/10 via-amber-500/5 to-transparent border-b border-slate-200 flex items-center justify-between">
                <div className="flex items-center gap-3">
                  <div className="p-2 bg-amber-500 text-white rounded-xl shadow-sm">
                    <FileText className="w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-sm sm:text-base font-bold text-slate-800 flex items-center gap-2">
                      Pratinjau Dokumen Laporan Wali Kelas
                      <span className="text-[11px] font-bold px-2 py-0.5 bg-emerald-100 text-emerald-800 rounded-md border border-emerald-200">
                        Standar Legal Landscape
                      </span>
                    </h3>
                    <p className="text-xs text-slate-500">
                      Kelas: <span className="font-semibold text-slate-700">{selectedClass || '-'}</span> | Periode: <span className="font-semibold text-slate-700">{currentMonth}</span>
                    </p>
                  </div>
                </div>
                <button
                  type="button"
                  onClick={handleClosePreview}
                  className="p-2 text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded-xl transition-all"
                  title="Tutup Pratinjau"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              {/* Modal Body: PDF Viewer */}
              <div className="flex-1 p-3 sm:p-4 bg-slate-100 overflow-hidden flex flex-col">
                <div className="w-full flex-1 bg-white rounded-xl shadow-inner border border-slate-300 overflow-hidden">
                  <iframe
                    src={previewPdfUrl}
                    className="w-full h-full min-h-[50vh] sm:min-h-[62vh]"
                    title="Pratinjau PDF Laporan Wali Kelas"
                  />
                </div>
              </div>

              {/* Modal Footer: Simetris & Presisi */}
              <div className="px-5 py-3.5 bg-white border-t border-slate-200 flex flex-col sm:flex-row items-center justify-between gap-3">
                <div className="text-xs text-slate-500 hidden sm:block">
                  Kop surat dan tanda tangan telah diatur presisi, simetris, dan sesuai standar tata naskah sekolah.
                </div>
                <div className="flex items-center gap-3 w-full sm:w-auto justify-end">
                  <button
                    type="button"
                    onClick={handleClosePreview}
                    className="flex-1 sm:flex-none px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-xl transition-all"
                  >
                    Tutup Pratinjau
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      if (previewPdfUrl) {
                        const a = document.createElement('a');
                        a.href = previewPdfUrl;
                        a.download = previewFileName || `Laporan_Wali_Kelas_${selectedClass}_${currentMonth}.pdf`;
                        a.click();
                        showToast?.('Dokumen PDF berhasil diunduh.', 'success');
                      }
                    }}
                    className="flex-1 sm:flex-none px-5 py-2.5 bg-slate-900 hover:bg-black text-white font-bold text-xs rounded-xl flex items-center justify-center gap-2 shadow-md transition-all active:scale-95"
                  >
                    <Printer className="w-4 h-4 text-amber-400" />
                    <span>Cetak / Unduh PDF Sekarang</span>
                  </button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

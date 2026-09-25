import React, { useState, useMemo, useEffect } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Pencil, 
  Plus, 
  Save, 
  Download, 
  Calendar, 
  Building2, 
  FileText, 
  CheckCircle2, 
  Trash2, 
  Search, 
  Filter, 
  Sparkles,
  Edit3,
  BookOpen,
  X,
  Printer,
  FileDown,
  UserCheck
} from 'lucide-react';
import { format } from 'date-fns';
import { id } from 'date-fns/locale';
import { collection, doc, setDoc, deleteDoc, onSnapshot, query, where, Firestore } from 'firebase/firestore';
import { Auth } from 'firebase/auth';
import * as XLSX from 'xlsx-js-style';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { kopSuratBase64 } from './kop-surat-b64';
import { compareClasses, compareStudentsByClass } from './classSortUtils';
import { handleFirestoreError, OperationType } from './firebase';

export type SignerRoleType = 'kepala_sekolah' | 'kurikulum' | 'kesiswaan' | 'humas' | 'guru_wali' | 'guru_bk' | 'none';

export interface Student {
  id: string;
  nisn: string; // Displays as "Nis" in UI as requested
  name: string;
  class: string;
  userId?: string;
}

export interface StudentAssignment {
  id: string;
  title: string;
  className: string;
  date: string;
  scores: Record<string, number | string>; // { [studentId]: score }
  userId: string;
  createdAt?: string;
}

export interface AttendanceSession {
  id: string;
  date: string;
  className: string;
  meetingNumber: number;
  records: Record<string, string>;
  userId?: string;
}

export interface ProfileRecord {
  namaGuruMapel?: string;
  nipGuruMapel?: string;
  mataPelajaran?: string;
  namaKepalaSekolah?: string;
  nipKepalaSekolah?: string;
  jabatanKepalaSekolah?: string;
  namaKurikulum?: string;
  nipKurikulum?: string;
  jabatanKurikulum?: string;
  namaKesiswaan?: string;
  nipKesiswaan?: string;
  jabatanKesiswaan?: string;
  namaHumas?: string;
  nipHumas?: string;
  jabatanHumas?: string;
  namaGuruWali?: string;
  nipGuruWali?: string;
  jabatanGuruWali?: string;
  namaBK?: string;
  nipBK?: string;
  jabatanBK?: string;
  role?: string;
  waliKelasClass?: string;
  tahunPelajaran?: string;
  semester?: string;
  [key: string]: unknown;
}

interface StudentGradesViewProps {
  classList: string[];
  students: Student[];
  attendanceSessions?: AttendanceSession[];
  profileData?: ProfileRecord;
  activeDb?: Firestore;
  activeAuth?: Auth;
  trackOp?: (type: 'read' | 'write', count?: number) => void;
  showToast?: (message: string, type: 'success' | 'info' | 'error') => void;
  initialSubTab?: 'input' | 'preview';
}

const safeSetLocalStorage = (key: string, value: string) => {
  try {
    localStorage.setItem(key, value);
  } catch (e) {
    console.warn(`LocalStorage quota exceeded for key "${key}":`, e);
  }
};

export default function StudentGradesView({
  classList,
  students,
  attendanceSessions = [],
  profileData,
  activeDb,
  activeAuth,
  trackOp,
  showToast,
  initialSubTab = 'input'
}: StudentGradesViewProps) {
  // Effective profile data from props or local storage fallback
  const effectiveProfile: ProfileRecord = useMemo(() => {
    if (profileData && (profileData.mataPelajaran || profileData.tahunPelajaran)) {
      return profileData;
    }
    try {
      const savedUser = localStorage.getItem('kaguci_active_custom_user');
      if (savedUser) {
        const parsed = JSON.parse(savedUser);
        const username = parsed.username || '';
        const savedProfile = localStorage.getItem(`kaguci_profile_${username.toLowerCase()}`);
        if (savedProfile) {
          return { ...profileData, ...JSON.parse(savedProfile) };
        }
      }
    } catch {
      // Ignore
    }
    return (profileData as ProfileRecord) || {};
  }, [profileData]);

  // Sub-menu Tab state: 'input' | 'preview'
  const [activeSubTab, setActiveSubTab] = useState<'input' | 'preview'>(initialSubTab);

  useEffect(() => {
    if (initialSubTab) {
      setActiveSubTab(initialSubTab);
    }
  }, [initialSubTab]);

  // Firestore Assignments state
  const [assignments, setAssignments] = useState<StudentAssignment[]>([]);
  const [isSaving, setIsSaving] = useState(false);

  // Input Nilai Form States
  const [inputDate, setInputDate] = useState(() => format(new Date(), 'yyyy-MM-dd'));
  const [selectedClassInput, setSelectedClassInput] = useState<string>(() => {
    if (classList.length > 0) {
      const sorted = classList.slice().sort(compareClasses);
      return sorted[0];
    }
    return '';
  });
  const [newAssignmentTitle, setNewAssignmentTitle] = useState('');
  const [selectedAssignmentId, setSelectedAssignmentId] = useState<string>('');
  const [scoresInput, setScoresInput] = useState<Record<string, string>>({});
  const [isScoresDisabled, setIsScoresDisabled] = useState<boolean>(false);

  // Preview Nilai Filter States
  const [selectedClassPreview, setSelectedClassPreview] = useState<string>(() => {
    if (classList.length > 0) {
      const sorted = classList.slice().sort(compareClasses);
      return sorted[0];
    }
    return '';
  });
  const [assignmentFilterPreview, setAssignmentFilterPreview] = useState<string>('all');
  const [searchStudentQuery, setSearchStudentQuery] = useState<string>('');

  // Signer states for official reports & Excel export
  const [leftSignerRole, setLeftSignerRole] = useState<SignerRoleType>('kepala_sekolah');
  const [midSignerRole, setMidSignerRole] = useState<SignerRoleType>('none');

  // Print & PDF modal state (Ukuran Kertas F4)
  const [isPrintModalOpen, setIsPrintModalOpen] = useState<boolean>(false);
  const [isExportingPDF, setIsExportingPDF] = useState<boolean>(false);
  const [printOrientationOption, setPrintOrientationOption] = useState<'auto' | 'portrait' | 'landscape'>('auto');
  const [printTargetSubTab, setPrintTargetSubTab] = useState<'preview' | 'input'>('preview');

  // Helper: Retrieve details for a given signer role
  const getSignerDetails = (role: SignerRoleType) => {
    switch (role) {
      case 'kurikulum':
        return {
          title: effectiveProfile.jabatanKurikulum || 'Wakasek Kurikulum',
          name: effectiveProfile.namaKurikulum || '(________________________)',
          nip: effectiveProfile.nipKurikulum ? `NIP. ${effectiveProfile.nipKurikulum}` : '',
          enabled: true
        };
      case 'kesiswaan':
        return {
          title: effectiveProfile.jabatanKesiswaan || 'Wakasek Kesiswaan',
          name: effectiveProfile.namaKesiswaan || '(________________________)',
          nip: effectiveProfile.nipKesiswaan ? `NIP. ${effectiveProfile.nipKesiswaan}` : '',
          enabled: true
        };
      case 'humas':
        return {
          title: effectiveProfile.jabatanHumas || 'Wakasek Humas',
          name: effectiveProfile.namaHumas || '(________________________)',
          nip: effectiveProfile.nipHumas ? `NIP. ${effectiveProfile.nipHumas}` : '',
          enabled: true
        };
      case 'guru_wali':
        return {
          title: effectiveProfile.jabatanGuruWali || 'Guru Wali',
          name: effectiveProfile.namaGuruWali || '(________________________)',
          nip: effectiveProfile.nipGuruWali ? `NIP. ${effectiveProfile.nipGuruWali}` : '',
          enabled: true
        };
      case 'guru_bk':
        return {
          title: effectiveProfile.jabatanBK || 'Guru BK',
          name: effectiveProfile.namaBK || '(________________________)',
          nip: effectiveProfile.nipBK ? `NIP. ${effectiveProfile.nipBK}` : '',
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
          title: effectiveProfile.jabatanKepalaSekolah || 'Kepala Sekolah',
          name: effectiveProfile.namaKepalaSekolah || '(________________________)',
          nip: effectiveProfile.nipKepalaSekolah ? `NIP. ${effectiveProfile.nipKepalaSekolah}` : '',
          enabled: true
        };
    }
  };

  // Helper: Retrieve teacher signer info
  const getTeacherSignerInfo = (targetMapel?: string) => {
    const rawMapel = targetMapel !== undefined ? targetMapel : (effectiveProfile.mataPelajaran || '').trim();
    const mapel = rawMapel ? rawMapel.toUpperCase() : '';
    const teacherTitle = effectiveProfile?.role === 'Wali Kelas'
      ? (effectiveProfile?.waliKelasClass ? `Wali Kelas ${effectiveProfile.waliKelasClass}` : 'Wali Kelas')
      : (mapel ? `Guru Mata Pelajaran ${mapel}` : 'Guru Mata Pelajaran');
    const teacherName = effectiveProfile.namaGuruMapel || '(________________________)';
    const teacherNIP = effectiveProfile.nipGuruMapel ? `NIP. ${effectiveProfile.nipGuruMapel}` : '';
    return { teacherTitle, teacherName, teacherNIP };
  };

  // Auto select default class if list changes
  useEffect(() => {
    if (classList.length > 0) {
      const sorted = classList.slice().sort(compareClasses);
      if (!selectedClassInput) setSelectedClassInput(sorted[0]);
      if (!selectedClassPreview) setSelectedClassPreview(sorted[0]);
    }
  }, [classList, selectedClassInput, selectedClassPreview]);

  // Sync Assignments with Firestore & LocalStorage
  useEffect(() => {
    const uid = activeAuth?.currentUser?.uid || 'guest';
    const storageKey = `kaguci_assignments_${uid}`;

    // 1. Initial local restore
    const saved = localStorage.getItem(storageKey);
    if (saved) {
      try {
        setAssignments(JSON.parse(saved));
      } catch (e) {
        console.error('Error parsing local assignments:', e);
      }
    }

    if (!activeDb || !activeAuth?.currentUser) return;

    // 2. Realtime Firestore listener
    const q = query(
      collection(activeDb, 'gradeAssignments'),
      where('userId', '==', uid)
    );

    const unsubscribe = onSnapshot(
      q,
      (snapshot) => {
        const list: StudentAssignment[] = [];
        snapshot.forEach((docSnap) => {
          list.push({ id: docSnap.id, ...docSnap.data() } as StudentAssignment);
        });
        setAssignments(list);
        safeSetLocalStorage(storageKey, JSON.stringify(list));
      },
      (err) => {
        console.warn('Firestore assignments snapshot error (operating offline):', err);
      }
    );

    return () => unsubscribe();
  }, [activeDb, activeAuth]);

  // Filtered assignments for currently selected class in Input mode
  const currentClassAssignmentsInput = useMemo(() => {
    return assignments
      .filter((a) => a.className === selectedClassInput)
      .sort((a, b) => b.date.localeCompare(a.date));
  }, [assignments, selectedClassInput]);

  // Auto pick first assignment when class changes in Input mode if none selected
  useEffect(() => {
    if (currentClassAssignmentsInput.length > 0) {
      const exists = currentClassAssignmentsInput.some((a) => a.id === selectedAssignmentId);
      if (!exists) {
        setSelectedAssignmentId(currentClassAssignmentsInput[0].id);
      }
    } else {
      setSelectedAssignmentId('');
      setScoresInput({});
    }
  }, [selectedClassInput, currentClassAssignmentsInput, selectedAssignmentId]);

  // Currently selected assignment object
  const selectedAssignmentObj = useMemo(() => {
    return assignments.find((a) => a.id === selectedAssignmentId);
  }, [assignments, selectedAssignmentId]);

  // Load scores when selected assignment changes
  useEffect(() => {
    if (!selectedAssignmentId) {
      setScoresInput({});
      setIsScoresDisabled(false);
      return;
    }
    const found = assignments.find((a) => a.id === selectedAssignmentId);
    if (found) {
      const strScores: Record<string, string> = {};
      const hasScores = Object.keys(found.scores || {}).length > 0;
      Object.entries(found.scores || {}).forEach(([stId, val]) => {
        strScores[stId] = val !== undefined && val !== null ? String(val) : '';
      });
      setScoresInput(strScores);
      setIsScoresDisabled(hasScores);
    } else {
      setScoresInput({});
      setIsScoresDisabled(false);
    }
  }, [selectedAssignmentId, assignments]);

  // Students in selected class (Input mode) sorted by class & name
  const studentsInInputClass = useMemo(() => {
    const classSt = students.filter((s) => s.class === selectedClassInput);
    return classSt.slice().sort(compareStudentsByClass);
  }, [students, selectedClassInput]);

  // Edit Assignment Modal States
  const [isEditingAssignmentModal, setIsEditingAssignmentModal] = useState(false);
  const [editAssignmentTitle, setEditAssignmentTitle] = useState('');
  const [editAssignmentDate, setEditAssignmentDate] = useState('');

  // Delete Assignment Modal States
  const [deleteConfirmAssignment, setDeleteConfirmAssignment] = useState<StudentAssignment | null>(null);
  const [isDeleting, setIsDeleting] = useState(false);

  const handleOpenEditModal = () => {
    const found = assignments.find((a) => a.id === selectedAssignmentId);
    if (!found) {
      showToast?.('Pilih tugas terlebih dahulu untuk diedit.', 'error');
      return;
    }
    setEditAssignmentTitle(found.title);
    setEditAssignmentDate(found.date || inputDate);
    setIsEditingAssignmentModal(true);
  };

  const handleSaveEditAssignment = async () => {
    if (!selectedAssignmentId) return;
    const target = assignments.find((a) => a.id === selectedAssignmentId);
    if (!target) return;

    const trimmedTitle = editAssignmentTitle.trim();
    if (!trimmedTitle) {
      showToast?.('Judul tugas tidak boleh kosong.', 'error');
      return;
    }

    const updatedDoc: StudentAssignment = {
      ...target,
      title: trimmedTitle,
      date: editAssignmentDate || target.date
    };

    setIsSaving(true);
    const uid = activeAuth?.currentUser?.uid || 'guest';
    try {
      if (activeDb && activeAuth?.currentUser) {
        trackOp?.('write', 1);
        await setDoc(doc(activeDb, 'gradeAssignments', selectedAssignmentId), updatedDoc, { merge: true });
      }

      const updatedList = assignments.map((a) => (a.id === selectedAssignmentId ? updatedDoc : a));
      setAssignments(updatedList);
      safeSetLocalStorage(`kaguci_assignments_${uid}`, JSON.stringify(updatedList));

      setIsEditingAssignmentModal(false);
      showToast?.(`Tugas berhasil diperbarui menjadi "${trimmedTitle}"!`, 'success');
    } catch (err) {
      console.error('Error editing assignment:', err);
      if (activeDb) handleFirestoreError(err, OperationType.WRITE, 'gradeAssignments');
      showToast?.('Gagal memperbarui tugas.', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  // Handler: Save New Assignment ("Simpan Tugas Baru")
  const handleSaveNewAssignment = async () => {
    const titleTrimmed = newAssignmentTitle.trim();
    if (!titleTrimmed) {
      showToast?.('Mohon isi Judul Tugas Baru terlebih dahulu.', 'error');
      return;
    }
    if (!selectedClassInput) {
      showToast?.('Mohon pilih kelas terlebih dahulu.', 'error');
      return;
    }

    const uid = activeAuth?.currentUser?.uid || 'guest';
    const newId = `assign_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;
    const newDoc: StudentAssignment = {
      id: newId,
      title: titleTrimmed,
      className: selectedClassInput,
      date: inputDate,
      scores: {}, // Always empty for a new assignment
      userId: uid,
      createdAt: new Date().toISOString()
    };

    setIsSaving(true);
    try {
      if (activeDb && activeAuth?.currentUser) {
        trackOp?.('write', 1);
        await setDoc(doc(activeDb, 'gradeAssignments', newId), newDoc);
      } else {
        // Local only fallback
        const updated = [...assignments, newDoc];
        setAssignments(updated);
        safeSetLocalStorage(`kaguci_assignments_${uid}`, JSON.stringify(updated));
      }

      setScoresInput({});
      setIsScoresDisabled(false);
      setSelectedAssignmentId(newId);
      setNewAssignmentTitle('');
      showToast?.(`Tugas "${titleTrimmed}" berhasil disimpan! Silakan input nilai siswa.`, 'success');
    } catch (err) {
      console.error('Error saving assignment:', err);
      if (activeDb) handleFirestoreError(err, OperationType.WRITE, 'gradeAssignments');
      showToast?.('Gagal menyimpan tugas baru.', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  // Handler: Save Student Scores ("Simpan Nilai")
  const handleSaveScores = async () => {
    if (!selectedAssignmentId) {
      showToast?.('Mohon pilih tugas yang akan dinilai.', 'error');
      return;
    }

    const targetAssignment = assignments.find((a) => a.id === selectedAssignmentId);
    if (!targetAssignment) return;

    setIsSaving(true);
    const uid = activeAuth?.currentUser?.uid || 'guest';

    // Format clean scores object
    const cleanScores: Record<string, number | string> = {};
    Object.entries(scoresInput).forEach(([stId, val]) => {
      const v = (val || '').trim();
      if (v !== '') {
        const num = Number(v);
        cleanScores[stId] = !isNaN(num) ? num : v;
      }
    });

    const updatedAssignment: StudentAssignment = {
      ...targetAssignment,
      scores: cleanScores
    };

    try {
      if (activeDb && activeAuth?.currentUser) {
        trackOp?.('write', 1);
        await setDoc(doc(activeDb, 'gradeAssignments', selectedAssignmentId), updatedAssignment, { merge: true });
      }

      // Update local state and storage
      const updatedList = assignments.map((a) => (a.id === selectedAssignmentId ? updatedAssignment : a));
      setAssignments(updatedList);
      safeSetLocalStorage(`kaguci_assignments_${uid}`, JSON.stringify(updatedList));

      setIsScoresDisabled(true);
      showToast?.(`Nilai siswa untuk "${targetAssignment.title}" berhasil disimpan!`, 'success');
    } catch (err) {
      console.error('Error saving scores:', err);
      if (activeDb) handleFirestoreError(err, OperationType.WRITE, 'gradeAssignments');
      showToast?.('Gagal menyimpan nilai siswa.', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  // Handler: Open Delete Assignment Confirmation Modal
  const handleOpenDeleteConfirm = (assignId: string) => {
    const found = assignments.find((a) => a.id === assignId);
    if (!found) {
      showToast?.('Pilih tugas terlebih dahulu untuk dihapus.', 'error');
      return;
    }
    setDeleteConfirmAssignment(found);
  };

  // Handler: Perform Delete Assignment
  const handleConfirmDeleteAssignment = async () => {
    if (!deleteConfirmAssignment) return;
    const target = deleteConfirmAssignment;
    const uid = activeAuth?.currentUser?.uid || 'guest';
    setIsDeleting(true);

    try {
      if (activeDb && activeAuth?.currentUser) {
        trackOp?.('write', 1);
        await deleteDoc(doc(activeDb, 'gradeAssignments', target.id));
      }
      const updatedList = assignments.filter((a) => a.id !== target.id);
      setAssignments(updatedList);
      safeSetLocalStorage(`kaguci_assignments_${uid}`, JSON.stringify(updatedList));

      if (selectedAssignmentId === target.id) {
        setSelectedAssignmentId('');
      }
      showToast?.(`Tugas "${target.title}" telah berhasil dihapus.`, 'info');
      setDeleteConfirmAssignment(null);
    } catch (err) {
      console.error('Error deleting assignment:', err);
      if (activeDb) handleFirestoreError(err, OperationType.DELETE, 'gradeAssignments');
      showToast?.('Gagal menghapus tugas.', 'error');
    } finally {
      setIsDeleting(false);
    }
  };

  // --- PREVIEW NILAI DATA COMPUTATIONS ---

  // Assignments in current Preview selected class
  const classAssignmentsPreview = useMemo(() => {
    return assignments
      .filter((a) => a.className === selectedClassPreview)
      .sort((a, b) => a.date.localeCompare(b.date));
  }, [assignments, selectedClassPreview]);

  // Filtered assignments based on dropdown "Pilih Tugas (Filter)"
  const activeAssignmentsColumns = useMemo(() => {
    if (assignmentFilterPreview === 'all' || !assignmentFilterPreview) {
      return classAssignmentsPreview;
    }
    return classAssignmentsPreview.filter((a) => a.id === assignmentFilterPreview);
  }, [classAssignmentsPreview, assignmentFilterPreview]);

  // Students in selected preview class
  const studentsInPreviewClass = useMemo(() => {
    let list = students.filter((s) => s.class === selectedClassPreview);
    if (searchStudentQuery.trim()) {
      const q = searchStudentQuery.toLowerCase().trim();
      list = list.filter((s) => s.name.toLowerCase().includes(q) || (s.nisn && s.nisn.toLowerCase().includes(q)));
    }
    return list.slice().sort(compareStudentsByClass);
  }, [students, selectedClassPreview, searchStudentQuery]);

  // Helper: Calculate attendance percentage (rounded integer, no '%' symbol, no decimals)
  const getStudentAttendanceScore = (studentId: string, className: string): number => {
    if (!attendanceSessions || attendanceSessions.length === 0) return 100;
    const classSessions = attendanceSessions.filter((s) => s.className === className);
    if (classSessions.length === 0) return 100;

    let totalRecorded = 0;
    let hadirCount = 0;

    classSessions.forEach((session) => {
      const status = session.records?.[studentId];
      if (status) {
        totalRecorded++;
        if (status === 'Hadir' || status === 'Dispen') {
          hadirCount++;
        }
      }
    });

    if (totalRecorded === 0) return 100;
    return Math.round((hadirCount / totalRecorded) * 100);
  };

  // Handler: Build and export Styled Excel Workbook with official school header
  const generateAndExportExcel = (
    targetClass: string,
    studentList: Student[],
    assignmentList: StudentAssignment[]
  ) => {
    if (studentList.length === 0) {
      showToast?.('Tidak ada data siswa untuk diekspor.', 'error');
      return;
    }

    const rawMapel = (effectiveProfile.mataPelajaran || '').trim();
    const mapel = rawMapel ? rawMapel.toUpperCase() : '';
    const tahun = (effectiveProfile.tahunPelajaran || '').trim();
    const tanggalUpdateStr = format(new Date(), 'dd MMMM yyyy', { locale: id });

    // 1. Table Headers: with "Nilai Rapor" replacing "Rata-Rata"
    const tableHeaders = ['No', 'Nama Lengkap Siswa', 'Nis', 'Nilai Kehadiran'];
    assignmentList.forEach((assign) => {
      tableHeaders.push(assign.title);
    });
    if (assignmentList.length > 0) {
      tableHeaders.push('Nilai Rapor');
    }

    const numCols = Math.max(tableHeaders.length, 4);

    // 2. Data Rows
    const tableDataRows = studentList.map((student, idx) => {
      const rowData: (string | number)[] = [
        idx + 1,
        student.name,
        student.nisn || '-',
        getStudentAttendanceScore(student.id, targetClass)
      ];

      let totalScore = 0;
      let gradedCount = 0;

      assignmentList.forEach((assign) => {
        const val = assign.scores?.[student.id];
        if (val !== undefined && val !== null && val !== '') {
          const num = Number(val);
          if (!isNaN(num)) {
            rowData.push(num);
            totalScore += num;
            gradedCount++;
          } else {
            rowData.push(String(val));
          }
        } else {
          rowData.push('-');
        }
      });

      if (assignmentList.length > 0) {
        rowData.push(gradedCount > 0 ? Math.round(totalScore / gradedCount) : '-');
      }

      return rowData;
    });

    // 3. Official Kop Surat & Title Rows
    const rowTitle1 = ['PEMERINTAH PROVINSI JAWA BARAT'];
    const rowTitle2 = ['DINAS PENDIDIKAN PROVINSI JAWA BARAT'];
    const rowTitle3 = ['CABANG DINAS PENDIDIKAN WILAYAH VI'];
    const rowTitle4 = ['SMA NEGERI 1 CILILIN'];
    const rowTitle5 = ['Jalan Radio Cililin Telp. (022) 6940049 Cililin, Kab. Bandung Barat 40562'];
    const rowTitle6 = ['Website: www.sman1cililin.sch.id | Email: sman1_cililin@yahoo.co.id'];
    const rowTitle7 = [mapel ? `REKAPITULASI NILAI MATA PELAJARAN ${mapel}` : 'REKAPITULASI NILAI MATA PELAJARAN'];
    const rowTitle8 = [tahun ? `TAHUN PELAJARAN ${tahun}` : 'TAHUN PELAJARAN'];

    // Row 9: Split into Left: Tanggal Update Data, Right: Kelas
    const splitCol = Math.max(Math.floor(numCols / 2), 2);
    const rowTitle9 = Array(numCols).fill('');
    rowTitle9[0] = `Tanggal Update Data : ${tanggalUpdateStr}`;
    rowTitle9[splitCol] = `Kelas : ${targetClass}`;

    // Pad title rows with empty strings so all merged cells are initialized
    for (let i = 1; i < numCols; i++) {
      rowTitle1.push('');
      rowTitle2.push('');
      rowTitle3.push('');
      rowTitle4.push('');
      rowTitle5.push('');
      rowTitle6.push('');
      rowTitle7.push('');
      rowTitle8.push('');
    }

    const emptyRow1 = Array(numCols).fill('');
    const emptyRow2 = Array(numCols).fill('');

    // 4. Symmetrical Signature Rows (Seperti Laporan Kehadiran Siswa)
    const leftSigner = getSignerDetails(leftSignerRole);
    const midSigner = getSignerDetails(midSignerRole);
    const is3Signers = leftSigner.enabled && midSigner.enabled;
    const { teacherTitle, teacherName, teacherNIP } = getTeacherSignerInfo(mapel);

    // Calculate symmetrical column span for signatures
    // Ensure both left and right blocks have the exact same width (span)
    const colSpan = is3Signers 
      ? Math.max(2, Math.floor(numCols / 4))
      : Math.max(2, Math.min(4, Math.floor(numCols / 3)));
    
    const leftStartCol = 0;
    const leftEndCol = leftStartCol + colSpan - 1;

    const rightStartCol = numCols - colSpan;
    const rightEndCol = numCols - 1;

    const midStartCol = Math.floor((numCols - colSpan) / 2);
    const midEndCol = midStartCol + colSpan - 1;

    const createSigRow = () => Array(numCols).fill('');

    const sigRow1 = createSigRow();
    if (leftSigner.enabled) sigRow1[leftStartCol] = 'Mengetahui,';
    if (is3Signers) sigRow1[midStartCol] = 'Mengetahui,';
    sigRow1[rightStartCol] = `Cililin, ${tanggalUpdateStr}`;

    const sigRow2 = createSigRow();
    if (leftSigner.enabled) sigRow2[leftStartCol] = leftSigner.title;
    if (is3Signers) sigRow2[midStartCol] = midSigner.title;
    sigRow2[rightStartCol] = teacherTitle;

    const sigRow3 = createSigRow(); // Ruang TTD 1
    const sigRow4 = createSigRow(); // Ruang TTD 2
    const sigRow5 = createSigRow(); // Ruang TTD 3

    const sigRow6 = createSigRow();
    if (leftSigner.enabled) sigRow6[leftStartCol] = leftSigner.name;
    if (is3Signers) sigRow6[midStartCol] = midSigner.name;
    sigRow6[rightStartCol] = teacherName;

    const sigRow7 = createSigRow();
    if (leftSigner.enabled) sigRow7[leftStartCol] = leftSigner.nip;
    if (is3Signers) sigRow7[midStartCol] = midSigner.nip;
    sigRow7[rightStartCol] = teacherNIP;

    const emptySigSpacer1 = createSigRow();
    const emptySigSpacer2 = createSigRow();

    const aoa: (string | number)[][] = [
      rowTitle1, // Row 0
      rowTitle2, // Row 1
      rowTitle3, // Row 2
      rowTitle4, // Row 3
      rowTitle5, // Row 4 (Alamat)
      rowTitle6, // Row 5 (Kontak/Website dengan garis ganda pembatas kop surat)
      emptyRow1, // Row 6 (Pemisah)
      rowTitle7, // Row 7 (Judul Rekapitulasi Mapel)
      rowTitle8, // Row 8 (Tahun Pelajaran)
      rowTitle9, // Row 9 (Tanggal Update Data di kiri, Kelas di kanan)
      emptyRow2, // Row 10 (Pemisah sebelum tabel)
      tableHeaders, // Row 11 (Table Header)
      ...tableDataRows, // Rows 12+ (Data Siswa)
      emptySigSpacer1, // Spacer 1
      emptySigSpacer2, // Spacer 2
      sigRow1, // Row Sig 1: Mengetahui & Tanggal
      sigRow2, // Row Sig 2: Jabatan
      sigRow3, // Row Sig 3: Ruang TTD
      sigRow4, // Row Sig 4: Ruang TTD
      sigRow5, // Row Sig 5: Ruang TTD
      sigRow6, // Row Sig 6: Nama (Bold Underline)
      sigRow7  // Row Sig 7: NIP
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(aoa);

    // 5. Merge Title Rows & Signature Rows across used columns for perfect symmetry
    const startDataRow = 12;
    const endDataRow = startDataRow + tableDataRows.length - 1;
    const sigStartRow = endDataRow + 3; // +1 & +2 are spacers
    const sigRowIndices = [0, 1, 2, 3, 4, 5, 6].map(offset => sigStartRow + offset);

    worksheet['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: numCols - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: numCols - 1 } },
      { s: { r: 2, c: 0 }, e: { r: 2, c: numCols - 1 } },
      { s: { r: 3, c: 0 }, e: { r: 3, c: numCols - 1 } },
      { s: { r: 4, c: 0 }, e: { r: 4, c: numCols - 1 } },
      { s: { r: 5, c: 0 }, e: { r: 5, c: numCols - 1 } },
      { s: { r: 7, c: 0 }, e: { r: 7, c: numCols - 1 } },
      { s: { r: 8, c: 0 }, e: { r: 8, c: numCols - 1 } },
      // Row 9: Left merged for Tanggal Update Data, Right merged for Kelas
      { s: { r: 9, c: 0 }, e: { r: 9, c: splitCol - 1 } },
      { s: { r: 9, c: splitCol }, e: { r: 9, c: numCols - 1 } }
    ];

    // Add Symmetrical Signature Merges
    sigRowIndices.forEach((r) => {
      if (leftSigner.enabled) {
        worksheet['!merges'].push({ s: { r, c: leftStartCol }, e: { r, c: leftEndCol } });
      }
      if (is3Signers) {
        worksheet['!merges'].push({ s: { r, c: midStartCol }, e: { r, c: midEndCol } });
      }
      worksheet['!merges'].push({ s: { r, c: rightStartCol }, e: { r, c: rightEndCol } });
    });

    // 6. Column Widths: Kolom tugas-tugas disamakan ukuran 15 dengan wrap text
    worksheet['!cols'] = tableHeaders.map((header) => {
      if (header === 'No') return { wch: 6 };
      if (header === 'Nis') return { wch: 14 };
      if (header === 'Nilai Kehadiran') return { wch: 16 };
      if (header === 'Nilai Rapor') return { wch: 15 };
      if (header === 'Nama Lengkap Siswa' || header === 'Nama Siswa') return { wch: 28 };
      // Seluruh kolom Tugas-Tugas berukuran presisi 15
      return { wch: 15 };
    });

    // 7. Row Heights
    worksheet['!rows'] = [
      { hpt: 20 }, // Row 0
      { hpt: 20 }, // Row 1
      { hpt: 20 }, // Row 2
      { hpt: 24 }, // Row 3 (SMA NEGERI 1 CILILIN)
      { hpt: 16 }, // Row 4
      { hpt: 18 }, // Row 5 (Kop line dengan double border)
      { hpt: 10 }, // Row 6 (Pemisah)
      { hpt: 22 }, // Row 7
      { hpt: 20 }, // Row 8
      { hpt: 18 }, // Row 9
      { hpt: 10 }, // Row 10 (Pemisah)
      { hpt: 40 }  // Row 11 (Table Header dengan wrapText yang lapang)
    ];

    // Data rows heights
    for (let r = 0; r < tableDataRows.length; r++) {
      worksheet['!rows'].push({ hpt: 20 });
    }

    // Signature spacers and rows heights
    worksheet['!rows'].push(
      { hpt: 12 }, // emptySigSpacer1
      { hpt: 12 }, // emptySigSpacer2
      { hpt: 20 }, // sigRow1
      { hpt: 20 }, // sigRow2
      { hpt: 16 }, // sigRow3
      { hpt: 16 }, // sigRow4
      { hpt: 16 }, // sigRow5
      { hpt: 22 }, // sigRow6 (Nama - bold)
      { hpt: 18 }  // sigRow7 (NIP)
    );

    // 8. Styling Cells
    // Style Kop Surat & Title Rows (0 to 9)
    for (let c = 0; c < numCols; c++) {
      // Rows 0-2: Font 13, Bold, Centered
      for (let r = 0; r <= 2; r++) {
        const ref = XLSX.utils.encode_cell({ r, c });
        if (worksheet[ref]) {
          worksheet[ref].s = {
            font: { bold: true, sz: 13, name: 'Calibri', color: { rgb: '0F172A' } },
            alignment: { horizontal: 'center', vertical: 'center' }
          };
        }
      }

      // Row 3: SMA NEGERI 1 CILILIN (Font 16, Bold, Centered)
      const ref3 = XLSX.utils.encode_cell({ r: 3, c });
      if (worksheet[ref3]) {
        worksheet[ref3].s = {
          font: { bold: true, sz: 16, name: 'Calibri', color: { rgb: '0F172A' } },
          alignment: { horizontal: 'center', vertical: 'center' }
        };
      }

      // Row 4: Alamat (Font 9.5, Italic, Centered)
      const ref4 = XLSX.utils.encode_cell({ r: 4, c });
      if (worksheet[ref4]) {
        worksheet[ref4].s = {
          font: { italic: true, sz: 9.5, name: 'Calibri', color: { rgb: '334155' } },
          alignment: { horizontal: 'center', vertical: 'center' }
        };
      }

      // Row 5: Kontak & Website dengan Double Bottom Border (Garis Resmi Kop Surat)
      const ref5 = XLSX.utils.encode_cell({ r: 5, c });
      if (worksheet[ref5]) {
        worksheet[ref5].s = {
          font: { sz: 9, name: 'Calibri', color: { rgb: '475569' } },
          alignment: { horizontal: 'center', vertical: 'center' },
          border: { bottom: { style: 'double', color: { rgb: '0F172A' } } }
        };
      }

      // Row 7: REKAPITULASI NILAI MATA PELAJARAN (Font 14, Bold, Centered)
      const ref7 = XLSX.utils.encode_cell({ r: 7, c });
      if (worksheet[ref7]) {
        worksheet[ref7].s = {
          font: { bold: true, sz: 14, name: 'Calibri', color: { rgb: '0F172A' } },
          alignment: { horizontal: 'center', vertical: 'center' }
        };
      }

      // Row 8: TAHUN PELAJARAN (Font 12, Bold, Centered)
      const ref8 = XLSX.utils.encode_cell({ r: 8, c });
      if (worksheet[ref8]) {
        worksheet[ref8].s = {
          font: { bold: true, sz: 12, name: 'Calibri', color: { rgb: '0F172A' } },
          alignment: { horizontal: 'center', vertical: 'center' }
        };
      }

      // Row 9: Left side: Tanggal Update Data, Right side: Kelas
      if (c < splitCol) {
        const ref9Left = XLSX.utils.encode_cell({ r: 9, c });
        if (!worksheet[ref9Left]) worksheet[ref9Left] = { t: 's', v: '' };
        worksheet[ref9Left].s = {
          font: { bold: true, sz: 10, name: 'Calibri', color: { rgb: '334155' } },
          alignment: { horizontal: 'left', vertical: 'center' }
        };
      } else {
        const ref9Right = XLSX.utils.encode_cell({ r: 9, c });
        if (!worksheet[ref9Right]) worksheet[ref9Right] = { t: 's', v: '' };
        worksheet[ref9Right].s = {
          font: { bold: true, sz: 11, name: 'Calibri', color: { rgb: '0F172A' } },
          alignment: { horizontal: 'right', vertical: 'center' }
        };
      }
    }

    // Solid default borders for all table columns and rows (No, Nama, Nis, Nilai Kehadiran, Tugas-Tugas, Nilai Rapor)
    const tableBorder = {
      top: { style: 'thin', color: { rgb: '000000' } },
      bottom: { style: 'thin', color: { rgb: '000000' } },
      left: { style: 'thin', color: { rgb: '000000' } },
      right: { style: 'thin', color: { rgb: '000000' } }
    };

    // Row 11: Table Header dengan Wrap Text dan border lengkap
    for (let c = 0; c < numCols; c++) {
      const ref = XLSX.utils.encode_cell({ r: 11, c });
      if (!worksheet[ref]) {
        worksheet[ref] = { t: 's', v: tableHeaders[c] || '' };
      }
      const isRaporCol = tableHeaders[c] === 'Nilai Rapor';
      worksheet[ref].s = {
        font: { bold: true, color: { rgb: isRaporCol ? '15803D' : '0F172A' }, sz: 11, name: 'Calibri' },
        fill: { fgColor: { rgb: isRaporCol ? 'DCFCE7' : 'E2E8F0' } },
        alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
        border: tableBorder
      };
    }

    // Rows 12+: Data Rows dengan border lengkap di setiap sel
    for (let r = startDataRow; r <= endDataRow; r++) {
      for (let c = 0; c < numCols; c++) {
        const ref = XLSX.utils.encode_cell({ r, c });
        if (!worksheet[ref]) {
          worksheet[ref] = { t: 's', v: '-' };
        }

        const headerTitle = tableHeaders[c];
        const isNameCol = headerTitle === 'Nama Lengkap Siswa' || headerTitle === 'Nama Siswa';
        const isRaporCol = headerTitle === 'Nilai Rapor';

        worksheet[ref].s = {
          font: {
            bold: isRaporCol,
            color: { rgb: isRaporCol ? '15803D' : '1E293B' },
            sz: isRaporCol ? 14 : 10,
            name: 'Calibri'
          },
          alignment: { horizontal: isNameCol ? 'left' : 'center', vertical: 'center', wrapText: true },
          border: tableBorder,
          ...(isRaporCol ? { fill: { fgColor: { rgb: 'F0FDF4' } } } : {})
        };
      }
    }

    // Symmetrical Signature Cells: Presisi, Rapih, Centered, Tanpa Border Tabel
    sigRowIndices.forEach((r, idx) => {
      const isNameRow = idx === 5;
      const isNipRow = idx === 6;
      const isTitleRow = idx === 1;

      for (let c = 0; c < numCols; c++) {
        const ref = XLSX.utils.encode_cell({ r, c });
        if (!worksheet[ref]) {
          worksheet[ref] = { t: 's', v: '' };
        }
        worksheet[ref].s = {
          font: {
            name: 'Calibri',
            sz: isNipRow ? 10 : 11,
            bold: isNameRow || isTitleRow,
            underline: isNameRow,
            color: { rgb: isNipRow ? '334155' : '0F172A' }
          },
          alignment: { horizontal: 'center', vertical: 'center' }
        };
      }
    });

    // 9. Page Setup: Ukuran Kertas F4 (Folio 215 x 330 mm)
    // Posisi Potrait, jika kolom tugas banyak (> 3 tugas / > 7 kolom) posisi otomatis jadi Landscape
    const isAutoLandscape = assignmentList.length > 3 || numCols > 7;

    worksheet['!pageSetup'] = {
      orientation: isAutoLandscape ? 'landscape' : 'portrait',
      paperSize: 41, // 41 is Folio / F4 (8.5 x 13 in / 215 x 330 mm)
      fitToWidth: 1,
      fitToHeight: 0,
      fitToPage: true
    };

    worksheet['!margins'] = {
      left: 0.5,
      right: 0.5,
      top: 0.6,
      bottom: 0.6,
      header: 0.3,
      footer: 0.3
    };

    worksheet['!sheetPr'] = {
      pageSetUpPr: { fitToPage: true }
    };

    // 10. Create workbook & save
    const workbook = XLSX.utils.book_new();
    const cleanSheetName = `Nilai ${targetClass}`.substring(0, 31);
    XLSX.utils.book_append_sheet(workbook, worksheet, cleanSheetName);

    const dateStr = format(new Date(), 'yyyyMMdd');
    const filename = `Rekap_Nilai_Siswa_Kelas_${targetClass.replace(/\s+/g, '_')}_${dateStr}.xlsx`;
    XLSX.writeFile(workbook, filename);

    showToast?.(`Rekap nilai kelas ${targetClass} berhasil diekspor (Format Cetak F4 ${isAutoLandscape ? 'Landscape' : 'Portrait'})!`, 'success');
  };

  // Handler: Build and export / print PDF on F4 paper with auto landscape/portrait
  const generateAndExportPDF = (
    targetClass: string,
    studentList: Student[],
    assignmentList: StudentAssignment[],
    action: 'download' | 'print' = 'download',
    forcedOrientation?: 'auto' | 'portrait' | 'landscape'
  ) => {
    if (studentList.length === 0) {
      showToast?.('Tidak ada data siswa untuk dicetak.', 'error');
      return;
    }

    const rawMapel = (effectiveProfile.mataPelajaran || '').trim();
    const mapel = rawMapel ? rawMapel.toUpperCase() : '';
    const tahun = (effectiveProfile.tahunPelajaran || '').trim();
    const tanggalUpdateStr = format(new Date(), 'dd MMMM yyyy', { locale: id });

    const numCols = assignmentList.length + 4 + (assignmentList.length > 0 ? 1 : 0);
    const orientation = forcedOrientation && forcedOrientation !== 'auto'
      ? forcedOrientation
      : (assignmentList.length > 3 || numCols > 7 ? 'landscape' : 'portrait');
    const isLandscape = orientation === 'landscape';

    // Kertas ukuran F4: 215 mm x 330 mm
    const doc = new jsPDF({
      orientation,
      unit: 'mm',
      format: [215, 330]
    });

    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const marginX = 14;

    // Kop Surat Header
    const imgX = marginX;
    const imgY = 8.5; // 2 spasi dari tepi atas
    const imgWidth = pageWidth - (marginX * 2);
    const imgHeight = isLandscape ? 48 : (imgWidth * (341 / 1450));

    try {
      doc.addImage(kopSuratBase64, 'PNG', imgX, imgY, imgWidth, imgHeight);
    } catch (e) {
      console.error('Failed to add custom header image in PDF', e);
    }

    const titleY = imgY + imgHeight + 7;
    doc.setFontSize(13);
    doc.setFont('helvetica', 'bold');
    const titleText = mapel ? `REKAPITULASI NILAI MATA PELAJARAN ${mapel}` : 'REKAPITULASI NILAI MATA PELAJARAN';
    doc.text(titleText, pageWidth / 2, titleY, { align: 'center' });

    doc.setFontSize(11);
    const subTitleText = tahun ? `TAHUN PELAJARAN ${tahun}` : 'TAHUN PELAJARAN';
    doc.text(subTitleText, pageWidth / 2, titleY + 6, { align: 'center' });

    doc.setFontSize(9.5);
    doc.setFont('helvetica', 'bold');
    doc.text(`Tanggal Update Data : ${tanggalUpdateStr}`, marginX, titleY + 14);
    doc.text(`Kelas : ${targetClass}`, pageWidth - marginX, titleY + 14, { align: 'right' });

    const tableHeaders = ['No', 'Nama Lengkap Siswa', 'Nis', 'Nilai Kehadiran'];
    assignmentList.forEach((assign) => {
      tableHeaders.push(assign.title);
    });
    if (assignmentList.length > 0) {
      tableHeaders.push('Nilai Rapor');
    }

    const tableBody = studentList.map((student, idx) => {
      const row: (string | number)[] = [
        idx + 1,
        student.name,
        student.nisn || '-',
        getStudentAttendanceScore(student.id, targetClass)
      ];

      let totalScore = 0;
      let gradedCount = 0;

      assignmentList.forEach((assign) => {
        const val = assign.scores?.[student.id];
        if (val !== undefined && val !== null && val !== '') {
          const num = Number(val);
          if (!isNaN(num)) {
            row.push(num);
            totalScore += num;
            gradedCount++;
          } else {
            row.push(String(val));
          }
        } else {
          row.push('-');
        }
      });

      if (assignmentList.length > 0) {
        row.push(gradedCount > 0 ? Math.round(totalScore / gradedCount) : '-');
      }

      return row;
    });

    const tableStartY = titleY + 18;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const columnStyles: Record<number, any> = {
      0: { halign: 'center', cellWidth: 10 },
      1: { halign: 'left', cellWidth: isLandscape ? 45 : 36 },
      2: { halign: 'center', cellWidth: 18 },
      3: { halign: 'center', cellWidth: 20 }
    };

    // Ensure all assignment and score columns are centered
    for (let i = 4; i < tableHeaders.length; i++) {
      columnStyles[i] = { halign: 'center' };
    }

    if (assignmentList.length > 0) {
      columnStyles[tableHeaders.length - 1] = {
        halign: 'center',
        fontStyle: 'bold',
        fillColor: [240, 253, 244]
      };
    }

    const autoTableOpts = {
      startY: tableStartY,
      head: [tableHeaders],
      body: tableBody,
      theme: 'grid' as const,
      headStyles: {
        fillColor: [226, 232, 240],
        textColor: [15, 23, 42],
        fontStyle: 'bold',
        halign: 'center',
        fontSize: 8.5,
        lineWidth: 0.2,
        lineColor: [0, 0, 0]
      },
      styles: {
        fontSize: 8,
        cellPadding: 2,
        lineWidth: 0.2,
        lineColor: [0, 0, 0],
        textColor: [30, 41, 59]
      },
      columnStyles
    };

    if (typeof autoTable === 'function') {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      autoTable(doc, autoTableOpts as any);
    } else if (typeof (doc as Record<string, unknown>).autoTable === 'function') {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (doc as any).autoTable(autoTableOpts);
    }

    // Signatures
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const lastTableY = (doc as any).lastAutoTable?.finalY || (tableStartY + 60);
    let sigY = lastTableY + 12;
    if (sigY + 45 > pageHeight) {
      doc.addPage();
      sigY = 25;
    }

    const leftSigner = getSignerDetails(leftSignerRole);
    const midSigner = getSignerDetails(midSignerRole);
    const is3Signers = leftSigner.enabled && midSigner.enabled;
    const { teacherTitle, teacherName, teacherNIP } = getTeacherSignerInfo(mapel);

    const d = new Date();
    const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    const dateStr = `Cililin, ${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;

    doc.setFontSize(9.5);
    doc.setFont('helvetica', 'normal');

    if (is3Signers) {
      const leftX = 14;
      const midX = (pageWidth / 2) - 25;
      const rightX = pageWidth - 65;

      // Kiri
      doc.text('Mengetahui,', leftX, sigY);
      doc.text(leftSigner.title, leftX, sigY + 5);
      doc.setFont('helvetica', 'bold');
      doc.text(leftSigner.name, leftX, sigY + 25);
      doc.setFont('helvetica', 'normal');
      if (leftSigner.nip) doc.text(leftSigner.nip, leftX, sigY + 30);

      // Tengah
      doc.text('Mengetahui,', midX, sigY);
      doc.text(midSigner.title, midX, sigY + 5);
      doc.setFont('helvetica', 'bold');
      doc.text(midSigner.name, midX, sigY + 25);
      doc.setFont('helvetica', 'normal');
      if (midSigner.nip) doc.text(midSigner.nip, midX, sigY + 30);

      // Kanan
      doc.text(dateStr, rightX, sigY);
      doc.text(teacherTitle, rightX, sigY + 5);
      doc.setFont('helvetica', 'bold');
      doc.text(teacherName, rightX, sigY + 25);
      doc.setFont('helvetica', 'normal');
      if (teacherNIP) doc.text(teacherNIP, rightX, sigY + 30);
    } else {
      const activeLeft = leftSigner.enabled ? leftSigner : (midSigner.enabled ? midSigner : null);
      const rightX = pageWidth - 80;

      if (activeLeft) {
        const leftX = 20;
        doc.text('Mengetahui,', leftX, sigY);
        doc.text(activeLeft.title, leftX, sigY + 5);
        doc.setFont('helvetica', 'bold');
        doc.text(activeLeft.name, leftX, sigY + 25);
        doc.setFont('helvetica', 'normal');
        if (activeLeft.nip) doc.text(activeLeft.nip, leftX, sigY + 30);
      }

      doc.text(dateStr, rightX, sigY);
      doc.text(teacherTitle, rightX, sigY + 5);
      doc.setFont('helvetica', 'bold');
      doc.text(teacherName, rightX, sigY + 25);
      doc.setFont('helvetica', 'normal');
      if (teacherNIP) doc.text(teacherNIP, rightX, sigY + 30);
    }

    const dateFile = format(new Date(), 'yyyyMMdd');
    const fileName = `Rekap_Nilai_${targetClass.replace(/\s+/g, '_')}_F4_${dateFile}.pdf`;

    // Generate PDF and handle action
    if (action === 'print') {
      try {
        const blob = doc.output('blob');
        const url = URL.createObjectURL(blob);
        
        // Open a blank window and inject an iframe for the PDF
        // This is extremely compatible and avoids many "blank page" issues with direct blob URL tabs
        const win = window.open('', '_blank');
        if (win) {
          win.document.write(`
            <html>
              <head>
                <title>Cetak PDF F4 - ${targetClass}</title>
                <style>
                  body, html { margin: 0; padding: 0; height: 100%; overflow: hidden; }
                  iframe { width: 100%; height: 100%; border: none; }
                </style>
              </head>
              <body>
                <iframe src="${url}"></iframe>
              </body>
            </html>
          `);
          win.document.close();
          showToast?.('Dokumen sedang dibuka untuk dicetak...', 'info');
        } else {
          doc.save(fileName);
          showToast?.('Gagal membuka jendela cetak (Popup terblokir). File diunduh otomatis.', 'warning');
        }
      } catch (err) {
        console.error('Print error:', err);
        doc.save(fileName);
        showToast?.('Gagal memproses cetak. File diunduh sebagai cadangan.', 'warning');
      }
    } else {
      doc.save(fileName);
      showToast?.(`File PDF ${fileName} (Ukuran F4) berhasil diunduh!`, 'success');
    }
  };

  // Handler: Export Preview Nilai to Excel (.xlsx)
  const handleExportExcelPreview = () => {
    generateAndExportExcel(selectedClassPreview, studentsInPreviewClass, activeAssignmentsColumns);
  };

  // Handler: Export Input Nilai to Excel (.xlsx)
  const handleExportExcelInput = () => {
    const assignmentsForClass = currentClassAssignmentsInput.slice().sort((a, b) => a.date.localeCompare(b.date));
    generateAndExportExcel(selectedClassInput, studentsInInputClass, assignmentsForClass);
  };

  // Handler: Print / Download PDF
  const handleOpenPrintModal = (target: 'preview' | 'input') => {
    setPrintTargetSubTab(target);
    setIsPrintModalOpen(true);
  };

  const handleExecutePDFAction = (action: 'download' | 'print') => {
    setIsExportingPDF(true);
    setTimeout(() => {
      try {
        if (printTargetSubTab === 'preview') {
          generateAndExportPDF(
            selectedClassPreview,
            studentsInPreviewClass,
            activeAssignmentsColumns,
            action,
            printOrientationOption
          );
        } else {
          const assignmentsForClass = currentClassAssignmentsInput.slice().sort((a, b) => a.date.localeCompare(b.date));
          generateAndExportPDF(
            selectedClassInput,
            studentsInInputClass,
            assignmentsForClass,
            action,
            printOrientationOption
          );
        }
        setIsPrintModalOpen(false);
      } catch (err) {
        console.error('Error generating PDF:', err);
        showToast?.('Gagal memproses dokumen PDF.', 'error');
      } finally {
        setIsExportingPDF(false);
      }
    }, 400);
  };

  return (
    <div className="space-y-6">
      {/* Main Content Area based on activeSubTab */}
      <AnimatePresence mode="wait">
        {activeSubTab === 'input' ? (
          <motion.div
            key="subtab-input"
            initial={{ opacity: 0, y: 10 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -10 }}
            transition={{ duration: 0.2 }}
            className="space-y-6"
          >
            {/* Form Control Box */}
            <div className="bg-white rounded-3xl p-6 sm:p-8 border-2 border-slate-400 shadow-sm space-y-6">
              <div className="flex items-center gap-3 pb-4 border-b-2 border-slate-200">
                <div className="w-10 h-10 rounded-2xl bg-[#8dc63f]/15 text-[#7bc025] flex items-center justify-center font-bold shrink-0">
                  <BookOpen className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base sm:text-lg font-black text-slate-800">
                    Konfigurasi Tugas & Kelas
                  </h3>
                  <p className="text-xs text-slate-500 font-medium">
                    Atur tanggal, pilih kelas, buat tugas baru, atau pilih &amp; kelola tugas yang sudah ada.
                  </p>
                </div>
              </div>

              {/* Form Section */}
              <div className="space-y-5">
                {/* Row 1: Tanggal & Pilih Kelas */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
                  {/* 1. Tanggal (Date) */}
                  <div className="space-y-1.5">
                    <label className="text-xs font-extrabold text-slate-700 flex items-center gap-1.5">
                      <Calendar className="w-3.5 h-3.5 text-[#8dc63f]" />
                      Tanggal Tugas
                    </label>
                    <input
                      type="date"
                      value={inputDate}
                      onChange={(e) => setInputDate(e.target.value)}
                      className="w-full p-3.5 bg-slate-50 border-2 border-slate-200 rounded-2xl font-bold text-slate-800 text-xs sm:text-sm focus:ring-2 focus:ring-[#8dc63f]/20 focus:border-[#8dc63f] transition-all outline-none"
                    />
                  </div>

                  {/* 2. Pilih Kelas (Dropdown) */}
                  <div className="space-y-1.5">
                    <label className="text-xs font-extrabold text-slate-700 flex items-center gap-1.5">
                      <Building2 className="w-3.5 h-3.5 text-[#8dc63f]" />
                      Pilih Kelas
                    </label>
                    <select
                      value={selectedClassInput}
                      onChange={(e) => setSelectedClassInput(e.target.value)}
                      className="w-full p-3.5 bg-slate-50 border-2 border-slate-200 rounded-2xl font-bold text-slate-800 text-xs sm:text-sm focus:ring-2 focus:ring-[#8dc63f]/20 focus:border-[#8dc63f] transition-all outline-none"
                    >
                      {classList.length === 0 ? (
                        <option value="">-- Belum ada kelas --</option>
                      ) : (
                        classList
                          .slice()
                          .sort(compareClasses)
                          .map((c) => (
                            <option key={c} value={c}>
                              Kelas {c}
                            </option>
                          ))
                      )}
                    </select>
                  </div>
                </div>

                {/* Row 2: Input Judul Tugas Baru & Simpan Tugas Baru (Symmetrical) */}
                <div className="space-y-1.5 pt-1">
                  <label className="text-xs font-extrabold text-slate-700 flex items-center gap-1.5">
                    <Plus className="w-3.5 h-3.5 text-[#8dc63f]" />
                    Input Judul Tugas Baru
                  </label>
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 items-center">
                    <input
                      type="text"
                      placeholder="Contoh: Tugas 1 Bangun Ruang / UH 1 Trigonometri"
                      value={newAssignmentTitle}
                      onChange={(e) => setNewAssignmentTitle(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') {
                          e.preventDefault();
                          handleSaveNewAssignment();
                        }
                      }}
                      className="sm:col-span-2 w-full p-3.5 bg-slate-50 border-2 border-slate-200 rounded-2xl font-bold text-slate-800 text-xs sm:text-sm focus:ring-2 focus:ring-[#8dc63f]/20 focus:border-[#8dc63f] transition-all outline-none placeholder:text-slate-400"
                    />
                    <button
                      type="button"
                      onClick={handleSaveNewAssignment}
                      disabled={isSaving || !newAssignmentTitle.trim()}
                      className="sm:col-span-1 w-full p-3.5 bg-[#8dc63f] hover:bg-[#7bc025] disabled:bg-slate-300 text-white font-black text-xs sm:text-sm rounded-2xl transition-all shadow-md hover:shadow-lg active:scale-95 flex items-center justify-center gap-2 cursor-pointer"
                    >
                      <Save className="w-4 h-4" />
                      <span>Simpan Tugas Baru</span>
                    </button>
                  </div>
                </div>

                {/* Row 3: Pilih Tugas Dibuat (Below Input Judul Tugas Baru) + Edit & Hapus Buttons */}
                <div className="space-y-1.5 pt-3 border-t border-slate-100">
                  <label className="text-xs font-extrabold text-slate-700 flex items-center gap-1.5">
                    <FileText className="w-3.5 h-3.5 text-[#8dc63f]" />
                    Pilih Tugas Dibuat
                  </label>
                  <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-3">
                    <select
                      value={selectedAssignmentId}
                      onChange={(e) => setSelectedAssignmentId(e.target.value)}
                      disabled={currentClassAssignmentsInput.length === 0}
                      className="flex-1 p-3.5 bg-slate-50 border-2 border-slate-200 rounded-2xl font-bold text-slate-800 text-xs sm:text-sm focus:ring-2 focus:ring-[#8dc63f]/20 focus:border-[#8dc63f] transition-all outline-none disabled:bg-slate-100 disabled:text-slate-400"
                    >
                      {currentClassAssignmentsInput.length === 0 ? (
                        <option value="">-- Belum ada tugas dibuat untuk kelas {selectedClassInput} --</option>
                      ) : (
                        currentClassAssignmentsInput.map((a) => (
                          <option key={a.id} value={a.id}>
                            {a.title} ({a.date})
                          </option>
                        ))
                      )}
                    </select>

                    {selectedAssignmentId && (
                      <div className="flex items-center gap-2 shrink-0">
                        <button
                          type="button"
                          onClick={handleOpenEditModal}
                          className="px-4 py-3.5 bg-[#8dc63f] hover:bg-[#7bc025] text-white font-black text-xs sm:text-sm rounded-2xl transition-all flex items-center justify-center gap-1.5 cursor-pointer active:scale-95 shadow-md hover:shadow-lg"
                          title="Edit Judul & Tanggal Tugas"
                        >
                          <Edit3 className="w-4 h-4 text-white" />
                          <span>Edit Tugas</span>
                        </button>

                        <button
                          type="button"
                          onClick={() => handleOpenDeleteConfirm(selectedAssignmentId)}
                          className="px-4 py-3.5 bg-[#8dc63f] hover:bg-[#7bc025] text-white font-black text-xs sm:text-sm rounded-2xl transition-all flex items-center justify-center gap-1.5 cursor-pointer active:scale-95 shadow-md hover:shadow-lg"
                          title="Hapus Tugas Ini"
                        >
                          <Trash2 className="w-4 h-4 text-white" />
                          <span>Hapus Tugas</span>
                        </button>
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </div>

            {/* Table Input Nilai Box */}
            <div className="bg-white rounded-3xl p-6 sm:p-8 border-2 border-slate-400 shadow-sm space-y-6">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b-2 border-slate-200">
                <div>
                  <div className="flex items-center gap-2">
                    <h3 className="text-base sm:text-lg font-black text-slate-800">
                      Tabel Input Nilai Siswa
                    </h3>
                    <span className="px-2.5 py-0.5 bg-[#8dc63f]/15 text-[#7bc025] text-[10px] font-black rounded-full uppercase tracking-wider">
                      Kelas {selectedClassInput || '-'}
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 font-medium mt-0.5">
                    {selectedAssignmentId ? (
                      <>
                        Menginput nilai untuk:{' '}
                        <span className="font-bold text-[#7bc025]">
                          {assignments.find((a) => a.id === selectedAssignmentId)?.title}
                        </span>
                      </>
                    ) : (
                      'Menampilkan daftar seluruh siswa pada kelas terpilih.'
                    )}
                  </p>
                </div>

                <div className="flex items-center gap-2">
                  {isScoresDisabled && selectedAssignmentId && (
                    <button
                      type="button"
                      onClick={() => {
                        setIsScoresDisabled(false);
                        showToast?.('Mode edit nilai diaktifkan.', 'info');
                      }}
                      className="px-4 py-2.5 bg-[#8dc63f] hover:bg-[#7bc025] text-white font-extrabold text-xs sm:text-sm rounded-xl transition-all flex items-center gap-2 cursor-pointer shadow-md active:scale-[0.98]"
                    >
                      <Pencil className="w-4 h-4 text-white" />
                      <span>Edit Nilai</span>
                    </button>
                  )}

                  <button
                    type="button"
                    onClick={handleSaveScores}
                    disabled={isSaving || isScoresDisabled}
                    className="px-6 py-2.5 bg-[#8dc63f] hover:bg-[#7bc025] disabled:bg-slate-300 disabled:cursor-not-allowed text-white font-black text-xs sm:text-sm rounded-xl transition-all shadow-md hover:shadow-lg active:scale-95 flex items-center gap-2 cursor-pointer"
                  >
                    <Save className="w-4 h-4" />
                    <span>{isSaving ? 'Menyimpan...' : isScoresDisabled ? 'Nilai Tersimpan' : 'Simpan Nilai'}</span>
                  </button>
                </div>
              </div>

              {/* Table rendering: ALWAYS shows student list for selected class */}
              {studentsInInputClass.length === 0 ? (
                <div className="text-center py-12 text-slate-500 italic">
                  Belum ada data siswa di kelas {selectedClassInput}. Silakan tambahkan data siswa terlebih dahulu.
                </div>
              ) : (
                <div className="space-y-4">
                  {!selectedAssignmentId && (
                    <div className="p-3.5 bg-amber-50 border border-amber-200 rounded-2xl flex items-center gap-2 text-xs font-bold text-amber-800">
                      <Sparkles className="w-4 h-4 text-amber-600 shrink-0" />
                      <span>
                        Siswa kelas <b>{selectedClassInput}</b> siap dinilai. Ketik judul tugas baru di atas dan klik <b>Simpan Tugas Baru</b>, atau pilih tugas yang sudah ada.
                      </span>
                    </div>
                  )}

                  <div className="overflow-x-auto max-h-[550px] border border-slate-300 rounded-2xl scrollbar-thin shadow-xs">
                    <table className="w-full text-left text-xs sm:text-sm border-collapse border border-slate-300">
                      <thead className="sticky top-0 bg-slate-100 z-10">
                        <tr>
                          <th className="p-3 sm:p-3.5 font-black text-slate-800 w-16 text-center border border-slate-300 bg-slate-100">No</th>
                          <th className="p-3 sm:p-3.5 font-black text-slate-800 border border-slate-300 bg-slate-100">Nama Lengkap Siswa</th>
                          <th className="p-3 sm:p-3.5 font-black text-slate-800 w-32 text-center border border-slate-300 bg-slate-100">Nis</th>
                          <th className="p-3 sm:p-3.5 font-black text-slate-800 min-w-[160px] text-center border border-slate-300 bg-slate-100">
                            {selectedAssignmentObj ? (
                              <div className="flex flex-col items-center">
                                <span className="text-slate-900 font-extrabold">{selectedAssignmentObj.title}</span>
                                <span className="text-[10px] text-slate-500 font-medium">{selectedAssignmentObj.date || inputDate}</span>
                              </div>
                            ) : (
                              <span>Nilai (0-100)</span>
                            )}
                          </th>
                        </tr>
                      </thead>
                      <tbody className="bg-white">
                        {studentsInInputClass.map((student, idx) => {
                          const currentVal = scoresInput[student.id] ?? '';
                          return (
                            <tr key={student.id} className="hover:bg-[#8dc63f]/10 transition-colors">
                              <td className="p-3 sm:p-3.5 text-center font-bold text-slate-600 border border-slate-300">{idx + 1}</td>
                              <td className="p-3 sm:p-3.5 font-extrabold text-slate-800 border border-slate-300">{student.name}</td>
                              <td className="p-3 sm:p-3.5 font-semibold text-slate-600 text-center border border-slate-300">{student.nisn || '-'}</td>
                              <td className="p-3 sm:p-3.5 text-center border border-slate-300">
                                <input
                                  type="number"
                                  min="0"
                                  max="100"
                                  placeholder="0-100"
                                  disabled={isScoresDisabled}
                                  value={currentVal}
                                  onChange={(e) => {
                                    const val = e.target.value;
                                    setScoresInput((prev) => ({
                                      ...prev,
                                      [student.id]: val
                                    }));
                                  }}
                                  className="w-28 p-2.5 text-center font-black bg-white border-2 border-slate-200 rounded-xl focus:border-[#8dc63f] focus:ring-2 focus:ring-[#8dc63f]/20 disabled:bg-slate-100 disabled:text-slate-500 disabled:border-slate-200 disabled:cursor-not-allowed transition-all outline-none text-slate-900"
                                />
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {studentsInInputClass.length > 0 && (
                <div className="flex flex-col sm:flex-row justify-end items-center gap-3 pt-2">
                  {isScoresDisabled && selectedAssignmentId && (
                    <button
                      type="button"
                      onClick={() => {
                        setIsScoresDisabled(false);
                        showToast?.('Mode edit nilai diaktifkan.', 'info');
                      }}
                      className="w-full sm:w-auto px-6 py-3.5 bg-[#8dc63f] hover:bg-[#7bc025] text-white font-extrabold text-xs sm:text-sm rounded-xl transition-all flex items-center justify-center gap-2 cursor-pointer active:scale-95 shadow-md"
                    >
                      <Pencil className="w-4 h-4 text-white" />
                      <span>Edit Nilai</span>
                    </button>
                  )}

                  <button
                    type="button"
                    onClick={handleSaveScores}
                    disabled={isSaving || isScoresDisabled}
                    className="w-full sm:w-auto px-8 py-3.5 bg-[#8dc63f] hover:bg-[#7bc025] disabled:bg-slate-300 disabled:cursor-not-allowed text-white font-black text-sm rounded-xl transition-all shadow-md hover:shadow-lg active:scale-95 flex items-center justify-center gap-2 cursor-pointer"
                  >
                    <CheckCircle2 className="w-5 h-5" />
                    <span>{isSaving ? 'Menyimpan...' : isScoresDisabled ? 'Nilai Tersimpan' : 'Simpan Nilai Siswa'}</span>
                  </button>
                </div>
              )}
            </div>
          </motion.div>
        ) : (
          /* PREVIEW NILAI SUB-MENU */
          <motion.div
            key="subtab-preview"
            initial={{ opacity: 0, y: 10 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -10 }}
            transition={{ duration: 0.2 }}
            className="space-y-6"
          >
            {/* Filter & Action Controls Bar */}
            <div className="bg-white rounded-3xl p-6 sm:p-8 border-2 border-slate-400 shadow-sm space-y-6">
              <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-4 pb-4 border-b-2 border-slate-200">
                <div>
                  <h3 className="text-base sm:text-lg font-black text-slate-800 flex items-center gap-2">
                    Preview & Rekapitulasi Nilai Siswa
                  </h3>
                  <p className="text-xs text-slate-500 font-medium mt-0.5">
                    Lihat matriks nilai lengkap per kelas, filter berdasarkan tugas, atau ekspor ke file Excel.
                  </p>
                </div>

                <div className="flex flex-wrap items-center gap-2 w-full md:w-auto">
                  {/* Export Excel Button */}
                  <button
                    type="button"
                    onClick={handleExportExcelPreview}
                    className="flex-1 md:flex-none px-5 py-2.5 bg-[#8dc63f] hover:bg-[#7bc025] text-white font-black text-xs sm:text-sm rounded-xl transition-all flex items-center justify-center gap-2 cursor-pointer shadow-md shadow-[#8dc63f]/30 active:scale-95"
                    title="Eksport Rekap Nilai ke File Excel (.xlsx) dengan Format Kertas F4 & Tanda Tangan"
                  >
                    <Download className="w-4 h-4" />
                    <span>Eksport Excel</span>
                  </button>

                  {/* Print / PDF Button */}
                  <button
                    type="button"
                    onClick={() => handleOpenPrintModal('preview')}
                    className="flex-1 md:flex-none px-5 py-2.5 bg-rose-600 hover:bg-rose-700 text-white font-black text-xs sm:text-sm rounded-xl transition-all flex items-center justify-center gap-2 cursor-pointer shadow-md shadow-rose-600/30 active:scale-95"
                    title="Cetak Rekap Nilai Siswa ke Kertas F4 (PDF / Print)"
                  >
                    <Printer className="w-4 h-4" />
                    <span>Cetak / PDF (F4)</span>
                  </button>
                </div>
              </div>

              {/* Filters Grid */}
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                {/* 1. Pilih Kelas Dropdown */}
                <div className="space-y-1.5">
                  <label className="text-xs font-extrabold text-slate-700 flex items-center gap-1.5">
                    <Building2 className="w-3.5 h-3.5 text-[#8dc63f]" />
                    Pilih Kelas
                  </label>
                  <select
                    value={selectedClassPreview}
                    onChange={(e) => {
                      setSelectedClassPreview(e.target.value);
                      setAssignmentFilterPreview('all');
                    }}
                    className="w-full p-3 bg-slate-50 border-2 border-slate-200 rounded-2xl font-bold text-slate-800 text-xs sm:text-sm focus:ring-2 focus:ring-[#8dc63f]/20 focus:border-[#8dc63f] transition-all outline-none"
                  >
                    {classList.length === 0 ? (
                      <option value="">-- Belum ada kelas --</option>
                    ) : (
                      classList
                        .slice()
                        .sort(compareClasses)
                        .map((c) => (
                          <option key={c} value={c}>
                            Kelas {c}
                          </option>
                        ))
                    )}
                  </select>
                </div>

                {/* 2. Pilih Tugas (Filter) Dropdown */}
                <div className="space-y-1.5">
                  <label className="text-xs font-extrabold text-slate-700 flex items-center gap-1.5">
                    <Filter className="w-3.5 h-3.5 text-[#8dc63f]" />
                    Pilih Tugas (Filter)
                  </label>
                  <select
                    value={assignmentFilterPreview}
                    onChange={(e) => setAssignmentFilterPreview(e.target.value)}
                    className="w-full p-3 bg-slate-50 border-2 border-slate-200 rounded-2xl font-bold text-slate-800 text-xs sm:text-sm focus:ring-2 focus:ring-[#8dc63f]/20 focus:border-[#8dc63f] transition-all outline-none"
                  >
                    <option value="all">Semua Tugas ({classAssignmentsPreview.length})</option>
                    {classAssignmentsPreview.map((a) => (
                      <option key={a.id} value={a.id}>
                        {a.title} ({a.date})
                      </option>
                    ))}
                  </select>
                </div>

                {/* Search Student Filter */}
                <div className="space-y-1.5">
                  <label className="text-xs font-extrabold text-slate-700 flex items-center gap-1.5">
                    <Search className="w-3.5 h-3.5 text-[#8dc63f]" />
                    Cari Nama Siswa
                  </label>
                  <input
                    type="text"
                    placeholder="Cari nama atau Nis..."
                    value={searchStudentQuery}
                    onChange={(e) => setSearchStudentQuery(e.target.value)}
                    className="w-full p-3 bg-slate-50 border-2 border-slate-200 rounded-2xl font-bold text-slate-800 text-xs sm:text-sm focus:ring-2 focus:ring-[#8dc63f]/20 focus:border-[#8dc63f] transition-all outline-none placeholder:text-slate-400"
                  />
                </div>
              </div>
            </div>

            {/* Table Preview Nilai */}
            <div className="bg-white rounded-3xl p-6 sm:p-8 border-2 border-slate-400 shadow-sm space-y-4">
              {studentsInPreviewClass.length === 0 ? (
                <div className="text-center py-12 text-slate-500 italic">
                  Tidak ada data siswa ditemukan di kelas {selectedClassPreview}.
                </div>
              ) : (
                <div className="space-y-4">
                  {activeAssignmentsColumns.length === 0 && (
                    <div className="flex flex-col sm:flex-row items-center justify-between gap-3 p-4 bg-slate-50 border border-dashed border-slate-200 rounded-2xl">
                      <p className="text-xs font-bold text-slate-600">
                        Belum ada tugas dibuat untuk kelas {selectedClassPreview}. Kolom nilai kehadiran siswa tetap ditampilkan di bawah.
                      </p>
                      <button
                        type="button"
                        onClick={() => setActiveSubTab('input')}
                        className="px-4 py-2 bg-[#8dc63f] hover:bg-[#7bc025] text-white font-black text-xs rounded-xl shadow-xs transition-all cursor-pointer shrink-0"
                      >
                        + Buat Tugas Sekarang
                      </button>
                    </div>
                  )}

                  <div className="overflow-x-auto max-h-[600px] border border-slate-300 rounded-2xl scrollbar-thin shadow-xs">
                    <table className="w-full text-left text-xs sm:text-sm border-collapse border border-slate-300">
                      <thead className="sticky top-0 bg-slate-100 z-10">
                        <tr>
                          <th className="p-3 sm:p-3.5 font-black text-slate-800 w-12 text-center border border-slate-300 bg-slate-100">
                            No
                          </th>
                          <th className="p-3 sm:p-3.5 font-black text-slate-800 min-w-[200px] border border-slate-300 bg-slate-100">
                            Nama Lengkap Siswa
                          </th>
                          <th className="p-3 sm:p-3.5 font-black text-slate-800 w-28 text-center border border-slate-300 bg-slate-100">
                            Nis
                          </th>
                          <th className="p-3 sm:p-3.5 font-black text-slate-800 text-center w-36 border border-slate-300 bg-slate-100">
                            Nilai Kehadiran
                          </th>

                          {/* Dynamic Columns for each Assignment */}
                          {activeAssignmentsColumns.map((assign) => (
                            <th
                              key={assign.id}
                              className="p-3 sm:p-3.5 font-black text-slate-800 text-center min-w-[130px] border border-slate-300 bg-slate-100"
                            >
                              <div className="flex flex-col items-center">
                                <span className="text-slate-900 font-extrabold">{assign.title}</span>
                                <span className="text-[10px] text-slate-500 font-medium mt-0.5">{assign.date}</span>
                              </div>
                            </th>
                          ))}

                          {activeAssignmentsColumns.length > 0 && (
                            <th className="p-3 sm:p-3.5 font-black text-[#5a8c20] text-center w-28 bg-[#8dc63f]/20 border border-slate-300">
                              Nilai Rapor
                            </th>
                          )}
                        </tr>
                      </thead>
                      <tbody className="bg-white">
                        {studentsInPreviewClass.map((student, idx) => {
                          let totalScore = 0;
                          let gradedCount = 0;

                          return (
                            <tr key={student.id} className="hover:bg-slate-50/80 transition-colors">
                              <td className="p-3 sm:p-3.5 text-center font-bold text-slate-600 border border-slate-300">{idx + 1}</td>
                              <td className="p-3 sm:p-3.5 font-extrabold text-slate-800 border border-slate-300">{student.name}</td>
                              <td className="p-3 sm:p-3.5 font-semibold text-slate-600 text-center border border-slate-300">{student.nisn || '-'}</td>
                              <td className="p-3 sm:p-3.5 text-center font-bold text-slate-800 text-xs sm:text-sm border border-slate-300">
                                {getStudentAttendanceScore(student.id, selectedClassPreview)}
                              </td>

                              {/* Scores for each assignment */}
                              {activeAssignmentsColumns.map((assign) => {
                                const scoreVal = assign.scores[student.id];
                                const hasScore = scoreVal !== undefined && scoreVal !== null && scoreVal !== '';
                                const numScore = Number(scoreVal);

                                if (hasScore && !isNaN(numScore)) {
                                  totalScore += numScore;
                                  gradedCount++;
                                }

                                return (
                                  <td key={assign.id} className="p-3 sm:p-3.5 text-center font-extrabold text-slate-800 text-xs sm:text-sm border border-slate-300">
                                    {hasScore ? (
                                      scoreVal
                                    ) : (
                                      <span className="text-slate-400 font-medium">-</span>
                                    )}
                                  </td>
                                );
                              })}

                              {/* Nilai Rapor Column */}
                              {activeAssignmentsColumns.length > 0 && (
                                <td className="p-3 sm:p-3.5 text-center font-black text-[#5a8c20] bg-[#8dc63f]/10 text-xs sm:text-sm border border-slate-300">
                                  {gradedCount > 0 ? (
                                    Math.round(totalScore / gradedCount)
                                  ) : (
                                    <span className="text-slate-400 font-medium">-</span>
                                  )}
                                </td>
                              )}
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>

                  {/* Symmetrical Signatures Preview & Settings */}
                  <div className="mt-8 pt-6 border-t border-slate-200 space-y-6">
                    {/* Header bar of signature section */}
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-slate-50 p-4 rounded-2xl border border-slate-200">
                      <div className="flex items-center gap-2.5">
                        <div className="w-8 h-8 rounded-xl bg-[#8dc63f]/20 text-[#7bc025] flex items-center justify-center font-bold">
                          <UserCheck className="w-4 h-4" />
                        </div>
                        <div>
                          <h4 className="text-xs font-black text-slate-800 uppercase tracking-wider">
                            Format Tanda Tangan Dokumen (Presisi &amp; Simetris)
                          </h4>
                          <p className="text-[11px] text-slate-500 font-medium">
                            Format teks tanda tangan disesuaikan dengan laporan kehadiran siswa resmi.
                          </p>
                        </div>
                      </div>

                      {/* Paper format status indicator */}
                      <div className="inline-flex items-center gap-2 px-3 py-1.5 bg-white rounded-xl border border-slate-200 text-slate-700 text-xs font-bold shadow-xs">
                        <span className="w-2 h-2 rounded-full bg-emerald-500"></span>
                        <span>Kertas F4 ({activeAssignmentsColumns.length > 3 ? 'Otomatis Landscape' : 'Portrait'})</span>
                      </div>
                    </div>

                    {/* Signers selection options */}
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 bg-slate-50/50 p-4 rounded-2xl border border-slate-200/80">
                      <div>
                        <label className="block text-[11px] font-bold text-slate-700 mb-1">
                          1. Penandatangan Kiri (Mengetahui 1)
                        </label>
                        <select
                          value={leftSignerRole}
                          onChange={(e) => setLeftSignerRole(e.target.value as SignerRoleType)}
                          className="w-full p-2.5 bg-white border border-slate-300 rounded-xl text-xs font-bold text-slate-800 focus:ring-2 focus:ring-[#8dc63f]/20 focus:border-[#8dc63f] outline-none"
                        >
                          <option value="kepala_sekolah">🏫 Kepala Sekolah ({effectiveProfile.namaKepalaSekolah || 'Belum diisi'})</option>
                          <option value="kurikulum">📚 Wakasek Kurikulum ({effectiveProfile.namaKurikulum || 'Belum diisi'})</option>
                          <option value="kesiswaan">👥 Wakasek Kesiswaan ({effectiveProfile.namaKesiswaan || 'Belum diisi'})</option>
                          <option value="humas">🤝 Wakasek Humas ({effectiveProfile.namaHumas || 'Belum diisi'})</option>
                          <option value="guru_wali">👨‍🏫 Guru Wali ({effectiveProfile.namaGuruWali || 'Belum diisi'})</option>
                          <option value="guru_bk">🧭 Guru BK ({effectiveProfile.namaBK || 'Belum diisi'})</option>
                          <option value="none">🚫 Tanpa Tanda Tangan Kiri</option>
                        </select>
                      </div>

                      <div>
                        <label className="block text-[11px] font-bold text-slate-700 mb-1 flex items-center justify-between">
                          <span>2. Penandatangan Tengah (Mengetahui 2)</span>
                          <span className="text-[10px] text-slate-400 font-normal">Bila perlu 3 tanda tangan</span>
                        </label>
                        <select
                          value={midSignerRole}
                          onChange={(e) => setMidSignerRole(e.target.value as SignerRoleType)}
                          className="w-full p-2.5 bg-white border border-slate-300 rounded-xl text-xs font-bold text-slate-800 focus:ring-2 focus:ring-[#8dc63f]/20 focus:border-[#8dc63f] outline-none"
                        >
                          <option value="none">🚫 Tanpa Tanda Tangan Tengah (Default 2 TTD)</option>
                          <option value="kurikulum">📚 Wakasek Kurikulum ({effectiveProfile.namaKurikulum || 'Belum diisi'})</option>
                          <option value="kesiswaan">👥 Wakasek Kesiswaan ({effectiveProfile.namaKesiswaan || 'Belum diisi'})</option>
                          <option value="humas">🤝 Wakasek Humas ({effectiveProfile.namaHumas || 'Belum diisi'})</option>
                          <option value="guru_bk">🧭 Guru BK ({effectiveProfile.namaBK || 'Belum diisi'})</option>
                          <option value="kepala_sekolah">🏫 Kepala Sekolah ({effectiveProfile.namaKepalaSekolah || 'Belum diisi'})</option>
                        </select>
                      </div>
                    </div>

                    {/* Symmetrical Visual Signature Box */}
                    {(() => {
                      const leftSig = getSignerDetails(leftSignerRole);
                      const midSig = getSignerDetails(midSignerRole);
                      const is3Sig = leftSig.enabled && midSig.enabled;
                      const { teacherTitle, teacherName, teacherNIP } = getTeacherSignerInfo();
                      const dateNowStr = format(new Date(), 'dd MMMM yyyy', { locale: id });

                      return (
                        <div className={`p-6 sm:p-8 bg-white border-2 border-dashed border-slate-200 rounded-3xl grid gap-8 text-center text-xs ${
                          is3Sig ? 'grid-cols-1 md:grid-cols-3' : 'grid-cols-1 sm:grid-cols-2'
                        }`}>
                          {/* Kiri */}
                          {leftSig.enabled ? (
                            <div className="flex flex-col items-center justify-between min-h-[160px]">
                              <div>
                                <p className="text-slate-600 font-medium">Mengetahui,</p>
                                <p className="font-extrabold text-slate-800">{leftSig.title}</p>
                              </div>
                              <div className="my-6 text-slate-300 italic text-[11px]">
                                (Ruang Tanda Tangan)
                              </div>
                              <div>
                                <p className="font-black text-slate-900 underline text-sm tracking-wide">{leftSig.name}</p>
                                {leftSig.nip ? (
                                  <p className="text-slate-600 font-semibold text-[11px]">{leftSig.nip}</p>
                                ) : (
                                  <p className="text-slate-400 text-[11px]">-</p>
                                )}
                              </div>
                            </div>
                          ) : (
                            <div className="hidden sm:block"></div>
                          )}

                          {/* Tengah (jika 3 TTD) */}
                          {is3Sig && (
                            <div className="flex flex-col items-center justify-between min-h-[160px]">
                              <div>
                                <p className="text-slate-600 font-medium">Mengetahui,</p>
                                <p className="font-extrabold text-slate-800">{midSig.title}</p>
                              </div>
                              <div className="my-6 text-slate-300 italic text-[11px]">
                                (Ruang Tanda Tangan)
                              </div>
                              <div>
                                <p className="font-black text-slate-900 underline text-sm tracking-wide">{midSig.name}</p>
                                {midSig.nip ? (
                                  <p className="text-slate-600 font-semibold text-[11px]">{midSig.nip}</p>
                                ) : (
                                  <p className="text-slate-400 text-[11px]">-</p>
                                )}
                              </div>
                            </div>
                          )}

                          {/* Kanan: Guru Mata Pelajaran / Wali Kelas */}
                          <div className="flex flex-col items-center justify-between min-h-[160px]">
                            <div>
                              <p className="text-slate-600 font-medium">Cililin, {dateNowStr}</p>
                              <p className="font-extrabold text-slate-800">{teacherTitle}</p>
                            </div>
                            <div className="my-6 text-slate-300 italic text-[11px]">
                              (Ruang Tanda Tangan)
                            </div>
                            <div>
                              <p className="font-black text-slate-900 underline text-sm tracking-wide">{teacherName}</p>
                              {teacherNIP ? (
                                <p className="text-slate-600 font-semibold text-[11px]">{teacherNIP}</p>
                              ) : (
                                <p className="text-slate-400 text-[11px]">-</p>
                              )}
                            </div>
                          </div>
                        </div>
                      );
                    })()}
                  </div>
                </div>
              )}
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Edit Assignment Modal */}
      <AnimatePresence>
        {isEditingAssignmentModal && (
          <div className="fixed inset-0 z-[200] flex items-center justify-center p-4">
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setIsEditingAssignmentModal(false)}
              className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
            />
            <motion.div
              initial={{ scale: 0.95, opacity: 0, y: 10 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 10 }}
              className="relative bg-white rounded-3xl p-6 sm:p-8 max-w-lg w-full shadow-2xl border-2 border-slate-400 z-10 space-y-6"
            >
              <div className="flex items-center justify-between pb-4 border-b border-slate-100">
                <div className="flex items-center gap-2.5">
                  <div className="w-9 h-9 rounded-xl bg-amber-100 text-amber-700 flex items-center justify-center font-bold">
                    <Edit3 className="w-5 h-5" />
                  </div>
                  <h3 className="text-lg font-black text-slate-800">Edit Tugas</h3>
                </div>
                <button
                  type="button"
                  onClick={() => setIsEditingAssignmentModal(false)}
                  className="p-2 text-slate-400 hover:text-slate-600 rounded-full hover:bg-slate-100 transition-colors"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              <div className="space-y-4">
                <div className="space-y-1.5">
                  <label className="text-xs font-extrabold text-slate-700">Tanggal Tugas</label>
                  <input
                    type="date"
                    value={editAssignmentDate}
                    onChange={(e) => setEditAssignmentDate(e.target.value)}
                    className="w-full p-3 bg-slate-50 border-2 border-slate-200 rounded-2xl font-bold text-slate-800 text-xs sm:text-sm focus:border-[#8dc63f] outline-none"
                  />
                </div>

                <div className="space-y-1.5">
                  <label className="text-xs font-extrabold text-slate-700">Judul / Nama Tugas</label>
                  <input
                    type="text"
                    value={editAssignmentTitle}
                    onChange={(e) => setEditAssignmentTitle(e.target.value)}
                    placeholder="Masukkan judul tugas..."
                    className="w-full p-3 bg-slate-50 border-2 border-slate-200 rounded-2xl font-bold text-slate-800 text-xs sm:text-sm focus:border-[#8dc63f] outline-none"
                  />
                </div>
              </div>

              <div className="flex items-center justify-end gap-3 pt-4 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setIsEditingAssignmentModal(false)}
                  className="px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs sm:text-sm rounded-xl transition-all cursor-pointer"
                >
                  Batal
                </button>
                <button
                  type="button"
                  onClick={handleSaveEditAssignment}
                  disabled={isSaving || !editAssignmentTitle.trim()}
                  className="px-6 py-2.5 bg-[#8dc63f] hover:bg-[#7bc025] disabled:bg-slate-300 text-white font-black text-xs sm:text-sm rounded-xl transition-all shadow-md active:scale-95 flex items-center gap-2 cursor-pointer"
                >
                  <Save className="w-4 h-4" />
                  <span>{isSaving ? 'Menyimpan...' : 'Simpan Perubahan'}</span>
                </button>
              </div>
            </motion.div>
          </div>
        )}

        {/* Delete Assignment Confirmation Modal */}
        {deleteConfirmAssignment && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => !isDeleting && setDeleteConfirmAssignment(null)}
              className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
            />
            <motion.div
              initial={{ scale: 0.95, opacity: 0, y: 10 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 10 }}
              className="relative bg-white rounded-3xl p-6 sm:p-8 max-w-md w-full shadow-2xl border-2 border-slate-400 z-10 space-y-5"
            >
              <div className="flex items-center gap-3">
                <div className="w-12 h-12 rounded-2xl bg-rose-100 text-rose-600 flex items-center justify-center font-bold shrink-0">
                  <Trash2 className="w-6 h-6" />
                </div>
                <div>
                  <h3 className="text-base sm:text-lg font-black text-slate-800">
                    Konfirmasi Hapus Tugas
                  </h3>
                  <p className="text-xs text-slate-500 font-medium">
                    Tindakan ini tidak dapat dibatalkan.
                  </p>
                </div>
              </div>

              <div className="p-4 bg-rose-50/80 border border-rose-200 rounded-2xl space-y-1">
                <p className="text-xs font-bold text-rose-900">
                  Apakah Anda yakin ingin menghapus tugas ini?
                </p>
                <p className="text-xs font-extrabold text-slate-800">
                  &ldquo;{deleteConfirmAssignment.title}&rdquo; (Kelas {deleteConfirmAssignment.className})
                </p>
                <p className="text-[11px] text-rose-700 font-medium pt-1">
                  Seluruh data nilai siswa yang tersimpan untuk tugas ini akan dihapus secara permanen.
                </p>
              </div>

              <div className="flex items-center justify-end gap-3 pt-2">
                <button
                  type="button"
                  disabled={isDeleting}
                  onClick={() => setDeleteConfirmAssignment(null)}
                  className="px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-extrabold text-xs sm:text-sm rounded-xl transition-all cursor-pointer disabled:opacity-50"
                >
                  Batal
                </button>
                <button
                  type="button"
                  disabled={isDeleting}
                  onClick={handleConfirmDeleteAssignment}
                  className="px-6 py-2.5 bg-rose-600 hover:bg-rose-700 disabled:bg-rose-300 text-white font-black text-xs sm:text-sm rounded-xl transition-all shadow-md active:scale-95 flex items-center gap-2 cursor-pointer"
                >
                  <Trash2 className="w-4 h-4" />
                  <span>{isDeleting ? 'Menghapus...' : 'Ya, Hapus Tugas'}</span>
                </button>
              </div>
            </motion.div>
          </div>
        )}

        {/* Print & PDF Export Modal (Ukuran Kertas F4) */}
        {isPrintModalOpen && (
          <div className="fixed inset-0 z-[220] flex items-center justify-center p-4">
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => !isExportingPDF && setIsPrintModalOpen(false)}
              className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
            />
            <motion.div
              initial={{ scale: 0.95, opacity: 0, y: 10 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 10 }}
              className="relative bg-white rounded-3xl p-6 sm:p-8 max-w-xl w-full shadow-2xl border-2 border-slate-400 z-10 space-y-6"
            >
              <div className="flex items-center justify-between pb-4 border-b border-slate-100">
                <div className="flex items-center gap-2.5">
                  <div className="w-10 h-10 rounded-2xl bg-rose-100 text-rose-600 flex items-center justify-center font-bold">
                    <Printer className="w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-base sm:text-lg font-black text-slate-800">Cetak Rekap Nilai (Kertas F4)</h3>
                    <p className="text-xs text-slate-500 font-medium">Standar tata naskah &amp; kop surat resmi sekolah</p>
                  </div>
                </div>
                <button
                  type="button"
                  onClick={() => setIsPrintModalOpen(false)}
                  disabled={isExportingPDF}
                  className="p-2 text-slate-400 hover:text-slate-600 rounded-full hover:bg-slate-100 transition-colors"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              {/* Status info box */}
              {(() => {
                const targetClass = printTargetSubTab === 'preview' ? selectedClassPreview : selectedClassInput;
                const studentCount = printTargetSubTab === 'preview' ? studentsInPreviewClass.length : studentsInInputClass.length;
                const taskCount = printTargetSubTab === 'preview' ? activeAssignmentsColumns.length : currentClassAssignmentsInput.length;
                const isAutoLand = taskCount > 3;

                return (
                  <div className="space-y-4">
                    <div className="p-4 bg-slate-50 border border-slate-200/80 rounded-2xl grid grid-cols-2 sm:grid-cols-4 gap-3 text-center">
                      <div className="p-2.5 bg-white rounded-xl border border-slate-200 shadow-xs">
                        <span className="block text-[10px] font-bold text-slate-400 uppercase">Kelas</span>
                        <span className="text-xs font-black text-slate-800">{targetClass || '-'}</span>
                      </div>
                      <div className="p-2.5 bg-white rounded-xl border border-slate-200 shadow-xs">
                        <span className="block text-[10px] font-bold text-slate-400 uppercase">Total Siswa</span>
                        <span className="text-xs font-black text-slate-800">{studentCount} Siswa</span>
                      </div>
                      <div className="p-2.5 bg-white rounded-xl border border-slate-200 shadow-xs">
                        <span className="block text-[10px] font-bold text-slate-400 uppercase">Kolom Tugas</span>
                        <span className="text-xs font-black text-slate-800">{taskCount} Tugas</span>
                      </div>
                      <div className="p-2.5 bg-white rounded-xl border border-slate-200 shadow-xs">
                        <span className="block text-[10px] font-bold text-slate-400 uppercase">Kertas Cetak</span>
                        <span className="text-xs font-black text-rose-700">F4 (Folio)</span>
                      </div>
                    </div>

                    {/* Auto-orientation alert notice */}
                    <div className={`p-4 rounded-2xl border text-xs leading-relaxed flex items-start gap-3 ${
                      isAutoLand ? 'bg-amber-50/80 border-amber-200 text-amber-900' : 'bg-emerald-50/80 border-emerald-200 text-emerald-900'
                    }`}>
                      <CheckCircle2 className={`w-5 h-5 shrink-0 mt-0.5 ${isAutoLand ? 'text-amber-600' : 'text-emerald-600'}`} />
                      <div>
                        <p className="font-extrabold text-[13px]">
                          {isAutoLand ? 'Posisi Otomatis: Landscape (Mendatar)' : 'Posisi Otomatis: Portrait (Tegak)'}
                        </p>
                        <p className="mt-1 font-medium opacity-90">
                          {isAutoLand
                            ? 'Karena jumlah kolom tugas lebih dari 3, posisi cetak diatur secara otomatis ke Landscape pada kertas F4 agar seluruh kolom nilai presisi, simetris, dan tidak terpotong.'
                            : 'Kolom nilai mencukupi pada posisi Portrait (tegak) pada kertas F4 dengan proporsi presisi.'}
                        </p>
                      </div>
                    </div>

                    {/* Orientation manual override selector */}
                    <div className="space-y-1.5">
                      <label className="text-xs font-bold text-slate-700">Pilihan Orientasi Kertas F4</label>
                      <div className="grid grid-cols-3 gap-2">
                        <button
                          type="button"
                          onClick={() => setPrintOrientationOption('auto')}
                          className={`p-2.5 rounded-xl border text-xs font-extrabold transition-all cursor-pointer ${
                            printOrientationOption === 'auto'
                              ? 'bg-rose-50 border-rose-500 text-rose-700 shadow-xs'
                              : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                          }`}
                        >
                          Otomatis ({isAutoLand ? 'Landscape' : 'Portrait'})
                        </button>
                        <button
                          type="button"
                          onClick={() => setPrintOrientationOption('portrait')}
                          className={`p-2.5 rounded-xl border text-xs font-extrabold transition-all cursor-pointer ${
                            printOrientationOption === 'portrait'
                              ? 'bg-rose-50 border-rose-500 text-rose-700 shadow-xs'
                              : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                          }`}
                        >
                          Paksa Portrait
                        </button>
                        <button
                          type="button"
                          onClick={() => setPrintOrientationOption('landscape')}
                          className={`p-2.5 rounded-xl border text-xs font-extrabold transition-all cursor-pointer ${
                            printOrientationOption === 'landscape'
                              ? 'bg-rose-50 border-rose-500 text-rose-700 shadow-xs'
                              : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                          }`}
                        >
                          Paksa Landscape
                        </button>
                      </div>
                    </div>

                    {/* Penandatangan Selectors in Modal */}
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-2">
                      <div className="space-y-1">
                        <label className="text-[11px] font-bold text-slate-600">Tanda Tangan Kiri</label>
                        <select
                          value={leftSignerRole}
                          onChange={(e) => setLeftSignerRole(e.target.value as SignerRoleType)}
                          className="w-full p-2 bg-slate-50 border border-slate-300 rounded-xl text-xs font-bold text-slate-800"
                        >
                          <option value="kepala_sekolah">🏫 Kepala Sekolah</option>
                          <option value="kurikulum">📚 Wakasek Kurikulum</option>
                          <option value="kesiswaan">👥 Wakasek Kesiswaan</option>
                          <option value="humas">🤝 Wakasek Humas</option>
                          <option value="guru_wali">👨‍🏫 Guru Wali</option>
                          <option value="guru_bk">🧭 Guru BK</option>
                          <option value="none">🚫 Tanpa TTD Kiri</option>
                        </select>
                      </div>
                      <div className="space-y-1">
                        <label className="text-[11px] font-bold text-slate-600">Tanda Tangan Tengah</label>
                        <select
                          value={midSignerRole}
                          onChange={(e) => setMidSignerRole(e.target.value as SignerRoleType)}
                          className="w-full p-2 bg-slate-50 border border-slate-300 rounded-xl text-xs font-bold text-slate-800"
                        >
                          <option value="none">🚫 Tanpa TTD Tengah (2 TTD)</option>
                          <option value="kurikulum">📚 Wakasek Kurikulum</option>
                          <option value="kesiswaan">👥 Wakasek Kesiswaan</option>
                          <option value="humas">🤝 Wakasek Humas</option>
                          <option value="guru_bk">🧭 Guru BK</option>
                        </select>
                      </div>
                    </div>
                  </div>
                );
              })()}

              <div className="flex flex-col-reverse sm:flex-row items-center justify-end gap-3 pt-3 border-t border-slate-100">
                <button
                  type="button"
                  disabled={isExportingPDF}
                  onClick={() => setIsPrintModalOpen(false)}
                  className="w-full sm:w-auto px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-extrabold text-xs sm:text-sm rounded-xl transition-all cursor-pointer disabled:opacity-50"
                >
                  Batal
                </button>
                <button
                  type="button"
                  disabled={isExportingPDF}
                  onClick={() => handleExecutePDFAction('download')}
                  className="w-full sm:w-auto px-5 py-2.5 bg-rose-600 hover:bg-rose-700 disabled:bg-rose-300 text-white font-black text-xs sm:text-sm rounded-xl transition-all shadow-md active:scale-95 flex items-center justify-center gap-2 cursor-pointer"
                >
                  <FileDown className="w-4 h-4" />
                  <span>{isExportingPDF ? 'Memproses...' : 'Unduh PDF (Kertas F4)'}</span>
                </button>
                <button
                  type="button"
                  disabled={isExportingPDF}
                  onClick={() => handleExecutePDFAction('print')}
                  className="w-full sm:w-auto px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 disabled:bg-emerald-300 text-white font-black text-xs sm:text-sm rounded-xl transition-all shadow-md active:scale-95 flex items-center justify-center gap-2 cursor-pointer"
                >
                  <Printer className="w-4 h-4" />
                  <span>{isExportingPDF ? 'Mencetak...' : 'Cetak Langsung'}</span>
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

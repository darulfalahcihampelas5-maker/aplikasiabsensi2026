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
  X
} from 'lucide-react';
import { format } from 'date-fns';
import { collection, doc, setDoc, deleteDoc, onSnapshot, query, where, Firestore } from 'firebase/firestore';
import { Auth } from 'firebase/auth';
import * as XLSX from 'xlsx';
import { compareClasses, compareStudentsByClass } from './classSortUtils';
import { handleFirestoreError, OperationType } from './firebase';

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

interface StudentGradesViewProps {
  classList: string[];
  students: Student[];
  profileData?: { role?: string; [key: string]: unknown };
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
  activeDb,
  activeAuth,
  trackOp,
  showToast,
  initialSubTab = 'input'
}: StudentGradesViewProps) {
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

  // Handler: Export Preview Nilai to Excel (.xlsx)
  const handleExportExcelPreview = () => {
    if (studentsInPreviewClass.length === 0) {
      showToast?.('Tidak ada data siswa untuk diekspor.', 'error');
      return;
    }

    const exportRows = studentsInPreviewClass.map((student, idx) => {
      /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
      const row: Record<string, any> = {
        'No': idx + 1,
        'Nama Lengkap Siswa': student.name,
        'Nis': student.nisn || '-'
      };

      let totalScore = 0;
      let gradedCount = 0;

      activeAssignmentsColumns.forEach((assign) => {
        const val = assign.scores[student.id];
        if (val !== undefined && val !== null && val !== '') {
          const num = Number(val);
          row[assign.title] = !isNaN(num) ? num : val;
          if (!isNaN(num)) {
            totalScore += num;
            gradedCount++;
          }
        } else {
          row[assign.title] = '-';
        }
      });

      if (activeAssignmentsColumns.length > 1) {
        row['Rata-Rata Nilai'] = gradedCount > 0 ? Math.round(totalScore / gradedCount) : '-';
      }

      return row;
    });

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, `Nilai ${selectedClassPreview}`);

    const dateStr = format(new Date(), 'yyyyMMdd');
    const filename = `Rekap_Nilai_Siswa_Kelas_${selectedClassPreview.replace(/\s+/g, '_')}_${dateStr}.xlsx`;
    XLSX.writeFile(workbook, filename);

    showToast?.(`Rekap nilai kelas ${selectedClassPreview} berhasil diekspor!`, 'success');
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
            <div className="bg-white rounded-3xl p-6 sm:p-8 border border-slate-100 shadow-sm space-y-6">
              <div className="flex items-center gap-3 pb-4 border-b border-slate-100">
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
            <div className="bg-white rounded-3xl p-6 sm:p-8 border border-slate-100 shadow-sm space-y-6">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-slate-100">
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
                      className="px-4 py-2.5 bg-[#8dc63f] hover:bg-[#7bc025] text-white font-extrabold text-xs sm:text-sm rounded-xl transition-all flex items-center gap-2 cursor-pointer shadow-md active:scale-95"
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

                  <div className="overflow-x-auto max-h-[550px] border border-slate-100 rounded-2xl scrollbar-thin">
                    <table className="w-full text-left text-xs sm:text-sm border-collapse">
                      <thead className="sticky top-0 bg-slate-100/90 backdrop-blur-md z-10 shadow-[0_1px_0_0_rgba(226,232,240,1)]">
                        <tr>
                          <th className="p-3.5 sm:p-4 font-black text-slate-700 w-16 text-center">No</th>
                          <th className="p-3.5 sm:p-4 font-black text-slate-700">Nama Lengkap Siswa</th>
                          <th className="p-3.5 sm:p-4 font-black text-slate-700 w-32">Nis</th>
                          <th className="p-3.5 sm:p-4 font-black text-slate-700 min-w-[160px] text-center">
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
                      <tbody className="divide-y divide-slate-100">
                        {studentsInInputClass.map((student, idx) => {
                          const currentVal = scoresInput[student.id] ?? '';
                          return (
                            <tr key={student.id} className="hover:bg-[#8dc63f]/10 transition-colors">
                              <td className="p-3.5 sm:p-4 text-center font-bold text-slate-500">{idx + 1}</td>
                              <td className="p-3.5 sm:p-4 font-extrabold text-slate-800">{student.name}</td>
                              <td className="p-3.5 sm:p-4 font-semibold text-slate-600">{student.nisn || '-'}</td>
                              <td className="p-3.5 sm:p-4 text-center">
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
            <div className="bg-white rounded-3xl p-6 sm:p-8 border border-slate-100 shadow-sm space-y-6">
              <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-4 pb-4 border-b border-slate-100">
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
                    className="flex-1 md:flex-none px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs sm:text-sm rounded-xl transition-all flex items-center justify-center gap-2 cursor-pointer shadow-md active:scale-95"
                  >
                    <Download className="w-4 h-4" />
                    <span>Eksport Excel</span>
                  </button>
                </div>
              </div>

              {/* Filters Grid */}
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                {/* 1. Pilih Kelas Dropdown */}
                <div className="space-y-1.5">
                  <label className="text-xs font-extrabold text-slate-700 flex items-center gap-1.5">
                    <Building2 className="w-3.5 h-3.5 text-emerald-600" />
                    Pilih Kelas
                  </label>
                  <select
                    value={selectedClassPreview}
                    onChange={(e) => {
                      setSelectedClassPreview(e.target.value);
                      setAssignmentFilterPreview('all');
                    }}
                    className="w-full p-3 bg-slate-50 border-2 border-slate-200 rounded-2xl font-bold text-slate-800 text-xs sm:text-sm focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600 transition-all outline-none"
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
                    <Filter className="w-3.5 h-3.5 text-emerald-600" />
                    Pilih Tugas (Filter)
                  </label>
                  <select
                    value={assignmentFilterPreview}
                    onChange={(e) => setAssignmentFilterPreview(e.target.value)}
                    className="w-full p-3 bg-slate-50 border-2 border-slate-200 rounded-2xl font-bold text-slate-800 text-xs sm:text-sm focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600 transition-all outline-none"
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
                    <Search className="w-3.5 h-3.5 text-emerald-600" />
                    Cari Nama Siswa
                  </label>
                  <input
                    type="text"
                    placeholder="Cari nama atau Nis..."
                    value={searchStudentQuery}
                    onChange={(e) => setSearchStudentQuery(e.target.value)}
                    className="w-full p-3 bg-slate-50 border-2 border-slate-200 rounded-2xl font-bold text-slate-800 text-xs sm:text-sm focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600 transition-all outline-none placeholder:text-slate-400"
                  />
                </div>
              </div>
            </div>

            {/* Table Preview Nilai */}
            <div className="bg-white rounded-3xl p-6 sm:p-8 border border-slate-100 shadow-sm space-y-4">
              {studentsInPreviewClass.length === 0 ? (
                <div className="text-center py-12 text-slate-500 italic">
                  Tidak ada data siswa ditemukan di kelas {selectedClassPreview}.
                </div>
              ) : activeAssignmentsColumns.length === 0 ? (
                <div className="text-center py-12 bg-slate-50 rounded-2xl border border-dashed border-slate-200 p-6 space-y-3">
                  <p className="text-xs font-bold text-slate-600">
                    Belum ada tugas dibuat untuk kelas {selectedClassPreview}.
                  </p>
                  <button
                    type="button"
                    onClick={() => setActiveSubTab('input')}
                    className="px-4 py-2 bg-emerald-600 text-white font-bold text-xs rounded-xl shadow-xs hover:bg-emerald-700 transition-all cursor-pointer"
                  >
                    + Buat Tugas Sekarang
                  </button>
                </div>
              ) : (
                <div className="overflow-x-auto max-h-[600px] border border-slate-100 rounded-2xl scrollbar-thin">
                  <table className="w-full text-left text-xs sm:text-sm border-collapse">
                    <thead className="sticky top-0 bg-slate-100/90 backdrop-blur-md z-10 shadow-[0_1px_0_0_rgba(226,232,240,1)]">
                      <tr>
                        <th className="p-3.5 sm:p-4 font-black text-slate-700 w-12 text-center border-b border-slate-200">
                          No
                        </th>
                        <th className="p-3.5 sm:p-4 font-black text-slate-700 min-w-[180px] border-b border-slate-200">
                          Nama Lengkap Siswa
                        </th>
                        <th className="p-3.5 sm:p-4 font-black text-slate-700 w-28 border-b border-slate-200">
                          Nis
                        </th>

                        {/* Dynamic Columns for each Assignment */}
                        {activeAssignmentsColumns.map((assign) => (
                          <th
                            key={assign.id}
                            className="p-3.5 sm:p-4 font-black text-slate-700 text-center min-w-[130px] border-b border-slate-200"
                          >
                            <div className="flex flex-col items-center">
                              <span className="text-slate-900 font-extrabold">{assign.title}</span>
                              <span className="text-[10px] text-slate-500 font-medium mt-0.5">{assign.date}</span>
                            </div>
                          </th>
                        ))}

                        {activeAssignmentsColumns.length > 1 && (
                          <th className="p-3.5 sm:p-4 font-black text-emerald-800 text-center w-28 bg-emerald-50/80 border-b border-slate-200">
                            Rata-Rata
                          </th>
                        )}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {studentsInPreviewClass.map((student, idx) => {
                        let totalScore = 0;
                        let gradedCount = 0;

                        return (
                          <tr key={student.id} className="hover:bg-slate-50/80 transition-colors">
                            <td className="p-3.5 sm:p-4 text-center font-bold text-slate-500">{idx + 1}</td>
                            <td className="p-3.5 sm:p-4 font-extrabold text-slate-800">{student.name}</td>
                            <td className="p-3.5 sm:p-4 font-semibold text-slate-600">{student.nisn || '-'}</td>

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
                                <td key={assign.id} className="p-3.5 sm:p-4 text-center font-extrabold text-slate-800 text-xs sm:text-sm">
                                  {hasScore ? (
                                    scoreVal
                                  ) : (
                                    <span className="text-slate-400 font-medium">-</span>
                                  )}
                                </td>
                              );
                            })}

                            {/* Average Score Column */}
                            {activeAssignmentsColumns.length > 1 && (
                              <td className="p-3.5 sm:p-4 text-center font-black text-slate-800 text-xs sm:text-sm">
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
              className="relative bg-white rounded-3xl p-6 sm:p-8 max-w-lg w-full shadow-2xl border border-slate-100 z-10 space-y-6"
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
              className="relative bg-white rounded-3xl p-6 sm:p-8 max-w-md w-full shadow-2xl border border-slate-100 z-10 space-y-5"
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
      </AnimatePresence>
    </div>
  );
}

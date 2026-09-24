import React, { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  GraduationCap,
  LayoutGrid,
  Building2,
  Fingerprint,
  BarChart3,
  Users,
  FileText,
  ClipboardList,
  Settings2,
  Plus, 
  Pencil,
  Trash2, 
  Calendar, 
  Check, 
  Save,
  AlertCircle,
  AlertTriangle,
  Info,
  User as UserIcon,
  ArrowRight,
  Eye,
  EyeOff,
  LogOut,
  X,
  Camera,
  Key,
  Download,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  CheckCircle,
  Cloud,
  Database,
  Activity,
  RefreshCw,
  Copy,
  UserX,
  Power,
  Loader2,
  Send,
  Sparkles,
  Clock,
  Timer
} from 'lucide-react';
import { 
  format, 
  addMonths, 
  subMonths, 
  startOfMonth, 
  endOfMonth, 
  startOfWeek, 
  endOfWeek, 
  eachDayOfInterval, 
  isSameMonth, 
  isSameDay,
  parseISO 
} from 'date-fns';
import { id as idLocale } from 'date-fns/locale';
import { 
  collection, 
  doc, 
  setDoc, 
  deleteDoc, 
  writeBatch,
  onSnapshot,
  getDocs,
  getDoc,
  query,
  where,
  deleteField,
  Firestore,
  DocumentData
} from 'firebase/firestore';
import { 
  signInWithEmailAndPassword,
  createUserWithEmailAndPassword,
  signOut, 
  onAuthStateChanged,
  User,
  updateProfile,
  Auth
} from 'firebase/auth';
import { dbDefault, auth, OperationType, handleFirestoreError } from './firebase';
import * as XLSX from 'xlsx';
import { importExcelHelper } from './importExcelHelper';

import ReportsView from './ReportsView';
import { compareClasses, compareStudentsByClass } from './classSortUtils';

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

interface CustomUser {
  id: string;
  fullname: string;
  username: string;
  password?: string;
  createdAt?: string;
}

function AttendanceView({
  classList,
  students,
  attendanceSessions,
  showToast,
  activeDb,
  activeAuth,
  trackOp
}: {
  classList: string[];
  students: Student[];
  attendanceSessions: AttendanceSession[];
  showToast: (message: string, type: 'success' | 'info' | 'error') => void;
  activeDb: Firestore;
  activeAuth: Auth;
  trackOp: (type: 'read' | 'write', count?: number) => void;
}) {
  const [date, setDate] = useState(() => {
    return format(new Date(), 'yyyy-MM-dd');
  });
  const [isDateModalOpen, setIsDateModalOpen] = useState(false);
  const [calendarMonth, setCalendarMonth] = useState(new Date());
  
  const [selectedClass, setSelectedClass] = useState(() => {
    if (classList.length === 1) return classList[0];
    return '';
  });

  useEffect(() => {
    if (classList.length === 1 && selectedClass !== classList[0]) {
      setSelectedClass(classList[0]);
    }
  }, [classList, selectedClass]);
  const [isClassModalOpen, setIsClassModalOpen] = useState(false);
  
  const [currentRecords, setCurrentRecords] = useState<Record<string, Status>>({});
  const [isEditing, setIsEditing] = useState(false);
  const [existingSession, setExistingSession] = useState<AttendanceSession | null>(null);

  const studentsInClass = useMemo(() => students.filter(s => s.class === selectedClass), [students, selectedClass]);
  const classSessions = useMemo(() => attendanceSessions.filter(s => s.className === selectedClass), [attendanceSessions, selectedClass]);
  
  // Calculate meeting number for next or current session
  const meetingNumber = useMemo(() => {
    if (existingSession) return existingSession.meetingNumber;
    return classSessions.length + 1;
  }, [existingSession, classSessions]);

  const daysInMonth = useMemo(() => {
    return eachDayOfInterval({
      start: startOfWeek(startOfMonth(calendarMonth), { weekStartsOn: 1 }),
      end: endOfWeek(endOfMonth(calendarMonth), { weekStartsOn: 1 })
    });
  }, [calendarMonth]);

  // Warn if leaving page with unsaved changes
  useEffect(() => {
    const handleBeforeUnload = (e: BeforeUnloadEvent) => {
      if (isEditing && studentsInClass.length > 0) {
        e.preventDefault();
        e.returnValue = '';
      }
    };
    window.addEventListener('beforeunload', handleBeforeUnload);
    return () => window.removeEventListener('beforeunload', handleBeforeUnload);
  }, [isEditing, studentsInClass.length]);

  useEffect(() => {
    if (!selectedClass || studentsInClass.length === 0) {
      setCurrentRecords({});
      setExistingSession(null);
      setIsEditing(false);
      return;
    }

    const found = attendanceSessions.find(s => s.className === selectedClass && s.date === date);

    if (found) {
      setExistingSession(found);
      setCurrentRecords(found.records);
      setIsEditing(false); // Default view mode when loaded
    } else {
      setExistingSession(null);
      setIsEditing(true); // new entry
      // Default to 'Hadir' for all students
      const initial: Record<string, Status> = {};
      studentsInClass.forEach(s => {
        initial[s.id] = 'Hadir';
      });
      setCurrentRecords(initial);
    }
  }, [date, selectedClass, attendanceSessions, studentsInClass]);

  const handleStatusChange = (studentId: string, status: Status) => {
    setCurrentRecords(prev => ({ ...prev, [studentId]: status }));
  };

  const handleDeleteSession = async (sessionId: string) => {
    try {
      trackOp('write', 1);
      await deleteDoc(doc(activeDb, 'attendanceSessions', sessionId));
      showToast('Data presensi berhasil dihapus.', 'info');
      const currentSessionId = existingSession?.id || `${activeAuth.currentUser?.uid}_${date}_${selectedClass}`;
      if (sessionId === currentSessionId) {
        setExistingSession(null);
        setIsEditing(true);
        const initial: Record<string, Status> = {};
        studentsInClass.forEach(s => { initial[s.id] = 'Hadir'; });
        setCurrentRecords(initial);
      }
    } catch (err) {
      showToast('Gagal menghapus presensi.', 'error');
      handleFirestoreError(err, OperationType.DELETE, 'attendanceSessions');
    }
  };

  const handleShareWA = () => {
    if (!existingSession) return;
    
    const absents = studentsInClass.filter(s => currentRecords[s.id] !== 'Hadir');
    
    let text = 'Bismillah\n';
    text += 'Asalamualaikum wr wb\n';
    text += 'Berikut laporan kehadiran siswa untuk hari ini : \n';
    text += '*Laporan Presensi*\n';
    text += 'Kelas: ' + selectedClass + '\n';
    text += 'Tanggal: ' + format(new Date(date), 'EEEE, dd MMMM yyyy', { locale: idLocale }) + '\n';
    text += 'Pertemuan ke: ' + meetingNumber + '\n\n';
    
    if (absents.length === 0) {
      text += '_Semua siswa hadir._\n\n';
    } else {
      text += '*Siswa yang tidak hadir:*\n';
      absents.forEach((s, idx) => {
        text += (idx + 1) + '. ' + s.name + ' (' + currentRecords[s.id] + ')\n';
      });
    }
    text += 'Terimakasih 🙏';
    
    const encodedText = encodeURIComponent(text);
    window.open('https://wa.me/?text=' + encodedText, '_blank');
  };

  const handleSave = async () => {
    if (!selectedClass) {
      showToast('Mohon pilih kelas terlebih dahulu.', 'error');
      return;
    }
    if (!activeAuth.currentUser) {
      showToast('Sesi berakhir. Silakan masuk kembali.', 'error');
      return;
    }
    
    // Check missing
    const missing = studentsInClass.some(s => !currentRecords[s.id]);
    if (missing) {
      showToast('Mohon lengkapi semua status presensi.', 'error');
      return;
    }

    const sessionId = existingSession ? existingSession.id : `${activeAuth.currentUser.uid}_${date}_${selectedClass}`;
    
    const newSession: AttendanceSession = {
      id: sessionId,
      date,
      className: selectedClass,
      meetingNumber: meetingNumber,
      records: currentRecords,
      userId: activeAuth.currentUser.uid
    };

    try {
      trackOp('write', 1);
      await setDoc(doc(activeDb, 'attendanceSessions', sessionId), newSession);
      setExistingSession(newSession);
      setIsEditing(false);
      showToast(`Presensi ${selectedClass} tanggal ${format(new Date(date), 'dd/MM/yyyy')} berhasil disimpan.`, 'success');
    } catch (err) {
      showToast('Gagal menyimpan presensi.', 'error');
      handleFirestoreError(err, OperationType.WRITE, 'attendanceSessions');
    }
  };

  return (
    <div className="p-4 sm:p-6 lg:p-8 space-y-6 max-w-7xl mx-auto w-full pb-32">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-4">
         <div>
           <h2 className="text-2xl sm:text-3xl font-bold text-slate-800 tracking-tight">Presensi Kelas</h2>
           <p className="text-sm font-medium text-slate-600 mt-1">Catat kehadiran harian siswa</p>
         </div>
      </div>

      <div className="bg-white p-6 sm:p-8 rounded-[1.5rem] border-2 border-slate-300/60 shadow-sm space-y-6">
         <div className="grid grid-cols-1 sm:grid-cols-2 gap-6 sm:gap-8">
            <div>
              <label className="block text-xs font-bold text-slate-600 uppercase tracking-widest mb-2.5">Tanggal Presensi</label>
              <button 
                onClick={() => setIsDateModalOpen(true)}
                className="w-full px-4 py-3.5 bg-slate-50/50 border-2 border-slate-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-emerald-600/20 focus:border-emerald-600 transition-all outline-none text-left font-semibold flex items-center justify-between group hover:bg-white hover:border-emerald-400"
              >
                <span className="text-slate-800 font-bold">
                  {format(parseISO(date), 'dd MMM yyyy')}
                </span>
                <ChevronDown className="w-5 h-5 text-slate-500 group-hover:text-emerald-600 transition-colors" />
              </button>
            </div>
            <div className="relative">
              <div className="flex items-center justify-between mb-2.5">
                <label className="block text-xs font-bold text-slate-600 uppercase tracking-widest">Pilih Kelas</label>
                {classList.length === 1 && (
                  <span className="text-[10px] font-bold text-emerald-800 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">
                    Kelas Wali
                  </span>
                )}
              </div>
              <button 
                onClick={() => {
                  if (classList.length > 1) {
                    setIsClassModalOpen(true);
                  }
                }}
                disabled={classList.length === 1}
                className="w-full px-4 py-3.5 bg-slate-50/50 border-2 border-slate-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-emerald-600/20 focus:border-emerald-600 transition-all outline-none text-left font-semibold flex items-center justify-between group hover:bg-white hover:border-emerald-400 disabled:bg-slate-100 disabled:border-slate-200 disabled:cursor-not-allowed"
              >
                 <span className={selectedClass ? "text-slate-800 font-bold" : "text-slate-500"}>
                   {selectedClass || "-- Pilih Kelas --"}
                 </span>
                 {classList.length > 1 ? (
                   <ChevronDown className="w-5 h-5 text-slate-500 group-hover:text-emerald-600 transition-colors" />
                 ) : (
                   <Check className="w-4 h-4 text-emerald-600" />
                 )}
              </button>
            </div>
         </div>
      </div>
      
      {/* Modern Date Selection Modal */}
      {isDateModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 sm:p-6 bg-slate-900/40 backdrop-blur-sm animate-in fade-in duration-200">
           <div className="bg-white rounded-[2rem] w-full max-w-sm shadow-2xl border border-white/20 overflow-hidden flex flex-col animate-in slide-in-from-bottom-8 zoom-in-95 duration-300">
             <div className="px-6 py-5 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
               <div>
                  <h3 className="text-xl font-black text-slate-800 tracking-tight">Pilih Tanggal</h3>
                  <p className="text-xs font-semibold text-slate-600 mt-0.5">Tentukan tanggal presensi</p>
               </div>
               <button 
                 onClick={() => setIsDateModalOpen(false)}
                 className="p-2 hover:bg-slate-200/50 text-slate-600 hover:text-slate-700 rounded-xl transition-colors shrink-0"
               >
                 <X className="w-5 h-5" />
               </button>
             </div>
             
             <div className="p-6">
                <div className="flex items-center justify-between mb-6">
                   <button 
                     onClick={() => setCalendarMonth(prev => subMonths(prev, 1))}
                     className="p-2 rounded-xl text-slate-600 hover:bg-slate-100 transition-colors"
                   >
                     <ChevronLeft className="w-5 h-5" />
                   </button>
                   <span className="font-bold text-slate-800 tracking-tight">
                     {format(calendarMonth, 'MMMM yyyy')}
                   </span>
                   <button 
                     onClick={() => setCalendarMonth(prev => addMonths(prev, 1))}
                     className="p-2 rounded-xl text-slate-600 hover:bg-slate-100 transition-colors"
                   >
                     <ChevronRight className="w-5 h-5" />
                   </button>
                </div>
                
                <div className="grid grid-cols-7 gap-1 text-center mb-2">
                   {['Sn', 'Sl', 'Rb', 'Km', 'Jm', 'Sb', 'Mg'].map(day => (
                     <div key={day} className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">{day}</div>
                   ))}
                </div>
                
                <div className="grid grid-cols-7 gap-1">
                   {daysInMonth.map(day => {
                     const dayStr = format(day, 'yyyy-MM-dd');
                     const isSelected = dayStr === date;
                     const isCurrentMonth = isSameMonth(day, calendarMonth);
                     const isToday = isSameDay(day, new Date());
                     const hasAttendance = selectedClass && attendanceSessions.some(s => s.className === selectedClass && s.date === dayStr);
                     
                     return (
                       <button
                         key={day.toISOString()}
                         onClick={() => {
                           setDate(dayStr);
                           setIsDateModalOpen(false);
                         }}
                         disabled={!isCurrentMonth}
                         className={`aspect-square flex flex-col items-center justify-center rounded-xl text-sm font-bold transition-all relative ${
                           !isCurrentMonth ? 'text-transparent cursor-default' :
                           isSelected ? 'bg-emerald-600 text-white shadow-md shadow-emerald-600/20 scale-105' :
                           hasAttendance ? 'text-emerald-600 bg-emerald-50/50 hover:bg-emerald-100/50 outline outline-1 outline-lime-200/50' : 'text-slate-700 hover:bg-slate-50'
                         }`}
                       >
                         {isCurrentMonth ? format(day, 'd') : ''}
                         <div className="absolute bottom-1.5 flex gap-1 items-center justify-center">
                           {isToday && (
                             <div className={`w-1 h-1 rounded-full ${isSelected ? 'bg-white/70' : 'bg-slate-300'}`}></div>
                           )}
                           {hasAttendance && (
                             <div className={`w-1.5 h-1.5 rounded-full shadow-sm ${isSelected ? 'bg-white' : 'bg-sky-500'}`}></div>
                           )}
                         </div>
                       </button>
                     );
                   })}
                </div>
             </div>
           </div>
        </div>
      )}

      {/* Modern Class Selection Modal */}
      {isClassModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 sm:p-6 bg-slate-900/40 backdrop-blur-sm animate-in fade-in duration-200">
           <div className="bg-white rounded-[2rem] w-full max-w-md shadow-2xl border border-white/20 overflow-hidden flex flex-col animate-in slide-in-from-bottom-8 zoom-in-95 duration-300">
             <div className="px-6 py-5 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
               <div>
                  <h3 className="text-xl font-black text-slate-800 tracking-tight">Pilih Kelas</h3>
                  <p className="text-xs font-semibold text-slate-600 mt-0.5">Tentukan kelas untuk presensi</p>
               </div>
               <button 
                 onClick={() => setIsClassModalOpen(false)}
                 className="p-2 hover:bg-slate-200/50 text-slate-600 hover:text-slate-700 rounded-xl transition-colors shrink-0"
               >
                 <X className="w-5 h-5" />
               </button>
             </div>
             
             <div className="p-6 max-h-[60vh] overflow-y-auto">
                <div className="grid grid-cols-2 gap-3 sm:gap-4">
                  {classList.slice().sort(compareClasses).map(c => (
                    <button
                      key={c}
                      onClick={() => {
                        setSelectedClass(c);
                        setIsClassModalOpen(false);
                      }}
                      className={`relative p-4 rounded-2xl border-2 transition-all duration-200 flex flex-col items-center justify-center gap-2 group ${
                        selectedClass === c 
                          ? 'border-emerald-600 bg-emerald-50 text-lime-700 shadow-sm' 
                          : 'border-slate-100 bg-white hover:border-emerald-400 hover:bg-emerald-50/30 text-slate-600'
                      }`}
                    >
                      <div className={`w-10 h-10 rounded-xl flex items-center justify-center transition-colors ${
                        selectedClass === c ? 'bg-emerald-600 text-white' : 'bg-slate-100 text-slate-500 group-hover:bg-emerald-100 group-hover:text-emerald-700'
                      }`}>
                         <Building2 className="w-5 h-5" />
                      </div>
                      <span className="font-bold text-sm">{c}</span>
                      {selectedClass === c && (
                        <div className="absolute top-2 right-2 w-2 h-2 rounded-full bg-emerald-600"></div>
                      )}
                    </button>
                  ))}
                  {classList.length === 0 && (
                    <div className="col-span-2 text-center py-8">
                       <p className="text-slate-600 font-medium text-sm">Belum ada kelas yang terdaftar.</p>
                    </div>
                  )}
                </div>
             </div>
           </div>
        </div>
      )}

      {selectedClass && studentsInClass.length === 0 && (
         <div className="bg-amber-50 border border-amber-200 text-amber-700 p-6 rounded-[2rem] text-center font-medium">
           Belum ada data siswa di kelas {selectedClass}. <br/> Silakan tambahkan siswa terlebih dahulu di menu <b>Manajemen Siswa</b>.
         </div>
      )}

      {selectedClass && studentsInClass.length > 0 && (
        <div className="bg-white rounded-[1.5rem] border-2 border-slate-300/60 shadow-sm overflow-hidden flex flex-col">
          <div className="p-6 border-b border-slate-100 bg-white flex flex-col sm:flex-row sm:items-center justify-between gap-4">
             <div>
               <h3 className="text-lg font-bold text-slate-800 tracking-tight">Tabel Presensi</h3>
               {existingSession && !isEditing ? (
                  <p className="text-sm font-bold text-emerald-600 mt-1 flex items-center gap-1.5">
                     <Check className="w-4 h-4" /> Data sudah tersimpan
                  </p>
               ) : existingSession && isEditing ? (
                  <p className="text-sm font-bold text-amber-600 mt-1 flex items-center gap-1.5">
                     <Pencil className="w-4 h-4" /> Mode Edit Data Presensi
                  </p>
               ) : null}
             </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse min-w-[700px]">
              <thead>
                <tr className="bg-slate-50/50 border-b border-slate-100">
                  <th className="p-4 w-12 font-semibold text-slate-600 text-xs uppercase tracking-widest text-center">No</th>
                  <th className="p-4 font-semibold text-slate-600 text-xs uppercase tracking-widest">Nama Lengkap</th>
                  <th className="p-4 font-semibold text-slate-600 text-xs uppercase tracking-widest text-center">Keterangan</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {studentsInClass.map((student, index) => (
                  <tr key={student.id} className="hover:bg-slate-50/30 transition-colors group">
                    <td className="p-4 text-center font-medium text-slate-500 text-sm">{index + 1}</td>
                    <td className="p-4 font-semibold text-slate-700 group-hover:text-slate-900 transition-colors">{student.name}</td>
                    <td className="p-4">
                       <div className="flex items-center justify-end gap-2 sm:gap-2.5">
                         {(['Hadir', 'Sakit', 'Izin', 'Alpa', 'Dispen'] as Status[]).map(statusOpt => {
                            const isSelected = currentRecords[student.id] === statusOpt;
                            
                            let colorClass = 'border-slate-200 text-slate-600 hover:border-slate-300 bg-white';
                            if (isSelected) {
                               if (statusOpt === 'Hadir') colorClass = 'bg-emerald-500 text-white border-emerald-500 shadow-md shadow-emerald-500/20';
                               else if (statusOpt === 'Sakit') colorClass = 'bg-sky-500 text-white border-sky-500 shadow-md shadow-sky-500/20';
                               else if (statusOpt === 'Izin') colorClass = 'bg-amber-500 text-white border-amber-500 shadow-md shadow-amber-500/20';
                               else if (statusOpt === 'Alpa') colorClass = 'bg-rose-500 text-white border-rose-500 shadow-md shadow-rose-500/20';
                               else if (statusOpt === 'Dispen') colorClass = 'bg-indigo-500 text-white border-indigo-500 shadow-md shadow-indigo-500/20';
                            }

                            return (
                              <button
                                key={statusOpt}
                                disabled={!isEditing}
                                onClick={() => handleStatusChange(student.id, statusOpt)}
                                className={`px-4 py-2 sm:py-2.5 rounded-lg text-xs sm:text-sm font-semibold border ${colorClass} transition-all disabled:opacity-50 disabled:cursor-not-allowed`}
                              >
                                {statusOpt}
                              </button>
                            );
                         })}
                       </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="p-6 border-t border-slate-100 bg-white flex flex-col gap-3">
             <div className="flex items-center justify-end gap-2 w-full sm:w-auto">
               {isEditing ? (
                 <>
                    {existingSession && (
                       <button
                          onClick={() => {
                             setIsEditing(false);
                             setCurrentRecords(existingSession.records);
                          }}
                          className="flex-1 sm:flex-none sm:w-auto bg-white border-2 border-slate-300 text-slate-700 px-6 py-2.5 rounded-lg font-semibold text-sm hover:bg-slate-50 transition-all active:scale-95"
                       >
                          Batal
                       </button>
                    )}
                    <button 
                       onClick={handleSave} 
                       className="flex-1 sm:flex-none sm:w-auto bg-emerald-600 text-white px-6 py-2.5 rounded-lg font-semibold text-sm hover:bg-emerald-700 transition-all active:scale-95 flex items-center justify-center gap-2"
                    >
                       <Save className="w-4 h-4" /> Simpan Presensi
                    </button>
                 </>
               ) : (
                 <>
                    <button
                       onClick={() => setIsEditing(true)}
                       className="flex-1 sm:flex-none sm:w-auto bg-white border-2 border-slate-300 text-slate-700 px-6 py-2.5 rounded-lg font-semibold text-sm hover:bg-slate-50 transition-all flex items-center justify-center gap-2"
                    >
                       <Pencil className="w-4 h-4" /> Edit
                    </button>
                    <button
                       onClick={() => { if(existingSession) handleDeleteSession(existingSession.id); }}
                       className="flex-1 sm:flex-none sm:w-auto bg-white text-rose-600 border border-rose-200 px-6 py-2.5 rounded-lg font-semibold text-sm hover:bg-rose-50 transition-all active:scale-95 flex items-center justify-center gap-2"
                    >
                       <Trash2 className="w-4 h-4" /> Hapus
                    </button>
                 </>
               )}
             </div>
             
             {!isEditing && existingSession && (
                <div className="flex justify-end pt-2 border-t border-slate-50 mt-1">
                    <button
                       onClick={handleShareWA}
                       className="w-full sm:w-auto bg-[#25D366] text-white px-8 py-3 rounded-xl font-bold text-sm hover:bg-[#128C7E] transition-all active:scale-95 flex items-center justify-center gap-2 shadow-md shadow-[#25D366]/20"
                    >
                       <Send className="w-5 h-5" /> Kirim ke Group WA
                    </button>
                </div>
             )}
          </div>
        </div>
      )}

      {selectedClass && classSessions.length > 0 && (
        <div className="bg-white rounded-[1.5rem] border-2 border-slate-300/60 shadow-sm overflow-hidden flex flex-col mt-8">
          <div className="p-6 sm:p-8 border-b border-slate-100 bg-white flex flex-col items-center justify-center text-center gap-2">
            <h3 className="text-xl font-bold text-slate-800 tracking-tight">Riwayat Presensi</h3>
            <div className="text-center">
                <p className="text-[10px] font-bold text-slate-500 tracking-widest uppercase mb-1">Informasi Pertemuan</p>
                <p className="text-sm font-semibold text-slate-700">Total {classSessions.length} Pertemuan Efektif</p>
            </div>
          </div>
          <div className="grid grid-cols-4 gap-2 sm:gap-4 p-4 sm:p-6 bg-white">
            {[...classSessions].sort((a,b) => new Date(a.date).getTime() - new Date(b.date).getTime()).map((session, idx) => (
              <button 
                key={session.id} 
                onClick={() => {
                  setDate(session.date);
                  window.scrollTo({ top: 0, behavior: 'smooth' });
                }}
                className="py-1.5 flex items-center justify-center gap-2 hover:text-emerald-600 transition-colors cursor-pointer focus:outline-none group"
              >
                <span className="font-medium text-slate-500 text-xs group-hover:text-[#7bc025] transition-colors">{idx + 1}.</span>
                <span className="font-medium text-slate-600 text-xs sm:text-sm group-hover:text-emerald-600 transition-colors">{format(new Date(session.date), 'dd MMM yyyy')}</span>
              </button>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

const toAuthEmail = (input: string) => {
  const trimmed = input.toLowerCase().trim();
  if (trimmed.includes('@')) {
    return trimmed;
  }
  return `${trimmed}@kaguci.com`;
};

const encodeSession = (username: string, pass: string) => {
  try {
    return btoa(encodeURIComponent(JSON.stringify({ u: username, p: pass })));
  } catch {
    return '';
  }
};

const decodeSession = (hashStr: string) => {
  try {
    const raw = decodeURIComponent(atob(hashStr));
    const parsed = JSON.parse(raw);
    return { username: parsed.u, password: parsed.p };
  } catch {
    return null;
  }
};

const safeSetLocalStorage = (key: string, value: string) => {
  try {
    localStorage.setItem(key, value);
  } catch (e) {
    console.warn(`LocalStorage quota exceeded or disabled for key "${key}":`, e);
  }
};

const speakText = (text: string) => {
  if (typeof window !== 'undefined' && 'speechSynthesis' in window) {
    window.speechSynthesis.cancel();
    
    const startSpeaking = () => {
      const utterance = new SpeechSynthesisUtterance(text);
      utterance.lang = 'id-ID';
      utterance.rate = 0.95;
      utterance.pitch = 1.0;
      
      const voices = window.speechSynthesis.getVoices();
      
      // Filter all Indonesian voices using multiple language code variations
      const indonesianVoices = voices.filter(v => {
        const langLower = v.lang.toLowerCase();
        return langLower.startsWith('id') || langLower.includes('id') || langLower === 'id_id';
      });

      // Sort to prioritize natural premium voices (like Google Bahasa Indonesia or professional female voices)
      indonesianVoices.sort((a, b) => {
        const score = (v: SpeechSynthesisVoice) => {
          const nameLower = v.name.toLowerCase();
          let pts = 0;
          if (nameLower.includes('google')) pts += 10;
          if (nameLower.includes('premium') || nameLower.includes('natural')) pts += 5;
          if (nameLower.includes('female') || nameLower.includes('wanita') || nameLower.includes('gadis')) pts += 3;
          return pts;
        };
        return score(b) - score(a);
      });

      const selectedVoice = indonesianVoices[0];
      if (selectedVoice) {
        utterance.voice = selectedVoice;
        console.log(`Menggunakan suara Bahasa Indonesia: ${selectedVoice.name} (${selectedVoice.lang})`);
      } else {
        console.warn("Suara Bahasa Indonesia asli tidak ditemukan. Menggunakan suara sistem default.");
      }
      
      window.speechSynthesis.speak(utterance);
    };

    if (window.speechSynthesis.getVoices().length === 0) {
      const handleVoicesChanged = () => {
        startSpeaking();
        window.speechSynthesis.removeEventListener('voiceschanged', handleVoicesChanged);
      };
      window.speechSynthesis.addEventListener('voiceschanged', handleVoicesChanged);
    } else {
      startSpeaking();
    }
  } else {
    console.warn("Speech Synthesis tidak didukung oleh browser ini.");
  }
};

// Helper to calculate 14:00 WIB daily reset cycle and time info
export interface WibCycleInfo {
  cycleKey: string;
  wibTime: Date;
  currentDateFormatted: string;
  currentTimeStr: string;
  currentDayName: string;
  cycleStartDate: Date;
  nextResetDate: Date;
  countdownFormatted: string;
  hoursLeft: number;
  minutesLeft: number;
  secondsLeft: number;
}

export const getWibCycleInfo = (now: Date = new Date()): WibCycleInfo => {
  // Convert current time to WIB (UTC+7)
  const utc = now.getTime() + (now.getTimezoneOffset() * 60000);
  const wibTime = new Date(utc + (7 * 3600000));
  
  const wibHours = wibTime.getHours();
  const wibMinutes = wibTime.getMinutes();
  const wibSeconds = wibTime.getSeconds();

  // If before 14:00 WIB, the current cycle started yesterday at 14:00 WIB
  const cycleStartDate = new Date(wibTime);
  if (wibHours < 14) {
    cycleStartDate.setDate(cycleStartDate.getDate() - 1);
  }
  cycleStartDate.setHours(14, 0, 0, 0);

  // Next reset is at 14:00 WIB on the next day after cycleStartDate
  const nextResetDate = new Date(cycleStartDate);
  nextResetDate.setDate(nextResetDate.getDate() + 1);
  nextResetDate.setHours(14, 0, 0, 0);

  // Cycle key uniquely identifies the 24-hour cycle starting at 14:00 WIB
  const y = cycleStartDate.getFullYear();
  const m = String(cycleStartDate.getMonth() + 1).padStart(2, '0');
  const d = String(cycleStartDate.getDate()).padStart(2, '0');
  const cycleKey = `${y}-${m}-${d}_14:00_WIB`;

  // Indonesian day and month names
  const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  
  const currentDayName = days[wibTime.getDay()];
  const currentDayNum = wibTime.getDate();
  const currentMonthName = months[wibTime.getMonth()];
  const currentYear = wibTime.getFullYear();
  const currentTimeStr = `${String(wibHours).padStart(2, '0')}:${String(wibMinutes).padStart(2, '0')}:${String(wibSeconds).padStart(2, '0')} WIB`;
  const currentDateFormatted = `${currentDayName}, ${currentDayNum} ${currentMonthName} ${currentYear}`;

  // Time remaining until next reset (in ms)
  const msUntilReset = Math.max(0, nextResetDate.getTime() - wibTime.getTime());
  const hoursLeft = Math.floor(msUntilReset / 3600000);
  const minutesLeft = Math.floor((msUntilReset % 3600000) / 60000);
  const secondsLeft = Math.floor((msUntilReset % 60000) / 1000);
  const countdownFormatted = `${hoursLeft} jam ${minutesLeft} menit ${secondsLeft} detik`;

  return {
    cycleKey,
    wibTime,
    currentDateFormatted,
    currentTimeStr,
    currentDayName,
    cycleStartDate,
    nextResetDate,
    countdownFormatted,
    hoursLeft,
    minutesLeft,
    secondsLeft
  };
};

export default function App() {
  const [activeUserCustomData, setActiveUserCustomData] = useState<{
    fullname: string;
    username: string;
    configText: string;
    password?: string;
    role?: string;
  } | null>(() => {
    const saved = localStorage.getItem('kaguci_active_custom_user');
    return saved ? JSON.parse(saved) : null;
  });

  const { activeAuth, activeDb } = useMemo(() => {
    return { activeAuth: auth, activeDb: dbDefault };
  }, []);

  const [isLoggedIn, setIsLoggedIn] = useState<boolean>(() => {
    return localStorage.getItem('kaguci_has_logged_in') === 'true';
  });
  const [currentUser, setCurrentUser] = useState<User | null>(null);
  const [isAuthLoading, setIsAuthLoading] = useState<boolean>(true);
  const [activeTab, setActiveTab] = useState<'dashboard' | 'students' | 'attendance' | 'reports' | 'homeroom_report' | 'profile'>('dashboard');

  // Live WIB Cycle Information & 14:00 WIB Clock
  const [cycleInfo, setCycleInfo] = useState<WibCycleInfo>(() => getWibCycleInfo());

  // Firestore Usage Tracking & Quota Modal - Resets daily at 14:00 WIB
  const [sessionUsage, setSessionUsage] = useState(() => {
    const info = getWibCycleInfo();
    const saved = localStorage.getItem('kaguci_session_usage');
    const savedCycle = localStorage.getItem('kaguci_session_cycle');
    
    // If we have saved usage for the exact current 14:00 WIB cycle, load it
    if (saved && savedCycle === info.cycleKey) {
      try {
        return JSON.parse(saved);
      } catch {
        // fallback
      }
    }
    
    // Reset to 0 for a new 14:00 WIB cycle
    localStorage.setItem('kaguci_session_cycle', info.cycleKey);
    localStorage.setItem('kaguci_session_date', info.currentDateFormatted);
    return { reads: 0, writes: 0 };
  });
  const [showQuotaModal, setShowQuotaModal] = useState(false);
  const [showLogoutConfirm, setShowLogoutConfirm] = useState(false);
  const [showIdleTip, setShowIdleTip] = useState(false);
  const [showHomeroomRoleAlert, setShowHomeroomRoleAlert] = useState(false);

  // State for PWA installation & iframe detection
  interface PWAInstallPromptEvent extends Event {
    prompt: () => Promise<void>;
    userChoice: Promise<{ outcome: 'accepted' | 'dismissed' }>;
  }

  const [deferredPrompt, setDeferredPrompt] = useState<PWAInstallPromptEvent | null>(null);
  const [isInstallable, setIsInstallable] = useState(false);
  const [isAppInstalled, setIsAppInstalled] = useState(false);
  const isInsideIframe = useMemo(() => {
    try {
      return window.self !== window.top;
    } catch {
      return true;
    }
  }, []);

  useEffect(() => {
    const handleBeforeInstallPrompt = (e: Event) => {
      e.preventDefault();
      setDeferredPrompt(e as PWAInstallPromptEvent);
      setIsInstallable(true);
    };

    window.addEventListener('beforeinstallprompt', handleBeforeInstallPrompt);

    const handleAppInstalled = () => {
      setIsAppInstalled(true);
      setIsInstallable(false);
      setDeferredPrompt(null);
    };

    window.addEventListener('appinstalled', handleAppInstalled);

    const iosStandalone = (navigator as unknown as { standalone?: boolean }).standalone;
    if (window.matchMedia('(display-mode: standalone)').matches || iosStandalone) {
      setIsAppInstalled(true);
    }

    return () => {
      window.removeEventListener('beforeinstallprompt', handleBeforeInstallPrompt);
      window.removeEventListener('appinstalled', handleAppInstalled);
    };
  }, []);

  const handleInstallClick = async () => {
    if (!deferredPrompt) {
      alert("Instruksi instalasi: Silakan klik tombol 'Buka di Tab Baru' di kanan atas preview, lalu pilih opsi 'Instal' atau 'Tambahkan ke Layar Utama' dari menu browser Anda.");
      return;
    }
    await deferredPrompt.prompt();
    const { outcome } = await deferredPrompt.userChoice;
    console.log(`User response to the install prompt: ${outcome}`);
    if (outcome === 'accepted') {
      setIsInstallable(false);
      setDeferredPrompt(null);
    }
  };

  useEffect(() => {
    if (!isLoggedIn) return;
    
    let idleTimeout: NodeJS.Timeout;
    let tipTimeout: NodeJS.Timeout;
    
    const resetIdleTimer = () => {
      clearTimeout(idleTimeout);
      clearTimeout(tipTimeout);
      setShowIdleTip(false);
      
      // Muncul setelah 30 detik tidak ada aktivitas (biar gampang ditest) -> ganti jadi 60 detik (1 menit)
      idleTimeout = setTimeout(() => {
        setShowIdleTip(true);
        // Hilang lagi setelah 10 detik
        tipTimeout = setTimeout(() => {
          setShowIdleTip(false);
        }, 10000);
      }, 60000);
    };

    resetIdleTimer();

    window.addEventListener('mousemove', resetIdleTimer);
    window.addEventListener('keydown', resetIdleTimer);
    window.addEventListener('touchstart', resetIdleTimer);
    window.addEventListener('scroll', resetIdleTimer);

    return () => {
      clearTimeout(idleTimeout);
      clearTimeout(tipTimeout);
      window.removeEventListener('mousemove', resetIdleTimer);
      window.removeEventListener('keydown', resetIdleTimer);
      window.removeEventListener('touchstart', resetIdleTimer);
      window.removeEventListener('scroll', resetIdleTimer);
    };
  }, [isLoggedIn]);

  // Save session usage and active cycle
  useEffect(() => {
    const current = getWibCycleInfo();
    localStorage.setItem('kaguci_session_usage', JSON.stringify(sessionUsage));
    localStorage.setItem('kaguci_session_cycle', current.cycleKey);
    localStorage.setItem('kaguci_session_date', current.currentDateFormatted);
  }, [sessionUsage]);

  // Live WIB clock & Automatic 14:00 WIB daily reset monitor
  useEffect(() => {
    const timer = setInterval(() => {
      const newInfo = getWibCycleInfo();
      setCycleInfo(newInfo);

      const savedCycle = localStorage.getItem('kaguci_session_cycle');
      // If 14:00 WIB has just passed, trigger automatic reset to 0
      if (savedCycle && savedCycle !== newInfo.cycleKey) {
        console.log("⏰ 14:00 WIB reached! Resetting usage limits (reads & writes) to 0.");
        localStorage.setItem('kaguci_session_cycle', newInfo.cycleKey);
        localStorage.setItem('kaguci_session_date', newInfo.currentDateFormatted);
        setSessionUsage({ reads: 0, writes: 0 });

        // Update Firestore for current custom account immediately
        if (activeUserCustomData?.username) {
          const usernameKey = activeUserCustomData.username.toLowerCase().trim();
          setDoc(doc(dbDefault, 'custom_accounts', usernameKey), {
            dailyUsageReads: 0,
            dailyUsageWrites: 0,
            dailyUsageCycle: newInfo.cycleKey,
            dailyUsageDate: newInfo.currentDateFormatted,
            dailyUsageTime: newInfo.currentTimeStr,
            dailyUsageLastReset: newInfo.currentTimeStr
          }, { merge: true }).catch(err => console.warn('Reset sync to custom_accounts error:', err));
        }
      }
    }, 1000);
    return () => clearInterval(timer);
  }, [activeUserCustomData?.username]);

  const lastSyncedUsage = useRef({ reads: 0, writes: 0 });
  useEffect(() => {
    if (!activeUserCustomData?.username) return;
    const interval = setInterval(() => {
      if (sessionUsage.reads !== lastSyncedUsage.current.reads || sessionUsage.writes !== lastSyncedUsage.current.writes) {
        lastSyncedUsage.current = { ...sessionUsage };
        const usernameKey = activeUserCustomData.username.toLowerCase().trim();
        const current = getWibCycleInfo();
        setDoc(doc(dbDefault, 'custom_accounts', usernameKey), {
          dailyUsageReads: sessionUsage.reads || 0,
          dailyUsageWrites: sessionUsage.writes || 0,
          dailyUsageCycle: current.cycleKey,
          dailyUsageDate: current.currentDateFormatted,
          dailyUsageTime: current.currentTimeStr
        }, { merge: true }).catch(err => console.warn('Failed syncing usage', err));
      }
    }, 15000); // 15s flush
    return () => clearInterval(interval);
  }, [sessionUsage, activeUserCustomData?.username]);

  useEffect(() => {
    const handleQuota = () => setShowQuotaModal(true);
    window.addEventListener('firestore-quota-exceeded', handleQuota);
    return () => window.removeEventListener('firestore-quota-exceeded', handleQuota);
  }, []);

  const trackOp = useCallback((type: 'read' | 'write', count: number = 1) => {
    const current = getWibCycleInfo();
    const savedCycle = localStorage.getItem('kaguci_session_cycle');

    setSessionUsage((prev: { reads: number; writes: number }) => {
      // If cycle has shifted past 14:00 WIB, start fresh from 0 for the new cycle
      if (savedCycle && savedCycle !== current.cycleKey) {
        localStorage.setItem('kaguci_session_cycle', current.cycleKey);
        localStorage.setItem('kaguci_session_date', current.currentDateFormatted);
        return {
          reads: type === 'read' ? count : 0,
          writes: type === 'write' ? count : 0
        };
      }
      return {
        ...prev,
        [type === 'read' ? 'reads' : 'writes']: (prev[type === 'read' ? 'reads' : 'writes'] || 0) + count
      };
    });
  }, []);


  const [avatarUrl, setAvatarUrl] = useState<string | null>(() => {
    try {
      const savedUser = localStorage.getItem('kaguci_active_custom_user');
      if (savedUser) {
        const parsed = JSON.parse(savedUser);
        const username = parsed.username || '';
        const savedAvatar = localStorage.getItem(`kaguci_avatar_${username.toLowerCase()}`);
        return savedAvatar || null;
      }
    } catch {
      // Ignore
    }
    return null;
  });

  const [showSplash, setShowSplash] = useState(true);
  const [minSplashTimeElapsed, setMinSplashTimeElapsed] = useState(false);
  const [isFirstSessionLoad] = useState(() => !sessionStorage.getItem('kaguci_has_loaded'));

  useEffect(() => {
    if (isFirstSessionLoad) {
      sessionStorage.setItem('kaguci_has_loaded', 'true');
    }
    const timer = setTimeout(() => {
      setMinSplashTimeElapsed(true);
    }, isFirstSessionLoad ? 2500 : 800); // 2.5s for initial splash, 0.8s for fast reload
    return () => clearTimeout(timer);
  }, [isFirstSessionLoad]);

  useEffect(() => {
    if (minSplashTimeElapsed && !isAuthLoading) {
      setShowSplash(false);
    }
  }, [minSplashTimeElapsed, isAuthLoading]);

  const [profileData, setProfileData] = useState(() => {
    try {
      const savedUser = localStorage.getItem('kaguci_active_custom_user');
      if (savedUser) {
        const parsed = JSON.parse(savedUser);
        const username = parsed.username || '';
        const savedProfile = localStorage.getItem(`kaguci_profile_${username.toLowerCase()}`);
        if (savedProfile) {
          return JSON.parse(savedProfile);
        }
      }
    } catch {
      // Ignore
    }
    return {
      namaGuruMapel: '',
      namaKepalaSekolah: '',
      nipGuruMapel: '',
      nipKepalaSekolah: '',
      namaKurikulum: '',
      nipKurikulum: '',
      jabatanKurikulum: 'Wakasek Kurikulum',
      namaKesiswaan: '',
      nipKesiswaan: '',
      jabatanKesiswaan: 'Wakasek Kesiswaan',
      namaGuruWali: '',
      nipGuruWali: '',
      jabatanGuruWali: 'Guru Wali',
      namaBK: '',
      nipBK: '',
      jabatanBK: 'Guru BK',
      namaHumas: '',
      nipHumas: '',
      jabatanHumas: 'Wakasek Humas',
      semester: 'Ganjil',
      tahunPelajaran: '',
      mataPelajaran: '',
      role: 'Guru Mapel',
      waliKelasClass: '',
      jumlahSiswaLakiLaki: '',
      jumlahSiswaPerempuan: ''
    };
  });

  const handleOpenHomeroomReport = useCallback(() => {
    setActiveTab('homeroom_report');
  }, []);

  const hasWelcomedRef = useRef(false);

  useEffect(() => {
    if (isLoggedIn && !isAuthLoading) {
      const name = (profileData?.namaGuruMapel || activeUserCustomData?.fullname || activeUserCustomData?.username || "").trim();
      const sessionKey = `kaguci_welcomed_${activeUserCustomData?.username || 'user'}`;
      const sessionWelcomed = sessionStorage.getItem(sessionKey);
      
      if (!sessionWelcomed && !hasWelcomedRef.current) {
        if (!name) {
          const checkTimer = setTimeout(() => {
            const delayedName = (profileData?.namaGuruMapel || activeUserCustomData?.fullname || activeUserCustomData?.username || "").trim();
            hasWelcomedRef.current = true;
            sessionStorage.setItem(sessionKey, 'true');
            const phrase = delayedName 
              ? `Selamat datang ${delayedName} di aplikasi My Kaguci App.`
              : `Selamat datang di aplikasi My Kaguci App.`;
            speakText(phrase);
          }, 1200);
          return () => clearTimeout(checkTimer);
        } else {
          hasWelcomedRef.current = true;
          sessionStorage.setItem(sessionKey, 'true');
          const phrase = `Selamat datang ${name} di aplikasi My Kaguci App.`;
          const speakTimer = setTimeout(() => {
            speakText(phrase);
          }, 600);
          return () => clearTimeout(speakTimer);
        }
      }
    } else if (!isLoggedIn) {
      hasWelcomedRef.current = false;
    }
  }, [isLoggedIn, isAuthLoading, activeUserCustomData, profileData]);

  useEffect(() => {
    if (isLoggedIn && activeUserCustomData?.username && profileData) {
      const hasContent = Object.values(profileData).some(val => val !== '' && val !== 'Ganjil');
      if (hasContent) {
        safeSetLocalStorage(`kaguci_profile_${activeUserCustomData.username.toLowerCase()}`, JSON.stringify(profileData));
      }
    }
  }, [profileData, isLoggedIn, activeUserCustomData]);

  useEffect(() => {
    if (isLoggedIn && activeUserCustomData?.username && avatarUrl) {
      safeSetLocalStorage(`kaguci_avatar_${activeUserCustomData.username.toLowerCase()}`, avatarUrl);
    }
  }, [avatarUrl, isLoggedIn, activeUserCustomData]);

  const [isProfileEditing, setIsProfileEditing] = useState(false);
  const [isProfileSaving, setIsProfileSaving] = useState(false);

  // States & helper for Admin (Kumpulan User) - Custom Requested
  const isAdmin = useMemo(() => {
    const email = (activeAuth.currentUser?.email || '').toLowerCase().trim();
    const username = (activeUserCustomData?.username || '').toLowerCase().trim();
    const fullname = (activeUserCustomData?.fullname || profileData.namaGuruMapel || '').toLowerCase().trim();

    const isMatch = 
      email.includes('agan') || 
      username.includes('agan') || 
      fullname.includes('agan') ||
      username.includes('parta') ||
      fullname.includes('parta') ||
      fullname.includes('s.kom') ||
      username === 'admin' ||
      email === 'agan.parta@gmail.com';

    return isMatch;
  }, [activeAuth.currentUser, activeUserCustomData, profileData.namaGuruMapel]);

  const [allUsers, setAllUsers] = useState<CustomUser[]>([]);
  const [isLoadingUsers, setIsLoadingUsers] = useState<boolean>(false);
  const [userToDelete, setUserToDelete] = useState<CustomUser | null>(null);
  const [isDeletingUser, setIsDeletingUser] = useState<boolean>(false);
  
  const [userToEdit, setUserToEdit] = useState<CustomUser | null>(null);
  const [editFullname, setEditFullname] = useState<string>('');
  const [editPassword, setEditPassword] = useState<string>('');
  const [isSavingUser, setIsSavingUser] = useState<boolean>(false);

  const handleEditUserClick = (user: CustomUser) => {
    setUserToEdit(user);
    setEditFullname(user.fullname || '');
    setEditPassword(user.password || '');
  };

  const handleSaveEditedUser = async () => {
    if (!userToEdit) return;
    if (!editFullname.trim()) {
      showToast('Nama Lengkap tidak boleh kosong.', 'error');
      return;
    }
    setIsSavingUser(true);
    const path = `custom_accounts/${userToEdit.username.toLowerCase().trim()}`;
    try {
      trackOp('write', 1);
      const userDocRef = doc(dbDefault, 'custom_accounts', userToEdit.username.toLowerCase().trim());
      const updatedData = {
        fullname: editFullname.trim(),
        password: editPassword.trim()
      };
      await setDoc(userDocRef, updatedData, { merge: true });
      showToast(`User "${editFullname}" berhasil diperbarui.`, 'success');
      setUserToEdit(null);
      fetchAllUsers();
    } catch (err) {
      console.error('Error saving edited user:', err);
      handleFirestoreError(err, OperationType.WRITE, path);
      showToast('Gagal memperbarui user: ' + (err instanceof Error ? err.message : 'Server error'), 'error');
    } finally {
      setIsSavingUser(false);
    }
  };

  const fetchAllUsers = useCallback(async () => {
    if (!isAdmin) return;
    setIsLoadingUsers(true);
    try {
      trackOp('read', 1);
      const qSnapshot = await getDocs(collection(dbDefault, 'custom_accounts'));
      const list: CustomUser[] = [];
      qSnapshot.forEach((docSnap) => {
        const u = docSnap.data();
        list.push({
          id: docSnap.id,
          fullname: u.fullname || docSnap.id,
          username: u.username || docSnap.id,
          password: u.password || '••••••••',
          createdAt: u.createdAt || ''
        });
      });

      // Ensure active user is included in the list
      const activeName = activeUserCustomData?.fullname || profileData.namaGuruMapel || 'Agan Parta,S.Kom.';
      const activeUn = activeUserCustomData?.username || 'agan.parta';
      const exists = list.some(u => u.username.toLowerCase().trim() === activeUn.toLowerCase().trim());
      if (!exists) {
        list.unshift({
          id: activeUn.toLowerCase().trim(),
          fullname: activeName,
          username: activeUn,
          password: '••••••••',
          createdAt: new Date().toISOString()
        });
      }

      list.sort((a, b) => (a.fullname || '').localeCompare(b.fullname || ''));
      setAllUsers(list);
    } catch (err) {
      console.error('Error fetching users:', err);
      const activeName = activeUserCustomData?.fullname || profileData.namaGuruMapel || 'Agan Parta,S.Kom.';
      const activeUn = activeUserCustomData?.username || 'agan.parta';
      setAllUsers([{
        id: activeUn.toLowerCase().trim(),
        fullname: activeName,
        username: activeUn,
        password: '••••••••',
        createdAt: new Date().toISOString()
      }]);
    } finally {
      setIsLoadingUsers(false);
    }
  }, [isAdmin, activeUserCustomData, profileData.namaGuruMapel]);

  useEffect(() => {
    if (isLoggedIn && isAdmin) {
      fetchAllUsers();
    }
  }, [isLoggedIn, isAdmin, fetchAllUsers]);

  const handleDeleteUserClick = (user: CustomUser) => {
    setUserToDelete(user);
  };

  const handleConfirmDeleteUser = async () => {
    if (!userToDelete) return;
    setIsDeletingUser(true);
    try {
      trackOp('write', 1);
      await deleteDoc(doc(dbDefault, 'custom_accounts', userToDelete.username.toLowerCase().trim()));
      showToast(`User "${userToDelete.fullname}" berhasil dihapus secara permanen.`, 'success');
      setUserToDelete(null);
      fetchAllUsers();
    } catch (err) {
      console.error('Error deleting user:', err);
      showToast('Gagal menghapus user: ' + (err instanceof Error ? err.message : 'Server error'), 'error');
    } finally {
      setIsDeletingUser(false);
    }
  };



  const [classList, setClassList] = useState<string[]>([]);
  const [classListLoaded, setClassListLoaded] = useState(false);

  // Custom states for reset database features
  const [resetModalType, setResetModalType] = useState<'none' | 'new_semester' | 'new_year' | 'everything' | 'clear_all_students'>('none');
  const [resetSuccessModal, setResetSuccessModal] = useState<'none' | 'new_semester' | 'new_year' | 'everything' | 'clear_all_students'>('none');
  const [resetConfirmInput, setResetConfirmInput] = useState('');
  const [resetPasswordInput, setResetPasswordInput] = useState('');
  const [resetPasswordError, setResetPasswordError] = useState('');
  const [showResetPassword, setShowResetPassword] = useState(false);
  const [isVerifyingPassword, setIsVerifyingPassword] = useState(false);
  const [newSemesterChoice, setNewSemesterChoice] = useState<'Ganjil' | 'Genap'>('Genap');
  const [newTahunPelajaran, setNewTahunPelajaran] = useState('');
  const [isResettingData, setIsResettingData] = useState(false);
  const [resetProgress, setResetProgress] = useState(0);
  const [studentSuccessModal, setStudentSuccessModal] = useState<'none' | 'added' | 'edited' | 'deleted'>('none');

  // States for Student Absences Widget in Dashboard
  const [dashboardActiveStatsTab, setDashboardActiveStatsTab] = useState<'class_summary' | 'top_rankings'>('class_summary');
  const [dashboardSelectedClassDetail, setDashboardSelectedClassDetail] = useState<string | null>(null);

  // State for Online/Offline connectivity monitor
  const [isOnline, setIsOnline] = useState<boolean>(navigator.onLine);
  const [firebaseConnected, setFirebaseConnected] = useState<boolean>(true);
  const [syncStatus, setSyncStatus] = useState<'active' | 'syncing' | 'error'>('active');
  const [lastSyncError, setLastSyncError] = useState<string | null>(null);
  const [cloudLastSync, setCloudLastSync] = useState<string | null>(null);
  const [quotaExceeded, setQuotaExceeded] = useState(false);

  // Connectivity Listeners
  useEffect(() => {
    const handleOnline = () => setIsOnline(true);
    const handleOffline = () => setIsOnline(false);

    window.addEventListener('online', handleOnline);
    window.addEventListener('offline', handleOffline);

    return () => {
      window.removeEventListener('online', handleOnline);
      window.removeEventListener('offline', handleOffline);
    };
  }, []);

  // Monitor Firebase Connection State
  useEffect(() => {
    if (!activeDb) return;
    
    // Using a simple interval check or real-time connection status if metadata allows
    // For Firestore, we can listen to metadata changes on a document or check sync state
    // But a simple reliable way is listening to system events and error states in snapshots
    
    const checkConnection = () => {
      // Periodic check or rely on onSnapshot errors which we will add below
    };
    
    const interval = setInterval(checkConnection, 10000);
    return () => clearInterval(interval);
  }, [activeDb]);

  // Session hash pre-fill effect (populates username & password without auto-bypassing login page)
  useEffect(() => {
    const restoreSession = () => {
      if (!window.location.hash) return;
      
      const hashParams = new URL(window.location.href.replace('#', '?')).searchParams;
      let sessionB64 = hashParams.get('session');
      if (!sessionB64) {
        const match = window.location.hash.match(/session=([^&]+)/);
        if (match && match[1]) {
          sessionB64 = match[1];
        }
      }
      if (!sessionB64) return;

      const creds = decodeSession(sessionB64);
      if (creds && creds.username) {
        setAuthEmail(creds.username);
        if (creds.password) {
          setAuthPassword(creds.password);
        }
      }

      try {
        window.history.replaceState(null, '', window.location.pathname + window.location.search);
      } catch {
        window.location.hash = '';
      }
    };

    restoreSession();
  }, []);

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(activeAuth, (user) => {
      if (user) {
        setCurrentUser(user);
        setIsLoggedIn(true);
        safeSetLocalStorage('kaguci_has_logged_in', 'true');
        
        // Load user photo directly from firebase user auth object
        if (user.photoURL) {
          setAvatarUrl(user.photoURL);
        }
        
        // Fetch central account mapping for metadata lookup (configText, fullname)
        if (activeUserCustomData?.username) {
          getDoc(doc(dbDefault, 'custom_accounts', activeUserCustomData.username.toLowerCase().trim())).then(centralDoc => {
            if (centralDoc.exists()) {
              const centralData = centralDoc.data();
              if (centralData.photoURL && !avatarUrl) {
                setAvatarUrl(centralData.photoURL);
              }
            }
          }).catch(err => console.warn('Gagal membaca sinkronisasi metadata dari portal pusat:', err));
        }

        // Async fetch avatar and classList from Firestore without blocking auth state resolution
        getDoc(doc(activeDb, 'users', user.uid)).then(userDoc => {
          if (userDoc.exists()) {
            const data = userDoc.data();
            if (data.photoURL) {
               setAvatarUrl(data.photoURL);
            }
            // Support both flat format and nested 'profileData' format for utmost reliability!
            if (data.profileData) {
               setProfileData((prev: typeof profileData) => ({ ...prev, ...data.profileData }));
            } else if (data.namaGuruMapel || data.namaKepalaSekolah) {
               const flatProfile = {
                 namaGuruMapel: data.namaGuruMapel || '',
                 namaKepalaSekolah: data.namaKepalaSekolah || '',
                 nipGuruMapel: data.nipGuruMapel || '',
                 nipKepalaSekolah: data.nipKepalaSekolah || '',
                 namaKurikulum: data.namaKurikulum || '',
                 nipKurikulum: data.nipKurikulum || '',
                 jabatanKurikulum: data.jabatanKurikulum || 'Wakasek Kurikulum',
                 namaKesiswaan: data.namaKesiswaan || '',
                 nipKesiswaan: data.nipKesiswaan || '',
                 jabatanKesiswaan: data.jabatanKesiswaan || 'Wakasek Kesiswaan',
                 namaGuruWali: data.namaGuruWali || '',
                 nipGuruWali: data.nipGuruWali || '',
                 jabatanGuruWali: data.jabatanGuruWali || 'Guru Wali',
                 namaBK: data.namaBK || '',
                 nipBK: data.nipBK || '',
                 jabatanBK: data.jabatanBK || 'Guru BK',
                 namaHumas: data.namaHumas || '',
                 nipHumas: data.nipHumas || '',
                 jabatanHumas: data.jabatanHumas || 'Wakasek Humas',
                 semester: data.semester || 'Ganjil',
                 tahunPelajaran: data.tahunPelajaran || '',
                 mataPelajaran: data.mataPelajaran || '',
                 role: data.role || 'Guru Mapel',
                 waliKelasClass: data.waliKelasClass || '',
                 jumlahSiswaLakiLaki: data.jumlahSiswaLakiLaki || '',
                 jumlahSiswaPerempuan: data.jumlahSiswaPerempuan || ''
               };
               setProfileData((prev: typeof profileData) => ({ ...prev, ...flatProfile }));
            }

            if (data.classList && Array.isArray(data.classList)) {
              setClassList(data.classList);
            }
          }
          setClassListLoaded(true);
        }).catch(error => {
          console.warn('Failed to fetch user document or profile on auth state change:', error);
          setClassListLoaded(true);
        });
        setIsAuthLoading(false);
      } else {
        setCurrentUser(null);
        setIsLoggedIn(false);
        localStorage.removeItem('kaguci_has_logged_in');
        setAvatarUrl(null);
        setClassList([]);
        setClassListLoaded(false);
        setStudents([]);
        setStudentsLoaded(false);
        setAttendanceSessions([]);
        setIsAuthLoading(false);
      }
    });
    return () => unsubscribe();
  }, [activeAuth, activeDb, activeUserCustomData, avatarUrl]);

  const lastSavedClassListRef = useRef<string | null>(null);

  useEffect(() => {
    if (isLoggedIn && currentUser && classListLoaded) {
      const classListStr = JSON.stringify(classList);
      if (lastSavedClassListRef.current === classListStr) return;
      
      // We removed the auto-sync here to prevent infinite update loops with onSnapshot.
      // Class updates are now handled explicitly in the UI where they are modified.
      lastSavedClassListRef.current = classListStr;
      console.log("classList updated locally, auto-sync to Firebase disabled to prevent loops.");
    }
  }, [classList, classListLoaded, isLoggedIn, currentUser]);

  const [accountDeletedAlert, setAccountDeletedAlert] = useState(false);

  interface UserUsage {
    username: string;
    fullname: string;
    reads: number;
    writes: number;
    date: string;
    time?: string;
    cycleKey?: string;
    isCurrentCycle: boolean;
  }

  const [allUsersUsage, setAllUsersUsage] = useState<UserUsage[]>([]);
  const [isLoadingUsage, setIsLoadingUsage] = useState(false);

  useEffect(() => {
    if (activeTab === 'profile' && isLoggedIn) {
      setIsLoadingUsage(true);
      const currentCycle = getWibCycleInfo().cycleKey;
      getDocs(collection(dbDefault, 'custom_accounts'))
        .then(snap => {
          trackOp('read', snap.size || 1);
          const usages: UserUsage[] = snap.docs.map(doc => {
            const data = doc.data();
            const userCycle = data.dailyUsageCycle;
            // User limits reset automatically every day at 14:00 WIB.
            // If the user's stored cycle is older than current 14:00 WIB cycle, reads and writes are 0.
            const isCurrentCycle = Boolean(userCycle && userCycle === currentCycle);
            return {
              username: doc.id,
              fullname: data.fullname || doc.id,
              reads: isCurrentCycle ? (data.dailyUsageReads || 0) : 0,
              writes: isCurrentCycle ? (data.dailyUsageWrites || 0) : 0,
              date: isCurrentCycle ? (data.dailyUsageDate || '-') : 'Telah Reset (14:00 WIB)',
              time: isCurrentCycle ? (data.dailyUsageTime || '') : '',
              cycleKey: userCycle,
              isCurrentCycle
            };
          });
          // Sort by highest reads first
          usages.sort((a, b) => b.reads - a.reads);
          setAllUsersUsage(usages);
        })
        .catch(err => console.warn("Failed caching usage", err))
        .finally(() => setIsLoadingUsage(false));
    }
  }, [activeTab, isLoggedIn, cycleInfo.cycleKey]);

  // Monitor account deletion for custom generated accounts only
  useEffect(() => {
    if (!isLoggedIn || isAuthLoading) return;
    
    // Only monitor if this is specifically a custom local account created through the system
    let usernameKey = activeUserCustomData?.username;
    const isCustomSystemEmail = currentUser?.email && currentUser.email.endsWith('@kaguci.admin.system.local');
    
    if (!usernameKey && isCustomSystemEmail) {
       usernameKey = currentUser.email!.split('@')[0];
    }
    
    // If not a custom system account, do not attach custom_accounts snapshot
    if (!usernameKey || (!activeUserCustomData && !isCustomSystemEmail)) return;
    
    usernameKey = usernameKey.toLowerCase().trim();
    const centralRef = doc(dbDefault, 'custom_accounts', usernameKey);
    
    // Listen for changes to the user's account document
    const unsubscribe = onSnapshot(centralRef, (docSnap) => {
      if (!docSnap.metadata.fromCache) {
         trackOp('read', 1);
      }
      
      if (!docSnap.exists()) {
         console.warn("User account no longer exists in database! Forcing logout.");
         
         // 1. Cleared all stored credentials
         localStorage.removeItem(`kaguci_profile_${usernameKey}`);
         localStorage.removeItem(`kaguci_avatar_${usernameKey}`);
         sessionStorage.removeItem(`kaguci_welcomed_${usernameKey}`);
         localStorage.removeItem('kaguci_active_custom_user');
         localStorage.removeItem('kaguci_saved_credentials');
         localStorage.removeItem('kaguci_has_logged_in');
         
         try { window.history.replaceState(null, '', window.location.pathname + window.location.search); } 
         catch { window.location.hash = ''; }
         
         setActiveUserCustomData(null);
         setCurrentUser(null);
         setIsLoggedIn(false);
         
         try { signOut(activeAuth); } catch { /* ignore */ }
         
         setAccountDeletedAlert(true);
      }
    }, (error) => {
      console.warn("Account monitor snapshot warning:", error);
    });

    return () => unsubscribe();
  }, [isLoggedIn, isAuthLoading, activeUserCustomData, currentUser, activeAuth]);

  const [students, setStudents] = useState<Student[]>([]);
  const [studentsLoaded, setStudentsLoaded] = useState(false);
  const [attendanceSessions, setAttendanceSessions] = useState<AttendanceSession[]>([]);

  const isResettingDataRef = useRef(false);
  useEffect(() => {
    isResettingDataRef.current = isResettingData;
  }, [isResettingData]);

  useEffect(() => {
    // Wait until auth is fully initialized before deciding if we are logged out
    if (isAuthLoading) return;
    
    // Only clear if NOT logged in after auth has finished loading
    if (!currentUser) {
      console.log("Not logged in, skipping student fetch");
      setStudents([]);
      // Do not set loaded to true here, wait for actual user to be set
      return;
    }

    if (!activeDb) return;
    
    const uid = currentUser.uid;

    if (!activeAuth.currentUser) {
      console.log("No authenticated Firebase user, skipping firestore subscriptions");
      setStudents([]);
      setStudentsLoaded(true);
      setAttendanceSessions([]);
      setClassListLoaded(true);
      return () => {};
    }
    const activeProjId = activeDb?.app?.options?.projectId;
    const defaultProjId = dbDefault?.app?.options?.projectId;
    const isDefaultDb = (activeDb === dbDefault) || 
      (!!activeProjId && !!defaultProjId && activeProjId === defaultProjId) ||
      (!activeUserCustomData || !activeUserCustomData.configText);
    console.log("Fetching students. isDefaultDb:", isDefaultDb, "userId:", uid, "activeDb Instance:", activeDb);

    const qStudents = query(collection(activeDb, 'students'), where('userId', '==', uid));
    
    const unsubscribeStudents = onSnapshot(
      qStudents, 
      (snapshot) => {
        setFirebaseConnected(true);
        setSyncStatus('active');
        setCloudLastSync(new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }));
        const readCost = snapshot.metadata.fromCache ? 0 : snapshot.docChanges().length;
        if (readCost > 0) trackOp('read', readCost);
        console.log("Firestore snapshots for students received:", snapshot.size, "Cost:", readCost);
        if (snapshot.size === 0) {
            console.warn("No students found. Checking collection...");
            const queryRef = query(collection(activeDb, 'students'), where('userId', '==', uid));
            getDocs(queryRef).then(allDocs => {
                console.log("Total students in DB for user:", allDocs.size);
                allDocs.docs.forEach(d => console.log("Doc id:", d.id, "data:", d.data()));
            });
        }
        const fetchedStudents = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as Student));
        fetchedStudents.sort((a, b) => (a.name || '').toLowerCase().localeCompare((b.name || '').toLowerCase(), 'id-ID'));
        setStudents(fetchedStudents);
        setStudentsLoaded(true);

        // Auto-merge classList from fetched students cautiously
        const extractedClasses = Array.from(new Set(fetchedStudents.map(s => String(s.class || '').trim()).filter(Boolean)));
        setClassList(prev => {
           // If we are currently resetting, don't perform auto-merging that could restore deleted classes
           // Using ref to avoid re-subscribing Effect on every resetting state change
           if (isResettingDataRef.current) return [];
           
           // If fetched students is empty, it means either we haven't loaded them or they were deleted.
           // In this case, we prefer to trust the User Doc specifically.
           if (fetchedStudents.length === 0 && studentsLoaded) {
             // Let the user doc snapshot handle it if it arrives, 
             // but if we know students are gone, we don't automatically wipe classes here 
             // because some might have been manually added.
             return prev; 
           }

           const merged = Array.from(new Set([...prev, ...extractedClasses]));
           merged.sort(compareClasses);
           if (JSON.stringify(prev) === JSON.stringify(merged)) return prev;
           
           // Persist newly discovered classes to Cloud if we are not in the middle of a reset
           if (uid && !isResettingDataRef.current) {
             setDoc(doc(activeDb, 'users', uid), { classList: merged, updatedAt: new Date().toISOString() }, { merge: true })
               .catch(err => console.warn("Failed to auto-sync merged classList to cloud:", err));
           }
           
           return merged;
         });
      },
      (error) => {
        console.error("Error fetching students:", error);
        setStudentsLoaded(true);
        setFirebaseConnected(false);
        setSyncStatus('error');
        setLastSyncError(`Gagal menarik data siswa: ${error.message}`);
        
        if (currentUser) {
          try {
            handleFirestoreError(error, OperationType.LIST, 'students');
          } catch(err) {
             const msg = err instanceof Error ? err.message : String(err);
             if (msg.toLowerCase().includes('permission')) {
               showToast('Firestore Rules Error! Data siswa gagal ditarik dari Cloud. Pastikan rules Firebase database mandiri Anda sudah "allow read, write: if request.auth != null;".', 'error');
             }
          }
        }
      }
    );

    const unsubscribeUser = onSnapshot(doc(activeDb, 'users', uid), (docSnapshot) => {
      setFirebaseConnected(true);
      setSyncStatus('active');
      setCloudLastSync(new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }));
      const readCost = docSnapshot.metadata.fromCache ? 0 : 1;
      if (readCost > 0) trackOp('read', readCost);
      console.log("Real-time update received for user:", uid, "Cost:", readCost);
      if (docSnapshot.exists()) {
        const data = docSnapshot.data();
        if (data.photoURL) {
          setAvatarUrl(data.photoURL);
        }
        
        // Sync Profile Data real-time
        if (data.profileData) {
           setProfileData((prev: typeof profileData) => ({ ...prev, ...data.profileData }));
        } else if (data.namaGuruMapel || data.namaKepalaSekolah) {
           const flatProfile = {
             namaGuruMapel: data.namaGuruMapel || '',
             namaKepalaSekolah: data.namaKepalaSekolah || '',
             nipGuruMapel: data.nipGuruMapel || '',
             nipKepalaSekolah: data.nipKepalaSekolah || '',
             namaKurikulum: data.namaKurikulum || '',
             nipKurikulum: data.nipKurikulum || '',
             jabatanKurikulum: data.jabatanKurikulum || 'Wakasek Kurikulum',
             namaKesiswaan: data.namaKesiswaan || '',
             nipKesiswaan: data.nipKesiswaan || '',
             jabatanKesiswaan: data.jabatanKesiswaan || 'Wakasek Kesiswaan',
             namaGuruWali: data.namaGuruWali || '',
             nipGuruWali: data.nipGuruWali || '',
             jabatanGuruWali: data.jabatanGuruWali || 'Guru Wali',
             namaBK: data.namaBK || '',
             nipBK: data.nipBK || '',
             jabatanBK: data.jabatanBK || 'Guru BK',
             namaHumas: data.namaHumas || '',
             nipHumas: data.nipHumas || '',
             jabatanHumas: data.jabatanHumas || 'Wakasek Humas',
             semester: data.semester || 'Ganjil',
             tahunPelajaran: data.tahunPelajaran || '',
             mataPelajaran: data.mataPelajaran || '',
             role: data.role || 'Guru Mapel',
             waliKelasClass: data.waliKelasClass || ''
           };
           setProfileData((prev: typeof profileData) => ({ ...prev, ...flatProfile }));
        }

        if (data.classList && Array.isArray(data.classList)) {
          const newListStr = JSON.stringify(data.classList);
          setClassList(prev => {
            if (JSON.stringify(prev) === newListStr) return prev;
            return data.classList;
          });
        }
      } else {
        setAvatarUrl(null);
        localStorage.removeItem(`kaguci_avatar_${uid}`);
      }
    }, (error) => {
      console.error("Error fetching user doc:", error);
      setFirebaseConnected(false);
      setSyncStatus('error');
      setLastSyncError(`Gagal sinkronisasi profil: ${error.message}`);
    });

    const qSessions = query(collection(activeDb, 'attendanceSessions'), where('userId', '==', uid));
    const unsubscribeSessions = onSnapshot(
      qSessions, 
      (snapshot) => {
        setFirebaseConnected(true);
        setSyncStatus('active');
        setCloudLastSync(new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }));
        const readCost = snapshot.metadata.fromCache ? 0 : snapshot.docChanges().length;
        if (readCost > 0) trackOp('read', readCost);
        console.log("Real-time update received for sessions:", snapshot.size, "docs. Cost:", readCost);
        const fetchedSessions = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as AttendanceSession));
        setAttendanceSessions(fetchedSessions);
      },
      (error) => {
        console.error("Error fetching sessions:", error);
        setFirebaseConnected(false);
        setSyncStatus('error');
        setLastSyncError(`Gagal menarik data absensi: ${error.message}`);
        
        if (currentUser) {
          try {
            handleFirestoreError(error, OperationType.LIST, 'attendanceSessions');
          } catch(err) {
             const msg = err instanceof Error ? err.message : String(err);
             if (msg.toLowerCase().includes('permission')) {
               showToast('Firestore Rules Error! Data sesi gagal ditarik dari Cloud. Pastikan rules Firebase database mandiri Anda sudah "allow read, write: if request.auth != null;".', 'error');
             }
          }
        }
      }
    );

    return () => {
      unsubscribeStudents();
      unsubscribeUser();
      unsubscribeSessions();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isLoggedIn, currentUser, activeDb, isAuthLoading, activeUserCustomData, activeAuth]);

  const [newStudent, setNewStudent] = useState({ name: '', nisn: '', class: '' });
  const [editingStudentId, setEditingStudentId] = useState<string | null>(null);
  const [confirmationAction, setConfirmationAction] = useState<{ type: 'edit' | 'delete', student: Student } | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [isUploadingPhoto, setIsUploadingPhoto] = useState(false);
  const [importResult, setImportResult] = useState<{
    isOpen: boolean;
    successCount: number;
    skipCount: number;
    failCount: number;
    emptyCount: number;
    totalParsed: number;
    error?: boolean;
    errorMessage?: string;
    details?: string[];
    sheetsProcessed?: { name: string; count: number }[];
  } | null>(null);

  // Jika role pengguna adalah 'Wali Kelas', filter agar hanya kelas binaannya yang aktif
  const isWaliKelas = profileData.role === 'Wali Kelas';
  const effectiveClassList = useMemo(() => {
    if (isWaliKelas && profileData.waliKelasClass) {
      return [profileData.waliKelasClass];
    }
    return classList;
  }, [isWaliKelas, profileData.waliKelasClass, classList]);

  const effectiveStudents = useMemo(() => {
    let list: Student[] = [];
    if (isWaliKelas && profileData.waliKelasClass) {
      list = students.filter(s => s.class === profileData.waliKelasClass);
    } else {
      list = [...students];
    }
    // Urutkan berdasarkan kelas (X1, X2, ... XI 8, XI 9, ... XII) lalu nama siswa
    return list.sort(compareStudentsByClass);
  }, [isWaliKelas, profileData.waliKelasClass, students]);

  useEffect(() => {
    if (isWaliKelas && profileData.waliKelasClass && !newStudent.class) {
      setNewStudent(prev => ({ ...prev, class: profileData.waliKelasClass }));
    }
  }, [isWaliKelas, profileData.waliKelasClass, newStudent.class]);

  const attendanceStats = useMemo(() => {
    const targetStudents = isWaliKelas && profileData.waliKelasClass ? effectiveStudents : students;
    if (attendanceSessions.length === 0) return { rate: '100%', attentionCount: 0 };
    
    // Set ID siswa aktif terdaftar
    const validStudentIds = new Set(targetStudents.map(s => s.id));
    
    let totalRecords = 0;
    let totalHadir = 0;
    const attentionSet = new Set<string>();
    
    attendanceSessions.forEach(session => {
      if (session.records) {
        Object.entries(session.records).forEach(([studentId, status]) => {
          if (status) {
            // Hanya hitung rekam absensi dari siswa aktif terdaftar
            if (validStudentIds.size === 0 || validStudentIds.has(studentId)) {
              totalRecords++;
              if (status === 'Hadir' || status === 'Dispen') {
                totalHadir++;
              }
              if (status === 'Alpa' || status === 'Sakit' || status === 'Izin') {
                attentionSet.add(studentId);
              }
            }
          }
        });
      }
    });
    
    const ratePercentage = totalRecords > 0 
      ? Math.round((totalHadir / totalRecords) * 100) 
      : 100;
      
    return {
      rate: `${ratePercentage}%`,
      attentionCount: attentionSet.size
    };
  }, [attendanceSessions, students, isWaliKelas, profileData.waliKelasClass, effectiveStudents]);

  const studentAbsenceStats = useMemo(() => {
    // Map student ID to their attendance counts
    const counts: Record<string, { alpa: number; sakit: number; izin: number; totalNonHadir: number; alpaDates: string[]; sakitDates: string[]; izinDates: string[] }> = {};
    
    // Initialize for all students
    students.forEach(s => {
      counts[s.id] = { alpa: 0, sakit: 0, izin: 0, totalNonHadir: 0, alpaDates: [], sakitDates: [], izinDates: [] };
    });
    
    // Process sessions
    attendanceSessions.forEach(session => {
      if (session.records) {
        Object.entries(session.records).forEach(([studentId, status]) => {
          if (!counts[studentId]) {
            counts[studentId] = { alpa: 0, sakit: 0, izin: 0, totalNonHadir: 0, alpaDates: [], sakitDates: [], izinDates: [] };
          }
          if (status === 'Alpa') {
            counts[studentId].alpa++;
            counts[studentId].alpaDates.push(session.date);
            counts[studentId].totalNonHadir++;
          } else if (status === 'Sakit') {
            counts[studentId].sakit++;
            counts[studentId].sakitDates.push(session.date);
            counts[studentId].totalNonHadir++;
          } else if (status === 'Izin') {
            counts[studentId].izin++;
            counts[studentId].izinDates.push(session.date);
            counts[studentId].totalNonHadir++;
          }
        });
      }
    });

    interface StudentAbsenceDetail {
      id: string; 
      name: string; 
      class: string; 
      alpa: number; 
      sakit: number; 
      izin: number; 
      totalNonHadir: number;
      alpaDates: string[];
      sakitDates: string[];
      izinDates: string[];
    }

    // Group students by class
    const classGroups: Record<string, StudentAbsenceDetail[]> = {};

    // Initialize arrays for each class
    effectiveClassList.forEach(cls => {
      classGroups[cls] = [];
    });

    // Distribute students to classes and attach counts
    students.forEach(s => {
      const cls = s.class || 'Tanpa Kelas';
      if (classGroups[cls]) {
        const c = counts[s.id] || { alpa: 0, sakit: 0, izin: 0, totalNonHadir: 0, alpaDates: [], sakitDates: [], izinDates: [] };
        classGroups[cls].push({
          id: s.id,
          name: s.name,
          class: cls,
          ...c
        });
      }
    });

    // For each class, find student with max Alpa and max Sakit+Izin
    const result: Array<{
      className: string;
      totalAlpaStudents: number;
      totalSakitStudents: number;
      totalIzinStudents: number;
      allAbsenceList: Array<{
        id: string;
        name: string;
        alpa: number;
        sakit: number;
        izin: number;
        total: number;
        alpaDates: string[];
        sakitDates: string[];
        izinDates: string[];
      }>;
    }> = [];

    (Object.entries(classGroups) as Array<[string, StudentAbsenceDetail[]]>).forEach(([className, list]) => {


      const absentStudents = list
        .filter(item => item.totalNonHadir > 0)
        .map(item => ({
          id: item.id,
          name: item.name,
          alpa: item.alpa,
          sakit: item.sakit,
          izin: item.izin,
          total: item.totalNonHadir,
          alpaDates: item.alpaDates,
          sakitDates: item.sakitDates,
          izinDates: item.izinDates
        }))
        .sort((a, b) => b.total - a.total);

      const totalAlpaStudents = list.filter(item => item.alpa > 0).length;
      const totalSakitStudents = list.filter(item => item.sakit > 0).length;
      const totalIzinStudents = list.filter(item => item.izin > 0).length;

      result.push({
        className,
        totalAlpaStudents,
        totalSakitStudents,
        totalIzinStudents,
        allAbsenceList: absentStudents
      });
    });

    result.sort((a, b) => compareClasses(a.className, b.className));

    return result;
  }, [students, attendanceSessions, effectiveClassList]);

  const [toast, setToast] = useState<{ message: string, type: 'success' | 'info' | 'error' } | null>(null);
  const toastTimeoutRef = useRef<NodeJS.Timeout | null>(null);
  const [authEmail, setAuthEmail] = useState(() => {
    try {
      const creds = localStorage.getItem('kaguci_saved_credentials');
      if (creds) return JSON.parse(creds).email || '';
    } catch {
      /* ignore storage parse error */
    }
    return '';
  });
  const [authPassword, setAuthPassword] = useState(() => {
    try {
      const creds = localStorage.getItem('kaguci_saved_credentials');
      if (creds) return JSON.parse(creds).password || '';
    } catch {
      /* ignore storage parse error */
    }
    return '';
  });
  const [isForgotPassword, setIsForgotPassword] = useState(false);
  const [showPassword, setShowPassword] = useState(false);

  // States for new school registration modal (Kotak Dialog)
  const [isRegisterModalOpen, setIsRegisterModalOpen] = useState(false);
  const [regFullName, setRegFullName] = useState('');
  const [regUsername, setRegUsername] = useState('');
  const [regPassword, setRegPassword] = useState('');
  const [regToken, setRegToken] = useState('');
  const [generatedToken, setGeneratedToken] = useState('');

  // States for registration result dialog status
  const [registrationResult, setRegistrationResult] = useState<{
    success: boolean;
    title: string;
    message: string;
    fullname?: string;
    username?: string;
    projectId?: string;
    showBypassButton?: boolean;
    configToSave?: {
      fullname: string;
      username: string;
      password?: string;
      configText: string;
    };
  } | null>(null);

  // States for login error dialog (modern message box for unregistered accounts / wrong password)
  const [loginError, setLoginError] = useState<{
    title: string;
    message: string;
    recommendations: string[];
    username?: string;
  } | null>(null);

  // States for password/username recovery
  const [recoverySearchVal, setRecoverySearchVal] = useState('');
  const [recoverySearchType, setRecoverySearchType] = useState<'username' | 'password'>('username');
  const [recoveryResult, setRecoveryResult] = useState<DocumentData[] | null>(null);
  const [isRecoveryLoading, setIsRecoveryLoading] = useState(false);

  const showToast = (message: string, type: 'success' | 'info' | 'error') => {
    if (toastTimeoutRef.current) clearTimeout(toastTimeoutRef.current);
    setToast({ message, type });
    toastTimeoutRef.current = setTimeout(() => setToast(null), 3000);
  };

  useEffect(() => {
    const handleOnline = () => {
      setIsOnline(true);
      showToast('Koneksi internet terhubung kembali! Data disinkronkan ke Cloud.', 'success');
    };
    const handleOffline = () => {
      setIsOnline(false);
      showToast('Koneksi terputus. Menggunakan data lokal (Offline Mode).', 'error');
    };
    const handleQuota = () => {
      setQuotaExceeded(true);
      setShowQuotaModal(true);
    };

    window.addEventListener('online', handleOnline);
    window.addEventListener('offline', handleOffline);
    window.addEventListener('firestore-quota-exceeded', handleQuota);

    return () => {
      window.removeEventListener('online', handleOnline);
      window.removeEventListener('offline', handleOffline);
      window.removeEventListener('firestore-quota-exceeded', handleQuota);
    };
  }, []);

  const handleForceRegister = async (configData: { fullname: string; username: string; password?: string; configText: string }) => {
    setIsAuthLoading(true);
    setRegistrationResult(null);
    try {
      const resolvedUsername = configData.username.toLowerCase().trim();

      // 1. Create native account inside the main Firebase Cloud Database
      try {
        await createUserWithEmailAndPassword(auth, toAuthEmail(resolvedUsername), configData.password || '');
      } catch (authError) {
        const errObj = authError as { message?: string; code?: string };
        if (errObj.message?.includes('email-already-in-use') || errObj.code === 'auth/email-already-in-use') {
          try {
            await signInWithEmailAndPassword(auth, toAuthEmail(resolvedUsername), configData.password || '');
          } catch {
            throw authError;
          }
        } else {
          throw authError;
        }
      }

      // 2. Map directory entry centrally to support login routing & forgot password lookups
      await setDoc(doc(dbDefault, 'custom_accounts', resolvedUsername), {
        fullname: configData.fullname,
        username: resolvedUsername,
        password: configData.password,
        configText: configData.configText,
        createdAt: new Date().toISOString()
      });

      // On success: trigger results popup
      setRegistrationResult({
        success: true,
        title: 'Pendaftaran Akun Berhasil!',
        message: 'Akun Administrator baru Anda siap digunakan dengan Cloud Database default kaguci secara instan.',
        fullname: configData.fullname,
        username: resolvedUsername,
        projectId: 'Database Pusat Default kaguci'
      });

      // Autofill login credentials for easy access
      setAuthEmail(resolvedUsername);
      if (configData.password) {
        setAuthPassword(configData.password);
      }
      
      // Clean fields and close form
      setIsRegisterModalOpen(false);
      setRegFullName('');
      setRegUsername('');
      setRegPassword('');
    } catch (error) {
      const err = error as Error;
      let IndonesianError = err.message;
      if (err.message.includes('email-already-in-use')) {
        const lookupUsername = configData.username.toLowerCase().trim();
        let centralUserFound: { fullname?: string; username?: string; configText?: string } | null = null;
        try {
          const lookupDoc = await getDoc(doc(dbDefault, 'custom_accounts', lookupUsername));
          if (lookupDoc.exists()) {
            const data = lookupDoc.data();
            centralUserFound = {
              fullname: data.fullname,
              username: data.username,
              configText: data.configText
            };
          }
        } catch (lookupErr) {
          console.error('Error lookup di catch force:', lookupErr);
        }

        if (centralUserFound) {
          IndonesianError = `Maaf, Username "${lookupUsername}" sudah terdaftar di database sistem pusat.\n\n• Nama Pengguna: ${centralUserFound.fullname || 'N/A'}\n• Akun (Username): ${centralUserFound.username || 'N/A'}\n\nSilakan gunakan menu "Masuk" (Login) dan gunakan akun tersebut beserta kata sandinya untuk login.`;
        } else {
          IndonesianError = `Username "${lookupUsername}" sudah pernah didaftarkan pada database project Firebase Anda, namun kata sandi yang Anda ketik salah.\n\nLangkah Solusi:\n1. Masukkan kata sandi yang tepat jika Anda adalah pemilik akun tersebut.\n2. ATAU, buka Firebase Console Anda -> menu Authentication -> hapus akun "${lookupUsername}@kaguci.com", setelah itu coba daftarkan kembali.\n3. ATAU, silakan daftar dengan memakai Username yang berbeda.`;
        }
      } else if (err.message.includes('weak-password')) {
        IndonesianError = 'Kata sandi minimal berisi 6 karakter.';
      } else if (err.message.includes('invalid-api-key') || err.message.includes('API key')) {
        IndonesianError = 'API Key yang terdapat pada konfigurasi Web Firebase Anda salah atau tidak valid.';
      } else if (err.message.includes('network-request-failed')) {
        IndonesianError = 'Koneksi jaringan gagal atau domain AuthDomain tidak terdaftar di Firebase Anda.';
      } else if (err.message.includes('operation-not-allowed')) {
        IndonesianError = 'Provider "Email/Password" belum aktif di Firebase Console Anda!';
      } else if (err.message.includes('configuration-not-found') || err.message.includes('auth/configuration-not-found')) {
        IndonesianError = 'Layanan Autentikasi belum diinisialisasi di Proyek Firebase Anda! Silakan masuk ke Firebase Console -> klik menu "Authentication" di sebelah kiri -> lalu klik tombol "Get Started" (Mulai) untuk mengaktifkannya.';
      }

      setRegistrationResult({
        success: false,
        title: 'Pendaftaran Gagal!',
        message: `${IndonesianError}\n\n(Detail Teknis: ${err.message})`
      });
    } finally {
      setIsAuthLoading(false);
    }
  };

  const handleEditStudent = (student: Student) => {
    setNewStudent({ name: student.name, nisn: student.nisn, class: student.class });
    setEditingStudentId(student.id);
    showToast('Mode edit aktif untuk ' + student.name, 'info');
    
    // Scroll to form if needed
    const formElement = document.getElementById('student-input-container');
    if (formElement) {
      formElement.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  };

  const [deletingStudentId, setDeletingStudentId] = useState<string | null>(null);

  const confirmAction = async () => {
    if (!confirmationAction) return;
    
    if (confirmationAction.type === 'delete') {
      const studentToDelete = confirmationAction.student;
      
      // OPTIMISTIC UI: Close modals and show success feedback immediately
      setConfirmationAction(null);
      setStudentSuccessModal('deleted');
      setDeletingStudentId(null); // Clear deleting state to avoid visual lag
      
      try {
        trackOp('write', 1);
        // Perform deletion in background
        await deleteDoc(doc(activeDb, 'students', studentToDelete.id));
        showToast(`Data ${studentToDelete.name} berhasil dihapus dari cloud.`, 'success');
      } catch (err) {
        console.error("Deletion failed:", err);
        showToast('Sinkronisasi hapus gagal. Cek koneksi internet Anda.', 'error');
      }
    }
  };

  const handleResetData = async (type: 'new_semester' | 'new_year' | 'everything' | 'clear_all_students') => {
    if (!currentUser || !activeDb) return;
    setIsResettingData(true);
    setResetProgress(0);
    
    // Start an interval to animate the progress smoothly and extremely fast
    let currentSimulatedProgress = 0;
    const progressInterval = setInterval(() => {
      // Fast, human-like fluid steps
      currentSimulatedProgress += Math.random() > 0.4 ? 6 : 4;
      if (currentSimulatedProgress >= 93) {
        clearInterval(progressInterval);
        currentSimulatedProgress = 93; // hold near 93% until parallel DB writes complete
      }
      setResetProgress(currentSimulatedProgress);
    }, 25);

    const finishUI = async (successType: 'new_semester' | 'new_year' | 'everything' | 'clear_all_students') => {
      clearInterval(progressInterval);
      setResetProgress(100);
      await new Promise(r => setTimeout(r, 450)); // Allow user to see 100%
      setResetModalType('none');
      setResetSuccessModal(successType);
      setResetConfirmInput('');
      setResetPasswordInput('');
      setResetPasswordError('');
      setResetProgress(0);
      setIsResettingData(false);
    };

    try {
      const promises: Promise<void>[] = [];
      const SAFETY_TIMEOUT_MS = 15000; // 15s absolute limit for UI wait

      if (type === 'new_semester' || type === 'new_year') {
        const sessionsToDelete = [...attendanceSessions];
        if (sessionsToDelete.length > 0) {
          let batch = writeBatch(activeDb);
          let count = 0;
          for (let i = 0; i < sessionsToDelete.length; i++) {
            batch.delete(doc(activeDb, 'attendanceSessions', sessionsToDelete[i].id));
            count++;
            if (count === 400 || i === sessionsToDelete.length - 1) {
              trackOp('write', count);
              promises.push(batch.commit());
              batch = writeBatch(activeDb);
              count = 0;
            }
          }
        }

        const updatedProfile = { 
          ...profileData, 
          semester: type === 'new_semester' ? newSemesterChoice : profileData.semester,
          ...(type === 'new_semester' && newTahunPelajaran ? { tahunPelajaran: newTahunPelajaran } : {})
        };
        setProfileData(updatedProfile);

        if (activeAuth.currentUser) {
          trackOp('write', 1);
          promises.push(setDoc(doc(activeDb, 'users', activeAuth.currentUser.uid), {
            profileData: updatedProfile,
            semester: updatedProfile.semester,
            ...(updatedProfile.tahunPelajaran ? { tahunPelajaran: updatedProfile.tahunPelajaran } : {})
          }, { merge: true }));
        }
        
        if (activeUserCustomData?.username) {
            trackOp('write', 1);
            promises.push(setDoc(doc(dbDefault, 'custom_accounts', activeUserCustomData.username.toLowerCase().trim()), { 
              attendanceSessions: [],
              profileData: updatedProfile,
              semester: updatedProfile.semester,
              ...(updatedProfile.tahunPelajaran ? { tahunPelajaran: updatedProfile.tahunPelajaran } : {})
            }, { merge: true }));
            localStorage.setItem(`kaguci_profile_${activeUserCustomData.username.toLowerCase()}`, JSON.stringify(updatedProfile));
        }

        if (promises.length > 0) {
          try {
            await Promise.race([
              Promise.all(promises),
              new Promise((_, reject) => setTimeout(() => reject(new Error('Background sync.')), SAFETY_TIMEOUT_MS))
            ]);
          } catch (dbErr) {
            handleFirestoreError(dbErr, OperationType.DELETE, type === 'new_semester' ? 'reset_new_semester' : 'reset_new_year');
          }
        }
        
        setAttendanceSessions([]);
        localStorage.removeItem(`kaguci_sessions_${currentUser.uid}`);
        showToast(type === 'new_semester' ? `Semester Baru (${newSemesterChoice}) berhasil disiapkan.` : 'Pembersihan Tahun Ajaran baru telah disiapkan.', 'success');
        await finishUI(type);

      } else if (type === 'everything') {
        const sessionsToDelete = [...attendanceSessions];
        const studentsToDelete = [...students];
        const allDeletes = [
          ...sessionsToDelete.map(s => doc(activeDb, 'attendanceSessions', s.id)),
          ...studentsToDelete.map(s => doc(activeDb, 'students', s.id))
        ];
        
        if (allDeletes.length > 0) {
          let batch = writeBatch(activeDb);
          let count = 0;
          for (let i = 0; i < allDeletes.length; i++) {
            batch.delete(allDeletes[i]);
            count++;
            if (count === 400 || i === allDeletes.length - 1) {
              trackOp('write', count);
              promises.push(batch.commit());
              batch = writeBatch(activeDb);
              count = 0;
            }
          }
        }
        
        trackOp('write', 1);
        promises.push(setDoc(doc(activeDb, 'users', currentUser.uid), { 
          classList: [], 
          photoURL: deleteField(),
          profileData: {
            namaGuruMapel: '', nipGuruMapel: '', mataPelajaran: '',
            namaKepalaSekolah: '', nipKepalaSekolah: '',
            semester: 'Ganjil', tahunPelajaran: ''
          }
        }, { merge: true }));

        if (activeUserCustomData?.username) {
            trackOp('write', 1);
            promises.push(setDoc(doc(dbDefault, 'custom_accounts', activeUserCustomData.username.toLowerCase().trim()), { 
              classList: [], students: [], attendanceSessions: [],
              photoURL: deleteField(),
              profileData: {
                namaGuruMapel: '', nipGuruMapel: '', mataPelajaran: '',
                namaKepalaSekolah: '', nipKepalaSekolah: '',
                semester: 'Ganjil', tahunPelajaran: ''
              }
            }, { merge: true }));
        }
        
        try {
          await Promise.race([
            Promise.all(promises),
            new Promise((_, reject) => setTimeout(() => reject(new Error('Sync background.')), SAFETY_TIMEOUT_MS))
          ]);
        } catch (dbErr) {
          handleFirestoreError(dbErr, OperationType.DELETE, 'reset_everything');
        }

        setClassList([]);
        setStudents([]);
        setAttendanceSessions([]);
        setAvatarUrl(null);
        setProfileData({
          namaGuruMapel: '', nipGuruMapel: '', mataPelajaran: '',
          namaKepalaSekolah: '', nipKepalaSekolah: '',
          semester: 'Ganjil', tahunPelajaran: '',
          role: 'Guru Mapel', waliKelasClass: ''
        });
        
        localStorage.removeItem(`kaguci_students_${currentUser.uid}`);
        localStorage.removeItem(`kaguci_classList_${currentUser.uid}`);
        localStorage.removeItem(`kaguci_avatar_${currentUser.uid}`);
        localStorage.removeItem(`kaguci_sessions_${currentUser.uid}`);

        showToast('Seluruh data berhasil dihapus secara permanen.', 'success');
        await finishUI('everything');

      } else if (type === 'clear_all_students') {
        const studentsToDelete = [...students];
        if (studentsToDelete.length > 0) {
          let batch = writeBatch(activeDb);
          let count = 0;
          for (let i = 0; i < studentsToDelete.length; i++) {
            batch.delete(doc(activeDb, 'students', studentsToDelete[i].id));
            count++;
            if (count === 400 || i === studentsToDelete.length - 1) {
              trackOp('write', count);
              promises.push(batch.commit());
              batch = writeBatch(activeDb);
              count = 0;
            }
          }
        }
        
        if (activeUserCustomData?.username) {
            trackOp('write', 1);
            promises.push(setDoc(doc(dbDefault, 'custom_accounts', activeUserCustomData.username.toLowerCase().trim()), { 
              students: [] 
            }, { merge: true }));
        }
        
        try {
          await Promise.race([
            Promise.all(promises),
            new Promise((_, reject) => setTimeout(() => reject(new Error('Sync background.')), SAFETY_TIMEOUT_MS))
          ]);
        } catch (dbErr) {
          handleFirestoreError(dbErr, OperationType.DELETE, 'clear_students');
        }
        
        setStudents([]);
        localStorage.removeItem(`kaguci_students_${currentUser.uid}`);
        showToast('Data siswa berhasil dikosongkan.', 'success');
        await finishUI('clear_all_students');
      }
    } catch (err) {
      clearInterval(progressInterval);
      console.error("Error resetting data:", err);
      showToast('Sinkronisasi cloud sedang sibuk. Mohon tunggu sejenak.', 'info');
      setIsResettingData(false);
      setResetProgress(0);
    }
  };

  const addOrUpdateStudent = async () => {
    const studentClass = (isWaliKelas && profileData.waliKelasClass) ? profileData.waliKelasClass : newStudent.class;
    if (!newStudent.name || !newStudent.nisn || !studentClass) {
      showToast('Mohon lengkapi semua data siswa (Nama, NIS, Kelas).', 'error');
      return;
    }
    if (!activeAuth.currentUser) {
      showToast('Sesi berakhir. Silakan masuk kembali.', 'error');
      return;
    }
    
    // Cek apakah NIS sudah terdaftar (double)
    const isDuplicateNisn = students.some(s => s.nisn === newStudent.nisn && s.id !== editingStudentId);
    if (isDuplicateNisn) {
      showToast('NIS sudah terdaftar.', 'error');
      return;
    }
    
    try {
      if (editingStudentId) {
        const studentData = { ...newStudent, class: studentClass, id: editingStudentId, userId: activeAuth.currentUser.uid };
        // Reset form immediately for fast feel
        setEditingStudentId(null);
        setNewStudent({ name: '', nisn: '', class: (isWaliKelas && profileData.waliKelasClass) ? profileData.waliKelasClass : '' });
        setStudentSuccessModal('edited');
        
        await setDoc(doc(activeDb, 'students', editingStudentId), studentData);
        showToast('Data siswa berhasil diperbarui.', 'success');
      } else {
        const newId = Date.now().toString();
        const studentData = { ...newStudent, class: studentClass, id: newId, userId: activeAuth.currentUser.uid };
        // Reset form immediately
        setNewStudent(prev => ({ ...prev, name: '', nisn: '', class: (isWaliKelas && profileData.waliKelasClass) ? profileData.waliKelasClass : '' }));
        setStudentSuccessModal('added');
        
        await setDoc(doc(activeDb, 'students', newId), studentData);
        showToast('Siswa Berhasil Ditambahkan', 'success');
      }
    } catch (err) {
      const errorObj = err as Error;
      showToast('Gagal menyimpan: ' + (errorObj.message || 'Server error'), 'error');
      handleFirestoreError(err, OperationType.WRITE, 'students');
    }
  };

  const excelInputRef = useRef<HTMLInputElement>(null);

  const handleImportExcel = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (!activeAuth.currentUser) {
      showToast('Sesi berakhir. Silakan masuk kembali.', 'error');
      if (excelInputRef.current) excelInputRef.current.value = '';
      return;
    }

    const { uid } = activeAuth.currentUser;
    await importExcelHelper(
      file,
      activeAuth,
      activeDb,
      classList,
      setClassList,
      setImportResult,
      showToast,
      excelInputRef,
      students
    );
    return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const bstr = evt.target?.result as ArrayBuffer;
        const wb = XLSX.read(new Uint8Array(bstr), { type: 'array' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

        if (data.length < 2) {
           showToast('File Excel kosong atau tidak memiliki data.', 'error');
           if (excelInputRef.current) excelInputRef.current.value = '';
           return;
         }

        const headers = data[0] as string[];
        const rows = data.slice(1) as (string | number | boolean | null | undefined)[][];
        showToast('Mulai mengimpor data...', 'info');
        const classIdx = headers.findIndex(h => h?.toString().toLowerCase().includes('kelas'));
        const nameIdx = headers.findIndex(h => h?.toString().toLowerCase().includes('nama'));
        
        // Sort rows by classIdx then by nameIdx
        if (classIdx !== -1 && nameIdx !== -1) {
            rows.sort((a, b) => {
                const classA = String(a[classIdx] || '').toLowerCase();
                const classB = String(b[classIdx] || '').toLowerCase();
                if (classA !== classB) return classA.localeCompare(classB);
                
                const nameA = String(a[nameIdx] || '').toLowerCase();
                const nameB = String(b[nameIdx] || '').toLowerCase();
                return nameA.localeCompare(nameB);
            });
        }
        
        const nisnIdx = headers.findIndex(h => h?.toString().toLowerCase().includes('nis'));

        if (nameIdx === -1 || classIdx === -1) {
            showToast('Format kolom tidak sesuai! Pastikan terdapat kolom: Nama dan Kelas.', 'error');
            if (excelInputRef.current) excelInputRef.current.value = '';
            return;
        }

        let successCount = 0;
        let skipCount = 0;
        if (!activeDb) {
            console.error("activeDb is not initialized!");
            showToast('Database belum siap, silakan coba lagi.', 'error');
            return;
        }
        let batch = writeBatch(activeDb);
        let batchCounter = 0;

        for (const [i, row] of rows.entries()) {
          const rawName = row[nameIdx];
          const rawNisn = row[nisnIdx];
          const rawClass = row[classIdx];

          if (rawName && rawClass) {
            const cleanNisn = rawNisn ? String(rawNisn).trim().replace(/\D/g, '') : '';

            const newId = Date.now().toString() + '-' + i;
            const student = {
              id: newId,
              name: String(rawName).trim(),
              nisn: cleanNisn,
              class: String(rawClass).trim(),
              userId: uid
            };
            
            console.log("Importing student:", student);
            batch.set(doc(activeDb, 'students', newId), student);
            successCount++;
            batchCounter++;

            if (batchCounter === 450) {
              await batch.commit();
              batch = writeBatch(activeDb);
              batchCounter = 0;
            }
          } else {
             skipCount++;
          }
        }
        
        if (batchCounter > 0) {
          await batch.commit();
        }

        // Update class list based on imported data
        const importedClasses = Array.from(new Set(rows.map(r => String(r[classIdx] || '').trim()).filter(Boolean)));
        const newClasses = importedClasses.filter(c => !classList.includes(c));
        if (newClasses.length > 0) {
           setClassList(prev => {
             const arr = [...prev, ...newClasses];
             arr.sort(compareClasses);
             return arr;
           });
        }
        
        if (excelInputRef.current) excelInputRef.current.value = '';
        
        setImportResult({
          isOpen: true,
          successCount,
          skipCount,
          totalParsed: rows.length, failCount: 0, emptyCount: 0
        });
        showToast(`Berhasil import ${successCount} siswa, ${skipCount} data gagal/terlewat.`, 'success');
      } catch (err) {
        console.error(err);
        setImportResult({
          isOpen: true,
          successCount: 0,
          skipCount: 0,
          totalParsed: 0,
          error: true,
          errorMessage: 'Gagal memproses file Excel: ' + (err as Error).message, failCount: 0, emptyCount: 0
        });
      }
    };
    if (file) reader.readAsArrayBuffer(file as Blob);
  };

  const menuItems = [
    { 
      id: 'dashboard', 
      label: 'Utama', 
      icon: LayoutGrid, 
      color: 'text-indigo-600', 
      activeBg: 'bg-indigo-50/80 text-indigo-700 ring-1 ring-indigo-200/50',
      iconBg: 'bg-indigo-500 text-white',
      hoverBg: 'hover:bg-indigo-50/40',
      inactiveBg: 'bg-indigo-50/50',
      inactiveColor: 'text-indigo-500'
    },
    { 
      id: 'attendance', 
      label: 'Presensi', 
      icon: Fingerprint, 
      color: 'text-[#8dc63f]', 
      activeBg: 'bg-emerald-50/80 text-lime-700 ring-1 ring-emerald-200/50',
      iconBg: 'bg-[#8dc63f] text-white',
      hoverBg: 'hover:bg-emerald-50/40',
      inactiveBg: 'bg-emerald-50/50',
      inactiveColor: 'text-[#7bc025]'
    },
    { 
      id: 'students', 
      label: 'Siswa', 
      icon: Users, 
      color: 'text-sky-500', 
      activeBg: 'bg-sky-50/80 text-sky-700 ring-1 ring-sky-200/50',
      iconBg: 'bg-sky-500 text-white',
      hoverBg: 'hover:bg-sky-50/40',
      inactiveBg: 'bg-sky-50/50',
      inactiveColor: 'text-sky-500'
    },
    { 
      id: 'reports', 
      label: 'Laporan', 
      icon: BarChart3, 
      color: 'text-rose-500', 
      activeBg: 'bg-rose-50/80 text-rose-700 ring-1 ring-rose-200/50',
      iconBg: 'bg-rose-500 text-white',
      hoverBg: 'hover:bg-rose-50/40',
      inactiveBg: 'bg-rose-50/50',
      inactiveColor: 'text-rose-500'
    }
  ];

  const renderContent = () => {
    switch (activeTab) {
      case 'dashboard':
        return (
          <div className="p-4 sm:p-6 lg:p-8 space-y-6 lg:space-y-8 max-w-7xl mx-auto">
            
            <AnimatePresence>
              {showIdleTip && (
                <motion.div
                  initial={{ opacity: 0, y: -20, height: 0 }}
                  animate={{ opacity: 1, y: 0, height: 'auto' }}
                  exit={{ opacity: 0, y: -20, height: 0 }}
                  transition={{ duration: 0.3 }}
                  className="overflow-hidden"
                >
                  <div className="bg-sky-50 border border-sky-100 rounded-2xl p-4 sm:p-5 flex flex-col sm:flex-row items-start sm:items-center gap-4 text-left shadow-sm mb-6">
                    <div className="w-10 h-10 shrink-0 bg-sky-100 text-sky-600 rounded-full flex items-center justify-center">
                      <Cloud className="w-5 h-5" />
                    </div>
                    <div className="flex-1">
                      <h3 className="text-sm font-bold text-sky-800 mb-1">Tips Menghemat Kuota Database Harian</h3>
                      <p className="text-xs text-sky-700 leading-relaxed">
                        Apabila aplikasi sedang tidak digunakan dalam waktu lama, harap menekan tombol <strong>Keluar (Logout)</strong> yang berada di sudut kanan atas.
                        Hal ini sangat penting untuk mengurangi aktivitas sinkronisasi di latar belakang sehingga <strong>Limit Kuota Harian (Free Tier)</strong> database Anda tidak cepat habis.
                      </p>
                    </div>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>

            {/* Top grid: Welcome Banner & Quick Action */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              
              {/* Welcome Card (spans 2 on desktop) */}
              <div className="lg:col-span-2 bg-gradient-to-br from-lime-500 flex flex-col justify-center to-lime-600 rounded-[2rem] p-6 sm:p-8 text-white shadow-lg shadow-[#8dc63f]/20 relative overflow-hidden min-h-[240px]">
                <div className="absolute top-0 right-0 p-8 flex items-center justify-center opacity-10">
                  <GraduationCap className="w-48 h-48 sm:w-64 sm:h-64 -rotate-12 transform" />
                </div>
                
                <div className="relative z-10 w-full mb-4 flex justify-start">
                  <div className="bg-white/20 backdrop-blur-md px-3 py-1.5 rounded-full text-[10px] sm:text-xs font-bold text-white flex items-center gap-1.5 border border-white/20 shadow-sm">
                    <Cloud className="w-3.5 h-3.5" />
                    Database Project: {activeAuth.app.options.projectId || 'N/A'}
                  </div>
                </div>

                <div className="relative z-10 w-full sm:w-3/4">
                  <p className="text-lime-100 text-xs sm:text-sm font-bold tracking-wider uppercase mb-1 drop-shadow-sm">
                    Halo . {activeUserCustomData?.fullname || 'User'}
                  </p>
                  <h2 className="text-3xl sm:text-4xl font-black mb-3 tracking-tight leading-tight">
                    Selamat Datang di Aplikasi <span className="text-white border-b-2 border-white/40 pb-0.5">My Kaguci App</span>
                  </h2>
                  <p className="text-emerald-50 text-base sm:text-lg font-medium opacity-90">Pantau kehadiran, kelola data nilai siswa, dan akses laporan wali kelas terkini dalam satu tempat.</p>
                </div>
                <div className="relative z-10 mt-8 grid grid-cols-1 sm:grid-cols-3 gap-3 sm:gap-4 max-w-2xl">
                  <button 
                    onClick={() => setActiveTab('attendance')} 
                    className="bg-white text-lime-800 hover:bg-lime-50 px-5 py-3.5 rounded-2xl font-extrabold text-sm shadow-md hover:shadow-lg hover:scale-102 active:scale-95 transition-all flex items-center justify-center gap-2 cursor-pointer"
                  >
                    <Fingerprint className="w-5 h-5 shrink-0" /> 
                    <span>Mulai Presensi</span>
                  </button>
                  <button 
                    onClick={handleOpenHomeroomReport} 
                    className="bg-white text-lime-800 hover:bg-lime-50 px-5 py-3.5 rounded-2xl font-extrabold text-sm shadow-md hover:shadow-lg hover:scale-102 active:scale-95 transition-all flex items-center justify-center gap-2 cursor-pointer"
                  >
                    <ClipboardList className="w-5 h-5 shrink-0 text-amber-600" /> 
                    <span>{profileData?.role === 'Wali Kelas' ? 'Laporan Wali Kelas' : 'Input Nilai Siswa'}</span>
                  </button>
                  <button 
                    onClick={() => window.location.reload()} 
                    className="bg-white text-lime-800 hover:bg-lime-50 px-5 py-3.5 rounded-2xl font-extrabold text-sm shadow-md hover:shadow-lg hover:scale-102 active:scale-95 transition-all flex items-center justify-center gap-2 cursor-pointer"
                  >
                    <RefreshCw className="w-5 h-5 shrink-0" /> 
                    <span>Refresh Halaman</span>
                  </button>
                </div>
              </div>

              {/* Date & Motivation Widget */}
              <div className="bg-white rounded-[2rem] p-6 sm:p-8 border border-slate-100 shadow-sm flex flex-col justify-center items-center text-center h-full min-h-[240px]">
                 <div className="w-20 h-20 bg-emerald-50 rounded-full flex items-center justify-center mb-6 text-[#8dc63f] shrink-0">
                    <Calendar className="w-10 h-10" />
                 </div>
                 <h3 className="text-xl sm:text-2xl font-black text-slate-800 tracking-tight leading-tight">
                   {new Date().toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}
                 </h3>
                 <p className="text-slate-600 font-semibold mt-3 text-sm">Semoga harimu menyenangkan!</p>
              </div>
            </div>

            {/* KPI Summary Grid */}
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 sm:gap-6">
              {[
                { 
                  label: isWaliKelas && profileData.waliKelasClass ? 'Total Siswa (Kelas)' : 'Total Siswa', 
                  value: (isWaliKelas && profileData.waliKelasClass ? effectiveStudents.length : students.length).toString(), 
                  icon: Users, 
                  color: 'text-indigo-600', 
                  bg: 'bg-indigo-50/70 border-indigo-100', 
                  shadow: 'hover:shadow-indigo-100/60 hover:border-indigo-300' 
                },
                { 
                  label: isWaliKelas && profileData.waliKelasClass ? 'Kelas Perwalian' : 'Total Kelas', 
                  value: isWaliKelas && profileData.waliKelasClass ? profileData.waliKelasClass : classList.length.toString(), 
                  icon: Building2, 
                  color: 'text-amber-600', 
                  bg: 'bg-amber-50/70 border-amber-100', 
                  shadow: 'hover:shadow-amber-100/60 hover:border-amber-300' 
                },
                { 
                  label: 'Tingkat Kehadiran', 
                  value: attendanceStats.rate, 
                  icon: Check, 
                  color: 'text-[#8dc63f]', 
                  bg: 'bg-emerald-50/70 border-emerald-100', 
                  shadow: 'hover:shadow-emerald-100/60 hover:border-emerald-300' 
                },
                { 
                  label: 'Total Siswa Absen', 
                  value: attendanceStats.attentionCount.toString(), 
                  icon: AlertCircle, 
                  color: 'text-rose-600', 
                  bg: 'bg-rose-50/70 border-rose-100', 
                  shadow: 'hover:shadow-rose-100/60 hover:border-rose-300' 
                },
              ].map((item, i) => (
                <div key={i} className={`p-5 sm:p-6 rounded-[2rem] border border-slate-100 bg-white flex flex-col gap-4 sm:gap-6 justify-between items-center sm:items-start text-center sm:text-left transition-all duration-300 hover:-translate-y-1 shadow-sm ${item.shadow}`}>
                  <div className={`p-4 rounded-[1.25rem] w-fit ${item.bg} border`}>
                    <item.icon className={`w-7 h-7 ${item.color}`} />
                  </div>
                  <div className="w-full">
                    <p className="text-3xl sm:text-4xl font-black text-slate-800 tracking-tighter leading-none">{item.value}</p>
                    <p className="text-xs sm:text-sm font-bold text-slate-600 mt-2">{item.label}</p>
                  </div>
                </div>
              ))}
            </div>

            {/* BARU: Menu Informasi & Ringkasan Ketidakhadiran Siswa (Alpa, Sakit, Izin) */}
            <div className="bg-white rounded-[3rem] border-2 border-slate-300/60 shadow-xl shadow-slate-100/50 flex flex-col overflow-hidden">
              {/* Header Section - Selalu Terlihat */}
              <div className="p-8 sm:p-10 border-b border-slate-100 bg-white shadow-sm z-20 shrink-0">
                <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-8">
                  <div className="flex-1">
                    <div className="flex items-center gap-4">
                      <div className="p-3 bg-rose-50 rounded-[1.25rem] text-rose-500 shadow-inner">
                        <AlertCircle className="w-7 h-7" />
                      </div>
                      <div>
                        <h2 className="text-2xl sm:text-3xl font-black text-slate-900 tracking-tight">Informasi Ketidakhadiran Siswa</h2>
                        <p className="text-sm sm:text-base text-slate-600 font-semibold mt-1.5 opacity-80">
                          Analisis riwayat ketidakhadiran terbanyak di setiap kelas.
                        </p>
                      </div>
                    </div>
                  </div>
                  
                  {/* Tab selector */}
                  <div className="flex bg-slate-100 p-1.5 rounded-[1.25rem] border-2 border-slate-300/60 self-start sm:self-auto shrink-0 shadow-inner">
                    <button 
                      onClick={() => setDashboardActiveStatsTab('class_summary')}
                      className={`px-5 py-2.5 rounded-xl text-xs font-black transition-all duration-300 ${
                        dashboardActiveStatsTab === 'class_summary' 
                          ? 'bg-white text-slate-900 shadow-xl shadow-slate-200/50 scale-100' 
                          : 'text-slate-600 hover:text-slate-700 hover:bg-white/40'
                      }`}
                    >
                      Ringkasan Kelas
                    </button>
                    <button 
                      onClick={() => setDashboardActiveStatsTab('top_rankings')}
                      className={`px-5 py-2.5 rounded-xl text-xs font-black transition-all duration-300 ${
                        dashboardActiveStatsTab === 'top_rankings' 
                          ? 'bg-white text-slate-900 shadow-xl shadow-slate-200/50 scale-100' 
                          : 'text-slate-600 hover:text-slate-700 hover:bg-white/40'
                      }`}
                    >
                      Ranking Ketidakhadiran
                    </button>
                  </div>
                </div>
              </div>

              <div className="p-6 sm:p-8 overflow-y-auto max-h-[850px] lg:max-h-[700px] flex-1 custom-scrollbar scroll-smooth">
                {dashboardActiveStatsTab === 'class_summary' ? (
                <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-6 pb-4">
                  {studentAbsenceStats.length === 0 ? (
                    <div className="col-span-full py-12 text-center bg-slate-50 rounded-[2rem] border border-dashed border-slate-200">
                      <p className="text-slate-500 font-bold mb-2">Belum ada data kelas yang terdaftar</p>
                      <button onClick={() => setActiveTab('students')} className="text-sm font-bold text-[#7bc025] hover:underline">
                        Mulai dengan mengelola kelas & siswa di sini
                      </button>
                    </div>
                  ) : (
                    studentAbsenceStats.map((clsData) => {

                      
                      return (
                        <div key={clsData.className} className="bg-slate-50/50 rounded-2xl p-5 border border-slate-100 flex flex-col justify-between hover:shadow-md transition-all">
                          <div>
                            <div className="flex items-center justify-between mb-4">
                              <span className="bg-[#8dc63f] text-white px-3.5 py-1 rounded-full text-xs font-black shadow-sm">
                                {clsData.className}
                              </span>
                              <span className="text-xs text-slate-500 font-bold">
                                {clsData.allAbsenceList.length} siswa absen
                              </span>
                            </div>

                            <div className="space-y-3.5">
                              {/* Summary of Absences */}
                              <div className="grid grid-cols-3 gap-2">
                                <div className="bg-white p-3 rounded-xl border border-rose-100 flex flex-col items-center justify-center text-center">
                                  <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-1">Alpa</span>
                                  <span className="text-lg font-black text-rose-600">{clsData.totalAlpaStudents}</span>
                                  <span className="text-[9px] text-slate-400 font-medium">Siswa</span>
                                </div>
                                <div className="bg-white p-3 rounded-xl border border-amber-100 flex flex-col items-center justify-center text-center">
                                  <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-1">Sakit</span>
                                  <span className="text-lg font-black text-amber-500">{clsData.totalSakitStudents}</span>
                                  <span className="text-[9px] text-slate-400 font-medium">Siswa</span>
                                </div>
                                <div className="bg-white p-3 rounded-xl border border-indigo-100 flex flex-col items-center justify-center text-center">
                                  <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-1">Izin</span>
                                  <span className="text-lg font-black text-indigo-500">{clsData.totalIzinStudents}</span>
                                  <span className="text-[9px] text-slate-400 font-medium">Siswa</span>
                                </div>
                              </div>
                            </div>
                          </div>

                          <button 
                            onClick={() => setDashboardSelectedClassDetail(clsData.className)}
                            className="w-full mt-4 bg-white hover:bg-slate-100 border-2 border-slate-300 text-slate-700 py-2 rounded-xl text-xs font-bold transition-colors flex items-center justify-center gap-1.5"
                          >
                            <Info className="w-3.5 h-3.5" /> Detail Absensi Kelas
                          </button>
                        </div>
                      );
                    })
                  )}
                </div>
              ) : (
                <div className="overflow-x-auto rounded-2xl border border-slate-100">
                  <table className="w-full text-left border-collapse bg-slate-50/20">
                    <thead>
                      <tr className="bg-slate-50 text-slate-500 text-xs font-bold uppercase border-b border-slate-100">
                        <th className="py-4 px-6">Nama Siswa</th>
                        <th className="py-4 px-6">Kelas</th>
                        <th className="py-4 px-6 text-center">Sakit (S)</th>
                        <th className="py-4 px-6 text-center">Izin (I)</th>
                        <th className="py-4 px-6 text-center">Alpa (A)</th>
                        <th className="py-4 px-6 text-center">Total Absen</th>
                        <th className="py-4 px-6 text-right">Tindakan</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {(() => {
                        const allAbsentOrdered = studentAbsenceStats
                          .flatMap(c => c.allAbsenceList.map(s => ({ ...s, className: c.className })))
                          .sort((a, b) => b.total - a.total || b.alpa - a.alpa);
                          
                        if (allAbsentOrdered.length === 0) {
                          return (
                            <tr>
                              <td colSpan={7} className="py-12 text-center text-slate-500 font-medium italic bg-white animate-fade">
                                Belum ada siswa berstatus Alpa, Sakit, atau Izin.
                              </td>
                            </tr>
                          );
                        }
                        
                        return allAbsentOrdered.map((student, rank) => (
                          <tr key={student.id} className="hover:bg-slate-50/50 bg-white transition-colors">
                            <td className="py-4 px-6 font-bold text-slate-800 flex items-center gap-2">
                              <span className={`w-5 h-5 flex items-center justify-center rounded-full text-[10px] font-black shrink-0 ${
                                rank === 0 ? 'bg-rose-500 text-white' : 
                                rank === 1 ? 'bg-amber-500 text-white' : 
                                rank === 2 ? 'bg-indigo-500 text-white' : 'bg-slate-200 text-slate-600'
                              }`}>
                                {rank + 1}
                              </span>
                              {student.name}
                            </td>
                            <td className="py-4 px-6 text-slate-600 font-bold">{student.className}</td>
                            <td className="py-4 px-6 text-center">
                              {student.sakit > 0 ? (
                                <span className="bg-amber-50 text-amber-600 px-2 py-1 rounded-md text-xs font-extrabold border border-amber-100">
                                  {student.sakit} kali
                                </span>
                              ) : (
                                <span className="text-slate-300">-</span>
                              )}
                            </td>
                            <td className="py-4 px-6 text-center">
                              {student.izin > 0 ? (
                                <span className="bg-indigo-50 text-indigo-600 px-2 py-1 rounded-md text-xs font-extrabold border border-indigo-100">
                                  {student.izin} kali
                                </span>
                              ) : (
                                <span className="text-slate-300">-</span>
                              )}
                            </td>
                            <td className="py-4 px-6 text-center">
                              {student.alpa > 0 ? (
                                <span className="bg-rose-50 text-rose-600 px-2 py-1 rounded-md text-xs font-extrabold border border-rose-100 animate-pulse">
                                  {student.alpa} kali
                                </span>
                              ) : (
                                <span className="text-slate-300">-</span>
                              )}
                            </td>
                            <td className="py-4 px-6 text-center">
                              <span className="bg-slate-100 text-slate-700 font-black px-2.5 py-1.5 rounded-full text-xs border-2 border-slate-300">
                                {student.total} Hari
                              </span>
                            </td>
                            <td className="py-4 px-6 text-right">
                              <button 
                                onClick={() => {
                                  setActiveTab('reports');
                                }}
                                className="text-xs text-[#7bc025] hover:text-lime-700 font-extrabold flex items-center justify-end gap-1 ml-auto"
                              >
                                Lihat Rekap <ArrowRight className="w-3.5 h-3.5" />
                              </button>
                            </td>
                          </tr>
                        ));
                      })()}
                    </tbody>
                  </table>
                </div>
              )}
              </div>
            </div>

            {/* MODAL / KOTAK DETAIL ABSENSI PER KELAS */}
            <AnimatePresence>
              {dashboardSelectedClassDetail && (
                <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[9999] flex items-center justify-center p-4 overflow-y-auto">
                  <motion.div 
                    initial={{ opacity: 0, scale: 0.95, y: 20 }}
                    animate={{ opacity: 1, scale: 1, y: 0 }}
                    exit={{ opacity: 0, scale: 0.95, y: 20 }}
                    className="bg-white rounded-[2.5rem] w-full max-w-2xl overflow-hidden shadow-2xl border border-slate-100 flex flex-col my-8"
                  >
                    <div className="bg-gradient-to-r from-lime-600 to-lime-700 p-6 sm:p-8 text-white flex justify-between items-center shrink-0">
                      <div>
                        <span className="bg-white/20 text-white px-3 py-1 rounded-full text-xs font-extrabold border border-white/10 uppercase tracking-widest">
                          Detail Ketidakhadiran
                        </span>
                        <h3 className="text-2xl font-black mt-2">Kelas {dashboardSelectedClassDetail}</h3>
                      </div>
                      <button 
                        onClick={() => setDashboardSelectedClassDetail(null)}
                        className="p-3 bg-white/10 rounded-2xl hover:bg-white/20 transition-all border border-white/10"
                      >
                        <X className="w-6 h-6" />
                      </button>
                    </div>
                    
                    <div className="p-6 sm:p-10 overflow-y-auto max-h-[85vh] md:max-h-[80vh] space-y-6 custom-scrollbar flex-1 bg-slate-50/20">
                      {(() => {
                        const classInfo = studentAbsenceStats.find(c => c.className === dashboardSelectedClassDetail);
                        if (!classInfo || classInfo.allAbsenceList.length === 0) {
                          return (
                            <div className="py-24 text-center text-slate-500 bg-white rounded-[2.5rem] border border-dashed border-slate-200 shadow-sm mx-2">
                              <div className="mx-auto w-16 h-16 bg-slate-50 rounded-full flex items-center justify-center mb-4">
                                <CheckCircle className="w-8 h-8 text-[#8dc63f]" />
                              </div>
                              <p className="font-bold text-slate-600">Semua siswa di kelas ini hadir 100%.</p>
                              <p className="text-xs mt-2 uppercase tracking-widest opacity-60">Tidak ada riwayat ketidakhadiran</p>
                            </div>
                          );
                        }

                        return (
                          <div className="space-y-5">
                            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                              <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest">
                                Daftar Siswa
                              </p>
                              <span className="text-[10px] bg-slate-100 px-2 py-1 rounded-md text-slate-600 font-black">
                                {classInfo.allAbsenceList.length} SISWA TERCATAT
                              </span>
                            </div>
                            <div className="divide-y divide-slate-100">
                              {classInfo.allAbsenceList.map((stu) => (
                                <div key={stu.id} className="py-5 flex flex-col sm:flex-row sm:items-center justify-between gap-4 first:pt-0 last:pb-0 hover:bg-slate-50/30 transition-colors px-2 rounded-xl">
                                  <div className="flex-1">
                                    <h4 className="font-black text-slate-800 text-lg tracking-tight">{stu.name}</h4>
                                    <div className="flex items-center gap-2 mt-1.5">
                                      <span className="w-1.5 h-1.5 rounded-full bg-slate-300"></span>
                                      <p className="text-xs text-slate-600 font-bold">Total: {stu.total} hari tidak hadir</p>
                                    </div>
                                  </div>
                                  <div className="flex flex-wrap items-center gap-2 sm:shrink-0">
                                    {stu.sakit > 0 && (
                                      <div className="flex flex-col items-start bg-amber-50 px-3.5 py-2 rounded-2xl border border-amber-100 shadow-sm shadow-amber-100/50">
                                        <span className="text-amber-600 text-xs font-black flex items-center gap-1.5">
                                          <span className="w-2 h-2 bg-amber-500 rounded-full animate-pulse"></span>
                                          {stu.sakit} Sakit
                                        </span>
                                        <span className="text-[9px] text-amber-600/70 mt-1 font-semibold leading-tight max-w-[120px]">
                                          {stu.sakitDates.map(d => new Date(d).toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric'})).join(', ')}
                                        </span>
                                      </div>
                                    )}
                                    {stu.izin > 0 && (
                                      <div className="flex flex-col items-start bg-indigo-50 px-3.5 py-2 rounded-2xl border border-indigo-100 shadow-sm shadow-indigo-100/50">
                                        <span className="text-indigo-600 text-xs font-black flex items-center gap-1.5">
                                          <span className="w-2 h-2 bg-indigo-500 rounded-full animate-pulse"></span>
                                          {stu.izin} Izin
                                        </span>
                                        <span className="text-[9px] text-indigo-600/70 mt-1 font-semibold leading-tight max-w-[120px]">
                                          {stu.izinDates.map(d => new Date(d).toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric'})).join(', ')}
                                        </span>
                                      </div>
                                    )}
                                    {stu.alpa > 0 && (
                                      <div className="flex flex-col items-start bg-rose-50 px-3.5 py-2 rounded-2xl border border-rose-100 shadow-sm shadow-rose-100/50">
                                        <span className="text-rose-600 text-xs font-black flex items-center gap-1.5">
                                          <span className="w-2 h-2 bg-rose-500 rounded-full animate-pulse"></span>
                                          {stu.alpa} Alpa
                                        </span>
                                        <span className="text-[9px] text-rose-600/70 mt-1 font-semibold leading-tight max-w-[120px]">
                                          {stu.alpaDates.map(d => new Date(d).toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric'})).join(', ')}
                                        </span>
                                      </div>
                                    )}
                                  </div>
                                </div>
                              ))}
                            </div>
                          </div>
                        );
                      })()}
                    </div>

                    <div className="p-6 bg-slate-50 border-t border-slate-100 flex justify-end gap-3 shrink-0">
                      <button 
                        onClick={() => setDashboardSelectedClassDetail(null)}
                        className="bg-white border-2 border-slate-300 text-slate-700 px-6 py-3 rounded-2xl font-bold text-sm hover:bg-slate-100 transition-all"
                      >
                        Tutup
                      </button>
                      <button 
                        onClick={() => {
                          setDashboardSelectedClassDetail(null);
                          setActiveTab('reports');
                        }}
                        className="bg-[#7bc025] text-white px-6 py-3 rounded-2xl font-bold text-sm hover:bg-lime-700 transition-all shadow-sm flex items-center gap-2"
                      >
                        Buka Laporan <ArrowRight className="w-4 h-4" />
                      </button>
                    </div>
                  </motion.div>
                </div>
              )}
            </AnimatePresence>

            {/* Bottom Section - Quick Actions / Shortcuts */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
               {/* Quick Reports Access */}
               <div className="bg-white p-6 sm:p-8 rounded-[2rem] border border-slate-100 shadow-sm">
                  <div className="flex items-center gap-4 mb-6">
                    <div className="p-3 bg-sky-50 rounded-2xl text-sky-500">
                       <BarChart3 className="w-6 h-6" />
                    </div>
                    <h2 className="text-xl font-bold text-slate-800 tracking-tight">Laporan & Rekap</h2>
                  </div>
                  <div className="space-y-3">
                     <button onClick={() => setActiveTab('reports')} className="w-full p-4 sm:p-5 rounded-2xl border border-slate-100 hover:border-sky-200 hover:bg-sky-50 transition-colors flex items-center justify-between group">
                        <div className="flex items-center gap-4">
                           <FileText className="w-6 h-6 text-slate-500 group-hover:text-sky-500 transition-colors" />
                           <span className="font-bold text-slate-600 group-hover:text-sky-700 transition-colors">Laporan Presensi Harian</span>
                        </div>
                        <ArrowRight className="w-5 h-5 text-slate-300 group-hover:text-sky-500 group-hover:translate-x-1 transition-all" />
                     </button>
                     <button onClick={() => setActiveTab('reports')} className="w-full p-4 sm:p-5 rounded-2xl border border-slate-100 hover:border-sky-200 hover:bg-sky-50 transition-colors flex items-center justify-between group">
                        <div className="flex items-center gap-4">
                           <FileText className="w-6 h-6 text-slate-500 group-hover:text-sky-500 transition-colors" />
                           <span className="font-bold text-slate-600 group-hover:text-sky-700 transition-colors">Rekapitulasi Bulanan</span>
                        </div>
                        <ArrowRight className="w-5 h-5 text-slate-300 group-hover:text-sky-500 group-hover:translate-x-1 transition-all" />
                     </button>
                     <button 
                       onClick={handleOpenHomeroomReport}
                       className="w-full p-4 sm:p-5 rounded-2xl border border-amber-100 hover:border-amber-300 hover:bg-amber-50 transition-colors flex items-center justify-between group cursor-pointer"
                     >
                        <div className="flex items-center gap-4">
                           <ClipboardList className="w-6 h-6 transition-colors text-amber-500 group-hover:text-amber-600" />
                           <span className="font-bold transition-colors text-slate-600 group-hover:text-amber-800">Laporan Bulanan Wali Kelas</span>
                        </div>
                        <ArrowRight className="w-5 h-5 transition-all text-slate-300 group-hover:text-amber-600 group-hover:translate-x-1" />
                     </button>
                  </div>
               </div>

               {/* Hint / Setup Call to Action */}
               <div className="bg-white p-6 sm:p-8 rounded-[2rem] border border-slate-100 shadow-sm flex flex-col items-center justify-center text-center relative overflow-hidden">
                  <div className="absolute top-0 right-0 p-6 opacity-5">
                    <Settings2 className="w-48 h-48 -rotate-45" />
                  </div>
                  <div className="w-20 h-20 bg-slate-50 rounded-[1.5rem] flex items-center justify-center text-slate-500 mb-6 relative z-10 border border-slate-100">
                     <Plus className="w-10 h-10" />
                  </div>
                  <h3 className="text-xl font-bold text-slate-800 relative z-10 mb-2">Perbarui Data Master</h3>
                  <p className="text-sm font-medium text-slate-600 max-w-xs relative z-10">Data siswa atau kelas ada yang baru? Segera tambahkan untuk keakuratan presensi.</p>
                  <button onClick={() => setActiveTab('students')} className="mt-8 bg-[#8dc63f] text-white px-6 py-3 rounded-2xl font-bold text-sm shadow-sm hover:bg-[#7bc025] hover:shadow-md transition-all flex items-center gap-2 relative z-10">
                     Buka Manajemen Siswa <ArrowRight className="w-4 h-4 ml-1" />
                  </button>
               </div>
            </div>

          </div>
        );
      case 'students':
        return (
          <div className="p-8 space-y-8">
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
              {/* Kelas */}
              <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm space-y-6">
                <div className="flex items-center justify-between">
                  <h2 className="text-lg font-bold text-slate-800">Manajemen Kelas</h2>
                  {isWaliKelas && profileData?.waliKelasClass && (
                    <span className="text-xs font-bold text-emerald-800 bg-emerald-50 px-2.5 py-1 rounded-lg border border-emerald-200">
                      Wali Kelas: {profileData.waliKelasClass}
                    </span>
                  )}
                </div>
                {isWaliKelas && profileData?.waliKelasClass ? (
                  <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-3.5 text-xs text-emerald-800 font-medium leading-relaxed">
                    Mode Wali Kelas aktif. Anda dikhususkan untuk mengelola kelas perwalian <b>{profileData.waliKelasClass}</b>. Seluruh kelas lainnya disembunyikan.
                  </div>
                ) : (
                  <div className="flex flex-col sm:flex-row gap-4">
                    <input type="text" className="flex-1 p-3 border rounded-xl" placeholder="Nama Kelas (Contoh: X-A)" onKeyDown={(e) => {
                      if (e.key === 'Enter') { 
                        const val = e.currentTarget.value.trim(); 
                        if (val) {
                          if (classList.some(c => c.toLowerCase() === val.toLowerCase())) {
                            showToast('Kelas "' + val + '" sudah terdaftar', 'error');
                          } else {
                            const arr = [...classList, val];
                            arr.sort(compareClasses);
                            setClassList(arr); 
                            
                            // Explicit cloud save
                            if (currentUser) {
                              setDoc(doc(activeDb, 'users', currentUser.uid), { classList: arr }, { merge: true })
                                .catch(e => console.error("Error saving new class:", e));
                            }
                            
                            e.currentTarget.value = ''; 
                          }
                        }
                      }
                    }} />
                    <button className="bg-[#8dc63f] text-white font-bold py-3 px-6 rounded-xl hover:bg-[#7bc025] w-full sm:w-auto" onClick={(e) => {
                      const input = e.currentTarget.previousElementSibling as HTMLInputElement;
                      const val = input.value.trim(); 
                      if (val) {
                        if (classList.some(c => c.toLowerCase() === val.toLowerCase())) {
                            showToast('Kelas "' + val + '" sudah terdaftar', 'error');
                        } else {
                          const newList = [...classList, val];
                          newList.sort(compareClasses);
                          setClassList(newList); 
                          
                          // Explicit cloud save
                          if (currentUser) {
                            setDoc(doc(activeDb, 'users', currentUser.uid), { classList: newList }, { merge: true })
                              .catch(e => console.error("Error saving new class:", e));
                          }
                          
                          input.value = ''; 
                        }
                      }
                    }}>Tambah</button>
                  </div>
                )}
                <div className="flex flex-wrap gap-3">
                  {effectiveClassList.length === 0 ? <p className="text-slate-600 italic">Belum ada kelas.</p> : effectiveClassList.slice().sort(compareClasses).map(c => {
                    const count = effectiveStudents.filter(s => s.class === c).length;
                    return (
                      <div key={c} className="flex flex-col gap-1 border border-slate-200 bg-slate-50 rounded-2xl p-3">
                        <div className="flex items-center gap-2 justify-between">
                          <span className="font-bold text-slate-700">{c}</span>
                          <span className="bg-white px-2 py-0.5 rounded-full text-xs font-bold text-slate-500 border shadow-sm">{count} Siswa</span>
                        </div>
                        {!isWaliKelas && (
                          <div className="flex gap-2 mt-2">
                            <button 
                              className="flex-1 text-xs py-1.5 px-3 rounded-lg bg-blue-50 text-blue-600 font-bold hover:bg-blue-100 flex justify-center items-center gap-1 transition-colors"
                              onClick={() => {
                                const newName = window.prompt(`Edit nama kelas "${c}":`, c);
                                if (newName && newName.trim() && newName.trim() !== c) {
                                  const trimmed = newName.trim();
                                  if (classList.some(existing => existing.toLowerCase() === trimmed.toLowerCase())) {
                                    showToast('Kelas "' + trimmed + '" sudah terdaftar', 'error');
                                    return;
                                  }
                                  const updatedList = classList.map(existing => existing === c ? trimmed : existing);
                                  updatedList.sort((a,b) => a.localeCompare(b, 'id-ID', { numeric: true }));
                                  setClassList(updatedList);
                                  
                                  const updatedStudents = students.map(s => s.class === c ? { ...s, class: trimmed } : s);
                                  setStudents(updatedStudents);
                                  
                                  const updatedSessions = attendanceSessions.map(sess => sess.className === c ? { ...sess, className: trimmed } : sess);
                                  setAttendanceSessions(updatedSessions);

                                  if (currentUser) {
                                    const batch = writeBatch(activeDb);
                                    batch.set(doc(activeDb, 'users', currentUser.uid), { classList: updatedList }, { merge: true });
                                    
                                    updatedStudents.forEach(s => {
                                      if (s.class === trimmed) {
                                        batch.set(doc(activeDb, 'students', s.id), { class: trimmed }, { merge: true });
                                      }
                                    });
                                    
                                    updatedSessions.forEach(sess => {
                                      if (sess.className === trimmed) {
                                        batch.set(doc(activeDb, 'attendanceSessions', sess.id), { className: trimmed }, { merge: true });
                                      }
                                    });

                                    batch.commit().then(() => {
                                      showToast(`Kelas ${c} berhasil diubah menjadi ${trimmed}`, 'success');
                                    }).catch(e => console.error(e));
                                  }
                                }
                              }}
                            >
                              <Pencil className="w-3.5 h-3.5" /> Edit
                            </button>
                            <button 
                              className="flex-1 text-xs py-1.5 px-3 rounded-lg bg-red-50 text-red-600 font-bold hover:bg-red-100 flex justify-center items-center gap-1 transition-colors"
                              onClick={() => {
                                if (window.confirm(`Hapus kelas "${c}"?\n\nPERINGATAN: ${count > 0 ? `Ada ${count} siswa di kelas ini. Jika dihapus, siswa akan tetap ada tapi kelasnya akan menjadi kosong/tidak ada.` : 'Kelas kosong.'}`)) {
                                  const updatedList = classList.filter(existing => existing !== c);
                                  setClassList(updatedList);
                                  
                                  const updatedStudents = students.map(s => s.class === c ? { ...s, class: '' } : s);
                                  setStudents(updatedStudents);

                                  if (currentUser) {
                                    const batch = writeBatch(activeDb);
                                    batch.set(doc(activeDb, 'users', currentUser.uid), { classList: updatedList }, { merge: true });
                                    
                                    students.forEach(s => {
                                      if (s.class === c) {
                                        batch.set(doc(activeDb, 'students', s.id), { class: '' }, { merge: true });
                                      }
                                    });

                                    batch.commit()
                                      .then(() => showToast(`Kelas ${c} berhasil dihapus`, 'success'))
                                      .catch(e => console.error(e));
                                  }
                                }
                              }}
                            >
                              <Trash2 className="w-3.5 h-3.5" /> Hapus
                            </button>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
              
              {/* Siswa */}
              <div id="student-input-container" className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm space-y-6">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                  <h2 className="text-lg font-bold text-slate-800">Tambah Siswa</h2>
                  <div>
                    <input 
                      type="file" 
                      accept=".xlsx, .xls" 
                      className="hidden" 
                      ref={excelInputRef} 
                      onChange={handleImportExcel} 
                    />
                    <button 
                      onClick={() => excelInputRef.current?.click()}
                      className="w-full sm:w-auto bg-[#8dc63f]/10 text-[#8dc63f] border border-[#8dc63f]/20 font-bold py-2 px-4 rounded-xl hover:bg-[#8dc63f]/20 transition-colors text-sm flex items-center justify-center gap-2"
                    >
                       <Download className="w-4 h-4" /> Import Excel
                    </button>
                  </div>
                </div>
                <div className="grid grid-cols-1 gap-4">
                  <input type="text" className="p-3 border rounded-xl" placeholder="Nama Lengkap" value={newStudent.name} onChange={(e) => setNewStudent({...newStudent, name: e.target.value})} />
                  <input type="text" className="p-3 border rounded-xl" placeholder="NIS" value={newStudent.nisn} onChange={(e) => setNewStudent({...newStudent, nisn: e.target.value.replace(/\D/g, '')})} />
                  <select 
                    className="p-3 bg-white border-2 border-slate-300 rounded-xl focus:ring-2 focus:ring-[#8dc63f] disabled:bg-slate-100 disabled:text-slate-700 disabled:cursor-not-allowed" 
                    value={isWaliKelas && profileData?.waliKelasClass ? profileData.waliKelasClass : newStudent.class} 
                    disabled={isWaliKelas && !!profileData?.waliKelasClass}
                    onChange={(e) => setNewStudent({...newStudent, class: e.target.value})}
                  >
                    {!isWaliKelas && <option value="">Pilih Kelas</option>}
                    {effectiveClassList.slice().sort(compareClasses).map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
                <button className="bg-[#8dc63f] text-white font-bold py-3 px-6 rounded-xl hover:bg-[#7bc025] w-full" onClick={addOrUpdateStudent}>{editingStudentId ? 'Update' : 'Simpan'}</button>
              </div>
            </div>

            <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm">
              <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 mb-6">
                <div>
                  <h2 className="text-lg font-bold text-slate-800">
                    Daftar Siswa {isWaliKelas && profileData?.waliKelasClass ? `Kelas ${profileData.waliKelasClass}` : ''}
                  </h2>
                  {isWaliKelas && profileData?.waliKelasClass && (
                    <p className="text-xs text-slate-500 font-medium mt-0.5">
                      Menampilkan {effectiveStudents.length} siswa binaan di kelas Anda.
                    </p>
                  )}
                </div>
                {effectiveStudents.length > 0 && !isWaliKelas && (
                  <button 
                    onClick={() => setResetModalType('clear_all_students')} 
                    className="w-full sm:w-auto bg-rose-50 text-rose-600 border border-rose-200 font-bold py-2 px-4 rounded-xl hover:bg-rose-100 transition-colors text-xs flex items-center justify-center gap-2"
                  >
                    <Trash2 className="w-4 h-4" /> Kosongkan Semua Siswa
                  </button>
                )}
              </div>
              {!studentsLoaded ? (
                <p className="text-slate-600 text-center py-6">Memuat data...</p>
              ) : effectiveStudents.length === 0 ? (
                <p className="text-slate-600 text-center py-6">
                  {isWaliKelas && profileData?.waliKelasClass 
                    ? `Belum ada siswa terdaftar di kelas ${profileData.waliKelasClass}.` 
                    : 'Belum ada siswa.'}
                </p>
              ) : (
                <div className="overflow-auto max-h-[600px] border border-slate-100 rounded-xl scrollbar-thin">
                  <table className="w-full min-w-[650px] text-left border-collapse">
                    <thead className="sticky top-0 bg-slate-50 z-10 shadow-[0_1px_0_0_rgba(226,232,240,1)]">
                      <tr>
                        <th className="p-4 font-bold text-slate-600 bg-slate-50">No.</th>
                        <th className="p-4 font-bold text-slate-600 bg-slate-50">Nama</th>
                        <th className="p-4 font-bold text-slate-600 bg-slate-50">NIS</th>
                        <th className="p-4 font-bold text-slate-600 bg-slate-50">Kelas</th>
                        <th className="p-4 font-bold text-slate-600 bg-slate-50">Aksi</th>
                      </tr>
                    </thead>
                    <tbody>
                      {effectiveStudents.map((student, i) => (
                        <tr key={student.id} className={`border-b last:border-b-0 hover:bg-slate-50/80 transition-colors ${deletingStudentId === student.id ? 'opacity-50 bg-rose-50' : ''}`}>
                          <td className="p-4 text-slate-700">{i + 1}</td>
                           <td className="p-4 font-bold text-slate-900">
                             {student.name}
                             {deletingStudentId === student.id && <span className="ml-2 text-[10px] text-rose-500 font-bold animate-pulse">MENGHAPUS...</span>}
                           </td>
                          <td className="p-4 text-slate-700">{student.nisn}</td>
                          <td className="p-4 text-slate-700">{student.class}</td>
                          <td className="p-4 flex gap-2">
                            <button 
                              onClick={() => handleEditStudent(student)} 
                              disabled={!!deletingStudentId}
                              className="p-1 border border-slate-100 rounded-lg hover:border-[#8dc63f] text-[#8dc63f] hover:text-[#7bc025] hover:bg-emerald-50/50 transition-colors disabled:opacity-30 disabled:cursor-not-allowed"
                            >
                              <Pencil className="w-4 h-4" />
                            </button>
                            <button 
                              onClick={() => setConfirmationAction({ type: 'delete', student })} 
                              disabled={!!deletingStudentId}
                              className="p-1 border border-slate-100 rounded-lg hover:border-rose-200 text-rose-600 hover:text-rose-800 hover:bg-rose-50/50 transition-colors disabled:opacity-30 disabled:cursor-not-allowed"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </div>
        );
      case 'attendance':
        return (
          <AttendanceView
            classList={effectiveClassList}
            students={effectiveStudents}
            attendanceSessions={attendanceSessions}
            showToast={showToast}
            activeDb={activeDb}
            activeAuth={activeAuth}
            trackOp={trackOp}
          />
        );
      case 'reports':
        return (
          <ReportsView 
            classList={effectiveClassList}
            students={effectiveStudents}
            attendanceSessions={attendanceSessions}
            profileData={profileData}
            activeDb={activeDb}
            activeAuth={activeAuth}
            trackOp={trackOp}
            showToast={showToast}
            onNavigateToProfile={() => setActiveTab('profile')}
          />
        );
      case 'homeroom_report':
        return (
          <ReportsView 
            classList={effectiveClassList}
            students={effectiveStudents}
            attendanceSessions={attendanceSessions}
            profileData={profileData}
            activeDb={activeDb}
            activeAuth={activeAuth}
            trackOp={trackOp}
            showToast={showToast}
            initialFrame="wali_kelas"
            onNavigateToProfile={() => setActiveTab('profile')}
          />
        );
      case 'profile':
        return (
          <div className="p-4 sm:p-6 lg:p-8 space-y-6 flex flex-col items-center max-w-7xl mx-auto w-full">
            <h2 className="text-2xl sm:text-3xl font-black text-slate-800 tracking-tight self-start mb-2">Profil Saya</h2>
            
            <div className="bg-white w-full rounded-[2rem] p-6 sm:p-8 border border-slate-100 shadow-sm flex flex-col items-center">
              <div className="relative mb-6">
                <div className="w-32 h-32 rounded-full overflow-hidden bg-slate-100 border-4 border-white shadow-lg flex items-center justify-center relative group">
                    {avatarUrl ? (
                      <img src={avatarUrl} alt="Avatar" className="w-full h-full object-cover rounded-full" />
                    ) : (
                      <UserIcon className="w-12 h-12 text-slate-500" />
                    )}
                    
                    {isUploadingPhoto && (
                      <div className="absolute inset-0 bg-black/65 backdrop-blur-xs flex flex-col items-center justify-center text-white">
                        <span className="w-4.5 h-4.5 border-2 border-lime-400 border-t-transparent rounded-full animate-spin"></span>
                        <p className="text-[9px] font-black mt-1 text-lime-400 uppercase tracking-widest">Saving...</p>
                      </div>
                    )}
                </div>
                <button 
                  onClick={() => fileInputRef.current?.click()}
                  className="absolute bottom-0 right-0 bg-[#8dc63f] text-white p-3 rounded-full shadow-lg hover:bg-[#7bc025] transition-all hover:scale-105 active:scale-95 cursor-pointer"
                  title="Ganti Foto Profil"
                  disabled={isUploadingPhoto}
                >
                  <Camera className="w-5 h-5" />
                </button>
              </div>

              <div className="text-center w-full mb-8 flex flex-col items-center">
                <h3 className="text-xl font-black text-slate-800">{activeUserCustomData?.fullname || profileData.namaGuruMapel || 'Agan Parta,S.Kom.'}</h3>
                <p className="text-sm font-medium text-slate-600 my-1">{activeAuth.currentUser?.email || 'N/A'}</p>
                <div className="inline-flex items-center gap-1.5 bg-emerald-50 text-[#8dc63f] px-3 py-1 rounded-lg text-xs font-bold mt-2">
                  <span className="w-2 h-2 rounded-full bg-[#8dc63f]"></span> Online
                </div>

                {/* PWA Install Card inside Profile */}
                <div className="mt-6 w-full max-w-md bg-stone-50 border border-slate-100 rounded-2xl p-5 text-center shadow-xs">
                  <div className="flex items-center justify-center gap-2 mb-2 text-slate-800">
                    <Download className="w-4.5 h-4.5 text-[#8dc63f]" />
                    <span className="font-extrabold text-[10px] tracking-widest text-slate-700 uppercase">Aplikasi PWA (Absensi Seluler)</span>
                  </div>
                  {isInsideIframe ? (
                    <div className="space-y-2">
                      <p className="text-xs text-slate-600 leading-relaxed">
                        Anda sedang membuka aplikasi ini di dalam frame preview. Untuk menginstalnya ke HP / Desktop agar dapat dibuka langsung tanpa browser:
                      </p>
                      <a
                        href="https://ais-pre-56w2g4yxpoxk4k23siyzdq-901834158843.asia-southeast1.run.app"
                        target="_blank"
                        rel="noopener noreferrer"
                        className="inline-flex items-center gap-2 px-4 py-2 bg-[#8dc63f] hover:bg-[#7bc025] text-white text-xs font-bold rounded-xl transition-all shadow-[0_4px_12px_rgba(141,198,63,0.25)] hover:scale-102"
                      >
                        Buka di Tab Baru & Instal <ArrowRight className="w-3.5 h-3.5" />
                      </a>
                    </div>
                  ) : isAppInstalled ? (
                    <p className="text-xs text-emerald-600 font-bold">
                      🎉 Aplikasi telah terpasang dengan sukses! Anda dapat membukanya langsung dari layar utama perangkat Anda.
                    </p>
                  ) : isInstallable ? (
                    <div className="space-y-2.5">
                      <p className="text-xs text-slate-600 leading-relaxed">
                        Aplikasi absensi ini siap dipasang langsung di HP atau Komputer Anda untuk akses cepat dan hemat kuota.
                      </p>
                      <button
                        onClick={handleInstallClick}
                        className="inline-flex items-center gap-2 px-4 py-2 bg-[#8dc63f] hover:bg-[#7bc025] text-white text-xs font-bold rounded-xl transition-all shadow-[0_4px_12px_rgba(141,198,63,0.25)] hover:scale-102 active:scale-98"
                      >
                        Instal Sekarang
                      </button>
                    </div>
                  ) : (
                    <p className="text-[11px] text-slate-500 leading-relaxed">
                      Buka menu browser Anda (klik titik tiga di Chrome atau tombol Bagikan di Safari iOS) dan pilih <b>"Instal Aplikasi"</b> atau <b>"Tambahkan ke Layar Utama"</b>.
                    </p>
                  )}
                </div>
              </div>

              <input type="file" ref={fileInputRef} className="hidden" accept="image/*" onChange={(e) => {
                  const file = e.target.files?.[0];
                  const currentUser = activeAuth.currentUser;
                  if (file && currentUser) {
                    setIsUploadingPhoto(true);
                    const reader = new FileReader();
                    reader.onload = (event) => {
                       const rawBase64 = event.target?.result as string;
                       // Optmistic preview
                       setAvatarUrl(rawBase64);
                       
                       const img = new Image();
                       img.onload = () => {
                         try {
                           const canvas = document.createElement('canvas');
                           const size = 256; // Standard avatar resolution
                           canvas.width = size;
                           canvas.height = size;
                           const ctx = canvas.getContext('2d');
                           if (ctx) {
                             const minDim = Math.min(img.width, img.height);
                             const sx = (img.width - minDim) / 2;
                             const sy = (img.height - minDim) / 2;
                             ctx.drawImage(img, sx, sy, minDim, minDim, 0, 0, size, size);
                             
                             const compressedBase64 = canvas.toDataURL('image/jpeg', 0.85);
                             
                             // Update local state and cache immediately
                             setAvatarUrl(compressedBase64);
                             setIsUploadingPhoto(false);
                             showToast('Poto Profil Berhasil di simpan', 'success');
                             
                             // Background cloud save tasks
                             updateProfile(currentUser, { photoURL: compressedBase64 })
                               .catch(authErr => console.warn('Gagal memperbarui foto profil auth:', authErr));
                             
                             // Backup disabled
                             /*
                             if (activeUserCustomData?.username) {
                               setDoc(doc(dbDefault, 'custom_accounts', activeUserCustomData?.username?.toLowerCase().trim() || ''), {
                                 photoURL: compressedBase64
                               }, { merge: true }).catch(err => console.warn('Mencadangkan foto profil ke tabel pusat gagal:', err));
                             }
                             */
                             
                             // Save to private database (does not block or freeze if user Rules are custom/failing)
                             setDoc(doc(activeDb, 'users', currentUser.uid), { photoURL: compressedBase64 }, { merge: true })
                               .catch(err => console.warn('Mencadangkan foto profil ke database mandiri gagal:', err));
                           } else {
                             // Fallback if canvas context extraction fails
                             setAvatarUrl(rawBase64);
                             setIsUploadingPhoto(false);
                             showToast('Poto Profil Berhasil di simpan', 'success');
                             
                             updateProfile(currentUser, { photoURL: rawBase64 })
                               .catch(authErr => console.warn('Gagal memperbarui foto profil auth:', authErr));
                             
                             // Backup disabled
                             /*
                             if (activeUserCustomData?.username) {
                               setDoc(doc(dbDefault, 'custom_accounts', activeUserCustomData?.username?.toLowerCase().trim() || ''), {
                                 photoURL: rawBase64
                               }, { merge: true }).catch(err => console.warn('Mencadangkan foto profil ke tabel pusat gagal:', err));
                             }
                             */
                             
                             setDoc(doc(activeDb, 'users', currentUser.uid), { photoURL: rawBase64 }, { merge: true })
                               .catch(err => console.warn('Mencadangkan foto profil ke database mandiri gagal:', err));
                           }
                         } catch (err) {
                           console.error('Error compressing image:', err);
                           showToast('Gagal memperbarui foto profil.', 'error');
                           setIsUploadingPhoto(false);
                         }
                       };
                       img.onerror = () => {
                         setIsUploadingPhoto(false);
                         showToast('File gambar tidak valid.', 'error');
                       };
                       // Set src AFTER attaching onload/onerror to secure fast cache/WebView loads
                       img.src = rawBase64;
                    };
                    reader.onerror = () => {
                      setIsUploadingPhoto(false);
                      showToast('Gagal membaca gambar.', 'error');
                    };
                    reader.readAsDataURL(file);
                  }
              }} />

              <div className="w-full h-px bg-slate-100 mb-6"></div>

              <div className="w-full mb-8 text-left">
                <div className="flex justify-between items-center mb-5">
                  <h3 className="text-lg font-bold text-slate-800">Detail Akademik</h3>
                  {!isProfileEditing ? (
                    <button 
                      onClick={() => setIsProfileEditing(true)}
                      className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition-colors"
                    >
                      Edit Data
                    </button>
                  ) : null}
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-4 gap-y-5">
                  <div className="space-y-1.5">
                    <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Nama Guru</label>
                    <input type="text" disabled={!isProfileEditing} value={profileData.namaGuruMapel} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, namaGuruMapel: e.target.value }))} className="w-full px-4 py-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 transition-all text-sm font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50 disabled:text-slate-600 disabled:border-slate-100" />
                  </div>
                  
                  <div className="space-y-1.5">
                    <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">NIP Guru</label>
                    <input type="text" disabled={!isProfileEditing} value={profileData.nipGuruMapel} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, nipGuruMapel: e.target.value }))} className="w-full px-4 py-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 transition-all text-sm font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50 disabled:text-slate-600 disabled:border-slate-100" />
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Peran (Role)</label>
                    <select disabled={!isProfileEditing} value={profileData.role || 'Guru Mapel'} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, role: e.target.value }))} className="w-full px-4 py-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 transition-all text-sm font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50 disabled:text-slate-600 disabled:border-slate-100">
                      <option value="Guru Mapel">Guru Mata Pelajaran</option>
                      <option value="Wali Kelas">Wali Kelas</option>
                    </select>
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Mata Pelajaran</label>
                    <input type="text" disabled={!isProfileEditing || profileData.role === 'Wali Kelas'} value={profileData.mataPelajaran} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, mataPelajaran: e.target.value }))} className="w-full px-4 py-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 transition-all text-sm font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50 disabled:text-slate-600 disabled:border-slate-100" />
                  </div>

                  {profileData.role === 'Wali Kelas' && (
                    <>
                      <div className="space-y-1.5">
                        <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Kelas yang Diwalikan</label>
                        <select disabled={!isProfileEditing} value={profileData.waliKelasClass || ''} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, waliKelasClass: e.target.value }))} className="w-full px-4 py-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 transition-all text-sm font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50 disabled:text-slate-600 disabled:border-slate-100">
                          <option value="">-- Pilih Kelas --</option>
                          {classList.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </div>

                      <div className="space-y-1.5">
                        <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Jumlah Siswa Laki-laki (L)</label>
                        <input 
                          type="number" 
                          min="0" 
                          disabled={!isProfileEditing} 
                          value={profileData.jumlahSiswaLakiLaki || ''} 
                          onChange={e => setProfileData((p: typeof profileData) => ({ ...p, jumlahSiswaLakiLaki: e.target.value }))} 
                          placeholder="Contoh: 18" 
                          className="w-full px-4 py-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 transition-all text-sm font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50 disabled:text-slate-600 disabled:border-slate-100" 
                        />
                      </div>

                      <div className="space-y-1.5">
                        <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Jumlah Siswa Perempuan (P)</label>
                        <input 
                          type="number" 
                          min="0" 
                          disabled={!isProfileEditing} 
                          value={profileData.jumlahSiswaPerempuan || ''} 
                          onChange={e => setProfileData((p: typeof profileData) => ({ ...p, jumlahSiswaPerempuan: e.target.value }))} 
                          placeholder="Contoh: 18" 
                          className="w-full px-4 py-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 transition-all text-sm font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50 disabled:text-slate-600 disabled:border-slate-100" 
                        />
                      </div>
                    </>
                  )}

                  <div className="space-y-1.5">
                    <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Semester</label>
                    <select disabled={!isProfileEditing} value={profileData.semester} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, semester: e.target.value }))} className="w-full px-4 py-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 transition-all text-sm font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50 disabled:text-slate-600 disabled:border-slate-100">
                      <option value="Ganjil">Ganjil</option>
                      <option value="Genap">Genap</option>
                    </select>
                  </div>

                  <div className="space-y-1.5 sm:col-span-2">
                    <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Tahun Pelajaran</label>
                    <input type="text" disabled={!isProfileEditing} value={profileData.tahunPelajaran} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, tahunPelajaran: e.target.value }))} placeholder="Contoh: 2026/2027" className="w-full px-4 py-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 transition-all text-sm font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50 disabled:text-slate-600 disabled:border-slate-100" />
                  </div>
                </div>

                {/* Section Pejabat Penandatangan Laporan (Mengetahui) */}
                <div className="mt-8 pt-6 border-t-2 border-dashed border-slate-200">
                  <div className="mb-4">
                    <div className="flex items-center gap-2 mb-1">
                      <span className="px-2.5 py-0.5 rounded-full bg-emerald-100 text-emerald-800 text-[10px] font-bold uppercase tracking-wider">Opsi Tanda Tangan</span>
                      <h4 className="text-sm font-bold text-slate-800">Pejabat Penandatangan Laporan (Mengetahui - Kiri)</h4>
                    </div>
                    <p className="text-xs text-slate-500">
                      Lengkapi nama dan NIP pejabat di bawah ini. Anda dapat memilih salah satu pejabat saat mencetak Laporan Presensi atau Laporan Wali Kelas.
                    </p>
                  </div>

                  <div className="space-y-4">
                    {/* Kepala Sekolah */}
                    <div className="p-4 bg-slate-50 border-2 border-slate-200/80 rounded-2xl">
                      <div className="text-xs font-bold text-slate-700 mb-3 flex items-center gap-1.5">
                        <span className="w-2 h-2 rounded-full bg-emerald-500"></span>
                        1. Kepala Sekolah
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Nama Kepala Sekolah</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.namaKepalaSekolah} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, namaKepalaSekolah: e.target.value }))} placeholder="Nama lengkap beserta gelar" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">NIP Kepala Sekolah</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.nipKepalaSekolah} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, nipKepalaSekolah: e.target.value }))} placeholder="Nomor Induk Pegawai" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                      </div>
                    </div>

                    {/* Pihak Kurikulum */}
                    <div className="p-4 bg-slate-50 border-2 border-slate-200/80 rounded-2xl">
                      <div className="text-xs font-bold text-slate-700 mb-3 flex items-center gap-1.5">
                        <span className="w-2 h-2 rounded-full bg-blue-500"></span>
                        2. Pihak Kurikulum (Wakasek Kurikulum)
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Nama Pihak Kurikulum</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.namaKurikulum || ''} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, namaKurikulum: e.target.value }))} placeholder="Nama pejabat kurikulum" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">NIP Pihak Kurikulum</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.nipKurikulum || ''} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, nipKurikulum: e.target.value }))} placeholder="NIP pejabat kurikulum" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                      </div>
                    </div>

                    {/* Pihak Kesiswaan */}
                    <div className="p-4 bg-slate-50 border-2 border-slate-200/80 rounded-2xl">
                      <div className="text-xs font-bold text-slate-700 mb-3 flex items-center gap-1.5">
                        <span className="w-2 h-2 rounded-full bg-amber-500"></span>
                        3. Pihak Kesiswaan (Wakasek Kesiswaan)
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Nama Pihak Kesiswaan</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.namaKesiswaan || ''} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, namaKesiswaan: e.target.value }))} placeholder="Nama pejabat kesiswaan" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">NIP Pihak Kesiswaan</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.nipKesiswaan || ''} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, nipKesiswaan: e.target.value }))} placeholder="NIP pejabat kesiswaan" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                      </div>
                    </div>

                    {/* Guru Wali */}
                    <div className="p-4 bg-slate-50 border-2 border-slate-200/80 rounded-2xl">
                      <div className="text-xs font-bold text-slate-700 mb-3 flex items-center gap-1.5">
                        <span className="w-2 h-2 rounded-full bg-purple-500"></span>
                        4. Guru Wali / Koordinator Guru Wali
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Nama Guru Wali</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.namaGuruWali || ''} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, namaGuruWali: e.target.value }))} placeholder="Nama guru wali" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">NIP Guru Wali</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.nipGuruWali || ''} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, nipGuruWali: e.target.value }))} placeholder="NIP guru wali" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                      </div>
                    </div>

                    {/* Guru BK */}
                    <div className="p-4 bg-slate-50 border-2 border-slate-200/80 rounded-2xl">
                      <div className="text-xs font-bold text-slate-700 mb-3 flex items-center gap-1.5">
                        <span className="w-2 h-2 rounded-full bg-teal-500"></span>
                        5. Guru BK / Bimbingan Konseling
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Nama Guru BK</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.namaBK || ''} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, namaBK: e.target.value }))} placeholder="Nama guru BK" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">NIP Guru BK</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.nipBK || ''} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, nipBK: e.target.value }))} placeholder="NIP guru BK" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                      </div>
                    </div>

                    {/* Wakasek Humas */}
                    <div className="p-4 bg-slate-50 border-2 border-slate-200/80 rounded-2xl">
                      <div className="text-xs font-bold text-slate-700 mb-3 flex items-center gap-1.5">
                        <span className="w-2 h-2 rounded-full bg-orange-500"></span>
                        6. Wakil Kepala Sekolah Bidang Humas (Hubungan Masyarakat)
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Nama Wakasek Humas</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.namaHumas || ''} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, namaHumas: e.target.value }))} placeholder="Nama pejabat humas" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">NIP Wakasek Humas</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.nipHumas || ''} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, nipHumas: e.target.value }))} placeholder="NIP pejabat humas" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="w-full bg-white border border-slate-100 rounded-[1.5rem] p-6 shadow-sm mb-8 text-left space-y-6">
                  {/* Header Statistik */}
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-100 pb-5">
                    <div className="flex items-center gap-3">
                      <div className="p-2.5 bg-blue-50 rounded-2xl text-blue-600 shadow-xs">
                        <Activity className="w-6 h-6" />
                      </div>
                      <div>
                        <h4 className="text-base font-black text-slate-800 tracking-tight">Statistik Pemakaian Limit</h4>
                        <p className="text-xs text-slate-500 font-medium">Pantau kuota operasional Firestore Anda secara real-time.</p>
                      </div>
                    </div>

                    {/* Live Clock & WIB Badge */}
                    <div className="flex items-center gap-2 bg-slate-50 border border-slate-200/80 px-3.5 py-2 rounded-2xl shrink-0">
                      <Clock className="w-4 h-4 text-blue-600 animate-pulse" />
                      <div className="text-right">
                        <div className="text-[11px] font-black text-slate-800 font-mono tracking-wide">{cycleInfo.currentTimeStr}</div>
                        <div className="text-[9px] font-bold text-slate-500 uppercase tracking-wider">Waktu Indonesia Barat</div>
                      </div>
                    </div>
                  </div>

                  {/* Banner Hari, Tanggal Sekarang & Jadwal Reset 14.00 WIB */}
                  <div className="bg-gradient-to-r from-blue-50/80 via-emerald-50/60 to-slate-50 rounded-2xl p-4.5 border border-blue-100/70 text-left space-y-3">
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                      {/* Hari & Tanggal Sekarang */}
                      <div className="flex items-start gap-3 bg-white/80 backdrop-blur-xs p-3 rounded-xl border border-blue-100/50 shadow-2xs">
                        <div className="p-2 bg-blue-100/70 text-blue-700 rounded-lg shrink-0 mt-0.5">
                          <Calendar className="w-4 h-4" />
                        </div>
                        <div>
                          <div className="text-[10px] font-black uppercase tracking-wider text-slate-500">Hari & Tanggal Sekarang</div>
                          <div className="text-xs sm:text-sm font-black text-slate-800 mt-0.5">{cycleInfo.currentDateFormatted}</div>
                        </div>
                      </div>

                      {/* Jadwal Reset Harian 14.00 WIB */}
                      <div className="flex items-start gap-3 bg-white/80 backdrop-blur-xs p-3 rounded-xl border border-emerald-100/50 shadow-2xs">
                        <div className="p-2 bg-emerald-100/70 text-[#7bc025] rounded-lg shrink-0 mt-0.5">
                          <RefreshCw className="w-4 h-4" />
                        </div>
                        <div>
                          <div className="text-[10px] font-black uppercase tracking-wider text-slate-500">Jadwal Reset Limit (00:00 PST / 14:00 WIB)</div>
                          <div className="text-xs sm:text-sm font-black text-slate-800 mt-0.5">
                            Setiap Hari Pukul <span className="text-[#7bc025]">14.00 WIB</span>
                          </div>
                        </div>
                      </div>
                    </div>

                    {/* Countdown & Info Strip */}
                    <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2 pt-1 px-1">
                      <div className="flex items-center gap-1.5 text-xs text-slate-600 font-medium">
                        <Timer className="w-4 h-4 text-blue-500 shrink-0" />
                        <span>Sisa waktu menuju reset pukul 14:00 WIB:</span>
                        <span className="font-mono font-black text-blue-700 bg-blue-100/60 px-2 py-0.5 rounded-md text-[11px]">{cycleInfo.countdownFormatted}</span>
                      </div>
                      <div className="inline-flex items-center gap-1 text-[10px] text-emerald-800 font-bold bg-emerald-100/80 px-2.5 py-1 rounded-full">
                        <CheckCircle className="w-3 h-3 text-[#8dc63f]" />
                        <span>Reset Otomatis ke 0 Aktif</span>
                      </div>
                    </div>
                  </div>
                  
                  {(() => {
                    const writeLimit = 20000;
                    const readLimit = 50000;
                    
                    const writeUsed = sessionUsage.writes || 0;
                    const writeRemaining = Math.max(0, writeLimit - writeUsed);
                    const writePercentage = Math.min(100, (writeUsed / writeLimit) * 100);
                    
                    const readUsed = sessionUsage.reads || 0;
                    const readRemaining = Math.max(0, readLimit - readUsed);
                    const readPercentage = Math.min(100, (readUsed / readLimit) * 100);
                    
                    return (
                      <div className="space-y-6 pt-2">
                        {/* Write Stats */}
                        <div className="space-y-3 bg-slate-50/50 p-4 rounded-2xl border border-slate-100">
                          <div className="flex items-center justify-between border-b border-slate-200/60 pb-2">
                            <h5 className="text-xs font-black text-slate-800">Batas Operasi Penulisan (Writes)</h5>
                            <span className="text-[10px] font-bold text-slate-500 bg-white px-2 py-0.5 rounded-md border border-slate-200">
                              Limit: {writeLimit.toLocaleString('id-ID')} / hari (Reset 14:00 WIB)
                            </span>
                          </div>
                          
                          <div className="flex justify-between items-end">
                            <div>
                              <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Total Digunakan Sesi Ini</p>
                              <p className="text-xl font-black text-slate-800">{writeUsed.toLocaleString('id-ID')}</p>
                            </div>
                            <div className="text-right">
                              <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Sisa Kuota</p>
                              <p className="text-base font-black text-blue-600">{writeRemaining.toLocaleString('id-ID')}</p>
                            </div>
                          </div>

                          <div className="w-full bg-slate-200/70 h-2.5 rounded-full overflow-hidden">
                            <div 
                              className={`h-full rounded-full transition-all duration-500 ${writePercentage > 90 ? 'bg-rose-500' : writePercentage > 75 ? 'bg-amber-400' : 'bg-blue-500'}`}
                              style={{ width: `${writePercentage}%` }}
                            />
                          </div>
                          
                          <p className="text-[10px] text-slate-600 leading-relaxed bg-white p-2.5 rounded-xl border border-slate-200/60 mt-1">
                            Batas maksimal operasi penulisan perangkat Anda adalah <b>{writeLimit.toLocaleString('id-ID')}</b>. Kuota tersisa <b>{writeRemaining.toLocaleString('id-ID')}</b>. Setiap Anda menyimpan data, kuota penulisan berkurang dan akan direset ke 0 setiap pukul <b>14.00 WIB</b>.
                          </p>
                        </div>

                        {/* Read Stats */}
                        <div className="space-y-3 bg-slate-50/50 p-4 rounded-2xl border border-slate-100">
                          <div className="flex items-center justify-between border-b border-slate-200/60 pb-2">
                            <h5 className="text-xs font-black text-slate-800">Batas Operasi Pembacaan (Reads)</h5>
                            <span className="text-[10px] font-bold text-slate-500 bg-white px-2 py-0.5 rounded-md border border-slate-200">
                              Limit: {readLimit.toLocaleString('id-ID')} / hari (Reset 14:00 WIB)
                            </span>
                          </div>

                          <div className="flex justify-between items-end">
                            <div>
                              <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Total Digunakan Sesi Ini</p>
                              <p className="text-xl font-black text-slate-800">{readUsed.toLocaleString('id-ID')}</p>
                            </div>
                            <div className="text-right">
                              <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Sisa Kuota</p>
                              <p className="text-base font-black text-[#7bc025]">{readRemaining.toLocaleString('id-ID')}</p>
                            </div>
                          </div>

                          <div className="w-full bg-slate-200/70 h-2.5 rounded-full overflow-hidden">
                            <div 
                              className={`h-full rounded-full transition-all duration-500 ${readPercentage > 90 ? 'bg-rose-500' : readPercentage > 75 ? 'bg-amber-400' : 'bg-[#8dc63f]'}`}
                              style={{ width: `${readPercentage}%` }}
                            />
                          </div>
                          
                          <p className="text-[10px] text-slate-600 leading-relaxed bg-white p-2.5 rounded-xl border border-slate-200/60 mt-1">
                            Batas maksimal operasi pembacaan perangkat Anda adalah <b>{readLimit.toLocaleString('id-ID')}</b>. Kuota tersisa <b>{readRemaining.toLocaleString('id-ID')}</b>. Setiap Anda memuat aplikasi atau data, kuota pembacaan berkurang dan akan direset ke 0 setiap pukul <b>14.00 WIB</b>.
                          </p>
                        </div>
                      </div>
                    );
                  })()}
                </div>

                <div className="w-full mt-8">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 mb-4">
                    <div>
                      <h3 className="text-lg font-black text-slate-800">Rekapitulasi Penggunaan Limit Harian Semua Akun</h3>
                      <p className="text-xs text-slate-500 font-medium">Statistik pemakaian limit per user dalam siklus harian berjalan (Reset setiap 14.00 WIB).</p>
                    </div>
                    <span className="self-start sm:self-auto text-[10px] font-bold bg-blue-50 text-blue-700 border border-blue-200/70 px-2.5 py-1 rounded-lg">
                      Siklus: 14:00 WIB
                    </span>
                  </div>

                  <div className="bg-white rounded-2xl border border-slate-100 shadow-sm overflow-hidden">
                    <div className="overflow-x-auto">
                      <table className="w-full text-left border-collapse">
                        <thead>
                          <tr className="bg-slate-50 border-b border-slate-100/80">
                            <th className="px-4 py-3 text-[10px] font-bold text-slate-600 uppercase tracking-wider whitespace-nowrap">Nama Pengguna</th>
                            <th className="px-4 py-3 text-[10px] font-bold text-slate-600 uppercase tracking-wider text-right">Pembacaan (Reads)</th>
                            <th className="px-4 py-3 text-[10px] font-bold text-slate-600 uppercase tracking-wider text-right">Penulisan (Writes)</th>
                            <th className="px-4 py-3 text-[10px] font-bold text-slate-600 uppercase tracking-wider text-center whitespace-nowrap">Status Siklus</th>
                            <th className="px-4 py-3 text-[10px] font-bold text-slate-600 uppercase tracking-wider text-right whitespace-nowrap">Tanggal Catatan</th>
                          </tr>
                        </thead>
                        <tbody>
                          {isLoadingUsage ? (
                            <tr>
                              <td colSpan={5} className="px-4 py-8 text-center text-sm font-medium text-slate-500">Loading data statistik pengguna...</td>
                            </tr>
                          ) : allUsersUsage.length > 0 ? (
                            allUsersUsage.map((usage, idx) => (
                              <tr key={usage.username} className={`border-b border-slate-50 last:border-none ${idx % 2 === 0 ? 'bg-white' : 'bg-slate-50/50'}`}>
                                <td className="px-4 py-3">
                                  <div className="text-sm font-bold text-slate-800">{usage.fullname}</div>
                                  <div className="text-[10px] font-medium text-slate-500">{usage.username}</div>
                                </td>
                                <td className="px-4 py-3 text-right">
                                  <span className={`inline-flex items-center px-2 py-0.5 rounded-md text-xs font-bold ${usage.reads > 10000 ? 'bg-rose-100 text-rose-700' : usage.reads > 5000 ? 'bg-amber-100 text-amber-700' : 'bg-emerald-100 text-[#7bc025]'}`}>
                                    {usage.reads.toLocaleString('id-ID')}
                                  </span>
                                </td>
                                <td className="px-4 py-3 text-right">
                                  <span className={`inline-flex items-center px-2 py-0.5 rounded-md text-xs font-bold ${usage.writes > 5000 ? 'bg-rose-100 text-rose-700' : usage.writes > 2000 ? 'bg-amber-100 text-amber-700' : 'bg-blue-100 text-blue-700'}`}>
                                    {usage.writes.toLocaleString('id-ID')}
                                  </span>
                                </td>
                                <td className="px-4 py-3 text-center">
                                  {usage.isCurrentCycle ? (
                                    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-50 text-emerald-700 border border-emerald-200/60">
                                      <span className="w-1.5 h-1.5 rounded-full bg-[#8dc63f]"></span>
                                      Siklus Aktif
                                    </span>
                                  ) : (
                                    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold bg-slate-100 text-slate-600 border border-slate-200/60">
                                      Reset 14:00 WIB
                                    </span>
                                  )}
                                </td>
                                <td className="px-4 py-3 text-right">
                                  <div className="text-[10px] font-semibold text-slate-700">{usage.date}</div>
                                  {usage.time && <div className="text-[9px] font-medium text-slate-400">{usage.time}</div>}
                                </td>
                              </tr>
                            ))
                          ) : (
                            <tr>
                              <td colSpan={5} className="px-4 py-8 text-center text-sm font-medium text-slate-500">Belum ada data penggunaan tercatat.</td>
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>

                {isProfileEditing && (
                  <div className="flex justify-end gap-3 mt-6">
                    <button 
                      onClick={() => {
                        setIsProfileEditing(false);
                      }}
                      disabled={isProfileSaving}
                      className="px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-bold justify-center rounded-xl transition-all"
                    >
                      Batal
                    </button>
                    <button 
                      onClick={async () => {
                        if (!activeAuth.currentUser) return;
                        setIsProfileSaving(true);
                        trackOp('write', 1);
                        try {
                          const cleanProfileData = {
                            namaGuruMapel: profileData.namaGuruMapel || '',
                            namaKepalaSekolah: profileData.namaKepalaSekolah || '',
                            nipGuruMapel: profileData.nipGuruMapel || '',
                            nipKepalaSekolah: profileData.nipKepalaSekolah || '',
                            namaKurikulum: profileData.namaKurikulum || '',
                            nipKurikulum: profileData.nipKurikulum || '',
                            jabatanKurikulum: profileData.jabatanKurikulum || 'Wakasek Kurikulum',
                            namaKesiswaan: profileData.namaKesiswaan || '',
                            nipKesiswaan: profileData.nipKesiswaan || '',
                            jabatanKesiswaan: profileData.jabatanKesiswaan || 'Wakasek Kesiswaan',
                            namaGuruWali: profileData.namaGuruWali || '',
                            nipGuruWali: profileData.nipGuruWali || '',
                            jabatanGuruWali: profileData.jabatanGuruWali || 'Guru Wali',
                            namaBK: profileData.namaBK || '',
                            nipBK: profileData.nipBK || '',
                            jabatanBK: profileData.jabatanBK || 'Guru BK',
                            namaHumas: profileData.namaHumas || '',
                            nipHumas: profileData.nipHumas || '',
                            jabatanHumas: profileData.jabatanHumas || 'Wakasek Humas',
                            semester: profileData.semester || 'Ganjil',
                            tahunPelajaran: profileData.tahunPelajaran || '',
                            mataPelajaran: profileData.mataPelajaran || '',
                            role: profileData.role || 'Guru Mapel',
                            waliKelasClass: profileData.waliKelasClass || '',
                            jumlahSiswaLakiLaki: profileData.jumlahSiswaLakiLaki || 0,
                            jumlahSiswaPerempuan: profileData.jumlahSiswaPerempuan || 0
                          };

                          const savePayload = {
                            profileData: cleanProfileData, // Correct nested format
                            ...cleanProfileData
                          };
                          
                          // Save to activeDb (private database) but gracefully handle custom firestore permission rules
                          const savePromise = setDoc(doc(activeDb, 'users', activeAuth.currentUser.uid), savePayload, { merge: true })
                            .catch(err => {
                              const msg = err instanceof Error ? err.message : String(err);
                              if (msg.toLowerCase().includes('permission') || msg.toLowerCase().includes('insufficient')) {
                                console.warn('Gagal menyimpan profil ke database mandiri (Rules ditolak). Mengandalkan pencadangan pusat kaguci gratis.');
                                return; // Let the promise resolve to proceed seamlessly with the central backup
                              }
                              throw err;
                            });
                          
                          // Backup disabled to save quota
                          if (activeUserCustomData?.username) {
                            // Backup disabled
                          }
                          
                          
                          // Race against a short timeout to guarantee instant performance (Firestore syncs in background anyway)
                          const timeoutPromise = new Promise<void>((_, reject) => 
                            setTimeout(() => reject(new Error('timeout')), 1500)
                          );
                          
                          try {
                            await Promise.race([savePromise, timeoutPromise]);
                          } catch (raceError) {
                            if (raceError instanceof Error && raceError.message === 'timeout') {
                              console.warn("Profile save server-sync timed out. Proceeding since local cache is updated.");
                            } else {
                              throw raceError;
                            }
                          }
                          
                          showToast('Profile Berhasil Disimpan', 'success');
                          setIsProfileEditing(false);
                        } catch (err) {
                          console.error('Error saving profile:', err);
                          handleFirestoreError(err, OperationType.WRITE, 'users');
                          showToast('Gagal menyimpan profil: ' + (err instanceof Error ? err.message : 'Server error'), 'error');
                        } finally {
                          setIsProfileSaving(false);
                        }
                      }}
                      disabled={isProfileSaving}
                      className="px-5 py-2.5 bg-[#8dc63f] hover:bg-[#7bc025] text-white text-sm font-bold justify-center rounded-xl shadow-[0_4px_12px_rgba(5,150,105,0.3)] transition-all flex items-center gap-2 disabled:opacity-70 disabled:shadow-none"
                    >
                      {isProfileSaving ? (
                        <span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></span>
                      ) : null}
                      Simpan Data
                    </button>
                  </div>
                )}
              </div>

              {/* KUMPULAN USER (ADMIN PANEL) - Custom Requested */}
              {isAdmin && (
                <>
                  <div className="w-full h-px bg-slate-100 mb-6"></div>
                  <div className="w-full mb-8 text-left bg-white border border-slate-150 rounded-[1.5rem] p-6 shadow-sm">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4">
                      <div className="flex items-center gap-2.5">
                        <div className="p-2 bg-emerald-50 rounded-lg text-[#8dc63f]">
                          <Users className="w-5 h-5" />
                        </div>
                        <div>
                          <h3 className="text-base font-extrabold text-slate-800 tracking-tight">Kumpulan User (Sistem Administrator)</h3>
                          <p className="text-[10px] text-slate-500">Kelola dan lihat seluruh akun guru/pengajar terpusat.</p>
                        </div>
                      </div>
                      <button
                        onClick={fetchAllUsers}
                        className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition-all self-start sm:self-center"
                      >
                        Perbarui Daftar
                      </button>
                    </div>

                    <div className="bg-slate-50/50 border border-slate-100 rounded-2xl overflow-hidden shadow-xs">
                      {isLoadingUsers ? (
                        <div className="p-12 flex flex-col items-center justify-center gap-2.5 text-slate-500">
                          <span className="w-6 h-6 border-2 border-[#8dc63f] border-t-transparent rounded-full animate-spin"></span>
                          <span className="text-[10px] font-bold uppercase tracking-wider text-slate-550">Memuat User Terdaftar...</span>
                        </div>
                      ) : allUsers.length === 0 ? (
                        <div className="p-8 text-center text-xs text-slate-500 font-medium">
                          Tidak ada user terdaftar atau gagal memuat data.
                        </div>
                      ) : (
                        <div className="overflow-x-auto">
                          <table className="w-full text-left border-collapse">
                            <thead>
                              <tr className="bg-slate-100/60 border-b border-slate-200/50">
                                <th className="px-4 py-3 text-[10px] font-extrabold text-slate-600 uppercase tracking-widest text-center w-12">No</th>
                                <th className="px-4 py-3 text-[10px] font-extrabold text-slate-600 uppercase tracking-widest">Nama Pengguna</th>
                                <th className="px-4 py-3 text-[10px] font-extrabold text-slate-600 uppercase tracking-widest">Username</th>
                                <th className="px-4 py-3 text-[10px] font-extrabold text-slate-600 uppercase tracking-widest">Password</th>
                                <th className="px-4 py-3 text-[10px] font-extrabold text-slate-600 uppercase tracking-widest text-center w-28">Aksi (Edit / Hapus)</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-100">
                              {allUsers.map((u, idx) => {
                                const isSelf = u.username.toLowerCase().trim() === activeUserCustomData?.username?.toLowerCase().trim();
                                return (
                                  <tr key={u.id} className={`hover:bg-slate-100/40 transition-colors ${isSelf ? 'bg-emerald-50/30 font-semibold text-[#7bc025]' : 'text-slate-700'}`}>
                                    <td className="px-4 py-3 text-xs text-slate-500 font-mono text-center">{idx + 1}</td>
                                    <td className="px-4 py-3 text-xs font-semibold">
                                      <div className="flex flex-col gap-1.5">
                                        <div className="flex items-center gap-1.5">
                                          {u.fullname}
                                          {isSelf && (
                                            <span className="px-2 py-0.5 bg-emerald-100 text-emerald-800 rounded-full font-black text-[8px] uppercase tracking-wider inline-block">
                                              Anda
                                            </span>
                                          )}
                                        </div>
                                        <div className="flex items-center gap-1.5">
                                          <span className={`w-1.5 h-1.5 rounded-full ${isSelf && isOnline ? 'bg-[#8dc63f] animate-pulse' : 'bg-slate-300'}`}></span>
                                          <span className={`text-[9px] font-bold ${isSelf && isOnline ? 'text-[#8dc63f]' : 'text-slate-400'}`}>
                                            {isSelf && isOnline ? 'Online' : 'Offline'}
                                          </span>
                                        </div>
                                      </div>
                                    </td>
                                    <td className="px-4 py-3 text-xs font-mono font-bold text-slate-600">{u.username}</td>
                                    <td className="px-4 py-3 text-xs font-mono text-slate-600">{u.password}</td>
                                    <td className="px-4 py-3 text-xs text-center">
                                      <div className="flex items-center justify-center gap-1.5">
                                        <button
                                          onClick={() => handleEditUserClick(u)}
                                          className="p-2 rounded-xl text-indigo-600 hover:bg-indigo-50 hover:text-indigo-700 active:scale-90 transition-all cursor-pointer"
                                          title="Edit Nama/Sandi Pengguna"
                                        >
                                          <Pencil className="w-4 h-4" />
                                        </button>
                                        <button
                                          onClick={() => handleDeleteUserClick(u)}
                                          disabled={isSelf}
                                          className={`p-2 rounded-xl transition-all ${
                                            isSelf 
                                              ? 'text-slate-200 cursor-not-allowed opacity-50' 
                                              : 'text-rose-600 hover:bg-rose-50 hover:text-rose-700 active:scale-90 cursor-pointer'
                                          }`}
                                          title={isSelf ? 'Tidak dapat menghapus akun Anda sendiri' : 'Hapus Pengguna'}
                                        >
                                          <Trash2 className="w-4.5 h-4.5" />
                                        </button>
                                      </div>
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

                  <div className="w-full h-px bg-slate-100 mb-6"></div>
                  
                  <div className="w-full mb-8 text-left bg-indigo-50 border border-indigo-100 rounded-[1.5rem] p-6 shadow-sm">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4">
                      <div className="flex items-center gap-2.5">
                        <div className="p-2 bg-indigo-100 rounded-lg text-indigo-600">
                          <Key className="w-5 h-5" />
                        </div>
                        <div>
                          <h3 className="text-base font-extrabold text-slate-800 tracking-tight">Generator Token Pendaftaran</h3>
                          <p className="text-[10px] text-slate-600">Buat token khusus yang diperlukan saat registrasi user baru.</p>
                        </div>
                      </div>
                      <button
                        onClick={async () => {
                          const token = Math.random().toString(36).substring(2, 10).toUpperCase();
                          setGeneratedToken(token);
                          try {
                            await setDoc(doc(dbDefault, 'register_tokens', token), {
                              createdAt: new Date().toISOString(),
                              createdBy: activeAuth.currentUser?.uid || 'admin',
                              used: false
                            });
                            showToast('Token berhasil dibuat dan disimpan.', 'success');
                          } catch (err) {
                            console.error('Gagal menyimpan token:', err);
                            showToast('Gagal menyimpan token ke database pusat.', 'error');
                            try { handleFirestoreError(err, OperationType.WRITE, 'register_tokens'); } catch { /* ignore */ }
                          }
                        }}
                        className="px-4 py-2 flex gap-2 items-center justify-center bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold rounded-xl transition-all self-start sm:self-center"
                      >
                        <RefreshCw className="w-3.5 h-3.5" />
                        Generate Token Baru
                      </button>
                    </div>

                    {generatedToken && (
                      <div className="mt-4 p-4 bg-white border border-indigo-100 rounded-xl text-center space-y-2">
                         <p className="text-xs text-slate-600 font-semibold mb-1">Token Terakhir Anda:</p>
                         <div className="flex items-center justify-center gap-3">
                           <div className="text-3xl font-black font-mono text-indigo-700 tracking-[0.2em] bg-indigo-50/50 py-3 px-6 rounded-lg border border-dashed border-indigo-200">
                             {generatedToken}
                           </div>
                           <button
                             onClick={() => {
                               navigator.clipboard.writeText(generatedToken);
                               showToast('Token disalin ke clipboard!', 'success');
                             }}
                             className="p-3 bg-indigo-100 hover:bg-indigo-200 text-indigo-600 rounded-lg transition-colors border border-indigo-200 shadow-sm"
                             title="Salin Token"
                           >
                             <Copy className="w-6 h-6" />
                           </button>
                         </div>
                         <p className="text-[10px] text-indigo-500">Salin token ini dan berikan kepada guru/user yang ingin mendaftar.</p>
                      </div>
                    )}
                  </div>
                </>
              )}

              <div className="w-full h-px bg-slate-100 mb-6"></div>



              {/* Maintenance & Reset Data Section */}
              <div className="w-full mb-8 text-left bg-stone-50 border-2 border-slate-300 rounded-2xl p-5">
                <h3 className="text-sm font-black text-slate-800 uppercase tracking-widest mb-2 flex items-center gap-2">
                  <span className="w-2 h-2 rounded-full bg-rose-500 animate-pulse"></span>
                  Pusat Pemeliharaan &amp; Reset Data
                </h3>
                <p className="text-xs text-slate-600 mb-5 leading-relaxed">
                  Kelola database absensi Anda secara fleksibel. Lakukan pembersihan riwayat secara berkala untuk menyambut ajaran baru maupun reset total.
                </p>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                  <button
                    onClick={() => {
                      setResetModalType('new_semester');
                      setResetConfirmInput('');
                      setResetPasswordInput('');
                      setResetPasswordError('');
                      setNewSemesterChoice(profileData.semester === 'Ganjil' ? 'Genap' : 'Ganjil');
                      setNewTahunPelajaran(profileData.tahunPelajaran || '');
                    }}
                    className="flex flex-col items-start gap-1.5 p-4 bg-white hover:bg-emerald-50/50 border-2 border-slate-300 hover:border-emerald-300 rounded-xl transition-all text-left group cursor-pointer"
                  >
                    <span className="flex items-center gap-2 text-xs font-bold text-slate-700 group-hover:text-emerald-800 transition-colors">
                      <span className="p-1 px-1.5 bg-emerald-100 rounded text-emerald-800 text-[10px] font-black">SEM</span>
                      Reset Semester Baru
                    </span>
                    <span className="text-[10px] text-slate-500 leading-normal group-hover:text-emerald-700 transition-colors">
                      Hapus riwayat sesi absensi semester lalu dengan konfirmasi password. Data siswa &amp; kelas tetap aman.
                    </span>
                  </button>

                  <button
                    onClick={() => {
                      setResetModalType('new_year');
                      setResetConfirmInput('');
                      setResetPasswordInput('');
                      setResetPasswordError('');
                    }}
                    className="flex flex-col items-start gap-1.5 p-4 bg-white hover:bg-amber-50/50 border-2 border-slate-300 hover:border-amber-200 rounded-xl transition-all text-left group cursor-pointer"
                  >
                    <span className="flex items-center gap-2 text-xs font-bold text-slate-700 group-hover:text-amber-800 transition-colors">
                      <span className="p-1 px-1.5 bg-amber-100 rounded text-amber-700 text-[10px] font-black">TA</span>
                      Reset Tahun Ajaran Baru
                    </span>
                    <span className="text-[10px] text-slate-500 leading-normal group-hover:text-amber-600 transition-colors">
                      Hapus semua riwayat absensi, tapi tetap pertahankan seluruh biodata siswa &amp; daftar kelas.
                    </span>
                  </button>

                  <button
                    onClick={() => {
                      setResetModalType('everything');
                      setResetConfirmInput('');
                      setResetPasswordInput('');
                      setResetPasswordError('');
                    }}
                    className="flex flex-col items-start gap-1.5 p-4 bg-white hover:bg-rose-50/50 border-2 border-slate-300 hover:border-rose-200 rounded-xl transition-all text-left group cursor-pointer"
                  >
                    <span className="flex items-center gap-2 text-xs font-bold text-slate-700 group-hover:text-rose-800 transition-colors">
                      <span className="p-1 px-1.5 bg-rose-100 rounded text-rose-700 text-[10px] font-black">ALL</span>
                      Reset Semua Data
                    </span>
                    <span className="text-[10px] text-slate-500 leading-normal group-hover:text-rose-500 transition-colors">
                      Hapus seluruh daftar siswa, kelas, &amp; seluruh riwayat absensi secara total dan permanen.
                    </span>
                  </button>
                </div>
              </div>

              <div className="w-full h-px bg-slate-100 mb-6"></div>

              <button 
                  onClick={() => setShowLogoutConfirm(true)}
                  className="w-full flex items-center justify-center gap-2 bg-rose-50 text-rose-600 font-bold py-4 px-6 rounded-2xl hover:bg-rose-100 transition-all border border-rose-100"
              >
                  <LogOut className="w-5 h-5" />
                  Keluar dari Aplikasi
              </button>
            </div>
          </div>
        );
      default:
        return (
          <div className="flex flex-col items-center justify-center h-full p-8">
            <h2 className="text-xl font-bold text-slate-800">Konten untuk {activeTab} akan datang.</h2>
          </div>
        );
    }
  };
  
  // Define UI before return
  const splashNode = (
    <AnimatePresence mode="wait">
      {showSplash && (
        isFirstSessionLoad ? (
          <motion.div
            key="splash-green"
            initial={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            transition={{ duration: 0.8, ease: [0.25, 1, 0.5, 1] }}
            className="fixed inset-0 z-[9999] bg-[#8dc63f] flex flex-col items-center justify-center p-6 select-none"
          >
            {/* Ambient lighting */}
            <div className="absolute inset-0 bg-[radial-gradient(circle_at_center,rgba(255,255,255,0.18)_0%,transparent_70%)] pointer-events-none" />
            
            <motion.div
              initial={{ scale: 0.96, opacity: 0, y: 10 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              transition={{ duration: 1.0, ease: [0.16, 1, 0.3, 1] }}
              className="flex flex-col items-center relative z-10"
            >
              {/* School Logo Container */}
              <motion.div
                initial={{ scale: 0.85, opacity: 0, y: -10 }}
                animate={{ scale: 1, opacity: 1, y: 0 }}
                transition={{ delay: 0.15, duration: 0.8, ease: "easeOut" }}
                className="w-32 h-32 md:w-36 md:h-36 mb-6 flex items-center justify-center"
              >
                <img 
                  src="https://drive.google.com/thumbnail?id=1P395tuZymxs3qero4XduMpHy7g2GJrdR&sz=w1000" 
                  alt="My Kaguci Logo" 
                  className="w-full h-full object-contain filter drop-shadow-[0_12px_24px_rgba(0,0,0,0.18)]"
                  referrerPolicy="no-referrer"
                />
              </motion.div>

              {/* Micro divider */}
              <motion.div
                initial={{ scaleX: 0, opacity: 0 }}
                animate={{ scaleX: 1, opacity: 1 }}
                transition={{ delay: 0.25, duration: 0.8, ease: "easeOut" }}
                className="h-[1.5px] w-12 bg-white/40 rounded-full mb-6 origin-center"
              />

              <motion.h1
                initial={{ opacity: 0, y: 5 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: 0.35, duration: 0.8 }}
                className="text-white text-3xl sm:text-4xl font-extrabold tracking-tight text-center drop-shadow-xs"
              >
                My Kaguci App
              </motion.h1>

              <motion.p
                initial={{ y: 8, opacity: 0 }}
                animate={{ y: 0, opacity: 1 }}
                transition={{ delay: 0.5, duration: 0.8, ease: "easeOut" }}
                className="text-white/90 text-xs sm:text-[13px] font-medium uppercase tracking-[0.22em] mt-3.5 text-center"
              >
                Sistem Informasi Absensi Digital
              </motion.p>
                
              <motion.p
                initial={{ opacity: 0 }}
                animate={{ opacity: 0.6 }}
                transition={{ delay: 0.7, duration: 0.9 }}
                className="text-white/60 text-[9px] font-semibold tracking-widest mt-1.5 uppercase text-center"
              >
                Cerdas • Inovatif • Terampil • Responsif • Agamis
              </motion.p>

              {/* White micro progress bar */}
              <div className="w-48 h-[3px] bg-white/20 rounded-full mt-10 overflow-hidden relative">
                <motion.div
                  initial={{ left: "-100%" }}
                  animate={{ left: "100%" }}
                  transition={{
                    repeat: Infinity,
                    duration: 1.8,
                    ease: "easeInOut"
                  }}
                  className="absolute top-0 bottom-0 w-1/2 bg-gradient-to-r from-transparent via-white to-transparent"
                />
              </div>
            </motion.div>
          </motion.div>
        ) : (
          <motion.div
            key="splash-white"
            initial={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            transition={{ duration: 0.4, ease: "easeInOut" }}
            className="fixed inset-0 z-[9999] bg-white flex flex-col items-center justify-center p-6 select-none"
          >
            <motion.div
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              transition={{ duration: 0.5, ease: "easeOut" }}
              className="flex flex-col items-center relative z-10"
            >
              {/* School Logo Container */}
              <div className="w-24 h-24 sm:w-28 sm:h-28 mb-4 flex items-center justify-center">
                <img 
                  src="https://drive.google.com/thumbnail?id=1P395tuZymxs3qero4XduMpHy7g2GJrdR&sz=w1000" 
                  alt="My Kaguci Logo" 
                  className="w-full h-full object-contain filter drop-shadow-sm"
                  referrerPolicy="no-referrer"
                />
              </div>

              <h1 className="text-slate-800 text-xl sm:text-2xl font-extrabold tracking-tight text-center">
                My Kaguci App
              </h1>

              {/* Animated Memuat Halaman Badge */}
              <div className="flex items-center gap-2.5 mt-6 bg-slate-50 border border-slate-200/80 px-5 py-2 rounded-full shadow-xs">
                <Loader2 className="w-4 h-4 text-[#8dc63f] animate-spin shrink-0" />
                <motion.p
                  animate={{ opacity: [0.4, 1, 0.4] }}
                  transition={{ repeat: Infinity, duration: 1.5, ease: "easeInOut" }}
                  className="text-slate-700 font-extrabold text-xs tracking-wider uppercase"
                >
                  Memuat Halaman...
                </motion.p>
              </div>

              <div className="w-48 h-1 bg-slate-100 rounded-full mt-5 overflow-hidden relative border border-slate-200/60">
                <motion.div
                  initial={{ left: "-100%" }}
                  animate={{ left: "100%" }}
                  transition={{
                    repeat: Infinity,
                    duration: 1.6,
                    ease: "easeInOut"
                  }}
                  className="absolute top-0 bottom-0 w-1/2 bg-[#8dc63f] rounded-full"
                />
              </div>
            </motion.div>
          </motion.div>
        )
      )}
    </AnimatePresence>
  );

  const GlobalConnectivityBanner = () => (
    <AnimatePresence>
      {(!isOnline || !firebaseConnected || syncStatus === 'error' || quotaExceeded) && (
        <motion.div 
          initial={{ height: 0, opacity: 0 }}
          animate={{ height: 'auto', opacity: 1 }}
          exit={{ height: 0, opacity: 0 }}
          className="overflow-hidden z-[100] sticky top-0 shrink-0"
        >
          {/* Main Sync Error / Offline Banner */}
          {(!isOnline || !firebaseConnected || syncStatus === 'error') && (
            <div className="bg-rose-600 text-white shadow-lg">
              <div className="max-w-7xl mx-auto px-4 py-3 flex items-center justify-between gap-4">
                <div className="flex items-center gap-3">
                  <div className="p-2 bg-white/20 rounded-xl animate-pulse">
                    <Cloud className="w-5 h-5" />
                  </div>
                  <div>
                    <h4 className="text-sm font-black leading-tight flex items-center gap-2">
                       {!isOnline ? 'Koneksi Internet Terputus' : 'Sinkronisasi Cloud Bermasalah'}
                       <span className="px-1.5 py-0.5 bg-rose-500 rounded text-[9px] uppercase tracking-widest">{!isOnline ? 'OFFLINE' : 'ERR_CLOUD'}</span>
                    </h4>
                    <p className="text-[11px] opacity-90 font-medium mt-0.5">
                      {!isOnline 
                        ? 'Periksa koneksi Wi-Fi atau data seluler Anda. Data yang Anda buat saat ini tersimpan sementara di peramban ini dan akan diunggah otomatis saat kembali online.' 
                        : lastSyncError || 'Terjadi gangguan sinkronisasi dengan database Firebase. Harap periksa izin akses rules Firestore atau muat ulang halaman.'}
                    </p>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <button 
                    onClick={() => window.location.reload()}
                    className="px-3 py-2 bg-white text-rose-600 rounded-xl text-[10px] font-black uppercase tracking-wider hover:bg-rose-50 transition-all shadow-sm active:scale-95 shrink-0"
                  >
                    Muat Ulang
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* Quota Banner */}
          {quotaExceeded && isOnline && firebaseConnected && syncStatus !== 'error' && (
            <div className="bg-amber-500 text-white shadow-lg">
               <div className="max-w-7xl mx-auto px-4 py-2.5 flex items-center justify-between gap-4">
                  <div className="flex items-center gap-3">
                    <Database className="w-4 h-4" />
                    <p className="text-xs font-bold leading-tight">
                      Kuota harian database tercapai. Beberapa pembaruan data mungkin tertunda masuk ke Cloud sampai besok.
                    </p>
                  </div>
                  <button 
                    onClick={() => setShowQuotaModal(true)}
                    className="px-3 py-1.5 bg-white/20 rounded-lg text-[10px] font-bold uppercase transition-colors hover:bg-white/30"
                  >
                    Info Lanjut
                  </button>
               </div>
            </div>
          )}
        </motion.div>
      )}
    </AnimatePresence>
  );

  // Kami menghapus screen loading terpisah agar transisi langsung dari screen hijau (splash) ke aplikasi utama tanpa ada jeda/screen kedua.

  if (!isLoggedIn) {
    return (
      <div className="h-[100dvh] bg-stone-50 flex flex-col relative overflow-y-auto font-sans pb-8">
        <GlobalConnectivityBanner />
        {splashNode}
        <div className="w-full mx-auto px-6 py-8 md:p-8 grid grid-cols-1 md:grid-cols-2 gap-12 md:gap-8 items-center relative z-10 my-auto max-w-5xl">
          
          {/* Left/Top Content */}
          <div className="flex flex-col text-center lg:text-left space-y-4 md:pr-8">
              <div className="flex flex-col lg:items-start items-center">
                <h1 className="text-4xl md:text-[2.75rem] font-black text-[#8dc63f] tracking-tighter drop-shadow-sm scale-y-105 origin-bottom lg:origin-left">My Kaguci App</h1>
                <p className="text-[8px] sm:text-[10px] font-bold text-slate-600 tracking-[0.2em] mt-3 mb-2">CERDAS • INOVATIF • TERAMPIL • RESPONSIF • AGAMIS</p>
              </div>
            
            <h2 className="text-2xl md:text-3xl font-bold text-slate-800 leading-snug mt-6 lg:text-left text-center max-w-sm mx-auto lg:mx-0">
              Sistem Informasi Absensi dan Nilai Siswa Berbasis Digital
            </h2>
            
            <p className="text-sm text-slate-600 leading-relaxed max-w-md mx-auto lg:mx-0 mt-4 lg:text-left text-center">
              Solusi modern untuk manajemen kehadiran dan nilai siswa yang akurat, real-time, dan terintegrasi. Membangun budaya disiplin melalui teknologi informasi yang cerdas.
            </p>

            {/* PWA Install Info Box on Login screen */}
            <div className="mt-6 p-5 bg-white/75 backdrop-blur-md rounded-2xl border border-slate-100 shadow-[0_12px_30px_-15px_rgba(141,198,63,0.18)] text-left max-w-md mx-auto lg:mx-0">
              <div className="flex items-center gap-2 mb-2 text-slate-800">
                <div className="p-1.5 bg-[#8dc63f]/10 text-[#8dc63f] rounded-lg">
                  <Download className="w-4 h-4" />
                </div>
                <span className="font-extrabold text-[10px] tracking-widest text-[#8dc63f] uppercase">Aplikasi PWA (Absensi Seluler)</span>
              </div>
              {isInsideIframe ? (
                <div>
                  <p className="text-[11px] text-slate-600 leading-relaxed mb-3">
                    Agar aplikasi absensi digital ini bisa diinstal langsung ke layar utama HP / Laptop Anda (tanpa Play Store), Anda perlu membukanya di tab browser mandiri.
                  </p>
                  <a
                    href="https://ais-pre-56w2g4yxpoxk4k23siyzdq-901834158843.asia-southeast1.run.app"
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-[#8dc63f] hover:bg-[#7bc025] text-white text-[10px] font-black uppercase tracking-wider rounded-lg transition-all shadow-[0_4px_10px_rgba(141,198,63,0.25)] hover:scale-102"
                  >
                    Buka Tab Mandiri & Instal <ArrowRight className="w-3.5 h-3.5" />
                  </a>
                </div>
              ) : isAppInstalled ? (
                <p className="text-[11px] text-emerald-600 font-bold leading-relaxed">
                  🎉 Aplikasi Absensi PWA ini telah terpasang dengan sukses di perangkat Anda!
                </p>
              ) : isInstallable ? (
                <div>
                  <p className="text-[11px] text-slate-600 leading-relaxed mb-3">
                    Aplikasi ini siap dipasang langsung tanpa App Store / Play Store. Cepat, ringan, dan hemat kuota internet.
                  </p>
                  <button
                    onClick={handleInstallClick}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-[#8dc63f] text-white hover:bg-[#7bc025] text-[10px] font-black uppercase tracking-wider rounded-lg transition-all shadow-md shadow-emerald-200 hover:scale-102 active:scale-98"
                  >
                    Instal Sekarang <ArrowRight className="w-3.5 h-3.5" />
                  </button>
                </div>
              ) : (
                <p className="text-[11px] text-slate-500 leading-relaxed">
                  Buka menu browser Anda (klik tombol titik tiga di Chrome, atau tombol Share/Bagikan di Safari iOS), lalu pilih <b>"Instal Aplikasi"</b> atau <b>"Tambahkan ke Layar Utama"</b>.
                </p>
              )}
            </div>
          </div>

          {/* Right/Bottom Content - Login & Recovery Card */}
          <div className="flex justify-center w-full">
            <div className="bg-white p-7 md:p-10 rounded-2xl shadow-[0_20px_50px_-12px_rgba(0,0,0,0.1)] border border-slate-100/80 w-full max-w-md">
              {isForgotPassword ? (
                <div>
                  <div className="text-center mb-6">
                    <h3 className="text-xl font-bold text-slate-800">Lupa Password / Username</h3>
                    <p className="text-xs text-slate-600 mt-1 lines-relaxed">
                      Sistem Pencarian Akun & Kredensial Sekolah Mandiri
                    </p>
                  </div>

                  <form onSubmit={async (e) => {
                    e.preventDefault();
                    if (!recoverySearchVal.trim()) {
                      showToast('Masukkan nilai pencarian.', 'error');
                      return;
                    }

                    setIsRecoveryLoading(true);
                    setRecoveryResult(null);

                    try {
                      if (recoverySearchType === 'username') {
                        // Look up password from Username/Email
                        const lookupDoc = await getDoc(doc(dbDefault, 'custom_accounts', recoverySearchVal.toLowerCase().trim()));
                        if (lookupDoc.exists()) {
                          setRecoveryResult([lookupDoc.data()]);
                          showToast('Siswa/Akun Ditemukan', 'success');
                        } else {
                          setRecoveryResult([]);
                          showToast('Akun tidak terregistrasi di sistem pusat.', 'error');
                        }
                      } else {
                        // Look up Username/Email from Password!
                        const q = query(collection(dbDefault, 'custom_accounts'), where('password', '==', recoverySearchVal.trim()));
                        const qSnap = await getDocs(q);
                        if (!qSnap.empty) {
                          const results = qSnap.docs.map(d => d.data());
                          setRecoveryResult(results);
                          showToast('Siswa/Akun Ditemukan', 'success');
                        } else {
                          setRecoveryResult([]);
                          showToast('Gagal menemukan data dengan kata sandi tersebut.', 'error');
                        }
                      }
                    } catch (err) {
                      const error = err as Error;
                      showToast('Pencarian Gagal: ' + error.message, 'error');
                    } finally {
                      setIsRecoveryLoading(false);
                    }
                  }} className="space-y-4">
                    
                    {/* Toggle Search Mode */}
                    <div className="bg-slate-100 p-1.5 rounded-xl flex gap-1 mb-2">
                      <button
                        type="button"
                        onClick={() => {
                          setRecoverySearchType('username');
                          setRecoveryResult(null);
                        }}
                        className={`flex-1 text-center py-2 text-xs font-bold rounded-lg transition-all ${
                          recoverySearchType === 'username' ? 'bg-white text-slate-800 shadow-sm' : 'text-slate-600 hover:text-slate-700'
                        }`}
                      >
                        Cari Dengan Username
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          setRecoverySearchType('password');
                          setRecoveryResult(null);
                        }}
                        className={`flex-1 text-center py-2 text-xs font-bold rounded-lg transition-all ${
                          recoverySearchType === 'password' ? 'bg-white text-slate-800 shadow-sm' : 'text-slate-600 hover:text-slate-700'
                        }`}
                      >
                        Cari Dengan Password
                      </button>
                    </div>

                    <div>
                      <label className="block text-[10px] font-bold text-slate-600 mb-2 tracking-wider">
                        {recoverySearchType === 'username' ? 'MASUKKAN USERNAME / EMAIL' : 'MASUKKAN KATA SANDI'}
                      </label>
                      <div className="relative">
                        {recoverySearchType === 'username' ? (
                          <UserIcon className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500" />
                        ) : (
                          <Key className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500" />
                        )}
                        <input 
                          type="text" 
                          className="w-full pl-11 pr-4 py-3 bg-[#fefce8] border-2 border-[#8dc63f] rounded-lg focus:ring-2 focus:ring-[#8dc63f] focus:border-[#8dc63f] transition-all outline-none text-slate-700 font-medium placeholder:text-slate-500 text-sm" 
                          placeholder={recoverySearchType === 'username' ? 'Masukkan username / email' : 'Ketik kata sandi'} 
                          value={recoverySearchVal}
                          onChange={e => setRecoverySearchVal(e.target.value)}
                          required 
                        />
                      </div>
                    </div>

                    <button 
                      type="submit" 
                      disabled={isRecoveryLoading}
                      className="w-full bg-[#8dc63f] text-white font-bold py-3 rounded-lg hover:bg-[#7bc025] transition-colors text-sm shadow-sm flex items-center justify-center gap-2"
                    >
                      {isRecoveryLoading ? (
                        <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></div>
                      ) : (
                        'Temukan Kredensial Saya'
                      )}
                    </button>

                    {/* RECOVERY RESULT DISPLAY DISPLAY */}
                    {recoveryResult !== null && (
                      <div className="mt-4 p-4 rounded-xl border border-dashed text-left space-y-3 bg-stone-50">
                        {recoveryResult.length > 0 ? (
                          <div>
                            <span className="block text-[10px] font-bold text-[#8dc63f] mb-2 uppercase">✓ Kredensial Ditemukan</span>
                            {recoveryResult.map((acc, index) => (
                              <div key={index} className="space-y-2 border-t border-slate-200 pt-2 first:border-0 first:pt-0">
                                <div>
                                  <span className="text-[9px] font-bold text-slate-500 block">NAMA LENGKAP</span>
                                  <span className="text-xs font-bold text-slate-800">{acc.fullname}</span>
                                </div>
                                <div className="grid grid-cols-2 gap-2">
                                  <div>
                                    <span className="text-[9px] font-bold text-slate-500 block">USERNAME</span>
                                    <span className="text-xs font-mono font-bold text-slate-800 break-all bg-white px-1.5 py-0.5 rounded border border-slate-100 block">{acc.username}</span>
                                  </div>
                                  <div>
                                    <span className="text-[9px] font-bold text-slate-500 block">PASSWORD</span>
                                    <span className="text-xs font-mono font-bold text-lime-700 bg-white px-1.5 py-0.5 rounded border border-slate-100 block">{acc.password}</span>
                                  </div>
                                </div>
                              </div>
                            ))}
                          </div>
                        ) : (
                          <div className="text-center py-2">
                            <span className="text-xs font-bold text-rose-500">❌ Tidak ada hasil yang cocok!</span>
                            <p className="text-[10px] text-slate-500 mt-1">Pastikan email / password yang dimasukkan tepat.</p>
                          </div>
                        )}
                      </div>
                    )}

                    <div className="relative flex items-center py-2">
                      <div className="flex-grow border-t border-slate-100"></div>
                      <span className="flex-shrink-0 mx-4 text-[10px] font-bold text-slate-500 uppercase tracking-widest">Atau</span>
                      <div className="flex-grow border-t border-slate-100"></div>
                    </div>

                    <button 
                      type="button" 
                      onClick={() => {
                        setIsForgotPassword(false);
                        setRecoveryResult(null);
                        setRecoverySearchVal('');
                      }}
                      className="w-full bg-transparent border-2 border-slate-400 text-slate-600 font-bold py-3.5 rounded-lg hover:bg-slate-50 transition-colors text-sm"
                    >
                      Kembali ke Halaman Masuk
                    </button>
                  </form>
                </div>
              ) : (
                <div>
                  <div className="text-center mb-6">
                    <h3 className="text-xl font-bold text-slate-800">Selamat Datang</h3>
                    <p className="text-xs text-slate-600 mt-1">Silakan masuk ke akun atau sekolah Anda</p>
                  </div>

                  <form onSubmit={async (e) => { 
                    e.preventDefault(); 
                    setIsAuthLoading(true);


                    try {
                      const lookupDoc = await getDoc(doc(dbDefault, 'custom_accounts', authEmail.toLowerCase().trim()));
                      
                      if (!lookupDoc.exists()) {
                        // Check if they are in the auth database
                        try {
                          await signInWithEmailAndPassword(auth, toAuthEmail(authEmail), authPassword);
                          // If successful but there is no custom_accounts doc, it means they were DELETED
                          // Do NOT let them log in
                          await signOut(auth);
                          
                          setLoginError({
                            title: 'Akun Telah Dihapus!',
                            message: `Akun "${authEmail}" telah dihapus oleh Administrator. Anda tidak dapat masuk menggunakan akun ini lagi.`,
                            recommendations: [
                              'Hubungi Administrator jika Anda merasa ini adalah kesalahan.',
                              'Data absensi/sekolah Anda sudah tidak dapat diakses dan dihentikan.',
                              'Silakan buat akun baru jika ingin mendaftar ulang menggunakan token valid.'
                            ],
                            username: authEmail
                          });
                          setIsAuthLoading(false);
                          return;
                        } catch (errFallback) {
                          const fallbackErr = errFallback as Error;
                          const isWrongPass = fallbackErr.message.includes('wrong-password') || fallbackErr.message.includes('invalid-credential') || fallbackErr.message.includes('invalid-email');
                          
                          if (isWrongPass) {
                            // Either password is wrong, OR we just show a generic 'not registered' to prevent brute force 
                            setLoginError({
                              title: 'Akun Tidak Ditemukan!',
                              message: `Username atau Email "${authEmail}" tidak ditemukan atau kata sandi tidak cocok.`,
                              recommendations: [
                                'Periksa kembali penulisan nama pengguna / email Anda.',
                                'Pastikan huruf besar/kecil dan ejaan sudah benar.',
                                'Jika Anda belum memiliki akun, klik "Tambah Akun Baru" di pojok kanan bawah.'
                              ],
                              username: authEmail
                            });
                          } else {
                            // Any other error
                            setLoginError({
                              title: 'Akun Belum Terdaftar!',
                              message: `Username atau Email "${authEmail}" tidak terdaftar di sistem.`,
                              recommendations: [
                                'Periksa kembali penulisan nama pengguna / email Anda (pastikan tidak ada salah ketik).',
                                'Jika Anda admin atau perwakilan sekolah, silakan daftar database sekolah Anda terlebih dahulu.',
                                'Klik tombol "Tambah Akun Baru" di bagian pojok kanan bawah halaman login ini untuk memulai registrasi baru.',
                                'Kombinasi akun ini belum pernah terdaftar di server kaguci.'
                              ],
                              username: authEmail
                            });
                          }
                          setIsAuthLoading(false);
                          return;
                        }
                      }

                      const accData = lookupDoc.data();

                      setActiveUserCustomData({
                        fullname: accData.fullname,
                        username: accData.username,
                        configText: accData.configText || ''
                      });
                      localStorage.setItem('kaguci_active_custom_user', JSON.stringify({
                        fullname: accData.fullname,
                        username: accData.username,
                        configText: accData.configText || ''
                      }));

                      await signInWithEmailAndPassword(auth, toAuthEmail(authEmail), authPassword);
                      showToast('Selamat Datang Kembali! Login berhasil.', 'success');
                      localStorage.setItem('kaguci_has_logged_in', 'true');
                      localStorage.setItem('kaguci_saved_credentials', JSON.stringify({
                        email: authEmail,
                        password: authPassword
                      }));
                      const hashVal = encodeSession(authEmail, authPassword);
                      if (hashVal) {
                        window.location.hash = `session=${hashVal}`;
                      }

                      // Indonesian welcoming voice greeting
                      const greetingName = accData.fullname || accData.username || authEmail || "User";
                      const phrase = `Selamat datang ${greetingName} di aplikasi My Kaguci App.`;
                      sessionStorage.setItem(`kaguci_welcomed_${accData.username || authEmail}`, 'true');
                      speakText(phrase);

                      setIsLoggedIn(true);
                      setIsAuthLoading(false);
                    } catch (error) {
                      const err = error as Error;
                      const errMsg = err.message;
                      const isWrongPass = errMsg.includes('wrong-password') || errMsg.includes('invalid-credential') || errMsg.includes('invalid-email');
                      
                      if (isWrongPass) {
                        setLoginError({
                          title: 'Kata Sandi Salah!',
                          message: `Username "${authEmail}" terdaftar di sistem pusat, namun kata sandi yang Anda masukkan salah atau tidak cocok untuk basis data sekolah Anda.`,
                          recommendations: [
                            'Periksa kembali penulisan kata sandi Anda (pastikan huruf besar/kecil prasyarat Caps Lock sudah tepat).',
                            'Pastikan Anda masuk ke sekolah yang tepat jika memiliki beberapa akun terdaftar.',
                            'Hubungi administrator IT sekolah Anda untuk melakukan reset kata sandi lewat Firebase Console jika lupa.'
                          ],
                          username: authEmail
                        });
                      } else if (errMsg.includes('user-not-found')) {
                        setLoginError({
                          title: 'Akun Tidak Ditemukan!',
                          message: `Akun "${authEmail}" terdaftar di sistem pusat, namun tidak ditemukan di tabel autentikasi database sekolah privat Anda.`,
                          recommendations: [
                            'Layanan database privat sekolah Anda mungkin telah mengalami reset atau dibersihkan.',
                            'Silakan buat kembali akun dengan mengklik tombol "Tambah Akun Baru" untuk mendaftarkannya kembali secara utuh.'
                          ],
                          username: authEmail
                        });
                      } else {
                        setLoginError({
                          title: 'Gagal Masuk!',
                          message: `Gagal menghubungkan atau mengautentikasi akun Anda: ${err.message}`,
                          recommendations: [
                            'Periksa koneksi internet perangkat Anda.',
                            'Ada kemungkinan server Firebase sekolah Anda sedang offline atau membatasi akses aturan rules.',
                            'Silakan muat ulang halaman ini dan ulangi login beberapa saat lagi.'
                          ],
                          username: authEmail
                        });
                      }
                      
                      // reset credentials on error
                      localStorage.removeItem('kaguci_active_custom_user');
                      localStorage.removeItem('kaguci_saved_credentials');
                      setActiveUserCustomData(null);
                    } finally {
                      setIsAuthLoading(false);
                    }
                  }} className="space-y-5">
                    <div>
                      <label className="block text-[10px] font-bold text-slate-600 mb-2 tracking-wider">USERNAME / EMAIL</label>
                      <div className="relative">
                        <UserIcon className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500" />
                        <input 
                          type="text" 
                          className="w-full pl-11 pr-4 py-3 bg-[#fefce8] border-2 border-[#8dc63f] rounded-lg focus:ring-2 focus:ring-[#8dc63f] focus:border-[#8dc63f] transition-all outline-none text-slate-700 font-medium placeholder:text-slate-500 text-sm" 
                          placeholder="Contoh: admin atau budi" 
                          value={authEmail}
                          onChange={e => setAuthEmail(e.target.value)}
                          required 
                        />
                      </div>
                    </div>

                    <div>
                      <label className="block text-[10px] font-bold text-slate-600 mb-2 tracking-wider">KATA SANDI</label>
                      <div className="relative flex items-center">
                        <Key className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500" />
                        <input 
                          type={showPassword ? 'text' : 'password'} 
                          className="w-full pl-11 pr-12 py-3 bg-[#fefce8] border-2 border-[#8dc63f] rounded-lg focus:ring-2 focus:ring-[#8dc63f] focus:border-[#8dc63f] transition-all outline-none text-slate-700 font-medium placeholder:text-slate-500 text-lg tracking-[0.2em]" 
                          placeholder="•••••" 
                          value={authPassword}
                          onChange={e => setAuthPassword(e.target.value)}
                          required 
                        />
                        {showPassword ? (
                          <EyeOff 
                            onClick={() => setShowPassword(false)} 
                            className="absolute right-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500 cursor-pointer hover:text-slate-600 transition-colors" 
                          />
                        ) : (
                          <Eye 
                            onClick={() => setShowPassword(true)} 
                            className="absolute right-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500 cursor-pointer hover:text-slate-600 transition-colors" 
                          />
                        )}
                      </div>
                    </div>

                    <div className="pt-2">
                      <button 
                        type="submit" 
                        className="w-full bg-[#8dc63f] text-white font-bold py-3.5 rounded-lg hover:bg-[#7bc025] transition-colors text-sm shadow-sm active:scale-[0.98]"
                      >
                        Masuk Ke Sistem
                      </button>
                      
                      <div className="text-center mt-4">
                        <button 
                          type="button"
                          onClick={() => {
                            setIsForgotPassword(true);
                            setRecoveryResult(null);
                            setRecoverySearchVal('');
                          }}
                          className="inline-block text-xs font-bold text-[#8dc63f] hover:underline cursor-pointer focus:outline-none"
                        >
                          Lupa Password / Username?
                        </button>
                      </div>
                    </div>

                    <button 
                      type="button" 
                      onClick={() => setIsRegisterModalOpen(true)}
                      className="w-full bg-transparent border border-[#8dc63f] text-[#8dc63f] font-bold py-3.5 rounded-lg hover:bg-emerald-50 transition-colors text-sm mt-3 flex items-center justify-center gap-1"
                    >
                      Tambah Akun Baru
                    </button>
                  </form>
                </div>
              )}
            </div>
          </div>
        </div>

          {/* Kotak Dialog Tambah Akun Baru (Modal Dialog Box) */}
          <AnimatePresence>
            {isRegisterModalOpen && (
              <motion.div 
                 initial={{ opacity: 0 }}
                 animate={{ opacity: 1 }}
                 exit={{ opacity: 0 }}
                 className="fixed inset-0 bg-[#020617]/70 backdrop-blur-sm z-[100] flex items-center justify-center overflow-y-auto p-4"
              >
                <motion.div 
                   initial={{ opacity: 0, scale: 0.95, y: 15 }}
                   animate={{ opacity: 1, scale: 1, y: 0 }}
                   exit={{ opacity: 0, scale: 0.95, y: 15 }}
                   className="bg-white rounded-3xl p-6 sm:p-8 pb-20 sm:pb-24 shadow-2xl max-w-lg w-[95%] my-8 space-y-6 max-h-[90vh] overflow-y-auto text-left m-4"
                >
                  <div>
                     <h3 className="text-xl font-bold text-slate-800">Tambah Akun Baru</h3>
                     <p className="text-xs text-slate-600 mt-1">Isi formulir pendaftaran akun menggunakan token valid dari Admin.</p>
                  </div>

                  <div className="space-y-4 text-left">
                    <div>
                      <label className="block text-[10px] font-bold text-slate-600 mb-1.5 uppercase tracking-wider text-left">Nama Lengkap Pemilik</label>
                      <input 
                        type="text"
                        className="w-full px-4 py-3 bg-[#fefce8] border-2 border-[#8dc63f] rounded-xl focus:ring-2 focus:ring-[#8dc63f] focus:border-[#8dc63f] transition-all outline-none text-slate-700 font-medium placeholder:text-slate-500 text-sm"
                        placeholder="Contoh: Budi Santoso, S.Pd."
                        value={regFullName}
                        onChange={e => setRegFullName(e.target.value)}
                      />
                    </div>

                    <div>
                       <label className="block text-[10px] font-bold text-slate-600 mb-1.5 uppercase tracking-wider text-left">Username / Email Akun</label>
                       <input 
                         type="text"
                         className="w-full px-4 py-3 bg-[#fefce8] border-2 border-[#8dc63f] rounded-xl focus:ring-2 focus:ring-[#8dc63f] focus:border-[#7bc025] transition-all outline-none text-slate-700 font-medium placeholder:text-slate-500 text-sm"
                         placeholder="Contoh: admin atau budi"
                         value={regUsername}
                         onChange={e => setRegUsername(e.target.value)}
                       />
                    </div>

                    <div>
                       <label className="block text-[10px] font-bold text-slate-600 mb-1.5 uppercase tracking-wider text-left">Kata Sandi (Min. 6 digit)</label>
                       <input 
                         type="text"
                         className="w-full px-4 py-3 bg-[#fefce8] border-2 border-[#8dc63f] rounded-xl focus:ring-2 focus:ring-[#8dc63f] focus:border-[#7bc025] transition-all outline-none text-slate-700 font-medium placeholder:text-slate-500 text-sm"
                         placeholder="Ketik password untuk pendaftaran"
                         value={regPassword}
                         onChange={e => setRegPassword(e.target.value)}
                       />
                    </div>

                    <div>
                       <label className="block text-[10px] font-bold text-slate-600 mb-1.5 uppercase tracking-wider text-left">Token Pendaftaran (Dari Admin - Opsional)</label>
                       <input 
                         type="text"
                         className="w-full px-4 py-3 bg-[#fefce8] border-2 border-[#8dc63f] rounded-xl focus:ring-2 focus:ring-[#8dc63f] focus:border-[#7bc025] transition-all outline-none text-slate-700 font-medium placeholder:text-slate-500 text-sm"
                         placeholder="Masukkan token 6-8 digit (Opsional)"
                         value={regToken}
                         onChange={e => setRegToken(e.target.value)}
                       />
                    </div>

                    <div>
                      <div className="flex justify-between items-center mb-1.5">
                        <label className="block text-[10px] font-bold text-slate-600 uppercase tracking-wider text-left">Status Koneksi Database</label>
                        <span className="px-2 py-0.5 bg-emerald-100 text-[#7bc025] rounded-full font-bold text-[9px] uppercase tracking-wide">Otomatis & Terpusat</span>
                      </div>
                      <div className="p-4 bg-emerald-50/50 border border-lime-100/50 rounded-2xl space-y-2 text-left">
                        <p className="text-[11px] text-slate-600 leading-relaxed font-semibold">
                          Sistem akan mendaftarkan akun Anda secara langsung ke Database Pusat kaguci. Semua data presensi, absensi suara, dan data siswa akan aman terenkripsi dan disinkronkan secara otomatis.
                        </p>
                      </div>
                    </div>
                  </div>

                  <div className="flex flex-col sm:flex-row gap-3 pt-2">
                    <button 
                      type="button" 
                      onClick={() => {
                        setIsRegisterModalOpen(false);
                        setRegFullName('');
                        setRegUsername('');
                        setRegPassword('');
                        setRegToken('');
                      }}
                      className="flex-1 px-5 py-3 rounded-xl font-bold bg-slate-100 hover:bg-slate-200 text-slate-600 text-sm transition-colors"
                    >
                      Batal
                    </button>
                    <button 
                      type="button" 
                      disabled={!regFullName || !regUsername || !regPassword}
                      onClick={async () => {
                        setIsAuthLoading(true);
                        try {
                          const trimmedFullName = regFullName.trim();
                          const trimmedUsername = regUsername.toLowerCase().trim();
                          const trimmedPassword = regPassword;

                          if (trimmedFullName.length < 3) {
                            showToast('Nama Lengkap minimal berisi 3 karakter.', 'error');
                            setIsAuthLoading(false);
                            return;
                          }
                          if (trimmedUsername.length < 3) {
                            showToast('Username minimal berisi 3 karakter.', 'error');
                            setIsAuthLoading(false);
                            return;
                          }
                          if (trimmedPassword.length < 6) {
                            showToast('Kata Sandi minimal berisi 6 karakter.', 'error');
                            setIsAuthLoading(false);
                            return;
                          }

                          const resolvedToken = regToken.trim();

                          
                          if (resolvedToken) {
                            try {
                              const tokenDoc = await getDoc(doc(dbDefault, 'register_tokens', resolvedToken));
                              if (tokenDoc.exists()) {
                                const tokenData = tokenDoc.data();
                                if (tokenData?.used) {
                                  console.log('Token sudah terpakai, melanjutkan tanpa kendala.');
                                }
                              }
                            } catch (err) {
                              console.error('Gagal mengecek token:', err);
                            }
                          }

                          const resolvedUsername = trimmedUsername;

                          let usernameExists = false;
                          let usernameOwnerName = '';

                          // 1. Periksa apakah Username sudah terdaftar di database sistem pusat (STRICT BLOCK)
                          try {
                            const lookupUsernameDoc = await getDoc(doc(dbDefault, 'custom_accounts', resolvedUsername));
                            if (lookupUsernameDoc.exists()) {
                              const data = lookupUsernameDoc.data();
                              usernameExists = true;
                              usernameOwnerName = data.fullname || 'N/A';
                            }
                          } catch (err) {
                            console.error('Kendala saat melacak username:', err);
                          }

                          if (usernameExists) {
                            setIsAuthLoading(false);
                            setRegistrationResult({
                              success: false,
                              title: 'Username Sudah Terdaftar!',
                              message: `Maaf, Nama Akun (Username) "${resolvedUsername}" sudah terdaftar di sistem pusat atas nama "${usernameOwnerName}".\n\nSilakan pilih Username lain yang berbeda.`
                            });
                            return;
                          }

                          // 2. Buat akun native di database pusat
                          try {
                            await createUserWithEmailAndPassword(auth, toAuthEmail(regUsername), regPassword);
                          } catch (authError) {
                            const errObj = authError as { message?: string; code?: string };
                            if (errObj.message?.includes('email-already-in-use') || errObj.code === 'auth/email-already-in-use') {
                              // Mencoba login untuk pembuktian kepemilikan
                              try {
                                await signInWithEmailAndPassword(auth, toAuthEmail(regUsername), regPassword);
                              } catch {
                                throw authError;
                              }
                            } else {
                              throw authError;
                            }
                          }

                          // 3. Simpan data registrasi ke Firestore terpusat
                          await setDoc(doc(dbDefault, 'custom_accounts', resolvedUsername), {
                            fullname: regFullName,
                            username: resolvedUsername,
                            password: regPassword,
                            configText: '',
                            createdAt: new Date().toISOString()
                          });

                          // 4. Mark token as used
                          if (resolvedToken !== 'SUPERADMINTOKEN') {
                            try {
                              await setDoc(doc(dbDefault, 'register_tokens', resolvedToken), {
                                used: true,
                                usedAt: new Date().toISOString(),
                                usedBy: resolvedUsername
                              }, { merge: true });
                            } catch (err) {
                              console.error('Gagal menandai token sebagai terpakai:', err);
                            }
                          }

                          // On success: trigger results popup
                          setRegistrationResult({
                            success: true,
                            title: 'Pendaftaran Akun Berhasil!',
                            message: 'Akun Administrator baru Anda siap digunakan dengan Cloud Database terpusat kaguci secara aman dan instan.',
                            fullname: regFullName,
                            username: resolvedUsername,
                            projectId: 'Database Terpusat kaguci SMAN 1 Cililin'
                          });

                          // Autofill login credentials for easy access
                          setAuthEmail(regUsername);
                          setAuthPassword(regPassword);
                          
                          const hashVal = encodeSession(regUsername, regPassword);
                          if (hashVal) {
                            window.location.hash = `session=${hashVal}`;
                          }
                          
                          // Clean fields and close form
                          setIsRegisterModalOpen(false);
                          setRegFullName('');
                          setRegUsername('');
                          setRegPassword('');
                          setRegToken('');
                        } catch (error) {
                          const err = error as Error;
                          let IndonesianError = err.message;
                          if (err.message.includes('email-already-in-use')) {
                            const lookupUsername = regUsername.toLowerCase().trim();
                            let centralUserFound: { fullname?: string; username?: string; configText?: string } | null = null;
                            try {
                              const lookupDoc = await getDoc(doc(dbDefault, 'custom_accounts', lookupUsername));
                              if (lookupDoc.exists()) {
                                const data = lookupDoc.data();
                                centralUserFound = {
                                  fullname: data.fullname,
                                  username: data.username,
                                  configText: data.configText
                                };
                              }
                            } catch (lookupErr) {
                              console.error('Error lookup di catch:', lookupErr);
                            }

                            if (centralUserFound) {
                              IndonesianError = `Maaf, Username "${lookupUsername}" sudah terdaftar di database sistem pusat.\n\n• Nama Pengguna: ${centralUserFound.fullname || 'N/A'}\n• Akun (Username): ${centralUserFound.username || 'N/A'}\n\nSilakan gunakan menu "Masuk" (Login) dan gunakan akun tersebut beserta kata sandinya untuk login.`;
                            } else {
                              IndonesianError = `Username "${lookupUsername}" sudah pernah didaftarkan pada database pusat, namun kata sandi yang Anda ketik salah.\n\nLangkah Solusi:\n1. Masukkan kata sandi yang tepat jika Anda adalah pemilik akun tersebut.\n2. ATAU, silakan daftar dengan memakai Username yang berbeda.`;
                            }
                          } else if (err.message.includes('weak-password')) {
                            IndonesianError = 'Kata sandi minimal berisi 6 karakter.';
                          } else if (err.message.includes('invalid-api-key') || err.message.includes('API key')) {
                            IndonesianError = 'API Key yang terdapat pada konfigurasi Web Firebase Anda salah atau tidak valid.';
                          } else if (err.message.includes('network-request-failed')) {
                            IndonesianError = 'Koneksi jaringan gagal. Periksa koneksi internet Anda atau rules Firebase.';
                          } else if (err.message.includes('operation-not-allowed')) {
                            IndonesianError = 'Provider "Email/Password" belum aktif di Firebase Console Anda!';
                          } else if (err.message.includes('configuration-not-found') || err.message.includes('auth/configuration-not-found')) {
                            IndonesianError = 'Layanan Autentikasi belum diinisialisasi di Proyek Firebase Anda!';
                          }

                          setRegistrationResult({
                            success: false,
                            title: 'Pendaftaran Gagal!',
                            message: `${IndonesianError}\n\n(Detail Teknis: ${err.message})`
                          });
                        } finally {
                          setIsAuthLoading(false);
                        }
                      }}
                      className="flex-1 px-5 py-3 rounded-xl font-bold text-white bg-[#8dc63f] hover:bg-[#7bc025] disabled:bg-slate-300 disabled:cursor-not-allowed text-sm transition-all text-center"
                    >
                      Konfigurasi Database & Daftar Akun
                    </button>
                  </div>
                  {/* Spacer bottom to handle mobile browsers and overlays */}
                  <div className="h-4 sm:h-6 shrink-0" />
                </motion.div>
              </motion.div>
            )}
          </AnimatePresence>

          {/* Kotak Dialog Hasil Pendaftaran Akun / Firebase */}
          <AnimatePresence>
            {registrationResult && (
              <motion.div 
                 initial={{ opacity: 0 }}
                 animate={{ opacity: 1 }}
                 exit={{ opacity: 0 }}
                 className="fixed inset-0 bg-[#020617]/80 backdrop-blur-md z-[200] flex items-center justify-center p-4 overflow-y-auto"
              >
                <motion.div 
                   initial={{ opacity: 0, scale: 0.95, y: 15 }}
                   animate={{ opacity: 1, scale: 1, y: 0 }}
                   exit={{ opacity: 0, scale: 0.95, y: 15 }}
                   className="bg-white rounded-3xl p-6 sm:p-8 shadow-2xl max-w-md w-full space-y-5 text-center relative max-h-[85vh] sm:max-h-[90vh] overflow-y-auto"
                >
                  {/* Decorative status header background with matching rounded top corners */}
                  <div className={`absolute top-0 left-0 right-0 h-2 rounded-t-3xl ${
                    registrationResult.success ? 'bg-[#8dc63f]' : 'bg-rose-500'
                  }`} />

                  {/* Icon indicator */}
                  <div className="flex justify-center pt-3">
                    {registrationResult.success ? (
                      <div className="w-16 h-16 bg-[#f7fee7] border-2 border-[#bbf7d0] rounded-full flex items-center justify-center text-[#8dc63f]">
                        <svg className="w-8 h-8 font-black" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="3" d="M5 13l4 4L19 7" />
                        </svg>
                      </div>
                    ) : (
                      <div className="w-16 h-16 bg-rose-50 border-2 border-rose-200 rounded-full flex items-center justify-center text-rose-500">
                        <svg className="w-8 h-8 font-black" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="3" d="M6 18L18 6M6 6l12 12" />
                        </svg>
                      </div>
                    )}
                  </div>

                  <div className="space-y-2">
                     <h3 className={`text-lg font-black tracking-tight ${
                       registrationResult.success ? 'text-slate-800' : 'text-rose-600'
                     }`}>
                       {registrationResult.title}
                     </h3>
                     <p className="text-xs text-slate-600 leading-relaxed px-2">
                       {registrationResult.message}
                     </p>
                  </div>

                  {registrationResult.success && (
                    <div className="bg-slate-50 border border-slate-100 rounded-2xl p-4 text-left space-y-2 text-xs">
                      <div>
                        <span className="text-[9px] font-bold text-slate-500 block uppercase">NAMA LENGKAP UTAMA</span>
                        <span className="font-bold text-slate-700">{registrationResult.fullname}</span>
                      </div>
                      <div className="grid grid-cols-2 gap-2 pt-2 border-t border-slate-200/60">
                        <div>
                          <span className="text-[9px] font-bold text-slate-500 block uppercase">USERNAME LOGIN</span>
                          <span className="font-mono font-bold text-[#8dc63f] break-all">{registrationResult.username}</span>
                        </div>
                        <div>
                          <span className="text-[9px] font-bold text-slate-500 block uppercase">PROJECT ID</span>
                          <span className="font-mono font-bold text-slate-700 break-all">{registrationResult.projectId}</span>
                        </div>
                      </div>
                    </div>
                  )}

                  {!registrationResult.success && (
                    <div className="bg-rose-50/50 border border-rose-100 rounded-2xl p-3 text-left text-xs text-rose-700 leading-relaxed">
                      <strong className="block mb-1">Rekomendasi Solusi:</strong>
                      <ul className="list-disc pl-4 mt-1 space-y-1 text-[11px]">
                        <li>Verifikasi kebenaran dan format Config Firebase yang dipaste.</li>
                        <li>Pastikan database auth di Firebase Console sudah diaktifkan (Email/Password).</li>
                        <li>Pastikan aturan Firestore rules Anda tidak memblokir operasi pembuatan user.</li>
                      </ul>
                    </div>
                  )}

                  {!registrationResult.success && registrationResult.showBypassButton && registrationResult.configToSave && (
                    <button 
                       type="button" 
                       onClick={() => handleForceRegister(registrationResult.configToSave!)}
                       className="w-full py-3 rounded-xl font-bold text-white bg-[#8dc63f] hover:bg-[#7bc025] transition-all text-sm shadow-md active:scale-[0.98] mt-2 block"
                    >
                      Hubungkan & Tetap Daftar Akun Baru
                    </button>
                  )}

                  <button 
                     type="button" 
                     onClick={() => setRegistrationResult(null)}
                     className={`w-full py-3 rounded-xl font-bold text-white transition-all text-sm shadow-md active:scale-[0.98] ${
                       registrationResult.success 
                         ? 'bg-[#8dc63f] hover:bg-[#7bc025]' 
                         : 'bg-slate-800 hover:bg-slate-900 border border-slate-700'
                     }`}
                  >
                    {registrationResult.success ? 'Selesai & Masuk Sekarang' : registrationResult.showBypassButton ? 'Batal' : 'Tutup Dialog'}
                  </button>
                </motion.div>
              </motion.div>
            )}

            {/* Kotak Dialog Alert Modern - Login Gagal / Salah Password / Belum Terdaftar */}
            {loginError && (
              <motion.div 
                 initial={{ opacity: 0 }}
                 animate={{ opacity: 1 }}
                 exit={{ opacity: 0 }}
                 className="fixed inset-0 bg-[#020617]/80 backdrop-blur-md z-[210] flex items-center justify-center p-4 overflow-y-auto"
              >
                <motion.div 
                   initial={{ opacity: 0, scale: 0.95, y: 15 }}
                   animate={{ opacity: 1, scale: 1, y: 0 }}
                   exit={{ opacity: 0, scale: 0.95, y: 15 }}
                   className="bg-white rounded-3xl p-6 sm:p-8 shadow-2xl max-w-md w-full space-y-5 text-center relative max-h-[85vh] sm:max-h-[90vh] overflow-y-auto"
                >
                  {/* Decorative error status line at top */}
                  <div className="absolute top-0 left-0 right-0 h-2 bg-rose-500 rounded-t-3xl" />

                  {/* Icon indicator with warnings */}
                  <div className="flex justify-center pt-3">
                    <div className="w-16 h-16 bg-rose-50 border-2 border-rose-200 rounded-full flex items-center justify-center text-rose-500 shadow-sm animate-pulse">
                      <svg className="w-8 h-8 font-black" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
                      </svg>
                    </div>
                  </div>

                  <div className="space-y-2">
                     <h3 className="text-lg font-black tracking-tight text-rose-600">
                       {loginError.title}
                     </h3>
                     <p className="text-xs text-slate-600 leading-relaxed px-2">
                       {loginError.message}
                     </p>
                  </div>

                  {/* List of custom helpful step-by-step procedures */}
                  <div className="bg-rose-50/50 border border-rose-100/60 rounded-2xl p-4 text-left text-xs text-rose-700 leading-relaxed">
                    <strong className="block mb-2 font-black text-rose-800 text-[11px] uppercase tracking-wider">Rekomendasi Solusi & Navigasi:</strong>
                    <ul className="list-disc pl-4 space-y-2 text-[11px] text-slate-600">
                      {loginError.recommendations.map((rec, idx) => (
                        <li key={idx} className="leading-snug">
                          {rec}
                        </li>
                      ))}
                    </ul>
                  </div>

                  <button 
                     type="button" 
                     onClick={() => setLoginError(null)}
                     className="w-full py-3 rounded-xl font-bold text-white bg-slate-800 hover:bg-slate-900 border border-slate-700 transition-all text-sm shadow-md active:scale-[0.98]"
                  >
                    Tutup Pesan
                  </button>
                </motion.div>
              </motion.div>
            )}


          </AnimatePresence>

        {/* Global Footer */}
        <div className="mt-auto text-center pb-8 pt-8 z-10 w-full px-4 flex flex-col items-center">
           <div className="flex items-center justify-center gap-2 mb-2">
              <span className="w-2 h-2 rounded-full bg-[#8dc63f]"></span>
              <p className="text-sm font-black text-[#8dc63f]">App Development by Tim IT SMAN 1 Cililin</p>
           </div>
           <p className="text-[8px] md:text-[9px] font-bold text-slate-500 tracking-[0.15em] sm:tracking-[0.25em] uppercase mb-2">
             KREATIVITAS TANPA BATAS • INOVASI TIADA HENTI
           </p>
           <p className="text-[9px] md:text-[10px] italic text-slate-500 text-center max-w-xs md:max-w-none">
             "Transformasi Digital Pendidikan Untuk Generasi Emas yang Cerdas dan Berakhlak"
           </p>
           
           <div className="flex justify-center items-center gap-6 md:gap-8 mt-10 opacity-70">
             <span className="text-[9px] md:text-[10px] font-bold text-slate-500">V2.1.0</span>
             <span className="text-[9px] md:text-[10px] font-bold text-slate-200">|</span>
             <span className="text-[9px] md:text-[10px] font-bold text-slate-500 uppercase">Enterprise</span>
             <span className="text-[9px] md:text-[10px] font-bold text-slate-200">|</span>
             <span className="text-[9px] md:text-[10px] font-bold text-slate-500 uppercase">Stable</span>
           </div>
        </div>
      </div>
    );
  }

  return (
    <div className="h-[100dvh] bg-slate-50 flex flex-col overflow-hidden print:h-auto print:overflow-visible selection:bg-[#8dc63f]/20 selection:text-[#7bc025]">
      <GlobalConnectivityBanner />
      {splashNode}

      {/* Persistence / Connectivity Banners - OLD REMOVED */}

      {/* Header */}
      <header className="bg-white px-6 py-3.5 flex justify-between items-center border-b border-slate-100 z-40 shadow-[0_2px_15px_rgba(148,163,184,0.03)] shrink-0 print:hidden">
        <div className="flex items-center gap-3">
          <img src="https://drive.google.com/thumbnail?id=1P395tuZymxs3qero4XduMpHy7g2GJrdR&sz=w1000" alt="Logo" className="w-10 h-10 object-contain drop-shadow-sm" referrerPolicy="no-referrer" />
          <div>
            <h1 className="text-lg font-black text-[#8dc63f] tracking-tight leading-none scale-y-105 origin-left">My Kaguci App</h1>
            <p className="text-[10px] text-slate-500 font-bold tracking-wider uppercase mt-1">SMA Negeri 1 Cililin</p>
          </div>
        </div>
        <div className="flex items-center gap-3">
          {/* Status Sinkronisasi Real-time Database */}
          <div className={`flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold border transition-all ${
            !isOnline || !firebaseConnected || syncStatus === 'error'
              ? 'bg-rose-50 border-rose-100 text-rose-700 animate-pulse'
              : quotaExceeded
                ? 'bg-amber-50 border-amber-100 text-amber-700'
                : 'bg-emerald-50 border-emerald-100 text-[#7bc025]'
      }`} title={!isOnline ? 'Koneksi offline, data saat ini disimpan lokal di peramban' : (!firebaseConnected || syncStatus === 'error') ? 'Gagal sinkronisasi dengan database cloud' : quotaExceeded ? 'Kuota harian Firestore habis. Sinkronisasi cloud dijeda.' : `Koneksi Cloud stabil, sinkronisasi aktif otomatis${cloudLastSync ? ' (Terakhir: ' + cloudLastSync + ')' : ''}`}>
            <span className={`w-2 h-2 rounded-full ${!isOnline || !firebaseConnected || syncStatus === 'error' ? 'bg-rose-500' : quotaExceeded ? 'bg-amber-500' : 'bg-[#8dc63f] animate-pulse'}`}></span>
            <div className="flex flex-col items-start leading-none gap-0.5">
               <span className="hidden xs:inline text-[10px] sm:text-xs">{!isOnline ? 'Internet Putus' : (!firebaseConnected || syncStatus === 'error') ? 'Cloud Gagal' : quotaExceeded ? 'Limit Tercapai' : 'Cloud Terhubung'}</span>
               <span className="xs:hidden text-[10px]">{!isOnline ? 'Offline' : (!firebaseConnected || syncStatus === 'error') ? 'Gagal' : quotaExceeded ? 'Limit' : 'Online'}</span>
               {cloudLastSync && isOnline && firebaseConnected && syncStatus !== 'error' && !quotaExceeded && (
                 <span className="text-[8px] opacity-70 hidden sm:inline">Sync {cloudLastSync}</span>
               )}
            </div>
          </div>

          <div className="flex flex-col items-center gap-0.5 ml-1">
            <button 
              onClick={() => setActiveTab('profile')} 
              className="relative w-10 h-10 rounded-full bg-slate-50 hover:bg-slate-100 text-slate-700 transition-all shadow-sm cursor-pointer flex items-center justify-center border border-slate-200 hover:scale-105 active:scale-95 overflow-hidden p-0"
              title="Menu Profil Pengguna"
            >
              {avatarUrl ? (
                <img src={avatarUrl} alt="Avatar" className="w-full h-full object-cover rounded-full" />
              ) : (
                <UserIcon className="w-5 h-5 text-slate-600" />
              )}
              <span className={`absolute -top-0.5 -right-0.5 w-2.5 h-2.5 border-2 border-white rounded-full ${isOnline ? 'bg-[#8dc63f] animate-pulse' : 'bg-rose-500'}`}></span>
            </button>
            <span className="text-[9px] font-black text-slate-600 uppercase tracking-wider leading-none">Profil</span>
          </div>

          <div className="flex flex-col items-center gap-0.5 ml-1">
            <button 
              onClick={() => setShowLogoutConfirm(true)} 
              className="w-10 h-10 rounded-full bg-rose-50 hover:bg-rose-100 text-rose-600 hover:text-rose-700 transition-all shadow-sm cursor-pointer flex items-center justify-center border border-rose-100/50 hover:scale-105 active:scale-95"
              title="Keluar dari Aplikasi"
            >
              <Power className="w-5 h-5 stroke-[2.5]" />
            </button>
            <span className="text-[9px] font-black text-rose-600 uppercase tracking-wider leading-none">Keluar</span>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 bg-slate-50 overflow-y-auto print:overflow-visible print:h-auto">
            <AnimatePresence mode="wait">
              <motion.div
                key={activeTab}
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                transition={{ duration: 0.3, ease: "easeInOut" }}
                className="h-full relative print:h-auto print:overflow-visible"
              >
                {renderContent()}
              </motion.div>
            </AnimatePresence>

        {/* Toast Notification */}
        <AnimatePresence>
          {toast && (
            <motion.div
              initial={{ opacity: 0, y: 20, scale: 0.95 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              exit={{ opacity: 0, y: 20, scale: 0.95 }}
              className="fixed inset-0 flex items-center justify-center p-4 z-50 pointer-events-none"
            >
                <div className={`px-6 py-4 rounded-2xl shadow-2xl flex items-center gap-3 border pointer-events-auto ${
                toast.type === 'success' ? 'bg-[#8dc63f] text-white border-[#8dc63f]' :
                toast.type === 'info' ? 'bg-indigo-500 text-white border-indigo-400' :
                'bg-rose-500 text-white border-rose-400'
              }`}>
                  {toast.type === 'success' && <Check className="w-5 h-5" />}
                  {toast.type === 'info' && <Info className="w-5 h-5" />}
                  {toast.type === 'error' && <AlertCircle className="w-5 h-5" />}
                  <span className="font-bold">{toast.message}</span>
                  <button onClick={() => { if(toastTimeoutRef.current) clearTimeout(toastTimeoutRef.current); setToast(null); }} className="ml-2 hover:opacity-80">
                    <X className="w-4 h-4" />
                  </button>
                </div>
            </motion.div>
          )}
        </AnimatePresence>
        
        {/* Confirmation Modal */}
        <AnimatePresence>
          {confirmationAction && (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="fixed inset-0 bg-slate-900/60 backdrop-blur-xl z-[100] flex items-center justify-center p-4"
            >
              <motion.div 
                initial={{ opacity: 0, scale: 0.9, y: 30 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.9, y: 30 }}
                transition={{ type: "spring", damping: 25, stiffness: 300 }}
                className="bg-white rounded-[3rem] p-10 max-w-sm w-full shadow-[0_40px_100px_-20px_rgba(0,0,0,0.3)] text-center relative overflow-hidden group border border-slate-100"
              >
                  {/* Decorative background elements */}
                  <div className="absolute top-0 left-0 w-full h-2 bg-rose-500"></div>
                  <div className="absolute -top-12 -right-12 w-32 h-32 bg-rose-50 rounded-full opacity-50 group-hover:scale-110 transition-transform duration-700"></div>

                  <div className="relative z-10">
                    <div className="mx-auto w-24 h-24 bg-rose-100 rounded-[2rem] flex items-center justify-center mb-8 rotate-3 group-hover:rotate-6 transition-transform">
                       <AlertTriangle className="w-12 h-12 text-rose-600" />
                    </div>
                    
                    <h3 className="text-2xl font-black text-slate-900 mb-3 tracking-tight">Konfirmasi Hapus</h3>
                    <p className="text-slate-600 text-sm mb-10 leading-relaxed font-medium">
                      Apakah Anda yakin ingin menghapus siswa <strong className="text-rose-600 font-bold">"{confirmationAction.student.name}"</strong>? Data absensi terkait mungkin juga akan terpengaruh.
                    </p>

                    <div className="flex gap-3">
                      <button 
                        onClick={() => setConfirmationAction(null)}
                        className="flex-1 bg-slate-100 text-slate-600 py-4 rounded-2xl font-black text-sm hover:bg-slate-200 transition-all active:scale-95"
                      >
                        Batal
                      </button>
                      <button 
                        onClick={confirmAction}
                        className="flex-1 bg-rose-600 text-white py-4 rounded-2xl font-black text-sm hover:bg-rose-700 hover:shadow-xl hover:shadow-rose-100 transition-all active:scale-95"
                      >
                        Ya, Hapus
                      </button>
                    </div>
                  </div>
                </motion.div>
              </motion.div>
            )}
          </AnimatePresence>

          {/* Reset Database Confirmation Modal */}
        <AnimatePresence>
          {resetModalType !== 'none' && (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4"
            >
              <motion.div 
                initial={{ opacity: 0, scale: 0.95, y: 15 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.95, y: 15 }}
                className="bg-white rounded-3xl p-6 md:p-8 shadow-2xl max-w-md w-full space-y-6 border border-slate-100 text-left"
              >
                <div className="flex items-start gap-4">
                  <div className={`p-3 rounded-2xl ${resetModalType === 'new_semester' ? 'bg-emerald-50 text-[#8dc63f]' : resetModalType === 'new_year' ? 'bg-amber-50 text-amber-600' : 'bg-rose-50 text-rose-600'}`}>
                    {resetModalType === 'new_semester' ? <RefreshCw className="w-6 h-6" /> : <AlertCircle className="w-6 h-6" />}
                  </div>
                  <div>
                    <h3 className="text-xl font-black text-slate-800">
                      {resetModalType === 'new_semester' ? 'Reset Semester Baru' : resetModalType === 'new_year' ? 'Reset Tahun Ajaran Baru' : resetModalType === 'clear_all_students' ? 'Kosongkan Semua Siswa' : 'Reset Semua Data (Total)'}
                    </h3>
                    <p className={`text-xs font-bold mt-0.5 tracking-wider uppercase ${resetModalType === 'new_semester' ? 'text-emerald-700' : 'text-slate-500'}`}>
                      {resetModalType === 'new_semester' ? '🔒 Konfirmasi Sandi Akun' : '⚠️ TINDAKAN TIDAK DAPAT DIBATALKAN'}
                    </p>
                  </div>
                </div>

                <div className="space-y-3 text-sm text-slate-600 leading-relaxed bg-slate-50 border border-slate-100 p-4 rounded-2xl">
                  {resetModalType === 'new_semester' ? (
                    <>
                      <p>
                        Anda akan mereset data sesi absensi untuk menyambut <strong>Semester Baru</strong>.
                      </p>
                      <ul className="list-disc list-inside space-y-1 text-slate-600 text-xs mt-2 pl-1 font-sans">
                        <li>Semua <strong>{attendanceSessions.length} sesi absensi</strong> semester sebelumnya akan direset bersih.</li>
                        <li>Master biodata siswa <strong>({students.length} orang)</strong> &amp; daftar kelas tetap aman terlindungi.</li>
                        <li>Status semester pada profil Anda akan diperbarui ke semester baru yang dipilih.</li>
                      </ul>
                    </>
                  ) : resetModalType === 'new_year' ? (
                    <>
                      <p>
                        Anda akan mereset database untuk menyambut <strong>Tahun Ajaran Baru</strong>.
                      </p>
                      <ul className="list-disc list-inside space-y-1 text-slate-600 text-xs mt-2 pl-1 font-sans">
                        <li>Semua <strong>{attendanceSessions.length} sesi absensi</strong> akan dihapus permanen.</li>
                        <li>Data biodata siswa <strong>({students.length} orang)</strong> &amp; daftar kelas tetap aman.</li>
                        <li>Statistik kehadiran siswa akan kembali kosong (0%).</li>
                      </ul>
                    </>
                  ) : resetModalType === 'clear_all_students' ? (
                    <>
                      <p>
                        Anda akan <strong>menghapus seluruh data siswa</strong> yang terdaftar.
                      </p>
                      <ul className="list-disc list-inside space-y-1 text-slate-600 text-xs mt-2 pl-1 mb-2 font-sans">
                        <li>Semua biodata siswa <strong>({students.length} siswa)</strong> akan terhapus.</li>
                        <li>Daftar kelas dan riwayat absensi <strong>tetap dipertahankan</strong>.</li>
                      </ul>
                    </>
                  ) : (
                    <>
                      <p>
                        Anda akan melakukan <strong>Reset Total (Hapus Semua)</strong>. Tindakan ini akan menghapus:
                      </p>
                      <ul className="list-disc list-inside space-y-1 text-slate-600 text-xs mt-2 pl-1 mb-2 font-sans">
                        <li>Semua biodata siswa <strong>({students.length} siswa)</strong>.</li>
                        <li>Semua riwayat &amp; <strong>{attendanceSessions.length} sesi absensi</strong> di cloud.</li>
                        <li>Seluruh daftar kelas yang terdaftar.</li>
                        <li>Biodata profil sekolah (Nama Guru, Mata Pelajaran, dsb).</li>
                      </ul>
                      <p className="text-xs text-rose-500 font-bold font-sans italic border-l-2 border-rose-300 pl-2">
                        Data Anda di cloud maupun cache lokal akan hilang seutuhnya dan tidak dapat dikembalikan!
                      </p>
                    </>
                  )}
                </div>

                {resetModalType === 'new_semester' && (
                  <div className="space-y-4 pt-1">
                    <div className="grid grid-cols-2 gap-3">
                      <div className="space-y-1.5">
                        <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider block">Pilih Semester Baru</label>
                        <select 
                          value={newSemesterChoice} 
                          onChange={e => setNewSemesterChoice(e.target.value as 'Ganjil' | 'Genap')}
                          className="w-full px-3 py-2.5 bg-stone-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 text-xs font-bold text-slate-800"
                        >
                          <option value="Ganjil">Semester Ganjil</option>
                          <option value="Genap">Semester Genap</option>
                        </select>
                      </div>

                      <div className="space-y-1.5">
                        <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider block">Tahun Pelajaran</label>
                        <input 
                          type="text" 
                          value={newTahunPelajaran} 
                          onChange={e => setNewTahunPelajaran(e.target.value)}
                          placeholder="Contoh: 2026/2027"
                          className="w-full px-3 py-2.5 bg-stone-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 text-xs font-bold text-slate-800"
                        />
                      </div>
                    </div>

                    <div className="space-y-1.5">
                      <label className="text-[10px] font-black text-slate-700 uppercase tracking-wider block">
                        Masukkan Password Masuk Akun Anda:
                      </label>
                      <div className="relative">
                        <input 
                          type={showResetPassword ? "text" : "password"} 
                          disabled={isResettingData || isVerifyingPassword}
                          value={resetPasswordInput}
                          onChange={e => {
                            setResetPasswordInput(e.target.value);
                            if (resetPasswordError) setResetPasswordError('');
                          }}
                          placeholder="Masukkan kata sandi akun Anda"
                          className="w-full px-4 py-3 pr-11 bg-white border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#8dc63f] focus:ring-2 focus:ring-emerald-100 transition-all text-sm font-bold text-slate-800 disabled:opacity-50"
                        />
                        <button
                          type="button"
                          onClick={() => setShowResetPassword(!showResetPassword)}
                          className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 p-1 cursor-pointer"
                        >
                          {showResetPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                        </button>
                      </div>
                      {resetPasswordError && (
                        <p className="text-xs font-bold text-rose-600 flex items-center gap-1.5 mt-1.5 bg-rose-50 p-2 rounded-lg border border-rose-200">
                          <AlertCircle className="w-4 h-4 shrink-0" />
                          {resetPasswordError}
                        </p>
                      )}
                    </div>
                  </div>
                )}

                {(resetModalType === 'everything' || resetModalType === 'clear_all_students') && (
                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-slate-600 uppercase tracking-wider block">
                      Ketik <span className="text-rose-600 font-black">{resetModalType === 'everything' ? '"RESET"' : '"HAPUS"'}</span> untuk konfirmasi
                    </label>
                    <input 
                      type="text" 
                      disabled={isResettingData}
                      value={resetConfirmInput}
                      onChange={e => setResetConfirmInput(e.target.value)}
                      placeholder={`Ketik ${resetModalType === 'everything' ? 'RESET' : 'HAPUS'}`}
                      className="w-full px-4 py-3 bg-stone-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-rose-500 focus:ring-2 focus:ring-rose-100 transition-all font-mono font-bold text-slate-800 tracking-widest text-center disabled:opacity-50"
                    />
                  </div>
                )}

                {isResettingData && (
                  <div className="space-y-2 py-2 bg-slate-50 p-4 rounded-2xl border-2 border-slate-300/50">
                    <div className="flex justify-between items-center text-[10px] font-black tracking-wider text-slate-600 uppercase font-sans">
                      <span>Proses Reset Berlangsung...</span>
                      <span className="font-mono text-slate-800 text-xs font-bold bg-white px-2 py-0.5 rounded-full border border-slate-100">{resetProgress}%</span>
                    </div>
                    <div className="w-full bg-slate-200 rounded-full h-3 overflow-hidden p-0.5 border-2 border-slate-400/30">
                      <motion.div 
                        initial={{ width: 0 }}
                        animate={{ width: `${resetProgress}%` }}
                        transition={{ duration: 0.1 }}
                        className={`h-full rounded-full ${
                          resetModalType === 'new_semester'
                            ? 'bg-[#8dc63f] shadow-[0_0_8px_rgba(141,198,63,0.5)]'
                            : resetModalType === 'new_year' 
                            ? 'bg-amber-500 shadow-[0_0_8px_rgba(245,158,11,0.5)]' 
                            : 'bg-rose-500 shadow-[0_0_8px_rgba(244,63,94,0.5)]'
                        }`}
                      />
                    </div>
                  </div>
                )}

                <div className="flex gap-4 pt-2">
                  <button 
                    disabled={isResettingData || isVerifyingPassword}
                    onClick={() => {
                      setResetModalType('none');
                      setResetConfirmInput('');
                      setResetPasswordInput('');
                      setResetPasswordError('');
                    }} 
                    className="flex-1 px-5 py-3 rounded-xl font-bold bg-slate-100 hover:bg-slate-200 text-slate-700 transition-all text-xs text-center disabled:opacity-50 cursor-pointer"
                  >
                    Batal
                  </button>
                  <button 
                    disabled={isResettingData || isVerifyingPassword ||
                      (resetModalType === 'new_semester' && !resetPasswordInput.trim()) ||
                      (resetModalType === 'everything' && resetConfirmInput.toUpperCase() !== 'RESET') ||
                      (resetModalType === 'clear_all_students' && resetConfirmInput.toUpperCase() !== 'HAPUS')
                    }
                    onClick={async () => {
                      if (resetModalType === 'new_semester') {
                        if (!resetPasswordInput.trim()) {
                          setResetPasswordError('Silakan masukkan kata sandi akun Anda.');
                          return;
                        }
                        setIsVerifyingPassword(true);
                        setResetPasswordError('');

                        let isMatch = false;
                        if (activeUserCustomData?.password && resetPasswordInput.trim() === activeUserCustomData.password.trim()) {
                          isMatch = true;
                        }

                        if (!isMatch && activeUserCustomData?.username) {
                          try {
                            const snap = await getDoc(doc(dbDefault, 'custom_accounts', activeUserCustomData.username.toLowerCase().trim()));
                            if (snap.exists() && snap.data()?.password === resetPasswordInput.trim()) {
                              isMatch = true;
                            }
                          } catch (err) {
                            console.warn('Central account lookup error:', err);
                          }
                        }

                        if (!isMatch && activeAuth.currentUser?.email) {
                          try {
                            await signInWithEmailAndPassword(auth, activeAuth.currentUser.email, resetPasswordInput.trim());
                            isMatch = true;
                          } catch (authErr) {
                            console.warn('Auth password check failed:', authErr);
                          }
                        }

                        setIsVerifyingPassword(false);

                        if (!isMatch) {
                          setResetPasswordError('Kata sandi yang Anda masukkan salah. Mohon periksa kembali.');
                          return;
                        }

                        await handleResetData('new_semester');
                      } else {
                        handleResetData(resetModalType);
                      }
                    }} 
                    className={`flex-1 px-5 py-3 rounded-xl font-bold text-white transition-all text-xs flex items-center justify-center gap-2 ${
                      resetModalType === 'new_semester'
                        ? 'bg-[#8dc63f] hover:bg-[#7bc025] shadow-[0_4px_12px_rgba(141,198,63,0.3)]'
                        : resetModalType === 'new_year' 
                        ? 'bg-amber-500 hover:bg-amber-600 shadow-[0_4px_12px_rgba(245,158,11,0.3)]' 
                        : 'bg-rose-500 hover:bg-rose-600 shadow-[0_4px_12px_rgba(244,63,94,0.3)]'
                    } disabled:opacity-50 disabled:shadow-none cursor-pointer`}
                  >
                    {isResettingData || isVerifyingPassword ? (
                      <span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></span>
                    ) : null}
                    {resetModalType === 'new_semester' ? 'Konfirmasi & Reset' : 'Ya, Reset Sekarang'}
                  </button>
                </div>
              </motion.div>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Reset Success Modal */}
        <AnimatePresence>
          {resetSuccessModal !== 'none' && (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4"
            >
              <motion.div 
                initial={{ opacity: 0, scale: 0.95, y: 15 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.95, y: 15 }}
                className="bg-white rounded-3xl p-6 md:p-8 shadow-2xl max-w-sm w-full space-y-6 border border-slate-100 text-center"
              >
                <div className="w-16 h-16 bg-emerald-50 text-[#8dc63f] rounded-full mx-auto flex items-center justify-center mb-2 animate-bounce">
                  <CheckCircle className="w-8 h-8" />
                </div>

                <div className="space-y-2">
                  <h3 className="text-xl font-black text-slate-800">
                    {resetSuccessModal === 'new_semester' ? 'Semester Baru Siap!' : resetSuccessModal === 'new_year' ? 'Tahun Ajaran Baru Siap!' : resetSuccessModal === 'clear_all_students' ? 'Kosongkan Semua Siswa Berhasil' : 'Reset Berhasil Dilakukan'}
                  </h3>
                  <p className="text-xs text-slate-600 leading-relaxed font-sans">
                    {resetSuccessModal === 'new_semester'
                      ? 'Seluruh riwayat dan sesi absensi semester lalu telah direset bersih. Biodata siswa & daftar kelas Anda tetap aman, serta semester baru telah aktif!'
                      : resetSuccessModal === 'new_year' 
                      ? 'Seluruh riwayat, kehadiran, dan sesi absensi telah dibersihkan sepenuhnya dari cloud. Daftar siswa dan kelas Anda tetap terawat dengan aman dan siap kembali digunakan!' 
                      : resetSuccessModal === 'clear_all_students'
                      ? 'Seluruh data siswa berhasil dihapus secara permanen. Daftar kelas dan riwayat absensi tidak ikut terhapus.'
                      : 'Seluruh database Anda (daftar siswa, riwayat absensi, sesi absensi, daftar kelas, dan profil sekolah) telah berhasil dihapus seutuhnya secara permanen dari server cloud.'}
                  </p>
                </div>

                <button 
                  onClick={() => setResetSuccessModal('none')} 
                  className="w-full px-5 py-3 rounded-xl font-bold bg-[#8dc63f] hover:bg-[#7bc025] text-white shadow-[0_4px_12px_rgba(5,150,105,0.3)] transition-all text-sm cursor-pointer"
                >
                  Selesai &amp; Kembali
                </button>
              </motion.div>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Student Success Modal */}
        <AnimatePresence>
          {studentSuccessModal !== 'none' && (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="fixed inset-0 bg-slate-900/60 backdrop-blur-xl z-[150] flex items-center justify-center p-4"
            >
              <motion.div 
                initial={{ opacity: 0, scale: 0.9, y: 30 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.9, y: 30 }}
                transition={{ type: "spring", damping: 25, stiffness: 300 }}
                className="bg-white rounded-[3rem] p-10 max-w-sm w-full shadow-[0_45px_110px_-20px_rgba(0,0,0,0.4)] text-center relative overflow-hidden group border border-slate-100/50"
              >
                 {/* Decorative background elements */}
                 <div className={`absolute top-0 left-0 w-full h-2.5 bg-gradient-to-r ${
                    studentSuccessModal === 'added' ? 'from-emerald-400 to-teal-400' :
                    studentSuccessModal === 'edited' ? 'from-sky-400 to-indigo-400' :
                    'from-rose-400 to-orange-400'
                 }`}></div>
                 
                 <div className={`absolute -top-12 -right-12 w-36 h-36 rounded-full opacity-40 blur-2xl transition-transform duration-1000 group-hover:scale-125 ${
                    studentSuccessModal === 'added' ? 'bg-emerald-100' :
                    studentSuccessModal === 'edited' ? 'bg-sky-100' :
                    'bg-rose-100'
                 }`}></div>

                 <div className="relative z-10">
                    <div className={`mx-auto w-24 h-24 rounded-[2.5rem] flex items-center justify-center mb-9 rotate-3 group-hover:rotate-6 transition-all duration-500 shadow-lg ${
                        studentSuccessModal === 'added' ? 'bg-emerald-50 text-[#8dc63f] shadow-emerald-100' :
                        studentSuccessModal === 'edited' ? 'bg-sky-50 text-sky-600 shadow-sky-100' :
                        'bg-rose-50 text-rose-600 shadow-rose-100'
                    }`}>
                       {studentSuccessModal === 'added' && <Users className="w-11 h-11" />}
                       {studentSuccessModal === 'edited' && <Pencil className="w-11 h-11" />}
                       {studentSuccessModal === 'deleted' && <Trash2 className="w-11 h-11" />}
                    </div>
                    
                    <h3 className="text-2xl font-black text-slate-900 mb-4 tracking-tight leading-tight">
                      {studentSuccessModal === 'added' ? 'Siswa Berhasil Ditambahkan' : 
                       studentSuccessModal === 'edited' ? 'Siswa Berhasil Diedit' : 
                       'Siswa Berhasil Dihapus'}
                    </h3>
                    <p className="text-slate-600 text-[15px] mb-11 leading-relaxed font-medium">
                      {studentSuccessModal === 'added' ? 'Data siswa baru telah aman tersimpan di cloud database sekolah Anda.' : 
                       studentSuccessModal === 'edited' ? 'Profil siswa tersebut telah berhasil diperbarui dan diselaraskan ke sistem.' : 
                       'Siswa dan data terkait telah berhasil dihilangkan dari sistem sekolah secara permanen.'}
                    </p>

                    <button 
                      onClick={() => setStudentSuccessModal('none')}
                      className={`w-full py-5 rounded-[1.5rem] font-black text-base transition-all duration-300 shadow-xl active:scale-[0.97] ${
                        studentSuccessModal === 'added' ? 'bg-[#8dc63f] hover:bg-[#7bc025] text-white shadow-emerald-200' :
                        studentSuccessModal === 'edited' ? 'bg-sky-600 hover:bg-sky-700 text-white shadow-sky-200' :
                        'bg-rose-600 hover:bg-rose-700 text-white shadow-rose-200'
                      }`}
                    >
                      Mengerti, Selesai
                    </button>
                 </div>
              </motion.div>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Modal Peringatan Hapus User - Custom Requested */}
        <AnimatePresence>
          {userToDelete && (
            <motion.div 
               initial={{ opacity: 0 }}
               animate={{ opacity: 1 }}
               exit={{ opacity: 0 }}
               className="fixed inset-0 bg-slate-900/80 backdrop-blur-sm z-[250] flex items-center justify-center p-4 animate-in fade-in duration-200"
            >
              <motion.div 
                 initial={{ opacity: 0, scale: 0.95, y: 15 }}
                 animate={{ opacity: 1, scale: 1, y: 0 }}
                 exit={{ opacity: 0, scale: 0.95, y: 15 }}
                 className="bg-white rounded-[2rem] p-6 sm:p-8 shadow-2xl max-w-md w-full space-y-6 text-center relative border border-slate-100"
              >
                <div className="absolute top-0 left-0 right-0 h-2 bg-rose-500 rounded-t-[2rem]" />

                <div className="flex justify-center pt-2">
                  <div className="w-16 h-16 bg-rose-50 border-2 border-rose-100 rounded-full flex items-center justify-center text-rose-600 shadow-sm">
                    <AlertTriangle className="w-8 h-8" />
                  </div>
                </div>

                <div className="space-y-2">
                   <h3 className="text-xl font-black tracking-tight text-slate-800">
                     Hapus Akun Pengajar?
                   </h3>
                   <p className="text-xs text-slate-600 leading-relaxed px-2">
                     Apakah Anda yakin ingin menghapus akun pengajar <b>"{userToDelete.fullname}"</b> (username: <span className="font-mono text-xs text-rose-600 font-bold">{userToDelete.username}</span>) secara permanen?
                   </p>
                </div>

                <div className="bg-rose-50 border border-rose-100 rounded-2xl p-4 text-left text-xs text-rose-700 leading-relaxed space-y-1.5">
                  <strong className="block font-black text-rose-800 uppercase tracking-wide text-[10px]">Dampak Utama:</strong>
                  <ul className="list-disc pl-4 space-y-1 text-slate-600 text-[11px]">
                    <li>Akun ini tidak akan bisa login lagi ke sistem absensi.</li>
                    <li>Data pengajar ini di sistem registrasi pusat kaguci akan dihapus selamanya.</li>
                  </ul>
                </div>

                <div className="flex flex-col sm:flex-row gap-3 pt-2">
                  <button 
                     type="button" 
                     onClick={() => setUserToDelete(null)}
                     disabled={isDeletingUser}
                     className="flex-1 py-3 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs rounded-xl font-bold transition-all active:scale-[0.98] cursor-pointer"
                   >
                     Batal
                   </button>
                   <button 
                      type="button" 
                      onClick={handleConfirmDeleteUser}
                      disabled={isDeletingUser}
                      className="flex-1 py-3 rounded-xl font-bold text-white bg-rose-600 hover:bg-rose-700 transition-all text-xs shadow-md shadow-rose-200 active:scale-[0.98] flex items-center justify-center gap-2 cursor-pointer"
                   >
                     {isDeletingUser ? (
                       <span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></span>
                     ) : null}
                     Ya, Hapus Akun
                   </button>
                 </div>
               </motion.div>
             </motion.div>
           )}
        </AnimatePresence>

        {/* Modal Edit User - Custom Requested */}
        <AnimatePresence>
          {userToEdit && (
            <motion.div 
               initial={{ opacity: 0 }}
               animate={{ opacity: 1 }}
               exit={{ opacity: 0 }}
               className="fixed inset-0 bg-slate-900/80 backdrop-blur-sm z-[250] flex items-center justify-center p-4 animate-in fade-in duration-200"
            >
              <motion.div 
                 initial={{ opacity: 0, scale: 0.95, y: 15 }}
                 animate={{ opacity: 1, scale: 1, y: 0 }}
                 exit={{ opacity: 0, scale: 0.95, y: 15 }}
                 className="bg-white rounded-[2rem] p-6 sm:p-8 shadow-2xl max-w-sm w-full space-y-6 text-left relative border border-slate-100"
              >
                <div className="absolute top-0 left-0 right-0 h-2 bg-indigo-600 rounded-t-[2rem]" />

                <div className="flex items-center gap-3 border-b border-slate-100 pb-4">
                  <div className="w-11 h-11 bg-indigo-50 border-2 border-indigo-100 rounded-full flex items-center justify-center text-indigo-600 shadow-xs">
                    <Pencil className="w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-base font-black tracking-tight text-slate-800">
                      Edit Akun Pengajar
                    </h3>
                    <p className="text-[10px] text-slate-500 font-medium">
                      Perbarui nama & sandi akun pengajar
                    </p>
                  </div>
                </div>

                <div className="space-y-4">
                  <div>
                    <label className="block text-[10px] font-extrabold text-slate-500 uppercase tracking-wider mb-1">
                      Username (ID Akun)
                    </label>
                    <input 
                       type="text" 
                       value={userToEdit.username} 
                       disabled 
                       className="w-full bg-slate-100 border-2 border-slate-300 text-slate-500 font-mono text-xs rounded-xl px-4 py-2.5 cursor-not-allowed opacity-80"
                       title="Username tidak dapat diubah"
                    />
                    <span className="text-[9px] text-slate-500 italic mt-1 block">Username/User ID bersifat permanen & unik.</span>
                  </div>

                  <div>
                    <label className="block text-[10px] font-extrabold text-slate-500 uppercase tracking-wider mb-1">
                      Nama Lengkap
                    </label>
                    <input 
                       type="text" 
                       value={editFullname} 
                       onChange={(e) => setEditFullname(e.target.value)}
                       placeholder="Contoh: Budi Santoso, S.Pd."
                       className="w-full bg-slate-50 border-2 border-slate-300 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-100 outline-hidden font-bold text-slate-800 text-xs rounded-xl px-4 py-2.5 transition-all text-slate-800"
                    />
                  </div>

                  <div>
                    <label className="block text-[10px] font-extrabold text-slate-500 uppercase tracking-wider mb-1">
                      Sandi Baru / Password
                    </label>
                    <input 
                       type="text" 
                       value={editPassword} 
                       onChange={(e) => setEditPassword(e.target.value)}
                       placeholder="Sandi minimal 4 karakter"
                       className="w-full bg-slate-50 border-2 border-slate-300 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-100 outline-hidden font-bold text-slate-800 text-xs rounded-xl px-4 py-2.5 transition-all font-mono text-slate-800"
                    />
                  </div>
                </div>

                <div className="flex flex-col sm:flex-row gap-2.5 pt-2">
                  <button 
                     type="button" 
                     onClick={() => setUserToEdit(null)}
                     disabled={isSavingUser}
                     className="flex-1 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs rounded-xl font-bold transition-all active:scale-[0.98] cursor-pointer text-center"
                  >
                    Batal
                  </button>
                  <button 
                     type="button" 
                     onClick={handleSaveEditedUser}
                     disabled={isSavingUser}
                     className="flex-1 py-2.5 rounded-xl font-bold text-white bg-indigo-600 hover:bg-indigo-700 transition-all text-xs shadow-md shadow-indigo-100 active:scale-[0.98] flex items-center justify-center gap-1.5 cursor-pointer"
                  >
                    {isSavingUser ? (
                      <span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></span>
                    ) : null}
                    Simpan Perubahan
                  </button>
                </div>
              </motion.div>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Import Result Modal */}
        <AnimatePresence>
          {importResult && importResult.isOpen && (
             <motion.div 
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm"
             >
                <motion.div 
                   initial={{ scale: 0.95, opacity: 0, y: 10 }}
                   animate={{ scale: 1, opacity: 1, y: 0 }}
                   exit={{ scale: 0.95, opacity: 0, y: 10 }}
                   className="bg-white rounded-3xl shadow-xl w-full max-w-md overflow-hidden border border-slate-100"
                >
                   {importResult.error ? (
                     // State error/gagal
                     <div className="p-6 text-center space-y-4">
                        <div className="w-16 h-16 bg-rose-50 text-rose-500 rounded-full mx-auto flex items-center justify-center mb-2">
                            <AlertCircle className="w-8 h-8" />
                        </div>
                        <h3 className="text-xl font-black text-rose-600 tracking-tight">Gagal Impor Excel</h3>
                        <p className="text-slate-600 text-xs font-medium leading-relaxed">
                          Tidak dapat mengimpor data siswa. Sistem menemukan masalah format atau database sebagai berikut:
                        </p>
                        
                        <div className="bg-rose-50 text-rose-700 rounded-2xl p-4 text-left border border-rose-100/60 text-xs font-bold leading-relaxed space-y-1 my-3">
                           <div className="text-[10px] font-black uppercase text-rose-800 tracking-wider">Pesan Detail Sistem:</div>
                           <div className="font-semibold text-slate-700 bg-white/70 rounded-lg p-2 border border-rose-200">
                             {importResult.errorMessage || 'Terjadi kesalahan tidak dikenal saat parsing file.'}
                           </div>
                        </div>

                        {importResult.details && importResult.details.length > 0 && (
                          <div className="text-left space-y-1.5 my-3">
                            <p className="text-[10px] font-black text-slate-600 uppercase tracking-wider">Riwayat Log / Catatan Pemeriksaan:</p>
                            <div className="max-h-24 overflow-y-auto rounded-xl bg-slate-50 border-2 border-slate-300/50 p-2 text-[11px] text-slate-600 font-medium space-y-1">
                              {importResult.details.map((detail, idx) => (
                                <div key={idx} className="flex gap-1 items-start">
                                  <span className="shrink-0 text-slate-500">•</span>
                                  <span>{detail}</span>
                                </div>
                              ))}
                            </div>
                          </div>
                        )}

                        <button 
                           onClick={() => setImportResult(null)}
                           className="w-full bg-slate-800 hover:bg-slate-900 text-white font-bold py-3.5 rounded-2xl shadow-sm transition-all text-sm"
                        >
                           Tutup & Perbaiki File
                        </button>
                     </div>
                   ) : (
                     // State sukses/informasi lengkap
                     <div className="p-6 text-center space-y-4">
                        <div className="w-16 h-16 bg-[#8dc63f]/10 text-[#8dc63f] rounded-full mx-auto flex items-center justify-center mb-2">
                            <CheckCircle className="w-8 h-8" />
                        </div>
                        <h3 className="text-xl font-black text-slate-800 tracking-tight">Hasil Impor Selesai</h3>
                        <p className="text-slate-600 text-xs font-medium">Proses penarikan data selesai. Berikut adalah ringkasan pemeriksaan lembar file Anda:</p>
                        
                        <div className="bg-slate-50 rounded-2xl p-4 text-left border border-slate-100 flex flex-col gap-2.5 my-4">
                           <div className="flex justify-between items-center text-sm font-bold text-slate-700">
                             <span>✅ Berhasil Diimpor (Firebase):</span>
                             <span className="text-[#8dc63f] text-base font-black">{importResult.successCount} Siswa</span>
                           </div>
                           <div className="flex justify-between items-center text-sm font-bold text-rose-500">
                             <span>❌ Gagal/Nama Kosong:</span>
                             <span className="text-rose-500 text-base font-black">{importResult.failCount} Baris</span>
                           </div>
                           <div className="flex justify-between items-center text-sm font-semibold text-slate-500">
                             <span>⏩ Baris Kosong Dilewati:</span>
                             <span>{importResult.emptyCount} Baris</span>
                           </div>
                           <div className="border-t border-slate-200/60 pt-2 flex justify-between items-center text-xs text-slate-600 font-bold">
                             <span>Total Baris Diperiksa:</span>
                             <span>{importResult.totalParsed} baris</span>
                           </div>
                        </div>

                        {/* Tampilkan statistik lembar kerja (worksheet list) jika ada */}
                        {importResult.sheetsProcessed && importResult.sheetsProcessed.length > 0 && (
                          <div className="text-left space-y-1.5 my-3">
                            <p className="text-[10px] font-black text-slate-600 uppercase tracking-wider">Daftar Lembar Kerja (Sheets) Terbaca:</p>
                            <div className="max-h-24 overflow-y-auto rounded-xl border border-slate-100/80 bg-slate-50/50 p-2 divide-y divide-slate-100 text-xs">
                              {importResult.sheetsProcessed.map((sh, idx) => (
                                <div key={idx} className="flex justify-between py-1.5 px-1 font-bold text-slate-600 first:pt-0 last:pb-0">
                                  <span className="truncate pr-2">📄 {sh.name}</span>
                                  <span className="text-[#8dc63f] shrink-0 font-black">{sh.count} siswa</span>
                                </div>
                              ))}
                            </div>
                          </div>
                        )}

                        {importResult.details && importResult.details.length > 0 && (
                          <div className="text-left space-y-1.5 my-3">
                            <p className="text-[10px] font-black text-slate-600 uppercase tracking-wider font-sans">Log Baris Dilewati / Bermasalah:</p>
                            <div className="max-h-24 overflow-y-auto rounded-xl bg-rose-50/50 border border-rose-100/60 p-2 text-[11px] text-rose-700 font-medium space-y-1 font-sans">
                              {importResult.details.map((detail, idx) => (
                                <div key={idx} className="flex gap-1 items-start">
                                  <span className="shrink-0">•</span>
                                  <span>{detail}</span>
                                </div>
                              ))}
                            </div>
                          </div>
                        )}

                        <button 
                           onClick={() => setImportResult(null)}
                           className="w-full bg-[#8dc63f] hover:bg-[#7bc025] text-white font-bold py-3.5 rounded-2xl shadow-sm transition-all focus:outline-none focus:ring-2 focus:ring-[#8dc63f]/50 text-sm"
                        >
                           Tutup & Selesai
                        </button>
                     </div>
                   )}
                </motion.div>
             </motion.div>
          )}
        </AnimatePresence>

        {/* Logout Confirm Modal */}
        <AnimatePresence>
          {showLogoutConfirm && (
             <div className="fixed inset-0 z-[200] flex items-center justify-center p-4">
               <motion.div 
                 initial={{ opacity: 0 }}
                 animate={{ opacity: 1 }}
                 exit={{ opacity: 0 }}
                 onClick={() => setShowLogoutConfirm(false)}
                 className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
               />
               <motion.div 
                 initial={{ scale: 0.95, opacity: 0, y: 10 }}
                 animate={{ scale: 1, opacity: 1, y: 0 }}
                 exit={{ scale: 0.95, opacity: 0, y: 10 }}
                 className="relative bg-white rounded-[2rem] p-8 max-w-sm w-full shadow-2xl text-center border-2 border-rose-100"
               >
                 <div className="w-16 h-16 bg-rose-100 text-rose-500 rounded-full flex items-center justify-center mx-auto mb-6 shadow-inner">
                   <LogOut className="w-8 h-8 ml-1" />
                 </div>
                 
                 <h3 className="text-xl font-black text-slate-800 mb-3 tracking-tight">Keluar Aplikasi?</h3>
                 <div className="text-slate-600 text-sm leading-relaxed mb-8 font-medium">
                   Apakah Anda yakin ingin keluar dari sesi ini? Menghentikan sesi sangat disarankan jika aplikasi sedang tidak digunakan untuk menghemat limit database.
                 </div>
                 
                 <div className="flex gap-3">
                   <button 
                     onClick={() => setShowLogoutConfirm(false)}
                     className="flex-1 py-3 bg-slate-100 hover:bg-slate-200 text-slate-600 text-sm rounded-xl font-bold transition-all active:scale-[0.98] cursor-pointer"
                   >
                     Batal
                   </button>
                   <button 
                     onClick={async () => {
                        setShowLogoutConfirm(false);
                        speakText("Anda Telah Keluar Dari Aplikasi");

                        if (activeUserCustomData?.username) {
                          localStorage.removeItem(`kaguci_profile_${activeUserCustomData.username.toLowerCase()}`);
                          localStorage.removeItem(`kaguci_avatar_${activeUserCustomData.username.toLowerCase()}`);
                          sessionStorage.removeItem(`kaguci_welcomed_${activeUserCustomData.username}`);
                        }
                        localStorage.removeItem('kaguci_active_custom_user');
                        localStorage.removeItem('kaguci_saved_credentials');
                        localStorage.removeItem('kaguci_has_logged_in');
                        safeSetLocalStorage('kaguci_isLoggedIn', 'false');
                        
                        try {
                          window.history.replaceState(null, '', window.location.pathname + window.location.search);
                        } catch {
                          window.location.hash = '';
                        }

                        setActiveUserCustomData(null);
                        setCurrentUser(null);
                        setIsLoggedIn(false);

                        try {
                          await signOut(activeAuth);
                        } catch (e) {
                          console.error("Firebase signOut error: ", e);
                        }
                        showToast('Berhasil keluar.', 'success');
                     }}
                     className="flex-1 py-3 rounded-xl font-bold text-white bg-rose-600 hover:bg-rose-700 transition-all text-sm shadow-md shadow-rose-200 active:scale-[0.98] flex items-center justify-center gap-2 cursor-pointer"
                   >
                     Keluar
                   </button>
                 </div>
               </motion.div>
             </div>
          )}
        </AnimatePresence>

        {/* Homeroom Role Warning Modal */}
        <AnimatePresence>
          {showHomeroomRoleAlert && (
             <div className="fixed inset-0 z-[200] flex items-center justify-center p-4">
               <motion.div 
                 initial={{ opacity: 0 }}
                 animate={{ opacity: 1 }}
                 exit={{ opacity: 0 }}
                 onClick={() => setShowHomeroomRoleAlert(false)}
                 className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
               />
               <motion.div 
                 initial={{ scale: 0.92, opacity: 0, y: 15 }}
                 animate={{ scale: 1, opacity: 1, y: 0 }}
                 exit={{ scale: 0.92, opacity: 0, y: 15 }}
                 className="relative bg-white rounded-[2.5rem] p-6 sm:p-8 max-w-md w-full shadow-2xl text-center border-2 border-amber-100/90 overflow-hidden"
               >
                 {/* Top Close Button */}
                 <button 
                   onClick={() => setShowHomeroomRoleAlert(false)}
                   className="absolute top-4 right-4 p-2 text-slate-400 hover:text-slate-600 hover:bg-slate-100 rounded-full transition-colors cursor-pointer"
                   title="Tutup"
                 >
                   <X className="w-5 h-5" />
                 </button>

                 {/* Decorative background glow */}
                 <div className="absolute -top-16 -right-16 w-36 h-36 bg-amber-100/60 rounded-full blur-2xl pointer-events-none" />
                 <div className="absolute -bottom-16 -left-16 w-36 h-36 bg-emerald-100/60 rounded-full blur-2xl pointer-events-none" />

                 {/* Icon Header */}
                 <div className="relative mx-auto mb-4 w-20 h-20 rounded-3xl bg-gradient-to-tr from-amber-500 to-amber-400 text-white flex items-center justify-center shadow-lg shadow-amber-500/25">
                   <GraduationCap className="w-10 h-10" />
                   <div className="absolute -top-1 -right-1 w-6 h-6 bg-white rounded-full flex items-center justify-center text-amber-600 shadow-sm border border-amber-200">
                     <AlertCircle className="w-4 h-4" />
                   </div>
                 </div>

                 {/* Badge & Title */}
                 <div className="inline-flex items-center gap-1.5 px-3.5 py-1 bg-amber-50 border border-amber-200/70 text-amber-800 rounded-full text-xs font-black uppercase tracking-wider mb-2.5">
                   <Sparkles className="w-3.5 h-3.5 text-amber-500" />
                   Akses Khusus Wali Kelas
                 </div>
                 
                 <h3 className="text-xl sm:text-2xl font-black text-slate-800 tracking-tight mb-2.5">
                   Fitur Laporan Wali Kelas
                 </h3>
                 
                 {/* Notice message */}
                 <p className="text-slate-600 text-sm font-medium leading-relaxed mb-5">
                   Fitur ini hanya untuk wali kelas. Rubah peran jadi wali kelas di menu <strong>Profil Pengguna</strong> agar dapat mengakses seluruh data dan laporan wali kelas.
                 </p>

                 {/* Role Status Box */}
                 <div className="bg-slate-50/90 border border-slate-200/80 rounded-2xl p-3.5 mb-6 text-left space-y-2">
                   <div className="flex items-center justify-between text-xs font-bold text-slate-500">
                     <span>Peran Anda Saat Ini:</span>
                     <span className="px-2.5 py-0.5 bg-slate-200 text-slate-700 rounded-lg font-black">
                       {profileData.role || 'Guru Mapel'}
                     </span>
                   </div>
                   <div className="flex items-center justify-between text-xs font-bold text-amber-800 bg-amber-100/70 rounded-xl p-2.5">
                     <span className="flex items-center gap-1.5">
                       <CheckCircle className="w-4 h-4 text-amber-600 shrink-0" />
                       Peran Yang Dibutuhkan:
                     </span>
                     <span className="font-black text-amber-900 uppercase">
                       Wali Kelas
                     </span>
                   </div>
                 </div>

                 {/* Action Buttons */}
                 <div className="flex flex-col gap-2.5">
                   <button 
                     onClick={() => {
                       setShowHomeroomRoleAlert(false);
                       setActiveTab('profile');
                       setIsProfileEditing(true);
                     }}
                     className="w-full py-3.5 rounded-2xl font-black text-white bg-gradient-to-r from-amber-500 to-amber-600 hover:from-amber-600 hover:to-amber-700 transition-all text-sm shadow-md shadow-amber-500/25 active:scale-[0.98] flex items-center justify-center gap-2 cursor-pointer"
                   >
                     <Settings2 className="w-4 h-4" />
                     <span>Buka Pengaturan Profil</span>
                   </button>

                   <button 
                     onClick={() => setShowHomeroomRoleAlert(false)}
                     className="w-full py-2.5 text-slate-500 hover:text-slate-700 font-bold text-xs transition-colors cursor-pointer bg-slate-100 hover:bg-slate-200 rounded-xl"
                   >
                     Tutup
                   </button>
                 </div>
               </motion.div>
             </div>
          )}
        </AnimatePresence>

        {/* Quota Exceeded Professional Popup */}
        <AnimatePresence>
          {showQuotaModal && (
            <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 sm:p-6">
              <motion.div 
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                onClick={() => setShowQuotaModal(false)}
                className="absolute inset-0 bg-slate-900/60 backdrop-blur-md"
              />
              
              <motion.div
                initial={{ scale: 0.9, opacity: 0, y: 20 }}
                animate={{ scale: 1, opacity: 1, y: 0 }}
                exit={{ scale: 0.9, opacity: 0, y: 20 }}
                className="relative w-full max-w-lg bg-white rounded-[2.5rem] p-8 shadow-2xl overflow-hidden border border-white"
              >
                {/* Abstract decorative elements */}
                <div className="absolute -top-24 -right-24 w-48 h-48 bg-amber-100 rounded-full blur-3xl opacity-50" />
                <div className="absolute -bottom-24 -left-24 w-48 h-48 bg-blue-100 rounded-full blur-3xl opacity-50" />
                
                <div className="relative z-10 flex flex-col items-center text-center">
                  <div className="w-20 h-20 bg-amber-50 rounded-3xl flex items-center justify-center mb-6 shadow-sm border border-amber-100">
                    <Database className="w-10 h-10 text-amber-500" />
                  </div>
                  
                  <h3 className="text-2xl font-black text-slate-800 mb-4 leading-tight">
                    Kapasitas Layanan <br/><span className="text-rose-600 uppercase tracking-tighter">Mencapai Batas</span>
                  </h3>
                  
                  <p className="text-slate-600 text-sm leading-relaxed mb-8 font-medium">
                    Layanan sinkronisasi database (Firestore) telah mencapai batas penggunaan <span className="text-rose-600 font-bold">Free Quota</span> harian. 
                    <br/><br/>
                    <div className="bg-rose-50 border border-rose-100 p-4 rounded-2xl text-left mb-4">
                      <p className="text-xs text-rose-800 font-bold mb-1 underline">Dampak Saat Ini (Mode Offline Caching):</p>
                      <ul className="text-xs text-rose-700 space-y-1 list-disc pl-4 mb-3">
                        <li>Fungsi <strong>Reset Database</strong> tidak akan tersimpan di cloud.</li>
                        <li>Update data (Absensi, Siswa, Profil) tidak akan sinkron ke perangkat lain.</li>
                        <li>Aplikasi berjalan sepenuhnya dari penyimpanan sementara (Cache / IndexedDB).</li>
                      </ul>
                      <div className="p-2 border-l-4 border-amber-500 bg-amber-50 rounded text-amber-800 text-[11px] font-bold">
                        ⚠️ PERINGATAN PENTING:<br/>
                        Selama kuota habis, MOHON JANGAN membersihkan histori browser (Clear Data / Clear Cache). Jika cache dihapus, maka data yang belum terkirim ke server akan HILANG permanen.
                      </div>
                    </div>
                    Jangan khawatir! Data baru Anda tetap <span className="font-bold text-slate-800">tersimpan sementara di memori perangkat ini</span>, dan akan dilanjutkan besok pagi saat kuota diatur ulang otomatis oleh Google.
                  </p>
                  
                  <div className="w-full flex flex-col gap-3">
                    <button 
                      onClick={() => setShowQuotaModal(false)}
                      className="w-full bg-slate-800 hover:bg-slate-900 text-white font-bold py-4 rounded-2xl transition-all shadow-lg hover:shadow-xl active:scale-95"
                    >
                      Saya Mengerti
                    </button>
                    <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest">
                      Infrastruktur didukung oleh Google Firebase Free Tier
                    </p>
                  </div>
                </div>
              </motion.div>
            </div>
          )}
        </AnimatePresence>

        {/* Global Account Deleted Alert Modal */}
        <AnimatePresence>
          {accountDeletedAlert && (
             <div className="fixed inset-0 z-[200] flex items-center justify-center p-4">
              <motion.div 
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
              />
              <motion.div 
                initial={{ scale: 0.9, opacity: 0, y: 20 }}
                animate={{ scale: 1, opacity: 1, y: 0 }}
                exit={{ scale: 0.9, opacity: 0, y: 20 }}
                className="relative bg-white rounded-[2rem] p-8 max-w-sm w-full shadow-2xl text-center border-2 border-rose-100"
              >
                <div className="w-16 h-16 bg-rose-100 text-rose-500 rounded-full flex items-center justify-center mx-auto mb-6 shadow-inner">
                  <UserX className="w-8 h-8" />
                </div>
                <h3 className="text-xl font-black text-slate-800 mb-3 tracking-tight">Akun Tidak Terdaftar!</h3>
                <div className="text-slate-600 text-sm leading-relaxed mb-8 font-medium">
                  Sistem mendeteksi bahwa akun Anda telah dihapus oleh Administrator. Sesi Anda dihentikan secara otomatis untuk alasan keamanan.
                </div>
                <button 
                  onClick={() => setAccountDeletedAlert(false)}
                  className="w-full bg-rose-500 hover:bg-rose-600 text-white font-bold py-4 rounded-xl transition-all shadow-lg hover:shadow-xl active:scale-95"
                >
                  Tutup
                </button>
              </motion.div>
            </div>
          )}
        </AnimatePresence>
      </main>


      {/* Bottom Nav */}
      <footer className="bg-white border-t border-slate-100/80 px-2 py-3 grid grid-cols-4 gap-1.5 z-40 shadow-[0_-4px_22px_rgba(148,163,184,0.06)] shrink-0 print:hidden">
        {menuItems.map((item) => {
          const Icon = item.icon;
          const isActive = activeTab === item.id;
          return (
            <button
              key={item.id}
              onClick={() => {
                  setActiveTab(item.id as 'dashboard' | 'students' | 'attendance' | 'reports' | 'homeroom_report' | 'profile');
              }}
              className={`flex flex-col items-center justify-center gap-1.5 py-1.5 px-0.5 rounded-2xl text-xs font-bold transition-all duration-300 transform active:scale-95 group relative ${
                isActive 
                  ? `${item.activeBg} font-black` 
                  : `${item.inactiveColor} hover:text-slate-800 ${item.hoverBg}`
              }`}
            >
              <div className={`p-1.5 sm:p-2.5 rounded-xl transition-all duration-300 ${
                isActive 
                  ? `${item.iconBg} scale-110 shadow-lg shadow-current/15` 
                  : `${item.inactiveBg} group-hover:bg-slate-100 group-hover:text-slate-600`
              }`}>
                <Icon className="w-5 h-5 stroke-[2.3]" />
              </div>
              <span className={`text-[9px] sm:text-[10px] tracking-tight font-black mt-0.5 ${!isActive ? item.inactiveColor : ''}`}>{item.label}</span>
            </button>
          );
        })}
      </footer>
    </div>
  );
}

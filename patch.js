import fs from 'fs';
let content = fs.readFileSync('src/App.tsx', 'utf-8');

content = content.replace(
`  Power,
  Loader2
} from 'lucide-react';`,
`  Power,
  Loader2,
  Send
} from 'lucide-react';`
);

content = content.replace(
`    } catch (err) {
      showToast('Gagal menghapus presensi.', 'error');
      handleFirestoreError(err, OperationType.DELETE, 'attendanceSessions');
    }
  };

  const handleSave = async () => {`,
`    } catch (err) {
      showToast('Gagal menghapus presensi.', 'error');
      handleFirestoreError(err, OperationType.DELETE, 'attendanceSessions');
    }
  };

  const handleShareWA = () => {
    if (!existingSession) return;
    
    const absents = studentsInClass.filter(s => currentRecords[s.id] !== 'Hadir');
    
    let text = '*Laporan Presensi*\\n';
    text += 'Kelas: ' + selectedClass + '\\n';
    text += 'Tanggal: ' + format(new Date(date), 'dd/MM/yyyy') + '\\n';
    text += 'Pertemuan ke: ' + meetingNumber + '\\n\\n';
    
    if (absents.length === 0) {
      text += '_Semua siswa hadir._';
    } else {
      text += '*Siswa yang tidak hadir:*\\n';
      absents.forEach((s, idx) => {
        text += (idx + 1) + '. ' + s.name + ' (' + currentRecords[s.id] + ')\\n';
      });
    }
    
    const encodedText = encodeURIComponent(text);
    window.open('https://wa.me/?text=' + encodedText, '_blank');
  };

  const handleSave = async () => {`
);

content = content.replace(
`                 <>
                    <button
                       onClick={() => setIsEditing(true)}
                       className="w-full sm:w-auto bg-white border-2 border-slate-300 text-slate-700 px-6 py-2.5 rounded-lg font-semibold text-sm hover:bg-slate-50 transition-all flex items-center justify-center gap-2"
                    >
                       <Pencil className="w-4 h-4" /> Edit
                    </button>
                    <button
                       onClick={() => { if(existingSession) handleDeleteSession(existingSession.id); }}
                       className="w-full sm:w-auto bg-white text-rose-600 border border-rose-200 px-6 py-2.5 rounded-lg font-semibold text-sm hover:bg-rose-50 transition-all active:scale-95 flex items-center justify-center gap-2"
                    >
                       <Trash2 className="w-4 h-4" /> Hapus
                    </button>
                 </>`,
`                 <>
                    <button
                       onClick={() => setIsEditing(true)}
                       className="w-full sm:w-auto bg-white border-2 border-slate-300 text-slate-700 px-6 py-2.5 rounded-lg font-semibold text-sm hover:bg-slate-50 transition-all flex items-center justify-center gap-2"
                    >
                       <Pencil className="w-4 h-4" /> Edit
                    </button>
                    <button
                       onClick={handleShareWA}
                       className="w-full sm:w-auto bg-[#25D366] text-white px-6 py-2.5 rounded-lg font-semibold text-sm hover:bg-[#128C7E] transition-all active:scale-95 flex items-center justify-center gap-2 shadow-sm"
                    >
                       <Send className="w-4 h-4" /> Kirim ke WA
                    </button>
                    <button
                       onClick={() => { if(existingSession) handleDeleteSession(existingSession.id); }}
                       className="w-full sm:w-auto bg-white text-rose-600 border border-rose-200 px-6 py-2.5 rounded-lg font-semibold text-sm hover:bg-rose-50 transition-all active:scale-95 flex items-center justify-center gap-2"
                    >
                       <Trash2 className="w-4 h-4" /> Hapus
                    </button>
                 </>`
);

fs.writeFileSync('src/App.tsx', content);

import fs from 'fs';
let content = fs.readFileSync('src/App.tsx', 'utf-8');

content = content.replace(
`          <div className="p-6 border-t border-slate-100 bg-white flex justify-end gap-4">
             <div className="flex items-center gap-2 w-full sm:w-auto">
               {isEditing ? (
                 <>
                    {existingSession && (
                       <button
                          onClick={() => {
                             setIsEditing(false);
                             setCurrentRecords(existingSession.records);
                          }}
                          className="w-full sm:w-auto bg-white border-2 border-slate-300 text-slate-700 px-6 py-2.5 rounded-lg font-semibold text-sm hover:bg-slate-50 transition-all active:scale-95"
                       >
                          Batal
                       </button>
                    )}
                    <button 
                       onClick={handleSave} 
                       className="w-full sm:w-auto bg-emerald-600 text-white px-6 py-2.5 rounded-lg font-semibold text-sm hover:bg-emerald-700 transition-all active:scale-95 flex items-center justify-center gap-2"
                    >
                       <Save className="w-4 h-4" /> Simpan Presensi
                    </button>
                 </>
               ) : (
                 <>
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
                 </>
               )}
             </div>
          </div>`,
`          <div className="p-6 border-t border-slate-100 bg-white flex flex-col gap-3">
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
          </div>`
);

fs.writeFileSync('src/App.tsx', content);

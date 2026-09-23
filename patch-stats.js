import fs from 'fs';
let content = fs.readFileSync('src/App.tsx', 'utf-8');

// Update attentionSet to include all absent students (Alpa, Sakit, Izin)
content = content.replace(
`            if (status === 'Alpa') {
              attentionSet.add(studentId);
            }`,
`            if (status === 'Alpa' || status === 'Sakit' || status === 'Izin') {
              attentionSet.add(studentId);
            }`
);

// Update result type
content = content.replace(
`    const result: Array<{
      className: string;
      maxAlpaStudent: { name: string; count: number } | null;
      maxSakitIzinStudent: { name: string; sakitCount: number; izinCount: number; totalCount: number } | null;
      allAbsenceList: Array<{`,
`    const result: Array<{
      className: string;
      totalAlpaStudents: number;
      totalSakitStudents: number;
      totalIzinStudents: number;
      allAbsenceList: Array<{`
);

// Update result.push
content = content.replace(
`      result.push({
        className,
        maxAlpaStudent: maxAlpaS && maxAlpaS.alpa > 0 ? { name: maxAlpaS.name, count: maxAlpaS.alpa } : null,
        maxSakitIzinStudent: maxSIS && (maxSIS.sakit + maxSIS.izin) > 0 ? { 
          name: maxSIS.name, 
          sakitCount: maxSIS.sakit, 
          izinCount: maxSIS.izin, 
          totalCount: maxSIS.sakit + maxSIS.izin 
        } : null,
        allAbsenceList: absentStudents
      });`,
`      const totalAlpaStudents = list.filter(item => item.alpa > 0).length;
      const totalSakitStudents = list.filter(item => item.sakit > 0).length;
      const totalIzinStudents = list.filter(item => item.izin > 0).length;

      result.push({
        className,
        totalAlpaStudents,
        totalSakitStudents,
        totalIzinStudents,
        allAbsenceList: absentStudents
      });`
);

// Now update the UI for Ringkasan Kelas
content = content.replace(
`                            <div className="space-y-3.5">
                              {/* Alpa Terbanyak */}
                              <div className="bg-white p-3.5 rounded-xl border border-slate-100/80">
                                <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-2">Siswa Alpa Terbanyak</p>
                                {hasAlpa ? (
                                  <div className="flex justify-between items-center">
                                    <span className="font-bold text-sm text-slate-800 truncate max-w-[150px]">
                                      {clsData.maxAlpaStudent?.name}
                                    </span>
                                    <span className="bg-rose-50 text-rose-600 px-2.5 py-1 rounded-lg text-xs font-black shrink-0 border border-rose-100 flex items-center gap-1">
                                      <span className="w-1.5 h-1.5 bg-rose-500 rounded-full animate-pulse"></span>
                                      {clsData.maxAlpaStudent?.count} Alpa
                                    </span>
                                  </div>
                                ) : (
                                  <span className="text-xs text-slate-500 italic font-medium">Tidak ada siswa alpa</span>
                                )}
                              </div>

                              {/* Sakit/Izin Terbanyak */}
                              <div className="bg-white p-3.5 rounded-xl border border-slate-100/80">
                                <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-2">Sering Sakit / Izin</p>
                                {hasSakitIzin ? (
                                  <div className="flex justify-between items-center">
                                    <span className="font-bold text-sm text-slate-800 truncate max-w-[130px]" title={clsData.maxSakitIzinStudent?.name}>
                                      {clsData.maxSakitIzinStudent?.name}
                                    </span>
                                    <div className="flex items-center gap-1.5 shrink-0">
                                      {clsData.maxSakitIzinStudent?.sakitCount ? (
                                        <span className="bg-amber-50 text-amber-600 px-2 py-0.5 rounded text-[10px] font-extrabold border border-amber-100">
                                          {clsData.maxSakitIzinStudent.sakitCount}S
                                        </span>
                                      ) : null}
                                      {clsData.maxSakitIzinStudent?.izinCount ? (
                                        <span className="bg-indigo-50 text-indigo-600 px-2 py-0.5 rounded text-[10px] font-extrabold border border-indigo-100">
                                          {clsData.maxSakitIzinStudent.izinCount}I
                                        </span>
                                      ) : null}
                                      <span className="bg-slate-100 text-slate-600 px-2 py-0.5 rounded text-[10px] font-extrabold border-2 border-slate-300">
                                        Total: {clsData.maxSakitIzinStudent?.totalCount}
                                      </span>
                                    </div>
                                  </div>
                                ) : (
                                  <span className="text-xs text-slate-500 italic font-medium">Tidak ada sakit/izin</span>
                                )}
                              </div>
                            </div>`,
`                            <div className="space-y-3.5">
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
                            </div>`
);

// Remove unused variables
content = content.replace(
`                      const hasAlpa = clsData.maxAlpaStudent && clsData.maxAlpaStudent.count > 0;
                      const hasSakitIzin = clsData.maxSakitIzinStudent && clsData.maxSakitIzinStudent.totalCount > 0;`,
``
);

// Also remove maxAlpaS logic which is now unused
content = content.replace(
`      let maxAlpaVal = 0;
      let maxAlpaS: StudentAbsenceDetail | null = null;
      
      let maxSIVal = 0;
      let maxSIS: StudentAbsenceDetail | null = null;

      for (const item of list) {
        if (item.alpa > maxAlpaVal) {
          maxAlpaVal = item.alpa;
          maxAlpaS = item;
        }
        const siVal = item.sakit + item.izin;
        if (siVal > maxSIVal) {
          maxSIVal = siVal;
          maxSIS = item;
        }
      }`,
``
);

fs.writeFileSync('src/App.tsx', content);

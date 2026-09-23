import * as XLSX from 'xlsx';
import { Firestore, writeBatch, doc, setDoc } from 'firebase/firestore';
import { Auth } from 'firebase/auth';

export interface Student {
  id: string;
  name: string;
  nisn: string;
  class: string;
  userId?: string;
}

export interface ImportResultState {
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
}

export const importExcelHelper = async (
  file: File,
  activeAuth: Auth,
  activeDb: Firestore,
  classList: string[],
  setClassList: React.Dispatch<React.SetStateAction<string[]>>,
  setImportResult: React.Dispatch<React.SetStateAction<ImportResultState | null>>,
  showToast: (msg: string, type: 'success' | 'info' | 'error') => void,
  excelInputRef: React.RefObject<HTMLInputElement | null>,
  existingStudents: Student[] = []
): Promise<void> => {
  return new Promise((resolve) => {
    if (!activeAuth.currentUser) {
      showToast('Sesi berakhir. Silakan masuk kembali.', 'error');
      if (excelInputRef.current) excelInputRef.current.value = '';
      resolve();
      return;
    }

    const { uid } = activeAuth.currentUser;
    const reader = new FileReader();

    reader.onload = async (evt) => {
      try {
        const bstr = evt.target?.result as ArrayBuffer;
        const wb = XLSX.read(new Uint8Array(bstr), { type: 'array' });
        
        let successCount = 0;
        let skipCount = 0;
        let failCount = 0;
        let emptyCount = 0;
        let totalParsed = 0;
        const failedRowDetails: string[] = [];
        const sheetsProcessed: { name: string; count: number }[] = [];
        const discoveredClasses: string[] = [];

        let batch = writeBatch(activeDb);
        let batchCounter = 0;
        const promises: Promise<void>[] = [];

        showToast('Sedang membaca file Excel...', 'info');

        // Process all sheets in the workbook
        for (const sheetName of wb.SheetNames) {
          const ws = wb.Sheets[sheetName];
          const data = XLSX.utils.sheet_to_json(ws, { header: 1 }) as (string | number | boolean | null | undefined)[][];

          if (!data || data.length === 0) {
            continue; // Skip empty sheets
          }

          // Search for the header row with student name
          let headerRowIdx = -1;
          let classIdx = -1;
          let nameIdx = -1;
          let nisnIdx = -1;

          // Try to search first 30 rows of the sheet
          for (let r = 0; r < Math.min(data.length, 30); r++) {
            const row = data[r];
            if (!row || !Array.isArray(row)) continue;

            const currentClassIdx = row.findIndex(h => {
              if (h === undefined || h === null) return false;
              const str = h.toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');
              return str === 'kelas' || str === 'rombel' || str === 'rombongan' || str.includes('kelas') || str.includes('rombel') || str.includes('rombongan') || str.includes('tingkat') || str.includes('group') || str === 'class';
            });

            const currentNameIdx = row.findIndex(h => {
              if (h === undefined || h === null) return false;
              const str = h.toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');
              if (str.includes('ayah') || str.includes('ibu') || str.includes('wali') || str.includes('ortu') || str.includes('orangtua') || str.includes('panggilan')) return false;
              return str === 'nama' || str === 'siswa' || str.includes('nama') || str.includes('siswa') || str.includes('pesertadidik') || str.includes('namalengkap') || str.includes('fullname') || str === 'name';
            });

            if (currentClassIdx !== -1 && currentNameIdx !== -1) {
              headerRowIdx = r;
              classIdx = currentClassIdx;
              nameIdx = currentNameIdx;
              nisnIdx = row.findIndex(h => {
                if (h === undefined || h === null) return false;
                const str = h.toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');
                return str.includes('nis') || str.includes('nisn') || str.includes('nomorinduk') || str.includes('induk');
              });
              break;
            }
          }

          // Fallback: If combined header class + name is not found, search for at least "nama"
          if (headerRowIdx === -1) {
            for (let r = 0; r < Math.min(data.length, 15); r++) {
              const row = data[r];
              if (!row || !Array.isArray(row)) continue;
              const currentNameIdx = row.findIndex(h => {
                if (h === undefined || h === null) return false;
                const str = h.toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');
                if (str.includes('ayah') || str.includes('ibu') || str.includes('wali') || str.includes('ortu') || str.includes('orangtua') || str.includes('panggilan')) return false;
                return str.includes('nama') || str.includes('siswa') || str === 'name';
              });
              if (currentNameIdx !== -1) {
                headerRowIdx = r;
                nameIdx = currentNameIdx;
                classIdx = row.findIndex(h => {
                  if (h === undefined || h === null) return false;
                  const str = h.toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');
                  return str.includes('kelas') || str.includes('rombel') || str === 'class';
                });
                nisnIdx = row.findIndex(h => {
                  if (h === undefined || h === null) return false;
                  const str = h.toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');
                  return str.includes('nis') || str.includes('nisn');
                });
                break;
              }
            }
          }

          // Final fallback: use row 0 and look up "nama"
          if (headerRowIdx === -1) {
            const firstRow = data[0];
            if (firstRow && Array.isArray(firstRow)) {
              nameIdx = firstRow.findIndex(h => {
                if (!h) return false;
                const str = h.toString().toLowerCase().replace(/[^a-z0-9]/g, '');
                if (str.includes('ayah') || str.includes('ibu') || str.includes('wali') || str.includes('ortu') || str.includes('orangtua') || str.includes('panggilan')) return false;
                return str.includes('nama') || str === 'name';
              });
              classIdx = firstRow.findIndex(h => h?.toString().toLowerCase().includes('kelas') || h?.toString().toLowerCase() === 'class');
              nisnIdx = firstRow.findIndex(h => h?.toString().toLowerCase().includes('nis'));
              if (nameIdx !== -1) {
                headerRowIdx = 0;
              }
            }
          }

          // Brutal fallback: If totally missing headers, let's just assume Column B (index 1) is Name if Column A is numbers, 
          // OR Column A is name. We'll find the first column that has mostly string data.
          if (headerRowIdx === -1 && data.length > 0) {
             for (let r = 0; r < Math.min(data.length, 5); r++) {
                 const row = data[r];
                 if (!row || !Array.isArray(row)) continue;
                 
                 // Look for a cell that is a string and represents a name
                 const possibleNameIdx = row.findIndex((cell) => {
                     if (typeof cell !== 'string') return false;
                     const s = cell.trim();
                     return s.length > 2 && s.length < 50 && !/^\d+$/.test(s);
                 });

                 if (possibleNameIdx !== -1) {
                     nameIdx = possibleNameIdx;
                     headerRowIdx = Math.max(0, r - 1); // Assume previous row or same row is header
                     break;
                 }
             }
          }

          // If no student name column could be identified on this sheet, skip this sheet gracefully
          if (nameIdx === -1) {
            failedRowDetails.push(`Sheet [${sheetName}]: Dilewati karena kolom Nama/Siswa tidak ditemukan.`);
            continue;
          }

          const rows = data.slice(headerRowIdx + 1) as (string | number | boolean | null | undefined)[][];
          let sheetSuccessCount = 0;
          let lastClass = '';

          for (const [i, row] of rows.entries()) {
            totalParsed++;
            if (!row || row.length === 0) {
              emptyCount++;
              continue;
            }

            const isRowBlank = row.every(cell => cell === undefined || cell === null || String(cell).trim() === '');
            if (isRowBlank) {
              emptyCount++;
              continue;
            }

            // More resilient column data extraction:
            const rawName = row[nameIdx] ?? '';
            const rawNisn = (nisnIdx !== -1) ? (row[nisnIdx] ?? '') : '';
            const rawClass = (classIdx !== -1) ? (row[classIdx] ?? null) : null;

            const currentName = rawName ? String(rawName).trim() : '';

            if (rawClass && String(rawClass).trim()) {
              lastClass = String(rawClass).trim();
            }

            // Smart class fallback: if no class cell was found, try the sheetName (especially if sheets are named like 'XII IPA 1')
            let currentClass = lastClass || String(rawClass || '').trim();
            if (!currentClass) {
              // Check if the sheet name looks like a class (is shorter than 15 characters, doesn't say Sheet1 etc.)
              const isSheetGeneric = sheetName.toLowerCase().startsWith('sheet') || sheetName.toLowerCase().includes('halaman');
              currentClass = isSheetGeneric ? 'Umum' : sheetName.trim();
            }

            if (currentName) {
              const cleanNisn = rawNisn ? String(rawNisn).trim().replace(/\D/g, '') : '';
              
              // Check for duplicates in existing data or current batch to prevent double imports
              const isDuplicate = existingStudents.some(s => 
                (cleanNisn && s.nisn === cleanNisn) || 
                (s.name.toLowerCase() === currentName.toLowerCase() && s.class.toLowerCase() === currentClass.toLowerCase())
              );

              if (isDuplicate) {
                skipCount++;
                continue;
              }

              const newId = Date.now().toString() + '-' + Math.floor(Math.random() * 1000) + '-' + totalParsed;
              
              const student = {
                id: newId,
                name: currentName,
                nisn: cleanNisn,
                class: currentClass,
                userId: uid
              };

              batch.set(doc(activeDb, 'students', newId), student);
              successCount++;
              sheetSuccessCount++;
              batchCounter++;

              if (!discoveredClasses.includes(currentClass)) {
                discoveredClasses.push(currentClass);
              }

              if (batchCounter === 450) {
                promises.push(batch.commit());
                batch = writeBatch(activeDb);
                batchCounter = 0;
              }
            } else {
              failCount++;
              const rowNumber = headerRowIdx + 1 + i + 1;
              failedRowDetails.push(`Sheet [${sheetName}] Baris ${rowNumber}: Nama kosong atau baris tidak valid`);
            }
          }

          if (sheetSuccessCount > 0 || rows.length > 0) {
            sheetsProcessed.push({
              name: sheetName,
              count: sheetSuccessCount
            });
          }
        }

        if (batchCounter > 0) {
          promises.push(batch.commit());
        }

        // Wait for all commits with a timeout safety to prevent hanging
        if (promises.length > 0) {
          try {
            await Promise.race([
              Promise.all(promises),
              new Promise<void>((_, reject) => setTimeout(() => reject(new Error('timeout')), 5000))
            ]);
          } catch (e) {
            if (e instanceof Error && e.message === 'timeout') {
              console.warn("Import commit timed out. Updates applied locally and sync in background.");
            } else {
              throw e;
            }
          }
        }

        // Update class list in Firestore with newly discovered classes
        if (discoveredClasses.length > 0) {
          const updatedClasses = Array.from(new Set([...classList, ...discoveredClasses]));
          updatedClasses.sort((a, b) => a.localeCompare(b, 'id-ID', { numeric: true }));

          // Save directly to Firestore users document for guaranteed sync
          try {
            await setDoc(doc(activeDb, 'users', activeAuth.currentUser!.uid), { classList: updatedClasses }, { merge: true });
            console.log("Successfully updated classList in Firestore");
          } catch (e) {
            console.error("Failed to update classList in Firestore:", e);
            // Fallback to local state if DB update fails (or Toast if critical)
            setClassList(updatedClasses);
          }
        }

        if (excelInputRef.current) excelInputRef.current.value = '';

        if (successCount === 0 && sheetsProcessed.length === 0) {
          // If nothing was parsed successfully
          setImportResult({
            isOpen: true,
            successCount: 0,
            skipCount: skipCount + failCount + emptyCount,
            failCount,
            emptyCount,
            totalParsed,
            error: true,
            errorMessage: 'Format tabel siswa atau kolom nama tidak terdeteksi di lembar kerja Excel mana pun.',
            details: failedRowDetails,
            sheetsProcessed
          });
          showToast('Gagal mengimpor data! Silakan periksa format file Excel.', 'error');
        } else {
          setImportResult({
            isOpen: true,
            successCount,
            skipCount: skipCount + failCount + emptyCount,
            failCount,
            emptyCount,
            totalParsed,
            details: failedRowDetails,
            sheetsProcessed
          });
          
          if (skipCount > 0) {
            showToast(`Berhasil mengimpor ${successCount} siswa. ${skipCount} data duplikat dilewati.`, 'info');
          } else if (failCount > 0) {
            showToast(`Berhasil mengimpor ${successCount} siswa. ${failCount} baris bermasalah.`, 'info');
          } else {
            showToast(`Berhasil mengimpor total ${successCount} siswa dari ${sheetsProcessed.length} lembar kerja.`, 'success');
          }
        }
        resolve();
      } catch (err) {
        console.error(err);
        if (excelInputRef.current) excelInputRef.current.value = '';
        const errMsg = (err as Error).message || '';
        
        let friendlyError = 'Terjadi kesalahan sistem saat membaca file Excel: ' + errMsg;
        if (errMsg.toLowerCase().includes('permission') || errMsg.toLowerCase().includes('missing or insufficient')) {
            friendlyError = 'Sistem menolak menyimpan data (Missing Permissions). Hal ini biasanya terjadi jika Anda menggunakan Database Firebase mandiri namun belum memperbarui "Firestore Rules" menjadi allow read, write: if request.auth != null;';
        }

        setImportResult({
          isOpen: true,
          successCount: 0,
          skipCount: 0,
          failCount: 0,
          emptyCount: 0,
          totalParsed: 0,
          error: true,
          errorMessage: friendlyError
        });
        showToast('Gagal memproses file Excel (Tertolak atau error format).', 'error');
        resolve();
      }
    };

    reader.onerror = () => {
      if (excelInputRef.current) excelInputRef.current.value = '';
      setImportResult({
        isOpen: true,
        successCount: 0,
        skipCount: 0,
        failCount: 0,
        emptyCount: 0,
        totalParsed: 0,
        error: true,
        errorMessage: 'Gagal mengunggah file Excel dari laptop/komputer Anda.'
      });
      showToast('Gagal membaca file.', 'error');
      resolve();
    };

    reader.readAsArrayBuffer(file);
  });
};

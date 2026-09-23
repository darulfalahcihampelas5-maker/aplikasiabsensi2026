import fs from 'fs';
let content = fs.readFileSync('src/App.tsx', 'utf-8');

if (!content.includes("import { id as idLocale } from 'date-fns/locale';")) {
  content = content.replace(
    "} from 'date-fns';",
    "} from 'date-fns';\nimport { id as idLocale } from 'date-fns/locale';"
  );
}

content = content.replace(
`  const handleShareWA = () => {
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
  };`,
`  const handleShareWA = () => {
    if (!existingSession) return;
    
    const absents = studentsInClass.filter(s => currentRecords[s.id] !== 'Hadir');
    
    let text = 'Bismillah\\n';
    text += 'Asalamualaikum wr wb\\n';
    text += 'Berikut laporan kehadiran siswa untuk hari ini : \\n';
    text += '*Laporan Presensi*\\n';
    text += 'Kelas: ' + selectedClass + '\\n';
    text += 'Tanggal: ' + format(new Date(date), 'EEEE, dd MMMM yyyy', { locale: idLocale }) + '\\n';
    text += 'Pertemuan ke: ' + meetingNumber + '\\n\\n';
    
    if (absents.length === 0) {
      text += '_Semua siswa hadir._\\n\\n';
    } else {
      text += '*Siswa yang tidak hadir:*\\n';
      absents.forEach((s, idx) => {
        text += (idx + 1) + '. ' + s.name + ' (' + currentRecords[s.id] + ')\\n';
      });
    }
    text += 'Terimakasih 🙏';
    
    const encodedText = encodeURIComponent(text);
    window.open('https://wa.me/?text=' + encodedText, '_blank');
  };`
);

fs.writeFileSync('src/App.tsx', content);

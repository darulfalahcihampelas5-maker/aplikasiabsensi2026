import fs from 'fs';
let content = fs.readFileSync('src/App.tsx', 'utf-8');
content = content.replace(/collection\(db, 'students'\)/g, "collection(db, 'users', currentUser!.uid, 'students')");
content = content.replace(/doc\(db, 'students', /g, "doc(db, 'users', currentUser!.uid, 'students', ");

content = content.replace(/collection\(db, 'attendanceRecords'\)/g, "collection(db, 'users', currentUser!.uid, 'attendanceRecords')");
content = content.replace(/doc\(db, 'attendanceRecords', /g, "doc(db, 'users', currentUser!.uid, 'attendanceRecords', ");

fs.writeFileSync('src/App.tsx', content);

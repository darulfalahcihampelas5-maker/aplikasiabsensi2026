import fs from 'fs';
let content = fs.readFileSync('src/App.tsx', 'utf-8');

content = content.replace(
  "const isUserAdmin = usernameClean === 'agan.parta' || usernameClean === 'agan.parta@gmail.com' || usernameClean === 'admin';",
  "const isUserAdmin = usernameClean === 'agan.parta' || usernameClean === 'agan.parta@gmail.com';"
);

fs.writeFileSync('src/App.tsx', content);

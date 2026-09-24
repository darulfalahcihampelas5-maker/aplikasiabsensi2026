/**
 * Helper utility for sorting school classes and students according to Indonesian grade levels:
 * X (10) -> XI (11) -> XII (12), followed by natural numeric suffix order (e.g. X1, X2, ... XI 8, XI 9, ... XII 1)
 */

export function parseClassName(className?: string): { level: number; rest: string; raw: string } {
  if (!className) return { level: 999, rest: '', raw: '' };
  const str = className.trim();
  const romanMap: Record<string, number> = {
    'VII': 7, 'VIII': 8, 'IX': 9,
    'X': 10, 'XI': 11, 'XII': 12,
    '7': 7, '8': 8, '9': 9,
    '10': 10, '11': 11, '12': 12
  };
  
  // Notice the order in the regex: XII before XI before X, VIII before VII
  const match = str.match(/^(XII|XI|X|VIII|VII|IX|12|11|10|[789])[\s\-_.]*(.*)$/i);
  if (match) {
    const levelKey = match[1].toUpperCase();
    const level = romanMap[levelKey] !== undefined ? romanMap[levelKey] : 500;
    const rest = match[2] || '';
    return { level, rest, raw: str };
  }
  return { level: 999, rest: str, raw: str };
}

export function compareClasses(a?: string, b?: string): number {
  if (!a && !b) return 0;
  if (!a) return 1;
  if (!b) return -1;
  const pA = parseClassName(a);
  const pB = parseClassName(b);
  if (pA.level !== pB.level) {
    return pA.level - pB.level;
  }
  return (pA.rest || '').localeCompare(pB.rest || '', 'id-ID', { numeric: true, sensitivity: 'base' });
}

export function compareStudentsByClass(
  a: { class?: string; name?: string }, 
  b: { class?: string; name?: string }
): number {
  const classComp = compareClasses(a.class, b.class);
  if (classComp !== 0) return classComp;
  return (a.name || '').localeCompare(b.name || '', 'id-ID', { numeric: true, sensitivity: 'base' });
}

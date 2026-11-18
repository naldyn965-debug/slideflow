export const fmtBytes = (n) => (n < 1024 ? `${n} B` : n < 1048576 ? `${Math.round(n / 1024)} KB` : `${(n / 1048576).toFixed(1)} MB`);

const rtf = new Intl.RelativeTimeFormat(undefined, { numeric: 'auto' });
export function ago(ts) {
  const s = (ts - Date.now()) / 1000;
  if (Math.abs(s) > 30 * 86400) return new Date(ts).toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
  for (const [unit, sec] of [['day', 86400], ['hour', 3600], ['minute', 60]]) {
    if (Math.abs(s) >= sec) return rtf.format(Math.round(s / sec), unit);
  }
  return 'just now';
}

const EMU_PER_INCH = 914400;
const inch = (emu) => String(Number((emu / EMU_PER_INCH).toFixed(2)));
export function sizeLabel(w, h) {
  if (!w || !h) return 'Unknown size';
  const r = w / h;
  const known = [[16 / 9, '16:9'], [4 / 3, '4:3'], [16 / 10, '16:10']].find(([x]) => Math.abs(x - r) < 0.01);
  const dims = `${inch(w)} × ${inch(h)} in`;
  return known ? `${known[1]}, ${dims}` : dims;
}

export const plural = (n, one, many = one + 's') => `${n} ${n === 1 ? one : many}`;
export const stripExt = (name) => name.replace(/\.[^.]+$/, '');

export const CLASS_LABEL = { title: 'Title', section: 'Section divider', text: 'Text', image: 'Image', comparison: 'Comparison', 'chart-table': 'Chart or table', closing: 'Closing', other: 'Other' };
export const MODE_LABEL = { preserve: 'Preserve Design', enhance: 'Smart Enhance' };
export const MODE_HELP = {
  preserve: 'Keep the template exactly as it is: its fonts, sizes, colours and margins.',
  enhance: 'Keep the template\'s identity, but lift sizes, margins and contrast to readable minimums.',
};

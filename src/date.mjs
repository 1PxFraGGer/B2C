export function dateLabel(value) {
  if (!value) return '—';
  // Assignment dates are calendar dates; don't shift them into the previous day.
  const input = typeof value === 'string' && /^\d{4}-\d{2}-\d{2}/.test(value)
    ? value.slice(0, 10) + 'T12:00:00'
    : value;
  const date = input instanceof Date ? input : new Date(input);
  return Number.isNaN(date.getTime()) ? '—' : date.toLocaleDateString('en-GB', {day:'numeric',month:'short',year:'numeric'});
}

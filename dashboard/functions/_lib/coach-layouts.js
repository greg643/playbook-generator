export const COACH_LAYOUTS = [1, 2, 4, 6, 9, 16];
export const COACH_LAYOUT_FIELDS = ['offense_plays_per_page', 'defense_plays_per_page'];

// Omission preserves the print geometry of older clients and queued jobs.
export function coachLayouts(options) {
  const result = {};
  for (const field of COACH_LAYOUT_FIELDS) {
    if (options[field] === undefined) continue;
    if (!COACH_LAYOUTS.includes(options[field])) throw new Error('Invalid coach card layout');
    result[field] = options[field];
  }
  return result;
}

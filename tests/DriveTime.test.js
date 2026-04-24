const {
  APP_CONFIG,
  convertDistance,
  roundNumber,
  stripHtml,
  normalizePoint,
  flattenSheetInput,
  classifyCommuteLevel,
  classifyTrafficLevel,
  detectRouteAccidents,
  getCommuteThresholds,
  getTrafficThresholds,
  STRAIGHT_LINE_DISTANCE,
} = require('./gasLoader');

// ── convertDistance ───────────────────────────────────────────────────────────

describe('convertDistance', () => {
  test('converts meters to miles', () => {
    expect(convertDistance(1609.34, 'mi')).toBe(1);
  });

  test('converts meters to kilometers', () => {
    expect(convertDistance(1000, 'km')).toBe(1);
  });

  test('converts meters to meters', () => {
    expect(convertDistance(500, 'm')).toBe(500);
  });

  test('converts meters to feet', () => {
    expect(convertDistance(1, 'ft')).toBeCloseTo(3.28, 1);
  });

  test('converts meters to nautical miles', () => {
    expect(convertDistance(1852, 'nm')).toBe(1);
  });

  test('accepts full unit name "miles"', () => {
    expect(convertDistance(1609.34, 'miles')).toBe(1);
  });

  test('accepts full unit name "kilometers"', () => {
    expect(convertDistance(1000, 'kilometers')).toBe(1);
  });

  test('accepts full unit name "meters"', () => {
    expect(convertDistance(500, 'meters')).toBe(500);
  });

  test('accepts full unit name "feet"', () => {
    expect(convertDistance(1, 'feet')).toBeCloseTo(3.28, 1);
  });

  test('accepts full unit name "nautical miles"', () => {
    expect(convertDistance(1852, 'nautical miles')).toBe(1);
  });

  test('is case-insensitive', () => {
    expect(convertDistance(1000, 'KM')).toBe(1);
    expect(convertDistance(1000, 'Kilometers')).toBe(1);
  });

  test('defaults to miles when unit is omitted', () => {
    expect(convertDistance(1609.34, null)).toBe(1);
    expect(convertDistance(1609.34, undefined)).toBe(1);
  });

  test('throws on unknown unit', () => {
    expect(() => convertDistance(1000, 'furlongs')).toThrow('Unknown unit');
  });
});

// ── roundNumber ───────────────────────────────────────────────────────────────

describe('roundNumber', () => {
  test('rounds to specified decimal places', () => {
    expect(roundNumber(1.256, 2)).toBe(1.26);
    expect(roundNumber(1.251, 2)).toBe(1.25);
  });

  test('rounds to 1 decimal place', () => {
    expect(roundNumber(12.34, 1)).toBe(12.3);
  });

  test('handles whole numbers', () => {
    expect(roundNumber(5, 2)).toBe(5);
  });

  test('handles string numbers', () => {
    expect(roundNumber('3.14159', 2)).toBe(3.14);
  });
});

// ── stripHtml ─────────────────────────────────────────────────────────────────

describe('stripHtml', () => {
  test('removes simple tags', () => {
    expect(stripHtml('<b>Turn left</b>')).toBe('Turn left');
  });

  test('removes nested tags', () => {
    expect(stripHtml('<div><span>Hello</span></div>')).toBe('Hello');
  });

  test('collapses extra whitespace', () => {
    expect(stripHtml('foo  <br/>  bar')).toBe('foo bar');
  });

  test('handles empty string', () => {
    expect(stripHtml('')).toBe('');
  });

  test('handles null/undefined', () => {
    expect(stripHtml(null)).toBe('');
    expect(stripHtml(undefined)).toBe('');
  });

  test('passes through plain text unchanged', () => {
    expect(stripHtml('No tags here')).toBe('No tags here');
  });
});

// ── flattenSheetInput ─────────────────────────────────────────────────────────

describe('flattenSheetInput', () => {
  test('returns scalar as single-element array', () => {
    expect(flattenSheetInput('Boston, MA')).toEqual(['Boston, MA']);
  });

  test('flattens a 1-D array', () => {
    expect(flattenSheetInput([1, 2, 3])).toEqual([1, 2, 3]);
  });

  test('flattens a 2-D array (Sheets range)', () => {
    expect(flattenSheetInput([[40.7128, -74.006]])).toEqual([40.7128, -74.006]);
  });

  test('flattens deeply nested arrays', () => {
    expect(flattenSheetInput([[[1]], [[2]]])).toEqual([1, 2]);
  });
});

// ── normalizePoint ────────────────────────────────────────────────────────────

describe('normalizePoint', () => {
  test('normalizes an address string', () => {
    const point = normalizePoint('Boston, MA');
    expect(point.kind).toBe('address');
    expect(point.address).toBe('Boston, MA');
    expect(point.cacheValue).toBe('Boston, MA');
  });

  test('trims whitespace from address', () => {
    const point = normalizePoint('  Boston, MA  ');
    expect(point.address).toBe('Boston, MA');
  });

  test('throws on empty address string', () => {
    expect(() => normalizePoint('')).toThrow('cannot be empty');
    expect(() => normalizePoint('   ')).toThrow('cannot be empty');
  });

  test('normalizes a coordinate pair array', () => {
    const point = normalizePoint([[40.7128, -74.006]]);
    expect(point.kind).toBe('coordinates');
    expect(point.latitude).toBe(40.7128);
    expect(point.longitude).toBe(-74.006);
    expect(point.cacheValue).toBe('40.7128,-74.006');
  });

  test('normalizes coordinate strings within an array', () => {
    const point = normalizePoint([['40.7128', '-74.006']]);
    expect(point.kind).toBe('coordinates');
    expect(point.latitude).toBe(40.7128);
  });

  test('throws on invalid input', () => {
    expect(() => normalizePoint(12345)).toThrow();
    expect(() => normalizePoint([[]])).toThrow();
  });
});

// ── getCommuteThresholds ──────────────────────────────────────────────────────

describe('getCommuteThresholds', () => {
  test('returns defaults when called with no arguments', () => {
    const t = getCommuteThresholds();
    expect(t.shortMax).toBe(APP_CONFIG.commuteMinutes.shortMax);
    expect(t.moderateMax).toBe(APP_CONFIG.commuteMinutes.moderateMax);
  });

  test('accepts custom thresholds', () => {
    const t = getCommuteThresholds(15, 30);
    expect(t.shortMax).toBe(15);
    expect(t.moderateMax).toBe(30);
  });

  test('treats empty string as default', () => {
    const t = getCommuteThresholds('', '');
    expect(t.shortMax).toBe(APP_CONFIG.commuteMinutes.shortMax);
  });

  test('throws when shortMax >= moderateMax', () => {
    expect(() => getCommuteThresholds(45, 20)).toThrow();
    expect(() => getCommuteThresholds(30, 30)).toThrow();
  });

  test('throws on non-numeric input', () => {
    expect(() => getCommuteThresholds('abc', 45)).toThrow();
  });
});

// ── getTrafficThresholds ──────────────────────────────────────────────────────

describe('getTrafficThresholds', () => {
  test('returns defaults when called with no arguments', () => {
    const t = getTrafficThresholds();
    expect(t.lowMax).toBe(APP_CONFIG.trafficDelayRatio.lowMax);
    expect(t.moderateMax).toBe(APP_CONFIG.trafficDelayRatio.moderateMax);
  });

  test('accepts custom thresholds', () => {
    const t = getTrafficThresholds(1.2, 1.5);
    expect(t.lowMax).toBe(1.2);
    expect(t.moderateMax).toBe(1.5);
  });

  test('throws when lowMax < 1', () => {
    expect(() => getTrafficThresholds(0.9, 1.5)).toThrow();
  });

  test('throws when lowMax >= moderateMax', () => {
    expect(() => getTrafficThresholds(1.5, 1.2)).toThrow();
  });
});

// ── classifyCommuteLevel ──────────────────────────────────────────────────────

describe('classifyCommuteLevel', () => {
  test('returns "short" at the boundary', () => {
    expect(classifyCommuteLevel(20)).toBe('short');
  });

  test('returns "short" below the threshold', () => {
    expect(classifyCommuteLevel(10)).toBe('short');
  });

  test('returns "moderate" above short threshold', () => {
    expect(classifyCommuteLevel(21)).toBe('moderate');
  });

  test('returns "moderate" at the moderate boundary', () => {
    expect(classifyCommuteLevel(45)).toBe('moderate');
  });

  test('returns "heavy" above moderate threshold', () => {
    expect(classifyCommuteLevel(46)).toBe('heavy');
  });

  test('respects custom thresholds', () => {
    expect(classifyCommuteLevel(10, 15, 30)).toBe('short');
    expect(classifyCommuteLevel(20, 15, 30)).toBe('moderate');
    expect(classifyCommuteLevel(35, 15, 30)).toBe('heavy');
  });
});

// ── classifyTrafficLevel ──────────────────────────────────────────────────────

describe('classifyTrafficLevel', () => {
  const snap = (trafficSecs, baseSecs) => ({
    durationInTrafficSeconds: trafficSecs,
    durationSeconds: baseSecs,
  });

  test('returns "unknown" when traffic data is absent', () => {
    expect(classifyTrafficLevel({ durationInTrafficSeconds: null, durationSeconds: 600 })).toBe('unknown');
    expect(classifyTrafficLevel({ durationInTrafficSeconds: 0, durationSeconds: 600 })).toBe('unknown');
  });

  test('returns "low" at ratio <= lowMax', () => {
    expect(classifyTrafficLevel(snap(660, 600))).toBe('low');   // ratio 1.1 exactly
    expect(classifyTrafficLevel(snap(600, 600))).toBe('low');   // ratio 1.0
  });

  test('returns "moderate" between thresholds', () => {
    expect(classifyTrafficLevel(snap(720, 600))).toBe('moderate');  // ratio 1.2
    expect(classifyTrafficLevel(snap(810, 600))).toBe('moderate');  // ratio 1.35 exactly
  });

  test('returns "heavy" above moderateMax', () => {
    expect(classifyTrafficLevel(snap(900, 600))).toBe('heavy');  // ratio 1.5
  });

  test('respects custom thresholds', () => {
    // ratio 1.2: above lowMax 1.1 but below moderateMax 1.3 → moderate
    expect(classifyTrafficLevel(snap(720, 600), 1.1, 1.3)).toBe('moderate');
    // ratio 1.2: above both thresholds → heavy
    expect(classifyTrafficLevel(snap(720, 600), 1.1, 1.15)).toBe('heavy');
  });
});

// ── detectRouteAccidents ──────────────────────────────────────────────────────

describe('detectRouteAccidents', () => {
  const snap = (warnings, instructions) => ({ warnings, instructions });

  test('returns false when no keywords present', () => {
    expect(detectRouteAccidents(snap(['Take I-95 N'], ['Turn left', 'Merge right']))).toBe(false);
  });

  test('detects "accident" in warnings', () => {
    expect(detectRouteAccidents(snap(['accident reported ahead'], []))).toBe(true);
  });

  test('detects "crash" in instructions', () => {
    expect(detectRouteAccidents(snap([], ['crash on ramp']))).toBe(true);
  });

  test('detects "road closed" keyword', () => {
    expect(detectRouteAccidents(snap(['road closed ahead'], []))).toBe(true);
  });

  test('detects "lane blocked" keyword', () => {
    expect(detectRouteAccidents(snap([], ['lane blocked due to debris']))).toBe(true);
  });

  test('is case-insensitive', () => {
    expect(detectRouteAccidents(snap(['ACCIDENT ON I-95'], []))).toBe(true);
  });

  test('returns false for empty snapshot', () => {
    expect(detectRouteAccidents(snap([], []))).toBe(false);
  });
});

// ── STRAIGHT_LINE_DISTANCE ────────────────────────────────────────────────────

describe('STRAIGHT_LINE_DISTANCE', () => {
  // Boston: 42.3601, -71.0589  Providence: 41.8240, -71.4128
  // Known straight-line distance ≈ 43 miles
  test('Boston to Providence is approximately 43 miles', () => {
    const dist = STRAIGHT_LINE_DISTANCE(
      [[42.3601, -71.0589]],
      [[41.8240, -71.4128]],
      'mi',
    );
    expect(dist).toBeGreaterThan(41);
    expect(dist).toBeLessThan(45);
  });

  // Same point → 0
  test('same point returns 0', () => {
    expect(STRAIGHT_LINE_DISTANCE([[40.7128, -74.006]], [[40.7128, -74.006]], 'mi')).toBe(0);
  });

  // New York to London ≈ 3459 miles
  test('New York to London is approximately 3459 miles', () => {
    const dist = STRAIGHT_LINE_DISTANCE(
      [[40.7128, -74.006]],
      [[51.5074, -0.1278]],
      'mi',
    );
    expect(dist).toBeGreaterThan(3400);
    expect(dist).toBeLessThan(3520);
  });

  test('accepts full unit name "kilometers"', () => {
    const km = STRAIGHT_LINE_DISTANCE([[42.3601, -71.0589]], [[41.8240, -71.4128]], 'kilometers');
    const mi = STRAIGHT_LINE_DISTANCE([[42.3601, -71.0589]], [[41.8240, -71.4128]], 'mi');
    expect(km).toBeCloseTo(mi * 1.60934, 0);
  });
});

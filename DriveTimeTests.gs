/* global DRIVING_DISTANCE, STRAIGHT_LINE_DISTANCE, COMMUTE_MINUTES, TRAVEL_TIME,
   COMMUTE_LEVEL, TRAFFIC_LEVEL, HAS_ROUTE_ACCIDENTS, PRIMARY_ROUTE, COMMUTE_SCORE,
   IS_WITHIN_COMMUTE, ROUND_TRIP_MINUTES, DISTANCE_DIFFERENCE, ROUTE_DETAILS */

/**
 * DriveTimeTests.gs
 * Manual integration test harness — run RunAllTests() from the Apps Script editor.
 * Results are written to the execution log (View → Logs).
 *
 * These tests hit the live Google Maps API, so they require Maps service to be
 * enabled and consume quota. Do not run in a tight loop.
 */

const TEST_ROUTES = Object.freeze({
  bostonToProvidence: Object.freeze({
    start: 'Boston, MA',
    end: 'Providence, RI',
    expectedMinDistanceMi: 40,
    expectedMaxDistanceMi: 60,
    expectedMinMinutes: 40,
    expectedMaxMinutes: 120,
    expectedStraightLineMi: 39,
    expectedStraightLineMaxMi: 45,
  }),
  samePoint: Object.freeze({
    start: 'Boston, MA',
    end: 'Boston, MA',
    expectedMinDistanceMi: 0,
    expectedMaxDistanceMi: 2,
  }),
});

function assert(condition, message) {
  if (!condition) {
    throw new Error('FAIL: ' + message);
  }
  Logger.log('PASS: ' + message);
}

function assertBetween(value, min, max, label) {
  assert(
    value >= min && value <= max,
    label + ' (' + value + ' expected between ' + min + ' and ' + max + ')',
  );
}

function assertOneOf(value, allowed, label) {
  assert(
    allowed.indexOf(value) !== -1,
    label + ' ("' + value + '" expected to be one of: ' + allowed.join(', ') + ')',
  );
}

function runSuite(suiteName, fn) {
  Logger.log('\n── ' + suiteName + ' ──');
  try {
    fn();
  } catch (e) {
    Logger.log(e.message);
  }
}

// ── Test suites ───────────────────────────────────────────────────────────────

function testDrivingDistance() {
  runSuite('DRIVING_DISTANCE', function() {
    const r = TEST_ROUTES.bostonToProvidence;

    const miles = DRIVING_DISTANCE(r.start, r.end, 'mi');
    assertBetween(miles, r.expectedMinDistanceMi, r.expectedMaxDistanceMi, 'driving distance in miles');

    const km = DRIVING_DISTANCE(r.start, r.end, 'km');
    assertBetween(km, r.expectedMinDistanceMi * 1.609, r.expectedMaxDistanceMi * 1.609, 'driving distance in km');

    const miByName = DRIVING_DISTANCE(r.start, r.end, 'miles');
    assert(miByName === miles, 'full unit name "miles" matches abbreviation "mi"');

    const kmByName = DRIVING_DISTANCE(r.start, r.end, 'kilometers');
    assert(kmByName === km, 'full unit name "kilometers" matches abbreviation "km"');

    assert(typeof miles === 'number' && miles > 0, 'result is a positive number');
  });
}

function testStraightLineDistance() {
  runSuite('STRAIGHT_LINE_DISTANCE', function() {
    const r = TEST_ROUTES.bostonToProvidence;

    const straight = STRAIGHT_LINE_DISTANCE(r.start, r.end, 'mi');
    assertBetween(straight, r.expectedStraightLineMi, r.expectedStraightLineMaxMi, 'straight-line distance in miles');

    const driving = DRIVING_DISTANCE(r.start, r.end, 'mi');
    assert(driving >= straight, 'driving distance >= straight-line distance');

    const zero = STRAIGHT_LINE_DISTANCE(
      [[42.3601, -71.0589]],
      [[42.3601, -71.0589]],
      'mi',
    );
    assert(zero === 0, 'same coordinates returns 0');
  });
}

function testCommuteMinutes() {
  runSuite('COMMUTE_MINUTES', function() {
    const r = TEST_ROUTES.bostonToProvidence;
    const minutes = COMMUTE_MINUTES(r.start, r.end);
    assertBetween(minutes, r.expectedMinMinutes, r.expectedMaxMinutes, 'commute minutes');
    assert(typeof minutes === 'number', 'result is a number');
  });
}

function testTravelTime() {
  runSuite('TRAVEL_TIME', function() {
    const r = TEST_ROUTES.bostonToProvidence;
    const text = TRAVEL_TIME(r.start, r.end, 'driving');
    assert(typeof text === 'string' && text.length > 0, 'returns non-empty string for driving');

    const walkText = TRAVEL_TIME(r.start, r.end, 'walking');
    assert(typeof walkText === 'string' && walkText.length > 0, 'returns non-empty string for walking');
  });
}

function testCommuteLevel() {
  runSuite('COMMUTE_LEVEL', function() {
    const r = TEST_ROUTES.bostonToProvidence;
    const level = COMMUTE_LEVEL(r.start, r.end);
    assertOneOf(level, ['short', 'moderate', 'heavy'], 'commute level');

    const withThresholds = COMMUTE_LEVEL(r.start, r.end, 5, 200);
    assert(withThresholds === 'moderate', 'wide thresholds (5–200 min) yields moderate for Boston→Providence');
  });
}

function testTrafficLevel() {
  runSuite('TRAFFIC_LEVEL', function() {
    const r = TEST_ROUTES.bostonToProvidence;
    const level = TRAFFIC_LEVEL(r.start, r.end);
    assertOneOf(level, ['low', 'moderate', 'heavy', 'unknown'], 'traffic level');
  });
}

function testHasRouteAccidents() {
  runSuite('HAS_ROUTE_ACCIDENTS', function() {
    const r = TEST_ROUTES.bostonToProvidence;
    const result = HAS_ROUTE_ACCIDENTS(r.start, r.end);
    assert(typeof result === 'boolean', 'returns a boolean');
  });
}

function testPrimaryRoute() {
  runSuite('PRIMARY_ROUTE', function() {
    const r = TEST_ROUTES.bostonToProvidence;
    const summary = PRIMARY_ROUTE(r.start, r.end);
    assert(typeof summary === 'string' && summary.length > 0, 'returns non-empty string');
  });
}

function testCommuteScore() {
  runSuite('COMMUTE_SCORE', function() {
    const r = TEST_ROUTES.bostonToProvidence;
    const score = COMMUTE_SCORE(r.start, r.end);
    assert(typeof score === 'number' && score > 0, 'returns a positive number');
    const minutes = COMMUTE_MINUTES(r.start, r.end);
    assert(score >= minutes, 'score is at least as large as raw minutes');
  });
}

function testIsWithinCommute() {
  runSuite('IS_WITHIN_COMMUTE', function() {
    const r = TEST_ROUTES.bostonToProvidence;
    assert(IS_WITHIN_COMMUTE(r.start, r.end, 300) === true, 'within 300 minutes');
    assert(IS_WITHIN_COMMUTE(r.start, r.end, 1) === false, 'not within 1 minute');
  });
}

function testRoundTripMinutes() {
  runSuite('ROUND_TRIP_MINUTES', function() {
    const r = TEST_ROUTES.bostonToProvidence;
    const roundTrip = ROUND_TRIP_MINUTES(r.start, r.end);
    const oneWay = COMMUTE_MINUTES(r.start, r.end);
    assert(roundTrip >= oneWay, 'round trip >= one way');
    assertBetween(roundTrip, r.expectedMinMinutes * 2 * 0.8, r.expectedMaxMinutes * 2, 'round trip minutes in range');
  });
}

function testDistanceDifference() {
  runSuite('DISTANCE_DIFFERENCE', function() {
    const r = TEST_ROUTES.bostonToProvidence;
    const diff = DISTANCE_DIFFERENCE(r.start, r.end, 'mi');
    assert(diff >= 0, 'driving is never shorter than straight-line');
    assert(typeof diff === 'number', 'returns a number');
  });
}

function testRouteDetails() {
  runSuite('ROUTE_DETAILS', function() {
    const r = TEST_ROUTES.bostonToProvidence;
    const grid = ROUTE_DETAILS(r.start, r.end);
    assert(Array.isArray(grid), 'returns an array');
    assert(grid.length === 8, 'returns 8 rows including header');
    assert(grid[0][0] === 'metric' && grid[0][1] === 'value', 'first row is header');

    const metrics = grid.slice(1).map(function(row) { return row[0]; });
    const expected = ['distance_miles', 'duration_minutes', 'duration_text', 'traffic_level', 'commute_level', 'has_route_accidents', 'primary_route'];
    expected.forEach(function(m) {
      assert(metrics.indexOf(m) !== -1, 'grid contains metric: ' + m);
    });
  });
}

function testCoordinateInputs() {
  runSuite('Coordinate inputs', function() {
    // Boston City Hall coordinates
    const bostonCoords = [[42.3601, -71.0589]];
    // Providence City Hall coordinates
    const providenceCoords = [[41.8240, -71.4128]];

    const miles = DRIVING_DISTANCE(bostonCoords, providenceCoords, 'mi');
    assertBetween(miles, 40, 60, 'driving distance via coordinates');

    const straight = STRAIGHT_LINE_DISTANCE(bostonCoords, providenceCoords, 'mi');
    assertBetween(straight, 39, 45, 'straight-line distance via coordinates');
  });
}

// ── Entry point ───────────────────────────────────────────────────────────────

function RunAllTests() {
  Logger.log('DriveTime integration tests — ' + new Date().toISOString());

  testDrivingDistance();
  testStraightLineDistance();
  testCommuteMinutes();
  testTravelTime();
  testCommuteLevel();
  testTrafficLevel();
  testHasRouteAccidents();
  testPrimaryRoute();
  testCommuteScore();
  testIsWithinCommute();
  testRoundTripMinutes();
  testDistanceDifference();
  testRouteDetails();
  testCoordinateInputs();

  Logger.log('\nDone.');
}

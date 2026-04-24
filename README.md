# Drive Time Apps Script

Copyright (c) Brian Mansell · [MIT License](LICENSE)

Google Apps Script helpers for Google Sheets that calculate route distance, commute time, traffic severity, and route incident signals using the built-in Google Maps service.

## Background

I built this while planning my family's cross-country move last summer. We were comparing neighborhoods, commute times to work and airport, and distances between cities — all inside a Google Sheet. The built-in Maps functions weren't enough, so I wrote these to get the data I actually needed.

## Available custom functions

### Distance and duration

| Function | Returns |
|----------|---------|
| `=COMMUTE_MINUTES(start, end)` | Driving minutes with live traffic |
| `=DRIVING_DISTANCE(start, end, unit)` | Driving distance in the chosen unit |
| `=STRAIGHT_LINE_DISTANCE(start, end, unit)` | Haversine straight-line distance |
| `=TRAVEL_TIME(start, end, mode)` | Human-readable travel time string |
| `=ROUND_TRIP_MINUTES(start, end)` | Total minutes for both directions |
| `=DISTANCE_DIFFERENCE(start, end, unit)` | How much longer the road route is vs. straight line |

### Commute and traffic classification

| Function | Returns |
|----------|---------|
| `=COMMUTE_LEVEL(start, end [, shortMax, moderateMax])` | `short`, `moderate`, or `heavy` |
| `=TRAFFIC_LEVEL(start, end [, lowMaxRatio, moderateMaxRatio])` | `low`, `moderate`, `heavy`, or `unknown` |
| `=COMMUTE_SCORE(start, end)` | Numeric score — lower is better |
| `=IS_WITHIN_COMMUTE(start, end, maxMinutes)` | `TRUE` if the drive is within the time limit |

### Route insights

| Function | Returns |
|----------|---------|
| `=HAS_ROUTE_ACCIDENTS(start, end)` | `TRUE` when route text mentions likely incidents |
| `=PRIMARY_ROUTE(start, end)` | Main route summary (major roads) from Google Maps |
| `=ROUTE_WARNINGS(start, end)` | Route warnings joined as a single string |
| `=ROUTE_DETAILS(start, end)` | 2-column grid: distance, time, traffic, commute level, incidents, route name |

### Batch functions

| Function | Returns |
|----------|---------|
| `=BATCH_COMMUTE_LEVEL(origins, destination [, shortMax, moderateMax])` | Commute level for each origin in a range |

## Inputs

Both `start` and `end` accept:

- Address strings: `"New York, NY"`
- Coordinate pairs from two adjacent cells: `{40.7128, -74.0060}`

**Distance units** (`unit`): abbreviation or full name — `mi` / `miles` (default), `km` / `kilometers`, `m` / `meters`, `ft` / `feet`, `nm` / `nautical miles`

**Travel modes** (`mode`): `driving` (default), `walking`, `bicycling`, `transit`

## Examples

```
=COMMUTE_MINUTES("Boston, MA", "Providence, RI")
=DRIVING_DISTANCE("Boston, MA", "Providence, RI", "mi")
=DRIVING_DISTANCE("Boston, MA", "Providence, RI", "miles")
=STRAIGHT_LINE_DISTANCE("Boston, MA", "Providence, RI", "km")
=TRAVEL_TIME("Boston, MA", "Providence, RI", "walking")
=ROUND_TRIP_MINUTES("Boston, MA", "Providence, RI")
=DISTANCE_DIFFERENCE("Boston, MA", "Providence, RI", "mi")

=COMMUTE_LEVEL("Boston, MA", "Providence, RI")
=COMMUTE_LEVEL("Boston, MA", "Providence, RI", 15, 40)
=TRAFFIC_LEVEL("Boston, MA", "Providence, RI")
=TRAFFIC_LEVEL("Boston, MA", "Providence, RI", 1.05, 1.25)
=COMMUTE_SCORE("Boston, MA", "Providence, RI")
=IS_WITHIN_COMMUTE("Boston, MA", "Providence, RI", 45)

=HAS_ROUTE_ACCIDENTS("Boston, MA", "Providence, RI")
=PRIMARY_ROUTE("Boston, MA", "Providence, RI")
=ROUTE_WARNINGS("Boston, MA", "Providence, RI")
=ROUTE_DETAILS("Boston, MA", "Providence, RI")

=BATCH_COMMUTE_LEVEL(A2:A10, $B$1)
=BATCH_COMMUTE_LEVEL(A2:A10, $B$1, 15, 40)
```

## Classification thresholds

### COMMUTE_LEVEL

| Level | Default condition |
|-------|-------------------|
| `short` | ≤ 20 minutes |
| `moderate` | 21 – 45 minutes |
| `heavy` | > 45 minutes |

Override with optional parameters: `=COMMUTE_LEVEL(start, end, shortMax, moderateMax)`

### TRAFFIC_LEVEL

Compares the traffic-adjusted duration to the baseline duration as a ratio.

| Level | Default condition |
|-------|-------------------|
| `low` | ratio ≤ 1.10 |
| `moderate` | ratio 1.11 – 1.35 |
| `heavy` | ratio > 1.35 |
| `unknown` | Google did not return real-time traffic data |

Override with optional parameters: `=TRAFFIC_LEVEL(start, end, lowMaxRatio, moderateMaxRatio)`

### COMMUTE_SCORE

Starts from driving minutes, then adds penalties:

| Condition | Penalty |
|-----------|---------|
| Moderate traffic | +12 |
| Heavy traffic | +25 |
| Detected incidents | +20 |
| Route > 30 miles | +10 |

Lower scores are better.

## Caching

Route data is cached using `CacheService.getDocumentCache()` for 1 hour. The cache key is an MD5 hash of the normalized origin, destination, travel mode, and whether a departure time was set.

- Repeated calls with the same inputs in the same document return instantly without hitting the Maps API.
- Live-traffic functions (`COMMUTE_MINUTES`, `COMMUTE_LEVEL`, `TRAFFIC_LEVEL`, `HAS_ROUTE_ACCIDENTS`, etc.) share a `depart-now` cache entry, so they only call the API once per route pair per session.
- Static functions like `DRIVING_DISTANCE` and `PRIMARY_ROUTE` use a separate `no-departure` entry so traffic data is never mixed in.
- The document-level cache is not shared across spreadsheets.
- `CacheService` has a hard 6-hour TTL maximum; the current 1-hour setting stays well within that.

## Notes and limitations

- `HAS_ROUTE_ACCIDENTS` is a keyword-based heuristic on route warnings and step instructions. Apps Script does not expose a dedicated incidents feed, so treat it as a signal rather than a guarantee.
- `TRAFFIC_LEVEL` returns `unknown` when Google does not provide a real-time traffic duration, which is common off-peak or for non-driving modes.
- `STRAIGHT_LINE_DISTANCE` uses the Haversine formula and does not call the Maps API — it requires coordinate inputs or geocodes the addresses itself.

## Testing

**Unit tests** (pure math and classification logic — no Maps API required):

```sh
npm test
```

**Integration tests** (live Maps API):

Open `DriveTimeTests.gs` in the Apps Script editor and run `RunAllTests()`. Results appear in View → Logs.

## Project files

| File | Purpose |
|------|---------|
| `DriveTime.gs` | All custom functions |
| `DriveTimeTests.gs` | Manual integration test harness (run in Apps Script editor) |
| `tests/DriveTime.test.js` | Jest unit tests |
| `tests/gasLoader.js` | Loads DriveTime.gs into Node with GAS globals stubbed |
| `eslint.config.mjs` | ESLint config with GAS globals |
| `.github/workflows/ci.yml` | GitHub Actions — lint and unit tests on every PR |
| `LICENSE` | MIT |

/**
 * DriveTime.gs
 * Google Apps Script commute, distance, and route helpers for Google Sheets.
 * Copyright (c) Brian Mansell — MIT License
 *
 * Custom functions:
 *   =COMMUTE_MINUTES(start, end)
 *   =DRIVING_DISTANCE(start, end, unit)
 *   =STRAIGHT_LINE_DISTANCE(start, end, unit)
 *   =TRAVEL_TIME(start, end, mode)
 *   =ROUND_TRIP_MINUTES(start, end)
 *   =DISTANCE_DIFFERENCE(start, end, unit)
 *   =COMMUTE_LEVEL(start, end [, shortMax, moderateMax])
 *   =TRAFFIC_LEVEL(start, end [, lowMaxRatio, moderateMaxRatio])
 *   =COMMUTE_SCORE(start, end)
 *   =IS_WITHIN_COMMUTE(start, end, maxMinutes)
 *   =HAS_ROUTE_ACCIDENTS(start, end)
 *   =PRIMARY_ROUTE(start, end)
 *   =ROUTE_WARNINGS(start, end)
 *   =ROUTE_DETAILS(start, end)
 *   =BATCH_COMMUTE_LEVEL(origins, destination [, shortMax, moderateMax])
 *
 * Inputs accept address strings or coordinate pairs:
 *   "New York, NY"
 *   {40.7128, -74.0060}
 *
 * Distance units accept abbreviations or full names:
 *   mi / miles (default), km / kilometers, m / meters, ft / feet, nm / nautical miles
 */

const APP_CONFIG = Object.freeze({
  cacheTtlSeconds: 60 * 60,
  earthRadiusMeters: 6371230,
  distanceUnits: Object.freeze({
    km: { meters: 1000, label: 'kilometers' },
    m: { meters: 1, label: 'meters' },
    mi: { meters: 1609.34, label: 'miles' },
    ft: { meters: 1 / 3.28084, label: 'feet' },
    nm: { meters: 1852, label: 'nautical miles' },
  }),
  commuteMinutes: Object.freeze({
    shortMax: 20,
    moderateMax: 45,
  }),
  trafficDelayRatio: Object.freeze({
    lowMax: 1.1,
    moderateMax: 1.35,
  }),
  commuteScore: Object.freeze({
    heavyTrafficPenalty: 25,
    moderateTrafficPenalty: 12,
    accidentPenalty: 20,
    longRoutePenalty: 10,
    longRouteThresholdMiles: 30,
  }),
  incidentKeywords: Object.freeze([
    'accident',
    'crash',
    'collision',
    'incident',
    'disabled vehicle',
    'lane blocked',
    'lanes blocked',
    'lane closed',
    'lanes closed',
    'road closed',
    'closure',
  ]),
});

// ── Public custom functions ───────────────────────────────────────────────────

/**
 * Returns travel time as numeric minutes.
 *
 * =COMMUTE_MINUTES("New York, NY", "Hoboken, NJ")
 *
 * @param {*} startPointInput
 * @param {*} endPointInput
 * @returns {number}
 * @customFunction
 */
function COMMUTE_MINUTES(startPointInput, endPointInput) {
  return getDrivingMinutes(startPointInput, endPointInput);
}

/**
 * Returns driving distance between two points.
 *
 * =DRIVING_DISTANCE("New York, NY", "Hoboken, NJ", "mi")
 *
 * @param {*} startPointInput
 * @param {*} endPointInput
 * @param {string} [distanceUnit]
 * @returns {number}
 * @customFunction
 */
function DRIVING_DISTANCE(startPointInput, endPointInput, distanceUnit) {
  const routeSnapshot = getRouteSnapshot(startPointInput, endPointInput, {
    travelMode: 'driving',
    departNow: false,
  });

  return convertDistance(routeSnapshot.distanceMeters, distanceUnit || 'mi');
}

/**
 * Returns travel time text for the selected mode.
 *
 * =TRAVEL_TIME("New York, NY", "Hoboken, NJ", "walking")
 *
 * @param {*} startPointInput
 * @param {*} endPointInput
 * @param {string} [travelMode]
 * @returns {string}
 * @customFunction
 */
function TRAVEL_TIME(startPointInput, endPointInput, travelMode) {
  const normalizedMode = String(travelMode || 'driving').toLowerCase();
  const routeSnapshot = getRouteSnapshot(startPointInput, endPointInput, {
    travelMode: normalizedMode,
    departNow: normalizedMode === 'driving',
  });

  return routeSnapshot.durationInTrafficText || routeSnapshot.durationText;
}

/**
 * Classifies a driving commute as short, moderate, or heavy.
 *
 * =COMMUTE_LEVEL("New York, NY", "Hoboken, NJ")
 * =COMMUTE_LEVEL("New York, NY", "Hoboken, NJ", 15, 40)
 *
 * @param {*} startPointInput
 * @param {*} endPointInput
 * @param {number} [shortMaxMinutes] Minutes at or below which the commute is "short". Default: 20.
 * @param {number} [moderateMaxMinutes] Minutes at or below which the commute is "moderate". Default: 45.
 * @returns {string} "short", "moderate", or "heavy"
 * @customFunction
 */
function COMMUTE_LEVEL(startPointInput, endPointInput, shortMaxMinutes, moderateMaxMinutes) {
  const drivingMinutes = getDrivingMinutes(startPointInput, endPointInput);
  return classifyCommuteLevel(drivingMinutes, shortMaxMinutes, moderateMaxMinutes);
}

/**
 * Classifies current traffic as low, moderate, heavy, or unknown.
 * The ratio of traffic-adjusted duration to baseline duration determines the level.
 * Returns "unknown" when Google does not provide real-time traffic data.
 *
 * =TRAFFIC_LEVEL("New York, NY", "Hoboken, NJ")
 * =TRAFFIC_LEVEL("New York, NY", "Hoboken, NJ", 1.05, 1.25)
 *
 * @param {*} startPointInput
 * @param {*} endPointInput
 * @param {number} [lowMaxRatio] Delay ratio at or below which traffic is "low". Default: 1.10.
 * @param {number} [moderateMaxRatio] Delay ratio at or below which traffic is "moderate". Default: 1.35.
 * @returns {string} "low", "moderate", "heavy", or "unknown"
 * @customFunction
 */
function TRAFFIC_LEVEL(startPointInput, endPointInput, lowMaxRatio, moderateMaxRatio) {
  const routeSnapshot = getRouteSnapshot(startPointInput, endPointInput, {
    travelMode: 'driving',
    departNow: true,
  });
  return classifyTrafficLevel(routeSnapshot, lowMaxRatio, moderateMaxRatio);
}

/**
 * Returns TRUE when route warnings or turn instructions mention likely incidents.
 * This is a keyword-based heuristic because Apps Script does not expose a dedicated accidents feed.
 *
 * =HAS_ROUTE_ACCIDENTS("New York, NY", "Hoboken, NJ")
 *
 * @param {*} startPointInput
 * @param {*} endPointInput
 * @returns {boolean}
 * @customFunction
 */
function HAS_ROUTE_ACCIDENTS(startPointInput, endPointInput) {
  const routeSnapshot = getRouteSnapshot(startPointInput, endPointInput, {
    travelMode: 'driving',
    departNow: true,
  });
  return detectRouteAccidents(routeSnapshot);
}

/**
 * Returns the primary route summary, such as major roads used.
 *
 * =PRIMARY_ROUTE("New York, NY", "Hoboken, NJ")
 *
 * @param {*} startPointInput
 * @param {*} endPointInput
 * @returns {string}
 * @customFunction
 */
function PRIMARY_ROUTE(startPointInput, endPointInput) {
  const routeSnapshot = getRouteSnapshot(startPointInput, endPointInput, {
    travelMode: 'driving',
    departNow: false,
  });

  return routeSnapshot.summary || 'Primary route summary unavailable';
}

/**
 * Returns the straight-line distance between two points.
 *
 * =STRAIGHT_LINE_DISTANCE("New York, NY", "Hoboken, NJ", "mi")
 *
 * @param {*} startPointInput
 * @param {*} endPointInput
 * @param {string} [distanceUnit]
 * @returns {number}
 * @customFunction
 */
function STRAIGHT_LINE_DISTANCE(startPointInput, endPointInput, distanceUnit) {
  const startCoordinates = resolveCoordinates(normalizePoint(startPointInput));
  const endCoordinates = resolveCoordinates(normalizePoint(endPointInput));

  const toRadians = function(degrees) {
    return degrees * Math.PI / 180;
  };

  const latitudeDelta = toRadians(endCoordinates.latitude - startCoordinates.latitude);
  const longitudeDelta = toRadians(endCoordinates.longitude - startCoordinates.longitude);
  const startLatitude = toRadians(startCoordinates.latitude);
  const endLatitude = toRadians(endCoordinates.latitude);

  const haversineFactor =
    Math.sin(latitudeDelta / 2) * Math.sin(latitudeDelta / 2) +
    Math.sin(longitudeDelta / 2) * Math.sin(longitudeDelta / 2) *
      Math.cos(startLatitude) * Math.cos(endLatitude);

  const centralAngle =
    2 * Math.atan2(Math.sqrt(haversineFactor), Math.sqrt(1 - haversineFactor));

  return convertDistance(
    APP_CONFIG.earthRadiusMeters * centralAngle,
    distanceUnit || 'mi',
  );
}

/**
 * Returns a commute score where lower is better.
 *
 * =COMMUTE_SCORE("Boston, MA", "Providence, RI")
 *
 * @param {*} startPointInput
 * @param {*} endPointInput
 * @returns {number}
 * @customFunction
 */
function COMMUTE_SCORE(startPointInput, endPointInput) {
  const routeSnapshot = getRouteSnapshot(startPointInput, endPointInput, {
    travelMode: 'driving',
    departNow: true,
  });
  const seconds = routeSnapshot.durationInTrafficSeconds || routeSnapshot.durationSeconds;
  const minutes = roundNumber(seconds / 60, 1);
  const trafficLevel = classifyTrafficLevel(routeSnapshot);
  const hasAccidents = detectRouteAccidents(routeSnapshot);
  const drivingMiles = convertDistance(routeSnapshot.distanceMeters, 'mi');
  let score = minutes;

  if (trafficLevel === 'moderate') {
    score += APP_CONFIG.commuteScore.moderateTrafficPenalty;
  } else if (trafficLevel === 'heavy') {
    score += APP_CONFIG.commuteScore.heavyTrafficPenalty;
  }

  if (hasAccidents) {
    score += APP_CONFIG.commuteScore.accidentPenalty;
  }

  if (drivingMiles > APP_CONFIG.commuteScore.longRouteThresholdMiles) {
    score += APP_CONFIG.commuteScore.longRoutePenalty;
  }

  return roundNumber(score, 1);
}

/**
 * Returns TRUE when a drive is within the maximum target commute time.
 *
 * =IS_WITHIN_COMMUTE("Boston, MA", "Providence, RI", 45)
 *
 * @param {*} startPointInput
 * @param {*} endPointInput
 * @param {number} maxMinutes
 * @returns {boolean}
 * @customFunction
 */
function IS_WITHIN_COMMUTE(startPointInput, endPointInput, maxMinutes) {
  const targetMinutes = Number(maxMinutes);
  if (Number.isNaN(targetMinutes) || targetMinutes < 0) {
    throw new Error('maxMinutes must be a non-negative number.');
  }

  return getDrivingMinutes(startPointInput, endPointInput) <= targetMinutes;
}

/**
 * Returns total driving minutes for a round trip.
 *
 * =ROUND_TRIP_MINUTES("Boston, MA", "Providence, RI")
 *
 * @param {*} pointAInput
 * @param {*} pointBInput
 * @returns {number}
 * @customFunction
 */
function ROUND_TRIP_MINUTES(pointAInput, pointBInput) {
  const outboundMinutes = getDrivingMinutes(pointAInput, pointBInput);
  const returnMinutes = getDrivingMinutes(pointBInput, pointAInput);
  return roundNumber(outboundMinutes + returnMinutes, 1);
}

/**
 * Returns how much longer the driving route is than the straight-line distance.
 *
 * =DISTANCE_DIFFERENCE("Boston, MA", "Providence, RI", "mi")
 *
 * @param {*} startPointInput
 * @param {*} endPointInput
 * @param {string} [distanceUnit]
 * @returns {number}
 * @customFunction
 */
function DISTANCE_DIFFERENCE(startPointInput, endPointInput, distanceUnit) {
  const normalizedUnit = distanceUnit || 'mi';
  const drivingDistance = DRIVING_DISTANCE(startPointInput, endPointInput, normalizedUnit);
  const straightLineDistance = STRAIGHT_LINE_DISTANCE(startPointInput, endPointInput, normalizedUnit);
  return roundNumber(drivingDistance - straightLineDistance, 2);
}

/**
 * Returns route warnings as a single text string.
 *
 * =ROUTE_WARNINGS("Boston, MA", "Providence, RI")
 *
 * @param {*} startPointInput
 * @param {*} endPointInput
 * @returns {string}
 * @customFunction
 */
function ROUTE_WARNINGS(startPointInput, endPointInput) {
  const routeSnapshot = getRouteSnapshot(startPointInput, endPointInput, {
    travelMode: 'driving',
    departNow: true,
  });

  return routeSnapshot.warnings.length
    ? routeSnapshot.warnings.join(' | ')
    : 'No route warnings reported';
}

/**
 * Returns a 2-column grid of route details for dashboards and comparisons.
 *
 * =ROUTE_DETAILS("Boston, MA", "Providence, RI")
 *
 * @param {*} startPointInput
 * @param {*} endPointInput
 * @returns {Array<Array<*>>}
 * @customFunction
 */
function ROUTE_DETAILS(startPointInput, endPointInput) {
  return buildRouteDetailsGrid(startPointInput, endPointInput);
}

/**
 * Returns commute levels for a range of origins paired with one destination.
 *
 * =BATCH_COMMUTE_LEVEL(A2:A10, $B$1)
 * =BATCH_COMMUTE_LEVEL(A2:A10, $B$1, 15, 40)
 *
 * @param {*} originsInput A range of origin addresses or coordinate pairs.
 * @param {*} destinationInput A single destination address or coordinate pair.
 * @param {number} [shortMaxMinutes] Minutes at or below which the commute is "short". Default: 20.
 * @param {number} [moderateMaxMinutes] Minutes at or below which the commute is "moderate". Default: 45.
 * @returns {Array<Array<string>>} Grid of "short", "moderate", or "heavy" values.
 * @customFunction
 */
function BATCH_COMMUTE_LEVEL(originsInput, destinationInput, shortMaxMinutes, moderateMaxMinutes) {
  const rows = Array.isArray(originsInput) ? originsInput : [[originsInput]];

  return rows.map(function(row) {
    return row.map(function(originCell) {
      if (originCell === '' || originCell === null || originCell === undefined) {
        return '';
      }

      return COMMUTE_LEVEL(originCell, destinationInput, shortMaxMinutes, moderateMaxMinutes);
    });
  });
}

// ── Route engine ──────────────────────────────────────────────────────────────

/**
 * Builds a route snapshot containing only the fields this script uses.
 * @param {*} startPointInput
 * @param {*} endPointInput
 * @param {{travelMode: string, departNow: boolean}} options
 * @returns {{summary: string, warnings: string[], distanceMeters: number, distanceText: string, durationSeconds: number, durationText: string, durationInTrafficSeconds: number|null, durationInTrafficText: string|null, instructions: string[]}}
 */
function getRouteSnapshot(startPointInput, endPointInput, options) {
  const normalizedOptions = options || {};
  const travelMode = String(normalizedOptions.travelMode || 'driving').toLowerCase();
  const departNow = Boolean(normalizedOptions.departNow);

  const startPoint = normalizePoint(startPointInput);
  const endPoint = normalizePoint(endPointInput);
  const cacheKey = [
    'route',
    startPoint.cacheValue,
    endPoint.cacheValue,
    travelMode,
    departNow ? 'depart-now' : 'no-departure',
  ].join('|');

  const cachedValue = getCachedValue(cacheKey);
  if (cachedValue !== null) {
    return JSON.parse(cachedValue);
  }

  const directionFinder = Maps.newDirectionFinder().setMode(travelMode);
  applyPointToDirections(directionFinder, 'setOrigin', startPoint);
  applyPointToDirections(directionFinder, 'setDestination', endPoint);

  if (departNow && travelMode === 'driving') {
    directionFinder.setDepart(new Date());
  }

  const directions = directionFinder.getDirections();
  const primaryRoute = directions.routes && directions.routes[0];
  const primaryLeg = primaryRoute && primaryRoute.legs && primaryRoute.legs[0];

  if (!primaryRoute || !primaryLeg) {
    throw new Error(
      'No ' + travelMode + ' route found from "' + startPoint.cacheValue + '" to "' + endPoint.cacheValue + '".',
    );
  }

  const routeSnapshot = {
    summary: primaryRoute.summary || '',
    warnings: (primaryRoute.warnings || []).map(stripHtml),
    distanceMeters: primaryLeg.distance.value,
    distanceText: primaryLeg.distance.text,
    durationSeconds: primaryLeg.duration.value,
    durationText: primaryLeg.duration.text,
    durationInTrafficSeconds: primaryLeg.duration_in_traffic
      ? primaryLeg.duration_in_traffic.value
      : null,
    durationInTrafficText: primaryLeg.duration_in_traffic
      ? primaryLeg.duration_in_traffic.text
      : null,
    instructions: (primaryLeg.steps || []).map(function(step) {
      return stripHtml(step.html_instructions);
    }),
  };

  setCachedValue(cacheKey, JSON.stringify(routeSnapshot));
  return routeSnapshot;
}

/**
 * Returns the best available driving duration in minutes.
 * Prefers current traffic duration when available.
 * @param {*} startPointInput
 * @param {*} endPointInput
 * @returns {number}
 */
function getDrivingMinutes(startPointInput, endPointInput) {
  const routeSnapshot = getRouteSnapshot(startPointInput, endPointInput, {
    travelMode: 'driving',
    departNow: true,
  });

  const seconds = routeSnapshot.durationInTrafficSeconds || routeSnapshot.durationSeconds;
  return roundNumber(seconds / 60, 1);
}

/**
 * Builds a row of spreadsheet-friendly route details.
 * @param {*} startPointInput
 * @param {*} endPointInput
 * @returns {Array<Array<*>>}
 */
function buildRouteDetailsGrid(startPointInput, endPointInput) {
  const drivingRoute = getRouteSnapshot(startPointInput, endPointInput, {
    travelMode: 'driving',
    departNow: true,
  });
  const drivingMiles = convertDistance(drivingRoute.distanceMeters, 'mi');
  const drivingMinutes = getDrivingMinutes(startPointInput, endPointInput);
  const commuteLevel = classifyCommuteLevel(drivingMinutes);
  const trafficLevel = classifyTrafficLevel(drivingRoute);
  const hasAccidents = detectRouteAccidents(drivingRoute);

  return [
    ['metric', 'value'],
    ['distance_miles', drivingMiles],
    ['duration_minutes', drivingMinutes],
    ['duration_text', drivingRoute.durationInTrafficText || drivingRoute.durationText],
    ['traffic_level', trafficLevel],
    ['commute_level', commuteLevel],
    ['has_route_accidents', hasAccidents],
    ['primary_route', drivingRoute.summary || 'Primary route summary unavailable'],
  ];
}

// ── Classifiers ───────────────────────────────────────────────────────────────

/**
 * Classifies commute severity using the provided thresholds.
 * @param {number} drivingMinutes
 * @param {number=} shortMaxMinutes
 * @param {number=} moderateMaxMinutes
 * @returns {string}
 */
function classifyCommuteLevel(drivingMinutes, shortMaxMinutes, moderateMaxMinutes) {
  const thresholds = getCommuteThresholds(shortMaxMinutes, moderateMaxMinutes);

  if (drivingMinutes <= thresholds.shortMax) {
    return 'short';
  }

  if (drivingMinutes <= thresholds.moderateMax) {
    return 'moderate';
  }

  return 'heavy';
}

/**
 * Classifies traffic severity from the route delay ratio.
 * @param {{durationInTrafficSeconds: number|null, durationSeconds: number}} routeSnapshot
 * @param {number=} lowMaxRatio
 * @param {number=} moderateMaxRatio
 * @returns {string}
 */
function classifyTrafficLevel(routeSnapshot, lowMaxRatio, moderateMaxRatio) {
  if (!routeSnapshot.durationInTrafficSeconds || !routeSnapshot.durationSeconds) {
    return 'unknown';
  }

  const thresholds = getTrafficThresholds(lowMaxRatio, moderateMaxRatio);
  const delayRatio = routeSnapshot.durationInTrafficSeconds / routeSnapshot.durationSeconds;

  if (delayRatio <= thresholds.lowMax) {
    return 'low';
  }

  if (delayRatio <= thresholds.moderateMax) {
    return 'moderate';
  }

  return 'heavy';
}

/**
 * Returns whether route text suggests incidents along the way.
 * @param {{warnings: string[], instructions: string[]}} routeSnapshot
 * @returns {boolean}
 */
function detectRouteAccidents(routeSnapshot) {
  const routeText = routeSnapshot.warnings
    .concat(routeSnapshot.instructions)
    .join(' ')
    .toLowerCase();

  return APP_CONFIG.incidentKeywords.some(function(keyword) {
    return routeText.indexOf(keyword) !== -1;
  });
}

// ── Threshold helpers ─────────────────────────────────────────────────────────

/**
 * Returns the configured or user-supplied commute thresholds.
 * @param {number=} shortMaxMinutes
 * @param {number=} moderateMaxMinutes
 * @returns {{shortMax: number, moderateMax: number}}
 */
function getCommuteThresholds(shortMaxMinutes, moderateMaxMinutes) {
  const shortMax = shortMaxMinutes === undefined || shortMaxMinutes === null || shortMaxMinutes === ''
    ? APP_CONFIG.commuteMinutes.shortMax
    : Number(shortMaxMinutes);
  const moderateMax = moderateMaxMinutes === undefined || moderateMaxMinutes === null || moderateMaxMinutes === ''
    ? APP_CONFIG.commuteMinutes.moderateMax
    : Number(moderateMaxMinutes);

  if (Number.isNaN(shortMax) || Number.isNaN(moderateMax)) {
    throw new Error('Commute thresholds must be numeric minute values.');
  }

  if (shortMax < 0 || moderateMax < 0 || shortMax >= moderateMax) {
    throw new Error('Commute thresholds must satisfy 0 <= shortMax < moderateMax.');
  }

  return {
    shortMax: shortMax,
    moderateMax: moderateMax,
  };
}

/**
 * Returns the configured or user-supplied traffic thresholds.
 * @param {number=} lowMaxRatio
 * @param {number=} moderateMaxRatio
 * @returns {{lowMax: number, moderateMax: number}}
 */
function getTrafficThresholds(lowMaxRatio, moderateMaxRatio) {
  const lowMax = lowMaxRatio === undefined || lowMaxRatio === null || lowMaxRatio === ''
    ? APP_CONFIG.trafficDelayRatio.lowMax
    : Number(lowMaxRatio);
  const moderateMax = moderateMaxRatio === undefined || moderateMaxRatio === null || moderateMaxRatio === ''
    ? APP_CONFIG.trafficDelayRatio.moderateMax
    : Number(moderateMaxRatio);

  if (Number.isNaN(lowMax) || Number.isNaN(moderateMax)) {
    throw new Error('Traffic thresholds must be numeric ratio values.');
  }

  if (lowMax < 1 || moderateMax < 1 || lowMax >= moderateMax) {
    throw new Error('Traffic thresholds must satisfy 1 <= lowMax < moderateMax.');
  }

  return {
    lowMax: lowMax,
    moderateMax: moderateMax,
  };
}

// ── Low-level utilities ───────────────────────────────────────────────────────

/**
 * Creates a normalized cache key so equivalent inputs reuse the same entry.
 * @param {string} rawKey
 * @returns {string}
 */
function createCacheKey(rawKey) {
  const normalized = String(rawKey).toLowerCase().replace(/\s+/g, '');
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, normalized)
    .map(function(byteValue) {
      return (byteValue + 256).toString(16).slice(-2);
    })
    .join('');
}

/**
 * Reads a cached value for the provided raw key.
 * @param {string} rawKey
 * @returns {string|null}
 */
function getCachedValue(rawKey) {
  return CacheService.getDocumentCache().get(createCacheKey(rawKey));
}

/**
 * Stores a cached value for the provided raw key.
 * @param {string} rawKey
 * @param {string} value
 * @returns {void}
 */
function setCachedValue(rawKey, value) {
  CacheService.getDocumentCache().put(
    createCacheKey(rawKey),
    value,
    APP_CONFIG.cacheTtlSeconds,
  );
}

/**
 * Converts a nested Sheets range or array input into a flat list of values.
 * @param {*} value
 * @returns {Array<*>}
 */
function flattenSheetInput(value) {
  if (!Array.isArray(value)) {
    return [value];
  }

  return value.reduce(function(flatValues, item) {
    return flatValues.concat(flattenSheetInput(item));
  }, []);
}

/**
 * Normalizes an address or coordinate pair into a consistent point object.
 * @param {*} pointInput
 * @returns {{kind: string, cacheValue: string, address: string|null, latitude: number|null, longitude: number|null}}
 */
function normalizePoint(pointInput) {
  if (typeof pointInput === 'string') {
    const address = pointInput.trim();
    if (!address) {
      throw new Error('Point input cannot be empty.');
    }
    return {
      kind: 'address',
      cacheValue: address,
      address: address,
      latitude: null,
      longitude: null,
    };
  }

  if (Array.isArray(pointInput)) {
    const values = flattenSheetInput(pointInput)
      .filter(function(item) {
        return item !== '' && item !== null && item !== undefined;
      })
      .map(function(item) {
        return typeof item === 'number' ? item : Number(item);
      });

    if (values.length >= 2 && values.every(function(value) { return !Number.isNaN(value); })) {
      return {
        kind: 'coordinates',
        cacheValue: values[0] + ',' + values[1],
        address: null,
        latitude: values[0],
        longitude: values[1],
      };
    }
  }

  throw new Error(
    'Point inputs must be an address string or a two-value coordinate pair such as {40.7128, -74.0060}.',
  );
}

/**
 * Applies a point object to the direction finder as origin or destination.
 * @param {Object} directionFinder
 * @param {string} methodName
 * @param {{kind: string, address: string|null, latitude: number|null, longitude: number|null}} point
 * @returns {Object}
 */
function applyPointToDirections(directionFinder, methodName, point) {
  if (point.kind === 'coordinates') {
    return directionFinder[methodName](point.latitude, point.longitude);
  }

  return directionFinder[methodName](point.address);
}

/**
 * Geocodes an address or uses coordinates directly.
 * @param {{kind: string, address: string|null, latitude: number|null, longitude: number|null}} point
 * @returns {{latitude: number, longitude: number}}
 */
function resolveCoordinates(point) {
  if (point.kind === 'coordinates') {
    return {
      latitude: point.latitude,
      longitude: point.longitude,
    };
  }

  const geocodeResponse = Maps.newGeocoder().geocode(point.address);
  if (geocodeResponse.status !== 'OK') {
    throw new Error('Could not geocode "' + point.address + '".');
  }

  const location = geocodeResponse.results[0].geometry.location;
  return {
    latitude: location.lat,
    longitude: location.lng,
  };
}

/**
 * Converts meters into the requested distance unit.
 * @param {number} distanceMeters
 * @param {string} distanceUnit
 * @returns {number}
 */
function convertDistance(distanceMeters, distanceUnit) {
  const input = String(distanceUnit || 'mi').toLowerCase().trim();
  const byAbbrev = APP_CONFIG.distanceUnits[input];
  const resolvedKey = byAbbrev ? input : Object.keys(APP_CONFIG.distanceUnits).find(function(key) {
    return APP_CONFIG.distanceUnits[key].label === input;
  });
  const unitConfig = resolvedKey && APP_CONFIG.distanceUnits[resolvedKey];

  if (!unitConfig) {
    throw new Error('Unknown unit "' + distanceUnit + '". Use km, m, mi, ft, nm, or the full name (e.g. "miles").');
  }

  return Number((distanceMeters / unitConfig.meters).toFixed(2));
}

/**
 * Returns a rounded number with a fixed precision.
 * @param {number} value
 * @param {number} decimalPlaces
 * @returns {number}
 */
function roundNumber(value, decimalPlaces) {
  return Number(Number(value).toFixed(decimalPlaces));
}

/**
 * Removes HTML tags from directions instructions and warnings.
 * @param {string} htmlText
 * @returns {string}
 */
function stripHtml(htmlText) {
  return String(htmlText || '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

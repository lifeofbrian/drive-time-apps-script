/**
 * Loads DriveTime.gs into a plain Node context with GAS globals stubbed out.
 * Returns all top-level functions and APP_CONFIG for unit testing.
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const GAS_STUBS = {
  CacheService: {
    getDocumentCache: () => ({ get: () => null, put: () => {} }),
  },
  Maps: {
    newDirectionFinder: () => ({
      setMode: function() { return this; },
      setOrigin: function() { return this; },
      setDestination: function() { return this; },
      setDepart: function() { return this; },
      getDirections: () => ({ routes: [] }),
    }),
    newGeocoder: () => ({
      geocode: () => ({ status: 'ZERO_RESULTS', results: [] }),
    }),
  },
  Utilities: {
    DigestAlgorithm: { MD5: 'MD5' },
    computeDigest: (_alg, input) =>
      Array.from(Buffer.from(input)).map(b => b - 128),
  },
};

const src = fs.readFileSync(
  path.resolve(__dirname, '../DriveTime.gs'),
  'utf8',
);

// `const` at the top level of a vm script is block-scoped to the script and
// not added to the sandbox object. Wrapping in a function and returning an
// explicit exports object is the only reliable way to surface them.
const wrapped = `
(function(exports) {
  ${src}
  exports.APP_CONFIG = APP_CONFIG;
  exports.convertDistance = convertDistance;
  exports.roundNumber = roundNumber;
  exports.stripHtml = stripHtml;
  exports.flattenSheetInput = flattenSheetInput;
  exports.normalizePoint = normalizePoint;
  exports.classifyCommuteLevel = classifyCommuteLevel;
  exports.classifyTrafficLevel = classifyTrafficLevel;
  exports.detectRouteAccidents = detectRouteAccidents;
  exports.getCommuteThresholds = getCommuteThresholds;
  exports.getTrafficThresholds = getTrafficThresholds;
  exports.STRAIGHT_LINE_DISTANCE = STRAIGHT_LINE_DISTANCE;
})
`;

const sandbox = { ...GAS_STUBS };
vm.createContext(sandbox);
const factory = vm.runInContext(wrapped, sandbox);
const out = {};
factory(out);

module.exports = out;

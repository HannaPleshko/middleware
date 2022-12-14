/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* This was originally from the package node-ical.js: https://github.com/jens-maus/node-ical */
/* We pulled it into our project and created quattro-node-ical from it so we could modify the dependent */
/* ical.js (creating quattro-ical.js) modify the parsing to include the COMMENT field. */
/* We did modify the code to fix linting errors/warnings */


/* eslint-disable @typescript-eslint/no-floating-promises */
const fs = require('fs');
const request = require('request');

const ical = require('./quattro-ical.js');

/**
 * ICal event object.
 *
 * These two fields are always present:
 *  - type
 *  - params
 *
 * The rest of the fields may or may not be present depending on the input.
 * Do not assume any of these fields are valid and check them before using.
 * Most types are simply there as a general guide for IDEs and users.
 *
 * @typedef iCalEvent
 * @type {object}
 *
 * @property {string} type           - Type of event.
 * @property {Array} params          - Extra event parameters.
 *
 * @property {?object} start         - When this event starts.
 * @property {?object} end           - When this event ends.
 *
 * @property {?string} summary       - Event summary string.
 * @property {?string} description   - Event description.
 *
 * @property {?object} dtstamp       - DTSTAMP field of this event.
 *
 * @property {?object} created       - When this event was created.
 * @property {?object} lastmodified  - When this event was last modified.
 *
 * @property {?string} uid           - Unique event identifier.
 *
 * @property {?string} status        - Event status.
 *
 * @property {?string} sequence      - Event sequence.
 *
 * @property {?string} url           - URL of this event.
 *
 * @property {?string} location      - Where this event occurs.
 * @property {?{
 *     lat: number, lon: number
 * }} geo                            - Lat/lon location of this event.
 *
 * @property {?Array.<string>}       - Array of event catagories.
 */
/**
 * Object containing iCal events.
 * @typedef {Object.<string, iCalEvent>} iCalData
 */
/**
 * Callback for iCal parsing functions with error and iCal data as a JavaScript object.
 * @callback icsCallback
 * @param {Error} err
 * @param {iCalData} ics
 */
/**
 * A Promise that is undefined if a compatible callback is passed.
 * @typedef {(Promise.<iCalData>|undefined)} optionalPromise
 */

// utility to allow callbacks to be used for promises
function promiseCallback(fn, cb) {
  const promise = new Promise(fn);
  if (!cb) {
    return promise;
  }

  promise
    .then(returnValue => {
      cb(null, returnValue);
    })
    .catch(error => {
      cb(error, null);
    });
}

// Sync functions
const sync = {};
// Async functions
const async = {};
// Auto-detect functions for backwards compatibility.
const autodetect = {};

/**
 * Download an iCal file from the web and parse it.
 *
 * @param {string} url                - URL of file to request.
 * @param {Object|icsCallback} [opts] - Options to pass to request() from npm:request.
 *                                      Alternatively you can pass the callback function directly.
 *                                      If no callback is provided a promise will be returned.
 * @param {icsCallback} [cb]          - Callback function.
 *                                      If no callback is provided a promise will be returned.
 *
 * @returns {optionalPromise} Promise is returned if no callback is passed.
 */
async.fromURL = function (url, options, cb) {
  return promiseCallback((resolve, reject) => {
    request(url, options, (err, response, data) => {
      if (err) {
        reject(err);
        return;
      }

      // If (r.statusCode !== 200) {
      // all ok status codes should be accepted (any 2XX code)
      if (Math.floor(response.statusCode / 100) !== 2) {
        reject(new Error(`${response.statusCode} ${response.statusMessage}`));
        return;
      }

      ical.parseICS(data, (error, ics) => {
        if (error) {
          reject(error);
          return;
        }

        resolve(ics);
      });
    });
  }, cb);
};

/**
 * Load iCal data from a file and parse it.
 *
 * @param {string} filename   - File path to load.
 * @param {icsCallback} [cb]  - Callback function.
 *                              If no callback is provided a promise will be returned.
 *
 * @returns {optionalPromise} Promise is returned if no callback is passed.
 */
async.parseFile = function (filename, cb) {
  return promiseCallback((resolve, reject) => {
    fs.readFile(filename, 'utf8', (err, data) => {
      if (err) {
        reject(err);
        return;
      }

      ical.parseICS(data, (error, ics) => {
        if (error) {
          reject(error);
          return;
        }

        resolve(ics);
      });
    });
  }, cb);
};

/**
 * Parse iCal data from a string.
 *
 * @param {string} data       - String containing iCal data.
 * @param {icsCallback} [cb]  - Callback function.
 *                              If no callback is provided a promise will be returned.
 *
 * @returns {optionalPromise} Promise is returned if no callback is passed.
 */
async.parseICS = function (data, cb) {
  return promiseCallback((resolve, reject) => {
    ical.parseICS(data, (err, ics) => {
      if (err) {
        reject(err);
        return;
      }

      resolve(ics);
    });
  }, cb);
};

/**
 * Load iCal data from a file and parse it.
 *
 * @param {string} filename   - File path to load.
 *
 * @returns {iCalData} Parsed iCal data.
 */
sync.parseFile = function (filename) {
  const data = fs.readFileSync(filename, 'utf8');
  return ical.parseICS(data);
};

/**
 * Parse iCal data from a string.
 *
 * @param {string} data - String containing iCal data.
 *
 * @returns {iCalData} Parsed iCal data.
 */
sync.parseICS = function (data) {
  return ical.parseICS(data);
};

/**
 * Load iCal data from a file and parse it.
 *
 * @param {string} filename   - File path to load.
 * @param {icsCallback} [cb]  - Callback function.
 *                              If no callback is provided this function runs synchronously.
 *
 * @returns {iCalData|undefined} Parsed iCal data or undefined if a callback is being used.
 */
autodetect.parseFile = function (filename, cb) {
  if (!cb) {
    return sync.parseFile(filename);
  }

  async.parseFile(filename, cb);
};

/**
 * Parse iCal data from a string.
 *
 * @param {string} data       - String containing iCal data.
 * @param {icsCallback} [cb]  - Callback function.
 *                              If no callback is provided this function runs synchronously.
 *
 * @returns {iCalData|undefined} Parsed iCal data or undefined if a callback is being used.
 */
autodetect.parseICS = function (data, cb) {
  if (!cb) {
    return sync.parseICS(data);
  }

  async.parseICS(data, cb);
};

// Export api functions
module.exports = {
  // Autodetect
  fromURL: async.fromURL,
  parseFile: autodetect.parseFile,
  parseICS: autodetect.parseICS,
  // Sync
  sync,
  // Async
  async,
  // Other backwards compat things
  objectHandlers: ical.objectHandlers,
  handleObject: ical.handleObject,
  parseLines: ical.parseLines
};

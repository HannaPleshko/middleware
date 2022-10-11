/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */

// Config
// keepBaseUrl: process.env.KEEP_BASE_URL || 'https://frascati.projectkeep.io',
// keepBaseUrl: process.env.KEEP_BASE_URL || 'http://localhost:8880',

module.exports = {
    protocol: process.env.PROTOCOL || 'http',
    pfx: process.env.PFX ===  undefined ? '' : require('fs').readFileSync(process.env.PFX),
    passphrase: process.env.PASSPHRASE || '',
    port: process.env.PORT || 3000,
    host: process.env.HOST,
    maxRequestBodySize: process.env.MAX_REQUEST_BODY_SIZE || '5MB',
    serverUrl: process.env.SERVER_URL,
    keepAliveInterval: process.env.KEEP_ALIVE_INTERVAL || 30000, // 30 seconds
    // The `gracePeriodForClose` provides a graceful close for http/https
    // servers with keep-alive clients. The default value is `Infinity`
    // (don't force-close). If you want to immediately destroy all sockets
    // upon stop, set its value to `0`.
    // See https://www.npmjs.com/package/stoppable
    gracePeriodForClose: process.env.GRACE_PERIOD_FOR_CLOSE || 5000, // 5 seconds
    keepBaseUrl: process.env.KEEP_BASE_URL || 'http://localhost:8880', 
    logToFile: process.env.LOG_TO_FILE || "true",
    logLevel: process.env.LOG_LEVEL || "debug"
};
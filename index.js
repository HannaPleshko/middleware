/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
require('dotenv').config();
const config = require('./config');
const application = require('./dist');

// Run the application
application.main({ rest: config }).catch(err => {
    console.error('Cannot start the application.', err);
    process.exit(1);
});

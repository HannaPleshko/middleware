# Writing Unit Tests

## Unit test frameworks

The EWS middleware project uses [Mocha](https://mochajs.org/) as it's unit test runner. Mocha is a simple but highly extendable framework for Nodejs. Also the project uses [Should.js](https://shouldjs.github.io/) for an assertion framwork and [Sinon.js](http://sinonjs.org/) for it's mocking framework. Both Should.js and Sinon.js are included with loopback in the @looback/labtest package. 

## Setting up Mocha Test Explorer in VSCode

VSCode contains a test explorer that can be accessed from the left panel (flask icon). You will need to install Mocha Test Explorer extension to have the Mocha test show up in the Test Explorer panel. When Mocha Test Explorer is installed you can use this to run unit tests. Perform the following one time setup to configure the Mocha Test Explorer. 

1. Click on the Configuration icon a the bottom of the left panel  and click on Settings.
1. In the "Search Settings" box, search for `mocha` 
1. In the results, find "Mocha Explorer:Files" and click on `Edit in settings.json`.
1. In the settings.json file add the following lines:
    ```
    "mochaExplorer.files": "src/tests/**/*.ts",
    "mochaExplorer.require": "ts-node/register",
    "mochaExplorer.logpanel": true
    ```
1. Save the settings.json file and restart VSCode.
1. Click on the Test icon (flask) on the left panel to open the test explorer. You should see a list of unit tests that have been created in the project. By hovering your mouse over the test, you can run one or all test from this list and see the results. You can also jump to the source code for the test from this list.   

## Creating a unit test

Unit test files should be create under the `src/tests` directory. A standard practice is to create one unit test file for each code file. Inside a unit test file there should be a `describe()` statement for each function being tested. Each test for a function should be wrapped in an `it()` statement. For larger functions, you can also included imbedded `describe()` statements inside of the main `describe()`.

Mocha also supports before(), after(), beforeEach(), and afterEach() functions for setup or teardown before/after each test is called. A good descussion on using `describe()`, `it()`, and `beforeEach()` can be found here: https://www.bignerdranch.com/blog/why-do-javascript-test-frameworks-use-describe-and-beforeeach/

### Mocking

To mock certain objects and functions use Sinon.js. See the [Sinon.js documentation](http://sinonjs.org/) for more information. 

When you setup mocking, such as creating a stub, you can reset these by calling sinon.restore() in the afterEach function so they are not set when the next test starts. 

The [Loopback testlab package](https://www.npmjs.com/package/@loopback/testlab) contains helper methods for writing unit test for loopback. See [Test code expecting Express Request or Response](https://www.npmjs.com/package/@loopback/testlab#test-code-expecting-express-request-or-response) for examples for creating loopback request and responses. 

### Things to know

1. In order to reference `this` in a test case you must use `function() {}` instead of `() => {}` to get access to Mocha's test context object. For example to access the title of the currently running test case:
    ```
    beforeEach(function () {
            console.info(`Setup before "${this.currentTest?.title}"`);
    }
    ```
    Note that lint rules will tag the use of this. In your unit test files you can turn this off by placing the following at the top of the unit test file: `/* eslint-disable no-invalid-this */`

1. You can use `console` to write out statement in a test. In order to view the console output select a test case in the Test Explorer. It will show the output at the bottom under the OUTPUT tab. 
1. If test do not run, check the Mocha Explorer Log under the OUTPUT tab at the bottom. It will show any runtime exceptions. 
1. To run the test with different time zones you can edit the Settings and under `Mocha Exploer: Env` click on `Edit in settings.json`. Then set the time zone environment variable TZ to the IANA time zone name. For example:
   ```
   "mochaExplorer.env": {
    "TZ": "Europe/Moscow"
    }
    ```
    The IANA time zone names can be found here: https://en.wikipedia.org/wiki/List_of_tz_database_time_zones

## Running unit test from the command line

If you want to run all the unit test from the command line, use the following command:

```
npm run test:unit
```
const winston = require('winston');
const { LEVEL, MESSAGE } = require('triple-beam');

/**
 * This class is to fix the issue with the Winston Console transport (winston.transports.Console) not writing to the Visual Studio Code debug console. 
 * This issue is documented here: https://github.com/winstonjs/winston/issues/1544
 */
export class ConsoleTransport extends winston.transports.Console {
    constructor(opts: any) {
        super(opts);
    }

    log(info: any, callback: () => void): void {
        setImmediate(() => {
            this.emit('logged', info);
        });

        /**
         * This code was taken from winston.transports.Console (https://github.com/winstonjs/winston/blob/master/lib/winston/transports/console.js) and modified to always use console to write the message. 
         */
        const msg = info[MESSAGE];
        if (this.stderrLevels[info[LEVEL]]) {
            console.error(msg);

            if (callback) {
                callback(); 
            }
            return;

        } else if (this.consoleWarnLevels[info[LEVEL]]) {
            console.warn(msg);

            if (callback) {
                callback(); 
            }
            return;
        }

        console.log(msg);

        if (callback) {
            callback(); 
        }

    }
}
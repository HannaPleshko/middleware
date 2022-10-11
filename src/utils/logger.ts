import fs from 'fs';
const winston = require('winston');
require('winston-daily-rotate-file');

const config = require('../../config')
import { ConsoleTransport } from './';
import { OpenClientLogger } from '@hcllabs/openclientkeepcomponent';

/**
 * Identifies a log file. 
 */
export enum LogIdentifier {
    /**
     * Log message to the combined log
     */
    COMBINED = "ews-combined-log"
}

/**
 * Used for logging message to the log or console. 
 */
export class Logger implements OpenClientLogger {

    private static instance: Logger;

    public static getInstance(): Logger {
        if (!Logger.instance) {
            this.instance = new Logger();
        }
        return this.instance;
    }

    constructor() {
        const consoleTransport = new ConsoleTransport({
            handleExceptions: true,
            handleRejections: true,
            stderrLevels: ['error']
        }); // Used to send log messages to the console

        // Always log to console
        const logTransports: any[] = [consoleTransport];

        /**
         * Setup combined logger
         */

        // If logging to file is enabled, setup the files
        if (config.logToFile === "true") {
            const logDir = 'logs';
            if (!fs.existsSync(logDir)) {
                fs.mkdirSync(logDir);
            }

            /**
             * Setup the combined logger files
             */
            const transport = new (winston.transports.DailyRotateFile)({
                dirname: logDir,
                filename: 'ews-%DATE%.log',
                auditFile: 'logs/ews-audit.json',
                datePattern: 'YYYY-MM-DD',
                maxSize: '5m',
                maxFiles: '5',
                handleExceptions: true,
                handleRejections: true
            });
            logTransports.push(transport);
        }

        // Add logger
        winston.loggers.add(LogIdentifier.COMBINED, {
            level: config.logLevel,
            format: winston.format.combine(
                winston.format.timestamp({
                    format: 'YYYY-MM-DD HH:mm:ss.SSS'
                }),
                winston.format.printf((info: { timestamp: any; level: any; message: any }) => `${info.timestamp} ${info.level}: ${info.message}`)
            ),
            transports: logTransports
        });

    }

    /**
     * Log a message used for debugging. 
     * @param message The message to log
     * @param logIds Optional identifiers to indicate which logs to use. The default is the combined log. 
     */
    public debug(message: string, logIds: LogIdentifier[] = [LogIdentifier.COMBINED]): void {
        this.log("debug", message, logIds);

    }

    /**
     * Log a message when verbose logging is enabled. 
     * @param message The message to log
     * @param logIds Optional identifiers to indicate which logs to use. The default is the combined log. 
     */
    public verbose(message: string, logIds: LogIdentifier[] = [LogIdentifier.COMBINED]): void {
        this.log("verbose", message, logIds);

    }

    /**
     * Log a message that shows an http request/response. 
     * @param message The message to log
     * @param logIds Optional identifiers to indicate which logs to use. The default is the combined log. 
     */
    public http(message: string, logIds: LogIdentifier[] = [LogIdentifier.COMBINED]): void {
        this.log("http", message, logIds);

    }

    /**
     * Log an informational message. 
     * @param message The message to log
     * @param logIds Optional identifiers to indicate which logs to use. The default is the combined log. 
     */
    public info(message: string, logIds: LogIdentifier[] = [LogIdentifier.COMBINED]): void {
        this.log("info", message, logIds);

    }

    /**
     * Log a warning message. 
     * @param message The message to log
     * @param logIds Optional identifiers to indicate which logs to use. The default is the combined log. 
     */
    public warn(message: string, logIds: LogIdentifier[] = [LogIdentifier.COMBINED]): void {
        this.log("warn", message, logIds);

    }

    /**
     * Log an error message. 
     * @param message The message to log
     * @param logIds Optional identifiers to indicate which logs to use. The default is the combined log. 
     */
    public error(message: string, logIds: LogIdentifier[] = [LogIdentifier.COMBINED]): void {
        this.log("error", message, logIds);

    }

    /**
     * Log a message. 
     * @param level The logging level. For possible values see https://github.com/winstonjs/winston#logging-levels
     * @param message The message to log
     * @param logIds Identifiers to indicate which logs to log the message in.
     */
    private log(level: string, message: string, logIds: LogIdentifier[]): void {
        logIds.forEach(logId => {
            const wLogger = winston.loggers.get(logId);
            if (wLogger) {
                wLogger.log(level, message);
            }
        });
    }
}


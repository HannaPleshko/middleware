/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import path from 'path';
import { ApplicationConfig } from '@loopback/core';
import { Response } from 'express';
import { createSoapServer } from './soap.server';
import { EWSMiddlewareApplication } from './application';
import { ExchangeSoap, AutodiscoverSoap } from './soap';
import util from 'util';
import { Logger } from './utils';
import { OpenClientKeepComponent } from '@hcllabs/openclientkeepcomponent';
import { OpenClientKeepOptionsProvider } from './keepcomponent/OpenClientKeepOptionsProvider';
import i18n = require("i18n");

const config = require('../config')

/**
 * Main entrypoint of server.
 * @param options The server options
 */

export let appContext: EWSMiddlewareApplication;

export async function main(options: ApplicationConfig = {}): Promise<EWSMiddlewareApplication> {
    const app = new EWSMiddlewareApplication(options);
    appContext = app;
    app.component(OpenClientKeepComponent);
    OpenClientKeepComponent.keepOptions = new OpenClientKeepOptionsProvider();
    // OpenClientKeepOptionsProvider.setupOpenClient(appContext);

    // Boot app
    await app.boot();
    
    //Configure translation module 
    i18n.configure({
        directory: path.join(__dirname, '/locales')
    });

    // Create soap server
    await createSoapServer(app, ExchangeSoap);
    await createSoapServer(app, AutodiscoverSoap);

    // Serve static wsdl/xsd files
    app.static('/EWS', path.resolve(__dirname, '../wsdl/ews'), {
        setHeaders: (res: Response, uri: string) => {
            if (uri && (uri.endsWith('.wsdl') || uri.endsWith('.xsd'))) {
                res.setHeader('Content-Type', 'text/xml');
            }
        }
    });
    app.static('/autodiscover', path.resolve(__dirname, '../wsdl/autodiscover'), {
        setHeaders: (res: Response, uri: string) => {
            if (uri && (uri.endsWith('.wsdl') || uri.endsWith('.xsd'))) {
                res.setHeader('Content-Type', 'text/xml');
            }
        }
    });

    // Serve a static main page. Use to verify service is started.
    app.static('/', path.resolve(__dirname, '../html'), {
        setHeaders: (res: Response, uri: string) => {
            res.setHeader('Content-Type', 'text/html');
        }
    });

    // Start server
    await app.start();
    Logger.getInstance().debug(`Server started on ${new Date()}`);
    Logger.getInstance().debug(`Config settings: ${util.inspect(config, false, 5)}`);
    Logger.getInstance().debug(`Server is running at ${app.restServer.url}`);

    // Tell the OpenClient Keep Component to use our logger
    OpenClientKeepComponent.commonLogger = Logger.getInstance();

    return app;
}

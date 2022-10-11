/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import fs from 'fs';
import path from 'path';
import { Request } from '@loopback/rest';
import { SOAP } from './soap';
import { EWSMiddlewareApplication } from './application';
import { objectFactory } from './models/objectFactory';
import { ResponseCodeType } from './models/enum.model';
import { ServerVersionInfo, AttributesKey, ValueKey, XSITypeKey, SOAPFault } from './models/common.model';
import { Logger } from './utils';

const SOAPParser = require('strong-soap/src/parser');
const SOAPServer = require('strong-soap/src/server');
const QName = SOAPParser.QName;

/**
 * Create soap server.
 * @param app The application
 * @param soap The SOAP definition
 */
export async function createSoapServer(app: EWSMiddlewareApplication, soap: SOAP): Promise<void> {

    // Get services
    const serviceBinds = app.findByTag(soap.service);

    // Use proxy to get the service operation by name
    const operations: any = {};
    const ports = new Proxy(operations, {
        get: (proxy: any, name: string): any => {
            if (name) {
                if (operations[name]) {
                    return operations[name];
                }
                for (const serviceBind of serviceBinds) {
                    const serviceBound = app.getSync<any>(serviceBind.key);
                    let serviceMethod = serviceBound[name];
                    if (serviceMethod && typeof serviceMethod === 'function') {
                        serviceMethod = serviceMethod.bind(serviceBound);
                        operations[name] = (soapRequest: any, callback: Function, soapHeaders: object, request: Request): void => {
                            soapRequest.SOAPHeaders = soapHeaders;
                            
                            const cb = (err: any, result: any, streaming?: boolean): void => {
                                callback(err, result, streaming);
                            }
                            serviceMethod(soapRequest, request, cb)
                                .then((result: any) => cb(null, result))
                                .catch((err: any) => {
                                    Logger.getInstance().error(`${err}`);
                                    let error: any;
                                    if (err instanceof SOAPFault) {
                                        error = err;
                                    } else {
                                        if (err.status && err.status === 401) {
                                            error = new SOAPFault(ResponseCodeType.ERROR_INVALID_AUTHORIZATION_CONTEXT, err.message, 401);
                                        } else {
                                            error = new SOAPFault(ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR, err.message);
                                        }
                                    }
                                    callback(error);
                                });
                        };
                        return operations[name];
                    }
                }
            }
            Logger.getInstance().warn('Invalid request operation: ' + name);
            throw new SOAPFault(ResponseCodeType.ERROR_INVALID_REQUEST, 'The request is invalid.')
        }
    });

    // Load WSDL
    const options = {
        valueKey: ValueKey,
        xsiTypeKey: XSITypeKey,
        attributesKey: AttributesKey,
        objectFactory: objectFactory(soap),
        ignoreUnknownProperties: true,
    };

    const uri = path.resolve(__dirname, soap.wsdl);
    const wsdl = new SOAPParser.WSDL(fs.readFileSync(uri).toString(), uri, options);

    // Create SOAP server
    const server: any = await new Promise((resolve) => {
        const _server = new SOAPServer({
            listeners: (): any[] => [],
            removeAllListeners: (): void => { return; },
            addListener: (): void => resolve(_server),
        }, soap.path, { [soap.service]: { [soap.port]: ports } }, wsdl, options)
    });

    // Add ServerVersionInfo to SOAP header
    server.addSoapHeader(new ServerVersionInfo(), new QName(soap.namespace, 'ServerVersionInfo'));

    // Bind SOAP server to application.
    app.bind(soap.service).to(server);
}
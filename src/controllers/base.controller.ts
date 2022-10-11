/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import zlib from 'zlib';
import accepts from 'accepts';
import { Request, Response } from '@loopback/rest';
import { Context } from '@loopback/context';
import { SOAPFault } from '../models/common.model';
import { ResponseCodeType } from '../models/enum.model';
import { SOAP } from '../soap';
import format from 'xml-formatter';
import { Logger } from '../utils';
import { UserManager } from '@hcllabs/openclientkeepcomponent';
import { UserContext } from '../keepcomponent/UserContext';
import i18n = require("i18n");
/**
 * Base controller for SOAP middleware.
 */
export abstract class BaseController {
    
    /**
     * Constructor.
     * @param soap The SOAP definition
     */
    constructor(private readonly soap: SOAP) { }
    /**
     * Write response using gzip/deflate compression.
     * @param result The result to write
     * @param statusCode The status code
     * @param request The HTTP request object
     * @param response The HTTP response object
     * @param streaming Indicates whether this is a long-holding streaming response
     */
    private static writeResponse(result: string, statusCode: number, request: Request, response: Response, streaming?: boolean): void {
        let compress;

        if (!response.headersSent) {
            if (statusCode) {
                response.statusCode = statusCode;
            }
            response.setHeader('Content-Type', 'text/xml; charset=utf-8');
            if (!streaming) {
                const accept = accepts(request);
                if (accept.encoding('gzip')) {
                    response.setHeader('Content-Encoding', 'gzip');
                    compress = zlib.createGzip();
                } else if (accept.encoding('deflate')) {
                    response.setHeader('Content-Encoding', 'deflate');
                    compress = zlib.createDeflate();
                }
            }
        }

        if (streaming) {
            response.write(result);
            return;
        }

        if (compress) {
            compress.pipe(response);
            compress.write(result);
            compress.end();
        } else {
            response.end(result);
        }
    }

    /**
     * HTTP request handler.
     * @param xml The input xml string
     * @param request The HTTP request object
     * @param response The HTTP response object
     * @param ctx The request context
     */
    protected async handle(xml: string, request: Request, response: Response, ctx: Context): Promise<Response> {

        const soapServer = await ctx.get<any>(this.soap.service);

        // Get the authentication token from the request if it exist.
        const authToken = request.header("authorization")?.split(" ")[1];

        return new Promise<Response>((resolve, reject) => {

            if (!authToken) {
                // Apple Mail requires authentication. It does not include the X-User-Identity, but includes the authorization 
                // header with an auth token that can be used to get the user.  
                // TODO: Replace basic auth with OAuth authentication here
                response.setHeader("WWW-Authenticate", "Basic Realm=\"\"");
                response.statusCode = 401;
                response.setHeader("Content-Length", 0);
                Logger.getInstance().http(`${request.method} response for ${request.url}: 401\n`)
                response.end();
            }
            else {
                try {
                    //Initialize i18n middleware
                    i18n.init(request, response);

                    soapServer._process(xml, request, (result: string, statusCode: number, streaming: boolean) => {
                        if (401 === statusCode) {
                            const userInfo = UserContext.getUserInfo(request);
                            if (userInfo) {
                                UserManager.getInstance().clearToken(userInfo.userId); // Clear the locally cached access token
                            }
                            response.setHeader("WWW-Authenticate", "Basic Realm=\"\"");
                            response.statusCode = 401;
                            response.setHeader("Content-Length", 0);
                            Logger.getInstance().http(`${request.method} response for ${request.url}: 401\n`)
                            response.end();
                        }
                        if (response.connection.destroyed) {
                            if (streaming) {
                                throw new Error('Streaming connection closed');
                            }
                            return resolve(response);
                        }

                        if (response.statusCode !== 401) {
                            Logger.getInstance().http(`${request.method} response for ${request.url}:\n${format(result,{collapseContent: true})}\n`)
                        }
                        
                        BaseController.writeResponse(result, statusCode, request, response, streaming);
                        if (!streaming) {
                            resolve(response);
                        }
                    });
                } catch (err) {
                    Logger.getInstance().error(`Soap Server error: ${err}`);
                    const error = err instanceof SOAPFault ? err : new SOAPFault(ResponseCodeType.ERROR_INVALID_REQUEST, err.message);
                    if (error.Fault.statusCode === 401) {
                        const userInfo = UserContext.getUserInfo(request);
                        if (userInfo) {
                            UserManager.getInstance().clearToken(userInfo.userId); // Clear the locally cached access token
                        }
                        response.setHeader("WWW-Authenticate", "Basic Realm=\"\"");
                        response.statusCode = 401;
                        response.setHeader("Content-Length", 0);
                        Logger.getInstance().http(`${request.method} response for ${request.url}: 401\n`)
                        response.end();
                    } else {
                        const operation = soapServer.wsdl.definitions.services.ExchangeService
                        .ports.ExchangeServicePort.binding.operations.GetUserPhoto;
                        soapServer._sendError(operation, error, (faultXml: string) => {
                            BaseController.writeResponse(faultXml, 500, request, response);
                            resolve(response);
                        });
                    }
                }
            }
        });
    }
}

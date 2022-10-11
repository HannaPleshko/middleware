/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import xml2js from 'xml2js';
import { inject, Context } from '@loopback/context';
import { Request, Response, requestBody, RestBindings, post } from '@loopback/rest';
import { AutodiscoverSoap } from '../soap';
import { BaseController } from './base.controller';
import { PoxAutodiscoverRequestMessage, PoxAutodiscoverResponseMessage, PoxAutodiscoverResponse, PoxResponse } from '../models/autodiscover.model';
import format from 'xml-formatter';
import { Logger } from '../utils'

const config = require('../../config');

/**
 * Autodiscover controller as SOAP middleware.
 */
export class AutodiscoverController extends BaseController {

    /**
     * Constructor.
     */
    constructor() {
        super(AutodiscoverSoap);
    }

    /**
     * Autodiscover SOAP service.
     * @param xml The input xml string
     * @param request The HTTP request object
     * @param response The HTTP response object
     * @param ctx The request context
     */
    @post(AutodiscoverSoap.path)
    async autodiscover(
        @requestBody({ content: { '*/*': { 'x-parser': 'text' } } }) xml: string,
        @inject(RestBindings.Http.REQUEST) request: Request,
        @inject(RestBindings.Http.RESPONSE) response: Response,
        @inject.context() ctx: Context): Promise<Response> {

        Logger.getInstance().http(`${request.method} request for ${AutodiscoverSoap.path}:\n${format(xml,{collapseContent: true})}\n`)

        return super.handle(xml, request, response, ctx);
    }

    /**
     * Autodiscover POX service.
     * @param xml The input xml string
     * @param request The HTTP request object
     * @param response The HTTP response object
     * @param ctx The request context
     */
    @post('/autodiscover/autodiscover.xml')
    async pox(
        @requestBody({ content: { '*/*': { 'x-parser': 'text' } } }) xml: string,
        @inject(RestBindings.Http.REQUEST) request: Request,
        @inject(RestBindings.Http.RESPONSE) response: Response,
        @inject.context() ctx: Context): Promise<Response> {

        Logger.getInstance().http(`${request.method} request for ${request.url}:\n${format(xml,{collapseContent: true})}\n`)

        const parser = new xml2js.Parser({ explicitArray: false });
        const body: PoxAutodiscoverRequestMessage = await parser.parseStringPromise(xml);

        const email = body.Autodiscover.Request.EMailAddress

        const poxResponse = new PoxResponse();

        /*
        Customers should configure host autodiscovery.<SMTP-address-domain> to resolve to this endpoint address or rewrite 
        https://<SMTP-address-domain>/autodiscovery/autodiscovery to this endpoint.  
        
        See https://docs.microsoft.com/en-us/exchange/architecture/client-access/autodiscover for more information.
        */
        const hostname = email.split('@')[1];
        const serverUrl =  `https://${config.serverUrl ?? hostname}`;

        poxResponse.User = {
            DisplayName: email,
            AutoDiscoverSMTPAddress: email
        };
        poxResponse.Account = {
            AccountType: 'email',
            Action: 'settings',
            Protocol: [{
                Type: 'EXPR',
                Server: serverUrl,
                SSL: serverUrl.startsWith('https') ? 'on' : 'off',
                AuthPackage: 'basic',
                AuthRequired: 'on',
                ASUrl: `${serverUrl}/EWS/Exchange.asmx`,
                EwsUrl: `${serverUrl}/EWS/Exchange.asmx`,
                EmwsUrl: `${serverUrl}/EWS/Exchange.asmx`,
                OOFUrl: `${serverUrl}/EWS/Exchange.asmx`,
                SharingUrl: `${serverUrl}/EWS/Exchange.asmx`,
                EwsPartnerUrl: `${serverUrl}/EWS/Exchange.asmx`,
                OABUrl: `${serverUrl}/oab`,
                ServerExclusiveConnect: 'on'
            }]
        };

        const result = new PoxAutodiscoverResponseMessage(new PoxAutodiscoverResponse(poxResponse));
        const resultXml = new xml2js.Builder().buildObject(result);

        response.set('Content-Type', 'text/xml')
        Logger.getInstance().http(`${request.method} response for ${request.url}:\n${format(resultXml,{collapseContent: true})}\n`)
        response.end(resultXml);
        
        return response;
    }
}

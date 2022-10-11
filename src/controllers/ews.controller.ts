/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import { Request, Response, requestBody, RestBindings, post } from '@loopback/rest';
import { inject, Context } from '@loopback/context';
import { ExchangeSoap } from '../soap';
import { BaseController } from './base.controller';
import format from 'xml-formatter';
import { Logger } from '../utils';

/**
 * EWS controller as SOAP middleware.
 */
export class EWSController extends BaseController {

    /**
     * Constructor.
     */
    constructor() {
        super(ExchangeSoap);
    }

    /**
     * Exchange service.
     * @param xml The input xml string
     * @param request The HTTP request object
     * @param response The HTTP response object
     * @param ctx The request context
     */
    @post(ExchangeSoap.path)
    async exchange(
        @requestBody({ content: { '*/*': { 'x-parser': 'text' } } }) xml: string,
        @inject(RestBindings.Http.REQUEST) request: Request,
        @inject(RestBindings.Http.RESPONSE) response: Response,
        @inject.context() ctx: Context): Promise<Response> {
        Logger.getInstance().http(`${request.method} request for ${ExchangeSoap.path}:\n${format(xml,{collapseContent: true})}\n`)
        return super.handle(xml, request, response, ctx);
    }
}

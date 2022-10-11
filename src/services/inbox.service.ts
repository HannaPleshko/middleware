/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import assert from 'assert';
import { BindingScope, bind } from '@loopback/core';
import { ExchangeSoap } from '../soap';
import {
    GetInboxRulesRequestType, GetInboxRulesResponseType,
    UpdateInboxRulesRequestType, UpdateInboxRulesResponseType
} from '../models/inbox.model';
import { Request } from '@loopback/rest';
import { GetInboxRules, UpdateInboxRules } from '../data/inbox';

/**
 * The Inbox service.
 */
@bind({ scope: BindingScope.SINGLETON, tags: ExchangeSoap.service })
export class InboxService {

    /**
     * Get Inbox rules.
     * @param soapRequest The soap request, which is type of GetInboxRulesRequestType
     * @returns Inbox rules
     */
    async GetInboxRules(soapRequest: GetInboxRulesRequestType, request: Request): Promise<GetInboxRulesResponseType> {
        assert(soapRequest instanceof GetInboxRulesRequestType);
        return GetInboxRules(soapRequest, request)
    }

    /**
     * Update Inbox rules.
     * @param soapRequest The soap request, which is type of UpdateInboxRulesRequestType
     * @returns Update result
     */
    async UpdateInboxRules(soapRequest: UpdateInboxRulesRequestType, request: Request): Promise<UpdateInboxRulesResponseType> {
        assert(soapRequest instanceof UpdateInboxRulesRequestType);
        return UpdateInboxRules(soapRequest, request);
    }
}
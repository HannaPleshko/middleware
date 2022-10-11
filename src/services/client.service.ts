/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import assert from 'assert';
import { BindingScope, bind } from '@loopback/core';
import { ExchangeSoap } from '../soap';
import { GetClientAccessTokenType, GetClientAccessTokenResponseType } from '../models/client.model';

import { mockGetClientAccessTokenResponse } from '../data/client';

/**
 * The client service.
 */
@bind({ scope: BindingScope.SINGLETON, tags: ExchangeSoap.service })
export class ClientService {

    /**
     * Get client access token.
     * @param soapRequest The soap request, which is type of GetClientAccessTokenType
     * @returns client access token
     */
    async GetClientAccessToken(soapRequest: GetClientAccessTokenType): Promise<GetClientAccessTokenResponseType> {
        assert(soapRequest instanceof GetClientAccessTokenType);
        return mockGetClientAccessTokenResponse;
    }
}
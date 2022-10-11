/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import { ClientAccessTokenTypeType } from '../models/enum.model';
import { GetClientAccessTokenResponseType, GetClientAccessTokenResponseMessageType } from '../models/client.model';
import { GenericArray } from '../models/common.model';

// mock reponse for GetClientAccessToken
export const mockGetClientAccessTokenResponse: GetClientAccessTokenResponseType = {
    ResponseMessages: new GenericArray(
        [Object.assign(new GetClientAccessTokenResponseMessageType(), {
            Token: {
                Id: '1C50226D-04B5-4AB2-9FCD-42E236B59E4B',
                TokenType: ClientAccessTokenTypeType.CALLER_IDENTITY,
                TokenValue: 'eyJ0eXAmv0QitaJg',
                TTL: 479
            }
        }), Object.assign(new GetClientAccessTokenResponseMessageType(), {
            Token: {
                Id: '1C50226D-04B5-4AB2-9FCD-42E236B59E4B',
                TokenType: ClientAccessTokenTypeType.EXTENSION_CALLBACK,
                TokenValue: 'ldjOds4dlkjHdq3p',
                TTL: 579
            }
        })])
};

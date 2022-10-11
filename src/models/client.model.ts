/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import { ResponseMessageType, BaseRequestType, BaseResponseMessageType } from './common.model';
import { ClientAccessTokenTypeType } from './enum.model';

/**
 * Contains the response to a GetClientAccessToken operation request.
 */
export class GetClientAccessTokenResponseType extends BaseResponseMessageType<GetClientAccessTokenResponseMessageType> { }

/**
 * Specifies the response message for a GetClientAccessToken request.
 */
export class GetClientAccessTokenResponseMessageType extends ResponseMessageType {
    Token?: ClientAccessTokenType;
}

/**
 * Specifies a client access token.
 */
export class ClientAccessTokenType {
    Id: string;
    TokenType: ClientAccessTokenTypeType;
    TokenValue: string;
    TTL: number;
}

/**
 * Contains a request to get a client access token.
 */
export class GetClientAccessTokenType extends BaseRequestType {
    TokenRequests: NonEmptyArrayOfClientAccessTokenRequestsType;
}

/**
 * Contains an array of token requests.
 */
export class NonEmptyArrayOfClientAccessTokenRequestsType {
    TokenRequest: ClientAccessTokenRequestType[] = [];
}

/**
 * Specifies a single token request.
 */
export class ClientAccessTokenRequestType {
    Id: string;
    TokenType: ClientAccessTokenTypeType;
    Scope?: string;
    ResourceUri?: string;
}

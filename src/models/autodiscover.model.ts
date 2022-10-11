/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import { ExchangeVersion, AutodiscoverErrorCode, AutodiscoverSetting } from './enum.model';
import { AttributesKey, XSITypeKey } from './common.model';

/**
 * Request to get user settings.
 */
export class GetUserSettingsRequestMessage {
    Request: GetUserSettingsRequest;
}

/**
 * Request to get domain settings.
 */
export class GetDomainSettingsRequestMessage {
    Request: GetDomainSettingsRequest;
}

/**
 * Base autodiscover request.
 */
export abstract class AutodiscoverRequest {
}

/**
 * Get user settings request.
 */
export class GetUserSettingsRequest extends AutodiscoverRequest {
    Users: Users;
    RequestedSettings: RequestedSettings;
    RequestedVersion: ExchangeVersion;
}

/**
 * Get domain settings request.
 */
export class GetDomainSettingsRequest extends AutodiscoverRequest {
    Domains: Domains;
    RequestedSettings: RequestedSettings;
    RequestedVersion: ExchangeVersion;
}

/**
 * Users.
 */
export class Users {
    User: User[] = [];
}

/**
 * User.
 */
export class User {
    Mailbox: string;
}

/**
 * Requested settings.
 */
export class RequestedSettings {
    Setting: AutodiscoverSetting[] = [];
}

/**
 * Domains.
 */
export class Domains {
    Domain: string[] = [];
}

// response classes

/**
 * Response of user settings.
 */
export class GetUserSettingsResponseMessage {
    Response: GetUserSettingsResponse;
}

/**
 * Response of domain settings.
 */
export class GetDomainSettingsResponseMessage {
    Response: GetDomainSettingsResponse;
}

/**
 * Base autodiscover response.
 */
export abstract class AutodiscoverResponse {
    ErrorCode: AutodiscoverErrorCode;
    ErrorMessage?: string;
}

/**
 * Get user settings response.
 */
export class GetUserSettingsResponse extends AutodiscoverResponse {
    UserResponses: ArrayOfUserResponse;
}

/**
 * Get domain settings response.
 */
export class GetDomainSettingsResponse extends AutodiscoverResponse {
    DomainResponses: ArrayOfDomainResponse;
}

/**
 * Array of user response.
 */
export class ArrayOfUserResponse {
    UserResponse: UserResponse[];
}

/**
 * Array of domain response.
 */
export class ArrayOfDomainResponse {
    DomainResponse: DomainResponse[];
}

/**
 * User response.
 */
export class UserResponse extends AutodiscoverResponse {
    UserSettings: UserSettings;
    UserSettingErrors?: UserSettingErrors;
    RedirectTarget?: string;
}

/**
 * Domain response.
 */
export class DomainResponse extends AutodiscoverResponse {
    DomainSettings: DomainSettings;
    DomainSettingErrors?: DomainSettingErrors;
    RedirectTarget?: string;
}

/**
 * User setting errors.
 */
export class UserSettingErrors {
    UserSettingError: UserSettingError[];
}

/**
 * User settings.
 */
export class UserSettings {
    UserSetting: UserSetting[];
}

/**
 * Domain setting errors.
 */
export class DomainSettingErrors {
    DomainSettingError: DomainSettingError[];
}

/**
 * Domain settings.
 */
export class DomainSettings {
    DomainSetting: DomainSetting[];
}

/**
 * User setting error.
 */
export class UserSettingError {
    ErrorCode: AutodiscoverErrorCode;
    ErrorMessage: string;
    SettingName: string;
}

/**
 * User setting.
 */
export abstract class UserSetting {
    constructor(public Name: string) { }
}

/**
 * User string setting.
 */
export class StringSetting extends UserSetting {
    [AttributesKey] = {
        [XSITypeKey]: {
            type: 'StringSetting',
            xmlns: 'http://schemas.microsoft.com/exchange/2010/Autodiscover',
        }
    }
    constructor(public Name: string, public Value: string) {
        super(Name);
    }
}

/**
 * Domain setting error.
 */
export class DomainSettingError {
    ErrorCode: AutodiscoverErrorCode;
    ErrorMessage: string;
    SettingName: string;
}

/**
 * Domain setting.
 */
export abstract class DomainSetting {
    constructor(public Name: string) { }
}

/**
 * Domain string setting.
 */
export class DomainStringSetting extends DomainSetting {
    [AttributesKey] = {
        [XSITypeKey]: {
            type: 'DomainStringSetting',
            xmlns: 'http://schemas.microsoft.com/exchange/2010/Autodiscover',
        }
    }
    constructor(public Name: string, public Value: string) {
        super(Name);
    }
}

// POX models

/**
 * POX request message.
 */
export class PoxAutodiscoverRequestMessage {
    Autodiscover: PoxAutodiscoverRequest;
}

/**
 * POX autodiscover request.
 */
export class PoxAutodiscoverRequest {
    Request: PoxRequest;
}


/**
 * POX request.
 */
export class PoxRequest {
    AcceptableResponseSchema: string;
    EMailAddress: string;
}

/**
 * POX response message.
 */
export class PoxAutodiscoverResponseMessage {
    constructor(public Autodiscover: PoxAutodiscoverResponse) { }
}

/**
 * POX autodiscover response.
 */
export class PoxAutodiscoverResponse {
    $ = { xmlns: 'http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006' };
    constructor(public Response: PoxResponse) { }
}

/**
 * POX response.
 */
export class PoxResponse {
    $ = { xmlns: 'http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a' };
    User: PoxUser;
    Account: PoxAccount;
}

/**
 * POX user.
 */
export class PoxUser {
    DisplayName: string;
    AutoDiscoverSMTPAddress: string;
}

/**
 * POX account.
 */
export class PoxAccount {
    AccountType: string;
    Action: string;
    Protocol: PoxProtocol[];
}

/**
 * POX protocol.
 */
export class PoxProtocol {
    Type: string;
    Server: string;
    ServerDN?: string;
    PublicFolderServer?: string;
    EwsUrl: string;
    EmwsUrl: string;
    ASUrl: string;
    OOFUrl: string;
    SharingUrl: string;
    OABUrl: string;
    EwsPartnerUrl?: string;
    SSL?: string;
    AuthPackage?: string;
    AuthRequired?: string;
    ServerExclusiveConnect?: string;
}

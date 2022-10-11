/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import assert from 'assert';
import { bind, BindingScope } from '@loopback/core';
import { AutodiscoverSoap } from '../soap';
import { GetDomainSettingsRequestMessage, GetDomainSettingsResponseMessage, GetUserSettingsRequestMessage, GetUserSettingsResponseMessage } from '../models/autodiscover.model';

import { mockGetDomainSettingsResponseMessage, mockGetUserSettingsResponseMessage } from '../data/autodiscover';

/**
 * The autodiscover service.
 */
@bind({ scope: BindingScope.SINGLETON, tags: AutodiscoverSoap.service })
export class AutodiscoverService {

    /**
     * Get domain settings.
     * @param soapRequest The soap request, which is type of GetDomainSettingsRequestMessage
     * @returns domain settings
     */
    async GetDomainSettings(soapRequest: GetDomainSettingsRequestMessage): Promise<GetDomainSettingsResponseMessage> {
        assert(soapRequest instanceof GetDomainSettingsRequestMessage);
        return mockGetDomainSettingsResponseMessage;
    }

    /**
     * Get user settings.
     * @param soapRequest The soap request, which is type of GetUserSettingsRequestMessage
     * @returns user settings
     */
    async GetUserSettings(soapRequest: GetUserSettingsRequestMessage): Promise<GetUserSettingsResponseMessage> {
        assert(soapRequest instanceof GetUserSettingsRequestMessage);
        return mockGetUserSettingsResponseMessage;
    }
}
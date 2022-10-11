/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import { AutodiscoverErrorCode, AutodiscoverSetting } from '../models/enum.model';
import { GetDomainSettingsResponseMessage, DomainStringSetting, GetUserSettingsResponseMessage, StringSetting } from '../models/autodiscover.model';
import { appContext } from '../index';

// mock reponse for GetDomainSettings
export const mockGetDomainSettingsResponseMessage: GetDomainSettingsResponseMessage = {
    Response: {
        ErrorCode: AutodiscoverErrorCode.NO_ERROR,
        DomainResponses: {
            DomainResponse: [{
                ErrorCode: AutodiscoverErrorCode.NO_ERROR,
                DomainSettings: {
                    DomainSetting: [
                        new DomainStringSetting(AutodiscoverSetting.EXTERNAL_EWS_URL, `${appContext.options.rest.serverUrl}/EWS/Exchange.asmx`)
                    ]
                }
            }]
        }
    }
};

// mock reponse for GetUserSettings
export const mockGetUserSettingsResponseMessage: GetUserSettingsResponseMessage = {
    Response: {
        ErrorCode: AutodiscoverErrorCode.NO_ERROR,
        UserResponses: {
            UserResponse: [{
                ErrorCode: AutodiscoverErrorCode.NO_ERROR,
                UserSettings: {
                    UserSetting: [
                        new StringSetting(AutodiscoverSetting.USER_DISPLAY_NAME, 'Tom Mike'),
                        new StringSetting(AutodiscoverSetting.EWS_SUPPORTED_SCHEMAS, 'Exchange2007, Exchange2010'),
                        new StringSetting(AutodiscoverSetting.EXTERNAL_EWS_URL, `${appContext.options.rest.serverUrl}/EWS/Exchange.asmx`)
                    ]
                }
            }]
        }
    }
};

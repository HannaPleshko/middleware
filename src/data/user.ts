/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import { ResponseMessageType, Value, DaysOfWeekType } from '../models/common.model';
import {
    ExternalAudience, OofState, ResponseClassType, ResponseCodeType, SuggestionQuality, 
    LegacyFreeBusyType, FreeBusyViewEnum, DayOfWeekType, UserConfigurationPropertyTypeEnum } from '../models/enum.model';
import {
    GetUserPhotoResponseMessageType, SetUserPhotoResponseMessageType, GetUserOofSettingsResponseMessageType, 
    SetUserOofSettingsResponseMessageType, GetUserAvailabilityResponseMessageType, GroupAttendeeConflictData,
    FreeBusyViewType, ReplyBody, IndividualAttendeeConflictData, TooBigGroupAttendeeConflictData,
    UnknownAttendeeConflictData, ArrayOfAttendeeConflictData, GetUserConfigurationResponseMessageType,
    GetUserConfigurationType, GetUserConfigurationResponseType, CreateUserConfigurationType, UpdateUserConfigurationType,
    DeleteUserConfigurationType, CreateUserConfigurationResponseType, CreateUserConfigurationResponseMessage,
    UpdateUserConfigurationResponseType, UpdateUserConfigurationResponseMessage, DeleteUserConfigurationResponseMessage,
    DeleteUserConfigurationResponseType, UserConfigurationType, GetUserOofSettingsRequest, UserOofSettings, 
    Duration, SetUserOofSettingsRequest, GetUserPhotoType } from '../models/user.model';
import { UserConfigurationNameType } from '../models/mail.model';
import { getLabelByTarget, Logger, parseSizeRequestedValue, throwErrorIfClientResponse } from '../utils';
import { UserContext } from '../keepcomponent';
import { Request } from '@loopback/rest';
import { 
    KeepPimCalendarManager, OOOState, OOOExternalAudience, PimOOO, PimLabel, 
    KeepPimManager} from '@hcllabs/openclientkeepcomponent';


// mock reponse for SetUserPhoto
export const mockSetUserPhotoResponse: SetUserPhotoResponseMessageType = new SetUserPhotoResponseMessageType();

// mock reponse for GetUserOofSettings
export const mockGetUserOofSettingsResponse: GetUserOofSettingsResponseMessageType = {
    ResponseMessage: new ResponseMessageType(),
    AllowExternalOof: ExternalAudience.ALL,
    OofSettings: {
        OofState: OofState.DISABLED,
        ExternalAudience: ExternalAudience.ALL,
        Duration: {
            StartTime: new Date('2006-11-03T23:00:00'),
            EndTime: new Date('2006-11-04T23:00:00'),
        },
        InternalReply: new ReplyBody('<body>I am out of office. This is my internal reply.</body>'),
        ExternalReply: new ReplyBody('<body>大家好</body>', 'zh'),
    }
};

// mock reponse for SetUserOofSettings
export const mockGetUserAvailabilityResponse: GetUserAvailabilityResponseMessageType = {
    SuggestionsResponse: {
        SuggestionDayResultArray: {
            SuggestionDayResult: [{
                Date: new Date(),
                DayQuality: SuggestionQuality.EXCELLENT,
                SuggestionArray: {
                    Suggestion: [{
                        SuggestionQuality: SuggestionQuality.GOOD,
                        IsWorkTime: true,
                        MeetingTime: new Date(),
                        AttendeeConflictDataArray: new ArrayOfAttendeeConflictData([
                            new GroupAttendeeConflictData(3, 2, 1, 4),
                            new IndividualAttendeeConflictData(LegacyFreeBusyType.WORKING_ELSEWHERE),
                            new TooBigGroupAttendeeConflictData(),
                            new GroupAttendeeConflictData(2, 1, 3, 4),
                            new UnknownAttendeeConflictData(),
                        ])
                    }]
                }
            }]
        }
    },
    FreeBusyResponseArray: {
        FreeBusyResponse: [
            {
                ResponseMessage: new ResponseMessageType(),
                FreeBusyView: {
                    FreeBusyViewType: new FreeBusyViewType([FreeBusyViewEnum.DETAILED_MERGED]),
                    MergedFreeBusy: '000002220220000000000000',
                    CalendarEventArray: {
                        CalendarEvent: [
                            {
                                StartTime: new Date('2006-10-16T06:00:00-07:00'),
                                EndTime: new Date('2006-10-16T06:30:00-07:00'),
                                BusyType: LegacyFreeBusyType.BUSY,
                                CalendarEventDetails: {
                                    ID: '14B6414B0',
                                    Subject: 'Meet with Contoso Account Executives',
                                    IsMeeting: false,
                                    IsRecurring: false,
                                    IsException: false,
                                    IsReminderSet: false,
                                    IsPrivate: false
                                }
                            },
                            {
                                StartTime: new Date('2006-10-16T07:00:00-07:00'),
                                EndTime: new Date('2006-10-16T08:00:00-07:00'),
                                BusyType: LegacyFreeBusyType.BUSY,
                                CalendarEventDetails: {
                                    ID: 'E14B6414B0B',
                                    Subject: 'Pick up my groceries',
                                    IsMeeting: false,
                                    IsRecurring: false,
                                    IsException: false,
                                    IsReminderSet: false,
                                    IsPrivate: false
                                }
                            },
                            {
                                StartTime: new Date('2006-10-16T09:40:00-07:00'),
                                EndTime: new Date('2006-10-16T10:10:00-07:00'),
                                BusyType: LegacyFreeBusyType.BUSY,
                                CalendarEventDetails: {
                                    ID: '14B6414B0B1',
                                    Subject: 'Meet with doctor',
                                    Location: 'Kirkland',
                                    IsMeeting: false,
                                    IsRecurring: false,
                                    IsException: false,
                                    IsReminderSet: false,
                                    IsPrivate: false
                                }
                            }
                        ]
                    },
                    WorkingHours: {
                        TimeZone: {
                            Bias: 480,
                            StandardTime: {
                                Bias: 0,
                                Time: '02:00:00',
                                DayOrder: 5,
                                Month: 10,
                                DayOfWeek: DayOfWeekType.SUNDAY
                            },
                            DaylightTime: {
                                Bias: -60,
                                Time: '02:00:00',
                                DayOrder: 1,
                                Month: 4,
                                DayOfWeek: DayOfWeekType.SUNDAY
                            }
                        },
                        WorkingPeriodArray: {
                            WorkingPeriod: [{
                                DayOfWeek: new DaysOfWeekType([DayOfWeekType.MONDAY, DayOfWeekType.TUESDAY, DayOfWeekType.WEDNESDAY, DayOfWeekType.THURSDAY, DayOfWeekType.FRIDAY]),
                                StartTimeInMinutes: 480,
                                EndTimeInMinutes: 1020
                            }]
                        }
                    }
                }
            },
            {
                ResponseMessage: new ResponseMessageType(),
                FreeBusyView: {
                    FreeBusyViewType: new FreeBusyViewType([FreeBusyViewEnum.FREE_BUSY_MERGED, FreeBusyViewEnum.DETAILED_MERGED]),
                    MergedFreeBusy: '000000001100000000000000',
                    CalendarEventArray: {
                        CalendarEvent: [{
                            StartTime: new Date('2006-10-16T09:00:00-07:00'),
                            EndTime: new Date('2006-10-16T10:00:00-07:00'),
                            BusyType: LegacyFreeBusyType.TENTATIVE
                        }]
                    },
                    WorkingHours: {
                        TimeZone: {
                            Bias: 480,
                            StandardTime: {
                                Bias: 0,
                                Time: '02:00:00',
                                DayOrder: 5,
                                Month: 10,
                                DayOfWeek: DayOfWeekType.SUNDAY
                            },
                            DaylightTime: {
                                Bias: -60,
                                Time: '02:00:00',
                                DayOrder: 1,
                                Month: 4,
                                DayOfWeek: DayOfWeekType.SUNDAY
                            }
                        },
                        WorkingPeriodArray: {
                            WorkingPeriod: [{
                                DayOfWeek: new DaysOfWeekType([DayOfWeekType.MONDAY, DayOfWeekType.TUESDAY, DayOfWeekType.WEDNESDAY, DayOfWeekType.THURSDAY, DayOfWeekType.FRIDAY]),
                                StartTimeInMinutes: 480,
                                EndTimeInMinutes: 1020
                            }]
                        }
                    }
                }
            }
        ]
    }
};


/**
 * Get the default category list for a exchange calendar.
 * @returns A base64 encoded string containing the default category list xml. 
 */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function getDefaultCalendarCategoryList(): string {
    // This is the default category list for a calendar.
    const xmlconf = Buffer.from(`
     <?xml version="1.0" encoding="utf-8"?>
<categories lastSavedTime="2020-03-18T18:58:51.6978972Z"
 xmlns="CategoryList.xsd">
 <category renameOnFirstUse="1" name="Red category" color="0" keyboardShortcut="0" lastTimeUsedNotes="2020-03-18T18:58:51.6669147Z" lastTimeUsedJournal="2020-03-18T18:58:51.6669147Z" lastTimeUsedContacts="2020-03-18T18:58:51.6669147Z" lastTimeUsedTasks="2020-03-18T18:58:51.6669147Z" lastTimeUsedCalendar="2020-03-18T18:58:51.6669147Z" lastTimeUsedMail="2020-03-18T18:58:51.6669147Z" lastTimeUsed="2020-03-18T18:58:51.6669147Z" guid="{0de79840-6a0a-46de-955a-7533ccef5e5b}" />
 <category renameOnFirstUse="1" name="Orange category" color="1" keyboardShortcut="0" lastTimeUsedNotes="2020-03-18T18:58:51.6669147Z" lastTimeUsedJournal="2020-03-18T18:58:51.6669147Z" lastTimeUsedContacts="2020-03-18T18:58:51.6669147Z" lastTimeUsedTasks="2020-03-18T18:58:51.6669147Z" lastTimeUsedCalendar="2020-03-18T18:58:51.6669147Z" lastTimeUsedMail="2020-03-18T18:58:51.6669147Z" lastTimeUsed="2020-03-18T18:58:51.6669147Z" guid="{3b87a78d-3f4a-4990-93a0-262acad1875f}" />
 <category renameOnFirstUse="1" name="Yellow category" color="3" keyboardShortcut="0" lastTimeUsedNotes="2020-03-18T18:58:51.6669147Z" lastTimeUsedJournal="2020-03-18T18:58:51.6669147Z" lastTimeUsedContacts="2020-03-18T18:58:51.6669147Z" lastTimeUsedTasks="2020-03-18T18:58:51.6669147Z" lastTimeUsedCalendar="2020-03-18T18:58:51.6669147Z" lastTimeUsedMail="2020-03-18T18:58:51.6669147Z" lastTimeUsed="2020-03-18T18:58:51.6669147Z" guid="{4f2cc5c7-e6ba-4f37-ac14-a27ce5fd53d7}" />
 <category renameOnFirstUse="1" name="Green category" color="4" keyboardShortcut="0" lastTimeUsedNotes="2020-03-18T18:58:51.6669147Z" lastTimeUsedJournal="2020-03-18T18:58:51.6669147Z" lastTimeUsedContacts="2020-03-18T18:58:51.6669147Z" lastTimeUsedTasks="2020-03-18T18:58:51.6669147Z" lastTimeUsedCalendar="2020-03-18T18:58:51.6669147Z" lastTimeUsedMail="2020-03-18T18:58:51.6669147Z" lastTimeUsed="2020-03-18T18:58:51.6669147Z" guid="{41053fd7-9358-46be-9bb0-6213033c5934}" />
 <category renameOnFirstUse="1" name="Blue category" color="7" keyboardShortcut="0" lastTimeUsedNotes="2020-03-18T18:58:51.6669147Z" lastTimeUsedJournal="2020-03-18T18:58:51.6669147Z" lastTimeUsedContacts="2020-03-18T18:58:51.6669147Z" lastTimeUsedTasks="2020-03-18T18:58:51.6669147Z" lastTimeUsedCalendar="2020-03-18T18:58:51.6669147Z" lastTimeUsedMail="2020-03-18T18:58:51.6669147Z" lastTimeUsed="2020-03-18T18:58:51.6669147Z" guid="{5fb8d81f-58cf-4e79-9ed2-9638d615db4b}" />
 <category renameOnFirstUse="1" name="Purple category" color="8" keyboardShortcut="0" lastTimeUsedNotes="2020-03-18T18:58:51.6669147Z" lastTimeUsedJournal="2020-03-18T18:58:51.6669147Z" lastTimeUsedContacts="2020-03-18T18:58:51.6669147Z" lastTimeUsedTasks="2020-03-18T18:58:51.6669147Z" lastTimeUsedCalendar="2020-03-18T18:58:51.6669147Z" lastTimeUsedMail="2020-03-18T18:58:51.6669147Z" lastTimeUsed="2020-03-18T18:58:51.6669147Z" guid="{71be6f11-f48f-4c00-840a-297f8af5e78e}" />
</categories>
     `);

    return xmlconf.toString('base64');
}

/**
 * Get user OOO settings.
 * @param soapRequest The soap request, which is type of GetUserOofSettingsRequest
 * @param request The request being processed
 * @returns User OOF settings
 */
export async function GetUserOofSettings(soapRequest: GetUserOofSettingsRequest, request: Request): Promise<GetUserOofSettingsResponseMessageType> {
    const userInfo = UserContext.getUserInfo(request);

    const response = new GetUserOofSettingsResponseMessageType();
    response.ResponseMessage = new ResponseMessageType();

    if (soapRequest.Mailbox.Address !== userInfo.userId) { // Only return response for current user. TODO: There is a bug in Keep if you ask for OOO data for anyone other than the authenticated user (LABS-1418). 
        response.ResponseMessage.ResponseClass = ResponseClassType.ERROR;
        response.ResponseMessage.ResponseCode = ResponseCodeType.ERROR_ACCESS_DENIED;
        response.ResponseMessage.MessageText = `The authenticated user does not have access to the settings for ${soapRequest.Mailbox.Address}`;
        return response;
    }
    else {
        try {
            const settings = await KeepPimCalendarManager.getInstance().getOOO(userInfo.userId, userInfo);

            response.OofSettings = new UserOofSettings();
            if (settings) {
                switch (settings.externalAudience) {
                    case OOOExternalAudience.ALL:
                        response.AllowExternalOof = ExternalAudience.ALL;
                        response.OofSettings.ExternalAudience = ExternalAudience.ALL;
                        break;
                    case OOOExternalAudience.KNOWN:
                        response.AllowExternalOof = ExternalAudience.KNOWN;
                        response.OofSettings.ExternalAudience = ExternalAudience.KNOWN;
                        break;
                    default:
                        response.AllowExternalOof = ExternalAudience.NONE;
                        response.OofSettings.ExternalAudience = ExternalAudience.NONE;
                        break;
                }

                switch (settings.state) {
                    case OOOState.ENABLED:
                        response.OofSettings.OofState = OofState.ENABLED;
                        break;
                    case OOOState.SCHEDULED:
                        response.OofSettings.OofState = OofState.SCHEDULED;
                        break;
                    default:
                        response.OofSettings.OofState = OofState.DISABLED;
                        break;
                }

                if (settings.startDate && settings.endDate) {
                    response.OofSettings.Duration = new Duration();
                    response.OofSettings.Duration.StartTime = settings.startDate;
                    response.OofSettings.Duration.EndTime = settings.endDate;
                }

                if (settings.replyMessage) {
                    response.OofSettings.ExternalReply = new ReplyBody(settings.replyMessage);
                    response.OofSettings.InternalReply = new ReplyBody(settings.replyMessage);
                }
            }
            else {
                response.OofSettings.OofState = OofState.DISABLED;
                response.OofSettings.ExternalAudience = ExternalAudience.NONE;
            }

        } catch (err) {
            throwErrorIfClientResponse(err);
            response.ResponseMessage.ResponseClass = ResponseClassType.ERROR;
            response.ResponseMessage.ResponseCode = ResponseCodeType.ERROR_SERVICE_UNAVAILABLE;
            response.ResponseMessage.MessageText = `Failed to retrieve user out-of-office settings`;
            response.ResponseMessage.MessageXml = { Value: new Value("Keep Error", `${err}`) };
        }
    }

    return response;

}

/**
 * Set the OOO settings for a user.
 * @param soapRequest The soap request, which is type of SetUserOofSettingsRequest
 * @param request The request being processed
 * @returns User OOF settings
 */
export async function SetUserOofSettings(soapRequest: SetUserOofSettingsRequest, request: Request): Promise<SetUserOofSettingsResponseMessageType> {
    const userInfo = UserContext.getUserInfo(request);

    const response = new SetUserOofSettingsResponseMessageType();
    response.ResponseMessage = new ResponseMessageType();

    if (soapRequest.Mailbox.Address !== userInfo.userId) { // Only return response for current user. TODO: There is a bug in Keep if you access OOO data for anyone other than the authenticated user (LABS-1418). 
        response.ResponseMessage.ResponseClass = ResponseClassType.ERROR;
        response.ResponseMessage.ResponseCode = ResponseCodeType.ERROR_ACCESS_DENIED;
        response.ResponseMessage.MessageText = `The authenticated user does not have access to the settings for ${soapRequest.Mailbox.Address}`;
        return response;
    }
    else {
        try {
            const settings: any = {};

            if (soapRequest.UserOofSettings.OofState !== OofState.DISABLED) {
                settings.Enabled = true;
                settings.SystemState = true;
            }
            else {
                settings.Enabled = false;
                settings.SystemState = false;
            }

            if (soapRequest.UserOofSettings.Duration) {
                settings.StartDateTime = soapRequest.UserOofSettings.Duration.StartTime.toISOString();
                settings.EndDateTime = soapRequest.UserOofSettings.Duration.EndTime.toISOString();
            }

            settings.ExcludeInternet = (soapRequest.UserOofSettings.ExternalAudience !== ExternalAudience.ALL);

            if (soapRequest.UserOofSettings.InternalReply) {
                settings.InternalReply = soapRequest.UserOofSettings.InternalReply.Message;
            }

            if (soapRequest.UserOofSettings.ExternalReply) {
                settings.ExternalReply = soapRequest.UserOofSettings.ExternalReply.Message;
            }

            await KeepPimCalendarManager.getInstance().updateOOO(new PimOOO(settings), userInfo);

        } catch (err) {
            throwErrorIfClientResponse(err);
            response.ResponseMessage.ResponseClass = ResponseClassType.ERROR;
            response.ResponseMessage.ResponseCode = ResponseCodeType.ERROR_SERVICE_UNAVAILABLE;
            response.ResponseMessage.MessageText = `Failed to update user out-of-office settings`;
            response.ResponseMessage.MessageXml = { Value: new Value("Keep Error", `${err}`) };
        }
    }

    return response;

}
/**
 * GetUserConfiguration for item
 * @param soapRequest The GetUserConfiguration request
 * @param request The HTTP request
 * @returns A GetUserConfiguration response
 */
export async function GetUserConfigurationResponse(soapRequest: GetUserConfigurationType, request: Request): Promise<GetUserConfigurationResponseType> {
    const userInfo = UserContext.getUserInfo(request);

    const response = new GetUserConfigurationResponseType();
    const msg = new GetUserConfigurationResponseMessageType();
    response.ResponseMessages.push(msg);

    const targetUserConfig = soapRequest.UserConfigurationName;

    let toLabel: PimLabel | undefined;
    try {
        toLabel = await getLabelByTarget(targetUserConfig, userInfo);
    } catch (err) {
        if (err.status !== 404) {
            Logger.getInstance().debug(`An error occurred getting all label for GetUserConfigurationResponse: ${err}`);
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR;
            msg.MessageText = err.message;
            return response;
        }
    }

    if (!toLabel) {
        msg.ResponseClass = ResponseClassType.ERROR;
        msg.ResponseCode = ResponseCodeType.ERROR_FOLDER_NOT_FOUND;
        return response;
    }

    const additionalProperty = toLabel.getAdditionalProperty(targetUserConfig.Name);

    msg.UserConfiguration = getAdditionalUserConfiguration(soapRequest.UserConfigurationProperties[0], additionalProperty, targetUserConfig.Name);

    if (targetUserConfig.DistinguishedFolderId) {
        msg.UserConfiguration.UserConfigurationName.DistinguishedFolderId = targetUserConfig.DistinguishedFolderId;
    } else if (targetUserConfig.FolderId) {
        msg.UserConfiguration.UserConfigurationName.FolderId = targetUserConfig.FolderId;
    } else if (targetUserConfig.AddressListId) {
        msg.UserConfiguration.UserConfigurationName.AddressListId = targetUserConfig.AddressListId;
    }

    msg.ResponseClass = ResponseClassType.SUCCESS;
    msg.ResponseCode = ResponseCodeType.NO_ERROR;

    return response;
}

function getAdditionalUserConfiguration(property: UserConfigurationPropertyTypeEnum, value: any, name: string ): UserConfigurationType {
    const userConfig = new UserConfigurationType();
    userConfig.UserConfigurationName = new UserConfigurationNameType();
    userConfig.UserConfigurationName.Name = name;

    if (property === UserConfigurationPropertyTypeEnum.BINARY_DATA || property === UserConfigurationPropertyTypeEnum.ALL) {
        userConfig.BinaryData = value?.[UserConfigurationPropertyTypeEnum.BINARY_DATA];
    } 
    if (property === UserConfigurationPropertyTypeEnum.DICTIONARY || property === UserConfigurationPropertyTypeEnum.ALL) {
        userConfig.Dictionary = value?.[UserConfigurationPropertyTypeEnum.DICTIONARY];
    } 
    if (property === UserConfigurationPropertyTypeEnum.XML_DATA || property === UserConfigurationPropertyTypeEnum.ALL) {
        userConfig.XmlData = value?.[UserConfigurationPropertyTypeEnum.XML_DATA];
    } 
    if (property === UserConfigurationPropertyTypeEnum.ID || property === UserConfigurationPropertyTypeEnum.ALL) {
        userConfig.ItemId = value?.[UserConfigurationPropertyTypeEnum.ID];
    }     
    return userConfig;
}


/**
 * CreateUserConfiguration for item 
 * @param soapRequest The CreateUserConfiguration request
 * @param request The HTTP request
 * @returns A CreateUSerConfiguration response
 */
export async function CreateUserConfigurationResponse(soapRequest: CreateUserConfigurationType, request: Request): Promise<CreateUserConfigurationResponseType> {
    const response = new CreateUserConfigurationResponseType();
    const msg = new CreateUserConfigurationResponseMessage();
    response.ResponseMessages.push(msg);

    const userInfo = UserContext.getUserInfo(request);

    const targetUserConfiguration = soapRequest.UserConfiguration.UserConfigurationName;

    let toLabel: PimLabel | undefined;
    try {
        toLabel = await getLabelByTarget(targetUserConfiguration, userInfo);
    } catch (err) {
        if (err.status !== 404) {
            Logger.getInstance().debug(`An error occurred getting all label for CreateUserConfigurationResponse: ${err}`);
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR;
            msg.MessageText = err.message;
            return response;
        }
    }

    if (!toLabel) {
        msg.ResponseClass = ResponseClassType.ERROR;
        msg.ResponseCode = ResponseCodeType.ERROR_FOLDER_NOT_FOUND;
        return response;
    }

    const userConfiguration = getUserConfigurationFromRequest(soapRequest.UserConfiguration);
    toLabel.setAdditionalProperty(targetUserConfiguration.Name, userConfiguration);

    try {
        await KeepPimManager.getInstance().updatePimItem(toLabel.unid, userInfo, toLabel);
    } catch (err) {
        Logger.getInstance().debug(`An error occurred updating label for GetUserConfigurationResponse: ${err}`);
        msg.ResponseClass = ResponseClassType.ERROR;
        msg.ResponseCode = ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR;
        msg.MessageText = err.message;
        return response;
    }

    msg.ResponseClass = ResponseClassType.SUCCESS;
    msg.ResponseCode = ResponseCodeType.NO_ERROR;

    return response;
}

function getUserConfigurationFromRequest (configuration: UserConfigurationType): any {
    const userConfig: any = {};

    if (configuration.BinaryData) {
        userConfig.BinaryData = configuration.BinaryData;
    } 
    if (configuration.Dictionary) {
        userConfig.Dictionary = configuration.Dictionary;
    } 
    if (configuration.ItemId) {
        userConfig.ItemId = configuration.ItemId;
    } 
    if (configuration.XmlData) {
        userConfig.XmlData = configuration.XmlData;
    }

    return userConfig;
}

/**
 * UpdateUserConfiguration for item
 * @param soapRequest The UpdateUserConfiguration request
 * @param request The HTTP request
 * @returns An UpdateUserConfiguration response
 */
export async function UpdateUserConfigurationResponse(soapRequest: UpdateUserConfigurationType, request: Request): Promise<UpdateUserConfigurationResponseType> {
    const response = new UpdateUserConfigurationResponseType();
    const msg = new UpdateUserConfigurationResponseMessage();
    response.ResponseMessages.push(msg);

    const userInfo = UserContext.getUserInfo(request);

    const targetUserConfiguration = soapRequest.UserConfiguration.UserConfigurationName;

    let toLabel: PimLabel | undefined;
    try {
        toLabel = await getLabelByTarget(targetUserConfiguration, userInfo);
    } catch (err) {
        if (err.status !== 404) {
            Logger.getInstance().debug(`An error occurred getting all label for UpdateUserConfigurationResponse: ${err}`);
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR;
            msg.MessageText = err.message;
            return response;
        }
    }

    if (!toLabel) {
        msg.ResponseClass = ResponseClassType.ERROR;
        msg.ResponseCode = ResponseCodeType.ERROR_FOLDER_NOT_FOUND;
        return response;
    }

    //if user config with targetUserConfiguration.Name not exist add a new config from update
    const userConfiguration = getUserConfigurationFromRequest(soapRequest.UserConfiguration);
    toLabel.setAdditionalProperty(targetUserConfiguration.Name, userConfiguration);

    try {
        await KeepPimManager.getInstance().updatePimItem(toLabel.unid, userInfo, toLabel);
    } catch (err) {
        Logger.getInstance().debug(`An error occurred updating label for GetUserConfigurationResponse: ${err}`);
        msg.ResponseClass = ResponseClassType.ERROR;
        msg.ResponseCode = ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR;
        msg.MessageText = err.message;
        return response;
    }

    msg.ResponseClass = ResponseClassType.SUCCESS;
    msg.ResponseCode = ResponseCodeType.NO_ERROR;

    return response;
}

/**
 * DeleteUserConfiguration for item
 * @param soapRequest The DeleteUserConfiguration request
 * @param request The HTTP request
 * @returns A DeleteUserConfiguration response
 */
export async function DeleteUserConfigurationResponse(soapRequest: DeleteUserConfigurationType, request: Request): Promise<DeleteUserConfigurationResponseType> {
    const response = new DeleteUserConfigurationResponseType();
    const msg = new DeleteUserConfigurationResponseMessage();
    response.ResponseMessages.push(msg);

    const userInfo = UserContext.getUserInfo(request);

    const targetUserConfiguration = soapRequest.UserConfigurationName;

    let toLabel: PimLabel | undefined;
    try {
        toLabel = await getLabelByTarget(targetUserConfiguration, userInfo);
    } catch (err) {
        if (err.status !== 404) {
            Logger.getInstance().debug(`An error occurred getting all label for DeleteUserConfigurationResponse: ${err}`);
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR;
            msg.MessageText = err.message;
            return response;
        }
    }
    
    if (!toLabel) {
        msg.ResponseClass = ResponseClassType.ERROR;
        msg.ResponseCode = ResponseCodeType.ERROR_FOLDER_NOT_FOUND;
        return response;
    }
    const additionalProperty = toLabel.getAdditionalProperty(targetUserConfiguration.Name);

    //if user config with targetUserConfiguration.Name not exist return success
    if (!additionalProperty) {
        msg.ResponseClass = ResponseClassType.SUCCESS;
        msg.ResponseCode = ResponseCodeType.NO_ERROR;
        return response;
    }
    
    toLabel.deleteAdditionalProperty(targetUserConfiguration.Name);

    try {
        await KeepPimManager.getInstance().updatePimItem(toLabel.unid, userInfo, toLabel);
    } catch (err) {
        Logger.getInstance().debug(`An error occurred updating label for DeleteUserConfigurationResponse: ${err}`);
        msg.ResponseClass = ResponseClassType.ERROR;
        msg.ResponseCode = ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR;
        msg.MessageText = err.message;
        return response;
    }

    msg.ResponseClass = ResponseClassType.SUCCESS;
    msg.ResponseCode = ResponseCodeType.NO_ERROR;

    return response;
}

/**
 * GetUserPhoto 
 * @param soapRequest The GetUserPhoto request
 * @param request The HTTP request
 * @returns The GetUserPhoto response
 */
 export async function GetUserPhotoResponse(soapRequest: GetUserPhotoType, request: Request): Promise<GetUserPhotoResponseMessageType> {
    const userInfo = UserContext.getUserInfo(request);

    const response = new GetUserPhotoResponseMessageType();
    const [height, width] = parseSizeRequestedValue(soapRequest.SizeRequested);

    try {
        const userPhoto = await KeepPimManager.getInstance().getAvatar(userInfo, soapRequest.Email, height, width);
        response.ResponseClass = ResponseClassType.SUCCESS;
        response.ResponseCode = ResponseCodeType.NO_ERROR;
        response.HasChanged = false;
        response.PictureData = Buffer.from(userPhoto).toString('base64');
        return response;
    } catch (err) {
        Logger.getInstance().debug(`Error occurred retrieving the user photo for email ${soapRequest.Email}: ${err}`);
        throwErrorIfClientResponse(err);
        response.ResponseClass = ResponseClassType.ERROR;
        response.ResponseCode = ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR;
        response.MessageText = `Error occurred retrieving the user photo for email: ${soapRequest.Email}`;
        return response;
    }
}

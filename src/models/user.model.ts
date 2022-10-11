/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import { ResponseMessageType, XSList, AttributesKey, BaseRequestType, GenericArray, DaysOfWeekType, BaseResponseMessageType } from './common.model';
import {
    FreeBusyViewEnum, LegacyFreeBusyType, DayOfWeekType, SuggestionQuality, ExternalAudience,
    OofState, UserPhotoSizeType, UserPhotoTypeType, MeetingAttendeeType, UserConfigurationDictionaryObjectTypesType, UserConfigurationPropertyTypeEnum
} from './enum.model';
import { UserConfigurationNameType, ItemIdType } from './mail.model';

/**
 * Contains the response to a GetUserPhoto request.
 */
export class GetUserPhotoResponseMessageType extends ResponseMessageType {
    HasChanged: boolean;
    PictureData?: string; // base64 string
}

/**
 * Contains the response to a SetUserPhoto request.
 */
export class SetUserPhotoResponseMessageType extends ResponseMessageType { }

/**
 * Contains the response to a GetUserOofSettings request.
 */
export class GetUserOofSettingsResponseMessageType {
    ResponseMessage: ResponseMessageType;
    OofSettings?: UserOofSettings;
    AllowExternalOof?: ExternalAudience;
}

/**
 * Contains the response to a SetUserOofSettings request.
 */
export class SetUserOofSettingsResponseMessageType {
    ResponseMessage?: ResponseMessageType;
}

/**
 * Contains the response to a GetUserAvailability request.
 */
export class GetUserAvailabilityResponseMessageType {
    FreeBusyResponseArray?: ArrayOfFreeBusyResponse;
    SuggestionsResponse?: SuggestionsResponseType;
}

/**
 * Contains the requested users' availability information and the response status.
 */
export class ArrayOfFreeBusyResponse {
    FreeBusyResponse?: FreeBusyResponseType[];
}

/**
 * Contains the free/busy information for a single mailbox user and the response status.
 */
export class FreeBusyResponseType {
    ResponseMessage?: ResponseMessageType;
    FreeBusyView?: FreeBusyView;
}

/**
 * Represents list of free busy view types.
 */
export class FreeBusyViewType extends XSList<FreeBusyViewEnum> { }

/**
 * Contains availability information for a specific user.
 */
export class FreeBusyView {
    FreeBusyViewType: FreeBusyViewType = new FreeBusyViewType();
    MergedFreeBusy?: string;
    CalendarEventArray?: ArrayOfCalendarEvent;
    WorkingHours?: WorkingHours;
}

/**
 * Contains a set of unique calendar item occurrences that represent the requested user's availability.
 */
export class ArrayOfCalendarEvent {
    CalendarEvent?: CalendarEvent[];
}

/**
 * Represents a unique calendar item occurrence.
 */
export class CalendarEvent {
    StartTime: Date;
    EndTime: Date;
    BusyType: LegacyFreeBusyType;
    CalendarEventDetails?: CalendarEventDetails;
}

/**
 * Provides additional information about a calendar event.
 */
export class CalendarEventDetails {
    ID?: string;
    Subject?: string;
    Location?: string;
    IsMeeting: boolean;
    IsRecurring: boolean;
    IsException: boolean;
    IsReminderSet: boolean;
    IsPrivate: boolean;
}

/**
 * Represents the time zone settings and working hours for the requested mailbox user.
 */
export class WorkingHours {
    TimeZone: SerializableTimeZone;
    WorkingPeriodArray: ArrayOfWorkingPeriod;
}

/**
 * Contains elements that identify time zone information.
 */
export class SerializableTimeZone {
    Bias: number;
    StandardTime: SerializableTimeZoneTime;
    DaylightTime: SerializableTimeZoneTime;
}

/**
 * Contains working period information for the mailbox user.
 */
export class ArrayOfWorkingPeriod {
    WorkingPeriod?: WorkingPeriod[];
}

/**
 * Represents an offset from the time relative to Coordinated Universal Time (UTC)
 * that is represented by the Bias (UTC) element.
 */
export class SerializableTimeZoneTime {
    Bias: number;
    Time: string;
    DayOrder: number;
    Month: number;
    DayOfWeek: DayOfWeekType;
    Year?: string;
}

/**
 * Contains working period information for the mailbox user.
 */
export class WorkingPeriod {
    DayOfWeek: DaysOfWeekType = new DaysOfWeekType();
    StartTimeInMinutes: number;
    EndTimeInMinutes: number;
}

/**
 * Contains response status information and suggestion data for requested meeting suggestions.
 */
export class SuggestionsResponseType {
    ResponseMessage?: ResponseMessageType;
    SuggestionDayResultArray?: ArrayOfSuggestionDayResult;
}

/**
 * Contains an array of meeting suggestions organized by date.
 */
export class ArrayOfSuggestionDayResult {
    SuggestionDayResult?: SuggestionDayResult[];
}

/**
 * Represents a single day that contains suggested meeting times.
 */
export class SuggestionDayResult {
    Date: Date;
    DayQuality: SuggestionQuality;
    SuggestionArray?: ArrayOfSuggestion;
}

/**
 * Contains an array of meeting suggestions.
 */
export class ArrayOfSuggestion {
    Suggestion?: Suggestion[];
}

/**
 * Represents a single meeting suggestion.
 */
export class Suggestion {
    MeetingTime: Date;
    IsWorkTime: boolean;
    SuggestionQuality: SuggestionQuality;
    AttendeeConflictDataArray?: ArrayOfAttendeeConflictData;
}

/**
 * Contains an array of conflict data for queried attendees.
 */
export class ArrayOfAttendeeConflictData extends GenericArray<AttendeeConflictData> { }

/**
 * Base class of attendee conflict data.
 */
export abstract class AttendeeConflictData { }

/**
 * Represents an unresolvable attendee or an attendee that is not a user, distribution list, or contact.
 */
export class UnknownAttendeeConflictData extends AttendeeConflictData { }

/**
 * Represents an attendee that resolved as a distribution list that was too large to expand.
 */
export class TooBigGroupAttendeeConflictData extends AttendeeConflictData { }

/**
 * Contains a user's or contact's free/busy status for a time window that occurs at the same time as the
 * suggested meeting time identified in the Suggestion element.
 */
export class IndividualAttendeeConflictData extends AttendeeConflictData {
    constructor(public BusyType: LegacyFreeBusyType) {
        super();
    }
}

/**
 * Contains aggregate conflict information about the number of users available, the number of users who have conflicts,
 * and the number of users who do not have availability information in a distribution list for a suggested meeting time.
 */
export class GroupAttendeeConflictData extends AttendeeConflictData {
    constructor(public NumberOfMembers: number, public NumberOfMembersAvailable: number, public NumberOfMembersWithConflict: number, public NumberOfMembersWithNoData: number) {
        super();
    }
}

/**
 * Contains the Out of Office (OOF) settings.
 */
export class UserOofSettings {
    OofState: OofState;
    ExternalAudience: ExternalAudience;
    Duration?: Duration;
    InternalReply?: ReplyBody;
    ExternalReply?: ReplyBody;
    DeclineMeetingReply?: ReplyBody;
    DeclineEventsForScheduledOOF?: boolean;
    DeclineAllEventsForScheduledOOF?: boolean;
    CreateOOFEvent?: boolean;
    AutoDeclineFutureRequestsWhenOOF?: boolean;
    EventsToDeleteIDs?: ArrayOfEventIDType;
    OOFEventSubject?: string;
    OOFEventID?: string;
}

/**
 * Contains the duration for which the OOF status is enabled if the OofState element is set to Scheduled.
 * If the OofState element is set to Enabled or Disabled, the value of this element is ignored.
 */
export class Duration {
    StartTime: Date;
    EndTime: Date;
}

/**
 * Contains the OOF response sent to other users.
 */
const langKey = 'xml:lang';
export class ReplyBody {
    [AttributesKey]: {
        [langKey]?: string;
    } = {};
    Message: string;

    constructor(_message: string, _lang?: string) {
        if (_message !== undefined) {
            this.Message = _message;
        }
        if (_lang) {
            this.lang = _lang;
        }
    }
    get lang(): string {
        return this[AttributesKey][langKey] as string;
    }
    set lang(_lang: string) {
        this[AttributesKey][langKey] = _lang;
    }
}

/**
 * Array of event id types.
 */
export class ArrayOfEventIDType {
    EventToDeleteID: string[] = [];
}


//// Request models

/**
 * Contains the request to get a user's photo.
 */
export class GetUserPhotoType extends BaseRequestType {
    Email: string;
    SizeRequested: UserPhotoSizeType;
    TypeRequested?: UserPhotoTypeType;
}

/**
 * Contains the request to set a user's photo.
 */
export class SetUserPhotoType extends BaseRequestType {
    Email: string;
    Content: string; // base64 string
    TypeRequested?: UserPhotoTypeType;
}

/**
 * Contains the request to get a user's OOF settings.
 */
export class GetUserOofSettingsRequest extends BaseRequestType {
    Mailbox: EmailAddress;
}

/**
 * Contains the request to set a user's OOF settings.
 */
export class SetUserOofSettingsRequest extends BaseRequestType {
    Mailbox: EmailAddress;
    UserOofSettings: UserOofSettings;
}

/**
 * Contains the request to get a user's availability.
 */
export class GetUserAvailabilityRequestType extends BaseRequestType {
    TimeZone?: SerializableTimeZone;
    MailboxDataArray: ArrayOfMailboxData;
    FreeBusyViewOptions?: FreeBusyViewOptionsType;
    SuggestionsViewOptions?: SuggestionsViewOptionsType;
}

/**
 * Contains a list of mailboxes to query for availability information.
 */
export class ArrayOfMailboxData {
    MailboxData: MailboxData[] = [];
}

/**
 * Represents an individual mailbox user and options for the type of data to be returned about the mailbox user.
 */
export class MailboxData {
    Email: EmailAddress;
    AttendeeType: MeetingAttendeeType;
    ExcludeConflicts?: boolean;
}

/**
 * Represents the mailbox user for a GetUserAvailability query.
 */
export class EmailAddress {
    Name?: string;
    Address: string;
    RoutingType?: string;
}

/**
 * Specifies the type of free/busy information returned in the response.
 */
export class FreeBusyViewOptionsType {
    TimeWindow: Duration;
    MergedFreeBusyIntervalInMinutes?: number;
    RequestedView = new FreeBusyViewType();
}

/**
 * Contains the options for obtaining meeting suggestion information.
 */
export class SuggestionsViewOptionsType {
    GoodThreshold?: number;
    MaximumResultsByDay?: number;
    MaximumNonWorkHourResultsByDay?: number;
    MeetingDurationInMinutes?: number;
    MinimumSuggestionQuality?: SuggestionQuality;
    DetailedSuggestionsWindow: Duration;
    CurrentMeetingTime?: Date;
    GlobalObjectId?: string;
}

/**
 * Contains the request to get user configuration.
 */
export class GetUserConfigurationType extends BaseRequestType {
    UserConfigurationName: UserConfigurationNameType;
    UserConfigurationProperties: UserConfigurationPropertyType = new UserConfigurationPropertyType();
}

/**
 * Contains the response of create user configuration.
 */
export class GetUserConfigurationResponseType extends BaseResponseMessageType<GetUserConfigurationResponseMessageType> { }

/**
 * Contains the response message of get user configuration.
 */
export class GetUserConfigurationResponseMessageType extends ResponseMessageType {
    UserConfiguration?: UserConfigurationType;
}

/**
 * Contains the request of create user configuration
 */
export class CreateUserConfigurationType extends BaseRequestType { 
    UserConfiguration: UserConfigurationType;
}

/**
 * Contains the response of create user configuation.
 */
export class CreateUserConfigurationResponseType extends BaseResponseMessageType<CreateUserConfigurationResponseMessage> { }

/**
 * Contains the response message of create user configuration
 */
export class CreateUserConfigurationResponseMessage extends ResponseMessageType { }

/**
 * Contains the request of modify user configuration
 */
export class UpdateUserConfigurationType extends BaseRequestType {
    UserConfiguration: UserConfigurationType;
}

/**
 * Contains the response of modify user configuration
 */
export class UpdateUserConfigurationResponseType extends BaseResponseMessageType<UpdateUserConfigurationResponseMessage> { }

/**
 * Contains the response message of modify user configuration
 */
export class UpdateUserConfigurationResponseMessage extends ResponseMessageType { }

/**
 * Contains the request of delete user configuration
 */
export class DeleteUserConfigurationType extends BaseRequestType {
    UserConfigurationName: UserConfigurationNameType;
}

/**
 * Contains the response of delete user configuration
 */
export class DeleteUserConfigurationResponseType extends BaseResponseMessageType<DeleteUserConfigurationResponseMessage> { }

/**
 * Contains the response message for delete user configuation
 */
export class DeleteUserConfigurationResponseMessage extends ResponseMessageType { }

/**
 * User configuration property type.
 */
export class UserConfigurationPropertyType extends XSList<UserConfigurationPropertyTypeEnum> { }

/**
 * User configuration type.
 */
export class UserConfigurationType {
    UserConfigurationName: UserConfigurationNameType;
    ItemId?: ItemIdType;
    Dictionary?: UserConfigurationDictionaryType;
    XmlData?: string;
    BinaryData?: string;
}

/**
 * User configuration dictionary type.
 */
export class UserConfigurationDictionaryType {
    DictionaryEntry: UserConfigurationDictionaryEntryType[] = [];
}

/**
 * User configuration dictionary entry type.
 */
export class UserConfigurationDictionaryEntryType {
    DictionaryKey: UserConfigurationDictionaryObjectType;
    DictionaryValue: UserConfigurationDictionaryObjectType;
}

/**
 * User configuration dictionary object type.
 */
export class UserConfigurationDictionaryObjectType {
    Type: UserConfigurationDictionaryObjectTypesType;
    Value: string[] = [];
}

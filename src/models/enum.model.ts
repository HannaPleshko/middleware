/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */

/**
 * Exchange version.
 */
export const enum ExchangeVersion {
    EXCHANGE_2007 = 'Exchange2007',
    EXCHANGE_2007_SP_1 = 'Exchange2007_SP1',
    EXCHANGE_2009 = 'Exchange2009',
    EXCHANGE_2010 = 'Exchange2010',
    EXCHANGE_2010_SP_1 = 'Exchange2010_SP1',
    EXCHANGE_2010_SP_2 = 'Exchange2010_SP2',
    EXCHANGE_2012 = 'Exchange2012',
    EXCHANGE_2013 = 'Exchange2013',
    EXCHANGE_2013_SP_1 = 'Exchange2013_SP1',
    EXCHANGE_2015 = 'Exchange2015',
    EXCHANGE_2016 = 'Exchange2016',
    V_2015_10_05 = 'V2015_10_05',
    V_2016_01_06 = 'V2016_01_06',
    V_2016_04_13 = 'V2016_04_13',
    V_2016_07_13 = 'V2016_07_13',
    V_2016_10_10 = 'V2016_10_10',
    V_2017_01_07 = 'V2017_01_07',
    V_2017_04_14 = 'V2017_04_14',
    V_2017_07_11 = 'V2017_07_11',
    V_2017_10_09 = 'V2017_10_09',
    V_2018_01_08 = 'V2018_01_08',
}

/**
 * Error code for auto discover.
 */
export const enum AutodiscoverErrorCode {
    NO_ERROR = 'NoError',
    REDIRECT_ADDRESS = 'RedirectAddress',
    REDIRECT_URL = 'RedirectUrl',
    INVALID_USER = 'InvalidUser',
    INVALID_REQUEST = 'InvalidRequest',
    INVALID_SETTING = 'InvalidSetting',
    SETTING_IS_NOT_AVAILABLE = 'SettingIsNotAvailable',
    SERVER_BUSY = 'ServerBusy',
    INVALID_DOMAIN = 'InvalidDomain',
    NOT_FEDERATED = 'NotFederated',
    INTERNAL_SERVER_ERROR = 'InternalServerError',
}

/**
 * Auto discover setting.
 */
export const enum AutodiscoverSetting {
    USER_DISPLAY_NAME = 'UserDisplayName',
    EXTERNAL_EWS_URL = 'ExternalEwsUrl',
    EXTERNAL_OAB_URL = 'ExternalOABUrl',
    EWS_SUPPORTED_SCHEMAS = 'EwsSupportedSchemas',
    MAPI_HTTP_ENABLED = 'MapiHttpEnabled',
}

/**
 * Represents the status of the response.
 */
export const enum ResponseClassType {
    SUCCESS = 'Success',
    WARNING = 'Warning',
    ERROR = 'Error',
}

export const enum DistinguishedUserType {
    DEFAULT = 'Default',
    ANONYMOUS = 'Anonymous',
}

export const enum PermissionReadAccessType {
    NONE = 'None',
    FULL_DETAILS = 'FullDetails',
}

export const enum CalendarPermissionReadAccessType {
    NONE = 'None',
    TIME_ONLY = 'TimeOnly',
    TIME_AND_SUBJECT_AND_LOCATION = 'TimeAndSubjectAndLocation',
    FULL_DETAILS = 'FullDetails',
}

export const enum PermissionActionType {
    NONE = 'None',
    OWNED = 'Owned',
    ALL = 'All',
}

export const enum PermissionLevelType {
    NONE = 'None',
    OWNER = 'Owner',
    PUBLISHING_EDITOR = 'PublishingEditor',
    EDITOR = 'Editor',
    PUBLISHING_AUTHOR = 'PublishingAuthor',
    AUTHOR = 'Author',
    NONEDITING_AUTHOR = 'NoneditingAuthor',
    REVIEWER = 'Reviewer',
    CONTRIBUTOR = 'Contributor',
    CUSTOM = 'Custom',
}

export const enum CalendarPermissionLevelType {
    NONE = 'None',
    OWNER = 'Owner',
    PUBLISHING_EDITOR = 'PublishingEditor',
    EDITOR = 'Editor',
    PUBLISHING_AUTHOR = 'PublishingAuthor',
    AUTHOR = 'Author',
    NONEDITING_AUTHOR = 'NoneditingAuthor',
    REVIEWER = 'Reviewer',
    CONTRIBUTOR = 'Contributor',
    FREE_BUSY_TIME_ONLY = 'FreeBusyTimeOnly',
    FREE_BUSY_TIME_AND_SUBJECT_AND_LOCATION = 'FreeBusyTimeAndSubjectAndLocation',
    CUSTOM = 'Custom',
}

export const enum SearchFolderTraversalType {
    SHALLOW = 'Shallow',
    DEEP = 'Deep',
}

/**
 * Specifies the flag for action value that must appear on incoming
 * messages in order for the condition or exception to apply.
 */
export const enum FlaggedForActionType {
    ANY = 'Any',
    CALL = 'Call',
    DO_NOT_FORWARD = 'DoNotForward',
    FOLLOW_UP = 'FollowUp',
    FYI = 'FYI',
    FORWARD = 'Forward',
    NO_RESPONSE_NECESSARY = 'NoResponseNecessary',
    READ = 'Read',
    REPLY = 'Reply',
    REPLY_TO_ALL = 'ReplyToAll',
    REVIEW = 'Review',
}

/**
 * Specifies the importance that is stamped on incoming messages
 * in order for the condition or exception to apply.
 */
export const enum ImportanceChoicesType {
    LOW = 'Low',
    NORMAL = 'Normal',
    HIGH = 'High',
}

/**
 * Indicates the sensitivity that must be stamped on incoming messages
 * in order for the condition or exception to apply.
 */
export const enum SensitivityChoicesType {
    NORMAL = 'Normal',
    PERSONAL = 'Personal',
    PRIVATE = 'Private',
    CONFIDENTIAL = 'Confidential',
}

/**
 * Defines user photo size.
 */
export const enum UserPhotoSizeType {
    HR_48_X_48 = 'HR48x48',
    HR_64_X_64 = 'HR64x64',
    HR_96_X_96 = 'HR96x96',
    HR_120_X_120 = 'HR120x120',
    HR_240_X_240 = 'HR240x240',
    HR_360_X_360 = 'HR360x360',
    HR_432_X_432 = 'HR432x432',
    HR_504_X_504 = 'HR504x504',
    HR_648_X_648 = 'HR648x648',
    HR_1024_X_N = 'HR1024xN',
    HR_1920_X_N = 'HR1920xN',
}

/**
 * Defines user photo type.
 */
export const enum UserPhotoTypeType {
    USER_PHOTO = 'UserPhoto',
    PROFILE_HEADER_PHOTO = 'ProfileHeaderPhoto',
}

/**
 * Defines meeting attendee type.
 */
export const enum MeetingAttendeeType {
    ORGANIZER = 'Organizer',
    REQUIRED = 'Required',
    OPTIONAL = 'Optional',
    ROOM = 'Room',
    RESOURCE = 'Resource',
}

/**
 * List of suggestion qualities.
 */
export const enum SuggestionQuality {
    EXCELLENT = 'Excellent',
    GOOD = 'Good',
    FAIR = 'Fair',
    POOR = 'Poor',
}

/**
 * List of free busy types.
 */
export const enum LegacyFreeBusyType {
    FREE = 'Free',
    TENTATIVE = 'Tentative',
    BUSY = 'Busy',
    OOF = 'OOF',
    WORKING_ELSEWHERE = 'WorkingElsewhere',
    NO_DATA = 'NoData',
}

/**
 * List of free busy view.
 */
export const enum FreeBusyViewEnum {
    NONE = 'None',
    MERGED_ONLY = 'MergedOnly',
    FREE_BUSY = 'FreeBusy',
    FREE_BUSY_MERGED = 'FreeBusyMerged',
    DETAILED = 'Detailed',
    DETAILED_MERGED = 'DetailedMerged',
}

/**
 * List of working days scheduled for the mailbox user.
 */
export const enum DayOfWeekType {
    SUNDAY = 'Sunday',
    MONDAY = 'Monday',
    TUESDAY = 'Tuesday',
    WEDNESDAY = 'Wednesday',
    THURSDAY = 'Thursday',
    FRIDAY = 'Friday',
    SATURDAY = 'Saturday',
    DAY = 'Day',
    WEEKDAY = 'Weekday',
    WEEKEND_DAY = 'WeekendDay',
}

/**
 * Defines to whom external OOF messages are sent.
 */
export const enum ExternalAudience {
    NONE = "None",
    KNOWN = "Known",
    ALL = "All",
}

/**
 * Defines user's OOF state.
 */
export const enum OofState {
    DISABLED = "Disabled",
    ENABLED = "Enabled",
    SCHEDULED = "Scheduled",
}

/**
 * List of mail box types.
 */
export const enum MailboxTypeType {
    UNKNOWN = 'Unknown',
    ONE_OFF = 'OneOff',
    MAILBOX = 'Mailbox',
    PUBLIC_DL = 'PublicDL',
    PRIVATE_DL = 'PrivateDL',
    CONTACT = 'Contact',
    PUBLIC_FOLDER = 'PublicFolder',
    GROUP_MAILBOX = 'GroupMailbox',
    IMPLICIT_CONTACT = 'ImplicitContact',
    USER = 'User',
}

/**
 * List of mail body types.
 */
export const enum BodyTypeType {
    HTML = "HTML",
    TEXT = "Text",
}

/**
 * List of body response types.
 */
export const enum BodyTypeResponseType {
    BEST = "Best",
    HTML = "HTML",
    TEXT = "Text",
}

/**
 * Sync folder items scop types.
 */
export const enum SyncFolderItemsScopeType {
    NORMAL_ITEMS = "NormalItems",
    NORMAL_AND_ASSOCIATED_ITEMS = "NormalAndAssociatedItems",
}

/**
 * Icon index type.
 */
export const enum IconIndexType {
    DEFAULT = 'Default',
    POST_ITEM = 'PostItem',
    MAIL_READ = 'MailRead',
    MAIL_UNREAD = 'MailUnread',
    MAIL_REPLIED = 'MailReplied',
    MAIL_FORWARDED = 'MailForwarded',
    MAIL_ENCRYPTED = 'MailEncrypted',
    MAIL_SMIME_SIGNED = 'MailSmimeSigned',
    MAIL_ENCRYPTED_REPLIED = 'MailEncryptedReplied',
    MAIL_SMIME_SIGNED_REPLIED = 'MailSmimeSignedReplied',
    MAIL_ENCRYPTED_FORWARDED = 'MailEncryptedForwarded',
    MAIL_SMIME_SIGNED_FORWARDED = 'MailSmimeSignedForwarded',
    MAIL_ENCRYPTED_READ = 'MailEncryptedRead',
    MAIL_SMIME_SIGNED_READ = 'MailSmimeSignedRead',
    MAIL_IRM = 'MailIrm',
    MAIL_IRM_FORWARDED = 'MailIrmForwarded',
    MAIL_IRM_REPLIED = 'MailIrmReplied',
    SMS_SUBMITTED = 'SmsSubmitted',
    SMS_ROUTED_TO_DELIVERY_POINT = 'SmsRoutedToDeliveryPoint',
    SMS_ROUTED_TO_EXTERNAL_MESSAGING_SYSTEM = 'SmsRoutedToExternalMessagingSystem',
    SMS_DELIVERED = 'SmsDelivered',
    OUTLOOK_DEFAULT_FOR_CONTACTS = 'OutlookDefaultForContacts',
    APPOINTMENT_ITEM = 'AppointmentItem',
    APPOINTMENT_RECUR = 'AppointmentRecur',
    APPOINTMENT_MEET = 'AppointmentMeet',
    APPOINTMENT_MEET_RECUR = 'AppointmentMeetRecur',
    APPOINTMENT_MEET_NY = 'AppointmentMeetNY',
    APPOINTMENT_MEET_YES = 'AppointmentMeetYes',
    APPOINTMENT_MEET_NO = 'AppointmentMeetNo',
    APPOINTMENT_MEET_MAYBE = 'AppointmentMeetMaybe',
    APPOINTMENT_MEET_CANCEL = 'AppointmentMeetCancel',
    APPOINTMENT_MEET_INFO = 'AppointmentMeetInfo',
    TASK_ITEM = 'TaskItem',
    TASK_RECUR = 'TaskRecur',
    TASK_OWNED = 'TaskOwned',
    TASK_DELEGATED = 'TaskDelegated',
}

/**
 * Inference classification type.
 */
export const enum InferenceClassificationType {
    FOCUSED = 'Focused',
    OTHER = 'Other',
}

/**
 * Calendar item type.
 */
export const enum CalendarItemTypeType {
    SINGLE = 'Single',
    OCCURRENCE = 'Occurrence',
    EXCEPTION = 'Exception',
    RECURRING_MASTER = 'RecurringMaster',
}

/**
 * Response type.
 */
export const enum ResponseTypeType {
    UNKNOWN = 'Unknown',
    ORGANIZER = 'Organizer',
    TENTATIVE = 'Tentative',
    ACCEPT = 'Accept',
    DECLINE = 'Decline',
    NO_RESPONSE_RECEIVED = 'NoResponseReceived',
}

/**
 * File as mapping type.
 */
export const enum FileAsMappingType {
    NONE = 'None',
    LAST_COMMA_FIRST = 'LastCommaFirst',
    FIRST_SPACE_LAST = 'FirstSpaceLast',
    COMPANY = 'Company',
    LAST_COMMA_FIRST_COMPANY = 'LastCommaFirstCompany',
    COMPANY_LAST_FIRST = 'CompanyLastFirst',
    LAST_FIRST = 'LastFirst',
    LAST_FIRST_COMPANY = 'LastFirstCompany',
    COMPANY_LAST_COMMA_FIRST = 'CompanyLastCommaFirst',
    LAST_FIRST_SUFFIX = 'LastFirstSuffix',
    LAST_SPACE_FIRST_COMPANY = 'LastSpaceFirstCompany',
    COMPANY_LAST_SPACE_FIRST = 'CompanyLastSpaceFirst',
    LAST_SPACE_FIRST = 'LastSpaceFirst',
    DISPLAY_NAME = 'DisplayName',
    FIRST_NAME = 'FirstName',
    LAST_FIRST_MIDDLE_SUFFIX = 'LastFirstMiddleSuffix',
    LAST_NAME = 'LastName',
    EMPTY = 'Empty',
}

/**
 * Contact source type.
 */
export const enum ContactSourceType {
    ACTIVE_DIRECTORY = 'ActiveDirectory',
    STORE = 'Store',
}

/**
 * Physical address index type.
 */
export const enum PhysicalAddressIndexType {
    NONE = 'None',
    HOME = 'Home',
    BUSINESS = 'Business',
    OTHER = 'Other',
}

/**
 * Meeting request type.
 */
export const enum MeetingRequestTypeType {
    NONE = 'None',
    FULL_UPDATE = 'FullUpdate',
    INFORMATIONAL_UPDATE = 'InformationalUpdate',
    NEW_MEETING_REQUEST = 'NewMeetingRequest',
    OUTDATED = 'Outdated',
    SILENT_UPDATE = 'SilentUpdate',
    PRINCIPAL_WANTS_COPY = 'PrincipalWantsCopy',
}

/**
 * Task delegate state type.
 */
export const enum TaskDelegateStateType {
    NO_MATCH = 'NoMatch',
    OWN_NEW = 'OwnNew',
    OWNED = 'Owned',
    ACCEPTED = 'Accepted',
    DECLINED = 'Declined',
    MAX = 'Max',
}

/**
 * Task status type.
 */
export const enum TaskStatusType {
    NOT_STARTED = 'NotStarted',
    IN_PROGRESS = 'InProgress',
    COMPLETED = 'Completed',
    WAITING_ON_OTHERS = 'WaitingOnOthers',
    DEFERRED = 'Deferred',
}

/**
 * User configuration property type.
 */
export const enum UserConfigurationPropertyTypeEnum {
    ID = 'Id',
    DICTIONARY = 'Dictionary',
    XML_DATA = 'XmlData',
    BINARY_DATA = 'BinaryData',
    ALL = 'All',
}

/**
 * Create action type.
 */
export const enum CreateActionType {
    CREATE_NEW = 'CreateNew',
    UPDATE = 'Update',
    UPDATE_OR_CREATE = 'UpdateOrCreate',
}

/**
 * Id format type.
 */
export const enum IdFormatType {
    EWS_LEGACY_ID = 'EwsLegacyId',
    EWS_ID = 'EwsId',
    ENTRY_ID = 'EntryId',
    HEX_ENTRY_ID = 'HexEntryId',
    STORE_ID = 'StoreId',
    OWA_ID = 'OwaId',
}

/**
 * Message disposition type.
 */
export const enum MessageDispositionType {
    SAVE_ONLY = 'SaveOnly',
    SEND_ONLY = 'SendOnly',
    SEND_AND_SAVE_COPY = 'SendAndSaveCopy',
}

/**
 * Folder query traversal type.
 */
export const enum FolderQueryTraversalType {
    SHALLOW = 'Shallow',
    DEEP = 'Deep',
    SOFT_DELETED = 'SoftDeleted',
}

/**
 * Item query traversal type.
 */
export const enum ItemQueryTraversalType {
    SHALLOW = 'Shallow',
    ASSOCIATED = 'Associated',
    SOFT_DELETED = 'SoftDeleted'
}
/**
 * Report message action type.
 */
export const enum ReportMessageActionType {
    JUNK = 'Junk',
    NOT_JUNK = 'NotJunk',
    PHISH = 'Phish',
    UNSUBSCRIBE = 'Unsubscribe',
}

/**
 * Connection status type.
 */
export const enum ConnectionStatusType {
    OK = "OK",
    CLOSED = "Closed",
}

/**
 * Resolve name search scope type.
 */
export const enum ResolveNamesSearchScopeType {
    ACTIVE_DIRECTORY = 'ActiveDirectory',
    ACTIVE_DIRECTORY_CONTACTS = 'ActiveDirectoryContacts',
    CONTACTS = 'Contacts',
    CONTACTS_ACTIVE_DIRECTORY = 'ContactsActiveDirectory',
}

/**
 * Affected task occurrences type.
 */
export const enum AffectedTaskOccurrencesType {
    ALL_OCCURRENCES = 'AllOccurrences',
    SPECIFIED_OCCURRENCE_ONLY = 'SpecifiedOccurrenceOnly',
}

/**
 * Conflict resolution type.
 */
export const enum ConflictResolutionType {
    NEVER_OVERWRITE = 'NeverOverwrite',
    AUTO_RESOLVE = 'AutoResolve',
    ALWAYS_OVERWRITE = 'AlwaysOverwrite',
}

/**
 * Calendar item update operation type.
 */
export const enum CalendarItemUpdateOperationType {
    SEND_TO_NONE = 'SendToNone',
    SEND_ONLY_TO_ALL = 'SendOnlyToAll',
    SEND_ONLY_TO_CHANGED = 'SendOnlyToChanged',
    SEND_TO_ALL_AND_SAVE_COPY = 'SendToAllAndSaveCopy',
    SEND_TO_CHANGED_AND_SAVE_COPY = 'SendToChangedAndSaveCopy',
}

/**
 * Calendar item create/delete operation type.
 */
export const enum CalendarItemCreateOrDeleteOperationType {
    SEND_TO_NONE = 'SendToNone',
    SEND_ONLY_TO_ALL = 'SendOnlyToAll',
    SEND_TO_ALL_AND_SAVE_COPY = 'SendToAllAndSaveCopy',
}

/**
 * Role member type.
 */
export const enum RoleMemberTypeType {
    NONE = 'None',
    PASSPORT = 'Passport',
    EVERYONE = 'Everyone',
    EMAIL = 'Email',
    PHONE = 'Phone',
    SKYPE_ID = 'SkypeId',
    EXTERNAL_ID = 'ExternalId',
    GROUP = 'Group',
    GUID = 'Guid',
    ROLE = 'Role',
    SERVICE = 'Service',
    CIRCLE = 'Circle',
    DOMAIN = 'Domain',
    PARTNER = 'Partner',
}

/**
 * Flag status type.
 */
export const enum FlagStatusType {
    NOT_FLAGGED = 'NotFlagged',
    FLAGGED = 'Flagged',
    COMPLETE = 'Complete',
}

/**
 * Predicted action reason type.
 */
export const enum PredictedActionReasonType {
    NONE = 'None',
    CONVERSATION_STARTER_IS_YOU = 'ConversationStarterIsYou',
    ONLY_RECIPIENT = 'OnlyRecipient',
    CONVERSATION_CONTRIBUTIONS = 'ConversationContributions',
    MARKED_IMPORTANT_BY_SENDER = 'MarkedImportantBySender',
    SENDER_IS_MANAGER = 'SenderIsManager',
    SENDER_IS_IN_MANAGEMENT_CHAIN = 'SenderIsInManagementChain',
    SENDER_IS_DIRECT_REPORT = 'SenderIsDirectReport',
    ACTION_BASED_ON_SENDER = 'ActionBasedOnSender',
    NAME_ON_TO_LINE = 'NameOnToLine',
    NAME_ON_CC_LINE = 'NameOnCcLine',
    MANAGER_POSITION = 'ManagerPosition',
    REPLY_TO_A_MESSAGE_FROM_ME = 'ReplyToAMessageFromMe',
    PREVIOUSLY_FLAGGED = 'PreviouslyFlagged',
    ACTION_BASED_ON_RECIPIENTS = 'ActionBasedOnRecipients',
    ACTION_BASED_ON_SUBJECT_WORDS = 'ActionBasedOnSubjectWords',
    ACTION_BASED_ON_BASED_ON_BODY_WORDS = 'ActionBasedOnBasedOnBodyWords',
}

/**
 * Sharing action importance.
 */
export const enum SharingActionImportance {
    PRIMARY = 'Primary',
    SECONDARY = 'Secondary',
}

/**
 * Sharing action type.
 */
export const enum SharingActionType {
    ACCEPT = "Accept",
}

/**
 * Sharing action.
 */
export const enum SharingAction {
    ACCEPT_AND_VIEW_CALENDAR = 'AcceptAndViewCalendar',
    VIEW_CALENDAR = 'ViewCalendar',
    ADD_THIS_CALENDAR = 'AddThisCalendar',
    ACCEPT = 'Accept',
}

/**
 * Lobby bypass type.
 */
export const enum LobbyBypassType {
    DISABLED = 'Disabled',
    ENABLED_FOR_GATEWAY_PARTICIPANTS = 'EnabledForGatewayParticipants',
}

/**
 * User configuration dictionary object type.
 */
export const enum UserConfigurationDictionaryObjectTypesType {
    DATE_TIME = 'DateTime',
    BOOLEAN = 'Boolean',
    BYTE = 'Byte',
    STRING = 'String',
    INTEGER_32 = 'Integer32',
    UNSIGNED_INTEGER_32 = 'UnsignedInteger32',
    INTEGER_64 = 'Integer64',
    UNSIGNED_INTEGER_64 = 'UnsignedInteger64',
    STRING_ARRAY = 'StringArray',
    BYTE_ARRAY = 'ByteArray',
}

/**
 * Disposal type.
 */
export const enum DisposalType {
    HARD_DELETE = 'HardDelete',
    SOFT_DELETE = 'SoftDelete',
    MOVE_TO_DELETED_ITEMS = 'MoveToDeletedItems',
}

/**
 * Notification event type.
 */
export const enum NotificationEventTypeType {
    COPIED_EVENT = 'CopiedEvent',
    CREATED_EVENT = 'CreatedEvent',
    DELETED_EVENT = 'DeletedEvent',
    MODIFIED_EVENT = 'ModifiedEvent',
    MOVED_EVENT = 'MovedEvent',
    NEW_MAIL_EVENT = 'NewMailEvent',
    FREE_BUSY_CHANGED_EVENT = 'FreeBusyChangedEvent',
}

/**
 * Online meeting access level type.
 */
export const enum OnlineMeetingAccessLevelType {
    LOCKED = 'Locked',
    INVITED = 'Invited',
    INTERNAL = 'Internal',
    EVERYONE = 'Everyone',
}

/**
 * Presenters type.
 */
export const enum PresentersType {
    DISABLED = 'Disabled',
    INTERNAL = 'Internal',
    EVERYONE = 'Everyone',
}

/**
 * Email reminder change type.
 */
export const enum EmailReminderChangeType {
    NONE = 'None',
    ADDED = 'Added',
    OVERRIDE = 'Override',
    DELETED = 'Deleted',
}

/**
 * Email reminder send option.
 */
export const enum EmailReminderSendOption {
    NOT_SET = 'NotSet',
    USER = 'User',
    ALL_ATTENDEES = 'AllAttendees',
    STAFF = 'Staff',
    CUSTOMER = 'Customer',
}

/**
 * Day of week index type.
 */
export const enum DayOfWeekIndexType {
    FIRST = 'First',
    SECOND = 'Second',
    THIRD = 'Third',
    FOURTH = 'Fourth',
    LAST = 'Last',
}

/**
 * Month names type.
 */
export const enum MonthNamesType {
    JANUARY = 'January',
    FEBRUARY = 'February',
    MARCH = 'March',
    APRIL = 'April',
    MAY = 'May',
    JUNE = 'June',
    JULY = 'July',
    AUGUST = 'August',
    SEPTEMBER = 'September',
    OCTOBER = 'October',
    NOVEMBER = 'November',
    DECEMBER = 'December',
}

/**
 * Email address key type.
 */
export const enum EmailAddressKeyType {
    EMAIL_ADDRESS_1 = 'EmailAddress1',
    EMAIL_ADDRESS_2 = 'EmailAddress2',
    EMAIL_ADDRESS_3 = 'EmailAddress3',
}

/**
 * Abch email address type.
 */
export const enum AbchEmailAddressTypeType {
    PERSONAL = 'Personal',
    BUSINESS = 'Business',
    OTHER = 'Other',
    PASSPORT = 'Passport',
}

/**
 * Physical address key type.
 */
export const enum PhysicalAddressKeyType {
    HOME = 'Home',
    BUSINESS = 'Business',
    OTHER = 'Other',
}

/**
 * Phone number key type.
 */
export const enum PhoneNumberKeyType {
    ASSISTANT_PHONE = 'AssistantPhone',
    BUSINESS_FAX = 'BusinessFax',
    BUSINESS_PHONE = 'BusinessPhone',
    BUSINESS_PHONE_2 = 'BusinessPhone2',
    CALLBACK = 'Callback',
    CAR_PHONE = 'CarPhone',
    COMPANY_MAIN_PHONE = 'CompanyMainPhone',
    HOME_FAX = 'HomeFax',
    HOME_PHONE = 'HomePhone',
    HOME_PHONE_2 = 'HomePhone2',
    ISDN = 'Isdn',
    MOBILE_PHONE = 'MobilePhone',
    OTHER_FAX = 'OtherFax',
    OTHER_TELEPHONE = 'OtherTelephone',
    PAGER = 'Pager',
    PRIMARY_PHONE = 'PrimaryPhone',
    RADIO_PHONE = 'RadioPhone',
    TELEX = 'Telex',
    TTY_TDD_PHONE = 'TtyTddPhone',
    BUSINESS_MOBILE = 'BusinessMobile',
    IP_PHONE = 'IPPhone',
    MMS = 'Mms',
    MSN = 'Msn',
}

/**
 * IM address key type.
 */
export const enum ImAddressKeyType {
    IM_ADDRESS_1 = 'ImAddress1',
    IM_ADDRESS_2 = 'ImAddress2',
    IM_ADDRESS_3 = 'ImAddress3',
}

/**
 * Contact url key type.
 */
export const enum ContactUrlKeyType {
    PERSONAL = 'Personal',
    BUSINESS = 'Business',
    ATTACHMENT = 'Attachment',
    EBC_DISPLAY_DEFINITION = 'EbcDisplayDefinition',
    EBC_FINAL_IMAGE = 'EbcFinalImage',
    EBC_LOGO = 'EbcLogo',
    FEED = 'Feed',
    IMAGE = 'Image',
    INTERNAL_MARKER = 'InternalMarker',
    OTHER = 'Other',
}

/**
 * Member status type.
 */
export const enum MemberStatusType {
    UNRECOGNIZED = 'Unrecognized',
    NORMAL = 'Normal',
    DEMOTED = 'Demoted',
}

/**
 * Email position type.
 */
export const enum EmailPositionType {
    LATEST_REPLY = 'LatestReply',
    OTHER = 'Other',
    SUBJECT = 'Subject',
    SIGNATURE = 'Signature',
}

/**
 * Send prompt type.
 */
export const enum SendPromptType {
    NONE = 'None',
    SEND = 'Send',
    VOTING_OPTION = 'VotingOption',
}

/**
 * Transition target kind type.
 */
export const enum TransitionTargetKindType {
    PERIOD = 'Period',
    GROUP = 'Group',
}

/**
 * List of location source types.
 */
export const enum LocationSourceType {
    NONE = 'None',
    LOCATION_SERVICES = 'LocationServices',
    PHONEBOOK_SERVICES = 'PhonebookServices',
    DEVICE = 'Device',
    CONTACT = 'Contact',
    RESOURCE = 'Resource',
}

/**
 * Defines the well-known property set IDs for extended MAPI properties.
 */
export const enum DistinguishedPropertySetType {
    MEETING = 'Meeting',
    APPOINTMENT = 'Appointment',
    COMMON = 'Common',
    PUBLIC_STRINGS = 'PublicStrings',
    ADDRESS = 'Address',
    INTERNET_HEADERS = 'InternetHeaders',
    CALENDAR_ASSISTANT = 'CalendarAssistant',
    UNIFIED_MESSAGING = 'UnifiedMessaging',
    TASK = 'Task',
    SHARING = 'Sharing',
}

/**
 * Defines property type of a property tag.
 */
export const enum MapiPropertyTypeType {
    APPLICATION_TIME = 'ApplicationTime',
    APPLICATION_TIME_ARRAY = 'ApplicationTimeArray',
    BINARY = 'Binary',
    BINARY_ARRAY = 'BinaryArray',
    BOOLEAN = 'Boolean',
    CLSID = 'CLSID',
    CLSID_ARRAY = 'CLSIDArray',
    CURRENCY = 'Currency',
    CURRENCY_ARRAY = 'CurrencyArray',
    DOUBLE = 'Double',
    DOUBLE_ARRAY = 'DoubleArray',
    ERROR = 'Error',
    FLOAT = 'Float',
    FLOAT_ARRAY = 'FloatArray',
    INTEGER = 'Integer',
    INTEGER_ARRAY = 'IntegerArray',
    LONG = 'Long',
    LONG_ARRAY = 'LongArray',
    NULL = 'Null',
    OBJECT = 'Object',
    OBJECT_ARRAY = 'ObjectArray',
    SHORT = 'Short',
    SHORT_ARRAY = 'ShortArray',
    SYSTEM_TIME = 'SystemTime',
    SYSTEM_TIME_ARRAY = 'SystemTimeArray',
    STRING = 'String',
    STRING_ARRAY = 'StringArray',
}

/**
 * Defines Mapi property ids
 * The full list is documented here: https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxprops/37dd7329-97a4-42ff-974d-d805ac4d7211
 * 
 * From https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/mapi-property-identifier-overview :
 *      The range for properties defined by MAPI runs from 0x0001 to 0x3FFF. These properties are referred to as MAPI-defined properties. The range 0x4000 to 0x7FFF 
 *      belongs to message and recipient properties, and either clients or service providers can define properties in this range. Properties in the range of 0x0001 to 0x7FFF 
 *      are referred to as tagged properties. Beyond 0x8000 is the range for what is known as named properties, or properties that include a 32-bit globally unique identifier 
 *      (GUID) and either a Unicode character string or numeric value. Clients can use named properties to customize their property set.
 */
export const enum MapiPropertyIds {
    // Tagged Properties
    FLAG_STATUS = 4240,             // x1090

    // Named Properties
    REMINDER_TIME = 34050,          // x8502
    REMINDER_SET = 34051,           // x8503
    EMAIL1_DISPLAY_NAME = 32896,    // x8080
    EMAIL1_ADDRESS_TYPE = 32898,    // x8082
    EMAIL1_ADDRESS = 32899,         // x8083
    EMAIL1_ORIG_DISPLAY_NAME = 32900, // x8084
    EMAIL2_DISPLAY_NAME = 32912,    // x8090
    EMAIL2_ADDRESS_TYPE = 32914,    // x8092
    EMAIL2_ADDRESS = 32915,         // x8093
    EMAIL2_ORIG_DISPLAY_NAME = 32916, // x8094
    EMAIL3_DISPLAY_NAME = 32928,    // x80A0
    EMAIL3_ADDRESS_TYPE = 32930,    // x80A2
    EMAIL3_ADDRESS = 32931,         // x80A3
    EMAIL3_ORIG_DISPLAY_NAME = 32932 // x80A4
}

/**
 * Defines the type of client access token.
 */
export const enum ClientAccessTokenTypeType {
    CALLER_IDENTITY = 'CallerIdentity',
    EXTENSION_CALLBACK = 'ExtensionCallback',
    SCOPED_TOKEN = 'ScopedToken',
    EXTENSION_REST_API_CALLBACK = 'ExtensionRestApiCallback',
    CONNECTORS = 'Connectors',
    LOKI = 'Loki',
    DESKTOP_OUTLOOK = 'DesktopOutlook',
}

/**
 * Defines the set of properties to return in an item or folder response.
 */
export const enum DefaultShapeNamesType {
    ID_ONLY = 'IdOnly',
    DEFAULT = 'Default',
    ALL_PROPERTIES = 'AllProperties',
    PCX_PEOPLE_SEARCH = 'PcxPeopleSearch',
}

/**
 * Describes whether the page of items or conversations will start from the beginning or the
 * end of the set of items or conversations that are found by using the search criteria.
 * Seeking from the end always searches backward.
 */
export const enum IndexBasePointType {
    BEGINNING = "Beginning",
    END = "End",
}

/**
 * Defines the sort direction.
 */
export const enum SortDirectionType {
    ASCENDING = 'Ascending',
    DESCENDING = 'Descending',
}

/**
 * Defines the boundaries of a search.
 */
export const enum ContainmentModeType {
    FULL_STRING = 'FullString',
    PREFIXED = 'Prefixed',
    SUBSTRING = 'Substring',
    PREFIX_ON_WORDS = 'PrefixOnWords',
    EXACT_PHRASE = 'ExactPhrase',
}

/**
 * Defines whether the search ignores cases and spaces.
 */
export const enum ContainmentComparisonType {
    EXACT = 'Exact',
    IGNORE_CASE = 'IgnoreCase',
    IGNORE_NON_SPACING_CHARACTERS = 'IgnoreNonSpacingCharacters',
    LOOSE = 'Loose',
    IGNORE_CASE_AND_NON_SPACING_CHARACTERS = 'IgnoreCaseAndNonSpacingCharacters',
    LOOSE_AND_IGNORE_CASE = 'LooseAndIgnoreCase',
    LOOSE_AND_IGNORE_NON_SPACE = 'LooseAndIgnoreNonSpace',
    LOOSE_AND_IGNORE_CASE_AND_IGNORE_NON_SPACE = 'LooseAndIgnoreCaseAndIgnoreNonSpace',
}

/**
 * Identifies offending property paths in an error.
 * This will only occur within the MessageXml property of a ResponseMessage.
 */
export const enum ExceptionPropertyURIType {
    ATTACHMENT_NAME = 'attachment:Name',
    ATTACHMENT_CONTENT_TYPE = 'attachment:ContentType',
    ATTACHMENT_CONTENT = 'attachment:Content',
    RECURRENCE_MONTH = 'recurrence:Month',
    RECURRENCE_DAY_OF_WEEK_INDEX = 'recurrence:DayOfWeekIndex',
    RECURRENCE_DAYS_OF_WEEK = 'recurrence:DaysOfWeek',
    RECURRENCE_DAY_OF_MONTH = 'recurrence:DayOfMonth',
    RECURRENCE_INTERVAL = 'recurrence:Interval',
    RECURRENCE_NUMBER_OF_OCCURRENCES = 'recurrence:NumberOfOccurrences',
    TIMEZONE_OFFSET = 'timezone:Offset',
}

/**
 * Represents standard groupings for GroupBy queries.
 */
export const enum StandardGroupByType {
    CONVERSATION_TOPIC = "ConversationTopic"
}

/**
 * This max/min evaluation is applied to the field specified within the group by
 * instance for EACH item within that group.  This determines which item from each group
 * is to be selected as the representative for that group.
 */
export const enum AggregateType {
    MINIMUM = "Minimum",
    MAXIMUM = "Maxinum"
}

/**
 * Values for the FieldIndex in a IndexedFieldURI. 
 */
export const enum FieldIndexValue {
    CONTACT_EMAIL_ADDRESS1 = "EmailAddress1",
    CONTACT_EMAIL_ADDRESS2 = "EmailAddress2",
    CONTACT_EMAIL_ADDRESS3 = "EmailAddress3",
    CONTACT_IM_ADDRESS1 = "ImAddress1",
    CONTACT_IM_ADDRESS2 = "ImAddress2",
    CONTACT_IM_ADDRESS3 = "ImAddress3",
    CONTACT_PHYSICAL_ADDRESS_HOME = "Home",
    CONTACT_PHYSICAL_ADDRESS_BUSINESS = "Business",
    CONTACT_PHYSICAL_ADDRESS_OTHER = "Other",
    CONTACT_PHONE_NUMBER_BUSINESS = "BusinessPhone",
    CONTACT_PHONE_NUMBER_BUSINESS2 = "BusinessPhone2",
    CONTACT_PHONE_NUMBER_BUSINESS_FAX = "BusinessFax",
    CONTACT_PHONE_NUMBER_COMPANY = "CompanyMainPhone",
    CONTACT_PHONE_NUMBER_ASSISTANT = "AssistantPhone",
    CONTACT_PHONE_NUMBER_HOME = "HomePhone",
    CONTACT_PHONE_NUMBER_HOME2 = "HomePhone2",
    CONTACT_PHONE_NUMBER_HOME_FAX = "HomeFax",
    CONTACT_PHONE_NUMBER_MOBILE = "MobilePhone",
    CONTACT_PHONE_NUMBER_PAGER = "Pager",
    CONTACT_PHONE_NUMBER_PRIMARY = "PrimaryPhone",
    CONTACT_PHONE_NUMBER_OTHER_FAX = "OtherFax",
    CONTACT_PHONE_NUMBER_OTHER = "OtherTelephone",
    CONTACT_PHONE_NUMBER_CALLBACK = "Callback",
    CONTACT_PHONE_NUMBER_RADIO = "RadioPhone"
}

/**
 * Defines the dictionary that contains the member to return.
 */
export const enum DictionaryURIType {
    ITEM_INTERNET_MESSAGE_HEADER = 'item:InternetMessageHeader',
    CONTACTS_IM_ADDRESS = 'contacts:ImAddress',
    CONTACTS_PHYSICAL_ADDRESS_STREET = 'contacts:PhysicalAddress:Street',
    CONTACTS_PHYSICAL_ADDRESS_CITY = 'contacts:PhysicalAddress:City',
    CONTACTS_PHYSICAL_ADDRESS_STATE = 'contacts:PhysicalAddress:State',
    CONTACTS_PHYSICAL_ADDRESS_COUNTRY_OR_REGION = 'contacts:PhysicalAddress:CountryOrRegion',
    CONTACTS_PHYSICAL_ADDRESS_POSTAL_CODE = 'contacts:PhysicalAddress:PostalCode',
    CONTACTS_PHONE_NUMBER = 'contacts:PhoneNumber',
    CONTACTS_EMAIL_ADDRESS = 'contacts:EmailAddress',
    DISTRIBUTIONLIST_MEMBERS_MEMBER = 'distributionlist:Members:Member',
}

/**
 * Defines where to search mailboxes.
 */
export const enum MailboxSearchLocationType {
    PRIMARY_ONLY = 'PrimaryOnly',
    ARCHIVE_ONLY = 'ArchiveOnly',
    ALL = 'All'
}

/**
 * Types of sub-tree traversal for conversations
 */
export const enum ConversationQueryTraversalType {
    SHALLOW = 'Shallow',
    DEEP = 'Deep'
}

/**
 * Types of view filters for finding items/conversations
 */
export const enum ViewFilterType {
    ALL = 'All',
    FLAGGED = 'Flagged',
    HAS_ATTACHMENT = 'HasAttachment',
    TO_OR_CC_ME = 'ToOrCcMe',
    UNREAD = 'Unread',
    TASK_ACTIVE = 'TaskActive',
    TASK_OVERDUE= 'TaskOverdue',
    TASK_COMPLETED = 'TaskCompleted',
    NO_CLUTTER = 'NoClutter',
    CLUTTER = 'Clutter'
}

/**
 * Sort order for conversation nodes
 */
export const enum ConversationNodeSortOrder {
    TREE_ORDER_ASCENDING = 'TreeOrderAscending',
    TREE_ORDER_DESCENDING = 'TreeOrderDescending',
    DATE_ORDER_ASCENDING = 'DateOrderAscending',
    DATE_ORDER_DESCENDING = 'DateOrderDescending'
}

export const enum ConversationActionTypeType {
    ALWAYS_CATEGORIZE = 'AlwaysCategorize',
    ALWAYS_DELETE = 'AlwaysDelete',
    ALWAYS_MOVE = 'AlwaysMove',
    DELETE = 'Delete',
    MOVE = 'Move',
    COPY = 'Copy',
    SET_READ_STATE = 'SetReadState',
    SET_RETENTION_POLICY = 'SetRetentionPolicy',
    FLAG = 'Flag'
}

export const enum RetentionType {
    DELETE = 'Delete',
    ARCHIVE = 'Archive'
}

/**
 * Defines distinguished folder id.
 */
export const enum DistinguishedFolderIdNameType {
    CALENDAR = 'calendar',
    CONTACTS = 'contacts',
    DELETEDITEMS = 'deleteditems',
    DRAFTS = 'drafts',
    INBOX = 'inbox',
    JOURNAL = 'journal',
    NOTES = 'notes',
    OUTBOX = 'outbox',
    SENTITEMS = 'sentitems',
    TASKS = 'tasks',
    MSGFOLDERROOT = 'msgfolderroot',
    PUBLICFOLDERSROOT = 'publicfoldersroot',
    ROOT = 'root',
    JUNKEMAIL = 'junkemail',
    SEARCHFOLDERS = 'searchfolders',
    VOICEMAIL = 'voicemail',
    RECOVERABLEITEMSROOT = 'recoverableitemsroot',
    RECOVERABLEITEMSDELETIONS = 'recoverableitemsdeletions',
    RECOVERABLEITEMSVERSIONS = 'recoverableitemsversions',
    RECOVERABLEITEMSPURGES = 'recoverableitemspurges',
    RECOVERABLEITEMSDISCOVERYHOLDS = 'recoverableitemsdiscoveryholds',
    ARCHIVEROOT = 'archiveroot',
    ARCHIVEMSGFOLDERROOT = 'archivemsgfolderroot',
    ARCHIVEDELETEDITEMS = 'archivedeleteditems',
    ARCHIVEINBOX = 'archiveinbox',
    ARCHIVERECOVERABLEITEMSROOT = 'archiverecoverableitemsroot',
    ARCHIVERECOVERABLEITEMSDELETIONS = 'archiverecoverableitemsdeletions',
    ARCHIVERECOVERABLEITEMSVERSIONS = 'archiverecoverableitemsversions',
    ARCHIVERECOVERABLEITEMSPURGES = 'archiverecoverableitemspurges',
    ARCHIVERECOVERABLEITEMSDISCOVERYHOLDS = 'archiverecoverableitemsdiscoveryholds',
    SYNCISSUES = 'syncissues',
    CONFLICTS = 'conflicts',
    LOCALFAILURES = 'localfailures',
    SERVERFAILURES = 'serverfailures',
    RECIPIENTCACHE = 'recipientcache',
    QUICKCONTACTS = 'quickcontacts',
    CONVERSATIONHISTORY = 'conversationhistory',
    ADMINAUDITLOGS = 'adminauditlogs',
    TODOSEARCH = 'todosearch',
    MYCONTACTS = 'mycontacts',
    DIRECTORY = 'directory',
    IMCONTACTLIST = 'imcontactlist',
    PEOPLECONNECT = 'peopleconnect',
    FAVORITES = 'favorites',
    MECONTACT = 'mecontact',
    PERSONMETADATA = 'personmetadata',
    TEAMSPACEACTIVITY = 'teamspaceactivity',
    TEAMSPACEMESSAGING = 'teamspacemessaging',
    TEAMSPACEWORKITEMS = 'teamspaceworkitems',
    SCHEDULED = 'scheduled',
    ORIONNOTES = 'orionnotes',
    TAGITEMS = 'tagitems',
    ALLTAGGEDITEMS = 'alltaggeditems',
    ALLCATEGORIZEDITEMS = 'allcategorizeditems',
    EXTERNALCONTACTS = 'externalcontacts',
    TEAMCHAT = 'teamchat',
    TEAMCHATHISTORY = 'teamchathistory',
    YAMMERDATA = 'yammerdata',
    YAMMERROOT = 'yammerroot',
    YAMMERINBOUND = 'yammerinbound',
    YAMMEROUTBOUND = 'yammeroutbound',
    YAMMERFEEDS = 'yammerfeeds',
    KAIZALADATA = 'kaizaladata',
    MESSAGEINGESTION = 'messageingestion',
    ONEDRIVEROOT = 'onedriveroot',
    ONEDRIVERECYLEBIN = 'onedriverecylebin',
    ONEDRIVESYSTEM = 'onedrivesystem',
    ONEDRIVEVOLUME = 'onedrivevolume',
    IMPORTANT = 'important',
    STARRED = 'starred',
    ARCHIVE = 'archive',
}

/**
 * Defines the URI of the property.
 */
export const enum UnindexedFieldURIType {
    FOLDER_FOLDER_ID = 'folder:FolderId',
    FOLDER_PARENT_FOLDER_ID = 'folder:ParentFolderId',
    FOLDER_DISPLAY_NAME = 'folder:DisplayName',
    FOLDER_UNREAD_COUNT = 'folder:UnreadCount',
    FOLDER_TOTAL_COUNT = 'folder:TotalCount',
    FOLDER_CHILD_FOLDER_COUNT = 'folder:ChildFolderCount',
    FOLDER_FOLDER_CLASS = 'folder:FolderClass',
    FOLDER_SEARCH_PARAMETERS = 'folder:SearchParameters',
    FOLDER_MANAGED_FOLDER_INFORMATION = 'folder:ManagedFolderInformation',
    FOLDER_PERMISSION_SET = 'folder:PermissionSet',
    FOLDER_EFFECTIVE_RIGHTS = 'folder:EffectiveRights',
    FOLDER_SHARING_EFFECTIVE_RIGHTS = 'folder:SharingEffectiveRights',
    FOLDER_DISTINGUISHED_FOLDER_ID = 'folder:DistinguishedFolderId',
    FOLDER_POLICY_TAG = 'folder:PolicyTag',
    FOLDER_ARCHIVE_TAG = 'folder:ArchiveTag',
    FOLDER_REPLICA_LIST = 'folder:ReplicaList',
    ITEM_ITEM_ID = 'item:ItemId',
    ITEM_PARENT_FOLDER_ID = 'item:ParentFolderId',
    ITEM_ITEM_CLASS = 'item:ItemClass',
    ITEM_MIME_CONTENT = 'item:MimeContent',
    ITEM_ATTACHMENTS = 'item:Attachments',
    ITEM_SUBJECT = 'item:Subject',
    ITEM_DATE_TIME_RECEIVED = 'item:DateTimeReceived',
    ITEM_SIZE = 'item:Size',
    ITEM_CATEGORIES = 'item:Categories',
    ITEM_HAS_ATTACHMENTS = 'item:HasAttachments',
    ITEM_IMPORTANCE = 'item:Importance',
    ITEM_IN_REPLY_TO = 'item:InReplyTo',
    ITEM_INTERNET_MESSAGE_HEADERS = 'item:InternetMessageHeaders',
    ITEM_IS_ASSOCIATED = 'item:IsAssociated',
    ITEM_IS_DRAFT = 'item:IsDraft',
    ITEM_IS_FROM_ME = 'item:IsFromMe',
    ITEM_IS_RESEND = 'item:IsResend',
    ITEM_IS_SUBMITTED = 'item:IsSubmitted',
    ITEM_IS_UNMODIFIED = 'item:IsUnmodified',
    ITEM_DATE_TIME_SENT = 'item:DateTimeSent',
    ITEM_DATE_TIME_CREATED = 'item:DateTimeCreated',
    ITEM_BODY = 'item:Body',
    ITEM_RESPONSE_OBJECTS = 'item:ResponseObjects',
    ITEM_SENSITIVITY = 'item:Sensitivity',
    ITEM_REMINDER_DUE_BY = 'item:ReminderDueBy',
    ITEM_REMINDER_IS_SET = 'item:ReminderIsSet',
    ITEM_REMINDER_NEXT_TIME = 'item:ReminderNextTime',
    ITEM_REMINDER_MINUTES_BEFORE_START = 'item:ReminderMinutesBeforeStart',
    ITEM_DISPLAY_TO = 'item:DisplayTo',
    ITEM_DISPLAY_CC = 'item:DisplayCc',
    ITEM_DISPLAY_BCC = 'item:DisplayBcc',
    ITEM_CULTURE = 'item:Culture',
    ITEM_EFFECTIVE_RIGHTS = 'item:EffectiveRights',
    ITEM_LAST_MODIFIED_NAME = 'item:LastModifiedName',
    ITEM_LAST_MODIFIED_TIME = 'item:LastModifiedTime',
    ITEM_CONVERSATION_ID = 'item:ConversationId',
    ITEM_UNIQUE_BODY = 'item:UniqueBody',
    ITEM_FLAG = 'item:Flag',
    ITEM_STORE_ENTRY_ID = 'item:StoreEntryId',
    ITEM_INSTANCE_KEY = 'item:InstanceKey',
    ITEM_NORMALIZED_BODY = 'item:NormalizedBody',
    ITEM_ENTITY_EXTRACTION_RESULT = 'item:EntityExtractionResult',
    ITEM_POLICY_TAG = 'item:PolicyTag',
    ITEM_ARCHIVE_TAG = 'item:ArchiveTag',
    ITEM_RETENTION_DATE = 'item:RetentionDate',
    ITEM_PREVIEW = 'item:Preview',
    ITEM_PREDICTED_ACTION_REASONS = 'item:PredictedActionReasons',
    ITEM_IS_CLUTTER = 'item:IsClutter',
    ITEM_RIGHTS_MANAGEMENT_LICENSE_DATA = 'item:RightsManagementLicenseData',
    ITEM_BLOCK_STATUS = 'item:BlockStatus',
    ITEM_HAS_BLOCKED_IMAGES = 'item:HasBlockedImages',
    ITEM_WEB_CLIENT_READ_FORM_QUERY_STRING = 'item:WebClientReadFormQueryString',
    ITEM_WEB_CLIENT_EDIT_FORM_QUERY_STRING = 'item:WebClientEditFormQueryString',
    ITEM_TEXT_BODY = 'item:TextBody',
    ITEM_ICON_INDEX = 'item:IconIndex',
    ITEM_MIME_CONTENT_UTF_8 = 'item:MimeContentUTF8',
    ITEM_MENTIONS = 'item:Mentions',
    ITEM_MENTIONED_ME = 'item:MentionedMe',
    ITEM_MENTIONS_PREVIEW = 'item:MentionsPreview',
    ITEM_MENTIONS_EX = 'item:MentionsEx',
    ITEM_HASHTAGS = 'item:Hashtags',
    ITEM_APPLIED_HASHTAGS = 'item:AppliedHashtags',
    ITEM_APPLIED_HASHTAGS_PREVIEW = 'item:AppliedHashtagsPreview',
    ITEM_LIKES = 'item:Likes',
    ITEM_LIKES_PREVIEW = 'item:LikesPreview',
    ITEM_PENDING_SOCIAL_ACTIVITY_TAG_IDS = 'item:PendingSocialActivityTagIds',
    ITEM_AT_ALL_MENTION = 'item:AtAllMention',
    ITEM_CAN_DELETE = 'item:CanDelete',
    ITEM_INFERENCE_CLASSIFICATION = 'item:InferenceClassification',
    ITEM_FIRST_BODY = 'item:FirstBody',
    MESSAGE_CONVERSATION_INDEX = 'message:ConversationIndex',
    MESSAGE_CONVERSATION_TOPIC = 'message:ConversationTopic',
    MESSAGE_INTERNET_MESSAGE_ID = 'message:InternetMessageId',
    MESSAGE_IS_READ = 'message:IsRead',
    MESSAGE_IS_RESPONSE_REQUESTED = 'message:IsResponseRequested',
    MESSAGE_IS_READ_RECEIPT_REQUESTED = 'message:IsReadReceiptRequested',
    MESSAGE_IS_DELIVERY_RECEIPT_REQUESTED = 'message:IsDeliveryReceiptRequested',
    MESSAGE_RECEIVED_BY = 'message:ReceivedBy',
    MESSAGE_RECEIVED_REPRESENTING = 'message:ReceivedRepresenting',
    MESSAGE_REFERENCES = 'message:References',
    MESSAGE_REPLY_TO = 'message:ReplyTo',
    MESSAGE_FROM = 'message:From',
    MESSAGE_SENDER = 'message:Sender',
    MESSAGE_TO_RECIPIENTS = 'message:ToRecipients',
    MESSAGE_CC_RECIPIENTS = 'message:CcRecipients',
    MESSAGE_BCC_RECIPIENTS = 'message:BccRecipients',
    MESSAGE_APPROVAL_REQUEST_DATA = 'message:ApprovalRequestData',
    MESSAGE_VOTING_INFORMATION = 'message:VotingInformation',
    MESSAGE_REMINDER_MESSAGE_DATA = 'message:ReminderMessageData',
    MESSAGE_PUBLISHED_CALENDAR_ITEM_ICS = 'message:PublishedCalendarItemIcs',
    MESSAGE_PUBLISHED_CALENDAR_ITEM_NAME = 'message:PublishedCalendarItemName',
    MESSAGE_MESSAGE_SAFETY = 'message:MessageSafety',
    SHARING_MESSAGE_SHARING_MESSAGE_ACTION = 'sharingMessage:SharingMessageAction',
    MEETING_ASSOCIATED_CALENDAR_ITEM_ID = 'meeting:AssociatedCalendarItemId',
    MEETING_IS_DELEGATED = 'meeting:IsDelegated',
    MEETING_IS_OUT_OF_DATE = 'meeting:IsOutOfDate',
    MEETING_HAS_BEEN_PROCESSED = 'meeting:HasBeenProcessed',
    MEETING_RESPONSE_TYPE = 'meeting:ResponseType',
    MEETING_PROPOSED_START = 'meeting:ProposedStart',
    MEETING_PROPOSED_END = 'meeting:ProposedEnd',
    MEETING_DO_NOT_FORWARD_MEETING = 'meeting:DoNotForwardMeeting',
    MEETING_REQUEST_MEETING_REQUEST_TYPE = 'meetingRequest:MeetingRequestType',
    MEETING_REQUEST_INTENDED_FREE_BUSY_STATUS = 'meetingRequest:IntendedFreeBusyStatus',
    MEETING_REQUEST_CHANGE_HIGHLIGHTS = 'meetingRequest:ChangeHighlights',
    CALENDAR_START = 'calendar:Start',
    CALENDAR_END = 'calendar:End',
    CALENDAR_ORIGINAL_START = 'calendar:OriginalStart',
    CALENDAR_START_WALL_CLOCK = 'calendar:StartWallClock',
    CALENDAR_END_WALL_CLOCK = 'calendar:EndWallClock',
    CALENDAR_START_TIME_ZONE_ID = 'calendar:StartTimeZoneId',
    CALENDAR_END_TIME_ZONE_ID = 'calendar:EndTimeZoneId',
    CALENDAR_IS_ALL_DAY_EVENT = 'calendar:IsAllDayEvent',
    CALENDAR_LEGACY_FREE_BUSY_STATUS = 'calendar:LegacyFreeBusyStatus',
    CALENDAR_LOCATION = 'calendar:Location',
    CALENDAR_ENHANCED_LOCATION = 'calendar:EnhancedLocation',
    CALENDAR_WHEN = 'calendar:When',
    CALENDAR_IS_MEETING = 'calendar:IsMeeting',
    CALENDAR_IS_CANCELLED = 'calendar:IsCancelled',
    CALENDAR_IS_RECURRING = 'calendar:IsRecurring',
    CALENDAR_MEETING_REQUEST_WAS_SENT = 'calendar:MeetingRequestWasSent',
    CALENDAR_IS_RESPONSE_REQUESTED = 'calendar:IsResponseRequested',
    CALENDAR_CALENDAR_ITEM_TYPE = 'calendar:CalendarItemType',
    CALENDAR_MY_RESPONSE_TYPE = 'calendar:MyResponseType',
    CALENDAR_ORGANIZER = 'calendar:Organizer',
    CALENDAR_REQUIRED_ATTENDEES = 'calendar:RequiredAttendees',
    CALENDAR_OPTIONAL_ATTENDEES = 'calendar:OptionalAttendees',
    CALENDAR_RESOURCES = 'calendar:Resources',
    CALENDAR_CONFLICTING_MEETING_COUNT = 'calendar:ConflictingMeetingCount',
    CALENDAR_ADJACENT_MEETING_COUNT = 'calendar:AdjacentMeetingCount',
    CALENDAR_CONFLICTING_MEETINGS = 'calendar:ConflictingMeetings',
    CALENDAR_ADJACENT_MEETINGS = 'calendar:AdjacentMeetings',
    CALENDAR_INBOX_REMINDERS = 'calendar:InboxReminders',
    CALENDAR_DURATION = 'calendar:Duration',
    CALENDAR_TIME_ZONE = 'calendar:TimeZone',
    CALENDAR_APPOINTMENT_REPLY_TIME = 'calendar:AppointmentReplyTime',
    CALENDAR_APPOINTMENT_SEQUENCE_NUMBER = 'calendar:AppointmentSequenceNumber',
    CALENDAR_APPOINTMENT_STATE = 'calendar:AppointmentState',
    CALENDAR_RECURRENCE = 'calendar:Recurrence',
    CALENDAR_FIRST_OCCURRENCE = 'calendar:FirstOccurrence',
    CALENDAR_LAST_OCCURRENCE = 'calendar:LastOccurrence',
    CALENDAR_MODIFIED_OCCURRENCES = 'calendar:ModifiedOccurrences',
    CALENDAR_DELETED_OCCURRENCES = 'calendar:DeletedOccurrences',
    CALENDAR_MEETING_TIME_ZONE = 'calendar:MeetingTimeZone',
    CALENDAR_CONFERENCE_TYPE = 'calendar:ConferenceType',
    CALENDAR_ALLOW_NEW_TIME_PROPOSAL = 'calendar:AllowNewTimeProposal',
    CALENDAR_IS_ONLINE_MEETING = 'calendar:IsOnlineMeeting',
    CALENDAR_MEETING_WORKSPACE_URL = 'calendar:MeetingWorkspaceUrl',
    CALENDAR_NET_SHOW_URL = 'calendar:NetShowUrl',
    CALENDAR_UID = 'calendar:UID',
    CALENDAR_RECURRENCE_ID = 'calendar:RecurrenceId',
    CALENDAR_DATE_TIME_STAMP = 'calendar:DateTimeStamp',
    CALENDAR_START_TIME_ZONE = 'calendar:StartTimeZone',
    CALENDAR_END_TIME_ZONE = 'calendar:EndTimeZone',
    CALENDAR_JOIN_ONLINE_MEETING_URL = 'calendar:JoinOnlineMeetingUrl',
    CALENDAR_ONLINE_MEETING_SETTINGS = 'calendar:OnlineMeetingSettings',
    CALENDAR_IS_ORGANIZER = 'calendar:IsOrganizer',
    CALENDAR_CALENDAR_ACTIVITY_DATA = 'calendar:CalendarActivityData',
    CALENDAR_DO_NOT_FORWARD_MEETING = 'calendar:DoNotForwardMeeting',
    TASK_ACTUAL_WORK = 'task:ActualWork',
    TASK_ASSIGNED_TIME = 'task:AssignedTime',
    TASK_BILLING_INFORMATION = 'task:BillingInformation',
    TASK_CHANGE_COUNT = 'task:ChangeCount',
    TASK_COMPANIES = 'task:Companies',
    TASK_COMPLETE_DATE = 'task:CompleteDate',
    TASK_CONTACTS = 'task:Contacts',
    TASK_DELEGATION_STATE = 'task:DelegationState',
    TASK_DELEGATOR = 'task:Delegator',
    TASK_DUE_DATE = 'task:DueDate',
    TASK_IS_ASSIGNMENT_EDITABLE = 'task:IsAssignmentEditable',
    TASK_IS_COMPLETE = 'task:IsComplete',
    TASK_IS_RECURRING = 'task:IsRecurring',
    TASK_IS_TEAM_TASK = 'task:IsTeamTask',
    TASK_MILEAGE = 'task:Mileage',
    TASK_OWNER = 'task:Owner',
    TASK_PERCENT_COMPLETE = 'task:PercentComplete',
    TASK_RECURRENCE = 'task:Recurrence',
    TASK_START_DATE = 'task:StartDate',
    TASK_STATUS = 'task:Status',
    TASK_STATUS_DESCRIPTION = 'task:StatusDescription',
    TASK_TOTAL_WORK = 'task:TotalWork',
    CONTACTS_ALIAS = 'contacts:Alias',
    CONTACTS_ASSISTANT_NAME = 'contacts:AssistantName',
    CONTACTS_BIRTHDAY = 'contacts:Birthday',
    CONTACTS_BUSINESS_HOME_PAGE = 'contacts:BusinessHomePage',
    CONTACTS_CHILDREN = 'contacts:Children',
    CONTACTS_COMPANIES = 'contacts:Companies',
    CONTACTS_COMPANY_NAME = 'contacts:CompanyName',
    CONTACTS_COMPLETE_NAME = 'contacts:CompleteName',
    CONTACTS_CONTACT_SOURCE = 'contacts:ContactSource',
    CONTACTS_CULTURE = 'contacts:Culture',
    CONTACTS_DEPARTMENT = 'contacts:Department',
    CONTACTS_DISPLAY_NAME = 'contacts:DisplayName',
    CONTACTS_DIRECTORY_ID = 'contacts:DirectoryId',
    CONTACTS_DIRECT_REPORTS = 'contacts:DirectReports',
    CONTACTS_EMAIL_ADDRESSES = 'contacts:EmailAddresses',
    CONTACTS_ABCH_EMAIL_ADDRESSES = 'contacts:AbchEmailAddresses',
    CONTACTS_FILE_AS = 'contacts:FileAs',
    CONTACTS_FILE_AS_MAPPING = 'contacts:FileAsMapping',
    CONTACTS_GENERATION = 'contacts:Generation',
    CONTACTS_GIVEN_NAME = 'contacts:GivenName',
    CONTACTS_IM_ADDRESSES = 'contacts:ImAddresses',
    CONTACTS_INITIALS = 'contacts:Initials',
    CONTACTS_JOB_TITLE = 'contacts:JobTitle',
    CONTACTS_MANAGER = 'contacts:Manager',
    CONTACTS_MANAGER_MAILBOX = 'contacts:ManagerMailbox',
    CONTACTS_MIDDLE_NAME = 'contacts:MiddleName',
    CONTACTS_MILEAGE = 'contacts:Mileage',
    CONTACTS_MS_EXCHANGE_CERTIFICATE = 'contacts:MSExchangeCertificate',
    CONTACTS_NICKNAME = 'contacts:Nickname',
    CONTACTS_NOTES = 'contacts:Notes',
    CONTACTS_OFFICE_LOCATION = 'contacts:OfficeLocation',
    CONTACTS_PHONE_NUMBERS = 'contacts:PhoneNumbers',
    CONTACTS_PHONETIC_FULL_NAME = 'contacts:PhoneticFullName',
    CONTACTS_PHONETIC_FIRST_NAME = 'contacts:PhoneticFirstName',
    CONTACTS_PHONETIC_LAST_NAME = 'contacts:PhoneticLastName',
    CONTACTS_PHOTO = 'contacts:Photo',
    CONTACTS_PHYSICAL_ADDRESSES = 'contacts:PhysicalAddresses',
    CONTACTS_POSTAL_ADDRESS_INDEX = 'contacts:PostalAddressIndex',
    CONTACTS_PROFESSION = 'contacts:Profession',
    CONTACTS_SPOUSE_NAME = 'contacts:SpouseName',
    CONTACTS_SURNAME = 'contacts:Surname',
    CONTACTS_WEDDING_ANNIVERSARY = 'contacts:WeddingAnniversary',
    CONTACTS_USER_SMIME_CERTIFICATE = 'contacts:UserSMIMECertificate',
    CONTACTS_HAS_PICTURE = 'contacts:HasPicture',
    CONTACTS_ACCOUNT_NAME = 'contacts:AccountName',
    CONTACTS_IS_AUTO_UPDATE_DISABLED = 'contacts:IsAutoUpdateDisabled',
    CONTACTS_IS_MESSENGER_ENABLED = 'contacts:IsMessengerEnabled',
    CONTACTS_COMMENT = 'contacts:Comment',
    CONTACTS_CONTACT_SHORT_ID = 'contacts:ContactShortId',
    CONTACTS_CONTACT_TYPE = 'contacts:ContactType',
    CONTACTS_CREATED_BY = 'contacts:CreatedBy',
    CONTACTS_GENDER = 'contacts:Gender',
    CONTACTS_IS_HIDDEN = 'contacts:IsHidden',
    CONTACTS_OBJECT_ID = 'contacts:ObjectId',
    CONTACTS_PASSPORT_ID = 'contacts:PassportId',
    CONTACTS_IS_PRIVATE = 'contacts:IsPrivate',
    CONTACTS_SOURCE_ID = 'contacts:SourceId',
    CONTACTS_TRUST_LEVEL = 'contacts:TrustLevel',
    CONTACTS_URLS = 'contacts:Urls',
    CONTACTS_CID = 'contacts:Cid',
    CONTACTS_SKYPE_AUTH_CERTIFICATE = 'contacts:SkypeAuthCertificate',
    CONTACTS_SKYPE_CONTEXT = 'contacts:SkypeContext',
    CONTACTS_SKYPE_ID = 'contacts:SkypeId',
    CONTACTS_XBOX_LIVE_TAG = 'contacts:XboxLiveTag',
    CONTACTS_SKYPE_RELATIONSHIP = 'contacts:SkypeRelationship',
    CONTACTS_YOMI_NICKNAME = 'contacts:YomiNickname',
    CONTACTS_INVITE_FREE = 'contacts:InviteFree',
    CONTACTS_HIDE_PRESENCE_AND_PROFILE = 'contacts:HidePresenceAndProfile',
    CONTACTS_IS_PENDING_OUTBOUND = 'contacts:IsPendingOutbound',
    CONTACTS_SUPPORT_GROUP_FEEDS = 'contacts:SupportGroupFeeds',
    CONTACTS_USER_TILE_HASH = 'contacts:UserTileHash',
    CONTACTS_UNIFIED_INBOX = 'contacts:UnifiedInbox',
    CONTACTS_MRIS = 'contacts:Mris',
    CONTACTS_WLID = 'contacts:Wlid',
    CONTACTS_ABCH_CONTACT_ID = 'contacts:AbchContactId',
    CONTACTS_NOT_IN_BIRTHDAY_CALENDAR = 'contacts:NotInBirthdayCalendar',
    CONTACTS_SHELL_CONTACT_TYPE = 'contacts:ShellContactType',
    CONTACTS_IM_MRI = 'contacts:ImMri',
    CONTACTS_PRESENCE_TRUST_LEVEL = 'contacts:PresenceTrustLevel',
    CONTACTS_OTHER_MRI = 'contacts:OtherMri',
    CONTACTS_PROFILE_LAST_CHANGED = 'contacts:ProfileLastChanged',
    CONTACTS_MOBILE_IM_ENABLED = 'contacts:MobileIMEnabled',
    DISTRIBUTIONLIST_MEMBERS = 'distributionlist:Members',
    CONTACTS_PARTNER_NETWORK_PROFILE_PHOTO_URL = 'contacts:PartnerNetworkProfilePhotoUrl',
    CONTACTS_PARTNER_NETWORK_THUMBNAIL_PHOTO_URL = 'contacts:PartnerNetworkThumbnailPhotoUrl',
    CONTACTS_PERSON_ID = 'contacts:PersonId',
    CONTACTS_CONVERSATION_GUID = 'contacts:ConversationGuid',
    POSTITEM_POSTED_TIME = 'postitem:PostedTime',
    CONVERSATION_CONVERSATION_ID = 'conversation:ConversationId',
    CONVERSATION_CONVERSATION_TOPIC = 'conversation:ConversationTopic',
    CONVERSATION_UNIQUE_RECIPIENTS = 'conversation:UniqueRecipients',
    CONVERSATION_GLOBAL_UNIQUE_RECIPIENTS = 'conversation:GlobalUniqueRecipients',
    CONVERSATION_UNIQUE_UNREAD_SENDERS = 'conversation:UniqueUnreadSenders',
    CONVERSATION_GLOBAL_UNIQUE_UNREAD_SENDERS = 'conversation:GlobalUniqueUnreadSenders',
    CONVERSATION_UNIQUE_SENDERS = 'conversation:UniqueSenders',
    CONVERSATION_GLOBAL_UNIQUE_SENDERS = 'conversation:GlobalUniqueSenders',
    CONVERSATION_LAST_DELIVERY_TIME = 'conversation:LastDeliveryTime',
    CONVERSATION_GLOBAL_LAST_DELIVERY_TIME = 'conversation:GlobalLastDeliveryTime',
    CONVERSATION_CATEGORIES = 'conversation:Categories',
    CONVERSATION_GLOBAL_CATEGORIES = 'conversation:GlobalCategories',
    CONVERSATION_FLAG_STATUS = 'conversation:FlagStatus',
    CONVERSATION_GLOBAL_FLAG_STATUS = 'conversation:GlobalFlagStatus',
    CONVERSATION_HAS_ATTACHMENTS = 'conversation:HasAttachments',
    CONVERSATION_GLOBAL_HAS_ATTACHMENTS = 'conversation:GlobalHasAttachments',
    CONVERSATION_HAS_IRM = 'conversation:HasIrm',
    CONVERSATION_GLOBAL_HAS_IRM = 'conversation:GlobalHasIrm',
    CONVERSATION_MESSAGE_COUNT = 'conversation:MessageCount',
    CONVERSATION_GLOBAL_MESSAGE_COUNT = 'conversation:GlobalMessageCount',
    CONVERSATION_UNREAD_COUNT = 'conversation:UnreadCount',
    CONVERSATION_GLOBAL_UNREAD_COUNT = 'conversation:GlobalUnreadCount',
    CONVERSATION_SIZE = 'conversation:Size',
    CONVERSATION_GLOBAL_SIZE = 'conversation:GlobalSize',
    CONVERSATION_ITEM_CLASSES = 'conversation:ItemClasses',
    CONVERSATION_GLOBAL_ITEM_CLASSES = 'conversation:GlobalItemClasses',
    CONVERSATION_IMPORTANCE = 'conversation:Importance',
    CONVERSATION_GLOBAL_IMPORTANCE = 'conversation:GlobalImportance',
    CONVERSATION_ITEM_IDS = 'conversation:ItemIds',
    CONVERSATION_GLOBAL_ITEM_IDS = 'conversation:GlobalItemIds',
    CONVERSATION_LAST_MODIFIED_TIME = 'conversation:LastModifiedTime',
    CONVERSATION_INSTANCE_KEY = 'conversation:InstanceKey',
    CONVERSATION_PREVIEW = 'conversation:Preview',
    CONVERSATION_ICON_INDEX = 'conversation:IconIndex',
    CONVERSATION_GLOBAL_ICON_INDEX = 'conversation:GlobalIconIndex',
    CONVERSATION_DRAFT_ITEM_IDS = 'conversation:DraftItemIds',
    CONVERSATION_HAS_CLUTTER = 'conversation:HasClutter',
    CONVERSATION_MENTIONED_ME = 'conversation:MentionedMe',
    CONVERSATION_GLOBAL_MENTIONED_ME = 'conversation:GlobalMentionedMe',
    CONVERSATION_AT_ALL_MENTION = 'conversation:AtAllMention',
    CONVERSATION_GLOBAL_AT_ALL_MENTION = 'conversation:GlobalAtAllMention',
    PERSON_FULL_NAME = 'person:FullName',
    PERSON_GIVEN_NAME = 'person:GivenName',
    PERSON_SURNAME = 'person:Surname',
    PERSON_PHONE_NUMBER = 'person:PhoneNumber',
    PERSON_SMS_NUMBER = 'person:SMSNumber',
    PERSON_EMAIL_ADDRESS = 'person:EmailAddress',
    PERSON_ALIAS = 'person:Alias',
    PERSON_DEPARTMENT = 'person:Department',
    PERSON_LINKED_IN_PROFILE_LINK = 'person:LinkedInProfileLink',
    PERSON_SKILLS = 'person:Skills',
    PERSON_PROFESSIONAL_BIOGRAPHY = 'person:ProfessionalBiography',
    PERSON_MANAGEMENT_CHAIN = 'person:ManagementChain',
    PERSON_DIRECT_REPORTS = 'person:DirectReports',
    PERSON_PEERS = 'person:Peers',
    PERSON_TEAM_SIZE = 'person:TeamSize',
    PERSON_CURRENT_JOB = 'person:CurrentJob',
    PERSON_BIRTHDAY = 'person:Birthday',
    PERSON_HOMETOWN = 'person:Hometown',
    PERSON_CURRENT_LOCATION = 'person:CurrentLocation',
    PERSON_COMPANY_PROFILE = 'person:CompanyProfile',
    PERSON_OFFICE = 'person:Office',
    PERSON_HEADLINE = 'person:Headline',
    PERSON_MUTUAL_CONNECTIONS = 'person:MutualConnections',
    PERSON_TITLE = 'person:Title',
    PERSON_MUTUAL_MANAGER = 'person:MutualManager',
    PERSON_INSIGHTS = 'person:Insights',
    PERSON_USER_PROFILE_PICTURE = 'person:UserProfilePicture',
    PERSONA_PERSONA_ID = 'persona:PersonaId',
    PERSONA_PERSONA_TYPE = 'persona:PersonaType',
    PERSONA_GIVEN_NAME = 'persona:GivenName',
    PERSONA_COMPANY_NAME = 'persona:CompanyName',
    PERSONA_SURNAME = 'persona:Surname',
    PERSONA_DISPLAY_NAME = 'persona:DisplayName',
    PERSONA_EMAIL_ADDRESS = 'persona:EmailAddress',
    PERSONA_FILE_AS = 'persona:FileAs',
    PERSONA_HOME_CITY = 'persona:HomeCity',
    PERSONA_CREATION_TIME = 'persona:CreationTime',
    PERSONA_RELEVANCE_SCORE = 'persona:RelevanceScore',
    PERSONA_RANKING_WEIGHT = 'persona:RankingWeight',
    PERSONA_WORK_CITY = 'persona:WorkCity',
    PERSONA_PERSONA_OBJECT_STATUS = 'persona:PersonaObjectStatus',
    PERSONA_FILE_AS_ID = 'persona:FileAsId',
    PERSONA_DISPLAY_NAME_PREFIX = 'persona:DisplayNamePrefix',
    PERSONA_YOMI_COMPANY_NAME = 'persona:YomiCompanyName',
    PERSONA_YOMI_FIRST_NAME = 'persona:YomiFirstName',
    PERSONA_YOMI_LAST_NAME = 'persona:YomiLastName',
    PERSONA_TITLE = 'persona:Title',
    PERSONA_EMAIL_ADDRESSES = 'persona:EmailAddresses',
    PERSONA_PHONE_NUMBER = 'persona:PhoneNumber',
    PERSONA_IM_ADDRESS = 'persona:ImAddress',
    PERSONA_IM_ADDRESSES = 'persona:ImAddresses',
    PERSONA_IM_ADDRESSES_2 = 'persona:ImAddresses2',
    PERSONA_IM_ADDRESSES_3 = 'persona:ImAddresses3',
    PERSONA_FOLDER_IDS = 'persona:FolderIds',
    PERSONA_ATTRIBUTIONS = 'persona:Attributions',
    PERSONA_DISPLAY_NAMES = 'persona:DisplayNames',
    PERSONA_INITIALS = 'persona:Initials',
    PERSONA_FILE_ASES = 'persona:FileAses',
    PERSONA_FILE_AS_IDS = 'persona:FileAsIds',
    PERSONA_DISPLAY_NAME_PREFIXES = 'persona:DisplayNamePrefixes',
    PERSONA_GIVEN_NAMES = 'persona:GivenNames',
    PERSONA_MIDDLE_NAMES = 'persona:MiddleNames',
    PERSONA_SURNAMES = 'persona:Surnames',
    PERSONA_GENERATIONS = 'persona:Generations',
    PERSONA_NICKNAMES = 'persona:Nicknames',
    PERSONA_YOMI_COMPANY_NAMES = 'persona:YomiCompanyNames',
    PERSONA_YOMI_FIRST_NAMES = 'persona:YomiFirstNames',
    PERSONA_YOMI_LAST_NAMES = 'persona:YomiLastNames',
    PERSONA_BUSINESS_PHONE_NUMBERS = 'persona:BusinessPhoneNumbers',
    PERSONA_BUSINESS_PHONE_NUMBERS_2 = 'persona:BusinessPhoneNumbers2',
    PERSONA_HOME_PHONES = 'persona:HomePhones',
    PERSONA_HOME_PHONES_2 = 'persona:HomePhones2',
    PERSONA_MOBILE_PHONES = 'persona:MobilePhones',
    PERSONA_MOBILE_PHONES_2 = 'persona:MobilePhones2',
    PERSONA_ASSISTANT_PHONE_NUMBERS = 'persona:AssistantPhoneNumbers',
    PERSONA_CALLBACK_PHONES = 'persona:CallbackPhones',
    PERSONA_CAR_PHONES = 'persona:CarPhones',
    PERSONA_HOME_FAXES = 'persona:HomeFaxes',
    PERSONA_ORGANIZATION_MAIN_PHONES = 'persona:OrganizationMainPhones',
    PERSONA_OTHER_FAXES = 'persona:OtherFaxes',
    PERSONA_OTHER_TELEPHONES = 'persona:OtherTelephones',
    PERSONA_OTHER_PHONES_2 = 'persona:OtherPhones2',
    PERSONA_PAGERS = 'persona:Pagers',
    PERSONA_RADIO_PHONES = 'persona:RadioPhones',
    PERSONA_TELEX_NUMBERS = 'persona:TelexNumbers',
    PERSONA_WORK_FAXES = 'persona:WorkFaxes',
    PERSONA_EMAILS_1 = 'persona:Emails1',
    PERSONA_EMAILS_2 = 'persona:Emails2',
    PERSONA_EMAILS_3 = 'persona:Emails3',
    PERSONA_BUSINESS_HOME_PAGES = 'persona:BusinessHomePages',
    PERSONA_SCHOOL = 'persona:School',
    PERSONA_PERSONAL_HOME_PAGES = 'persona:PersonalHomePages',
    PERSONA_OFFICE_LOCATIONS = 'persona:OfficeLocations',
    PERSONA_BUSINESS_ADDRESSES = 'persona:BusinessAddresses',
    PERSONA_HOME_ADDRESSES = 'persona:HomeAddresses',
    PERSONA_OTHER_ADDRESSES = 'persona:OtherAddresses',
    PERSONA_TITLES = 'persona:Titles',
    PERSONA_DEPARTMENTS = 'persona:Departments',
    PERSONA_COMPANY_NAMES = 'persona:CompanyNames',
    PERSONA_MANAGERS = 'persona:Managers',
    PERSONA_ASSISTANT_NAMES = 'persona:AssistantNames',
    PERSONA_PROFESSIONS = 'persona:Professions',
    PERSONA_SPOUSE_NAMES = 'persona:SpouseNames',
    PERSONA_HOBBIES = 'persona:Hobbies',
    PERSONA_WEDDING_ANNIVERSARIES = 'persona:WeddingAnniversaries',
    PERSONA_BIRTHDAYS = 'persona:Birthdays',
    PERSONA_CHILDREN = 'persona:Children',
    PERSONA_LOCATIONS = 'persona:Locations',
    PERSONA_EXTENDED_PROPERTIES = 'persona:ExtendedProperties',
    PERSONA_POSTAL_ADDRESS = 'persona:PostalAddress',
    PERSONA_BODIES = 'persona:Bodies',
    PERSONA_IS_FAVORITE = 'persona:IsFavorite',
    PERSONA_INLINE_LINKS = 'persona:InlineLinks',
    PERSONA_ITEM_LINK_IDS = 'persona:ItemLinkIds',
    PERSONA_HAS_ACTIVE_DEALS = 'persona:HasActiveDeals',
    PERSONA_IS_BUSINESS_CONTACT = 'persona:IsBusinessContact',
    PERSONA_ATTRIBUTED_HAS_ACTIVE_DEALS = 'persona:AttributedHasActiveDeals',
    PERSONA_ATTRIBUTED_IS_BUSINESS_CONTACT = 'persona:AttributedIsBusinessContact',
    PERSONA_SOURCE_MAILBOX_GUIDS = 'persona:SourceMailboxGuids',
    PERSONA_LAST_CONTACTED_DATE = 'persona:LastContactedDate',
    PERSONA_EXTERNAL_DIRECTORY_OBJECT_ID = 'persona:ExternalDirectoryObjectId',
    PERSONA_MAPI_ENTRY_ID = 'persona:MapiEntryId',
    PERSONA_MAPI_EMAIL_ADDRESS = 'persona:MapiEmailAddress',
    PERSONA_MAPI_ADDRESS_TYPE = 'persona:MapiAddressType',
    PERSONA_MAPI_SEARCH_KEY = 'persona:MapiSearchKey',
    PERSONA_MAPI_TRANSMITTABLE_DISPLAY_NAME = 'persona:MapiTransmittableDisplayName',
    PERSONA_MAPI_SEND_RICH_INFO = 'persona:MapiSendRichInfo',
    ROLEMEMBER_MEMBER_TYPE = 'rolemember:MemberType',
    ROLEMEMBER_MEMBER_ID = 'rolemember:MemberId',
    ROLEMEMBER_DISPLAY_NAME = 'rolemember:DisplayName',
    NETWORK_TOKEN_REFRESH_LAST_COMPLETED = 'network:TokenRefreshLastCompleted',
    NETWORK_TOKEN_REFRESH_LAST_ATTEMPTED = 'network:TokenRefreshLastAttempted',
    NETWORK_SYNC_ENABLED = 'network:SyncEnabled',
    NETWORK_REJECTED_OFFERS = 'network:RejectedOffers',
    NETWORK_SESSION_HANDLE = 'network:SessionHandle',
    NETWORK_REFRESH_TOKEN_EXPIRY_2 = 'network:RefreshTokenExpiry2',
    NETWORK_REFRESH_TOKEN_2 = 'network:RefreshToken2',
    NETWORK_PSA_LAST_CHANGED = 'network:PsaLastChanged',
    NETWORK_OFFERS = 'network:Offers',
    NETWORK_LAST_WELCOME_CONTACT = 'network:LastWelcomeContact',
    NETWORK_LAST_VERSION_SAVED = 'network:LastVersionSaved',
    NETWORK_DOMAIN_TAG = 'network:DomainTag',
    NETWORK_FIRST_AUTH_ERROR_DATES = 'network:FirstAuthErrorDates',
    NETWORK_ERROR_OFFERS = 'network:ErrorOffers',
    NETWORK_CONTACT_SYNC_SUCCESS = 'network:ContactSyncSuccess',
    NETWORK_CONTACT_SYNC_ERROR = 'network:ContactSyncError',
    NETWORK_CLIENT_TOKEN_2 = 'network:ClientToken2',
    NETWORK_CLIENT_TOKEN = 'network:ClientToken',
    NETWORK_CLIENT_PUBLISH_SECRET = 'network:ClientPublishSecret',
    NETWORK_USER_EMAIL = 'network:UserEmail',
    NETWORK_AUTO_LINK_SUCCESS = 'network:AutoLinkSuccess',
    NETWORK_AUTO_LINK_ERROR = 'network:AutoLinkError',
    NETWORK_IS_DEFAULT = 'network:IsDefault',
    NETWORK_SETTINGS = 'network:Settings',
    NETWORK_PROFILE_URL = 'network:ProfileUrl',
    NETWORK_USER_TILE_URL = 'network:UserTileUrl',
    NETWORK_DOMAIN_ID = 'network:DomainId',
    NETWORK_DISPLAY_NAME = 'network:DisplayName',
    NETWORK_ACCOUNT_NAME = 'network:AccountName',
    NETWORK_SOURCE_ENTRY_ID = 'network:SourceEntryID',
    ABCHPERSON_FAVORITE_ORDER = 'abchperson:FavoriteOrder',
    ABCHPERSON_PERSON_ID = 'abchperson:PersonId',
    ABCHPERSON_EXCHANGE_PERSON_ID_GUID = 'abchperson:ExchangePersonIdGuid',
    ABCHPERSON_ANTI_LINK_INFO = 'abchperson:AntiLinkInfo',
    ABCHPERSON_RELEVANCE_ORDER_1 = 'abchperson:RelevanceOrder1',
    ABCHPERSON_RELEVANCE_ORDER_2 = 'abchperson:RelevanceOrder2',
    ABCHPERSON_CONTACT_HANDLES = 'abchperson:ContactHandles',
    ABCHPERSON_CATEGORIES = 'abchperson:Categories',
    BOOKING_SERVICE_IDS = 'booking:ServiceIds',
    BOOKING_STAFF_IDS = 'booking:StaffIds',
    BOOKING_STAFF_INITIALS = 'booking:StaffInitials',
    BOOKING_CUSTOMER_NAME = 'booking:CustomerName',
    BOOKING_CUSTOMER_EMAIL = 'booking:CustomerEmail',
    BOOKING_CUSTOMER_PHONE = 'booking:CustomerPhone',
    BOOKING_CUSTOMER_ID = 'booking:CustomerId',
    INSIGHT_INSIGHT_ID = 'insight:InsightId',
    INSIGHT_TYPE = 'insight:Type',
    INSIGHT_START_TIME_UTC = 'insight:StartTimeUtc',
    INSIGHT_END_TIME_UTC = 'insight:EndTimeUtc',
    INSIGHT_STATUS = 'insight:Status',
    INSIGHT_VERSION = 'insight:Version',
    INSIGHT_APPLICATIONS_IDS = 'insight:ApplicationsIds',
    INSIGHT_TEXT = 'insight:Text',
    INSIGHT_SUGGESTED_ACTIONS = 'insight:SuggestedActions',
    INSIGHT_APP_CONTEXTS = 'insight:AppContexts',
}

/**
 * Defines the keys ExtendedFieldURI attributes
 */
export const enum ExtendedPropertyKeyType {
    /** Property name key for ExtendedFieldURI attribute */
    PROPERTY_NAME = 'PropertyName',
    /** Property set id key for ExtendedFieldURI attribute */
    PROPERTY_SETID = 'PropertySetId',
    /** Distinguished property set id key for ExtendedFieldURI attribute */
    DISTINGUISHED_PROPERTY_SETID = 'DistinguishedPropertySetId',
    /** Property tag key for ExtendedFieldURI attribute */
    PROPERTY_TAG = 'PropertyTag',
    /** Property id key for ExtendedFieldURI attribute */
    PROPERTY_ID = 'PropertyId',
    /** Property type key for ExtendedFieldURI attribute */
    PROPERTY_TYPE = 'PropertyType' 
}

/**
 * Defines the rule field URI.
 */
export const enum RuleFieldURIType {
    RULE_ID = 'RuleId',
    DISPLAY_NAME = 'DisplayName',
    PRIORITY = 'Priority',
    IS_NOT_SUPPORTED = 'IsNotSupported',
    ACTIONS = 'Actions',
    CONDITION_CATEGORIES = 'Condition:Categories',
    CONDITION_CONTAINS_BODY_STRINGS = 'Condition:ContainsBodyStrings',
    CONDITION_CONTAINS_HEADER_STRINGS = 'Condition:ContainsHeaderStrings',
    CONDITION_CONTAINS_RECIPIENT_STRINGS = 'Condition:ContainsRecipientStrings',
    CONDITION_CONTAINS_SENDER_STRINGS = 'Condition:ContainsSenderStrings',
    CONDITION_CONTAINS_SUBJECT_OR_BODY_STRINGS = 'Condition:ContainsSubjectOrBodyStrings',
    CONDITION_CONTAINS_SUBJECT_STRINGS = 'Condition:ContainsSubjectStrings',
    CONDITION_FLAGGED_FOR_ACTION = 'Condition:FlaggedForAction',
    CONDITION_FROM_ADDRESSES = 'Condition:FromAddresses',
    CONDITION_FROM_CONNECTED_ACCOUNTS = 'Condition:FromConnectedAccounts',
    CONDITION_HAS_ATTACHMENTS = 'Condition:HasAttachments',
    CONDITION_IMPORTANCE = 'Condition:Importance',
    CONDITION_IS_APPROVAL_REQUEST = 'Condition:IsApprovalRequest',
    CONDITION_IS_AUTOMATIC_FORWARD = 'Condition:IsAutomaticForward',
    CONDITION_IS_AUTOMATIC_REPLY = 'Condition:IsAutomaticReply',
    CONDITION_IS_ENCRYPTED = 'Condition:IsEncrypted',
    CONDITION_IS_MEETING_REQUEST = 'Condition:IsMeetingRequest',
    CONDITION_IS_MEETING_RESPONSE = 'Condition:IsMeetingResponse',
    CONDITION_IS_NDR = 'Condition:IsNDR',
    CONDITION_IS_PERMISSION_CONTROLLED = 'Condition:IsPermissionControlled',
    CONDITION_IS_READ_RECEIPT = 'Condition:IsReadReceipt',
    CONDITION_IS_SIGNED = 'Condition:IsSigned',
    CONDITION_IS_VOICEMAIL = 'Condition:IsVoicemail',
    CONDITION_ITEM_CLASSES = 'Condition:ItemClasses',
    CONDITION_MESSAGE_CLASSIFICATIONS = 'Condition:MessageClassifications',
    CONDITION_NOT_SENT_TO_ME = 'Condition:NotSentToMe',
    CONDITION_SENT_CC_ME = 'Condition:SentCcMe',
    CONDITION_SENT_ONLY_TO_ME = 'Condition:SentOnlyToMe',
    CONDITION_SENT_TO_ADDRESSES = 'Condition:SentToAddresses',
    CONDITION_SENT_TO_ME = 'Condition:SentToMe',
    CONDITION_SENT_TO_OR_CC_ME = 'Condition:SentToOrCcMe',
    CONDITION_SENSITIVITY = 'Condition:Sensitivity',
    CONDITION_WITHIN_DATE_RANGE = 'Condition:WithinDateRange',
    CONDITION_WITHIN_SIZE_RANGE = 'Condition:WithinSizeRange',
    EXCEPTION_CATEGORIES = 'Exception:Categories',
    EXCEPTION_CONTAINS_BODY_STRINGS = 'Exception:ContainsBodyStrings',
    EXCEPTION_CONTAINS_HEADER_STRINGS = 'Exception:ContainsHeaderStrings',
    EXCEPTION_CONTAINS_RECIPIENT_STRINGS = 'Exception:ContainsRecipientStrings',
    EXCEPTION_CONTAINS_SENDER_STRINGS = 'Exception:ContainsSenderStrings',
    EXCEPTION_CONTAINS_SUBJECT_OR_BODY_STRINGS = 'Exception:ContainsSubjectOrBodyStrings',
    EXCEPTION_CONTAINS_SUBJECT_STRINGS = 'Exception:ContainsSubjectStrings',
    EXCEPTION_FLAGGED_FOR_ACTION = 'Exception:FlaggedForAction',
    EXCEPTION_FROM_ADDRESSES = 'Exception:FromAddresses',
    EXCEPTION_FROM_CONNECTED_ACCOUNTS = 'Exception:FromConnectedAccounts',
    EXCEPTION_HAS_ATTACHMENTS = 'Exception:HasAttachments',
    EXCEPTION_IMPORTANCE = 'Exception:Importance',
    EXCEPTION_IS_APPROVAL_REQUEST = 'Exception:IsApprovalRequest',
    EXCEPTION_IS_AUTOMATIC_FORWARD = 'Exception:IsAutomaticForward',
    EXCEPTION_IS_AUTOMATIC_REPLY = 'Exception:IsAutomaticReply',
    EXCEPTION_IS_ENCRYPTED = 'Exception:IsEncrypted',
    EXCEPTION_IS_MEETING_REQUEST = 'Exception:IsMeetingRequest',
    EXCEPTION_IS_MEETING_RESPONSE = 'Exception:IsMeetingResponse',
    EXCEPTION_IS_NDR = 'Exception:IsNDR',
    EXCEPTION_IS_PERMISSION_CONTROLLED = 'Exception:IsPermissionControlled',
    EXCEPTION_IS_READ_RECEIPT = 'Exception:IsReadReceipt',
    EXCEPTION_IS_SIGNED = 'Exception:IsSigned',
    EXCEPTION_IS_VOICEMAIL = 'Exception:IsVoicemail',
    EXCEPTION_ITEM_CLASSES = 'Exception:ItemClasses',
    EXCEPTION_MESSAGE_CLASSIFICATIONS = 'Exception:MessageClassifications',
    EXCEPTION_NOT_SENT_TO_ME = 'Exception:NotSentToMe',
    EXCEPTION_SENT_CC_ME = 'Exception:SentCcMe',
    EXCEPTION_SENT_ONLY_TO_ME = 'Exception:SentOnlyToMe',
    EXCEPTION_SENT_TO_ADDRESSES = 'Exception:SentToAddresses',
    EXCEPTION_SENT_TO_ME = 'Exception:SentToMe',
    EXCEPTION_SENT_TO_OR_CC_ME = 'Exception:SentToOrCcMe',
    EXCEPTION_SENSITIVITY = 'Exception:Sensitivity',
    EXCEPTION_WITHIN_DATE_RANGE = 'Exception:WithinDateRange',
    EXCEPTION_WITHIN_SIZE_RANGE = 'Exception:WithinSizeRange',
    ACTION_ASSIGN_CATEGORIES = 'Action:AssignCategories',
    ACTION_COPY_TO_FOLDER = 'Action:CopyToFolder',
    ACTION_DELETE = 'Action:Delete',
    ACTION_FORWARD_AS_ATTACHMENT_TO_RECIPIENTS = 'Action:ForwardAsAttachmentToRecipients',
    ACTION_FORWARD_TO_RECIPIENTS = 'Action:ForwardToRecipients',
    ACTION_MARK_IMPORTANCE = 'Action:MarkImportance',
    ACTION_MARK_AS_READ = 'Action:MarkAsRead',
    ACTION_MOVE_TO_FOLDER = 'Action:MoveToFolder',
    ACTION_PERMANENT_DELETE = 'Action:PermanentDelete',
    ACTION_REDIRECT_TO_RECIPIENTS = 'Action:RedirectToRecipients',
    ACTION_SEND_SMS_ALERT_TO_RECIPIENTS = 'Action:SendSMSAlertToRecipients',
    ACTION_SERVER_REPLY_WITH_MESSAGE = 'Action:ServerReplyWithMessage',
    ACTION_STOP_PROCESSING_RULES = 'Action:StopProcessingRules',
    IS_ENABLED = 'IsEnabled',
    IS_IN_ERROR = 'IsInError',
    CONDITIONS = 'Conditions',
    EXCEPTIONS = 'Exceptions',
}

/**
 * Defines the rule validation error code.
 */
export const enum RuleValidationErrorCodeType {
    AD_OPERATION_FAILURE = 'ADOperationFailure',
    CONNECTED_ACCOUNT_NOT_FOUND = 'ConnectedAccountNotFound',
    CREATE_WITH_RULE_ID = 'CreateWithRuleId',
    EMPTY_VALUE_FOUND = 'EmptyValueFound',
    DUPLICATED_PRIORITY = 'DuplicatedPriority',
    DUPLICATED_OPERATION_ON_THE_SAME_RULE = 'DuplicatedOperationOnTheSameRule',
    FOLDER_DOES_NOT_EXIST = 'FolderDoesNotExist',
    INVALID_ADDRESS = 'InvalidAddress',
    INVALID_DATE_RANGE = 'InvalidDateRange',
    INVALID_FOLDER_ID = 'InvalidFolderId',
    INVALID_SIZE_RANGE = 'InvalidSizeRange',
    INVALID_VALUE = 'InvalidValue',
    MESSAGE_CLASSIFICATION_NOT_FOUND = 'MessageClassificationNotFound',
    MISSING_ACTION = 'MissingAction',
    MISSING_PARAMETER = 'MissingParameter',
    MISSING_RANGE_VALUE = 'MissingRangeValue',
    NOT_SETTABLE = 'NotSettable',
    RECIPIENT_DOES_NOT_EXIST = 'RecipientDoesNotExist',
    RULE_NOT_FOUND = 'RuleNotFound',
    SIZE_LESS_THAN_ZERO = 'SizeLessThanZero',
    STRING_VALUE_TOO_BIG = 'StringValueTooBig',
    UNSUPPORTED_ADDRESS = 'UnsupportedAddress',
    UNEXPECTED_ERROR = 'UnexpectedError',
    UNSUPPORTED_RULE = 'UnsupportedRule',
}

/**
 * List of response codes.
 */
export const enum ResponseCodeType {
    NO_ERROR = 'NoError',
    ERROR_ACCESS_DENIED = 'ErrorAccessDenied',
    ERROR_ACCESS_MODE_SPECIFIED = 'ErrorAccessModeSpecified',
    ERROR_ACCOUNT_DISABLED = 'ErrorAccountDisabled',
    ERROR_ADD_DELEGATES_FAILED = 'ErrorAddDelegatesFailed',
    ERROR_ADDRESS_SPACE_NOT_FOUND = 'ErrorAddressSpaceNotFound',
    ERROR_AD_OPERATION = 'ErrorADOperation',
    ERROR_AD_SESSION_FILTER = 'ErrorADSessionFilter',
    ERROR_AD_UNAVAILABLE = 'ErrorADUnavailable',
    ERROR_SERVICE_UNAVAILABLE = 'ErrorServiceUnavailable',
    ERROR_AUTO_DISCOVER_FAILED = 'ErrorAutoDiscoverFailed',
    ERROR_AFFECTED_TASK_OCCURRENCES_REQUIRED = 'ErrorAffectedTaskOccurrencesRequired',
    ERROR_ATTACHMENT_NEST_LEVEL_LIMIT_EXCEEDED = 'ErrorAttachmentNestLevelLimitExceeded',
    ERROR_ATTACHMENT_SIZE_LIMIT_EXCEEDED = 'ErrorAttachmentSizeLimitExceeded',
    ERROR_ARCHIVE_FOLDER_PATH_CREATION = 'ErrorArchiveFolderPathCreation',
    ERROR_ARCHIVE_MAILBOX_NOT_ENABLED = 'ErrorArchiveMailboxNotEnabled',
    ERROR_ARCHIVE_MAILBOX_SERVICE_DISCOVERY_FAILED = 'ErrorArchiveMailboxServiceDiscoveryFailed',
    ERROR_AVAILABILITY_CONFIG_NOT_FOUND = 'ErrorAvailabilityConfigNotFound',
    ERROR_BATCH_PROCESSING_STOPPED = 'ErrorBatchProcessingStopped',
    ERROR_CALENDAR_CANNOT_MOVE_OR_COPY_OCCURRENCE = 'ErrorCalendarCannotMoveOrCopyOccurrence',
    ERROR_CALENDAR_CANNOT_UPDATE_DELETED_ITEM = 'ErrorCalendarCannotUpdateDeletedItem',
    ERROR_CALENDAR_CANNOT_USE_ID_FOR_OCCURRENCE_ID = 'ErrorCalendarCannotUseIdForOccurrenceId',
    ERROR_CALENDAR_CANNOT_USE_ID_FOR_RECURRING_MASTER_ID = 'ErrorCalendarCannotUseIdForRecurringMasterId',
    ERROR_CALENDAR_DURATION_IS_TOO_LONG = 'ErrorCalendarDurationIsTooLong',
    ERROR_CALENDAR_END_DATE_IS_EARLIER_THAN_START_DATE = 'ErrorCalendarEndDateIsEarlierThanStartDate',
    ERROR_CALENDAR_FOLDER_IS_INVALID_FOR_CALENDAR_VIEW = 'ErrorCalendarFolderIsInvalidForCalendarView',
    ERROR_CALENDAR_INVALID_ATTRIBUTE_VALUE = 'ErrorCalendarInvalidAttributeValue',
    ERROR_CALENDAR_INVALID_DAY_FOR_TIME_CHANGE_PATTERN = 'ErrorCalendarInvalidDayForTimeChangePattern',
    ERROR_CALENDAR_INVALID_DAY_FOR_WEEKLY_RECURRENCE = 'ErrorCalendarInvalidDayForWeeklyRecurrence',
    ERROR_CALENDAR_INVALID_PROPERTY_STATE = 'ErrorCalendarInvalidPropertyState',
    ERROR_CALENDAR_INVALID_PROPERTY_VALUE = 'ErrorCalendarInvalidPropertyValue',
    ERROR_CALENDAR_INVALID_RECURRENCE = 'ErrorCalendarInvalidRecurrence',
    ERROR_CALENDAR_INVALID_TIME_ZONE = 'ErrorCalendarInvalidTimeZone',
    ERROR_CALENDAR_IS_CANCELLED_FOR_ACCEPT = 'ErrorCalendarIsCancelledForAccept',
    ERROR_CALENDAR_IS_CANCELLED_FOR_DECLINE = 'ErrorCalendarIsCancelledForDecline',
    ERROR_CALENDAR_IS_CANCELLED_FOR_REMOVE = 'ErrorCalendarIsCancelledForRemove',
    ERROR_CALENDAR_IS_CANCELLED_FOR_TENTATIVE = 'ErrorCalendarIsCancelledForTentative',
    ERROR_CALENDAR_IS_DELEGATED_FOR_ACCEPT = 'ErrorCalendarIsDelegatedForAccept',
    ERROR_CALENDAR_IS_DELEGATED_FOR_DECLINE = 'ErrorCalendarIsDelegatedForDecline',
    ERROR_CALENDAR_IS_DELEGATED_FOR_REMOVE = 'ErrorCalendarIsDelegatedForRemove',
    ERROR_CALENDAR_IS_DELEGATED_FOR_TENTATIVE = 'ErrorCalendarIsDelegatedForTentative',
    ERROR_CALENDAR_IS_NOT_ORGANIZER = 'ErrorCalendarIsNotOrganizer',
    ERROR_CALENDAR_IS_ORGANIZER_FOR_ACCEPT = 'ErrorCalendarIsOrganizerForAccept',
    ERROR_CALENDAR_IS_ORGANIZER_FOR_DECLINE = 'ErrorCalendarIsOrganizerForDecline',
    ERROR_CALENDAR_IS_ORGANIZER_FOR_REMOVE = 'ErrorCalendarIsOrganizerForRemove',
    ERROR_CALENDAR_IS_ORGANIZER_FOR_TENTATIVE = 'ErrorCalendarIsOrganizerForTentative',
    ERROR_CALENDAR_OCCURRENCE_INDEX_IS_OUT_OF_RECURRENCE_RANGE = 'ErrorCalendarOccurrenceIndexIsOutOfRecurrenceRange',
    ERROR_CALENDAR_OCCURRENCE_IS_DELETED_FROM_RECURRENCE = 'ErrorCalendarOccurrenceIsDeletedFromRecurrence',
    ERROR_CALENDAR_OUT_OF_RANGE = 'ErrorCalendarOutOfRange',
    ERROR_CALENDAR_MEETING_REQUEST_IS_OUT_OF_DATE = 'ErrorCalendarMeetingRequestIsOutOfDate',
    ERROR_CALENDAR_VIEW_RANGE_TOO_BIG = 'ErrorCalendarViewRangeTooBig',
    ERROR_CALLER_IS_INVALID_AD_ACCOUNT = 'ErrorCallerIsInvalidADAccount',
    ERROR_CANNOT_ACCESS_DELETED_PUBLIC_FOLDER = 'ErrorCannotAccessDeletedPublicFolder',
    ERROR_CANNOT_ARCHIVE_CALENDAR_CONTACT_TASK_FOLDER_EXCEPTION = 'ErrorCannotArchiveCalendarContactTaskFolderException',
    ERROR_CANNOT_ARCHIVE_ITEMS_IN_PUBLIC_FOLDERS = 'ErrorCannotArchiveItemsInPublicFolders',
    ERROR_CANNOT_ARCHIVE_ITEMS_IN_ARCHIVE_MAILBOX = 'ErrorCannotArchiveItemsInArchiveMailbox',
    ERROR_CANNOT_CREATE_CALENDAR_ITEM_IN_NON_CALENDAR_FOLDER = 'ErrorCannotCreateCalendarItemInNonCalendarFolder',
    ERROR_CANNOT_CREATE_CONTACT_IN_NON_CONTACT_FOLDER = 'ErrorCannotCreateContactInNonContactFolder',
    ERROR_CANNOT_CREATE_POST_ITEM_IN_NON_MAIL_FOLDER = 'ErrorCannotCreatePostItemInNonMailFolder',
    ERROR_CANNOT_CREATE_TASK_IN_NON_TASK_FOLDER = 'ErrorCannotCreateTaskInNonTaskFolder',
    ERROR_CANNOT_DELETE_OBJECT = 'ErrorCannotDeleteObject',
    ERROR_CANNOT_DISABLE_MANDATORY_EXTENSION = 'ErrorCannotDisableMandatoryExtension',
    ERROR_CANNOT_FIND_USER = 'ErrorCannotFindUser',
    ERROR_CANNOT_GET_SOURCE_FOLDER_PATH = 'ErrorCannotGetSourceFolderPath',
    ERROR_CANNOT_GET_EXTERNAL_ECP_URL = 'ErrorCannotGetExternalEcpUrl',
    ERROR_CANNOT_OPEN_FILE_ATTACHMENT = 'ErrorCannotOpenFileAttachment',
    ERROR_CANNOT_DELETE_TASK_OCCURRENCE = 'ErrorCannotDeleteTaskOccurrence',
    ERROR_CANNOT_EMPTY_FOLDER = 'ErrorCannotEmptyFolder',
    ERROR_CANNOT_SET_CALENDAR_PERMISSION_ON_NON_CALENDAR_FOLDER = 'ErrorCannotSetCalendarPermissionOnNonCalendarFolder',
    ERROR_CANNOT_SET_NON_CALENDAR_PERMISSION_ON_CALENDAR_FOLDER = 'ErrorCannotSetNonCalendarPermissionOnCalendarFolder',
    ERROR_CANNOT_SET_PERMISSION_UNKNOWN_ENTRIES = 'ErrorCannotSetPermissionUnknownEntries',
    ERROR_CANNOT_SPECIFY_SEARCH_FOLDER_AS_SOURCE_FOLDER = 'ErrorCannotSpecifySearchFolderAsSourceFolder',
    ERROR_CANNOT_USE_FOLDER_ID_FOR_ITEM_ID = 'ErrorCannotUseFolderIdForItemId',
    ERROR_CANNOT_USE_ITEM_ID_FOR_FOLDER_ID = 'ErrorCannotUseItemIdForFolderId',
    ERROR_CHANGE_KEY_REQUIRED = 'ErrorChangeKeyRequired',
    ERROR_CHANGE_KEY_REQUIRED_FOR_WRITE_OPERATIONS = 'ErrorChangeKeyRequiredForWriteOperations',
    ERROR_CLIENT_DISCONNECTED = 'ErrorClientDisconnected',
    ERROR_CLIENT_INTENT_INVALID_STATE_DEFINITION = 'ErrorClientIntentInvalidStateDefinition',
    ERROR_CLIENT_INTENT_NOT_FOUND = 'ErrorClientIntentNotFound',
    ERROR_CONNECTION_FAILED = 'ErrorConnectionFailed',
    ERROR_CONTAINS_FILTER_WRONG_TYPE = 'ErrorContainsFilterWrongType',
    ERROR_CONTENT_CONVERSION_FAILED = 'ErrorContentConversionFailed',
    ERROR_CONTENT_INDEXING_NOT_ENABLED = 'ErrorContentIndexingNotEnabled',
    ERROR_CORRUPT_DATA = 'ErrorCorruptData',
    ERROR_CREATE_ITEM_ACCESS_DENIED = 'ErrorCreateItemAccessDenied',
    ERROR_CREATE_MANAGED_FOLDER_PARTIAL_COMPLETION = 'ErrorCreateManagedFolderPartialCompletion',
    ERROR_CREATE_SUBFOLDER_ACCESS_DENIED = 'ErrorCreateSubfolderAccessDenied',
    ERROR_CROSS_MAILBOX_MOVE_COPY = 'ErrorCrossMailboxMoveCopy',
    ERROR_CROSS_SITE_REQUEST = 'ErrorCrossSiteRequest',
    ERROR_DATA_SIZE_LIMIT_EXCEEDED = 'ErrorDataSizeLimitExceeded',
    ERROR_DATA_SOURCE_OPERATION = 'ErrorDataSourceOperation',
    ERROR_DELEGATE_ALREADY_EXISTS = 'ErrorDelegateAlreadyExists',
    ERROR_DELEGATE_CANNOT_ADD_OWNER = 'ErrorDelegateCannotAddOwner',
    ERROR_DELEGATE_MISSING_CONFIGURATION = 'ErrorDelegateMissingConfiguration',
    ERROR_DELEGATE_NO_USER = 'ErrorDelegateNoUser',
    ERROR_DELEGATE_VALIDATION_FAILED = 'ErrorDelegateValidationFailed',
    ERROR_DELETE_DISTINGUISHED_FOLDER = 'ErrorDeleteDistinguishedFolder',
    ERROR_DELETE_ITEMS_FAILED = 'ErrorDeleteItemsFailed',
    ERROR_DELETE_UNIFIED_MESSAGING_PROMPT_FAILED = 'ErrorDeleteUnifiedMessagingPromptFailed',
    ERROR_DISTINGUISHED_USER_NOT_SUPPORTED = 'ErrorDistinguishedUserNotSupported',
    ERROR_DISTRIBUTION_LIST_MEMBER_NOT_EXIST = 'ErrorDistributionListMemberNotExist',
    ERROR_DUPLICATE_INPUT_FOLDER_NAMES = 'ErrorDuplicateInputFolderNames',
    ERROR_DUPLICATE_USER_IDS_SPECIFIED = 'ErrorDuplicateUserIdsSpecified',
    ERROR_EMAIL_ADDRESS_MISMATCH = 'ErrorEmailAddressMismatch',
    ERROR_EVENT_NOT_FOUND = 'ErrorEventNotFound',
    ERROR_EXCEEDED_CONNECTION_COUNT = 'ErrorExceededConnectionCount',
    ERROR_EXCEEDED_SUBSCRIPTION_COUNT = 'ErrorExceededSubscriptionCount',
    ERROR_EXCEEDED_FIND_COUNT_LIMIT = 'ErrorExceededFindCountLimit',
    ERROR_EXPIRED_SUBSCRIPTION = 'ErrorExpiredSubscription',
    ERROR_EXTENSION_NOT_FOUND = 'ErrorExtensionNotFound',
    ERROR_EXTENSIONS_NOT_AUTHORIZED = 'ErrorExtensionsNotAuthorized',
    ERROR_FOLDER_CORRUPT = 'ErrorFolderCorrupt',
    ERROR_FOLDER_NOT_FOUND = 'ErrorFolderNotFound',
    ERROR_FOLDER_PROPERT_REQUEST_FAILED = 'ErrorFolderPropertRequestFailed',
    ERROR_FOLDER_SAVE = 'ErrorFolderSave',
    ERROR_FOLDER_SAVE_FAILED = 'ErrorFolderSaveFailed',
    ERROR_FOLDER_SAVE_PROPERTY_ERROR = 'ErrorFolderSavePropertyError',
    ERROR_FOLDER_EXISTS = 'ErrorFolderExists',
    ERROR_FREE_BUSY_GENERATION_FAILED = 'ErrorFreeBusyGenerationFailed',
    ERROR_GET_SERVER_SECURITY_DESCRIPTOR_FAILED = 'ErrorGetServerSecurityDescriptorFailed',
    ERROR_IM_CONTACT_LIMIT_REACHED = 'ErrorImContactLimitReached',
    ERROR_IM_GROUP_DISPLAY_NAME_ALREADY_EXISTS = 'ErrorImGroupDisplayNameAlreadyExists',
    ERROR_IM_GROUP_LIMIT_REACHED = 'ErrorImGroupLimitReached',
    ERROR_IMPERSONATE_USER_DENIED = 'ErrorImpersonateUserDenied',
    ERROR_IMPERSONATION_DENIED = 'ErrorImpersonationDenied',
    ERROR_IMPERSONATION_FAILED = 'ErrorImpersonationFailed',
    ERROR_INCORRECT_SCHEMA_VERSION = 'ErrorIncorrectSchemaVersion',
    ERROR_INCORRECT_UPDATE_PROPERTY_COUNT = 'ErrorIncorrectUpdatePropertyCount',
    ERROR_INDIVIDUAL_MAILBOX_LIMIT_REACHED = 'ErrorIndividualMailboxLimitReached',
    ERROR_INSUFFICIENT_RESOURCES = 'ErrorInsufficientResources',
    ERROR_INTERNAL_SERVER_ERROR = 'ErrorInternalServerError',
    ERROR_INTERNAL_SERVER_TRANSIENT_ERROR = 'ErrorInternalServerTransientError',
    ERROR_INVALID_ACCESS_LEVEL = 'ErrorInvalidAccessLevel',
    ERROR_INVALID_ARGUMENT = 'ErrorInvalidArgument',
    ERROR_INVALID_ATTACHMENT_ID = 'ErrorInvalidAttachmentId',
    ERROR_INVALID_ATTACHMENT_SUBFILTER = 'ErrorInvalidAttachmentSubfilter',
    ERROR_INVALID_ATTACHMENT_SUBFILTER_TEXT_FILTER = 'ErrorInvalidAttachmentSubfilterTextFilter',
    ERROR_INVALID_AUTHORIZATION_CONTEXT = 'ErrorInvalidAuthorizationContext',
    ERROR_INVALID_CHANGE_KEY = 'ErrorInvalidChangeKey',
    ERROR_INVALID_CLIENT_SECURITY_CONTEXT = 'ErrorInvalidClientSecurityContext',
    ERROR_INVALID_COMPLETE_DATE = 'ErrorInvalidCompleteDate',
    ERROR_INVALID_CONTACT_EMAIL_ADDRESS = 'ErrorInvalidContactEmailAddress',
    ERROR_INVALID_CONTACT_EMAIL_INDEX = 'ErrorInvalidContactEmailIndex',
    ERROR_INVALID_CROSS_FOREST_CREDENTIALS = 'ErrorInvalidCrossForestCredentials',
    ERROR_INVALID_DELEGATE_PERMISSION = 'ErrorInvalidDelegatePermission',
    ERROR_INVALID_DELEGATE_USER_ID = 'ErrorInvalidDelegateUserId',
    ERROR_INVALID_EXCLUDES_RESTRICTION = 'ErrorInvalidExcludesRestriction',
    ERROR_INVALID_EXPRESSION_TYPE_FOR_SUB_FILTER = 'ErrorInvalidExpressionTypeForSubFilter',
    ERROR_INVALID_EXTENDED_PROPERTY = 'ErrorInvalidExtendedProperty',
    ERROR_INVALID_EXTENDED_PROPERTY_VALUE = 'ErrorInvalidExtendedPropertyValue',
    ERROR_INVALID_FOLDER_ID = 'ErrorInvalidFolderId',
    ERROR_INVALID_FOLDER_TYPE_FOR_OPERATION = 'ErrorInvalidFolderTypeForOperation',
    ERROR_INVALID_FRACTIONAL_PAGING_PARAMETERS = 'ErrorInvalidFractionalPagingParameters',
    ERROR_INVALID_FREE_BUSY_VIEW_TYPE = 'ErrorInvalidFreeBusyViewType',
    ERROR_INVALID_ID = 'ErrorInvalidId',
    ERROR_INVALID_ID_EMPTY = 'ErrorInvalidIdEmpty',
    ERROR_INVALID_ID_MALFORMED = 'ErrorInvalidIdMalformed',
    ERROR_INVALID_ID_MALFORMED_EWS_LEGACY_ID_FORMAT = 'ErrorInvalidIdMalformedEwsLegacyIdFormat',
    ERROR_INVALID_ID_MONIKER_TOO_LONG = 'ErrorInvalidIdMonikerTooLong',
    ERROR_INVALID_ID_NOT_AN_ITEM_ATTACHMENT_ID = 'ErrorInvalidIdNotAnItemAttachmentId',
    ERROR_INVALID_ID_RETURNED_BY_RESOLVE_NAMES = 'ErrorInvalidIdReturnedByResolveNames',
    ERROR_INVALID_ID_STORE_OBJECT_ID_TOO_LONG = 'ErrorInvalidIdStoreObjectIdTooLong',
    ERROR_INVALID_ID_TOO_MANY_ATTACHMENT_LEVELS = 'ErrorInvalidIdTooManyAttachmentLevels',
    ERROR_INVALID_ID_XML = 'ErrorInvalidIdXml',
    ERROR_INVALID_IM_CONTACT_ID = 'ErrorInvalidImContactId',
    ERROR_INVALID_IM_DISTRIBUTION_GROUP_SMTP_ADDRESS = 'ErrorInvalidImDistributionGroupSmtpAddress',
    ERROR_INVALID_IM_GROUP_ID = 'ErrorInvalidImGroupId',
    ERROR_INVALID_INDEXED_PAGING_PARAMETERS = 'ErrorInvalidIndexedPagingParameters',
    ERROR_INVALID_INTERNET_HEADER_CHILD_NODES = 'ErrorInvalidInternetHeaderChildNodes',
    ERROR_INVALID_ITEM_FOR_OPERATION_ARCHIVE_ITEM = 'ErrorInvalidItemForOperationArchiveItem',
    ERROR_INVALID_ITEM_FOR_OPERATION_CREATE_ITEM_ATTACHMENT = 'ErrorInvalidItemForOperationCreateItemAttachment',
    ERROR_INVALID_ITEM_FOR_OPERATION_CREATE_ITEM = 'ErrorInvalidItemForOperationCreateItem',
    ERROR_INVALID_ITEM_FOR_OPERATION_ACCEPT_ITEM = 'ErrorInvalidItemForOperationAcceptItem',
    ERROR_INVALID_ITEM_FOR_OPERATION_DECLINE_ITEM = 'ErrorInvalidItemForOperationDeclineItem',
    ERROR_INVALID_ITEM_FOR_OPERATION_CANCEL_ITEM = 'ErrorInvalidItemForOperationCancelItem',
    ERROR_INVALID_ITEM_FOR_OPERATION_EXPAND_DL = 'ErrorInvalidItemForOperationExpandDL',
    ERROR_INVALID_ITEM_FOR_OPERATION_REMOVE_ITEM = 'ErrorInvalidItemForOperationRemoveItem',
    ERROR_INVALID_ITEM_FOR_OPERATION_SEND_ITEM = 'ErrorInvalidItemForOperationSendItem',
    ERROR_INVALID_ITEM_FOR_OPERATION_TENTATIVE = 'ErrorInvalidItemForOperationTentative',
    ERROR_INVALID_LOGON_TYPE = 'ErrorInvalidLogonType',
    ERROR_INVALID_LIKE_REQUEST = 'ErrorInvalidLikeRequest',
    ERROR_INVALID_MAILBOX = 'ErrorInvalidMailbox',
    ERROR_INVALID_MANAGED_FOLDER_PROPERTY = 'ErrorInvalidManagedFolderProperty',
    ERROR_INVALID_MANAGED_FOLDER_QUOTA = 'ErrorInvalidManagedFolderQuota',
    ERROR_INVALID_MANAGED_FOLDER_SIZE = 'ErrorInvalidManagedFolderSize',
    ERROR_INVALID_MERGED_FREE_BUSY_INTERVAL = 'ErrorInvalidMergedFreeBusyInterval',
    ERROR_INVALID_NAME_FOR_NAME_RESOLUTION = 'ErrorInvalidNameForNameResolution',
    ERROR_INVALID_OPERATION = 'ErrorInvalidOperation',
    ERROR_INVALID_NETWORK_SERVICE_CONTEXT = 'ErrorInvalidNetworkServiceContext',
    ERROR_INVALID_OOF_PARAMETER = 'ErrorInvalidOofParameter',
    ERROR_INVALID_PAGING_MAX_ROWS = 'ErrorInvalidPagingMaxRows',
    ERROR_INVALID_PARENT_FOLDER = 'ErrorInvalidParentFolder',
    ERROR_INVALID_PERCENT_COMPLETE_VALUE = 'ErrorInvalidPercentCompleteValue',
    ERROR_INVALID_PERMISSION_SETTINGS = 'ErrorInvalidPermissionSettings',
    ERROR_INVALID_PHONE_CALL_ID = 'ErrorInvalidPhoneCallId',
    ERROR_INVALID_PHONE_NUMBER = 'ErrorInvalidPhoneNumber',
    ERROR_INVALID_USER_INFO = 'ErrorInvalidUserInfo',
    ERROR_INVALID_PROPERTY_APPEND = 'ErrorInvalidPropertyAppend',
    ERROR_INVALID_PROPERTY_DELETE = 'ErrorInvalidPropertyDelete',
    ERROR_INVALID_PROPERTY_FOR_EXISTS = 'ErrorInvalidPropertyForExists',
    ERROR_INVALID_PROPERTY_FOR_OPERATION = 'ErrorInvalidPropertyForOperation',
    ERROR_INVALID_PROPERTY_REQUEST = 'ErrorInvalidPropertyRequest',
    ERROR_INVALID_PROPERTY_SET = 'ErrorInvalidPropertySet',
    ERROR_INVALID_PROPERTY_UPDATE_SENT_MESSAGE = 'ErrorInvalidPropertyUpdateSentMessage',
    ERROR_INVALID_PROXY_SECURITY_CONTEXT = 'ErrorInvalidProxySecurityContext',
    ERROR_INVALID_PULL_SUBSCRIPTION_ID = 'ErrorInvalidPullSubscriptionId',
    ERROR_INVALID_PUSH_SUBSCRIPTION_URL = 'ErrorInvalidPushSubscriptionUrl',
    ERROR_INVALID_RECIPIENTS = 'ErrorInvalidRecipients',
    ERROR_INVALID_RECIPIENT_SUBFILTER = 'ErrorInvalidRecipientSubfilter',
    ERROR_INVALID_RECIPIENT_SUBFILTER_COMPARISON = 'ErrorInvalidRecipientSubfilterComparison',
    ERROR_INVALID_RECIPIENT_SUBFILTER_ORDER = 'ErrorInvalidRecipientSubfilterOrder',
    ERROR_INVALID_RECIPIENT_SUBFILTER_TEXT_FILTER = 'ErrorInvalidRecipientSubfilterTextFilter',
    ERROR_INVALID_REFERENCE_ITEM = 'ErrorInvalidReferenceItem',
    ERROR_INVALID_REQUEST = 'ErrorInvalidRequest',
    ERROR_INVALID_RESTRICTION = 'ErrorInvalidRestriction',
    ERROR_INVALID_RETENTION_TAG_TYPE_MISMATCH = 'ErrorInvalidRetentionTagTypeMismatch',
    ERROR_INVALID_RETENTION_TAG_INVISIBLE = 'ErrorInvalidRetentionTagInvisible',
    ERROR_INVALID_RETENTION_TAG_INHERITANCE = 'ErrorInvalidRetentionTagInheritance',
    ERROR_INVALID_RETENTION_TAG_ID_GUID = 'ErrorInvalidRetentionTagIdGuid',
    ERROR_INVALID_ROUTING_TYPE = 'ErrorInvalidRoutingType',
    ERROR_INVALID_SCHEDULED_OOF_DURATION = 'ErrorInvalidScheduledOofDuration',
    ERROR_INVALID_SCHEMA_VERSION_FOR_MAILBOX_VERSION = 'ErrorInvalidSchemaVersionForMailboxVersion',
    ERROR_INVALID_SECURITY_DESCRIPTOR = 'ErrorInvalidSecurityDescriptor',
    ERROR_INVALID_SEND_ITEM_SAVE_SETTINGS = 'ErrorInvalidSendItemSaveSettings',
    ERROR_INVALID_SERIALIZED_ACCESS_TOKEN = 'ErrorInvalidSerializedAccessToken',
    ERROR_INVALID_SERVER_VERSION = 'ErrorInvalidServerVersion',
    ERROR_INVALID_SID = 'ErrorInvalidSid',
    ERROR_INVALID_SIP_URI = 'ErrorInvalidSIPUri',
    ERROR_INVALID_SMTP_ADDRESS = 'ErrorInvalidSmtpAddress',
    ERROR_INVALID_SUBFILTER_TYPE = 'ErrorInvalidSubfilterType',
    ERROR_INVALID_SUBFILTER_TYPE_NOT_ATTENDEE_TYPE = 'ErrorInvalidSubfilterTypeNotAttendeeType',
    ERROR_INVALID_SUBFILTER_TYPE_NOT_RECIPIENT_TYPE = 'ErrorInvalidSubfilterTypeNotRecipientType',
    ERROR_INVALID_SUBSCRIPTION = 'ErrorInvalidSubscription',
    ERROR_INVALID_SUBSCRIPTION_REQUEST = 'ErrorInvalidSubscriptionRequest',
    ERROR_INVALID_SYNC_STATE_DATA = 'ErrorInvalidSyncStateData',
    ERROR_INVALID_TIME_INTERVAL = 'ErrorInvalidTimeInterval',
    ERROR_INVALID_USER_OOF_SETTINGS = 'ErrorInvalidUserOofSettings',
    ERROR_INVALID_USER_PRINCIPAL_NAME = 'ErrorInvalidUserPrincipalName',
    ERROR_INVALID_USER_SID = 'ErrorInvalidUserSid',
    ERROR_INVALID_USER_SID_MISSING_UPN = 'ErrorInvalidUserSidMissingUPN',
    ERROR_INVALID_VALUE_FOR_PROPERTY = 'ErrorInvalidValueForProperty',
    ERROR_INVALID_WATERMARK = 'ErrorInvalidWatermark',
    ERROR_IP_GATEWAY_NOT_FOUND = 'ErrorIPGatewayNotFound',
    ERROR_IRRESOLVABLE_CONFLICT = 'ErrorIrresolvableConflict',
    ERROR_ITEM_CORRUPT = 'ErrorItemCorrupt',
    ERROR_ITEM_NOT_FOUND = 'ErrorItemNotFound',
    ERROR_ITEM_PROPERTY_REQUEST_FAILED = 'ErrorItemPropertyRequestFailed',
    ERROR_ITEM_SAVE = 'ErrorItemSave',
    ERROR_ITEM_SAVE_PROPERTY_ERROR = 'ErrorItemSavePropertyError',
    ERROR_LEGACY_MAILBOX_FREE_BUSY_VIEW_TYPE_NOT_MERGED = 'ErrorLegacyMailboxFreeBusyViewTypeNotMerged',
    ERROR_LOCAL_SERVER_OBJECT_NOT_FOUND = 'ErrorLocalServerObjectNotFound',
    ERROR_LOGON_AS_NETWORK_SERVICE_FAILED = 'ErrorLogonAsNetworkServiceFailed',
    ERROR_MAILBOX_CONFIGURATION = 'ErrorMailboxConfiguration',
    ERROR_MAILBOX_DATA_ARRAY_EMPTY = 'ErrorMailboxDataArrayEmpty',
    ERROR_MAILBOX_DATA_ARRAY_TOO_BIG = 'ErrorMailboxDataArrayTooBig',
    ERROR_MAILBOX_HOLD_NOT_FOUND = 'ErrorMailboxHoldNotFound',
    ERROR_MAILBOX_LOGON_FAILED = 'ErrorMailboxLogonFailed',
    ERROR_MAILBOX_MOVE_IN_PROGRESS = 'ErrorMailboxMoveInProgress',
    ERROR_MAILBOX_STORE_UNAVAILABLE = 'ErrorMailboxStoreUnavailable',
    ERROR_MAIL_RECIPIENT_NOT_FOUND = 'ErrorMailRecipientNotFound',
    ERROR_MAIL_TIPS_DISABLED = 'ErrorMailTipsDisabled',
    ERROR_MANAGED_FOLDER_ALREADY_EXISTS = 'ErrorManagedFolderAlreadyExists',
    ERROR_MANAGED_FOLDER_NOT_FOUND = 'ErrorManagedFolderNotFound',
    ERROR_MANAGED_FOLDERS_ROOT_FAILURE = 'ErrorManagedFoldersRootFailure',
    ERROR_MEETING_SUGGESTION_GENERATION_FAILED = 'ErrorMeetingSuggestionGenerationFailed',
    ERROR_MESSAGE_DISPOSITION_REQUIRED = 'ErrorMessageDispositionRequired',
    ERROR_MESSAGE_SIZE_EXCEEDED = 'ErrorMessageSizeExceeded',
    ERROR_MIME_CONTENT_CONVERSION_FAILED = 'ErrorMimeContentConversionFailed',
    ERROR_MIME_CONTENT_INVALID = 'ErrorMimeContentInvalid',
    ERROR_MIME_CONTENT_INVALID_BASE_64_STRING = 'ErrorMimeContentInvalidBase64String',
    ERROR_MISSING_ARGUMENT = 'ErrorMissingArgument',
    ERROR_MISSING_EMAIL_ADDRESS = 'ErrorMissingEmailAddress',
    ERROR_MISSING_EMAIL_ADDRESS_FOR_MANAGED_FOLDER = 'ErrorMissingEmailAddressForManagedFolder',
    ERROR_MISSING_INFORMATION_EMAIL_ADDRESS = 'ErrorMissingInformationEmailAddress',
    ERROR_MISSING_INFORMATION_REFERENCE_ITEM_ID = 'ErrorMissingInformationReferenceItemId',
    ERROR_MISSING_ITEM_FOR_CREATE_ITEM_ATTACHMENT = 'ErrorMissingItemForCreateItemAttachment',
    ERROR_MISSING_MANAGED_FOLDER_ID = 'ErrorMissingManagedFolderId',
    ERROR_MISSING_RECIPIENTS = 'ErrorMissingRecipients',
    ERROR_MISSING_USER_ID_INFORMATION = 'ErrorMissingUserIdInformation',
    ERROR_MORE_THAN_ONE_ACCESS_MODE_SPECIFIED = 'ErrorMoreThanOneAccessModeSpecified',
    ERROR_MOVE_COPY_FAILED = 'ErrorMoveCopyFailed',
    ERROR_MOVE_DISTINGUISHED_FOLDER = 'ErrorMoveDistinguishedFolder',
    ERROR_MULTI_LEGACY_MAILBOX_ACCESS = 'ErrorMultiLegacyMailboxAccess',
    ERROR_NAME_RESOLUTION_MULTIPLE_RESULTS = 'ErrorNameResolutionMultipleResults',
    ERROR_NAME_RESOLUTION_NO_MAILBOX = 'ErrorNameResolutionNoMailbox',
    ERROR_NAME_RESOLUTION_NO_RESULTS = 'ErrorNameResolutionNoResults',
    ERROR_NO_APPLICABLE_PROXY_CAS_SERVERS_AVAILABLE = 'ErrorNoApplicableProxyCASServersAvailable',
    ERROR_NO_CALENDAR = 'ErrorNoCalendar',
    ERROR_NO_DESTINATION_CAS_DUE_TO_KERBEROS_REQUIREMENTS = 'ErrorNoDestinationCASDueToKerberosRequirements',
    ERROR_NO_DESTINATION_CAS_DUE_TO_SSL_REQUIREMENTS = 'ErrorNoDestinationCASDueToSSLRequirements',
    ERROR_NO_DESTINATION_CAS_DUE_TO_VERSION_MISMATCH = 'ErrorNoDestinationCASDueToVersionMismatch',
    ERROR_NO_FOLDER_CLASS_OVERRIDE = 'ErrorNoFolderClassOverride',
    ERROR_NO_FREE_BUSY_ACCESS = 'ErrorNoFreeBusyAccess',
    ERROR_NON_EXISTENT_MAILBOX = 'ErrorNonExistentMailbox',
    ERROR_NON_PRIMARY_SMTP_ADDRESS = 'ErrorNonPrimarySmtpAddress',
    ERROR_NO_PROPERTY_TAG_FOR_CUSTOM_PROPERTIES = 'ErrorNoPropertyTagForCustomProperties',
    ERROR_NO_PUBLIC_FOLDER_REPLICA_AVAILABLE = 'ErrorNoPublicFolderReplicaAvailable',
    ERROR_NO_PUBLIC_FOLDER_SERVER_AVAILABLE = 'ErrorNoPublicFolderServerAvailable',
    ERROR_NO_RESPONDING_CAS_IN_DESTINATION_SITE = 'ErrorNoRespondingCASInDestinationSite',
    ERROR_NOT_DELEGATE = 'ErrorNotDelegate',
    ERROR_NOT_ENOUGH_MEMORY = 'ErrorNotEnoughMemory',
    ERROR_OBJECT_TYPE_CHANGED = 'ErrorObjectTypeChanged',
    ERROR_OCCURRENCE_CROSSING_BOUNDARY = 'ErrorOccurrenceCrossingBoundary',
    ERROR_OCCURRENCE_TIME_SPAN_TOO_BIG = 'ErrorOccurrenceTimeSpanTooBig',
    ERROR_OPERATION_NOT_ALLOWED_WITH_PUBLIC_FOLDER_ROOT = 'ErrorOperationNotAllowedWithPublicFolderRoot',
    ERROR_PARENT_FOLDER_ID_REQUIRED = 'ErrorParentFolderIdRequired',
    ERROR_PARENT_FOLDER_NOT_FOUND = 'ErrorParentFolderNotFound',
    ERROR_PASSWORD_CHANGE_REQUIRED = 'ErrorPasswordChangeRequired',
    ERROR_PASSWORD_EXPIRED = 'ErrorPasswordExpired',
    ERROR_PHONE_NUMBER_NOT_DIALABLE = 'ErrorPhoneNumberNotDialable',
    ERROR_PROPERTY_UPDATE = 'ErrorPropertyUpdate',
    ERROR_PROMPT_PUBLISHING_OPERATION_FAILED = 'ErrorPromptPublishingOperationFailed',
    ERROR_PROPERTY_VALIDATION_FAILURE = 'ErrorPropertyValidationFailure',
    ERROR_PROXIED_SUBSCRIPTION_CALL_FAILURE = 'ErrorProxiedSubscriptionCallFailure',
    ERROR_PROXY_CALL_FAILED = 'ErrorProxyCallFailed',
    ERROR_PROXY_GROUP_SID_LIMIT_EXCEEDED = 'ErrorProxyGroupSidLimitExceeded',
    ERROR_PROXY_REQUEST_NOT_ALLOWED = 'ErrorProxyRequestNotAllowed',
    ERROR_PROXY_REQUEST_PROCESSING_FAILED = 'ErrorProxyRequestProcessingFailed',
    ERROR_PROXY_SERVICE_DISCOVERY_FAILED = 'ErrorProxyServiceDiscoveryFailed',
    ERROR_PROXY_TOKEN_EXPIRED = 'ErrorProxyTokenExpired',
    ERROR_PUBLIC_FOLDER_MAILBOX_DISCOVERY_FAILED = 'ErrorPublicFolderMailboxDiscoveryFailed',
    ERROR_PUBLIC_FOLDER_OPERATION_FAILED = 'ErrorPublicFolderOperationFailed',
    ERROR_PUBLIC_FOLDER_REQUEST_PROCESSING_FAILED = 'ErrorPublicFolderRequestProcessingFailed',
    ERROR_PUBLIC_FOLDER_SERVER_NOT_FOUND = 'ErrorPublicFolderServerNotFound',
    ERROR_PUBLIC_FOLDER_SYNC_EXCEPTION = 'ErrorPublicFolderSyncException',
    ERROR_QUERY_FILTER_TOO_LONG = 'ErrorQueryFilterTooLong',
    ERROR_QUOTA_EXCEEDED = 'ErrorQuotaExceeded',
    ERROR_READ_EVENTS_FAILED = 'ErrorReadEventsFailed',
    ERROR_READ_RECEIPT_NOT_PENDING = 'ErrorReadReceiptNotPending',
    ERROR_RECURRENCE_END_DATE_TOO_BIG = 'ErrorRecurrenceEndDateTooBig',
    ERROR_RECURRENCE_HAS_NO_OCCURRENCE = 'ErrorRecurrenceHasNoOccurrence',
    ERROR_REMOVE_DELEGATES_FAILED = 'ErrorRemoveDelegatesFailed',
    ERROR_REQUEST_ABORTED = 'ErrorRequestAborted',
    ERROR_REQUEST_STREAM_TOO_BIG = 'ErrorRequestStreamTooBig',
    ERROR_REQUIRED_PROPERTY_MISSING = 'ErrorRequiredPropertyMissing',
    ERROR_RESOLVE_NAMES_INVALID_FOLDER_TYPE = 'ErrorResolveNamesInvalidFolderType',
    ERROR_RESOLVE_NAMES_ONLY_ONE_CONTACTS_FOLDER_ALLOWED = 'ErrorResolveNamesOnlyOneContactsFolderAllowed',
    ERROR_RESPONSE_SCHEMA_VALIDATION = 'ErrorResponseSchemaValidation',
    ERROR_RESTRICTION_TOO_LONG = 'ErrorRestrictionTooLong',
    ERROR_RESTRICTION_TOO_COMPLEX = 'ErrorRestrictionTooComplex',
    ERROR_RESULT_SET_TOO_BIG = 'ErrorResultSetTooBig',
    ERROR_INVALID_EXCHANGE_IMPERSONATION_HEADER_DATA = 'ErrorInvalidExchangeImpersonationHeaderData',
    ERROR_SAVED_ITEM_FOLDER_NOT_FOUND = 'ErrorSavedItemFolderNotFound',
    ERROR_SCHEMA_VALIDATION = 'ErrorSchemaValidation',
    ERROR_SEARCH_FOLDER_NOT_INITIALIZED = 'ErrorSearchFolderNotInitialized',
    ERROR_SEND_AS_DENIED = 'ErrorSendAsDenied',
    ERROR_SEND_MEETING_CANCELLATIONS_REQUIRED = 'ErrorSendMeetingCancellationsRequired',
    ERROR_SEND_MEETING_INVITATIONS_OR_CANCELLATIONS_REQUIRED = 'ErrorSendMeetingInvitationsOrCancellationsRequired',
    ERROR_SEND_MEETING_INVITATIONS_REQUIRED = 'ErrorSendMeetingInvitationsRequired',
    ERROR_SENT_MEETING_REQUEST_UPDATE = 'ErrorSentMeetingRequestUpdate',
    ERROR_SENT_TASK_REQUEST_UPDATE = 'ErrorSentTaskRequestUpdate',
    ERROR_SERVER_BUSY = 'ErrorServerBusy',
    ERROR_SERVICE_DISCOVERY_FAILED = 'ErrorServiceDiscoveryFailed',
    ERROR_STALE_OBJECT = 'ErrorStaleObject',
    ERROR_SUBMISSION_QUOTA_EXCEEDED = 'ErrorSubmissionQuotaExceeded',
    ERROR_SUBSCRIPTION_ACCESS_DENIED = 'ErrorSubscriptionAccessDenied',
    ERROR_SUBSCRIPTION_DELEGATE_ACCESS_NOT_SUPPORTED = 'ErrorSubscriptionDelegateAccessNotSupported',
    ERROR_SUBSCRIPTION_NOT_FOUND = 'ErrorSubscriptionNotFound',
    ERROR_SUBSCRIPTION_UNSUBSCRIBED = 'ErrorSubscriptionUnsubscribed',
    ERROR_SYNC_FOLDER_NOT_FOUND = 'ErrorSyncFolderNotFound',
    ERROR_TEAM_MAILBOX_NOT_FOUND = 'ErrorTeamMailboxNotFound',
    ERROR_TEAM_MAILBOX_NOT_LINKED_TO_SHARE_POINT = 'ErrorTeamMailboxNotLinkedToSharePoint',
    ERROR_TEAM_MAILBOX_URL_VALIDATION_FAILED = 'ErrorTeamMailboxUrlValidationFailed',
    ERROR_TEAM_MAILBOX_NOT_AUTHORIZED_OWNER = 'ErrorTeamMailboxNotAuthorizedOwner',
    ERROR_TEAM_MAILBOX_ACTIVE_TO_PENDING_DELETE = 'ErrorTeamMailboxActiveToPendingDelete',
    ERROR_TEAM_MAILBOX_FAILED_SENDING_NOTIFICATIONS = 'ErrorTeamMailboxFailedSendingNotifications',
    ERROR_TEAM_MAILBOX_ERROR_UNKNOWN = 'ErrorTeamMailboxErrorUnknown',
    ERROR_TIME_INTERVAL_TOO_BIG = 'ErrorTimeIntervalTooBig',
    ERROR_TIMEOUT_EXPIRED = 'ErrorTimeoutExpired',
    ERROR_TIME_ZONE = 'ErrorTimeZone',
    ERROR_TO_FOLDER_NOT_FOUND = 'ErrorToFolderNotFound',
    ERROR_TOKEN_SERIALIZATION_DENIED = 'ErrorTokenSerializationDenied',
    ERROR_TOO_MANY_OBJECTS_OPENED = 'ErrorTooManyObjectsOpened',
    ERROR_UPDATE_PROPERTY_MISMATCH = 'ErrorUpdatePropertyMismatch',
    ERROR_ACCESSING_PARTIAL_CREATED_UNIFIED_GROUP = 'ErrorAccessingPartialCreatedUnifiedGroup',
    ERROR_UNIFIED_GROUP_MAILBOX_AAD_CREATION_FAILED = 'ErrorUnifiedGroupMailboxAADCreationFailed',
    ERROR_UNIFIED_GROUP_MAILBOX_AAD_DELETE_FAILED = 'ErrorUnifiedGroupMailboxAADDeleteFailed',
    ERROR_UNIFIED_GROUP_MAILBOX_NAMING_POLICY = 'ErrorUnifiedGroupMailboxNamingPolicy',
    ERROR_UNIFIED_GROUP_MAILBOX_DELETE_FAILED = 'ErrorUnifiedGroupMailboxDeleteFailed',
    ERROR_UNIFIED_GROUP_MAILBOX_NOT_FOUND = 'ErrorUnifiedGroupMailboxNotFound',
    ERROR_UNIFIED_GROUP_MAILBOX_UPDATE_DELAYED = 'ErrorUnifiedGroupMailboxUpdateDelayed',
    ERROR_UNIFIED_GROUP_MAILBOX_UPDATED_PARTIAL_PROPERTIES = 'ErrorUnifiedGroupMailboxUpdatedPartialProperties',
    ERROR_UNIFIED_GROUP_MAILBOX_UPDATE_FAILED = 'ErrorUnifiedGroupMailboxUpdateFailed',
    ERROR_UNIFIED_GROUP_MAILBOX_PROVISION_FAILED = 'ErrorUnifiedGroupMailboxProvisionFailed',
    ERROR_UNIFIED_MESSAGING_DIAL_PLAN_NOT_FOUND = 'ErrorUnifiedMessagingDialPlanNotFound',
    ERROR_UNIFIED_MESSAGING_REPORT_DATA_NOT_FOUND = 'ErrorUnifiedMessagingReportDataNotFound',
    ERROR_UNIFIED_MESSAGING_PROMPT_NOT_FOUND = 'ErrorUnifiedMessagingPromptNotFound',
    ERROR_UNIFIED_MESSAGING_REQUEST_FAILED = 'ErrorUnifiedMessagingRequestFailed',
    ERROR_UNIFIED_MESSAGING_SERVER_NOT_FOUND = 'ErrorUnifiedMessagingServerNotFound',
    ERROR_UNABLE_TO_GET_USER_OOF_SETTINGS = 'ErrorUnableToGetUserOofSettings',
    ERROR_UNABLE_TO_REMOVE_IM_CONTACT_FROM_GROUP = 'ErrorUnableToRemoveImContactFromGroup',
    ERROR_UNSUPPORTED_SUB_FILTER = 'ErrorUnsupportedSubFilter',
    ERROR_UNSUPPORTED_CULTURE = 'ErrorUnsupportedCulture',
    ERROR_UNSUPPORTED_MAPI_PROPERTY_TYPE = 'ErrorUnsupportedMapiPropertyType',
    ERROR_UNSUPPORTED_MIME_CONVERSION = 'ErrorUnsupportedMimeConversion',
    ERROR_UNSUPPORTED_PATH_FOR_QUERY = 'ErrorUnsupportedPathForQuery',
    ERROR_UNSUPPORTED_PATH_FOR_SORT_GROUP = 'ErrorUnsupportedPathForSortGroup',
    ERROR_UNSUPPORTED_PROPERTY_DEFINITION = 'ErrorUnsupportedPropertyDefinition',
    ERROR_UNSUPPORTED_QUERY_FILTER = 'ErrorUnsupportedQueryFilter',
    ERROR_UNSUPPORTED_RECURRENCE = 'ErrorUnsupportedRecurrence',
    ERROR_UNSUPPORTED_TYPE_FOR_CONVERSION = 'ErrorUnsupportedTypeForConversion',
    ERROR_UPDATE_DELEGATES_FAILED = 'ErrorUpdateDelegatesFailed',
    ERROR_USER_NOT_UNIFIED_MESSAGING_ENABLED = 'ErrorUserNotUnifiedMessagingEnabled',
    ERROR_VOICE_MAIL_NOT_IMPLEMENTED = 'ErrorVoiceMailNotImplemented',
    ERROR_VALUE_OUT_OF_RANGE = 'ErrorValueOutOfRange',
    ERROR_VIRUS_DETECTED = 'ErrorVirusDetected',
    ERROR_VIRUS_MESSAGE_DELETED = 'ErrorVirusMessageDeleted',
    ERROR_WEB_REQUEST_IN_INVALID_STATE = 'ErrorWebRequestInInvalidState',
    ERROR_WIN_32_INTEROP_ERROR = 'ErrorWin32InteropError',
    ERROR_WORKING_HOURS_SAVE_FAILED = 'ErrorWorkingHoursSaveFailed',
    ERROR_WORKING_HOURS_XML_MALFORMED = 'ErrorWorkingHoursXmlMalformed',
    ERROR_WRONG_SERVER_VERSION = 'ErrorWrongServerVersion',
    ERROR_WRONG_SERVER_VERSION_DELEGATE = 'ErrorWrongServerVersionDelegate',
    ERROR_MISSING_INFORMATION_SHARING_FOLDER_ID = 'ErrorMissingInformationSharingFolderId',
    ERROR_DUPLICATE_SOAP_HEADER = 'ErrorDuplicateSOAPHeader',
    ERROR_SHARING_SYNCHRONIZATION_FAILED = 'ErrorSharingSynchronizationFailed',
    ERROR_SHARING_NO_EXTERNAL_EWS_AVAILABLE = 'ErrorSharingNoExternalEwsAvailable',
    ERROR_FREE_BUSY_DL_LIMIT_REACHED = 'ErrorFreeBusyDLLimitReached',
    ERROR_INVALID_GET_SHARING_FOLDER_REQUEST = 'ErrorInvalidGetSharingFolderRequest',
    ERROR_NOT_ALLOWED_EXTERNAL_SHARING_BY_POLICY = 'ErrorNotAllowedExternalSharingByPolicy',
    ERROR_USER_NOT_ALLOWED_BY_POLICY = 'ErrorUserNotAllowedByPolicy',
    ERROR_PERMISSION_NOT_ALLOWED_BY_POLICY = 'ErrorPermissionNotAllowedByPolicy',
    ERROR_ORGANIZATION_NOT_FEDERATED = 'ErrorOrganizationNotFederated',
    ERROR_MAILBOX_FAILOVER = 'ErrorMailboxFailover',
    ERROR_INVALID_EXTERNAL_SHARING_INITIATOR = 'ErrorInvalidExternalSharingInitiator',
    ERROR_MESSAGE_TRACKING_PERMANENT_ERROR = 'ErrorMessageTrackingPermanentError',
    ERROR_MESSAGE_TRACKING_TRANSIENT_ERROR = 'ErrorMessageTrackingTransientError',
    ERROR_MESSAGE_TRACKING_NO_SUCH_DOMAIN = 'ErrorMessageTrackingNoSuchDomain',
    ERROR_USER_WITHOUT_FEDERATED_PROXY_ADDRESS = 'ErrorUserWithoutFederatedProxyAddress',
    ERROR_INVALID_ORGANIZATION_RELATIONSHIP_FOR_FREE_BUSY = 'ErrorInvalidOrganizationRelationshipForFreeBusy',
    ERROR_INVALID_FEDERATED_ORGANIZATION_ID = 'ErrorInvalidFederatedOrganizationId',
    ERROR_INVALID_EXTERNAL_SHARING_SUBSCRIBER = 'ErrorInvalidExternalSharingSubscriber',
    ERROR_INVALID_SHARING_DATA = 'ErrorInvalidSharingData',
    ERROR_INVALID_SHARING_MESSAGE = 'ErrorInvalidSharingMessage',
    ERROR_NOT_SUPPORTED_SHARING_MESSAGE = 'ErrorNotSupportedSharingMessage',
    ERROR_APPLY_CONVERSATION_ACTION_FAILED = 'ErrorApplyConversationActionFailed',
    ERROR_INBOX_RULES_VALIDATION_ERROR = 'ErrorInboxRulesValidationError',
    ERROR_OUTLOOK_RULE_BLOB_EXISTS = 'ErrorOutlookRuleBlobExists',
    ERROR_RULES_OVER_QUOTA = 'ErrorRulesOverQuota',
    ERROR_NEW_EVENT_STREAM_CONNECTION_OPENED = 'ErrorNewEventStreamConnectionOpened',
    ERROR_MISSED_NOTIFICATION_EVENTS = 'ErrorMissedNotificationEvents',
    ERROR_DUPLICATE_LEGACY_DISTINGUISHED_NAME = 'ErrorDuplicateLegacyDistinguishedName',
    ERROR_INVALID_CLIENT_ACCESS_TOKEN_REQUEST = 'ErrorInvalidClientAccessTokenRequest',
    ERROR_UNAUTHORIZED_CLIENT_ACCESS_TOKEN_REQUEST = 'ErrorUnauthorizedClientAccessTokenRequest',
    ERROR_NO_SPEECH_DETECTED = 'ErrorNoSpeechDetected',
    ERROR_UM_SERVER_UNAVAILABLE = 'ErrorUMServerUnavailable',
    ERROR_RECIPIENT_NOT_FOUND = 'ErrorRecipientNotFound',
    ERROR_RECOGNIZER_NOT_INSTALLED = 'ErrorRecognizerNotInstalled',
    ERROR_SPEECH_GRAMMAR_ERROR = 'ErrorSpeechGrammarError',
    ERROR_INVALID_MANAGEMENT_ROLE_HEADER = 'ErrorInvalidManagementRoleHeader',
    ERROR_LOCATION_SERVICES_DISABLED = 'ErrorLocationServicesDisabled',
    ERROR_LOCATION_SERVICES_REQUEST_TIMED_OUT = 'ErrorLocationServicesRequestTimedOut',
    ERROR_LOCATION_SERVICES_REQUEST_FAILED = 'ErrorLocationServicesRequestFailed',
    ERROR_LOCATION_SERVICES_INVALID_REQUEST = 'ErrorLocationServicesInvalidRequest',
    ERROR_WEATHER_SERVICE_DISABLED = 'ErrorWeatherServiceDisabled',
    ERROR_MAILBOX_SCOPE_NOT_ALLOWED_WITHOUT_QUERY_STRING = 'ErrorMailboxScopeNotAllowedWithoutQueryString',
    ERROR_ARCHIVE_MAILBOX_SEARCH_FAILED = 'ErrorArchiveMailboxSearchFailed',
    ERROR_GET_REMOTE_ARCHIVE_FOLDER_FAILED = 'ErrorGetRemoteArchiveFolderFailed',
    ERROR_FIND_REMOTE_ARCHIVE_FOLDER_FAILED = 'ErrorFindRemoteArchiveFolderFailed',
    ERROR_GET_REMOTE_ARCHIVE_ITEM_FAILED = 'ErrorGetRemoteArchiveItemFailed',
    ERROR_EXPORT_REMOTE_ARCHIVE_ITEMS_FAILED = 'ErrorExportRemoteArchiveItemsFailed',
    ERROR_INVALID_PHOTO_SIZE = 'ErrorInvalidPhotoSize',
    ERROR_SEARCH_QUERY_HAS_TOO_MANY_KEYWORDS = 'ErrorSearchQueryHasTooManyKeywords',
    ERROR_SEARCH_TOO_MANY_MAILBOXES = 'ErrorSearchTooManyMailboxes',
    ERROR_INVALID_RETENTION_TAG_NONE = 'ErrorInvalidRetentionTagNone',
    ERROR_DISCOVERY_SEARCHES_DISABLED = 'ErrorDiscoverySearchesDisabled',
    ERROR_CALENDAR_SEEK_TO_CONDITION_NOT_SUPPORTED = 'ErrorCalendarSeekToConditionNotSupported',
    ERROR_CALENDAR_IS_GROUP_MAILBOX_FOR_ACCEPT = 'ErrorCalendarIsGroupMailboxForAccept',
    ERROR_CALENDAR_IS_GROUP_MAILBOX_FOR_DECLINE = 'ErrorCalendarIsGroupMailboxForDecline',
    ERROR_CALENDAR_IS_GROUP_MAILBOX_FOR_TENTATIVE = 'ErrorCalendarIsGroupMailboxForTentative',
    ERROR_CALENDAR_IS_GROUP_MAILBOX_FOR_SUPPRESS_READ_RECEIPT = 'ErrorCalendarIsGroupMailboxForSuppressReadReceipt',
    ERROR_ORGANIZATION_ACCESS_BLOCKED = 'ErrorOrganizationAccessBlocked',
    ERROR_INVALID_LICENSE = 'ErrorInvalidLicense',
    ERROR_MESSAGE_PER_FOLDER_COUNT_RECEIVE_QUOTA_EXCEEDED = 'ErrorMessagePerFolderCountReceiveQuotaExceeded',
    ERROR_INVALID_BULK_ACTION_TYPE = 'ErrorInvalidBulkActionType',
    ERROR_INVALID_KEEP_N_COUNT = 'ErrorInvalidKeepNCount',
    ERROR_INVALID_KEEP_N_TYPE = 'ErrorInvalidKeepNType',
    ERROR_NO_O_AUTH_SERVER_AVAILABLE_FOR_REQUEST = 'ErrorNoOAuthServerAvailableForRequest',
    ERROR_INSTANT_SEARCH_SESSION_EXPIRED = 'ErrorInstantSearchSessionExpired',
    ERROR_INSTANT_SEARCH_TIMEOUT = 'ErrorInstantSearchTimeout',
    ERROR_INSTANT_SEARCH_FAILED = 'ErrorInstantSearchFailed',
    ERROR_UNSUPPORTED_USER_FOR_EXECUTE_SEARCH = 'ErrorUnsupportedUserForExecuteSearch',
    ERROR_DUPLICATE_EXTENDED_KEYWORD_DEFINITION = 'ErrorDuplicateExtendedKeywordDefinition',
    ERROR_MISSING_EXCHANGE_PRINCIPAL = 'ErrorMissingExchangePrincipal',
    ERROR_UNEXPECTED_UNIFIED_GROUPS_COUNT = 'ErrorUnexpectedUnifiedGroupsCount',
    ERROR_PARSING_XML_RESPONSE = 'ErrorParsingXMLResponse',
    ERROR_INVALID_FEDERATION_ORGANIZATION_IDENTIFIER = 'ErrorInvalidFederationOrganizationIdentifier',
    ERROR_INVALID_SWEEP_RULE = 'ErrorInvalidSweepRule',
    ERROR_INVALID_SWEEP_RULE_OPERATION_TYPE = 'ErrorInvalidSweepRuleOperationType',
    ERROR_TARGET_DOMAIN_NOT_SUPPORTED = 'ErrorTargetDomainNotSupported',
    ERROR_INVALID_INTERNET_WEB_PROXY_ON_LOCAL_SERVER = 'ErrorInvalidInternetWebProxyOnLocalServer',
    ERROR_NO_SENDER_RESTRICTIONS_SETTINGS_FOUND_IN_REQUEST = 'ErrorNoSenderRestrictionsSettingsFoundInRequest',
    ERROR_DUPLICATE_SENDER_RESTRICTIONS_INPUT_FOUND = 'ErrorDuplicateSenderRestrictionsInputFound',
    ERROR_SENDER_RESTRICTIONS_UPDATE_FAILED = 'ErrorSenderRestrictionsUpdateFailed',
    ERROR_MESSAGE_SUBMISSION_BLOCKED = 'ErrorMessageSubmissionBlocked',
    ERROR_EXCEEDED_MESSAGE_LIMIT = 'ErrorExceededMessageLimit',
    ERROR_EXCEEDED_MAX_RECIPIENT_LIMIT_BLOCK = 'ErrorExceededMaxRecipientLimitBlock',
    ERROR_ACCOUNT_SUSPEND = 'ErrorAccountSuspend',
    ERROR_EXCEEDED_MAX_RECIPIENT_LIMIT = 'ErrorExceededMaxRecipientLimit',
    ERROR_MESSAGE_BLOCKED = 'ErrorMessageBlocked',
    ERROR_ACCOUNT_SUSPEND_SHOW_TIER_UPGRADE = 'ErrorAccountSuspendShowTierUpgrade',
    ERROR_EXCEEDED_MESSAGE_LIMIT_SHOW_TIER_UPGRADE = 'ErrorExceededMessageLimitShowTierUpgrade',
    ERROR_EXCEEDED_MAX_RECIPIENT_LIMIT_SHOW_TIER_UPGRADE = 'ErrorExceededMaxRecipientLimitShowTierUpgrade',
    ERROR_INVALID_LONGITUDE = 'ErrorInvalidLongitude',
    ERROR_INVALID_LATITUDE = 'ErrorInvalidLatitude',
    ERROR_PROXY_SOAP_EXCEPTION = 'ErrorProxySoapException',
    ERROR_UNIFIED_GROUP_ALREADY_EXISTS = 'ErrorUnifiedGroupAlreadyExists',
    ERROR_UNIFIED_GROUP_AAD_AUTHORIZATION_REQUEST_DENIED = 'ErrorUnifiedGroupAadAuthorizationRequestDenied',
    ERROR_UNIFIED_GROUP_CREATION_DISABLED = 'ErrorUnifiedGroupCreationDisabled',
    ERROR_MARKET_PLACE_EXTENSION_ALREADY_INSTALLED_FOR_ORG = 'ErrorMarketPlaceExtensionAlreadyInstalledForOrg',
    ERROR_EXTENSION_ALREADY_INSTALLED_FOR_ORG = 'ErrorExtensionAlreadyInstalledForOrg',
    ERROR_NEWER_EXTENSION_ALREADY_INSTALLED = 'ErrorNewerExtensionAlreadyInstalled',
    ERROR_NEWER_MARKET_PLACE_EXTENSION_ALREADY_INSTALLED = 'ErrorNewerMarketPlaceExtensionAlreadyInstalled',
    ERROR_INVALID_EXTENSION_ID = 'ErrorInvalidExtensionId',
    ERROR_CANNOT_UNINSTALL_PROVIDED_EXTENSIONS = 'ErrorCannotUninstallProvidedExtensions',
    ERROR_NO_RBAC_PERMISSION_TO_INSTALL_MARKET_PLACE_EXTENSIONS = 'ErrorNoRbacPermissionToInstallMarketPlaceExtensions',
    ERROR_NO_RBAC_PERMISSION_TO_INSTALL_READ_WRITE_MAILBOX_EXTENSIONS = 'ErrorNoRbacPermissionToInstallReadWriteMailboxExtensions',
    ERROR_INVALID_REPORT_MESSAGE_ACTION_TYPE = 'ErrorInvalidReportMessageActionType',
    ERROR_CANNOT_DOWNLOAD_EXTENSION_MANIFEST = 'ErrorCannotDownloadExtensionManifest',
    ERROR_CALENDAR_FORWARD_ACTION_NOT_ALLOWED = 'ErrorCalendarForwardActionNotAllowed',
    ERROR_UNIFIED_GROUP_ALIAS_NAMING_POLICY = 'ErrorUnifiedGroupAliasNamingPolicy',
    ERROR_SUBSCRIPTIONS_DISABLED_FOR_GROUP = 'ErrorSubscriptionsDisabledForGroup',
    ERROR_CANNOT_FIND_FILE_ATTACHMENT = 'ErrorCannotFindFileAttachment',
    ERROR_INVALID_VALUE_FOR_FILTER = 'ErrorInvalidValueForFilter',
    ERROR_QUOTA_EXCEEDED_ON_DELETE = 'ErrorQuotaExceededOnDelete',
    ERROR_ACCESS_DENIED_DUE_TO_COMPLIANCE = 'ErrorAccessDeniedDueToCompliance',
    ERROR_RECOVERABLE_ITEMS_ACCESS_DENIED = 'ErrorRecoverableItemsAccessDenied',
}

/**
 * List of values for FolderClass.
 * Folder classes are defined here: https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/folders-and-items-in-ews-in-exchange
 */
export const enum FolderClassType {
    CALENDAR = 'IPF.Appointment',
    CONTACT = 'IPF.Contact',
    TASK = 'IPF.Task',
    NOTE = 'IPF.StickyNote',
    MESSAGE = 'IPF.Note'

}

/**
 * List of values for ItemClass.
 * Item classes are defined here: https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb176446(v=office.12)
 */
export const enum ItemClassType {
    ITEM_CLASS_ACTIVITY = "IPM.Activity", // Journal entries
    ITEM_CLASS_CALENDAR = "IPM.Appointment", // Appointments
    ITEM_CLASS_CONTACT = "IPM.Contact", // Contacts
    ITEM_CLASS_DISTLIST = "IPM.DistList", // Distribution lists
    ITEM_CLASS_DOCUMENT = "IPM.Document", // Documents
    ITEM_CLASS_RECURRENCE_EXCEPTION = "IPM.OLE.Class", // The exception item of a recurrence series
    ITEM_CLASS_UNKNOWN = "IPM",	// Items for which the specified form cannot be found
    ITEM_CLASS_MESSAGE = "IPM.Note", // E-mail messages
    ITEM_CLASS_IMC_REPORT = "IPM.Note.IMC.Notification", // Reports from the Internet Mail Connect (the Exchange Server gateway to the Internet)
    ITEM_CLASS_OOO_TEMPLATE = "IPM.Note.Rules.Oof.Template.Microsoft", // Out-of-office templates
    ITEM_CLASS_POST_NOTE = "IPM.Post",	// Posting notes in a folder
    ITEM_CLASS_NOTE = "IPM.StickyNote", // Creating notes
    ITEM_CLASS_RECALL_REPORT = "IPM.Recall.Report", // Message recall reports
    ITEM_CLASS_RECALL = "IPM.Outlook.Recall", // Recalling sent messages from recipient Inboxes
    ITEM_CLASS_REMOTE_HEADERS = "IPM.Remote", //Remote Mail message headers
    ITEM_CLASS_REPLY_TEMPLATE = "IPM.Note.Rules.ReplyTemplate.Microsoft", // Editing rule reply templates
    ITEM_CLASS_ITEM_STATUS = "IPM.Report", // Reporting item status
    ITEM_CLASS_RESEND = "IPM.Resend", // Resending a failed message
    ITEM_CLASS_MEETING_CANCELED = "IPM.Schedule.Meeting.Canceled", // Meeting cancellations
    ITEM_CLASS_MEETING_REQUEST = "IPM.Schedule.Meeting.Request", // Meeting requests
    ITEM_CLASS_DECLINE_MEETING_RESPONSE = "IPM.Schedule.Meeting.Resp.Neg", // Responses to decline meeting requests
    ITEM_CLASS_ACCEPT_MEETING_RESPONSE = "IPM.Schedule.Meeting.Resp.Pos", // Responses to accept meeting requests
    ITEM_CLASS_TENTATIVE_MEETING_RESPONSE = "IPM.Schedule.Meeting.Resp.Tent", // Responses to tentatively accept meeting requests
    ITEM_CLASS_SECURE_NOTE = "IPM.Note.Secure", // Encrypted notes to other people
    ITEM_CLASS_SIGNED_NOTE = "IPM.Note.Secure.Sign", // Digitally signed notes to other people
    ITEM_CLASS_TASK = "IPM.Task", // Tasks
    ITEM_CLASS_ACCEPT_TASK_RESPONSE = "IPM.TaskRequest.Accept", // Responses to accept task requests
    ITEM_CLASS_DECLINE_TASK_RESPONSE = "IPM.TaskRequest.Decline", // Responses to decline task requests
    ITEM_CLASS_TASK_REQUEST = "IPM.TaskRequest", // Task requests
    ITEM_CLASS_TASK_REQUEST_UPDATE = "IPM.TaskRequest.Update" // Updates to requested tasks
}
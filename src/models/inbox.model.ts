/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import { ResponseMessageType, ArrayOfStringsType, BaseRequestType, GenericArray } from './common.model';
import {
    ImportanceChoicesType, FlaggedForActionType, SensitivityChoicesType, RuleFieldURIType,
    RuleValidationErrorCodeType
} from './enum.model';
import { ArrayOfEmailAddressesType, ItemIdType, TargetFolderIdType } from './mail.model';

/**
 * Defines a response to a GetInboxRules operation request.
 */
export class GetInboxRulesResponseType extends ResponseMessageType {
    OutlookRuleBlobExists?: boolean;
    InboxRules?: ArrayOfRulesType;
}

/**
 * Defines a response to an UpdateInboxRules request.
 */
export class UpdateInboxRulesResponseType extends ResponseMessageType {
    RuleOperationErrors?: ArrayOfRuleOperationErrorsType;
}

/**
 * Represents an array of rules in the user's mailbox.
 */
export class ArrayOfRulesType {
    Rule?: RuleType[];
}

/**
 * Contains a single rule and represents a rule in the user's mailbox.
 */
export class RuleType {
    RuleId?: string;
    DisplayName: string;
    Priority: number;
    IsEnabled: boolean;
    IsNotSupported?: boolean;
    IsInError?: boolean;
    Conditions?: RulePredicatesType;
    Exceptions?: RulePredicatesType;
    Actions?: RuleActionsType;
}

/**
 * Identifies the conditions that, when fulfilled, will trigger the rule actions for a rule.
 * Or identifies the exceptions that represent all the available rule exception conditions for
 * the inbox rule.
 */
export class RulePredicatesType {
    Categories?: ArrayOfStringsType;
    ContainsBodyStrings?: ArrayOfStringsType;
    ContainsHeaderStrings?: ArrayOfStringsType;
    ContainsRecipientStrings?: ArrayOfStringsType;
    ContainsSenderStrings?: ArrayOfStringsType;
    ContainsSubjectOrBodyStrings?: ArrayOfStringsType;
    ContainsSubjectStrings?: ArrayOfStringsType;
    FlaggedForAction?: FlaggedForActionType;
    FromAddresses?: ArrayOfEmailAddressesType;
    FromConnectedAccounts?: ArrayOfStringsType;
    HasAttachments?: boolean;
    Importance?: ImportanceChoicesType;
    IsApprovalRequest?: boolean;
    IsAutomaticForward?: boolean;
    IsAutomaticReply?: boolean;
    IsEncrypted?: boolean;
    IsMeetingRequest?: boolean;
    IsMeetingResponse?: boolean;
    IsNDR?: boolean;
    IsPermissionControlled?: boolean;
    IsReadReceipt?: boolean;
    IsSigned?: boolean;
    IsVoicemail?: boolean;
    ItemClasses?: ArrayOfStringsType;
    MessageClassifications?: ArrayOfStringsType;
    NotSentToMe?: boolean;
    SentCcMe?: boolean;
    SentOnlyToMe?: boolean;
    SentToAddresses?: ArrayOfEmailAddressesType;
    SentToMe?: boolean;
    SentToOrCcMe?: boolean;
    Sensitivity?: SensitivityChoicesType;
    WithinDateRange?: RulePredicateDateRangeType;
    WithinSizeRange?: RulePredicateSizeRangeType;
}

/**
 * Represents the actions to be taken on a message when the conditions are fulfilled.
 */
export class RuleActionsType {
    AssignCategories?: ArrayOfStringsType;
    CopyToFolder?: TargetFolderIdType;
    Delete?: boolean;
    ForwardAsAttachmentToRecipients?: ArrayOfEmailAddressesType;
    ForwardToRecipients?: ArrayOfEmailAddressesType;
    MarkImportance?: ImportanceChoicesType;
    MarkAsRead?: boolean;
    MoveToFolder?: TargetFolderIdType;
    PermanentDelete?: boolean;
    RedirectToRecipients?: ArrayOfEmailAddressesType;
    SendSMSAlertToRecipients?: ArrayOfEmailAddressesType;
    ServerReplyWithMessage?: ItemIdType;
    StopProcessingRules?: boolean;
}

/**
 * Specifies the date range.
 */
export class RulePredicateDateRangeType {
    StartDateTime?: Date;
    EndDateTime?: Date;
}

/**
 * Specifies the size range.
 */
export class RulePredicateSizeRangeType {
    MinimumSize?: number;
    MaximumSize?: number;
}

/**
 * Represents an array of rule validation errors on each rule field that has an error.
 */
export class ArrayOfRuleOperationErrorsType {
    RuleOperationError: RuleOperationErrorType[];
}

/**
 * Represents a rule operation error.
 */
export class RuleOperationErrorType {
    OperationIndex: number;
    ValidationErrors: ArrayOfRuleValidationErrorsType;
}

/**
 * Represents an array of rule validation errors on each rule field that has an error.
 */
export class ArrayOfRuleValidationErrorsType {
    Error: RuleValidationErrorType[];
}

/**
 * Represents a single validation error on a particular rule property value,
 * predicate property value, or action property value.
 */
export class RuleValidationErrorType {
    FieldURI: RuleFieldURIType;
    ErrorCode: RuleValidationErrorCodeType;
    ErrorMessage: string;
    FieldValue: string;
}


//// Request models

/**
 * Defines request for GetInboxRules operation.
 */
export class GetInboxRulesRequestType extends BaseRequestType {
    MailboxSmtpAddress: string;
}

/**
 * Defines request for UpdateInboxRules operation.
 */
export class UpdateInboxRulesRequestType extends BaseRequestType {
    MailboxSmtpAddress?: string;
    RemoveOutlookRuleBlob?: boolean;
    Operations: ArrayOfRuleOperationsType;
}

/**
 * Represents array of rule operations.
 */
export class ArrayOfRuleOperationsType extends GenericArray<RuleOperationType> { }

/**
 * The base rule operation type.
 */
export abstract class RuleOperationType { }

/**
 * Represents an operation to create rule.
 */
export class CreateRuleOperationType extends RuleOperationType {
    Rule: RuleType;
}

/**
 * Represents an operation to set rule.
 */
export class SetRuleOperationType extends RuleOperationType {
    Rule: RuleType;
}

/**
 * Represents an operation to delete rule.
 */
export class DeleteRuleOperationType extends RuleOperationType {
    RuleId: string;
}
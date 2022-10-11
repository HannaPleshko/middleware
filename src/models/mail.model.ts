/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import {
    BaseRequestType, ResponseMessageType, ArrayOfStringsType, AttributeSimpleType, NamedElement, GenericArray,
    BaseResponseMessageType, DaysOfWeekType, AttributesKey, XSITypeKey, PathChoiceType, PathToExtendedFieldType,
    BasePathToElementType, NonEmptyArrayOfFieldOrdersType, PathToUnindexedFieldType, PathToIndexedFieldType
} from './common.model';
import {
    DefaultShapeNamesType,  DistinguishedFolderIdNameType, MailboxTypeType,
    SearchFolderTraversalType, DistinguishedUserType, PermissionActionType, PermissionLevelType,
    PermissionReadAccessType, CalendarPermissionLevelType, CalendarPermissionReadAccessType, CreateActionType,
    SyncFolderItemsScopeType, BodyTypeResponseType, BodyTypeType, SensitivityChoicesType, SendPromptType,
    SharingActionImportance, SharingActionType, SharingAction, LegacyFreeBusyType, CalendarItemTypeType,
    ResponseTypeType, ImportanceChoicesType, IconIndexType, InferenceClassificationType, DayOfWeekType,
    DayOfWeekIndexType, MonthNamesType, FileAsMappingType, ContactSourceType, PhysicalAddressIndexType,
    EmailAddressKeyType, AbchEmailAddressTypeType, PhysicalAddressKeyType, PhoneNumberKeyType, ImAddressKeyType,
    ContactUrlKeyType, EmailReminderSendOption, EmailReminderChangeType, TransitionTargetKindType, LobbyBypassType,
    OnlineMeetingAccessLevelType, PresentersType, FlagStatusType, EmailPositionType, PredictedActionReasonType,
    MeetingRequestTypeType, TaskDelegateStateType, TaskStatusType, RoleMemberTypeType, MemberStatusType,
    MessageDispositionType, CalendarItemCreateOrDeleteOperationType, DisposalType, ConflictResolutionType,
    CalendarItemUpdateOperationType, AffectedTaskOccurrencesType, ReportMessageActionType, NotificationEventTypeType,
    ConnectionStatusType, FolderQueryTraversalType, IndexBasePointType, IdFormatType, ItemQueryTraversalType, 
    SortDirectionType, StandardGroupByType, AggregateType, MailboxSearchLocationType, ConversationQueryTraversalType,
    ViewFilterType, ConversationNodeSortOrder, ConversationActionTypeType, RetentionType
} from './enum.model';

import { RestrictionType, PersonaPostalAddressType } from './persona.model';

/**
 * Defines a request to get a folder from a mailbox in the Exchange store.
 */
export class GetFolderType extends BaseRequestType {
    FolderShape: FolderResponseShapeType;
    FolderIds: NonEmptyArrayOfBaseFolderIdsType;
}

/**
 * Defines a request to find folder from a mailbox in the Exchange store.
 */
export class FindFolderType extends BaseRequestType {
    FolderShape: FolderResponseShapeType;
    IndexedPageFolderView?: IndexedPageViewType;
    FractionalPageFolderView?: FractionalPageViewType;
    Restriction?: RestrictionType;
    ParentFolderIds: NonEmptyArrayOfBaseFolderIdsType;
    Traversal: FolderQueryTraversalType;
}

/**
 * Defines a request to create a folder path and includes a parent folder Id and a relative folder path.
 */
export class CreateFolderType extends BaseRequestType {
    ParentFolderId: TargetFolderIdType;
    Folders: NonEmptyArrayOfFoldersType;
}

/**
 * Defines a request to create a folder path and includes a parent folder Id and a relative folder path.
 */
export class CreateFolderPathType extends BaseRequestType {
    ParentFolderId: TargetFolderIdType;
    RelativeFolderPath: NonEmptyArrayOfFoldersType;
}

/**
 * Defines a request to delete folder.
 */
export class DeleteFolderType extends BaseRequestType {
    FolderIds: NonEmptyArrayOfBaseFolderIdsType;
    DeleteType: DisposalType;
}

/**
 * Defines a request to update properties for a specified folder.
 */
export class UpdateFolderType extends BaseRequestType {
    FolderChanges: NonEmptyArrayOfFolderChangesType;
}

/**
 * Base move/copy folder type.
 */
export abstract class BaseMoveCopyFolderType extends BaseRequestType {
    ToFolderId: TargetFolderIdType;
    FolderIds: NonEmptyArrayOfBaseFolderIdsType;
}

/**
 * Defines a request to move folder.
 */
export class MoveFolderType extends BaseMoveCopyFolderType { }

/**
 * Defines a request to copy folder.
 */
export class CopyFolderType extends BaseMoveCopyFolderType { }

/**
 * Defines a request to convert id.
 */
export class ConvertIdType extends BaseRequestType {
    SourceIds: NonEmptyArrayOfAlternateIdsType;
    DestinationFormat: IdFormatType;
}

/**
 * Defines a request to synchronize a folder hierarchy on a client.
 */
export class SyncFolderHierarchyType extends BaseRequestType {
    FolderShape: FolderResponseShapeType;
    SyncFolderId?: TargetFolderIdType;
    SyncState?: string;
}

/**
 * Defines a request to synchronize items in an Exchange store folder.
 */
export class SyncFolderItemsType extends BaseRequestType {
    ItemShape: ItemResponseShapeType;
    SyncFolderId: TargetFolderIdType;
    SyncState?: string;
    Ignore?: ArrayOfBaseItemIdsType;
    MaxChangesReturned: number;
    SyncScope?: SyncFolderItemsScopeType;
}

/**
 * Defines a request to update items in an Exchange store folder.
 */
export class UploadItemsType extends BaseRequestType {
    Items: NonEmptyArrayOfUploadItemsType;
}

/**
 * Defines a request to create item in an Exchange store folder.
 */
export class CreateItemType extends BaseRequestType {
    SavedItemFolderId?: TargetFolderIdType;
    Items: NonEmptyArrayOfAllItemsType;
    MessageDisposition?: MessageDispositionType;
    SendMeetingInvitations?: CalendarItemCreateOrDeleteOperationType;
}

/**
 * Defines a request to get item in an Exchange store folder.
 */
export class GetItemType extends BaseRequestType {
    ItemShape: ItemResponseShapeType;
    ItemIds: NonEmptyArrayOfBaseItemIdsType;
}

/**
 * Defines a request to send item in an Exchange store folder.
 */
export class SendItemType extends BaseRequestType {
    ItemIds: NonEmptyArrayOfBaseItemIdsType;
    SavedItemFolderId?: TargetFolderIdType;
    SaveItemToFolder: boolean;
}

/**
 * Defines a request to update item in an Exchange store folder.
 */
export class UpdateItemType extends BaseRequestType {
    SavedItemFolderId?: TargetFolderIdType;
    ItemChanges: NonEmptyArrayOfItemChangesType;
    ConflictResolution: ConflictResolutionType;
    MessageDisposition?: MessageDispositionType;
    SendMeetingInvitationsOrCancellations?: CalendarItemUpdateOperationType;
    SuppressReadReceipts?: boolean;
}

/**
 * Defines a request to delete item in an Exchange store folder.
 */
export class DeleteItemType extends BaseRequestType {
    ItemIds: NonEmptyArrayOfBaseItemIdsType;
    DeleteType: DisposalType;
    SendMeetingCancellations?: CalendarItemCreateOrDeleteOperationType;
    AffectedTaskOccurrences?: AffectedTaskOccurrencesType;
    SuppressReadReceipts?: boolean;
}

/**
 * Base move/copy item type.
 */
export abstract class BaseMoveCopyItemType extends BaseRequestType {
    ToFolderId: TargetFolderIdType;
    ItemIds: NonEmptyArrayOfBaseItemIdsType;
    ReturnNewItemIds?: boolean;
}

/**
 * Defines a request to move item in an Exchange store folder.
 */
export class MoveItemType extends BaseMoveCopyItemType { }

/**
 * Defines a request to copy item in an Exchange store folder.
 */
export class CopyItemType extends BaseMoveCopyItemType { }

/**
 * Defines a request to archive item in an Exchange store folder.
 */
export class ArchiveItemType extends BaseRequestType {
    ArchiveSourceFolderId: TargetFolderIdType;
    ItemIds: NonEmptyArrayOfBaseItemIdsType;
}

/**
 * Defines a request to mark item as junk in an Exchange store folder.
 */
export class MarkAsJunkType extends BaseRequestType {
    ItemIds: NonEmptyArrayOfBaseItemIdsType;
    IsJunk: boolean;
    MoveItem: boolean;
}

/**
 * Defines a request to mark all items as read in an Exchange store folder.
 */
export class MarkAllItemsAsReadType extends BaseRequestType {
    FolderIds: NonEmptyArrayOfBaseFolderIdsType;
    ReadFlag: boolean;
    SuppressReadReceipts: boolean;
}

/**
 * Defines a request to report message in an Exchange store folder.
 */
export class ReportMessageType extends BaseRequestType {
    ItemIds: NonEmptyArrayOfBaseItemIdsType;
    ReportAction: ReportMessageActionType;
}

/**
 * Defines a request to create an attachment to an item in the Exchange store.
 */
export class CreateAttachmentType extends BaseRequestType {
    ParentItemId: ItemIdType;
    Attachments: NonEmptyArrayOfAttachmentsType;
}

/**
 * Defines a request to get an attachment in the Exchange store.
 */
export class GetAttachmentType extends BaseRequestType {
    AttachmentShape?: AttachmentResponseShapeType;
    AttachmentIds: NonEmptyArrayOfRequestAttachmentIdsType;
}

/**
 * Defines a request to delete an attachment in the Exchange store.
 */
export class DeleteAttachmentType extends BaseRequestType {
    AttachmentIds: NonEmptyArrayOfRequestAttachmentIdsType;
}

/**
 * Defines a request to subscribe events in the Exchange store.
 */
export class SubscribeType extends BaseRequestType {
    PullSubscriptionRequest?: PullSubscriptionRequestType;
    PushSubscriptionRequest?: PushSubscriptionRequestType;
    StreamingSubscriptionRequest?: StreamingSubscriptionRequestType;
}

/**
 * Defines a request to get app market place url.
 */
export class GetAppMarketplaceUrlType extends BaseRequestType { }

/**
 * Defines a request to get streaming events.
 */
export class GetStreamingEventsType extends BaseRequestType {
    SubscriptionIds: NonEmptyArrayOfSubscriptionIdsType;
    ConnectionTimeout: number;
}

/**
 * Contains the status and result of a single GetStreamingEvents operation request.
 */
export class GetStreamingEventsResponseMessageType extends ResponseMessageType {
    Notifications?: NonEmptyArrayOfNotificationsType;
    ErrorSubscriptionIds?: NonEmptyArrayOfSubscriptionIdsType;
    ConnectionStatus?: ConnectionStatusType;
}

/**
 * Defines a response to a GetStreamingEvents request.
 */
export class GetStreamingEventsResponseType extends BaseResponseMessageType<GetStreamingEventsResponseMessageType> { }

/**
 * Contains the status and result of a single GetAppMarketplaceUrl operation request.
 */
export class GetAppMarketplaceUrlResponseMessageType extends ResponseMessageType {
    AppMarketplaceUrl?: string;
    ConnectorsManagementUrl?: string;
}

/**
 * Contains the status and result of a single Subscribe operation request.
 */
export class SubscribeResponseMessageType extends ResponseMessageType {
    SubscriptionId?: string;
    Watermark?: string;
}

/**
 * Defines a response to a Subscribe request.
 */
export class SubscribeResponseType extends BaseResponseMessageType<SubscribeResponseMessageType> { }

/**
 * Contains the status and result of a single GetAttachment/CreateAttachment operation request.
 */
export class AttachmentInfoResponseMessageType extends ResponseMessageType {
    Attachments: ArrayOfAttachmentsType = new ArrayOfAttachmentsType();
}

/**
 * Contains the status and result of a single GetAttachment operation request.
 */
export class GetAttachmentResponseMessage extends AttachmentInfoResponseMessageType { }

/**
 * Defines a response to a CreateAttachment request.
 */
export class GetAttachmentResponseType extends BaseResponseMessageType<GetAttachmentResponseMessage> { }

/**
 * Contains the status and result of a single CreateAttachment operation request.
 */
export class CreateAttachmentResponseMessage extends AttachmentInfoResponseMessageType { }

/**
 * Defines a response to a CreateAttachment request.
 */
export class CreateAttachmentResponseType extends BaseResponseMessageType<CreateAttachmentResponseMessage> { }

/**
 * Contains the status and result of a single DeleteAttachment operation request.
 */
export class DeleteAttachmentResponseMessageType extends ResponseMessageType {
    RootItemId?: RootItemIdType;
}

/**
 * Defines a response to a DeleteAttachment request.
 */
export class DeleteAttachmentResponseType extends BaseResponseMessageType<DeleteAttachmentResponseMessageType> { }

/**
 * Contains the status and result of a single ReportMessage operation request.
 */
export class ReportMessageResponseMessageType extends ResponseMessageType {
    MovedItemId?: ItemIdType;
}

/**
 * Defines a response to a ReportMessage request.
 */
export class ReportMessageResponseType extends BaseResponseMessageType<ReportMessageResponseMessageType> { }

/**
 * Contains the status and result of a single MarkAllItemsAsRead operation request.
 */
export class MarkAllItemsAsReadResponseMessage extends ResponseMessageType { }

/**
 * Defines a response to a MarkAllItemsAsRead request.
 */
export class MarkAllItemsAsReadResponseType extends BaseResponseMessageType<MarkAllItemsAsReadResponseMessage> { }

/**
 * Contains the status and result of a single MarkAsJunk operation request.
 */
export class MarkAsJunkResponseMessageType extends ResponseMessageType {
    MovedItemId?: ItemIdType;
}

/**
 * Defines a response to a MarkAsJunk request.
 */
export class MarkAsJunkResponseType extends BaseResponseMessageType<MarkAsJunkResponseMessageType> { }

/**
 * Item info response message type.
 */
export class ItemInfoResponseMessageType extends ResponseMessageType {
    Items: ArrayOfRealItemsType = new ArrayOfRealItemsType();
}

/**
 * Contains the status and result of a single CreateItem operation request.
 */
export class CreateItemResponseMessage extends ItemInfoResponseMessageType { }

/**
 * Defines a response to a CreateItem request.
 */
export class CreateItemResponseType extends BaseResponseMessageType<CreateItemResponseMessage> { }

/**
 * Contains the status and result of a single GetItem operation request.
 */
export class GetItemResponseMessage extends ItemInfoResponseMessageType { }

/**
 * Defines a response to a GetItem request.
 */
export class GetItemResponseType extends BaseResponseMessageType<GetItemResponseMessage> { }

/**
 * Contains the status and result of a single UpdateItem operation request.
 */
export class UpdateItemResponseMessageType extends ItemInfoResponseMessageType {
    ConflictResults?: ConflictResultsType;
}

/**
 * Defines a response to a UpdateItem request.
 */
export class UpdateItemResponseType extends BaseResponseMessageType<UpdateItemResponseMessageType> { }

/**
 * Contains the status and result of a single MoveItem operation request.
 */
export class MoveItemResponseMessage extends ItemInfoResponseMessageType { }

/**
 * Defines a response to a MoveItem request.
 */
export class MoveItemResponseType extends BaseResponseMessageType<MoveItemResponseMessage> { }

/**
 * Contains the status and result of a single CopyItem operation request.
 */
export class CopyItemResponseMessage extends ItemInfoResponseMessageType { }

/**
 * Defines a response to a CopyItem request.
 */
export class CopyItemResponseType extends BaseResponseMessageType<CopyItemResponseMessage> { }

/**
 * Contains the status and result of a single ArchiveItem operation request.
 */
export class ArchiveItemResponseMessage extends ItemInfoResponseMessageType { }

/**
 * Defines a response to a ArchiveItem request.
 */
export class ArchiveItemResponseType extends BaseResponseMessageType<ArchiveItemResponseMessage> { }

/**
 * Contains the status and result of a single SendItem operation request.
 */
export class SendItemResponseMessage extends ResponseMessageType { }

/**
 * Defines a response to a SendItem request.
 */
export class SendItemResponseType extends BaseResponseMessageType<SendItemResponseMessage> { }

/**
 * Contains the status and result of a single DeleteItem operation request.
 */
export class DeleteItemResponseMessageType extends ResponseMessageType { }

/**
 * Defines a response to a DeleteItem request.
 */
export class DeleteItemResponseType extends BaseResponseMessageType<DeleteItemResponseMessageType> { }

/**
 * Contains the status and result of a single UploadItems operation request.
 */
export class UploadItemsResponseMessageType extends ResponseMessageType {
    ItemId: ItemIdType;
}

/**
 * Defines a response to a UploadItems request.
 */
export class UploadItemsResponseType extends BaseResponseMessageType<UploadItemsResponseMessageType> { }

/**
 * Contains the status and result of a single SyncFolderItems operation request.
 */
export class SyncFolderItemsResponseMessageType extends ResponseMessageType {
    SyncState?: string;
    IncludesLastItemInRange: boolean;
    Changes: SyncFolderItemsChangesType = new SyncFolderItemsChangesType();
}

/**
 * Defines a response to a SyncFolderItems request.
 */
export class SyncFolderItemsResponseType extends BaseResponseMessageType<SyncFolderItemsResponseMessageType> { }

/**
 * Defines a response to a SyncFolderHierarchy request.
 */
export class SyncFolderHierarchyResponseType extends BaseResponseMessageType<SyncFolderHierarchyResponseMessageType> { }

/**
 * Contains the status and result of a single SyncFolderHierarchy operation request.
 */
export class SyncFolderHierarchyResponseMessageType extends ResponseMessageType {
    SyncState?: string;
    IncludesLastFolderInRange: boolean;
    Changes: SyncFolderHierarchyChangesType = new SyncFolderHierarchyChangesType();
}

/**
 * Contains the status and result of a single FindFolder operation request.
 */
export class FindFolderResponseMessageType extends ResponseMessageType {
    RootFolder?: FindFolderParentType;
}

/**
 * Defines a response to a FindFolder request.
 */
export class FindFolderResponseType extends BaseResponseMessageType<FindFolderResponseMessageType>{ }

/**
 * Folder response message type.
 */
export class FolderInfoResponseMessageType extends ResponseMessageType {
    Folders: ArrayOfFoldersType = new ArrayOfFoldersType();
}

/**
 * Contains the status and result of a single GetFolder operation request.
 */
export class GetFolderResponseMessage extends FolderInfoResponseMessageType { }

/**
 * Defines a response to a GetFolder request.
 */
export class GetFolderResponseType extends BaseResponseMessageType<GetFolderResponseMessage> { }

/**
 * Contains the status and result of a single CreateFolder operation request.
 */
export class CreateFolderResponseMessage extends FolderInfoResponseMessageType { }

/**
 * Defines a response to a CreateFolder request.
 */
export class CreateFolderResponseType extends BaseResponseMessageType<CreateFolderResponseMessage> { }

/**
 * Specifies the response message for a CreateFolderPath request.
 */
export class CreateFolderPathResponseMessage extends FolderInfoResponseMessageType { }

/**
 * Defines a response to a GetFolder request.
 */
export class CreateFolderPathResponseType extends BaseResponseMessageType<CreateFolderPathResponseMessage> { }

/**
 * Specifies the response message for a UpdateFolder request.
 */
export class UpdateFolderResponseMessage extends FolderInfoResponseMessageType { }

/**
 * Defines a response to a UpdateFolder request.
 */
export class UpdateFolderResponseType extends BaseResponseMessageType<UpdateFolderResponseMessage> { }

/**
 * Specifies the response message for a MoveFolder request.
 */
export class MoveFolderResponseMessage extends FolderInfoResponseMessageType { }

/**
 * Defines a response to a MoveFolder request.
 */
export class MoveFolderResponseType extends BaseResponseMessageType<MoveFolderResponseMessage> { }

/**
 * Specifies the response message for a CopyFolder request.
 */
export class CopyFolderResponseMessage extends FolderInfoResponseMessageType { }

/**
 * Defines a response to a CopyFolder request.
 */
export class CopyFolderResponseType extends BaseResponseMessageType<CopyFolderResponseMessage> { }

/**
 * Specifies the response message for a ConvertId request.
 */
export class ConvertIdResponseMessageType extends ResponseMessageType {
    AlternateId?: AlternateIdBaseType;
}

/**
 * Defines a response to a ConvertId request.
 */
export class ConvertIdResponseType extends BaseResponseMessageType<ConvertIdResponseMessageType> { }

/**
 * Defines a response to a DeleteFolder request.
 */
export class DeleteFolderResponseType extends BaseResponseMessageType<DeleteFolderResponseMessage> { }

/**
 * Specifies the response message for a DeleteFolder request.
 */
export class DeleteFolderResponseMessage extends ResponseMessageType { }

/**
 * Array of base item ids.
 */
export class NonEmptyArrayOfBaseItemIdsType extends GenericArray<
    ItemIdType | OccurrenceItemIdType | RecurringMasterItemIdType | RecurringMasterItemIdRangesType> {
}

/**
 * Defines a request to create a managed folder 
 */
export class CreateManagedFolderRequestType extends BaseRequestType {
    FolderNames: NonEmptyArrayOfFolderNamesType;
    Mailbox?: EmailAddressType;
}

/**
 * Defines a response to a CreateManagedFolder request.
 */
export class CreateManagedFolderResponseType extends BaseResponseMessageType<CreateManagedFolderResponseMessage>{ }

/**
 * Contains the status and result of a single CreateManagedFolder operation request.
 */
export class CreateManagedFolderResponseMessage extends FolderInfoResponseMessageType { }

/**
 * Defines a request to create a managed folder 
 */
export class EmptyFolderType extends BaseRequestType {
    FolderIds: NonEmptyArrayOfBaseFolderIdsType;
    DeleteType: DisposalType;
    DeleteSubFolders: boolean;
}

/**
 * Defines a response to a Empty request.
 */
export class EmptyFolderResponseType extends BaseResponseMessageType<EmptyFolderResponseMessage>{ }

/**
 * Contains the status and result of a single EmptyFolder operation request.
 */
export class EmptyFolderResponseMessage extends ResponseMessageType { }

/**
 * Array of real items.
 */
export class ArrayOfRealItemsType extends GenericArray<
    ItemType | MessageType | SharingMessageType | CalendarItemType | ContactItemType | DistributionListType |
    MeetingMessageType | MeetingRequestMessageType | MeetingResponseMessageType | MeetingCancellationMessageType |
    TaskType | PostItemType | RoleMemberItemType | NetworkItemType | AbchPersonItemType> {
}

/**
 * Grouped Summary.
 */
export class GroupSummaryType {
    GroupCount: number;
    UnreadCount: number;
    InstanceKey: string;   //xs:base64Binary
    GroupByValue: string;
}


/**
 * Grouped Items.
 */
export class GroupedItemsType {
    GroupIndex: string;
    Items?: ArrayOfRealItemsType;
    GroupSummary?: GroupSummaryType
}

/**
 * Array of grouped items.
 */
export class ArrayOfGroupedItemsType {
    GroupedItems?: GroupedItemsType[] = [];
}

/**
 * Array of upload items.
 */
export class NonEmptyArrayOfUploadItemsType {
    Item: UploadItemType[] = [];
}

/**
 * Upload item.
 */
export class UploadItemType {
    ParentFolderId: FolderIdType;
    ItemId?: ItemIdType;
    Data: string; // xs:base64Binary
    CreateAction: CreateActionType;
    IsAssociated?: boolean;
}

/**
 * Array of item changes.
 */
export class NonEmptyArrayOfItemChangesType {
    ItemChange: ItemChangeType[] = [];
}

/**
 * Conflict results.
 */
export class ConflictResultsType {
    Count: number;
}

/**
 * Item change.
 */
export class ItemChangeType {
    ItemId?: ItemIdType;
    OccurrenceItemId?: OccurrenceItemIdType;
    RecurringMasterItemId?: RecurringMasterItemIdType;
    Updates: NonEmptyArrayOfItemChangeDescriptionsType;
    CalendarActivityData?: CalendarActivityDataType;
}

export class NonEmptyArrayOfItemChangeDescriptionsType extends GenericArray<ItemChangeDescriptionType> { }

/**
 * Contains a sequenced array of change types that represent the type of differences
 * between the folders on the client and the folders on Exchange server.
 */
export class SyncFolderHierarchyChangesType extends GenericArray<SyncFolderHierarchyCreateType | SyncFolderHierarchyUpdateType | SyncFolderHierarchyDeleteType> { }

/**
 * Identifies a single folder to create/update in the local client store.
 */
export abstract class SyncFolderHierarchyCreateOrUpdateType {
    Folder?: FolderType;
    CalendarFolder?: CalendarFolderType;
    ContactsFolder?: ContactsFolderType;
    SearchFolder?: SearchFolderType;
    TasksFolder?: TasksFolderType;
}

/**
 * Identifies a single folder to create in the local client store.
 */
export class SyncFolderHierarchyCreateType extends SyncFolderHierarchyCreateOrUpdateType implements NamedElement {
    readonly $qname = 'Create';
}

/**
 * Identifies a single folder to update in the local client store.
 */
export class SyncFolderHierarchyUpdateType extends SyncFolderHierarchyCreateOrUpdateType implements NamedElement {
    readonly $qname = 'Update';
}

/**
 * Identifies a single folder to delete in the local client store.
 */
export class SyncFolderHierarchyDeleteType implements NamedElement {
    readonly $qname = 'Delete';
    FolderId: FolderIdType;
}

/**
 * Folder response shape type.
 */
export class FolderResponseShapeType {
    BaseShape: DefaultShapeNamesType;
    AdditionalProperties?: NonEmptyArrayOfPathsToElementType;
}

/**
 * Contains a collection of changes for a specified folder.
 */
export class NonEmptyArrayOfFolderChangesType {
    FolderChange: FolderChangeType[] = [];
}

/**
 * Represents a collection of changes to be performed on a single folder
 */
export class FolderChangeType {
    FolderId?: FolderIdType;
    DistinguishedFolderId?: DistinguishedFolderIdType;
    Updates: NonEmptyArrayOfFolderChangeDescriptionsType;
}

/**
 * Contains a set of elements that define append, set, and delete changes to folder properties.
 */
export class NonEmptyArrayOfFolderChangeDescriptionsType extends GenericArray<FolderChangeDescriptionType> { }

/**
 * Represents a change.
 */
export abstract class ChangeDescriptionType extends PathChoiceType { }

/**
 * Represents an item change.
 */
export abstract class ItemChangeDescriptionType extends ChangeDescriptionType { }

/**
 * Delete item field type.
 */
export class DeleteItemFieldType extends ItemChangeDescriptionType { }

/**
 * Set item field type.
 */
export class SetItemFieldType extends ItemChangeDescriptionType {
    Item?: ItemType;
    Message?: MessageType;
    SharingMessage?: SharingMessageType;
    CalendarItem?: CalendarItemType;
    Contact?: ContactItemType;
    DistributionList?: DistributionListType;
    MeetingMessage?: MeetingMessageType;
    MeetingRequest?: MeetingRequestMessageType;
    MeetingResponse?: MeetingResponseMessageType;
    MeetingCancellation?: MeetingCancellationMessageType;
    Task?: TaskType;
    PostItem?: PostItemType;
    RoleMember?: RoleMemberItemType;
    Network?: NetworkItemType;
    Person?: AbchPersonItemType;
}

/**
 * Append item field type.
 */
export class AppendToItemFieldType extends ItemChangeDescriptionType {
    Item?: ItemType;
    Message?: MessageType;
    SharingMessage?: SharingMessageType;
    CalendarItem?: CalendarItemType;
    Contact?: ContactItemType;
    DistributionList?: DistributionListType;
    MeetingMessage?: MeetingMessageType;
    MeetingRequest?: MeetingRequestMessageType;
    MeetingResponse?: MeetingResponseMessageType;
    MeetingCancellation?: MeetingCancellationMessageType;
    Task?: TaskType;
    PostItem?: PostItemType;
    RoleMember?: RoleMemberItemType;
    Network?: NetworkItemType;
    Person?: AbchPersonItemType;
}

/**
 * Represents a folder change.
 */
export abstract class FolderChangeDescriptionType extends ChangeDescriptionType { }

/**
 * Represents data to append to a folder property during an UpdateFolder operation.
 */
export class AppendToFolderFieldType extends FolderChangeDescriptionType {
    Folder?: FolderType;
    CalendarFolder?: CalendarFolderType;
    ContactsFolder?: ContactsFolderType;
    SearchFolder?: SearchFolderType;
    TasksFolder?: TasksFolderType;
}

/**
 * Represents an update to a single property on a folder in an UpdateFolder operation.
 */
export class SetFolderFieldType extends FolderChangeDescriptionType {
    Folder?: FolderType;
    CalendarFolder?: CalendarFolderType;
    ContactsFolder?: ContactsFolderType;
    SearchFolder?: SearchFolderType;
    TasksFolder?: TasksFolderType;
}

/**
 * Represents an operation to delete a given property from a folder during an UpdateFolder operation.
 */
export class DeleteFolderFieldType extends FolderChangeDescriptionType { }

/**
 * Identifies additional properties for use.
 */
export class NonEmptyArrayOfPathsToElementType extends GenericArray<BasePathToElementType> { }

/**
 * Base item id type.
 */
export abstract class BaseItemIdType { }

/**
 * Item id type.
 */
export class ItemIdType extends BaseItemIdType {
    constructor(public Id: string, public ChangeKey?: string) {
        super();
    }
}

/**
 * Root item id type used solely in DeleteAttachment responses.
 */
export class RootItemIdType extends BaseItemIdType {
    RootItemId: string;
    RootItemChangeKey: string;
}

/**
 * Array of item ids.
 */
export class ArrayOfBaseItemIdsType {
    ItemId: ItemIdType[] = [];
}

/**
 * Occurrence item id.
 */
export class OccurrenceItemIdType extends BaseItemIdType {
    RecurringMasterId: string;
    ChangeKey?: string;
    InstanceIndex: number;
}

/**
 * Recurring master item id.
 */
export class RecurringMasterItemIdType extends BaseItemIdType {
    OccurrenceId: string;
    ChangeKey?: string;
}

/**
 * Recurring master item ranges.
 */
export class RecurringMasterItemIdRangesType extends ItemIdType {
    Ranges?: ArrayOfOccurrenceRangesType;
}

/**
 * Array of occurrence ranges.
 */
export class ArrayOfOccurrenceRangesType {
    Range: OccurrencesRangeType[] = [];
}

/**
 * Occurrence range.
 */
export class OccurrencesRangeType {
    Start?: Date;
    End?: Date;
    Count?: number;
    CompareOriginalStartTime?: boolean;
}

/**
 * Base email address type.
 */
export abstract class BaseEmailAddressType {
}

/**
 * Represents an email address of person.
 */
export class EmailAddressType extends BaseEmailAddressType {
    Name?: string;
    EmailAddress: string;
    RoutingType?: string;
    MailboxType?: MailboxTypeType;
    ItemId?: ItemIdType;
    OriginalDisplayName?: string;
}

/**
 * Represents an array of email addresses.
 */
export class ArrayOfEmailAddressesType {
    Address: EmailAddressType[] = [];
}

/**
 * Array of alternate ids.
 */
export class NonEmptyArrayOfAlternateIdsType extends GenericArray<AlternateIdBaseType> { }

/**
 * Base alternate id type.
 */
export abstract class AlternateIdBaseType {
    Format: IdFormatType;
}

/**
 * Alternate id type.
 */
export class AlternateIdType extends AlternateIdBaseType {
    [AttributesKey] = {
        [XSITypeKey]: {
            type: 'AlternateIdType',
            xmlns: 'http://schemas.microsoft.com/exchange/services/2006/types',
        }
    }
    Id: string;
    Mailbox: string;
    IsArchive: boolean;
}

/**
 * Alternate public folder id type.
 */
export class AlternatePublicFolderIdType extends AlternateIdBaseType {
    [AttributesKey] = {
        [XSITypeKey]: {
            type: 'AlternatePublicFolderIdType',
            xmlns: 'http://schemas.microsoft.com/exchange/services/2006/types',
        }
    }
    FolderId: string;
}

/**
 * Alternate public folder item id type.
 */
export class AlternatePublicFolderItemIdType extends AlternatePublicFolderIdType {
    [AttributesKey] = {
        [XSITypeKey]: {
            type: 'AlternatePublicFolderItemIdType',
            xmlns: 'http://schemas.microsoft.com/exchange/services/2006/types',
        }
    }
    ItemId: string;
}

/**
 * Base folder id type.
 */
export abstract class BaseFolderIdType { }

/**
 * Represents folder id.
 */
export class FolderIdType extends BaseFolderIdType {
    constructor(public Id: string, public ChangeKey?: string) {
        super();
    }
}

/**
 * Identifies default Microsoft Exchange Server 2007 folders.
 * Using this element excludes the use of the FolderId element.
 */
export class DistinguishedFolderIdType extends BaseFolderIdType {
    Id: DistinguishedFolderIdNameType;
    Mailbox?: EmailAddressType;
    ChangeKey?: string;
}

/**
 * Represents address list id.
 */
export class AddressListIdType extends BaseFolderIdType {
    Id: string;
}

/**
 * Represents an array of folder ids.
 */
export class ArrayOfFolderIdType {
    FolderId?: FolderIdType[] = [];
}

/**
 * Array of base folder ids.
 */
export class NonEmptyArrayOfBaseFolderIdsType extends GenericArray<BaseFolderIdType>{
}

/**
 * Identifies the folder in which a new folder is created or the folder to search.
 */
export class TargetFolderIdType {
    FolderId?: FolderIdType;
    DistinguishedFolderId?: DistinguishedFolderIdType;
    AddressListId?: AddressListIdType;
}

/**
 * Represents the name of a user configuration object. The user configuration object name
 * is the identifier for a user configuration object.
 */
export class UserConfigurationNameType extends TargetFolderIdType {
    Name: string;
}

/**
 * Specifies an array of extended properties for a persona.
 */
export class ExtendedPropertyType {
    ExtendedFieldURI: PathToExtendedFieldType;
    Value?: string;
    Values?: NonEmptyArrayOfPropertyValuesType;
}

/**
 * Contains a collection of values for an extended property.
 */
export class NonEmptyArrayOfPropertyValuesType {
    Value: string[] = [];
}

/**
 * User id type.
 */
export class UserIdType {
    Sid?: string;
    PrimarySmtpAddress?: string;
    DisplayName?: string;
    DistinguishedUser?: DistinguishedUserType;
    ExternalUserIdentity?: string;
}

/**
 * Base permission type.
 */
export abstract class BasePermissionType {
    UserId: UserIdType;
    CanCreateItems?: boolean;
    CanCreateSubFolders?: boolean;
    IsFolderOwner?: boolean;
    IsFolderVisible?: boolean;
    IsFolderContact?: boolean;
    EditItems?: PermissionActionType;
    DeleteItems?: PermissionActionType;
}

/**
 * Permission type.
 */
export class PermissionType extends BasePermissionType {
    ReadItems?: PermissionReadAccessType;
    PermissionLevel: PermissionLevelType;
}

/**
 * Calendar permission type.
 */
export class CalendarPermissionType extends BasePermissionType {
    ReadItems?: CalendarPermissionReadAccessType;
    CalendarPermissionLevel: CalendarPermissionLevelType;
}

/**
 * Retention tag.
 */
export class RetentionTagType extends AttributeSimpleType<string> {
    constructor(_value: string, public IsExplicit: boolean) {
        super(_value);
    }
}

/**
 * Base folder type.
 */
export abstract class BaseFolderType {
    FolderId?: FolderIdType;
    ParentFolderId?: FolderIdType;
    FolderClass?: string;
    DisplayName?: string;
    TotalCount?: number;
    ChildFolderCount?: number;
    ExtendedProperty?: ExtendedPropertyType[] = [];
    ManagedFolderInformation?: ManagedFolderInformationType;
    EffectiveRights?: EffectiveRightsType;
    DistinguishedFolderId?: DistinguishedFolderIdNameType;
    PolicyTag?: RetentionTagType;
    ArchiveTag?: RetentionTagType;
    ReplicaList?: ArrayOfStringsType;
}

/**
 * Array of folders (used in response).
 */
export class ArrayOfFoldersType extends GenericArray<BaseFolderType> { }

/**
 * Array of folders (used in request).
 */
export class NonEmptyArrayOfFoldersType extends GenericArray<BaseFolderType> { }

/**
 * Array of folders (used in request).
 */
export class NonEmptyArrayOfFolderNamesType { 
    FolderName: string[] = [];
}

/**
 * Find folder parent type.
 */
export class FindFolderParentType {
    Folders?: ArrayOfFoldersType;
    IndexedPagingOffset?: number;
    NumeratorOffset?: number;
    AbsoluteDenominator?: number;
    IncludesLastItemInRange?: boolean;
    TotalItemsInView?: number;
}

/**
 * Permisson set.
 */
export class PermissionSetType {
    Permissions: ArrayOfPermissionsType;
    UnknownEntries?: ArrayOfUnknownEntriesType;
}

/**
 * Calendar permission set.
 */
export class CalendarPermissionSetType {
    CalendarPermissions: ArrayOfCalendarPermissionsType;
    UnknownEntries?: ArrayOfUnknownEntriesType;
}

/**
 * Array of permissions.
 */
export class ArrayOfPermissionsType {
    Permission: PermissionType[] = [];
}

/**
 * Array of unknown entries.
 */
export class ArrayOfUnknownEntriesType {
    UnknownEntry: string[] = [];
}

/**
 * Array of calendar permissions.
 */
export class ArrayOfCalendarPermissionsType {
    CalendarPermission: CalendarPermissionType[] = [];
}

/**
 * Folder type.
 */
export class FolderType extends BaseFolderType {
    PermissionSet?: PermissionSetType;
    UnreadCount?: number;
}

/**
 * Search paramters type.
 */
export class SearchParametersType {
    Restriction: RestrictionType;
    BaseFolderIds: NonEmptyArrayOfBaseFolderIdsType;
    Traversal: SearchFolderTraversalType;
}

/**
 * Search folder.
 */
export class SearchFolderType extends FolderType {
    SearchParameters?: SearchParametersType;
}

/**
 * Tasks folder.
 */
export class TasksFolderType extends FolderType { }

/**
 * Calendar folder.
 */
export class CalendarFolderType extends BaseFolderType {
    PermissionSet?: CalendarPermissionSetType;
    SharingEffectiveRights?: CalendarPermissionReadAccessType;
}

/**
 * Contacts folder.
 */
export class ContactsFolderType extends BaseFolderType {
    PermissionSet?: PermissionSetType;
    SharingEffectiveRights?: PermissionReadAccessType;
    SourceId?: string;
    AccountName?: string;
}

/**
 * Managed folder information.
 */
export class ManagedFolderInformationType {
    CanDelete?: boolean;
    CanRenameOrMove?: boolean;
    MustDisplayComment?: boolean;
    HasQuota?: boolean;
    IsManagedFoldersRoot?: boolean;
    ManagedFolderId?: string;
    Comment?: string;
    StorageQuota?: number;
    FolderSize?: number;
    HomePage?: string;
}

/**
 * Effective rights.
 */
export class EffectiveRightsType {
    CreateAssociated: boolean;
    CreateContents: boolean;
    CreateHierarchy: boolean;
    Delete: boolean;
    Modify: boolean;
    Read: boolean;
    ViewPrivateItems?: boolean;
}

/**
 * Identifies a set of properties to return in a GetItem operation, FindItem operation, or SyncFolderItems operation response.
 */
export class ItemResponseShapeType {
    BaseShape: DefaultShapeNamesType;
    IncludeMimeContent?: boolean;
    BodyType?: BodyTypeResponseType;
    UniqueBodyType?: BodyTypeResponseType;
    NormalizedBodyType?: BodyTypeResponseType;
    FilterHtmlContent?: boolean;
    ConvertHtmlCodePageToUTF8?: boolean;
    InlineImageUrlTemplate?: string;
    BlockExternalImages?: boolean;
    AddBlankTargetToLinks?: boolean;
    MaximumBodySize?: number;
    AdditionalProperties?: NonEmptyArrayOfPathsToElementType;
}

/**
 * Identifies additional properties to return in a response to a GetAttachment request.
 */
export class AttachmentResponseShapeType {
    IncludeMimeContent?: boolean;
    BodyType?: BodyTypeResponseType;
    FilterHtmlContent?: boolean;
    AdditionalProperties?: NonEmptyArrayOfPathsToElementType;
}

/**
 * Contains the ASCII MIME stream of an object that is represented in base64Binary format and supports [RFC2045].
 */
export class MimeContentType extends AttributeSimpleType<string> {
    constructor(_value: string, public CharacterSet?: string) {
        super(_value);
    }
}

/**
 * This type is used in GetAttachment.
 */
export class RequestAttachmentIdType extends BaseItemIdType {
    Id: string;
}

/**
 * This type is used in CreateAttachment responses. CreateAttachment modifies the
 * change key of the root item of the created attachment. We have to return that new change key.
 */
export class AttachmentIdType extends RequestAttachmentIdType {
    RootItemId?: string;
    RootItemChangeKey?: string;
}

/**
 * Contains an array of attachment identifiers.
 */
export class NonEmptyArrayOfRequestAttachmentIdsType {
    AttachmentId: RequestAttachmentIdType[] = [];
}

/**
 * Attachment type.
 */
export abstract class AttachmentType {
    AttachmentId?: AttachmentIdType;
    Name?: string;
    ContentType?: string;
    ContentId?: string;
    ContentLocation?: string;
    AttachmentOriginalUrl?: string;
    Size?: number;
    LastModifiedTime?: Date;
    IsInline?: boolean;
}

/**
 * File Attachment type.
 */
export class FileAttachmentType extends AttachmentType {
    IsContactPhoto?: boolean;
    Content?: string; // xs:base64Binary
}

/**
 * Item Attachment type.
 */
export class ItemAttachmentType extends AttachmentType {
    Item?: ItemType;
    Message?: MessageType;
    SharingMessage?: SharingMessageType;
    CalendarItem?: CalendarItemType;
    Contact?: ContactItemType;
    MeetingMessage?: MeetingMessageType;
    MeetingRequest?: MeetingRequestMessageType;
    MeetingResponse?: MeetingResponseMessageType;
    MeetingCancellation?: MeetingCancellationMessageType;
    Task?: TaskType;
    PostItem?: PostItemType;
    RoleMember?: RoleMemberItemType;
    Network?: NetworkItemType;
    Person?: AbchPersonItemType;
}

/**
 * Reference Attachment type.
 */
export class ReferenceAttachmentType extends AttachmentType {
    AttachLongPathName?: string;
    ProviderType?: string;
    ProviderEndpointUrl?: string;
    AttachmentThumbnailUrl?: string;
    AttachmentPreviewUrl?: string;
    PermissionType?: number;
    OriginalPermissionType?: number;
    AttachmentIsFolder?: boolean;
}

/**
 * Contains the items or files that are attached to an item in the Exchange store.
 */
export class ArrayOfAttachmentsType extends GenericArray<ItemAttachmentType | FileAttachmentType | ReferenceAttachmentType> { }

/**
 * Contains the items or files that are attached to an item in the Exchange store.
 */
export class NonEmptyArrayOfAttachmentsType extends GenericArray<FileAttachmentType | ItemAttachmentType | ReferenceAttachmentType> { }

/**
 * Represents the actual body content of a message.
 */
export class BodyType extends AttributeSimpleType<string> {
    // eslint-disable-next-line no-shadow
    constructor(_value: string, public BodyType: BodyTypeType, public IsTruncated?: boolean) {
        super(_value);
    }
}

/**
 * Represents a generic item in the Exchange store.
 */
export class ItemType {
    MimeContent?: MimeContentType;
    ItemId?: ItemIdType;
    ParentFolderId?: FolderIdType;
    ItemClass?: string;
    Subject?: string;
    Sensitivity?: SensitivityChoicesType;
    Body?: BodyType;
    Attachments?: NonEmptyArrayOfAttachmentsType;
    DateTimeReceived?: Date;
    Size?: number;
    Categories?: ArrayOfStringsType;
    Importance?: ImportanceChoicesType;
    InReplyTo?: string;
    IsSubmitted?: boolean;
    IsDraft?: boolean;
    IsFromMe?: boolean;
    IsResend?: boolean;
    IsUnmodified?: boolean;
    InternetMessageHeaders?: NonEmptyArrayOfInternetHeadersType;
    DateTimeSent?: Date;
    DateTimeCreated?: Date;
    ResponseObjects?: NonEmptyArrayOfResponseObjectsType;
    ReminderDueBy?: Date;
    ReminderIsSet?: boolean;
    ReminderNextTime?: Date;
    ReminderMinutesBeforeStart?: string;
    DisplayCc?: string;
    DisplayTo?: string;
    DisplayBcc?: string;
    HasAttachments?: boolean;
    ExtendedProperty?: ExtendedPropertyType[] = [];
    Culture?: string;
    EffectiveRights?: EffectiveRightsType;
    LastModifiedName?: string;
    LastModifiedTime?: Date;
    IsAssociated?: boolean;
    WebClientReadFormQueryString?: string;
    WebClientEditFormQueryString?: string;
    ConversationId?: ItemIdType;
    UniqueBody?: BodyType;
    Flag?: FlagType;
    StoreEntryId?: string; // xs:base64Binary
    InstanceKey?: string; // xs:base64Binary
    NormalizedBody?: BodyType;
    EntityExtractionResult?: EntityExtractionResultType;
    PolicyTag?: RetentionTagType;
    ArchiveTag?: RetentionTagType;
    RetentionDate?: Date;
    Preview?: string;
    RightsManagementLicenseData?: RightsManagementLicenseDataType;
    PredictedActionReasons?: NonEmptyArrayOfPredictedActionReasonType;
    IsClutter?: boolean;
    BlockStatus?: boolean;
    HasBlockedImages?: boolean;
    TextBody?: BodyType;
    IconIndex?: IconIndexType;
    SearchKey?: string; // xs:base64Binary
    SortKey?: number;
    Hashtags?: ArrayOfStringsType;
    Mentions?: ArrayOfRecipientsType;
    MentionedMe?: boolean;
    MentionsPreview?: MentionsPreviewType;
    MentionsEx?: NonEmptyArrayOfMentionActionsType;
    AppliedHashtags?: NonEmptyArrayOfAppliedHashtagType;
    AppliedHashtagsPreview?: AppliedHashtagsPreviewType;
    Likes?: NonEmptyArrayOfLikeType;
    LikesPreview?: LikesPreviewType;
    PendingSocialActivityTagIds?: ArrayOfStringsType;
    AtAllMention?: boolean;
    CanDelete?: boolean;
    InferenceClassification?: InferenceClassificationType;
}

/**
 * Message item type.
 */
export class MessageType extends ItemType {
    Sender?: SingleRecipientType;
    ToRecipients?: ArrayOfRecipientsType;
    CcRecipients?: ArrayOfRecipientsType;
    BccRecipients?: ArrayOfRecipientsType;
    IsReadReceiptRequested?: boolean;
    IsDeliveryReceiptRequested?: boolean;
    ConversationIndex?: string; // xs:base64Binary
    ConversationTopic?: string;
    From?: SingleRecipientType;
    InternetMessageId?: string;
    IsRead?: boolean;
    IsResponseRequested?: boolean;
    References?: string;
    ReplyTo?: ArrayOfRecipientsType;
    ReceivedBy?: SingleRecipientType;
    ReceivedRepresenting?: SingleRecipientType;
    ApprovalRequestData?: ApprovalRequestDataType;
    VotingInformation?: VotingInformationType;
    ReminderMessageData?: ReminderMessageDataType;
    MessageSafety?: MessageSafetyType;
    SenderSMTPAddress?: SmtpAddressType;
    MailboxGuids?: MailboxGuids;
    PublishedCalendarItemIcs?: string;
    PublishedCalendarItemName?: string;
}

/**
 * Contact item type.
 */
export class ContactItemType extends ItemType {
    FileAs?: string;
    FileAsMapping?: FileAsMappingType;
    DisplayName?: string;
    GivenName?: string;
    Initials?: string;
    MiddleName?: string;
    Nickname?: string;
    CompleteName?: CompleteNameType;
    CompanyName?: string;
    EmailAddresses?: EmailAddressDictionaryType;
    AbchEmailAddresses?: AbchEmailAddressDictionaryType;
    PhysicalAddresses?: PhysicalAddressDictionaryType;
    PhoneNumbers?: PhoneNumberDictionaryType;
    AssistantName?: string;
    Birthday?: Date;
    BusinessHomePage?: string;
    Children?: ArrayOfStringsType;
    Companies?: ArrayOfStringsType;
    ContactSource?: ContactSourceType;
    Department?: string;
    Generation?: string;
    ImAddresses?: ImAddressDictionaryType;
    JobTitle?: string;
    Manager?: string;
    Mileage?: string;
    OfficeLocation?: string;
    PostalAddressIndex?: PhysicalAddressIndexType;
    Profession?: string;
    SpouseName?: string;
    Surname?: string;
    WeddingAnniversary?: Date;
    HasPicture?: boolean;
    PhoneticFullName?: string;
    PhoneticFirstName?: string;
    PhoneticLastName?: string;
    Alias?: string;
    Notes?: string;
    Photo?: string; // xs:base64Binary
    UserSMIMECertificate?: ArrayOfBinaryType;
    DirectoryId?: string;
    ManagerMailbox?: SingleRecipientType;
    DirectReports?: ArrayOfRecipientsType;
    AccountName?: string;
    IsAutoUpdateDisabled?: boolean;
    IsMessengerEnabled?: boolean;
    Comment?: string;
    ContactShortId?: number;
    ContactType?: string;
    Gender?: string;
    IsHidden?: boolean;
    ObjectId?: string;
    PassportId?: number;
    IsPrivate?: boolean;
    SourceId?: string;
    TrustLevel?: number;
    CreatedBy?: string;
    Urls?: ContactUrlDictionaryType;
    Cid?: number;
    SkypeAuthCertificate?: string;
    SkypeContext?: string;
    SkypeId?: string;
    SkypeRelationship?: string;
    YomiNickname?: string;
    XboxLiveTag?: string;
    InviteFree?: boolean;
    HidePresenceAndProfile?: boolean;
    IsPendingOutbound?: boolean;
    SupportGroupFeeds?: boolean;
    UserTileHash?: string;
    UnifiedInbox?: boolean;
    Mris?: ArrayOfStringsType;
    Wlid?: string;
    AbchContactId?: string;
    NotInBirthdayCalendar?: boolean;
    ShellContactType?: string;
    ImMri?: string;
    PresenceTrustLevel?: number;
    OtherMri?: string;
    ProfileLastChanged?: string;
    MobileIMEnabled?: boolean;
    PartnerNetworkProfilePhotoUrl?: string;
    PartnerNetworkThumbnailPhotoUrl?: string;
    PersonId?: string;
    ConversationGuid?: string;
    MsexchangeCertificate?: ArrayOfBinaryType;
}

/**
 * Calendar item type.
 */
export class CalendarItemType extends ItemType {
    UID?: string;
    RecurrenceId?: Date;
    RecurrenceIdString?: string;
    DateTimeStamp?: Date;
    Start?: Date;
    StartString?: string;
    End?: Date;
    EndString?: string; 
    OriginalStart?: Date;
    OriginalStartString?: string;
    IsAllDayEvent?: boolean;
    LegacyFreeBusyStatus?: LegacyFreeBusyType;
    Location?: string;
    When?: string;
    IsMeeting?: boolean;
    IsCancelled?: boolean;
    IsRecurring?: boolean;
    MeetingRequestWasSent?: boolean;
    IsResponseRequested?: boolean;
    CalendarItemType?: CalendarItemTypeType;
    MyResponseType?: ResponseTypeType;
    Organizer?: SingleRecipientType;
    RequiredAttendees?: NonEmptyArrayOfAttendeesType;
    OptionalAttendees?: NonEmptyArrayOfAttendeesType;
    Resources?: NonEmptyArrayOfAttendeesType;
    InboxReminders?: ArrayOfInboxReminderType;
    ConflictingMeetingCount?: number;
    AdjacentMeetingCount?: number;
    ConflictingMeetings?: NonEmptyArrayOfAllItemsType;
    AdjacentMeetings?: NonEmptyArrayOfAllItemsType;
    Duration?: string;
    TimeZone?: string;
    AppointmentReplyTime?: Date;
    AppointmentReplyTimeString?: string;
    AppointmentSequenceNumber?: number;
    AppointmentState?: number;
    Recurrence?: RecurrenceType;
    FirstOccurrence?: OccurrenceInfoType;
    LastOccurrence?: OccurrenceInfoType;
    ModifiedOccurrences?: NonEmptyArrayOfOccurrenceInfoType;
    DeletedOccurrences?: NonEmptyArrayOfDeletedOccurrencesType;
    MeetingTimeZone?: TimeZoneType;
    StartTimeZone?: TimeZoneDefinitionType;
    EndTimeZone?: TimeZoneDefinitionType;
    ConferenceType?: number;
    AllowNewTimeProposal?: boolean;
    IsOnlineMeeting?: boolean;
    MeetingWorkspaceUrl?: string;
    NetShowUrl?: string;
    EnhancedLocation?: EnhancedLocationType;
    StartWallClock?: Date;
    StartWallClockString?: string;
    EndWallClock?: Date;
    EndWallClockString?: string;
    StartTimeZoneId?: string;
    EndTimeZoneId?: string;
    IntendedFreeBusyStatus?: LegacyFreeBusyType;
    JoinOnlineMeetingUrl?: string;
    OnlineMeetingSettings?: OnlineMeetingSettingsType;
    IsOrganizer?: boolean;
    CalendarActivityData?: CalendarActivityDataType;
    DoNotForwardMeeting?: boolean;
}

/**
 * Task item type.
 */
export class TaskType extends ItemType {
    ActualWork?: number;
    AssignedTime?: Date;
    BillingInformation?: string;
    ChangeCount?: number;
    Companies?: ArrayOfStringsType;
    CompleteDate?: Date;
    Contacts?: ArrayOfStringsType;
    DelegationState?: TaskDelegateStateType;
    Delegator?: string;
    DueDate?: Date;
    IsAssignmentEditable?: number;
    IsComplete?: boolean;
    IsRecurring?: boolean;
    IsTeamTask?: boolean;
    Mileage?: string;
    Owner?: string;
    PercentComplete?: number;
    Recurrence?: TaskRecurrenceType;
    StartDate?: Date;
    Status?: TaskStatusType;
    StatusDescription?: string;
    TotalWork?: number;
}

/**
 * Post item type.
 */
export class PostItemType extends ItemType {
    ConversationIndex?: string; // xs:base64Binary
    ConversationTopic?: string;
    From?: SingleRecipientType;
    InternetMessageId?: string;
    IsRead?: boolean;
    PostedTime?: Date;
    References?: string;
    Sender?: SingleRecipientType;
}

/**
 * Role member item type.
 */
export class RoleMemberItemType extends ItemType {
    DisplayName?: string;
    Type?: RoleMemberTypeType;
    MemberId?: string;
}

/**
 * Network item type.
 */
export class NetworkItemType extends ItemType {
    DomainId?: number;
    DomainTag?: string;
    UserTileUrl?: string;
    ProfileUrl?: string;
    Settings?: number;
    IsDefault?: boolean;
    AutoLinkError?: string;
    AutoLinkSuccess?: string;
    UserEmail?: string;
    ClientPublishSecret?: string;
    ClientToken?: string;
    ClientToken2?: string;
    ContactSyncError?: string;
    ContactSyncSuccess?: string;
    ErrorOffers?: number;
    FirstAuthErrorDates?: string;
    LastVersionSaved?: number;
    LastWelcomeContact?: string;
    Offers?: number;
    PsaLastChanged?: Date;
    RefreshToken2?: string;
    RefreshTokenExpiry2?: string;
    SessionHandle?: string;
    RejectedOffers?: number;
    SyncEnabled?: boolean;
    TokenRefreshLastAttempted?: Date;
    TokenRefreshLastCompleted?: Date;
    PsaState?: string;
    SourceEntryID?: string; // xs:base64Binary
    AccountName?: string;
    LastSync?: Date;
}

/**
 * One drive item type.
 */
export class OneDriveItemType extends ItemType {
    ResourceId: string;
}

/**
 * Delve item type.
 */
export class DelveItemType extends ItemType {
    GraphNodeLogicalId: string;
}

/**
 * File item type.
 */
export class FileItemType extends ItemType {
    FileName?: string;
    FileExtension?: string;
    FileSize?: number;
    FileCreatedTime?: string;
    FileModifiedTime?: string;
    StorageProviderContext?: string;
    FileID?: string;
    ItemReferenceId?: string;
    ReferenceId?: string;
    Sender?: string;
    ItemReceivedTime?: string;
    ItemPath?: string;
    ItemSentTime?: string;
    FileContexts?: ArrayOfStringsType;
    VisualizationContainerUrl?: string;
    VisualizationContainerTitle?: string;
    VisualizationAccessUrl?: string;
    ReferenceAttachmentProviderEndpoint?: string;
    ReferenceAttachmentProviderType?: string;
    ItemConversationId?: string;
    SharepointItemListId?: string;
    SharepointItemListItemId?: string;
    SharepointItemSiteId?: string;
    SharepointItemSitePath?: string;
    SharepointItemWebId?: string;
    AttachmentId?: string;
}

/**
 * Abch person item type.
 */
export class AbchPersonItemType extends ItemType {
    AntiLinkInfo?: string;
    PersonId?: string;
    ContactHandles?: ArrayOfAbchPersonContactHandlesType;
    ContactCategories?: ArrayOfStringsType;
    RelevanceOrder1?: string;
    RelevanceOrder2?: string;
    TrustLevel?: number;
    FavoriteOrder?: number;
    ExchangePersonIdGuid?: string;
}

/**
 * Distribution list type.
 */
export class DistributionListType extends ItemType {
    DisplayName?: string;
    FileAs?: string;
    ContactSource?: ContactSourceType;
    Members?: MembersListType;
}

/**
 * Response object core type.
 */
export abstract class ResponseObjectCoreType extends MessageType {
    ReferenceItemId?: ItemIdType;
}

/**
 * Response object type.
 */
export abstract class ResponseObjectType extends ResponseObjectCoreType {
    ObjectName?: string;
}

/**
 * Smart response base type.
 */
export class SmartResponseBaseType extends ResponseObjectType { }

/**
 * Smart response type.
 */
export class SmartResponseType extends SmartResponseBaseType {
    NewBodyContent?: BodyType;
}

/**
 * Reply to item type.
 */
export class ReplyToItemType extends SmartResponseType { }

/**
 * Reply all to item type.
 */
export class ReplyAllToItemType extends SmartResponseType {
    IsSpecificMessageReply?: boolean;
}

/**
 * Forward item type.
 */
export class ForwardItemType extends SmartResponseType { }

/**
 * Cancel calendar item type.
 */
export class CancelCalendarItemType extends SmartResponseType { }

/**
 * Propose new time type.
 */
export class ProposeNewTimeType extends ResponseObjectType { }

/**
 * Remove item type.
 */
export class RemoveItemType extends ResponseObjectType { }

/**
 * Add item to my calendar type.
 */
export class AddItemToMyCalendarType extends ResponseObjectType { }

/**
 * Reference item response type.
 */
export class ReferenceItemResponseType extends ResponseObjectType { }

/**
 * Suppress read receipt type.
 */
export class SuppressReadReceiptType extends ReferenceItemResponseType { }

/**
 * Accept sharing invitation type.
 */
export class AcceptSharingInvitationType extends ReferenceItemResponseType { }

/**
 * Post reply item base type.
 */
export class PostReplyItemBaseType extends ResponseObjectType { }

/**
 * Post reply item type.
 */
export class PostReplyItemType extends PostReplyItemBaseType {
    NewBodyContent?: BodyType;
}

/**
 * Well known response object type.
 */
export class WellKnownResponseObjectType extends ResponseObjectType { }

/**
 * Meeting registration response object type.
 */
export class MeetingRegistrationResponseObjectType extends WellKnownResponseObjectType {
    ProposedStart?: Date;
    ProposedEnd?: Date;
}

/**
 * Accept item type.
 */
export class AcceptItemType extends MeetingRegistrationResponseObjectType { }

/**
 * Tentative accept item type.
 */
export class TentativelyAcceptItemType extends MeetingRegistrationResponseObjectType { }

/**
 * Decline item type.
 */
export class DeclineItemType extends MeetingRegistrationResponseObjectType { }

/**
 * Array of response objects.
 */
export class NonEmptyArrayOfResponseObjectsType extends GenericArray<
    AcceptItemType | TentativelyAcceptItemType | DeclineItemType | ReplyToItemType
    | ForwardItemType | ReplyAllToItemType | CancelCalendarItemType | RemoveItemType
    | SuppressReadReceiptType | PostReplyItemType | AcceptSharingInvitationType | AddItemToMyCalendarType | ProposeNewTimeType> { }

/**
 * Sharing message type.
 */
export class SharingMessageType extends MessageType {
    SharingMessageAction?: SharingMessageActionType;
    SharingMessageActions?: ArrayOfSharingMessageActionType;
}

/**
 * Meeting message type.
 */
export class MeetingMessageType extends MessageType {
    AssociatedCalendarItemId?: ItemIdType;
    IsDelegated?: boolean;
    IsOutOfDate?: boolean;
    HasBeenProcessed?: boolean;
    ResponseType?: ResponseTypeType;
    Uid?: string;
    RecurrenceId?: Date;
    DateTimeStamp?: Date;
    IsOrganizer?: boolean;
}

/**
 * Meeting request message type.
 */
export class MeetingRequestMessageType extends MeetingMessageType {
    MeetingRequestType?: MeetingRequestTypeType;
    IntendedFreeBusyStatus?: LegacyFreeBusyType;
    Start?: Date;
    StartString?: string;
    End?: Date;
    EndString?: string;
    OriginalStart?: Date;
    OriginalStartString?: string;
    IsAllDayEvent?: boolean;
    LegacyFreeBusyStatus?: LegacyFreeBusyType;
    Location?: string;
    When?: string;
    IsMeeting?: boolean;
    IsCancelled?: boolean;
    IsRecurring?: boolean;
    MeetingRequestWasSent?: boolean;
    CalendarItemType?: CalendarItemTypeType;
    MyResponseType?: ResponseTypeType;
    Organizer?: SingleRecipientType;
    RequiredAttendees?: NonEmptyArrayOfAttendeesType;
    OptionalAttendees?: NonEmptyArrayOfAttendeesType;
    Resources?: NonEmptyArrayOfAttendeesType;
    ConflictingMeetingCount?: number;
    AdjacentMeetingCount?: number;
    ConflictingMeetings?: NonEmptyArrayOfAllItemsType;
    AdjacentMeetings?: NonEmptyArrayOfAllItemsType;
    Duration?: string;
    TimeZone?: string;
    AppointmentReplyTime?: Date;
    AppointmentReplyTimeString?: string;
    AppointmentSequenceNumber?: number;
    AppointmentState?: number;
    Recurrence?: RecurrenceType;
    FirstOccurrence?: OccurrenceInfoType;
    LastOccurrence?: OccurrenceInfoType;
    ModifiedOccurrences?: NonEmptyArrayOfOccurrenceInfoType;
    DeletedOccurrences?: NonEmptyArrayOfDeletedOccurrencesType;
    MeetingTimeZone?: TimeZoneType;
    StartTimeZone?: TimeZoneDefinitionType;
    EndTimeZone?: TimeZoneDefinitionType;
    ConferenceType?: number;
    AllowNewTimeProposal?: boolean;
    IsOnlineMeeting?: boolean;
    MeetingWorkspaceUrl?: string;
    NetShowUrl?: string;
    EnhancedLocation?: EnhancedLocationType;
    ChangeHighlights?: ChangeHighlightsType;
    StartWallClock?: Date;
    StartWallClockString?: string;
    EndWallClock?: Date;
    EndWallClockString?: string;
    StartTimeZoneId?: string;
    EndTimeZoneId?: string;
    DoNotForwardMeeting?: boolean;
}

/**
 * Meeting response message type.
 */
export class MeetingResponseMessageType extends MeetingMessageType {
    Start?: Date;
    End?: Date;
    Location?: string;
    Recurrence?: RecurrenceType;
    CalendarItemType?: string;
    ProposedStart?: Date;
    ProposedEnd?: Date;
    EnhancedLocation?: EnhancedLocationType;
}

/**
 * Meeting cancellation message type.
 */
export class MeetingCancellationMessageType extends MeetingMessageType {
    Start?: Date;
    End?: Date;
    Location?: string;
    Recurrence?: RecurrenceType;
    CalendarItemType?: string;
    EnhancedLocation?: EnhancedLocationType;
    DoNotForwardMeeting?: boolean;
}

/**
 * Array of all items.
 */
export class NonEmptyArrayOfAllItemsType extends GenericArray<
    ItemType | MessageType | SharingMessageType | CalendarItemType | ContactItemType | DistributionListType |
    MeetingMessageType | MeetingRequestMessageType | MeetingResponseMessageType | MeetingCancellationMessageType |
    TaskType | PostItemType | ReplyToItemType | ForwardItemType | ReplyAllToItemType | AcceptItemType |
    TentativelyAcceptItemType | DeclineItemType | CancelCalendarItemType | RemoveItemType | SuppressReadReceiptType |
    PostReplyItemType | AcceptSharingInvitationType | RoleMemberItemType | NetworkItemType | AbchPersonItemType> {
}

/**
 * Abch person contact handle.
 */
export class AbchPersonContactHandle {
    SourceId: string;
    ObjectId: string;
    AccountName?: string;
}

/**
 * Array of Abch person contact handles.
 */
export class ArrayOfAbchPersonContactHandlesType {
    ContactHandle: AbchPersonContactHandle[] = [];
}

/**
 * Member type.
 */
export class MemberType {
    Mailbox?: EmailAddressType;
    Status?: MemberStatusType;
    Key?: string;
}

/**
 * Members list type.
 */
export class MembersListType {
    Member: MemberType[] = [];
}

/**
 * Array of predicted action reason.
 */
export class NonEmptyArrayOfPredictedActionReasonType {
    PredictedActionReason: PredictedActionReasonType[] = [];
}

/**
 * Email address extended type.
 */
export class EmailAddressExtendedType extends EmailAddressType {
    ExternalObjectId?: string;
    PrimaryEmailAddress?: string;
}

/**
 * Change highlights type.
 */
export class ChangeHighlightsType {
    HasLocationChanged?: boolean;
    Location?: string;
    HasStartTimeChanged?: boolean;
    Start?: Date;
    HasEndTimeChanged?: boolean;
    End?: Date;
}

/**
 * Mention action type.
 */
export class MentionActionType {
    Id: string;
    CreatedBy: EmailAddressExtendedType;
    CreatedDateTime?: string;
    ServerCreatedDateTime?: string;
    DeepLink?: string;
    Application?: string;
    Mentioned: EmailAddressExtendedType;
    MentionText?: string;
    ClientReference?: string;
}

/**
 * Array of mention actions type.
 */
export class NonEmptyArrayOfMentionActionsType {
    MentionAction: MentionActionType[] = [];
}

/**
 * Applied hashtag type.
 */
export class AppliedHashtagType {
    Id: string;
    CreatedBy: EmailAddressExtendedType;
    CreatedDateTime?: string;
    ServerCreatedDateTime?: string;
    DeepLink?: string;
    Application?: string;
    Tag: string;
    IsAutoTagged: boolean;
    IsInlined: boolean;
}

/**
 * Array of applied hashtag type.
 */
export class NonEmptyArrayOfAppliedHashtagType {
    AppliedHashtag: AppliedHashtagType[] = [];
}

/**
 * Link type.
 */
export class LikeType {
    Id: string;
    CreatedBy: EmailAddressExtendedType;
    CreatedDateTime?: string;
    ServerCreatedDateTime?: string;
    DeepLink?: string;
    Application?: string;
}

/**
 * Array of link types.
 */
export class NonEmptyArrayOfLikeType {
    Like: LikeType[] = [];
}

/**
 * Rights management license data type.
 */
export class RightsManagementLicenseDataType {
    RightsManagedMessageDecryptionStatus?: number;
    RmsTemplateId?: string;
    TemplateName?: string;
    TemplateDescription?: string;
    EditAllowed?: boolean;
    ReplyAllowed?: boolean;
    ReplyAllAllowed?: boolean;
    ForwardAllowed?: boolean;
    ModifyRecipientsAllowed?: boolean;
    ExtractAllowed?: boolean;
    PrintAllowed?: boolean;
    ExportAllowed?: boolean;
    ProgrammaticAccessAllowed?: boolean;
    IsOwner?: boolean;
    ContentOwner?: string;
    ContentExpiryDate?: string;
}

/**
 * Mentions preview type.
 */
export class MentionsPreviewType {
    IsMentioned: boolean;
}

/**
 * Likes preview type.
 */
export class LikesPreviewType {
    LikeCount: number;
}

/**
 * Applied hashtags preview type.
 */
export class AppliedHashtagsPreviewType {
    Hashtags: ArrayOfStringsType;
}

/**
 * Entity type.
 */
export class EntityType {
    Position: EmailPositionType[] = [];
}

/**
 * Email user type.
 */
export class EmailUserType {
    Name?: string;
    UserId?: string;
}

/**
 * Array of email users type.
 */
export class ArrayOfEmailUsersType {
    EmailUser: EmailUserType[] = [];
}

/**
 * Phone type.
 */
export class PhoneType {
    OriginalPhoneString?: string;
    PhoneString?: string;
    Type?: string;
}

/**
 * Array of phones type.
 */
export class ArrayOfPhonesType {
    Phone: PhoneType[] = [];
}

/**
 * Array of urls type.
 */
export class ArrayOfUrlsType {
    Url: string[] = [];
}

/**
 * Array of extracted email addresses.
 */
export class ArrayOfExtractedEmailAddresses {
    EmailAddress: string[] = [];
}

/**
 * Array of addresses type.
 */
export class ArrayOfAddressesType {
    Address: string[] = [];
}

/**
 * Address entity type.
 */
export class AddressEntityType extends EntityType {
    Address?: string;
}

/**
 * Meeting suggestion type.
 */
export class MeetingSuggestionType extends EntityType {
    Attendees?: ArrayOfEmailUsersType;
    Location?: string;
    Subject?: string;
    MeetingString?: string;
    StartTime?: Date;
    EndTime?: Date;
    TimeStringBeginIndex?: number;
    TimeStringLength?: number;
    EntityId?: string;
    ExtractionId?: string;
}

/**
 * Task suggestion type.
 */
export class TaskSuggestionType extends EntityType {
    TaskString?: string;
    Assignees?: ArrayOfEmailUsersType;
}

/**
 * Email address entity type.
 */
export class EmailAddressEntityType extends EntityType {
    EmailAddress?: string;
}

/**
 * Contact type.
 */
export class ContactType extends EntityType {
    PersonName?: string;
    BusinessName?: string;
    PhoneNumbers?: ArrayOfPhonesType;
    Urls?: ArrayOfUrlsType;
    EmailAddresses?: ArrayOfExtractedEmailAddresses;
    Addresses?: ArrayOfAddressesType;
    ContactString?: string;
}

/**
 * Url entity type.
 */
export class UrlEntityType extends EntityType {
    Url?: string;
}

/**
 * Phone entity type.
 */
export class PhoneEntityType extends EntityType {
    OriginalPhoneString?: string;
    PhoneString?: string;
    Type?: string;
}

/**
 * Array of address entities type.
 */
export class ArrayOfAddressEntitiesType {
    AddressEntity: AddressEntityType[] = [];
}

/**
 * Array of meeting suggestions type.
 */
export class ArrayOfMeetingSuggestionsType {
    MeetingSuggestion: MeetingSuggestionType[] = [];
}

/**
 * Array of task suggestions type.
 */
export class ArrayOfTaskSuggestionsType {
    TaskSuggestion: TaskSuggestionType[] = [];
}

/**
 * Array of email address entities type.
 */
export class ArrayOfEmailAddressEntitiesType {
    EmailAddressEntity: EmailAddressEntityType[] = [];
}

/**
 * Array of contacts type.
 */
export class ArrayOfContactsType {
    Contact: ContactType[] = [];
}

/**
 * Array of url entities type.
 */
export class ArrayOfUrlEntitiesType {
    UrlEntity: UrlEntityType[] = [];
}

/**
 * Array of phone entities type.
 */
export class ArrayOfPhoneEntitiesType {
    Phone: PhoneEntityType[] = [];
}

/**
 * Parcel delivery entity type.
 */
export class ParcelDeliveryEntityType {
    Carrier?: string;
    TrackingNumber?: string;
    TrackingUrl?: string;
    ExpectedArrivalFrom?: string;
    ExpectedArrivalUntil?: string;
    Product?: string;
    ProductUrl?: string;
    ProductImage?: string;
    ProductSku?: string;
    ProductDescription?: string;
    ProductBrand?: string;
    ProductColor?: string;
    OrderNumber?: string;
    Seller?: string;
    OrderStatus?: string;
    AddressName?: string;
    StreetAddress?: string;
    AddressLocality?: string;
    AddressRegion?: string;
    AddressCountry?: string;
    PostalCode?: string;
}

/**
 * Array of parcel delivery entities type.
 */
export class ArrayOfParcelDeliveryEntitiesType {
    ParcelDelivery: ParcelDeliveryEntityType[] = [];
}

/**
 * Flight entity type.
 */
export class FlightEntityType {
    FlightNumber?: string;
    AirlineIataCode?: string;
    DepartureTime?: string;
    WindowsTimeZoneName?: string;
    DepartureAirportIataCode?: string;
    ArrivalAirportIataCode?: string;
}

/**
 * Array of flights type.
 */
export class ArrayOfFlightsType {
    Flight: FlightEntityType[] = [];
}

/**
 * Flight reservation entity type.
 */
export class FlightReservationEntityType {
    ReservationId?: string;
    ReservationStatus?: string;
    UnderName?: string;
    BrokerName?: string;
    BrokerPhone?: string;
    Flights?: ArrayOfFlightsType;
}

/**
 * Array of flight reservations type.
 */
export class ArrayOfFlightReservationsType {
    FlightReservation: FlightReservationEntityType[] = [];
}

/**
 * Sender addin entity type.
 */
export class SenderAddInEntityType {
    ExtensionId?: string;
}

/**
 * Array of sender addins type.
 */
export class ArrayOfSenderAddInsType {
    ['Microsoft.OutlookServices.SenderApp']: SenderAddInEntityType[] = [];
}

/**
 * Entity extraction result type.
 */
export class EntityExtractionResultType {
    Addresses?: ArrayOfAddressEntitiesType;
    MeetingSuggestions?: ArrayOfMeetingSuggestionsType;
    TaskSuggestions?: ArrayOfTaskSuggestionsType;
    EmailAddresses?: ArrayOfEmailAddressEntitiesType;
    Contacts?: ArrayOfContactsType;
    Urls?: ArrayOfUrlEntitiesType;
    PhoneNumbers?: ArrayOfPhoneEntitiesType;
    ParcelDeliveries?: ArrayOfParcelDeliveryEntitiesType;
    FlightReservations?: ArrayOfFlightReservationsType;
    SenderAddIns?: ArrayOfSenderAddInsType;
}

/**
 * Flag type.
 */
export class FlagType {
    FlagStatus: FlagStatusType;
    StartDate?: Date;
    DueDate?: Date;
    CompleteDate?: Date;
}

/**
 * Internet header type.
 */
export class InternetHeaderType extends AttributeSimpleType<string> {
    constructor(_value: string, public HeaderName: string) {
        super(_value);
    }
}

/**
 * Array of internet headers type.
 */
export class NonEmptyArrayOfInternetHeadersType {
    InternetMessageHeader: InternetHeaderType[] = [];
}

/**
 * Online meeting settings type.
 */
export class OnlineMeetingSettingsType {
    LobbyBypass: LobbyBypassType;
    AccessLevel: OnlineMeetingAccessLevelType;
    Presenters: PresentersType;
}

/**
 * Calendar activity data type.
 */
export class CalendarActivityDataType {
    ActivityAction: string;
    ClientId: string;
    CasRequestId: string;
    IndexSelected: number;
}

/**
 * Enhanced location type.
 */
export class EnhancedLocationType {
    DisplayName: string;
    Annotation?: string;
    PostalAddress?: PersonaPostalAddressType;
}

/**
 * Time change type.
 */
export class TimeChangeType {
    Offset: string; // xs:duration
    RelativeYearlyRecurrence?: RelativeYearlyRecurrencePatternType;
    AbsoluteDate?: Date;
    AbsoluteDateString?: string;
    Time: Date;
    TimeString: string;
    TimeZoneName?: string;
}

/**
 * Time zone type.
 */
export class TimeZoneType {
    BaseOffset?: string; // xs:duration
    Standard?: TimeChangeType;
    Daylight?: TimeChangeType;
    TimeZoneName?: string;
}

/**
 * Period type.
 */
export class PeriodType {
    Bias?: string; // xs:duration
    Name?: string;
    Id?: string;
}

/**
 * Transition target type.
 */
export class TransitionTargetType extends AttributeSimpleType<string> {
    constructor(_value: string, public Kind: TransitionTargetKindType) {
        super(_value);
    }
}

/**
 * Transition type.
 */
export class TransitionType {
    To: TransitionTargetType;
}

/**
 * Absolute date transition type.
 */
export class AbsoluteDateTransitionType extends TransitionType {
    DateTime: Date;
}

/**
 * Recurring time transition type.
 */
export abstract class RecurringTimeTransitionType extends TransitionType {
    TimeOffset: string; // xs:duration
    Month: number;
}

/**
 * Recurring day transition type.
 */
export class RecurringDayTransitionType extends RecurringTimeTransitionType {
    DayOfWeek: DayOfWeekType;
    Occurrence: number;
}

/**
 * Recurring date transition type.
 */
export class RecurringDateTransitionType extends RecurringTimeTransitionType {
    Day: number;
}

/**
 * Array of periods type.
 */
export class NonEmptyArrayOfPeriodsType {
    Period: PeriodType[] = [];
}

/**
 * Array of transitions groups type.
 */
export class ArrayOfTransitionsGroupsType {
    TransitionsGroup: ArrayOfTransitionsType[] = [];
}

/**
 * Array of transitions type.
 */
export class ArrayOfTransitionsType extends GenericArray<TransitionType> {
    Id?: string;
}

/**
 * Time zone definition type.
 */
export class TimeZoneDefinitionType {
    Periods?: NonEmptyArrayOfPeriodsType;
    TransitionsGroups?: ArrayOfTransitionsGroupsType;
    Transitions?: ArrayOfTransitionsType;
    Id?: string;
    Name?: string;
}

/**
 * Deleted occurrence info type.
 */
export class DeletedOccurrenceInfoType {
    Start: Date;
}

/**
 * Array of deleted occurrences type.
 */
export class NonEmptyArrayOfDeletedOccurrencesType {
    DeletedOccurrence: DeletedOccurrenceInfoType[] = [];
}

/**
 * Occurrence info type.
 */
export class OccurrenceInfoType {
    ItemId: ItemIdType;
    Start: Date;
    End: Date;
    OriginalStart: Date;
}

/**
 * Array of occurrence info type.
 */
export class NonEmptyArrayOfOccurrenceInfoType {
    Occurrence: OccurrenceInfoType[] = [];
}

/**
 * Inbox reminder type.
 */
export class InboxReminderType {
    Id?: string;
    ReminderOffset?: number;
    Message?: string;
    IsOrganizerReminder?: boolean;
    OccurrenceChange?: EmailReminderChangeType;
    SendOption?: EmailReminderSendOption;
}

/**
 * Array of inbox reminder type.
 */
export class ArrayOfInboxReminderType {
    InboxReminder: InboxReminderType[] = [];
}

/**
 * Attendee type.
 */
export class AttendeeType {
    Mailbox: EmailAddressType;
    ResponseType?: ResponseTypeType;
    LastResponseTime?: Date;
    ProposedStart?: Date;
    ProposedEnd?: Date;
}

/**
 * Array of attendees type.
 */
export class NonEmptyArrayOfAttendeesType {
    Attendee: AttendeeType[] = [];
}

/**
 * Email address dictionary entry type.
 */
export class EmailAddressDictionaryEntryType extends AttributeSimpleType<string> {
    constructor(_value: string, public Key: EmailAddressKeyType, public Name?: string, public RoutingType?: string, public MailboxType?: MailboxTypeType) {
        super(_value);
    }
}

/**
 * Email address dictionary type.
 */
export class EmailAddressDictionaryType {
    Entry: EmailAddressDictionaryEntryType[] = [];
}

/**
 * Abch email address dictionary entry type.
 */
export class AbchEmailAddressDictionaryEntryType {
    Type: AbchEmailAddressTypeType;
    Address: string;
    IsMessengerEnabled?: boolean;
    Capabilities?: number;
}

/**
 * Abch email address dictionary type.
 */
export class AbchEmailAddressDictionaryType {
    Email: AbchEmailAddressDictionaryEntryType[] = [];
}

/**
 * Physical address dictionary entry type.
 */
export class PhysicalAddressDictionaryEntryType {
    Street?: string;
    City?: string;
    State?: string;
    CountryOrRegion?: string;
    PostalCode?: string;
    Key: PhysicalAddressKeyType;
}

/**
 * Physical address dictionary type.
 */
export class PhysicalAddressDictionaryType {
    Entry: PhysicalAddressDictionaryEntryType[] = [];
}

/**
 * Phone number dictionary entry type.
 */
export class PhoneNumberDictionaryEntryType extends AttributeSimpleType<string> {
    constructor(_value: string, public Key: PhoneNumberKeyType) {
        super(_value);
    }
}

/**
 * Phone number dictionary type.
 */
export class PhoneNumberDictionaryType {
    Entry: PhoneNumberDictionaryEntryType[] = [];
}

/**
 * Im address dictionary entry type.
 */
export class ImAddressDictionaryEntryType extends AttributeSimpleType<string> {
    constructor(_value: string, public Key: ImAddressKeyType) {
        super(_value);
    }
}

/**
 * Im address dictionary type.
 */
export class ImAddressDictionaryType {
    Entry: ImAddressDictionaryEntryType[] = [];
}

/**
 * Contact url dictionary entry type.
 */
export class ContactUrlDictionaryEntryType {
    Type: ContactUrlKeyType;
    Name?: string;
    Address?: string;
}

/**
 * Contact url dictionary type.
 */
export class ContactUrlDictionaryType {
    Url: ContactUrlDictionaryEntryType[] = [];
}

/**
 * Array of binary type.
 */
export class ArrayOfBinaryType {
    Base64Binary: string[] = []; // xs:base64Binary
}

/**
 * Complete name type.
 */
export class CompleteNameType {
    Title?: string;
    FirstName?: string;
    MiddleName?: string;
    LastName?: string;
    Suffix?: string;
    Initials?: string;
    FullName?: string;
    Nickname?: string;
    YomiFirstName?: string;
    YomiLastName?: string;
}

/**
 * Recurrence pattern base type.
 */
export abstract class RecurrencePatternBaseType { }

/**
 * Relative yearly recurrence pattern type.
 */
export class RelativeYearlyRecurrencePatternType extends RecurrencePatternBaseType {
    DaysOfWeek: DayOfWeekType;
    DayOfWeekIndex: DayOfWeekIndexType;
    Month: MonthNamesType;
}

/**
 * Absolute yearly recurrence pattern type.
 */
export class AbsoluteYearlyRecurrencePatternType extends RecurrencePatternBaseType {
    DayOfMonth: number;
    Month: MonthNamesType;
}

/**
 * Interval recurrence pattern type.
 */
export class IntervalRecurrencePatternBaseType extends RecurrencePatternBaseType {
    Interval: number;
}

/**
 * Relative monthly recurrence pattern type.
 */
export class RelativeMonthlyRecurrencePatternType extends IntervalRecurrencePatternBaseType {
    DaysOfWeek: DayOfWeekType;
    DayOfWeekIndex: DayOfWeekIndexType;
}

/**
 * Absolute monthly recurrence pattern type.
 */
export class AbsoluteMonthlyRecurrencePatternType extends IntervalRecurrencePatternBaseType {
    DayOfMonth: number;
}

/**
 * Weekly recurrence pattern type.
 */
export class WeeklyRecurrencePatternType extends IntervalRecurrencePatternBaseType {
    DaysOfWeek: DaysOfWeekType = new DaysOfWeekType();
    FirstDayOfWeek?: DayOfWeekType;
}

/**
 * Daily recurrence pattern type.
 */
export class DailyRecurrencePatternType extends IntervalRecurrencePatternBaseType { }

/**
 * Regenerating recurrence pattern type.
 */
export class RegeneratingPatternBaseType extends IntervalRecurrencePatternBaseType { }

/**
 * Daily regenerating recurrence pattern type.
 */
export class DailyRegeneratingPatternType extends RegeneratingPatternBaseType { }

/**
 * Weekly regenerating recurrence pattern type.
 */
export class WeeklyRegeneratingPatternType extends RegeneratingPatternBaseType { }

/**
 * Monthly regenerating recurrence pattern type.
 */
export class MonthlyRegeneratingPatternType extends RegeneratingPatternBaseType { }

/**
 * Yearly regenerating recurrence pattern type.
 */
export class YearlyRegeneratingPatternType extends RegeneratingPatternBaseType { }

/**
 * Recurrence range base type.
 */
export abstract class RecurrenceRangeBaseType {
    StartDate: Date;
}

/**
 * No end recurrence range type.
 */
export class NoEndRecurrenceRangeType extends RecurrenceRangeBaseType { }

/**
 * End date recurrence range type.
 */
export class EndDateRecurrenceRangeType extends RecurrenceRangeBaseType {
    EndDate: Date;
}

/**
 * Numbered recurrence range type.
 */
export class NumberedRecurrenceRangeType extends RecurrenceRangeBaseType {
    NumberOfOccurrences: number;
}

/**
 * Recurrence type.
 */
export class RecurrenceType {
    RelativeYearlyRecurrence?: RelativeYearlyRecurrencePatternType;
    AbsoluteYearlyRecurrence?: AbsoluteYearlyRecurrencePatternType;
    RelativeMonthlyRecurrence?: RelativeMonthlyRecurrencePatternType;
    AbsoluteMonthlyRecurrence?: AbsoluteMonthlyRecurrencePatternType;
    WeeklyRecurrence?: WeeklyRecurrencePatternType;
    DailyRecurrence?: DailyRecurrencePatternType;
    NoEndRecurrence?: NoEndRecurrenceRangeType;
    EndDateRecurrence?: EndDateRecurrenceRangeType;
    NumberedRecurrence?: NumberedRecurrenceRangeType;
}

/**
 * Task recurrence type.
 */
 export class TaskRecurrenceType extends RecurrenceType {
    DailyRegeneration?: DailyRegeneratingPatternType;
    WeeklyRegeneration?: WeeklyRegeneratingPatternType;
    MonthlyRegeneration?: MonthlyRegeneratingPatternType;
    YearlyRegeneration?: YearlyRegeneratingPatternType;
}

/**
 * Sharing message action type.
 */
export class SharingMessageActionType {
    Importance?: SharingActionImportance;
    ActionType?: SharingActionType;
    Action?: SharingAction;
}

/**
 * Array of sharing message action type.
 */
export class ArrayOfSharingMessageActionType {
    SharingMessageAction?: SharingMessageActionType;
}

/**
 * Smtp address type.
 */
export class SmtpAddressType extends AttributeSimpleType<string> {
    constructor(_value: string) {
        super(_value);
    }
}

/**
 * Message safety type.
 */
export class MessageSafetyType {
    MessageSafetyLevel?: number;
    MessageSafetyReason?: number;
    Info?: string;
}

/**
 * Reminder message data type.
 */
export class ReminderMessageDataType {
    ReminderText?: string;
    Location?: string;
    StartTime?: Date;
    EndTime?: Date;
    AssociatedCalendarItemId?: ItemIdType;
}

/**
 * Voting option data type.
 */
export class VotingOptionDataType {
    DisplayName?: string;
    SendPrompt?: SendPromptType;
}

/**
 * Array of voting option data type.
 */
export class ArrayOfVotingOptionDataType {
    VotingOptionData: VotingOptionDataType[] = [];
}

/**
 * Voting information type.
 */
export class VotingInformationType {
    UserOptions?: ArrayOfVotingOptionDataType;
    VotingResponse?: string;
}

/**
 * Approval request data type.
 */
export class ApprovalRequestDataType {
    IsUndecidedApprovalRequest?: boolean;
    ApprovalDecision?: number;
    ApprovalDecisionMaker?: string;
    ApprovalDecisionTime?: Date;
}

/**
 * Single recipient type.
 */
export class SingleRecipientType {
    Mailbox: EmailAddressType;
}

/**
 * Array of recipients type.
 */
export class ArrayOfRecipientsType {
    Mailbox: EmailAddressType[] = [];
}

/**
 * Array of distribution list type.
 */
export class ArrayOfDLExpansionType {
    Mailbox: EmailAddressType[] = [];
}

/**
 * Mailbox Guids.
 */
export class MailboxGuids {
    MailboxGuid: string[] = [];
}

/**
 * Identifies a single item to create/update in the local client store.
 */
export abstract class SyncFolderItemsCreateOrUpdateType {
    Item?: ItemType;
    Message?: MessageType;
    SharingMessage?: SharingMessageType;
    CalendarItem?: CalendarItemType;
    Contact?: ContactItemType;
    DistributionList?: DistributionListType;
    MeetingMessage?: MeetingMessageType;
    MeetingRequest?: MeetingRequestMessageType;
    MeetingResponse?: MeetingResponseMessageType;
    MeetingCancellation?: MeetingCancellationMessageType;
    Task?: TaskType;
    PostItem?: PostItemType;
    RoleMember?: RoleMemberItemType;
    Network?: NetworkItemType;
    Person?: AbchPersonItemType;
}

/**
 * Identifies a single item to create in the local client store.
 */
export class SyncFolderItemsCreateType extends SyncFolderItemsCreateOrUpdateType implements NamedElement {
    readonly $qname = 'Create';
}

/**
 * Identifies a single item to update in the local client store.
 */
export class SyncFolderItemsUpdateType extends SyncFolderItemsCreateOrUpdateType implements NamedElement {
    readonly $qname = 'Update';
}

/**
 * Identifies a single item to delete in the local client store.
 */
export class SyncFolderItemsDeleteType implements NamedElement {
    readonly $qname = 'Delete';
    ItemId: ItemIdType;
}

/**
 * Identifies when an item has been read.
 */
export class SyncFolderItemsReadFlagType implements NamedElement {
    readonly $qname = 'ReadFlagChange';
    ItemId: ItemIdType;
    IsRead: boolean;
}

/**
 * Contains a sequence array of change types that represent the types of differences
 * between the items on the client and the items on the Exchange server.
 */
export class SyncFolderItemsChangesType extends GenericArray<SyncFolderItemsCreateOrUpdateType | SyncFolderItemsDeleteType | SyncFolderItemsReadFlagType> { }

/**
 * Base subscription request type.
 */
export abstract class BaseSubscriptionRequestType {
    FolderIds?: NonEmptyArrayOfBaseFolderIdsType;
    EventTypes: NonEmptyArrayOfNotificationEventTypesType;
    Watermark?: string;
    SubscribeToAllFolders?: boolean;
}

/**
 * Pull subscription request type.
 */
export class PullSubscriptionRequestType extends BaseSubscriptionRequestType {
    Timeout: number;
}

/**
 * Push subscription request type.
 */
export class PushSubscriptionRequestType extends BaseSubscriptionRequestType {
    StatusFrequency: number;
    Url: string;
    CallerData?: string;
}

/**
 * Streaming subscription request type.
 */
export class StreamingSubscriptionRequestType {
    FolderIds?: NonEmptyArrayOfBaseFolderIdsType;
    EventTypes: NonEmptyArrayOfNotificationEventTypesType;
    SubscribeToAllFolders?: boolean;
}

/**
 * Array of notification event types.
 */
export class NonEmptyArrayOfNotificationEventTypesType {
    EventType: NotificationEventTypeType[] = [];
}

/**
 * Base paging type.
 */
export abstract class BasePagingType {
    MaxEntriesReturned?: number;
}

/**
 * Indexed page view type.
 */
export class IndexedPageViewType extends BasePagingType {
    Offset: number;
    BasePoint: IndexBasePointType;
}

/**
 * Fractional page view type.
 */
export class FractionalPageViewType extends BasePagingType {
    Numerator: number;
    Denominator: number;
}

/**
 * Calendar view type.
 */
export class CalendarViewType extends BasePagingType {
    StartDate: Date;
    EndDate: Date;
}

/**
 * Contacts view type.
 */
export class ContactsViewType extends BasePagingType {
    InitialName?: string;
    FinalName?: string;
}

/**
 * Seek to condition page view type.
 */
export class SeekToConditionPageViewType extends BasePagingType {
    Condition: RestrictionType;
    BasePoint: IndexBasePointType;
}

/**
 * Array of notifications type.
 */
export class NonEmptyArrayOfNotificationsType {
    Notification: NotificationType[] = [];
}

/**
 * Array of subscription ids type.
 */
export class NonEmptyArrayOfSubscriptionIdsType {
    SubscriptionId: string[] = [];
}

/**
 * Notification type.
 */
export class NotificationType extends GenericArray<BaseNotificationEventType> {
    SubscriptionId: string;
    PreviousWatermark?: string;
    MoreEvents?: boolean;
}

/**
 * Base notification event type.
 */
export abstract class BaseNotificationEventType {
    Watermark?: string;
}

/**
 * Base object changed event type.
 */
export abstract class BaseObjectChangedEventType extends BaseNotificationEventType {
    TimeStamp: Date;
    FolderId?: FolderIdType;
    ItemId?: ItemIdType;
    ParentFolderId: FolderIdType;
}

/**
 * Moved/copied event type.
 */
export abstract class MovedCopiedEventType extends BaseObjectChangedEventType {
    OldFolderId?: FolderIdType;
    OldItemId?: ItemIdType;
    OldParentFolderId: FolderIdType;
}

/**
 * modified event type.
 */
export abstract class ModifiedEventType extends BaseObjectChangedEventType {
    UnreadCount?: number;
}

/**
 * Status event.
 */
export class StatusEvent extends BaseNotificationEventType { }

/**
 * Copied event.
 */
export class CopiedEvent extends MovedCopiedEventType { }

/**
 * Moved event.
 */
export class MovedEvent extends MovedCopiedEventType { }

/**
 * Modified event.
 */
export class ModifiedEvent extends ModifiedEventType { }

/**
 * Created event.
 */
export class CreatedEvent extends BaseObjectChangedEventType { }

/**
 * Deleted event.
 */
export class DeletedEvent extends BaseObjectChangedEventType { }

/**
 * New mail event.
 */
export class NewMailEvent extends BaseObjectChangedEventType { }

/**
 * Free busy changed event.
 */
export class FreeBusyChangedEvent extends BaseObjectChangedEventType { }

/**
 * Highlight term
 */
export class HighlightTermType {
    Scope: string;
    Value: string;
}

/**
 * Array of highlight terms
 */
export class ArrayOfHighlightTermsType {
    Term?: HighlightTermType[] = [];
}

export abstract class BaseGroupByType {
    Order: SortDirectionType;
}

/**
 * Represents the field of each item to aggregate on and the qualifier to apply to that
 *   field in determining which item will represent the group.
 */
export class AggregateOnType extends PathChoiceType {
    Aggregate: AggregateType;
}

/**
 * Represents an array of GroupIds.
 */
export class ArrayOfGroupIdType {
    GroupId?: string; //xs:bae64Binary
}

/**
 * Allows consumers to specify arbitrary groupings for FindItem queries.
 */
export class GroupByType extends BaseGroupByType {
    FieldURI?: PathToUnindexedFieldType;
    IndexedFieldURI?: PathToIndexedFieldType;
    ExtendedFieldURI?: PathToExtendedFieldType;
    AggregateOn: AggregateOnType;
    UseCollapsibleGroups?: boolean;
    ItemsPerGroup?: number;
    MaxItemsPerGroup?: number;
    GroupsToExpand?: ArrayOfGroupIdType;
}

/**
 * Allows consumers to access standard groupings for FindItem queries.  This is in
 *   contrast to the arbitrary (custom) groupings available via the GroupByType
 */
export class DistinguishedGroupByType extends BaseGroupByType {
    StandardGroupBy: StandardGroupByType;
}

 /**
 * Defines a request to find folder from a mailbox in the Exchange store.
 */
export class FindItemType extends BaseRequestType {
    ItemShape: ItemResponseShapeType;
    IndexedPageItemView?: IndexedPageViewType;
    FractionalPageItemView?: FractionalPageViewType;
    SeekToConditionPageItemView?: SeekToConditionPageViewType;
    CalendarView?: CalendarViewType;
    ContactsView?: ContactsViewType;
    GroupBy?: GroupByType;
    DistinguishedGroupBy?: DistinguishedGroupByType;
    Restriction?: RestrictionType;
    SortOrder?: NonEmptyArrayOfFieldOrdersType;
    ParentFolderIds: NonEmptyArrayOfBaseFolderIdsType;
    QueryString?: string;
    Traversal: ItemQueryTraversalType;
}

/**
 * Find item parent type.
 */
export class FindItemParentType {
    Items?: ArrayOfRealItemsType;
    Groups?: ArrayOfGroupedItemsType;
    IndexedPagingOffset?: number;
    NumeratorOffset?: number;
    AbsoluteDenominator?: number;
    IncludesLastItemInRange?: boolean;
    TotalItemsInView?: number;
}

/**
 * Contains the status and result of a single FindItem operation request.
 */
export class FindItemResponseMessageType extends ResponseMessageType {
    RootFolder?: FindItemParentType;
    HighlightTerms?: ArrayOfHighlightTermsType;
}

/**
 * Defines a response to a FindItem request.
 */
export class FindItemResponseType extends BaseResponseMessageType<FindItemResponseMessageType>{ }

/**
 * Defines a conversation shape.
 */
export class ConversationResponseShapeType {
    BaseShape: DefaultShapeNamesType;
    AdditionalProperties?: NonEmptyArrayOfPathsToElementType;
}

/**
 * Defines a request to a find a conversation in a mailbox.
 */
export class FindConversationType extends BaseRequestType {
    IndexedPageItemView?: IndexedPageViewType;
    SeekToConditionPageItemView?: SeekToConditionPageViewType;
    SortOrder?: NonEmptyArrayOfFieldOrdersType;
    ParentFolderId: TargetFolderIdType;
    MailboxScope?: MailboxSearchLocationType;
    QueryString?: string;
    ConversationShape?: ConversationResponseShapeType;
    Traversal: ConversationQueryTraversalType;
    ViewFilter: ViewFilterType;

}

export class ArrayOfItemClassType {
    ItemClass?: string[] = [];
}

export class ConversationType {
    ConversaitonId?: ItemIdType;
    ConversationTopic?: string;
    UniqueRecipients?: ArrayOfStringsType;
    GlobalUniqueRecipients?: ArrayOfStringsType;
    UniqueUnreadSenders?: ArrayOfStringsType;
    GlobalUniqueUnreadSenders?: ArrayOfStringsType;
    UniqueSenders?: ArrayOfStringsType;
    GlobalUniqueSenders?: ArrayOfStringsType;
    LastDeliveryTime?: Date;
    GlobalLastDeliveryTime?: Date;
    Categories?: ArrayOfStringsType;
    GlobalCategories?: ArrayOfStringsType;
    FlagStatus?: FlagStatusType;
    GlobalFlagStatus?: FlagStatusType;
    HasAttachments?: boolean;
    GlobalHasAttachments?: boolean;
    MessageCount?: number;
    GlobalMessageCount?: number;
    UnreadCount?: number;
    GlobalUnreadCount?: number;
    Size?: number;
    GlobalSize?: number;
    ItemClasses?: ArrayOfItemClassType;
    GlobalItemClasses?: ArrayOfItemClassType;
    Importance?: ImportanceChoicesType
    GlobalImportance?: ImportanceChoicesType
    ItemIds?: NonEmptyArrayOfBaseItemIdsType;
    GlobalItemIds?: NonEmptyArrayOfBaseItemIdsType;
    LastModifiedTime?: Date;
    InstanceKey?: string;  //base64Binary
    Preview?: string;
    MailboxScope?: MailboxSearchLocationType;
    IconIndex?: IconIndexType;
    GlobalIconIndex?: IconIndexType;
    DraftItemIds?: NonEmptyArrayOfBaseItemIdsType;
    HasIrm?: boolean;
    GlobalHasIrm?: boolean;
    InferenceClassification?: InferenceClassificationType;
    SortKey?: number;
    MentionedMe?: boolean;
    GlobalMentionedMe?: boolean;
    SenderSMTPAddress?: SmtpAddressType
    MailboxGuids?: MailboxGuids;
    From?: SingleRecipientType
    AtAllMention?: boolean;
    GlobalAtAllMention?: boolean;
}

export class ArrayOfConversationsType {
    Conversaion?: ConversationType[] = [];
}

/**
 * Returns the conversations and requested highlighted terms
 */
export class FindConversationResponseMessageType extends ResponseMessageType{
    Conversations?: ArrayOfConversationsType;
    HIghlightTerms?: ArrayOfHighlightTermsType;
    TotalConversationsInView?: number;
    IndexedOffset?: number;
}


export class ConversationRequestsType {
    ConversationId: ItemIdType;
    SyncState?: string //xs:base64Binary
}

export class ArrayOfConversationRequestsType {
    Conversation?: ConversationRequestsType[] = [];
}

/**
 * Defines a conversation from which to retrieve the conversation items
 */
export class GetConversationItemsType extends BaseRequestType {
    ItemShape: ItemResponseShapeType;
    FoldersToIgnore?: NonEmptyArrayOfBaseFolderIdsType;
    MaxItemsToReturn?: number;
    SortOrder?: ConversationNodeSortOrder;
    MailboxScope?: MailboxSearchLocationType;
    Conversations: ArrayOfConversationRequestsType;
} 

export class ConversationNodeType {
    InternetMessageId?: string;
    ParentInternetMessageId?: string;
    Items?: NonEmptyArrayOfAllItemsType;
}

export class ArrayOfConversationNodesType {
    ConversationNode?: ConversationNodeType[] = [];
}

export class ConversationResponseType {
    ConversaionId: ItemIdType;
    SyncState?: string;  // xs:base64Binary
    ConversationNodes?: ArrayOfConversationNodesType;
    CanDelete?: boolean;
}

/**
 * Defines resulting conversation items response
 */
export class GetConversationItemsResponseMessageType extends ResponseMessageType {
    Conversation?: ConversationResponseType;
}

/**
 * Result response of conversation items 
 */
export class GetConversationItemsResponseType extends BaseResponseMessageType<GetConversationItemsResponseMessageType> { }

export class ConversationActionType {
    Action: ConversationActionTypeType;
    ConversationId: ItemIdType;
    ContextFolderId?: TargetFolderIdType;
    ConversationLastSyncTime?: Date;
    ProcessRightAway?: boolean;
    DestinationFolderId?: TargetFolderIdType;
    Categories?: ArrayOfStringsType;
    EnableAlwaysDelete?: boolean;
    IsRead?: boolean;
    DeleteType?: DisposalType;
    RetentionPolicyType?: RetentionType;
    RetentionPolicyTagId?: string;
    Flag?: FlagType;
    SuppressReadReceipts?: boolean;
}

export class NonEmptyArrayOfApplyConversationActionType {
    ConversationAction: ConversationActionType[] = [];
}

/**
 * Defines conversation actions to apply.
 */
export class ApplyConversationActionType extends BaseRequestType {
    ConversationActions: NonEmptyArrayOfApplyConversationActionType;
}

/**
 * Message response for each applied action .
 */
export class ApplyConversationActionResponseMessageType extends ResponseMessageType { }

/**
 * Defines a result from applying the actions to the conversation.
 */
export class ApplyConversationActionResponseType extends BaseResponseMessageType<ApplyConversationActionResponseMessageType> { }

/**
 * Defines email identifying the distribution list to expand
 */
export class ExpandDLType extends BaseRequestType {
    Mailbox: EmailAddressType;
}

/**
 * Resulting distribution list expansion
 */
export class ExpandDLResponseMessageType extends ResponseMessageType { 
    DLExpansion?: ArrayOfDLExpansionType;
    IndexedPagingOffset?: number;
    NumeratorOffset?: number;
    AbsoluteDenominator?: number;
    IncludesLastItemInRange?: boolean;
    TotalItemsInView?: number;
}

/**
 * Response to request to expand distribution list
 */
export class ExpandDLResponseType extends BaseResponseMessageType<ExpandDLResponseMessageType> { }

/**
 * Defines the request for the mailbox to retrieve the password expiration
 */
export class GetPasswordExpirationDateType extends BaseRequestType {
    MailboxSmtpAddress?: string;
}

/**
 * Resulting expiration date for the password
 */
export class GetPasswordExpirationDateResponseMessageType extends ResponseMessageType { 
    PasswordEExpirationDate: Date;
}

/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import EventEmitter from 'events';

import fastcopy from 'fast-copy';
import { v4 as uuidv4 } from 'uuid';
import { Request } from '@loopback/rest';
import { UserContext } from '../keepcomponent';

import {
    KeepPimMessageManager, KeepPimConstants, KeepPimLabelManager, PimLabel, KeepPimAttachmentManager, UserInfo,
    KeepPimManager, PimItem, base64Encode, PimLabelTypes, PimMessage
} from '@hcllabs/openclientkeepcomponent';

import {
    FolderType, CreateFolderPathResponseType, CreateFolderPathResponseMessage, UploadItemsResponseType, FolderIdType,
    UploadItemsResponseMessageType, CalendarFolderType, ContactsFolderType, TasksFolderType, BaseFolderType, ItemIdType,
    SyncFolderHierarchyCreateType, SyncFolderItemsCreateType, SyncFolderHierarchyDeleteType, SyncFolderHierarchyUpdateType,
    SyncFolderItemsUpdateType, SyncFolderItemsDeleteType, SyncFolderItemsReadFlagType, ItemType, GetFolderType, MessageType,
    DistinguishedFolderIdType, GetFolderResponseMessage, ArrayOfFoldersType, SyncFolderHierarchyResponseType, TaskType,
    SyncFolderHierarchyResponseMessageType, SyncFolderHierarchyType, SearchFolderType, SyncFolderHierarchyCreateOrUpdateType,
    CreateFolderType, CreateFolderResponseType, CreateFolderResponseMessage, DeleteFolderType, DeleteFolderResponseType,
    DeleteFolderResponseMessage, UpdateFolderType, UpdateFolderResponseType, SetFolderFieldType, UpdateFolderResponseMessage,
    SyncFolderItemsResponseType, SyncFolderItemsResponseMessageType, SyncFolderItemsCreateOrUpdateType, CalendarItemType,
    MeetingCancellationMessageType, MeetingResponseMessageType, MeetingRequestMessageType, MeetingMessageType, SharingMessageType,
    PostItemType, NetworkItemType, AbchPersonItemType, RoleMemberItemType, SyncFolderItemsType, GetItemType, GetItemResponseType,
    GetItemResponseMessage, CreateItemType, CreateItemResponseType, CreateItemResponseMessage, GetFolderResponseType, ContactItemType,
    UpdateItemType, UpdateItemResponseType, UpdateItemResponseMessageType,
    DeleteItemType, DeleteItemResponseType, DeleteItemResponseMessageType, MoveItemType, MoveItemResponseType,
    MoveItemResponseMessage, CopyItemType, CopyItemResponseType, CopyItemResponseMessage, ArchiveItemResponseType,
    ArchiveItemResponseMessage, ArchiveItemType, MarkAsJunkType, MarkAsJunkResponseType, MarkAsJunkResponseMessageType, SubscribeType,
    MarkAllItemsAsReadResponseType, MarkAllItemsAsReadType, MarkAllItemsAsReadResponseMessage, ReportMessageResponseType, ReportMessageType,
    ReportMessageResponseMessageType, NonEmptyArrayOfAttachmentsType, FileAttachmentType, AttachmentIdType, GetAttachmentType,
    GetAttachmentResponseType, GetAttachmentResponseMessage, DeleteAttachmentType, DeleteAttachmentResponseType, CreateAttachmentType,
    DeleteAttachmentResponseMessageType, RootItemIdType, CreateAttachmentResponseType, CreateAttachmentResponseMessage, ItemAttachmentType,
    GetStreamingEventsType, GetStreamingEventsResponseType, GetStreamingEventsResponseMessageType, StreamingSubscriptionRequestType,
    SubscribeResponseType, SubscribeResponseMessageType, NonEmptyArrayOfNotificationsType, BaseNotificationEventType, NotificationType,
    ModifiedEvent, DeletedEvent, CreatedEvent, FindFolderType, FindFolderResponseType, FindFolderResponseMessageType, FindFolderParentType,
    MovedEvent, UploadItemsType, SendItemType, SendItemResponseType, SendItemResponseMessage, CreateFolderPathType, ConvertIdType,
    ConvertIdResponseType, ConvertIdResponseMessageType, AlternateIdType, MoveFolderType, MoveFolderResponseType, MoveFolderResponseMessage,
    CopyFolderType, CopyFolderResponseType, CopyFolderResponseMessage,
    GetAppMarketplaceUrlType, GetAppMarketplaceUrlResponseMessageType, FindItemType, FindItemResponseType, FindItemResponseMessageType,
    FindItemParentType, ArrayOfHighlightTermsType, ArrayOfRealItemsType, CreateManagedFolderRequestType, CreateManagedFolderResponseType,
    CreateManagedFolderResponseMessage, EmptyFolderType, EmptyFolderResponseType, EmptyFolderResponseMessage,
    FindConversationType, FindConversationResponseMessageType,
    GetConversationItemsType, GetConversationItemsResponseType, GetConversationItemsResponseMessageType, ApplyConversationActionType,
    ApplyConversationActionResponseType, ApplyConversationActionResponseMessageType, ExpandDLType, ExpandDLResponseType,
    ExpandDLResponseMessageType, GetPasswordExpirationDateType, GetPasswordExpirationDateResponseMessageType,
    MimeContentType, ItemResponseShapeType, UploadItemType, AttachmentType, ExtendedPropertyType
} from '../models/mail.model';
import {
    DistinguishedFolderIdNameType, ResponseClassType, ResponseCodeType, ConnectionStatusType, DefaultShapeNamesType,
    MessageDispositionType, CalendarItemCreateOrDeleteOperationType, UnindexedFieldURIType, ContainmentModeType, 
    ContainmentComparisonType, FolderQueryTraversalType, CreateActionType, DisposalType
} from '../models/enum.model';
import { Value } from '../models/common.model';
import util from 'util';
import {
    Logger, throwErrorIfClientResponse, findLabel, getDistinguishedFolderId, findLabelsWithParent, findLabelByTarget,
    isRootFolderId, isTargetRootFolderId, rootFolderIdForRequest, getAttachmentId, getAttachmentIdPair, pimLabelFromItem,
    EWSServiceManager, getUserId, getUserMailData, populateItemFromMime
} from '../utils';
import { RestrictionType, ContainsExpressionType } from '../models/persona.model';
import { SyncStateInfo } from './SyncStateInfo';
import { 
    labelToFolder, includeUnReadCountForFolders, rootFolderForUser, labelTypeFromFolderClass, getLabelByTarget, 
    getKeepIdPair, combineFolderIdAndMailbox, findMailboxId, getFolderLabelsFromHash, getKeepLabels, getEWSId
} from '../utils/pimHelper';

const config = require('../../config');

/**
 * Populate folder change object.
 * @param change The folder change object
 * @param folder The changed folder
 * @returns given folder change object
 */
function populateSyncFolderChange(change: SyncFolderHierarchyCreateOrUpdateType, folder: BaseFolderType): SyncFolderHierarchyCreateOrUpdateType {
    if (folder instanceof SearchFolderType) {
        change.SearchFolder = folder;
    } else if (folder instanceof TasksFolderType) {
        change.TasksFolder = folder;
    } else if (folder instanceof CalendarFolderType) {
        change.CalendarFolder = folder;
    } else if (folder instanceof ContactsFolderType) {
        change.ContactsFolder = folder;
    } else if (folder instanceof FolderType) {
        // Both Notes and Messages use FolderType
        change.Folder = folder;
    } else {
        Logger.getInstance().debug("Unexpected FolderType");
    }
    return change;
}

/**
 * Populate item change object.
 * @param change The item change object
 * @param item The changed item
 * @returns given item change object
 */
function populateSyncItemChange(change: SyncFolderItemsCreateOrUpdateType, item: ItemType): SyncFolderItemsCreateOrUpdateType {
    if (item instanceof CalendarItemType) {
        change.CalendarItem = item;
    } else if (item instanceof TaskType) {
        change.Task = item;
    } else if (item instanceof ContactItemType) {
        change.Contact = item;
    } else if (item instanceof MeetingCancellationMessageType) {
        change.MeetingCancellation = item;
    } else if (item instanceof MeetingResponseMessageType) {
        change.MeetingResponse = item;
    } else if (item instanceof MeetingRequestMessageType) {
        change.MeetingRequest = item;
    } else if (item instanceof MeetingMessageType) {
        change.MeetingMessage = item;
    } else if (item instanceof SharingMessageType) {
        change.SharingMessage = item;
    } else if (item instanceof MessageType) {
        change.Message = item;
    } else if (item instanceof PostItemType) {
        change.PostItem = item;
    } else if (item instanceof NetworkItemType) {
        change.Network = item;
    } else if (item instanceof AbchPersonItemType) {
        change.Person = item;
    } else if (item instanceof RoleMemberItemType) {
        change.RoleMember = item;
    } else if (item instanceof ItemType) {
        change.Item = item;
    }
    return change;
}

/**
 * Folder event.
 */
export class FolderEvent {
    constructor(
        public timestamp: number,
        public change: SyncFolderHierarchyCreateType | SyncFolderHierarchyUpdateType | SyncFolderHierarchyDeleteType) { }
}
/**
 * Item event.
 */
export class ItemEvent {
    constructor(
        public timestamp: number,
        public change: SyncFolderItemsCreateType | SyncFolderItemsUpdateType | SyncFolderItemsDeleteType | SyncFolderItemsReadFlagType) { }
}

/**
 * User mailbox data.
 */
export class UserMailboxData {

    // Folder events.
    readonly FOLDER_EVENTS: { [folderId: string]: FolderEvent } = {};

    // Folder item events.
    readonly FOLDER_ITEM_EVENTS: { [folderId: string]: { [itemId: string]: ItemEvent } } = {};

    // Items sent on the last SyncFolderItemsRequest
    public LAST_SYNCFOLDER_ITEMS: { [folderId: string]: ItemType[] } = {};

    // The folders on the last SyncFolderHierarchy 
    public LAST_SYNCFOLDERS: PimLabel[] = [];

    public PROBLEM_ITEMS: string[] = [];
    public PROBLEM_ITEM_NEXT_SYNC_FLUSH: number = Date.now();

    /**
     * Create a mailbox for a user to hold cached data. 
     * @param userId The owning user id.
     */
    constructor(public userId: string) { }

}

export const USER_MAILBOX_DATA: { [email: string]: UserMailboxData } = {};

async function getInitializedUserMailData(request: Request): Promise<UserMailboxData> {
    try {
        const userId = getUserId(request);
        if (!userId) {
            throw new Error('User identity not found');
        }

        return getUserMailData(request);

    } catch (err) {
        Logger.getInstance().debug(" error executing getLabels=" + err);
        throwErrorIfClientResponse(err);
        return getUserMailData(request) // Use mock data
    }
}

/**
 * Create a new instance with type of given object.
 * @param obj Object
 * @returns new instance
 */
function newInstance<T extends object>(obj: T): T {
    const classType = obj.constructor as (new () => T);
    return new classType();
}

/**
 * Shallow copy given object.
 * @param obj Object
 * @returns shallow copy
 */
function shallowCopy<T extends object>(obj: T): T {
    const copied = newInstance(obj);
    Object.assign(copied, obj);
    return copied;
}

/**
 * Shallow copy given object, and omit given field.
 * @param obj Object
 * @returns shallow copy of given object with given field omitted
 */
function omit<T extends object>(obj: T, field: string): T {
    const copied = shallowCopy(obj);
    delete (copied as any)[field];
    return copied;
}

/**
 * Deep copy given object.
 * @param obj Object
 * @returns deep copy
 */
function deepcoy<T>(obj: T): T {
    return fastcopy(obj, { isStrict: true });
}

/**
 * Get folder.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns folder retrieved
 */
export async function GetFolder(soapRequest: GetFolderType, request: Request): Promise<GetFolderResponseType> {
    const response = new GetFolderResponseType();

    const userInfo = UserContext.getUserInfo(request);
    if (
        soapRequest.FolderIds.items.length === 1 && isRootFolderId(userInfo, soapRequest.FolderIds.items[0], request) 
        && soapRequest.FolderShape.BaseShape === DefaultShapeNamesType.ID_ONLY
    ) {
        // In the case where the request is just for the root folder and we don't have to return the count of child folders, avoid making a Keep call
        const mailboxId = findMailboxId(userInfo, soapRequest.FolderIds.items[0]);
        const root = rootFolderForUser(userInfo, soapRequest.FolderShape, [], mailboxId, request);
        const msg = new GetFolderResponseMessage();
        msg.Folders = new ArrayOfFoldersType();
        msg.Folders.push(root); 
        response.ResponseMessages.push(msg);
        return response;
    }


    if (soapRequest.FolderIds.items.length === 0) {
        console.error("Missing ids");
        const msg = new GetFolderResponseMessage();
        msg.ResponseClass = ResponseClassType.ERROR;
        msg.ResponseCode = ResponseCodeType.ERROR_MISSING_ARGUMENT;
        response.ResponseMessages.push(msg);
        return response;
    }

    const firstFolderMailboxId = findMailboxId(userInfo, soapRequest.FolderIds.items[0]);
    const getFolderLabels = await getFolderLabelsFromHash(
        userInfo,  
        firstFolderMailboxId,
        includeUnReadCountForFolders(soapRequest.FolderShape)
        ).catch(err => {
            Logger.getInstance().error(`GetFolder is unable to get PIM labels: ${err}`);
            throw err;
        });

    for (const item of soapRequest.FolderIds.items) {
        const itemMailboxId = findMailboxId(userInfo, item);
        const labels = await getFolderLabels(itemMailboxId)
            .catch(err => {
                Logger.getInstance().error(`GetFolder is unable to get PIM labels: ${err}`);
                throw err;
            });

        // Since Domino does not have a root label, first check if this is the root folder and always return it. 
        if (isRootFolderId(userInfo, item, request)) {
            const root = rootFolderForUser(userInfo, soapRequest.FolderShape, labels, itemMailboxId, request);
            const msg = new GetFolderResponseMessage();
            msg.Folders = new ArrayOfFoldersType();
            msg.Folders.push(root);
            response.ResponseMessages.push(msg);
            continue;
        }

        const found = findLabel(labels, item);
        if (found) {
            //Get folder with combined EWS id
            const folder = labelToFolder(userInfo, found, request, item, soapRequest.FolderShape, labels);
            const msg = new GetFolderResponseMessage();
            msg.Folders = new ArrayOfFoldersType();
            msg.Folders.push(folder);
            response.ResponseMessages.push(msg);
        } else {
            const msg = new GetFolderResponseMessage();
            msg.ResponseClass = ResponseClassType.ERROR;

            // Include the folder id in the response to assist with debugging
            let folder: string;
            if (item instanceof FolderIdType) {
                folder = item.Id;
            } else if (item instanceof DistinguishedFolderIdType) {
                folder = item.Id;
            }
            else {
                folder = "unknown";
            }

            const isPublicRoot = item instanceof DistinguishedFolderIdType
                && item.Id === DistinguishedFolderIdNameType.PUBLICFOLDERSROOT;
            msg.ResponseCode = isPublicRoot ? ResponseCodeType.ERROR_NO_PUBLIC_FOLDER_REPLICA_AVAILABLE
                : ResponseCodeType.ERROR_FOLDER_NOT_FOUND;
            msg.MessageText = isPublicRoot ? 'There are no public folder servers available.'
                : 'The specified folder could not be found in the store';
            msg.DescriptiveLinkKey = 0;
            msg.MessageXml = { Value: new Value('FolderId', folder) };
            msg.Folders = new ArrayOfFoldersType();
            response.ResponseMessages.push(msg);
        }
    }
    return response;
}

/**
 * Find folder.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns folder found
 */
export async function FindFolder(soapRequest: FindFolderType, request: Request): Promise<FindFolderResponseType> {

    await getInitializedUserMailData(request);
    const userInfo = UserContext.getUserInfo(request);

    const response = new FindFolderResponseType();

    const firstFolderMailboxId = findMailboxId(userInfo, soapRequest.ParentFolderIds.items[0]);
    const getFolderLabels = await getFolderLabelsFromHash(
        userInfo, 
        firstFolderMailboxId,
        includeUnReadCountForFolders(soapRequest.FolderShape)
        ).catch(err => {
            Logger.getInstance().error(`FindFolder is unable to get PIM labels: ${err}`);
            throw err;
        });

    // Generate a response for each parent folder requested
    for (const folderId of soapRequest.ParentFolderIds.items) {
        const folderMailboxId = findMailboxId(userInfo, folderId);
        const labels = await getFolderLabels(folderMailboxId)
            .catch(err => {
                Logger.getInstance().error(`FindFolder is unable to get PIM labels: ${err}`);
                throw err;
            });               

        const msg = new FindFolderResponseMessageType();
        response.ResponseMessages.push(msg);
        msg.RootFolder = new FindFolderParentType();
        msg.RootFolder.IncludesLastItemInRange = true;

        const deepSearch = (soapRequest.Traversal === FolderQueryTraversalType.DEEP)
        const children = findLabelsWithParent(userInfo, labels, folderId, deepSearch, request);
        if (children.length > 0) {
            msg.RootFolder.Folders = new ArrayOfFoldersType();

            for (const label of children) {
                //Get folder with combined EWS id
                const found = labelToFolder(userInfo, label, request, folderId, soapRequest.FolderShape, labels);

                let include = true;
                if (soapRequest.Restriction) {
                    include = matchesRestrictions(found, soapRequest.Restriction);
                }

                if (include) {
                    msg.RootFolder.Folders.push(found);
                }
            }

            msg.RootFolder.IndexedPagingOffset = msg.RootFolder.Folders.items.length;
            msg.RootFolder.TotalItemsInView = msg.RootFolder.Folders.items.length;
        }
        else {
            if (findLabel(labels, folderId)) {
                // Parent exists but has no child folders
                msg.RootFolder.IndexedPagingOffset = 0;
                msg.RootFolder.TotalItemsInView = 0;
            }
            else {
                // Parent does not exist
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_FOLDER_NOT_FOUND;
            }
        }
    }


    return response;
}

/**
 * Test if a folder matches restrictions set on a request. 
 * @param folder The folder to test
 * @param restrictions The restrictions from the request. 
 * @returns True if the folder matches the restrictions, otherwise fale. 
 */
function matchesRestrictions(folder: FolderType, restrictions: RestrictionType): boolean {

    if (restrictions.Or) {
        for (const expression of restrictions.Or.items) {
            let value: string | undefined = undefined;
            if (expression instanceof ContainsExpressionType) {
                if (expression.FieldURI?.FieldURI === UnindexedFieldURIType.FOLDER_FOLDER_CLASS && folder.FolderClass) {
                    value = folder.FolderClass
                }
                else {
                    return true; // TODO: Implement others. Will ignore the restriction so return a match.
                }

                if (value) {
                    if (expression.ContainmentMode === ContainmentModeType.FULL_STRING && expression.ContainmentComparison === ContainmentComparisonType.EXACT) {
                        if (value === expression.Constant.Value) {
                            return true; // Found a match for OR expression, no need to continue
                        }
                    }
                    else if (expression.ContainmentMode === ContainmentModeType.PREFIXED && expression.ContainmentComparison === ContainmentComparisonType.EXACT) {
                        if (value.startsWith(expression.Constant.Value)) {
                            return true; // Found a match for OR expression, no need to continue
                        }
                    }
                }
            }
            else {
                return true; // TODO: Implement others. Will ignore the restriction so return a match.
            }
        }

        return false; // No match found
    }

    return true; // TODO: Implement others. Will ignore the restriction so return a match.
}

/**
 * Find Item.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns items found
 */
export function FindItem(soapRequest: FindItemType, request: Request): FindItemResponseType {
    // ##EWSTODO Fill in details to find the items for the response rather than an empty response
    // const userData = getUserMailData(request);
    Logger.getInstance().debug('FindItem is not completely implemented yet');
    const response = new FindItemResponseType();
    const msg = new FindItemResponseMessageType();
    response.ResponseMessages.push(msg);
    msg.RootFolder = new FindItemParentType();
    msg.RootFolder.Items = new ArrayOfRealItemsType();
    msg.HighlightTerms = new ArrayOfHighlightTermsType();
    msg.RootFolder.IncludesLastItemInRange = true;
    msg.RootFolder.IndexedPagingOffset = msg.RootFolder.Items.items.length;
    msg.RootFolder.TotalItemsInView = msg.RootFolder.Items.items.length;
    return response;
}



/**
 * Find conversation
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns The conversation
 */
export function FindConversation(soapRequest: FindConversationType, request: Request): FindConversationResponseMessageType {
    // ##EWSTODO Fill in details to find the items for the response rather than an empty response
    // const userData = getUserMailData(request);
    Logger.getInstance().debug('FindConversation is not completely implemented yet');
    return new FindConversationResponseMessageType();
}

/**
 * Get conversation items
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns The conversation items
 */
export function GetConversationItems(soapRequest: GetConversationItemsType, request: Request): GetConversationItemsResponseType {
    // ##EWSTODO Fill in details to find the items for the response rather than an empty response
    // const userData = getUserMailData(request);
    Logger.getInstance().debug('GetConversationItems is not completely implemented yet');
    const response = new GetConversationItemsResponseType();
    const msg = new GetConversationItemsResponseMessageType();
    response.ResponseMessages.push(msg);
    return response;
}

/**
 * Apply a conversation action
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns applied converation responses
 */
export function ApplyConversationAction(soapRequest: ApplyConversationActionType, request: Request): ApplyConversationActionResponseType {
    // ##EWSTODO Fill in details to find the items for the response rather than an empty response
    // const userData = getUserMailData(request);
    Logger.getInstance().debug('ApplyConversationAction is not completely implemented yet');
    const response = new ApplyConversationActionResponseType();
    const msg = new ApplyConversationActionResponseMessageType();
    response.ResponseMessages.push(msg);
    return response;
}

/**
 * Expand the distribution list
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns Array of distribution list expansions
 */
export function ExpandDL(soapRequest: ExpandDLType, request: Request): ExpandDLResponseType {
    // ##EWSTODO Fill in details to find the items for the response rather than an empty response
    // const userData = getUserMailData(request);
    Logger.getInstance().debug('ExpandDL is not completely implemented yet');
    const response = new ExpandDLResponseType();
    const msg = new ExpandDLResponseMessageType();
    response.ResponseMessages.push(msg);
    return response;
}

/**
 * Get the password expiration date
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns password expiration date
 */
export function GetPasswordExpirationDate(soapRequest: GetPasswordExpirationDateType, request: Request): GetPasswordExpirationDateResponseMessageType {
    // ##EWSTODO Fill in details to find the items for the response rather than an empty response
    // const userData = getUserMailData(request);
    Logger.getInstance().debug('GetPasswordExpirationDate is not completely implemented yet');
    return new GetPasswordExpirationDateResponseMessageType();
}

/**
 * Create Managed Folder.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns items found
 */

export function CreateManagedFolder(soapRequest: CreateManagedFolderRequestType, request: Request): CreateManagedFolderResponseType {
    // ##EWSTODO Fill in details to add the managed folder for the response 

    // Exchange methodology....
    // Use FindFolder to see if the folder exists.
    // User GetFolder to get the folder information to determine if it is managed.  If not, fail.
    // If the folder is found and it is managed, add the folder to the mailbox if provided, otherwise add
    // to the current user's mailbox

    // Does the concept of managed folders exist in domino????
    // What authority do I need to add a managed folder to a mailbox?

    const response = new CreateManagedFolderResponseType();

    soapRequest.FolderNames.FolderName.forEach(folderName => {
        const msg = new CreateManagedFolderResponseMessage();
        response.ResponseMessages.push(msg);
        msg.ResponseClass = ResponseClassType.ERROR;
        msg.ResponseCode = ResponseCodeType.ERROR_MANAGED_FOLDERS_ROOT_FAILURE;
        msg.MessageText = "Managed folders are not supported.";
    });

    return response;
}

/**
 * Empty the contents of a folder. The folder will only contatins MessageType items and other folders. 
 * @param soapRequest The EmptyFolder request. 
 * @param request The HTTP request being processed.
 */
export async function EmptyFolder(soapRequest: EmptyFolderType, request: Request): Promise<EmptyFolderResponseType> {
    const hardDelete = soapRequest.DeleteType === DisposalType.HARD_DELETE;

    const response = new EmptyFolderResponseType();

    try {
        const userInfo = UserContext.getUserInfo(request);

        const firstFolderMailboxId = findMailboxId(userInfo, soapRequest.FolderIds.items[0]);
        const getFolderLabels = await getFolderLabelsFromHash(userInfo, firstFolderMailboxId);

        for (const _folderId of soapRequest.FolderIds.items) {
            const msg = new EmptyFolderResponseMessage();
            response.ResponseMessages.push(msg);

            try {
                const folderMailboxId = findMailboxId(userInfo, _folderId);
                const labels = await getFolderLabels(folderMailboxId);
            
                const folder = findLabel(labels, _folderId);

                if (!folder) {
                    msg.ResponseClass = ResponseClassType.ERROR;
                    msg.ResponseCode = ResponseCodeType.ERROR_FOLDER_NOT_FOUND;
                    continue;
                }

                await doRemoveMessagesFromFolder(userInfo, folder, hardDelete, folderMailboxId);

                // Hack because includeSubFolders, while typed as boolean is coming in as a string
                const includeSubFolders = soapRequest.DeleteSubFolders;
                if ((typeof includeSubFolders === 'string' && includeSubFolders === "true")
                    || (typeof includeSubFolders === 'boolean' && includeSubFolders === true)) {
                    await doRemoveSubFoldersFromFolder(userInfo, folder, labels, hardDelete, folderMailboxId);
                }
            }
            catch (error) {
                const err: any = error; 
                Logger.getInstance().error(`Error occurred emptying folder ${util.inspect(_folderId, false, 2)}: ${err}`);
                throwErrorIfClientResponse(err);

                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_CANNOT_EMPTY_FOLDER;
                msg.MessageXml = { Value: new Value("Error", err.message) }
            }
        }
    }
    catch (error) {
        const err: any = error; 
        Logger.getInstance().error(`Error getting labels: ${err}`);
        throwErrorIfClientResponse(err);

        const msg = new EmptyFolderResponseMessage();
        response.ResponseMessages.push(msg);
        msg.ResponseClass = ResponseClassType.ERROR;
        msg.ResponseCode = ResponseCodeType.ERROR_CANNOT_EMPTY_FOLDER;
        msg.MessageXml = { Value: new Value("Error", err.message) }

    }
    return response;
}

/**
 * Remove all messages from a folder. 
 * @param userInfo The user's credentials to be passed to Keep.
 * @param folder The folder containing the messages.
 * @param hardDelete Indicator if the messages should be deleted from trash as well
 * @throws An error if a request to Keep fails.
 */
async function doRemoveMessagesFromFolder(userInfo: UserInfo, folder: PimLabel, hardDelete: boolean, mailboxId? : string): Promise<void> {
    // Remove all the messages from the folder
    // Need a single API to permanently delete all messages in a folder. moveMessage will leave them in other folders. LABS-1470

    async function deleteMessages(pimMessages: PimMessage[]): Promise<void> {
        const COUNT_OF_PARALLEL = 100;
        let promises: any = [];
        for (const msg of pimMessages) {
            promises.push(
                KeepPimMessageManager.getInstance().deleteMessage(userInfo, msg.unid, mailboxId, hardDelete)
            );
            if (promises.length === COUNT_OF_PARALLEL) {
                await Promise.all(promises);
                promises = [];
            }
        }
        if (promises.length > 0) await Promise.all(promises);
    }

    let messages: PimMessage[];
    // getMailMessages returns only 1000 messages - default value for messages/{labelId} keep API. see keep API documentation
    do {
        messages = await KeepPimMessageManager.getInstance()
            .getMailMessages(userInfo, folder.view, undefined, undefined, undefined, mailboxId);
        await deleteMessages(messages);
    } while (messages.length > 0)
}

/**
 * Remove folder with messages w/o removing of subfolders. 
 * @param userInfo The user's credentials to be passed to Keep.
 * @param folder The folder to be removed
 * @param hardDelete Indicator if the messages should be deleted from trash as well
 * @throws An error if a request to Keep fails.
 */
async function removeFolder(userInfo: UserInfo, pimFolder: PimLabel, hardDelete: boolean, mailboxId?: string): Promise<void> {
    await doRemoveMessagesFromFolder(userInfo, pimFolder, hardDelete, mailboxId);
    await KeepPimLabelManager.getInstance().deleteLabel(userInfo, pimFolder.folderId, undefined, mailboxId);
}

/**
 * Remove all folders with messages w/o removing of subfolders. 
 * @param userInfo The user's credentials to be passed to Keep.
 * @param folders The array of folders to be removed
 * @param hardDelete Indicator if the messages should be deleted from trash as well
 * @throws An error if a request to Keep fails.
 */
async function removeFolders(userInfo: UserInfo, pimFolders: PimLabel[], hardDelete: boolean, mailboxId?: string): Promise<void> {
    const promises = pimFolders.map(pimFolder => removeFolder(userInfo, pimFolder, hardDelete, mailboxId));
    await Promise.all(promises);
}

/**
 * Remove all sub-folders from a folder. 
 * @param userInfo The user's credentials to be passed to Keep.
 * @param folder The parent folder 
 * @param folders The complete list of folders (e.g. labels)
 * @param hardDelete Indicator if the messages should be deleted from trash as well
 * @param levels The array of levels. Each level contains the list of folders belonging to it.
 * @param currentLevel the current level of nested folder
 * @throws An error if a request to Keep fails.
 */
async function doRemoveSubFoldersFromFolder(
    userInfo: UserInfo,
    folder: PimLabel,
    folders: PimLabel[],
    hardDelete: boolean,
    mailboxId: string | undefined,
    levels: Array<PimLabel[]> = [],
    currentLevel = 0
): Promise<void> {
    for (const pimFolder of folders) {
        if (pimFolder.parentFolderId === folder.folderId) {
            await doRemoveSubFoldersFromFolder(userInfo, pimFolder, folders, hardDelete, mailboxId, levels, currentLevel + 1); // Remove any subFolders in this folder
            if (!levels[currentLevel]) levels[currentLevel] = [];
            levels[currentLevel].push(pimFolder);
        }
    }
    if (currentLevel > 0) return;
    levels.reverse(); //reverse because we need to delete from the down to the top.
    for (const pimFolders of levels) {
        await removeFolders(userInfo, pimFolders, hardDelete, mailboxId);
    }
}

/**
 * Convert id.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns id converted
 */
export function ConvertId(soapRequest: ConvertIdType, request: Request): ConvertIdResponseType {
    const response = new ConvertIdResponseType();
    soapRequest.SourceIds.items.forEach(sourceId => {
        const msg = new ConvertIdResponseMessageType();
        msg.AlternateId = sourceId;
        sourceId.Format = soapRequest.DestinationFormat;
        if (sourceId instanceof AlternateIdType) {
            if (sourceId.Id === 'id-archive') {
                sourceId.IsArchive = true;
            }
            sourceId.Id = Buffer.from(sourceId.Id).toString('base64');
        }
        response.ResponseMessages.push(msg);
    });
    return response;
}

/**
 * Get app market place url.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns app market place url
 */
export function GetAppMarketplaceUrl(soapRequest: GetAppMarketplaceUrlType, request: Request): GetAppMarketplaceUrlResponseMessageType {
    const response = new GetAppMarketplaceUrlResponseMessageType();

    response.AppMarketplaceUrl = 'https://appsource.microsoft.com/en-us/marketplace/apps?product=office%3boutlook';
    return response;
}

/**
 * Sync folder hierarchy.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns sync result
 */
export async function SyncFolderHierarchy(soapRequest: SyncFolderHierarchyType, request: Request): Promise<SyncFolderHierarchyResponseType> {
    const userData: UserMailboxData = await getInitializedUserMailData(request);

    const response = new SyncFolderHierarchyResponseType();
    const msg = new SyncFolderHierarchyResponseMessageType();
    response.ResponseMessages.push(msg);
    msg.IncludesLastFolderInRange = true;

    msg.SyncState = '' + Date.now();

    if (soapRequest.SyncState) {
        if (soapRequest.SyncState > msg.SyncState) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_INVALID_SYNC_STATE_DATA;
            msg.MessageText = "Invalid SyncState";
            return response;
        }
        for (const event of Object.values(userData.FOLDER_EVENTS)) {
            if (event.timestamp > Number(soapRequest.SyncState)) {
                msg.Changes.push(event.change);
            }
        }
    }

    try {
        const userInfo = UserContext.getUserInfo(request);
        //Get folder labels specifying a mailboxId of the delegate or non-delegate
        const mailboxId = soapRequest.SyncFolderId ? findMailboxId(userInfo, soapRequest.SyncFolderId) : undefined;
        const pimLabels = await getKeepLabels(userInfo, mailboxId, includeUnReadCountForFolders(soapRequest.FolderShape))
            .catch(err => {
                Logger.getInstance().error(`Syncfolderhierarchy is unable to get PIM labels: ${err}`);
                throw err;
            })
        Logger.getInstance().debug("Syncfolderhierarchy pimLabels = " + util.inspect(pimLabels, false, 5));

        if (soapRequest.SyncState === undefined) {
            // If not sync state, include all folders
            for (const label of pimLabels) {
                //Get folder with combined EWS id
                const folder = labelToFolder(userInfo, label, request, soapRequest.SyncFolderId, soapRequest.FolderShape);
                const change = new SyncFolderHierarchyCreateType();
                populateSyncFolderChange(change, folder);
                msg.Changes.push(change);
            }
        }

        const previousItems = userData.LAST_SYNCFOLDERS ?? [];

        // Look for any existing folders that should be removed
        const deletedItems: PimLabel[] = [];
        previousItems.forEach(label => {
            const exists = pimLabels.find(newLabel => { return newLabel.folderId && label.folderId && newLabel.folderId === label.folderId });
            if (!exists) {
                deletedItems.push(label);
            }
        });

        deletedItems.forEach(label => {
            const change = new SyncFolderHierarchyDeleteType();
            //Get combined EWS id for change
            const combinedChangeId = combineFolderIdAndMailbox(userInfo, label.folderId, soapRequest.SyncFolderId);
            change.FolderId = new FolderIdType(combinedChangeId, `ck-${combinedChangeId}`);

            delete userData.FOLDER_ITEM_EVENTS[change.FolderId.Id];
            delete userData.FOLDER_EVENTS[change.FolderId.Id];

            msg.Changes.push(change);
        });

        if (soapRequest.SyncState) { // If no sync state, then all already included above. 
            // Look for any new labels on the server that don't exist in our local cache and create changes for them
            const newItems: PimLabel[] = [];
            pimLabels.forEach(newLabel => {
                const previousLabel = previousItems.find(oldLabel => { return oldLabel.folderId && newLabel.folderId && newLabel.folderId === oldLabel.folderId });
                const folderEvent = Object.values(userData.FOLDER_EVENTS)
                    .find(event => {
                        if (event.change instanceof SyncFolderHierarchyCreateType && event.change.Folder?.FolderId !== undefined) {
                            //Get change id from combined EWS id
                            const [changeId] = getKeepIdPair(event.change.Folder.FolderId.Id);
                            //Check newlabel folder id equal to change label id
                            return changeId === newLabel.folderId; 
                        } 
                    });
                // Only add if we have not previously returned this folder or there is an outstanding create event (those are added above)
                if (previousLabel === undefined && folderEvent === undefined) {
                    newItems.push(newLabel);
                }
            });

            newItems.forEach(newLabel => {
                const change = new SyncFolderHierarchyCreateType();
                //Get new folder with combined EWS id
                const newFolder = labelToFolder(userInfo, newLabel, request, soapRequest.SyncFolderId, soapRequest.FolderShape);
                if (newFolder) {
                    populateSyncFolderChange(change, newFolder);
                    msg.Changes.push(change);
                }
            });
        }

        userData.LAST_SYNCFOLDERS = pimLabels;
        return response;

    } catch (err) {
        Logger.getInstance().debug("SyncFolderHierarchy error retrieving pim labels:" + err);
        throw err;
    }
}

/**
 * Create folder.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns folder create result
 */
export async function CreateFolder(soapRequest: CreateFolderType, request: Request): Promise<CreateFolderResponseType> {
    const userData = getUserMailData(request);

    const response = new CreateFolderResponseType();

    let parentPimLabel: PimLabel | undefined;
    let parentFolderViewName: string | undefined;
    let labelParentId: string | undefined;
    let parentLabelType: PimLabelTypes | undefined;

    const userInfo = UserContext.getUserInfo(request);
    
    try {
        if (soapRequest.ParentFolderId) {
            if (isTargetRootFolderId(userInfo, soapRequest.ParentFolderId, request)) {
                parentFolderViewName = "";
            } else {
                parentPimLabel = await getLabelByTarget(soapRequest.ParentFolderId, userInfo);

                if (parentPimLabel !== undefined) {
                    parentFolderViewName = parentPimLabel.view;
                    labelParentId = parentPimLabel.folderId;
                    parentLabelType = parentPimLabel.type;
                }
                else {
                    Logger.getInstance().error(`Did not find label for parent ${util.inspect(soapRequest.ParentFolderId, false, 5)}`);
                }

            }
        }
    } catch (err) {
        Logger.getInstance().debug(`CreateFolder: Error getting current folders for parentage.  ${err}`);
        throwErrorIfClientResponse(err);
        // Falling through with parent not found for each item
    }

    for (let folder of soapRequest.Folders.items) {
        const msg = new CreateFolderResponseMessage();
        response.ResponseMessages.push(msg);
        if (parentFolderViewName === undefined) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_PARENT_FOLDER_NOT_FOUND;
            continue;
        }

        try {
            let targetFolder: BaseFolderType | undefined;
            const fDisplayName = folder.DisplayName ?? folder.FolderId?.Id;
            if (fDisplayName) {

                // If no label type at this point, map the folder class to the appropriate label type
                if (parentLabelType === undefined) {
                    parentLabelType = labelTypeFromFolderClass(folder);
                }

                const pimLabelIn = pimLabelFromItem(folder, {}, parentLabelType);
                if (labelParentId) {
                    pimLabelIn.parentFolderId = labelParentId;
                }
                
                //Get mailboxId from combined EWS id
                const  mailboxId = findMailboxId(userInfo, soapRequest.ParentFolderId);

                const folderPimLabel = await KeepPimLabelManager.getInstance().createLabel(userInfo, pimLabelIn, parentLabelType, mailboxId);
                Logger.getInstance().debug(" CreateFolder folderPimLabel=" + util.inspect(folderPimLabel, false, 5));

                targetFolder = labelToFolder(userInfo, folderPimLabel, request, soapRequest.ParentFolderId);
                targetFolder.FolderClass = folder.FolderClass;
                folder = targetFolder;
            }
            else {
                throw new Error("Folder name is not set for " + util.inspect(folder, false, 5));
            }
        } catch (err) {
            Logger.getInstance().error("CreateFolder error creating pimLabel " + err);
            throw err;
        }

        if (folder.FolderId === undefined) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_INVALID_FOLDER_ID;
            continue;
        }

        if (folder.ParentFolderId === undefined) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_PARENT_FOLDER_ID_REQUIRED;
            continue;
        }

        // Setup create folder event
        const event = new SyncFolderHierarchyCreateType();
        event.Folder = folder;
        userData.FOLDER_EVENTS[folder.FolderId.Id] = new FolderEvent(Date.now(), event);

        const notifyEvent = new CreatedEvent();
        notifyEvent.FolderId = folder.FolderId;
        notifyEvent.ParentFolderId = folder.ParentFolderId;
        notifyEvent.TimeStamp = new Date();
        sendStreamingEvent(userData.userId, notifyEvent);

        msg.Folders = new ArrayOfFoldersType();
        msg.Folders.push(folder);
    }

    return response;
}

/**
 * Copy folder not supported.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns folder copy result
 */
export async function CopyFolder(soapRequest: CopyFolderType, request: Request): Promise<CopyFolderResponseType> {
    const userInfo = UserContext.getUserInfo(request);

    let labels: PimLabel[] = [];
    try {
        labels = await KeepPimLabelManager.getInstance().getLabels(userInfo);
    } catch (err) {
        Logger.getInstance().error(`CopyFolder is unable to get PIM labels: ${err}`);
        throw err;
    }

    const response = new CopyFolderResponseType();

    soapRequest.FolderIds.items.forEach(_folderId => {
        const msg = new CopyFolderResponseMessage();
        response.ResponseMessages.push(msg);

        const folder = findLabel(labels, _folderId);

        if (!folder) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_FOLDER_NOT_FOUND;
            return;
        }

        let toFolderId: FolderIdType | undefined;
        if (isTargetRootFolderId(userInfo, soapRequest.ToFolderId)) {
            toFolderId = rootFolderIdForRequest(request); // Copying to Root
        }
        else {
            const toFolder = findLabelByTarget(labels, soapRequest.ToFolderId);
            if (toFolder) {
                toFolderId = new FolderIdType(toFolder.folderId, `ck-${toFolder.folderId}`);
            }
        }

        if (!toFolderId) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_TO_FOLDER_NOT_FOUND;
            return;
        }

        const clone = labelToFolder(userInfo, folder, request);
        clone.FolderId = new FolderIdType(uuidv4(), uuidv4());
        clone.ParentFolderId = toFolderId;

        // TODO: Need Keep API to copy label

        // Remove this code and uncomment out code when Keep API is avaliable. 
        msg.ResponseClass = ResponseClassType.ERROR;
        msg.ResponseCode = ResponseCodeType.ERROR_MOVE_COPY_FAILED;
        msg.MessageText = "CopyFolder is not supported yet.";

        // const uchange = new SyncFolderHierarchyCreateType();
        // populateSyncFolderChange(uchange, clone);
        // userData.FOLDER_EVENTS[clone.FolderId!.Id] = new FolderEvent(Date.now(), uchange);

        // const notifyEvent = new CopiedEvent();
        // notifyEvent.FolderId = clone.FolderId;
        // notifyEvent.ParentFolderId = clone.ParentFolderId!;
        // notifyEvent.OldFolderId = new FolderIdType(folder.folderId, `ck-${folder.folderId}`);
        // notifyEvent.OldParentFolderId = new FolderIdType(folder.parentId!, `ck-${folder.parentId!}`);
        // notifyEvent.TimeStamp = new Date();
        // sendStreamingEvent(userData.userId, notifyEvent);

        // msg.Folders.push(clone);
    });

    return response;
}

/**
 * Move folder.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns folder move result
 */
export async function MoveFolder(soapRequest: MoveFolderType, request: Request): Promise<MoveFolderResponseType> {
    const userData = getUserMailData(request);
    const userInfo = UserContext.getUserInfo(request);

    const toFolderMailboxId = findMailboxId(userInfo, soapRequest.ToFolderId);

    const labels = await getKeepLabels(userInfo, toFolderMailboxId)
        .catch(err => {
            Logger.getInstance().error(`MoveFolder is unable to get PIM labels: ${err}`);
            throw err;
        });

    const response = new MoveFolderResponseType();
    let toFolderId: FolderIdType | undefined;
    let targetMoveParentId: string | undefined;

    if (isTargetRootFolderId(userInfo, soapRequest.ToFolderId, request)) {
        toFolderId = rootFolderIdForRequest(request); // Copying to Root
        targetMoveParentId = KeepPimConstants.MOVE_ROOT_ID;
    }
    else {
        const toFolder = findLabelByTarget(labels, soapRequest.ToFolderId);
        if (toFolder) {
            const toFolderEWSId = combineFolderIdAndMailbox(userInfo, toFolder.folderId, soapRequest.ToFolderId);
            toFolderId = new FolderIdType(toFolderEWSId, `ck-${toFolderEWSId}`);
            targetMoveParentId = toFolder.unid;
        }
    }

    for (const _folderId of soapRequest.FolderIds.items) {
        const folderMailboxId = findMailboxId(userInfo, _folderId);

        if (toFolderMailboxId !== folderMailboxId) {
            throw new Error(`Folder cannot be moved from ${folderMailboxId} to ${toFolderMailboxId} mailbox`);
        }

        const msg = new MoveFolderResponseMessage();
        response.ResponseMessages.push(msg);

        const label = findLabel(labels, _folderId);

        if (!label) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_FOLDER_NOT_FOUND;
            msg.MessageText = `Folder ${util.inspect(_folderId, false, 5)} not found.`;
            continue;
        }

        if (!toFolderId || !targetMoveParentId) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_TO_FOLDER_NOT_FOUND;
            msg.MessageText = `Folder ${util.inspect(soapRequest.ToFolderId, false, 5)} not found.`;
            continue;
        }

        try {
            //Move the item in keep.  
            const moveItems: string[] = [];
            moveItems.push(label.unid);
            const oldParentFolderId = label.parentFolderId ? new FolderIdType(label.parentFolderId, `ck-${label.parentFolderId}`) : undefined;
            
            //Get mailboxId from combined EWS id
            const  mailboxId = findMailboxId(userInfo, soapRequest.FolderIds.items[0]);
            await KeepPimLabelManager.getInstance().moveLabel(userInfo, targetMoveParentId, moveItems, mailboxId);
            //Get folder with combined EWS id
            //Pass the labels here for counting children
            const folder = labelToFolder(userInfo, label, request, soapRequest.ToFolderId, undefined, labels);
            folder.ParentFolderId = toFolderId;

            if (folder.FolderId === undefined) {
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_INVALID_FOLDER_ID;
                continue;
            }

            const uchange = new SyncFolderHierarchyUpdateType();
            populateSyncFolderChange(uchange, folder);
            userData.FOLDER_EVENTS[folder.FolderId.Id] = new FolderEvent(Date.now(), uchange);

            const notifyEvent = new MovedEvent();
            notifyEvent.FolderId = folder.FolderId;
            notifyEvent.ParentFolderId = folder.ParentFolderId ? folder.ParentFolderId: toFolderId;
            notifyEvent.OldFolderId = folder.FolderId;
            if (oldParentFolderId) {
                notifyEvent.OldParentFolderId = oldParentFolderId;
            }
            notifyEvent.TimeStamp = new Date();
            sendStreamingEvent(userData.userId, notifyEvent);

            msg.Folders.push(folder);
        } catch (error) {
            const err: any = error; 
            // If there was an error, delete the message on the server since it didn't make it to the correct location
            Logger.getInstance().debug(`MoveFolder error moving folder, ${label.unid},  to desired folder ${toFolderId.Id}: ${err}`);
            throwErrorIfClientResponse(err);

            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_MOVE_COPY_FAILED;
            msg.MessageText = `${err}`;
        }
    }

    return response;
}

/**
 * Delete folder.
 * @param useData User's mailbox data
 * @param folder The folder's folderType
 * @returns
 */
function doDeleteFolder(request: Request, folder: FolderType): void {
    if (folder.FolderId === undefined) {
        throw new Error(`Folder id is not defined for ${util.inspect(folder, false, 5)}`);
    }

    if (folder.ParentFolderId === undefined) {
        throw new Error(`Parent Folder id is not defined for ${util.inspect(folder, false, 5)}`);
    }

    const userData = getUserMailData(request);
    const userInfo = UserContext.getUserInfo(request);

    const folderId = folder.FolderId;

    delete userData.FOLDER_ITEM_EVENTS[folderId.Id];

    //Get labelId and mailboxId from combined EWS id
    const [labelId, mailboxId] = getKeepIdPair(folderId.Id);

    if (labelId === undefined) {
        throw new Error(`Unable to determine label unid for folder id ${folderId.Id}`);
    }

    KeepPimLabelManager.getInstance()
        .deleteLabel(userInfo, labelId, undefined, mailboxId)
        .catch((error: any) => {
            Logger.getInstance().error(`DeleteLabel failed for ${labelId}: ${util.inspect(error, false, 5)}`);
            throwErrorIfClientResponse(error); // When mockdata is gone this should change to just throw err unless other special soap responses need to be handled
        });

    const change = new SyncFolderHierarchyDeleteType();
    change.FolderId = folderId;
    userData.FOLDER_EVENTS[folderId.Id] = new FolderEvent(Date.now(), change);

    const notifyEvent = new DeletedEvent();
    notifyEvent.FolderId = folder.FolderId;
    notifyEvent.ParentFolderId = folder.ParentFolderId;
    notifyEvent.TimeStamp = new Date();
    sendStreamingEvent(userData.userId, notifyEvent);
}

/**
 * Delete folder.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns folder delete result
 */
export async function DeleteFolder(soapRequest: DeleteFolderType, request: Request): Promise<DeleteFolderResponseType> {
    getUserMailData(request);
    const userInfo = UserContext.getUserInfo(request);

    const response = new DeleteFolderResponseType();

    const firstFolderMailboxId = findMailboxId(userInfo, soapRequest.FolderIds.items[0]);
    const getFolderLabels = await getFolderLabelsFromHash(userInfo, firstFolderMailboxId)
        .catch(err => {
            Logger.getInstance().error(`DeleteFolder is unable to get PIM labels: ${err}`);
            throw err;
        });

    for (const folderId of soapRequest.FolderIds.items) {
        const folderMailboxId = findMailboxId(userInfo, folderId);
        const msg = new DeleteFolderResponseMessage();
        response.ResponseMessages.push(msg);

        const labels = await getFolderLabels(folderMailboxId)
            .catch(err => {
                Logger.getInstance().error(`DeleteFolder is unable to get PIM labels: ${err}`);
                throw err;
            });

        const label = findLabel(labels, folderId);

        if (!label) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_FOLDER_NOT_FOUND;
            continue;
        }

        if (getDistinguishedFolderId(label)) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_DELETE_DISTINGUISHED_FOLDER;
            continue;
        }
        
        doDeleteFolder(request, labelToFolder(userInfo, label, request, folderId));
    }

    return response;
}

/**
 * Update folder.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns folder update result
 */
export async function UpdateFolder(soapRequest: UpdateFolderType, request: Request): Promise<UpdateFolderResponseType> {
    const userData = getUserMailData(request);
    const userInfo = UserContext.getUserInfo(request);

    const response = new UpdateFolderResponseType();

    const firstChangeMailboxId = findMailboxId(userInfo, soapRequest.FolderChanges.FolderChange[0]);
    const getFolderLabels = await getFolderLabelsFromHash(userInfo, firstChangeMailboxId)
        .catch(err => {
            Logger.getInstance().error(`UpdateFolder is unable to get PIM labels: ${err}`);
            throw err;
        });

    for (const change of soapRequest.FolderChanges.FolderChange) {
        const msg = new UpdateFolderResponseMessage();
        response.ResponseMessages.push(msg);

        const changeMailboxId = findMailboxId(userInfo, change);
        const labels = await getFolderLabels(changeMailboxId)
            .catch(err => {
                Logger.getInstance().error(`UpdateFolder is unable to get PIM labels: ${err}`);
                throw err;
            });

        let label: PimLabel | undefined;
        if (change.FolderId) {
            label = findLabel(labels, change.FolderId);
        }
        else if (change.DistinguishedFolderId) {
            label = findLabel(labels, change.DistinguishedFolderId);
        }

        if (!label) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_FOLDER_NOT_FOUND;
            continue;
        }
        
        //Get folder with combined EWS id
        const folder = labelToFolder(userInfo, label, request, change);
        if (folder.FolderId === undefined) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_INVALID_FOLDER_ID;
            break;
        }

        if (folder.ParentFolderId === undefined) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_PARENT_FOLDER_ID_REQUIRED;
            break;
        }
        
        for (const update of change.Updates.items) {
            if (update instanceof SetFolderFieldType) {
                if (update.Folder === undefined) {
                    msg.ResponseClass = ResponseClassType.ERROR;
                    msg.ResponseCode = ResponseCodeType.ERROR_FOLDER_NOT_FOUND;
                    break;
                }

                if (update.FieldURI) {
                    const field = update.FieldURI.FieldURI.split(':')[1];
                    if (update.FieldURI.FieldURI === UnindexedFieldURIType.FOLDER_DISPLAY_NAME) {
                        try {
                            label.displayName = (update.Folder as any)[field];
                            
                            //Get mailboxId from combined EWS id
                            const  mailboxId = findMailboxId(userInfo, change);
                            await KeepPimLabelManager.getInstance().updateLabel(userInfo, label, mailboxId)
                        } catch (err) {
                            Logger.getInstance().debug(`An error occurred renaming folder, ${label.displayName} to ${(update.Folder as any)[field]}: ${err}`);
                            throwErrorIfClientResponse(err);
                            msg.ResponseClass = ResponseClassType.ERROR;
                            msg.ResponseCode = ResponseCodeType.ERROR_FOLDER_SAVE_PROPERTY_ERROR;
                            msg.MessageText = `${err}`;
                            break;
                        }
                    }
                    (folder as any)[field] = (update.Folder as any)[field];
                } else if (update.ExtendedFieldURI) {
                    let foundep: ExtendedPropertyType | undefined;
                    const newep = update.Folder.ExtendedProperty !== undefined ? update.Folder.ExtendedProperty[0] : undefined;
                    if (newep) {
                        if (folder.ExtendedProperty) {
                            foundep = folder.ExtendedProperty.find(ep => {
                                return ep.ExtendedFieldURI.PropertyType === update.ExtendedFieldURI?.PropertyType
                                    && (ep.ExtendedFieldURI.PropertyId === update.ExtendedFieldURI?.PropertyId
                                        || ep.ExtendedFieldURI.PropertySetId === update.ExtendedFieldURI?.PropertySetId
                                        || ep.ExtendedFieldURI.PropertyTag === update.ExtendedFieldURI?.PropertyTag);
                            });
                            if (foundep) {
                                foundep.Value = newep.Value;
                                foundep.Values = newep.Values;
                                if (!foundep.Value) {
                                    delete foundep.Value;
                                }
                                if (!foundep.Values) {
                                    delete foundep.Values;
                                }
                            }
                        }
                        if (!foundep) {
                            folder.ExtendedProperty = folder.ExtendedProperty ?? [];
                            folder.ExtendedProperty.push(newep);
                        }
                    }
                    else {
                        Logger.getInstance().error(`Extended property not found for update folder extended property for :${util.inspect(update.Folder, false, 5)}`);
                    }
                }
            }
        }

        folder.FolderId.ChangeKey = uuidv4();

        const uchange = new SyncFolderHierarchyUpdateType();
        populateSyncFolderChange(uchange, folder);
        userData.FOLDER_EVENTS[folder.FolderId.Id] = new FolderEvent(Date.now(), uchange);

        const notifyEvent = new ModifiedEvent();
        notifyEvent.FolderId = folder.FolderId;
        notifyEvent.ParentFolderId = folder.ParentFolderId;
        notifyEvent.TimeStamp = new Date();
        sendStreamingEvent(userData.userId, notifyEvent);

        msg.Folders.push(folder);
    }

    return response;
}

/**
 * Sync folder items.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns folder items sync result
 */
export async function SyncFolderItems(soapRequest: SyncFolderItemsType, request: Request): Promise<SyncFolderItemsResponseType> {
    const userData: UserMailboxData = await getInitializedUserMailData(request);
    const userInfo = UserContext.getUserInfo(request);

    const response = new SyncFolderItemsResponseType();
    const msg = new SyncFolderItemsResponseMessageType();
    response.ResponseMessages.push(msg);

    if (isTargetRootFolderId(userInfo, soapRequest.SyncFolderId, request)) {
        // There will not be any items in the root folder. 
        msg.SyncState = new SyncStateInfo(Date.now().toString(), 0).getBase64SyncState();
        msg.IncludesLastItemInRange = true;
        return response;
    }
    
    let label: PimLabel | undefined;
    try {
        label = await getLabelByTarget(soapRequest.SyncFolderId, userInfo);
    } catch (err) {
        Logger.getInstance().error(`SyncFolderItems is unable to get PIM labels: ${err}`);
        msg.ResponseClass = ResponseClassType.ERROR;
        msg.ResponseCode = ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR;
        msg.MessageText = `${err}`;
        return response;
    }

    if (!label) {
        msg.ResponseClass = ResponseClassType.ERROR;
        msg.ResponseCode = ResponseCodeType.ERROR_SYNC_FOLDER_NOT_FOUND;
        return response;
    }
    //Get combined EWS id
    const folderEWSId = combineFolderIdAndMailbox(userInfo, label.folderId, soapRequest.SyncFolderId);

    let startIndex = 0;
    let nextSyncIndex = 0;
    const count = Math.min(soapRequest.MaxChangesReturned, 512);
    const newSyncStateInfo = new SyncStateInfo(Date.now().toString(), nextSyncIndex);
    const previousSyncStateInfo = SyncStateInfo.parse(soapRequest.SyncState);
    if (previousSyncStateInfo) {
        startIndex = previousSyncStateInfo.lastIndex;
    }
    msg.IncludesLastItemInRange = true;
    if (previousSyncStateInfo?.timestamp && newSyncStateInfo.timestamp && 
        (previousSyncStateInfo.timestamp > newSyncStateInfo.timestamp)) {
        msg.ResponseClass = ResponseClassType.ERROR;
        msg.ResponseCode = ResponseCodeType.ERROR_INVALID_SYNC_STATE_DATA;
        msg.MessageText = "Invalid SyncState";
        return response;
    }

    const includeMimeContent = soapRequest.ItemShape.IncludeMimeContent ?? false;

    // const ignores: string[] = [];
    // if (soapRequest.Ignore) {
    //     soapRequest.Ignore.ItemId.forEach(i => {
    //         ignores.push(i.Id);
    //     })
    // }
    // 
    // TODO: Something similar to this can be used when we have event support in the Event Arc API. 
    // const distinguishedFolderId = getDistinguishedFolderId(label);
    // const events = userData.FOLDER_ITEM_EVENTS[folderId] ?? {};
    //
    //Bypass inbox and drafts for now....FIXME This is just a hack to ensure we sync the inbox for return emails.  Once we get events, we must remove this check for th inbox
    // const isBypassSyncState = distinguishedFolderId === DistinguishedFolderIdNameType.INBOX || distinguishedFolderId === DistinguishedFolderIdNameType.DRAFTS;
    // if (soapRequest.SyncState && !isBypassSyncState) {
    //     for (const id of Object.keys(events)) {
    //         const event = events[id];
    //         if (event.timestamp > Number(soapRequest.SyncState)) {
    //             if (event.change instanceof SyncFolderItemsDeleteType) {
    //                 msg.Changes.push(event.change);
    //                 continue;
    //             }
    //             if (ignores.indexOf(id) > -1) {
    //                 continue;
    //             }
    //             if (!includeMimeContent && (event.change instanceof SyncFolderItemsCreateType
    //                 || event.change instanceof SyncFolderItemsUpdateType)) {
    //                 const itemId = new ItemIdType(id, `ck-${id}`);
    //                 try {

    //                     const item = await getItem(userData, itemId, userInfo, soapRequest.ItemShape, includeMimeContent);
    //                     if (item) {
    //                         const change = populateSyncItemChange(newInstance(event.change), item);
    //                         msg.Changes.push(change);
    //                     }
    //                 } catch (err) {
    //                     // Log the error, but do not push the change.
    //                     Logger.getInstance().debug(`Error occurrect getting item for SynchFolderItems create event for id ${itemId}: ${err}`);
    //                     throwErrorIfClientResponse(err);
    //                 }

    //             } else {
    //                 msg.Changes.push(event.change);
    //             }
    //         }
    //     }
    //     return response;
    // }

    let ewsItems: ItemType[] = []; // List of items to process

    // First try using a service manager
    const manager = EWSServiceManager.getInstanceFromPimItem(label);
    if (manager) {
        const mailboxId = findMailboxId(userInfo, soapRequest.SyncFolderId);
        ewsItems = await manager.getItems(userInfo, request, soapRequest.ItemShape, startIndex, count, label, mailboxId);
        if (ewsItems.length === 0 || ewsItems.length < count) {
            msg.IncludesLastItemInRange = true;
            nextSyncIndex = 0;  // Start at the beginning again if synching the same folder....when event processing is active this does not need to restart at 0
        } else {
            nextSyncIndex = startIndex + ewsItems.length;
        }
    }
    else {
        Logger.getInstance().error(`Unable to create service manager for SyncFolderItems with label ${label.folderId}`);
    }

    newSyncStateInfo.lastIndex = nextSyncIndex;
    msg.SyncState = newSyncStateInfo.getBase64SyncState();
    // Only compare against previous items if sync state is set
    const previousItems = soapRequest.SyncState ? userData.LAST_SYNCFOLDER_ITEMS[folderEWSId] ?? [] : [];
    // Add the fetched items to the response message
    addChanges(folderEWSId, ewsItems, previousItems, includeMimeContent, msg, request);

    userData.LAST_SYNCFOLDER_ITEMS[folderEWSId] = ewsItems; // Save for next time

    return response;
}

/**
* Add changes to SyncFolderItems response. On return the changes will be added to the msg object. 
* @param folderEWSId The folder containing the items.
* @param ewsItems The current EWS items retrieved from Keep and processed.
* @param previousItems The items retrieved from the last SyncFolderItems
* @param includeMimeContent True to include the mime content in the changes; otherwise false
* @param msg The SyncFolderItems response message.
* @param userData Current user's data. 
*/
function addChanges(
    folderEWSId: string, 
    ewsItems: ItemType[], 
    previousItems: ItemType[], 
    includeMimeContent: boolean, 
    msg: SyncFolderItemsResponseMessageType, 
    request: Request
): void {
    const userData = getUserMailData(request);

    const deletedItems: ItemType[] = [];
    previousItems.forEach(item => {
        const exists = ewsItems.find(newItem => { return newItem.ItemId && item.ItemId && newItem.ItemId.Id === item.ItemId.Id });
        if (!exists) {
            deletedItems.push(item);
        }
    });

    deletedItems.forEach(item => {
        if (item.ItemId) {
            const change = new SyncFolderItemsDeleteType();
            change.ItemId = item.ItemId;
            if (userData.FOLDER_ITEM_EVENTS && userData.FOLDER_ITEM_EVENTS[folderEWSId] && userData.FOLDER_ITEM_EVENTS[folderEWSId][item.ItemId.Id]) {
                delete userData.FOLDER_ITEM_EVENTS[folderEWSId][item.ItemId.Id];
            }
            msg.Changes.push(change);
        }
    });

    if (Date.now() > userData.PROBLEM_ITEM_NEXT_SYNC_FLUSH) {
        // 5 minutes has passed....we'll flush the PROBLEM_ITEM_CACHE and let it check again
        userData.PROBLEM_ITEMS = [];
        userData.PROBLEM_ITEM_NEXT_SYNC_FLUSH = Date.now() + (5 * 60000);
    }
    //Look for any new messages on the server that don't exist the last SyncFolderItems and create changes for them
    const newItems: ItemType[] = [];
    ewsItems.forEach(newItem => {
        const exists = previousItems.find(oldItem => { return oldItem.ItemId && newItem.ItemId && newItem.ItemId.Id === oldItem.ItemId.Id });
        if (newItem.ItemId && userData.PROBLEM_ITEMS.includes(newItem.ItemId.Id)) {
            const change = new SyncFolderItemsDeleteType();
            change.ItemId = newItem.ItemId;
            if (userData.FOLDER_ITEM_EVENTS && userData.FOLDER_ITEM_EVENTS[folderEWSId] && userData.FOLDER_ITEM_EVENTS[folderEWSId][newItem.ItemId.Id]) {
                delete userData.FOLDER_ITEM_EVENTS[folderEWSId][newItem.ItemId.Id];
            }
            msg.Changes.push(change);
        } else if (!exists) {
            newItems.push(newItem);
        }
    });

    newItems.forEach(newItem => {
        const change = new SyncFolderItemsCreateType();
        populateSyncItemChange(change, includeMimeContent ? newItem : omit(newItem, 'MimeContent'));
        msg.Changes.push(change);
    });
}

/**
 * Get item.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns item retrieved
 */
export async function GetItem(soapRequest: GetItemType, request: Request): Promise<GetItemResponseType> {
    const userData = getUserMailData(request);

    const response = new GetItemResponseType();

    const includeMimeContent = soapRequest.ItemShape.IncludeMimeContent;
    for (const itemIdType of soapRequest.ItemIds.items) {
        const msg = new GetItemResponseMessage();
        response.ResponseMessages.push(msg);

        if (!(itemIdType instanceof ItemIdType)) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_INVALID_ID;
            msg.MessageText = `Item id type is not supported`;
            continue;
        }
        
        let item: ItemType | undefined;

        try {
            if (userData.PROBLEM_ITEMS.includes(itemIdType.Id)) {
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_ITEM_CORRUPT;
                msg.MessageText = `GetItem: Error occurred getting item ${itemIdType.Id}`;
                continue;
            }

            // First try using the service manager
            const userInfo = UserContext.getUserInfo(request);
            item = await EWSServiceManager.getItem(itemIdType.Id, userInfo, request, soapRequest.ItemShape);
        } catch (error) {
            const err: any = error; 
            Logger.getInstance().error(`GetItem: Error getting item: ${err}`)
            if (err.ResponseCode) {
                userData.PROBLEM_ITEMS.push(itemIdType.Id);
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_ITEM_CORRUPT;
                // msg.ResponseCode = err.ResponseCode;
                msg.MessageText = err.MessageText ?? `GetItem: Error occurred getting item ${itemIdType.Id}`;
                continue;
            } else {
                if (err.status && err.status !== 401 && err.status !== 403 && err.status !== 404) {
                    userData.PROBLEM_ITEMS.push(itemIdType.Id);
                }
                throwErrorIfClientResponse(err);
            }
        }

        if (!item) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_ITEM_NOT_FOUND;
            msg.MessageText = `Item id does not exist: ${itemIdType.Id}`;
            continue;
        }

        // Apple Mail will crash if asking for mimeContent and nothing is returned....
        // Debug code to check if we reported a different type on SyncFolderItems than on GetItem
        // outerLoop:
        // for (const fItems of Object.values(userData.LAST_SYNCFOLDER_ITEMS)) {
        //     for (const fItem of fItems) {
        //         if (item.ItemId && item.ItemId.Id === fItem.ItemId?.Id) {
        //             if (typeof(item) !== typeof(fItem)){
        //                 Logger.getInstance().debug('invalidType');
        //             }
        //             break outerLoop;
        //         }
        //     }
        // }

        // If it thinks its item ItemClass is a MessageType and we're returning a different type
        // This 'if includeMimeContent with bogusMime' can be removed once LABS-1281 is fixed
        if (includeMimeContent && item instanceof MessageType && !item.MimeContent) {
            const bogusMime = "";
            item.MimeContent = new MimeContentType(base64Encode(bogusMime), 'UTF-8');
        }
        Logger.getInstance().debug("GetItem msg item=" + util.inspect(item, false, 5));
        msg.Items.push(item);
    }

    return response;
}

/**
 * Do create item. Assign id, add to user mailbox, send events.
 * @param userData The user mailbox data
 * @param item The item to create
 * @param folder The parent folder
 */
function doMockDataCreateItem(userInfo: UserInfo, userData: UserMailboxData, items: ItemType[], folderIdType?: FolderIdType): ItemType[] {
    for (const item of items) {
        if (!item.ItemId) {
            const mailboxId = folderIdType ? findMailboxId(userInfo, folderIdType) : undefined;
            const itemEWSId = getEWSId(uuidv4(), mailboxId);
            item.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
        }
        const itemId = item.ItemId.Id;
        if (folderIdType) {
            item.ParentFolderId = folderIdType;
        }
    
        if (item.ParentFolderId) {
            const folderId = item.ParentFolderId.Id;
    
            const change = new SyncFolderItemsCreateType();
            populateSyncItemChange(change, item);
            userData.FOLDER_ITEM_EVENTS[folderId] = userData.FOLDER_ITEM_EVENTS[folderId] ?? {};
            userData.FOLDER_ITEM_EVENTS[folderId][itemId] = new ItemEvent(Date.now(), change);
        }
    
        const notifyEvent = new CreatedEvent();
        notifyEvent.ItemId = item.ItemId;
        if (item.ParentFolderId) {
            notifyEvent.ParentFolderId = item.ParentFolderId;
        }
        notifyEvent.TimeStamp = new Date();
        sendStreamingEvent(userData.userId, notifyEvent);    
    }

    return items;
}

/**
 * Do create item. Assign id, add to user mailbox, send events.
 * @param request The request being processed
 * @param item The item to create
 * @param folderIdType Folder id forthe parent folder
 * @param messageDispositionType Optional save/send disposition for a message item
 * @param toLabel Optional parent label for the new item
 * @param calendarOperationType Optional save/send disposition for a calendar invite item
 * @returns 
 */
async function doCreateItem(
    request: Request, 
    item: ItemType, 
    folderIdType?: FolderIdType, 
    messageDispositionType?: MessageDispositionType, 
    toLabel?: PimLabel | undefined, 
    calendarOperationType?: CalendarItemCreateOrDeleteOperationType
): Promise<ItemType[]> {
    const userData = getUserMailData(request);
    const userInfo = UserContext.getUserInfo(request);

    let targetItems: ItemType[];

    try {
        const manager = EWSServiceManager.getInstanceFromItem(item);
        if (manager) {
            let targetLabel: PimLabel | undefined = toLabel;
            if (!toLabel && folderIdType) {
                const labels = await KeepPimLabelManager.getInstance().getLabels(userInfo);
                targetLabel = findLabel(labels, folderIdType);
            }
            let msgDispositionType = messageDispositionType;
            if (calendarOperationType) {
                if (calendarOperationType === CalendarItemCreateOrDeleteOperationType.SEND_ONLY_TO_ALL) {
                    msgDispositionType = MessageDispositionType.SEND_ONLY;
                } else if (calendarOperationType === CalendarItemCreateOrDeleteOperationType.SEND_TO_ALL_AND_SAVE_COPY) {
                    msgDispositionType = MessageDispositionType.SEND_AND_SAVE_COPY;
                }
            }
            const mailboxId = folderIdType ? findMailboxId(userInfo, folderIdType) : undefined;
            targetItems = await manager.createItem(item, userInfo, request, msgDispositionType, targetLabel, mailboxId);
        }
        else {
            throw new Error(`Unknown item type for create item.`);
        }
    } catch (err) {
        Logger.getInstance().debug("doCreateItem error creating message " + err);
        throw err;
    }

    doMockDataCreateItem(userInfo, userData, targetItems, folderIdType);
    return targetItems;
}

/**
 * Do update item. Assign change key, send events.
 * @param userData The user mailbox data
 * @param item The item to create
 */
function doUpdateItem(userData: UserMailboxData, item: ItemType): void {

    if (!item.ParentFolderId) {
        Logger.getInstance().error(`Cannot send modify event for item with no parent id: ${util.inspect(item, false, 5)}`);
        return;
    }

    if (!item.ItemId) {
        Logger.getInstance().error(`Cannot send modify event for item with no item id: ${util.inspect(item, false, 5)}`);
        return;
    }

    item.ItemId.ChangeKey = `ck-${item.ItemId.Id}`;

    const itemId = item.ItemId.Id;
    const folderId = item.ParentFolderId.Id;

    const change = new SyncFolderItemsUpdateType();
    populateSyncItemChange(change, item);
    userData.FOLDER_ITEM_EVENTS[folderId] = userData.FOLDER_ITEM_EVENTS[folderId] ?? {};
    userData.FOLDER_ITEM_EVENTS[folderId][itemId] = new ItemEvent(Date.now(), change);

    const notifyEvent = new ModifiedEvent();
    notifyEvent.ItemId = item.ItemId;
    notifyEvent.ParentFolderId = item.ParentFolderId;
    notifyEvent.TimeStamp = new Date();
    sendStreamingEvent(userData.userId, notifyEvent);
}

/**
 * Delete the mock data for an item.  Delete related attachments also, send events.
 * @param userData The current user's mailbox data.
 * @param item The item to delete
 * @param updateEvents True to send event updates, false to skip sending events. Default is true. 
 */
function doMockDataDeleteItem(userData: UserMailboxData, itemId: ItemIdType, parentFolderId: FolderIdType | undefined, updateEvents = true): void {
    const itemIdBase = itemId.Id;

    if (!parentFolderId) {
        return;
    }

    const folderIdBase = parentFolderId.Id;
    if (updateEvents) {
        const dchange = new SyncFolderItemsDeleteType();
        dchange.ItemId = deepcoy(itemId);
        userData.FOLDER_ITEM_EVENTS[folderIdBase] = userData.FOLDER_ITEM_EVENTS[folderIdBase] ?? {};
        userData.FOLDER_ITEM_EVENTS[folderIdBase][itemIdBase] = new ItemEvent(Date.now(), dchange);

        const notifyEvent = new DeletedEvent();
        notifyEvent.ItemId = itemId;
        notifyEvent.ParentFolderId = parentFolderId;
        notifyEvent.TimeStamp = new Date();
        sendStreamingEvent(userData.userId, notifyEvent);
    }
}

/**
 * Do delete item from server.
 * @param request The request being processed
 * @param itemIdType The item id type with EWS id to delete
 * @param parentFolderId The item parent folder to delete
 * @param hardDelete The type of deletion 
 * @throws an error if the delete fails
 */
async function doDeleteItem(
    request: Request, 
    itemIdType: ItemIdType,
    parentFolderId: FolderIdType | undefined, 
    hardDelete = false
): Promise<void> {
    const userData = getUserMailData(request);
    const userInfo = UserContext.getUserInfo(request);

    const [itemId, mailboxId] = getKeepIdPair(itemIdType.Id);

    let pimItem: PimItem | undefined;

    try {
        if (itemId) {
            pimItem = await KeepPimManager.getInstance().getPimItem(itemId, userInfo, mailboxId);
        }
    } catch (error) {
        const err: any = error; 
        Logger.getInstance().error(`Failed to get item ${itemId}: ${err}`);
        const msg: string = err.message;
        // Temporary work around for LABS-1551. Should not fail if item is already delete. 
        if (msg.indexOf("com.hcl.domino.exception.DocumentDeletedException") === -1) {
            throw err;
        }
    }

    if (pimItem) {
        const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
        if (manager) {
            await manager.deleteItem(pimItem, userInfo, mailboxId, hardDelete);
        }
        else {
            throw new Error(`Unknown item type for delete item.`);
        }
    }
    doMockDataDeleteItem(userData, itemIdType, parentFolderId, true);

}

/**
 * Create folder path.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns folder path created
 */
export async function CreateFolderPath(soapRequest: CreateFolderPathType, request: Request): Promise<CreateFolderPathResponseType> {
    const userInfo = UserContext.getUserInfo(request);

    const response = new CreateFolderPathResponseType();

    let noParent = false; // True if the parent does not exist
    let parentId: string | undefined;

    if (!isTargetRootFolderId(userInfo, soapRequest.ParentFolderId, request)) {
        try {
            const root = await getLabelByTarget(soapRequest.ParentFolderId, userInfo);
            if (root === undefined) {
                noParent = true;
            }
            else {
                parentId = root.folderId;
            }
        } catch (err) {
            Logger.getInstance().error(`CreateFolderPath is unable to get PIM label: ${err}`);
            throw err;
        }

    }

    for (const newFolder of soapRequest.RelativeFolderPath.items) {
        const msg = new CreateFolderPathResponseMessage();

        if (noParent) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_PARENT_FOLDER_NOT_FOUND;
            response.ResponseMessages.push(msg);
            continue;
        }

        if (newFolder.DisplayName === undefined) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_INVALID_FOLDER_TYPE_FOR_OPERATION;
            response.ResponseMessages.push(msg);
            noParent = true;
            continue;
        }

        try {
            const pimLabelIn = pimLabelFromItem(newFolder, {});

            if (parentId) {
                pimLabelIn.parentFolderId = parentId;
            }
            //Get mailboxId from combined EWS id
            const mailboxId = findMailboxId(userInfo, soapRequest.ParentFolderId);
            const newLabel = await KeepPimLabelManager.getInstance().createLabel(userInfo, pimLabelIn, undefined, mailboxId);

            const folder = labelToFolder(userInfo, newLabel, request, soapRequest.ParentFolderId);
            msg.Folders.push(folder);
            response.ResponseMessages.push(msg);

            parentId = newLabel.folderId; // Use this as parent for next reqeust. 
        }
        catch (error) {
            const err: any = error; 
            Logger.getInstance().error(`Failed to create new folder ${util.inspect(newFolder, false, 5)}: ${err}`);
            throwErrorIfClientResponse(err);

            noParent = true;
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_FOLDER_SAVE_FAILED;
            msg.MessageText = err.Message;
            response.ResponseMessages.push(msg);
        }
    }

    return response;
}


/**
 * Upload items.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns upload result
 */
export async function UploadItems(soapRequest: UploadItemsType, request: Request): Promise<UploadItemsResponseType> {
    const userData = getUserMailData(request);
    const userInfo = UserContext.getUserInfo(request);

    const response = new UploadItemsResponseType();
    try {
        for (const uploadItem of soapRequest.Items.Item) {
            const msg = new UploadItemsResponseMessageType();
            response.ResponseMessages.push(msg);

            const [parentFolderId] = getKeepIdPair(uploadItem.ParentFolderId.Id);

            if (uploadItem.IsAssociated) {
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR;
                msg.MessageText = `UploadItems for Folders is not implemented`;
                continue;
            }

            switch (uploadItem.CreateAction) {
                case CreateActionType.CREATE_NEW: {
                    msg.ItemId = await uploadCreateItem(userData, request, uploadItem);
                    break;
                }
                case CreateActionType.UPDATE: {
                    if (!uploadItem.ItemId) {
                        msg.ResponseClass = ResponseClassType.ERROR;
                        msg.ResponseCode = ResponseCodeType.ERROR_INVALID_ID_EMPTY;
                        msg.MessageText = `UploadItems update request did not include an item id.`;
                        break;
                    }
                    try {
                        const [pimItemId, mailboxId] = getKeepIdPair(uploadItem.ItemId.Id);
                        if (pimItemId) {
                            const idString = await KeepPimMessageManager
                                .getInstance()
                                .updateMimeMessage(userInfo, pimItemId, uploadItem.Data, true, parentFolderId);
                            const itemEWSId = getEWSId(idString, mailboxId);
                            msg.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
                        } else {
                            throw new Error(`Invalid PIM item id ${uploadItem.ItemId.Id}`);
                        }
                    } catch (err) {
                        Logger.getInstance().debug(`An error occurred updating an item for upload: ${err}`);
                        throw err;
                    }
                    break;
                }
                case CreateActionType.UPDATE_OR_CREATE: {
                    if (!uploadItem.ItemId) {
                        // No item id.  Just try create.
                        msg.ItemId = await uploadCreateItem(userData, request, uploadItem);
                    } else {
                        try {
                            const [pimItemId, mailboxId] = getKeepIdPair(uploadItem.ItemId.Id);
                            if (pimItemId) {
                                const idString = await KeepPimMessageManager
                                    .getInstance()
                                    .updateMimeMessage(userInfo, pimItemId, uploadItem.Data, true, parentFolderId);
                                const itemEWSId = getEWSId(idString, mailboxId);
                                msg.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
                            } else {
                                throw new Error(`Invalid PIM item id ${uploadItem.ItemId.Id}`);
                            }
                        } catch (error) {
                            const err: any = error; 
                            Logger.getInstance().debug(`An error occurred updating an item for upload: ${err}.  Checking create.`);
                            if (err.status === 404) {
                                // Item was not found....create it
                                msg.ItemId = await uploadCreateItem(userData, request, uploadItem);
                            } else {
                                throw err;
                            }
                        }
                    }
                    break;
                }
            }
        }
    } catch (err) {
        Logger.getInstance().debug(`An error occurred attempting to upload items: ${err}`);
        throw err;
    }

    return response;
}

async function uploadCreateItem(userData: UserMailboxData, request: Request, uploadItem: UploadItemType): Promise<ItemIdType> {
    let item = new MessageType();
    item.DateTimeCreated = new Date();
    item.DateTimeReceived = new Date();

    item.MimeContent = new MimeContentType(uploadItem.Data);
    item = await populateItemFromMime(userData, item, item.MimeContent.Value);

    const resultItems = await doCreateItem(request, item, uploadItem.ParentFolderId, MessageDispositionType.SAVE_ONLY);
    if (resultItems[0].ItemId === undefined) {
        throw new Error(`No item id set for created uploaded item: ${uploadItem.ItemId}`)
    }
    return resultItems[0].ItemId;
}

/**
 * Send item.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns send item result
 */
export async function SendItem(soapRequest: SendItemType, request: Request): Promise<SendItemResponseType> {
    const response = new SendItemResponseType();

    const userInfo = UserContext.getUserInfo(request);
    for (const itemIdType of soapRequest.ItemIds.items) {
        const msg = new SendItemResponseMessage();
        response.ResponseMessages.push(msg);
        if (itemIdType instanceof ItemIdType) {
            try {
                const [itemId, mailboxId] = getKeepIdPair(itemIdType.Id);

                let mimeContent: string;
                if (itemId) {
                    mimeContent = await KeepPimMessageManager.getInstance().getMimeMessage(userInfo, itemId, mailboxId);
                    if (mimeContent) {
                        let folderId: string | undefined;
                        if (soapRequest.SavedItemFolderId?.FolderId) {
                            [folderId] = getKeepIdPair(soapRequest.SavedItemFolderId?.FolderId?.Id);
                        }
                        await KeepPimMessageManager
                            .getInstance()
                            .updateMimeMessage(userInfo, itemId, base64Encode(mimeContent), true, folderId, undefined, mailboxId);
                    }
                }
                else {
                    msg.ResponseCode = ResponseCodeType.ERROR_INVALID_ITEM_FOR_OPERATION_SEND_ITEM;
                    msg.ResponseClass = ResponseClassType.ERROR;
                    msg.MessageText = "Could not find item to send";
                }
            } catch (error) {
                const err: any = error; 
                msg.ResponseCode = ResponseCodeType.ERROR_INVALID_ITEM_FOR_OPERATION_SEND_ITEM;
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.MessageText = err.message;
            }
        }
    }
    return response;
}

/**
 * Create item.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns create item result
 */
export async function CreateItem(soapRequest: CreateItemType, request: Request): Promise<CreateItemResponseType> {
    const userInfo = UserContext.getUserInfo(request);
    const response = new CreateItemResponseType();

    // SavedItemFolder is optional.  
    let toFolderIdType: FolderIdType | undefined;
    let toLabel: PimLabel | undefined;
    if (soapRequest.SavedItemFolderId) {
        try {
            toLabel = await getLabelByTarget(soapRequest.SavedItemFolderId, userInfo);
            if (toLabel) {
                const folderEWSIdType = combineFolderIdAndMailbox(userInfo, toLabel.folderId, soapRequest.SavedItemFolderId);
                toFolderIdType = new FolderIdType(folderEWSIdType, `ck-${folderEWSIdType}`);
            }
        } catch (error) {
            const err: any = error; 
            if (err.status !== 404) {
                Logger.getInstance().debug(`An error occurred getting all labels for CreateItem: ${err}`);
                throw err;
            }
        }
    }

    for (const item of soapRequest.Items.items) {

        const msg = new CreateItemResponseMessage();
        response.ResponseMessages.push(msg);

        if (soapRequest.SavedItemFolderId && !toFolderIdType) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_FOLDER_NOT_FOUND;
            continue;
        }

        try {
            item.DateTimeCreated = new Date();
            item.DateTimeReceived = new Date();

            const resultItems = await doCreateItem(request, item, toFolderIdType, soapRequest.MessageDisposition, toLabel, soapRequest.SendMeetingInvitations);
            for (const resultItem of resultItems) {
                msg.Items.push(resultItem);
            }
        } catch (error) {
            const err: any = error; 
            Logger.getInstance().debug(`CreateItem: An error occurred creating item ${item}`);
            if (err.status && err.status === 403) {
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_CREATE_ITEM_ACCESS_DENIED;
                msg.MessageText = `CreateItem: An error occurred creating item ${item}`;
            } else {
                throw err;
            }
        }
    }

    return response;
}

/**
 * Update item.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns update item result
 */
export async function UpdateItem(soapRequest: UpdateItemType, request: Request): Promise<UpdateItemResponseType> {
    const userData = getUserMailData(request);
    const userInfo = UserContext.getUserInfo(request);

    const response = new UpdateItemResponseType();
    const responseShapeType = new ItemResponseShapeType();
    responseShapeType.BaseShape = DefaultShapeNamesType.DEFAULT;

    // SavedItemFolder is optional.  
    let toLabel: PimLabel | undefined;
    if (soapRequest.SavedItemFolderId) {
        try {
            toLabel = await getLabelByTarget(soapRequest.SavedItemFolderId, userInfo);
        } catch (err) {
            Logger.getInstance().debug(`An error occurred getting all labels for UpdateItem: ${err}`);
            throw err;
        }
    }

    const savedFolderMailboxId = soapRequest.SavedItemFolderId ? findMailboxId(userInfo, soapRequest.SavedItemFolderId) : undefined;

    for (const change of soapRequest.ItemChanges.ItemChange) {
        const msg = new UpdateItemResponseMessageType();
        response.ResponseMessages.push(msg);

        if (change.ItemId === undefined) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_INVALID_REQUEST;
            msg.MessageText = 'Item id is missing in ItemChange';
            continue; 
        }

        const itemEWSId = change.ItemId.Id;
        const [, itemMailboxId] = getKeepIdPair(itemEWSId);
        //Should we throw an error when updates include new saved folder from another mailbox
        if (soapRequest.SavedItemFolderId && itemMailboxId !== savedFolderMailboxId) {
            throw new Error(`Item cannot be moved from ${itemMailboxId} to ${savedFolderMailboxId} mailbox`);
        }

        let currentItem: ItemType | undefined = undefined;

        try {
            currentItem = await EWSServiceManager.updateItem(change, userInfo, request, toLabel);
            if (currentItem) {
                // Update the local cache
                doUpdateItem(userData, currentItem);
                msg.Items.push(currentItem);
            }
            else {
                Logger.getInstance().warn(`Service manager did not return an item for ${itemEWSId}`);
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_ITEM_NOT_FOUND;
                msg.MessageText = `Item could not be found to be updated: ${itemEWSId}`;
            }
        } catch (error) {
            const err: any = error; 
            Logger.getInstance().warn(`Error updating item ${itemEWSId}: ${err}`)
            msg.ResponseClass = err.ResponseClass ?? ResponseClassType.ERROR;
            msg.ResponseCode = err.ResponseCode ?? ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR;
            msg.MessageText = err.MessageText ?? 'Error updating item';
        }
    }

    return response;
}

/**
 * Delete item.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns delete item result
 */
export async function DeleteItem(soapRequest: DeleteItemType, request: Request): Promise<DeleteItemResponseType> {
    const response = new DeleteItemResponseType();

    for (const itemIdType of soapRequest.ItemIds.items) {
        const msg = new DeleteItemResponseMessageType();
        response.ResponseMessages.push(msg);

        if (!(itemIdType instanceof ItemIdType)) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_INVALID_ID;
            msg.MessageText = `Item id type is not supported`;
            continue;
        }

        try {
            await doDeleteItem(request, itemIdType, undefined, soapRequest.DeleteType === DisposalType.HARD_DELETE);
        } catch (error) {
            const err: any = error; 
            if (err.status === 400 && err.message && err.message.includes("error code: 549")) {
                // item has already been deleted
                continue;
            }
            throwErrorIfClientResponse(err);
            if (err.status === 404) {
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_ITEM_NOT_FOUND;
                msg.MessageText = `Item id does not exist: ${itemIdType.Id}`;
                continue;
            } else {
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_CANNOT_DELETE_OBJECT;
                msg.MessageText = err.message;
                continue;
            }
        }
    }

    return response;
}

/**
 * Copy item.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns copy item result
 */
export async function CopyItem(soapRequest: CopyItemType, request: Request): Promise<CopyItemResponseType> {
    const userData = getUserMailData(request);
    const userInfo = UserContext.getUserInfo(request);

    const response = new CopyItemResponseType();
    const responseShapeType = new ItemResponseShapeType();
    // For copy, we fetch the existing item and create a new one based on it, so populate the existing item with ALL_PROPERTIES shape
    responseShapeType.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;

    let toLabel: PimLabel | undefined;
    try {
        toLabel = await getLabelByTarget(soapRequest.ToFolderId, userInfo);
    } catch (err) {
        Logger.getInstance().debug(`An error occurred getting all labels for CopyItem: ${err}`);
        throw err;
    }

    for (const itemIdType of soapRequest.ItemIds.items) {

        const msg = new CopyItemResponseMessage();
        response.ResponseMessages.push(msg);

        // We won't do the folder check when we remove the mock data.  Let Keep do the folder check
        if (!toLabel) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_TO_FOLDER_NOT_FOUND;
            continue;
        }

        if (!(itemIdType instanceof ItemIdType)) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_INVALID_ID;
            msg.MessageText = `Item id type is not supported`;
            continue;
        }

        // FIXME Hack for Apple Mail which will do a CopyItem and a DeleteItem to move an item to trash.  If the target folder is trash
        // Return the existing item id delete the item.  The subsequent DeleteItem will fail since we deleted it here.
        if (toLabel.view === KeepPimConstants.TRASH) {
            try {
                const result = await EWSServiceManager.getItem(itemIdType.Id, userInfo, request, responseShapeType);
                if (result) {
                    await doDeleteItem(request, itemIdType, result.ParentFolderId);
                    msg.Items.push(result);
                }
                else {
                    msg.ResponseClass = ResponseClassType.ERROR;
                    msg.ResponseCode = ResponseCodeType.ERROR_ITEM_NOT_FOUND;
                    msg.MessageText = `Item id does not exist: ${itemIdType.Id}`;
                }
            } catch (error) {
                const err: any = error; 
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_MOVE_COPY_FAILED;
                msg.MessageText = err.message ?? `Item ${itemIdType.Id} could not be copied to the target folder ${soapRequest.ToFolderId}`;
            }
            continue;
        }

        let pimItem: PimItem | undefined;

        const [itemIdBase, mailboxId] = getKeepIdPair(itemIdType.Id);

        try {
            if (itemIdBase) {
                pimItem = await KeepPimManager.getInstance().getPimItem(itemIdBase, userInfo, mailboxId);
            }
        } catch (error) {
            const err: any = error; 
            if (err.ResponseCode) {
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = err.ResponseCode;
                msg.MessageText = err.MessageText ?? `Error occurred getting item to copy ${itemIdType.Id}`;
                continue;
            } else {
                throwErrorIfClientResponse(err);
            }
        }

        if (!pimItem) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_ITEM_NOT_FOUND;
            msg.MessageText = `Item id does not exist: ${itemIdType.Id}`;
            continue;
        }

        let item: ItemType | undefined;
        // Create in new folder with new id
        try {
            const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
            if (manager) {
                item = await manager.copyItem(pimItem, toLabel.unid, userInfo, request, responseShapeType, mailboxId);
                if (item) {
                    // Update local cache
                    doMockDataCreateItem(userInfo, userData, [item]);
                }
            }
            else {
                throw new Error(`Unknown item type for copy item.`);
            }
        } catch (error) {
            const err: any = error; 
            Logger.getInstance().error(`Copy failed: ${err}`);
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_MOVE_COPY_FAILED;
            msg.MessageText = err.message ?? `New copy of ${itemIdType.Id} could not be created in the target folder ${util.inspect(soapRequest.ToFolderId, false, 5)}.`;
            continue;
        }

        if (!item) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_MOVE_COPY_FAILED;
            msg.MessageText = `New copy of ${itemIdType.Id} could not be created in the target folder ${util.inspect(soapRequest.ToFolderId, false, 5)}.`;
            continue;
        }

        msg.Items.push(item);
    }

    return response;
}

/**
 * Move item.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns move item result
 */
export async function MoveItem(soapRequest: MoveItemType, request: Request): Promise<MoveItemResponseType> {
    const userData = getUserMailData(request);
    const userInfo = UserContext.getUserInfo(request);

    const response = new MoveItemResponseType();
    const responseShapeType = new ItemResponseShapeType();
    responseShapeType.BaseShape = DefaultShapeNamesType.DEFAULT;

    let toLabel: PimLabel | undefined;
    try {
        toLabel = await getLabelByTarget(soapRequest.ToFolderId, userInfo);
    } catch (err) {
        Logger.getInstance().debug(`An error occurred getting all labels for MoveItem: ${err}`);
        throw err;
    }
    const toFolderMailboxId = findMailboxId(userInfo, soapRequest.ToFolderId);

    for (const itemIdType of soapRequest.ItemIds.items) {
        const msg = new MoveItemResponseMessage();
        response.ResponseMessages.push(msg);

        const itemMailboxId = findMailboxId(userInfo, itemIdType);
        if (itemMailboxId !== toFolderMailboxId) {
            throw new Error(`Item cannot be moved from ${itemMailboxId} to ${toFolderMailboxId} mailbox`);
        }

        if (!toLabel) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_TO_FOLDER_NOT_FOUND;
            continue;
        }

        if (!(itemIdType instanceof ItemIdType)) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_INVALID_ID;
            msg.MessageText = `Item id type is not supported`;
            continue;
        }

        const itemEWSId = itemIdType.Id;

        //FIXME:  This getItem can be removed when the Keep event code is hooked in.  Here we need the old parent folder to add events to.
        let beforeItem: ItemType | undefined;
        let oldFolder: FolderIdType | undefined;
        try {
            beforeItem = await EWSServiceManager.getItem(itemEWSId, userInfo, request, responseShapeType);
            if (beforeItem) {
                oldFolder = beforeItem.ParentFolderId;
                const afterItem = await EWSServiceManager.moveItemWithResult(beforeItem, toLabel.folderId, userInfo, request);
                Logger.getInstance().debug("MoveItem afterItem:" + util.inspect(afterItem, false, 5));
                if (afterItem) {
                    msg.Items.push(afterItem);
                    // Move the mock data
                    if (oldFolder) {
                        const toFolderEWSId = combineFolderIdAndMailbox(userInfo, toLabel.folderId, oldFolder);
                        const toFolderId = new FolderIdType(toFolderEWSId, `ck-${toFolderEWSId}`);
                        doMoveItem(userData, afterItem, oldFolder, toFolderId);
                    }
                } else {
                    msg.ResponseClass = ResponseClassType.ERROR;
                    msg.ResponseCode = ResponseCodeType.ERROR_MOVE_COPY_FAILED;
                    msg.MessageText = `Error: Could not retrieve moved item ${itemEWSId}`;
                }
            } else {
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_MOVE_COPY_FAILED;
                msg.MessageText = `Error: Could not retrieve item to move ${itemEWSId}`;
            }
        } catch (error) {
            const err: any = error; 
            if (err.ResponseCode) {
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = err.ResponseCode;
                msg.MessageText = err.MessageText ?? `Error occurred moving item ${itemEWSId}: ${err}`;
                continue;
            } else {
                throwErrorIfClientResponse(err);
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_MOVE_COPY_FAILED;
                msg.MessageText = `Error moving item, ${itemEWSId}: ${err}`;
                continue;
            }
        }
    }
    return response;
}

/**
 * Do move item.
 * @param userData The user mailbox data
 * @param item The item to move
 * @param folder The folder to move to
 */
function doMoveItem(userData: UserMailboxData, item: ItemType, oldFolderIdType: FolderIdType, newFolderIdType: FolderIdType): void {
    if (item.ItemId === undefined) {
        throw new Error(`Missing item id for ${util.inspect(item, false, 5)}`);
    }
    
    const itemId = item.ItemId.Id;

    // Delete from old folder
    const dchange = new SyncFolderItemsDeleteType();
    dchange.ItemId = deepcoy(item.ItemId);
    userData.FOLDER_ITEM_EVENTS[oldFolderIdType.Id] = userData.FOLDER_ITEM_EVENTS[oldFolderIdType.Id] ?? {};
    userData.FOLDER_ITEM_EVENTS[oldFolderIdType.Id][itemId] = new ItemEvent(Date.now(), dchange);

    // Create in new folder
    item.ParentFolderId = newFolderIdType;

    const newFolderId = item.ParentFolderId.Id;

    const change = new SyncFolderItemsCreateType();
    populateSyncItemChange(change, item);
    userData.FOLDER_ITEM_EVENTS[newFolderId] = userData.FOLDER_ITEM_EVENTS[newFolderId] ?? {};
    userData.FOLDER_ITEM_EVENTS[newFolderId][itemId] = new ItemEvent(Date.now(), change);

    const notifyEvent = new MovedEvent();
    notifyEvent.ItemId = item.ItemId;
    notifyEvent.ParentFolderId = item.ParentFolderId;
    notifyEvent.OldItemId = item.ItemId;
    notifyEvent.OldParentFolderId = oldFolderIdType;
    notifyEvent.TimeStamp = new Date();
    sendStreamingEvent(userData.userId, notifyEvent);
}

/**
 * Archive item.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns archive item result
 */
export async function ArchiveItem(soapRequest: ArchiveItemType, request: Request): Promise<ArchiveItemResponseType> {
    const userData = getUserMailData(request);
    const userInfo = UserContext.getUserInfo(request);

    const response = new ArchiveItemResponseType();
    const responseShapeType = new ItemResponseShapeType();
    responseShapeType.BaseShape = DefaultShapeNamesType.DEFAULT;

    let toLabel: PimLabel | undefined;
    
    try {
        const labels = await KeepPimLabelManager.getInstance().getLabels(userInfo);
        // TODO: Change to DistinguishedFolderIdType when LABS-2565 is complete
        const mailboxId = findMailboxId(userInfo, soapRequest.ItemIds.items[0]);
        toLabel = findLabel(labels, new FolderIdType(getEWSId(DistinguishedFolderIdNameType.ARCHIVE, mailboxId)));
    } catch (err) {
        Logger.getInstance().debug(`An error occurred getting all labels for ArchiveItem: ${err}`);
        throw err;
    }

    for (const itemIdType of soapRequest.ItemIds.items) {

        const msg = new ArchiveItemResponseMessage();
        response.ResponseMessages.push(msg);

        if (!toLabel) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_ARCHIVE_MAILBOX_NOT_ENABLED;
            continue;
        }

        if (!(itemIdType instanceof ItemIdType)) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_INVALID_ID;
            msg.MessageText = `Item id type is not supported`;
            continue;
        }

        const itemEWSId = itemIdType.Id;

        //FIXME:  This getItem can be removed when the Keep event code is hooked in.  Here we need the old parent folder to add events to.
        let beforeItem: ItemType | undefined;
        let oldFolder: FolderIdType | undefined;
        try {
            beforeItem = await EWSServiceManager.getItem(itemEWSId, userInfo, request, responseShapeType);
            if (beforeItem) {
                oldFolder = beforeItem.ParentFolderId;
                const afterItem = await EWSServiceManager.moveItemWithResult(beforeItem, toLabel.folderId, userInfo, request);
                Logger.getInstance().debug("ArchiveItem afterItem:" + util.inspect(afterItem, false, 5));
                if (afterItem) {
                    msg.Items.push(afterItem);
                    // Move the mock data
                    if (oldFolder) {
                        const toFolderEWSId = combineFolderIdAndMailbox(userInfo, toLabel.folderId, oldFolder);
                        const toFolderId = new FolderIdType(toFolderEWSId, `ck-${toFolderEWSId}`);
                        doMoveItem(userData, afterItem, oldFolder, toFolderId);
                    }
                } else {
                    msg.ResponseClass = ResponseClassType.ERROR;
                    msg.ResponseCode = ResponseCodeType.ERROR_MOVE_COPY_FAILED;
                    msg.MessageText = `Error: Could not retrieve archived item ${itemEWSId}`;
                }
            } else {
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_MOVE_COPY_FAILED;
                msg.MessageText = `Error: Could not retrieve item to archive ${itemEWSId}`;
            }
        } catch (error) {
            const err: any = error; 
            Logger.getInstance().debug(`ArchiveItem error archiving item ${itemEWSId}: ${err}`);
            if (err.ResponseCode) {
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = err.ResponseCode;
                msg.MessageText = err.MessageText ?? `Error occurred archiving item ${itemEWSId}: ${err}`;
                continue;
            } else {
                throwErrorIfClientResponse(err);
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_MOVE_COPY_FAILED;
                msg.MessageText = `Error moving item, ${itemEWSId}: ${err}`;
                continue;
            }
        }
    }
    return response;
}

/**
 * Mark as junk.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns mark as junk result
 */
export async function MarkAsJunk(soapRequest: MarkAsJunkType, request: Request): Promise<MarkAsJunkResponseType> {

    const userData = getUserMailData(request);
    const userInfo = UserContext.getUserInfo(request);

    const response = new MarkAsJunkResponseType();
    const responseShapeType = new ItemResponseShapeType();
    responseShapeType.BaseShape = DefaultShapeNamesType.DEFAULT;

    let toLabel: PimLabel | undefined;
    try {
        const labels = await KeepPimLabelManager.getInstance().getLabels(userInfo);
        const folderId = new DistinguishedFolderIdType(); 
        folderId.Id = soapRequest.IsJunk ? DistinguishedFolderIdNameType.JUNKEMAIL : DistinguishedFolderIdNameType.INBOX; 
        toLabel = findLabel(labels, folderId);
    } catch (err) {
        Logger.getInstance().debug(`An error occurred getting all labels for MarkAsJunk: ${err}`);
        throw err;
    }

    for (const itemIdType of soapRequest.ItemIds.items) {
        const msg = new MarkAsJunkResponseMessageType();
        response.ResponseMessages.push(msg);

        if (!soapRequest.MoveItem) {
            continue;
        }

        if (!toLabel) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_FOLDER_NOT_FOUND;
            continue;
        }

        if (!(itemIdType instanceof ItemIdType)) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_INVALID_ID;
            msg.MessageText = `Item id type is not supported`;
            continue;
        }

        const itemEWSId = itemIdType.Id;

        //FIXME:  This getItem can be removed when the Keep event code is hooked in.  Here we need the old parent folder to add events to.
        let beforeItem: ItemType | undefined;
        let oldFolder: FolderIdType | undefined;
        try {
            beforeItem = await EWSServiceManager.getItem(itemEWSId, userInfo, request, responseShapeType);
            if (beforeItem) {
                oldFolder = beforeItem.ParentFolderId;
                const afterItem = await EWSServiceManager.moveItemWithResult(beforeItem, toLabel.folderId, userInfo, request);
                Logger.getInstance().debug("MarkAsJunk afterItem:" + util.inspect(afterItem, false, 5));
                if (afterItem) {
                    msg.MovedItemId = afterItem.ItemId;
                    // Move the mock data
                    if (oldFolder) {
                        //Get EWS toFolder Id
                        const toFolderEWSId = combineFolderIdAndMailbox(userInfo, toLabel.folderId, oldFolder);
                        const toFolderId = new FolderIdType(toFolderEWSId, `ck-${toFolderEWSId}`);
                        doMoveItem(userData, afterItem, oldFolder, toFolderId);
                    }
                } else {
                    msg.ResponseClass = ResponseClassType.ERROR;
                    msg.ResponseCode = ResponseCodeType.ERROR_MOVE_COPY_FAILED;
                    msg.MessageText = `Error: Could not retrieve item marked as junk ${itemEWSId}`;
                }
            } else {
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_MOVE_COPY_FAILED;
                msg.MessageText = `Error: Could not retrieve item to mark as junk ${itemEWSId}`;
            }
        } catch (error) {
            const err: any = error; 
            Logger.getInstance().debug("MarkAsJunk error moving item to junk email:" + itemEWSId);
            if (err.ResponseCode) {
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = err.ResponseCode;
                msg.MessageText = err.MessageText ?? `Error occurred moving item to junk email: ${itemEWSId}`;
                continue;
            } else {
                throwErrorIfClientResponse(err);
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_MOVE_COPY_FAILED;
                msg.MessageText = `Error moving item, ${itemEWSId} to junk email: ${err}`;
                continue;
            }
        }
    }
    return response;

}

/**
 * Mark all items as read.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns mark as read result
 */
export async function MarkAllItemsAsRead(soapRequest: MarkAllItemsAsReadType, request: Request): Promise<MarkAllItemsAsReadResponseType> {
    const userData = getUserMailData(request);
    const userInfo = UserContext.getUserInfo(request);

    const response = new MarkAllItemsAsReadResponseType();

    try {
        const labels = await KeepPimLabelManager.getInstance().getLabels(userInfo);

        for (const folderIdType of soapRequest.FolderIds.items) {
            const msg = new MarkAllItemsAsReadResponseMessage();
            response.ResponseMessages.push(msg);

            const label = findLabel(labels, folderIdType);
            if (!label) {
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_FOLDER_NOT_FOUND;
                msg.MessageXml = { Value: new Value("FolderId", util.inspect(folderIdType, false, 5)) }
                continue;
            }

            try {
                // TODO: Need API to mark items as read (LABS-1463)
                const messages = await KeepPimMessageManager.getInstance().getMailMessages(userInfo, label.view);
                for (const message of messages) {
                    const item = new MessageType();
                    item.ItemId = new ItemIdType(message.unid, `ck-${message.unid}`);
                    item.ParentFolderId = new FolderIdType(label.folderId, `ck-${label.folderId}`);
                    item.IsRead = soapRequest.ReadFlag;
                    doUpdateItem(userData, item);
                }
            } catch (err) {
                const messageText = `Error getting messages for folder ${label.folderId}: ${err}`;
                Logger.getInstance().error(messageText);
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR;
                msg.MessageText = messageText;
            }
        }

        return response;

    } catch (err) {
        Logger.getInstance().error(`Error getting labels: ${err}`)
        throw err;
    }
}

/**
 * Report message.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns report result
 */
export async function ReportMessage(soapRequest: ReportMessageType, request: Request): Promise<ReportMessageResponseType> {
    const userInfo = UserContext.getUserInfo(request);

    const response = new ReportMessageResponseType();

    for (const itemIdType of soapRequest.ItemIds.items) {
        const msg = new ReportMessageResponseMessageType();
        response.ResponseMessages.push(msg);

        if (!(itemIdType instanceof ItemIdType)) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_INVALID_ID;
            msg.MessageText = `Item id type is not supported`;
            continue;
        }

        const responseShapeType = new ItemResponseShapeType();
        responseShapeType.BaseShape = DefaultShapeNamesType.ID_ONLY;

        let item: ItemType | undefined;
        const itemEWSId = itemIdType.Id;

        try {
            item = await EWSServiceManager.getItem(itemEWSId, userInfo, request, responseShapeType);
        } catch (error) {
            const err: any = error; 
            if (err.ResponseCode) {
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = err.ResponseCode;
                msg.MessageText = err.MessageText ?? `ReportMessage: Error occurred getting item ${itemEWSId}`;
                continue;
            } else {
                throwErrorIfClientResponse(err);
            }
        }

        if (!item) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_ITEM_NOT_FOUND;
            msg.MessageText = `Item id does not exist: ${itemEWSId}`;
            continue;
        }
        msg.MovedItemId = itemIdType;
    }
    return response;
}

/**
 * Create attachment.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns create attachment result
 */
export async function CreateAttachment(soapRequest: CreateAttachmentType, request: Request): Promise<CreateAttachmentResponseType> {
    const userData = getUserMailData(request);
    const userInfo = UserContext.getUserInfo(request);

    const response = new CreateAttachmentResponseType();

    let itemUpdated = false;

    const responseShapeType = new ItemResponseShapeType();
    responseShapeType.BaseShape = DefaultShapeNamesType.ID_ONLY;
    let itemEWSId: string | undefined;
    let itemId: string | undefined;
    let item: ItemType | undefined;
    if (soapRequest.ParentItemId) {
        itemEWSId = soapRequest.ParentItemId.Id;
        [itemId] = getKeepIdPair(itemEWSId);
        try {
            item = await EWSServiceManager.getItem(itemEWSId, userInfo, request, responseShapeType);
        } catch (error) {
            const err: any = error; 
            if (err.ResponseCode !== ResponseCodeType.ERROR_ITEM_NOT_FOUND) {
                throw (err);
            }
        }
    }
    for (const att of soapRequest.Attachments.items) {
        Logger.getInstance().debug(`Request to add attachment ${att.Name ?? att.ContentId} to item ${itemId}`);

        const msg = new CreateAttachmentResponseMessage();
        response.ResponseMessages.push(msg);

        if (!item) {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_ITEM_NOT_FOUND;
            msg.MessageText = `Item id does not exist: ${itemEWSId}`;
            continue;
        }

        // FIXME:  There is no api for create attachemnt in Keep LABS-1075
        // NOTE: For contact photos, Domino using the name ContactPhoto and Exchange uses ContactPicture.jpg. The EWS client will send a CreateAttachment with the name ContactPicture.jpg 
        // and isContactPhoto = true. When saving this to Keep, the name should be changed to ContactPhoto.         

        // Create the attachment.
        if (att instanceof FileAttachmentType) {
            if (!att.Content || !att.Name || !itemId) {
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_INVALID_ITEM_FOR_OPERATION_CREATE_ITEM_ATTACHMENT;
                msg.MessageText = `Attachment does not have name, id or content: ${att.Name} / ${itemId} / ${att.Content}`;
                continue;
            }
            try {
                
                await KeepPimAttachmentManager.getInstance().createAttachment(userInfo, itemId, att.Content, att.Name, att.ContentType);
            } catch (error) {
                const err: any = error; 
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_INVALID_ITEM_FOR_OPERATION_CREATE_ITEM_ATTACHMENT;
                msg.MessageText = err.message ?? `An error occurred creating the attachment.`;
                continue;
            }

            // If successful create, retrieve the attachment for the response.
            try {
                const attachment = await getItemAttachment(request, userData, itemId, att.Name, false);
                msg.Attachments.push(attachment);
                item.HasAttachments = true;
                item.Attachments = item.Attachments ?? new NonEmptyArrayOfAttachmentsType();
                item.Attachments.push(att);
                itemUpdated = true;
            } catch (error) {
                const gErr: any = error; 
                if (gErr.ResponseClass && gErr.ResponseCode) {
                    msg.ResponseClass = gErr.ResponseClass;
                    msg.ResponseCode = gErr.ResponseCode;
                    msg.MessageText = gErr.MessageText ?? `An error occurred retrieving the created attachment.`;
                    continue;
                } else {
                    throw gErr;
                }
            }
        } else {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_INVALID_ITEM_FOR_OPERATION_CREATE_ITEM_ATTACHMENT;
            msg.MessageText = `Only File Attachments are supported at this time: ${itemEWSId}`;
        }
    }

    if (item?.ItemId !== undefined) {
        if (itemUpdated) {
            doUpdateItem(userData, item);
        }
        for (const msg of response.ResponseMessages.items) {
            if (msg.Attachments && msg.Attachments.items.length > 0) {
                for (const att of msg.Attachments.items) {
                    if (att.AttachmentId === undefined) {
                        att.AttachmentId = new AttachmentIdType();
                    }
                    att.AttachmentId.RootItemChangeKey = item.ItemId.ChangeKey;
                }
            } 
        }
    }

    return response;
}

/**
 * Get attachment.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns get attachment result
 */
export async function GetAttachment(soapRequest: GetAttachmentType, request: Request): Promise<GetAttachmentResponseType> {
    const userData = getUserMailData(request);

    const response = new GetAttachmentResponseType();
    const includeMimeContent = soapRequest.AttachmentShape?.IncludeMimeContent;

    for (const attId of soapRequest.AttachmentIds.AttachmentId) {
        const msg = new GetAttachmentResponseMessage();
        response.ResponseMessages.push(msg);

        
        let attachment: AttachmentType | undefined;
        try {
            const [parentId, attachmentName] = getAttachmentIdPair(attId.Id);
            if (!parentId) {
                Logger.getInstance().debug(`Error occurred retrieving the attachment from the server for attachment ${attId.Id}:  Id of parent item is required`);
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_CANNOT_FIND_FILE_ATTACHMENT;
                msg.MessageText = `Error occurred retrieving the attachment from the server for attachment ${attId.Id}:  Id of parent item is required`;
                continue;
            } else if (!attachmentName) {
                Logger.getInstance().debug(`Error occurred retrieving the attachment from the server for attachment ${attId.Id}:  attachment name is required`);
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_CANNOT_FIND_FILE_ATTACHMENT;
                msg.MessageText = `Error occurred retrieving the attachment from the server for attachment ${attId.Id}:  attachment name is required`;
                continue;
            } else {
                try {
                    attachment = await getItemAttachment(request, userData, parentId, attachmentName, includeMimeContent ?? true);
                    msg.Attachments.push(attachment);
                } catch (err) {
                    const gErr: any = err; 
                    if (gErr.ResponseClass && gErr.ResponseCode) {
                        msg.ResponseClass = gErr.ResponseClass;
                        msg.ResponseCode = gErr.ResponseCode;
                        msg.MessageText = gErr.MessageText ?? `An error occurred retrieving the created attachment.`;
                        continue;
                    } else {
                        throw gErr;
                    }
                }
            }
        } catch (error) {
            const err: any = error; 
            Logger.getInstance().debug(`Error occurred retrieving the attachment from the server for attachment ${attId.Id}: ` + err);
            throwErrorIfClientResponse(err);

            msg.ResponseClass = ResponseClassType.ERROR;
            msg.MessageText = `Error occurred retrieving the attachment from the server for attachment ${attId.Id}:  ${err}`;
            if (err.status && err.status === 404) {
                msg.ResponseCode = ResponseCodeType.ERROR_CANNOT_FIND_FILE_ATTACHMENT;

            } else {
                msg.ResponseCode = ResponseCodeType.ERROR_CANNOT_OPEN_FILE_ATTACHMENT;
            }
            continue;
        }
    }
    return response;
}


async function getItemAttachment(request: Request, userData: UserMailboxData, parentId: string, attachmentName: string, includeMimeContent: boolean): Promise<ItemAttachmentType> {
    let attachment: AttachmentType | undefined;
    //We're currently assuming all attachments are File Attachments.  No ItemAttachment
    const userInfo = UserContext.getUserInfo(request);
    const attInfo = await KeepPimAttachmentManager.getInstance().getAttachment(userInfo, parentId, attachmentName);
    if (attInfo) {
        // Set lastModifiedTime
        // Set Size
        attachment = new FileAttachmentType();
        attachment.AttachmentId = new AttachmentIdType();
        // Set attachmentId
        attachment.AttachmentId.Id = getAttachmentId(parentId, attachmentName);
        // Set attachment name
        attachment.Name = attachmentName;
        // Set attachment RootItemId
        attachment.AttachmentId.RootItemId = parentId;
        // Set Content
        const attachmentString = attInfo.getBase64AttachmentData();
        if (!attachmentString) {
            Logger.getInstance().debug(`Error occurred retrieving the attachment from the server for attachment ${parentId}, attachment name ${attachmentName}`);
            const err: any = new Error();
            err.ResponseClass = ResponseClassType.ERROR;
            err.ResponseCode = ResponseCodeType.ERROR_CANNOT_OPEN_FILE_ATTACHMENT;
            err.MessageText = `Error occurred retrieving the attachment from the server for attachment ${parentId}, attachment name ${attachmentName}`;
            throw err;
        } else {
            attachment.ContentId = getAttachmentId(parentId, attachmentName);
            attachment.ContentType = attInfo.contentType;
            if (attachment instanceof FileAttachmentType) {
                attachment.Content = attachmentString;
            } else if (attachment instanceof ItemAttachmentType) {
                // FIXME: Currently this is dead code since we're assuming all attachments are FileAtttachmentType above
                // Can domino have a Message, Calendar, Contact, Task as an attachment?  What if Outlook or Apple Mail client
                // allow them to be created?  Will they be just file attachments in domino?
                let msgType = new MessageType();
                msgType.ItemId = new ItemIdType(attachmentName, `ck-${attachmentName}`);
                msgType = await populateItemFromMime(userData, msgType, attInfo.getBase64AttachmentData());
                if (includeMimeContent) {
                    msgType.MimeContent = new MimeContentType(attachmentString, 'UTF-8');
                }
                attachment.Message = msgType;
            }
        }
    }
    return new Promise<AttachmentType>((resolve, reject) => {
        if (attachment) {
            resolve(attachment);
        } else {
            reject(new Error(`Attachment ${attachmentName} could not be found.`));
        }
    });
}

/**
 * Delete attachment.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns delete attachment result
 */
export async function DeleteAttachment(soapRequest: DeleteAttachmentType, request: Request): Promise<DeleteAttachmentResponseType> {
    const userData = getUserMailData(request);
    const userInfo = UserContext.getUserInfo(request);

    const response = new DeleteAttachmentResponseType();

    for (const attId of soapRequest.AttachmentIds.AttachmentId) {
        const msg = new DeleteAttachmentResponseMessageType();
        response.ResponseMessages.push(msg);

        let parentId: string | undefined;
        let attachmentName: string | undefined;
        try {
            [parentId, attachmentName] = getAttachmentIdPair(attId.Id);
            if (!parentId) {
                Logger.getInstance().debug(`Error occurred deleting the attachment from the server for attachment ${attId.Id}:  Id of parent item is required`);
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_CANNOT_FIND_FILE_ATTACHMENT;
                msg.MessageText = `Error occurred deleting the attachment from the server for attachment ${attId.Id}:  Id of parent item is required`;
                continue;
            } else if (!attachmentName) {
                Logger.getInstance().debug(`Error occurred deleting the attachment from the server for attachment ${attId.Id}:  attachment name is required`);
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_CANNOT_FIND_FILE_ATTACHMENT;
                msg.MessageText = `Error occurred deleting the attachment from the server for attachment ${attId.Id}:  attachment name is required`;
                continue;
            } else {
                await KeepPimAttachmentManager.getInstance().deleteAttachment(userInfo, parentId, attachmentName);
            }
        } catch (err) {
            // Error occurred deleting the attachment from the server...fallback to the mock data only
            Logger.getInstance().debug(`Error occurred deleting the attachment from the server for attachment ${attId.Id}: ${err}`);
            throwErrorIfClientResponse(err);

            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_CANNOT_FIND_FILE_ATTACHMENT;
            msg.MessageText = `Error occurred deleting the attachment from the server for attachment ${attId.Id}:  attachment name is required`;
            continue;
        }

        //FIXME:  We shouldn't need to get the item to call doUpdateItem if server eventing is in place.  This block should go away and we get the event
        // from the server that the item was modified
        if (parentId) {
            const responseShapeType = new ItemResponseShapeType();
            responseShapeType.BaseShape = DefaultShapeNamesType.ID_ONLY;
            try {
                const parentEWSId = getEWSId(parentId);
                const item = await EWSServiceManager.getItem(parentEWSId, userInfo, request, responseShapeType);
                if (item) {
                    doUpdateItem(userData, item);
                    msg.RootItemId = new RootItemIdType();
                    msg.RootItemId.RootItemId = parentEWSId;
                    msg.RootItemId.RootItemChangeKey = item.ItemId?.ChangeKey ?? `ck-${parentEWSId}`;
                }
            } catch (error) {
                const err: any = error; 
                if (err.ResponseCode) {
                    msg.ResponseClass = ResponseClassType.ERROR;
                    msg.ResponseCode = err.ResponseCode;
                    msg.MessageText = err.MessageText ?? `DeleteAttachment: Error occurred getting item ${parentId}`;
                    continue;
                } else {
                    throw err;
                }
            }
        }
    }

    return response;
}

/**
 * Send streaming notification event to user's subscription.
 * @param userId The user id
 * @param event The event to send
 */
function sendStreamingEvent(userId: string, event: BaseNotificationEventType): void {
    if (STREAMING_SUBSCRIPTIONS[userId]) {
        for (const sub of Object.values(STREAMING_SUBSCRIPTIONS[userId])) {
            sub.sendEvent(event);
        }
    }
}

/**
 * Streaming subscription.
 */
class StreamingSubscription {

    // Whether the subscription has been started.
    started: boolean;

    /**
     * Constructor.
     * @param emitter The event emitter
     * @param request The subscription request
     */
    constructor(public emitter: EventEmitter, public request: StreamingSubscriptionRequestType) { }

    /**
     * Send notification event.
     * @param event The event to send
     */
    sendEvent(event: BaseNotificationEventType): void {
        if (!this.started) {
            return;
        }
        const found = this.request.EventTypes.EventType.find(et => et === event.constructor.name)
        if (found) {
            this.emitter.emit('notification', event);
        }
    }
}
const STREAMING_SUBSCRIPTIONS: { [userId: string]: { [subscriptionId: string]: StreamingSubscription } } = {};

/**
 * Subscribe.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns subscribe result
 */
export function Subscribe(soapRequest: SubscribeType, request: Request): SubscribeResponseType {
    const userData = getUserMailData(request);
    const userId = userData.userId;

    const response = new SubscribeResponseType();
    const msg = new SubscribeResponseMessageType();
    response.ResponseMessages.push(msg);
    if (soapRequest.StreamingSubscriptionRequest) {
        STREAMING_SUBSCRIPTIONS[userId] = STREAMING_SUBSCRIPTIONS[userId] ?? {};
        const subscriptionId = msg.SubscriptionId = uuidv4();
        STREAMING_SUBSCRIPTIONS[userId][subscriptionId]
            = new StreamingSubscription(new EventEmitter(), soapRequest.StreamingSubscriptionRequest);
    }
    return response;
}

/**
 * Get streaming result. Will hold a long connection.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns get streaming result
 */
export async function GetStreamingEvents(soapRequest: GetStreamingEventsType, request: Request, callback: Function): Promise<GetStreamingEventsResponseType> {
    const userData = getUserMailData(request);
    const userId = userData.userId;

    const result = new GetStreamingEventsResponseType();
    const msg = new GetStreamingEventsResponseMessageType();
    result.ResponseMessages.push(msg);

    if (!soapRequest.SubscriptionIds || !soapRequest.SubscriptionIds.SubscriptionId || !soapRequest.SubscriptionIds.SubscriptionId.length) {
        msg.ResponseClass = ResponseClassType.ERROR;
        msg.ResponseCode = ResponseCodeType.ERROR_INVALID_SUBSCRIPTION_REQUEST;
        return result;
    }

    const subscriptionId = soapRequest.SubscriptionIds.SubscriptionId[0];
    if (!STREAMING_SUBSCRIPTIONS[userId] || !STREAMING_SUBSCRIPTIONS[userId][subscriptionId]) {
        msg.ResponseClass = ResponseClassType.ERROR;
        msg.ResponseCode = ResponseCodeType.ERROR_SUBSCRIPTION_NOT_FOUND;
        return result;
    }

    const subscription = STREAMING_SUBSCRIPTIONS[userId][subscriptionId];
    if (subscription.started) {
        msg.ResponseClass = ResponseClassType.ERROR;
        msg.ResponseCode = ResponseCodeType.ERROR_INVALID_SUBSCRIPTION;
        return result;
    }

    subscription.started = true;
    return new Promise<GetStreamingEventsResponseType>((resolve) => {

        const listener = (event: BaseNotificationEventType): void => {
            const sresult = new GetStreamingEventsResponseType();
            const smsg = new GetStreamingEventsResponseMessageType();
            sresult.ResponseMessages.push(smsg);

            smsg.Notifications = new NonEmptyArrayOfNotificationsType();
            const noti = new NotificationType();
            noti.SubscriptionId = subscriptionId;
            noti.items.push(event);
            smsg.Notifications.Notification.push(noti);

            sendNoti(sresult);
        };
        subscription.emitter.on('notification', listener);

        const sendNoti = (toSend: GetStreamingEventsResponseType): void => {
            setImmediate(() => {
                try {
                    callback(null, toSend, true);
                } catch (error) {
                    const err: any = error; 
                    Logger.getInstance().error(err);
                    cleanup(err);
                }
            })
        }

        let cleaned = false;
        const cleanup = (err?: Error): void => {
            if (cleaned) {
                return;
            }
            Logger.getInstance().debug('Cleanup streaming connection');
            cleaned = true;
            clearInterval(intervalHandle);
            subscription.emitter.off('notification', listener);
            if (err) {
                // Probably connection closed error, client will order new subscription
                delete STREAMING_SUBSCRIPTIONS[userId][subscriptionId];
            } else {
                // Normal cleanup, the subscription will be reused
                subscription.started = false;
            }

            const sresult = new GetStreamingEventsResponseType();
            const smsg = new GetStreamingEventsResponseMessageType();
            sresult.ResponseMessages.push(smsg);
            smsg.ConnectionStatus = ConnectionStatusType.CLOSED;
            resolve(sresult);
        }

        const sendOK = (): void => {
            const sresult = new GetStreamingEventsResponseType();
            const smsg = new GetStreamingEventsResponseMessageType();
            sresult.ResponseMessages.push(smsg);
            smsg.ConnectionStatus = ConnectionStatusType.OK;
            sendNoti(sresult);
        }

        const connectionTimeout = soapRequest.ConnectionTimeout * 60 * 1000;
        const startTime = Date.now();
        const intervalHandle = setInterval(() => {
            if (Date.now() - startTime >= connectionTimeout) {
                cleanup();
            } else {
                sendOK();
            }
        }, config.keepAliveInterval);

        sendOK();
    });
}

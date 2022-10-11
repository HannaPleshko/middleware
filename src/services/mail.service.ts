/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import assert from 'assert';
import { Request } from '@loopback/rest';
import { BindingScope, bind } from '@loopback/core';

import { ExchangeSoap } from '../soap';
import {
    GetFolderType, GetFolderResponseType, SyncFolderHierarchyType, SyncFolderHierarchyResponseType,
    MoveItemResponseType, CopyItemType, CopyItemResponseType, ArchiveItemType, ArchiveItemResponseType,
    CreateItemResponseType, SendItemResponseType, SendItemType, DeleteFolderType, DeleteFolderResponseType,
    UploadItemsResponseType, UploadItemsType, GetItemResponseType, GetItemType, CreateItemType, SubscribeType,
    UpdateItemType, UpdateItemResponseType, DeleteItemType, DeleteItemResponseType, MoveItemType, MarkAsJunkType,
    MarkAsJunkResponseType, MarkAllItemsAsReadType, MarkAllItemsAsReadResponseType, CreateAttachmentResponseType,
    CreateFolderType, CreateFolderResponseType, CreateFolderPathType, CreateFolderPathResponseType, ConvertIdType,
    SubscribeResponseType, GetAppMarketplaceUrlType, GetAppMarketplaceUrlResponseMessageType, GetStreamingEventsType,
    GetStreamingEventsResponseType, FindFolderType, FindFolderResponseType, CreateAttachmentType, DeleteAttachmentType,
    UpdateFolderType, UpdateFolderResponseType, SyncFolderItemsType, SyncFolderItemsResponseType, ConvertIdResponseType,
    ReportMessageType, ReportMessageResponseType, GetAttachmentType, GetAttachmentResponseType, DeleteAttachmentResponseType,
    MoveFolderType, MoveFolderResponseType, CopyFolderType, CopyFolderResponseType, FindItemType, FindItemResponseType,
    CreateManagedFolderRequestType, CreateManagedFolderResponseType, EmptyFolderType, EmptyFolderResponseType,
    FindConversationType, FindConversationResponseMessageType, GetConversationItemsType, GetConversationItemsResponseType,
    ApplyConversationActionType, ApplyConversationActionResponseType, ExpandDLType, ExpandDLResponseType, 
    GetPasswordExpirationDateType, GetPasswordExpirationDateResponseMessageType
} from '../models/mail.model';

import {
    GetFolder, CreateFolder, UpdateItem, DeleteItem, DeleteFolder, UpdateFolder, SyncFolderItems, GetItem, CreateItem, CopyFolder,
    SyncFolderHierarchy, MoveItem, CopyItem, ArchiveItem, MarkAsJunk, MarkAllItemsAsRead, GetAttachment, GetStreamingEvents, MoveFolder,
    Subscribe, FindFolder, CreateAttachment, DeleteAttachment, UploadItems, SendItem, CreateFolderPath, ConvertId, ReportMessage, 
    GetAppMarketplaceUrl, FindItem, CreateManagedFolder, EmptyFolder, FindConversation, GetConversationItems, ApplyConversationAction, ExpandDL, 
    GetPasswordExpirationDate
} from '../data/mail';

/**
 * The mail service.
 */
@bind({ scope: BindingScope.SINGLETON, tags: ExchangeSoap.service })
export class MailService {

    /**
     * Get folders.
     * @param soapRequest The soap request, which is type of GetFolderType
     * @param request The HTTP request
     * @returns folders
     */
    async GetFolder(soapRequest: GetFolderType, request: Request): Promise<GetFolderResponseType> {
        assert(soapRequest instanceof GetFolderType);

        return GetFolder(soapRequest, request);
    }

    /**
     * Find folders.
     * @param soapRequest The soap request, which is type of FindFolderType
     * @param request The HTTP request
     * @returns folders
     */
    async FindFolder(soapRequest: FindFolderType, request: Request): Promise<FindFolderResponseType> {
        assert(soapRequest instanceof FindFolderType);

        return FindFolder(soapRequest, request);
    }

    /**
     * Sync folder hierarchy.
     * @param soapRequest The soap request, which is type of SyncFolderHierarchyType
     * @param request The HTTP request
     * @returns folder hierarchy changes to sync
     */
    async SyncFolderHierarchy(soapRequest: SyncFolderHierarchyType, request: Request): Promise<SyncFolderHierarchyResponseType> {
        assert(soapRequest instanceof SyncFolderHierarchyType);

        return SyncFolderHierarchy(soapRequest, request);
    }

    /**
     * Create folder.
     * @param soapRequest The soap request, which is type of CreateFolderType
     * @param request The HTTP request
     * @returns folder created
     */
    async CreateFolder(soapRequest: CreateFolderType, request: Request): Promise<CreateFolderResponseType> {
        assert(soapRequest instanceof CreateFolderType);

        return CreateFolder(soapRequest, request);
    }

    /**
     * Update folder.
     * @param soapRequest The soap request, which is type of UpdateFolderType
     * @param request The HTTP request
     * @returns folder updated
     */
    async UpdateFolder(soapRequest: UpdateFolderType, request: Request): Promise<UpdateFolderResponseType> {
        assert(soapRequest instanceof UpdateFolderType);

        return UpdateFolder(soapRequest, request);
    }

    /**
     * Move folder.
     * @param soapRequest The soap request, which is type of MoveFolderType
     * @param request The HTTP request
     * @returns folder move result
     */
    async MoveFolder(soapRequest: MoveFolderType, request: Request): Promise<MoveFolderResponseType> {
        assert(soapRequest instanceof MoveFolderType);

        return MoveFolder(soapRequest, request);
    }

    /**
     * Copy folder.
     * @param soapRequest The soap request, which is type of CopyFolderType
     * @param request The HTTP request
     * @returns folder move result
     */
    async CopyFolder(soapRequest: CopyFolderType, request: Request): Promise<CopyFolderResponseType> {
        assert(soapRequest instanceof CopyFolderType);

        return CopyFolder(soapRequest, request);
    }

    /**
     * Delete folder.
     * @param soapRequest The soap request, which is type of DeleteFolderType
     * @param request The HTTP request
     * @returns delete result
     */
    async DeleteFolder(soapRequest: DeleteFolderType, request: Request): Promise<DeleteFolderResponseType> {
        assert(soapRequest instanceof DeleteFolderType);

        return DeleteFolder(soapRequest, request);
    }

    /**
     * Sync folder items.
     * @param soapRequest The soap request, which is type of SyncFolderItemsType
     * @param request The HTTP request
     * @returns folder items changes to sync
     */
    async SyncFolderItems(soapRequest: SyncFolderItemsType, request: Request): Promise<SyncFolderItemsResponseType> {
        assert(soapRequest instanceof SyncFolderItemsType);

        return SyncFolderItems(soapRequest, request);
    }

    /**
     * Get item.
     * @param soapRequest The soap request, which is type of GetItemType
     * @param request The HTTP request
     * @returns item retrieved
     */
    async GetItem(soapRequest: GetItemType, request: Request): Promise<GetItemResponseType> {
        assert(soapRequest instanceof GetItemType);

        return GetItem(soapRequest, request);
    }

    /**
     * Create item.
     * @param soapRequest The soap request, which is type of CreateItemType
     * @param request The HTTP request
     * @returns item created
     */
    async CreateItem(soapRequest: CreateItemType, request: Request): Promise<CreateItemResponseType> {
        assert(soapRequest instanceof CreateItemType);

        return CreateItem(soapRequest, request);
    }

    /**
     * Update item.
     * @param soapRequest The soap request, which is type of UpdateItemType
     * @param request The HTTP request
     * @returns item update result
     */
    async UpdateItem(soapRequest: UpdateItemType, request: Request): Promise<UpdateItemResponseType> {
        assert(soapRequest instanceof UpdateItemType);

        return UpdateItem(soapRequest, request);
    }

    /**
     * Copy item.
     * @param soapRequest The soap request, which is type of CopyItemType
     * @param request The HTTP request
     * @returns item copy result
     */
    async CopyItem(soapRequest: CopyItemType, request: Request): Promise<CopyItemResponseType> {
        assert(soapRequest instanceof CopyItemType);

        return CopyItem(soapRequest, request);
    }

    /**
     * Move item.
     * @param soapRequest The soap request, which is type of MoveItemType
     * @param request The HTTP request
     * @returns item move result
     */
    async MoveItem(soapRequest: MoveItemType, request: Request): Promise<MoveItemResponseType> {
        assert(soapRequest instanceof MoveItemType);

        return MoveItem(soapRequest, request);
    }

    /**
     * Mark item as junk.
     * @param soapRequest The soap request, which is type of MarkAsJunkType
     * @param request The HTTP request
     * @returns item mark as junk result
     */
    async MarkAsJunk(soapRequest: MarkAsJunkType, request: Request): Promise<MarkAsJunkResponseType> {
        assert(soapRequest instanceof MarkAsJunkType);

        return MarkAsJunk(soapRequest, request);
    }

    /**
     * Mark all items as read.
     * @param soapRequest The soap request, which is type of MarkAllItemsAsReadType
     * @param request The HTTP request
     * @returns items mark as read result
     */
    async MarkAllItemsAsRead(soapRequest: MarkAllItemsAsReadType, request: Request): Promise<MarkAllItemsAsReadResponseType> {
        assert(soapRequest instanceof MarkAllItemsAsReadType);

        return MarkAllItemsAsRead(soapRequest, request);
    }

    /**
     * Archive item.
     * @param soapRequest The soap request, which is type of ArchiveItemType
     * @param request The HTTP request
     * @returns item archive result
     */
    async ArchiveItem(soapRequest: ArchiveItemType, request: Request): Promise<ArchiveItemResponseType> {
        assert(soapRequest instanceof ArchiveItemType);

        return ArchiveItem(soapRequest, request);
    }

    /**
     * Delete item.
     * @param soapRequest The soap request, which is type of DeleteItemType
     * @param request The HTTP request
     * @returns item delete result
     */
    async DeleteItem(soapRequest: DeleteItemType, request: Request): Promise<DeleteItemResponseType> {
        assert(soapRequest instanceof DeleteItemType);

        return DeleteItem(soapRequest, request);
    }

    /**
     * Report message.
     * @param soapRequest The soap request, which is type of ReportMessageType
     * @param request The HTTP request
     * @returns report message result
     */
    async ReportMessage(soapRequest: ReportMessageType, request: Request): Promise<ReportMessageResponseType> {
        assert(soapRequest instanceof ReportMessageType);

        return ReportMessage(soapRequest, request);
    }

    /**
     * Create attachment.
     * @param soapRequest The soap request, which is type of CreateAttachmentType
     * @param request The HTTP request
     * @returns attachment create result
     */
    async CreateAttachment(soapRequest: CreateAttachmentType, request: Request): Promise<CreateAttachmentResponseType> {
        assert(soapRequest instanceof CreateAttachmentType);

        return CreateAttachment(soapRequest, request);
    }

    /**
     * Get attachment.
     * @param soapRequest The soap request, which is type of GetAttachmentType
     * @param request The HTTP request
     * @returns attachment retrieved
     */
    async GetAttachment(soapRequest: GetAttachmentType, request: Request): Promise<GetAttachmentResponseType> {
        assert(soapRequest instanceof GetAttachmentType);

        return GetAttachment(soapRequest, request);
    }

    /**
     * Delete attachment.
     * @param soapRequest The soap request, which is type of DeleteAttachmentType
     * @param request The HTTP request
     * @returns attachment delete result
     */
    async DeleteAttachment(soapRequest: DeleteAttachmentType, request: Request): Promise<DeleteAttachmentResponseType> {
        assert(soapRequest instanceof DeleteAttachmentType);

        return DeleteAttachment(soapRequest, request);
    }

    /**
     * Send item.
     * @param soapRequest The soap request, which is type of SendItemType
     * @param request The HTTP request
     * @returns item sent result
     */
    async SendItem(soapRequest: SendItemType, request: Request): Promise<SendItemResponseType> {
        assert(soapRequest instanceof SendItemType);

        return SendItem(soapRequest, request);
    }

    /**
     * Upload items.
     * @param soapRequest The soap request, which is type of UploadItemsType
     * @param request The HTTP request
     * @returns items upload result
     */
    async UploadItems(soapRequest: UploadItemsType, request: Request): Promise<UploadItemsResponseType> {
        assert(soapRequest instanceof UploadItemsType);

        return UploadItems(soapRequest, request);
    }

    /**
     * Create folder path.
     * @param soapRequest The soap request, which is type of CreateFolderPathType
     * @param request The HTTP request
     * @returns folder path created
     */
    async CreateFolderPath(soapRequest: CreateFolderPathType, request: Request): Promise<CreateFolderPathResponseType> {
        assert(soapRequest instanceof CreateFolderPathType);

        return CreateFolderPath(soapRequest, request);
    }

    /**
     * Convert Id.
     * @param soapRequest The soap request, which is type of ConvertIdType
     * @param request The HTTP request
     * @returns convert result
     */
    async ConvertId(soapRequest: ConvertIdType, request: Request): Promise<ConvertIdResponseType> {
        assert(soapRequest instanceof ConvertIdType);

        return ConvertId(soapRequest, request);
    }

    /**
     * Subscribe.
     * @param soapRequest The soap request, which is type of SubscribeType
     * @param request The HTTP request
     * @returns subscribe result
     */
    async Subscribe(soapRequest: SubscribeType, request: Request): Promise<SubscribeResponseType> {
        assert(soapRequest instanceof SubscribeType);

        return Subscribe(soapRequest, request);
    }

    /**
     * Get streaming events.
     * @param soapRequest The soap request, which is type of GetStreamingEventsType
     * @param request The HTTP request
     * @param callback The callback function
     * @returns streaming events result
     */
    async GetStreamingEvents(soapRequest: GetStreamingEventsType, request: Request, callback: Function): Promise<GetStreamingEventsResponseType> {
        assert(soapRequest instanceof GetStreamingEventsType);

        return GetStreamingEvents(soapRequest, request, callback);
    }

    /**
     * Get app marketplace Url.
     * @param soapRequest The soap request, which is type of GetAppMarketplaceUrlType
     * @param request The HTTP request
     * @returns app marketplace Url
     */
    async GetAppMarketplaceUrl(soapRequest: GetAppMarketplaceUrlType, request: Request): Promise<GetAppMarketplaceUrlResponseMessageType> {
        assert(soapRequest instanceof GetAppMarketplaceUrlType);

        return GetAppMarketplaceUrl(soapRequest, request);
    }

    /**
     * Find item.
     * @param soapRequest The soap request, which is type of FindItemType
     * @param request The HTTP request
     * @returns find item result
     */
    async FindItem(soapRequest: FindItemType, request: Request): Promise<FindItemResponseType> {
        assert(soapRequest instanceof FindItemType);

        return FindItem(soapRequest, request);
    }

    /**
     * Create Managed Folder.
     * A managed folder is one abiding by retention policies.  
     * In Microsoft Exchange Server 2013, messaging records management (MRM) is performed by using 
     * retention tags and retention policies. A retention policy is a group of retention tags that 
     * can be applied to a mailbox. For more details, see Retention tags and retention policies. Managed 
     * folders, the MRM technology introduced in Exchange Server 2007, aren't supported.  
     * @param soapRequest The soap request, which is type of CreateManagedFolderRequestType
     * @param request The HTTP request
     * @returns Create managed folder result
     */
    async CreateManagedFolder(soapRequest: CreateManagedFolderRequestType, request: Request): Promise<CreateManagedFolderResponseType> {
        assert(soapRequest instanceof CreateManagedFolderRequestType);

        return CreateManagedFolder(soapRequest, request);
    }

    /**
     * Empty Folder.
     * @param soapRequest The soap request, which is type of EmptyFolder
     * @param request The HTTP request
     * @returns Empty folder result
     */
    async EmptyFolder(soapRequest: EmptyFolderType, request: Request): Promise<EmptyFolderResponseType> {
        assert(soapRequest instanceof EmptyFolderType);

        return EmptyFolder(soapRequest, request);
    }

    /**
     * Find a conversation
     * @param soapRequest The soap request, which is type of FindConversationType
     * @param request The HTTP request
     * @returns a conversation result
     */
    async FindConversation(soapRequest: FindConversationType, request: Request): Promise<FindConversationResponseMessageType> {
        assert(soapRequest instanceof FindConversationType);

        return FindConversation(soapRequest, request);
    }

    /**
     * Get conversation items
     * @param soapRequest The soap request, which is type of GetConversationItemsType
     * @param request The HTTP request
     * @returns The items in the conversation
     */
    async GetConversationItems(soapRequest: GetConversationItemsType, request: Request): Promise<GetConversationItemsResponseType> {
        assert(soapRequest instanceof GetConversationItemsType);

        return GetConversationItems(soapRequest, request);
    }

    /**
     * Apply a conversation action
     * @param soapRequest The soap request, which is type of ApplyConversationActionType
     * @param request The HTTP request
     * @returns ApplyConversationActionResponseType
     */
    async ApplyConversationAction(soapRequest: ApplyConversationActionType, request: Request): Promise<ApplyConversationActionResponseType> {
        assert(soapRequest instanceof ApplyConversationActionType);

        return ApplyConversationAction(soapRequest, request);
    }

    /**
     * Expand DL (Distribution List)
     * @param soapRequest The soap request, which is type of ExpandDLType
     * @param request The HTTP request
     * @returns The items in the conversation
     */
    async ExpandDL(soapRequest: ExpandDLType, request: Request): Promise<ExpandDLResponseType> {
        assert(soapRequest instanceof ExpandDLType);

        return ExpandDL(soapRequest, request);
    }

    /**
     * Get password expiration date
     * @param soapRequest The soap request, which is type of ExpandDLType
     * @param request The HTTP request
     * @returns The password expirsation date
     */
    async GetPasswordExpirationDate(soapRequest: GetPasswordExpirationDateType, request: Request): Promise<GetPasswordExpirationDateResponseMessageType> {
        assert(soapRequest instanceof GetPasswordExpirationDateType);

        return GetPasswordExpirationDate(soapRequest, request);
    }
}

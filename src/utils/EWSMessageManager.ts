import {
    PimLabel, KeepPimConstants, PimItemFormat, UserInfo, PimImportance, PimMessage, KeepPimMessageManager, base64Encode, 
    PimReceiptType, PimItemFactory, KeepPimCalendarManager, PimParticipationStatus, KeepPimManager, PimItem, PimParticipationInfo
} from '@hcllabs/openclientkeepcomponent';

import { EWSServiceManager } from '.';
import { UserContext } from '../keepcomponent';
import { Logger } from './logger';
import {
    PathToUnindexedFieldType, PathToExtendedFieldType
} from '../models/common.model';
import {
    DefaultShapeNamesType, UnindexedFieldURIType, ImportanceChoicesType, BodyTypeType, MessageDispositionType,
    MeetingRequestTypeType, CalendarItemTypeType, MapiPropertyIds, MapiPropertyTypeType
} from '../models/enum.model';
import {
    ItemResponseShapeType, MessageType, ItemIdType, FolderIdType, ItemChangeType, ItemType, ArrayOfRecipientsType,
    SetItemFieldType, AppendToItemFieldType, DeleteItemFieldType, MimeContentType, SingleRecipientType,
    NonEmptyArrayOfAttachmentsType, FileAttachmentType, ItemAttachmentType, BodyType, MeetingRequestMessageType,
    MeetingResponseMessageType, MeetingMessageType, AcceptItemType, DeclineItemType, TentativelyAcceptItemType, CancelCalendarItemType, MeetingCancellationMessageType
} from '../models/mail.model';
import {
    addExtendedPropertiesToPIM, addExtendedPropertyToPIM, identifiersForPathToExtendedFieldType, 
    getValueForExtendedFieldURI, getEmail, getMimeRecipients, addPimExtendedPropertyObjectToEWS
} from '../utils';
import { ewsItemClassType, getEWSId, getKeepIdPair } from './pimHelper';
import NodeICalSync, { CalendarResponse } from './../utils/icalutil/quattro-node-ical';
import { Request } from '@loopback/rest';
import * as util from 'util';
import { simpleParser, ParsedMail } from 'mailparser';
import { v4 as uuidv4 } from 'uuid';
import mimemessage = require("mimemessage");

// EWSServiceManager subclass implemented for Message related operations
export class EWSMessageManager extends EWSServiceManager {

    /**
     * Static objects
     */
    // Our singleton instance
    private static instance: EWSMessageManager;

    // Fields included for DEFAULT shape for notes
    private static defaultFields = [
        ...EWSServiceManager.idOnlyFields,
        UnindexedFieldURIType.ITEM_BODY, // Including body for now.  Body is not included on FindItem though
        UnindexedFieldURIType.ITEM_SUBJECT,
        UnindexedFieldURIType.ITEM_DATE_TIME_CREATED,
        UnindexedFieldURIType.ITEM_DATE_TIME_SENT,
        UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS,
        UnindexedFieldURIType.ITEM_SENSITIVITY,
        UnindexedFieldURIType.ITEM_SIZE,
        UnindexedFieldURIType.ITEM_IS_ASSOCIATED,
        UnindexedFieldURIType.ITEM_RESPONSE_OBJECTS,
        UnindexedFieldURIType.MESSAGE_FROM,
        UnindexedFieldURIType.MESSAGE_IS_DELIVERY_RECEIPT_REQUESTED
    ];

    // Fields included for ALL_PROPERTIES shape for notes
    private static allPropertiesFields = [
        ...EWSMessageManager.defaultFields,
        UnindexedFieldURIType.ITEM_CATEGORIES,
        UnindexedFieldURIType.ITEM_IMPORTANCE,
        UnindexedFieldURIType.ITEM_LAST_MODIFIED_TIME,
        UnindexedFieldURIType.CONTACTS_BUSINESS_HOME_PAGE,
        UnindexedFieldURIType.MESSAGE_BCC_RECIPIENTS,
        UnindexedFieldURIType.MESSAGE_CC_RECIPIENTS,
        UnindexedFieldURIType.CONVERSATION_CONVERSATION_ID,
        UnindexedFieldURIType.MESSAGE_CONVERSATION_INDEX,
        UnindexedFieldURIType.MESSAGE_CONVERSATION_TOPIC,
        UnindexedFieldURIType.ITEM_CULTURE,
        UnindexedFieldURIType.ITEM_DATE_TIME_RECEIVED,
        UnindexedFieldURIType.ITEM_DISPLAY_CC,
        UnindexedFieldURIType.ITEM_DISPLAY_TO,
        UnindexedFieldURIType.ITEM_EFFECTIVE_RIGHTS,
        UnindexedFieldURIType.ITEM_IN_REPLY_TO,
        UnindexedFieldURIType.MESSAGE_INTERNET_MESSAGE_ID,
        UnindexedFieldURIType.ITEM_INTERNET_MESSAGE_HEADERS,
        UnindexedFieldURIType.ITEM_IS_DRAFT,
        UnindexedFieldURIType.ITEM_IS_FROM_ME,
        UnindexedFieldURIType.MESSAGE_IS_READ,
        UnindexedFieldURIType.MESSAGE_IS_READ_RECEIPT_REQUESTED,
        UnindexedFieldURIType.ITEM_IS_RESEND,
        UnindexedFieldURIType.MESSAGE_IS_RESPONSE_REQUESTED,
        UnindexedFieldURIType.ITEM_IS_SUBMITTED,
        UnindexedFieldURIType.ITEM_IS_UNMODIFIED,
        UnindexedFieldURIType.ITEM_LAST_MODIFIED_TIME,
        UnindexedFieldURIType.MESSAGE_RECEIVED_BY,
        UnindexedFieldURIType.MESSAGE_RECEIVED_REPRESENTING,
        UnindexedFieldURIType.MESSAGE_REFERENCES,
        UnindexedFieldURIType.ITEM_REMINDER_DUE_BY,
        UnindexedFieldURIType.ITEM_REMINDER_IS_SET,
        UnindexedFieldURIType.ITEM_REMINDER_MINUTES_BEFORE_START,
        UnindexedFieldURIType.MESSAGE_REPLY_TO,
        UnindexedFieldURIType.MESSAGE_SENDER,
        UnindexedFieldURIType.MESSAGE_TO_RECIPIENTS,
        UnindexedFieldURIType.ITEM_WEB_CLIENT_EDIT_FORM_QUERY_STRING,
        UnindexedFieldURIType.ITEM_WEB_CLIENT_READ_FORM_QUERY_STRING
    ];

    /**
     * Static functions
     */

    /**
     * Get a singleton instance of this class.
     * @returns A signleton instance of this class.
     */
    public static getInstance(): EWSMessageManager {
        if (!EWSMessageManager.instance) {
            this.instance = new EWSMessageManager();
        }
        return this.instance;
    }

    /**
     * Public functions
     */

    /**
     * This function will fetech a group of messages using the Keep API and return an
     * array of corresponding EWS items populated with the fields requested in the shape.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param request The original SOAP request for the get.
     * @param shape EWS shape describing the information requested for each item.
     * @param startIndex Optional start index for items when paging.  Defaults to 0.
     * @param count Optional count of objects to request.  Defaults to 512.
     * @param fromLabel The unid of the label (folder) we are querying.  If not provided, we will query against all messages.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @returns An array of EWS items built from the returned PimItems.
     */
    async getItems(
        userInfo: UserInfo, 
        request: Request, 
        shape: ItemResponseShapeType, 
        startIndex?: number, 
        count?: number, 
        fromLabel?: PimLabel,
        mailboxId?: string
    ): Promise<MessageType[]> {

        const items: MessageType[] = [];
        let pimMessages: PimMessage[] | undefined = undefined;
        try {
            // FIXME: if documents = false, then filtering out non-mail messages does not work. This can be uncommented with LABS-2411 is delivered by Keep
            // const needDocuments = (shape.BaseShape === DefaultShapeNamesType.ID_ONLY &&
            //     (shape.AdditionalProperties === undefined || shape.AdditionalProperties.items.length === 0)) ? false : true;
            const needDocuments = true; 
            pimMessages = await KeepPimMessageManager.getInstance().getMailMessages(userInfo, fromLabel?.unid, needDocuments, startIndex, count, mailboxId);
        } catch (err) {
            Logger.getInstance().debug("Error retrieving PIM message entries: " + err);
            // If we throw the err here the client will continue in a loop with SyncFolderHierarchy and SyncFolderItems asking for messages.
        }
        if (pimMessages) {
            for (const pimMessage of pimMessages) {
                // For getItems, we don't want to fetch the mime for every message.
                const messageItem = await this.pimItemToEWSItem(pimMessage, userInfo, request, shape, mailboxId, fromLabel?.unid, false);
                items.push(messageItem);
            }
        }
        return items;
    }

    /**
     * Process updates for a message.
     * @param pimMessage The populated PimMessage containing the message data before the change.
     * @param change The EWS change type object that describes the changes to be made.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param toLabel Optional label of the item's parent folder.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @returns An updated message or undefined if the request could not be completed.
     */
    async updateItem(
        pimMessage: PimMessage, 
        change: ItemChangeType, 
        userInfo: UserInfo, 
        request: Request, 
        toLabel?: PimLabel,
        mailboxId?: string
    ): Promise<MessageType | undefined> {

        for (const fieldUpdate of change.Updates.items) {
            let newItem: ItemType | undefined = undefined;
            if (fieldUpdate instanceof SetItemFieldType || fieldUpdate instanceof AppendToItemFieldType) {
                // Should only have to worry about Item and Message field updates for notes
                newItem = fieldUpdate.Item ?? fieldUpdate.Message ?? fieldUpdate.SharingMessage ?? fieldUpdate.RoleMember
                    ?? fieldUpdate.MeetingMessage ?? fieldUpdate.MeetingRequest ?? fieldUpdate.MeetingResponse ?? fieldUpdate.MeetingCancellation;
            }

            if (fieldUpdate instanceof SetItemFieldType) {
                if (newItem === undefined) {
                    Logger.getInstance().error(`No new item set for update field: ${util.inspect(newItem, false, 5)}`);
                    continue;
                }

                if (fieldUpdate.FieldURI) {
                    const field = fieldUpdate.FieldURI.FieldURI.split(':')[1];
                    this.updatePimItemFieldValue(pimMessage, fieldUpdate.FieldURI.FieldURI, (newItem as any)[field]);
                } else if (fieldUpdate.ExtendedFieldURI && newItem.ExtendedProperty) {
                    const identifiers = identifiersForPathToExtendedFieldType(fieldUpdate.ExtendedFieldURI);
                    const newValue = getValueForExtendedFieldURI(newItem, identifiers);

                    const current = pimMessage.findExtendedProperty(identifiers);
                    if (current) {
                        current.Value = newValue;
                        pimMessage.updateExtendedProperty(identifiers, current);
                    }
                    else {
                        addExtendedPropertyToPIM(pimMessage, newItem.ExtendedProperty[0]);
                    }

                    if (fieldUpdate.ExtendedFieldURI.PropertyTag && fieldUpdate.ExtendedFieldURI.PropertyType === MapiPropertyTypeType.INTEGER) {
                        const tagNumber = Number.parseInt(fieldUpdate.ExtendedFieldURI.PropertyTag);
                        if (tagNumber === MapiPropertyIds.FLAG_STATUS) {
                            // Values defined here: https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxoflag/eda9fd25-6407-4cec-9e62-26e4f9d6a098
                            pimMessage.isFlaggedForFollowUp = (newValue === 2) ? true : false;
                        }
                    }
                } else if (fieldUpdate.IndexedFieldURI) {
                    Logger.getInstance().warn(`Unhandled SetItemField request for messages for field:  ${fieldUpdate.IndexedFieldURI}`);
                }
            }
            else if (fieldUpdate instanceof AppendToItemFieldType) {
                /*
                    The following properties are supported for the append action for notes:
                    - message:ReplyTo
                    - item:Body
                    - All the recipient and attendee collection properties
                */
                if (fieldUpdate.FieldURI?.FieldURI === UnindexedFieldURIType.ITEM_BODY && newItem) {
                    pimMessage.body = `${pimMessage.body}${newItem?.Body?.Value ?? ""}`
                } else {
                    Logger.getInstance().warn(`Unhandled AppendToItemField request for notes for field:  ${fieldUpdate.FieldURI?.FieldURI}`);
                }

                // TODO: Implement message:ReplyTo and message recipients with LABS-506
            }
            else if (fieldUpdate instanceof DeleteItemFieldType) {
                if (fieldUpdate.FieldURI) {
                    this.updatePimItemFieldValue(pimMessage, fieldUpdate.FieldURI.FieldURI, undefined);
                }
                else if (fieldUpdate.ExtendedFieldURI) {
                    pimMessage.deleteExtendedProperty(identifiersForPathToExtendedFieldType(fieldUpdate.ExtendedFieldURI));

                    if (fieldUpdate.ExtendedFieldURI.PropertyTag) {
                        const tagNumber = Number.parseInt(fieldUpdate.ExtendedFieldURI.PropertyTag);
                        if (tagNumber === MapiPropertyIds.FLAG_STATUS) {
                            pimMessage.isFlaggedForFollowUp = false;
                        }
                    }
                }
                else if (fieldUpdate.IndexedFieldURI) {
                    /* 
                        From https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/indexedfielduri
                        These are the fields that are supported in indexed field URI:
                            item:InternetMessageHeader
                            contacts:ImAddress
                            contacts:PhysicalAddress:Street
                            contacts:PhysicalAddress:City
                            contacts:PhysicalAddress:State
                            contacts:PhysicalAddress:Country
                            contacts:PhysicalAddress:PostalCode
                            contacts:PhoneNumber
                            contacts:EmailAddress
                            distributionlist:Members:Member
    
                        We don't support:
                            item:InternetMessageHeader
                            distributionlist:Members:Member
                    */
                    Logger.getInstance().warn(`Unhandled DeleteItemField request for notes for field:  ${fieldUpdate.IndexedFieldURI}`);
                }
            }
        }

        // Set the target folder if passed in
        if (toLabel) {
            pimMessage.parentFolderIds = [toLabel.folderId];  // May be possible we've lost other parent ids set by another client.  
            // Even if stored as an extra property for the client, it could have been updated on the server since stored on the client.
            // To shrink the window, we'd need to request for the item from the server and update it right away
        }

        // The PimMessage should now be updated with the new information.  Send it to Keep.
        await KeepPimMessageManager.getInstance().updateMessageData(pimMessage, userInfo, mailboxId);

        // For updates, we only return the ID_ONLY shape.  Just build that item here
        const item = new MessageType();
        const itemEWSId = getEWSId(pimMessage.unid, mailboxId);
        item.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);

        let folderId = pimMessage.parentFolderIds ? pimMessage.parentFolderIds[0] : undefined;
        if (toLabel) {
            folderId = toLabel.folderId;
        }


        //Add EWS parent folder Id
        if (folderId) {
            const parentEWSId = getEWSId(folderId, mailboxId);
            const parentId = new FolderIdType(parentEWSId, `ck-${parentEWSId}`);
            item.ParentFolderId = parentId;
        }

        item.ItemClass = ewsItemClassType(pimMessage);

        return item;
    }

    /**
     * This function issues a Keep API call to create a message based on the EWS item passed in.
     * @param item The EWS message item to create.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param request The original SOAP request for the create.
     * @param toLabel Optional target label (folder).
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @returns A new EWS message populated with ID_ONLY shape information about the created entry.
     */
    async createItem(
        item: MessageType, 
        userInfo: UserInfo, 
        request: Request, 
        disposition?: MessageDispositionType, 
        toLabel?: PimLabel,
        mailboxId?: string
    ): Promise<ItemType[]> {
        let newTargetItems: ItemType[] | undefined;
        const sendMessage = (disposition === MessageDispositionType.SEND_AND_SAVE_COPY || disposition === MessageDispositionType.SEND_ONLY);
        if (item instanceof AcceptItemType) {
            if (item.ReferenceItemId) {
                // Create the calendar response for the accept
                const participationInfo = new PimParticipationInfo();
                participationInfo.participantAction = PimParticipationStatus.ACCEPTED;
                newTargetItems = await this.getCalendarResponseItem(item.ReferenceItemId.Id, participationInfo, userInfo, request, mailboxId);
            }
        } else if (item instanceof DeclineItemType) {
            if (item.ReferenceItemId) {
                // Create the calendar response for the decline
                const participationInfo = new PimParticipationInfo();
                participationInfo.participantAction = PimParticipationStatus.DECLINED;
                newTargetItems = await this.getCalendarResponseItem(item.ReferenceItemId.Id, participationInfo, userInfo, request, mailboxId);
            }
        } else if (item instanceof TentativelyAcceptItemType) {
            if (item.ReferenceItemId) {
                    const participationInfo = new PimParticipationInfo();
                    participationInfo.participantAction = PimParticipationStatus.TENTATIVE;
                    // Create the calendar response for the decline
                    newTargetItems = await this.getCalendarResponseItem(item.ReferenceItemId.Id, participationInfo, userInfo, request, mailboxId);
                }
        } else if (item instanceof CancelCalendarItemType) {
            if (item.ReferenceItemId) {
                // Create the calendar response for the meeting cancellation
                newTargetItems = await this.getCancelCalendarItemResponse(item.ReferenceItemId.Id, userInfo, request);
            }
        } else if (item.MimeContent) {
            // Apple Mail is sending MimeContent as a string value directly.
            if (typeof (item.MimeContent) === 'string') {
                item.MimeContent = new MimeContentType(item.MimeContent);
            }

            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
            // Add information from the mime content to the item.
            await this.addRequestedPropertiesToEWSMessage(undefined, item, item.MimeContent.Value, 'base64', shape, mailboxId);
            // item = await populateItemFromMime(getUserMailData(request), item, item.MimeContent.Value);

            const pimMessage = this.pimItemFromEWSItem(item, request, undefined, toLabel);
            let receipts: PimReceiptType[] | undefined = undefined;
            if (pimMessage.returnReceipt) {
                receipts = receipts ?? [];
                receipts.push(PimReceiptType.READ);
            }
            if (pimMessage.deliveryReceipt) {
                receipts = receipts ?? [];
                receipts.push(PimReceiptType.DELIVERY);
            }
            const msgId = await KeepPimMessageManager
                .getInstance()
                .createMimeMessage(
                    userInfo, 
                    item.MimeContent.Value, 
                    sendMessage, 
                    receipts,
                    mailboxId
                );
            // After the mime message is created, call updateMessageData to add any extended fields
            if (pimMessage.extendedProperties) {
                pimMessage.unid = msgId;
                await KeepPimMessageManager.getInstance().updateMessageData(pimMessage, userInfo, mailboxId);
            }
            // Only need to return the ID of the new message
            newTargetItems = [];
            const newTargetItem = new MessageType();
            const msgEWSId = getEWSId(msgId, mailboxId);
            newTargetItem.ItemId = new ItemIdType(msgEWSId, `ck-${msgEWSId}`);
            newTargetItems.push(newTargetItem);
        } else {
            const pimMessage = this.pimItemFromEWSItem(item, request, undefined, toLabel);
            const msgId = await KeepPimMessageManager.getInstance().createMessage(userInfo, pimMessage, sendMessage, mailboxId);
            // Only need to return the ID of the new message
            newTargetItems = [];
            const newTargetItem = new MessageType();
            const msgEWSId = getEWSId(msgId, mailboxId);
            newTargetItem.ItemId = new ItemIdType(msgEWSId, `ck-${msgEWSId}`);
            newTargetItems.push(newTargetItem);
        }

        // Move the message to the desired folder
        if (newTargetItems && newTargetItems.length > 0) {
            for (const newTargetItem of newTargetItems) {
                if (newTargetItem?.ItemId) {
                    const newTargetItemId = newTargetItem.ItemId.Id;
                    if (toLabel) {
                        // Don't move to drafts or sent folders.  Domino does not allow and is handled automatically by state of the message.  See LABS-588
                        if (toLabel.view !== KeepPimConstants.SENT && toLabel.view !== KeepPimConstants.DRAFTS) {
                            await this.moveItem(newTargetItemId, toLabel.unid, userInfo);
                        }
                        if (((toLabel.view === KeepPimConstants.SENT && disposition === MessageDispositionType.SAVE_ONLY) ||
                            (toLabel.view === KeepPimConstants.DRAFTS && disposition === MessageDispositionType.SEND_ONLY))) {
                            // Delete the created draft item since we assume it has already been sent (we can't move the message to the sent folder by save...only by send).  
                            // See LABS-588 - can't move to drafts or sent
                            // Return the item information though in response to the client
                            // The next sync will correct the client with the real saved message.
                            Logger.getInstance().debug(`createItem deleting sent message that was for save only, ${newTargetItemId}`);
                            await this.deleteItem(newTargetItem.ItemId.Id, userInfo, mailboxId);
                            // await doDeleteItem(request, newTargetItem.ItemId!, undefined);
                        }
                    }
        
                    // Apple Mail sends emails with message diposition of SEND_ONLY.  Doing that should not save the message in the Sent Items folder according to 
                    // the EWS docs: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/createitem
                    // Using Exchange as a backend (with hcl account), while testing, appears to ignore the message disposition too.  We are choosing not to delete
                    // the message as well.
                    //
                    // if (!folderIdType && messageDispositionType === MessageDispositionType.SEND_ONLY) {
                    //     // Delete the item left in draft that we just created.  It should have been moved to the Sent folder
                    //     Logger.getInstance().debug(`doCreateItem deleting sent message that was sent and didn't have a target folder, ${newTargetItemId}`);
                    //     await doDeleteItem(request, newTargetItem.ItemId!, undefined);
                    // }
                }
            }
        }
        return newTargetItems ?? [];
    }

    /**
     * 
     * @param referencedItemId A referenced item Id
     * @param participationStatus Response options for a calendar or invitation
     * @param userInfo The user's credentials to be passed to Keep.
     * @param request The original SOAP request for the create.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @returns A populated EWS message
     */
    async getCalendarResponseItem(
        referencedItemId: string, 
        participationInfo: PimParticipationInfo, 
        userInfo: UserInfo, 
        request: Request,
        mailboxId?: string
    ): Promise<ItemType[] | undefined> {
        const newTargetItems: ItemType[] = [];
        let newTargetItem: ItemType | undefined;

        // Do not copy the pimitem.  Notes and the router, when sending from notes to notes, relies on the unid to remain the same throughout
        // the calendar invitation accept/decline/update/cancel flow.

        const [refItemId] = getKeepIdPair(referencedItemId);
        if (!refItemId) {
            const messageText = `Error sending calendar response.  referencedItemId, ${referencedItemId}, is an invalid id`;
            Logger.getInstance().debug(messageText);
            throw new Error(messageText);
        }

        // // Keep bug preventDelete=true causing error 500 https://jira01.hclpnp.com/browse/LABS-3137
        // // preventDelete must be true here because the client will try to delete the item after responding to the invitation.
        // // (i.e. ews clients expect the invitation item to be different from the calender item.  In domino the invitation item
        // // is converted to the calendar item)
        // const resp = await KeepPimCalendarManager
        //     .getInstance()
        //     .createCalendarResponse('default', refItemId, participationInfo, userInfo, true, mailboxId);
        const resp = await KeepPimCalendarManager
            .getInstance()
            .createCalendarResponse('default', refItemId, participationInfo, userInfo, false, mailboxId);

        Logger.getInstance().debug (`createCalendarResponse: ${util.inspect(resp, false, 5)}`);

        const pimItem = await KeepPimManager.getInstance().getPimItem(refItemId, userInfo, mailboxId);
        if (pimItem) {
            // Create a manager based on the item and convert it to an ItemType subclass object
            const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
            if (manager) {
                // Get the updated item.  It should have been converted to a calendar entry
                // For the returned calendar item we only need the ID_ONLY shape.

                const itemShape = new ItemResponseShapeType();
                itemShape.BaseShape = DefaultShapeNamesType.ID_ONLY;        
                newTargetItem = await manager.pimItemToEWSItem(pimItem, userInfo, request, itemShape, mailboxId);
                // // Set the parent id to the calendar  (HCL notes had it originally as the inbox)
                // try {
                //     const labels = await KeepPimLabelManager.getInstance().getLabels(userInfo);
                //     if (labels) {
                //         const folderId = new DistinguishedFolderIdType();
                //         folderId.Id = DistinguishedFolderIdNameType.CALENDAR;
                //         const label = findLabel(labels, folderId);
                //         if (label) {
                //             newTargetItem.ParentFolderId = new FolderIdType(label.unid, `ck-${label.unid}`);
                //         }
                //     }    
                // } catch (err) {
                //     // If there was an error getting the calendar, keep the original parent id.
                // }
                // newTargetItem = await manager.copyItem(pimItem, newTargetItem.ParentFolderId!.Id, userInfo, request, itemShape);
                newTargetItems.push(newTargetItem);

            }
        }

        return newTargetItems.length > 0 ? newTargetItems : undefined;
    }

    /**
     * Retrieve the copied item from Keep
     * @param copiedId The id of the item in Keep.
     * @param originalPimItem The pim item that was copied.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param request The original SOAP request for the get.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @returns The PimItem from Keep that corresponds to the copiedId
     * @throws An error if the copy fails
     */
    public async getCopiedPimItem(
        copiedId: string, 
        originalPimItem: PimItem, 
        userInfo: UserInfo,
        request: Request, 
        mailboxId?: string
    ): Promise<PimItem | undefined> {
        // // TODO: Update (hopefully it can be removed) this function when LABS-2759 is avalaible
        const newItem = await KeepPimManager.getInstance().getPimItem(copiedId, userInfo, mailboxId);
        if (newItem) {
            newItem.isRead = originalPimItem.isRead;
            await KeepPimManager.getInstance().updatePimItem(copiedId, userInfo, newItem, mailboxId); // Update message flags
        }
        return super.getCopiedPimItem(copiedId, originalPimItem, userInfo, request, mailboxId);
    }
    
    
    /**
     * Creates a new pim item. 
     * @param pimItem The note item to create.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @returns The new Pim note item created with Keep.
     * @throws An error if the create fails
     */
    protected async createNewPimItem(pimItem: PimMessage, userInfo: UserInfo, mailboxId?: string): Promise<PimMessage> {
        const mime = await KeepPimMessageManager.getInstance().getMimeMessage(userInfo, pimItem.unid, mailboxId);
        if (mime) {
            const mimeBuffer = Buffer.from(mime);
            const unid = await KeepPimMessageManager
                .getInstance()
                .createMimeMessage(userInfo, mimeBuffer.toString('base64'), undefined, undefined, mailboxId);
            const newItem = await KeepPimManager.getInstance().getPimItem(unid, userInfo, mailboxId);
            if (newItem === undefined || !newItem.isPimMessage()) {
                // Try to delete item since it may have been created
                try {
                    await KeepPimMessageManager.getInstance().deleteMessage(userInfo, unid, mailboxId);
                }
                catch {
                    // Ignore errors
                }

                Logger.getInstance().error(`New item not found or is not a message: ${newItem?.unid}`);
                throw new Error(`Unable to retrieve message ${unid} after create`);
            }

            return newItem;
        }
        else {
            throw new Error(`Did not find message ${pimItem.unid}`);
        }
    }

    /**
     * This function will make a Keep API request to delete a message
     * @param item The pim message or the message EWSId to delete.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @param hardDelete Indicator if the message should be deleted from trash as well
     */
    async deleteItem(item: string | PimMessage, userInfo: UserInfo, mailboxId?: string, hardDelete = false): Promise<void> {
        let unid;
        if (typeof item === 'string') {
            [unid] = getKeepIdPair(item);
        } else unid = item.unid;
        
        if (unid) {
            await KeepPimMessageManager.getInstance().deleteMessage(userInfo, unid, mailboxId, hardDelete);
        }
    }

    /**
     * Getter for the list of fields included in the default shape for notes
     * @returns Array of fields for the default shape
     */
    fieldsForDefaultShape(): UnindexedFieldURIType[] {
        return EWSMessageManager.defaultFields;
    }

    /**
     * Getter for the list of fields included in the all properties shape for notes
     * @returns Array of fields for the all properties shape
     */
    fieldsForAllPropertiesShape(): UnindexedFieldURIType[] {
        return EWSMessageManager.allPropertiesFields;
    }

    /**
     * Create and return a MessageType EWS item based on the passed in PimMessage, filtered by the passed in shape.
     * @param pimItem The PimMessage source for populating the EWS MessageType.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param request The original request received.
     * @param shape An EWS shape inicating the fields to include in the returned MessageType object.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @param targetParentFolderId An optional parent folder Id to apply to the returned item.
     * @param includeMime Set to true to fetch the mime content when building the item.
     * @returns A MessageType object populated with the requested data.
     */
    async pimItemToEWSItem(
        pimItem: PimMessage, 
        userInfo: UserInfo, 
        request: Request, 
        shape: ItemResponseShapeType, 
        mailboxId?: string,
        targetParentFolderId?: string, 
        includeMime = true,
    ): Promise<MessageType> {

        let message = new MessageType();

        if (pimItem.noticeType !== undefined) {
            if (pimItem.isMeetingRequest()) {
                message = new MeetingRequestMessageType();
            } else if (pimItem.isMeetingCancellation()) {
                message = new MeetingCancellationMessageType();
            } else { // Has to be a meeting response
                message = new MeetingResponseMessageType();
            }
            this.addRequestedPropertiesToEWSMeetingMessage(pimItem, message, shape, request, mailboxId);
        } else {
            if (includeMime === false) {
                // Don't fetch the mime for the message, just use the pimItem
                await this.addRequestedPropertiesToEWSMessage(pimItem, message, undefined, 'utf8', shape, mailboxId);
            } else {
                // We need to fetch the mime to get all the details we might need
                const mimeContent = await KeepPimMessageManager.getInstance().getMimeMessage(userInfo, pimItem.unid, mailboxId);
                if (mimeContent) {
                    // if (itemId.ChangeKey === undefined) {
                    //     itemId.ChangeKey = `ck-${itemId.Id}`;
                    // }
                    // itemRef.ItemId = itemId;

                    await this.addRequestedPropertiesToEWSMessage(pimItem, message, mimeContent, 'utf8', shape, mailboxId);

                    if (shape.IncludeMimeContent === true) {
                        message.MimeContent = new MimeContentType(base64Encode(mimeContent), 'UTF-8');
                    }
                } else {
                    Logger.getInstance().debug("No mime content falling through");
                    // It's a valid message, we just didn't get the data.  Throw the error to look in the mock data
                    throw new Error("No mime content falling through");
                }
            }
        }
        if (targetParentFolderId) {
            const parentFolderEWSId = getEWSId(targetParentFolderId, mailboxId);
            message.ParentFolderId = new FolderIdType(parentFolderEWSId, `ck-${parentFolderEWSId}`);
        }
        return message;
    }

    /**
     * Protected functions
     */

    /**
     * Create and populate a PimMessage item based on the passed in EWS MessageType item.
     * @param item EWS item used to create the PimMessage.
     * @param request The original EWS request.
     * @returns A new fully populated PimMessage representing the passed in EWS message.
     */
    pimItemFromEWSItem(item: MessageType, request: Request, existing?: object, parentLabel?: PimLabel): PimMessage {

        const result = PimItemFactory.newPimMessage({}, PimItemFormat.DOCUMENT);

        if (item.ItemId) {
            const [itemId] = getKeepIdPair(item.ItemId.Id);
            if (itemId) result.unid = itemId;
        }

        if (item.ParentFolderId) {
            const [parentFolderId] = getKeepIdPair(item.ParentFolderId.Id);
            if (parentFolderId) result.parentFolderIds = [parentFolderId];
        }

        // Copy any extended fields to the PimNote to preserve them.
        addExtendedPropertiesToPIM(item, result);

        result.createdDate = item.DateTimeCreated;
        result.lastModifiedDate = item.LastModifiedTime ?? item.DateTimeCreated;
        if (item.DateTimeSent) {
            result.sentDate = item.DateTimeSent;
        }

        if (item.Body) {
            result.body = item.Body.Value;
            if (item.Body.BodyType === BodyTypeType.HTML) {
                result.bodyType = "text/html; charset=utf-8";
            } else if (item.Body.BodyType === BodyTypeType.TEXT) {
                result.bodyType = "text/plain; charset=utf-8";
            }
        }

        if (item.Subject) {
            result.subject = item.Subject;
        }

        if (item.Size) {
            result.size = item.Size;
        }


        result.from = item.From?.Mailbox.EmailAddress;    //"From": "CN=David Kennedy/OU=USA/O=PNPHCL",
        if (!result.from) {
            if (item.From && item.From.Mailbox && item.From.Mailbox.Name) {
                const convertedEmailType = getEmail(item.From.Mailbox.Name);
                if (convertedEmailType.EmailAddress) {
                    result.from = convertedEmailType.EmailAddress;
                } else if (convertedEmailType.Name) {
                    result.from = convertedEmailType.Name;
                }
            }

            // If still no from, we must set it.  It is optional is EWS MessageType, however it is required by Keep
            // ...we assume it is from the current user
            if (!result.from) {
                const userInfo = UserContext.getUserInfo(request);
                if (userInfo?.userId) {
                    result.from = userInfo.userId;
                }
            }
        }

        let emailAddresses: string[] = [];
        if (item.ToRecipients) {
            item.ToRecipients.Mailbox.forEach(eAddress => {
                if (!eAddress.EmailAddress) {
                    if (eAddress.Name) {
                        const convertedEmailType = getEmail(eAddress.Name);
                        if (convertedEmailType.EmailAddress) {
                            emailAddresses.push(convertedEmailType.EmailAddress);
                        } else if (convertedEmailType.Name) {
                            emailAddresses.push(convertedEmailType.Name);
                        }
                    }
                } else {
                    emailAddresses.push(eAddress.EmailAddress);
                }
            });
        }
        if (emailAddresses.length > 0) {
            result.to = emailAddresses;
        }

        emailAddresses = [];
        if (item.BccRecipients) {
            item.BccRecipients.Mailbox.forEach(eAddress => {
                if (!eAddress.EmailAddress) {
                    if (eAddress.Name) {
                        const convertedEmailType = getEmail(eAddress.Name);
                        if (convertedEmailType.EmailAddress) {
                            emailAddresses.push(convertedEmailType.EmailAddress);
                        } else if (convertedEmailType.Name) {
                            emailAddresses.push(convertedEmailType.Name);
                        }
                    }
                } else {
                    emailAddresses.push(eAddress.EmailAddress);
                }
            });
        }
        if (emailAddresses.length > 0) {
            result.bcc = emailAddresses;
        }

        emailAddresses = [];
        if (item.CcRecipients) {
            item.CcRecipients.Mailbox.forEach(eAddress => {
                if (!eAddress.EmailAddress) {
                    if (eAddress.Name) {
                        const convertedEmailType = getEmail(eAddress.Name);
                        if (convertedEmailType.EmailAddress) {
                            emailAddresses.push(convertedEmailType.EmailAddress);
                        } else if (convertedEmailType.Name) {
                            emailAddresses.push(convertedEmailType.Name);
                        }
                    }
                } else {
                    emailAddresses.push(eAddress.EmailAddress);
                }
            });
        }
        if (emailAddresses.length > 0) {
            result.cc = emailAddresses;
        }

        if (item.IsReadReceiptRequested) {
            result.returnReceipt = item.IsReadReceiptRequested;
        }

        if (item.IsDeliveryReceiptRequested) {
            result.deliveryReceipt = item.IsDeliveryReceiptRequested;
        }

        if (item.Importance) {
            switch (item.Importance) {
                case ImportanceChoicesType.HIGH:
                    result.importance = PimImportance.HIGH;
                    break;
                case ImportanceChoicesType.LOW:
                    result.importance = PimImportance.LOW;
                    break;
                default:
                    result.importance = PimImportance.NONE;
                    break;
            }
        }
        // FIXME - Add Attachment information

        return result;
    }

    /**
     * Build mime content for the passed in PimMessage representing a meeting message.  Our code that does this
     * first builds out a full MessageType object and then constructs the mime from that.  We might want to 
     * investigate making this more efficient.
     * @param pimItem The PimMessage to convert to mime.
     * @returns A MimeContentType populated with the information about the message from the PimMessage.
     */
    protected buildMimeContentForMeetingMessage(pimItem: PimMessage): MimeContentType {

        // Create a temporary MeetingMessage and populate it with the common fields we need
        // to construct the mime
        const baseMessage = new MessageType();

        // We only need to, cc, bcc and replyTo from the MessageType to create the mime.
        const props = [
            UnindexedFieldURIType.MESSAGE_TO_RECIPIENTS,
            UnindexedFieldURIType.MESSAGE_CC_RECIPIENTS,
            UnindexedFieldURIType.MESSAGE_BCC_RECIPIENTS,
            UnindexedFieldURIType.MESSAGE_REPLY_TO
        ];

        props.forEach(field => {
            this.updateEWSItemFieldValue(baseMessage, pimItem, field);
        })

        // Use the populated MeetingMessage to build the mime content
        // Build the top-level multipart MIME message.
        const msg = mimemessage.factory({
            contentType: 'multipart/mixed',
            body: []
        });
        msg.header('Message-ID', pimItem.icalMessageId);

        let data = [];
        if (baseMessage.ToRecipients) {
            for (const recip of baseMessage.ToRecipients.Mailbox) {
                data.push(recip.EmailAddress);
            }
        }
        msg.header('To', data.join(", "));

        data = [];
        if (baseMessage.CcRecipients) {
            for (const recip of baseMessage.CcRecipients.Mailbox) {
                data.push(recip.EmailAddress);
            }
        }
        msg.header('Cc', data.join(", "));

        data = [];
        if (baseMessage.BccRecipients) {
            for (const recip of baseMessage.BccRecipients.Mailbox) {
                data.push(recip.EmailAddress);
            }
        }
        msg.header('Bcc', data.join(", "));

        data = [];
        if (baseMessage.ReplyTo) {
            for (const recip of baseMessage.ReplyTo.Mailbox) {
                data.push(recip.EmailAddress);
            }
        }
        msg.header('ReplyTo', data.join(", "));

        msg.header('Date', pimItem.createdDate);
        if (pimItem.subject) {
            msg.header('Subject', pimItem.subject);
        }
        if (pimItem.from) {
            let msgFrom = getEmail(pimItem.from).EmailAddress;
            if (!msgFrom) {
                msgFrom = pimItem.from;
            }
            msg.header('From', msgFrom);
        }

        // // Build the multipart/alternate MIME entity containing both the HTML and plain text entities.
        // const alternateEntity = mimemessage.factory({
        //     contentType: 'multipart/alternate',
        //     body: []
        // });

        // Build the HTML MIME entity.
        let pBody = pimItem.body;
        if (!pBody && pimItem.icalStream) {
            const icalResponse: CalendarResponse = NodeICalSync.sync.parseICS(pimItem.icalStream);
            if (icalResponse) {
                const iEvents = Object.values(icalResponse);
                if (iEvents && Array.isArray(iEvents) && iEvents.length > 0) {
                    for (const iEvent of iEvents) {
                        if (iEvent.type && iEvent.type === 'VEVENT') {
                            if (iEvent.comment) {
                                if (typeof (iEvent.comment) === 'string') {
                                    pBody = iEvent.comment;
                                } else if ((iEvent.comment as any).val) {
                                    pBody = (iEvent.comment as any).val;
                                    if (pBody) {
                                        // Comment may not parse correctly.
                                        // "COMMENT;ALTREP=CID:<FFFF__=0ABB0C8CDFE5A1D78f9e8a93df938690918c0AB@>:Go Buckeyes!"
                                        // will parse to <FFFF__=0ABB0C8CDFE5A1D78f9e8a93df938690918c0AB@>:Go Buckeyes!
                                        const idx = pBody.indexOf(">:");
                                        if (idx >= 0) {
                                            pBody = pBody.substring(idx + 2);  // +2 for the >:
                                        }
                                    }
                                }
                            }
                            break;
                        }
                    }
                }
            }
        }

        let htmlEntity: any | undefined;
        if (pBody) {
            htmlEntity = mimemessage.factory({
                contentType: 'text/html;charset=utf-8',
                body: pBody
            });
        }

        // // Build the plain text MIME entity.
        // const plainEntity = mimemessage.factory({
        //     body: fromString(pimItem.body)
        // });

        // // Add both the HTML and plain text entities to the multipart/alternate entity.
        // alternateEntity.body.push(htmlEntity);
        // alternateEntity.body.push(plainEntity);

        // // Add the multipart/alternate entity to the top-level MIME message.
        // msg.body.push(alternateEntity);
        if (htmlEntity) {
            msg.body.push(htmlEntity);
        }

        if (pimItem.icalStream) {
            // Build the CALENDAR MIME entity.
            const calEntity = mimemessage.factory({
                contentType: 'text/calendar; charset="utf-8"',
                contentTransferEncoding: 'base64',
                body: base64Encode(pimItem.icalStream)
            });
            // calEntity.header('Content-Disposition', 'attachment ;filename="mypicture.png"');

            // Add the calendar entity to the top-level MIME message.
            msg.body.push(calEntity);
        }

        return new MimeContentType(base64Encode(msg.toString()), 'UTF-8');
    }

    /**
    * Adds requested properties to an EWS meeting message item, based on the shape passed in.
    * @param pimMessage The PIM item containing the  properties.
    * @param message The EWS item where the properties should be set. 
    * @param shape The shape setting used to determine which properties to add. If not specified, only extended properties saved in the PIM item will be added.
    * @param request The original EWS request.
    * @param mailboxId SMTP mailbox delegator or delegatee address.
    */
    protected addRequestedPropertiesToEWSMeetingMessage(
        pimMessage: PimMessage,
        meeting: MeetingMessageType,
        shape: ItemResponseShapeType, 
        request: Request,
        mailboxId?: string
    ): void {
        // First include all the fields required by the base shape type

        const baseFields = this.defaultPropertiesForShape(shape);

        // We have either a MeetingRequestMessageType or MeetingResponseMessageType.  Both an iCal response
        // with events.  In order to avoid processing the iCal multiple times, we will do it here and store
        // any info we might need from it and pass it along to the updateEWSFieldValueForMeeing call
        const iCalObject: any = {};
        if (pimMessage.icalStream) {
            const icalResponse: CalendarResponse = NodeICalSync.sync.parseICS(pimMessage.icalStream);
            if (icalResponse) {
                const iEvents = Object.values(icalResponse);
                if (iEvents && Array.isArray(iEvents) && iEvents.length > 0) {
                    for (const iEvent of iEvents) {
                        if (iEvent.type && iEvent.type === 'VEVENT') {
                            iCalObject['Start'] = iEvent.start;
                            iCalObject['End'] = iEvent.end;
                            iCalObject['IsCancelled'] = false; // FIXME: determine if cancelled or not ...maybe in ical?
                            if (!pimMessage.body) {
                                let pBody;
                                if (iEvent.comment) {
                                    if (typeof (iEvent.comment) === 'string') {
                                        pBody = iEvent.comment;
                                    } else if ((iEvent.comment as any).val) {
                                        pBody = (iEvent.comment as any).val;
                                        if (pBody) {
                                            // Comment may not parse correctly.
                                            // "COMMENT;ALTREP=CID:<FFFF__=0ABB0C8CDFE5A1D78f9e8a93df938690918c0AB@>:Go Buckeyes!"
                                            // will parse to <FFFF__=0ABB0C8CDFE5A1D78f9e8a93df938690918c0AB@>:Go Buckeyes!
                                            const idx = pBody.indexOf(">:");
                                            if (idx >= 0) {
                                                pBody = pBody.substring(idx + 2);  // +2 for the >:
                                            }
                                        }
                                    }
                                    if (pBody) {
                                        pimMessage.body = pBody;
                                    }
                                }
                            }
                            break;
                        }
                    }
                }
            }
        }

        // Populate the default fields for this shape
        for (const fieldName of baseFields) {
            this.updateEWSItemFieldValueForMeeting(meeting, pimMessage, fieldName, request, iCalObject, mailboxId);
        }

        // Now add the reqeusted additional and extended fields
        if (shape.AdditionalProperties) {
            for (const property of shape.AdditionalProperties.items) {
                if (property instanceof PathToUnindexedFieldType) {
                    this.updateEWSItemFieldValueForMeeting(meeting, pimMessage, property.FieldURI, request, iCalObject, mailboxId);
                }
                else if (property instanceof PathToExtendedFieldType) {
                    if (pimMessage) {
                        const pimProperty = pimMessage.findExtendedProperty(identifiersForPathToExtendedFieldType(property));
                        if (pimProperty) {
                            addPimExtendedPropertyObjectToEWS(pimProperty, meeting);
                        }
                    }
                }
            }
        }
        else if (shape.BaseShape === DefaultShapeNamesType.ALL_PROPERTIES) {
            // If shape was not passed in, or it is AllProperties, add all exteneded properties to the EWS item
            if (pimMessage) {
                pimMessage.extendedProperties.forEach((property: any) => {
                    addPimExtendedPropertyObjectToEWS(property, meeting);
                });
            }
        }

        // If we're asked to include mime content, we have to build it from the PimItem separately.
        // Add it if we haven't already.
        if (shape.IncludeMimeContent && meeting.MimeContent === undefined) {
            meeting.MimeContent = this.buildMimeContentForMeetingMessage(pimMessage);
        }

    }

    /**
    * Adds requested properties to an EWS item, based on the shape passed in.  We will first attempt to 
    * update the fields from the mime, then fall back to any data set in the optional pimMessage.
    * @param pimMessage The PIM item containing the  properties.  If this parameter is undefined, we will only use the mime content.
    * @param message The EWS item where the properties should be set. 
    * @param mimeContent The encoded mime for the message.
    * @param encoding The encoding of the mime content.
    * @param shape The shape setting used to determine which properties to add. If not specified, only extended properties saved in the PIM item will be added.
    * @param mailboxId SMTP mailbox delegator or delegatee address.
    */
    protected async addRequestedPropertiesToEWSMessage(
        pimMessage: PimMessage | undefined,
        message: MessageType, 
        mimeContent: string | undefined, 
        encoding = 'base64', 
        shape?: ItemResponseShapeType,
        mailboxId?: string
    ): Promise<void> {
        let messageBuffer: Buffer | undefined;
        let parsed: ParsedMail | undefined;
        if (mimeContent !== undefined) {
            // Parse the mime
            messageBuffer = Buffer.from(mimeContent, encoding);
            parsed = await simpleParser(messageBuffer);
            Logger.getInstance().debug(`Mime content: ${util.inspect(parsed, false, 5)}`);
        }

        // First include all the fields required by the base shape type
        if (shape) {
            const baseFields = this.defaultPropertiesForShape(shape);
            for (const fieldName of baseFields) {
                await this.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, fieldName, mailboxId);
                // Need to handle size here
                if (fieldName === UnindexedFieldURIType.ITEM_SIZE) {
                    if (message.Size === undefined && messageBuffer !== undefined) {
                        message.Size = messageBuffer.length;
                    }
                }
            }
        }

        // Now add the reqeusted additional and extended fields
        if (shape?.AdditionalProperties) {
            for (const property of shape.AdditionalProperties.items) {
                if (property instanceof PathToUnindexedFieldType) {
                    await this.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, property.FieldURI, mailboxId);
                    // Need to handle size here
                    if (property.FieldURI === UnindexedFieldURIType.ITEM_SIZE) {
                        if (message.Size === undefined && messageBuffer !== undefined) {
                            message.Size = messageBuffer.length;
                        }
                    }
                }
                else if (property instanceof PathToExtendedFieldType) {
                    if (pimMessage) {
                        const pimProperty = pimMessage.findExtendedProperty(identifiersForPathToExtendedFieldType(property));
                        if (pimProperty) {
                            addPimExtendedPropertyObjectToEWS(pimProperty, message);
                        }
                    }
                }
            }
        } else if (shape === undefined || shape.BaseShape === DefaultShapeNamesType.ALL_PROPERTIES) {
            // If shape was not passed in, or it is AllProperties, add all exteneded properties to the EWS item
            if (pimMessage) {
                pimMessage.extendedProperties.forEach((property: any) => {
                    addPimExtendedPropertyObjectToEWS(property, message);
                });
            }
        }

        //Add EWS parent folder Id
        if (pimMessage?.parentFolderIds && pimMessage?.parentFolderIds[0]) {
            const parentFolderEWSId = getEWSId(pimMessage?.parentFolderIds[0], mailboxId);
            message.ParentFolderId = new FolderIdType(parentFolderEWSId, `ck-${parentFolderEWSId}`);
        }
    }


    /**
    * Update the passed in EWS message's specific field with the information stored in the pimItem.  This function is
    * used to populate fields requested in a GetItem/FindItem/SyncX EWS operation.
    * @param message The EWS message to update with a value.
    * @param pimMessage The PIM message from which to read the value.
    * @param fieldIdentifier The EWS field we are mapping.
    * @param mailboxId SMTP mailbox delegator or delegatee address.
    * @returns True if the field was processed.
    */
    protected updateEWSItemFieldValue(
        message: MessageType, 
        pimMessage: PimMessage, 
        fieldIdentifier: UnindexedFieldURIType,
        mailboxId?: string
    ): boolean {

        let handled = true;

        switch (fieldIdentifier) {
            case UnindexedFieldURIType.MESSAGE_CONVERSATION_INDEX:
                const DELIVERY_FILETIME = '01D7C2E9ED';
                message.ConversationIndex ='01'+DELIVERY_FILETIME+pimMessage.threadId;
                const respLevel = pimMessage.conversationIndex ? parseInt(pimMessage.conversationIndex, 10) : 0;
                if (respLevel > 0) {
                    let respLevelsString = '';
                    for (let ind=1; ind<= respLevel; ind++){
                        respLevelsString += String(ind).padStart(10, '0');
                    }
                    message.ConversationIndex += respLevelsString;
                }
                // encode hex string to base64
                message.ConversationIndex = Buffer.from(message.ConversationIndex, 'hex').toString('base64');
                break;
            case UnindexedFieldURIType.ITEM_DATE_TIME_SENT:
                if (pimMessage.sentDate) {
                    message.DateTimeSent = pimMessage.sentDate;
                }
                break;
            case UnindexedFieldURIType.MESSAGE_FROM:
                {
                    // From?: SingleRecipientType;
                    if (pimMessage.from) {
                        message.From = new SingleRecipientType();
                        message.From.Mailbox = getEmail(pimMessage.from);
                        if (!message.From.Mailbox.EmailAddress) {
                            message.From.Mailbox.EmailAddress = pimMessage.from;
                        }
                        // FIXME: Hack for digital week demo to avoid domino names.
                        if (message.From.Mailbox.EmailAddress) {
                            message.From.Mailbox.Name = message.From.Mailbox.EmailAddress;
                        }
                    }
                }
                break;
            case UnindexedFieldURIType.MESSAGE_IS_DELIVERY_RECEIPT_REQUESTED:
                // IsDeliveryReceiptRequested?: boolean;
                if (pimMessage.deliveryReceipt) {
                    message.IsDeliveryReceiptRequested = pimMessage.deliveryReceipt;
                }
                break;
            case UnindexedFieldURIType.MESSAGE_BCC_RECIPIENTS:
                {
                    // BccRecipients?: ArrayOfRecipientsType;
                    if (pimMessage.bcc) {
                        const bccRecipients = new ArrayOfRecipientsType();
                        for (const addr of pimMessage.bcc) {
                            const recipientMail = getEmail(addr);
                            if (!recipientMail.EmailAddress) {
                                recipientMail.EmailAddress = addr;
                            }
                            // FIXME: Hack for digital week demo to avoid domino names.
                            if (recipientMail.EmailAddress) {
                                recipientMail.Name = recipientMail.EmailAddress;
                            }
                            bccRecipients.Mailbox.push(recipientMail);
                        }
                        message.BccRecipients = bccRecipients;
                    }
                }
                break;
            case UnindexedFieldURIType.MESSAGE_CC_RECIPIENTS:
                {
                    // CcRecipients?: ArrayOfRecipientsType;
                    if (pimMessage.cc) {
                        const ccRecipients = new ArrayOfRecipientsType();
                        for (const addr of pimMessage.cc) {
                            const recipientMail = getEmail(addr);
                            if (!recipientMail.EmailAddress) {
                                recipientMail.EmailAddress = addr;
                            }
                            // FIXME: Hack for digital week demo to avoid domino names.
                            if (recipientMail.EmailAddress) {
                                recipientMail.Name = recipientMail.EmailAddress;
                            }
                            ccRecipients.Mailbox.push(recipientMail);
                        }
                        message.CcRecipients = ccRecipients;
                    }
                }
                break;
            case UnindexedFieldURIType.ITEM_IMPORTANCE:
                // Message support HIGH, LOW and undefined
                if (pimMessage.importance === PimImportance.HIGH) {
                    message.Importance = ImportanceChoicesType.HIGH;
                } else if (pimMessage.importance === PimImportance.LOW) {
                    message.Importance = ImportanceChoicesType.LOW;
                }
                break;
            case UnindexedFieldURIType.ITEM_SIZE:
                message.Size = pimMessage.size;
                break;
            case UnindexedFieldURIType.MESSAGE_IS_READ:
                message.IsRead = pimMessage.isRead;
                break;
            case UnindexedFieldURIType.ITEM_DATE_TIME_RECEIVED:
                if (pimMessage.receivedDate) {
                    message.DateTimeReceived = pimMessage.receivedDate;
                }
                break;
            case UnindexedFieldURIType.MESSAGE_IS_READ_RECEIPT_REQUESTED:
                // IsReadReceiptRequested?: boolean;
                if (pimMessage.returnReceipt) {
                    message.IsReadReceiptRequested = pimMessage.returnReceipt;
                }
                break;
            case UnindexedFieldURIType.MESSAGE_REPLY_TO:
                {
                    // ReplyTo?: ArrayOfRecipientsType;
                    if (pimMessage.replyTo) {
                        const replyToRecipients = new ArrayOfRecipientsType();
                        for (const addr of pimMessage.replyTo) {
                            const recipientMail = getEmail(addr);
                            if (!recipientMail.EmailAddress) {
                                recipientMail.EmailAddress = addr;
                            }
                            // FIXME: Hack for digital week demo to avoid domino names.
                            if (recipientMail.EmailAddress) {
                                recipientMail.Name = recipientMail.EmailAddress;
                            }
                            replyToRecipients.Mailbox.push(recipientMail);
                        }
                        message.ReplyTo = replyToRecipients;
                    }
                }
                break;
            case UnindexedFieldURIType.MESSAGE_SENDER:
                {
                    // Sender?: SingleRecipientType;
                    const singleRecipient = new SingleRecipientType();
                    if (pimMessage.from) {
                        singleRecipient.Mailbox = getEmail(pimMessage.from);
                        if (!singleRecipient.Mailbox.EmailAddress) {
                            singleRecipient.Mailbox.EmailAddress = pimMessage.from;
                        }
                    }

                    // FIXME: Hack for digital week demo to avoid domino names.
                    if (singleRecipient.Mailbox && singleRecipient.Mailbox.EmailAddress) {
                        singleRecipient.Mailbox.Name = singleRecipient.Mailbox.EmailAddress;
                    }

                    message.Sender = singleRecipient;
                }
                break;
            case UnindexedFieldURIType.MESSAGE_TO_RECIPIENTS:
                {
                    if (pimMessage.to) {
                        const toRecipients = new ArrayOfRecipientsType();
                        for (const addr of pimMessage.to) {
                            const recipientMail = getEmail(addr);
                            if (!recipientMail.EmailAddress) {
                                recipientMail.EmailAddress = addr;
                            }
                            // FIXME: Hack for digital week demo to avoid domino names.
                            if (recipientMail.EmailAddress) {
                                recipientMail.Name = recipientMail.EmailAddress;
                            }
                            toRecipients.Mailbox.push(recipientMail);
                        }
                        message.ToRecipients = toRecipients;
                    }
                }
                break;

            case UnindexedFieldURIType.ITEM_WEB_CLIENT_EDIT_FORM_QUERY_STRING:
            case UnindexedFieldURIType.ITEM_WEB_CLIENT_READ_FORM_QUERY_STRING:
            case UnindexedFieldURIType.ITEM_IS_RESEND:
            case UnindexedFieldURIType.MESSAGE_IS_RESPONSE_REQUESTED:
            case UnindexedFieldURIType.ITEM_IS_SUBMITTED:
            case UnindexedFieldURIType.ITEM_IS_UNMODIFIED:
            case UnindexedFieldURIType.MESSAGE_RECEIVED_BY:
            case UnindexedFieldURIType.MESSAGE_RECEIVED_REPRESENTING:
            case UnindexedFieldURIType.ITEM_REMINDER_DUE_BY:
            case UnindexedFieldURIType.ITEM_REMINDER_IS_SET:
            case UnindexedFieldURIType.ITEM_REMINDER_MINUTES_BEFORE_START:
            case UnindexedFieldURIType.ITEM_DISPLAY_CC:
            case UnindexedFieldURIType.ITEM_DISPLAY_TO:
            case UnindexedFieldURIType.ITEM_IN_REPLY_TO:
            case UnindexedFieldURIType.MESSAGE_INTERNET_MESSAGE_ID:
            case UnindexedFieldURIType.ITEM_INTERNET_MESSAGE_HEADERS:
            case UnindexedFieldURIType.ITEM_IS_FROM_ME:
            case UnindexedFieldURIType.ITEM_IS_ASSOCIATED:
            case UnindexedFieldURIType.ITEM_RESPONSE_OBJECTS:
            case UnindexedFieldURIType.CONVERSATION_CONVERSATION_ID:
            case UnindexedFieldURIType.MESSAGE_CONVERSATION_TOPIC:
            case UnindexedFieldURIType.ITEM_CULTURE:
            case UnindexedFieldURIType.CONTACTS_BUSINESS_HOME_PAGE:
            default:
                handled = false;
        }

        if (handled === false) {
            handled = super.updateEWSItemFieldValue(message, pimMessage, fieldIdentifier, mailboxId);
        }

        if (handled === false) {
            Logger.getInstance().warn(`Unrecognized field ${fieldIdentifier} not processed for Message`);
        }

        return handled;
    }

    /**
     * Populate requested fields in an EWS meeting request or response from the passed in PimMessage.
     * @param message The message we are populating.
     * @param pimMessage PimMessage object also containing information about the message.
     * @param fieldIdentifier The field we are mapping.
     * @param request The original EWS request for the information.
     * @param iCalObject iCal response
     * @param mailboxId SMTP mailbox address of the delegate or non-delegate.
     * @returns True if the field was processed.
     */
    protected updateEWSItemFieldValueForMeeting(
        message: MeetingMessageType,
        pimMessage: PimMessage,
        fieldIdentifier: UnindexedFieldURIType,
        request: Request, 
        iCalObject: any,
        mailboxId?: string
    ): boolean {

        // First process the field using the PimMessage
        let handled = true;

        // We will handle both MeetingRequestMessageType and MeetingResponseMessageType here
        const meetingRequest = message instanceof MeetingRequestMessageType ? message : undefined;
        const meetingResponse = message instanceof MeetingResponseMessageType ? message : undefined;

        switch (fieldIdentifier) {
            // Common to both types
            case UnindexedFieldURIType.CALENDAR_IS_ORGANIZER:
                {
                    // See if the current user is the sender
                    if (message.Sender === undefined) {
                        this.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.MESSAGE_SENDER, mailboxId);
                    }

                    const currentUser = UserContext.getUserInfo(request);
                    message.IsOrganizer = currentUser && message.Sender?.Mailbox.EmailAddress === currentUser.userId;
                }
                break;
            case UnindexedFieldURIType.MEETING_IS_DELEGATED:
                if (pimMessage.isDelegatedRequest()) {
                    message.IsDelegated = true;
                }
                break;
            // These need special handling depending on the type
            case UnindexedFieldURIType.MEETING_ASSOCIATED_CALENDAR_ITEM_ID:
                {
                    const refCalId = pimMessage.referencedCalendarItemUnid;
                    if (refCalId) {
                        if (meetingResponse) {
                            const refCalEWSId = getEWSId(refCalId, mailboxId);
                            meetingResponse.AssociatedCalendarItemId = new ItemIdType(refCalEWSId, `ck-${refCalEWSId}`);
                        }

                    }
                }
                break;
            case UnindexedFieldURIType.CALENDAR_START:
                {
                    const start = iCalObject['Start'] ?? pimMessage.start;
                    if (start) {
                        if (meetingRequest) {
                            meetingRequest.Start = start;
                        } else if (meetingResponse) {
                            meetingResponse.Start = start;
                        }
                    }
                }
                break;
            case UnindexedFieldURIType.CALENDAR_END:
                {
                    const end = iCalObject['End'] ?? pimMessage.end;
                    if (end) {
                        if (meetingRequest) {
                            meetingRequest.End = end;
                        } else if (meetingResponse) {
                            meetingResponse.End = end;
                        }
                    }
                }
                break;
            case UnindexedFieldURIType.MEETING_PROPOSED_START:
                if (meetingResponse) {
                    if (pimMessage.isCounterProposalRequest()) {
                        meetingResponse.ProposedStart = pimMessage.newStartDate;
                    }
                }
                break;
            case UnindexedFieldURIType.MEETING_PROPOSED_END:
                if (meetingResponse) {
                    if (pimMessage.isCounterProposalRequest()) {
                        meetingResponse.ProposedEnd = pimMessage.newEndDate;
                    }
                }
                break;
            case UnindexedFieldURIType.MEETING_REQUEST_MEETING_REQUEST_TYPE:
                if (meetingRequest) {
                    // TODO:  Are there other values we should be considering here?
                    meetingRequest.MeetingRequestType = MeetingRequestTypeType.NEW_MEETING_REQUEST;
                }
                break;
            case UnindexedFieldURIType.CALENDAR_IS_MEETING:
                if (meetingRequest) {
                    meetingRequest.IsMeeting = true;
                }
                break;
            case UnindexedFieldURIType.CALENDAR_IS_CANCELLED:
                if (meetingRequest) {
                    meetingRequest.IsCancelled = false;
                }
                break;
            case UnindexedFieldURIType.CALENDAR_CALENDAR_ITEM_TYPE:
                if (meetingRequest) {
                    meetingRequest.CalendarItemType = CalendarItemTypeType.SINGLE;
                }
                break;
            case UnindexedFieldURIType.ITEM_MIME_CONTENT: {
                    // If we're asked to include mime content, we have to build it from the PimItem separately
                    message.MimeContent = this.buildMimeContentForMeetingMessage(pimMessage);
                }
                break;
            default:
                handled = false;
        }

        let handledFromPimMessage = false;
        if (!handled) {
            handledFromPimMessage = this.updateEWSItemFieldValue(message, pimMessage, fieldIdentifier, mailboxId);
        }

        if (handled === false && handledFromPimMessage === false) {
            Logger.getInstance().warn(`Unrecognized field ${fieldIdentifier} not processed for Mime`);
        }

        return handled || handledFromPimMessage;
    }


    /**
     * Populate requested fields in an EWS message from the mime content of the message.
     * @param message The message we are populating.
     * @param parsed The parsed mime content.
     * @param pimMessage PimMessage object also containing information about the message.  If a requested field is not
     *                   contained in the mime, this will be used to populate it.
     * @param fieldIdentifier The field we are mapping.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @returns True if the field was processed.
     */
    protected async updateEWSItemFieldValueFromMime(
        message: MessageType, 
        parsed: ParsedMail | undefined, 
        pimMessage: PimMessage | undefined, 
        fieldIdentifier: UnindexedFieldURIType,
        mailboxId?: string
    ): Promise<boolean> {

        // First process the field using the PimMessage
        let handledFromPimMessage = false;
        if (pimMessage) {
            handledFromPimMessage = this.updateEWSItemFieldValue(message, pimMessage, fieldIdentifier, mailboxId);
        }

        let handled = false;
        if (parsed !== undefined) {
            handled = true;
            // After processing using the data in the PimMessage, try processing using the parsed mime.  The values in the mime override those in the PimMessage
            switch (fieldIdentifier) {
                // Fields handled by EWSServiceManager for all other paths
                case UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS:
                    if (parsed.attachments && parsed.attachments.length && parsed.attachments.some(attachment => attachment.contentDisposition !== 'inline')) {
                        message.HasAttachments = true;
                    } else {
                        message.HasAttachments = false;
                    }
                    break;
                case UnindexedFieldURIType.ITEM_ATTACHMENTS:
                    {
                        if (parsed.attachments && parsed.attachments.length) {
                            message.Attachments = new NonEmptyArrayOfAttachmentsType();

                            for (const att of parsed.attachments) {
                                if (att.contentDisposition === 'inline') {
                                    continue;
                                }

                                if (att.contentType === 'message/rfc822') {
                                    const parsedItem = await simpleParser(att.content);
                                    // for (const _item of Object.values(userData.ITEMS)) {
                                    //     if (_item instanceof MessageType && _item.InternetMessageId === parsedItem.messageId) {

                                    const attchment = new ItemAttachmentType();
                                    message.Attachments.push(attchment);

                                    // FIXME: We should lookup the attachment by InternetMessageId and store the unid as the ContentId
                                    attchment.ContentId = uuidv4();
                                    attchment.ContentType = att.contentType;
                                    attchment.Name = parsedItem.subject;

                                    //         UserMailboxData.ATTACHMENTS_CONTENTS[attchment.ContentId] = deepcoy(_item);
                                    //         break;
                                    //     }
                                    // }
                                } else {
                                    const attchment = new FileAttachmentType();
                                    message.Attachments.push(attchment);

                                    attchment.ContentId = uuidv4();
                                    attchment.ContentType = att.contentType;
                                    attchment.Size = att.size;
                                    attchment.Name = att.filename;

                                    // UserMailboxData.ATTACHMENTS_CONTENTS[attchment.ContentId] = att.content.toString('base64');
                                }
                            }
                        }
                    }
                    break;
                case UnindexedFieldURIType.ITEM_SUBJECT:
                    message.Subject = parsed.subject;
                    break;
                case UnindexedFieldURIType.ITEM_IN_REPLY_TO:
                    message.InReplyTo = parsed.inReplyTo;
                    break;
                case UnindexedFieldURIType.ITEM_BODY:
                case UnindexedFieldURIType.ITEM_UNIQUE_BODY:
                case UnindexedFieldURIType.ITEM_NORMALIZED_BODY:
                case UnindexedFieldURIType.ITEM_TEXT_BODY:
                    if (parsed.html) {
                        message.Body = new BodyType(parsed.html, BodyTypeType.HTML);
                    } else {
                        message.Body = new BodyType(parsed.text ?? '', BodyTypeType.TEXT);
                    }
                    break;
                case UnindexedFieldURIType.MESSAGE_FROM:
                    if (parsed.from && parsed.from.value) {
                        let val: any = {};
                        if (Array.isArray(parsed.from.value)) {
                            val = parsed.from.value[0];
                        } else {
                            val = parsed.from.value;
                        }
                        const recipient = new SingleRecipientType();
                        recipient.Mailbox = getEmail(val.address.length > 0 ? val.address : val.name);
                        if (!recipient.Mailbox.EmailAddress) {
                            recipient.Mailbox.EmailAddress = val.address;
                        }
                        if (val.name && val.name.length > 0) {
                            recipient.Mailbox.Name = val.name;
                        }
                        // FIXME: Hack to avoid domino name showing for digital week
                        if (recipient.Mailbox.EmailAddress) {
                            recipient.Mailbox.Name = recipient.Mailbox.EmailAddress;
                        }

                        message.From = recipient;
                    }
                    break;
                case UnindexedFieldURIType.ITEM_DATE_TIME_SENT:
                    if (parsed.date) {
                        message.DateTimeSent = parsed.date;
                    }
                    break;
                case UnindexedFieldURIType.MESSAGE_BCC_RECIPIENTS:
                    if (parsed.bcc && parsed.bcc.value && Array.isArray(parsed.bcc.value)) {
                        message.BccRecipients = getMimeRecipients(parsed.bcc.value);
                    }
                    break;
                case UnindexedFieldURIType.MESSAGE_CC_RECIPIENTS:
                    if (parsed.cc && parsed.cc.value && Array.isArray(parsed.cc.value)) {
                        message.CcRecipients = getMimeRecipients(parsed.cc.value);
                    }
                    break;
                // FIXME:  Can we get this setting?
                // case UnindexedFieldURIType.CONVERSATION_CONVERSATION_ID:
                //     break;
                case UnindexedFieldURIType.MESSAGE_CONVERSATION_TOPIC:
                    message.ConversationTopic = parsed.headers.get('thread-topic')?.toString();
                    break;
                case UnindexedFieldURIType.MESSAGE_INTERNET_MESSAGE_ID:
                    message.InternetMessageId = parsed.messageId;
                    break;
                case UnindexedFieldURIType.MESSAGE_REFERENCES:
                    {
                        const references = parsed.references;
                        if (references) {
                            // it can be array or string
                            message.References = Array.isArray(references) ? references.join(' ') : references;
                        }
                    }
                    break;
                case UnindexedFieldURIType.MESSAGE_REPLY_TO:
                    if (parsed.replyTo && parsed.replyTo.value && Array.isArray(parsed.replyTo.value)) {
                        message.ReplyTo = getMimeRecipients(parsed.replyTo.value);
                    }
                    break;
                case UnindexedFieldURIType.MESSAGE_SENDER:
                    if (parsed.from && parsed.from.value) {
                        let val: any = {};
                        if (Array.isArray(parsed.from.value)) {
                            val = parsed.from.value[0];
                        } else {
                            val = parsed.from.value;
                        }
                        const recipient = new SingleRecipientType();
                        recipient.Mailbox = getEmail(val.address.length > 0 ? val.address : val.name);
                        if (!recipient.Mailbox.EmailAddress) {
                            recipient.Mailbox.EmailAddress = val.address;
                        }
                        if (val.name && val.name.length > 0) {
                            recipient.Mailbox.Name = val.name;
                        }
                        // FIXME: Hack to avoid domino name showing for digital week
                        if (recipient.Mailbox.EmailAddress) {
                            recipient.Mailbox.Name = recipient.Mailbox.EmailAddress;
                        }

                        message.Sender = recipient;
                    }
                    break;
                case UnindexedFieldURIType.MESSAGE_TO_RECIPIENTS:
                    if (parsed.to && parsed.to.value && Array.isArray(parsed.to.value)) {
                        message.ToRecipients = getMimeRecipients(parsed.to.value);
                    }
                    break;
                default:
                    handled = false;
            }
        }

        if (handled === false && handledFromPimMessage === false) {
            Logger.getInstance().warn(`Unrecognized field ${fieldIdentifier} not processed for Mime`);
        }

        return handled || handledFromPimMessage;
    }

        /**
     * Update a value for an EWS field in a Keep PIM message. 
     * @param pimItem The pim message that will be updated.
     * @param fieldIdentifier The EWS unindexed field identifier
     * @param newValue The new value to set. The type is based on what fieldIdentifier is set to. To delete a field, pass in undefined. 
     * @returns true if field was handled
     */
    protected  updatePimItemFieldValue(pimMessage: PimMessage, fieldIdentifier: UnindexedFieldURIType, newValue: any | undefined): boolean {
        return super.updatePimItemFieldValue(pimMessage, fieldIdentifier, newValue);
        // TODO: Implement with LABS-506

        /*
        MESSAGE_CONVERSATION_INDEX = 'message:ConversationIndex',
        MESSAGE_CONVERSATION_TOPIC = 'message:ConversationTopic',
        MESSAGE_INTERNET_MESSAGE_ID = 'message:InternetMessageId',
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
        */
    }

    /**
     * 
     * @param referencedItemId A referenced item Id
     * @param userInfo The user's credentials to be passed to Keep.
     * @param request The original SOAP request for the create.
     * @returns A populated EWS message
     */
    async getCancelCalendarItemResponse(
        referencedItemId: string, 
        userInfo: UserInfo, 
        request: Request,
    ): Promise<ItemType[] | undefined> {
        const newTargetItems: ItemType[] = [];
        let newTargetItem: ItemType | undefined;

        const [refItemId, mailboxId] = getKeepIdPair(referencedItemId);
        
        if (!refItemId) {
            const messageText = `Error cancelling calendar item. referencedItemId, ${referencedItemId}, is an invalid id`;
            Logger.getInstance().debug(messageText);
            throw new Error(messageText);
        }

        // TODO:
        // Make here a meeting cancellation request to the domino server through a KeepPimCalendarManager
        // after such API will be implemented on the domino side (LABS-3067)

        const pimItem = await KeepPimManager.getInstance().getPimItem(refItemId, userInfo, mailboxId);
        if (pimItem) {
            // Create a manager based on the item and convert it to an ItemType subclass object
            const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
            if (manager) {
                // Get the updated item.  It should have been converted to a calendar entry
                // For the returned calendar item we only need the ID_ONLY shape.
                const itemShape = new ItemResponseShapeType();
                itemShape.BaseShape = DefaultShapeNamesType.ID_ONLY;        
                newTargetItem = await manager.pimItemToEWSItem(pimItem, userInfo, request, itemShape, mailboxId);
                newTargetItems.push(newTargetItem);

            }
        }

        return newTargetItems.length > 0 ? newTargetItems : undefined;
    }
}
import {
    PimNote, PimLabel, KeepPimNotebookManager, KeepPimConstants, PimItemFormat, UserInfo, KeepPimLabelManager, 
    PimItemFactory, base64Encode
} from '@hcllabs/openclientkeepcomponent';
import { EWSServiceManager } from '.';
import { Logger } from './logger';
import { DefaultShapeNamesType, UnindexedFieldURIType, MessageDispositionType } from '../models/enum.model';
import {
    ItemResponseShapeType, MessageType, ItemIdType, FolderIdType, ItemChangeType, ItemType,SetItemFieldType,
    AppendToItemFieldType, DeleteItemFieldType
} from '../models/mail.model';
import { getKeepIdPair } from './pimHelper';
import {
    addExtendedPropertiesToPIM, addExtendedPropertyToPIM, identifiersForPathToExtendedFieldType, 
    getValueForExtendedFieldURI, getEWSId
} from '../utils';
import { Request } from '@loopback/rest';
import * as util from 'util';

// EWSServiceManager subclass implemented for Notes related operations
export class EWSNotesManager extends EWSServiceManager {

    private static instance: EWSNotesManager;

    public static getInstance(): EWSNotesManager {
        if (!EWSNotesManager.instance) {
            this.instance = new EWSNotesManager();
        }
        return this.instance;
    }

    pimItemFromEWSItem(item: MessageType, request: Request, existing?: object): PimNote {
        const noteObject: any = existing ? existing : {};

        if (item.ItemId) {
            const [itemId] = getKeepIdPair(item.ItemId.Id);
            noteObject.uid = itemId;
        }

        //Get parent folder id from EWS parent folder Id
        if (item.ParentFolderId) {
            const [parentFolderId] = getKeepIdPair(item.ParentFolderId.Id);
            noteObject.ParentFolder = parentFolderId;
        }

        if (item.Subject) {
            noteObject.Subject = item.Subject;
        }

        // FIXME:  Do we need to handle HTML and/or MIME body content specially?
        if (item.Body) {
            noteObject.Body = base64Encode(item.Body.Value);
        }

        if (item.DateTimeCreated) {
            noteObject.TimeCreated = item.DateTimeCreated.toISOString();
        }

        if (item.Categories && item.Categories.String.length > 0) {
            // TODO: Always set to array when LABS-1550 fixed
            if (item.Categories.String.length === 1) {
                noteObject.Categories = item.Categories.String[0];
            }
            else {
                noteObject.Categories = item.Categories.String;
            }
        }

        if (item.IsRead !== undefined) {
            noteObject["@unread"] = (item.IsRead === false);
        }

        const rtn = PimItemFactory.newPimNote(noteObject, PimItemFormat.DOCUMENT);

        if (item.LastModifiedTime) {
            rtn.lastModifiedDate = item.LastModifiedTime;
        }

        // Copy any extended fields to the PimNote to preserve them.
        addExtendedPropertiesToPIM(item, rtn);

        return rtn;
    }

    // Fields included for DEFAULT shape for notes
    private static defaultFields = [
        ...EWSServiceManager.idOnlyFields,
        UnindexedFieldURIType.ITEM_BODY, // Including body for now.  Body is not included on FindItem though
        UnindexedFieldURIType.ITEM_SUBJECT,
        UnindexedFieldURIType.ITEM_DATE_TIME_CREATED,
        UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS,
        UnindexedFieldURIType.ITEM_SENSITIVITY,
        UnindexedFieldURIType.ITEM_SIZE
    ];

    // Fields included for ALL_PROPERTIES shape for notes
    private static allPropertiesFields = [
        ...EWSNotesManager.defaultFields,
        UnindexedFieldURIType.ITEM_CATEGORIES,
        UnindexedFieldURIType.ITEM_IMPORTANCE,
        UnindexedFieldURIType.MESSAGE_IS_READ,
        UnindexedFieldURIType.ITEM_LAST_MODIFIED_TIME
    ];

    async getItems(
        userInfo: UserInfo, 
        request: Request, 
        shape: ItemResponseShapeType, 
        startIndex?: number, 
        count?: number, 
        fromLabel?: PimLabel,
        mailboxId?: string
    ): Promise<MessageType[]> {

        const noteItems: MessageType[] = [];
        let pimItems: PimNote[] | undefined = undefined;
        try {
            const needDocuments = (shape.BaseShape === DefaultShapeNamesType.ID_ONLY 
                && (shape.AdditionalProperties === undefined || shape.AdditionalProperties.items.length === 0)
            ) ? false : true;

            pimItems = await KeepPimNotebookManager
                .getInstance()
                .getNotes(userInfo, needDocuments, startIndex, count, fromLabel?.unid, mailboxId);
        } catch (err) {
            Logger.getInstance().debug("Error retrieving PIM note entries: " + err);
            // If we throw the err here the client will continue in a loop with SyncFolderHierarchy and SyncFolderItems asking for messages.
        }
        if (pimItems) {
            for (const pimItem of pimItems) {
                const noteItem = await this.pimItemToEWSItem(pimItem, userInfo, request, shape, mailboxId, fromLabel?.unid);
                if (noteItem) {
                    noteItems.push(noteItem);
                } else {
                    Logger.getInstance().error(`An EWS note type could not be created for PIM note ${pimItem.unid}`);
                }
            }
        }
        return noteItems;
    }

    async updateItem(
        pimNote: PimNote, 
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
                newItem = fieldUpdate.Item ?? fieldUpdate.Message;
            }

            if (fieldUpdate instanceof SetItemFieldType) {
                if (newItem === undefined) {
                    Logger.getInstance().error(`No new item set for update field: ${util.inspect(newItem, false, 5)}`);
                    continue;
                }

                if (fieldUpdate.FieldURI) {
                    const field = fieldUpdate.FieldURI.FieldURI.split(':')[1];
                    this.updatePimItemFieldValue(pimNote, fieldUpdate.FieldURI.FieldURI, (newItem as any)[field]);
                } else if (fieldUpdate.ExtendedFieldURI && newItem.ExtendedProperty) {
                    const identifiers = identifiersForPathToExtendedFieldType(fieldUpdate.ExtendedFieldURI);
                    const newValue = getValueForExtendedFieldURI(newItem, identifiers);

                    const current = pimNote.findExtendedProperty(identifiers);
                    if (current) {
                        current.Value = newValue;
                        pimNote.updateExtendedProperty(identifiers, current);
                    }
                    else {
                        addExtendedPropertyToPIM(pimNote, newItem.ExtendedProperty[0]);
                    }
                } else if (fieldUpdate.IndexedFieldURI) {
                    Logger.getInstance().warn(`Unhandled SetItemField request for notes for field:  ${fieldUpdate.IndexedFieldURI}`);
                }
            }
            else if (fieldUpdate instanceof AppendToItemFieldType) {
                /*
                    The following properties are supported for the append action for notes:
                    - item:Body
                */
                if (fieldUpdate.FieldURI?.FieldURI === UnindexedFieldURIType.ITEM_BODY && newItem) {
                    pimNote.body = `${pimNote.body}${newItem?.Body ?? ""}`
                } else {
                    Logger.getInstance().warn(`Unhandled AppendToItemField request for notes for field:  ${fieldUpdate.FieldURI?.FieldURI}`);
                }
            }
            else if (fieldUpdate instanceof DeleteItemFieldType) {
                if (fieldUpdate.FieldURI) {
                    this.updatePimItemFieldValue(pimNote, fieldUpdate.FieldURI.FieldURI, undefined);
                }
                else if (fieldUpdate.ExtendedFieldURI) {
                    pimNote.deleteExtendedProperty(identifiersForPathToExtendedFieldType(fieldUpdate.ExtendedFieldURI));
                }
                else if (fieldUpdate.IndexedFieldURI) {
                    // Only handling these for contacts currently
                    Logger.getInstance().warn(`Unhandled DeleteItemField request for notes for field:  ${fieldUpdate.IndexedFieldURI}`);
                }
            }
        }

        // Set the target folder if passed in
        if (toLabel) {
            pimNote.parentFolderIds = [toLabel.folderId];  // May be possible we've lost other parent ids set by another client.  
            // Even if stored as an extra property for the client, it could have been updated on the server since stored on the client.
            // To shrink the window, we'd need to request for the item from the server and update it right away
        }

        // The PimNote should now be updated with the new information.  Send it to Keep.
        await KeepPimNotebookManager.getInstance().updateNote(pimNote, userInfo, mailboxId);

        if (pimNote.parentFolderIds === undefined) {
            try {
                const labels = await KeepPimLabelManager.getInstance().getLabels(userInfo, undefined, undefined, mailboxId);
                const journal = labels.find(label => { return label.view === KeepPimConstants.JOURNAL });
                pimNote.parentFolderIds = journal ? [journal.folderId] : undefined;
            }
            catch (err) {
                Logger.getInstance().error(`Unable to get labels to determine note parent folder: ${err}`);
            }
        }
        const shape = new ItemResponseShapeType();
        shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
        return this.pimItemToEWSItem(pimNote, userInfo, request, shape, mailboxId);
    }

    async createItem(
        item: MessageType, 
        userInfo: UserInfo, 
        request: Request, 
        disposition?: MessageDispositionType, 
        toLabel?: PimLabel,
        mailboxId?: string
    ): Promise<MessageType[]> {
        const pimNote = this.pimItemFromEWSItem(item, request);
        const unid = await KeepPimNotebookManager.getInstance().createNote(pimNote, userInfo, mailboxId);
        const targetItem = new MessageType();
        const itemEWSId = getEWSId(unid, mailboxId);
        targetItem.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);

        // TODO:  Uncomment this line when we can set the parent folder to the base notes folder
        //        See LABS-866
        // If no toLabel is passed in, we will create the note in the top level notes folder
        // if (toLabel === undefined) {
        //     const labels = await KeepPimLabelManager.getInstance().getLabels(userInfo);
        //     toLabel = labels.find(label => { return label.view === KeepPimConstants.JOURNAL });
        // }

        // The Keep API does not support creating a note IN a folder, so we must now move the 
        // note to the desired folder
        // TODO:  If this note is being created in the top level notes folder, do not attempt to 
        //        move the contact to the top level folder as this will fail.
        //        See LABS-866
        if (toLabel !== undefined && toLabel.view !== KeepPimConstants.JOURNAL) {
            //should improve in notes compose/decompose
            await this.moveItem(itemEWSId, toLabel.folderId, userInfo);
        }
        return [targetItem];
    }

    /**
     * Creates a new pim item. 
     * @param pimItem The note item to create.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @returns The new Pim note item created with Keep.
     * @throws An error if the create fails
     */
    protected async createNewPimItem(pimItem: PimNote, userInfo: UserInfo, mailboxId?: string): Promise<PimNote> {
        const unid = await KeepPimNotebookManager.getInstance().createNote(pimItem, userInfo, mailboxId);
        const newItem = await KeepPimNotebookManager.getInstance().getNote(userInfo, unid, mailboxId);
        if (newItem === undefined) {
            // Try to delete item since it may have been created
            try {
                await KeepPimNotebookManager.getInstance().deleteNote(unid, userInfo, undefined, mailboxId);
            }
            catch {
                // Ignore errors
            }
            throw new Error(`Unable to retrieve note ${unid} after create`);
        }

        return newItem;
    }

    /**
     * This function will make a Keep API request to delete a note.
     * @param item The pim note or unid of the note to delete.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @param hardDelete Indicator if the note should be deleted from trash as well
     */
    async deleteItem(item: string | PimNote, userInfo: UserInfo, mailboxId?: string, hardDelete = false): Promise<void> {
        const unid = (typeof item === 'string') ? item : item.unid;
        await KeepPimNotebookManager.getInstance().deleteNote(unid, userInfo, hardDelete, mailboxId);
    }

    /**
     * Getter for the list of fields included in the default shape for notes
     * @returns Array of fields for the default shape
     */
    fieldsForDefaultShape(): UnindexedFieldURIType[] {
        return EWSNotesManager.defaultFields;
    }

    /**
     * Getter for the list of fields included in the all properties shape for notes
     * @returns Array of fields for the all properties shape
     */
    fieldsForAllPropertiesShape(): UnindexedFieldURIType[] {
        return EWSNotesManager.allPropertiesFields;
    }

    async pimItemToEWSItem(
        pimNote: PimNote, 
        userInfo: UserInfo, 
        request: Request, 
        shape: ItemResponseShapeType,
        mailboxId?: string,
        targetParentFolderId?: string
    ): Promise<MessageType> {
        // Notes are treated as MessageTypes by EWS
        const note = new MessageType();

        // Add all requested properties based on the shape
        this.addRequestedPropertiesToEWSItem(pimNote, note, shape, mailboxId);
        
        if (targetParentFolderId) {
            const parentFolderEWSId = getEWSId(targetParentFolderId, mailboxId);
            note.ParentFolderId = new FolderIdType(parentFolderEWSId, `ck-${parentFolderEWSId}`);
        }
        
        return note;
    }
}
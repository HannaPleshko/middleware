import {
    PimItem, PimLabel, KeepPimMessageManager, UserInfo, PimImportance, PimTaskPriority, KeepPimManager, PimLabelTypes,
    PimRule, PimCommonEventsJmap, hasHTML
} from '@hcllabs/openclientkeepcomponent';
import { 
    PimRecurrenceDayOfWeek, PimRecurrenceFrequency, PimRecurrenceRule 
} from '@hcllabs/openclientkeepcomponent/dist/keep/pim/jmap/PimRecurrenceRule';
import { Logger } from './logger';
import { 
    ArrayOfStringsType, PathToUnindexedFieldType, PathToExtendedFieldType, DaysOfWeekType, PathToIndexedFieldType
} from '../models/common.model';
import {
    DefaultShapeNamesType, UnindexedFieldURIType, ItemClassType, ImportanceChoicesType, SensitivityChoicesType,
    BodyTypeType, ResponseClassType, ResponseCodeType, MessageDispositionType, DayOfWeekType, MonthNamesType, 
    DayOfWeekIndexType
} from '../models/enum.model';
import {
    ItemResponseShapeType, ItemType, MessageType, TaskType, ItemIdType, FolderIdType, BodyType, ItemChangeType,
    ContactItemType, CalendarItemType, RecurrenceType, AbsoluteYearlyRecurrencePatternType, 
    RelativeYearlyRecurrencePatternType, AbsoluteMonthlyRecurrencePatternType, RelativeMonthlyRecurrencePatternType,
    WeeklyRecurrencePatternType, DailyRecurrencePatternType, EndDateRecurrenceRangeType, NumberedRecurrenceRangeType,
    NoEndRecurrenceRangeType, MimeContentType
} from '../models/mail.model';
import {
    ewsItemClassType, getItemAttachments, addPimExtendedPropertyObjectToEWS, effectiveRightsForItem,
    identifiersForPathToExtendedFieldType, EWSNotesManager, EWSTaskManager, EWSMessageManager, EWSContactsManager,
    getFallbackCreatedDate, EWSCalendarManager, getEWSId
} from '../utils';
import { getKeepIdPair } from './pimHelper';
import { Request } from '@loopback/rest';
import * as util from 'util';
import { DateTime } from 'luxon';

/**
 * EWSServiceManager defines an abstract base class for service managers used in our EWS implemenation.
 * We will subclass this class with managers for the different types of PIM item we manage (Messages,
 * Contacts, Calendar Items, Tasks and Notes).  Common functions will be implemented in this class and
 * many functions marked as abstract will be implemented by the subclasses.  The goal is to hide the
 * Keep specific implementation inside the managers so that callers aren't required to know anything
 * about Keep and Keep PIM internals.
 * 
 * Current design is for users of these managers to request a singleton instance of a manager using 
 * getInstanceFromItem() or getInstanceFromPimItem(), passing in either an EWS ItemType or a PimItem
 * subclass object.  The getInstance functions will return an appropriate manager based on the type 
 * of the object passed in.  The goal is to maintain a single public signature so that the caller 
 * doesn't have to make any decisions based on the type of the manager being used.
 * 
 * Managers will handle common operations like createItem(), updateItem() and deleteItem().  In some
 * cases, a single static function in this class will suffice--for instance, getItem() or moveItem()
 * where the item type is either irrelevent or not known at the time of the call.
 */

export abstract class EWSServiceManager {

    /**
     * Static data
     */

    // Fields included for ID_ONLY shape for all item types
    protected static idOnlyFields = [
        UnindexedFieldURIType.ITEM_ITEM_ID,
        UnindexedFieldURIType.ITEM_PARENT_FOLDER_ID,
        UnindexedFieldURIType.ITEM_ITEM_CLASS
    ];

    /**
     * Static functions
     */

    /**
     * Create an instance of an EWSServiceManager based on a passed in ItemType subclass object.
     * @param pimItem The ItemType subclass object identifying which manager to return.
     * @returns A specific service manager to handle the ItemType type.
     */
    static getInstanceFromItem(ewsItem: ItemType): EWSServiceManager | undefined {

        if ((ewsItem instanceof MessageType || ewsItem instanceof ItemType) && ewsItem.ItemClass === ItemClassType.ITEM_CLASS_NOTE) {
            return EWSNotesManager.getInstance();
        }
        else if (ewsItem instanceof TaskType) {
            return EWSTaskManager.getInstance();
        }
        else if (ewsItem instanceof MessageType) {
            return EWSMessageManager.getInstance();
        }
        else if (ewsItem instanceof ContactItemType) {
            return EWSContactsManager.getInstance();
        }
        else if (ewsItem instanceof CalendarItemType) {
            return EWSCalendarManager.getInstance();
        }
        else {
            Logger.getInstance().debug(`Unrecognized ItemType: ${typeof ewsItem}`);
        }

        return undefined;
    }

    /**
     * Create an instance of an EWSServiceManager based on a passed in PimItem subclass object.
     * @param pimItem The PimItem subclass object identifying which manager to return.
     * @returns A specific service manager to handle the PimItem type.
     */
    static getInstanceFromPimItem(pimItem: PimItem): EWSServiceManager | undefined {

        if (pimItem.isPimNote()) {
            return EWSNotesManager.getInstance();
        } else if (pimItem.isPimTask()) {
            return EWSTaskManager.getInstance();
        } else if (pimItem.isPimMessage()) {
            return EWSMessageManager.getInstance();
        } else if (pimItem.isPimContact()) {
            return EWSContactsManager.getInstance();
        } else if (pimItem.isPimCalendarItem()) {
            return EWSCalendarManager.getInstance();
        } else if (pimItem.isPimLabel()) {
            // Map a pim label to a manager
            if (pimItem.type === PimLabelTypes.JOURNAL) {
                return EWSNotesManager.getInstance();
            } else if (pimItem.type === PimLabelTypes.TASKS) {
                return EWSTaskManager.getInstance();
            } else if (pimItem.type === PimLabelTypes.MAIL) {
                return EWSMessageManager.getInstance();
            } else if (pimItem.type === PimLabelTypes.CONTACTS) {
                return EWSContactsManager.getInstance();
            } else if (pimItem.type === PimLabelTypes.CALENDAR) {
                return EWSCalendarManager.getInstance();
            }
        } else {
            Logger.getInstance().debug(`Unrecognized ItemType for ${pimItem.unid}`);
        }

        return undefined;
    }

    /**
     * Fetch a PIM item from Keep and return an EWS ItemType subclass object representing the response.
     * @param itemEWSId EWSId of item to fetch.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param request The original SOAP request for the get.
     * @param shape Optional EWS shape describing the fields to be populated in the returned item.
     * @returns An EWS ItemType subclass object 
     */
    static async getItem(itemEWSId: string, userInfo: UserInfo, request: Request, shape?: ItemResponseShapeType): Promise<ItemType | undefined> {
        // Get PimItem Id from EWS itemId
        const [itemId, mailboxId] = getKeepIdPair(itemEWSId);
        
        // Fetch the item from Keep
        let pimItem: PimItem | undefined;
        if (itemId) {
            pimItem = await KeepPimManager.getInstance().getPimItem(itemId, userInfo, mailboxId);
        }

        if (pimItem) {
            // Create a manager based on the item and convert it to an ItemType subclass object
            const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
            if (manager) {
                // TODO: WHat should the default shape be?
                if (shape === undefined) {
                    shape = new ItemResponseShapeType();
                    shape.BaseShape = DefaultShapeNamesType.DEFAULT;
                }
                return manager.pimItemToEWSItem(pimItem, userInfo, request, shape, mailboxId);
            }
        }

        return undefined;
    }

    /**
     * Process updates for an item.
     * @param change The EWS change type object that describes the item and the changes to be made.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param toLabel Optional label of the item's parent folder.
     * @returns An updated item or undefined if the request could not be completed.
     */
    static async updateItem(
        change: ItemChangeType, 
        userInfo: UserInfo, 
        request: Request, 
        toLabel?: PimLabel,
    ): Promise<ItemType | undefined> {

        const itemEWSId = change.ItemId?.Id;

        if (itemEWSId) {
            const [itemId, mailboxId] = getKeepIdPair(itemEWSId);
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            let pimItem;
            if (itemId) {
                pimItem = await KeepPimManager.getInstance().getPimItem(itemId, userInfo, mailboxId);
            }
            if (pimItem) {
                const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                if (manager) {
                    const updatedItem = await manager.updateItem(pimItem, change, userInfo, request, toLabel, mailboxId);
                    if (updatedItem) {
                        return updatedItem;
                    }
                } else {
                    // Throw an item not found error
                    const err: any = new Error();
                    err.ResponseClass = ResponseClassType.ERROR;
                    err.ResponseCode = ResponseCodeType.ERROR_ITEM_NOT_FOUND;
                    err.MessageText = `Item not found: ${itemId}`;
                    throw err;
                }
            }
            else {
                // Throw an item id not found
                const err: any = new Error();
                err.ResponseClass = ResponseClassType.ERROR;
                err.ResponseCode = ResponseCodeType.ERROR_ITEM_NOT_FOUND;
                err.MessageText = `Unid id not found in EWS id: ${itemEWSId}`;
                throw err;
            }

        } else {
            // Throw an error for no itemId
            const err: any = new Error();
            err.ResponseClass = ResponseClassType.ERROR;
            err.ResponseCode = ResponseCodeType.ERROR_ITEM_NOT_FOUND;
            err.MessageText = `Missing ItemId in updateItem`;
            throw err;
        }
        return undefined;
    }

    /**
     * Handle moving an item to a new folder and return an the base shape of the updated item
     * @param itemId UID of the item to move.
     * @param toFolderId ID of the target folder.
     * @param userInfo The user's credentials to be passed to Keep.
     * @returns An updated ItemType object with ID_ONLY shape fields.
     */
    static async moveItemWithResult(
        ewsItemType: ItemType, 
        toFolderId: string, 
        userInfo: UserInfo, 
        request: Request
    ): Promise<ItemType | undefined> {
        const resEWSId = await this.moveItem(ewsItemType, toFolderId, userInfo, request);
        if (resEWSId) {
            const [resId, mailboxId] = getKeepIdPair(resEWSId);
            let updatedPimItem: PimItem | undefined;

            if (resId) {
                updatedPimItem = await KeepPimManager.getInstance().getPimItem(resId, userInfo, mailboxId);
            }
        
            if (updatedPimItem) {
                const itemShape = new ItemResponseShapeType();
                itemShape.BaseShape = DefaultShapeNamesType.ID_ONLY;
                const manager = EWSServiceManager.getInstanceFromPimItem(updatedPimItem);
                return manager?.pimItemToEWSItem(updatedPimItem, userInfo, request, itemShape, mailboxId);
            }
        }
        return undefined;
    }
    /**
     * Handle moving an item to a new folder.
     * @param ewsItemType itemType of the item to move.
     * @param toFolderId ID of the target folder.
     * @param userInfo The user's credentials to be passed to Keep.
     * @returns The unid of the moved item or undefined.
     */
    static async moveItem(ewsItemType: ItemType, toFolderId: string, userInfo: UserInfo, request: Request): Promise<string | undefined> {

        const itemEWSId = ewsItemType.ItemId?.Id;
        if (itemEWSId) {
            const manager = EWSServiceManager.getInstanceFromItem(ewsItemType);
            if (manager) {
                await manager.moveItem(itemEWSId, toFolderId, userInfo);
                return itemEWSId;
            }
        }
        return undefined;
    }

    /**
     * Public functions
     */
    /**
     * This function issues a Keep API call to create a PIM entry based on the EWS item passed in.
     * @param item The EWS item to create
     * @param userInfo The user's credentials to be passed to Keep.
     * @param request The original SOAP request for the create.
     * @param toLabel Optional target label (folder).
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @returns A new EWS item populated with information about the created entry.
     */
    abstract createItem(item: ItemType, userInfo: UserInfo, request: Request, disposition?: MessageDispositionType, toLabel?: PimLabel, mailboxId?: string): Promise<ItemType[]>;

    /**
     * This function will make a Keep API request to delete the item with the passed in itemId.
     * @param item The pim item or unid of the item to delete.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @param hardDelete Indicator if the item should be deleted from trash as well
     */
    abstract deleteItem(item: string | PimItem, userInfo: UserInfo, mailboxId?: string, hardDelete?: boolean): Promise<void>;

    /**
     * This function will make a Keep API PUT request to update an item.
     * TODO:  How should we distinguish between a PUT and a PATCH update?
     * @param pimItem Existing PimItem from Keep being updated.
     * @param change The changes to apply to the item.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param request The original SOAP request for the get.
     * @param toLabel Optional label of the parent folder.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     */
    abstract updateItem(pimItem: PimItem, change: ItemChangeType, userInfo: UserInfo, request: Request, toLabel?: PimLabel, mailboxId?: string): Promise<ItemType | undefined>;

    /**
     * Make a copy of an item
     * @param pimItem The source of the copy.
     * @param toFolderId The folder ID of the target folder.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param request The original SOAP request for the get.
     * @param shape EWS shape describing the information requested for each item.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @returns The EWS item for the new item
     * @throws An error if the copy fails
     */
    public async copyItem(
        pimItem: PimItem, 
        toFolderId: string, 
        userInfo: UserInfo, 
        request: Request, 
        shape: ItemResponseShapeType,
        mailboxId?: string
    ): Promise<ItemType> {
        pimItem.parentFolderIds = [toFolderId];

        const results = await KeepPimMessageManager
            .getInstance()
            .moveMessages(userInfo, toFolderId, undefined, undefined, undefined, [pimItem.unid], mailboxId);

        if (results.copiedIds && results.copiedIds.length > 0) {
            const result = results.copiedIds[0];
            if (result.unid && result.status === 200) {
                try {
                    const newItem = await this.getCopiedPimItem(result.unid, pimItem, userInfo, request, mailboxId);
                    if (newItem) {
                        return await this.pimItemToEWSItem(newItem, userInfo, request, shape, mailboxId, toFolderId); // Success
                    } else {
                        Logger.getInstance().warn(`Copy of item ${pimItem.unid} could not be retrieved`);
                    }
                } catch (err) {
                    Logger.getInstance().warn(`Copy of item ${pimItem.unid} could not be created`);
                }
            }
        }

        Logger.getInstance().error(`Move of item being copied failed: ${util.inspect(results, false, 5)}`);

        throw new Error(`Copy item failed for ${pimItem.unid}`);
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
        return KeepPimManager.getInstance().getPimItem(copiedId, userInfo, mailboxId);
    }

    /**
     * Creates a new pim item. 
     * @param pimItem The item to create.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @returns The new Pim item created with Keep.
     * @throws An error if the create fails
     */
    protected abstract createNewPimItem(pimItem: PimItem, userInfo: UserInfo, mailboxId?: string): Promise<PimItem>;

    /**
     * This function issues a Keep API call to move an item to a new folder.
     * NOTES:  1) Should we take an array of items instead of just the one?
     *         2) Does this even belong in this class?  Should move operations be handled by a label manager class or other common code?
     * @param itemId The EWS id of the item being moved.
     * @param toFolderId The folder ID of the target folder.
     * @param userInfo The user's credentials to be passed to Keep.
     */
    protected async moveItem(itemEWSId: string, toFolderId: string, userInfo: UserInfo): Promise<void> {
        let itemId: string | undefined,
            mailboxId: string | undefined;

        try {
            [itemId, mailboxId] = getKeepIdPair(itemEWSId);
            // Move the item in keep.  
            const moveItems: string[] = [];
            if (itemId) moveItems.push(itemId);
            if (toFolderId) {
                // Since we're creating the item, we don't need to worry about multiple parents so we can just move the item to the right folder
                const results = await KeepPimMessageManager
                    .getInstance()
                    .moveMessages(userInfo, toFolderId, moveItems, undefined, undefined, undefined, mailboxId);

                if (results.movedIds && results.movedIds.length > 0) {
                    const result = results.movedIds[0];
                    if (result.unid !== itemId || result.status !== 200) {
                        Logger.getInstance().error(`move ${itemId} to folder ${toFolderId} failed: ${util.inspect(results, false, 5)}`);
                        throw new Error(`Move of item ${itemId} to folder ${toFolderId} failed`);
                    }
                }

            } else {
                Logger.getInstance().debug(`error moving item: ${itemId}.  Could not find the target folder.`);
                throw new Error(`error moving item: ${itemId}.  Could not find the target folder.`);
            }
        } catch (err) {
            // TODO:  Is this the desired behavior?
            // If there was an error, delete the note on the server since it didn't make it to the correct location
            Logger.getInstance().debug(`error moving item, ${itemId}, to desired folder ${toFolderId}: ${err}`);
            if (itemId) await this.deleteItem(itemId, userInfo, mailboxId);
            throw err;
        }
    }

    /**
     * This function will fetech a group of items using the Keep API and return an
     * array of corresponding EWS items populated with the fields requested in the shape.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param request The original SOAP request for the get.
     * @param shape EWS shape describing the information requested for each item.
     * @param startIndex Optional start index for items when paging.  Defaults to 0.
     * @param count Optional count of objects to request.  Defaults to 512.
     * @param fromLabel The label (folder) we are querying.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @returns An array of EWS items built from the returned PimItems.
     */
    abstract getItems(userInfo: UserInfo, request: Request, shape: ItemResponseShapeType, startIndex?: number, count?: number, fromLabel?: PimLabel, mailboxId?: string): Promise<ItemType[]>;


    /**
     * Internal functions
     */

    /**
     * Getter for the array of fields returned for the default shape for the subclass.
     * @returns An array of UnindexedFieldURIType descriptors corresponding to the fields requested by the default shape.
     */
    protected abstract fieldsForDefaultShape(): UnindexedFieldURIType[];

    /**
     * Getter for the array of fields returned for the all properties shape for the subclass.
     * @returns An array of UnindexedFieldURIType descriptors corresponding to the fields requested by the all properties shape.
     */
    protected abstract fieldsForAllPropertiesShape(): UnindexedFieldURIType[];


    /**
     * Create an EWS ItemType object based on the provided PIM item.
     * @param pimItem The source PimItem from which to build an EWS item.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param request The original SOAP request for the get.
     * @param shape The EWS shape describing the fields that should be populated on the returned object.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @param targetParentFolderId An optional parent folder Id to apply to the returned item.
     * @returns An EWS ItemType subclass object based on the passed PimItem and shape in shape.
     */
    abstract pimItemToEWSItem(pimITem: PimItem, userInfo: UserInfo, request: Request, shape: ItemResponseShapeType, mailboxId?: string, targetFolderId?: string): Promise<ItemType>;

    /**
     * This is a helper function for each manager to implement to convert an EWS item to a PimItem used to communicate with the Keep API.
     * @param ewsItem The source EWS item to convert to a PimItem.
     * @param request The original SOAP request for the get.
     * @param existing Object containing existing fields to apply.
     * @returns A PimItem populated with the EWS item's data.
     */
    abstract pimItemFromEWSItem(ewsItem: ItemType, request: Request, existing?: object): PimItem;

    /**
    * Adds requested properties to an EWS item, based on the shape passed in.
    * @param pimItem The PIM item containing the  properties
    * @param toItem The EWS item where the properties should be set. 
    * @param shape The shape setting used to determine which properties to add. If not specified, only extended properties saved in the PIM item will be added.
    * @param mailboxId SMTP mailbox delegator or delegatee address
    */
    protected addRequestedPropertiesToEWSItem(pimItem: PimItem, toItem: ItemType, shape?: ItemResponseShapeType, mailboxId?: string): void {

        // First include all the fields required by the base shape type
        if (shape) {
            const baseFields = this.defaultPropertiesForShape(shape);
            for (const fieldName of baseFields) {
                this.updateEWSItemFieldValue(toItem, pimItem, fieldName, mailboxId);
            }
        }

        // Now add the reqeusted additional and extended fields
        if (shape?.AdditionalProperties) {
            for (const property of shape.AdditionalProperties.items) {
                if (property instanceof PathToUnindexedFieldType) {
                    this.updateEWSItemFieldValue(toItem, pimItem, property.FieldURI, mailboxId);
                }
                else if (property instanceof PathToExtendedFieldType) {
                    const pimProperty = this.getExtendedField(property, pimItem);
                    if (pimProperty) {
                        addPimExtendedPropertyObjectToEWS(pimProperty, toItem);
                    }
                }
                else if (property instanceof PathToIndexedFieldType) {
                    this.updateEWSIndexedItemFieldValue(toItem, pimItem, property);
                }
                else {
                    Logger.getInstance().error(`Unknown additional property: ${util.inspect(property, false, 5)}`)
                }
            }
        } else if (shape === undefined || shape.BaseShape === DefaultShapeNamesType.ALL_PROPERTIES) {
            // If shape was not passed in, or it is AllProperties, add all exteneded properties to the EWS item
            pimItem.extendedProperties.forEach(property => {
                addPimExtendedPropertyObjectToEWS(property, toItem);
            });
        }
        
        //Add EWS parent folder Id
        if (pimItem.parentFolderIds && pimItem.parentFolderIds[0]) {
            const parentFolderEWSId = getEWSId(pimItem.parentFolderIds[0], mailboxId);
            toItem.ParentFolderId = new FolderIdType(parentFolderEWSId, `ck-${parentFolderEWSId}`);
        }
    }

    /**
     * Update the passed in EWS item's indexed field with information stored in the pimItem. This function is
     * used to populate fields requested in a GetItem/FindItem/SyncX EWS operation. 
     * 
     * Subclasses should override this if the EWS items they support need to handle indexed fields. 
     * 
     * @param item The EWS item to update with a value.
     * @param pimItem The PIM item from which to read the value.
     * @param indexedField The EWS indexed field we are mapping.
     * @param mailboxId SMTP mailbox address of the delegate or non-delegate.
     * @returns True if the index field was process or false if it was not. 
     */
    protected updateEWSIndexedItemFieldValue(item: ItemType, pimItem: PimItem, indexedField: PathToIndexedFieldType, mailboxId?: string): boolean {

        // Add common indexed fileds here
        
        return false; 
    }

    /**
     * Get the value for an extended field for a Pim item. 
     * 
     * Subclasses should override this to pull the value from the Pim item data if pimItem.findExtendedProperty returns undefined.
     * 
     * @param path The extended field item. 
     * @param pimItem The PIM item from which to read the value.
     * @returns The properties and value for the extended field or undefined if there is no value
     */
    protected getExtendedField(path: PathToExtendedFieldType, pimItem: PimItem): any | undefined {

        // Add common extended fileds here
        
        return pimItem.findExtendedProperty(identifiersForPathToExtendedFieldType(path)); 
    }

    /**
    * Update the passed in EWS item's specific field with the information stored in the pimItem.  This function is
    * used to populate fields requested in a GetItem/FindItem/SyncX EWS operation.  It will process common fields
    * and return true if the request was handled.
    * @param item The EWS item to update with a value.
    * @param pimItem The PIM item from which to read the value.
    * @param fieldIdentifier The EWS field we are mapping.
    * @param mailboxId SMTP mailbox delegator or delegatee address.
    * 
    * Common fields that are processed by this function are:
    * UnindexedFieldURIType.ITEM_ITEM_ID
    * UnindexedFieldURIType.ITEM_PARENT_FOLDER_ID
    * UnindexedFieldURIType.ITEM_ITEM_CLASS
    * UnindexedFieldURIType.ITEM_DATE_TIME_CREATED
    * UnindexedFieldURIType.ITEM_MIME_CONTENT -- Not currently supported.  Should move to type specific processing.
    * UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS
    * UnindexedFieldURIType.ITEM_ATTACHMENTS
    * UnindexedFieldURIType.ITEM_SUBJECT
    * UnindexedFieldURIType.ITEM_CATEGORIES
    * UnindexedFieldURIType.ITEM_IMPORTANCE
    * UnindexedFieldURIType.ITEM_IN_REPLY_TO -- Not currently supported.  Should move to type specific processing.
    * UnindexedFieldURIType.ITEM_IS_DRAFT
    * UnindexedFieldURIType.ITEM_BODY
    * UnindexedFieldURIType.ITEM_UNIQUE_BODY -- Not currently processed.
    * UnindexedFieldURIType.ITEM_NORMALIZED_BODY -- Not currently processed.
    * UnindexedFieldURIType.ITEM_TEXT_BODY -- Not currently processed.
    * UnindexedFieldURIType.ITEM_SENSITIVITY
    * UnindexedFieldURIType.ITEM_REMINDER_DUE_BY
    * UnindexedFieldURIType.ITEM_REMINDER_NEXT_TIME
    * UnindexedFieldURIType.ITEM_REMINDER_IS_SET
    * UnindexedFieldURIType.ITEM_REMINDER_MINUTES_BEFORE_START
    * UnindexedFieldURIType.ITEM_LAST_MODIFIED_TIME
    * UnindexedFieldURIType.ITEM_SIZE -- Need to add size to all Pim* objects
    * UnindexedFieldURIType.MESSAGE_IS_READ -- Should move to message-specific processing.
    * UnindexedFieldURIType.ITEM_EFFECTIVE_RIGHTS -- Currently returning full access
    */
    protected updateEWSItemFieldValue(item: ItemType, pimItem: PimItem, fieldIdentifier: UnindexedFieldURIType, mailboxId?: string): boolean {

        if (fieldIdentifier === UnindexedFieldURIType.ITEM_ITEM_ID) {
            const itemEWSId = getEWSId(pimItem.unid, mailboxId);
            item.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_PARENT_FOLDER_ID) {
            if (pimItem.parentFolderIds) {
                // If there are multiple parents, we can only take one for ews 
                const pFolderId = pimItem.parentFolderIds.length > 0 ? pimItem.parentFolderIds[0] : undefined;
                if (pFolderId) {
                    const parentFolderEWSId = getEWSId(pFolderId, mailboxId);
                    item.ParentFolderId = new FolderIdType(parentFolderEWSId, `ck-${parentFolderEWSId}`);
                }
            }
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_ITEM_CLASS) {
            item.ItemClass = ewsItemClassType(pimItem);
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_MIME_CONTENT) {
            // TODO: Implement with LABS-506
            Logger.getInstance().error(`Fetching ${fieldIdentifier} is not supported`);
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_DATE_TIME_CREATED) {
            item.DateTimeCreated = getFallbackCreatedDate(pimItem);
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS) {
            const pAttachments = getItemAttachments(pimItem);
            item.HasAttachments = pAttachments ? true : false;
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_ATTACHMENTS) {
            const pAttachments = getItemAttachments(pimItem);
            if (pAttachments) {
                item.Attachments = pAttachments;
            }
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_SUBJECT) {
            item.Subject = pimItem.subject === undefined ? "" : pimItem.subject;
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_CATEGORIES) {
            if (pimItem.categories.length > 0) {
                const categories = new ArrayOfStringsType();
                categories.String = pimItem.categories;
                item.Categories = categories;
            }
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_IMPORTANCE) {
            if (pimItem.isPimCalendarItem()) {
                // Calendar items only support HIGH or undefined
                if (pimItem.importance !== undefined && pimItem.importance === PimImportance.HIGH) {
                    item.Importance = ImportanceChoicesType.HIGH
                }
            }
            else if (pimItem.isPimTask()) {
                // Tasks use the priority field in PIM and support HIGH, NORMAL and LOW  in EWS
                switch (pimItem.priority) {
                    case PimTaskPriority.HIGH:
                        item.Importance = ImportanceChoicesType.HIGH;
                        break;
                    case PimTaskPriority.MEDIUM:
                        item.Importance = ImportanceChoicesType.NORMAL;
                        break;
                    case PimTaskPriority.LOW:
                        item.Importance = ImportanceChoicesType.LOW;
                        break;
                    default:
                        item.Importance = ImportanceChoicesType.NORMAL;
                        break;
                }
            }
            else {
                Logger.getInstance().warn(`Unexpected PIM item type ${pimItem.constructor.name} for field ${fieldIdentifier} in base updateEWSItemFieldValue`);
                // Let the caller handle this one
                return false;
            }
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_IN_REPLY_TO) {
            // TODO: Implement with LABS-506...To what types of items does this apply?
            Logger.getInstance().error(`Fetching ${fieldIdentifier} is not supported`);
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_IS_DRAFT) {
            if (pimItem.isPimCalendarItem()) {
                // Just leave it not set if false
                if (pimItem.isDraft) {
                    item.IsDraft = true;
                }
            }
            else {
                Logger.getInstance().error(`Unexpected PIM item type ${pimItem.constructor.name} for field ${fieldIdentifier} in base updateEWSItemFieldValue`);
                // Let the caller handle this one
                return false;
            }
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_BODY) {
            if (pimItem.body) {
                // Defaulting to HTML body type as PimItem does not have a bodyType property.
                item.Body = new BodyType(pimItem.body, hasHTML(pimItem.body) ? BodyTypeType.HTML : BodyTypeType.TEXT);
            }
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_UNIQUE_BODY || fieldIdentifier === UnindexedFieldURIType.ITEM_NORMALIZED_BODY ||
            fieldIdentifier === UnindexedFieldURIType.ITEM_TEXT_BODY) {
            // TODO:  How do we handle these body fields?
            Logger.getInstance().error(`Fetching ${fieldIdentifier} is not supported`);
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_SENSITIVITY) {
            if (pimItem.isPrivate !== undefined) {
                if (pimItem.isConfidential) {
                    item.Sensitivity = SensitivityChoicesType.CONFIDENTIAL;
                } else if (pimItem.isPrivate) {
                    item.Sensitivity = SensitivityChoicesType.PRIVATE;
                } else {
                    item.Sensitivity = SensitivityChoicesType.NORMAL;
                }
            }
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_REMINDER_DUE_BY || fieldIdentifier === UnindexedFieldURIType.ITEM_REMINDER_NEXT_TIME) {
            if (pimItem.isPimCalendarItem()) {
                if (pimItem.alarm) {
                    item.ReminderIsSet = true;
                    item.ReminderMinutesBeforeStart = pimItem.alarm < 0 ? `${Math.abs(pimItem.alarm)}` : "0";
                    if (pimItem.start) {
                        item.ReminderDueBy = new Date(pimItem.start);
                    }
                }
                else {
                    item.ReminderIsSet = false;
                }
            }
            else if (pimItem.isPimTask() && item instanceof TaskType) {
                // Need the due date in order to compute the reminder
                if (pimItem.due) {
                    item.DueDate = new Date(pimItem.due);
                }

                item.ReminderIsSet = (pimItem.alarm !== undefined && item.DueDate !== undefined) ? true : false;
                if (item.ReminderIsSet) {
                    item.ReminderMinutesBeforeStart = pimItem.alarm ? `${Math.abs(pimItem.alarm)}` : '0';
                    item.ReminderDueBy = item.DueDate;
                }
            }
            else {
                Logger.getInstance().error(`Unexpected PIM item type ${pimItem.constructor.name} for field ${fieldIdentifier}`);
                // Let the caller handle this one
                return false;
            }
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_REMINDER_IS_SET) {
            if (pimItem.isPimCalendarItem() || pimItem.isPimTask()) {
                item.ReminderIsSet = pimItem.alarm ? true : false;
            }
            else {
                Logger.getInstance().error(`Unexpected PIM item type ${pimItem.constructor.name} for field ${fieldIdentifier} in base updateEWSItemFieldValue`);
                // Let the caller handle this one
                return false;
            }
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_REMINDER_MINUTES_BEFORE_START) {
            if (pimItem.isPimCalendarItem() || pimItem.isPimTask()) {
                if (pimItem.alarm) {
                    item.ReminderMinutesBeforeStart = pimItem.alarm < 0 ? `${Math.abs(pimItem.alarm)}` : "0";
                }
            }
            else {
                Logger.getInstance().error(`Unexpected PIM item type ${pimItem.constructor.name} for field ${fieldIdentifier} in base updateEWSItemFieldValue`);
                // Let the caller handle this one
                return false;
            }
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_LAST_MODIFIED_TIME) {
            if (pimItem.lastModifiedDate) {
                item.LastModifiedTime = pimItem.lastModifiedDate;
            }
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_EFFECTIVE_RIGHTS) {
            item.EffectiveRights = effectiveRightsForItem(pimItem);
        }
        else {
            // Not processed
            return false;
        }
        return true;
    }

    /**
    * Return an array of UnindexedFieldURIType strings indicating the default fields requested for the BaseShape 
    * of the shape that is passed in with respect to this object's classes EWS objects.
    * See https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/property-sets-and-response-shapes-in-ews-in-exchange for details.
    * @param shape The shape describing what fields are being requested.
    * @returns An array of URI types indicating the default fields returned for the shape.
    */
    protected defaultPropertiesForShape(shape: ItemResponseShapeType): UnindexedFieldURIType[] {
        switch (shape.BaseShape) {
            case DefaultShapeNamesType.ID_ONLY:
                return EWSServiceManager.idOnlyFields;
            case DefaultShapeNamesType.DEFAULT:
                return this.fieldsForDefaultShape();
            case DefaultShapeNamesType.ALL_PROPERTIES:
                return this.fieldsForAllPropertiesShape();
            default:
                return [];
        }
    }

    /**
     * Update a value for an EWS field in a Keep PIM item. 
     * @param pimItem The pim item that will be updated.
     * @param fieldIdentifier The EWS unindexed field identifier
     * @param newValue The new value to set. The type is based on what fieldIdentifier is set to. To delete a field, pass in undefined. 
     * @returns true if it was handled 
     */
    protected updatePimItemFieldValue(pimItem: PimItem, fieldIdentifier: UnindexedFieldURIType, newValue: any | undefined): boolean {
        if (fieldIdentifier === UnindexedFieldURIType.ITEM_PARENT_FOLDER_ID) {
            const parent: FolderIdType = newValue;
            pimItem.parentFolderIds = (newValue === undefined) ? [] : [parent.Id];
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_MIME_CONTENT) {
            // TODO: Implement with LABS-506
            //Logger.getInstance().error(`Updating ${fieldIdentifier} is not supported`);
            if(newValue instanceof MimeContentType){
                pimItem.mimeUpdate =  newValue.Value;
             }
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_ATTACHMENTS) {
            if (newValue === undefined) {
                pimItem.attachments = [];
            }
            else {
                // TODO: Implement with LABS-555
                Logger.getInstance().error(`Updating ${fieldIdentifier} is not supported`);
            }
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_SUBJECT) {
            pimItem.subject = newValue === undefined ? "" : newValue;
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_CATEGORIES) {
            if (newValue === undefined) {
                pimItem.categories = [];
            }
            else {
                const categories: ArrayOfStringsType = newValue;
                pimItem.categories = categories.String;
            }
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_IMPORTANCE) {
            if (pimItem.isPimMessage()) {
                if (newValue === undefined) {
                    pimItem.importance = PimImportance.NONE;
                }
                else {
                    switch (newValue) {
                        case ImportanceChoicesType.HIGH:
                            pimItem.importance = PimImportance.HIGH;
                            break;
                        case ImportanceChoicesType.LOW:
                            pimItem.importance = PimImportance.LOW;
                            break;
                        default:
                            pimItem.importance = PimImportance.NONE
                    }
                }
            }
            else if (pimItem.isPimCalendarItem()) {
                if (newValue === undefined) {
                    pimItem.importance = PimImportance.NONE;
                }
                else if (newValue === ImportanceChoicesType.HIGH) {
                    pimItem.importance = PimImportance.HIGH;
                }
                else {
                    pimItem.importance = PimImportance.NONE;
                }
            }
            else if (pimItem.isPimTask()) {
                if (newValue === undefined) {
                    pimItem.priority = PimTaskPriority.NONE;
                }
                else {
                    switch (newValue) {
                        case ImportanceChoicesType.HIGH:
                            pimItem.priority = PimTaskPriority.HIGH;
                            break;
                        case ImportanceChoicesType.LOW:
                            pimItem.priority = PimTaskPriority.LOW;
                            break;
                        default:
                            pimItem.priority = PimTaskPriority.NONE;
                    }
                }
            }
            else {
                Logger.getInstance().error(`Unexpected PIM item type ${pimItem.constructor.name} for field ${fieldIdentifier}`)
            }
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_IN_REPLY_TO) {
            // TODO: Implement with LABS-506
            Logger.getInstance().error(`Updating ${fieldIdentifier} is not supported`);
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_IS_DRAFT) {
            if (pimItem.isPimCalendarItem()) {
                pimItem.isDraft = newValue === undefined ? false : newValue;
            }
            else {
                Logger.getInstance().error(`Unexpected PIM item type ${pimItem.constructor.name} for field ${fieldIdentifier}`)
            }
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_UNIQUE_BODY || fieldIdentifier === UnindexedFieldURIType.ITEM_NORMALIZED_BODY ||
            fieldIdentifier === UnindexedFieldURIType.ITEM_BODY || fieldIdentifier === UnindexedFieldURIType.ITEM_TEXT_BODY) {
            if (newValue === undefined) {
                pimItem.body = "";
                pimItem.bodyType = undefined;
            }
            else if (newValue instanceof BodyType) {
                pimItem.bodyType = newValue.BodyType;
                pimItem.body = newValue.Value;
            }
            else {
                Logger.getInstance().error(`Unexpected body value for PIM item ${pimItem.unid}`)
            }
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_SENSITIVITY) {
            if (pimItem instanceof PimRule) {
                pimItem.sensitivity = newValue ?? PimRule.DefaultSensitivity;
            }
            else {
                pimItem.isPrivate = newValue === SensitivityChoicesType.PRIVATE ? true : false;
            }
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_REMINDER_DUE_BY || fieldIdentifier === UnindexedFieldURIType.ITEM_REMINDER_NEXT_TIME) {
            if (pimItem.isPimCalendarItem()) {
                if (newValue === undefined) {
                    pimItem.alarm = undefined;
                }
                else if (pimItem.start) {
                    const startDate = new Date(pimItem.start);
                    pimItem.alarm = -(startDate.getTime() - newValue.getTime()) * 1000 * 60; // Convert to minutes
                }
            }
            else if (pimItem.isPimTask()) {
                if (newValue === undefined) {
                    pimItem.alarm = undefined;
                    pimItem.due = undefined;
                }
                else if (pimItem.start) {
                    const startDate = new Date(pimItem.start);
                    pimItem.alarm = -(startDate.getTime() - newValue.getTime()) * 1000 * 60; // Convert to minutes
                    pimItem.due = newValue;
                }
            }
            else {
                Logger.getInstance().error(`Unexpected PIM item type ${pimItem.constructor.name} for field ${fieldIdentifier}`)
            }
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_REMINDER_IS_SET) {
            if (pimItem.isPimCalendarItem() || pimItem.isPimTask()) {
                if (newValue === undefined || newValue === false) {
                    pimItem.alarm = undefined;
                }
            }
            else {
                Logger.getInstance().error(`Unexpected PIM item type ${pimItem.constructor.name} for field ${fieldIdentifier}`)
            }
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_REMINDER_MINUTES_BEFORE_START) {
            if (pimItem.isPimCalendarItem() || pimItem.isPimTask()) {
                pimItem.alarm = -newValue;
            }
            else {
                Logger.getInstance().error(`Unexpected PIM item type ${pimItem.constructor.name} for field ${fieldIdentifier}`)
            }
        }
        else if (fieldIdentifier === UnindexedFieldURIType.ITEM_LAST_MODIFIED_TIME) {
            pimItem.lastModifiedDate = newValue === undefined ? undefined : new Date(newValue);
        }
        else if (fieldIdentifier === UnindexedFieldURIType.MESSAGE_IS_READ) {
            if (pimItem.isPimMessage()) {
                pimItem.isRead = newValue ?? false;
            }
            else {
                Logger.getInstance().error(`Unexpected PIM item type ${pimItem.constructor.name} for field ${fieldIdentifier}`)
            }
        }
        else {
            return false;
        }

        return true;

        /*
        The following item fields are not currently supported:

        ITEM_ITEM_ID = 'item:ItemId',
        ITEM_ITEM_CLASS = 'item:ItemClass',
        ITEM_SUBJECT = 'item:Subject',
        ITEM_DATE_TIME_RECEIVED = 'item:DateTimeReceived',
        ITEM_SIZE = 'item:Size',
        ITEM_HAS_ATTACHMENTS = 'item:HasAttachments',
        ITEM_INTERNET_MESSAGE_HEADERS = 'item:InternetMessageHeaders',
        ITEM_IS_ASSOCIATED = 'item:IsAssociated',
        ITEM_IS_FROM_ME = 'item:IsFromMe',
        ITEM_IS_RESEND = 'item:IsResend',
        ITEM_IS_SUBMITTED = 'item:IsSubmitted',
        ITEM_IS_UNMODIFIED = 'item:IsUnmodified',
        ITEM_DATE_TIME_SENT = 'item:DateTimeSent',
        ITEM_RESPONSE_OBJECTS = 'item:ResponseObjects',
        ITEM_DISPLAY_TO = 'item:DisplayTo',
        ITEM_DISPLAY_CC = 'item:DisplayCc',
        ITEM_DISPLAY_BCC = 'item:DisplayBcc',
        ITEM_CULTURE = 'item:Culture',
        ITEM_EFFECTIVE_RIGHTS = 'item:EffectiveRights',
        ITEM_LAST_MODIFIED_NAME = 'item:LastModifiedName',
        ITEM_CONVERSATION_ID = 'item:ConversationId',
        ITEM_FLAG = 'item:Flag',
        ITEM_STORE_ENTRY_ID = 'item:StoreEntryId',
        ITEM_INSTANCE_KEY = 'item:InstanceKey',
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
        ITEM_DATE_TIME_CREATED = 'time:DateTimeCreated'
        */
    }


    /**
     * Convert jsCalendar rules in a recurring Pim item to EWS recurrence type. 
     * 
     * EWS recurrence type is documented at [Recurrence (RecurrenceType)](https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/recurrence-recurrencetype)
     * and [Recurrence (TaskRecurrenceType)](https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/recurrence-taskrecurrencetype). 
     * 
     * Note: TaskRecurrenceType defines YearlyRegeneration, MonthlyRegeneration, WeeklyRegeneration, DailyRegeneration which seem to be 
     * alternatives to *YearlyRecurrence, *MonthlyRecurrence, WeeklyRecurrence, and DailyRecurrence. So YearlyRegeneration, MonthlyRegeneration, WeeklyRegeneration, DailyRegeneration
     * will not be added by this method. If they are needed in the furture this method should be overriden in EWSTaskManager and they should be implemented there. 
     *    
     * @param pimItem A pim item that supports recurrence (calendar item or task). 
     * @returns A base EWS recurrence structure. Rules from the pimItem will be added to this structure. If the item is not recurring, undefined will be returned. 
     * @throws An error if the jsCalendar rules can't be converted to EWS recurrence type. 
     */
    protected getEWSRecurrence(pimItem: PimCommonEventsJmap): RecurrenceType | undefined {

        let startString = pimItem.start;
        if (pimItem.isPimTask() && startString === undefined) {
            // If task does not have a start date, recurrence applies to due date. 
            startString = pimItem.due;
        }

        if (startString === undefined) {
            // Recurrence requires a start date
            Logger.getInstance().warn(`PIM item ${pimItem.unid} does not contain a starting date for the recurrence`);
            return undefined;
        }

        if (pimItem.recurrenceRules !== undefined) {

            if (pimItem.recurrenceRules?.length !== 1) {
                throw new Error(`Unable to calculate recurrence for PIM item ${pimItem.unid} because multiple recurrence rules are not supported`);
            }

            const rule = pimItem.recurrenceRules[0];

            if (rule.byHour || rule.byMinute || rule.bySecond || rule.byWeekNo || rule.byYearDay) {
                Logger.getInstance().error(`PIM item ${pimItem.unid} rule contains unsupported elements: ${util.inspect(rule, false, 10)}`);
                throw new Error(`Unable to calculate recurrence for PIM item ${pimItem.unid} because the rule is not supported in EWS`);
            }

            if (rule.bySetPosition && rule.bySetPosition.length > 1) {
                Logger.getInstance().error(`PIM item ${pimItem.unid} rule contains unsupported bySetPosition: ${util.inspect(rule, false, 10)}`);
                throw new Error(`Unable to calculate recurrence for PIM item ${pimItem.unid} because multiple bySetPostion values is not supported in EWS`);
            }

            const recurrence = new RecurrenceType();

            /*
             The following code is based of the published rules in the [iCalendar to Appointment Object Conversion Algorithm](https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcical/a685a040-5b69-4c84-b084-795113fb4012)  
             using the recurrence templates defined in section 2.1.3.2.2. 

             NOTE: Domino does not support byhour or byminute recurrence, so byhour-part and byminute-part are ignored.

                common-parts = [interval-part] [byhour-part] [byminute-part]
                                [(until-part / count-part)] [wkst-part]

                interval-part       = ";INTERVAL=" 1*DIGIT          ; See 2.3.1.2
                byminute-part       = ";BYMINUTE=" 1*2DIGIT          ; See 2.3.1.3
                byhour-part         = ";BYHOUR=" 1*2DIGIT          ; See 2.3.1.4
                bymonthday-part     = ";BYMONTHDAY=" ["-"]1*2DIGIT     ; See 2.3.1.5
                byday-part          = ";BYDAY=" byday-list          ; See 2.3.1.6
                byday-nth-part      = ";BYDAY=" byday-nth-list     ; See 2.3.1.6
                bymonth-part        = ";BYMONTH=" 1*2DIGIT          ; See 2.3.1.7
                bysetpos-part       = ";BYSETPOS=" weeknum          ; See 2.3.1.8
                wkst-part           = ";WKST=" dayofweek          ; See 2.3.1.9
                until-part          = ";UNTIL=" datetime          ; See 2.3.1.10
                count-part          = ";COUNT=" 1*3DIGIT          ; See 2.3.1.11
                
                byday-list          = byday-elm *("," byday-elm)
                byday-elm           = [weeknum] dayofweek
                dayofweek           = "SU" / "MO" / "TU" / "WE" / "TH" / "FR" / "SA"
                
                weeknum             = "-1" / "1" / "2" / "3" / "4"
                byday-nth-list      = dayofweek
                byday-nth-list      =/ "SA,SU"                    ; Any ordering
                byday-nth-list      =/ "MO,TU,WE,TH,FR"          ; Any ordering
                byday-nth-list      =/ "SU,MO,TU,WE,TH,FR,SA"          ; Any ordering
                datetime            = year month day "T" hour minute second ["Z"]
                year                = 4DIGIT
                month               = 2DIGIT
                day                 = 2DIGIT
                hour                = 2DIGIT
                minute              = 2DIGIT
                second              = "00"

             */

            switch (rule.frequency) {
                /*
                 Supported Yearly Recurrence Templates (see above):
                    yearly-template= "FREQ=YEARLY" [bymonthday-part] [bymonth-part] [common-parts]
                    yearlynth-template= "FREQ=YEARLY" byday-nth-part bysetpos-part bymonth-part [common-parts] 
                 */
                case PimRecurrenceFrequency.YEARLY:
                    /*
                    <AbsoluteYearlyRecurrence>
                        <DayOfMonth/>
                        <Month/>
                    </AbsoluteYearlyRecurrence>
                    */
                    if (rule.byMonthDay && rule.byMonthDay.length === 1 &&
                        rule.byMonth && rule.byMonth.length === 1) {
                        recurrence.AbsoluteYearlyRecurrence = new AbsoluteYearlyRecurrencePatternType();
                        recurrence.AbsoluteYearlyRecurrence.DayOfMonth = rule.byMonthDay[0];
                        recurrence.AbsoluteYearlyRecurrence.Month = this.byMonthToEWSMonth(rule.byMonth[0]);
                    }
                    /*
                     <RelativeYearlyRecurrence>
                        <DaysOfWeek/>
                        <DayOfWeekIndex/>
                        <Month/>
                     </RelativeYearlyRecurrence>
                    */
                    else if (rule.byDay && rule.byMonth && rule.byMonth.length === 1) {
                        recurrence.RelativeYearlyRecurrence = new RelativeYearlyRecurrencePatternType();
                        recurrence.RelativeYearlyRecurrence.DaysOfWeek = this.byDayToEWSDay(rule);
                        recurrence.RelativeYearlyRecurrence.Month = this.byMonthToEWSMonth(rule.byMonth[0]);
                        let index = 1; // Default to first
                        if (rule.bySetPosition && rule.bySetPosition.length === 1) {
                            index = rule.bySetPosition[0];
                        }
                        else if (rule.byDay.length === 1 && rule.byDay[0].nthOfPeriod !== undefined) {
                            index = rule.byDay[0].nthOfPeriod;
                        }
                        recurrence.RelativeYearlyRecurrence.DayOfWeekIndex = this.instanceToEWSDayOfWeekIndex(index);
                    }
                    else {
                        Logger.getInstance().error(`Unable to convert yearly recurrence rule for PIM item ${pimItem.unid}: ${util.inspect(rule, false, 5)}`);
                        throw new Error(`Unable to convert yearly recurrence rule for PIM item ${pimItem.unid}`);
                    }

                    break;

                /*
                 Supported Monthly Recurrence Templates (see above):
                    monthly-template= "FREQ=MONTHLY" [bymonthday-part] [common-parts]
                    monthlynth-template= "FREQ=MONTHLY" byday-nth-part bysetpos-part [common-parts]
                 */
                case PimRecurrenceFrequency.MONTHLY:
                    /* 
                     <AbsoluteMonthlyRecurrence>
                        <Interval/>
                        <DayOfMonth/>
                     </AbsoluteMonthlyRecurrence>
                     */
                    if (rule.byMonthDay && rule.byMonthDay.length === 1) {
                        recurrence.AbsoluteMonthlyRecurrence = new AbsoluteMonthlyRecurrencePatternType();
                        recurrence.AbsoluteMonthlyRecurrence.DayOfMonth = rule.byMonthDay[0];
                        recurrence.AbsoluteMonthlyRecurrence.Interval = rule.interval;
                    }
                    /*
                     <RelativeMonthlyRecurrence>
                        <Interval/>
                        <DaysOfWeek/>
                        <DayOfWeekIndex/>
                     </RelativeMonthlyRecurrence>
                     */
                    else if (rule.byDay) {
                        recurrence.RelativeMonthlyRecurrence = new RelativeMonthlyRecurrencePatternType();
                        recurrence.RelativeMonthlyRecurrence.DaysOfWeek = this.byDayToEWSDay(rule); // TODO: See LABS-2448
                        recurrence.RelativeMonthlyRecurrence.Interval = rule.interval;
                        let index = 1; // Default to first
                        if (rule.bySetPosition && rule.bySetPosition.length === 1) {
                            index = rule.bySetPosition[0];
                        }
                        else if (rule.byDay.length === 1 && rule.byDay[0].nthOfPeriod !== undefined) {
                            index = rule.byDay[0].nthOfPeriod;
                        }
                        recurrence.RelativeMonthlyRecurrence.DayOfWeekIndex = this.instanceToEWSDayOfWeekIndex(index);
                    }
                    else {
                        Logger.getInstance().error(`PIM item ${pimItem.unid} is missing byMonthDay or byDay in monthly recurrence rule: ${util.inspect(rule, false, 5)}`);
                        throw new Error(`PIM item ${pimItem.unid} is missing byMonthDay or byDay in monthly recurrence rule`);
                    }
                    break;

                /*
                 Supported Weekly Recurrence Template (see above):
                    weekly-template= "FREQ=WEEKLY" [byday-part] [common-parts]
                 */
                case PimRecurrenceFrequency.WEEKLY:
                    /*
                     <WeeklyRecurrence>
                        <Interval/>
                        <DaysOfWeek/>
                        <FirstDayOfWeek/>
                     </WeeklyRecurrence>
                     */
                    recurrence.WeeklyRecurrence = new WeeklyRecurrencePatternType();
                    recurrence.WeeklyRecurrence.FirstDayOfWeek = this.pimDaytoEWSDay(rule.firstDayOfWeek);
                    recurrence.WeeklyRecurrence.Interval = rule.interval;
                    if (rule.byDay) {
                        // NOTE: The iCalendar to Appointment Object Conversion Algorithm indicates that byDay entries can have a week number (represented with nthOfPeriod in PimRecurrenceNDay). 
                        // However there is no way to set this in the EWS WeeklyRecurrence.
                        recurrence.WeeklyRecurrence.DaysOfWeek = new DaysOfWeekType(rule.byDay.map(NDay => this.pimDaytoEWSDay(NDay.day)));
                    }
                    else {
                        Logger.getInstance().error(`PIM item ${pimItem.unid} is missing byDay in weekly recurrence rule: ${util.inspect(rule, false, 5)}`);
                        throw new Error(`PIM item ${pimItem.unid} is missing byDay in weekly recurrence rule`);
                    }
                    break;

                /*
                 Supported Daily Recurrence Template (see above):
                    daily-template= "FREQ=DAILY" [common-parts]
                 */
                case PimRecurrenceFrequency.DAILY:
                    /*
                     <DailyRecurrence>
                        <Interval/>
                     </DailyRecurrence>
                     */
                    recurrence.DailyRecurrence = new DailyRecurrencePatternType();
                    recurrence.DailyRecurrence.Interval = rule.interval;
                    break;

                default:
                    Logger.getInstance().error(`PIM item ${pimItem.unid} has unknown recurrence frequency in recurrence rule: ${util.inspect(rule, false, 5)}`);
                    throw new Error(`PIM item ${pimItem.unid} has unknown recurrence frequency in rule`);
            }

            // Set start and end of recurrence
            if (rule.until) {
                // Ends on a specific date
                recurrence.EndDateRecurrence = new EndDateRecurrenceRangeType();
                recurrence.EndDateRecurrence.StartDate = new Date(startString);

                // Apply timezone to until date
                let st = DateTime.fromISO(rule.until);
                if (pimItem.startTimeZone) {
                    st = st.setZone(pimItem.startTimeZone, { keepLocalTime: true });
                }
                // Don't include the timezone if it is local. This means no startTimeZone set. 
                const untilString = st.toISO({ includeOffset: st.zone.type === 'local' ? false : true }) ?? rule.until;
                recurrence.EndDateRecurrence.EndDate = new Date(untilString);
            }
            else if (rule.count) {
                // Ends after a number of recurrences
                recurrence.NumberedRecurrence = new NumberedRecurrenceRangeType();
                recurrence.NumberedRecurrence.NumberOfOccurrences = rule.count;
                recurrence.NumberedRecurrence.StartDate = new Date(startString);
            }
            else {
                // No end date
                recurrence.NoEndRecurrence = new NoEndRecurrenceRangeType();
                recurrence.NoEndRecurrence.StartDate = new Date(startString);
            }

            return recurrence;
        }
        else {
            Logger.getInstance().warn(`PIM item ${pimItem.unid} does not contain recurrence rules`);
        }

        return undefined;

    }

    /**
     * Convert a jsCalendar byMonth value to a EWS month value.
     * @param byMonth The jsCalendar byMonth value
     * @returns The EWS month name
     * @throws An error if the jsCalendar byMonth value can't be converted to an EWS month
     */
    protected byMonthToEWSMonth(byMonth: string): MonthNamesType {
        if (byMonth.length === 0) {
            throw new Error(`byMonth '${byMonth}' is not valid`);
        }

        const monthNames: MonthNamesType[] = [
            MonthNamesType.JANUARY, MonthNamesType.FEBRUARY, MonthNamesType.MARCH, MonthNamesType.APRIL,
            MonthNamesType.MAY, MonthNamesType.JUNE, MonthNamesType.JULY, MonthNamesType.AUGUST,
            MonthNamesType.SEPTEMBER, MonthNamesType.OCTOBER, MonthNamesType.NOVEMBER, MonthNamesType.DECEMBER
        ];

        let monthString = byMonth;
        if (monthString.slice(-1) === 'L') {
            monthString = monthString.substring(0, monthString.length - 1);
            // TODO: LABS-2105 What do do with Leap Month?
        }
        const monthIndex = Number.parseInt(monthString) - 1; // Convert to zero based number
        if (monthIndex < 0 && monthIndex >= monthNames.length) {
            throw new Error(`Unable to convert byMonth '${byMonth}' to EWS month name. Month index ${monthIndex} out of range.`);
        }

        return monthNames[monthIndex];

    }

    /**
     * Converts a pim day to an EWS day.
     * @param day The Pim day 
     * @returns The EWS day
     */
    protected pimDaytoEWSDay(day: PimRecurrenceDayOfWeek): DayOfWeekType {
        const dayOfWeekMap: any = {};
        dayOfWeekMap[PimRecurrenceDayOfWeek.MONDAY] = DayOfWeekType.MONDAY;
        dayOfWeekMap[PimRecurrenceDayOfWeek.TUESDAY] = DayOfWeekType.TUESDAY;
        dayOfWeekMap[PimRecurrenceDayOfWeek.WEDNESDAY] = DayOfWeekType.WEDNESDAY;
        dayOfWeekMap[PimRecurrenceDayOfWeek.THURSDAY] = DayOfWeekType.THURSDAY;
        dayOfWeekMap[PimRecurrenceDayOfWeek.FRIDAY] = DayOfWeekType.FRIDAY;
        dayOfWeekMap[PimRecurrenceDayOfWeek.SATURDAY] = DayOfWeekType.SATURDAY;
        dayOfWeekMap[PimRecurrenceDayOfWeek.SUNDAY] = DayOfWeekType.SUNDAY;

        return dayOfWeekMap[day];
    }

    /**
     * Convert a byDay setting in a recurrence rule to the EWS day of week type.
     * @param rule The recurrence rule containing the byDay setting.
     * @returns The EWS day of the week representing the byDay setting.
     * @throws An error if byDay is not set in the recurrence rule or the byDay value is not supported by EWS. 
     */
    protected byDayToEWSDay(rule: PimRecurrenceRule): DayOfWeekType {

        const byDay = rule.byDay;
        if (byDay) {

            if (byDay.length === 1) {
                return this.pimDaytoEWSDay(byDay[0].day);
            }
            else {
                const pimDays = byDay.map(NDay => NDay.day);
                const weekend = [PimRecurrenceDayOfWeek.SATURDAY, PimRecurrenceDayOfWeek.SUNDAY];
                const weekday = [PimRecurrenceDayOfWeek.MONDAY, PimRecurrenceDayOfWeek.TUESDAY, PimRecurrenceDayOfWeek.WEDNESDAY, PimRecurrenceDayOfWeek.THURSDAY, PimRecurrenceDayOfWeek.FRIDAY];
                const allDays = [...weekday, ...weekend];

                if (weekend.filter(day => !pimDays.includes(day)).concat(pimDays.filter(day => !weekend.includes(day))).length === 0) {
                    return DayOfWeekType.WEEKEND_DAY;
                }
                else if (weekday.filter(day => !pimDays.includes(day)).concat(pimDays.filter(day => !weekday.includes(day))).length === 0) {
                    return DayOfWeekType.WEEKDAY;
                }
                else if (allDays.filter(day => !pimDays.includes(day)).concat(pimDays.filter(day => !allDays.includes(day))).length === 0) {
                    return DayOfWeekType.DAY;
                }

                Logger.getInstance().error(`byDay ${pimDays} is not supported for rule: ${util.inspect(rule, false, 5)}`);
                throw new Error(`byDay days ${pimDays} are not supported by EWS`)
            }
        }
        else {
            Logger.getInstance().error(`byDay is required for rule: ${util.inspect(rule, false, 5)}`);
            throw new Error('byDay is required for recurrence rule');
        }

    }

    /**
     * Convert a jsCalendar instance position to the EWS day of week index.
     * @param instance The instance position from jsCalendar recurrence rule
     * @returns The EWS day of week index
     * @throws An error if the instance number is not supported by EWS. 
     */
    protected instanceToEWSDayOfWeekIndex(instance: number): DayOfWeekIndexType {

        if (instance === 1) {
            return DayOfWeekIndexType.FIRST;
        }
        else if (instance === 2) {
            return DayOfWeekIndexType.SECOND;
        }
        else if (instance === 3) {
            return DayOfWeekIndexType.THIRD;
        }
        else if (instance === 4) {
            return DayOfWeekIndexType.FOURTH;
        }
        else if (instance === -1) {
            return DayOfWeekIndexType.LAST;
        }

        Logger.getInstance().error(`Unsupported instance index: ${instance}`);
        throw new Error(`Unsupported instance index: ${instance}`);
    }
}

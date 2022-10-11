import {
    PimLabel, UserInfo, PimItemFactory, PimItemFormat, PimImportance, PimCalendarItem,
    KeepPimConstants, KeepPimCalendarManager, KeepPimMessageManager, hasHTML, PimParticipationStatus, hasTimeZone
} from '@hcllabs/openclientkeepcomponent';
import { 
    EWSServiceManager, addExtendedPropertiesToPIM, addExtendedPropertyToPIM, getEmail, getValueForExtendedFieldURI,
    identifiersForPathToExtendedFieldType 
} from '.';
import { UserContext } from '../keepcomponent';
import { Logger } from './logger';
import {
    BodyTypeType,
    CalendarItemTypeType, DefaultShapeNamesType, ExtendedPropertyKeyType, ImportanceChoicesType, MessageDispositionType,
    ResponseTypeType, SensitivityChoicesType, UnindexedFieldURIType
} from '../models/enum.model';
import {
    AppendToItemFieldType, AttendeeType, BodyType, CalendarItemType, DeletedOccurrenceInfoType, DeleteItemFieldType, 
    EmailAddressType, FolderIdType, ItemChangeType, ItemIdType, ItemResponseShapeType, ItemType, 
    NonEmptyArrayOfAttendeesType, NonEmptyArrayOfDeletedOccurrencesType, NonEmptyArrayOfOccurrenceInfoType, 
    OccurrenceInfoType, SetItemFieldType, SingleRecipientType, TimeZoneDefinitionType, TimeZoneType
} from '../models/mail.model';
import { getEWSId, getFallbackCreatedDate, pimRecurrenceRuleToRRule } from './pimHelper';
import { Request } from '@loopback/rest';
import { DateTime, DateTimeJSOptions, IANAZone } from 'luxon';
import * as moment from 'moment-timezone';
import fastcopy from 'fast-copy';
import RRule from 'rrule';
import * as JsonPointer from 'json-pointer';
import * as util from 'util';
import { findIana, findWindows } from 'windows-iana';

/**
 * This class is an EWSServicemanager subclass responsible for managing calendar items.
 */
export class EWSCalendarManager extends EWSServiceManager {

    private static instance: EWSCalendarManager;

    /**
     * Shape is defined at:
     * See https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/baseshape
     * 
     * The properties included for each Shape is defined at:
     * https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/property-sets-and-response-shapes-in-ews-in-exchange
     */
    // Fields included for DEFAULT shape for calendar
    private static defaultFields = [
        ...EWSServiceManager.idOnlyFields,
        UnindexedFieldURIType.CALENDAR_END,
        UnindexedFieldURIType.ITEM_IS_ASSOCIATED,
        UnindexedFieldURIType.CALENDAR_LEGACY_FREE_BUSY_STATUS,
        UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS,
        UnindexedFieldURIType.CALENDAR_LOCATION,
        UnindexedFieldURIType.CALENDAR_ORGANIZER,
        UnindexedFieldURIType.ITEM_SUBJECT,
        UnindexedFieldURIType.CALENDAR_UID,
        // Not as documented but testing indicates they are needed
        UnindexedFieldURIType.ITEM_DATE_TIME_CREATED,
        UnindexedFieldURIType.CALENDAR_START
    ];

    // Fields included for ALL_PROPERTIES shape for notes
    private static allPropertiesFields = [
        ...EWSCalendarManager.defaultFields,
        UnindexedFieldURIType.ITEM_PARENT_FOLDER_ID,
        UnindexedFieldURIType.CALENDAR_ADJACENT_MEETING_COUNT,
        UnindexedFieldURIType.CALENDAR_ADJACENT_MEETINGS,
        UnindexedFieldURIType.CALENDAR_ALLOW_NEW_TIME_PROPOSAL,
        UnindexedFieldURIType.CALENDAR_APPOINTMENT_REPLY_TIME,
        UnindexedFieldURIType.CALENDAR_APPOINTMENT_SEQUENCE_NUMBER,
        UnindexedFieldURIType.CALENDAR_APPOINTMENT_STATE,
        UnindexedFieldURIType.ITEM_BODY,
        UnindexedFieldURIType.CALENDAR_CALENDAR_ITEM_TYPE,
        UnindexedFieldURIType.ITEM_CATEGORIES,
        UnindexedFieldURIType.CALENDAR_CONFERENCE_TYPE,
        UnindexedFieldURIType.CALENDAR_CONFLICTING_MEETING_COUNT,
        UnindexedFieldURIType.CALENDAR_CONFLICTING_MEETINGS,
        UnindexedFieldURIType.ITEM_CONVERSATION_ID,
        UnindexedFieldURIType.ITEM_CULTURE,
        UnindexedFieldURIType.ITEM_DATE_TIME_RECEIVED,
        UnindexedFieldURIType.ITEM_DATE_TIME_SENT,
        UnindexedFieldURIType.CALENDAR_DATE_TIME_STAMP,
        UnindexedFieldURIType.CALENDAR_DELETED_OCCURRENCES,
        UnindexedFieldURIType.ITEM_DISPLAY_CC,
        UnindexedFieldURIType.ITEM_DISPLAY_TO,
        UnindexedFieldURIType.CALENDAR_DURATION,
        UnindexedFieldURIType.ITEM_EFFECTIVE_RIGHTS,
        UnindexedFieldURIType.CALENDAR_END_TIME_ZONE,
        UnindexedFieldURIType.CALENDAR_FIRST_OCCURRENCE,
        UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS,
        UnindexedFieldURIType.ITEM_IMPORTANCE,
        UnindexedFieldURIType.ITEM_IN_REPLY_TO,
        UnindexedFieldURIType.ITEM_INTERNET_MESSAGE_HEADERS,
        UnindexedFieldURIType.CALENDAR_IS_ALL_DAY_EVENT,
        UnindexedFieldURIType.ITEM_IS_ASSOCIATED,
        UnindexedFieldURIType.CALENDAR_IS_CANCELLED,
        UnindexedFieldURIType.ITEM_IS_DRAFT,
        UnindexedFieldURIType.ITEM_IS_FROM_ME,
        UnindexedFieldURIType.CALENDAR_IS_MEETING,
        UnindexedFieldURIType.CALENDAR_IS_ONLINE_MEETING,
        UnindexedFieldURIType.CALENDAR_IS_RECURRING,
        UnindexedFieldURIType.ITEM_IS_RESEND,
        UnindexedFieldURIType.CALENDAR_IS_RESPONSE_REQUESTED,
        UnindexedFieldURIType.ITEM_IS_SUBMITTED,
        UnindexedFieldURIType.ITEM_IS_UNMODIFIED,
        UnindexedFieldURIType.ITEM_LAST_MODIFIED_NAME,
        UnindexedFieldURIType.ITEM_LAST_MODIFIED_TIME,
        UnindexedFieldURIType.CALENDAR_LAST_OCCURRENCE,
        UnindexedFieldURIType.CALENDAR_LEGACY_FREE_BUSY_STATUS,
        UnindexedFieldURIType.CALENDAR_LOCATION,
        UnindexedFieldURIType.CALENDAR_MEETING_REQUEST_WAS_SENT,
        UnindexedFieldURIType.CALENDAR_MEETING_TIME_ZONE,
        UnindexedFieldURIType.CALENDAR_MEETING_WORKSPACE_URL,
        UnindexedFieldURIType.CALENDAR_MODIFIED_OCCURRENCES,
        UnindexedFieldURIType.CALENDAR_MY_RESPONSE_TYPE,
        UnindexedFieldURIType.CALENDAR_NET_SHOW_URL,
        UnindexedFieldURIType.CALENDAR_OPTIONAL_ATTENDEES,
        UnindexedFieldURIType.CALENDAR_ORGANIZER,
        UnindexedFieldURIType.CALENDAR_ORIGINAL_START,
        UnindexedFieldURIType.CALENDAR_RECURRENCE,
        UnindexedFieldURIType.ITEM_REMINDER_DUE_BY,
        UnindexedFieldURIType.ITEM_REMINDER_IS_SET,
        UnindexedFieldURIType.ITEM_REMINDER_MINUTES_BEFORE_START,
        UnindexedFieldURIType.CALENDAR_REQUIRED_ATTENDEES,
        UnindexedFieldURIType.CALENDAR_RESOURCES,
        UnindexedFieldURIType.ITEM_RESPONSE_OBJECTS,
        UnindexedFieldURIType.ITEM_SENSITIVITY,
        UnindexedFieldURIType.ITEM_SIZE,
        UnindexedFieldURIType.CALENDAR_START_TIME_ZONE,
        UnindexedFieldURIType.ITEM_SUBJECT,
        UnindexedFieldURIType.CALENDAR_TIME_ZONE,
        UnindexedFieldURIType.ITEM_WEB_CLIENT_EDIT_FORM_QUERY_STRING,
        UnindexedFieldURIType.ITEM_WEB_CLIENT_READ_FORM_QUERY_STRING
    ];

    public static getInstance(): EWSCalendarManager {
        if (!EWSCalendarManager.instance) {
            this.instance = new EWSCalendarManager();
        }
        return this.instance;
    }

    /**
     * Public functions
     */
    /**
     * This function issues a Keep API call to create a PIM entry based on the EWS item passed in.
     * @param item The EWS item to create
     * @param userInfo The user's credentials to be passed to Keep.
     * @param request The original SOAP request for the create.
     * @param toLabel Optional target label (calendar folder). If not specified, the item will be created in the default calendar. 
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @returns A new EWS item populated with information about the created entry.
     */
    async createItem(
        item: CalendarItemType, 
        userInfo: UserInfo, 
        request: Request, 
        disposition?: MessageDispositionType, 
        toLabel?: PimLabel, 
        mailboxId?: string
    ): Promise<CalendarItemType[]> {
        const pimItem = this.pimItemFromEWSItem(item, request);
        pimItem.calendarName = toLabel?.calendarName ? toLabel.calendarName : KeepPimConstants.DEFAULT_CALENDAR_NAME;

        // Make the Keep API call to create the task
        const send = disposition !== MessageDispositionType.SAVE_ONLY;
        const unid = await KeepPimCalendarManager.getInstance().createCalendarItem(pimItem, userInfo, send, mailboxId);
        const targetItem = new CalendarItemType();
        const itemEWSId = getEWSId(unid, mailboxId);
        targetItem.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);

        // TODO:  Uncomment this line when we can set the parent folder to the base calendar folder
        //        See LABS-866
        // If no toLabel is passed in, we will create the contact in the top level calendar folder
        // if (toLabel === undefined) {
        //     const labels = await KeepPimLabelManager.getInstance().getLabels(userInfo);
        //     toLabel = labels.find(label => { return label.view === KeepPimConstants.CALENDAR });
        // }

        // The Keep API does not support creating a calendar item IN a folder, so we must now move the 
        // calendar item to the desired folder
        // TODO:  If this calendar item is being created in the top level contacts folder, do not attempt to 
        //        move the calendar item to the top level folder as this will fail.
        //        See LABS-866
        if (toLabel !== undefined && toLabel.view !== KeepPimConstants.CALENDAR) {
            const folderId = toLabel.folderId;
            try {
                //Move the item in keep.  
                const moveItems: string[] = [];
                moveItems.push(unid);

                // Since we're creating the item, we don't need to worry about multiple parents so we can just move the item to the right folder
                const results = await KeepPimMessageManager
                    .getInstance()
                    .moveMessages(userInfo, folderId, moveItems, undefined, undefined, undefined, mailboxId);
                if (results.movedIds && results.movedIds.length > 0) {
                    const result = results.movedIds[0];
                    if (result.unid !== unid || result.status !== 200) {
                        Logger.getInstance().error(`move for calendar item ${unid} to folder ${folderId} failed: ${util.inspect(results, false, 5)}`);
                        throw new Error(`Move of item ${unid} failed`);
                    }
                }

            } catch (err) {
                // If there was an error, delete the contact on the server since it didn't make it to the correct location
                Logger.getInstance().debug(`createItem error moving item, ${unid},  to desired folder ${folderId}: ${err}`);
                await this.deleteItem(unid, userInfo, mailboxId);
                throw err;
            }
        }

        return [targetItem];
    }

    /**
     * This function will make a Keep API request to delete the item with the passed in itemId.
     * @param item The calendar item or unid of the item of the item to delete. If you pass in the unid, then the calendar item is delete
     * from the default calendar. If the item is not in the default calendar, pass in the PimCalendarItem. 
     * @param userInfo The user's credentials to be passed to Keep.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @param hardDelete Indicator if the item should be deleted from trash as well
     */
    async deleteItem(item: string | PimCalendarItem, userInfo: UserInfo, mailboxId?: string, hardDelete?: boolean): Promise<void> {
        let calendarName = KeepPimConstants.DEFAULT_CALENDAR_NAME;
        let unid: string;
        if (typeof item === 'string') {
            unid = item;
        }
        else {
            calendarName = item.calendarName;
            unid = item.unid;
        }
        await KeepPimCalendarManager.getInstance().deleteCalendarItem(calendarName, unid, userInfo, hardDelete, mailboxId);
    }

    /**
     * This function will make a Keep API PUT request to update an item.
     * @param pimItem Existing PimItem from Keep being updated.
     * @param change The changes to apply to the item.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param toLabel Optional label of the parent folder.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     */
    async updateItem(
        pimItem: PimCalendarItem, 
        change: ItemChangeType, 
        userInfo: UserInfo, 
        request: Request, 
        toLabel: PimLabel,
        mailboxId?: string
    ): Promise<CalendarItemType | undefined> {

        // Set the calendar name if the parent label provided
        if (toLabel?.calendarName !== undefined) {
            pimItem.calendarName = toLabel.calendarName;
        }

        const calendarUpdates = new CalendarItemType();
        let deletesProcessed = false; // True if fields were deleted in pimItem
        const removedCalendarAttendees: string[] = [];

        for (const fieldUpdate of change.Updates.items) {
            let newItem: ItemType | undefined = undefined;
            if (fieldUpdate instanceof SetItemFieldType || fieldUpdate instanceof AppendToItemFieldType) {
                // Should only have to worry about Item and calendar related field updates for calendar
                newItem = fieldUpdate.Item ?? fieldUpdate.CalendarItem ?? fieldUpdate.MeetingMessage ??
                    fieldUpdate.MeetingRequest ?? fieldUpdate.MeetingResponse ?? fieldUpdate.MeetingCancellation;
            }

            if (fieldUpdate instanceof SetItemFieldType) {
                if (newItem === undefined) {
                    Logger.getInstance().error(`No new item set for update field: ${util.inspect(newItem, false, 5)}`);
                    continue;
                }

                if (fieldUpdate.FieldURI) {
                    if (newItem instanceof CalendarItemType) {
                        /*
                         When an attendee is removed we get an set field with the remaining attendees, so we have to figure out which attendees were removed.  
                        */
                        if (fieldUpdate.FieldURI.FieldURI === UnindexedFieldURIType.CALENDAR_REQUIRED_ATTENDEES || fieldUpdate.FieldURI.FieldURI === UnindexedFieldURIType.CALENDAR_OPTIONAL_ATTENDEES) {
                            // Compare with current before we update the Pim item
                            if (newItem.RequiredAttendees) {
                                const updated = newItem.RequiredAttendees.Attendee.map(attendee => { return attendee.Mailbox.EmailAddress });
                                const removed = pimItem.requiredAttendees.filter((attendee: any) => { return !updated.includes(attendee) });
                                removed.forEach((attendee: any) => { removedCalendarAttendees.push(attendee) });
                            }
                            if (newItem.OptionalAttendees) {
                                const updated = newItem.OptionalAttendees.Attendee.map(attendee => { return attendee.Mailbox.EmailAddress });
                                const removed = pimItem.optionalAttendees.filter((attendee: any) => { return !updated.includes(attendee) });
                                removed.forEach((attendee: any) => { removedCalendarAttendees.push(attendee) });
                            }
                        } 
                    }

                    /*
                     The start and end time zones from EWS need to be converted to Windows time zones. So we delay setting the start/end and corresponding 
                     time zones until after all updates have been gathered (see below).
                     */
                    const field = fieldUpdate.FieldURI.FieldURI.split(':')[1];
                    if (fieldUpdate.FieldURI.FieldURI !== UnindexedFieldURIType.CALENDAR_START_TIME_ZONE && 
                        fieldUpdate.FieldURI.FieldURI !== UnindexedFieldURIType.CALENDAR_START &&
                        fieldUpdate.FieldURI.FieldURI !== UnindexedFieldURIType.CALENDAR_END && 
                        fieldUpdate.FieldURI.FieldURI !== UnindexedFieldURIType.CALENDAR_END_TIME_ZONE) {
                        this.updatePimItemFieldValue(pimItem, fieldUpdate.FieldURI.FieldURI, (newItem as any)[field]);
                    }

                    if (newItem instanceof CalendarItemType) {
                        // Add the updates to the calendar updates
                        (calendarUpdates as any)[field] = (newItem as any)[field];
                        if (field === 'Start') {
                            calendarUpdates.StartString = (newItem as any)['StartString'];
                        }
                        else if (field === 'End') {
                            calendarUpdates.EndString = (newItem as any)['EndString'];
                        }
                    }
                }
                else if (fieldUpdate.ExtendedFieldURI && newItem.ExtendedProperty) {
                    const identifiers = identifiersForPathToExtendedFieldType(fieldUpdate.ExtendedFieldURI);
                    const newValue = getValueForExtendedFieldURI(newItem, identifiers);

                    const current = pimItem.findExtendedProperty(identifiers);
                    if (current) {
                        current.Value = newValue;
                        pimItem.updateExtendedProperty(identifiers, current);
                    }
                    else {
                        addExtendedPropertyToPIM(pimItem, newItem.ExtendedProperty[0]);
                    }

                    if (newItem instanceof CalendarItemType && newItem.ExtendedProperty) {
                        // Add the updates to the calendar updates
                        if (calendarUpdates.ExtendedProperty === undefined) {
                            calendarUpdates.ExtendedProperty = [];
                        }
                        calendarUpdates.ExtendedProperty.push(newItem.ExtendedProperty[0]);
                    }
                }
                else if (fieldUpdate.IndexedFieldURI) {
                    Logger.getInstance().warn(`Unhandled SetItemField request for calendar item field:  ${fieldUpdate.IndexedFieldURI}`);
                }
            }
            else if (fieldUpdate instanceof AppendToItemFieldType) {
                /*
                    The following properties are supported for the append action for notes:
                    - item:Body
                    - All the recipient and attendee collection properties

                    So the values will either be arrays or strings.
                */
                if (fieldUpdate.FieldURI?.FieldURI === UnindexedFieldURIType.ITEM_BODY && newItem?.Body) {
                    pimItem.body = `${pimItem.body}${newItem.Body.Value ?? ""}`;
                    calendarUpdates.Body = new BodyType(pimItem.body, hasHTML(pimItem.body) ? BodyTypeType.HTML : BodyTypeType.TEXT);

                } else if (fieldUpdate.FieldURI?.FieldURI === UnindexedFieldURIType.CALENDAR_REQUIRED_ATTENDEES && pimItem.isPimCalendarItem() && newItem instanceof CalendarItemType) {
                    if (newItem.RequiredAttendees) {
                        // Update the required attendees in the PIM item and EWS item
                        pimItem.requiredAttendees = this.mergeAttendees(UnindexedFieldURIType.CALENDAR_REQUIRED_ATTENDEES, newItem.RequiredAttendees.Attendee, pimItem.requiredAttendees, calendarUpdates);
                    }
                    else {
                        Logger.getInstance().error(`No required attendees set for append update for calendar item ${pimItem.unid}`);
                    }
                }
                else if (fieldUpdate.FieldURI?.FieldURI === UnindexedFieldURIType.CALENDAR_OPTIONAL_ATTENDEES && pimItem.isPimCalendarItem() && newItem instanceof CalendarItemType) {
                    if (newItem.OptionalAttendees) {
                        // Update the optional attendees in the PIM item and EWS item
                        pimItem.optionalAttendees = this.mergeAttendees(UnindexedFieldURIType.CALENDAR_OPTIONAL_ATTENDEES, newItem.OptionalAttendees.Attendee, pimItem.optionalAttendees, calendarUpdates);
                    }
                    else {
                        Logger.getInstance().error(`No optional attendees set for append update for calendar item ${pimItem.unid}`);
                    }
                }
                else {
                    Logger.getInstance().warn(`Unhandled AppendToItemField request for calendar item field:  ${fieldUpdate.FieldURI?.FieldURI}`);
                }
            }
            else if (fieldUpdate instanceof DeleteItemFieldType) {
                if (fieldUpdate.FieldURI) {
                    this.updatePimItemFieldValue(pimItem, fieldUpdate.FieldURI.FieldURI, undefined);
                    deletesProcessed = true;
                }
                else if (fieldUpdate.ExtendedFieldURI) {
                    pimItem.deleteExtendedProperty(identifiersForPathToExtendedFieldType(fieldUpdate.ExtendedFieldURI));
                    deletesProcessed = true;
                }
                else if (fieldUpdate.IndexedFieldURI) {
                    // Only handling these for contacts currently
                    Logger.getInstance().warn(`Unhandled DeleteItemField request for calendar item field:  ${fieldUpdate.IndexedFieldURI}`);
                }
            }
        }

        // Set the target folder if passed in
        if (toLabel) {
            pimItem.parentFolderIds = [toLabel.unid];  // May be possible we've lost other parent ids set by another client.  
            // Even if stored as an extra property for the client, it could have been updated on the server since stored on the client.
            // To shrink the window, we'd need to request for the item from the server and update it right away
        }

        // Special code to handle setting start and end times with time zones. Microsoft time zone values need to be convered to IANA time zones. 
        if (calendarUpdates.Start !== undefined || calendarUpdates.End !== undefined) {
            const times = this.getTimesWithZone(calendarUpdates);
            if (times.Start !== undefined) {
                pimItem.start = times.Start; 
            }

            if (times.End !== undefined) {
                pimItem.end = times.End; 
            }
        }

        // The pimItem should now be updated with the new information.  First get the structure for items that need to 
        // be updated by modifyCalendarItem. 
        const modifyStructure = this.modifyCalendarStructure(calendarUpdates, pimItem, removedCalendarAttendees);
        if (modifyStructure) {
            await KeepPimCalendarManager.getInstance().modifyCalendarItem(pimItem.calendarName, pimItem.unid, modifyStructure, userInfo, mailboxId);
        }

        if (Object.getOwnPropertyNames(calendarUpdates).length > 0 || deletesProcessed) {
            // There are updates outside of those handled by modifyCalendarItem
            await KeepPimCalendarManager.getInstance().updateCalendarItem(pimItem, userInfo, mailboxId);
        }

        const shape = new ItemResponseShapeType();
        shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
        return this.pimItemToEWSItem(pimItem, userInfo, request, shape, mailboxId);
    }

    /**
   * This function will fetech a group of items using the Keep API and return an
   * array of corresponding EWS items populated with the fields requested in the shape.
   * @param userInfo The user's credentials to be passed to Keep.
   * @param request The original SOAP request for the get.
   * @param shape EWS shape describing the information requested for each item.
   * @param startIndex Optional start index for items when paging.  Defaults to 0.
   * @param count Optional count of objects to request.  Defaults to 512.
   * @param fromLabel The unid of the label (folder) we are querying.
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
    ): Promise<CalendarItemType[]> {
        const calName = fromLabel?.calendarName ?? KeepPimConstants.DEFAULT_CALENDAR_NAME;
        const calItems: CalendarItemType[] = [];
        let pimItems: PimCalendarItem[] | undefined = undefined;
        try {
            pimItems = await KeepPimCalendarManager
                .getInstance()
                .getCalendarItems(calName, userInfo, undefined, undefined, startIndex, count, mailboxId);
        } catch (err) {
            Logger.getInstance().debug("Error retrieving PIM Calendar entries: " + err);
        }

        if (pimItems && pimItems.length > 0) {
            for (const pimItem of pimItems) {
                const calItem = await this.pimItemToEWSItem(pimItem, userInfo, request, shape, mailboxId, fromLabel?.unid);
                if (calItem) {
                    calItems.push(calItem);
                } else {
                    Logger.getInstance().error(`An EWS calendar item could not be created for PIM calendar item ${pimItem.unid}`);
                }
            }
        }

        return calItems;
    }

    /**
     * Internal functions
     */

    /**
     * Creates a new Pim Calendar item. 
     * @param pimItem The calendar item to create.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @returns The new Pim calendar item created with Keep.
     * @throws An error if the create fails
     */
    protected async createNewPimItem(pimItem: PimCalendarItem, userInfo: UserInfo, mailboxId?: string): Promise<PimCalendarItem> {
        const unid = await KeepPimCalendarManager.getInstance().createCalendarItem(pimItem, userInfo, undefined, mailboxId);

        let newItem: PimCalendarItem | undefined;
        try {
            newItem = await KeepPimCalendarManager.getInstance().getCalendarItem(unid, pimItem.calendarName, userInfo, mailboxId);
        } catch (err) {
            Logger.getInstance().error(`Error retreiving copy of item: ${err}`);
        }

        if (newItem === undefined) {
            // Try to delete item since it may have been created
            try {
                await KeepPimCalendarManager.getInstance().deleteCalendarItem(pimItem.calendarName, unid, userInfo, undefined, mailboxId);
            }
            catch {
                // Ignore errors
            }
            throw new Error(`Unable to retrieve calendar item ${unid} after create`);
        }

        return newItem;
    }

    /**
     * Getter for the array of fields returned for the default shape for the subclass.
     * @returns An array of UnindexedFieldURIType descriptors corresponding to the fields requested by the default shape.
     */
    protected fieldsForDefaultShape(): UnindexedFieldURIType[] {
        return EWSCalendarManager.defaultFields;
    }

    /**
     * Getter for the array of fields returned for the all properties shape for the subclass.
     * @returns An array of UnindexedFieldURIType descriptors corresponding to the fields requested by the all properties shape.
     */
    protected fieldsForAllPropertiesShape(): UnindexedFieldURIType[] {
        return EWSCalendarManager.allPropertiesFields;
    }


    /**
     * Create an EWS ItemType object based on the provided PIM item.
     * @param pimItem The source PimItem from which to build an EWS item.
     * @param shape The EWS shape describing the fields that should be populated on the returned object.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @returns An EWS ItemType subclass object based on the passed PimItem and shape in shape.
     */
    public async pimItemToEWSItem(
        pimCalendarItem: PimCalendarItem, 
        userInfo: UserInfo, 
        request: Request, 
        shape: ItemResponseShapeType, 
        mailboxId?: string,
        targetFolderId?: string
    ): Promise<CalendarItemType> {
        const item = new CalendarItemType();
    
        // Add all requested properties based on the shape
        this.addRequestedPropertiesToEWSItem(pimCalendarItem, item, shape, mailboxId);

        if (targetFolderId) {
            const parentFolderEWSId = getEWSId(targetFolderId, mailboxId);
            item.ParentFolderId = new FolderIdType(parentFolderEWSId, `ck-${parentFolderEWSId}`);
        }
        return item;
    }

    /**
     * This is a helper function for each manager to implement to convert an EWS item to a PimItem used to communicate with the Keep API.
     * @param ewsItem The source EWS item to convert to a PimItem.
     * @param request The original SOAP request for the get.
     * @param existing An optional object containing existing fields to apply.
     * @param parentLabel The parent lable of the pim item.
     * @returns A PimItem populated with the EWS item's data.
     */
    public pimItemFromEWSItem(ewsItem: CalendarItemType, request: Request, existing?: object, parentLabel?: PimLabel): PimCalendarItem {
        let calendarName = KeepPimConstants.DEFAULT_CALENDAR_NAME;
        if (parentLabel?.calendarName) {
            calendarName = parentLabel.calendarName;
        }
        const rtn = PimItemFactory.newPimCalendarItem({}, calendarName, PimItemFormat.DOCUMENT);

        if (ewsItem.Subject) {
            rtn.subject = ewsItem.Subject;
        }

        // Keep Calendar items are either high importance or not
        if (ewsItem.Importance === ImportanceChoicesType.HIGH) {
            rtn.importance = PimImportance.HIGH;
        }

        if (ewsItem.Sensitivity) {
            rtn.isPrivate = ewsItem.Sensitivity !== SensitivityChoicesType.NORMAL;
        }

        if (ewsItem.Body) {
            rtn.body = ewsItem.Body.Value;
            rtn.bodyType = ewsItem.Body.BodyType;
        }

        const times = this.getTimesWithZone(ewsItem);
        if (times.Start) {
            rtn.start = times.Start;
        }
        if (times.End) {
            rtn.end = times.End;
        }

        if (ewsItem.IsDraft) {
            rtn.isDraft = true;
        }

        // TODO:  For attendees and optional attendees, should we be sending more than just email addresses?  Shouldn't we preserve
        //        names if they're present?
        let attendeesSet = false;
        if (ewsItem.RequiredAttendees && ewsItem.RequiredAttendees.Attendee.length > 0) {
            const attendees: string[] = [];
            ewsItem.RequiredAttendees.Attendee.forEach(attendee => { attendees.push(attendee.Mailbox.EmailAddress) });
            rtn.requiredAttendees = attendees;
            attendeesSet = true;
        }

        if (ewsItem.OptionalAttendees && ewsItem.OptionalAttendees.Attendee.length > 0) {
            const attendees: string[] = [];
            ewsItem.OptionalAttendees.Attendee.forEach(attendee => { attendees.push(attendee.Mailbox.EmailAddress) });
            rtn.optionalAttendees = attendees;
            attendeesSet = true;
        }

        // Only set the organizer if there are other attendees or it'a a meeting.  i.e. if Appointment or all day event, we don't want any participants.
        if (ewsItem.IsMeeting || (rtn.requiredAttendees && rtn.requiredAttendees.length > 0) || (rtn.optionalAttendees && rtn.optionalAttendees.length > 0)) {
            if (ewsItem.Organizer) {
                rtn.organizer = ewsItem.Organizer.Mailbox.EmailAddress;
            }
            else if (existing === undefined) {
                const userInfo = UserContext.getUserInfo(request);
                if (userInfo?.userId) {
                    rtn.organizer = userInfo.userId;
                }
            }
        }

        if (ewsItem.Location) {
            rtn.location = ewsItem.Location;
        }

        if (ewsItem.Categories && ewsItem.Categories.String.length > 0) {
            rtn.categories = ewsItem.Categories.String;
        }

        if (ewsItem.ReminderMinutesBeforeStart && ewsItem.ReminderIsSet) {
            rtn.alarm = -parseInt(ewsItem.ReminderMinutesBeforeStart);
        }

        // FIXME:  This is not used in PimCalendarItemClassic.  Should it be implemented in PimItemJmap
        // if (ewsItem.AllowNewTimeProposal !== undefined) {
        //     calObject["PreventCounter"] = ewsItem.AllowNewTimeProposal ? "1" : "0";
        // }

        addExtendedPropertiesToPIM(ewsItem, rtn); // Add extended properties from the calendar item

        if (ewsItem.UID) {
            /*
             When an EWS client creates a calendar item it will set a UID for the ewsItem. This is different from the UNID when the item is created in Keep. 
             We will save the UID generated by the client so that it can be returned when the client asked for the ewsItem. 
            */
            rtn.setAdditionalProperty(UnindexedFieldURIType.CALENDAR_UID, ewsItem.UID);
        }

        if (ewsItem.MeetingRequestWasSent) {
            /*
             Was the meeting request sent?  Typically this would be managed by the server.  Here if the state is known by the client already
            */
            rtn.setAdditionalProperty(UnindexedFieldURIType.CALENDAR_MEETING_REQUEST_WAS_SENT, ewsItem.MeetingRequestWasSent);
        }

        if (ewsItem.IsResponseRequested) {
            rtn.setAdditionalProperty(UnindexedFieldURIType.CALENDAR_IS_RESPONSE_REQUESTED, ewsItem.IsResponseRequested);
        }

        if (ewsItem.AppointmentSequenceNumber) {
            rtn.setAdditionalProperty(UnindexedFieldURIType.CALENDAR_APPOINTMENT_SEQUENCE_NUMBER, ewsItem.AppointmentSequenceNumber);
        }

        if (ewsItem.AppointmentState) {
            rtn.setAdditionalProperty(UnindexedFieldURIType.CALENDAR_APPOINTMENT_STATE, ewsItem.AppointmentState);
        }
        
        // TODO:  Map this to JMAP?  Not sure why we are using an additional property for this setting
        if (ewsItem.MeetingTimeZone?.TimeZoneName) {
            rtn.setAdditionalProperty(UnindexedFieldURIType.CALENDAR_MEETING_TIME_ZONE, ewsItem.MeetingTimeZone.TimeZoneName);
        }

        if (ewsItem.IsAllDayEvent) {
            if (!attendeesSet) {
                rtn.isAllDayEvent = true; // Keep only supports all day events with no attendees
            }
            // If the EWS item was set as an all day event, remember it was originally set as all day
            rtn.setAdditionalProperty(UnindexedFieldURIType.CALENDAR_IS_ALL_DAY_EVENT, true);
        }

        return rtn;
    }

    protected updateEWSItemFieldValue(
        item: CalendarItemType, 
        pimItem: PimCalendarItem, 
        fieldIdentifier: UnindexedFieldURIType,
        mailboxId?: string
    ): boolean {
        let handled = true;

        switch (fieldIdentifier) {

            case UnindexedFieldURIType.CALENDAR_LOCATION:
                if (pimItem.location) {
                    item.Location = pimItem.location;
                }
                break;

            case UnindexedFieldURIType.CALENDAR_ORGANIZER:
                if (pimItem.organizer) {
                    const organizer = new SingleRecipientType();
                    organizer.Mailbox = new EmailAddressType();
                    organizer.Mailbox.EmailAddress = pimItem.organizer;
                    organizer.Mailbox.Name = pimItem.organizer;
                    item.Organizer = organizer;
                }
                break;

            case UnindexedFieldURIType.CALENDAR_DURATION:
                if (pimItem.duration) {
                    item.Duration = pimItem.duration;
                }
                break;


            case UnindexedFieldURIType.CALENDAR_START:
                if (pimItem.start) {
                    item.Start = new Date(pimItem.start);
                }
                break;

            case UnindexedFieldURIType.CALENDAR_START_TIME_ZONE:
                if (pimItem.startTimeZone) {
                    const tz = new TimeZoneDefinitionType();
                    const zones = findWindows(pimItem.startTimeZone);
                    if (zones.length > 0) {
                        tz.Name = zones[0]; // Take the first one
                        tz.Id = tz.Name;
                        item.StartTimeZone = tz;
                    }
                    else {
                        Logger.getInstance().error(`Unable to convert start time zone ${pimItem.startTimeZone} to a Windows time zone`);
                    }
                }
                break;

            case UnindexedFieldURIType.CALENDAR_END:
                if (pimItem.end) {
                    item.End = new Date(pimItem.end);
                }
                break;

            case UnindexedFieldURIType.CALENDAR_END_TIME_ZONE:
                {
                    const timeZone = pimItem.endTimeZone ?? pimItem.startTimeZone;
                    if (timeZone) {
                        const zones = findWindows(timeZone);
                        if (zones.length > 0) {
                            const tz = new TimeZoneDefinitionType();
                            tz.Name = zones[0]; // Take the first one
                            tz.Id = tz.Name;
                            item.EndTimeZone = tz;
                        }
                        else {
                            Logger.getInstance().error(`Unable to convert end time zone ${timeZone} to a Windows time zone`);
                        }
                    }
                }
                break;

            case UnindexedFieldURIType.CALENDAR_IS_ALL_DAY_EVENT:
                item.IsAllDayEvent = pimItem.isAllDayEvent;
                break;

            case UnindexedFieldURIType.CALENDAR_IS_MEETING:
                item.IsMeeting = pimItem.isMeeting;
                break;

            case UnindexedFieldURIType.CALENDAR_IS_RECURRING:
                item.IsRecurring = (pimItem.recurrenceRules ?? []).length > 0;
                break;

            case UnindexedFieldURIType.CALENDAR_FIRST_OCCURRENCE:
            case UnindexedFieldURIType.CALENDAR_LAST_OCCURRENCE:
                this.addOccurrenceInstance(pimItem, fieldIdentifier, item);
                break;

            case UnindexedFieldURIType.CALENDAR_DELETED_OCCURRENCES:
                this.addDeletedOccurrences(pimItem, item);
                break;

            case UnindexedFieldURIType.CALENDAR_MODIFIED_OCCURRENCES:
                this.processOccurrenceOverrides(pimItem, item);
                break;

            case UnindexedFieldURIType.CALENDAR_IS_CANCELLED:
                item.IsCancelled = false; // TODO How to determine this
                break;

            case UnindexedFieldURIType.CALENDAR_MEETING_TIME_ZONE:
                {
                    const tz = pimItem.getAdditionalProperty(fieldIdentifier);
                    if (tz) {
                        item.MeetingTimeZone = new TimeZoneType();
                        item.MeetingTimeZone.TimeZoneName = tz;
                    }
                }
                break;

            case UnindexedFieldURIType.CALENDAR_OPTIONAL_ATTENDEES:
                if (pimItem.optionalAttendees.length > 0) {
                    item.OptionalAttendees = new NonEmptyArrayOfAttendeesType();
                    item.OptionalAttendees.Attendee = this.getAttendees(pimItem.optionalAttendees, pimItem);
                }
                break;

            case UnindexedFieldURIType.CALENDAR_REQUIRED_ATTENDEES:
                if (pimItem.requiredAttendees.length > 0) {
                    item.RequiredAttendees = new NonEmptyArrayOfAttendeesType();
                    item.RequiredAttendees.Attendee = this.getAttendees(pimItem.requiredAttendees, pimItem);
                }
                break;

            case UnindexedFieldURIType.CALENDAR_RECURRENCE: {
                    // Outlook calendar items do not show if the recurrence field is set but undefined.
                    const recurrence = this.getEWSRecurrence(pimItem);
                    if (recurrence) {
                        item.Recurrence = recurrence;
                    }
                }
                break;

            case UnindexedFieldURIType.CALENDAR_LEGACY_FREE_BUSY_STATUS:
                // TODO blocked by OOO 
                break;

            case UnindexedFieldURIType.CALENDAR_CALENDAR_ITEM_TYPE:
                // TODO How to set EXCEPTION item Type

                if ((pimItem.recurrenceRules ?? []).length > 0) {
                    item.CalendarItemType = CalendarItemTypeType.RECURRING_MASTER;
                }
                else if (pimItem.recurrenceId) {
                    item.CalendarItemType = CalendarItemTypeType.OCCURRENCE;
                }
                else {
                    item.CalendarItemType = CalendarItemTypeType.SINGLE;
                }
                break;

            case UnindexedFieldURIType.CALENDAR_TIME_ZONE:
                // TODO: Blocked by LABS-1949
                break;

            case UnindexedFieldURIType.CALENDAR_DATE_TIME_STAMP:
                item.DateTimeStamp = getFallbackCreatedDate(pimItem);
                break;

            case UnindexedFieldURIType.ITEM_BODY:
                item.Body = new BodyType(pimItem.body, hasHTML(pimItem.body) ? BodyTypeType.HTML : BodyTypeType.TEXT);
                break;

            case UnindexedFieldURIType.CALENDAR_UID:
                item.UID = pimItem.getAdditionalProperty(fieldIdentifier) ?? pimItem.unid;
                break;
    
            case UnindexedFieldURIType.CALENDAR_MEETING_REQUEST_WAS_SENT: {
                    const value = pimItem.getAdditionalProperty(fieldIdentifier);
                    if (value !== undefined) {
                        item.MeetingRequestWasSent = value;
                    }
                }
                break;

            case UnindexedFieldURIType.CALENDAR_IS_RESPONSE_REQUESTED: {
                    const value = pimItem.getAdditionalProperty(fieldIdentifier);
                    if (value !== undefined) {
                        item.IsResponseRequested = value;
                    }
                }
                break;

            case UnindexedFieldURIType.CALENDAR_APPOINTMENT_SEQUENCE_NUMBER: {
                    const value = pimItem.getAdditionalProperty(fieldIdentifier);
                    if (value !== undefined) {
                        item.AppointmentSequenceNumber = value;
                    }
                }
                break;

            case UnindexedFieldURIType.CALENDAR_APPOINTMENT_STATE: {
                    const value = pimItem.getAdditionalProperty(fieldIdentifier);
                    if (value !== undefined) {
                        item.AppointmentState = value;
                    } else {
                        // None 0x0000 No flags have been set. This is only used for an appointment that does not include attendees.
                        // Meeting 0x0001 This appointment is a meeting.
                        // Received 0x0002 This appointment has been received.
                        // Canceled 0x0004 This appointment has been canceled.
                        if (pimItem.isMeeting){
                            item.AppointmentState = 1;
                        } else if (pimItem.isAppointment) {
                            item.AppointmentState = 0;
                        }
                    }
                }
                break;  

            default:
                // Assume these will be saved in additional properties:
                // UnindexedFieldURIType.CALENDAR_ADJACENT_MEETING_COUNT,
                // UnindexedFieldURIType.CALENDAR_ADJACENT_MEETINGS,
                // UnindexedFieldURIType.CALENDAR_APPOINTMENT_REPLY_TIME,
                // UnindexedFieldURIType.CALENDAR_APPOINTMENT_SEQUENCE_NUMBER,
                // UnindexedFieldURIType.CALENDAR_APPOINTMENT_STATE,
                // UnindexedFieldURIType.CALENDAR_CONFLICTING_MEETING_COUNT,
                // UnindexedFieldURIType.CALENDAR_CONFLICTING_MEETINGS,
                // UnindexedFieldURIType.CALENDAR_DATE_TIME_STAMP,
                // UnindexedFieldURIType.CALENDAR_IS_ONLINE_MEETING,
                // UnindexedFieldURIType.CALENDAR_LEGACY_FREE_BUSY_STATUS,
                // UnindexedFieldURIType.CALENDAR_MEETING_WORKSPACE_URL,
                // UnindexedFieldURIType.CALENDAR_MY_RESPONSE_TYPE,
                // UnindexedFieldURIType.CALENDAR_NET_SHOW_URL,
                // UnindexedFieldURIType.CALENDAR_ORIGINAL_START,
                // UnindexedFieldURIType.CALENDAR_RESOURCES,
                // UnindexedFieldURIType.CALENDAR_ALLOW_NEW_TIME_PROPOSAL
                {
                    // Check if the field identifier is stored in additional properties
                    const value = pimItem.getAdditionalProperty(fieldIdentifier);
                    if (value) {
                        const split = fieldIdentifier.split(':');
                        if (split.length > 1) {
                            (item as any)[split[1]] = value;
                        }
                        else {
                            Logger.getInstance().error(`Unexpected field identifier format: ${fieldIdentifier}`);
                            handled = false;
                        }
                    }
                    else {
                        handled = false;
                    }
                }
                break;
        }

        if (handled === false) {
            handled = super.updateEWSItemFieldValue(item, pimItem, fieldIdentifier, mailboxId);
        }

        if (handled === false) {
            Logger.getInstance().warn(`Unrecognized field ${fieldIdentifier} not processed for Calendar item`);
        }

        return handled;
    }

    /**
     * Retuns the start and end time of a calendar item in ISO format with the correct time zone set. 
     * @param item An EWS calendar item containing the start/end times and the time zone settings.
     * @returns An object contain the ISO8601 formatted start and end times with the times zone offset included. 
     */
    protected getTimesWithZone(item: CalendarItemType): { Start: string | undefined; End: string | undefined } {

        /**
         The start and end times are in ISO format, but they do not contain a timezone setting for the meeting. The meeting time zone is set in extended properites. So the Start and End date objects will be have the server's time zone set. 
         We must change the time zone without causing the date/time to be adjusted to the new time zone. 
         */
        const tzIdentifiers: any = {};
        tzIdentifiers[ExtendedPropertyKeyType.PROPERTY_NAME] = "CalendarTimeZone";
        tzIdentifiers[ExtendedPropertyKeyType.PROPERTY_SETID] = "A7B529B5-4B75-47A7-A24F-20743D6C55CD";
        const calendarTimeZone: string | undefined = getValueForExtendedFieldURI(item, tzIdentifiers);

        const rtn: { Start: string | undefined; End: string | undefined } = { Start: undefined, End: undefined };
        if (item.StartString) {
            let timeZone = calendarTimeZone; 
            if (item.StartTimeZone?.Id) {
                // This overrides CalendarTimeZone
                const zones = findIana(item.StartTimeZone.Id);
                if (zones.length > 0) {
                    timeZone = zones[0];
                }
            }
            if (timeZone === undefined) {
                Logger.getInstance().warn(`Timezone information not included for start in calender item ${item.ItemId?.Id}`);
            }
            const opts: DateTimeJSOptions = hasTimeZone(item.StartString) ? {} : {zone: timeZone};
            let st = DateTime.fromISO(item.StartString, opts);
            st = st.setZone('Etc/UTC'); // Convert to UTC
            rtn.Start = st.toISO({includeOffset: true}) ?? undefined;
        }

        if (item.EndString) {
            let timeZone = calendarTimeZone; 
            if (item.EndTimeZone?.Id) {
                // This overrides CalendarTimeZone
                const zones = findIana(item.EndTimeZone.Id);
                if (zones.length > 0) {
                    timeZone = zones[0];
                }
            }
            if (timeZone === undefined) {
                Logger.getInstance().warn(`Timezone information not included for end in calender item ${item.ItemId?.Id}`);
            }
            const opts: DateTimeJSOptions = hasTimeZone(item.EndString) ? {} : {zone: timeZone};
            let et = DateTime.fromISO(item.EndString, opts);
            et = et.setZone('Etc/UTC'); // Conver to UTC
            rtn.End = et.toISO({includeOffset: true}) ?? undefined;
        }

        return rtn;
    }


    /**
     * Returns a structure to modify a calendar entry. 
     * 
     * @param item The calendar item containing the fields to be modified. Upon return any fields set in the returned object will be removed.  
     * @param pimItem The Keep Pim item for the calendar item being updated.
     * @param removedAttendees A list of attendees that should be removed from the meeting. 
     * @returns An object to use on the KEEP modify API to update the calendar fields. It will return undefined if the calendar item does not contain fields supported on modify. The returned object will have the following format:
     * ```
     * "Removed": [
     *       "email_address", ...
     *   ],
     *   "Update": { 
     *       "RequiredAttendees": ["email_address", "email_address", ... ], 
     *       "OptionalAttendees": "email_address",
     *       "STARTDATETIME": "yyyy-mm-ddThh:mm:ss.SSZ", 
     *       "CalendarDateTime": "yyyy-mm-ddThh:mm:ss.SSZ", 
     *       "EndDateTime": "yyyy-mm-ddThh:mm:ss.SSZ"
     *   }
     * ```
     * Removed is optional. Rescheduled and Update are mutually exclusive. Update is used if both attendees and date/time are changed. 
     * 
     */
    protected modifyCalendarStructure(item: CalendarItemType, pimItem: PimCalendarItem, removedAttendees: string[]): any | undefined {

        let hasUpdates = false;
        const newPimCalItem = PimItemFactory.newPimCalendarItem({}, pimItem.calendarName);
        const times = this.getTimesWithZone(item);
        if (times.Start) {
            pimItem.start = times.Start;
            newPimCalItem.start = times.Start;
            hasUpdates = true;

            delete item.Start; // Remove from EWS item so it is not processed again
        }
        if (times.End) {
            pimItem.end = times.End;
            newPimCalItem.end = times.End;
            hasUpdates = true;

            delete item.End; // Remove from EWS item so it is not processed again
        }

        if (item.RequiredAttendees) {
            const attendees = item.RequiredAttendees.Attendee.map(attendee => { return attendee.Mailbox.EmailAddress });

            pimItem.requiredAttendees = attendees; // Update in PimItem
            newPimCalItem.requiredAttendees = attendees;
            hasUpdates = true;

            delete item.RequiredAttendees; // Remove from EWS item so it is not processed again
        }

        if (item.OptionalAttendees) {
            const attendees = item.OptionalAttendees.Attendee.map(attendee => { return attendee.Mailbox.EmailAddress });

            pimItem.optionalAttendees = attendees; // Update in PimItem
            newPimCalItem.optionalAttendees = attendees;
            hasUpdates = true;

            delete item.OptionalAttendees;// Remove from EWS item so it is not processed again
        }

        if (item.ExtendedProperty && item.ExtendedProperty.length === 0) {
            // Remove empty extended properties
            delete item.ExtendedProperty;
        }

        return hasUpdates ? newPimCalItem.toPimStructure() : undefined;
    }

    /**
     * Converts an string of attendees returned from Keep to a list attendee objects.
     * @param attendees A comma seperated list of attendees returned from Keep.
     * @returns An array of attendee objects.
     */
    protected getAttendees(attendees: string[], pimCalendarItem: PimCalendarItem): AttendeeType[] {
        const rtn: AttendeeType[] = [];
        attendees.forEach(attendee => {
            const attendeeType = new AttendeeType();
            attendeeType.Mailbox = getEmail(attendee);
            attendeeType.ResponseType = this.getAttendeeResponseType(attendee, pimCalendarItem);
            rtn.push(attendeeType);
        });
        return rtn;
    }

    /**
     * Merge a EWS attendee list with an email list and add it to an EWS calendar item.
     * @param attendees A list of EWS attendees
     * @param emails A list of email addresses
     * @param toItem The calendar item which to add the combined attendees 
     * @returns A combined list of email addresses
     */
    protected mergeAttendees(type: UnindexedFieldURIType, attendees: AttendeeType[], emails: string[], toItem: CalendarItemType): string[] {

        const newEmails = [...emails];

        // Merge the new attendees with the current email list
        attendees.forEach(attendee => {
            if (!newEmails.includes(attendee.Mailbox.EmailAddress)) {
                newEmails.push(attendee.Mailbox.EmailAddress);
            }
        });

        // Add the new required attendee list to the calendar item
        let target: NonEmptyArrayOfAttendeesType;
        if (type === UnindexedFieldURIType.CALENDAR_OPTIONAL_ATTENDEES) {
            if (toItem.OptionalAttendees === undefined) {
                toItem.OptionalAttendees = new NonEmptyArrayOfAttendeesType();
            }
            target = toItem.OptionalAttendees;
        }
        else {
            if (toItem.RequiredAttendees === undefined) {
                toItem.RequiredAttendees = new NonEmptyArrayOfAttendeesType();
            }
            target = toItem.RequiredAttendees;
        }

        for (const email of newEmails) {
            const attendee = new AttendeeType();
            attendee.Mailbox = new EmailAddressType();
            attendee.Mailbox.EmailAddress = email;
            target.Attendee.push(attendee);
        }

        return newEmails;

    }

    /**
     * Update a value for an EWS field in a Keep PIM calendar item. 
     * @param pimItem The pim calendar item that will be updated.
     * @param fieldIdentifier The EWS unindexed field identifier
     * @param newValue The new value to set. The type is based on what fieldIdentifier is set to. To delete a field, pass in undefined. 
     * @returns true if field was handled
     */
    protected updatePimItemFieldValue(pimCalendarItem: PimCalendarItem, fieldIdentifier: UnindexedFieldURIType, newValue: any | undefined): boolean {
        if (!super.updatePimItemFieldValue(pimCalendarItem, fieldIdentifier, newValue)) {
            if (fieldIdentifier === UnindexedFieldURIType.CALENDAR_START) {
                pimCalendarItem.start = newValue instanceof Date ? newValue.toISOString() : newValue;
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CALENDAR_END) {
                pimCalendarItem.end = newValue instanceof Date ? newValue.toISOString() : newValue;
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CALENDAR_IS_ALL_DAY_EVENT) {
                const allDayEvent: boolean = newValue ?? false;
                if ((pimCalendarItem.requiredAttendees.length === 0 &&
                    pimCalendarItem.optionalAttendees.length === 0 &&
                    pimCalendarItem.fyiAttendees.length === 0) ||
                    allDayEvent === false) {
                    pimCalendarItem.isAllDayEvent = allDayEvent;
                }
                else {
                    Logger.getInstance().warn(`Not setting calendar item ${pimCalendarItem.unid} to all day event since it has attendees`);
                }
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CALENDAR_LOCATION) {
                pimCalendarItem.location = newValue ?? "";
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CALENDAR_ORGANIZER) {
                pimCalendarItem.organizer = newValue ?? "";
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CALENDAR_REQUIRED_ATTENDEES) {
                if (newValue === undefined) {
                    pimCalendarItem.requiredAttendees = [];
                }
                else {
                    const attendees: NonEmptyArrayOfAttendeesType = newValue;
                    pimCalendarItem.requiredAttendees = attendees.Attendee.map(attendee => { return attendee.Mailbox.EmailAddress });
                }
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CALENDAR_OPTIONAL_ATTENDEES) {
                if (newValue === undefined) {
                    pimCalendarItem.optionalAttendees = [];
                }
                else {
                    const attendees: NonEmptyArrayOfAttendeesType = newValue;
                    pimCalendarItem.optionalAttendees = attendees.Attendee.map(attendee => { return attendee.Mailbox.EmailAddress });
                }
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CALENDAR_MEETING_TIME_ZONE) {
                if (newValue === undefined) {
                    pimCalendarItem.deleteAdditionalProperty(UnindexedFieldURIType.CALENDAR_MEETING_TIME_ZONE);
                }
                else {
                    const tzType: TimeZoneType = newValue;
                    if (tzType.TimeZoneName) {
                        pimCalendarItem.setAdditionalProperty(UnindexedFieldURIType.CALENDAR_MEETING_TIME_ZONE, tzType.TimeZoneName);
                    }
                    else {
                        Logger.getInstance().warn(`No timezone name set for ${util.inspect(tzType, false, 5)}`);
                    }
                }
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CALENDAR_UID) {
                if (newValue === undefined) {
                    pimCalendarItem.deleteAdditionalProperty(UnindexedFieldURIType.CALENDAR_UID);
                }
                else {
                    pimCalendarItem.setAdditionalProperty(UnindexedFieldURIType.CALENDAR_UID, newValue);
                }
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CALENDAR_MEETING_REQUEST_WAS_SENT) {
                if (newValue === undefined) {
                    pimCalendarItem.deleteAdditionalProperty(UnindexedFieldURIType.CALENDAR_MEETING_REQUEST_WAS_SENT);
                }
                else {
                    pimCalendarItem.setAdditionalProperty(UnindexedFieldURIType.CALENDAR_MEETING_REQUEST_WAS_SENT, newValue);
                }
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CALENDAR_IS_RESPONSE_REQUESTED) {
                if (newValue === undefined) {
                    pimCalendarItem.deleteAdditionalProperty(UnindexedFieldURIType.CALENDAR_IS_RESPONSE_REQUESTED);
                }
                else {
                    pimCalendarItem.setAdditionalProperty(UnindexedFieldURIType.CALENDAR_IS_RESPONSE_REQUESTED, newValue);
                }
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CALENDAR_APPOINTMENT_SEQUENCE_NUMBER) {
                if (newValue === undefined) {
                    pimCalendarItem.deleteAdditionalProperty(UnindexedFieldURIType.CALENDAR_APPOINTMENT_SEQUENCE_NUMBER);
                }
                else {
                    pimCalendarItem.setAdditionalProperty(UnindexedFieldURIType.CALENDAR_APPOINTMENT_SEQUENCE_NUMBER, newValue);
                }
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CALENDAR_APPOINTMENT_STATE) {
                if (newValue === undefined) {
                    pimCalendarItem.deleteAdditionalProperty(UnindexedFieldURIType.CALENDAR_APPOINTMENT_STATE);
                }
                else {
                    pimCalendarItem.setAdditionalProperty(UnindexedFieldURIType.CALENDAR_APPOINTMENT_STATE, newValue);
                }
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CALENDAR_START_TIME_ZONE) {
                pimCalendarItem.startTimeZone = newValue?.Id ?? "";
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CALENDAR_END_TIME_ZONE) {
                pimCalendarItem.endTimeZone = newValue.Id ?? "";
            }
            else {
                Logger.getInstance().error(`Unsupported field ${fieldIdentifier} for PIM calendar item ${pimCalendarItem.unid}`);
                return false;
            }
        }

        /*
        The following item fields are not currently supported:

        CALENDAR_ORIGINAL_START = 'calendar:OriginalStart',
        CALENDAR_START_WALL_CLOCK = 'calendar:StartWallClock',
        CALENDAR_END_WALL_CLOCK = 'calendar:EndWallClock',
        CALENDAR_START_TIME_ZONE_ID = 'calendar:StartTimeZoneId',
        CALENDAR_END_TIME_ZONE_ID = 'calendar:EndTimeZoneId',
        CALENDAR_LEGACY_FREE_BUSY_STATUS = 'calendar:LegacyFreeBusyStatus',
        CALENDAR_ENHANCED_LOCATION = 'calendar:EnhancedLocation',
        CALENDAR_WHEN = 'calendar:When',
        CALENDAR_IS_CANCELLED = 'calendar:IsCancelled',
        CALENDAR_CALENDAR_ITEM_TYPE = 'calendar:CalendarItemType',
        CALENDAR_MY_RESPONSE_TYPE = 'calendar:MyResponseType',
        CALENDAR_RESOURCES = 'calendar:Resources',
        CALENDAR_CONFLICTING_MEETING_COUNT = 'calendar:ConflictingMeetingCount',
        CALENDAR_ADJACENT_MEETING_COUNT = 'calendar:AdjacentMeetingCount',
        CALENDAR_CONFLICTING_MEETINGS = 'calendar:ConflictingMeetings',
        CALENDAR_ADJACENT_MEETINGS = 'calendar:AdjacentMeetings',
        CALENDAR_INBOX_REMINDERS = 'calendar:InboxReminders',
        CALENDAR_DURATION = 'calendar:Duration',
        CALENDAR_TIME_ZONE = 'calendar:TimeZone',
        CALENDAR_APPOINTMENT_REPLY_TIME = 'calendar:AppointmentReplyTime',
        CALENDAR_CONFERENCE_TYPE = 'calendar:ConferenceType',
        CALENDAR_ALLOW_NEW_TIME_PROPOSAL = 'calendar:AllowNewTimeProposal',
        CALENDAR_IS_ONLINE_MEETING = 'calendar:IsOnlineMeeting',
        CALENDAR_MEETING_WORKSPACE_URL = 'calendar:MeetingWorkspaceUrl',
        CALENDAR_NET_SHOW_URL = 'calendar:NetShowUrl',
        CALENDAR_RECURRENCE_ID = 'calendar:RecurrenceId',
        CALENDAR_JOIN_ONLINE_MEETING_URL = 'calendar:JoinOnlineMeetingUrl',
        CALENDAR_ONLINE_MEETING_SETTINGS = 'calendar:OnlineMeetingSettings',
        CALENDAR_IS_ORGANIZER = 'calendar:IsOrganizer',
        CALENDAR_CALENDAR_ACTIVITY_DATA = 'calendar:CalendarActivityData',
        CALENDAR_DO_NOT_FORWARD_MEETING = 'calendar:DoNotForwardMeeting',
        CALENDAR_DATE_TIME_STAMP = 'calendar:DateTimeStamp'
        */
        return true;
    }

    /**
     * Adds deleted occurrences to an EWS calendar item if a Pim item contains excludedRecurrenceRules. 
     * @param pimItem The PIM item containing the excludedRecurrenceRules
     * @param item The EWS Calendar item 
     */
    protected addDeletedOccurrences(pimItem: PimCalendarItem, item: CalendarItemType): void {
        // Check for deleted occurrences
        if (pimItem.excludedRecurrenceRules) {
            const deletedRules = new NonEmptyArrayOfDeletedOccurrencesType();

            // Need to add a deleted occurrence for each occurence in the excludedRecurrenceRules
            for (const rule of pimItem.excludedRecurrenceRules) {
                const exRule = fastcopy(rule, { isStrict: true }); // Make a copy in case count/until is updated
                if (rule.count === undefined && rule.until === undefined && (pimItem.recurrenceRules ?? []).length === 1) {
                    // The exclude rule does not have an end. Get the end date from the recurrence rule.
                    let occurrence = item.LastOccurrence;
                    if (occurrence === undefined) {
                        const ewsItem = new CalendarItemType();
                        this.addOccurrenceInstance(pimItem, UnindexedFieldURIType.CALENDAR_LAST_OCCURRENCE, ewsItem);
                        occurrence = ewsItem.LastOccurrence;
                    }
                    if (occurrence) {
                        exRule.until = occurrence.Start.toISOString();
                    }
                }

                if (exRule.count !== undefined || exRule.until !== undefined) {
                    const rrule = pimRecurrenceRuleToRRule(exRule, pimItem);
                    for (const ruleDate of rrule.all()) {
                        const deleteRule = new DeletedOccurrenceInfoType();
                        deleteRule.Start = ruleDate;
                        deletedRules.DeletedOccurrence.push(deleteRule);
                    }
                }
                else {
                    Logger.getInstance().warn(`PIM calendar item ${pimItem.unid} exclude rule does not end: ${util.inspect(rule, false, 10)}`);
                }
            }

            if (deletedRules.DeletedOccurrence.length > 0) {
                item.DeletedOccurrences = deletedRules;
            }
            else {
                Logger.getInstance().warn(`No deleted occurrences created for item ${pimItem.unid} from excludedRecurrenceRules: ${util.inspect(pimItem.excludedRecurrenceRules, false, 10)}`);
            }
        }
    }

    /**
     * Adds either modified occurrences or deleted occurrences to an EWS calendar item if a PIM item contains recurrenceOverrides. 
     * @param pimItem The PIM item containing the recurrenceOverrides.
     * @param item The EWS Calendar item to update. 
     * @throw An error if the PIM item recurrence data is not valid. 
     */
    protected processOccurrenceOverrides(pimItem: PimCalendarItem, item: CalendarItemType): void {
        const ids = Object.keys(pimItem.recurrenceOverrides ?? {});
        if (ids.length === 0) {
            Logger.getInstance().warn(`PIM calendar item ${pimItem.unid} does not contain occurrence overrides.`);
            return;
        }

        if (pimItem.start === undefined) {
            Logger.getInstance().warn(`Occurrence overrides not set for PIM calendar item ${pimItem.unid} because it does not have a start time.`);
            return;
        }

        let rrule: RRule | undefined = undefined;

        if (pimItem.recurrenceRules && pimItem.recurrenceRules.length === 1) {
            rrule = pimRecurrenceRuleToRRule(pimItem.recurrenceRules[0], pimItem);
        }

        /**
         * Add EWS Occurrence to ModifiedOccurrences for each patch object in the recurrenceOverrides. An EWS Occurrence must be created for each recurrenceOverrides
         * even if the start/end date does not change so the client knows issue GetItem for the single occurrence calendar item to get the updates. 
         *  
         * Also the code will look for excluded patch object and create a DeletedOccurrence for it. 
         */
        const modified = new NonEmptyArrayOfOccurrenceInfoType();

        for (const id of ids) {
            // Convert the id time string to a Date object used to find a matching instance
            const dateTime = DateTime.fromISO(id, {
                zone: pimItem.startTimeZone === undefined ? undefined : IANAZone.create(pimItem.startTimeZone)
            });
            const idDate = dateTime.toJSDate();
            const match = (rrule && rrule.between(idDate, idDate, true).length > 0);

            const patchObject = pimItem.recurrenceOverrides[id];
            const pointers = Object.keys(patchObject);
            /*
                If the recurrence id does not match a date-time from the recurrence
                rule (or no rule is specified), it is to be treated as an additional
                occurrence (like an RDATE from iCalendar).  The patch object may
                often be empty in this case.
             */
            if (!match) {
                if (pimItem.recurrenceRules === undefined) {
                    // TODO: LABS-2234 - Additional date - how to handle
                    throw new Error(`Adding additional occurrence with recurrenceOverrides is not supported:\n${id} : ${patchObject}`);
                }
                else {
                    Logger.getInstance().warn(`No matching occurrence found for patch id ${id} in PIM calendar item ${pimItem.unid}`);
                }
                continue;
            }

            if (pointers.length > 0) {
                let updatedJMap: any | undefined = undefined; // Updated JMap object containing start/end changes

                for (const pointer of pointers) { // Process each JSON pointer in the patch object
                    if (pointer === "excluded" && patchObject["excluded"] === true) {
                        // This date is excluded, so created a deleted Occurrence
                        let deleted = item.DeletedOccurrences;
                        if (deleted === undefined) {
                            deleted = new NonEmptyArrayOfDeletedOccurrencesType();
                            item.DeletedOccurrences = deleted;
                        }
                        const info = new DeletedOccurrenceInfoType();
                        info.Start = idDate;
                        deleted.DeletedOccurrence.push(info);
                    }
                    else  {
                        if (updatedJMap === undefined) {
                            updatedJMap = fastcopy(pimItem.toPimStructure());
                            updatedJMap.start = id; // Initialize the start date with the instance being processed
                        }
                        let jsonPointer = pointer; 
                        if (!jsonPointer.startsWith('/')) {
                            jsonPointer = `/${jsonPointer}`; // Make sure pointer starts with '/'
                        }
                        JsonPointer.set(updatedJMap, jsonPointer, patchObject[pointer]);
                    }
                }

                if (updatedJMap) {
                    const updatePimItem = PimItemFactory.newPimCalendarItem(updatedJMap, pimItem.calendarName);

                    const occurrence = new OccurrenceInfoType();
                    occurrence.ItemId = new ItemIdType(updatePimItem.unid, `ck-${updatePimItem.unid}`); // TODO: LABS-2233 - Should this be the UNID of the single occurrence
                    occurrence.OriginalStart = new Date(pimItem.start);

                    if (updatePimItem.start) {
                        occurrence.Start = new Date(updatePimItem.start);
                    }
                    else {
                        // Should not occur
                        Logger.getInstance().error(`PIM calendar item ${updatePimItem.unid} reccurrence override requires start to be set: ${util.inspect(patchObject, false, 5)}`);
                        throw new Error(`PIM calendar item ${updatePimItem.unid} reccurrence override requires start to be set`);
                    }

                    if (updatePimItem.end) {
                        occurrence.End = new Date(updatePimItem.end);
                    }
                    else {
                        // Should not occur
                        Logger.getInstance().error(`PIM calendar item ${updatePimItem.unid} reccurrence override requires end to be set: ${util.inspect(patchObject, false, 5)}`);
                        throw new Error(`PIM calendar item ${updatePimItem.unid} reccurrence override requires end to be set`);
                    }

                    modified.Occurrence.push(occurrence);
                }
            }
            else {
                throw new Error(`PIM calendar item ${pimItem.unid} patch object ${id} does not contain data.`);
            }

        }

        if (modified.Occurrence.length > 0) {
            item.ModifiedOccurrences = modified;
        }
    }

    /**
     * Add either the first or last occurrence of a recurring meeting to a EWS item.
     * @param pimItem The PIM calendar item for the recurring meeting.
     * @param type The type of occurence to add
     * @param ewsItem The EWS calendar item that will be updated with the occurence. 
     * @throws An error if the recurrence rule is invalid.
     */
    protected addOccurrenceInstance(pimItem: PimCalendarItem, type: UnindexedFieldURIType.CALENDAR_FIRST_OCCURRENCE | UnindexedFieldURIType.CALENDAR_LAST_OCCURRENCE, ewsItem: CalendarItemType): void {

        if ((pimItem.recurrenceRules ?? []).length === 0) {
            Logger.getInstance().warn(`Pim calendar item ${pimItem.unid} does not recur.`);
            return;
        }

        if (pimItem.start === undefined) {
            Logger.getInstance().warn(`Unable to calculate recurrence for PIM calendar item ${pimItem.unid} because it does not have a start date.`);
            return;
        }

        if (pimItem.duration === undefined) {
            Logger.getInstance().warn(`Unable to calculate recurrence for PIM calendar item ${pimItem.unid} because it does not have a duration.`);
            return;
        }

        if (pimItem.recurrenceRules?.length !== 1) {
            throw new Error(`Unable to calculate recurrence for PIM calendar item ${pimItem.unid} because multiple recurrence rules are not supported`);
        }

        const rule = pimItem.recurrenceRules[0];

        const rrule = pimRecurrenceRuleToRRule(rule, pimItem);

        const info = new OccurrenceInfoType();
        info.ItemId = ewsItem.ItemId ?? new ItemIdType(pimItem.unid); // TODO: LABS-2233 - Should this be the UNID of the single occurrence

        if (type === UnindexedFieldURIType.CALENDAR_FIRST_OCCURRENCE) {
            // Use a date way in the past to get the first date
            const first = rrule.after(new Date(Date.UTC(1900, 0)));
            info.Start = first;
            // Add duration to the start for the first date to get the end of the first occurence.
            let dt = DateTime.fromJSDate(first);
            const millis = moment.duration(pimItem.duration).asMilliseconds();
            dt = dt.plus(millis);
            info.End = dt.toJSDate();
            ewsItem.FirstOccurrence = info;

        }
        else {
            // If count or until not defined this recurrence has no end
            if (rule.count !== undefined || rule.until !== undefined) {
                // Use a date way in the future to get the last date
                const last = rrule.before(new Date(Date.UTC(6000, 0)));
                info.Start = last;
                // Add duration to the start for the last date to get the end of the last occurence.
                let dt = DateTime.fromJSDate(last);
                const millis = moment.duration(pimItem.duration).asMilliseconds();
                dt = dt.plus(millis);
                info.End = dt.toJSDate();
                ewsItem.LastOccurrence = info;
            }
            else {
                Logger.getInstance().warn(`Unable to calculate last occurrence because PIM calendar item ${pimItem.unid} rule does not an ending`);
            }

        }
    }

    protected getAttendeeResponseType(attendee: string, pimCalendarItem: PimCalendarItem): ResponseTypeType {
        let respType: ResponseTypeType;

        if (attendee === pimCalendarItem.organizer) {
            respType = ResponseTypeType.ORGANIZER;
        } else {
            const participationResponse: PimParticipationStatus = pimCalendarItem.getParticipationStatus(attendee);

            switch (participationResponse) {
                case PimParticipationStatus.NEEDS_ACTION:
                    respType = ResponseTypeType.NO_RESPONSE_RECEIVED;
                    break;
    
                case PimParticipationStatus.ACCEPTED:
                    respType = ResponseTypeType.ACCEPT;
                    break;
    
                case PimParticipationStatus.DECLINED:
                    respType = ResponseTypeType.DECLINE;
                    break;
    
                case PimParticipationStatus.TENTATIVE:
                    respType = ResponseTypeType.TENTATIVE;
                    break;
    
                default:
                    respType = ResponseTypeType.UNKNOWN;
    
            }
        }

        return respType;
    }
}
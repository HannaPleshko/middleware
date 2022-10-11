import { FolderChangeType, ItemIdType } from './../models/mail.model';
import {
    EmailAddressType, ContactItemType, ItemResponseShapeType, CalendarItemType, EffectiveRightsType, TaskType, ItemType,
    FolderType, BaseFolderIdType, DistinguishedFolderIdType, FolderIdType,TargetFolderIdType, FileAttachmentType,
    AttachmentIdType, NonEmptyArrayOfAttachmentsType, MessageType, ExtendedPropertyType, NonEmptyArrayOfPropertyValuesType,
    BaseFolderType, TimeZoneType, ContactsFolderType, FolderResponseShapeType, CalendarFolderType, TasksFolderType
} from '../models/mail.model';

import {
    MailboxTypeType, DefaultShapeNamesType, MapiPropertyTypeType, DistinguishedFolderIdNameType, UnindexedFieldURIType, FolderClassType,
    ItemClassType, ExtendedPropertyKeyType
} from '../models/enum.model';
import {EWSContactsManager} from '../utils';

import { Request } from '@loopback/rest';
import {
    PimItemFormat, PimItem, PimContact, KeepPimConstants, PimLabel,
    PimLabelTypes, PimNoticeTypes, isEmail,
    base64Encode, base64Decode, PimItemFactory, KeepPimLabelManager, UserInfo, PimCommonEventsJmap, isDevelopment, hasTimeZone
} from '@hcllabs/openclientkeepcomponent';
import { Logger } from './logger';
import { PathToUnindexedFieldType, PathToExtendedFieldType, SOAPFault } from '../models/common.model';
import util from 'util';
import { UserContext } from '../keepcomponent';
import RRule, { Frequency, Options, Weekday } from 'rrule';
import { 
    PimRecurrenceDayOfWeek, PimRecurrenceFrequency, PimRecurrenceRule 
} from '@hcllabs/openclientkeepcomponent/dist/keep/pim/jmap/PimRecurrenceRule';
import { DateTime, DateTimeOptions } from 'luxon';

//Default view to get labels
export const DEFAULT_LABEL_EXCLUSIONS = { views: [KeepPimConstants.ALL] };

// Type of labels to use for emails. Values are for extended properties with DistinguishedPropertySetId = "PublicStrings" 
// and PropertyName = "http://schemas.microsoft.com/entourage/emaillabeln", where n is 1, 2, or 3
export enum EmailLabelType {
    HOME = 'home',
    WORK = 'work',
    OTHER = 'other'
}

// TODO: How to do this with Keep? See https://jira.cwp.pnp-hcl.com/browse/LABS-1186
export function getEmail(user: string): EmailAddressType {
    const rtn = new EmailAddressType();
    rtn.Name = user;

    if (isEmail(user)) {
        rtn.EmailAddress = user;
    } else if (user.indexOf("Jane Doe") !== -1) {
        rtn.EmailAddress = "jane.doe@xxxxx.ngrok.io";
    } else if (user.indexOf("John Doe") !== -1) {
        rtn.EmailAddress = "john.doe@xxxxx.ngrok.io";
    } else if (user.indexOf("James Godwin") !== -1) {
        rtn.EmailAddress = "rusty.godwin@pnp-hcl.com";
    } else if (user.indexOf("David Kennedy") !== -1) {
        rtn.EmailAddress = "david.kennedy@pnp-hcl.com";
    } else if (user.indexOf("Tanya Anurag") !== -1) {
        rtn.EmailAddress = "tanya-anurag@pnp-hcl.com";
    } else if (user.indexOf("rustyg_mail") !== -1) {
        rtn.EmailAddress = "rustyg.mail@mail.quattro.rocks";
    } else if (user.toLowerCase().includes("rustyg mail")) {
        rtn.EmailAddress = "rustyg.mail@mail.quattro.rocks";
    } else if (user.indexOf("rogerw_mail") !== -1) {
        rtn.EmailAddress = "rogerw.mail@mail.quattro.rocks";
    } else if (user.toLowerCase().includes("rogerw mail")) {
        rtn.EmailAddress = "rogerw.mail@mail.quattro.rocks";
    } else if (user.indexOf("davek_mail") !== -1) {
        rtn.EmailAddress = "davek.mail@mail.quattro.rocks";
    } else if (user.toLowerCase().includes("davek mail")) {
        rtn.EmailAddress = "davek.mail@mail.quattro.rocks";
    }

    if (rtn.EmailAddress) {
        rtn.RoutingType = "SMTP";
        rtn.MailboxType = MailboxTypeType.MAILBOX;
    }
    else {
        rtn.MailboxType = MailboxTypeType.UNKNOWN;
    }

    return rtn;
}


/**
 * Get the value for an ExtendedFieldURI that uses a property set id and name.
 * @param item The item containing the extended properties.
 * @param identifiers An object containing the names of the identifiers, as keys, to check and what their values should be. Keys are defined by the ExtendedPropertyKeyType constants.
 */
export function getValueForExtendedFieldURI(item: ItemType, identifiers: any): any | undefined {
    if (item.ExtendedProperty) {
        for (const property of item.ExtendedProperty) {
            let match = true;
            for (const key in identifiers) {
                const value = identifiers[key];

                if ((key === ExtendedPropertyKeyType.PROPERTY_NAME && property.ExtendedFieldURI.PropertyName !== value) ||
                    (key === ExtendedPropertyKeyType.PROPERTY_SETID && property.ExtendedFieldURI.PropertySetId !== value) ||
                    (key === ExtendedPropertyKeyType.DISTINGUISHED_PROPERTY_SETID && property.ExtendedFieldURI.DistinguishedPropertySetId !== value) ||
                    (key === ExtendedPropertyKeyType.PROPERTY_TAG && property.ExtendedFieldURI.PropertyTag !== value) ||
                    (key === ExtendedPropertyKeyType.PROPERTY_ID && property.ExtendedFieldURI.PropertyId !== value)) {
                    match = false;
                }
            }

            if (match) {
                const numberTypes = [MapiPropertyTypeType.INTEGER, MapiPropertyTypeType.FLOAT, MapiPropertyTypeType.DOUBLE, MapiPropertyTypeType.LONG];

                if (property.ExtendedFieldURI.PropertyType === MapiPropertyTypeType.BOOLEAN && property.Value) {
                    return property.Value === "true" ? true : false;
                }
                else if (property.ExtendedFieldURI.PropertyType === MapiPropertyTypeType.SYSTEM_TIME && property.Value) {
                    return new Date(property.Value);
                }
                else if (property.ExtendedFieldURI.PropertyType === MapiPropertyTypeType.STRING) {
                    return property.Value;
                }
                else if (numberTypes.includes(property.ExtendedFieldURI.PropertyType) && property.Value) {
                    return Number.parseInt(property.Value);
                }
                else {
                    Logger.getInstance().warn(`Unsupported property type: ${property.ExtendedFieldURI.PropertyType}`);
                }
            }
        }
    }

    return undefined;
}

/**
 * Adds the extended properties in the EWS item to a PIM item.
 * @param item The EWS item containing the extended properties.
 * @param pimItem The pim item where the property is added.
 * @param identifiers An object containing the names of the identifiers, as keys, to check and what their values should be. Keys are defined by the ExtendedPropertyKeyType constants.
 */
export function addExtendedPropertiesToPIM(item: ItemType, pimItem: PimItem, identifiers?: any[]): void {

    if (item.ExtendedProperty) {
        item.ExtendedProperty.forEach(property => {
            let match = false;

            if (identifiers) {
                identifiers.forEach(identifier => {
                    if (identifier[ExtendedPropertyKeyType.DISTINGUISHED_PROPERTY_SETID] === property.ExtendedFieldURI.DistinguishedPropertySetId &&
                        identifier[ExtendedPropertyKeyType.PROPERTY_ID] === property.ExtendedFieldURI.PropertyId &&
                        identifier[ExtendedPropertyKeyType.PROPERTY_NAME] === property.ExtendedFieldURI.PropertyName &&
                        identifier[ExtendedPropertyKeyType.PROPERTY_TAG] === property.ExtendedFieldURI.PropertyTag &&
                        identifier[ExtendedPropertyKeyType.PROPERTY_SETID] === property.ExtendedFieldURI.PropertySetId) {
                        match = true;
                    }
                });
            }
            else {
                match = true;
            }

            if (match) {
                addExtendedPropertyToPIM(pimItem, property);
            }
        });
    }
}

/**
 * Add an EWS extended property to a PIM item.
 * @param pimItem The pim item where the property is added.
 * @param property The EWS extended property to add. 
 */
export function addExtendedPropertyToPIM(pimItem: PimItem, property: ExtendedPropertyType): void {
    const rtn: any = {};
    if (property.ExtendedFieldURI.DistinguishedPropertySetId) {
        rtn[ExtendedPropertyKeyType.DISTINGUISHED_PROPERTY_SETID] = property.ExtendedFieldURI.DistinguishedPropertySetId;
    }
    if (property.ExtendedFieldURI.PropertyId) {
        rtn[ExtendedPropertyKeyType.PROPERTY_ID] = property.ExtendedFieldURI.PropertyId;
    }
    if (property.ExtendedFieldURI.PropertySetId) {
        rtn[ExtendedPropertyKeyType.PROPERTY_SETID] = property.ExtendedFieldURI.PropertySetId;
    }
    if (property.ExtendedFieldURI.PropertyName) {
        rtn[ExtendedPropertyKeyType.PROPERTY_NAME] = property.ExtendedFieldURI.PropertyName;
    }
    if (property.ExtendedFieldURI.PropertyTag) {
        rtn[ExtendedPropertyKeyType.PROPERTY_TAG] = property.ExtendedFieldURI.PropertyTag;
    }
    if (property.ExtendedFieldURI.PropertyType) {
        rtn[ExtendedPropertyKeyType.PROPERTY_TYPE] = property.ExtendedFieldURI.PropertyType;
    }

    if (property.Value) {
        rtn["Value"] = property.Value;
    }
    else if (property.Values) {
        rtn["Value"] = property.Values.Value;
    }
    pimItem.addExtendedProperty(rtn);
}

/**
 * Return an object used to search for Extended properties
 * @param path The PathToExtendedFieldType property.
 */
export function identifiersForPathToExtendedFieldType(path: PathToExtendedFieldType): any {
    const identifiers: any = {};
    if (path.DistinguishedPropertySetId) {
        identifiers[ExtendedPropertyKeyType.DISTINGUISHED_PROPERTY_SETID] = path.DistinguishedPropertySetId;
    }
    if (path.PropertyId) {
        identifiers[ExtendedPropertyKeyType.PROPERTY_ID] = path.PropertyId;
    }
    if (path.PropertyName) {
        identifiers[ExtendedPropertyKeyType.PROPERTY_NAME] = path.PropertyName;
    }
    if (path.PropertySetId) {
        identifiers[ExtendedPropertyKeyType.PROPERTY_SETID] = path.PropertySetId;
    }
    if (path.PropertyTag) {
        identifiers[ExtendedPropertyKeyType.PROPERTY_TAG] = path.PropertyTag;
    }

    return identifiers;
}

/**
 * Returns the type of label to use for an email. 
 * @param email The email address
 * @param pimContact The Pim contact containing all email addresses
 * @returns The type of label for the email address or undefined if it can't be determined
 */
export function getLabelTypeForEmail(email: string, pimContact: PimContact): EmailLabelType | undefined {
    if (pimContact.mobileEmail === email || pimContact.schoolEmail === email) {
        return EmailLabelType.OTHER; 
    }
    else if (pimContact.workEmails.findIndex(aEmail => aEmail === email ) !== -1) {
        return EmailLabelType.WORK; 
    }
    else if (pimContact.homeEmails.findIndex(aEmail => aEmail === email) !== -1) {
        return EmailLabelType.HOME; 
    }
    else if (pimContact.otherEmails.findIndex(aEmail => aEmail === email) !== -1) {
        return EmailLabelType.OTHER; 
    }
    return undefined; 
}

/**
 * Returns the EWS access right for a PIM item.
 * @param pimItem The Pim item for the access rights
 */
export function effectiveRightsForItem(pimItem: PimItem): EffectiveRightsType {
    // FIXME: Temporarly allow all right. Need logic to calculate what is correct. 
    const rights = new EffectiveRightsType();
    rights.CreateAssociated = true;
    rights.CreateContents = true;
    rights.CreateHierarchy = true;
    rights.Delete = true;
    rights.Modify = true;
    rights.Read = true;
    rights.ViewPrivateItems = true;

    return rights;
}

/**
 * Adds additional properties to an EWS item. 
 * @param pimItem The PIM item containing the  properties
 * @param toItem The EWS item where the properties should be set. 
 * @param shape The shape setting used to determine which properties to add. If not specified, only extended properties saved in the PIM item will be added.
 */
export function addAdditionalPropertiesToEWSItem(pimItem: PimItem, toItem: ItemType, shape?: ItemResponseShapeType): void {

    if (shape?.AdditionalProperties) {
        for (const property of shape.AdditionalProperties.items) {
            if (property instanceof PathToUnindexedFieldType) {
                // Check for items that need special processing first
                if (property.FieldURI === UnindexedFieldURIType.ITEM_ITEM_CLASS) {
                    toItem.ItemClass = ewsItemClassType(pimItem);
                }
                else if (property.FieldURI === UnindexedFieldURIType.ITEM_EFFECTIVE_RIGHTS) {
                    toItem.EffectiveRights = effectiveRightsForItem(pimItem);
                }
                // TODO: Add other item fields defined in UnindexedFieldURIType here 

                else if (toItem instanceof CalendarItemType && pimItem.isPimCalendarItem()) {
                    // Process calendar FieldURIs here

                    if (property.FieldURI === UnindexedFieldURIType.CALENDAR_UID) {
                        const uid = pimItem.getAdditionalProperty(UnindexedFieldURIType.CALENDAR_UID);
                        if (uid) {
                            toItem.UID = uid;
                        }
                        else {
                            // Special processing for calendar item UID. If not saved in the PimItem, it may be created by a Domino client, so use the item UNID. 
                            toItem.UID = pimItem.unid;
                        }
                    }
                    else if (property.FieldURI === UnindexedFieldURIType.CALENDAR_IS_ALL_DAY_EVENT) {
                        toItem.IsAllDayEvent = pimItem.isAllDayEvent;
                    }
                    else if (property.FieldURI === UnindexedFieldURIType.CALENDAR_MEETING_TIME_ZONE) {
                        const tz = pimItem.getAdditionalProperty(UnindexedFieldURIType.CALENDAR_MEETING_TIME_ZONE);
                        if (tz) {
                            toItem.MeetingTimeZone = new TimeZoneType();
                            toItem.MeetingTimeZone.TimeZoneName = tz;
                        }
                    }
                    else if (property.FieldURI === UnindexedFieldURIType.CALENDAR_MEETING_REQUEST_WAS_SENT) {
                        const value = pimItem.getAdditionalProperty(UnindexedFieldURIType.CALENDAR_MEETING_REQUEST_WAS_SENT);
                        if (value !== undefined) {
                            toItem.MeetingRequestWasSent = value;
                        }
                    }
                    else if (property.FieldURI === UnindexedFieldURIType.CALENDAR_IS_RESPONSE_REQUESTED) {
                        const value = pimItem.getAdditionalProperty(UnindexedFieldURIType.CALENDAR_IS_RESPONSE_REQUESTED);
                        if (value !== undefined) {
                            toItem.IsResponseRequested = value;
                        }
                    }
                    else if (property.FieldURI === UnindexedFieldURIType.CALENDAR_APPOINTMENT_SEQUENCE_NUMBER) {
                        const value = pimItem.getAdditionalProperty(UnindexedFieldURIType.CALENDAR_APPOINTMENT_SEQUENCE_NUMBER);
                        if (value !== undefined) {
                            toItem.AppointmentSequenceNumber = value;
                        }
                    }
                    else if (property.FieldURI === UnindexedFieldURIType.CALENDAR_APPOINTMENT_STATE) {
                        const value = pimItem.getAdditionalProperty(UnindexedFieldURIType.CALENDAR_APPOINTMENT_STATE);
                        if (value !== undefined) {
                            toItem.AppointmentState = value;
                        }
                    }
    
                    // TODO: Add other calendar fields defined in UnindexedFieldURIType  
                }
                else if (toItem instanceof ContactItemType && pimItem.isPimContact()) {
                    // Process contact FieldURIs here

                    if (property.FieldURI === UnindexedFieldURIType.CONTACTS_COMPANY_NAME) {
                        const companyName = pimItem.companyName;
                        if (companyName) {
                            toItem.CompanyName = companyName;
                        }
                    }
                    // TODO: Add other contact fields defined in UnindexedFieldURIType
                }
                else if (toItem instanceof TaskType && pimItem.isPimTask()) {
                    // Process Task FieldURIs here

                    if (property.FieldURI === UnindexedFieldURIType.TASK_DUE_DATE) {
                        if (pimItem.due) {
                            toItem.DueDate = new Date(pimItem.due);
                        }
                    }
                    // TODO: Add other task fields defined in UnindexedFieldURIType
                }
                else if (toItem instanceof MessageType && (pimItem.isPimMessage() || pimItem.isPimNote())) {
                    // Process Message FieldURIs here

                    if (property.FieldURI === UnindexedFieldURIType.MESSAGE_IS_READ) {
                        toItem.IsRead = pimItem.isRead;
                    }
                    // TODO: Add other message fields defined in UnindexedFieldURIType
                }
                else {
                    Logger.getInstance().warn(`addExtendedPropertiesToEWSItem: Unsupported item type on FieldURI ${property.FieldURI}`);
                }
            }
            else if (property instanceof PathToExtendedFieldType) {
                const pimProperty = pimItem.findExtendedProperty(identifiersForPathToExtendedFieldType(property));
                if (pimProperty) {
                    addPimExtendedPropertyObjectToEWS(pimProperty, toItem);
                }
            }
        }
    }
    else if (shape === undefined || shape.BaseShape === DefaultShapeNamesType.ALL_PROPERTIES) {
        // If shape was not passed in, or it is AllProperties, add all exteneded properties to the EWS item
        pimItem.extendedProperties.forEach(property => {
            addPimExtendedPropertyObjectToEWS(property, toItem);
        });
    }
}

/**
 * Return the EWS ItemClassType for the pimItem instance.
 * @param pimItem  The instance of PimItem or one of its subclasses to determine the ItemClassType .
 * @returns The ItemClassType for the PimItem.
 */
export function ewsItemClassType(pimItem: PimItem): ItemClassType {
    if (pimItem.isPimCalendarItem()) {
        return ItemClassType.ITEM_CLASS_CALENDAR;
    } else if (pimItem.isPimContact()) {
        return ItemClassType.ITEM_CLASS_CONTACT;
    } else if (pimItem.isPimMessage()) {
        if (pimItem.isMeetingRequest()) {
            return ItemClassType.ITEM_CLASS_MEETING_REQUEST;
        } else if (pimItem.isMeetingResponse()) {

            if (pimItem.noticeType === PimNoticeTypes.USER_ACCEPTED || pimItem.noticeType === PimNoticeTypes.COUNTER_PROPOSAL_ACCEPTED) {
                return ItemClassType.ITEM_CLASS_ACCEPT_MEETING_RESPONSE;
            } else if (pimItem.noticeType === PimNoticeTypes.USER_DECLINED
                || pimItem.noticeType === PimNoticeTypes.USER_DELETED
                || pimItem.noticeType === PimNoticeTypes.COUNTER_PROPOSAL_DECLINED) {
                return ItemClassType.ITEM_CLASS_DECLINE_MEETING_RESPONSE;
            } else if (pimItem.noticeType === PimNoticeTypes.USER_TENTATIVELY_ACCEPTED) {
                return ItemClassType.ITEM_CLASS_TENTATIVE_MEETING_RESPONSE;
            } else if (pimItem.noticeType === PimNoticeTypes.EVENT_CANCELLED) {
                return ItemClassType.ITEM_CLASS_MEETING_CANCELED;
            } else {
                return ItemClassType.ITEM_CLASS_ACCEPT_MEETING_RESPONSE;
            }
        } else {
            return ItemClassType.ITEM_CLASS_MESSAGE;
        }
    } else if (pimItem.isPimNote()) {
        return ItemClassType.ITEM_CLASS_NOTE;
    } else if (pimItem.isPimTask()) {
        return ItemClassType.ITEM_CLASS_TASK;
    }
    return ItemClassType.ITEM_CLASS_UNKNOWN;
}

/**
 * Returns an EWS folder class for a PIM label. 
 * @param pimLabel The PIM label. 
 * @returns The FolderClassType of the label.
 */
export function folderClassForLabel(pimLabel: PimLabel): FolderClassType {

    if (pimLabel.type === PimLabelTypes.CALENDAR) {
        return FolderClassType.CALENDAR;
    }
    else if (pimLabel.type === PimLabelTypes.CONTACTS) {
        return FolderClassType.CONTACT;
    }
    else if (pimLabel.type === PimLabelTypes.JOURNAL) {
        return FolderClassType.NOTE;
    }
    else if (pimLabel.type === PimLabelTypes.TASKS) {
        return FolderClassType.TASK;
    }

    return FolderClassType.MESSAGE; // Must be a mail folder

}

/**
 * Map a FolderType object to a pim label type.
 * @param folderType The folder type object to map to pim label type 
 * @returns The corresponding pim label type for the folder type object.
 */
export function labelTypeFromFolderClass(folderType: BaseFolderType): PimLabelTypes {

    // Default to MAIL.  Note that calendars are handled differently.
    // TODO:  Update when Journal folders are implemented
    let labelType = PimLabelTypes.MAIL;

    if (folderType instanceof ContactsFolderType) {
        labelType = PimLabelTypes.CONTACTS;
    } else if (folderType instanceof TasksFolderType) {
        labelType = PimLabelTypes.TASKS;
    } else if (folderType instanceof CalendarFolderType) {
        labelType = PimLabelTypes.CALENDAR;
    } else if (folderType instanceof FolderType && (folderType.FolderClass !== undefined && folderType.FolderClass === FolderClassType.NOTE)) {
        labelType = PimLabelTypes.JOURNAL;
    }

    return labelType;
}

/**
 * Adds additional properties to an EWS folder. 
 * @param userInfo Information about the current user. 
 * @param pimItem The PIM label containing the  properties
 * @param toItem The EWS folder where the properties should be set. 
 * @param shape The shape setting used to determine which properties to add. If not specified, only extended properties saved in the PIM folder will be added. 
 */
export function addAdditionalPropertiesToEWSFolder(userInfo: UserInfo, pimItem: PimLabel, toItem: BaseFolderType, shape?: FolderResponseShapeType): void {

    if (shape?.AdditionalProperties) {
        for (const property of shape.AdditionalProperties.items) {
            if (property instanceof PathToUnindexedFieldType) {
                // Check for items that need special processing first
                if (property.FieldURI === UnindexedFieldURIType.FOLDER_FOLDER_CLASS) {
                    toItem.FolderClass = folderClassForLabel(pimItem);
                }
                else if (property.FieldURI === UnindexedFieldURIType.FOLDER_EFFECTIVE_RIGHTS) {
                    toItem.EffectiveRights = effectiveRightsForItem(pimItem);
                }
                else if (property.FieldURI === UnindexedFieldURIType.FOLDER_PARENT_FOLDER_ID) {
                    if (pimItem.parentFolderId) {
                        const combinedParentId = combineFolderIdAndMailbox(userInfo, pimItem.parentFolderId, toItem);
                        toItem.ParentFolderId = new FolderIdType(combinedParentId, `ck-${combinedParentId}`);
                    } else if (userInfo.userId) {
                        toItem.ParentFolderId = rootFolderIdForUser(userInfo.userId);
                    }
                }
                else if (property.FieldURI === UnindexedFieldURIType.FOLDER_DISPLAY_NAME) {
                    toItem.DisplayName = pimItem.displayName;
                }
                // TODO: Add other folder fields defined in UnindexedFieldURIType  

                else {
                    Logger.getInstance().warn(`addExtendedPropertiesToEWSFolder: Unsupported item type on FieldURI ${property.FieldURI}`);
                }
            }
            else if (property instanceof PathToExtendedFieldType) {
                const pimProperty = pimItem.findExtendedProperty(identifiersForPathToExtendedFieldType(property));
                if (pimProperty) {
                    addPimExtendedPropertyObjectToEWS(pimProperty, toItem);
                }
            }
        }
    }
    else if (shape === undefined) {
        // If shape was not passed in, add all exteneded properties to the EWS item
        pimItem.extendedProperties.forEach(property => {
            addPimExtendedPropertyObjectToEWS(property, toItem);
        });
    }

}

/**
 * Add an extended properties object for a Keep Pim Item to and EWS item. 
 * @param property The property object saved in a Keep PIM item.
 * @param toItem The EWS item.
 */
export function addPimExtendedPropertyObjectToEWS(property: any, toItem: ItemType): void {
    const propertyType = new ExtendedPropertyType();
    propertyType.ExtendedFieldURI = new PathToExtendedFieldType();
    if (property[ExtendedPropertyKeyType.DISTINGUISHED_PROPERTY_SETID]) {
        propertyType.ExtendedFieldURI.DistinguishedPropertySetId = property[ExtendedPropertyKeyType.DISTINGUISHED_PROPERTY_SETID];
    }
    if (property[ExtendedPropertyKeyType.PROPERTY_ID]) {
        propertyType.ExtendedFieldURI.PropertyId = property[ExtendedPropertyKeyType.PROPERTY_ID];
    }
    if (property[ExtendedPropertyKeyType.PROPERTY_NAME]) {
        propertyType.ExtendedFieldURI.PropertyName = property[ExtendedPropertyKeyType.PROPERTY_NAME];
    }
    if (property[ExtendedPropertyKeyType.PROPERTY_SETID]) {
        propertyType.ExtendedFieldURI.PropertySetId = property[ExtendedPropertyKeyType.PROPERTY_SETID];
    }
    if (property[ExtendedPropertyKeyType.PROPERTY_TAG]) {
        propertyType.ExtendedFieldURI.PropertyTag = property[ExtendedPropertyKeyType.PROPERTY_TAG];
    }
    if (property[ExtendedPropertyKeyType.PROPERTY_TYPE]) {
        propertyType.ExtendedFieldURI.PropertyType = property[ExtendedPropertyKeyType.PROPERTY_TYPE];
    }
    if (property[ExtendedPropertyKeyType.PROPERTY_ID]) {
        propertyType.ExtendedFieldURI.PropertyId = property[ExtendedPropertyKeyType.PROPERTY_ID];
    }

    const value: any = property["Value"];
    if (value instanceof Array) {
        const arrayValue = new NonEmptyArrayOfPropertyValuesType();
        arrayValue.Value = value;
        propertyType.Values = arrayValue;
    }
    else {
        propertyType.Value = value;
    }

    if (toItem.ExtendedProperty === undefined) {
        toItem.ExtendedProperty = [];
    }

    toItem.ExtendedProperty.push(propertyType);

}

/**
 * Returns the folder id for the root folder. 
 * @param request The current request being processed.
 * @param mailboxId SMTP mailbox delegator or delegatee address.
 */
export function rootFolderIdForRequest(request: Request, mailboxId?: string): FolderIdType {
    const userInfo = UserContext.getUserInfo(request);
    return rootFolderIdForUser(userInfo?.userId ?? "", mailboxId);
}

/**
 * Returns the folder id for the root folder. 
 * @param userId The current userid.
 * @param mailboxId SMTP mailbox delegator or delegatee address.
 */
export function rootFolderIdForUser(userId: string, mailboxId?: string): FolderIdType {
    const idEWS = getEWSId(`id-root-${userId}`, mailboxId);
    return new FolderIdType(idEWS, `ck-${idEWS}`);
}

/**
 * Determines if a folder id is for the root folder. 
 * @param folderId The folder id to check.
 * @param request The current request being processed. If not supplied then folderId must be a DistinguishedFolderIdName. 
 * @returns True if the folder id is for the root folder, otherwise false. 
 */
export function isRootFolderId(userInfo: UserInfo, folderIdType: BaseFolderIdType, request?: Request): boolean {
    if (folderIdType instanceof DistinguishedFolderIdType) {
        return (folderIdType.Id === DistinguishedFolderIdNameType.ROOT || folderIdType.Id === DistinguishedFolderIdNameType.MSGFOLDERROOT);
    }
    else if (request !== undefined && folderIdType instanceof FolderIdType) {
        const mailboxId = findMailboxId(userInfo, folderIdType);
        return (folderIdType.Id === rootFolderIdForRequest(request, mailboxId).Id);
    }

    return false;
}

/**
 * Determines if a target folder id is for the root folder. 
 * @param targetIdType The target folder id to check.
 * @param request The current request being processed. If not supplied then targetId must be a DistinguishedFolderIdName. 
 * @returns True if the target folder id is for the root folder, otherwise false. 
 */
export function isTargetRootFolderId(userInfo: UserInfo, targetId: BaseFolderIdType, request?: Request): boolean {
    if (targetId instanceof TargetFolderIdType) {
        if (targetId.DistinguishedFolderId) {
            return isRootFolderId(userInfo, targetId.DistinguishedFolderId, request);
        }
        else if (targetId.FolderId) {
            return isRootFolderId(userInfo, targetId.FolderId, request);
        }
        else if (targetId.AddressListId) {
            const folderId = new FolderIdType(targetId.AddressListId.Id);
            return isRootFolderId(userInfo, folderId);
        }
    } else if (targetId instanceof FolderIdType || targetId instanceof DistinguishedFolderIdType) {
        return isRootFolderId(userInfo, targetId, request);
    }
    return false;
}

/**
 * Returns a representation of the root folder. 
 * @param userInfo The current user information.
 * @param shape The shape of the returned folder.
 * @param labels The current list of labels. Use to calulate the child folder count. If not specified the child folder count will be zero.
 * @param mailboxId SMTP mailbox delegator or delegatee address.
 * @param request The oritinal request containing the item. 
 */
export function rootFolderForUser(
    userInfo: UserInfo, 
    shape: FolderResponseShapeType, 
    labels: PimLabel[] = [],
    mailboxId?: string,
    request?: Request
    ): FolderType {

    const folderId = rootFolderIdForUser(userInfo.userId, mailboxId);

    // Create a PimLabel representation of the root folder so we can use labelToFolder to convert it to a FolderType with the proper shape. 
    // Note: The display name never actually gets displayed,
    const pimRoot = PimItemFactory.newPimLabel({ "FolderId": folderId.Id, "DocumentCount": 0, "DisplayName": "Top of Information Store", "Type": PimLabelTypes.MAIL }, PimItemFormat.DOCUMENT);

    // Add mapi properties included with root folder. Defined in https://interoperability.blob.core.windows.net/files/MS-OXPROPS/%5BMS-OXPROPS%5D.pdf
    let property: any = {};
    // Root is not hidden
    property[ExtendedPropertyKeyType.PROPERTY_TAG] = '0x10f4';
    property[ExtendedPropertyKeyType.PROPERTY_TYPE] = MapiPropertyTypeType.BOOLEAN;
    property["Value"] = 'false';
    pimRoot.addExtendedProperty(property);

    property = {};
    // Set time of the most recent message change within the folder container to now. 
    property[ExtendedPropertyKeyType.PROPERTY_TAG] = '0x670a';
    property[ExtendedPropertyKeyType.PROPERTY_TYPE] = MapiPropertyTypeType.SYSTEM_TIME;
    property["Value"] = new Date().toISOString();
    pimRoot.addExtendedProperty(property);

    property = {};
    // Set total count of messages that have been deleted from a folder to zero. 
    property[ExtendedPropertyKeyType.PROPERTY_TAG] = '0x670b';
    property[ExtendedPropertyKeyType.PROPERTY_TYPE] = MapiPropertyTypeType.INTEGER;
    property["Value"] = '0';
    pimRoot.addExtendedProperty(property);

    // Convert the label to a EWS folder object
    return labelToFolder(userInfo, pimRoot, request, folderId, shape, labels) as FolderType;
}

/**
 * Maps a distinguished folder id to a Domino label view.  
 * @param folder The folder who's view name should be returned.
 * @returns The name of the view. If the distinguished folder id does not map to a Domino label view then the folder's display name will be returned or undefined if no display name is set.  
 */
export function getViewName(folder: FolderType): string | undefined {
    const dnToView: any = {};
    dnToView[DistinguishedFolderIdNameType.INBOX] = KeepPimConstants.INBOX;
    dnToView[DistinguishedFolderIdNameType.SENTITEMS] = KeepPimConstants.SENT;
    dnToView[DistinguishedFolderIdNameType.DELETEDITEMS] = KeepPimConstants.TRASH;
    dnToView[DistinguishedFolderIdNameType.DRAFTS] = KeepPimConstants.DRAFTS;
    dnToView[DistinguishedFolderIdNameType.JUNKEMAIL] = KeepPimConstants.JUNKMAIL;
    dnToView[DistinguishedFolderIdNameType.CALENDAR] = KeepPimConstants.CALENDAR;
    dnToView[DistinguishedFolderIdNameType.CONTACTS] = KeepPimConstants.CONTACTS;
    dnToView[DistinguishedFolderIdNameType.TASKS] = KeepPimConstants.TASKS;
    dnToView[DistinguishedFolderIdNameType.NOTES] = KeepPimConstants.JOURNAL;

    return (folder.DistinguishedFolderId !== undefined) ? (dnToView[folder.DistinguishedFolderId] ?? folder.DisplayName) : folder.DisplayName;
}

/**
 * Maps the Domino label view to a distinguished folder id.
 * TODO: Figure out what to do with these EWS folders: outbox, contacts, calendar, notes, tasks, journal, root, archiveroot, archivemsgfolderroot, archivedeleteditems
 * @param label The label who's DistinguishedFolderId should be returned. 
 * @returns The DistinguishedFolderId for the label or undefined if this label does not match to a DistinguishedFolderId.
 */
export function getDistinguishedFolderId(label: PimLabel): DistinguishedFolderIdNameType | undefined {
    const viewToFolder: any = {};
    viewToFolder[KeepPimConstants.INBOX] = DistinguishedFolderIdNameType.INBOX;
    viewToFolder[KeepPimConstants.SENT] = DistinguishedFolderIdNameType.SENTITEMS;
    viewToFolder[KeepPimConstants.TRASH] = DistinguishedFolderIdNameType.DELETEDITEMS;
    viewToFolder[KeepPimConstants.DRAFTS] = DistinguishedFolderIdNameType.DRAFTS;
    viewToFolder[KeepPimConstants.JUNKMAIL] = DistinguishedFolderIdNameType.JUNKEMAIL;
    viewToFolder[KeepPimConstants.CALENDAR] = DistinguishedFolderIdNameType.CALENDAR;
    viewToFolder[KeepPimConstants.CONTACTS] = DistinguishedFolderIdNameType.CONTACTS;
    viewToFolder[KeepPimConstants.TASKS] = DistinguishedFolderIdNameType.TASKS;
    viewToFolder[KeepPimConstants.JOURNAL] = DistinguishedFolderIdNameType.NOTES;

    return viewToFolder[label.view];
}

/**
 * Find a label matching a EWS folder id in a list of labels. 
 * @param labels The list of PIM labels to search.
 * @param folderId The id of a EWS folder. 
 * @returns The matching PIM label of undefined if not found. 
 */
export function findLabel(labels: PimLabel[], folderId: BaseFolderIdType): PimLabel | undefined {
    if (folderId instanceof DistinguishedFolderIdType) {
        return labels.find(f => {
            const labelDistinguishedFolderId = getDistinguishedFolderId(f);
            return folderId.Id === labelDistinguishedFolderId;
        });
    }
    else if (folderId instanceof FolderIdType) {
        const [itemId] = getKeepIdPair(folderId.Id);
        return labels.find(f => (itemId === f.folderId));
    }

    return undefined;
}

/**
 * Find all labels with a parent folder in a list of labels. 
 * @param labels The list of PIM labels to search.
 * @param parentId The folder id type of the parent folder. If the parent is the root, always pass in a DistinguishedFolderIdType for the root folder. 
 * @param deepSearch True will return all child folders under the parent folder. False will only return child folder 1 level deep. 
 * @param isRoot True for the root parent polder
 * @returns A list of PIM labels with the parent id. An empty list will be returned if none found.  
 */
export function findLabelsWithParent(
    userInfo: UserInfo,
    labels: PimLabel[], 
    parentId: BaseFolderIdType, 
    deepSearch = false, 
    request?: Request
): PimLabel[] {
    if (parentId instanceof FolderIdType) {
        if (isTargetRootFolderId(userInfo, parentId, request)) {
            if (deepSearch) return labels;
            else return labels.filter(label => { return (label.parentFolderId === undefined) });
        } else {
            const [itemId, mailboxId] = getKeepIdPair(parentId.Id);
            let found = labels.filter(label => { 
                return (label.parentFolderId && label.parentFolderId === itemId) 
            });
        
            if (deepSearch) {
                found.forEach(label => {
                    const parentEWSId = getEWSId(label.folderId, mailboxId);
                    const sub = findLabelsWithParent(userInfo, labels, new FolderIdType(parentEWSId), true, request);
                    if (sub.length > 0) {
                        found = found.concat(sub);
                    }
                });
            }
            return found;
        }
    } else if (parentId instanceof DistinguishedFolderIdType) {
        if (isTargetRootFolderId(userInfo, parentId)) {
            if (deepSearch) return labels; // A deep search of the root returns all labels
            else return labels.filter(label => { return (label.parentFolderId === undefined) }); // Return labels with no parent
        }
        else {
            const parentLabel = findLabel(labels, parentId);
            if (parentLabel) {
                let found = labels.filter(label => { return (label.parentFolderId && label.parentFolderId === parentLabel.folderId) });
                if (deepSearch) {
                    found.forEach(label => {
                        const sub = findLabelsWithParent(userInfo, labels, new FolderIdType(label.folderId), true, request);
                        if (sub.length > 0) {
                            found = found.concat(sub);
                        }
                    });
                }
                return found;
            }
        }
    }

    return [];
}

/**
 * Find label by target id.
 * @param targetId The target id
 * @returns label found or undefined if not found. 
 */
export function findLabelByTarget(labels: PimLabel[], targetId: TargetFolderIdType): PimLabel | undefined {
    if (targetId.FolderId) {
        return findLabel(labels, targetId.FolderId);
    }
    else if (targetId.DistinguishedFolderId) {
        return findLabel(labels, targetId.DistinguishedFolderId);
    }
    else if (targetId.AddressListId) {
        return labels.find(f => targetId.AddressListId?.Id === f.folderId);
    }

    return undefined;
}

export async function getLabelByTarget(targetId: TargetFolderIdType, userInfo: UserInfo): Promise<PimLabel | undefined> {

    if (targetId.DistinguishedFolderId) {
        // If the target is a distinguished folder id, we have to get all the labels to find the correct one
        const mailboxId = findMailboxId(userInfo, targetId.DistinguishedFolderId);
        const labels = await getKeepLabels(userInfo, mailboxId);
    
        return findLabelByTarget(labels, targetId);
    }
    else {
        // If the target is a folder id then we can use it to retrieve the single label. 
        let folderId: string | undefined = undefined;
        let mailboxId: string | undefined;

        if (targetId.FolderId) {
            [folderId, mailboxId] = getKeepIdPair(targetId.FolderId.Id);
        } else if (targetId.AddressListId) {
            folderId = targetId.AddressListId?.Id;
        }

        if (folderId) {
            return KeepPimLabelManager.getInstance().getLabel(userInfo, folderId, false, mailboxId)      
        }
    }

    return undefined;
}
/**
 * Create an EWS folder from a Keep label. 
 * @param userId The user for the request being processed. 
 * @param label The Keep label.
 * @param request The oritinal request containing the item.
 * @param item The source EWS item with combined EWS id.
 * @param folderShape Specifies the properties that should be included in the returned folder. All properties if not specified.
 * @param labels The list of Keep labels last retrieve from Keep. 
 * @returns A EWS folder representing the label.
 */
export function labelToFolder(
    userInfo: UserInfo,
    label: PimLabel,
    request?: Request,
    item?: BaseFolderIdType,
    folderShape?: FolderResponseShapeType,
    labels: PimLabel[] = []
): BaseFolderType {

    const distinguishedFolderId = getDistinguishedFolderId(label);

    const baseShape = folderShape?.BaseShape ?? DefaultShapeNamesType.ALL_PROPERTIES;

    /*
     When label represents the root folder and the shape is not ID ONLY we must return the following format. Note: No distinguished id. 
        <t:Folder>
            <t:FolderId Id="<id>" ChangeKey="<ck-id>" />
            <t:DisplayName>Top of Information Store</t:DisplayName>
            <t:TotalCount>0</t:TotalCount>
            <t:ChildFolderCount>nn</t:ChildFolderCount>
            <t:UnreadCount>0</t:UnreadCount>
        </t:Folder>
     */
    let mailboxId: string | undefined;
    if (item) mailboxId = findMailboxId(userInfo, item);
    const isRootFolder = rootFolderIdForUser(userInfo.userId, mailboxId).Id === label.folderId;
    const folder = newFolderInstance(baseShape, isRootFolder ? DistinguishedFolderIdNameType.MSGFOLDERROOT : distinguishedFolderId, label.type);

    // Always include folder id
    if (label.folderId) {
        const combineFolderIds = combineFolderIdAndMailbox(userInfo, label.folderId, item);
        folder.FolderId = isRootFolder ? rootFolderIdForUser(userInfo.userId, mailboxId) : new FolderIdType(combineFolderIds, `ck-${combineFolderIds}`);

        if (baseShape !== DefaultShapeNamesType.ID_ONLY) {
            // Include child folder count for everything but ID ONLY
            let folderId: BaseFolderIdType;
            if (isRootFolder) {
                // For the root folder, findLabelsWithParent requires a distinguished folder id. Only use this to search, don't set it in the folder.
                const distinguishedFolderIdType = new DistinguishedFolderIdType();
                distinguishedFolderIdType.Id = DistinguishedFolderIdNameType.MSGFOLDERROOT;
                folderId = distinguishedFolderIdType;
            }
            else {
                folderId = folder.FolderId;
            }

            const children = findLabelsWithParent(userInfo, labels, folderId, false, request);
            folder.ChildFolderCount = children.length;
        }
    }
    else {
        Logger.getInstance().error(`Folder id is not set for label: ${util.inspect(label, false, 5)}`);
    }

    if (baseShape !== DefaultShapeNamesType.ID_ONLY) {
        /*
        The data to return for shapes is defined here: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/baseshape
        */

        folder.DisplayName = label.displayName;

        if (folder instanceof CalendarFolderType === false || baseShape === DefaultShapeNamesType.ALL_PROPERTIES) {
            folder.TotalCount = label.documentCount;
        }

        if (folder instanceof FolderType && label.type !== PimLabelTypes.JOURNAL) {
            folder.UnreadCount = label.unreadCount;
        }

        if (folder instanceof TasksFolderType) {
            // TODO: How to get/set the past due count?
        }

        if (label.parentFolderId) {
            const combineFolderIds = combineFolderIdAndMailbox(userInfo, label.parentFolderId, item);
            folder.ParentFolderId = new FolderIdType(combineFolderIds, `ck-${combineFolderIds}`);
        } else {
            // No parent set, it will be a child of the root if it is not the root.
            const parent = rootFolderIdForUser(userInfo.userId, mailboxId);
            if (parent.Id !== label.folderId) {
                folder.ParentFolderId = parent; // The label is not for the root folder
            }
        }
    }

    if (folderShape !== undefined) {
        // Set additional properties specified on shape
        addAdditionalPropertiesToEWSFolder(userInfo, label, folder, folderShape);
    }

    return folder;
}

/**
 * Create and return a new instance of a EWS folder. 
 * @param shape The information that needs to be included in the new folder.
 * @param distinguishedFolderId Optional distinguished folder name for the new folder. When creating a BaseFolderType for the root folder, 
 * always set this to DistinguishedFolderIdNameType.MSGFOLDERROOT or DistinguishedFolderIdNameType.ROOT so the returned folder is formatted correctly.
 * @param labelType PIM label type of the folder. Required only if distinguishedFolderId is not set.
 */
function newFolderInstance(shape: DefaultShapeNamesType, distinguishedFolderIdName?: DistinguishedFolderIdNameType, labelType?: PimLabelTypes): BaseFolderType {

    let folder: BaseFolderType;
    let folderClass: string | undefined = undefined; // Folder Class is defined here: https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxosfld/68a85898-84fe-43c4-b166-4711c13cdd61
    let isRootFolder = false;

    if (distinguishedFolderIdName !== undefined) {
        isRootFolder = (distinguishedFolderIdName === DistinguishedFolderIdNameType.ROOT || distinguishedFolderIdName === DistinguishedFolderIdNameType.MSGFOLDERROOT);
        if (isRootFolder) {
            folder = new FolderType();
            // No folder class is defined for the root folder
        }
        else if (distinguishedFolderIdName === DistinguishedFolderIdNameType.CALENDAR) {
            folder = new CalendarFolderType();
            folderClass = FolderClassType.CALENDAR;
        }
        else if (distinguishedFolderIdName === DistinguishedFolderIdNameType.CONTACTS) {
            folder = new ContactsFolderType();
            folderClass = FolderClassType.CONTACT;
        }
        else if (distinguishedFolderIdName === DistinguishedFolderIdNameType.TASKS) {
            folder = new TasksFolderType();
            folderClass = FolderClassType.TASK;
        }
        else if (distinguishedFolderIdName === DistinguishedFolderIdNameType.NOTES) {
            folder = new FolderType();
            folderClass = FolderClassType.NOTE;
        }
        else {
            folder = new FolderType();
            folderClass = FolderClassType.MESSAGE;
        }
    }
    else {
        if (labelType === undefined) {
            Logger.getInstance().error("The label type is not specified.");
            // Return a default.
            folder = new FolderType();
            folderClass = FolderClassType.MESSAGE;
        }
        else {
            switch (labelType) {
                case PimLabelTypes.CONTACTS: {
                    folder = new ContactsFolderType();
                    folderClass = FolderClassType.CONTACT
                    break;
                }
                case PimLabelTypes.JOURNAL: {
                    folder = new FolderType();
                    folderClass = FolderClassType.NOTE;
                    break;
                }
                case PimLabelTypes.TASKS: {
                    folder = new TasksFolderType();
                    folderClass = FolderClassType.TASK;
                    break;
                }
                case PimLabelTypes.CALENDAR: {
                    folder = new CalendarFolderType();
                    folderClass = FolderClassType.CALENDAR;
                    break;
                }
                default: {
                    folder = new FolderType();
                    folderClass = FolderClassType.MESSAGE;
                    break;
                }
            }
        }
    }

    if (shape !== DefaultShapeNamesType.ID_ONLY) {
        if (folderClass !== undefined) {
            folder.FolderClass = folderClass;
        }
        if (distinguishedFolderIdName && !isRootFolder) {
            // The returned root folder should not contain a distinguished folder id
            folder.DistinguishedFolderId = distinguishedFolderIdName;
        }
    }

    return folder;
}

/**
 * Returns if the read count should be include in a folder response. 
 * @param shape The shape of the expected folder.
 */
export function includeUnReadCountForFolders(shape: FolderResponseShapeType): boolean {
    return (shape.BaseShape !== undefined && shape.BaseShape !== DefaultShapeNamesType.ID_ONLY);
}

/**
 * Return a pair with the first entry being the parentId and the second being the attachmentName. 
 * @param combinedIds The combined parentId and attachmentName in a base64Encoded string.
 * @returns Return a pair with the first entry being the parentId and the second being the attachmentName. 
 */
export function getAttachmentIdPair(combinedIds: string): [string | undefined, string | undefined] {
    const separator = "##";
    let parentId: string | undefined;
    let attachmentName: string | undefined;
    if (combinedIds) {
        const idInfo = base64Decode(combinedIds);
        const separatorIndex = idInfo.indexOf(separator);
        if (separatorIndex >= 0) {
            parentId = idInfo.substring(0, separatorIndex);
            if (parentId.length === 0) parentId = undefined;
            attachmentName = idInfo.substring(separatorIndex + separator.length);
            if (attachmentName && attachmentName.length === 0) {
                attachmentName = undefined;
            }
        } else {
            attachmentName = idInfo;
        }
    }
    return [parentId, attachmentName];
}

/**
 * Return an id of combined parent id and attachmentName. 
 * @param parentId Id of the item containing  the attachment
 * @param attachmentName Name of the attachment in the containing item.
 * @returns The combined parentId and attachmentName in a base64Encoded string. 
 */
export function getAttachmentId(parentId: string, attachmentName: string): string {
    const separator = "##";
    if (!parentId) return base64Encode(attachmentName);
    return base64Encode((parentId ?? "") + separator + (attachmentName ?? ""));
}

/**
 * Convert PimItem attachments to an EWS AttachmentType array 
 * @param pimItem The subclass of pimItem holding the attachments
 * @param attachmentName Name of the attachment in the containing item.
 * @returns Return an array of attachment types or undefined. 
 */
export function getItemAttachments(pimItem: PimItem): NonEmptyArrayOfAttachmentsType | undefined {

    const pAttachments = pimItem.attachments;
    if (pAttachments && pAttachments.length > 0) {
        const attTypes = new NonEmptyArrayOfAttachmentsType();
        for (const pAtt of pAttachments) {
            const fAttachmentType = new FileAttachmentType();
            fAttachmentType.AttachmentId = new AttachmentIdType();
            fAttachmentType.AttachmentId.Id = getAttachmentId(pimItem.unid, pAtt);
            fAttachmentType.AttachmentId.RootItemId = pimItem.unid;
            fAttachmentType.AttachmentId.RootItemChangeKey = `ck-${pimItem.unid}`;
            fAttachmentType.Name = pAtt;
            attTypes.push(fAttachmentType);
        }
        return attTypes;
    }

    return undefined;
}

/**
 * Create a PimLabel from an EWS folder item. 
 * @param folder The source EWS folder item.
 * @param request The oritinal request containing the item.
 * @returns A PimLabel object representing the item.
 */
export function pimLabelFromItem(item: FolderType, existing?: any, parentLabelType?: PimLabelTypes): PimLabel {

    const labelObject: any = existing ? existing : {};
    if (parentLabelType) {
        labelObject["DesignType"] = parentLabelType;
    }

    if (item.DisplayName) {
        labelObject["DisplayName"] = item.DisplayName;
    }

    if (item.ParentFolderId) {
        labelObject["ParentId"] = item.ParentFolderId;
    }

    const pimLabel = PimItemFactory.newPimLabel(labelObject, PimItemFormat.DOCUMENT);

    // Add any extended fields
    addExtendedPropertiesToPIM(item, pimLabel);

    return pimLabel;

}

/**
 * Do common error handling for response to client 
 */
export function throwErrorIfClientResponse(error: any): void {
    if (error instanceof SOAPFault || (error.status && error.status >= 400 && error.status !== 404)) {
        throw error;
    }
}

/**
 * Return the date the item was created...or fallback to other dates if there is no created date
 * @param pimItem The PIM item from which to get the created date
 * @returns a Date....if there is not createdDate, it returns the lastModifiedDate and if that does not exist it returns today
 */
export function getFallbackCreatedDate(pimItem: PimItem): Date {
    // This function can be changed to return just the pimItem.createdDate when LABS_1863 is fixed.
    if (!pimItem.createdDate) {
        pimItem.createdDate = pimItem.lastModifiedDate ?? new Date();
    }
    return pimItem.createdDate;
}

/**
 * Get a RRule from a PIM recurrenceRule. 
 * @param rule The PIM recurrence rule.
 * @param pimItem The pim Item the rule belongs to. 
 * @returns A RRule for the PIM recurrence rule.
 * @throws An error if the PIM item does not have a start time or the rule is not valid. 
 */
export function pimRecurrenceRuleToRRule(rule: PimRecurrenceRule, pimItem: PimCommonEventsJmap): RRule {

    if (pimItem.start === undefined) {
        throw new Error(`Unable to determine recurrence for item ${pimItem.unid} since it has no start time`);
    }

    const weekdayMap: any = {}
    weekdayMap[PimRecurrenceDayOfWeek.MONDAY] = 0;
    weekdayMap[PimRecurrenceDayOfWeek.TUESDAY] = 1;
    weekdayMap[PimRecurrenceDayOfWeek.WEDNESDAY] = 2;
    weekdayMap[PimRecurrenceDayOfWeek.THURSDAY] = 3;
    weekdayMap[PimRecurrenceDayOfWeek.FRIDAY] = 4;
    weekdayMap[PimRecurrenceDayOfWeek.SATURDAY] = 5;
    weekdayMap[PimRecurrenceDayOfWeek.SUNDAY] = 6;

    let until: Date | null = null; 
    if (rule.until !== undefined) {
        const opts: DateTimeOptions = hasTimeZone(rule.until) ? {setZone: true} : {zone: pimItem.startTimeZone}
        const dt = DateTime.fromISO(rule.until, opts);
        until = dt.toJSDate();
    }
    
    const options: Options = {
        freq: getRRuleFrequency(rule.frequency),
        interval: rule.interval,
        dtstart: new Date(pimItem.start),
        wkst: null,
        count: rule.count ?? null,
        until,
        tzid: pimItem.startTimeZone ?? null,
        bysetpos: rule.bySetPosition ?? null,
        byyearday: rule.byYearDay ?? null,
        bymonth: rule.byMonth === undefined ? null : rule.byMonth.map(month => Number.parseInt(month)),
        bymonthday: rule.byMonthDay ?? null,
        bynmonthday: null, // RRule internal use only
        byweekno: rule.byWeekNo ?? null,
        byweekday: rule.byDay === undefined ? null : rule.byDay.map(day => new Weekday(weekdayMap[day.day], day.nthOfPeriod)),
        bynweekday: null, // RRule internal use only
        byhour: rule.byHour ?? null,
        byminute: rule.byMinute ?? null,
        bysecond: rule.bySecond ?? null,
        byeaster: null // Not supported
    };

    return new RRule(options);
}

/**
 * Conver the Keep Pim recurrence frequency to a RRule recurrence frequency. 
 * @param frequency The recurrence frequency from the Pim calendar item. 
 * @returns The RRule recurrence frequency 
 * @throws An error if frequency is set to an unknown value.  
 */
export function getRRuleFrequency(frequency: PimRecurrenceFrequency): Frequency {
    switch (frequency) {
        case PimRecurrenceFrequency.YEARLY:
            return Frequency.YEARLY;

        case PimRecurrenceFrequency.MONTHLY:
            return Frequency.MONTHLY;

        case PimRecurrenceFrequency.WEEKLY:
            return Frequency.WEEKLY;

        case PimRecurrenceFrequency.DAILY:
            return Frequency.DAILY;

        case PimRecurrenceFrequency.HOURLY:
            return Frequency.HOURLY;

        case PimRecurrenceFrequency.MINUTELY:
            return Frequency.MINUTELY;

        case PimRecurrenceFrequency.SECONDLY:
            return Frequency.SECONDLY;

        default:
            throw new Error(`Unknown recurrence frequency: ${frequency}`);
    }
}

/**
 * Return an id of combined unid and mailboxId to determine the delegator's or delegate's folder. 
 * @param itemId The delegator or delegatee item or the folder unid.
 * @param mailboxId SMTP mailbox delegator or delegatee address.
 * @returns The delegator or delegatee combined unid and mailboxId in a base64Encoded string. 
 */
 export function getEWSId(itemId: string, mailboxId?: string): string {
    if (!itemId) throw new Error('Item id cannot be undefined');
    if (!mailboxId) return base64Encode(itemId);
    const separator = "##";
    return base64Encode(itemId + separator + mailboxId);
}

/**
 * Return a pair with the first entry being the unid and the second being the mailboxId of the delegatee's folder. 
 * @param combinedIds The combined delegator's or delegate's unid and mailboxId in a base64Encoded string.
 * @returns A pair with the first entry being the unid and the second being the mailboxId of the delegator or delegatee.
 */
 export function getKeepIdPair(combinedIds: string): [string | undefined, string | undefined] {
    const separator = "##";
    let itemId: string | undefined;
    let mailboxId: string | undefined;

    if (combinedIds) {
        const idInfo = base64Decode(combinedIds);
        const separatorIndex = idInfo.indexOf(separator);

        if (separatorIndex >= 0) {
            itemId = idInfo.substring(0, separatorIndex);
            mailboxId = idInfo.substring(separatorIndex + separator.length);
        } else {
            itemId = idInfo;
        }
    }
     
    return [itemId, mailboxId];
}

/**
 * Closure with a collection of labels, where the keys are the  delegator's or delegate's mailboxIds .
 * @param userInfo The Domino username and password.
 * @param initialMailboxId The first EWS folder mailbox id.
 * @param includeUnread The read count should be include in a folder or not.
 * @returns A function that finds labels by delegator's or delegate's mailboxId inside a collection
 * Return labels from collection with the mailboxId key 
 * @param itemMailboxId The mailbox id of each folder
 * @returns delegator or delegatee PIM labels. 
 */
export async function getFolderLabelsFromHash(
    userInfo: UserInfo,
    initialMailboxId?: string,
    includeUnread = false,
): Promise<(itemMailboxId: string | undefined) => Promise<PimLabel[]>> {
    const labelsHash = new Map();
    const labelsInitial = await getKeepLabels(userInfo, initialMailboxId, includeUnread);
    const initialKey = initialMailboxId ? initialMailboxId : 'initialKey';
    labelsHash.set(initialKey, labelsInitial);

    return async function(itemMailboxId: string | undefined): Promise<PimLabel[]> {
        if (!itemMailboxId) return labelsHash.get(initialKey);
        if (!labelsHash.has(itemMailboxId)) {
            const labels = await getKeepLabels(userInfo, itemMailboxId, includeUnread);
            labelsHash.set(itemMailboxId, labels);
        }

        return labelsHash.get(itemMailboxId);
    }
}

/**
* Get the list of labels.
* @param userInfo The Domino username and password
* @param mailboxId SMTP mailbox delegator or delegatee address.
* @param includeUnread The read count should be include in a folder or not. 
* @returns An array of label objects.
*/
export async function getKeepLabels(userInfo: UserInfo, mailboxId?: string, includeUnread = false): Promise<PimLabel[]> {
    return KeepPimLabelManager
        .getInstance()
        .getLabels(userInfo, includeUnread, DEFAULT_LABEL_EXCLUSIONS, mailboxId);
}

/**
 * Find mailboxId of the folder.
 * @param folder The source EWS folder item with combined EWS id.
 * @returns MailboxId SMTP mailbox delegator or delegatee address.
 */
export function findMailboxId(userInfo: UserInfo, folder: BaseFolderIdType): string | undefined {
    let mailboxId: string | undefined;

    if (folder instanceof DistinguishedFolderIdType) {
        mailboxId = folder.Mailbox?.EmailAddress;
    } else if (folder instanceof FolderIdType || folder instanceof ItemIdType) {
        [, mailboxId] = getKeepIdPair(folder.Id);
    } else if (folder instanceof FolderChangeType 
        || folder instanceof TargetFolderIdType) {
        if (folder.FolderId) {
            [, mailboxId] = getKeepIdPair(folder.FolderId.Id);
        } else if (folder.DistinguishedFolderId) {
            mailboxId = folder.DistinguishedFolderId.Mailbox?.EmailAddress;
        }
    }

    // Current authenticated user's mailbox is always undefined
    if (mailboxId === userInfo.userId) {
        return undefined;
    }

    if (isDevelopment() && mailboxId && (mailboxId.endsWith('ngrok.io') || mailboxId.indexOf('@') < 0)) {
        const atIndex = mailboxId.indexOf('@');
        let mbId = mailboxId;
        if (atIndex > 0) {
            mbId = mbId.substring(0, atIndex);
        }
        mailboxId = mbId.replace('.',' ');
    }

    return mailboxId;
}

/**
 * Combine mailboxId and target folder id.
 * @param folder The source EWS folder item with combined EWS id.
 * @param targetId The uid of the source EWS folder item that shoul be combined with mailboxId.
 * @returns The combined EWS Id for the target.
 */
export function combineFolderIdAndMailbox(userInfo: UserInfo, targetId: string, folder?: BaseFolderIdType): string {
    let mailboxId: string | undefined;
    if (folder) mailboxId = findMailboxId(userInfo, folder);

    return getEWSId(targetId, mailboxId);
}

/**
 * Extract height and width from SizeRequested value
 * @param sizeRequestedValue The text value of SizeRequested element
 * @returns A [height, width] array
 */
 export function parseSizeRequestedValue(sizeRequestedValue: string): [number, number] | [undefined, undefined] {
    // To ensure that the recieved value has correct form (HR48x48, HR240x240 and etc.)
    const regexp = /hr[0-9]+x[0-9]+/i;
    if (!sizeRequestedValue.match(regexp)) {
        return [undefined, undefined]
    }
    const [height, width] = sizeRequestedValue.toLowerCase().slice(2).split('x').map(str => parseInt(str, 10));
    return [height, width];
}
  
/**
 * Filter PIM items by folder Ids. 
 * @param pimItems The list of PIM items to filter.
 * @param parentId The folder id type of the parent folder. 
 * @returns A list of PIM items with the parent id. An empty list will be returned if none found.  
 */
 export function filterPimContactsByFolderIds(pimItems: PimContact[], parentIds: BaseFolderIdType[]): PimContact[] {
    let resPimItems: PimContact[] = [];
    for (const parentId of parentIds) {
        let parentFolderId: any;
        if (parentId instanceof FolderIdType) {
            const [_parentFolderId] = getKeepIdPair(parentId.Id);
            parentFolderId = _parentFolderId;
        }else if (parentId instanceof DistinguishedFolderIdType) {    
            parentFolderId = parentId.Id;
        }
        if (parentFolderId){
            const foundPimItems = pimItems.filter(pimItem => { return (pimItem.parentFolderIds && pimItem.parentFolderIds.includes(parentFolderId)) });
            resPimItems = resPimItems.concat(foundPimItems);
        }
    }
    return resPimItems;
}

export async function createContactItemFromPimContact(pimContact: PimContact, userInfo: UserInfo, request: Request, returnFullContactData: boolean, contactDataShape: DefaultShapeNamesType): Promise<ContactItemType> {
    const shape = new ItemResponseShapeType();
    shape.BaseShape = returnFullContactData ? DefaultShapeNamesType.ALL_PROPERTIES : contactDataShape;

    return EWSContactsManager.getInstance().pimItemToEWSItem(pimContact, userInfo, request, shape);
}
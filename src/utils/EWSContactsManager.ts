import { 
    ContactAddressKey, convertToListObject, KeepPimConstants, KeepPimContactsManager, KeepPimLabelManager, PimAddress,
    PimContact, PimItemFactory, PimItemFormat, PimLabel, UserInfo 
} from '@hcllabs/openclientkeepcomponent';
import { 
    addExtendedPropertiesToPIM, addExtendedPropertyToPIM, EWSServiceManager, getItemAttachments, getLabelTypeForEmail,
    getValueForExtendedFieldURI, identifiersForPathToExtendedFieldType, Logger 
} from '.';
import { PathToExtendedFieldType, PathToIndexedFieldType } from '../models/common.model';
import { 
    BodyTypeType, DefaultShapeNamesType, DictionaryURIType, DistinguishedPropertySetType, EmailAddressKeyType, 
    ExtendedPropertyKeyType, FieldIndexValue, ImAddressKeyType, MailboxTypeType, MapiPropertyIds, MapiPropertyTypeType,
    MessageDispositionType, PhoneNumberKeyType, PhysicalAddressKeyType, SensitivityChoicesType, UnindexedFieldURIType 
} from '../models/enum.model';
import { 
    AppendToItemFieldType, BodyType, CompleteNameType, ContactItemType, DeleteItemFieldType, EmailAddressDictionaryEntryType,
    EmailAddressDictionaryType, FolderIdType, ImAddressDictionaryEntryType, ImAddressDictionaryType, ItemChangeType, 
    ItemIdType, ItemResponseShapeType, ItemType, PhoneNumberDictionaryEntryType, PhoneNumberDictionaryType, 
    PhysicalAddressDictionaryEntryType, PhysicalAddressDictionaryType, SetItemFieldType 
} from '../models/mail.model';
import { getEWSId, getKeepIdPair } from './pimHelper';
import { Request } from '@loopback/rest';
import * as util from 'util';
import { fromString } from 'html-to-text';

/**
 * This class is an EWSServicemanager subclass responsible for managing Contacts.
 */
export class EWSContactsManager extends EWSServiceManager {

    /**
     * Static functions and properties
     */
    private static instance: EWSContactsManager;

    public static getInstance(): EWSContactsManager {
        if (!this.instance) {
            this.instance = new EWSContactsManager();
        }
        return this.instance;
    }

    // Fields included for DEFAULT shape for notes
    private static defaultFields = [
        ...EWSServiceManager.idOnlyFields,
        UnindexedFieldURIType.CONTACTS_COMPANY_NAME,
        UnindexedFieldURIType.CONTACTS_COMPLETE_NAME,
        UnindexedFieldURIType.CONTACTS_EMAIL_ADDRESSES,
        UnindexedFieldURIType.CONTACTS_FILE_AS,
        UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS,
        UnindexedFieldURIType.CONTACTS_IM_ADDRESSES,
        UnindexedFieldURIType.CONTACTS_JOB_TITLE,
        UnindexedFieldURIType.CONTACTS_PHONE_NUMBERS,
        UnindexedFieldURIType.CONTACTS_PHYSICAL_ADDRESSES
    ];

    // Fields included for ALL_PROPERTIES shape for notes
    private static allPropertiesFields = [
        ...EWSContactsManager.defaultFields,
        UnindexedFieldURIType.CONTACTS_ALIAS,
        UnindexedFieldURIType.CONTACTS_ASSISTANT_NAME,
        UnindexedFieldURIType.CONTACTS_BIRTHDAY,
        UnindexedFieldURIType.CONTACTS_WEDDING_ANNIVERSARY,
        UnindexedFieldURIType.ITEM_BODY,
        UnindexedFieldURIType.CONTACTS_BUSINESS_HOME_PAGE,
        UnindexedFieldURIType.CONTACTS_CHILDREN,
        UnindexedFieldURIType.CONTACTS_CONTACT_SOURCE,
        UnindexedFieldURIType.ITEM_CONVERSATION_ID,
        UnindexedFieldURIType.CONTACTS_CULTURE,
        UnindexedFieldURIType.ITEM_DATE_TIME_CREATED,
        UnindexedFieldURIType.ITEM_DATE_TIME_RECEIVED,
        UnindexedFieldURIType.ITEM_DATE_TIME_SENT,
        UnindexedFieldURIType.CONTACTS_DEPARTMENT,
        UnindexedFieldURIType.CONTACTS_DIRECTORY_ID,
        UnindexedFieldURIType.CONTACTS_DIRECT_REPORTS,
        UnindexedFieldURIType.ITEM_DISPLAY_CC,
        UnindexedFieldURIType.CONTACTS_DISPLAY_NAME,
        UnindexedFieldURIType.ITEM_DISPLAY_TO,
        UnindexedFieldURIType.ITEM_EFFECTIVE_RIGHTS,
        UnindexedFieldURIType.CONTACTS_FILE_AS_MAPPING,
        UnindexedFieldURIType.CONTACTS_GENERATION,
        UnindexedFieldURIType.CONTACTS_GIVEN_NAME,
        UnindexedFieldURIType.CONTACTS_HAS_PICTURE,
        UnindexedFieldURIType.ITEM_IMPORTANCE,
        UnindexedFieldURIType.CONTACTS_INITIALS,
        UnindexedFieldURIType.ITEM_IN_REPLY_TO,
        UnindexedFieldURIType.ITEM_INTERNET_MESSAGE_HEADERS,
        UnindexedFieldURIType.ITEM_IS_ASSOCIATED,
        UnindexedFieldURIType.ITEM_IS_DRAFT,
        UnindexedFieldURIType.ITEM_IS_FROM_ME,
        UnindexedFieldURIType.ITEM_IS_RESEND,
        UnindexedFieldURIType.ITEM_IS_SUBMITTED,
        UnindexedFieldURIType.ITEM_IS_UNMODIFIED,
        UnindexedFieldURIType.ITEM_LAST_MODIFIED_NAME,
        UnindexedFieldURIType.ITEM_LAST_MODIFIED_TIME,
        UnindexedFieldURIType.CONTACTS_MANAGER,
        UnindexedFieldURIType.CONTACTS_MIDDLE_NAME,
        UnindexedFieldURIType.CONTACTS_MILEAGE,
        UnindexedFieldURIType.CONTACTS_NICKNAME,
        UnindexedFieldURIType.CONTACTS_NOTES,
        UnindexedFieldURIType.CONTACTS_OFFICE_LOCATION,
        UnindexedFieldURIType.CONTACTS_PHONETIC_FIRST_NAME,
        UnindexedFieldURIType.CONTACTS_PHONETIC_FULL_NAME,
        UnindexedFieldURIType.CONTACTS_PHONETIC_LAST_NAME,
        UnindexedFieldURIType.CONTACTS_PHOTO,
        UnindexedFieldURIType.CONTACTS_POSTAL_ADDRESS_INDEX,
        UnindexedFieldURIType.CONTACTS_PROFESSION,
        UnindexedFieldURIType.ITEM_REMINDER_DUE_BY,
        UnindexedFieldURIType.ITEM_REMINDER_IS_SET,
        UnindexedFieldURIType.ITEM_REMINDER_MINUTES_BEFORE_START,
        UnindexedFieldURIType.ITEM_RESPONSE_OBJECTS,
        UnindexedFieldURIType.ITEM_SENSITIVITY,
        UnindexedFieldURIType.ITEM_SIZE,
        UnindexedFieldURIType.ITEM_CATEGORIES,
        UnindexedFieldURIType.ITEM_ATTACHMENTS,
        UnindexedFieldURIType.CONTACTS_SPOUSE_NAME,
        UnindexedFieldURIType.ITEM_SUBJECT,
        UnindexedFieldURIType.CONTACTS_SURNAME,
        UnindexedFieldURIType.ITEM_WEB_CLIENT_EDIT_FORM_QUERY_STRING,
        UnindexedFieldURIType.ITEM_WEB_CLIENT_READ_FORM_QUERY_STRING
    ];

    /**
     * Public functions
     */

    /**
     * This function will fetech a group of contacts using the Keep API and return an
     * array of corresponding EWS items populated with the fields requested in the shape.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param request The original SOAP request for the get.
     * @param shape EWS shape describing the information requested for each item.
     * @param startIndex Optional start index for items when paging.  Defaults to 0.
     * @param count Optional count of objects to request.  Defaults to 512.
     * @param fromLabel The unid of the label (folder) we are querying.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @returns An array of EWS ContactItemType objects built from the returned PimItems.
     */
    async getItems(
        userInfo: UserInfo, 
        request: Request, 
        shape: ItemResponseShapeType, 
        startIndex?: number, 
        count?: number, 
        fromLabel?: PimLabel,
        mailboxId?: string
    ): Promise<ContactItemType[]> {

        const contactItems: ContactItemType[] = [];
        let pimItems: PimContact[] | undefined = undefined;
        try {
            const needDocuments = (shape.BaseShape === DefaultShapeNamesType.ID_ONLY &&
                (shape.AdditionalProperties === undefined || shape.AdditionalProperties.items.length === 0)) ? false : true;
            pimItems = await KeepPimContactsManager
                .getInstance()
                .getContacts(userInfo, needDocuments, startIndex, count, fromLabel?.unid, mailboxId);
        } catch (err) {
            Logger.getInstance().debug("Error retrieving pim contact entries: " + err);
            // If we throw the err here the client will continue in a loop with SyncFolderHierarchy and SyncFolderItems asking for contacts.
        }
        if (pimItems) {
            for (const pimContact of pimItems) {
                if (pimContact.isGroup) {
                    Logger.getInstance().warn('Group contacts are currently not supported');
                    continue;
                }

                if (pimContact.isPimContact()) {
                    this.updateContactEmailProperties(pimContact);
                }

                const contact = await this.pimItemToEWSItem(pimContact, userInfo, request, shape, mailboxId, fromLabel?.unid);
                if (contact) {
                    contactItems.push(contact);
                } else {
                    Logger.getInstance().error(`An EWS contact type could not be created for PIM contact ${pimContact.unid}`);
                }
            }
        }
        return contactItems;
    }

    /**
     * Process updates for a contact.
     * @param pimItem Existing PimItem from Keep being updated.
     * @param change The EWS change type object that describes the item and the changes to be made.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param request The original SOAP request for the get.
     * @param toLabel Optional label of the item's parent folder.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @returns An updated contact or undefined if the request could not be completed.
     */
    async updateItem(
        pimContact: PimContact, 
        change: ItemChangeType, 
        userInfo: UserInfo, 
        request: Request, 
        toLabel?: PimLabel,
        mailboxId?: string
    ): Promise<ContactItemType | undefined> {

        let originalEmails: string[] = []; // For contacts, the list of original 3 emails set (in order EmailAddress1, EmailAddress2, EmailAddress3)
        originalEmails = this.orderedEmails(pimContact);

        for (const fieldUpdate of change.Updates.items) {
            let newItem: ItemType | undefined = undefined;
            if (fieldUpdate instanceof SetItemFieldType || fieldUpdate instanceof AppendToItemFieldType) {
                newItem = fieldUpdate.Item ?? fieldUpdate.Contact;
            }

            if (fieldUpdate instanceof SetItemFieldType) {
                if (newItem === undefined) {
                    Logger.getInstance().error(`No new item set for update field: ${util.inspect(newItem, false, 5)}`);
                    continue;
                }

                if (fieldUpdate.FieldURI) {
                    const field = fieldUpdate.FieldURI.FieldURI.split(':')[1];
                    this.updatePimItemFieldValue(pimContact, fieldUpdate.FieldURI.FieldURI, (newItem as any)[field]);
                } else if (fieldUpdate.ExtendedFieldURI && newItem.ExtendedProperty) {
                    const identifiers = identifiersForPathToExtendedFieldType(fieldUpdate.ExtendedFieldURI);
                    const newValue = getValueForExtendedFieldURI(newItem, identifiers);

                    const current = pimContact.findExtendedProperty(identifiers);
                    if (current) {
                        current.Value = newValue;
                        pimContact.updateExtendedProperty(identifiers, current);
                    }
                    else {
                        addExtendedPropertyToPIM(pimContact, newItem.ExtendedProperty[0]);
                    }
                } else if (fieldUpdate.IndexedFieldURI) {
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
                    switch (fieldUpdate.IndexedFieldURI.FieldURI) {
                        case DictionaryURIType.CONTACTS_IM_ADDRESS:
                            // Update IM addresses from newItem
                            this.updateIMAddresses(newItem, pimContact);
                            break;
                        case DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_STREET:
                        case DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_CITY:
                        case DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_STATE:
                        case DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_COUNTRY_OR_REGION:
                        case DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_POSTAL_CODE:
                            // Update physical addresses from newItem
                            this.updatePhysicalAddresses(newItem, pimContact);
                            break;
                        case DictionaryURIType.CONTACTS_PHONE_NUMBER:
                            // Update phone numbers from newItem
                            this.updatePhoneNumbers(newItem, pimContact);
                            break;
                        case DictionaryURIType.CONTACTS_EMAIL_ADDRESS:
                            // Update email addres from newItem
                            this.updateEmailAddresses(newItem, pimContact);
                            break;
                        case DictionaryURIType.DISTRIBUTIONLIST_MEMBERS_MEMBER:
                        case DictionaryURIType.ITEM_INTERNET_MESSAGE_HEADER:
                        default:
                            Logger.getInstance().warn(`Unhandled SetItemField request for notes for field:  ${fieldUpdate.IndexedFieldURI}`);
                    }



                }
            }
            else if (fieldUpdate instanceof AppendToItemFieldType) {
                /*
                    The following properties are supported for the append action for contacts:
                    - item:Body
                */
                if (fieldUpdate.FieldURI?.FieldURI === UnindexedFieldURIType.ITEM_BODY && newItem) {
                    pimContact.body = `${pimContact.body}${newItem?.Body ?? ""}`
                } else {
                    Logger.getInstance().warn(`Unhandled AppendToItemField request for contact for field:  ${fieldUpdate.FieldURI?.FieldURI}`);
                }
            }
            else if (fieldUpdate instanceof DeleteItemFieldType) {
                if (fieldUpdate.FieldURI) {
                    this.updatePimItemFieldValue(pimContact, fieldUpdate.FieldURI.FieldURI, undefined);
                }
                else if (fieldUpdate.ExtendedFieldURI) {
                    pimContact.deleteExtendedProperty(identifiersForPathToExtendedFieldType(fieldUpdate.ExtendedFieldURI));
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
                    const type = fieldUpdate.IndexedFieldURI.FieldURI.split(':')[0];
                    if (type === "contacts" && pimContact.isPimContact()) {
                        if (fieldUpdate.IndexedFieldURI.FieldURI === DictionaryURIType.CONTACTS_EMAIL_ADDRESS) {
                            // fieldUpdate.IndexedFieldURI.FieldIndex will be EmailAddressKeyType (EmailAddress1, EmailAddress2, or EmailAddress3).
                            if (originalEmails.length > 0) {
                                let index = -1;
                                if (fieldUpdate.IndexedFieldURI.FieldIndex === EmailAddressKeyType.EMAIL_ADDRESS_1) {
                                    index = 0;
                                }
                                else if (fieldUpdate.IndexedFieldURI.FieldIndex === EmailAddressKeyType.EMAIL_ADDRESS_2) {
                                    index = 1;
                                }
                                else if (fieldUpdate.IndexedFieldURI.FieldIndex === EmailAddressKeyType.EMAIL_ADDRESS_3) {
                                    index = 2;
                                }

                                if (index !== -1 && originalEmails.length > index) {
                                    pimContact.removeEmailAddress(originalEmails[index]);
                                }
                            }
                        }
                        else if (fieldUpdate.IndexedFieldURI.FieldURI === DictionaryURIType.CONTACTS_PHONE_NUMBER) {
                            let number: string | undefined = undefined;
                            if (fieldUpdate.IndexedFieldURI.FieldIndex === PhoneNumberKeyType.HOME_PHONE) {
                                number = pimContact.homePhone;
                            }
                            else if (fieldUpdate.IndexedFieldURI.FieldIndex === PhoneNumberKeyType.HOME_FAX) {
                                number = pimContact.homeFax;
                            }
                            else if (fieldUpdate.IndexedFieldURI.FieldIndex === PhoneNumberKeyType.MOBILE_PHONE) {
                                number = pimContact.cellPhone;
                            }
                            else if (fieldUpdate.IndexedFieldURI.FieldIndex === PhoneNumberKeyType.BUSINESS_PHONE) {
                                number = pimContact.officePhone;
                            }
                            else if (fieldUpdate.IndexedFieldURI.FieldIndex === PhoneNumberKeyType.BUSINESS_FAX) {
                                number = pimContact.officeFax;
                            }

                            if (number) {
                                pimContact.removePhoneNumber(number);
                            }
                        }
                        else if (fieldUpdate.IndexedFieldURI.FieldURI === DictionaryURIType.CONTACTS_IM_ADDRESS) {
                            const currentIMs = pimContact.imAddresses;
                            let address: string | undefined = undefined;
                            if (fieldUpdate.IndexedFieldURI.FieldIndex === ImAddressKeyType.IM_ADDRESS_1 && currentIMs.length > 0) {
                                address = currentIMs[0];
                            }
                            else if (fieldUpdate.IndexedFieldURI.FieldIndex === ImAddressKeyType.IM_ADDRESS_2 && currentIMs.length > 1) {
                                address = currentIMs[1];
                            }
                            else if (fieldUpdate.IndexedFieldURI.FieldIndex === ImAddressKeyType.IM_ADDRESS_3 && currentIMs.length > 2) {
                                address = currentIMs[2];
                            }
                            if (address) {
                                pimContact.removeImAddress(address);
                            }
                        }
                        else if (fieldUpdate.IndexedFieldURI.FieldURI === DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_STREET) {
                            pimContact.removeAddressKey(ContactAddressKey.Street, fieldUpdate.IndexedFieldURI.FieldIndex);
                        }
                        else if (fieldUpdate.IndexedFieldURI.FieldURI === DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_CITY) {
                            pimContact.removeAddressKey(ContactAddressKey.City, fieldUpdate.IndexedFieldURI.FieldIndex);
                        }
                        else if (fieldUpdate.IndexedFieldURI.FieldURI === DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_STATE) {
                            pimContact.removeAddressKey(ContactAddressKey.State, fieldUpdate.IndexedFieldURI.FieldIndex);
                        }
                        else if (fieldUpdate.IndexedFieldURI.FieldURI === DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_POSTAL_CODE) {
                            pimContact.removeAddressKey(ContactAddressKey.PostalCode, fieldUpdate.IndexedFieldURI.FieldIndex);
                        }
                        else if (fieldUpdate.IndexedFieldURI.FieldURI === DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_COUNTRY_OR_REGION) {
                            pimContact.removeAddressKey(ContactAddressKey.CountryOrRegion, fieldUpdate.IndexedFieldURI.FieldIndex);
                        }
                        else {
                            Logger.getInstance().warn(`Did not process delete for indexed field ${fieldUpdate.IndexedFieldURI.FieldURI}`);
                        }
                    }
                    else {
                        Logger.getInstance().error(`Unexpected PIM item type ${pimContact.constructor.name} for remove index field ${fieldUpdate.IndexedFieldURI.FieldURI}`)
                    }
                }
            }
        }

        // Set the target folder if passed in
        if (toLabel) {
            pimContact.parentFolderIds = [toLabel.folderId];  // May be possible we've lost other parent ids set by another client.  
            // Even if stored as an extra property for the client, it could have been updated on the server since stored on the client.
            // To shrink the window, we'd need to request for the item from the server and update it right away
        }

        // The PimContact should now be updated with the new information.  Send it to Keep.
        await KeepPimContactsManager.getInstance().updateContact(pimContact, userInfo, mailboxId);

        if (pimContact.parentFolderIds === undefined) {
            try {
                const labels = await KeepPimLabelManager.getInstance().getLabels(userInfo, undefined, undefined, mailboxId) ;
                const contacts = labels.find(label => { return label.view === KeepPimConstants.CONTACTS });
                pimContact.parentFolderIds = contacts ? [contacts.folderId] : undefined;
            }
            catch (err) {
                Logger.getInstance().error(`Unable to get labels to determine contact parent folder: ${err}`);
            }
        }
        const shape = new ItemResponseShapeType();
        shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
        return this.pimItemToEWSItem(pimContact, userInfo, request, shape, mailboxId);
    }

    /**
     * This function issues a Keep API call to create a contact based on the EWS item passed in.
     * @param item The EWS contact to create
     * @param userInfo The user's credentials to be passed to Keep.
     * @param request The original SOAP request for the create.
     * @param toLabel Optional target label (folder).
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @returns A new EWS ContactItemType object populated with information about the created entry.
     */
    async createItem(
        item: ContactItemType, 
        userInfo: UserInfo, 
        request: Request, 
        disposition?: MessageDispositionType, 
        toLabel?: PimLabel,
        mailboxId?: string
    ): Promise<ContactItemType[]> {
        const pimContact = this.pimItemFromEWSItem(item, request);
        const unid = await KeepPimContactsManager.getInstance().createContact(pimContact, userInfo, mailboxId);
        const targetItem = new ContactItemType();
        const itemEWSId = getEWSId(unid, mailboxId);
        targetItem.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);

        // TODO:  Uncomment this line when we can set the parent folder to the base contacts folder
        //        See LABS-866
        // If no toLabel is passed in, we will create the contact in the top level contacts folder
        // if (toLabel === undefined) {
        //     const labels = await KeepPimLabelManager.getInstance().getLabels(userInfo);
        //     toLabel = labels.find(label => { return label.view === KeepPimConstants.CONTACTS });
        // }

        // The Keep API does not support creating a contact IN a folder, so we must now move the 
        // contact to the desired folder
        // TODO:  If this contact is being created in the top level contacts folder, do not attempt to 
        //        move the contact to the top level folder as this will fail.
        //        See LABS-866
        if (toLabel !== undefined && toLabel.view !== KeepPimConstants.CONTACTS) {
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
    protected async createNewPimItem(pimItem: PimContact, userInfo: UserInfo, mailboxId?: string): Promise<PimContact> {
        const unid = await KeepPimContactsManager.getInstance().createContact(pimItem, userInfo, mailboxId);
        const newItem = await KeepPimContactsManager.getInstance().getContact(unid, userInfo, mailboxId);
        if (newItem === undefined) {
            // Try to delete item since it may have been created
            try {
                await KeepPimContactsManager.getInstance().deleteContact(unid, userInfo, undefined, mailboxId);
            }
            catch {
                // Ignore errors
            }

            throw new Error(`Unable to retrieve contact ${unid} after create`);
        }

        return newItem;
    }

    /**
      * This function will make a Keep API request to delete a contact.
      * @param contact The contact or unid of the contact to delete.
      * @param userInfo The user's credentials to be passed to Keep.
      * @param mailboxId SMTP mailbox delegator or delegatee address.
      * @param hardDelete Indicator if the contact should be deleted from trash as well
      */
    async deleteItem(item: string | PimContact, userInfo: UserInfo, mailboxId?: string, hardDelete = false): Promise<void> {
        const unid = (typeof item === 'string') ? item : item.unid;
        await KeepPimContactsManager.getInstance().deleteContact(unid, userInfo, hardDelete, mailboxId);
    }


    /**
     * Getter for the list of fields included in the default shape for notes
     * @returns Array of fields for the default shape
     */
    fieldsForDefaultShape(): UnindexedFieldURIType[] {
        return EWSContactsManager.defaultFields;
    }

    /**
     * Getter for the list of fields included in the all properties shape for notes
     * @returns Array of fields for the all properties shape
     */
    fieldsForAllPropertiesShape(): UnindexedFieldURIType[] {
        return EWSContactsManager.allPropertiesFields;
    }

    /**
     * Internal functions
     */

    /**
     * This is a helper function to convert an EWS ContactItemType object to a PimContact used to communicate with the Keep API.
     * @param item The source EWS item to convert to a PimContact.
     * @param request The original SOAP request for the get.
     * @param existing Object containing existing fields to apply.
     * @returns A PimContact populated with the EWS contact's data.
     */
    pimItemFromEWSItem(item: ContactItemType, request: Request, existing?: object): PimContact {
        const contactObject: any = existing ? existing : { "Type": "Person" };

        if (item.ItemId) {
            const [itemId] = getKeepIdPair(item.ItemId.Id);
            contactObject.uid = itemId;
        }

        if (item.ParentFolderId) {
            //Get parent folder id from EWS parent folder Id
            const [parentFolderId] = getKeepIdPair(item.ParentFolderId.Id);
            contactObject.ParentFolder = parentFolderId;
        }

        if (item.Categories) {
            contactObject.Categories = convertToListObject(item.Categories.String);
        }

        if (item.CompleteName) {
            contactObject.FullName = item.CompleteName.FullName;
            contactObject.FirstName = item.CompleteName.FirstName;
            contactObject.MiddleInitial = item.CompleteName.MiddleName;
            contactObject.LastName = item.CompleteName.LastName;
            contactObject.Title = item.CompleteName.Title;
            contactObject.Suffix = item.CompleteName.Suffix;
        }

        if (item.GivenName) {
            contactObject.FirstName = item.GivenName;
        }
        if (item.MiddleName) {
            contactObject.MiddleInitial = item.MiddleName;
        }
        if (item.Surname) {
            contactObject.LastName = item.Surname;
        }

        if (item.Body) {
            contactObject.Comment = item.Body.Value;
        }
        if (item.JobTitle) {
            contactObject.JobTitle = item.JobTitle;
        }
        if (item.Manager) {
            contactObject.Manager = item.Manager;
        }
        if (item.CompanyName) {
            contactObject.CompanyName = item.CompanyName;
        }
        else if (item.Companies && item.Companies.String.length > 0) {
            // Domino only shows one company
            contactObject.CompanyName = item.Companies.String[0];
        }
        if (item.AssistantName) {
            contactObject.Assistant = item.AssistantName;
        }
        if (item.Birthday) {
            contactObject.Birthday = item.Birthday.toISOString();
        }
        if (item.BusinessHomePage) {
            contactObject.WebSite = item.BusinessHomePage;
        }
        if (item.Department) {
            contactObject.Department = item.Department;
        }

        if (item.OfficeLocation) {
            contactObject.Location = item.OfficeLocation;
        }
        if (item.SpouseName) {
            contactObject.Spouse = item.SpouseName;
        }


        // TODO: How to set contact photo. Photo shows as attachment:
        // "$FILES": [
        //     "ContactPhoto"
        //   ],
        // Contact item has these settings: 
        //    <HasPicture/>
        //    <Photo/>

        const rtn = PimItemFactory.newPimContact(contactObject, PimItemFormat.DOCUMENT);

        if (item.WeddingAnniversary) {
            rtn.anniversary = item.WeddingAnniversary;
        }

        if (item.DateTimeCreated) {
            rtn.createdDate = item.DateTimeCreated;
        }

        if (item.LastModifiedTime) {
            rtn.lastModifiedDate = item.LastModifiedTime;
        }

        if (item.IsPrivate === undefined || item.IsPrivate === false) {
            rtn.isPrivate = false;
        } else {
            rtn.isPrivate = true;
        }

        // Handle the contact's physical addresses
        this.updatePhysicalAddresses(item, rtn);

        // Handle the contact's phone numbers
        this.updatePhoneNumbers(item, rtn);

        // Handle the contact's email addresses
        this.updateEmailAddresses(item, rtn);

        // Handle the contact's IM addresses
        this.updateIMAddresses(item, rtn);

        // Copy any extended fields to the PimMessage to preserve them.
        addExtendedPropertiesToPIM(item, rtn);

        return rtn;
    }

    /**
     * Create an EWS ContactItem object based on the provided PIM item.
     * @param pimContact The source PimContact from which to build an EWS item.
     * @param shape The EWS shape describing the fields that should be populated on the returned object.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @param targetParentFolderId An optional parent folder Id to apply to the returned item.
     * @returns An EWS ContactItemType object based on the passed PimContact and shape in shape.
     */
    async pimItemToEWSItem(
        pimContact: PimContact, 
        userInfo: UserInfo, 
        request: Request, 
        shape: ItemResponseShapeType, 
        mailboxId?: string,
        targetParentFolderId?: string
    ): Promise<ContactItemType> {
        // Notes are treated as MessageTypes by EWS
        const contact = new ContactItemType();

        // Add all requested properties based on the shape
        this.addRequestedPropertiesToEWSItem(pimContact, contact, shape, mailboxId);

        if (targetParentFolderId) {
            const parentFolderEWSId = getEWSId(targetParentFolderId, mailboxId);
            contact.ParentFolderId = new FolderIdType(parentFolderEWSId, `ck-${parentFolderEWSId}`);
        }
        return contact;
    }

    /**
     * Update an field in an EWS contact for a indexed field. 
     * @param contact The EWS contatct item to update
     * @param pimContact The PIM contact item containing the data
     * @param indexedField The indexed field describing the field to update
     * @returns True if the field was processed and updated in the contact. False if the field was not processed. 
     */
    protected updateEWSIndexedItemFieldValue(contact: ContactItemType, pimContact: PimContact, indexedField: PathToIndexedFieldType): boolean {
        let handled = true; 

        switch (indexedField.FieldURI) {
            case DictionaryURIType.CONTACTS_EMAIL_ADDRESS:
                {
                    const emails = this.orderedEmails(pimContact);
                    const fields = [EmailAddressKeyType.EMAIL_ADDRESS_1, EmailAddressKeyType.EMAIL_ADDRESS_2, EmailAddressKeyType.EMAIL_ADDRESS_3];
                    const index = fields.findIndex(field => field === indexedField.FieldIndex);

                    let email: string | undefined; 
                    if (index !== -1) {
                        if (index < emails.length) {
                            email = emails[index];
                        }

                        if (email) {
                            if (contact.EmailAddresses === undefined) {
                                const entry = new EmailAddressDictionaryEntryType(email, fields[index], email, 'SMTP', MailboxTypeType.CONTACT);
                                contact.EmailAddresses = new EmailAddressDictionaryType();
                                contact.EmailAddresses.Entry = [entry];
                            }
                            else if (contact.EmailAddresses) {
                                const existingIndex = contact.EmailAddresses.Entry.findIndex(entry => entry.Key === indexedField.FieldIndex);
                                if (existingIndex !== -1) {
                                    const existingEntry = contact.EmailAddresses.Entry[existingIndex];
                                    existingEntry.Value = email; 
                                    existingEntry.Name = email; 
                                    existingEntry.RoutingType = 'SMTP';
                                    existingEntry.MailboxType = MailboxTypeType.CONTACT;
                                }
                                else {
                                    const entry = new EmailAddressDictionaryEntryType(email, fields[index], email, 'SMTP', MailboxTypeType.CONTACT);
                                    contact.EmailAddresses.Entry.push(entry);
                                }
                            }
                        }
                    }
                    else {
                        Logger.getInstance().error(`Unknown Contact email field: ${util.inspect(indexedField, false, 5)}`);
                    }
                }
                break;
        
            case DictionaryURIType.CONTACTS_IM_ADDRESS: 
                {
                    const imAddresses: any = pimContact.getAdditionalProperty(UnindexedFieldURIType.CONTACTS_IM_ADDRESSES);
                    if (imAddresses) {
                        const fields = [ImAddressKeyType.IM_ADDRESS_1, ImAddressKeyType.IM_ADDRESS_2, ImAddressKeyType.IM_ADDRESS_3];
                        const index = fields.findIndex(field => field === indexedField.FieldIndex);

                        let im: string | undefined; 
                        if (index !== -1) {
                            if (index < Object.keys(imAddresses).length) {
                                im = imAddresses[indexedField.FieldIndex];
                            }

                            if (im) {
                                if (contact.ImAddresses === undefined) {
                                    const entry = new ImAddressDictionaryEntryType(im, fields[index]);
                                    contact.ImAddresses = new ImAddressDictionaryType();
                                    contact.ImAddresses.Entry = [entry];
                                }
                                else if (contact.ImAddresses) {
                                    const existingIndex = contact.ImAddresses.Entry.findIndex(entry => entry.Key === indexedField.FieldIndex);
                                    if (existingIndex !== -1) {
                                        const existingEntry = contact.ImAddresses.Entry[existingIndex];
                                        existingEntry.Value = im; 
                                    }
                                    else {
                                        const entry = new ImAddressDictionaryEntryType(im, fields[index]);
                                        contact.ImAddresses.Entry.push(entry);
                                    }
                                }
                            }
                        }
                        else {
                            Logger.getInstance().error(`Unknown Contact IM Address field: ${util.inspect(indexedField, false, 5)}`);
                        }
                    }
                }
                break; 

            case DictionaryURIType.CONTACTS_PHONE_NUMBER: 
                {
                    const phoneDictionary: PhoneNumberDictionaryEntryType[] = [];
                    switch (indexedField.FieldIndex) {
                        case FieldIndexValue.CONTACT_PHONE_NUMBER_BUSINESS:
                            if (pimContact.officePhone) {
                                phoneDictionary.push(new PhoneNumberDictionaryEntryType(pimContact.officePhone, PhoneNumberKeyType.BUSINESS_PHONE));
                            }
                            break;
                    
                        // TODO: Uncomment when LABS-1155 is implemented
                        // case FieldIndexValue.CONTACT_PHONE_NUMBER_PRIMARY:
                        //     if (pimContact.primaryPhone) {
                        //         phoneNumber.push(new PhoneNumberDictionaryEntryType(pimContact.primaryPhone, PhoneNumberKeyType.PRIMARY_PHONE));
                        //     }
                        //     break;
                        
                        case FieldIndexValue.CONTACT_PHONE_NUMBER_BUSINESS_FAX:
                            if (pimContact.officeFax) {
                                phoneDictionary.push(new PhoneNumberDictionaryEntryType(pimContact.officeFax, PhoneNumberKeyType.BUSINESS_FAX));
                            }
                            break;
                        case FieldIndexValue.CONTACT_PHONE_NUMBER_HOME:
                            if (pimContact.homePhone) {
                                phoneDictionary.push(new PhoneNumberDictionaryEntryType(pimContact.homePhone, PhoneNumberKeyType.HOME_PHONE));
                            }
                            break;
                        case FieldIndexValue.CONTACT_PHONE_NUMBER_HOME_FAX:
                            if (pimContact.homeFax) {
                                phoneDictionary.push(new PhoneNumberDictionaryEntryType(pimContact.homeFax, PhoneNumberKeyType.HOME_FAX));
                            }
                            break;
                        case FieldIndexValue.CONTACT_PHONE_NUMBER_MOBILE:
                            if (pimContact.cellPhone) {
                                phoneDictionary.push(new PhoneNumberDictionaryEntryType(pimContact.cellPhone, PhoneNumberKeyType.MOBILE_PHONE));
                            }
                            break;
                        case FieldIndexValue.CONTACT_PHONE_NUMBER_OTHER:
                            pimContact.otherPhones.forEach(item => {
                                const entry = new PhoneNumberDictionaryEntryType(item, PhoneNumberKeyType.OTHER_TELEPHONE);
                                phoneDictionary.push(entry);
                            });
                            break;
                        default:
                            Logger.getInstance().warn(`Unsupported field index ${util.inspect(indexedField, false, 5)}`);
                            break;
                    }

                    if (phoneDictionary.length > 0) {
                        if (contact.PhoneNumbers === undefined) {
                            contact.PhoneNumbers = new PhoneNumberDictionaryType();
                            contact.PhoneNumbers.Entry = phoneDictionary;
                        }
                        else {
                            // Merge with existing phone numbers
                            for (const phone of phoneDictionary) {
                                const exist = contact.PhoneNumbers.Entry.findIndex(number => number.Value === phone.Value && number.Key === phone.Key);
                                if (exist === -1) {
                                    contact.PhoneNumbers.Entry.push(phone); 
                                }
                            }
                        }
                    }
                }
                break; 

            case DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_CITY:
                {
                    const [address, keyType] = this.getAddressAndKeyType(indexedField, pimContact);
                    if (address && keyType) {
                        const entry = this.getEWSAddressEntry(contact, keyType);
                        entry.City = address.City; 
                    }
                }
                break; 

            case DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_COUNTRY_OR_REGION:
                {
                    const [address, keyType] = this.getAddressAndKeyType(indexedField, pimContact);
                    if (address && keyType) {
                        const entry = this.getEWSAddressEntry(contact, keyType);
                        entry.CountryOrRegion = address.Country; 
                       
                    }
                }
                break; 

            case DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_POSTAL_CODE:
                {
                    const [address, keyType] = this.getAddressAndKeyType(indexedField, pimContact);
                    if (address && keyType) {
                        const entry = this.getEWSAddressEntry(contact, keyType);
                        entry.PostalCode = address.PostalCode; 
                    }
                }
                break;

            case DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_STATE:
                {
                    const [address, keyType] = this.getAddressAndKeyType(indexedField, pimContact);
                    if (address && keyType) {
                        const entry = this.getEWSAddressEntry(contact, keyType);
                        entry.State = address.State; 
                    }
                }
                break; 

            case DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_STREET:
                {
                    const [address, keyType] = this.getAddressAndKeyType(indexedField, pimContact);
                    if (address && keyType) {
                        const entry = this.getEWSAddressEntry(contact, keyType);
                        entry.Street = address.Street; 
                    }
                }
                break; 

            default:
                handled = false; 
                break;
        }

        if (handled === false) {
            // Common fields handled in superclass
            handled = super.updateEWSIndexedItemFieldValue(contact, pimContact, indexedField);
        }

        if (handled === false) {
            Logger.getInstance().warn(`Unrecognized contact indexed field: ${util.inspect(indexedField, false, 5)}`);
        }

        return handled;
    }

    protected getExtendedField(path: PathToExtendedFieldType, pimContact: PimContact): any | undefined {

        const identifiers = identifiersForPathToExtendedFieldType(path); 
        let rtn = pimContact.findExtendedProperty(identifiers);
        if (rtn === undefined) {
            let value: string | undefined;  
            if (identifiers[ExtendedPropertyKeyType.DISTINGUISHED_PROPERTY_SETID] === DistinguishedPropertySetType.ADDRESS) {
                              
                const emails = this.orderedEmails(pimContact);
                const propertyId = identifiers[ExtendedPropertyKeyType.PROPERTY_ID];
                if (propertyId === MapiPropertyIds.EMAIL1_ORIG_DISPLAY_NAME && emails.length > 0) {
                    value = emails[0];
                }
                else if (propertyId === MapiPropertyIds.EMAIL2_ORIG_DISPLAY_NAME && emails.length > 1) {
                    value = emails[1];
                }
                else if (propertyId === MapiPropertyIds.EMAIL3_ORIG_DISPLAY_NAME && emails.length > 2) {
                    value = emails[2];
                }
            }
            else if (identifiers[ExtendedPropertyKeyType.DISTINGUISHED_PROPERTY_SETID] === DistinguishedPropertySetType.PUBLIC_STRINGS) {
                const emails = this.orderedEmails(pimContact);
                if (identifiers[ExtendedPropertyKeyType.PROPERTY_NAME] === "http://schemas.microsoft.com/entourage/emaillabel1" && emails.length > 0) {
                    value = getLabelTypeForEmail(emails[0], pimContact);
                }
                else if (identifiers[ExtendedPropertyKeyType.PROPERTY_NAME] === "http://schemas.microsoft.com/entourage/emaillabel2" && emails.length > 1) {
                    value = getLabelTypeForEmail(emails[1], pimContact);
                }
                else if (identifiers[ExtendedPropertyKeyType.PROPERTY_NAME] === "http://schemas.microsoft.com/entourage/emaillabel3" && emails.length > 2) {
                    value = getLabelTypeForEmail(emails[2], pimContact);
                }
            }
            else if (identifiers[ExtendedPropertyKeyType.PROPERTY_SETID] === "745b56bb-b118-4a7e-8020-264e35e87ceb") {
                if (identifiers[ExtendedPropertyKeyType.PROPERTY_NAME] === "DefaultEmailAddress") {
                    value = pimContact.primaryEmail; 
                }
            }

            if (value !== undefined) {
                rtn = Object.assign({'Value': value}, identifiers);
                rtn[ExtendedPropertyKeyType.PROPERTY_TYPE] = path.PropertyType; // identifiers don't contain type
            } 
        }

        return rtn;
    }

    /**
     * Get a EWS address entry for a address key type. 
     * @param contact The EWS contact contact containing the existing addresses.
     * @param keyType The physical address type to return.
     * @returns A physical address entry for the key type. If the contact does not contain a physical address entry for the key type a new one will be created in the contact. 
     */
    protected getEWSAddressEntry(contact: ContactItemType, keyType: PhysicalAddressKeyType): PhysicalAddressDictionaryEntryType  {
        let entry: PhysicalAddressDictionaryEntryType | undefined; 
        if (contact.PhysicalAddresses === undefined) {
            contact.PhysicalAddresses = new PhysicalAddressDictionaryType();
        }
        else {
            entry = contact.PhysicalAddresses.Entry.find(pAdd => pAdd.Key === keyType);
        }

        if (entry === undefined) {
            entry = new PhysicalAddressDictionaryEntryType(); 
            entry.Key = keyType; 
            contact.PhysicalAddresses.Entry.push(entry);
        }

        return entry; 
    }

    /**
     * Get the address object from a PIM contact that is identified by an indexed field. 
     * @param indexedField The indexed field identifing the address to return.
     * @param pimContact The PIM contact containing the addresses
     * @returns The PIM address object and the physical address key for the address.  
     */
    protected getAddressAndKeyType(indexedField: PathToIndexedFieldType, pimContact: PimContact): [PimAddress | undefined, PhysicalAddressKeyType | undefined] {
        let address: PimAddress | undefined; 
        let keyType: PhysicalAddressKeyType | undefined; 
        switch (indexedField.FieldIndex) {
            case FieldIndexValue.CONTACT_PHYSICAL_ADDRESS_HOME:
                address = pimContact.homeAddress;
                keyType = PhysicalAddressKeyType.HOME;
                break;
        
            case FieldIndexValue.CONTACT_PHYSICAL_ADDRESS_BUSINESS:
                address = pimContact.officeAddress; 
                keyType = PhysicalAddressKeyType.BUSINESS;
                break;
            
            case FieldIndexValue.CONTACT_PHYSICAL_ADDRESS_OTHER:
                address = pimContact.otherAddress; 
                keyType = PhysicalAddressKeyType.OTHER;
                break;
            default:
                Logger.getInstance().error(`Unknown field index ${util.inspect(indexedField, false, 5)}`);
                break;
        }

        return [address, keyType];
    }

    /**
    * Update the passed in EWS contact's specific field with the information stored in the pimItem.  This function is
    * used to populate fields requested in a GetItem/FindItem/SyncX EWS operation.  It will process common fields
    * and return true if the request was handled.
    * @param item The EWS item to update with a value.
    * @param pimContact The PIM item from which to read the value.
    * @param fieldIdentifier The EWS field we are mapping.
    * @param mailboxId SMTP mailbox delegator or delegatee address
    */
    protected updateEWSItemFieldValue(
        contact: ContactItemType, 
        pimContact: PimContact, 
        fieldIdentifier: UnindexedFieldURIType,
        mailboxId?: string
    ): boolean {
            
        let handled = true;

            switch (fieldIdentifier) {
                case UnindexedFieldURIType.CONTACTS_COMPANY_NAME:
                    if (pimContact.companyName) {
                        contact.CompanyName = pimContact.companyName;
                    }
                    break;
                case UnindexedFieldURIType.CONTACTS_COMPLETE_NAME:
                    {
                        const fullName = new CompleteNameType();
                        fullName.FirstName = pimContact.firstName;
                        fullName.LastName = pimContact.lastName;
                        fullName.MiddleName = pimContact.middleInitial;
                        fullName.Title = pimContact.title;
                        fullName.Suffix = pimContact.suffix;
                        contact.CompleteName = fullName;
                    }
                    break;
                case UnindexedFieldURIType.CONTACTS_EMAIL_ADDRESSES:
                    {
                        const emails = this.orderedEmails(pimContact);

                        if (emails.length > 0) {
                            const emailDictionary = new EmailAddressDictionaryType();
                            emailDictionary.Entry = [];
                            const keyTypes = [EmailAddressKeyType.EMAIL_ADDRESS_1, EmailAddressKeyType.EMAIL_ADDRESS_2, EmailAddressKeyType.EMAIL_ADDRESS_3];
                            for (let index = 0; index < emails.length && index < keyTypes.length; index++) {
                                const entry = new EmailAddressDictionaryEntryType(emails[index], keyTypes[index], emails[index], 'SMTP', MailboxTypeType.CONTACT);
                                emailDictionary.Entry.push(entry);
                            }
                            contact.EmailAddresses = emailDictionary;
                        }
                    }
                    break;
                case UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS:
                    // This is true if there are attachments ('ContactPhoto' does not count)
                    contact.HasAttachments = false;
                    if (pimContact.attachments && pimContact.attachments.length > 0) {
                        if (pimContact.attachments.includes('ContactPhoto')) {
                            contact.HasAttachments = pimContact.attachments.length > 1;
                        } else {
                            contact.HasAttachments = true;
                        }
                    }
                    break;
                case UnindexedFieldURIType.ITEM_ATTACHMENTS:
                    {
                        const pAttachments = getItemAttachments(pimContact);
                        if (pAttachments) {
                            contact.Attachments = pAttachments;
                        }
                    }
                    break;
                case UnindexedFieldURIType.CONTACTS_IM_ADDRESSES:
                    {
                        const keys = [ImAddressKeyType.IM_ADDRESS_1, ImAddressKeyType.IM_ADDRESS_2, ImAddressKeyType.IM_ADDRESS_3];
                        let imAddresses: any = pimContact.getAdditionalProperty(UnindexedFieldURIType.CONTACTS_IM_ADDRESSES);
                        if (imAddresses === undefined && pimContact.imAddresses.length > 0) {
                            imAddresses = {};
                            pimContact.imAddresses.slice(0,3).forEach((address, index) => imAddresses[keys[index]] = address);
                        }
                        
                        if (imAddresses !== undefined){
                            const ImDictionary = new ImAddressDictionaryType();
                            for (const key in imAddresses){
                                ImDictionary.Entry.push(new ImAddressDictionaryEntryType(imAddresses[key], key as ImAddressKeyType))
                            }
                            contact.ImAddresses = ImDictionary;
                        }
                    }
                    break;
                case UnindexedFieldURIType.CONTACTS_JOB_TITLE:
                    if (pimContact.jobTitle) {
                        contact.JobTitle = pimContact.jobTitle;
                    }
                    break;
                case UnindexedFieldURIType.CONTACTS_PHONE_NUMBERS:
                    {
                        const phoneDictionary = new PhoneNumberDictionaryType();
                        phoneDictionary.Entry = [];

                        // if (pimContact.primaryPhone) {
                        //     const entry = new PhoneNumberDictionaryEntryType(pimContact.primaryPhone, PhoneNumberKeyType.PRIMARY_PHONE);
                        //     phoneDictionary.Entry.push(entry);
                        // }
                        if (pimContact.officePhone) {
                            const entry = new PhoneNumberDictionaryEntryType(pimContact.officePhone, PhoneNumberKeyType.BUSINESS_PHONE);
                            phoneDictionary.Entry.push(entry);
                        }
                        if (pimContact.homePhone) {
                            const entry = new PhoneNumberDictionaryEntryType(pimContact.homePhone, PhoneNumberKeyType.HOME_PHONE);
                            phoneDictionary.Entry.push(entry);
                        }
                        if (pimContact.cellPhone) {
                            const entry = new PhoneNumberDictionaryEntryType(pimContact.cellPhone, PhoneNumberKeyType.MOBILE_PHONE);
                            phoneDictionary.Entry.push(entry);
                        }
                        if (pimContact.officeFax) {
                            const entry = new PhoneNumberDictionaryEntryType(pimContact.officeFax, PhoneNumberKeyType.BUSINESS_FAX);
                            phoneDictionary.Entry.push(entry);
                        }
                        if (pimContact.homeFax) {
                            const entry = new PhoneNumberDictionaryEntryType(pimContact.homeFax, PhoneNumberKeyType.HOME_FAX);
                            phoneDictionary.Entry.push(entry);
                        }
                        pimContact.otherPhones.forEach(item => {
                            const entry = new PhoneNumberDictionaryEntryType(item, PhoneNumberKeyType.OTHER_TELEPHONE);
                            phoneDictionary.Entry.push(entry);
                        });
                        if (phoneDictionary.Entry.length > 0) {
                            contact.PhoneNumbers = phoneDictionary;
                        }
                    }
                    break;
                case UnindexedFieldURIType.CONTACTS_PHYSICAL_ADDRESSES:
                    {
                        const addresses: PhysicalAddressDictionaryEntryType[] = [];
                        if (pimContact.officeAddress) {
                            const entry = new PhysicalAddressDictionaryEntryType();
                            entry.Street = pimContact.officeAddress.Street;
                            entry.State = pimContact.officeAddress.State;
                            entry.City = pimContact.officeAddress.City;
                            entry.PostalCode = pimContact.officeAddress.PostalCode;
                            entry.CountryOrRegion = pimContact.officeAddress.Country;
                            entry.Key = PhysicalAddressKeyType.BUSINESS;
                            addresses.push(entry);
                        }
                        if (pimContact.homeAddress) {
                            const entry = new PhysicalAddressDictionaryEntryType();
                            entry.Street = pimContact.homeAddress.Street;
                            entry.State = pimContact.homeAddress.State;
                            entry.City = pimContact.homeAddress.City;
                            entry.PostalCode = pimContact.homeAddress.PostalCode;
                            entry.CountryOrRegion = pimContact.homeAddress.Country;
                            entry.Key = PhysicalAddressKeyType.HOME;
                            addresses.push(entry);
                        }
                        if (pimContact.otherAddress) {
                            const entry = new PhysicalAddressDictionaryEntryType();
                            entry.Street = pimContact.otherAddress.Street;
                            entry.State = pimContact.otherAddress.State;
                            entry.City = pimContact.otherAddress.City;
                            entry.PostalCode = pimContact.otherAddress.PostalCode;
                            entry.CountryOrRegion = pimContact.otherAddress.Country;
                            entry.Key = PhysicalAddressKeyType.OTHER;
                            addresses.push(entry);
                        }
                        if (addresses.length > 0) {
                            const addressDictionary = new PhysicalAddressDictionaryType();
                            addressDictionary.Entry = addresses;
                            contact.PhysicalAddresses = addressDictionary;
                        }
                    }
                    break;
                case UnindexedFieldURIType.ITEM_BODY:
                    {
                        const comment = fromString(pimContact.body, { preserveNewlines: true });
                        contact.Body = new BodyType(comment, BodyTypeType.TEXT, false);
                    }
                    break;
                case UnindexedFieldURIType.CONTACTS_BUSINESS_HOME_PAGE:
                    if (pimContact.homepage) {
                        contact.BusinessHomePage = pimContact.homepage;
                    }
                    break;
                case UnindexedFieldURIType.CONTACTS_DEPARTMENT:
                    if (pimContact.department) {
                        contact.Department = pimContact.department;
                    }
                    break;
                case UnindexedFieldURIType.CONTACTS_GIVEN_NAME:
                    contact.GivenName = pimContact.firstName;
                    break;
                case UnindexedFieldURIType.CONTACTS_HAS_PICTURE:
                    contact.HasPicture = pimContact.attachments.includes('ContactPhoto');
                    break;
                case UnindexedFieldURIType.CONTACTS_IS_PRIVATE:
                    contact.IsPrivate = pimContact.isPrivate;
                    break;
                case UnindexedFieldURIType.CONTACTS_BIRTHDAY:
                    if (pimContact.birthday) {
                        contact.Birthday = pimContact.birthday;
                    }
                    break;
                case UnindexedFieldURIType.CONTACTS_WEDDING_ANNIVERSARY:
                    if (pimContact.anniversary) {
                        contact.WeddingAnniversary = pimContact.anniversary;
                    }
                    break;
                case UnindexedFieldURIType.CONTACTS_OFFICE_LOCATION:
                    if (pimContact.location) {
                        contact.OfficeLocation = pimContact.location;
                    }
                    break;
                case UnindexedFieldURIType.ITEM_SENSITIVITY:
                    if (pimContact.isConfidential) {
                        contact.Sensitivity = SensitivityChoicesType.CONFIDENTIAL;
                    } else {
                        contact.Sensitivity = SensitivityChoicesType.NORMAL;
                    }
                    break;
                case UnindexedFieldURIType.CONTACTS_SURNAME:
                    contact.Surname = pimContact.lastName;
                    break;
                case UnindexedFieldURIType.CONTACTS_ASSISTANT_NAME:
                    if (pimContact.spouse) {
                        contact.AssistantName = pimContact.assistant;
                    }
                    break;
                case UnindexedFieldURIType.CONTACTS_MANAGER:
                    if (pimContact.manager) {
                        contact.Manager = pimContact.manager;
                    }
                    break;
                case UnindexedFieldURIType.CONTACTS_SPOUSE_NAME:
                    if (pimContact.spouse) {
                        contact.SpouseName = pimContact.spouse;
                    }
                    break;
                case UnindexedFieldURIType.CONTACTS_FILE_AS:
                case UnindexedFieldURIType.CONTACTS_ALIAS:
                case UnindexedFieldURIType.CONTACTS_CHILDREN:
                case UnindexedFieldURIType.CONTACTS_CONTACT_SOURCE:
                case UnindexedFieldURIType.ITEM_CONVERSATION_ID:
                case UnindexedFieldURIType.CONTACTS_CULTURE:
                case UnindexedFieldURIType.ITEM_DATE_TIME_RECEIVED:
                case UnindexedFieldURIType.ITEM_DATE_TIME_SENT:
                case UnindexedFieldURIType.CONTACTS_DIRECTORY_ID:
                case UnindexedFieldURIType.CONTACTS_DIRECT_REPORTS:
                case UnindexedFieldURIType.ITEM_DISPLAY_CC:
                case UnindexedFieldURIType.CONTACTS_DISPLAY_NAME:
                case UnindexedFieldURIType.ITEM_DISPLAY_TO:
                case UnindexedFieldURIType.CONTACTS_FILE_AS_MAPPING:
                case UnindexedFieldURIType.CONTACTS_GENERATION:
                case UnindexedFieldURIType.ITEM_IMPORTANCE:
                case UnindexedFieldURIType.CONTACTS_INITIALS:
                case UnindexedFieldURIType.ITEM_INTERNET_MESSAGE_HEADERS:
                case UnindexedFieldURIType.ITEM_IS_ASSOCIATED:
                case UnindexedFieldURIType.ITEM_IS_DRAFT:
                case UnindexedFieldURIType.ITEM_IS_FROM_ME:
                case UnindexedFieldURIType.ITEM_IS_RESEND:
                case UnindexedFieldURIType.ITEM_IS_SUBMITTED:
                case UnindexedFieldURIType.ITEM_IS_UNMODIFIED:
                case UnindexedFieldURIType.ITEM_LAST_MODIFIED_NAME:
                case UnindexedFieldURIType.CONTACTS_MIDDLE_NAME:
                case UnindexedFieldURIType.CONTACTS_MILEAGE:
                case UnindexedFieldURIType.CONTACTS_NICKNAME:
                case UnindexedFieldURIType.CONTACTS_NOTES:
                case UnindexedFieldURIType.CONTACTS_PHONETIC_FIRST_NAME:
                case UnindexedFieldURIType.CONTACTS_PHONETIC_FULL_NAME:
                case UnindexedFieldURIType.CONTACTS_PHONETIC_LAST_NAME:
                case UnindexedFieldURIType.CONTACTS_PHOTO:
                case UnindexedFieldURIType.CONTACTS_POSTAL_ADDRESS_INDEX:
                case UnindexedFieldURIType.CONTACTS_PROFESSION:
                case UnindexedFieldURIType.ITEM_REMINDER_DUE_BY:
                case UnindexedFieldURIType.ITEM_REMINDER_IS_SET:
                case UnindexedFieldURIType.ITEM_REMINDER_MINUTES_BEFORE_START:
                case UnindexedFieldURIType.ITEM_RESPONSE_OBJECTS:
                case UnindexedFieldURIType.ITEM_SIZE:
                case UnindexedFieldURIType.ITEM_SUBJECT:
                case UnindexedFieldURIType.ITEM_WEB_CLIENT_EDIT_FORM_QUERY_STRING:
                case UnindexedFieldURIType.ITEM_WEB_CLIENT_READ_FORM_QUERY_STRING:
                default:
                    Logger.getInstance().debug(`Unhandled unindexed field type for contacts: ${fieldIdentifier}`);
                    handled = false;
            }
        
            if (handled === false) {
                // Common fields handled in superclass
                handled = super.updateEWSItemFieldValue(contact, pimContact, fieldIdentifier, mailboxId);
            }

            if (handled === false) {
                Logger.getInstance().warn(`Unrecognized field ${fieldIdentifier} not processed for contact`);
            }

            return handled;
    }

    // Helper functions
    /**
     * Update the extented properties for emails in a contact. 
     * @param contact The updated contact.
     */
    protected updateContactEmailProperties(contact: PimContact): void {

        const emails = this.orderedEmails(contact, 3);

        const identifiers: any = {};
        identifiers[ExtendedPropertyKeyType.DISTINGUISHED_PROPERTY_SETID] = DistinguishedPropertySetType.ADDRESS;

        const newProperty: any = {};
        newProperty[ExtendedPropertyKeyType.DISTINGUISHED_PROPERTY_SETID] = DistinguishedPropertySetType.ADDRESS;
        newProperty[ExtendedPropertyKeyType.PROPERTY_TYPE] = MapiPropertyTypeType.STRING;

        if (emails.length > 0) {
            identifiers[ExtendedPropertyKeyType.PROPERTY_ID] = MapiPropertyIds.EMAIL1_ADDRESS;
            let property = contact.findExtendedProperty(identifiers);
            if (property === undefined) {
                property = newProperty;
                property[ExtendedPropertyKeyType.PROPERTY_ID] = MapiPropertyIds.EMAIL1_ADDRESS;
            }

            property.Value = emails[0];
            contact.updateExtendedProperty(identifiers, property);
        }

        if (emails.length > 1) {
            identifiers[ExtendedPropertyKeyType.PROPERTY_ID] = MapiPropertyIds.EMAIL2_ADDRESS;
            let property = contact.findExtendedProperty(identifiers);
            if (property === undefined) {
                property = newProperty;
                property[ExtendedPropertyKeyType.PROPERTY_ID] = MapiPropertyIds.EMAIL2_ADDRESS;
            }

            property.Value = emails[1];
            contact.updateExtendedProperty(identifiers, property);
        }

        if (emails.length > 2) {
            identifiers[ExtendedPropertyKeyType.PROPERTY_ID] = MapiPropertyIds.EMAIL3_ADDRESS;
            let property = contact.findExtendedProperty(identifiers);
            if (property === undefined) {
                property = newProperty;
                property[ExtendedPropertyKeyType.PROPERTY_ID] = MapiPropertyIds.EMAIL3_ADDRESS;
            }

            property.Value = emails[2];
            contact.updateExtendedProperty(identifiers, property);
        }
    }

    /**
     * Returns the list of email address for a contact in the order they should be provided to EWS.
     * @param pimContact The Keep contact containing the emails.
     * @param max The maximum number of emails to return. The default is all defined email addresses. 
     * @returns The ordered list of emails. 
     */
    protected orderedEmails(pimContact: PimContact, max = Number.MAX_VALUE): string[] {
        const emails: string[] = [];

        const identifiers: any = {};
        identifiers[ExtendedPropertyKeyType.DISTINGUISHED_PROPERTY_SETID] = DistinguishedPropertySetType.ADDRESS;

        if (emails.length < max) {
            identifiers[ExtendedPropertyKeyType.PROPERTY_ID] = MapiPropertyIds.EMAIL1_ADDRESS;
            const property = pimContact.findExtendedProperty(identifiers);
            if (property?.["Value"] !== undefined) {
                emails.push(property["Value"]);
            }
        }

        if (emails.length < max) {
            identifiers[ExtendedPropertyKeyType.PROPERTY_ID] = MapiPropertyIds.EMAIL2_ADDRESS;
            const property = pimContact.findExtendedProperty(identifiers);
            if (property?.["Value"] !== undefined) {
                emails.push(property["Value"]);
            }
        }

        if (emails.length < max) {
            identifiers[ExtendedPropertyKeyType.PROPERTY_ID] = MapiPropertyIds.EMAIL3_ADDRESS;
            const property = pimContact.findExtendedProperty(identifiers);
            if (property?.["Value"] !== undefined) {
                emails.push(property["Value"]);
            }
        }

        // If don't have at least the max, look for other email settings in contact
        if (emails.length < max) {
            if (pimContact.primaryEmail && !emails.includes(pimContact.primaryEmail)) {
                emails.push(pimContact.primaryEmail);
            }
            // Alternate work/home first.
            if (emails.length < max) {
                const workEmailsLength = pimContact.workEmails?.length ?? 0;
                const homeEmailsLength = pimContact.homeEmails?.length ?? 0;
                for (let index = 0; index < Math.max(workEmailsLength, homeEmailsLength); index++) {
                    if (emails.length < max && workEmailsLength > index && emails.indexOf(pimContact.workEmails[index]) < 0) {
                        if (!emails.includes(pimContact.workEmails[index])) {
                            emails.push(pimContact.workEmails[index]);
                        }
                    }
                    if (emails.length < max && homeEmailsLength > index && emails.indexOf(pimContact.homeEmails[index]) < 0) {
                        if (!emails.includes(pimContact.homeEmails[index])) {
                            emails.push(pimContact.homeEmails[index]);
                        }
                    }
                }
            }

            if (emails.length < max && pimContact.schoolEmail && emails.indexOf(pimContact.schoolEmail) < 0) {
                if (!emails.includes(pimContact.schoolEmail)) {
                    emails.push(pimContact.schoolEmail);
                }
            }
            if (emails.length < max && pimContact.mobileEmail && emails.indexOf(pimContact.mobileEmail) < 0) {
                if (!emails.includes(pimContact.mobileEmail)) {
                    emails.push(pimContact.mobileEmail);
                }
            }

            // Finally add other emails
            if (pimContact.otherEmails) {
                pimContact.otherEmails.forEach(item => {
                    if (emails.length < max && !emails.includes(item)) {
                        emails.push(item);
                    }
                });
            }
        }
        return emails;
    }

    /**
     * Helpers for converting EWS contact fields to PimContact fields
     */

    protected updatePhysicalAddresses(item: ContactItemType, toPimItem: PimContact): void {
        if (item.PhysicalAddresses) {
            item.PhysicalAddresses.Entry.forEach(address => {
                if (address.Key === PhysicalAddressKeyType.BUSINESS) {
                    const officeAddress = toPimItem.officeAddress ?? new PimAddress();

                    if (address.Street) {
                        officeAddress.Street = address.Street;
                    }
                    if (address.City) {
                        officeAddress.City = address.City;
                    }
                    if (address.State) {
                        officeAddress.State = address.State;
                    }
                    if (address.CountryOrRegion) {
                        officeAddress.Country = address.CountryOrRegion;
                    }
                    if (address.PostalCode) {
                        officeAddress.PostalCode = address.PostalCode;
                    }
                    toPimItem.officeAddress = officeAddress;
                }
                else if (address.Key === PhysicalAddressKeyType.HOME) {
                    const homeAddress = toPimItem.homeAddress ?? new PimAddress();

                    if (address.Street) {
                        homeAddress.Street = address.Street;
                    }
                    if (address.City) {
                        homeAddress.City = address.City;
                    }
                    if (address.State) {
                        homeAddress.State = address.State;
                    }
                    if (address.CountryOrRegion) {
                        homeAddress.Country = address.CountryOrRegion;
                    }
                    if (address.PostalCode) {
                        homeAddress.PostalCode = address.PostalCode;
                    }
                    toPimItem.homeAddress = homeAddress;
                }
                else {
                    const otherAddress = toPimItem.otherAddress ?? new PimAddress();

                    if (address.Street) {
                        otherAddress.Street = address.Street;
                    }
                    if (address.City) {
                        otherAddress.City = address.City;
                    }
                    if (address.State) {
                        otherAddress.State = address.State;
                    }
                    if (address.CountryOrRegion) {
                        otherAddress.Country = address.CountryOrRegion;
                    }
                    if (address.PostalCode) {
                        otherAddress.PostalCode = address.PostalCode;
                    }
                    toPimItem.otherAddress = otherAddress;
                }
            });
        }
    }

    protected updateIMAddresses(item: ContactItemType, toPimItem: PimContact): void {
        if (item.ImAddresses && item.ImAddresses.Entry.length > 0) {
            // Domino only allows one IM address
            toPimItem.imAddresses = [item.ImAddresses.Entry[0].Value];

            const imAddressesObject: any = {};
            item.ImAddresses.Entry.forEach(entry => imAddressesObject[entry.Key] = entry.Value);
            toPimItem.setAdditionalProperty(UnindexedFieldURIType.CONTACTS_IM_ADDRESSES,imAddressesObject);
        }
    }

    protected updateEmailAddresses(item: ContactItemType, toPimItem: PimContact): void {
        if (item.EmailAddresses) {
            item.EmailAddresses.Entry.forEach(address => {

                switch (address.Key) {
                    case EmailAddressKeyType.EMAIL_ADDRESS_1:
                        toPimItem.primaryEmail = address.Value;
                        break;
                    case EmailAddressKeyType.EMAIL_ADDRESS_2:
                    case EmailAddressKeyType.EMAIL_ADDRESS_3:
                        toPimItem.otherEmails.push(address.Value);
                }
            });
        }
    }

    protected updatePhoneNumbers(item: ContactItemType, toPimItem: PimContact): void {
        if (item.PhoneNumbers) {
            const otherPhones: string[] = [];
            item.PhoneNumbers.Entry.forEach(phoneNumber => {
                switch (phoneNumber.Key) {
                    case PhoneNumberKeyType.PRIMARY_PHONE:
                        // FIXME:  There is no property on PimContact for primary phone number.  Should there be?
                        // toPimItem.primaryPhoneNumber = phoneNumber.Value;
                        break;
                    case PhoneNumberKeyType.BUSINESS_PHONE:
                        toPimItem.officePhone = phoneNumber.Value;
                        break;
                    case PhoneNumberKeyType.BUSINESS_FAX:
                        toPimItem.officeFax = phoneNumber.Value;
                        break;
                    case PhoneNumberKeyType.HOME_FAX:
                        toPimItem.homeFax = phoneNumber.Value;
                        break;
                    case PhoneNumberKeyType.HOME_PHONE:
                        toPimItem.homePhone = phoneNumber.Value;
                        break;
                    case PhoneNumberKeyType.MOBILE_PHONE:
                        toPimItem.cellPhone = phoneNumber.Value;
                        break;
                    default:
                        otherPhones.push(`${phoneNumber.Value}`);
                        break;
                }
            });

            if (otherPhones.length > 0) {
                otherPhones.forEach(phoneNumber => {
                    toPimItem.otherPhones.push(phoneNumber);
                });
            }
        }
    }

    /**
     * Update a value for an EWS field in a Keep PIM contact. 
     * @param pimItem The pim contact that will be updated.
     * @param fieldIdentifier The EWS unindexed field identifier
     * @param newValue The new value to set. The type is based on what fieldIdentifier is set to. To delete a field, pass in undefined. 
     * @returns true if field was handled
     */
    protected updatePimItemFieldValue(pimContact: PimContact, fieldIdentifier: UnindexedFieldURIType, newValue: any | undefined): boolean {
        if (!super.updatePimItemFieldValue(pimContact, fieldIdentifier, newValue)) {
            if (fieldIdentifier === UnindexedFieldURIType.CONTACTS_BIRTHDAY) {
                pimContact.birthday = newValue === undefined ? undefined : new Date(newValue);
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CONTACTS_BUSINESS_HOME_PAGE) {
                pimContact.homepage = newValue ?? "";
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CONTACTS_COMPANY_NAME) {
                pimContact.companyName = newValue ?? "";
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CONTACTS_COMPLETE_NAME) {
                pimContact.fullName = newValue ?? "";
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CONTACTS_DEPARTMENT) {
                pimContact.department = newValue ?? "";
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CONTACTS_EMAIL_ADDRESSES) {
                if (newValue === undefined) {
                    pimContact.removeEmailAddresses();
                }
                else {
                    const addresses: EmailAddressDictionaryType = newValue;
                    const keys = [EmailAddressKeyType.EMAIL_ADDRESS_1, EmailAddressKeyType.EMAIL_ADDRESS_2, EmailAddressKeyType.EMAIL_ADDRESS_3];
                    let emailNumber = 1;
                    const others: string[] = [];
                    addresses.Entry.forEach(address => {
                        if (keys.indexOf(address.Key) !== -1) {
                            if (emailNumber === 1) {
                                pimContact.primaryEmail = address.Value; // Treat first email as main email
                            }
                            else {
                                others.push(address.Value);
                            }
                            emailNumber = emailNumber + 1;
                        }
                    });
                    pimContact.otherEmails = others;
                }
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CONTACTS_GIVEN_NAME) {
                pimContact.firstName = newValue ?? "";
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CONTACTS_IM_ADDRESSES) {
                if (newValue === undefined) {
                    pimContact.imAddresses = [];
                }
                else {
                    const addresses: ImAddressDictionaryType = newValue;
                    pimContact.imAddresses = addresses.Entry.length > 0 ? [addresses.Entry[0].Value] : []; // Keep only supports one 
                }
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CONTACTS_JOB_TITLE) {
                pimContact.jobTitle = newValue ?? "";
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CONTACTS_MIDDLE_NAME) {
                pimContact.middleInitial = newValue ?? "";
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CONTACTS_PHONE_NUMBERS) {
                if (newValue === undefined) {
                    pimContact.removePhoneNumbers();
                }
                else {
                    const numbers: PhoneNumberDictionaryType = newValue;
                    const otherPhones: string[] = [];
                    numbers.Entry.forEach(phoneNumber => {
                        switch (phoneNumber.Key) {
                            // case PhoneNumberKeyType.PRIMARY_PHONE:
                            //     pimContact.primaryPhone = phoneNumber.Value;
                            //     break;
                            case PhoneNumberKeyType.BUSINESS_PHONE:
                                pimContact.officePhone = phoneNumber.Value;
                                break;
                            case PhoneNumberKeyType.BUSINESS_FAX:
                                pimContact.officeFax = phoneNumber.Value;
                                break;
                            case PhoneNumberKeyType.HOME_FAX:
                                pimContact.homeFax = phoneNumber.Value;
                                break;
                            case PhoneNumberKeyType.HOME_PHONE:
                                pimContact.homePhone = phoneNumber.Value;
                                break;
                            case PhoneNumberKeyType.MOBILE_PHONE:
                                pimContact.cellPhone = phoneNumber.Value;
                                break;
                            default:
                                otherPhones.push(`${phoneNumber.Value}`);
                                break;
                        }
                    });
    
                    pimContact.otherPhones = otherPhones;
                }
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CONTACTS_PHYSICAL_ADDRESSES) {
                if (newValue === undefined) {
                    pimContact.removeAddresses();
                }
                else {
                    const addresses: PhysicalAddressDictionaryType = newValue;
    
                    // These objects contains the keys "Street", "State", "City", "PostalCode", and "Country".
                    let officeAddress: any | undefined = undefined;
                    let homeAddress: any | undefined = undefined;
                    let otherAddress: any | undefined = undefined;
    
                    addresses.Entry.forEach(address => {
                        if (address.Key === PhysicalAddressKeyType.BUSINESS) {
                            if (officeAddress === undefined) {
                                officeAddress = new PimAddress();
                            }
    
                            if (address.Street) {
                                officeAddress.Street = address.Street;
                            }
                            if (address.City) {
                                officeAddress.City = address.City;
                            }
                            if (address.State) {
                                officeAddress.State = address.State;
                            }
                            if (address.CountryOrRegion) {
                                officeAddress.Country = address.CountryOrRegion;
                            }
                            if (address.PostalCode) {
                                officeAddress.PostalCode = address.PostalCode;
                            }
                        }
                        else if (address.Key === PhysicalAddressKeyType.HOME) {
                            if (homeAddress === undefined) {
                                homeAddress = new PimAddress();
                            }
    
                            if (address.Street) {
                                homeAddress.Street = address.Street;
                            }
                            if (address.City) {
                                homeAddress.City = address.City;
                            }
                            if (address.State) {
                                homeAddress.State = address.State;
                            }
                            if (address.CountryOrRegion) {
                                homeAddress.Country = address.CountryOrRegion;
                            }
                            if (address.PostalCode) {
                                homeAddress.PostalCode = address.PostalCode;
                            }
                        }
                        else {
                            if (otherAddress === undefined) {
                                otherAddress = new PimAddress();
                            }
    
                            if (address.Street) {
                                otherAddress.Street = address.Street;
                            }
                            if (address.City) {
                                otherAddress.City = address.City;
                            }
                            if (address.State) {
                                otherAddress.State = address.State;
                            }
                            if (address.CountryOrRegion) {
                                otherAddress.Country = address.CountryOrRegion;
                            }
                            if (address.PostalCode) {
                                otherAddress.PostalCode = address.PostalCode;
                            }
                        }
                    });
    
                    if (officeAddress) {
                        pimContact.officeAddress = officeAddress;
                    }
    
                    if (homeAddress) {
                        pimContact.homeAddress = homeAddress;
                    }
    
                    if (otherAddress) {
                        pimContact.otherAddress = otherAddress;
                    }
                }
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CONTACTS_SURNAME) {
                pimContact.lastName = newValue ?? "";
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CONTACTS_WEDDING_ANNIVERSARY) {
                pimContact.anniversary = (newValue === undefined) ? undefined : new Date(newValue);
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CONTACTS_COMMENT) {
                pimContact.body = newValue ?? "";
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CONTACTS_IS_PRIVATE) {
                pimContact.isPrivate = newValue ?? false;
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CONTACTS_MANAGER) {
                pimContact.manager = newValue ?? "";
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CONTACTS_ASSISTANT_NAME) {
                pimContact.assistant = newValue ?? "";
            }
            else if (fieldIdentifier === UnindexedFieldURIType.CONTACTS_SPOUSE_NAME) {
                pimContact.spouse = newValue ?? "";
            }
            else {
                Logger.getInstance().error(`Unsupported field ${fieldIdentifier} for PIM contact ${pimContact.unid}`);
                return false;
            }
        }

        /*
        The following item fields are not currently supported:

        CONTACTS_ALIAS = 'contacts:Alias',
        CONTACTS_ASSISTANT_NAME = 'contacts:AssistantName',
        CONTACTS_CHILDREN = 'contacts:Children',
        CONTACTS_COMPANIES = 'contacts:Companies',
        CONTACTS_CONTACT_SOURCE = 'contacts:ContactSource',
        CONTACTS_CULTURE = 'contacts:Culture',
        CONTACTS_DISPLAY_NAME = 'contacts:DisplayName',
        CONTACTS_DIRECTORY_ID = 'contacts:DirectoryId',
        CONTACTS_DIRECT_REPORTS = 'contacts:DirectReports',
        CONTACTS_ABCH_EMAIL_ADDRESSES = 'contacts:AbchEmailAddresses',
        CONTACTS_FILE_AS = 'contacts:FileAs',
        CONTACTS_FILE_AS_MAPPING = 'contacts:FileAsMapping',
        CONTACTS_GENERATION = 'contacts:Generation',
        CONTACTS_INITIALS = 'contacts:Initials',
        CONTACTS_MANAGER = 'contacts:Manager',
        CONTACTS_MANAGER_MAILBOX = 'contacts:ManagerMailbox',
        CONTACTS_MILEAGE = 'contacts:Mileage',
        CONTACTS_MS_EXCHANGE_CERTIFICATE = 'contacts:MSExchangeCertificate',
        CONTACTS_NICKNAME = 'contacts:Nickname',
        CONTACTS_NOTES = 'contacts:Notes',
        CONTACTS_OFFICE_LOCATION = 'contacts:OfficeLocation',
        CONTACTS_PHONETIC_FULL_NAME = 'contacts:PhoneticFullName',
        CONTACTS_PHONETIC_FIRST_NAME = 'contacts:PhoneticFirstName',
        CONTACTS_PHONETIC_LAST_NAME = 'contacts:PhoneticLastName',
        CONTACTS_PHOTO = 'contacts:Photo',
        CONTACTS_POSTAL_ADDRESS_INDEX = 'contacts:PostalAddressIndex',
        CONTACTS_PROFESSION = 'contacts:Profession',
        CONTACTS_SPOUSE_NAME = 'contacts:SpouseName',
        CONTACTS_USER_SMIME_CERTIFICATE = 'contacts:UserSMIMECertificate',
        CONTACTS_HAS_PICTURE = 'contacts:HasPicture',
        CONTACTS_ACCOUNT_NAME = 'contacts:AccountName',
        CONTACTS_IS_AUTO_UPDATE_DISABLED = 'contacts:IsAutoUpdateDisabled',
        CONTACTS_IS_MESSENGER_ENABLED = 'contacts:IsMessengerEnabled',
        CONTACTS_CONTACT_SHORT_ID = 'contacts:ContactShortId',
        CONTACTS_CONTACT_TYPE = 'contacts:ContactType',
        CONTACTS_CREATED_BY = 'contacts:CreatedBy',
        CONTACTS_GENDER = 'contacts:Gender',
        CONTACTS_IS_HIDDEN = 'contacts:IsHidden',
        CONTACTS_OBJECT_ID = 'contacts:ObjectId',
        CONTACTS_PASSPORT_ID = 'contacts:PassportId',
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
        */

       return true;
    }
}
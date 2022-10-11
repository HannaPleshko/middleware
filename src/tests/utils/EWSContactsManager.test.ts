/* eslint-disable @typescript-eslint/no-non-null-assertion */
import { expect, sinon } from '@loopback/testlab';
import {
    PimItem, PimLabel, UserInfo, KeepPimConstants,
    PimItemFormat, PimItemFactory, PimContact, KeepPimContactsManager, KeepPimLabelManager, PimLabelTypes, KeepPimMessageManager, PimAddress, KeepPimBaseResults
} from '@hcllabs/openclientkeepcomponent';
import {
    UnindexedFieldURIType, MessageDispositionType, BodyTypeType, DefaultShapeNamesType,
    DistinguishedPropertySetType, DictionaryURIType, EmailAddressKeyType, PhysicalAddressKeyType, PhoneNumberKeyType, ImAddressKeyType, MapiPropertyTypeType, SensitivityChoicesType, FieldIndexValue, MapiPropertyIds, MailboxTypeType
} from '../../models/enum.model';
import {
    ItemResponseShapeType, ItemType, ItemChangeType, ItemIdType, BodyType, NonEmptyArrayOfItemChangeDescriptionsType, DeleteItemFieldType, SetItemFieldType,
    ContactItemType, ExtendedPropertyType, AppendToItemFieldType, CompleteNameType, EmailAddressDictionaryType, EmailAddressDictionaryEntryType,
    PhysicalAddressDictionaryType, PhysicalAddressDictionaryEntryType, PhoneNumberDictionaryType, PhoneNumberDictionaryEntryType,
    ImAddressDictionaryType, ImAddressDictionaryEntryType, MessageType, FolderIdType, NonEmptyArrayOfPathsToElementType,
} from '../../models/mail.model';
import {
    ArrayOfStringsType, PathToExtendedFieldType, PathToIndexedFieldType, PathToUnindexedFieldType
} from '../../models/common.model';
import { EWSContactsManager, getEWSId, getKeepIdPair, getValueForExtendedFieldURI, identifiersForPathToExtendedFieldType } from '../../utils';
import { UserContext } from '../../keepcomponent';
import {
    generateTestContactsLabel,
    getContext, stubContactsManager, stubLabelsManager, stubMessageManager
} from '../unitTestHelper';
import { Request } from '@loopback/rest';

describe('EWSContactsManager tests', () => {

    const testUser = "test.user@test.org";

    // Mock subclass of EWSContacts manager for testing protected functions
    class MockEWSManager extends EWSContactsManager {

        createItem(
            item: ItemType, 
            userInfo: UserInfo, 
            request: Request,
            disposition?: MessageDispositionType,
            toLabel?: PimLabel,
            mailboxId?: string
        ): Promise<ItemType[]> {
            return super.createItem(item, userInfo, request, disposition, toLabel, mailboxId);
        }

        async deleteItem(item: string | PimContact, userInfo: UserInfo, mailboxId?: string, hardDelete = false): Promise<void> {
            return super.deleteItem(item, userInfo, mailboxId, hardDelete);
        }

        async updateItem(
            pimItem: PimContact, 
            change: ItemChangeType, 
            userInfo: UserInfo, 
            request: Request, 
            toLabel?: PimLabel,
            mailboxId?: string
        ): Promise<ItemType | undefined> {
            return super.updateItem(pimItem, change, userInfo, request, toLabel, mailboxId);
        }

        async getItems(
            userInfo: UserInfo, 
            request: Request, 
            shape: ItemResponseShapeType, 
            startIndex?: number, 
            count?: number, 
            fromLabel?: PimLabel,
            mailboxId?: string
        ): Promise<ItemType[]> {
            return super.getItems(userInfo, request, shape, startIndex, count, fromLabel, mailboxId);
        }

        addRequestedPropertiesToEWSItem(pimItem: PimItem, toItem: ItemType, shape?: ItemResponseShapeType, mailboxId?: string): void {
            super.addRequestedPropertiesToEWSItem(pimItem, toItem, shape, mailboxId);
        }

        updateEWSItemFieldValue(
            item: ItemType, 
            pimContact: PimContact, 
            fieldId: UnindexedFieldURIType, 
            mailboxId?: string
        ): boolean {
            return super.updateEWSItemFieldValue(item, pimContact, fieldId, mailboxId);
        }

        pimItemFromEWSItem(item: ContactItemType, request: Request, existing?: object): PimContact {
            return super.pimItemFromEWSItem(item, request, existing);
        }

        pimItemToEWSItem(
            pimItem: PimContact, 
            userInfo: UserInfo, 
            request: Request, 
            shape: ItemResponseShapeType,
            mailboxId?: string,
            parentFolderId?: string
        ): Promise<ContactItemType> {
            return super.pimItemToEWSItem(pimItem, userInfo, request, shape, mailboxId, parentFolderId);
        }

        updateContactEmailProperties(contact: PimContact): void {
            super.updateContactEmailProperties(contact);
        }

        orderedEmails(pimContact: PimContact, max = Number.MAX_VALUE): string[] {
            return super.orderedEmails(pimContact, max);
        }
    }

    describe('Test getInstance', () => {

        it('getInstance', function () {

            const manager = EWSContactsManager.getInstance();
            expect(manager).to.be.instanceof(EWSContactsManager);

            const manager2 = EWSContactsManager.getInstance();
            expect(manager).to.be.equal(manager2);
        });
    });

    describe('Test pimItemFromEWSItem', () => {

        let contact: ContactItemType;
        const created = new Date();
        const modified = new Date();
        const uid = 'uid';

        beforeEach(function () {
            contact = new ContactItemType();
            const mailboxId = 'test@test.com';
            const itemEWSId = getEWSId(uid, mailboxId);
            contact.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            const parentFolderEWSId = getEWSId('parent');
            contact.ParentFolderId = new FolderIdType(parentFolderEWSId, `ck-${parentFolderEWSId}`);
            contact.DateTimeCreated = created;
            contact.LastModifiedTime = modified;
            const sent = new Date();
            contact.DateTimeSent = sent;
            contact.Categories = new ArrayOfStringsType();
            contact.Categories.String = ['cat1', 'cat2'];
            contact.JobTitle = 'Engineer';
            contact.CompanyName = 'HCL';
            contact.Department = 'LNMA';
            contact.Body = new BodyType('BODY', BodyTypeType.TEXT);
            contact.Manager = 'Chris';
            contact.AssistantName = 'Assistant';
            contact.SpouseName = 'Spouse';

            // Fields not currently mapped to PimContact, but processed in the code
            contact.Birthday = new Date();
            contact.BusinessHomePage = 'https://www.hcl.com';
            contact.WeddingAnniversary = new Date();
            contact.OfficeLocation = 'Cary, NC';
        });

        it('empty ContactItemType', () => {
            const manager = new MockEWSManager();
            contact = new ContactItemType();
            const context = getContext(testUser, 'password');
            const pimContact = manager.pimItemFromEWSItem(contact, context.request);
            expect(pimContact.unid).to.be.undefined();
            expect(pimContact.createdDate).to.be.undefined();
        });

        it('common fields populated', () => {
            const manager = new MockEWSManager();
            const context = getContext(testUser, 'password');

            const pimContact = manager.pimItemFromEWSItem(contact, context.request);
            expect(pimContact.unid).to.be.equal(uid);
            const parentIds = pimContact.parentFolderIds;
            expect(parentIds).to.be.an.Array();
            if (parentIds) {
                expect(parentIds.length).to.be.equal(1);
                const [parentFolderId] = getKeepIdPair(contact.ParentFolderId!.Id);
                expect(parentIds[0]).to.be.equal(parentFolderId);
            }
            expect(pimContact.createdDate).to.be.eql(created);
            expect(pimContact.lastModifiedDate).to.be.eql(modified);
            expect(pimContact.categories).to.be.an.Array();
            expect(pimContact.categories.length).to.be.equal(2);
            expect(pimContact.jobTitle).to.be.equal(contact.JobTitle);
            expect(pimContact.companyName).to.be.equal(contact.CompanyName);
            expect(pimContact.department).to.be.equal(contact.Department);
            expect(pimContact.body).to.be.equal(contact.Body?.Value);
            expect(pimContact.manager).to.be.equal(contact.Manager);
            expect(pimContact.assistant).to.be.equal(contact.AssistantName);
            expect(pimContact.spouse).to.be.equal(contact.SpouseName);
        });

        it('test name fields', () => {
            contact.CompleteName = new CompleteNameType();
            contact.CompleteName.FirstName = 'First';
            contact.CompleteName.LastName = 'Last';
            contact.CompleteName.MiddleName = 'Middle';
            contact.CompleteName.Nickname = 'Nickname';
            contact.CompleteName.Suffix = 'Jr';
            contact.CompleteName.Title = 'Mrs';
            contact.CompleteName.Initials = 'FML';
            contact.CompleteName.FullName = 'Mrs First Middle Last Jr';

            const manager = new MockEWSManager();
            const context = getContext(testUser, 'password');
            let pimContact = manager.pimItemFromEWSItem(contact, context.request);

            expect(pimContact.firstName).to.be.equal(contact.CompleteName.FirstName);
            expect(pimContact.lastName).to.be.equal(contact.CompleteName.LastName);
            expect(pimContact.middleInitial).to.be.equal(contact.CompleteName.MiddleName);
            expect(pimContact.title).to.be.equal(contact.CompleteName.Title);
            expect(pimContact.suffix).to.be.equal(contact.CompleteName.Suffix);
            if(pimContact.fullName){
                expect(pimContact.fullName[0]).to.be.equal(contact.CompleteName.FullName);
            }

            contact.CompleteName = undefined;
            contact.GivenName = 'Given';
            contact.Surname = 'Surname';
            contact.MiddleName = 'Middle';
            pimContact = manager.pimItemFromEWSItem(contact, context.request);
            expect(pimContact.firstName).to.be.equal(contact.GivenName);
            expect(pimContact.lastName).to.be.equal(contact.Surname);
            expect(pimContact.middleInitial).to.be.equal(contact.MiddleName);
            expect(pimContact.title).to.be.undefined();
        });

        it('test email fields', () => {
            contact.EmailAddresses = new EmailAddressDictionaryType();
            contact.EmailAddresses.Entry = [];
            const entry = new EmailAddressDictionaryEntryType('email1@hcl.com', EmailAddressKeyType.EMAIL_ADDRESS_1);
            contact.EmailAddresses.Entry.push(entry);
            const entry2 = new EmailAddressDictionaryEntryType('email2@hcl.com', EmailAddressKeyType.EMAIL_ADDRESS_2);
            contact.EmailAddresses.Entry.push(entry2);
            const entry3 = new EmailAddressDictionaryEntryType('email3@hcl.com', EmailAddressKeyType.EMAIL_ADDRESS_3);
            contact.EmailAddresses.Entry.push(entry3);

            const manager = new MockEWSManager();
            const context = getContext(testUser, 'password');
            const pimContact = manager.pimItemFromEWSItem(contact, context.request);
            expect(pimContact).to.not.be.undefined();
            expect(pimContact.primaryEmail).to.be.equal('email1@hcl.com');
            expect(pimContact.otherEmails).to.be.an.Array();
            expect(pimContact.otherEmails.length).to.be.equal(2);
        });

        it('test physical addresses', () => {
            contact.PhysicalAddresses = new PhysicalAddressDictionaryType();
            contact.PhysicalAddresses.Entry = [];
            const home = new PhysicalAddressDictionaryEntryType();
            home.City = 'Apex';
            home.State = 'NC';
            home.CountryOrRegion = 'USA';
            home.PostalCode = '27523';
            home.Street = '1313 Mockingbird Ln';
            home.Key = PhysicalAddressKeyType.HOME
            contact.PhysicalAddresses.Entry.push(home);

            const work = new PhysicalAddressDictionaryEntryType();
            work.City = 'Cary';
            work.State = 'NC';
            work.CountryOrRegion = 'USA';
            work.PostalCode = '27511';
            work.Street = '200 Lucent Ln';
            work.Key = PhysicalAddressKeyType.BUSINESS;
            contact.PhysicalAddresses.Entry.push(work);

            const other = new PhysicalAddressDictionaryEntryType();
            other.City = 'Washington';
            other.State = 'DC';
            other.CountryOrRegion = 'USA';
            other.PostalCode = '20500';
            other.Street = '1600 Pennsylvania Ave';
            other.Key = PhysicalAddressKeyType.OTHER;
            contact.PhysicalAddresses.Entry.push(other);

            const manager = new MockEWSManager();
            const context = getContext(testUser, 'password');
            let pimContact = manager.pimItemFromEWSItem(contact, context.request);
            let pimAddr = pimContact.homeAddress;
            expect(pimAddr).to.not.be.undefined();
            if (pimAddr) {
                expect(pimAddr.City).to.be.equal(home.City);
                expect(pimAddr.State).to.be.equal(home.State);
                expect(pimAddr.Street).to.be.equal(home.Street);
                expect(pimAddr.PostalCode).to.be.equal(home.PostalCode);
                expect(pimAddr.Country).to.be.equal(home.CountryOrRegion);
            }

            pimAddr = pimContact.officeAddress;
            expect(pimAddr).to.not.be.undefined();
            if (pimAddr) {
                expect(pimAddr.City).to.be.equal(work.City);
                expect(pimAddr.State).to.be.equal(work.State);
                expect(pimAddr.Street).to.be.equal(work.Street);
                expect(pimAddr.PostalCode).to.be.equal(work.PostalCode);
                expect(pimAddr.Country).to.be.equal(work.CountryOrRegion);
            }

            pimAddr = pimContact.otherAddress;
            expect(pimAddr).to.not.be.undefined();
            if (pimAddr) {
                expect(pimAddr.City).to.be.equal(other.City);
                expect(pimAddr.State).to.be.equal(other.State);
                expect(pimAddr.Street).to.be.equal(other.Street);
                expect(pimAddr.PostalCode).to.be.equal(other.PostalCode);
                expect(pimAddr.Country).to.be.equal(other.CountryOrRegion);
            }

            // Try empty physical addresses
            contact.PhysicalAddresses = new PhysicalAddressDictionaryType();
            contact.PhysicalAddresses.Entry = [];
            const emptyHome = new PhysicalAddressDictionaryEntryType();
            emptyHome.Key = PhysicalAddressKeyType.HOME
            contact.PhysicalAddresses.Entry.push(emptyHome);

            const emptyWork = new PhysicalAddressDictionaryEntryType();
            emptyWork.Key = PhysicalAddressKeyType.BUSINESS;
            contact.PhysicalAddresses.Entry.push(emptyWork);

            const emptyOther = new PhysicalAddressDictionaryEntryType();
            emptyOther.Key = PhysicalAddressKeyType.OTHER;
            contact.PhysicalAddresses.Entry.push(emptyOther);
            pimContact = manager.pimItemFromEWSItem(contact, context.request);
            expect(pimContact.homeAddress).to.be.undefined();
        });

        it('test phone numbers', () => {
            contact.PhoneNumbers = new PhoneNumberDictionaryType();
            contact.PhoneNumbers.Entry = [];
            const primary = new PhoneNumberDictionaryEntryType('919-555-1212', PhoneNumberKeyType.PRIMARY_PHONE);
            contact.PhoneNumbers.Entry.push(primary);
            const business = new PhoneNumberDictionaryEntryType('919-555-1111', PhoneNumberKeyType.BUSINESS_PHONE);
            contact.PhoneNumbers.Entry.push(business);
            const businessFax = new PhoneNumberDictionaryEntryType('919-555-2121', PhoneNumberKeyType.BUSINESS_FAX);
            contact.PhoneNumbers.Entry.push(businessFax);
            const home = new PhoneNumberDictionaryEntryType('919-867-5309', PhoneNumberKeyType.HOME_PHONE);
            contact.PhoneNumbers.Entry.push(home);
            const homeFax = new PhoneNumberDictionaryEntryType('111-111-1111', PhoneNumberKeyType.HOME_FAX);
            contact.PhoneNumbers.Entry.push(homeFax);
            const mobile = new PhoneNumberDictionaryEntryType('919-333-3332', PhoneNumberKeyType.MOBILE_PHONE);
            contact.PhoneNumbers.Entry.push(mobile);
            const other = new PhoneNumberDictionaryEntryType('919-444-1212', PhoneNumberKeyType.BUSINESS_MOBILE);
            contact.PhoneNumbers.Entry.push(other);

            const manager = new MockEWSManager();
            const context = getContext(testUser, 'password');
            let pimContact = manager.pimItemFromEWSItem(contact, context.request);
            expect(pimContact.homePhone).to.be.equal(home.Value);
            // FIXME:  Primary phone not being used by PimContact
            // expect(pimContact.primaryPhone).to.be.equal(primary.Value);
            expect(pimContact.officePhone).to.be.equal(business.Value);
            expect(pimContact.officeFax).to.be.equal(businessFax.Value);
            expect(pimContact.homeFax).to.be.equal(homeFax.Value);
            expect(pimContact.cellPhone).to.be.equal(mobile.Value);
            expect(pimContact.otherPhones[0]).to.be.equal(other.Value);

            // No otherPHones
            contact.PhoneNumbers = new PhoneNumberDictionaryType();
            contact.PhoneNumbers.Entry = [];
            contact.PhoneNumbers.Entry.push(primary);
            contact.PhoneNumbers.Entry.push(business);
            pimContact = manager.pimItemFromEWSItem(contact, context.request);
            expect(pimContact.otherPhones).to.be.eql([]);

        });

        it('test variations', () => {
            contact.CompanyName = undefined;
            contact.Companies = new ArrayOfStringsType();
            contact.Companies.String = ['HCL', 'IBM'];
            contact.IsPrivate = true;
            contact.ImAddresses = new ImAddressDictionaryType();
            contact.ImAddresses.Entry = [new ImAddressDictionaryEntryType('user@hcl.com', ImAddressKeyType.IM_ADDRESS_1)];

            const manager = new MockEWSManager();
            const context = getContext(testUser, 'password');
            const pimContact = manager.pimItemFromEWSItem(contact, context.request, { "Type": "Person" });
            expect(pimContact.imAddresses[0]).to.be.equal('user@hcl.com');
            expect(pimContact.companyName).to.be.equal('HCL');
            expect(pimContact.isPrivate).to.be.true();
        });
    });

    describe('Test fieldsForDefaultShape', () => {

        it('test default fields', () => {
            const manager = EWSContactsManager.getInstance();
            const fields = manager.fieldsForDefaultShape();
            expect(fields).to.be.an.Array();
            // Verify size is correct
            expect(fields.length).to.be.equal(12);
            // Check a few of the fields
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS);
            expect(fields).to.containEql(UnindexedFieldURIType.CONTACTS_IM_ADDRESSES);
            expect(fields).to.containEql(UnindexedFieldURIType.CONTACTS_PHONE_NUMBERS);
        });

    });

    describe('Test fieldsForAllPropertiesShape', () => {
        it('test all properties fields', () => {
            const manager = EWSContactsManager.getInstance();
            const fields = manager.fieldsForAllPropertiesShape();
            expect(fields).to.be.an.Array();
            // Verify the size is correct
            expect(fields.length).to.be.equal(73);
            // Check a few of the fields
            expect(fields).to.containEql(UnindexedFieldURIType.CONTACTS_NOTES);
            expect(fields).to.containEql(UnindexedFieldURIType.CONTACTS_OFFICE_LOCATION);
            expect(fields).to.containEql(UnindexedFieldURIType.CONTACTS_PHONETIC_FIRST_NAME);
        });
    });


    describe('Test getItems', () => {

        // Clean up the stubs
        afterEach(() => {
            sinon.restore(); // Reset stubs
        });

        it('Expected results', async () => {
            const parentLabel = generateTestContactsLabel();
            const pimContact = PimItemFactory.newPimContact({}, PimItemFormat.DOCUMENT);
            pimContact.unid = 'UNID';
            pimContact.parentFolderIds = [parentLabel.folderId];
            pimContact.firstName = 'First';
            pimContact.lastName = 'Last';

            const pimGroup = PimItemFactory.newPimContact({ 'Type': 'Group' }, PimItemFormat.DOCUMENT);
            pimGroup.unid = 'GROUP-ID';

            const contactsManagerStub = sinon.createStubInstance(KeepPimContactsManager);
            contactsManagerStub.getContacts.resolves([pimContact, pimGroup]);
            stubContactsManager(contactsManagerStub);

            const manager = EWSContactsManager.getInstance();
            const userInfo = new UserContext();
            const context = getContext(testUser, 'password');
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            const mailboxId = 'test@test.com';
            const itemEWSId = getEWSId(pimContact.unid, mailboxId);
            let messages = await manager.getItems(userInfo, context.request, shape, 0, 10, parentLabel, mailboxId);

            expect(messages).to.be.an.Array();
            expect(messages.length).to.be.equal(1);
            let message = messages[0];
            expect(message.ItemId?.Id).to.be.equal(itemEWSId);

            // No labelId
            messages = await manager.getItems(userInfo, context.request, shape);

            expect(messages).to.be.an.Array();
            expect(messages.length).to.be.equal(1);
            message = messages[0];
            expect(message.ItemId?.Id).to.be.equal(getEWSId(pimContact.unid));
        });

        it('getContacts throws an error', async () => {
            const parentLabel = generateTestContactsLabel();
            const pimContact = PimItemFactory.newPimContact({}, PimItemFormat.DOCUMENT);
            pimContact.unid = 'UNID';
            pimContact.parentFolderIds = [parentLabel.folderId];
            pimContact.firstName = 'First';
            pimContact.lastName = 'Last';
            const contactsManagerStub = sinon.createStubInstance(KeepPimContactsManager);
            contactsManagerStub.getContacts.throws();

            stubContactsManager(contactsManagerStub);

            const manager = EWSContactsManager.getInstance();
            const userInfo = new UserContext();
            const context = getContext(testUser, 'password');
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            const messages = await manager.getItems(userInfo, context.request, shape, 0, 10, parentLabel);
            expect(messages).to.be.an.Array();
            expect(messages.length).to.be.equal(0);
        });


        describe('Test updateItem', () => {
            const context = getContext(testUser, 'password');
            const userInfo = new UserContext();
            const manager = new MockEWSManager();
            let pimContact: PimContact;
            let newContact: PimContact;
            let contactsManagerStub;

            beforeEach(function () {
                pimContact = PimItemFactory.newPimContact({}, PimItemFormat.DOCUMENT);
                pimContact.subject = 'SUBJECT';
                pimContact.body = 'BODY';
                pimContact.unid = 'UNID';
                pimContact.parentFolderIds = ['PARENT'];

                newContact = PimItemFactory.newPimContact({}, PimItemFormat.DOCUMENT);
                newContact.unid = 'UNID';
                newContact.parentFolderIds = ['PARENT'];
                contactsManagerStub = sinon.createStubInstance(KeepPimContactsManager);
                const updateBaseResults: KeepPimBaseResults = {
                    unid: 'UNID',
                    status: 200,
                    message: 'Success'
                };
                contactsManagerStub.updateContact.resolves(updateBaseResults);
                stubContactsManager(contactsManagerStub);
            });

            // Clean up the stubs
            afterEach(() => {
                sinon.restore(); // Reset stubs
            });

            it('test set field update', async () => {
                const changes = new ItemChangeType();
                changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

                const fieldURIChange = new SetItemFieldType()
                fieldURIChange.FieldURI = new PathToUnindexedFieldType();
                fieldURIChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
                fieldURIChange.Item = new ContactItemType();
                fieldURIChange.Item.ItemId = new ItemIdType(pimContact.unid, `ck-${pimContact.unid}`);
                fieldURIChange.Item.Subject = 'UPDATED-SUBJECT';

                changes.ItemId = new ItemIdType(pimContact.unid, `ck-${pimContact.unid}`);
                changes.Updates.push(fieldURIChange);

                const extendedFieldChange = new SetItemFieldType();
                extendedFieldChange.ExtendedFieldURI = new PathToExtendedFieldType();
                extendedFieldChange.ExtendedFieldURI.DistinguishedPropertySetId = DistinguishedPropertySetType.ADDRESS;

                extendedFieldChange.Item = new ContactItemType();
                extendedFieldChange.Item.ExtendedProperty = [];
                const prop = new ExtendedPropertyType();
                prop.ExtendedFieldURI = new PathToExtendedFieldType();
                prop.ExtendedFieldURI.DistinguishedPropertySetId = DistinguishedPropertySetType.ADDRESS;
                prop.ExtendedFieldURI.PropertyId = 32899;
                prop.ExtendedFieldURI.PropertyType = MapiPropertyTypeType.STRING;
                prop.Value = 'newemail@hcl.com';
                extendedFieldChange.Item.ExtendedProperty.push(prop);
                changes.Updates.push(extendedFieldChange);

                // Add this to drive a not handled warning
                const indextedFieldChange = new SetItemFieldType();
                indextedFieldChange.IndexedFieldURI = new PathToIndexedFieldType();
                indextedFieldChange.IndexedFieldURI.FieldURI = DictionaryURIType.CONTACTS_EMAIL_ADDRESS;
                indextedFieldChange.Item = new ContactItemType();
                changes.Updates.push(indextedFieldChange);

                const mailboxId = 'test@test.com';
                const parentEWSId = getEWSId('PARENT-LABEL', mailboxId);

                let contact = await manager.updateItem(pimContact, changes, userInfo, context.request);
                expect(contact).to.not.be.undefined();
                expect(contact?.ItemId?.Id).to.be.equal(getEWSId(pimContact.unid));

                const label = PimItemFactory.newPimLabel({FolderId: 'PARENT-LABEL'});
                // Try with parent folderId
                contact = await manager.updateItem(pimContact, changes, userInfo, context.request, label, mailboxId);
                expect(contact).to.not.be.undefined();
                expect(contact?.ParentFolderId?.Id).to.be.equal(parentEWSId);

                // Try with no existing parent folder
                pimContact.parentFolderIds = undefined;
                contact = await manager.updateItem(pimContact, changes, userInfo, context.request, label, mailboxId);
                expect(contact).to.not.be.undefined();
                expect(contact?.ParentFolderId?.Id).to.be.equal(parentEWSId);

                // Try with no parent folder at all
                pimContact.parentFolderIds = undefined;
                contact = await manager.updateItem(pimContact, changes, userInfo, context.request);
                expect(contact).to.not.be.undefined();
                expect(contact?.ParentFolderId?.Id).to.be.undefined();
            });

            it('test append field update', async () => {
                let changes = new ItemChangeType();
                changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

                // Try body
                let fieldURIChangeBody = new AppendToItemFieldType()
                fieldURIChangeBody.FieldURI = new PathToUnindexedFieldType();
                fieldURIChangeBody.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_BODY;
                fieldURIChangeBody.Contact = new ContactItemType();
                fieldURIChangeBody.Contact.ItemId = new ItemIdType(pimContact.unid, `ck-${pimContact.unid}`);
                fieldURIChangeBody.Contact.Body = new BodyType('APPENDED-BODY', BodyTypeType.TEXT)

                changes.ItemId = new ItemIdType(pimContact.unid, `ck-${pimContact.unid}`);
                changes.Updates.push(fieldURIChangeBody);

                // Try subject (not supported)
                const fieldURIChangeSubject = new AppendToItemFieldType()
                fieldURIChangeSubject.FieldURI = new PathToUnindexedFieldType();
                fieldURIChangeSubject.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
                fieldURIChangeSubject.Contact = new ContactItemType();
                fieldURIChangeSubject.Contact.ItemId = new ItemIdType(pimContact.unid, `ck-${pimContact.unid}`);
                fieldURIChangeSubject.Contact.Subject = 'APPENDED-SUBJECT';

                changes.ItemId = new ItemIdType(pimContact.unid, `ck-${pimContact.unid}`);
                changes.Updates.push(fieldURIChangeSubject);

                // No extended fields are supported
                const extendedFieldChange = new AppendToItemFieldType();
                extendedFieldChange.ExtendedFieldURI = new PathToExtendedFieldType();
                extendedFieldChange.ExtendedFieldURI.DistinguishedPropertySetId = DistinguishedPropertySetType.ADDRESS;
                extendedFieldChange.ExtendedFieldURI.PropertyId = 32899;
                extendedFieldChange.ExtendedFieldURI.PropertyType = MapiPropertyTypeType.STRING;

                extendedFieldChange.Item = new ContactItemType();
                extendedFieldChange.Item.ExtendedProperty = [];
                const prop = new ExtendedPropertyType();
                prop.ExtendedFieldURI = new PathToExtendedFieldType();
                prop.ExtendedFieldURI.DistinguishedPropertySetId = DistinguishedPropertySetType.ADDRESS;
                prop.ExtendedFieldURI.PropertyId = 32899;
                prop.ExtendedFieldURI.PropertyType = MapiPropertyTypeType.STRING;
                prop.Value = 'append';
                extendedFieldChange.Item.ExtendedProperty.push(prop);
                changes.Updates.push(extendedFieldChange);

                let contact = await manager.updateItem(pimContact, changes, userInfo, context.request);
                expect(contact).to.not.be.undefined();
                expect(contact?.ItemId?.Id).to.be.equal(getEWSId(pimContact.unid));

                // Try with no body
                changes = new ItemChangeType();
                changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

                fieldURIChangeBody = new AppendToItemFieldType()
                fieldURIChangeBody.FieldURI = new PathToUnindexedFieldType();
                fieldURIChangeBody.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_BODY;
                fieldURIChangeBody.Item = new ContactItemType();
                fieldURIChangeBody.Item.ItemId = new ItemIdType(pimContact.unid, `ck-${pimContact.unid}`);

                changes.ItemId = new ItemIdType(pimContact.unid, `ck-${pimContact.unid}`);
                changes.Updates.push(fieldURIChangeBody);

                contact = await manager.updateItem(pimContact, changes, userInfo, context.request);
                expect(contact).to.not.be.undefined();
                expect(contact?.ItemId?.Id).to.be.equal(getEWSId(pimContact.unid));
            });


            it('test delete field update', async () => {
                const changes = new ItemChangeType();
                changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();
                const fieldURIChange = new DeleteItemFieldType();
                fieldURIChange.FieldURI = new PathToUnindexedFieldType();
                fieldURIChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
                changes.ItemId = new ItemIdType(pimContact.unid, `ck-${pimContact.unid}`);
                changes.Updates.push(fieldURIChange);

                const extendedFieldChange = new DeleteItemFieldType();
                extendedFieldChange.ExtendedFieldURI = new PathToExtendedFieldType();
                extendedFieldChange.ExtendedFieldURI.DistinguishedPropertySetId = DistinguishedPropertySetType.COMMON;
                extendedFieldChange.ExtendedFieldURI.PropertyId = 32899;
                extendedFieldChange.ExtendedFieldURI.PropertyType = MapiPropertyTypeType.STRING;
                changes.Updates.push(extendedFieldChange);

                const indextedFieldChange = new DeleteItemFieldType();
                indextedFieldChange.IndexedFieldURI = new PathToIndexedFieldType();
                indextedFieldChange.IndexedFieldURI.FieldURI = DictionaryURIType.CONTACTS_IM_ADDRESS;
                changes.Updates.push(indextedFieldChange);

                const contact = await manager.updateItem(pimContact, changes, userInfo, context.request);
                expect(contact).to.not.be.undefined();
                expect(contact?.ItemId?.Id).to.be.equal(getEWSId(pimContact.unid));
            });

            it('test deleting im addresses', async () => {
                const changes = new ItemChangeType();
                changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();
                const fieldURIChange = new DeleteItemFieldType();
                fieldURIChange.FieldURI = new PathToUnindexedFieldType();
                fieldURIChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
                changes.ItemId = new ItemIdType(pimContact.unid, `ck-${pimContact.unid}`);
                changes.Updates.push(fieldURIChange);

                const change1 = new DeleteItemFieldType();
                change1.IndexedFieldURI = new PathToIndexedFieldType();
                change1.IndexedFieldURI.FieldURI = DictionaryURIType.CONTACTS_IM_ADDRESS;
                change1.IndexedFieldURI.FieldIndex = ImAddressKeyType.IM_ADDRESS_1;
                changes.Updates.push(change1);

                const change2 = new DeleteItemFieldType();
                change2.IndexedFieldURI = new PathToIndexedFieldType();
                change2.IndexedFieldURI.FieldURI = DictionaryURIType.CONTACTS_IM_ADDRESS;
                change2.IndexedFieldURI.FieldIndex = ImAddressKeyType.IM_ADDRESS_2;
                changes.Updates.push(change2);

                const change3 = new DeleteItemFieldType();
                change3.IndexedFieldURI = new PathToIndexedFieldType();
                change3.IndexedFieldURI.FieldURI = DictionaryURIType.CONTACTS_IM_ADDRESS;
                change3.IndexedFieldURI.FieldIndex = ImAddressKeyType.IM_ADDRESS_3;
                changes.Updates.push(change3);

                // First try with no existing im addresses
                let contact = await manager.updateItem(pimContact, changes, userInfo, context.request);
                expect(contact).to.not.be.undefined();
                expect(contact?.ItemId?.Id).to.be.equal(getEWSId(pimContact.unid));

                // Now with existing addresses
                pimContact.imAddresses = ['im1', 'im2', 'im3', 'im4', 'im5'];
                contact = await manager.updateItem(pimContact, changes, userInfo, context.request);
                expect(contact).to.not.be.undefined();
                expect(contact?.ItemId?.Id).to.be.equal(getEWSId(pimContact.unid));

            });

            it('test deleting email addresses', async () => {
                const changes = new ItemChangeType();
                changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

                const change1 = new DeleteItemFieldType();
                change1.IndexedFieldURI = new PathToIndexedFieldType();
                change1.IndexedFieldURI.FieldURI = DictionaryURIType.CONTACTS_EMAIL_ADDRESS;
                change1.IndexedFieldURI.FieldIndex = EmailAddressKeyType.EMAIL_ADDRESS_1;
                changes.Updates.push(change1);

                const change2 = new DeleteItemFieldType();
                change2.IndexedFieldURI = new PathToIndexedFieldType();
                change2.IndexedFieldURI.FieldURI = DictionaryURIType.CONTACTS_EMAIL_ADDRESS;
                change2.IndexedFieldURI.FieldIndex = EmailAddressKeyType.EMAIL_ADDRESS_2;
                changes.Updates.push(change2);

                const change3 = new DeleteItemFieldType();
                change3.IndexedFieldURI = new PathToIndexedFieldType();
                change3.IndexedFieldURI.FieldURI = DictionaryURIType.CONTACTS_EMAIL_ADDRESS;
                change3.IndexedFieldURI.FieldIndex = EmailAddressKeyType.EMAIL_ADDRESS_3;
                changes.Updates.push(change3);

                // First try with no existing email addresses
                let contact = await manager.updateItem(pimContact, changes, userInfo, context.request);
                expect(contact).to.not.be.undefined();
                expect(contact?.ItemId?.Id).to.be.equal(getEWSId(pimContact.unid));

                // Now with existing addresses
                pimContact.homeEmails = ['home1@home.com', 'home2@home.com'];
                pimContact.workEmails = ['work1@hcl.com', 'work2@hcl.com'];
                pimContact.mobileEmail = 'mobile@hcl.com';
                contact = await manager.updateItem(pimContact, changes, userInfo, context.request);
                expect(contact).to.not.be.undefined();
                expect(contact?.ItemId?.Id).to.be.equal(getEWSId(pimContact.unid));
            });

            it('test deleting phone numbers', async () => {
                const changes = new ItemChangeType();
                changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

                const change1 = new DeleteItemFieldType();
                change1.IndexedFieldURI = new PathToIndexedFieldType();
                change1.IndexedFieldURI.FieldURI = DictionaryURIType.CONTACTS_PHONE_NUMBER;
                change1.IndexedFieldURI.FieldIndex = PhoneNumberKeyType.BUSINESS_PHONE;
                changes.Updates.push(change1);

                const change2 = new DeleteItemFieldType();
                change2.IndexedFieldURI = new PathToIndexedFieldType();
                change2.IndexedFieldURI.FieldURI = DictionaryURIType.CONTACTS_PHONE_NUMBER;
                change2.IndexedFieldURI.FieldIndex = PhoneNumberKeyType.BUSINESS_FAX;
                changes.Updates.push(change2);

                const change3 = new DeleteItemFieldType();
                change3.IndexedFieldURI = new PathToIndexedFieldType();
                change3.IndexedFieldURI.FieldURI = DictionaryURIType.CONTACTS_PHONE_NUMBER;
                change3.IndexedFieldURI.FieldIndex = PhoneNumberKeyType.HOME_PHONE;
                changes.Updates.push(change3);

                const change4 = new DeleteItemFieldType();
                change4.IndexedFieldURI = new PathToIndexedFieldType();
                change4.IndexedFieldURI.FieldURI = DictionaryURIType.CONTACTS_PHONE_NUMBER;
                change4.IndexedFieldURI.FieldIndex = PhoneNumberKeyType.HOME_FAX;
                changes.Updates.push(change4);

                const change5 = new DeleteItemFieldType();
                change5.IndexedFieldURI = new PathToIndexedFieldType();
                change5.IndexedFieldURI.FieldURI = DictionaryURIType.CONTACTS_PHONE_NUMBER;
                change5.IndexedFieldURI.FieldIndex = PhoneNumberKeyType.MOBILE_PHONE;
                changes.Updates.push(change5);

                // not supported
                const change6 = new DeleteItemFieldType();
                change6.IndexedFieldURI = new PathToIndexedFieldType();
                change6.IndexedFieldURI.FieldURI = DictionaryURIType.CONTACTS_PHONE_NUMBER;
                change6.IndexedFieldURI.FieldIndex = PhoneNumberKeyType.BUSINESS_MOBILE;
                changes.Updates.push(change6);

                // First try with no existing email addresses
                let contact = await manager.updateItem(pimContact, changes, userInfo, context.request);
                expect(contact).to.not.be.undefined();
                expect(contact?.ItemId?.Id).to.be.equal(getEWSId(pimContact.unid));

                // Now with existing addresses
                pimContact.homePhone = '5551212';
                pimContact.homeFax = '5551212';
                pimContact.officePhone = '5551212';
                pimContact.officeFax = '5551212';
                pimContact.cellPhone = '5551212';

                contact = await manager.updateItem(pimContact, changes, userInfo, context.request);
                expect(contact).to.not.be.undefined();
                expect(contact?.ItemId?.Id).to.be.equal(getEWSId(pimContact.unid));
            });

            it('test deleting physical addresses', async () => {
                const changes = new ItemChangeType();
                changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

                const change1 = new DeleteItemFieldType();
                change1.IndexedFieldURI = new PathToIndexedFieldType();
                change1.IndexedFieldURI.FieldURI = DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_CITY;
                changes.Updates.push(change1);

                const change2 = new DeleteItemFieldType();
                change2.IndexedFieldURI = new PathToIndexedFieldType();
                change2.IndexedFieldURI.FieldURI = DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_COUNTRY_OR_REGION;
                changes.Updates.push(change2);

                const change3 = new DeleteItemFieldType();
                change3.IndexedFieldURI = new PathToIndexedFieldType();
                change3.IndexedFieldURI.FieldURI = DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_POSTAL_CODE;
                changes.Updates.push(change3);

                const change4 = new DeleteItemFieldType();
                change4.IndexedFieldURI = new PathToIndexedFieldType();
                change4.IndexedFieldURI.FieldURI = DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_STATE;
                changes.Updates.push(change4);

                const change5 = new DeleteItemFieldType();
                change5.IndexedFieldURI = new PathToIndexedFieldType();
                change5.IndexedFieldURI.FieldURI = DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_STREET;
                changes.Updates.push(change5);

                const contact = await manager.updateItem(pimContact, changes, userInfo, context.request);
                expect(contact).to.not.be.undefined();
                expect(contact?.ItemId?.Id).to.be.equal(getEWSId(pimContact.unid));
            });


            it('test different item types', async () => {
                const changes = new ItemChangeType();
                changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

                const fieldURIChange = new SetItemFieldType()
                fieldURIChange.FieldURI = new PathToUnindexedFieldType();
                fieldURIChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
                fieldURIChange.Message = new MessageType();
                fieldURIChange.Message.ItemId = new ItemIdType(pimContact.unid, `ck-${pimContact.unid}`);
                fieldURIChange.Message.Subject = 'UPDATED-SUBJECT';

                changes.ItemId = new ItemIdType(pimContact.unid, `ck-${pimContact.unid}`);
                changes.Updates.push(fieldURIChange);

                pimContact.parentFolderIds = undefined;

                // Stub label manager to return a contact folder
                const contactsLabel = PimItemFactory.newPimLabel({}, PimItemFormat.DOCUMENT);
                contactsLabel.displayName = 'Contacts';
                contactsLabel.view = KeepPimConstants.CONTACTS;
                contactsLabel.type = PimLabelTypes.CONTACTS;

                const labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
                labelManagerStub.getLabels.resolves([contactsLabel]);
                stubLabelsManager(labelManagerStub);

                const contact = await manager.updateItem(pimContact, changes, userInfo, context.request);
                expect(contact).to.not.be.undefined();
                expect(contact?.ItemId?.Id).to.be.equal(getEWSId(pimContact.unid));
            });

            it('test no contact labels', async () => {
                const changes = new ItemChangeType();
                changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

                const fieldURIChange = new SetItemFieldType()
                fieldURIChange.FieldURI = new PathToUnindexedFieldType();
                fieldURIChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
                fieldURIChange.Message = new MessageType();
                fieldURIChange.Message.ItemId = new ItemIdType(pimContact.unid, `ck-${pimContact.unid}`);
                fieldURIChange.Message.Subject = 'UPDATED-SUBJECT';

                changes.ItemId = new ItemIdType(pimContact.unid, `ck-${pimContact.unid}`);
                changes.Updates.push(fieldURIChange);

                pimContact.parentFolderIds = undefined;

                // Stub label manager to return a contact folder
                const notesLabel = PimItemFactory.newPimLabel({}, PimItemFormat.DOCUMENT);
                notesLabel.displayName = 'Test';
                notesLabel.view = KeepPimConstants.JOURNAL;
                notesLabel.type = PimLabelTypes.JOURNAL;

                const labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
                labelManagerStub.getLabels.resolves([notesLabel]);
                stubLabelsManager(labelManagerStub);

                const contact = await manager.updateItem(pimContact, changes, userInfo, context.request);
                expect(contact).to.not.be.undefined();
                expect(contact?.ItemId?.Id).to.be.equal(getEWSId(pimContact.unid));
            });
        });
    });

    describe('Test createItem', () => {

        // Clean up the stubs
        afterEach(() => {
            sinon.restore(); // Reset stubs
        });

        it('test create contact, no label', async () => {
            const contactsManagerStub = sinon.createStubInstance(KeepPimContactsManager);
            contactsManagerStub.createContact.resolves('UID');
            contactsManagerStub.deleteContact.resolves({});
            stubContactsManager(contactsManagerStub);

            const manager = new MockEWSManager();
            const userInfo = new UserContext();

            const contact = new ContactItemType();
            contact.ItemId = new ItemIdType('UID', 'ck-UID');

            const context = getContext(testUser, 'password');

            // No parent folder
            const newMessages = await manager.createItem(contact, userInfo, context.request);
            expect(newMessages).to.not.be.undefined();
            expect(newMessages.length).to.be.equal(1);
            expect(newMessages[0]?.ItemId?.Id).to.be.equal(getEWSId('UID'));
        });


        it('test create contact, with a label', async () => {
            const messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            messageManagerStub.moveMessages.resolves({movedIds: [{status: 200, message: 'success', unid: 'UID'}]});
            stubMessageManager(messageManagerStub);

            const contactsManagerStub = sinon.createStubInstance(KeepPimContactsManager);
            contactsManagerStub.createContact.resolves('UID');
            contactsManagerStub.deleteContact.resolves({});
            stubContactsManager(contactsManagerStub);

            const manager = new MockEWSManager();
            const userInfo = new UserContext();
            const toLabel = PimItemFactory.newPimLabel({}, PimItemFormat.DOCUMENT);
            toLabel.folderId = 'TO-LABEL';
            toLabel.type = PimLabelTypes.CONTACTS;

            const contact = new ContactItemType();
            contact.ItemId = new ItemIdType('UID', 'ck-UID');
            const context = getContext(testUser, 'password');

            const mailboxId = 'test@test.com';
            const itemEWSId = getEWSId('UID', mailboxId);

            const newMessages = await manager
                .createItem(contact, userInfo, context.request, MessageDispositionType.SAVE_ONLY, toLabel, mailboxId);
            expect(newMessages).to.not.be.undefined();
            expect(newMessages.length).to.be.equal(1);
            expect(newMessages[0]?.ItemId?.Id).to.be.equal(itemEWSId);
        });

        it('test create contact, with error thrown', async () => {
            const messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            messageManagerStub.moveMessages.throws();
            stubMessageManager(messageManagerStub);

            const contactsManagerStub = sinon.createStubInstance(KeepPimContactsManager);
            contactsManagerStub.createContact.resolves('UID');
            contactsManagerStub.deleteContact.resolves({});
            stubContactsManager(contactsManagerStub);

            const manager = new MockEWSManager();
            const userInfo = new UserContext();
            const toLabel = PimItemFactory.newPimLabel({}, PimItemFormat.DOCUMENT);
            toLabel.folderId = 'FOLDER-ID';
            toLabel.type = PimLabelTypes.CONTACTS;

            const contact = new ContactItemType();
            contact.ItemId = new ItemIdType('UID', 'ck-UID');
            const context = getContext(testUser, 'password');

            let success = false;
            try {
                await manager.createItem(contact, userInfo, context.request, MessageDispositionType.SAVE_ONLY, toLabel);
            } catch (err) {
                success = true;
            }
            expect(success).to.be.true();
        });
    });

    describe('Test deleteItem', () => {

        it('test delete', async () => {
            const contactsManagerStub = sinon.createStubInstance(KeepPimContactsManager);
            contactsManagerStub.createContact.resolves('UID');
            contactsManagerStub.deleteContact.resolves({});
            stubContactsManager(contactsManagerStub);

            const manager = EWSContactsManager.getInstance();
            const userInfo = new UserContext();
            await manager.deleteItem('item-id)', userInfo);
        });
    });

    describe('Test addRequestedPropertiesToEWSItem', () => {

        let pimContact: PimContact;
        let contact: ContactItemType;

        // Reset pimItem before each test
        beforeEach(function () {

            const extendedProp = {
                DistinguishedPropertySetId: 'Address',
                PropertyId: 32899,
                PropertyType: 'String',
                Value: 'kermit.frog@fakemail.com'
            }

            pimContact = PimItemFactory.newPimContact({}, PimItemFormat.DOCUMENT);
            pimContact.unid = 'UNID';
            pimContact.parentFolderIds = ['PARENT-FOLDER'];
            pimContact.subject = 'SUBJECT';
            pimContact.body = 'BODY';
            pimContact.addExtendedProperty(extendedProp);

            contact = new ContactItemType();
        });

        it("base fields for contact", function () {
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;

            const mockManager = new MockEWSManager();
            mockManager.addRequestedPropertiesToEWSItem(pimContact, contact, shape);
            expect(contact.ItemId?.Id).to.be.equal(getEWSId(pimContact.unid));
            expect(contact.ParentFolderId?.Id).to.be.equal(getEWSId('PARENT-FOLDER'));
        });
    });

    describe('Test updateEWSItemFieldValue', () => {

        let pimContact: PimContact;
        let contact: ContactItemType;
        let manager: MockEWSManager;

        beforeEach(() => {
            pimContact = PimItemFactory.newPimContact({}, PimItemFormat.DOCUMENT);
            contact = new ContactItemType();
            manager = new MockEWSManager();
        });

        it("Test unhandled items", function () {
            const unhandledFields = [
                UnindexedFieldURIType.CONTACTS_FILE_AS,
                UnindexedFieldURIType.CONTACTS_ALIAS,
                UnindexedFieldURIType.CONTACTS_CHILDREN,
                UnindexedFieldURIType.CONTACTS_CONTACT_SOURCE,
                UnindexedFieldURIType.ITEM_CONVERSATION_ID,
                UnindexedFieldURIType.CONTACTS_CULTURE,
                UnindexedFieldURIType.ITEM_DATE_TIME_RECEIVED,
                UnindexedFieldURIType.ITEM_DATE_TIME_SENT,
                UnindexedFieldURIType.CONTACTS_DIRECTORY_ID,
                UnindexedFieldURIType.CONTACTS_DIRECT_REPORTS,
                UnindexedFieldURIType.ITEM_DISPLAY_CC,
                UnindexedFieldURIType.CONTACTS_DISPLAY_NAME,
                UnindexedFieldURIType.ITEM_DISPLAY_TO,
                UnindexedFieldURIType.CONTACTS_FILE_AS_MAPPING,
                UnindexedFieldURIType.CONTACTS_GENERATION,
                UnindexedFieldURIType.ITEM_IMPORTANCE,
                UnindexedFieldURIType.CONTACTS_INITIALS,
                UnindexedFieldURIType.ITEM_INTERNET_MESSAGE_HEADERS,
                UnindexedFieldURIType.ITEM_IS_ASSOCIATED,
                UnindexedFieldURIType.ITEM_IS_DRAFT,
                UnindexedFieldURIType.ITEM_IS_FROM_ME,
                UnindexedFieldURIType.ITEM_IS_RESEND,
                UnindexedFieldURIType.ITEM_IS_SUBMITTED,
                UnindexedFieldURIType.ITEM_IS_UNMODIFIED,
                UnindexedFieldURIType.ITEM_LAST_MODIFIED_NAME,
                UnindexedFieldURIType.CONTACTS_MIDDLE_NAME,
                UnindexedFieldURIType.CONTACTS_MILEAGE,
                UnindexedFieldURIType.CONTACTS_NICKNAME,
                UnindexedFieldURIType.CONTACTS_NOTES,
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
                UnindexedFieldURIType.ITEM_SIZE,
                UnindexedFieldURIType.ITEM_WEB_CLIENT_EDIT_FORM_QUERY_STRING,
                UnindexedFieldURIType.ITEM_WEB_CLIENT_READ_FORM_QUERY_STRING
            ];

            unhandledFields.forEach(field => {
                const handled = manager.updateEWSItemFieldValue(contact, pimContact, field);
                console.log(field);
                expect(handled).to.be.false();
            });
        });


        it("UnindexedFieldURIType.CONTACTS_COMPANY_NAME", function () {
            // Not set
            let handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_COMPANY_NAME);
            expect(handled).to.be.true();
            expect(contact.CompanyName).to.be.undefined();

            pimContact.companyName = 'HCL';
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_COMPANY_NAME);
            expect(handled).to.be.true();
            expect(contact.CompanyName).to.be.eql(pimContact.companyName);
        });

        it("UnindexedFieldURIType.CONTACTS_COMPLETE_NAME", function () {
            // Not set
            let handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_COMPLETE_NAME);
            expect(handled).to.be.true();
            expect(contact.CompleteName).to.not.be.undefined();
            expect(contact.CompleteName?.FirstName).to.be.undefined();
            expect(contact.CompleteName?.LastName).to.be.undefined();
            expect(contact.CompleteName?.Title).to.be.undefined();
            expect(contact.CompleteName?.MiddleName).to.be.undefined();
            expect(contact.CompleteName?.Suffix).to.be.undefined();

            pimContact.firstName = 'First';
            pimContact.lastName = 'Last';
            pimContact.middleInitial = 'Middle';
            pimContact.suffix = 'Jr';
            pimContact.title = 'Dr';

            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_COMPLETE_NAME);
            expect(handled).to.be.true();
            expect(contact.CompleteName?.FirstName).to.be.eql(pimContact.firstName);
            expect(contact.CompleteName?.LastName).to.be.eql(pimContact.lastName);
            expect(contact.CompleteName?.MiddleName).to.be.eql(pimContact.middleInitial);
            expect(contact.CompleteName?.Title).to.be.eql(pimContact.title);
            expect(contact.CompleteName?.Suffix).to.be.eql(pimContact.suffix);
        });

        it("UnindexedFieldURIType.CONTACTS_EMAIL_ADDRESSES", function () {
            // Not set
            let handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_EMAIL_ADDRESSES);
            expect(handled).to.be.true();
            expect(contact.EmailAddresses).to.be.undefined();

            pimContact.workEmails = ['work@hcl.com', 'work2@hcl.com'];
            pimContact.mobileEmail = 'mobile@hcl.com';
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_EMAIL_ADDRESSES);
            expect(handled).to.be.true();
            expect(contact.EmailAddresses).to.not.be.undefined();
            expect(contact.EmailAddresses?.Entry).to.be.an.Array();
            expect(contact.EmailAddresses?.Entry?.length).to.be.equal(3);
        });

        it("UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS", function () {
            // Not set
            let handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS);
            expect(handled).to.be.true();
            expect(contact.HasAttachments).to.be.false();

            pimContact.attachments = ['ContactPhoto'];
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS);
            expect(handled).to.be.true();
            expect(contact.HasAttachments).to.be.false();

            pimContact.attachments = ['ContactPhoto', 'Att1'];
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS);
            expect(handled).to.be.true();
            expect(contact.HasAttachments).to.be.true();

            pimContact.attachments = ['Att1'];
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS);
            expect(handled).to.be.true();
            expect(contact.HasAttachments).to.be.true();

        });

        it("UnindexedFieldURIType.ITEM_ATTACHMENTS", function () {
            // Not set
            let handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.ITEM_ATTACHMENTS);
            expect(handled).to.be.true();
            expect(contact.Attachments).to.be.undefined();

            // FIXME:  We return false for HasAttachments if there's only a contact photo, but we include the contact photo in the array of attachments.  Should we?
            pimContact.attachments = ['ContactPhoto'];
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.ITEM_ATTACHMENTS);
            expect(handled).to.be.true();
            expect(contact.Attachments).to.not.be.undefined();

            pimContact.attachments = ['ContactPhoto', 'Att1'];
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.ITEM_ATTACHMENTS);
            expect(handled).to.be.true();
            expect(contact.Attachments).to.not.be.undefined();
            expect(contact.Attachments?.items).to.be.an.Array();
            expect(contact.Attachments?.items?.length).to.be.equal(2);
        });

        it("UnindexedFieldURIType.CONTACTS_IM_ADDRESSES", function () {
            // Not set
            let handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_IM_ADDRESSES);
            expect(handled).to.be.true();
            expect(contact.ImAddresses).to.be.undefined();

            pimContact.imAddresses = ['im1', 'im2'];
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_IM_ADDRESSES);
            expect(handled).to.be.true();
            expect(contact.ImAddresses).to.not.be.undefined();
            expect(contact.ImAddresses?.Entry).to.be.an.Array();
            expect(contact.ImAddresses?.Entry?.length).to.be.equal(2);
        });

        it("UnindexedFieldURIType.CONTACTS_IM_ADDRESSES when set", function () {
            // set
            let handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_IM_ADDRESSES);

            pimContact.imAddresses = ['im1'];
            const addresses: any = {};
            addresses[ImAddressKeyType.IM_ADDRESS_1] = 'im1';
            addresses[ImAddressKeyType.IM_ADDRESS_2] = 'im2';
            pimContact.setAdditionalProperty(UnindexedFieldURIType.CONTACTS_IM_ADDRESSES, addresses);

            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_IM_ADDRESSES);
            expect(handled).to.be.true();
            expect(contact.ImAddresses).to.not.be.undefined();
            expect(contact.ImAddresses?.Entry).to.be.an.Array();
            expect(contact.ImAddresses?.Entry?.length).to.be.equal(2);
        });

        it("UnindexedFieldURIType.CONTACTS_JOB_TITLE", function () {
            // Not set
            let handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_JOB_TITLE);
            expect(handled).to.be.true();
            expect(contact.JobTitle).to.be.undefined();

            pimContact.jobTitle = 'Developer';
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_JOB_TITLE);
            expect(handled).to.be.true();
            expect(contact.JobTitle).to.be.eql(pimContact.jobTitle);
        });

        it("UnindexedFieldURIType.CONTACTS_PHONE_NUMBERS", function () {
            // Not set
            let handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_PHONE_NUMBERS);
            expect(handled).to.be.true();
            expect(contact.PhoneNumbers).to.be.undefined();

            pimContact.homePhone = '919551212';
            pimContact.officeFax = '4444444';
            pimContact.officePhone = '99999999';
            pimContact.cellPhone = '777777';
            pimContact.homeFax = '111111';
            pimContact.otherPhones = ['1234', '4321'];
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_PHONE_NUMBERS);
            expect(handled).to.be.true();
            expect(contact.PhoneNumbers).to.not.be.undefined();
            expect(contact.PhoneNumbers?.Entry).to.be.an.Array();
            expect(contact.PhoneNumbers?.Entry?.length).to.be.equal(7);
        });

        it("UnindexedFieldURIType.CONTACTS_PHYSICAL_ADDRESSES", function () {
            // Not set
            let handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_PHYSICAL_ADDRESSES);
            expect(handled).to.be.true();
            expect(contact.PhysicalAddresses).to.be.undefined();

            const office = new PimAddress();
            office.Street = '1313 Mockingbird Ln';
            pimContact.officeAddress = office;
            const home = new PimAddress();
            home.Street = '1600 Pennsylvania Ave';
            pimContact.homeAddress = home;
            const other = new PimAddress();
            other.Street = '10 Downing Street';
            pimContact.otherAddress = other;

            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_PHYSICAL_ADDRESSES);
            expect(handled).to.be.true();
            expect(contact.PhysicalAddresses).to.not.be.undefined();
            expect(contact.PhysicalAddresses?.Entry).to.not.be.undefined();
            expect(contact.PhysicalAddresses?.Entry?.length).to.be.equal(3);

        });

        it("UnindexedFieldURIType.ITEM_BODY", function () {
            // Not set
            let handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.ITEM_BODY);
            expect(handled).to.be.true();
            expect(contact.Body).to.not.be.undefined();
            expect(contact.Body?.Value).to.be.equal(pimContact.body);

            pimContact.body = 'BODY';
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.ITEM_BODY);
            expect(handled).to.be.true();
            expect(contact.Body).to.not.be.undefined();
            expect(contact.Body?.Value).to.be.equal(pimContact.body);
        });

        it("UnindexedFieldURIType.CONTACTS_BUSINESS_HOME_PAGE", function () {
            // Not set
            let handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_BUSINESS_HOME_PAGE);
            expect(handled).to.be.true();
            expect(contact.BusinessHomePage).to.be.undefined();

            pimContact.homepage = 'www.hcl.com';
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_BUSINESS_HOME_PAGE);
            expect(handled).to.be.true();
            expect(contact.BusinessHomePage).to.be.eql(pimContact.homepage);
        });

        it("UnindexedFieldURIType.ITEM_DATE_TIME_CREATED", function () {
            // Not set
            let handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.ITEM_DATE_TIME_CREATED);
            expect(handled).to.be.true();
            // When LABS-1863 is resolved the next line can be removed and the 
            // expect to be undefined can be uncommented.
            expect(contact.DateTimeCreated).to.be.eql(pimContact.createdDate);
            // expect(contact.DateTimeCreated).to.be.undefined();

            pimContact.createdDate = new Date();
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.ITEM_DATE_TIME_CREATED);
            expect(handled).to.be.true();
            expect(contact.DateTimeCreated).to.be.eql(pimContact.createdDate);
        });

        it("UnindexedFieldURIType.CONTACTS_DEPARTMENT", function () {
            // Not set
            let handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_DEPARTMENT);
            expect(handled).to.be.true();
            expect(contact.Department).to.be.undefined();

            pimContact.department = 'LNMA';
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_DEPARTMENT);
            expect(handled).to.be.true();
            expect(contact.Department).to.be.eql(pimContact.department);
        });

        it("UnindexedFieldURIType.CONTACTS_GIVEN_NAME", function () {
            // Not set
            let handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_GIVEN_NAME);
            expect(handled).to.be.true();
            expect(contact.GivenName).to.be.equal(pimContact.firstName);

            contact.GivenName = undefined;
            pimContact.firstName = undefined;
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_GIVEN_NAME);
            expect(handled).to.be.true();
            expect(contact.GivenName).to.be.undefined();
        });

        it("UnindexedFieldURIType.CONTACTS_HAS_PICTURE", function () {
            // Not set
            let handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_HAS_PICTURE);
            expect(handled).to.be.true();
            expect(contact.HasPicture).to.be.false();

            pimContact.attachments = ['aaa', 'ContactPhoto'];
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_HAS_PICTURE);
            expect(handled).to.be.true();
            expect(contact.HasPicture).to.be.true();
        });

        it("UnindexedFieldURIType.CONTACTS_IS_PRIVATE", function () {
            // Not set
            let handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_IS_PRIVATE);
            expect(handled).to.be.true();
            expect(contact.IsPrivate).to.be.false();

            pimContact.isPrivate = true;
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_IS_PRIVATE);
            expect(handled).to.be.true();
            expect(contact.IsPrivate).to.be.true();

            pimContact.isPrivate = false;
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_IS_PRIVATE);
            expect(handled).to.be.true();
            expect(contact.IsPrivate).to.be.false();
        });

        it("UnindexedFieldURIType.CONTACTS_BIRTHDAY", function () {
            // Not set
            let handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_BIRTHDAY);
            expect(handled).to.be.true();
            expect(contact.Birthday).to.be.undefined();

            pimContact.birthday = new Date();
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_BIRTHDAY);
            expect(handled).to.be.true();
            expect(contact.Birthday).to.be.eql(pimContact.birthday);
        });

        it("UnindexedFieldURIType.CONTACTS_WEDDING_ANNIVERSARY", function () {
            // Not set
            let handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_WEDDING_ANNIVERSARY);
            expect(handled).to.be.true();
            expect(contact.WeddingAnniversary).to.be.undefined();

            pimContact.anniversary = new Date();
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_WEDDING_ANNIVERSARY);
            expect(handled).to.be.true();
            expect(contact.WeddingAnniversary).to.be.eql(pimContact.anniversary);
        });

        it("UnindexedFieldURIType.CONTACTS_OFFICE_LOCATION", function () {
            // Not set
            let handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_OFFICE_LOCATION);
            expect(handled).to.be.true();
            expect(contact.OfficeLocation).to.be.undefined();

            pimContact.location = 'Earth';
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_OFFICE_LOCATION);
            expect(handled).to.be.true();
            expect(contact.OfficeLocation).to.be.eql(pimContact.location);
        });

        it("UnindexedFieldURIType.ITEM_SENSITIVITY", function () {
            // Not set
            let handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.ITEM_SENSITIVITY);
            expect(handled).to.be.true();
            expect(contact.Sensitivity).to.be.equal(SensitivityChoicesType.NORMAL);

            pimContact.isConfidential = true;
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.ITEM_SENSITIVITY);
            expect(handled).to.be.true();
            expect(contact.Sensitivity).to.be.eql(SensitivityChoicesType.CONFIDENTIAL);
        });

        it("UnindexedFieldURIType.CONTACTS_SURNAME", function () {
            // Not set
            let handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_SURNAME);
            expect(handled).to.be.true();
            expect(contact.Surname).to.be.equal(pimContact.lastName);

            contact.Surname = undefined;
            pimContact.lastName = undefined;
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.CONTACTS_SURNAME);
            expect(handled).to.be.true();
            expect(contact.Surname).to.be.undefined();
        });
    });

    describe('Test pimItemToEWSItem', () => {
        const mailboxId = 'test@test.com';
        const itemEWSId = getEWSId('UNID', mailboxId);
        const parentEWSId = getEWSId('PARENT-FOLDER-ID', mailboxId);

        // Clean up the stubs
        afterEach(() => {
            sinon.restore(); // Reset stubs
        });

        it('simple contact, no parent folder', async () => {
            const pimItem = PimItemFactory.newPimContact({}, PimItemFormat.DOCUMENT);
            pimItem.unid = 'UNID';
            pimItem.firstName = 'First';
            pimItem.lastName = 'Last';

            const manager = new MockEWSManager();
            const userInfo = new UserContext();
            const context = getContext(testUser, 'password');
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            shape.IncludeMimeContent = true;

            const contact = await manager.pimItemToEWSItem(pimItem, userInfo, context.request, shape, mailboxId);
            expect(contact.ItemId?.Id).to.be.equal(itemEWSId);
        });

        it('simple contact, with parent folder', async () => {
            const pimItem = PimItemFactory.newPimContact({}, PimItemFormat.DOCUMENT);
            pimItem.unid = 'UNID';
            pimItem.firstName = 'First';
            pimItem.lastName = 'Last';

            const manager = new MockEWSManager();
            const userInfo = new UserContext();
            const context = getContext(testUser, 'password');
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            shape.IncludeMimeContent = true;

            const contact = await manager.pimItemToEWSItem(pimItem, userInfo, context.request, shape, mailboxId, 'PARENT-FOLDER-ID');
            expect(contact.ItemId?.Id).to.be.equal(itemEWSId);
            expect(contact.ParentFolderId?.Id).to.be.equal(parentEWSId);
        });
    });

    describe('EWSContactsManager.pimItemToEWSItem with shape', function () {
        const userInfo = new UserContext();
        const context = getContext("testID","testPass");

        const pimObject = {
                '@unid': 'testID',
                'ParentFolder':'parentID',
                'Comment':'testComment',
                'Suffix':'Mr',
                'Title':'Dr',
                'FirstName':'fname',
                'LastName':'lname',
                'MiddleInitial':'mid',
                'CompanyName':'HCL',
                "Confidential": "1",
                'Department':'IT',
                'Location':'Bangalore',
                'Birthday':'1991-12-25T11:00:00.000Z',
                'Anniversary':'2020-12-01T11:00:00.000Z'
            }
    
        it('for shape IdOnly', async function () {
            const pimContact = PimItemFactory.newPimContact(pimObject, PimItemFormat.DOCUMENT)
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            const ews = await EWSContactsManager.getInstance().pimItemToEWSItem(pimContact, userInfo, context.request, shape);
    
            expect(ews?.ItemId?.Id).to.be.equal(getEWSId('testID'));
            expect(ews?.ParentFolderId?.Id).to.be.equal(getEWSId('parentID'));
        });
    
        it('for shape AllProperties', async function () {
            const pimContact = PimItemFactory.newPimContact(pimObject, PimItemFormat.DOCUMENT)
            
            pimContact.jobTitle='Software Engineer';
            
            pimContact.primaryEmail = 'primaryemail@mytest.com';
            pimContact.homeEmails = ['homeemail@mytest.com'];
            pimContact.workEmails = ['workemail@mytest.com'];
    
            // TODO: Uncomment when LABS-1155 is implemented
            // pimContact.primaryPhone = '9898786767';
            pimContact.officePhone = '9878437438';
            pimContact.homePhone = '9898439538';
            pimContact.cellPhone = '89328429428';
            pimContact.otherPhones = ['91093283208'];
            pimContact.officeFax = '011200031212121';
            pimContact.homeFax = '01211129438943';
    
            let pimAddress = new PimAddress(); 
            pimAddress.Street = "Office Street";
            pimAddress.City = "Office City";
            pimAddress.State = "Office State";
            pimAddress.PostalCode = "Office Code";
            pimAddress.Country = "Office Country";
            pimContact.officeAddress = pimAddress;
    
            pimAddress = new PimAddress(); 
            pimAddress.Street = "Home Street";
            pimAddress.City = "Home City";
            pimAddress.State = "Home State";
            pimAddress.PostalCode = "Home Code";
            pimAddress.Country = "Home Country";
            pimContact.homeAddress = pimAddress; 
    
            pimAddress = new PimAddress(); 
            pimAddress.Street = "Other Street";
            pimAddress.City = "Other City";
            pimAddress.State = "Other State";
            pimAddress.PostalCode = "Other Code";
            pimAddress.Country = "Other Country";
            pimContact.otherAddress = pimAddress; 
    
            pimContact.categories = ['ews','keep']
    
            pimContact.homepage = 'myBusinesspage.com';
            pimContact.attachments = ['ContactPicture.jpg'];
    
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
            const ews = await EWSContactsManager.getInstance().pimItemToEWSItem(pimContact, userInfo, context.request, shape);
    
            expect(ews?.Body?.Value).to.be.equal(pimObject.Comment);
            expect(ews?.CompleteName?.Suffix).to.be.equal(pimObject.Suffix);
            expect(ews?.CompleteName?.Title).to.be.equal(pimObject.Title);
            expect(ews?.CompleteName?.FirstName).to.be.equal(pimObject.FirstName);
            expect(ews?.CompleteName?.LastName).to.be.equal(pimObject.LastName);
            expect(ews?.CompleteName?.MiddleName).to.be.equal(pimObject.MiddleInitial);
            expect(ews?.JobTitle).to.be.equal(pimContact.jobTitle);
            expect(ews?.CompanyName).to.be.equal(pimObject.CompanyName);
    
           expect(ews?.Sensitivity).to.be.equal(SensitivityChoicesType.CONFIDENTIAL);
    
           expect(ews.EmailAddresses).to.not.be.undefined(); 

           ews.EmailAddresses!.Entry.forEach(emailaddress => {
                if(emailaddress.Key === EmailAddressKeyType.EMAIL_ADDRESS_1){
                    expect(emailaddress.Value).to.be.equal(pimContact.primaryEmail);
                }else if(emailaddress.Key === EmailAddressKeyType.EMAIL_ADDRESS_2){
                    expect(emailaddress.Value).to.be.equal(pimContact.workEmails[0]);
                }else if(emailaddress.Key === EmailAddressKeyType.EMAIL_ADDRESS_3){
                    expect(emailaddress.Value).to.be.equal(pimContact.homeEmails[0]);
                }
           });
           
            expect(ews.PhoneNumbers).to.not.be.undefined(); 
            ews.PhoneNumbers!.Entry.forEach(phonenumber => {
                // TODO: Uncomment when LABS-1155 is implemented
               // if(phonenumber.Key == PhoneNumberKeyType.PRIMARY_PHONE){
                 //   expect(phonenumber.Value).to.be.equal(pimContact.primaryPhone);
                //}else 
                if(phonenumber.Key === PhoneNumberKeyType.BUSINESS_PHONE){
                    expect(phonenumber.Value).to.be.equal(pimContact.officePhone);
                }else if(phonenumber.Key === PhoneNumberKeyType.HOME_PHONE){
                    expect(phonenumber.Value).to.be.equal(pimContact.homePhone);
                }else if(phonenumber.Key === PhoneNumberKeyType.MOBILE_PHONE){
                    expect(phonenumber.Value).to.be.equal(pimContact.cellPhone);
                }else if(phonenumber.Key === PhoneNumberKeyType.BUSINESS_FAX){
                    expect(phonenumber.Value).to.be.equal(pimContact.officeFax);
                }else if(phonenumber.Key === PhoneNumberKeyType.HOME_FAX){
                    expect(phonenumber.Value).to.be.equal(pimContact.homeFax);
                }else if(phonenumber.Key === PhoneNumberKeyType.OTHER_TELEPHONE){
                    expect(phonenumber.Value).to.be.equal(pimContact.otherPhones[0]);
                }
            });   
           
           
            expect(ews.PhysicalAddresses).to.not.be.undefined(); 
            ews.PhysicalAddresses!.Entry.forEach(address => {
                if(address.Key === PhysicalAddressKeyType.BUSINESS){
                    expect(address.Street).to.be.equal('Office Street');
                    expect(address.State).to.be.equal('Office State');
                    expect(address.City).to.be.equal('Office City');
                    expect(address.PostalCode).to.be.equal('Office Code');
                    expect(address.CountryOrRegion).to.be.equal('Office Country');
                } 
                else if(address.Key === PhysicalAddressKeyType.HOME){
                    expect(address.Street).to.be.equal('Home Street');
                    expect(address.State).to.be.equal('Home State');
                    expect(address.City).to.be.equal('Home City');
                    expect(address.PostalCode).to.be.equal('Home Code');
                    expect(address.CountryOrRegion).to.be.equal('Home Country');
                }
                else if(address.Key === PhysicalAddressKeyType.OTHER){
                    expect(address.Street).to.be.equal('Other Street');
                    expect(address.State).to.be.equal('Other State');
                    expect(address.City).to.be.equal('Other City');
                    expect(address.PostalCode).to.be.equal('Other Code');
                    expect(address.CountryOrRegion).to.be.equal('Other Country');
                }
            });
    
            expect(ews?.Department).to.be.equal(pimObject.Department);
            expect(ews?.OfficeLocation).to.be.equal(pimObject.Location);
            expect(ews.Birthday!.getTime() - new Date("1991-12-25T11:00:00.000Z").getTime()).to.equal(0);
            expect(ews.WeddingAnniversary!.getTime() - new Date("2020-12-01T11:00:00.000Z").getTime()).to.equal(0);    
            expect(ews?.BusinessHomePage).to.be.equal(pimContact.homepage);
            expect(ews?.HasAttachments).to.be.equal(true);
            expect(ews?.Attachments?.items.length).to.be.equal(1);
            expect(ews?.Categories?.String).to.containDeep(pimContact.categories);
    
        });

        it('for shape with additional properties', async function () {
            const pimContact = PimItemFactory.newPimContact(pimObject, PimItemFormat.DOCUMENT)
            
            pimContact.jobTitle='Software Engineer';
            
            pimContact.primaryEmail = 'workemail@mytest.com';
            pimContact.homeEmails = ['homeemail@mytest.com'];
            pimContact.workEmails = ['workemail@mytest.com'];
            pimContact.otherEmails = ['otheremail@mytest.com'];

            pimContact.imAddresses = ['imaddress@mytest.com'];
            const imAddresses: any = {};
            imAddresses[ImAddressKeyType.IM_ADDRESS_1] = 'imaddress@mytest.com';
            pimContact.setAdditionalProperty(UnindexedFieldURIType.CONTACTS_IM_ADDRESSES, imAddresses);
    
            // TODO: Uncomment when LABS-1155 is implemented
            // pimContact.primaryPhone = '9898786767';
            pimContact.officePhone = '9878437438';
            pimContact.homePhone = '9898439538';
            pimContact.cellPhone = '89328429428';
            pimContact.otherPhones = ['91093283208'];
            pimContact.officeFax = '011200031212121';
            pimContact.homeFax = '01211129438943';
    
            let pimAddress = new PimAddress(); 
            pimAddress.Street = "Office Street";
            pimAddress.City = "Office City";
            pimAddress.State = "Office State";
            pimAddress.PostalCode = "Office Code";
            pimAddress.Country = "Office Country";
            pimContact.officeAddress = pimAddress;
    
            pimAddress = new PimAddress(); 
            pimAddress.Street = "Home Street";
            pimAddress.City = "Home City";
            pimAddress.State = "Home State";
            pimAddress.PostalCode = "Home Code";
            pimAddress.Country = "Home Country";
            pimContact.homeAddress = pimAddress; 
    
            pimAddress = new PimAddress(); 
            pimAddress.Street = "Other Street";
            pimAddress.City = "Other City";
            pimAddress.State = "Other State";
            pimAddress.PostalCode = "Other Code";
            pimAddress.Country = "Other Country";
            pimContact.otherAddress = pimAddress; 
    
            pimContact.categories = ['ews','keep']
    
            pimContact.homepage = 'myBusinesspage.com';
            pimContact.attachments = ['ContactPicture.jpg'];
    
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            shape.AdditionalProperties = new NonEmptyArrayOfPathsToElementType();

            const props: any = {};
            props[DictionaryURIType.CONTACTS_EMAIL_ADDRESS] = [
                FieldIndexValue.CONTACT_EMAIL_ADDRESS1, 
                FieldIndexValue.CONTACT_EMAIL_ADDRESS2,
                FieldIndexValue.CONTACT_EMAIL_ADDRESS3
            ];
            props[DictionaryURIType.CONTACTS_IM_ADDRESS] = [
                FieldIndexValue.CONTACT_IM_ADDRESS1, 
                FieldIndexValue.CONTACT_IM_ADDRESS2,
                FieldIndexValue.CONTACT_IM_ADDRESS3
            ];
            props[DictionaryURIType.CONTACTS_PHONE_NUMBER] = [
                FieldIndexValue.CONTACT_PHONE_NUMBER_BUSINESS,
                FieldIndexValue.CONTACT_PHONE_NUMBER_BUSINESS_FAX,
                FieldIndexValue.CONTACT_PHONE_NUMBER_HOME,
                FieldIndexValue.CONTACT_PHONE_NUMBER_HOME_FAX,
                FieldIndexValue.CONTACT_PHONE_NUMBER_MOBILE,
                FieldIndexValue.CONTACT_PHONE_NUMBER_OTHER
            ];
            props[DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_STREET] = [
                FieldIndexValue.CONTACT_PHYSICAL_ADDRESS_BUSINESS,
                FieldIndexValue.CONTACT_PHYSICAL_ADDRESS_HOME,
                FieldIndexValue.CONTACT_PHYSICAL_ADDRESS_OTHER
            ];
            props[DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_CITY] = [
                FieldIndexValue.CONTACT_PHYSICAL_ADDRESS_BUSINESS,
                FieldIndexValue.CONTACT_PHYSICAL_ADDRESS_HOME,
                FieldIndexValue.CONTACT_PHYSICAL_ADDRESS_OTHER
            ];
            props[DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_STATE] = [
                FieldIndexValue.CONTACT_PHYSICAL_ADDRESS_BUSINESS,
                FieldIndexValue.CONTACT_PHYSICAL_ADDRESS_HOME,
                FieldIndexValue.CONTACT_PHYSICAL_ADDRESS_OTHER
            ];
            props[DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_POSTAL_CODE] = [
                FieldIndexValue.CONTACT_PHYSICAL_ADDRESS_BUSINESS,
                FieldIndexValue.CONTACT_PHYSICAL_ADDRESS_HOME,
                FieldIndexValue.CONTACT_PHYSICAL_ADDRESS_OTHER
            ];
            props[DictionaryURIType.CONTACTS_PHYSICAL_ADDRESS_COUNTRY_OR_REGION] = [
                FieldIndexValue.CONTACT_PHYSICAL_ADDRESS_BUSINESS,
                FieldIndexValue.CONTACT_PHYSICAL_ADDRESS_HOME,
                FieldIndexValue.CONTACT_PHYSICAL_ADDRESS_OTHER
            ];

            Object.keys(props).forEach(key => {
                const values: FieldIndexValue[] = props[key];
                values.forEach(fieldIndex => {
                    const indexedProp = new PathToIndexedFieldType();
                    indexedProp.FieldURI = key as DictionaryURIType; 
                    indexedProp.FieldIndex = fieldIndex;
                    shape.AdditionalProperties!.push(indexedProp);
                });
            });

            const email1OrgDisplayName = new PathToExtendedFieldType();
            email1OrgDisplayName.DistinguishedPropertySetId = DistinguishedPropertySetType.ADDRESS;
            email1OrgDisplayName.PropertyId = MapiPropertyIds.EMAIL1_ORIG_DISPLAY_NAME;
            email1OrgDisplayName.PropertyType = MapiPropertyTypeType.STRING; 
            shape.AdditionalProperties!.push(email1OrgDisplayName);
            const email2OrgDisplayName = new PathToExtendedFieldType();
            email2OrgDisplayName.DistinguishedPropertySetId = DistinguishedPropertySetType.ADDRESS;
            email2OrgDisplayName.PropertyId = MapiPropertyIds.EMAIL2_ORIG_DISPLAY_NAME;
            email2OrgDisplayName.PropertyType = MapiPropertyTypeType.STRING; 
            shape.AdditionalProperties!.push(email2OrgDisplayName);
            const email3OrgDisplayName = new PathToExtendedFieldType();
            email3OrgDisplayName.DistinguishedPropertySetId = DistinguishedPropertySetType.ADDRESS;
            email3OrgDisplayName.PropertyId = MapiPropertyIds.EMAIL3_ORIG_DISPLAY_NAME;
            email3OrgDisplayName.PropertyType = MapiPropertyTypeType.STRING; 
            shape.AdditionalProperties!.push(email3OrgDisplayName);
            const email1Label = new PathToExtendedFieldType();
            email1Label.DistinguishedPropertySetId = DistinguishedPropertySetType.PUBLIC_STRINGS;
            email1Label.PropertyName = "http://schemas.microsoft.com/entourage/emaillabel1";
            email1Label.PropertyType = MapiPropertyTypeType.STRING; 
            shape.AdditionalProperties!.push(email1Label);
            const email2Label = new PathToExtendedFieldType();
            email2Label.DistinguishedPropertySetId = DistinguishedPropertySetType.PUBLIC_STRINGS;
            email2Label.PropertyName = "http://schemas.microsoft.com/entourage/emaillabel2";
            email2Label.PropertyType = MapiPropertyTypeType.STRING; 
            shape.AdditionalProperties!.push(email2Label);
            const email3Label = new PathToExtendedFieldType();
            email3Label.DistinguishedPropertySetId = DistinguishedPropertySetType.PUBLIC_STRINGS;
            email3Label.PropertyName = "http://schemas.microsoft.com/entourage/emaillabel3";
            email3Label.PropertyType = MapiPropertyTypeType.STRING; 
            shape.AdditionalProperties!.push(email3Label);

            const ews = await EWSContactsManager.getInstance().pimItemToEWSItem(pimContact, userInfo, context.request, shape);

            expect(ews.EmailAddresses).to.not.be.undefined(); 
            ews.EmailAddresses!.Entry.forEach(emailaddress => {
                if(emailaddress.Key === EmailAddressKeyType.EMAIL_ADDRESS_1){
                    expect(emailaddress.Value).to.be.equal(pimContact.primaryEmail);
                    expect(emailaddress.Name).to.be.equal(pimContact.primaryEmail);
                }else if(emailaddress.Key === EmailAddressKeyType.EMAIL_ADDRESS_2){
                    expect(emailaddress.Value).to.be.equal(pimContact.homeEmails[0]);
                    expect(emailaddress.Name).to.be.equal(pimContact.homeEmails[0]);
                }else if(emailaddress.Key === EmailAddressKeyType.EMAIL_ADDRESS_3){ 
                    expect(emailaddress.Value).to.be.equal(pimContact.otherEmails[0]);
                    expect(emailaddress.Name).to.be.equal(pimContact.otherEmails[0]);
                }
                expect(emailaddress.MailboxType).to.be.equal(MailboxTypeType.CONTACT);
                expect(emailaddress.RoutingType).to.be.equal('SMTP');
           });
           
            expect(ews.ImAddresses).to.not.be.undefined(); 
            ews.ImAddresses!.Entry.forEach(imAddress => {
                if(imAddress.Key === ImAddressKeyType.IM_ADDRESS_1){
                    expect(pimContact.imAddresses.length).to.be.greaterThan(0);
                    expect(imAddress.Value).to.be.equal(pimContact.imAddresses[0]);
                }else if(imAddress.Key === ImAddressKeyType.IM_ADDRESS_2){
                    expect(pimContact.imAddresses.length).to.be.greaterThan(1);
                    expect(imAddress.Value).to.be.equal(pimContact.imAddresses[1]);
                }else if(imAddress.Key === ImAddressKeyType.IM_ADDRESS_3){
                    expect(pimContact.imAddresses.length).to.be.greaterThan(2);
                    expect(imAddress.Value).to.be.equal(pimContact.imAddresses[2]);
                }
            });
            

            expect(ews.PhoneNumbers).to.not.be.undefined(); 
            ews.PhoneNumbers!.Entry.forEach(phonenumber => {
                // TODO: Uncomment when LABS-1155 is implemented
                // if(phonenumber.Key == PhoneNumberKeyType.PRIMARY_PHONE){
                //   expect(phonenumber.Value).to.be.equal(pimContact.primaryPhone);
                // }else 
                if(phonenumber.Key === PhoneNumberKeyType.BUSINESS_PHONE){
                    expect(phonenumber.Value).to.be.equal(pimContact.officePhone);
                }else if(phonenumber.Key === PhoneNumberKeyType.HOME_PHONE){
                    expect(phonenumber.Value).to.be.equal(pimContact.homePhone);
                }else if(phonenumber.Key === PhoneNumberKeyType.MOBILE_PHONE){
                    expect(phonenumber.Value).to.be.equal(pimContact.cellPhone);
                }else if(phonenumber.Key === PhoneNumberKeyType.BUSINESS_FAX){
                    expect(phonenumber.Value).to.be.equal(pimContact.officeFax);
                }else if(phonenumber.Key === PhoneNumberKeyType.HOME_FAX){
                    expect(phonenumber.Value).to.be.equal(pimContact.homeFax);
                }else if(phonenumber.Key === PhoneNumberKeyType.OTHER_TELEPHONE){
                    expect(phonenumber.Value).to.be.equal(pimContact.otherPhones[0]);
                }
            });    
            
           
            expect(ews.PhysicalAddresses).to.not.be.undefined();
            ews.PhysicalAddresses!.Entry.forEach(address => {
                if(address.Key === PhysicalAddressKeyType.BUSINESS){
                    expect(address.Street).to.be.equal('Office Street');
                    expect(address.State).to.be.equal('Office State');
                    expect(address.City).to.be.equal('Office City');
                    expect(address.PostalCode).to.be.equal('Office Code');
                    expect(address.CountryOrRegion).to.be.equal('Office Country');
                } 
                else if(address.Key === PhysicalAddressKeyType.HOME){
                    expect(address.Street).to.be.equal('Home Street');
                    expect(address.State).to.be.equal('Home State');
                    expect(address.City).to.be.equal('Home City');
                    expect(address.PostalCode).to.be.equal('Home Code');
                    expect(address.CountryOrRegion).to.be.equal('Home Country');
                }
                else if(address.Key === PhysicalAddressKeyType.OTHER){
                    expect(address.Street).to.be.equal('Other Street');
                    expect(address.State).to.be.equal('Other State');
                    expect(address.City).to.be.equal('Other City');
                    expect(address.PostalCode).to.be.equal('Other Code');
                    expect(address.CountryOrRegion).to.be.equal('Other Country');
                }
            });

            expect(ews?.ExtendedProperty).to.not.be.undefined();
            expect(ews?.ExtendedProperty?.length).to.be.equal(6);

            // Validate extended properties
            let identifiers = identifiersForPathToExtendedFieldType(email1OrgDisplayName); 
            let value = getValueForExtendedFieldURI(ews, identifiers);
            expect(value).to.be.equal(pimContact.primaryEmail);
            identifiers = identifiersForPathToExtendedFieldType(email2OrgDisplayName); 
            value = getValueForExtendedFieldURI(ews, identifiers);
            expect(value).to.be.equal(pimContact.homeEmails[0]);
            identifiers = identifiersForPathToExtendedFieldType(email3OrgDisplayName); 
            value = getValueForExtendedFieldURI(ews, identifiers);
            expect(value).to.be.equal(pimContact.otherEmails[0]);
            identifiers = identifiersForPathToExtendedFieldType(email1Label); 
            value = getValueForExtendedFieldURI(ews, identifiers);
            expect(value).to.be.equal('work');
            identifiers = identifiersForPathToExtendedFieldType(email2Label); 
            value = getValueForExtendedFieldURI(ews, identifiers);
            expect(value).to.be.equal('home');
            identifiers = identifiersForPathToExtendedFieldType(email3Label); 
            value = getValueForExtendedFieldURI(ews, identifiers);
            expect(value).to.be.equal('other');
            
        });
    
    });
    
    describe('EWSContactsManager.pimItemFromEWSItem', function(){
       
        const context = getContext("testID","testPass");
    
        it('no existing PIM contact', function(){
            const itemType =new ContactItemType();
            itemType.Categories = new ArrayOfStringsType();
            itemType.Categories.String = ['test1','test2'];
            itemType.IsPrivate = false;
    
            const pimItem = EWSContactsManager.getInstance().pimItemFromEWSItem(itemType, context.request);
    
            expect(pimItem.categories).to.be.deepEqual(['test1','test2']);
            expect(pimItem.isPrivate).to.be.equal(false);
    
        });

        it('existing PIM contact', function(){
    
            const itemType =new ContactItemType();
    
            const completeName = new CompleteNameType();
            completeName.FullName = 'fullname1';
            completeName.FirstName = 'fname1';
            completeName.MiddleName = 'midname1';
            completeName.LastName = 'lname1';
            completeName.Title = 'title1';
            completeName.Suffix = 'suffix1';
        
            itemType.CompleteName =  completeName;
    
            itemType.GivenName = 'fname2';
            itemType.MiddleName = 'midname2';
            itemType.Surname = 'lname2';
    
            const email1 = new EmailAddressDictionaryEntryType("email1@test.com", EmailAddressKeyType.EMAIL_ADDRESS_1);
            const email2 = new EmailAddressDictionaryEntryType("email2@test.com", EmailAddressKeyType.EMAIL_ADDRESS_2);
            const email3 = new EmailAddressDictionaryEntryType("email3@test.com", EmailAddressKeyType.EMAIL_ADDRESS_3);
            itemType.EmailAddresses = new EmailAddressDictionaryType();
            itemType.EmailAddresses.Entry = [email1, email2, email3];
    
            const office = new PhoneNumberDictionaryEntryType("654-334-2388", PhoneNumberKeyType.BUSINESS_PHONE);
            const home = new PhoneNumberDictionaryEntryType("879-324-4423", PhoneNumberKeyType.HOME_PHONE);
            const mobile = new PhoneNumberDictionaryEntryType("876-444-2345", PhoneNumberKeyType.MOBILE_PHONE);
            const busFax = new PhoneNumberDictionaryEntryType("443-452-0944", PhoneNumberKeyType.BUSINESS_FAX);
            const homeFax = new PhoneNumberDictionaryEntryType("499-221-9322", PhoneNumberKeyType.HOME_FAX);
            const other = new PhoneNumberDictionaryEntryType("367-221-8494", PhoneNumberKeyType.OTHER_TELEPHONE);
    
            itemType.PhoneNumbers = new PhoneNumberDictionaryType();
            itemType.PhoneNumbers.Entry = [office, home, mobile, busFax, homeFax, other];
    
            const busAddress = new PhysicalAddressDictionaryEntryType();
            busAddress.Key = PhysicalAddressKeyType.BUSINESS;
            busAddress.Street = "123 Office St";
            busAddress.City = "Office City";
            busAddress.State = "Office State";
            busAddress.PostalCode = "008476";
            busAddress.CountryOrRegion = "Office Country";
    
            const homeAddress = new PhysicalAddressDictionaryEntryType();
            homeAddress.Key = PhysicalAddressKeyType.HOME;
            homeAddress.Street = "123 Home St";
            homeAddress.City = "Home City";
            homeAddress.State = "Home State";
            homeAddress.PostalCode = "0995487";
            homeAddress.CountryOrRegion = "Home Country";
    
            const otherAddress = new PhysicalAddressDictionaryEntryType();
            otherAddress.Key = PhysicalAddressKeyType.OTHER;
            otherAddress.Street = "123 Other St";
            otherAddress.City = "Other City";
            otherAddress.State = "Other State";
            otherAddress.PostalCode = "855678";
            otherAddress.CountryOrRegion = "Other Country";
    
            itemType.PhysicalAddresses = new PhysicalAddressDictionaryType();
            itemType.PhysicalAddresses.Entry = [busAddress, homeAddress, otherAddress];
    
            itemType.Body = new BodyType('comment',BodyTypeType.TEXT,false);
            itemType.JobTitle = 'Software Engineer';
            itemType.CompanyName = 'companyName';
            itemType.Birthday = new Date("1991-12-25T11:00:00.000Z");
            itemType.WeddingAnniversary = new Date("2020-12-01T11:00:00.000Z");
            itemType.BusinessHomePage = 'myBusinesspage.com';
            itemType.Department = 'myDepartment';
            itemType.OfficeLocation = 'myOfficeLocation';
            
            const pimContact = PimItemFactory.newPimContact({"@testID":"testPass"}, PimItemFormat.DOCUMENT)
            const pimItem = EWSContactsManager.getInstance().pimItemFromEWSItem(itemType, context.request, pimContact);

            expect(pimItem.firstName).to.equal('fname2');
            expect(pimItem.middleInitial).to.equal('midname2');
            expect(pimItem.lastName).to.equal('lname2');
    
            expect(pimItem.primaryEmail).to.be.equal('email1@test.com');
            expect(pimItem.otherEmails).to.containDeep(['email2@test.com','email3@test.com'])
    
            // TODO: Uncomment when LABS-1155 is implemented
            // expect(pimItem.primaryPhone).to.be.equal('879-324-4423');
            expect(pimItem.officePhone).to.be.equal('654-334-2388');
            expect(pimItem.cellPhone).to.be.equal('876-444-2345');
            expect(pimItem.officeFax).to.be.equal('443-452-0944');
            expect(pimItem.homeFax).to.be.equal('499-221-9322');
            expect(pimItem.otherPhones).to.containDeep(['367-221-8494']);
    
            expect(pimItem.officeAddress).to.containDeep({ Street: "123 Office St", City: "Office City", State: "Office State", PostalCode: "008476", Country: "Office Country" });
            expect(pimItem.homeAddress).to.containDeep({ Street: "123 Home St", City: "Home City", State: "Home State", PostalCode: "0995487", Country: "Home Country" });
            expect(pimItem.otherAddress).to.containDeep({ Street: "123 Other St", City: "Other City", State: "Other State", PostalCode: "855678", Country: "Other Country" });
           
            expect(pimItem.body).to.be.equal('comment');
            expect(pimItem.jobTitle).to.be.equal('Software Engineer');
            expect(pimItem.companyName).to.be.equal('companyName');
            expect(pimItem.location).to.be.equal('myOfficeLocation');
            expect(pimItem.birthday?.getDate()).to.be.equal(itemType.Birthday.getDate());
          // expect(pimItem.homepage).to.be.equal('myBusinesspage.com');    TODO: Remove when LABS-1283 is implemented.
           expect(pimItem.department).to.be.equal('myDepartment')
           expect(pimItem.anniversary?.getDate()).to.be.equal(itemType.WeddingAnniversary.getDate());
        });
    
    
    });
    
    describe("pimHelper.updateContactEmailProperties", function () {

        it('contact with no emails', function () {
            const pimObject = {
                "@unid": "unit-test-contact",
                "FirstName": "Unit",
                "LastName": "Test"
            };

            const manager = new MockEWSManager();
            const contact = PimItemFactory.newPimContact(pimObject, PimItemFormat.DOCUMENT);
            manager.updateContactEmailProperties(contact);

            expect(contact.extendedProperties.length).to.be.equal(0);
        });

        it('contact with 1 email', function () {
            const pimObject = {
                "@unid": "unit-test-contact",
                "FirstName": "Unit",
                "LastName": "Test",
                "MailAddress": "unit.test@work.mycom.com"
            };

            const manager = new MockEWSManager();
            const contact = PimItemFactory.newPimContact(pimObject, PimItemFormat.DOCUMENT);
            manager.updateContactEmailProperties(contact);

            expect(contact.extendedProperties.length).to.be.equal(1);
            expect(contact.extendedProperties[0]).to.containDeep({ "Value": "unit.test@work.mycom.com" });
        });

        it('contact with 2 unique emails', function () {
            const pimObject = {
                "@unid": "unit-test-contact",
                "FirstName": "Unit",
                "LastName": "Test",
                "MailAddress": "unit.test@work.mycom.com",
                "Work0Email": "unit.test@work.mycom.com",
                "HomeEmail": "ut@home.com"
            };

            const manager = new MockEWSManager();
            const contact = PimItemFactory.newPimContact(pimObject, PimItemFormat.DOCUMENT);
            manager.updateContactEmailProperties(contact);

            expect(contact.extendedProperties.length).to.be.equal(2);
            expect(contact.extendedProperties).to.containDeep([{
                DistinguishedPropertySetId: 'Address',
                PropertyType: 'String',
                PropertyId: 32899,
                Value: 'unit.test@work.mycom.com'
            }]);
            expect(contact.extendedProperties).to.containDeep([{
                DistinguishedPropertySetId: 'Address',
                PropertyType: 'String',
                PropertyId: 32915,
                Value: 'ut@home.com'
            }]);
        });

        it('contact with 3 unique emails', function () {
            const pimObject = {
                "@unid": "unit-test-contact",
                "FirstName": "Unit",
                "LastName": "Test",
                "MailAddress": "unit.test@work.mycom.com",
                "Work0Email": "unit.test@work.mycom.com",
                "work1email": "unitT@work2.mycom.com",
                "HomeEmail": "ut@home.com"
            };

            const manager = new MockEWSManager();
            const contact = PimItemFactory.newPimContact(pimObject, PimItemFormat.DOCUMENT);
            manager.updateContactEmailProperties(contact);

            expect(contact.extendedProperties.length).to.be.equal(3);
            expect(contact.extendedProperties).to.containDeep([{
                DistinguishedPropertySetId: 'Address',
                PropertyType: 'String',
                PropertyId: 32899,
                Value: 'unit.test@work.mycom.com'
            }]);
            expect(contact.extendedProperties).to.containDeep([{
                DistinguishedPropertySetId: 'Address',
                PropertyType: 'String',
                PropertyId: 32915,
                Value: 'ut@home.com'
            }]);
            expect(contact.extendedProperties).to.containDeep([{
                DistinguishedPropertySetId: 'Address',
                PropertyType: 'String',
                PropertyId: 32931,
                Value: 'unitT@work2.mycom.com'
            }]);
        });
    });

    describe('EWSContactsManager.orderedEmails', () => {

        const manager = new MockEWSManager();

        it('with no email properties set', function () {
            const pimObject = {
                "@unid": "unit-test-contact",
                "FirstName": "Unit",
                "LastName": "Test"
            };
            const pimContact = PimItemFactory.newPimContact(pimObject, PimItemFormat.DOCUMENT, KeepPimConstants.CONTACTS);
            const results = manager.orderedEmails(pimContact);

            expect(results.length).to.be.equal(0);
        });

        it('with one email property set', function () {
            const expected = ["unit_test@myco.com"];

            const pimObject = {
                "@unid": "unit-test-contact",
                "FirstName": "Unit",
                "LastName": "Test",
                "MailAddress": "unit_test@myco.com",
                "AdditionalFields": {
                    "xHCL-extProp_0": {
                        "DistinguishedPropertySetId": "Address",
                        "PropertyId": 32899,
                        "PropertyType": "String",
                        "Value": "unit_test@myco.com"
                    }
                }
            };
            const pimContact = PimItemFactory.newPimContact(pimObject, PimItemFormat.DOCUMENT, KeepPimConstants.CONTACTS);
            const results = manager.orderedEmails(pimContact);

            expect(results).to.be.eql(expected);
        });

        it('with two email properties set', function () {
            const expected = ["unit_test@myco.com", "unit_test@myhome.com"];

            const pimObject = {
                "@unid": "unit-test-contact",
                "FirstName": "Unit",
                "LastName": "Test",
                "MailAddress": "unit_test@myco.com",
                "HomeEmail": "unit_test@myhome.com",
                "AdditionalFields": {
                    "xHCL-extProp_0": {
                        "DistinguishedPropertySetId": "Address",
                        "PropertyId": 32899,
                        "PropertyType": "String",
                        "Value": "unit_test@myco.com"
                    },
                    "xHCL-extProp_1": {
                        "DistinguishedPropertySetId": "Address",
                        "PropertyId": 32915,
                        "PropertyType": "String",
                        "Value": "unit_test@myhome.com"
                    }
                }
            };
            const pimContact = PimItemFactory.newPimContact(pimObject, PimItemFormat.DOCUMENT, KeepPimConstants.CONTACTS);
            const results = manager.orderedEmails(pimContact);

            expect(results).to.be.eql(expected);
        });

        it('with three email properties set', function () {
            const expected = ["unit_test@myco.com", "unit_test@myhome.com", "unit_test@myOther.com"];

            const pimObject = {
                "@unid": "unit-test-contact",
                "FirstName": "Unit",
                "LastName": "Test",
                "MailAddress": "unit_test@myco.com",
                "HomeEmail": "unit_test@myhome.com",
                "email_1": "unit_test@myOther.com",
                'AdditionalFields': {
                    "xHCL-extProp_0": {
                        "DistinguishedPropertySetId": "Address",
                        "PropertyId": 32899,
                        "PropertyType": "String",
                        "Value": "unit_test@myco.com"
                    },
                    "xHCL-extProp_1": {
                        "DistinguishedPropertySetId": "Address",
                        "PropertyId": 32915,
                        "PropertyType": "String",
                        "Value": "unit_test@myhome.com"
                    },
                    "xHCL-extProp_2": {
                        "DistinguishedPropertySetId": "Address",
                        "PropertyId": 32931,
                        "PropertyType": "String",
                        "Value": "unit_test@myOther.com"
                    }
                }
            };
            const pimContact = PimItemFactory.newPimContact(pimObject, PimItemFormat.DOCUMENT, KeepPimConstants.CONTACTS);
            const results = manager.orderedEmails(pimContact);

            expect(results).to.be.eql(expected);
        });

        it('with multiple work and home emails', function () {
            const expected = ["unit_test@myco.com", "unit_test@myhome.com", "unit_test@myco2.com", "unit_test@myhome2.com", "unit_test@myco3.com", "unit_test@school.com", "unit_test@other.com"];

            const pimObject = {
                "@unid": "unit-test-contact",
                "FirstName": "Unit",
                "LastName": "Test",
                "MailAddress": "unit_test@myco.com",
                "Work0Email": "unit_test@myco.com",
                "work1email": "unit_test@myco2.com",
                "work2email": "unit_test@myco3.com",
                "HomeEmail": "unit_test@myhome.com",
                "home0email": "unit_test@myhome2.com"
            };

            const pimContact = PimItemFactory.newPimContact(pimObject, PimItemFormat.DOCUMENT, KeepPimConstants.CONTACTS);
            pimContact.schoolEmail = 'unit_test@school.com';
            pimContact.otherEmails = ['unit_test@other.com'];
            const results = manager.orderedEmails(pimContact);

            expect(results).to.be.eql(expected);
        });
    });
});
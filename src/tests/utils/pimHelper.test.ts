/* eslint-disable @typescript-eslint/no-non-null-assertion */
import { base64Decode, base64Encode, PimAddress, UserInfo} from '@hcllabs/openclientkeepcomponent';
/* eslint-disable no-invalid-this */
import { expect } from '@loopback/testlab';
import { 
    KeepPimConstants, PimCalendarItem, PimItemFormat, PimLabelTypes, 
    PimItemFactory } from '@hcllabs/openclientkeepcomponent';
import { PathToExtendedFieldType, PathToUnindexedFieldType } from '../../models/common.model';
import { 
    DefaultShapeNamesType, DistinguishedFolderIdNameType, DistinguishedPropertySetType, MapiPropertyIds,
    MapiPropertyTypeType, ExtendedPropertyKeyType, UnindexedFieldURIType,
    EmailAddressKeyType, PhysicalAddressKeyType, PhoneNumberKeyType,  SensitivityChoicesType } from '../../models/enum.model';

import { 
    CalendarItemType, DistinguishedFolderIdType, ExtendedPropertyType, FolderIdType, FolderResponseShapeType,
    ItemIdType, ItemResponseShapeType, NonEmptyArrayOfPathsToElementType} from '../../models/mail.model';
import { 
    addExtendedPropertiesToPIM, findLabel, rootFolderIdForUser, rootFolderForUser, 
    addAdditionalPropertiesToEWSItem, getEWSId, getKeepIdPair, parseSizeRequestedValue, filterPimContactsByFolderIds, createContactItemFromPimContact, getLabelTypeForEmail, EmailLabelType} from '../../utils';
import { UserContext } from '../../keepcomponent';
import {getContext} from '../unitTestHelper';

describe('pimHelper tests', () => {
    describe('pimHelper.addExtendedPropertiesToPIM', () => {

        const startDate = "2021-07-16T04:00:00.000Z";
        const endDate = "2021-07-17T04:00:00.000Z";
        let item: CalendarItemType;
        let pimItem: PimCalendarItem;

        // Reset pimItem before each test
        beforeEach(function () {
            console.info(`Setup before "${this.currentTest?.title}"`);
            pimItem = PimItemFactory.newPimCalendarItem({}, "default", PimItemFormat.DOCUMENT);
            pimItem.unid = 'THISISATEST';
            pimItem.subject = 'Unit Test Meeting';
            pimItem.start = startDate;
            pimItem.end = endDate;


            item = new CalendarItemType();
            item.ItemId = new ItemIdType("unit-test");
            const property1 = new ExtendedPropertyType();
            property1.ExtendedFieldURI = new PathToExtendedFieldType();
            property1.ExtendedFieldURI.PropertyName = "CalendarTimeZone";
            property1.ExtendedFieldURI.PropertySetId = "A7B529B5-4B75-47A7-A24F-20743D6C55CD";
            property1.Value = "America/New_York";
            const property2 = new ExtendedPropertyType();
            property2.ExtendedFieldURI = new PathToExtendedFieldType();
            property2.ExtendedFieldURI.DistinguishedPropertySetId = DistinguishedPropertySetType.COMMON;
            property2.ExtendedFieldURI.PropertyId = MapiPropertyIds.REMINDER_SET;
            property2.Value = "true";
            item.ExtendedProperty = [property1, property2];

        });

        it("adds all extended properties", function () {

            addExtendedPropertiesToPIM(item, pimItem);

            expect(pimItem.extendedProperties.length).to.be.equal(2);
        });

        it("adds a matching extended property", function () {
            const identifiers: any = {};
            identifiers[ExtendedPropertyKeyType.PROPERTY_NAME] = "CalendarTimeZone";
            identifiers[ExtendedPropertyKeyType.PROPERTY_SETID] = "A7B529B5-4B75-47A7-A24F-20743D6C55CD";

            addExtendedPropertiesToPIM(item, pimItem, [identifiers]);

            expect(pimItem.extendedProperties.length).to.be.equal(1);

            const property = pimItem.extendedProperties[0];
            expect(property["Value"]).to.equal("America/New_York");

        });

        it("does not find a matching extened property", function () {
            const identifiers: any = {};
            identifiers[ExtendedPropertyKeyType.PROPERTY_NAME] = "UnitTest";
            identifiers[ExtendedPropertyKeyType.PROPERTY_SETID] = "A7B529B5-4B75-47A7-A24F-20743D6C55CD";

            addExtendedPropertiesToPIM(item, pimItem, [identifiers]);

            expect(pimItem.extendedProperties.length).to.be.equal(0);
        });

    });

    describe('pimHelper.findLabel', () => {
        const labels = [
            PimItemFactory.newPimLabel({ "FolderId": "222-444-5554", "View": KeepPimConstants.INBOX, "DisplayName": "Inbox" }, PimItemFormat.DOCUMENT),
            PimItemFactory.newPimLabel({ "FolderId": "ZGVmYXVsdF83NzctNDU1Ni0zMzMz", "View": KeepPimConstants.CALENDAR, "DisplayName": "Calendar" }, PimItemFormat.DOCUMENT),
            PimItemFactory.newPimLabel({ "FolderId": "folder-0002", "DisplayName": "Unit Test Folder Two" }, PimItemFormat.DOCUMENT)];

        it("with distinguished folder id", function () {
            const folderId = new DistinguishedFolderIdType();
            folderId.Id = DistinguishedFolderIdNameType.CALENDAR;
            const label = findLabel(labels, folderId);

            expect(label).to.not.be.null();
            expect(label?.view).to.equal(KeepPimConstants.CALENDAR);
        });

        it("with regular folder id", function () {
            const folderEWSId = getEWSId('222-444-5554');
            const folderId = new FolderIdType(folderEWSId, `ck-${folderEWSId}`);
            const label = findLabel(labels, folderId);

            expect(label).to.not.be.null();
            expect(label?.view).to.equal(KeepPimConstants.INBOX);
        });
    });

    describe('pimHelper.rootFolderForUser', () => {
        const uInfo: any = {};
        uInfo.userId = 'testUser'; 
        const userInfo: UserInfo = uInfo as UserInfo;

        const childLabels = [
            PimItemFactory.newPimLabel({ "FolderId": "folder1", "View": KeepPimConstants.INBOX, "DocumentCount": 0, "DisplayName": "Inbox", "Type": PimLabelTypes.MAIL }, PimItemFormat.DOCUMENT),
            PimItemFactory.newPimLabel({ "FolderId": "folder2", "View": KeepPimConstants.SENT, "DocumentCount": 0, "DisplayName": "Sent", "Type": PimLabelTypes.MAIL }, PimItemFormat.DOCUMENT),
            PimItemFactory.newPimLabel({ "FolderId": "folder3", "View": KeepPimConstants.CALENDAR, "DocumentCount": 0, "DisplayName": "Calendar", "Type": PimLabelTypes.CALENDAR }, PimItemFormat.DOCUMENT),
            PimItemFactory.newPimLabel({ "FolderId": "folder4", "View": KeepPimConstants.TASKS, "DocumentCount": 0, "DisplayName": "Tasks", "Type": PimLabelTypes.TASKS }, PimItemFormat.DOCUMENT),
            PimItemFactory.newPimLabel({ "FolderId": "folder5", "View": KeepPimConstants.CONTACTS, "DocumentCount": 0, "DisplayName": "Contacts", "Type": PimLabelTypes.CONTACTS }, PimItemFormat.DOCUMENT),
            PimItemFactory.newPimLabel({ "FolderId": "folder6", "View": KeepPimConstants.JOURNAL, "DocumentCount": 0, "DisplayName": "Notes", "Type": PimLabelTypes.JOURNAL }, PimItemFormat.DOCUMENT)
        ];

        describe('with no additional properties', function () {
            it('for shape IdOnly', function () {
                const shape = new FolderResponseShapeType();
                shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
                const root = rootFolderForUser(userInfo, shape, childLabels);

                expect(root.FolderId?.Id).to.be.eql(rootFolderIdForUser(userInfo.userId).Id);
                expect(root.ExtendedProperty?.length).to.eql(0);
                expect(root.DisplayName).to.be.undefined();
                expect(root.ChildFolderCount).to.be.undefined();
                expect(root.TotalCount).to.be.undefined();
            });

            it('for shape Default', function () {
                const shape = new FolderResponseShapeType();
                shape.BaseShape = DefaultShapeNamesType.DEFAULT;
                const root = rootFolderForUser(userInfo, shape, childLabels);

                expect(root.FolderId?.Id).to.be.eql(rootFolderIdForUser(userInfo.userId).Id);
                expect(root.ExtendedProperty?.length).to.eql(0);
                expect(root.DisplayName).to.eql("Top of Information Store");
                expect(root.ChildFolderCount).to.eql(6);
                expect(root.TotalCount).to.eql(0);
            });

            it('for shape All', function () {
                const shape = new FolderResponseShapeType();
                shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
                const root = rootFolderForUser(userInfo, shape, childLabels);

                expect(root.FolderId?.Id).to.be.eql(rootFolderIdForUser(userInfo.userId).Id);
                expect(root.ExtendedProperty?.length).to.eql(0);
                expect(root.DisplayName).to.eql("Top of Information Store");
                expect(root.ChildFolderCount).to.eql(6);
                expect(root.TotalCount).to.eql(0);
            });

        });

        describe('with additional properties', function () {
            const path1 = new PathToExtendedFieldType();
            path1.PropertyTag = "0x10f4";
            path1.PropertyType = MapiPropertyTypeType.BOOLEAN;
            const path2 = new PathToExtendedFieldType();
            path2.PropertyTag = "0x670a";
            path2.PropertyType = MapiPropertyTypeType.SYSTEM_TIME;
            const path3 = new PathToExtendedFieldType();
            path3.PropertyTag = "0x670b";
            path3.PropertyType = MapiPropertyTypeType.INTEGER;

            it('for shape IdOnly', function () {
                const shape = new FolderResponseShapeType();
                shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
                shape.AdditionalProperties = new NonEmptyArrayOfPathsToElementType([path1, path2, path3]);
                const root = rootFolderForUser(userInfo, shape, childLabels);

                expect(root.FolderId?.Id).to.be.eql(rootFolderIdForUser(userInfo.userId).Id);
                expect(root.ExtendedProperty?.length).to.eql(3);
                expect(root.DisplayName).to.be.undefined();
                expect(root.ChildFolderCount).to.be.undefined();
                expect(root.TotalCount).to.be.undefined();
            });

            it('for shape Default', function () {
                const shape = new FolderResponseShapeType();
                shape.BaseShape = DefaultShapeNamesType.DEFAULT;
                shape.AdditionalProperties = new NonEmptyArrayOfPathsToElementType([path1, path2, path3]);
                const root = rootFolderForUser(userInfo, shape, childLabels);

                expect(root.FolderId?.Id).to.be.eql(rootFolderIdForUser(userInfo.userId).Id);
                expect(root.ExtendedProperty?.length).to.eql(3);
                expect(root.DisplayName).to.eql("Top of Information Store");
                expect(root.ChildFolderCount).to.eql(6);
                expect(root.TotalCount).to.eql(0);
            });

            it('for shape All', function () {
                const shape = new FolderResponseShapeType();
                shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
                shape.AdditionalProperties = new NonEmptyArrayOfPathsToElementType([path1, path2, path3]);
                const root = rootFolderForUser(userInfo, shape, childLabels);

                expect(root.FolderId?.Id).to.be.eql(rootFolderIdForUser(userInfo.userId).Id);
                expect(root.ExtendedProperty?.length).to.eql(3);
                expect(root.DisplayName).to.eql("Top of Information Store");
                expect(root.ChildFolderCount).to.eql(6);
                expect(root.TotalCount).to.eql(0);
            });
        });
    });

    describe('addAdditionalPropertiesToEWSItem', function () {

        it('all day event', function () {
            const pimObject: any = {
                uid: 'unit-test-id',
                start: new Date(),
                duration: 'P1D',
                showWithoutTime: true
            }

            // Verify all day event in Keep, EWS all-day is true
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            shape.AdditionalProperties = new NonEmptyArrayOfPathsToElementType();
            const field = new PathToUnindexedFieldType();
            field.FieldURI = UnindexedFieldURIType.CALENDAR_IS_ALL_DAY_EVENT;
            shape.AdditionalProperties.items = [ field ];

            let ewsItem = new CalendarItemType();
            let pimCalendarItem = PimItemFactory.newPimCalendarItem(pimObject, "default", PimItemFormat.DOCUMENT);
            addAdditionalPropertiesToEWSItem(pimCalendarItem, ewsItem, shape); 
            expect(ewsItem.IsAllDayEvent).to.not.be.undefined();
            expect(ewsItem.IsAllDayEvent).to.be.true();

            // Verify if all day with participants, EWS all-day is false
            pimObject.participants = {
                "dG9tQGZvb2Jhci5xlLmNvbQ": {
                    "@type": "Participant",
                    "name": "Tom Tool",
                    "email": "tom@foobar.example.com",
                    "sendTo": {
                        "imip": "mailto:tom@calendar.example.com"
                    },
                    "participationStatus": "accepted",
                    "roles": {
                        "attendee": true
                    }
                }
            };

            ewsItem = new CalendarItemType();
            pimCalendarItem = PimItemFactory.newPimCalendarItem(pimObject, "default", PimItemFormat.DOCUMENT);
            addAdditionalPropertiesToEWSItem(pimCalendarItem, ewsItem, shape); 
            expect(ewsItem.IsAllDayEvent).to.not.be.undefined();
            expect(ewsItem.IsAllDayEvent).to.be.false();

        });
    });

    describe('pimHelper.getEWSId', () => {
        it('returns unid when mailboxId is undefiend', () => {
            const mockItemId = '738CB93AE2E1B7F8852564B5001283E2';

            const receivedResult = getEWSId(mockItemId, '');

            expect(base64Decode(receivedResult)).to.equal(mockItemId);
        });
        it('combines the unid and mailboxId', () => {
            const mockItemId = '738CB93AE2E1B7F8852564B5001283E2';
            const mockMailboxId = 'test@test.com';
            const expectedResult = mockItemId + '##' + mockMailboxId;

            const receivedResult = getEWSId(mockItemId, mockMailboxId);

            expect(base64Decode(receivedResult)).to.equal(expectedResult);
        });
    });

    describe('pimHelper.getKeepIdPair', () => {
        const mockItemId = '738CB93AE2E1B7F8852564B5001283E2';
        const mockMailboxId = 'test@test.com';

        it('returns a pair with the unid and mailboxId', () => {
            const mockCombinedIds = base64Encode(mockItemId + '##' + mockMailboxId);
            
            const [receivedItemId, receivedMailboxId] = getKeepIdPair(mockCombinedIds);

            expect(receivedItemId).to.equal(mockItemId);
            expect(receivedMailboxId).to.equal(mockMailboxId);
        });

        it('returns the unid and undefined when combinedIds includes only unid', () => {
            const mockCombinedIds = base64Encode(mockItemId);
            
            const [receivedItemId, receivedMailboxId] = getKeepIdPair(mockCombinedIds);

            expect(receivedItemId).to.equal(mockItemId);
            expect(receivedMailboxId).to.equal(undefined);
        });

        it('returns undefined when an empty combinedIds was passed', () => {
            const mockCombinedIds = '';

            const [receivedItemId, receivedMailboxId] = getKeepIdPair(mockCombinedIds);

            expect(receivedItemId).to.equal(undefined);
            expect(receivedMailboxId).to.equal(undefined);
        });
    });

    describe('pimHelper.parseSizeRequestedValue', () => {
        it('returns an array of two numbers if the value was provided in proper format', () => {
            const testData: [string, [number, number]][] = [
                ['HR233x233', [233, 233]],
                ['hr25x25', [25, 25]],
                ['hR3000X3000', [3000, 3000]],
            ]

            testData.forEach(([value, expectedResult]) => {
                expect(parseSizeRequestedValue(value)).to.deepEqual(expectedResult);
            })
        });

        it('returns [undefined, undefined] when invalid string was provided', () => {
            const invalidValues: string[] = ['230x200', 'hd340x300', 'test'];

            invalidValues.forEach(invalidValue => {
                expect(parseSizeRequestedValue(invalidValue)).to.deepEqual([undefined, undefined]);
            }) 
        });
    });

    describe('pimHelper.filterPimContactsByFolderIds', () => {
    
        const pimContact1 = PimItemFactory.newPimContact({}, PimItemFormat.DOCUMENT);
        pimContact1.unid = 'UNID1';
        pimContact1.parentFolderIds = ["222-444-5554"];
        
        
        const pimContact2 = PimItemFactory.newPimContact({}, PimItemFormat.DOCUMENT);
        pimContact2.unid = 'UNID2';
        pimContact2.parentFolderIds = [DistinguishedFolderIdNameType.INBOX];

        const pimContactsInitial = [pimContact1, pimContact2];

        
        it("with distinguished folder id", function () {
            const folderId = new DistinguishedFolderIdType();
            folderId.Id = DistinguishedFolderIdNameType.INBOX;

            const pimContacts = filterPimContactsByFolderIds(pimContactsInitial, [folderId]);

            expect(pimContacts[0]).to.not.be.null();
            expect(pimContacts[0]?.unid).to.equal('UNID2');
            if(pimContacts[0]?.parentFolderIds) expect(pimContacts[0]?.parentFolderIds[0]).to.equal(DistinguishedFolderIdNameType.INBOX);
        });

        it("with regular folder id", function () {
            const folderEWSId = getEWSId('222-444-5554');
            const folderId = new FolderIdType(folderEWSId, `ck-${folderEWSId}`);

            const pimContacts = filterPimContactsByFolderIds(pimContactsInitial, [folderId]);

            expect(pimContacts[0]).to.not.be.null();
            expect(pimContacts[0]?.unid).to.equal('UNID1');
            if(pimContacts[0]?.parentFolderIds) expect(pimContacts[0]?.parentFolderIds[0]).to.equal('222-444-5554');
        });
        
    });

    describe('pimHelper.createContactItemFromPimContact', () => {

        const userInfo = new UserContext();

        const testUser = "test.user@test.org"; 
        const context = getContext(testUser, 'password');

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

        const pimContact = PimItemFactory.newPimContact(pimObject, PimItemFormat.DOCUMENT)
        
        pimContact.jobTitle='Software Engineer';
        
        pimContact.primaryEmail = 'primaryemail@mytest.com';
        pimContact.homeEmails = ['homeemail@mytest.com'];
        pimContact.workEmails = ['workemail@mytest.com'];

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

        it("for shape IdOnly", async function () {

        const returnFullContactData = false;
        const contactDataShape = DefaultShapeNamesType.ID_ONLY
        const ews = await createContactItemFromPimContact(pimContact, userInfo, context.request, returnFullContactData, contactDataShape);

        expect(ews?.ItemId?.Id).to.be.equal(getEWSId('testID'));
        expect(ews?.ParentFolderId?.Id).to.be.equal(getEWSId("parentID"));

        expect(ews?.JobTitle).to.be.undefined();
        expect(ews?.CompanyName).to.be.undefined();

        });


        it("for shape Default", async function () {

            const returnFullContactData = false;
            const contactDataShape = DefaultShapeNamesType.DEFAULT
            const ews = await createContactItemFromPimContact(pimContact, userInfo, context.request, returnFullContactData, contactDataShape);
    
            expect(ews?.ItemId?.Id).to.be.equal(getEWSId('testID'));
            expect(ews?.ParentFolderId?.Id).to.be.equal(getEWSId("parentID"));
    
            expect(ews?.Body).to.be.undefined();
            expect(ews?.Sensitivity).to.be.undefined();
            expect(ews?.Department).to.be.undefined();
            expect(ews?.OfficeLocation).to.be.undefined();
            expect(ews?.Birthday).to.be.undefined();
            expect(ews?.WeddingAnniversary).to.be.undefined();
            expect(ews?.BusinessHomePage).to.be.undefined();
            expect(ews?.Attachments).to.be.undefined();
            expect(ews?.Categories).to.be.undefined();


            expect(ews?.CompleteName?.Suffix).to.be.equal(pimObject.Suffix);
            expect(ews?.CompleteName?.Title).to.be.equal(pimObject.Title);
            expect(ews?.CompleteName?.FirstName).to.be.equal(pimObject.FirstName);
            expect(ews?.CompleteName?.LastName).to.be.equal(pimObject.LastName);
            expect(ews?.CompleteName?.MiddleName).to.be.equal(pimObject.MiddleInitial);
            expect(ews?.JobTitle).to.be.equal(pimContact.jobTitle);
            expect(ews?.CompanyName).to.be.equal(pimObject.CompanyName);
    
        
    
            if(ews?.EmailAddresses?.Entry){
            ews.EmailAddresses.Entry.forEach(emailaddress => {
                if(emailaddress.Key === EmailAddressKeyType.EMAIL_ADDRESS_1){
                    expect(emailaddress.Value).to.be.equal(pimContact.primaryEmail);
                }else if(emailaddress.Key === EmailAddressKeyType.EMAIL_ADDRESS_2){
                    expect(emailaddress.Value).to.be.equal(pimContact.workEmails[0]);
                }else if(emailaddress.Key === EmailAddressKeyType.EMAIL_ADDRESS_3){
                    expect(emailaddress.Value).to.be.equal(pimContact.homeEmails[0]);
                }
            });
    
            } 
            
            if(ews?.PhoneNumbers?.Entry){
            ews.PhoneNumbers.Entry.forEach(phonenumber => {
    
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
            })    
            }
            
            if(ews?.PhysicalAddresses?.Entry){
            ews.PhysicalAddresses.Entry.forEach(address => {
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
            })
            }

        });

        it("for shape AllProperties", async function () {

        const returnFullContactData = false;
        const contactDataShape = DefaultShapeNamesType.ALL_PROPERTIES
        const ews = await createContactItemFromPimContact(pimContact, userInfo, context.request, returnFullContactData, contactDataShape);

        expect(ews?.ItemId?.Id).to.be.equal(getEWSId('testID'));
        expect(ews?.ParentFolderId?.Id).to.be.equal(getEWSId("parentID"));

        expect(ews?.Body?.Value).to.be.equal(pimObject.Comment);
        expect(ews?.CompleteName?.Suffix).to.be.equal(pimObject.Suffix);
        expect(ews?.CompleteName?.Title).to.be.equal(pimObject.Title);
        expect(ews?.CompleteName?.FirstName).to.be.equal(pimObject.FirstName);
        expect(ews?.CompleteName?.LastName).to.be.equal(pimObject.LastName);
        expect(ews?.CompleteName?.MiddleName).to.be.equal(pimObject.MiddleInitial);
        expect(ews?.JobTitle).to.be.equal(pimContact.jobTitle);
        expect(ews?.CompanyName).to.be.equal(pimObject.CompanyName);

        expect(ews?.Sensitivity).to.be.equal(SensitivityChoicesType.CONFIDENTIAL);

        if(ews?.EmailAddresses?.Entry){
            ews.EmailAddresses.Entry.forEach(emailaddress => {
                if(emailaddress.Key === EmailAddressKeyType.EMAIL_ADDRESS_1){
                    expect(emailaddress.Value).to.be.equal(pimContact.primaryEmail);
                }else if(emailaddress.Key === EmailAddressKeyType.EMAIL_ADDRESS_2){
                    expect(emailaddress.Value).to.be.equal(pimContact.workEmails[0]);
                }else if(emailaddress.Key === EmailAddressKeyType.EMAIL_ADDRESS_3){
                    expect(emailaddress.Value).to.be.equal(pimContact.homeEmails[0]);
                }
            });

        } 
        
        if(ews?.PhoneNumbers?.Entry){
            ews.PhoneNumbers.Entry.forEach(phonenumber => {

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
            })    
        }
        
        if(ews?.PhysicalAddresses?.Entry){
            ews.PhysicalAddresses.Entry.forEach(address => {
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
            })
        }

            expect(ews?.Department).to.be.equal(pimObject.Department);
            expect(ews?.OfficeLocation).to.be.equal(pimObject.Location);
            if(ews?.Birthday) expect(ews.Birthday!.getTime() - new Date("1991-12-25T11:00:00.000Z").getTime()).to.equal(0);
            if(ews?.WeddingAnniversary) expect(ews.WeddingAnniversary!.getTime() - new Date("2020-12-01T11:00:00.000Z").getTime()).to.equal(0);    
            expect(ews?.BusinessHomePage).to.be.equal(pimContact.homepage);
            expect(ews?.HasAttachments).to.be.equal(true);
            expect(ews?.Attachments?.items.length).to.be.equal(1);
            expect(ews?.Categories?.String).to.containDeep(pimContact.categories);
        });

        it("for returnFullContactData", async function () {

            const returnFullContactData = true;
            const contactDataShape = DefaultShapeNamesType.ID_ONLY
            const ews = await createContactItemFromPimContact(pimContact, userInfo, context.request, returnFullContactData, contactDataShape);
    
            expect(ews?.ItemId?.Id).to.be.equal(getEWSId('testID'));
            expect(ews?.ParentFolderId?.Id).to.be.equal(getEWSId("parentID"));
    
            expect(ews?.JobTitle).to.be.not.undefined();
            expect(ews?.CompanyName).to.be.not.undefined();
    
            expect(ews?.Body).to.be.not.undefined();
            expect(ews?.Sensitivity).to.be.not.undefined();
            expect(ews?.Department).to.be.not.undefined();
            expect(ews?.OfficeLocation).to.be.not.undefined();
            expect(ews?.Birthday).to.be.not.undefined();
            expect(ews?.WeddingAnniversary).to.be.not.undefined();
            expect(ews?.BusinessHomePage).to.be.not.undefined();
            expect(ews?.Attachments).to.be.not.undefined();
            expect(ews?.Categories).to.be.not.undefined();

        });

    });

    describe('pimHelper.getLabelTypeForEmail', () => {
        const pimObject: any = {
            "@unid": "AAB8D300562E552500258759006E41B1",
            "@noteid": 2578,
            "@created": "2021-09-23T20:04:17Z",
            "@lastmodified": "2021-10-18T20:27:39Z",
            "@lastaccessed": "2021-10-18T20:27:39Z",
            "@size": 1014,
            "@unread": false,
            "@etag": "W/\" 616dd8bb\"",
            "@threadid": "AAB8D300562E552500258759006E41B1",
            "$Disclaimed": "",
            "$NoPurge": "1",
            "$PublicAccess": "1",
            "AddressFldDisplayed": "0",
            "AltFullName": "",
            "AltFullNameLanguage": "",
            "AltFullNameSort": "",
            "Anniversary": "",
            "AreaCodeFromLoc": "",
            "Assistant": "John Doe",
            "Assistant2Phone": "",
            "Assistant3Phone": "",
            "AssistantPhone": "",
            "Birthday": "",
            "BusinessAddress": "",
            "CarPhone": "",
            "Categories": "demo1, acceptance",
            "CellPhoneNumber": "999-222-1111",
            "City": "Raleigh",
            "Comment": "This is a comments section.",
            "CompanyName": "A Software Company",
            "CompanyNameSort": "",
            "Confidential": "",
            "country": "USA",
            "Department": "A334",
            "email_1": "testuser@business.hcl.com",
            "FirstName": "Test",
            "Form": "Person",
            "From": "rustyg.miramare@quattro.rocks",
            "FullName": [
                "Test User",
                "User Test",
                "testuser@business.hcl.com"
            ],
            "FullNameInput": "Test User",
            "Home2Phone": "",
            "Home3Phone": "",
            "HomeAddress": "",
            "HomeEmail": "testuser@personal.hcl.com",
            "HomeFAXPhoneNumber": "",
            "HomeURL": "",
            "InternetAddress": "testuser@business.hcl.com",
            "InternetAddress1": "testuser@business.hcl.com",
            "JobTitle": "Senior Tester",
            "LastName": "User",
            "Location": "Cary, NC",
            "MailAddress": "testuser@business.hcl.com",
            "MailDomain": "",
            "MailSystem": "1",
            "Manager": "John Doe",
            "MiddleInitial": "",
            "mobileEmail": "testuser@mobile.com",
            "OfficeCity": "Cary",
            "OfficeCountry": "USA",
            "OfficeFAXPhoneNumber": "999-555-1111",
            "OfficePhoneNumber": "999-111-1111",
            "OfficeState": "NC",
            "OfficeStreetAddress": "99 Business St",
            "OfficeZIP": "27760",
            "OtherEmail": "testuser@other.com",
            "OtherFaxPhone": "",
            "OtherPhone": "",
            "OtherURL": "",
            "PersPager": "",
            "PhoneFldsDisplayed": "0,6,12,13,3,9,10,4",
            "PhoneNumber": "999-444-1111",
            "PhoneNumber_1": "999-333-1111",
            "PhoneNumber_10": "",
            "PhoneNumber_2": "999-333-1111",
            "PhoneNumber_6": "",
            "PhoneNumber_8": "",
            "primaryPhoneNumber": "999-111-1111",
            "SametimeLogin": "testy@im.test.io",
            "SchoolEmail": "testuser@school.com",
            "Spouse": "Jane Doe",
            "State": "NC",
            "StreetAddress": "99 Personal Ave",
            "Type": "Person",
            "WebSite": "https://www.mysoftware.com",
            "Work0Email": "testuser@business.hcl.com",
            "work1email": "testuser@business2.hcl.com",
            "work2email": "",
            "Work2Phone": "",
            "Work2URL": "",
            "Work3Email": "",
            "Work3Phone": "",
            "Work3URL": "",
            "WorkFax2Phone": "",
            "WorkFax3Phone": "",
            "Zip": "27765"
        };

        const pimContact = PimItemFactory.newPimContact(pimObject, PimItemFormat.DOCUMENT); 

        it('label type is work', () => {
            expect(getLabelTypeForEmail('testuser@business.hcl.com', pimContact)).to.be.equal(EmailLabelType.WORK);
            expect(getLabelTypeForEmail('testuser@business2.hcl.com', pimContact)).to.be.equal(EmailLabelType.WORK);
        });

        it('label type is home', () => {
            expect(getLabelTypeForEmail('testuser@personal.hcl.com', pimContact)).to.be.equal(EmailLabelType.HOME);
        });

        it('label type is other', () => {
            expect(getLabelTypeForEmail('testuser@school.com', pimContact)).to.be.equal(EmailLabelType.OTHER);
            expect(getLabelTypeForEmail('testuser@mobile.com', pimContact)).to.be.equal(EmailLabelType.OTHER);
            expect(getLabelTypeForEmail('testuser@other.com', pimContact)).to.be.equal(EmailLabelType.OTHER);
        });

        it('label type can not be determined', () => {
            expect(getLabelTypeForEmail('testuser@unknown.com', pimContact)).to.be.undefined(); 
        });
    });
});
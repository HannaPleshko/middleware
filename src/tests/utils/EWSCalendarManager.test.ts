/* eslint-disable @typescript-eslint/no-non-null-assertion */
/* eslint-disable no-invalid-this */
import { KeepPimCalendarManager, KeepPimConstants, KeepPimLabelManager, KeepPimManager, KeepPimMessageManager, PimCalendarItem, PimImportance, PimItemFactory, PimItemFormat } from '@hcllabs/openclientkeepcomponent';
import { expect, sinon } from '@loopback/testlab';
import { UserContext } from '../../keepcomponent';
import { ArrayOfStringsType, PathToExtendedFieldType, PathToUnindexedFieldType } from '../../models/common.model';
import {
    DefaultShapeNamesType, MessageDispositionType, UnindexedFieldURIType, ExtendedPropertyKeyType, MapiPropertyTypeType,
    BodyTypeType, ImportanceChoicesType, SensitivityChoicesType, MonthNamesType, DayOfWeekType, DayOfWeekIndexType, ResponseTypeType
} from '../../models/enum.model';
import {
    AppendToItemFieldType,
    AttendeeType, BodyType, CalendarItemType, DeleteItemFieldType, EmailAddressType, ExtendedPropertyType, ItemChangeType,
    ItemIdType, ItemResponseShapeType, NonEmptyArrayOfAttendeesType, NonEmptyArrayOfItemChangeDescriptionsType,
    NonEmptyArrayOfPathsToElementType,
    SetItemFieldType, SingleRecipientType, TimeZoneType
} from '../../models/mail.model';
import { EWSCalendarManager, EWSServiceManager, getEWSId } from '../../utils';
import {
    compareDateStrings, generateTestCalendarItems, generateTestCalendarLabel, generateTestLabels, getContext, stubCalendarManager,
    stubLabelsManager, stubMessageManager, stubPimManager
} from '../unitTestHelper';
import util from 'util';
import { DateTime } from 'luxon';

class MockEWSCalendarManager extends EWSCalendarManager {
    modifyCalendarStructure(item: CalendarItemType, pimItem: PimCalendarItem, removedAttendees: string[]): any | undefined {
        return super.modifyCalendarStructure(item, pimItem, removedAttendees);
    }
    updatePimItemFieldValue(pimCalendarItem: PimCalendarItem, fieldIdentifier: UnindexedFieldURIType, newValue: any | undefined): boolean {
        return super.updatePimItemFieldValue(pimCalendarItem, fieldIdentifier, newValue);
    }

    getAttendeeResponseType(attendee: string, pimCalendarItem: PimCalendarItem): ResponseTypeType {
        return super.getAttendeeResponseType(attendee, pimCalendarItem);
    }
}

describe('EWSCalendarManager tests', () => {

    const testUser = "test.user@test.org";
    const context = getContext(testUser, 'password');
    const userInfo = new UserContext(); // Don't set userid/password so if real Keep methods called it will fail

    afterEach(function () {
        sinon.restore();
    });

    it('getInstance', function () {
        const manager = EWSCalendarManager.getInstance();
        expect(manager).to.be.instanceof(EWSCalendarManager);

        const manager2 = EWSCalendarManager.getInstance();
        expect(manager).to.be.equal(manager2);
    });


    describe('Test getItems', () => {

        describe('getItems successful and validate pimItemToEWSItem', () => {
            let testData: PimCalendarItem[] = [];

            before(() => {
                // Setup the test data
                testData = generateTestCalendarItems(KeepPimConstants.DEFAULT_CALENDAR_NAME);

                const expectedRequired = ['req.email1@test.com', 'req.email2@test.com'];
                const expectedOptional = ['opt.email1@test.com'];
                const expectedOrganizer = 'organizer@test.com';
                const expectedLocation = 'Conference Room';
                const expectedStart = '2020-03-18T13:58:51-04:00';
                const expectedCategories = ["cat1", "cat2"];
                const expectedAlarm = -5;
                const expectedCalUID = 'CAL-UID';

                testData.forEach(pimItem => {
                    pimItem.requiredAttendees = expectedRequired;
                    pimItem.optionalAttendees = expectedOptional;
                    pimItem.organizer = expectedOrganizer;
                    pimItem.location = expectedLocation;
                    pimItem.start = expectedStart;
                    pimItem.categories = expectedCategories;
                    pimItem.alarm = expectedAlarm;
                    pimItem.setAdditionalProperty(UnindexedFieldURIType.CALENDAR_UID, expectedCalUID);
                    const prop: any = {};
                    prop[ExtendedPropertyKeyType.PROPERTY_SETID] = "A7B529B5-4B75-47A7-A24F-20743D6C55CD";
                    prop[ExtendedPropertyKeyType.PROPERTY_NAME] = "CalendarTimeZone";
                    prop[ExtendedPropertyKeyType.PROPERTY_TYPE] = MapiPropertyTypeType.STRING;
                    prop.Value = 'America/New_York';
                    pimItem.addExtendedProperty(prop);
                });
            });

            beforeEach(() => {
                const calManagerStub = sinon.createStubInstance(KeepPimCalendarManager);
                calManagerStub.getCalendarItems.withArgs(KeepPimConstants.DEFAULT_CALENDAR_NAME, sinon.match.any, sinon.match.any, sinon.match.any, 0, 10).resolves(testData);
                calManagerStub.getCalendarItems.withArgs('Calendar 2', sinon.match.any, sinon.match.any, sinon.match.any, 0, 10).resolves([]);
                calManagerStub.getCalendarItems.callThrough();
                stubCalendarManager(calManagerStub);

                const labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
                labelManagerStub.getLabels.resolves(generateTestLabels());
                stubLabelsManager(labelManagerStub);

            });

            it('with Shape ID only', async () => {
                const manager = EWSCalendarManager.getInstance();
                const shape = new ItemResponseShapeType();
                shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
                const parentLabel = generateTestCalendarLabel();

                let items = await manager.getItems(userInfo, context.request, shape, 0, 10, parentLabel);
                expect(items).to.be.an.Array();
                expect(items.length).to.be.equal(testData.length);
                for (let index = 0; index < items.length; index++) {
                    const item = items[index];
                    expect(item.ItemId?.Id).to.be.equal(getEWSId(testData[index].unid));
                    expect(item.ParentFolderId?.Id).to.be.equal(getEWSId(parentLabel.folderId));
                    expect(item.ItemClass).to.be.equal('IPM.Appointment');

                    // Should not be set
                    expect(item.Subject).to.be.undefined();
                    expect(item.UID).to.be.undefined();
                    expect(item.Organizer).to.be.undefined();
                    expect(item.Location).to.be.undefined();
                    expect(item.Body).to.be.undefined();
                    expect(item.Categories).to.be.undefined();
                    expect(item.ReminderMinutesBeforeStart).to.be.undefined();
                    expect(item.RequiredAttendees).to.be.undefined();
                    expect(item.OptionalAttendees).to.be.undefined();
                    expect(item.Start).to.be.undefined();
                    expect((item.ExtendedProperty ?? []).length).to.be.equal(0);
                }

                // No items returned
                items = await manager.getItems(userInfo, context.request, shape, 0, 10, generateTestCalendarLabel(true));
                expect(items).to.be.an.Array();
                expect(items.length).to.be.equal(0);
            });

            it('with Shape default', async () => {
                const manager = EWSCalendarManager.getInstance();
                const shape = new ItemResponseShapeType();
                shape.BaseShape = DefaultShapeNamesType.DEFAULT;

                let items = await manager.getItems(userInfo, context.request, shape, 0, 10);
                expect(items).to.be.an.Array();
                expect(items.length).to.be.equal(testData.length);
                for (let index = 0; index < items.length; index++) {
                    const item = items[index];
                    expect(item.ItemId?.Id).to.be.equal(getEWSId(testData[index].unid));
                    expect(item.ParentFolderId?.Id).to.be.equal(getEWSId(testData[index].parentFolderIds![0]));
                    expect(item.ItemClass).to.be.equal('IPM.Appointment');
                    expect(item.Subject).to.be.equal(testData[index].subject);
                    expect(item.UID).to.be.equal(testData[index].getAdditionalProperty(UnindexedFieldURIType.CALENDAR_UID));
                    expect(item.Organizer?.Mailbox.EmailAddress).to.be.equal(testData[index].organizer);
                    expect(item.Location).to.be.equal(testData[index].location);

                    // Should not be set
                    expect(item.Body).to.be.undefined();
                    expect(item.Categories).to.be.undefined();
                    expect(item.ReminderMinutesBeforeStart).to.be.undefined();
                    expect(item.RequiredAttendees).to.be.undefined();
                    expect(item.OptionalAttendees).to.be.undefined();
                    expect((item.ExtendedProperty ?? []).length).to.be.equal(0);
                }

                // No items returned
                items = await manager.getItems(userInfo, context.request, shape, 0, 10, generateTestCalendarLabel(true));
                expect(items).to.be.an.Array();
                expect(items.length).to.be.equal(0);
            });

            it('with Shape All', async () => {

                const manager = EWSCalendarManager.getInstance();
                const shape = new ItemResponseShapeType();
                shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
                const parentLabel = generateTestCalendarLabel();
                const mailboxId = 'test@test.com';

                let items = await manager.getItems(userInfo, context.request, shape, 0, 10, parentLabel, mailboxId);
                expect(items).to.be.an.Array();
                expect(items.length).to.be.equal(testData.length);
                for (let index = 0; index < items.length; index++) {
                    const item = items[index];
                    const itemEWSId = getEWSId(testData[index].unid, mailboxId);
                    expect(item.ItemId?.Id).to.be.equal(itemEWSId);
                    const parentEWSId = getEWSId(parentLabel.folderId, mailboxId);
                    expect(item.ParentFolderId?.Id).to.be.equal(parentEWSId);
                    expect(item.ItemClass).to.be.equal('IPM.Appointment');
                    expect(item.Subject).to.be.equal(testData[index].subject);
                    expect(item.Body?.Value).to.be.equal(testData[index].body);
                    expect(item.Organizer?.Mailbox.EmailAddress).to.be.equal(testData[index].organizer);
                    expect(item.Location).to.be.equal(testData[index].location);
                    expect(item.Categories?.String).to.be.deepEqual(testData[index].categories);
                    expect(parseInt(item.ReminderMinutesBeforeStart!)).to.be.equal(Math.abs(testData[index].alarm!));
                    expect(item.RequiredAttendees?.Attendee.map(attendee => attendee.Mailbox.EmailAddress)).to.be.deepEqual(testData[index].requiredAttendees);
                    expect(item.OptionalAttendees?.Attendee.map(attendee => attendee.Mailbox.EmailAddress)).to.be.deepEqual(testData[index].optionalAttendees);
                    expect(new Date(item.Start!).getTime() - new Date(testData[index].start!).getTime()).to.be.equal(0);
                    expect(item.UID).to.be.equal(testData[index].getAdditionalProperty(UnindexedFieldURIType.CALENDAR_UID));
                    const itemProp = item.ExtendedProperty![0];
                    const identifier: any = {};
                    identifier[ExtendedPropertyKeyType.PROPERTY_SETID] = "A7B529B5-4B75-47A7-A24F-20743D6C55CD";
                    identifier[ExtendedPropertyKeyType.PROPERTY_NAME] = "CalendarTimeZone";
                    identifier[ExtendedPropertyKeyType.PROPERTY_TYPE] = MapiPropertyTypeType.STRING;
                    const pimProp = testData[0].findExtendedProperty(identifier);
                    expect(pimProp).to.not.be.undefined();
                    expect(itemProp.ExtendedFieldURI.PropertySetId).to.be.equal(identifier[ExtendedPropertyKeyType.PROPERTY_SETID]);
                    expect(itemProp.ExtendedFieldURI.PropertyName).to.be.equal(identifier[ExtendedPropertyKeyType.PROPERTY_NAME]);
                    expect(itemProp.ExtendedFieldURI.PropertyType).to.be.equal(identifier[ExtendedPropertyKeyType.PROPERTY_TYPE]);
                    expect(itemProp.Value).to.be.equal('America/New_York');

                }

                // No items returned
                items = await manager.getItems(userInfo, context.request, shape, 0, 10, generateTestCalendarLabel(true));
                expect(items).to.be.an.Array();
                expect(items.length).to.be.equal(0);

            });
        });

        it('getCalendarEntries throws an error', async () => {
            const calManagerStub = sinon.createStubInstance(KeepPimCalendarManager);
            calManagerStub.getCalendarItems.rejects();
            stubCalendarManager(calManagerStub);

            const manager = EWSCalendarManager.getInstance();
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;

            const items = await manager.getItems(userInfo, context.request, shape, 0, 10, generateTestCalendarLabel());
            expect(items).to.be.an.Array();
            expect(items.length).to.be.equal(0);
        });

    });

    describe('Test updateItem', function () {

        it('updateItem with set', async function () {
            const toLabel = generateTestCalendarLabel();
            const pimItem = generateTestCalendarItems(toLabel.calendarName!, 1)[0];
            pimItem.requiredAttendees = ["req.email1@test.com", "req.email2@test.com"];
            pimItem.optionalAttendees = ["opt.email1@test.com", "opt.email2@test.com", "opt.email3@test.com"];
            pimItem.start = '2020-03-18T13:58:51-04:00';
            pimItem.startTimeZone = 'America/New_York';
            pimItem.endTimeZone = 'America/New_York';
            pimItem.duration = 'PT1H'
            pimItem.location = "Conference Room A";

            // Updated values
            const expectedStartTime = '2020-03-19T18:58:51.000-04:00';
            const expectedEndTime = '2020-03-19T19:58:51.000-04:00';
            const expectedLocation = "Conference Room B";
            const expectedRequired = ['req.email1@test.com', 'req.email3@test.com'];
            const expectedOptional = ['opt.email1@test.com'];

            const calManagerStub = sinon.createStubInstance(KeepPimCalendarManager);
            calManagerStub.modifyCalendarItem.withArgs(toLabel.calendarName!, pimItem.unid,
                sinon.match((value) => {
                    console.log(`modifyCalendarItem called with ${util.inspect(value, false, 5)}`)
                    // This will validate the code in protected function modifyCalendarStructure
                    const calItem = new PimCalendarItem(value, 'default');
                    
                    if (!Array.isArray(calItem.optionalAttendees) ||
                        !Array.isArray(calItem.requiredAttendees)) {
                        console.error(`Optional or required Attendees not an array`);
                        return false; 
                    }

                    // Verify optional attendees
                    let attendees = calItem.optionalAttendees.filter((a: string) => expectedOptional.includes(a));
                    if (attendees.length !== expectedOptional.length) {
                        console.error(`Optional Attendees not correct: ${calItem.optionalAttendees}`);
                        return false; 
                    }
                    
                    // Verify required attendees
                    attendees = calItem.requiredAttendees.filter((a: string) => expectedRequired.includes(a));
                    if (attendees.length !== expectedRequired.length) {
                        console.error(`Required Attendees not correct: ${calItem.requiredAttendees}`);
                        return false; 
                    }
                    
                    // Verify removed
                    const expected = ['req.email2@test.com', 'opt.email2@test.com', 'opt.email3@test.com'];
                    let combinedAttendees: string[] = [];
                    combinedAttendees = combinedAttendees.concat(calItem.optionalAttendees, calItem.requiredAttendees, calItem.fyiAttendees);
                    attendees = expected.filter((a: string) => !combinedAttendees.includes(a));
                    if (attendees.length !== expected.length) {
                        console.error(`Attendees were not removed.\nExpected: ${expected}, actual: ${attendees}`);
                        return false; 
                    }
                    
                    // Verify start date
                    if (calItem.start !== undefined && (new Date(calItem.start).getTime() - new Date(expectedStartTime).getTime()) !== 0) {
                        console.error(`Start date is not correct: ${calItem.start}`);
                        return false; 
                    }
                    
                    // Verify end date
                    if (calItem.end !== undefined && (new Date(calItem.end).getTime() - new Date(expectedEndTime).getTime()) !== 0) {
                        console.error(`End date is not correct: ${calItem.end}`);
                        return false; 
                    }

                    return true; 
                }),
                sinon.match.any).resolves();
            calManagerStub.modifyCalendarItem.callThrough();

            calManagerStub.updateCalendarItem.withArgs(
                sinon.match((value) => {
                    console.log(`updateCalendarItem called with ${util.inspect(value, false, 5)}`)
                    // This will validate that the PimItem was updated correctly
                    if (!(value instanceof PimCalendarItem)) {
                        console.error('Items is not a calendar item');
                        return false; 
                    }
                    if (value.location !== expectedLocation) {
                        console.error(`Calendar item location is not correct. Expected: ${expectedLocation}, actual: ${value.location}`);
                        return false; 
                    }
                    if (value.start !== expectedStartTime) {
                        console.error(`Calendar item start time is not correct. Expected: ${expectedStartTime}, actual: ${value.start}`);
                        return false; 
                    }
                    if (value.end !== expectedEndTime) {
                        console.error(`Calendar item end time is not correct. Expected: ${expectedEndTime}, actual: ${value.end}`);
                        return false; 
                    }
                    if (value.requiredAttendees.filter(email => !expectedRequired.includes(email)).length !== 0) {
                        console.error(`Required attendees is not correct. Expected: ${expectedRequired}, actual: ${value.requiredAttendees}`);
                        return false; 
                    }
                    if (value.optionalAttendees.filter(email => !expectedOptional.includes(email)).length !== 0) {
                        console.error(`Required attendees is not correct. Expected: ${expectedOptional}, actual: ${value.optionalAttendees}`);
                        return false; 
                    }

                    const identifier: any = {};
                    identifier[ExtendedPropertyKeyType.PROPERTY_SETID] = "A7B529B5-4B75-47A7-A24F-20743D6C55CD";
                    identifier[ExtendedPropertyKeyType.PROPERTY_NAME] = "CalendarTimeZone";
                    const tz = pimItem.findExtendedProperty(identifier);
                    if (tz.Value !== 'America/New_York') {
                        console.error(`CalendarTimeZone is not correct: ${tz.Value}`);
                        return false; 
                    }
                    
                    return true;
                }),
                sinon.match.any).resolves();
            calManagerStub.updateCalendarItem.callThrough();
            stubCalendarManager(calManagerStub);

            const manager = EWSCalendarManager.getInstance();

            const change = new ItemChangeType();
            change.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();
            let description = new SetItemFieldType();
            description.FieldURI = new PathToUnindexedFieldType();
            description.FieldURI.FieldURI = UnindexedFieldURIType.CALENDAR_START;
            description.CalendarItem = new CalendarItemType();
            description.CalendarItem.Start = new Date(expectedStartTime);
            description.CalendarItem.StartString = expectedStartTime; 
            change.Updates.push(description);

            description = new SetItemFieldType();
            description.FieldURI = new PathToUnindexedFieldType();
            description.FieldURI.FieldURI = UnindexedFieldURIType.CALENDAR_END;
            description.CalendarItem = new CalendarItemType();
            description.CalendarItem.End = new Date(expectedEndTime);
            description.CalendarItem.EndString = expectedEndTime;
            change.Updates.push(description);

            description = new SetItemFieldType();
            description.FieldURI = new PathToUnindexedFieldType();
            description.FieldURI.FieldURI = UnindexedFieldURIType.CALENDAR_REQUIRED_ATTENDEES;
            description.CalendarItem = new CalendarItemType();
            description.CalendarItem.RequiredAttendees = new NonEmptyArrayOfAttendeesType();
            for (const email of expectedRequired) {
                const attendee = new AttendeeType();
                attendee.Mailbox = new EmailAddressType();
                attendee.Mailbox.EmailAddress = email;
                description.CalendarItem.RequiredAttendees.Attendee.push(attendee);
            }
            change.Updates.push(description);

            description = new SetItemFieldType();
            description.FieldURI = new PathToUnindexedFieldType();
            description.FieldURI.FieldURI = UnindexedFieldURIType.CALENDAR_OPTIONAL_ATTENDEES;
            description.CalendarItem = new CalendarItemType();
            description.CalendarItem.OptionalAttendees = new NonEmptyArrayOfAttendeesType();
            for (const email of expectedOptional) {
                const attendee = new AttendeeType();
                attendee.Mailbox = new EmailAddressType();
                attendee.Mailbox.EmailAddress = email;
                description.CalendarItem.OptionalAttendees.Attendee.push(attendee);
            }
            change.Updates.push(description);

            description = new SetItemFieldType();
            description.FieldURI = new PathToUnindexedFieldType();
            description.FieldURI.FieldURI = UnindexedFieldURIType.CALENDAR_LOCATION;
            description.CalendarItem = new CalendarItemType();
            description.CalendarItem.Location = expectedLocation;
            change.Updates.push(description);

            description = new SetItemFieldType();
            description.ExtendedFieldURI = new PathToExtendedFieldType();
            description.ExtendedFieldURI.PropertyName = "CalendarTimeZone";
            description.ExtendedFieldURI.PropertySetId = "A7B529B5-4B75-47A7-A24F-20743D6C55CD";
            description.ExtendedFieldURI.PropertyType = MapiPropertyTypeType.STRING;
            description.CalendarItem = new CalendarItemType();
            const prop = new ExtendedPropertyType();
            prop.ExtendedFieldURI = description.ExtendedFieldURI;
            prop.Value = 'America/New_York';
            description.CalendarItem.ExtendedProperty = [prop];
            change.Updates.push(description);

            const mailboxId = 'test@test.com';

            const result = await manager.updateItem(pimItem, change, userInfo, context.request, toLabel, mailboxId);

            // Validate the returned results
            expect(result).to.not.be.undefined();
            const itemEWSId = getEWSId(pimItem.unid, mailboxId);
            expect(result?.ItemId?.Id).to.be.equal(itemEWSId);
            const parentEWSId = getEWSId(toLabel.folderId, mailboxId);
            expect(result?.ParentFolderId?.Id).to.be.equal(parentEWSId);

            // Validate Keep PIM functions called
            expect(calManagerStub.modifyCalendarItem.calledOnce).to.be.true();
            expect(calManagerStub.updateCalendarItem.calledOnce).to.be.true();

        });

        it('updateItem with append', async function () {
            const toLabel = generateTestCalendarLabel();
            const pimItem = generateTestCalendarItems(toLabel.calendarName!, 1)[0];
            pimItem.body = 'Original body.';
            pimItem.requiredAttendees = ["req.email1@test.com", "req.email2@test.com"];
            pimItem.optionalAttendees = ["opt.email1@test.com"];

            const expectedBody = pimItem.body + ' This is new!';
            const expectedRequired = pimItem.requiredAttendees;
            expectedRequired.push('req.email3@test.com');
            const expectedOptional = pimItem.optionalAttendees;
            expectedOptional.push('opt.email2@test.com');
            expectedOptional.push('opt.email3@test.com');


            const calManagerStub = sinon.createStubInstance(KeepPimCalendarManager);
            calManagerStub.modifyCalendarItem.withArgs(toLabel.calendarName!, pimItem.unid,
                sinon.match((value) => {
                    // This will validate the code in protected function modifyCalendarStructure
                    let valid = true;
                    const calItem = new PimCalendarItem(value, 'default');
                    // Verify optional attendees
                    valid = Array.isArray(calItem.optionalAttendees) &&
                        Array.isArray(calItem.requiredAttendees);
                    if (valid) {
                        const attendees = calItem.optionalAttendees.filter((a: string) => expectedOptional.includes(a));
                        valid = (attendees.length === expectedOptional.length);
                    }
                    // Verify required attendees
                    if (valid) {
                        const attendees = calItem.requiredAttendees.filter((a: string) => expectedRequired.includes(a));
                        valid = (attendees.length === expectedRequired.length);
                    }

                    if (!valid) {
                        console.log(`Modify Structure is not valid: ${util.inspect(value, false, 5)}`);
                    }
                    return valid;
                }),
                sinon.match.any).resolves();
            calManagerStub.modifyCalendarItem.callThrough();

            calManagerStub.updateCalendarItem.withArgs(
                sinon.match((value) => {
                    // This will validate that the PimItem was updated correctly
                    const valid = value instanceof PimCalendarItem &&
                        value.body === expectedBody &&
                        value.requiredAttendees.filter(email => !expectedRequired.includes(email)).length === 0 &&
                        value.optionalAttendees.filter(email => !expectedOptional.includes(email)).length === 0;

                    if (!valid) {
                        console.log(`Pim Item is not valid: ${util.inspect(value, false, 5)}`);
                    }
                    return valid;
                }),
                sinon.match.any).resolves();
            calManagerStub.updateCalendarItem.callThrough();
            stubCalendarManager(calManagerStub);

            const manager = EWSCalendarManager.getInstance();

            const change = new ItemChangeType();
            change.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();
            let description = new AppendToItemFieldType();
            description.CalendarItem = new CalendarItemType();
            description.CalendarItem.Body = new BodyType(' This is new!', BodyTypeType.TEXT);
            description.FieldURI = new PathToUnindexedFieldType();
            description.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_BODY;
            change.Updates.push(description);

            description = new AppendToItemFieldType();
            description.FieldURI = new PathToUnindexedFieldType();
            description.FieldURI.FieldURI = UnindexedFieldURIType.CALENDAR_REQUIRED_ATTENDEES;
            description.CalendarItem = new CalendarItemType();
            description.CalendarItem.RequiredAttendees = new NonEmptyArrayOfAttendeesType();
            const newAttendee = new AttendeeType();
            newAttendee.Mailbox = new EmailAddressType();
            newAttendee.Mailbox.EmailAddress = 'req.email3@test.com';
            description.CalendarItem.RequiredAttendees.Attendee.push(newAttendee);
            change.Updates.push(description);

            description = new AppendToItemFieldType();
            description.FieldURI = new PathToUnindexedFieldType();
            description.FieldURI.FieldURI = UnindexedFieldURIType.CALENDAR_OPTIONAL_ATTENDEES;
            description.CalendarItem = new CalendarItemType();
            description.CalendarItem.OptionalAttendees = new NonEmptyArrayOfAttendeesType();
            for (const email of ['opt.email2@test.com', 'opt.email3@test.com']) {
                const attendee = new AttendeeType();
                attendee.Mailbox = new EmailAddressType();
                attendee.Mailbox.EmailAddress = email;
                description.CalendarItem.OptionalAttendees.Attendee.push(attendee);
            }
            change.Updates.push(description);

            const result = await manager.updateItem(pimItem, change, userInfo, context.request, toLabel);

            // Validate the returned results
            expect(result).to.not.be.undefined();
            expect(result?.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));
            expect(result?.ParentFolderId?.Id).to.be.equal(getEWSId(toLabel.folderId));

            // Validate Keep PIM functions called
            expect(calManagerStub.modifyCalendarItem.calledOnce).to.be.true();
            expect(calManagerStub.updateCalendarItem.calledOnce).to.be.true();

        });

        it('updateItem with delete', async function () {
            const toLabel = generateTestCalendarLabel();
            const pimItem = generateTestCalendarItems(toLabel.calendarName!, 1)[0];
            pimItem.location = 'Conference Room B';
            const propertyIdentifiers: any = {};
            propertyIdentifiers[ExtendedPropertyKeyType.DISTINGUISHED_PROPERTY_SETID] = "Common";
            propertyIdentifiers[ExtendedPropertyKeyType.PROPERTY_ID] = "34144";
            propertyIdentifiers[ExtendedPropertyKeyType.PROPERTY_TYPE] = "SystemTime";
            pimItem.addExtendedProperty({ ...propertyIdentifiers, Value: "2020-09-29T17:15:00-04:00" });

            const calManagerStub = sinon.createStubInstance(KeepPimCalendarManager);
            calManagerStub.updateCalendarItem.withArgs(
                sinon.match((value) => {
                    // This will validate that the PimItem was updated correctly
                    const valid = value instanceof PimCalendarItem &&
                        value.location === undefined &&
                        value.findExtendedProperty(propertyIdentifiers) === undefined;

                    if (!valid) {
                        console.log(`Pim Item is not valid: ${util.inspect(value, false, 5)}`);
                    }
                    return valid;
                }),
                sinon.match.any).resolves();
            calManagerStub.updateCalendarItem.callThrough();
            stubCalendarManager(calManagerStub);

            const manager = EWSCalendarManager.getInstance();

            const change = new ItemChangeType();
            change.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();
            let description = new DeleteItemFieldType();
            description.FieldURI = new PathToUnindexedFieldType();
            description.FieldURI.FieldURI = UnindexedFieldURIType.CALENDAR_LOCATION;
            change.Updates.push(description);

            description = new DeleteItemFieldType();
            description.ExtendedFieldURI = new PathToExtendedFieldType();
            description.ExtendedFieldURI.DistinguishedPropertySetId = propertyIdentifiers[ExtendedPropertyKeyType.DISTINGUISHED_PROPERTY_SETID];
            description.ExtendedFieldURI.PropertyId = propertyIdentifiers[ExtendedPropertyKeyType.PROPERTY_ID];
            description.ExtendedFieldURI.PropertyType = propertyIdentifiers[ExtendedPropertyKeyType.PROPERTY_TYPE];
            change.Updates.push(description);

            const result = await manager.updateItem(pimItem, change, userInfo, context.request, toLabel);

            // Validate the returned results
            expect(result).to.not.be.undefined();
            expect(result?.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));
            expect(result?.ParentFolderId?.Id).to.be.equal(getEWSId(toLabel.folderId));

            // Validate Keep PIM functions called
            expect(calManagerStub.modifyCalendarItem.called).to.be.false();
            expect(calManagerStub.updateCalendarItem.calledOnce).to.be.true();

        });
    });


    describe('EWSCalendarManager.modifyCalendarStructure', function () {
        let pimItem: PimCalendarItem;

        beforeEach(function () {

            const startDate = "2021-12-25T15:00:00.000Z";
            const endDate = "2021-12-25T16:00:00.000Z";
            const createdDate = new Date("2020-12-25T15:00:00.000Z");
            pimItem = PimItemFactory.newPimCalendarItem({}, 'default', PimItemFormat.DOCUMENT);
            pimItem.unid = 'test-cal-item';
            pimItem.subject = 'Unit Test';
            pimItem.start = startDate;
            pimItem.end = endDate;
            pimItem.createdDate = createdDate;
            pimItem.requiredAttendees = ["tester1@ut.com", "tester2@ut.com"];
            pimItem.optionalAttendees = ["tester3@ut.com", "tester4@ut.com"];
            const prop: any = { 'Value': "America/New_York" };
            prop[ExtendedPropertyKeyType.PROPERTY_NAME] = "CalendarTimeZone";
            prop[ExtendedPropertyKeyType.PROPERTY_SETID] = "A7B529B5-4B75-47A7-A24F-20743D6C55CD";
            prop[ExtendedPropertyKeyType.PROPERTY_TYPE] = MapiPropertyTypeType.STRING;
            pimItem.addExtendedProperty(prop);
        });

        it('modify start/end times', function () {

            // EWS item containing the updates
            const ewsItem = new CalendarItemType();
            const timeZoneName = "America/New_York";
            const startString = "2021-01-15T01:00:00";
            let dt = DateTime.fromISO(startString, {zone: timeZoneName});
            const expectedStart = dt.toJSDate(); 
            ewsItem.Start = expectedStart;
            ewsItem.StartString  = startString;
            const endString = "2021-01-15T02:00:00";
            dt = DateTime.fromISO(endString, {zone: timeZoneName});
            const expectedEnd = dt.toJSDate(); 
            ewsItem.End = expectedEnd;
            ewsItem.EndString = endString; 
            const prop = new ExtendedPropertyType();
            prop.ExtendedFieldURI = new PathToExtendedFieldType();
            prop.ExtendedFieldURI.PropertySetId = "A7B529B5-4B75-47A7-A24F-20743D6C55CD";
            prop.ExtendedFieldURI.PropertyName = "CalendarTimeZone";
            prop.ExtendedFieldURI.PropertyType = MapiPropertyTypeType.STRING;
            prop.Value = timeZoneName;
            ewsItem.ExtendedProperty = [prop];
            ewsItem.MeetingTimeZone = new TimeZoneType();
            ewsItem.MeetingTimeZone.TimeZoneName = "Eastern Standard Time";

            const results = new MockEWSCalendarManager().modifyCalendarStructure(ewsItem, pimItem, []);

            // Verify the returned updates
            
            const updates = new PimCalendarItem(results, 'default');
            expect(updates.start !== undefined && new Date(updates.start).getTime() - expectedStart.getTime()).to.be.eql(0);
            expect(updates.end !== undefined && new Date(updates.end).getTime() - expectedEnd.getTime()).to.be.eql(0);

            // Verify processed values were removed from the EWS item
            expect(ewsItem.Start).to.be.undefined();
            expect(ewsItem.End).to.be.undefined();

            // Verify pim item was updated
            expect(pimItem.start).to.not.be.undefined(); 
            expect(pimItem.end).to.not.be.undefined(); 
            const pimStart = new Date(pimItem.start!);
            const pimEnd = new Date(pimItem.end!); 
            expect(pimStart.getTime() - expectedStart.getTime()).to.be.eql(0);
            expect(pimEnd.getTime() - expectedEnd.getTime()).to.be.eql(0);

        });

        it('modify attendees', function () {
            const ewsItem = new CalendarItemType();

            // Required attendees has a new attendee
            ewsItem.RequiredAttendees = new NonEmptyArrayOfAttendeesType();
            const attendee1 = new AttendeeType();
            attendee1.Mailbox = new EmailAddressType();
            attendee1.Mailbox.EmailAddress = "tester1@ut.com";
            const attendee2 = new AttendeeType();
            attendee2.Mailbox = new EmailAddressType();
            attendee2.Mailbox.EmailAddress = "tester2@ut.com";
            const newAttendee = new AttendeeType();
            newAttendee.Mailbox = new EmailAddressType();
            newAttendee.Mailbox.EmailAddress = "newtester@ut.com";
            ewsItem.RequiredAttendees.Attendee = [attendee1, attendee2, newAttendee];

            // An optional attendee is removed
            ewsItem.OptionalAttendees = new NonEmptyArrayOfAttendeesType();
            const attendee3 = new AttendeeType();
            attendee3.Mailbox = new EmailAddressType();
            attendee3.Mailbox.EmailAddress = "tester3@ut.com";
            ewsItem.OptionalAttendees.Attendee = [attendee3];

            const results = new MockEWSCalendarManager().modifyCalendarStructure(ewsItem, pimItem, ["tester4@ut.com"]);

            const calItem = new PimCalendarItem(results, 'default');
            // Verify removed
            let combinedAttendees: string[] = [];
            combinedAttendees = combinedAttendees.concat(calItem.optionalAttendees, calItem.requiredAttendees, calItem.fyiAttendees);
            expect(combinedAttendees.includes("tester4@ut.com")).to.be.false();

            // Verify updated
            expect(calItem.requiredAttendees).to.containDeep(["newtester@ut.com", "tester1@ut.com", "tester2@ut.com"]);
            expect(calItem.optionalAttendees).to.containDeep(["tester3@ut.com"]);

            // Verify pim item was updated
            expect(pimItem.requiredAttendees).to.containDeep(["newtester@ut.com", "tester1@ut.com", "tester2@ut.com"]);
            expect(pimItem.optionalAttendees).to.containDeep(["tester3@ut.com"]);

        });
    });

    describe('Test createItem', function () {
        beforeEach(() => {
            const calManagerStub = sinon.createStubInstance(KeepPimCalendarManager);
            calManagerStub.createCalendarItem.resolves('UNID');
            stubCalendarManager(calManagerStub);
        });

        it('createItem without a label', async function () {
            const manager = EWSCalendarManager.getInstance();
            const item = new CalendarItemType();
            item.ItemId = new ItemIdType('UNID');
            const results = await manager.createItem(item, userInfo, context.request);
            expect(results.length).to.be.equal(1);
            expect(results[0].ItemId?.Id).to.be.equal(getEWSId('UNID'));
            expect(results[0].ItemId?.ChangeKey).to.not.be.undefined();

        });

        it('createItem with a label', async function () {
            const msgManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            msgManagerStub.moveMessages.resolves({ movedIds: [{ unid: 'UNID', status: 200, message: 'success' }] })
            stubMessageManager(msgManagerStub);

            const manager = EWSCalendarManager.getInstance();
            const item = new CalendarItemType();
            item.ItemId = new ItemIdType('UNID');
            const label = generateTestCalendarLabel(true);
            const mailboxId = 'test@test.com';

            const results = await manager.createItem(item, userInfo, context.request, MessageDispositionType.SAVE_ONLY, label, mailboxId);
            
            expect(results.length).to.be.equal(1);
            const itemEWSId = getEWSId('UNID', mailboxId);
            expect(results[0].ItemId?.Id).to.be.equal(itemEWSId);
            expect(results[0].ItemId?.ChangeKey).to.not.be.undefined();
        });

        it('createItem with createCalendarItem error', async function () {
            sinon.restore();
            const error = new Error('Unit Test');
            const calManagerStub = sinon.createStubInstance(KeepPimCalendarManager);
            calManagerStub.createCalendarItem.rejects(error);
            stubCalendarManager(calManagerStub);

            const manager = EWSCalendarManager.getInstance();
            const item = new CalendarItemType();
            item.ItemId = new ItemIdType('UNID');
            await expect(manager.createItem(item, userInfo, context.request)).to.be.rejectedWith(error);

        });

        it('createItem and verify pimItemFromEWSItem', async function () {
            sinon.restore();

            const expectedItemUNID = 'item-UNID';
            const expectedItemEWSId = getEWSId(expectedItemUNID);
            const expectedRequired = ['req.email1@test.com', 'req.email2@test.com'];
            const expectedOptional = ['opt.email1@test.com'];
            const expectedSubject = 'Unit Test';
            const expectedBody = 'This is Unit Test';
            const expectedOrganizer = 'organizer@test.com';
            const expectedLocation = 'Conference Room';
            const expectedStart = '2020-03-18T13:58:51-04:00';
            const expectedCategories = ["cat1", "cat2"];
            const expectedAlarm = 5;
            const expectedCalUID = 'CAL-UID';

            const calManagerStub = sinon.createStubInstance(KeepPimCalendarManager);
            calManagerStub.createCalendarItem.withArgs(
                sinon.match((value) => {
                    if (value instanceof PimCalendarItem) {
                        let valid = value.calendarName === KeepPimConstants.DEFAULT_CALENDAR_NAME &&
                            value.subject === expectedSubject &&
                            value.importance === PimImportance.HIGH &&
                            value.isPrivate &&
                            value.isDraft &&
                            value.body === expectedBody &&
                            new Date(value.start ?? "").getTime() - new Date(expectedStart).getTime() === 0 &&
                            value.duration === 'PT1H' &&
                            value.isMeeting &&
                            value.requiredAttendees.filter(email => !expectedRequired.includes(email)).length === 0 &&
                            value.optionalAttendees.filter(email => !expectedOptional.includes(email)).length === 0 &&
                            value.organizer === expectedOrganizer &&
                            value.location === expectedLocation &&
                            value.categories.filter(cat => !expectedCategories.includes(cat)).length === 0 &&
                            value.alarm === -expectedAlarm &&
                            value.getAdditionalProperty(UnindexedFieldURIType.CALENDAR_UID) === expectedCalUID;

                        // Validate extended properties
                        if (valid) {
                            const identifier: any = {};
                            identifier[ExtendedPropertyKeyType.PROPERTY_SETID] = "A7B529B5-4B75-47A7-A24F-20743D6C55CD";
                            identifier[ExtendedPropertyKeyType.PROPERTY_NAME] = "CalendarTimeZone";
                            const zone = value.findExtendedProperty(identifier);
                            valid = (zone.Value === 'America/New_York');
                        }

                        if (!valid) {
                            console.log(`Pim Item is not valid: ${util.inspect(value, false, 5)}`);
                        }
                        return valid;
                    }
                    return false;
                }),
                sinon.match.any, true).resolves(expectedItemUNID);
            calManagerStub.createCalendarItem.callThrough();
            stubCalendarManager(calManagerStub);

            // EWS Calendar item 
            const ewsItem = new CalendarItemType();
            ewsItem.ItemId = new ItemIdType(expectedItemUNID);
            ewsItem.Subject = expectedSubject;
            ewsItem.Importance = ImportanceChoicesType.HIGH;
            ewsItem.Sensitivity = SensitivityChoicesType.PRIVATE;
            ewsItem.Body = new BodyType(expectedBody, BodyTypeType.TEXT);
            ewsItem.Start = new Date(expectedStart);
            ewsItem.StartString = expectedStart; 
            const endString = '2020-03-18T14:58:51-04:00';
            ewsItem.End = new Date(endString); // 1H duration
            ewsItem.EndString = endString; 
            const tz = new ExtendedPropertyType();
            tz.ExtendedFieldURI = new PathToExtendedFieldType();
            tz.ExtendedFieldURI.PropertyName = "CalendarTimeZone";
            tz.ExtendedFieldURI.PropertySetId = "A7B529B5-4B75-47A7-A24F-20743D6C55CD";
            tz.ExtendedFieldURI.PropertyType = MapiPropertyTypeType.STRING;
            tz.Value = 'America/New_York';
            ewsItem.ExtendedProperty = [tz];
            ewsItem.MeetingTimeZone = new TimeZoneType();
            ewsItem.MeetingTimeZone.TimeZoneName = tz.Value;
            ewsItem.IsDraft = true;
            ewsItem.RequiredAttendees = new NonEmptyArrayOfAttendeesType();
            for (const email of expectedRequired) {
                const attendee = new AttendeeType();
                attendee.Mailbox = new EmailAddressType();
                attendee.Mailbox.EmailAddress = email;
                ewsItem.RequiredAttendees.Attendee.push(attendee);
            }
            ewsItem.OptionalAttendees = new NonEmptyArrayOfAttendeesType();
            for (const email of expectedOptional) {
                const attendee = new AttendeeType();
                attendee.Mailbox = new EmailAddressType();
                attendee.Mailbox.EmailAddress = email;
                ewsItem.OptionalAttendees.Attendee.push(attendee);
            }
            ewsItem.Organizer = new SingleRecipientType();
            ewsItem.Organizer.Mailbox = new EmailAddressType();
            ewsItem.Organizer.Mailbox.EmailAddress = expectedOrganizer;
            ewsItem.Location = expectedLocation;
            ewsItem.Categories = new ArrayOfStringsType();
            ewsItem.Categories.String = expectedCategories;
            ewsItem.ReminderIsSet = true;
            ewsItem.ReminderMinutesBeforeStart = `${expectedAlarm}`;
            ewsItem.UID = expectedCalUID;

            const manager = EWSCalendarManager.getInstance();
            const results = await manager.createItem(ewsItem, userInfo, context.request);
            expect(results.length).to.be.equal(1);
            expect(results[0].ItemId?.Id).to.be.equal(expectedItemEWSId);
            expect(results[0].ItemId?.ChangeKey).to.not.be.undefined();
        });

        it('createItem with a label and move failed response', async function () {
            // The move messages returns a failed response
            const msgManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            msgManagerStub.moveMessages.resolves({ movedIds: [{ unid: 'UNID', status: 500, message: 'failed' }] });
            stubMessageManager(msgManagerStub);

            const manager = EWSCalendarManager.getInstance();
            const item = new CalendarItemType();
            item.ItemId = new ItemIdType('UNID');
            const label = generateTestCalendarLabel(true);
            await expect(manager.createItem(item, userInfo, context.request, MessageDispositionType.SAVE_ONLY, label)).to.be.rejected();
        });

        it('createItem with a label and move exception', async function () {
            // The move messages throws an exception
            const error = new Error('Unit test');
            const msgManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            msgManagerStub.moveMessages.rejects(error);
            stubMessageManager(msgManagerStub);

            const manager = EWSCalendarManager.getInstance();
            const item = new CalendarItemType();
            item.ItemId = new ItemIdType('UNID');
            const label = generateTestCalendarLabel(true);
            await expect(manager.createItem(item, userInfo, context.request, MessageDispositionType.SAVE_ONLY, label)).to.be.rejectedWith(error);
        });
    });

    describe('Test copyItem', function () {

        it('copyItem success with getPimItem success', async function () {

            const pimItem = generateTestCalendarItems(KeepPimConstants.DEFAULT_CALENDAR_NAME, 1)[0];
            const toCalendar = generateTestCalendarLabel(true);
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            const pimCopy = PimItemFactory.newPimCalendarItem(pimItem.toPimStructure(), KeepPimConstants.DEFAULT_CALENDAR_NAME, PimItemFormat.DOCUMENT);
            pimCopy.unid = 'new-UNID';

            const msgManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            msgManagerStub
                .moveMessages
                .withArgs(sinon.match.any, toCalendar.folderId, undefined, undefined, undefined, sinon.match.array.deepEquals([pimItem.unid]))
                .resolves({ copiedIds: [{ unid: pimCopy.unid, status: 200, message: 'success' }] });
            msgManagerStub.moveMessages.callThrough();
            stubMessageManager(msgManagerStub);

            const stub = sinon.createStubInstance(KeepPimManager);
            stub.getPimItem.resolves(pimCopy);
            stubPimManager(stub);
        
            const manager = EWSCalendarManager.getInstance();
            const mailboxId = 'test@test.com';

            const newPimItem = await manager.copyItem(pimItem, toCalendar.folderId, userInfo, context.request, shape, mailboxId);

            const itemEWSId = getEWSId(pimCopy.unid, mailboxId);
            expect(newPimItem.ItemId?.Id).to.be.equal(itemEWSId);
            const parentEWSId = getEWSId(toCalendar.folderId, mailboxId);
            expect(newPimItem.ParentFolderId?.Id).to.be.equal(parentEWSId);
        });

        it('copyItem fails due to moveMessages failure', async function () {

            const pimItem = generateTestCalendarItems(KeepPimConstants.DEFAULT_CALENDAR_NAME, 1)[0];
            const toCalendar = generateTestCalendarLabel(true);
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;

            const error = new Error('Unit Test');
            const calManagerStub = sinon.createStubInstance(KeepPimCalendarManager);
            stubCalendarManager(calManagerStub);

            const managerStub = sinon.createStubInstance(KeepPimMessageManager);
            managerStub.moveMessages.rejects(error);
            stubMessageManager(managerStub);

            const manager = EWSCalendarManager.getInstance();
            await expect(manager.copyItem(pimItem, toCalendar.folderId, userInfo, context.request, shape)).to.be.rejectedWith(error);
            expect(calManagerStub.getCalendarItem.called).to.be.false();

        });

        it('copyItem fails due to getPimItem failure', async function () {

            const pimItem = generateTestCalendarItems(KeepPimConstants.DEFAULT_CALENDAR_NAME, 1)[0];
            const toCalendar = generateTestCalendarLabel(true);
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            const pimCopy = PimItemFactory.newPimCalendarItem(pimItem.toPimStructure(), KeepPimConstants.DEFAULT_CALENDAR_NAME, PimItemFormat.DOCUMENT);
            pimCopy.unid = 'new-UNID';

            const msgManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            msgManagerStub.moveMessages.withArgs(sinon.match.any, toCalendar.folderId, undefined, undefined, undefined, sinon.match.array.deepEquals([pimItem.unid])).resolves({ copiedIds: [{ unid: pimCopy.unid, status: 200, message: 'success' }] });
            msgManagerStub.moveMessages.callThrough();
            stubMessageManager(msgManagerStub);

            const managerStub = sinon.createStubInstance(KeepPimManager);
            managerStub.getPimItem.rejects();
            stubPimManager(managerStub);

            const manager = EWSCalendarManager.getInstance();
            await expect(manager.copyItem(pimItem, toCalendar.folderId, userInfo, context.request, shape)).to.be.rejected();
            expect(managerStub.getPimItem.called).to.be.true();
        });
    });

    describe('Test deleteItem', function () {
        it('deleteItem with unid', async function () {
            const calManagerStub = sinon.createStubInstance(KeepPimCalendarManager);
            calManagerStub.deleteCalendarItem.withArgs(KeepPimConstants.DEFAULT_CALENDAR_NAME, 'UNID', sinon.match.any).resolves();
            calManagerStub.deleteCalendarItem.callThrough();
            stubCalendarManager(calManagerStub);

            const manager = EWSCalendarManager.getInstance();
            await expect(manager.deleteItem('UNID', userInfo)).to.not.be.rejected();
        });

        it('deleteItem with calendar item', async function () {
            const item = generateTestCalendarItems('My Calendar', 1)[0];
            const calManagerStub = sinon.createStubInstance(KeepPimCalendarManager);
            calManagerStub.deleteCalendarItem.withArgs(item.calendarName, item.unid, sinon.match.any, true).resolves();
            calManagerStub.deleteCalendarItem.callThrough();
            stubCalendarManager(calManagerStub);

            const manager = EWSCalendarManager.getInstance();
            const mailboxId = 'test@test.com';
            await expect(manager.deleteItem(item, userInfo, mailboxId, true)).to.not.be.rejected();
        });

        it('deleteItem error', async function () {
            const error = new Error('Unit Test');
            const calManagerStub = sinon.createStubInstance(KeepPimCalendarManager);
            calManagerStub.deleteCalendarItem.rejects(error);
            stubCalendarManager(calManagerStub);

            const manager = EWSCalendarManager.getInstance();
            const mailboxId = 'test@test.com';
            await expect(manager.deleteItem('UNID', userInfo, mailboxId, true)).to.be.rejectedWith(error);
        });
    });

    describe('EWSCalendarManager.updatePimItemFieldValue', function () {

        it('update calendar item', function () {
            const startDate = "2021-12-25T15:00:00.000Z";
            const endDate = "2021-12-25T16:00:00.000Z";
            const pimItem = PimItemFactory.newPimCalendarItem({}, 'default', PimItemFormat.DOCUMENT);
            pimItem.unid = 'test-cal-item';
            pimItem.subject = 'Unit Test';
            pimItem.start = startDate;
            pimItem.end = endDate;
            pimItem.requiredAttendees = ["tester1@ut.com", "tester2@ut.com"];
            pimItem.optionalAttendees = ["tester3@ut.com", "tester4@ut.com"];
            const prop: any = { 'Value': "America/New_York" };
            prop[ExtendedPropertyKeyType.PROPERTY_NAME] = "CalendarTimeZone";
            prop[ExtendedPropertyKeyType.PROPERTY_SETID] = "A7B529B5-4B75-47A7-A24F-20743D6C55CD";
            prop[ExtendedPropertyKeyType.PROPERTY_TYPE] = MapiPropertyTypeType.STRING;
            pimItem.addExtendedProperty(prop);
            pimItem.setAdditionalProperty(UnindexedFieldURIType.CALENDAR_UID, "unique-cal-id");

            // Update time zone
            const newTZ = new TimeZoneType();
            newTZ.TimeZoneName = "Eastern Standard Time";
            const calManager = new MockEWSCalendarManager();
            calManager.updatePimItemFieldValue(pimItem, UnindexedFieldURIType.CALENDAR_MEETING_TIME_ZONE, newTZ);

            // Update start/end time
            calManager.updatePimItemFieldValue(pimItem, UnindexedFieldURIType.CALENDAR_START, "2021-12-25T11:00:00.000Z");
            calManager.updatePimItemFieldValue(pimItem, UnindexedFieldURIType.CALENDAR_END, "2021-12-25T11:30:00.000Z");

            // Update location
            calManager.updatePimItemFieldValue(pimItem, UnindexedFieldURIType.CALENDAR_LOCATION, "My New Location");

            // Update calendar uid
            calManager.updatePimItemFieldValue(pimItem, UnindexedFieldURIType.CALENDAR_UID, "new-uid");

            // Update meeting type
            calManager.updatePimItemFieldValue(pimItem, UnindexedFieldURIType.CALENDAR_IS_MEETING, true);

            // Update organizer
            calManager.updatePimItemFieldValue(pimItem, UnindexedFieldURIType.CALENDAR_ORGANIZER, "ut.master@ut.com");

            // Update required attendees
            let attendees = new NonEmptyArrayOfAttendeesType();
            let attendee1 = new AttendeeType();
            attendee1.Mailbox = new EmailAddressType()
            attendee1.Mailbox.EmailAddress = "new1@ut.com";
            attendees.Attendee = [attendee1];
            calManager.updatePimItemFieldValue(pimItem, UnindexedFieldURIType.CALENDAR_REQUIRED_ATTENDEES, attendees);

            // Update required attendees
            attendees = new NonEmptyArrayOfAttendeesType();
            attendee1 = new AttendeeType();
            attendee1.Mailbox = new EmailAddressType()
            attendee1.Mailbox.EmailAddress = "new2@ut.com";
            const attendee2 = new AttendeeType();
            attendee2.Mailbox = new EmailAddressType()
            attendee2.Mailbox.EmailAddress = "new3@ut.com";
            attendees.Attendee = [attendee1, attendee2];
            calManager.updatePimItemFieldValue(pimItem, UnindexedFieldURIType.CALENDAR_OPTIONAL_ATTENDEES, attendees);


            // Verify updates
            const expectedStart = new Date("2021-12-25T11:00:00.000Z"); 
            const expectedEnd = new Date("2021-12-25T11:30:00.000Z"); 
            expect(new Date(pimItem.start).getTime() - expectedStart.getTime()).to.be.eql(0);
            expect(new Date(pimItem.end).getTime() - expectedEnd.getTime()).to.be.eql(0);
            expect(pimItem.organizer).to.be.eql("ut.master@ut.com");
            expect(pimItem.getAdditionalProperty(UnindexedFieldURIType.CALENDAR_MEETING_TIME_ZONE)).to.be.eql("Eastern Standard Time");
            expect(pimItem.getAdditionalProperty(UnindexedFieldURIType.CALENDAR_UID)).to.be.eql("new-uid");
            expect(pimItem.location).to.be.eql("My New Location");
            expect(pimItem.isMeeting).to.be.true();
            expect(pimItem.requiredAttendees).to.containDeep(["new1@ut.com"]);
            expect(pimItem.optionalAttendees).to.containDeep(["new2@ut.com", "new3@ut.com"]);

            let pimObject: any = {
                uid: 'unit-test-id',
                start: new Date(),
                duration: 'P1D',
            }

            let pimCalendarItem = PimItemFactory.newPimCalendarItem(pimObject, "default", PimItemFormat.DOCUMENT);
            expect(pimCalendarItem.isAllDayEvent).to.be.false();
            calManager.updatePimItemFieldValue(pimCalendarItem, UnindexedFieldURIType.CALENDAR_IS_ALL_DAY_EVENT, true);
            expect(pimCalendarItem.isAllDayEvent).to.be.true();

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

            pimCalendarItem = PimItemFactory.newPimCalendarItem(pimObject, "default", PimItemFormat.DOCUMENT);
            calManager.updatePimItemFieldValue(pimCalendarItem, UnindexedFieldURIType.CALENDAR_IS_ALL_DAY_EVENT, true);
            expect(pimCalendarItem.isAllDayEvent).to.be.false();

            pimObject = {
                uid: 'unit-test-id',
                start: new Date(),
                duration: 'P1D',
                showWithoutTime: true
            }

            pimCalendarItem = PimItemFactory.newPimCalendarItem(pimObject, "default", PimItemFormat.DOCUMENT);
            expect(pimCalendarItem.isAllDayEvent).to.be.true();
            calManager.updatePimItemFieldValue(pimCalendarItem, UnindexedFieldURIType.CALENDAR_IS_ALL_DAY_EVENT, false);
            expect(pimCalendarItem.isAllDayEvent).to.be.false();

        });
    });

    describe('Test recurrence', () => {

        const shape = new ItemResponseShapeType();
        shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
        const additions = [UnindexedFieldURIType.CALENDAR_RECURRENCE, UnindexedFieldURIType.CALENDAR_FIRST_OCCURRENCE,
        UnindexedFieldURIType.CALENDAR_LAST_OCCURRENCE, UnindexedFieldURIType.CALENDAR_IS_RECURRING];
        shape.AdditionalProperties = new NonEmptyArrayOfPathsToElementType(additions.map(type => {
            const fieldType = new PathToUnindexedFieldType();
            fieldType.FieldURI = type;
            return fieldType;
        }));

        describe('Unsupported rules', () => {
            it('multiple rules', async () => {
                // Start of first event
                const startString = '2021-05-11T09:00:00';
                const jsCalendar =
                {
                    "uid": "test-uid",
                    "title": "multiple rules Unit Test",
                    "start": startString,
                    "timeZone": "America/Los_Angeles",
                    "duration": "PT1H30M",
                    recurrenceRules: [{
                        "@type": "RecurrenceRule",
                        "frequency": "yearly",
                        "byMonthDay": [11],
                        "byMonth": ['5'],
                        "count": 2
                    },
                    {
                        "@type": "RecurrenceRule",
                        "frequency": "monthly",
                        "byMonthDay": [11],
                        "interval": 2,
                        "count": 3
                    }]
                };

                const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
                const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                expect(manager).to.not.be.undefined();

                await expect(manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape)).to.be.rejected();

            });

            it('unsupported byXXXX settings', async function () {
                // Start of first event
                const startString = '2021-05-11T09:00:00';
                const jsCalendars = [
                    {
                        "uid": "test-uid",
                        "title": "byHour Unit Test",
                        "start": startString,
                        "timeZone": "America/Los_Angeles",
                        "duration": "PT1H30M",
                        recurrenceRules: [{
                            "@type": "RecurrenceRule",
                            "frequency": "daily",
                            "byHour": [9, 10, 11],
                            "count": 50
                        }]
                    },
                    {
                        "uid": "test-uid",
                        "title": "byMinute Unit Test",
                        "start": startString,
                        "timeZone": "America/Los_Angeles",
                        "duration": "PT1H30M",
                        recurrenceRules: [{
                            "@type": "RecurrenceRule",
                            "frequency": "daily",
                            "byMinute": [45, 50, 55],
                            "count": 50
                        }]
                    },
                    {
                        "uid": "test-uid",
                        "title": "bySecond Unit Test",
                        "start": startString,
                        "timeZone": "America/Los_Angeles",
                        "duration": "PT1H30M",
                        recurrenceRules: [{
                            "@type": "RecurrenceRule",
                            "frequency": "daily",
                            "bySecond": [30],
                            "count": 50
                        }]
                    },
                    {
                        "uid": "test-uid",
                        "title": "byWeekNo Unit Test",
                        "start": startString,
                        "timeZone": "America/Los_Angeles",
                        "duration": "PT1H30M",
                        recurrenceRules: [{
                            "@type": "RecurrenceRule",
                            "frequency": "daily",
                            "byWeekNo": [4, 6],
                            "count": 50
                        }]
                    },
                    {
                        "uid": "test-uid",
                        "title": "byYearDay Unit Test",
                        "start": startString,
                        "timeZone": "America/Los_Angeles",
                        "duration": "PT1H30M",
                        recurrenceRules: [{
                            "@type": "RecurrenceRule",
                            "frequency": "daily",
                            "byYearDay": [130, 220],
                            "count": 50
                        }]
                    },
                ];

                for (const jsCalendar of jsCalendars) {
                    console.log(`${this.test?.title}: Testing item: ${util.inspect(jsCalendar, false, 10)}`);
                    const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
                    const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                    expect(manager).to.not.be.undefined();

                    await expect(manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape)).to.be.rejected();
                }
            });

            it('invalid bySetPosition', async function () {
                // Start of first event
                const startString = '2021-05-11T09:00:00';
                const jsCalendars = [
                    {
                        "uid": "test-uid",
                        "title": "Monthly Unit Test",
                        "start": startString,
                        "duration": "PT1H00M",
                        recurrenceRules: [{
                            "@type": "RecurrenceRule",
                            "frequency": "monthly",
                            "byDay": [{ "@type": "NDay", day: "tu" }],
                            "bySetPosition": [-2],
                            "count": 7
                        }]
                    },
                    {
                        "uid": "test-uid",
                        "title": "Monthly Unit Test",
                        "start": startString,
                        "duration": "PT1H00M",
                        recurrenceRules: [{
                            "@type": "RecurrenceRule",
                            "frequency": "monthly",
                            "byDay": [{ "@type": "NDay", day: "tu" }],
                            "bySetPosition": [-1, 1],
                            "count": 7
                        }]
                    }
                ];

                for (const jsCalendar of jsCalendars) {
                    console.log(`${this.test?.title}: Testing item: ${util.inspect(jsCalendar, false, 10)}`);
                    const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
                    const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                    expect(manager).to.not.be.undefined();

                    await expect(manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape)).to.be.rejected();
                }
            });

            it('invalid monthly', async () => {
                // Start of first event
                const startString = '2021-05-11T09:00:00';
                const jsCalendar =
                {
                    "uid": "test-uid",
                    "title": "Monthly Unit Test",
                    "start": startString,
                    "duration": "PT1H00M",
                    recurrenceRules: [{
                        "@type": "RecurrenceRule",
                        "frequency": "monthly",
                        "count": 7
                    }]
                };

                const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
                const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                expect(manager).to.not.be.undefined();

                await expect(manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape)).to.be.rejected();
            });

            it('invalid monthly rules', async function () {
                // Start of first event
                const startString = '2021-05-11T09:00:00';
                const jsCalendars = [
                    {
                        "uid": "test-uid",
                        "title": "Monthly Unit Test",
                        "start": startString,
                        "duration": "PT1H00M",
                        recurrenceRules: [{
                            "@type": "RecurrenceRule",
                            "frequency": "monthly",
                            "byMonth": ["13"],
                            "count": 7
                        }]
                    },
                    {
                        "uid": "test-uid",
                        "title": "Monthly Unit Test",
                        "start": startString,
                        "duration": "PT1H00M",
                        recurrenceRules: [{
                            "@type": "RecurrenceRule",
                            "frequency": "monthly",
                            "byMonth": ["7"],
                            "byDay": [{ "@type": "NDay", day: "tu" }, { "@type": "NDay", day: "we" }, { "@type": "NDay", day: "fr" }],
                            "count": 7
                        }]
                    }
                ];

                for (const jsCalendar of jsCalendars) {
                    console.log(`${this.test?.title}: Testing item: ${util.inspect(jsCalendar, false, 10)}`);
                    const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
                    const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                    expect(manager).to.not.be.undefined();

                    await expect(manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape)).to.be.rejected();
                }
            });

            it('invalid yearly', async () => {
                // Start of first event
                const startString = '2021-05-11T09:00:00';

                const jsCalendar = {
                    "uid": "test-uid",
                    "title": "Yearly Unit Test",
                    "start": startString,
                    "timeZone": "UTC",
                    "duration": "PT1H30M",
                    recurrenceRules: [{
                        "@type": "RecurrenceRule",
                        "frequency": "yearly",
                        "count": 2
                    }]
                };

                const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
                const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                expect(manager).to.not.be.undefined();
                const mailboxId = 'test@test.com';

                await expect(manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape, mailboxId)).to.be.rejected();
            });

            it('invalid weekly', async () => {
                // Start of first event
                const startString = '2021-05-04T09:00:00Z';

                const jsCalendar = {
                    "uid": "test-uid",
                    "title": "Weely Unit Test",
                    "start": startString,
                    "duration": "PT1H00M",
                    recurrenceRules: [{
                        "@type": "RecurrenceRule",
                        "frequency": "weekly",
                        "count": 9
                    }]
                };

                const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
                const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                expect(manager).to.not.be.undefined();

                await expect(manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape)).to.be.rejected();
            });

            it('Unsupported frequency', async () => {
                // Start of first event
                const startString = '2021-05-11T09:00:00';
                const jsCalendar =
                {
                    "uid": "test-uid",
                    "title": "Unsupported Frequency Unit Test",
                    "start": startString,
                    "duration": "PT1H00M",
                    recurrenceRules: [{
                        "@type": "RecurrenceRule",
                        "frequency": "hourly",
                        "count": 7
                    }]
                };

                const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
                const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                expect(manager).to.not.be.undefined();

                await expect(manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape)).to.be.rejected();
            });
        });

        describe('Yearly recurrences', () => {
            it('Yearly on day of month', async function () {
                // Start of first event
                const startString = '2021-05-11T09:00:00';

                const jsCalendars = [
                    {
                        "uid": "test-uid",
                        "title": "Yearly Unit Test ",
                        "start": startString,
                        "timeZone": "UTC",
                        "duration": "PT1H30M",
                        recurrenceRules: [{
                            "@type": "RecurrenceRule",
                            "frequency": "yearly",
                            "byMonthDay": [11],
                            "byMonth": ['5'],
                            "count": 2
                        }]
                    },
                    {
                        "uid": "test-uid",
                        "title": "Yearly Unit Test: no end",
                        "start": startString,
                        "timeZone": "UTC",
                        "duration": "PT1H30M",
                        recurrenceRules: [{
                            "@type": "RecurrenceRule",
                            "frequency": "yearly",
                            "byMonthDay": [11],
                            "byMonth": ['5']
                        }]
                    }
                ];

                for (const jsCalendar of jsCalendars) {
                    console.log(`${this.test?.title}: Testing item: ${util.inspect(jsCalendar, false, 10)}`);
                    const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
                    const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                    const mailboxId = 'test@test.com';

                    expect(manager).to.not.be.undefined();

                    const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape, mailboxId) as CalendarItemType;
                    expect(ewsItem.Recurrence).to.not.be.undefined();
                    expect(ewsItem.ItemId?.Id).to.not.be.undefined();
                    const itemEWSId = getEWSId(pimItem.unid, mailboxId);
                    expect(ewsItem.ItemId?.Id).to.be.equal(itemEWSId);
                    expect(ewsItem.IsRecurring).to.be.true();

                    // Validate recurrence rule
                    expect(ewsItem.Recurrence?.AbsoluteYearlyRecurrence?.DayOfMonth).to.be.equal(11);
                    expect(ewsItem.Recurrence?.AbsoluteYearlyRecurrence?.Month).to.be.equal(MonthNamesType.MAY);

                    // Validate first occurrence
                    const firstStartDate = new Date(pimItem.start!);
                    const firstEndDate = new Date(pimItem.end!);
                    let info = ewsItem.FirstOccurrence;
                    expect(info).to.not.be.undefined();
                    expect(info!.Start.getTime() - firstStartDate.getTime()).to.be.equal(0, `Expected First occurrence start ${info!.Start.toISOString()} to equal ${firstStartDate.toISOString()}`);
                    expect(info!.End.getTime() - firstEndDate.getTime()).to.be.equal(0, `Expected First occurrence end ${info!.End.toISOString()} to equal ${firstEndDate.toISOString()}`);
                    expect(info!.ItemId.Id).to.not.be.undefined();
                    expect(info!.ItemId.Id).to.be.equal(itemEWSId, `Expected first occurrence id ${info!.ItemId.Id} to equal ${getEWSId(pimItem.unid)}`);

                    if (!jsCalendar.title.endsWith('no end')) {
                        // Validate last occurrence
                        const lastStartDate = new Date(firstStartDate);
                        lastStartDate.setFullYear(lastStartDate.getFullYear() + 1);
                        const lastEndDate = new Date(firstEndDate);
                        lastEndDate.setFullYear(lastEndDate.getFullYear() + 1);
                        info = ewsItem.LastOccurrence;
                        expect(info).to.not.be.undefined();
                        expect(info!.Start.getTime() - lastStartDate.getTime()).to.be.equal(0, `Expected last occurrence start ${info!.Start.toISOString()} to equal ${lastStartDate.toISOString()}`);
                        expect(info!.End.getTime() - lastEndDate.getTime()).to.be.equal(0, `Expected last occurrence end ${info!.End.toISOString()} to equal ${lastEndDate.toISOString()}`);
                        expect(info!.ItemId.Id).to.not.be.undefined();
                        expect(info!.ItemId.Id).to.be.equal(itemEWSId, `Expected last occurrence id ${info!.ItemId.Id} to equal ${getEWSId(pimItem.unid)}`);

                        expect(ewsItem.Recurrence?.NumberedRecurrence).to.not.be.undefined();
                        expect(ewsItem.Recurrence?.NumberedRecurrence?.NumberOfOccurrences).to.be.equal(2, "NumberedRecurrence occurrences is not correct");
                        expect(ewsItem.Recurrence?.NumberedRecurrence?.StartDate.getTime()).to.be.equal(firstStartDate.getTime(), "NumberedRecurrence start is not correct");
                    }
                    else {
                        expect(ewsItem.Recurrence?.NoEndRecurrence).to.not.be.undefined();
                        expect(ewsItem.Recurrence?.NoEndRecurrence?.StartDate.getTime()).to.be.equal(firstStartDate.getTime(), "NoEndRecurrence start is not correct");
                    }

                }
            });

            it('Yearly on nth day of week in a month', async function () {
                const startString = '2021-04-08T09:00:00Z';
                const endString = '2022-04-14T09:00:00Z';

                const jsCalendars = [
                    {
                        "uid": "test-uid",
                        "title": "Yearly Unit Test",
                        "start": startString,
                        "duration": "PT1H00M",
                        recurrenceRules: [{
                            "@type": "RecurrenceRule",
                            "frequency": "yearly",
                            "byMonth": ['4'],
                            "byDay": [{ "@type": "NDay", day: "th", nthOfPeriod: 2 }],
                            "until": endString
                        }]
                    },
                    {
                        "uid": "test-uid",
                        "title": "Yearly Unit Test",
                        "start": startString,
                        "duration": "PT1H00M",
                        recurrenceRules: [{
                            "@type": "RecurrenceRule",
                            "frequency": "yearly",
                            "byMonth": ['4'],
                            "byDay": [{ "@type": "NDay", day: "th" }],
                            "bySetPosition": [2],
                            "until": endString
                        }]
                    }
                ];

                for (const jsCalendar of jsCalendars) {
                    console.log(`${this.test?.title}: Testing item: ${util.inspect(jsCalendar, false, 10)}`);
                    const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
                    const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                    expect(manager).to.not.be.undefined();

                    const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape) as CalendarItemType;
                    expect(ewsItem.Recurrence).to.not.be.undefined();
                    expect(ewsItem.ItemId?.Id).to.not.be.undefined();
                    expect(ewsItem.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));
                    expect(ewsItem.IsRecurring).to.be.true();

                    // Validate recurrence rule
                    expect(ewsItem.Recurrence?.RelativeYearlyRecurrence?.DayOfWeekIndex).to.be.equal(DayOfWeekIndexType.SECOND);
                    expect(ewsItem.Recurrence?.RelativeYearlyRecurrence?.DaysOfWeek).to.be.equal(DayOfWeekType.THURSDAY);
                    expect(ewsItem.Recurrence?.RelativeYearlyRecurrence?.Month).to.be.equal(MonthNamesType.APRIL);

                    // Validate first occurrence. The exact day is hard to determine but we can check year, month, month day, and time
                    let startDate = new Date(pimItem.start!);
                    let endDate = new Date(pimItem.end!);
                    let info = ewsItem.FirstOccurrence;
                    expect(info).to.not.be.undefined();
                    expect(info!.Start.getFullYear()).to.be.equal(startDate.getFullYear(), "First occurence start year is wrong");
                    expect(info!.Start.getMonth()).to.be.equal(startDate.getMonth(), "First occurence start month is wrong");
                    expect(info!.Start.getDay()).to.be.equal(startDate.getDay(), "First occurence start day is wrong")
                    expect(info!.Start.getHours()).to.be.equal(startDate.getHours(), "First occurence start hours is wrong");
                    expect(info!.Start.getMinutes()).to.be.equal(startDate.getMinutes(), "First occurence start minutes is wrong");
                    expect(info!.End.getFullYear()).to.be.equal(endDate.getFullYear(), "First occurence end year is wrong");
                    expect(info!.End.getMonth()).to.be.equal(endDate.getMonth(), "First occurence end month is wrong");
                    expect(info!.End.getDay()).to.be.equal(endDate.getDay(), "First occurence end day is wrong")
                    expect(info!.End.getHours()).to.be.equal(endDate.getHours(), "First occurence end hours is wrong");
                    expect(info!.End.getMinutes()).to.be.equal(endDate.getMinutes(), "First occurence end minutes is wrong");
                    expect(info!.ItemId.Id).to.not.be.undefined();
                    expect(info!.ItemId.Id).to.be.equal(getEWSId(pimItem.unid), `Expected first occurrence id ${info!.ItemId.Id} to equal ${getEWSId(pimItem.unid)}`);

                    // Validate last occurrence. The exact day is hard to determine but we can check year, month, and time
                    startDate = new Date(endString);
                    endDate = new Date(endString);
                    endDate.setHours(endDate.getHours() + 1);
                    info = ewsItem.LastOccurrence;
                    expect(info).to.not.be.undefined();
                    expect(info!.Start.getFullYear()).to.be.equal(startDate.getFullYear(), "Last occurence start year is wrong");
                    expect(info!.Start.getMonth()).to.be.equal(startDate.getMonth(), "Last occurence start month is wrong");
                    expect(info!.Start.getDay()).to.be.equal(startDate.getDay(), "Last occurence start day is wrong")
                    expect(info!.Start.getHours()).to.be.equal(startDate.getHours(), "Last occurence start hours is wrong");
                    expect(info!.Start.getMinutes()).to.be.equal(startDate.getMinutes(), "Last occurence start minutes is wrong");
                    expect(info!.End.getFullYear()).to.be.equal(endDate.getFullYear(), "Last occurence end year is wrong");
                    expect(info!.End.getMonth()).to.be.equal(endDate.getMonth(), "Last occurence end month is wrong");
                    expect(info!.End.getDay()).to.be.equal(endDate.getDay(), "Last occurence end day is wrong")
                    expect(info!.End.getHours()).to.be.equal(endDate.getHours(), "Last occurence end hours is wrong");
                    expect(info!.End.getMinutes()).to.be.equal(endDate.getMinutes(), "Last occurence end minutes is wrong");
                    expect(info!.ItemId.Id).to.not.be.undefined();
                    expect(info!.ItemId.Id).to.be.equal(getEWSId(pimItem.unid), `Expected last occurrence id ${info!.ItemId.Id} to equal ${getEWSId(pimItem.unid)}`);

                    expect(ewsItem.Recurrence?.EndDateRecurrence).to.not.be.undefined();
                    expect(ewsItem.Recurrence?.EndDateRecurrence?.StartDate.getTime()).to.be.equal(new Date(pimItem.start!).getTime(), "EndDateRecurrence start time is wrong");
                    expect(ewsItem.Recurrence?.EndDateRecurrence?.EndDate.getTime()).to.be.equal(startDate.getTime(), "EndDateRecurrence end time is wrong"); // Last start date
                }
            });

            it('Yearly in month on day group', async function () {
                const jsCalendars: any[] = [
                    {
                        "uid": "test-uid",
                        "title": "Yearly Unit Test - in April on 2nd weekend day",
                        "start": "2021-04-04T09:00:00",
                        "timeZone": "UTC",
                        "duration": "PT1H00M",
                        "recurrenceRules": [{
                            "@type": "RecurrenceRule",
                            "frequency": "yearly",
                            "byMonth": ['4'],
                            "byDay": [{ "@type": "NDay", day: "sa" }, { "@type": "NDay", day: "su" }],
                            "bySetPosition": [2],
                            "until": "2022-04-03T09:00:00"
                        }]
                    },
                    {
                        "uid": "test-uid",
                        "title": "Yearly Unit Test - in April on 3rd week day",
                        "start": "2021-04-05T09:00:00",
                        "timeZone": "UTC",
                        "duration": "PT1H00M",
                        "recurrenceRules": [{
                            "@type": "RecurrenceRule",
                            "frequency": "yearly",
                            "byMonth": ['4'],
                            "byDay": [{ "@type": "NDay", day: "mo" }, { "@type": "NDay", day: "tu" }, { "@type": "NDay", day: "we" }, { "@type": "NDay", day: "th" }, { "@type": "NDay", day: "fr" }],
                            "bySetPosition": [3],
                            "until": "2022-04-05T09:00:00"
                        }]
                    },
                    {
                        "uid": "test-uid",
                        "title": "Yearly Unit Test - in April on 4th day",
                        "start": "2021-04-04T09:00:00",
                        "timeZone": "UTC",
                        "duration": "PT1H00M",
                        recurrenceRules: [{
                            "@type": "RecurrenceRule",
                            "frequency": "yearly",
                            "byMonth": ['4'],
                            "byDay": [{ "@type": "NDay", day: "mo" }, { "@type": "NDay", day: "tu" }, { "@type": "NDay", day: "we" }, { "@type": "NDay", day: "th" }, { "@type": "NDay", day: "fr" }, { "@type": "NDay", day: "sa" }, { "@type": "NDay", day: "su" }],
                            "bySetPosition": [4],
                            "until": "2021-04-04T09:00:00"
                        }]
                    }
                ];

                for (const jsCalendar of jsCalendars) {
                    console.log(`${this.test?.title}: Testing item: ${util.inspect(jsCalendar, false, 10)}`);
                    const endString = `${jsCalendar.recurrenceRules[0].until}Z`;
                    const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
                    const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                    expect(manager).to.not.be.undefined();

                    const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape) as CalendarItemType;
                    expect(ewsItem.Recurrence).to.not.be.undefined();
                    expect(ewsItem.ItemId?.Id).to.not.be.undefined();
                    expect(ewsItem.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));
                    expect(ewsItem.IsRecurring).to.be.true();

                    // Validate recurrence rule
                    if (jsCalendar.title.endsWith('weekend day')) {
                        expect(ewsItem.Recurrence?.RelativeYearlyRecurrence?.DayOfWeekIndex).to.be.equal(DayOfWeekIndexType.SECOND);
                        expect(ewsItem.Recurrence?.RelativeYearlyRecurrence?.DaysOfWeek).to.be.equal(DayOfWeekType.WEEKEND_DAY);
                    }
                    else if (jsCalendar.title.endsWith('week day')) {
                        expect(ewsItem.Recurrence?.RelativeYearlyRecurrence?.DayOfWeekIndex).to.be.equal(DayOfWeekIndexType.THIRD);
                        expect(ewsItem.Recurrence?.RelativeYearlyRecurrence?.DaysOfWeek).to.be.equal(DayOfWeekType.WEEKDAY);
                    }
                    else {
                        expect(ewsItem.Recurrence?.RelativeYearlyRecurrence?.DayOfWeekIndex).to.be.equal(DayOfWeekIndexType.FOURTH);
                        expect(ewsItem.Recurrence?.RelativeYearlyRecurrence?.DaysOfWeek).to.be.equal(DayOfWeekType.DAY);
                    }
                    expect(ewsItem.Recurrence?.RelativeYearlyRecurrence?.Month).to.be.equal(MonthNamesType.APRIL);

                    // Validate first occurrence. The exact day is hard to determine but we can check year, month, month day, and time
                    let startDate = new Date(pimItem.start!);
                    let endDate = new Date(pimItem.end!);
                    let info = ewsItem.FirstOccurrence;
                    expect(info).to.not.be.undefined();
                    expect(info!.Start.getTime()).to.be.equal(startDate.getTime(), `Expected first occurrence start ${info!.Start.toISOString()} to equal ${startDate.toISOString()}`);
                    expect(info!.End.getTime()).to.be.equal(endDate.getTime(), `Expected first occurrence end ${info!.End.toISOString()} to equal ${endDate.toISOString()}`);
                    expect(info!.ItemId.Id).to.not.be.undefined();
                    expect(info!.ItemId.Id).to.be.equal(getEWSId(pimItem.unid), `Expected first occurrence id ${info!.ItemId.Id} to equal ${getEWSId(pimItem.unid)}`);

                    // Validate last occurrence. The exact day is hard to determine but we can check year, month, and time
                    startDate = new Date(endString);
                    endDate = new Date(endString);
                    endDate.setHours(endDate.getHours() + 1);
                    info = ewsItem.LastOccurrence;
                    expect(info).to.not.be.undefined();
                    expect(info!.Start.getTime()).to.be.equal(startDate.getTime(), `Expected last occurrence start ${info!.Start.toISOString()} to equal ${startDate.toISOString()}`);
                    expect(info!.End.getTime()).to.be.equal(endDate.getTime(), `Expected last occurrence end ${info!.End.toISOString()} to equal ${endDate.toISOString()}`);
                    expect(info!.ItemId.Id).to.not.be.undefined();
                    expect(info!.ItemId.Id).to.be.equal(getEWSId(pimItem.unid), `Expected last occurrence id ${info!.ItemId.Id} to equal ${getEWSId(pimItem.unid)}`);

                    expect(ewsItem.Recurrence?.EndDateRecurrence).to.not.be.undefined();
                    expect(ewsItem.Recurrence?.EndDateRecurrence?.StartDate.getTime()).to.be.equal(new Date(pimItem.start!).getTime(), "EndDateRecurrence start time is wrong");
                    expect(ewsItem.Recurrence?.EndDateRecurrence?.EndDate.getTime()).to.be.equal(startDate.getTime(), "EndDateRecurrence end time is wrong"); // Last start date
                }
            });
        });

        describe('Monthly recurrence', () => {

            it('Monthly, on day of month', async () => {
                // Start of first event
                const startString = '2021-05-11T09:00:00Z';

                // Every 2 Months on the 11th of the month
                const jsCalendar = {
                    "uid": "test-uid",
                    "title": "Monthly Unit Test",
                    "start": startString,
                    "duration": "PT1H00M",
                    recurrenceRules: [{
                        "@type": "RecurrenceRule",
                        "frequency": "monthly",
                        "byMonthDay": [11],
                        "interval": 2,
                        "count": 3
                    }]
                };

                const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
                const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                expect(manager).to.not.be.undefined();

                const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape) as CalendarItemType;
                expect(ewsItem.Recurrence).to.not.be.undefined();
                expect(ewsItem.ItemId?.Id).to.not.be.undefined();
                expect(ewsItem.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));
                expect(ewsItem.IsRecurring).to.be.true();

                // Validate recurrence rule
                expect(ewsItem.Recurrence?.AbsoluteMonthlyRecurrence?.DayOfMonth).to.be.equal(11);
                expect(ewsItem.Recurrence?.AbsoluteMonthlyRecurrence?.Interval).to.be.equal(2);

                // Validate first occurrence
                let info = ewsItem.FirstOccurrence;
                expect(info).to.not.be.undefined();
                let expectedStart = new Date(startString);
                let expectedEnd = new Date(startString);
                expectedEnd.setHours(expectedEnd.getHours() + 1); // 1 hr duration
                expect(info?.Start.getTime()).to.be.equal(expectedStart.getTime(), `Expected first occurrence start ${info!.Start.toISOString()} to equal ${expectedStart.toISOString()}`);
                expect(info?.End.getTime()).to.be.equal(expectedEnd.getTime(), `Expected first occurrence end ${info!.End.toISOString()} to equal ${expectedEnd.toISOString()}`);
                expect(info!.ItemId.Id).to.not.be.undefined();
                expect(info!.ItemId.Id).to.be.equal(getEWSId(pimItem.unid), `Expected first occurrence id ${info!.ItemId.Id} to equal ${getEWSId(pimItem.unid)}`);

                // Validate last occurrence
                info = ewsItem.LastOccurrence;
                expect(info).to.not.be.undefined();
                expectedStart = new Date(startString);
                expectedStart.setMonth(expectedStart.getMonth() + 4);
                expectedEnd = new Date(expectedStart);
                expectedEnd.setHours(expectedStart.getHours() + 1); // 1 hr duration
                expect(info?.Start.getTime()).to.be.equal(expectedStart.getTime(), `Expected last occurrence start ${info!.Start.toISOString()} to equal ${expectedStart.toISOString()}`);
                expect(info?.End.getTime()).to.be.equal(expectedEnd.getTime(), `Expected last occurrence end ${info!.End.toISOString()} to equal ${expectedEnd.toISOString()}`);
                expect(info!.ItemId.Id).to.not.be.undefined();
                expect(info!.ItemId.Id).to.be.equal(getEWSId(pimItem.unid), `Expected last occurrence id ${info!.ItemId.Id} to equal ${getEWSId(pimItem.unid)}`);

                expect(ewsItem.Recurrence?.NumberedRecurrence).to.not.be.undefined();
                expect(ewsItem.Recurrence?.NumberedRecurrence?.NumberOfOccurrences).to.be.equal(3, "NumberedRecurrence occurrences is not correct");
                expect(ewsItem.Recurrence?.NumberedRecurrence?.StartDate.getTime()).to.be.equal(new Date(pimItem.start!).getTime(), "NumberedRecurrence start is not correct");
            });

            it('Monthly, on nth day of week', async function () {
                const jsCalendars = [
                    {
                        "uid": "test-uid",
                        "title": "Monthly Unit Test - fourth tuesday",
                        "start": '2021-05-25T09:00:00Z',
                        "duration": "PT1H00M",
                        recurrenceRules: [{
                            "@type": "RecurrenceRule",
                            "frequency": "monthly",
                            "byDay": [{ "@type": "NDay", day: "tu", nthOfPeriod: 4 }],
                            "count": 3
                        }]
                    },
                    {
                        "uid": "test-uid",
                        "title": "Monthly Unit Test - first tuesday",
                        "start": '2021-05-04T09:00:00Z',
                        "duration": "PT1H00M",
                        recurrenceRules: [{
                            "@type": "RecurrenceRule",
                            "frequency": "monthly",
                            "byDay": [{ "@type": "NDay", day: "tu" }],
                            "bySetPosition": [1],
                            "count": 3
                        }]
                    }
                ];

                for (const jsCalendar of jsCalendars) {
                    console.log(`${this.test?.title}: Testing item: ${util.inspect(jsCalendar, false, 10)}`);
                    const startString = jsCalendar.start;
                    const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
                    const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                    expect(manager).to.not.be.undefined();

                    const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape) as CalendarItemType;
                    expect(ewsItem.Recurrence).to.not.be.undefined();
                    expect(ewsItem.ItemId?.Id).to.not.be.undefined();
                    expect(ewsItem.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));
                    expect(ewsItem.IsRecurring).to.be.true();

                    // Validate recurrence rule
                    if (jsCalendar.title.endsWith('fourth tuesday')) {
                        expect(ewsItem.Recurrence?.RelativeMonthlyRecurrence?.DayOfWeekIndex).to.be.equal(DayOfWeekIndexType.FOURTH);
                    } else {
                        expect(ewsItem.Recurrence?.RelativeMonthlyRecurrence?.DayOfWeekIndex).to.be.equal(DayOfWeekIndexType.FIRST);
                    }
                    expect(ewsItem.Recurrence?.RelativeMonthlyRecurrence?.DaysOfWeek).to.be.equal(DayOfWeekType.TUESDAY);
                    expect(ewsItem.Recurrence?.RelativeMonthlyRecurrence?.Interval).to.be.equal(1); // Default is 1

                    // Validate first occurrence. The exact day is hard to determine but we can check year, month, month day, and time
                    let expectedStart = new Date(startString);
                    let expectedEnd = new Date(startString);
                    expectedEnd.setHours(expectedEnd.getHours() + 1); // 1 hr duration
                    let info = ewsItem.FirstOccurrence;
                    expect(info).to.not.be.undefined();
                    expect(info!.Start.getFullYear()).to.be.equal(expectedStart.getFullYear(), "First occurence start year is wrong");
                    expect(info!.Start.getMonth()).to.be.equal(expectedStart.getMonth(), "First occurence start month is wrong");
                    expect(info!.Start.getDay()).to.be.equal(expectedStart.getDay(), "First occurence start day is wrong")
                    expect(info!.Start.getHours()).to.be.equal(expectedStart.getHours(), "First occurence start hours is wrong");
                    expect(info!.Start.getMinutes()).to.be.equal(expectedStart.getMinutes(), "First occurence start minutes is wrong");
                    expect(info!.End.getFullYear()).to.be.equal(expectedEnd.getFullYear(), "First occurence end year is wrong");
                    expect(info!.End.getMonth()).to.be.equal(expectedEnd.getMonth(), "First occurence end month is wrong");
                    expect(info!.End.getDay()).to.be.equal(expectedEnd.getDay(), "First occurence end day is wrong")
                    expect(info!.End.getHours()).to.be.equal(expectedEnd.getHours(), "First occurence end hours is wrong");
                    expect(info!.End.getMinutes()).to.be.equal(expectedEnd.getMinutes(), "First occurence end minutes is wrong");
                    expect(info!.ItemId.Id).to.not.be.undefined();
                    expect(info!.ItemId.Id).to.be.equal(getEWSId(pimItem.unid), `Expected first occurrence id ${info!.ItemId.Id} to equal ${getEWSId(pimItem.unid)}`);

                    // Validate last occurrence. The exact day is hard to determine but we can check year, month, and time
                    info = ewsItem.LastOccurrence;
                    expectedStart = new Date(startString);
                    expectedStart.setMonth(expectedStart.getMonth() + 2);
                    // Setting day of week is not supported by Date, so can't verify the exact date
                    expectedEnd = new Date(expectedStart);
                    expectedEnd.setHours(expectedEnd.getHours() + 1); // 1 hr duration
                    expect(info).to.not.be.undefined();
                    expect(info!.Start.getFullYear()).to.be.equal(expectedStart.getFullYear(), "First occurence start year is wrong");
                    expect(info!.Start.getMonth()).to.be.equal(expectedStart.getMonth(), "First occurence start month is wrong");
                    expect(info!.Start.getHours()).to.be.equal(expectedStart.getHours(), "First occurence start hours is wrong");
                    expect(info!.Start.getMinutes()).to.be.equal(expectedStart.getMinutes(), "First occurence start minutes is wrong");
                    expect(info!.End.getFullYear()).to.be.equal(expectedEnd.getFullYear(), "First occurence end year is wrong");
                    expect(info!.End.getMonth()).to.be.equal(expectedEnd.getMonth(), "First occurence end month is wrong");
                    expect(info!.End.getHours()).to.be.equal(expectedEnd.getHours(), "First occurence end hours is wrong");
                    expect(info!.End.getMinutes()).to.be.equal(expectedEnd.getMinutes(), "First occurence end minutes is wrong");
                    expect(info!.ItemId.Id).to.not.be.undefined();
                    expect(info!.ItemId.Id).to.be.equal(getEWSId(pimItem.unid), `Expected last occurrence id ${info!.ItemId.Id} to equal ${getEWSId(pimItem.unid)}`);

                    expect(ewsItem.Recurrence?.NumberedRecurrence).to.not.be.undefined();
                    expect(ewsItem.Recurrence?.NumberedRecurrence?.NumberOfOccurrences).to.be.equal(3, "NumberedRecurrence occurrences is not correct");
                    expect(ewsItem.Recurrence?.NumberedRecurrence?.StartDate.getTime()).to.be.equal(new Date(pimItem.start!).getTime(), "NumberedRecurrence start is not correct");
                }
            });

            it('Monthly, on nth day group', async function () {
                const jsCalendars = [
                    {
                        "uid": "test-uid",
                        "title": "Monthly Unit Test - on last weekday",
                        "start": "2021-05-31T09:00:00",
                        "timeZone": "UTC",
                        "duration": "PT1H00M",
                        "recurrenceRules": [{
                            "@type": "RecurrenceRule",
                            "frequency": "monthly",
                            "byDay": [{ "@type": "NDay", day: "mo" }, { "@type": "NDay", day: "tu" }, { "@type": "NDay", day: "we" }, { "@type": "NDay", day: "th" }, { "@type": "NDay", day: "fr" }],
                            "bySetPosition": [-1],
                            "count": 3
                        }]
                    },
                    {
                        "uid": "test-uid",
                        "title": "Monthly Unit Test - on 2nd week end day",
                        "start": "2021-05-02T09:00:00",
                        "timeZone": "UTC",
                        "duration": "PT1H00M",
                        recurrenceRules: [{
                            "@type": "RecurrenceRule",
                            "frequency": "monthly",
                            "byDay": [{ "@type": "NDay", day: "sa" }, { "@type": "NDay", day: "su" }],
                            "bySetPosition": [2],
                            "count": 3
                        }]
                    },
                    {
                        "uid": "test-uid",
                        "title": "Monthly Unit Test - on 1st day",
                        "start": "2021-05-01T09:00:00",
                        "timeZone": "UTC",
                        "duration": "PT1H00M",
                        recurrenceRules: [{
                            "@type": "RecurrenceRule",
                            "frequency": "monthly",
                            "byDay": [{ "@type": "NDay", day: "mo" }, { "@type": "NDay", day: "tu" }, { "@type": "NDay", day: "we" }, { "@type": "NDay", day: "th" }, { "@type": "NDay", day: "fr" }, { "@type": "NDay", day: "sa" }, { "@type": "NDay", day: "su" }],
                            "bySetPosition": [1],
                            "count": 3
                        }]
                    }
                ];

                for (const jsCalendar of jsCalendars) {
                    console.log(`${this.test?.title}: Testing item: ${util.inspect(jsCalendar, false, 10)}`);
                    const startString = `${jsCalendar.start}Z`;
                    const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
                    const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                    expect(manager).to.not.be.undefined();

                    const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape) as CalendarItemType;
                    expect(ewsItem.Recurrence).to.not.be.undefined();
                    expect(ewsItem.ItemId?.Id).to.not.be.undefined();
                    expect(ewsItem.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));
                    expect(ewsItem.IsRecurring).to.be.true();

                    // Validate recurrence rule
                    if (jsCalendar.title.endsWith('weekday')) {
                        expect(ewsItem.Recurrence?.RelativeMonthlyRecurrence?.DaysOfWeek).to.be.equal(DayOfWeekType.WEEKDAY);
                        expect(ewsItem.Recurrence?.RelativeMonthlyRecurrence?.DayOfWeekIndex).to.be.equal(DayOfWeekIndexType.LAST);
                    }
                    else if (jsCalendar.title.endsWith('week end day')) {
                        expect(ewsItem.Recurrence?.RelativeMonthlyRecurrence?.DaysOfWeek).to.be.equal(DayOfWeekType.WEEKEND_DAY);
                        expect(ewsItem.Recurrence?.RelativeMonthlyRecurrence?.DayOfWeekIndex).to.be.equal(DayOfWeekIndexType.SECOND);
                    }
                    else {
                        expect(ewsItem.Recurrence?.RelativeMonthlyRecurrence?.DaysOfWeek).to.be.equal(DayOfWeekType.DAY);
                        expect(ewsItem.Recurrence?.RelativeMonthlyRecurrence?.DayOfWeekIndex).to.be.equal(DayOfWeekIndexType.FIRST);
                    }

                    expect(ewsItem.Recurrence?.RelativeMonthlyRecurrence?.Interval).to.be.equal(1); // Default is 1

                    // Validate first occurrence. The exact day is hard to determine but we can check year, month, month day, and time
                    let expectedStart = new Date(startString);
                    let expectedEnd = new Date(startString);
                    expectedEnd.setHours(expectedEnd.getHours() + 1); // 1 hr duration
                    let info = ewsItem.FirstOccurrence;
                    expect(info).to.not.be.undefined();
                    expect(info!.Start.getFullYear()).to.be.equal(expectedStart.getFullYear(), "First occurence start year is wrong");
                    expect(info!.Start.getMonth()).to.be.equal(expectedStart.getMonth(), "First occurence start month is wrong");
                    expect(info!.Start.getDay()).to.be.equal(expectedStart.getDay(), "First occurence start day is wrong")
                    expect(info!.Start.getHours()).to.be.equal(expectedStart.getHours(), "First occurence start hours is wrong");
                    expect(info!.Start.getMinutes()).to.be.equal(expectedStart.getMinutes(), "First occurence start minutes is wrong");
                    expect(info!.End.getFullYear()).to.be.equal(expectedEnd.getFullYear(), "First occurence end year is wrong");
                    expect(info!.End.getMonth()).to.be.equal(expectedEnd.getMonth(), "First occurence end month is wrong");
                    expect(info!.End.getDay()).to.be.equal(expectedEnd.getDay(), "First occurence end day is wrong")
                    expect(info!.End.getHours()).to.be.equal(expectedEnd.getHours(), "First occurence end hours is wrong");
                    expect(info!.End.getMinutes()).to.be.equal(expectedEnd.getMinutes(), "First occurence end minutes is wrong");
                    expect(info!.ItemId.Id).to.not.be.undefined();
                    expect(info!.ItemId.Id).to.be.equal(getEWSId(pimItem.unid), `Expected first occurrence id ${info!.ItemId.Id} to equal ${getEWSId(pimItem.unid)}`);

                    // Validate last occurrence. The exact day is hard to determine but we can check year, month, and time
                    info = ewsItem.LastOccurrence;
                    expectedStart = new Date(startString);
                    expectedStart.setMonth(expectedStart.getMonth() + 2);
                    // Setting day of week is not supported by Date, so can't verify the exact date
                    expectedEnd = new Date(expectedStart);
                    expectedEnd.setHours(expectedEnd.getHours() + 1); // 1 hr duration
                    expect(info).to.not.be.undefined();
                    expect(info!.Start.getFullYear()).to.be.equal(expectedStart.getFullYear(), "First occurence start year is wrong");
                    expect(info!.Start.getMonth()).to.be.equal(expectedStart.getMonth(), "First occurence start month is wrong");
                    expect(info!.Start.getHours()).to.be.equal(expectedStart.getHours(), "First occurence start hours is wrong");
                    expect(info!.Start.getMinutes()).to.be.equal(expectedStart.getMinutes(), "First occurence start minutes is wrong");
                    expect(info!.End.getFullYear()).to.be.equal(expectedEnd.getFullYear(), "First occurence end year is wrong");
                    expect(info!.End.getMonth()).to.be.equal(expectedEnd.getMonth(), "First occurence end month is wrong");
                    expect(info!.End.getHours()).to.be.equal(expectedEnd.getHours(), "First occurence end hours is wrong");
                    expect(info!.End.getMinutes()).to.be.equal(expectedEnd.getMinutes(), "First occurence end minutes is wrong");
                    expect(info!.ItemId.Id).to.not.be.undefined();
                    expect(info!.ItemId.Id).to.be.equal(getEWSId(pimItem.unid), `Expected last occurrence id ${info!.ItemId.Id} to equal ${getEWSId(pimItem.unid)}`);

                    expect(ewsItem.Recurrence?.NumberedRecurrence).to.not.be.undefined();
                    expect(ewsItem.Recurrence?.NumberedRecurrence?.NumberOfOccurrences).to.be.equal(3, "NumberedRecurrence occurrences is not correct");
                    expect(ewsItem.Recurrence?.NumberedRecurrence?.StartDate.getTime()).to.be.equal(new Date(pimItem.start!).getTime(), "NumberedRecurrence start is not correct");
                }
            });
        });

        describe('Weekly recurrence', () => {

            it('Weekly on day of the week', async () => {
                // Start of first event
                const startString = '2021-05-04T09:00:00Z';

                // Every work day for 9 occurrences
                const jsCalendar = {
                    "uid": "test-uid",
                    "title": "Weely Unit Test",
                    "start": startString,
                    "duration": "PT1H00M",
                    recurrenceRules: [{
                        "@type": "RecurrenceRule",
                        "frequency": "weekly",
                        "byDay": [
                            { "@type": "NDay", day: "mo" },
                            { "@type": "NDay", day: "tu" },
                            { "@type": "NDay", day: "we" },
                            { "@type": "NDay", day: "th" },
                            { "@type": "NDay", day: "fr" },
                        ],
                        "count": 9
                    }]
                };

                const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
                const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                expect(manager).to.not.be.undefined();

                const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape) as CalendarItemType;
                expect(ewsItem.Recurrence).to.not.be.undefined();
                expect(ewsItem.ItemId?.Id).to.not.be.undefined();
                expect(ewsItem.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));
                expect(ewsItem.IsRecurring).to.be.true();

                // Validate recurrence rule
                expect(ewsItem.Recurrence?.WeeklyRecurrence?.FirstDayOfWeek).to.be.equal(DayOfWeekType.MONDAY);
                expect(ewsItem.Recurrence?.WeeklyRecurrence?.DaysOfWeek.$value).to.be.deepEqual(`${DayOfWeekType.MONDAY} ${DayOfWeekType.TUESDAY} ${DayOfWeekType.WEDNESDAY} ${DayOfWeekType.THURSDAY} ${DayOfWeekType.FRIDAY}`);
                expect(ewsItem.Recurrence?.WeeklyRecurrence?.Interval).to.be.equal(1); // Default is 1

                // Validate first occurrence
                let info = ewsItem.FirstOccurrence;
                expect(info).to.not.be.undefined();
                let expectedStart = new Date(startString);
                let expectedEnd = new Date(startString);
                expectedEnd.setHours(expectedEnd.getHours() + 1); // 1 hr duration
                expect(info?.Start.getTime()).to.be.equal(expectedStart.getTime(), `Expected first occurrence start ${info!.Start.toISOString()} to equal ${expectedStart.toISOString()}`);
                expect(info?.End.getTime()).to.be.equal(expectedEnd.getTime(), `Expected first occurrence end ${info!.End.toISOString()} to equal ${expectedEnd.toISOString()}`);
                expect(info!.ItemId.Id).to.not.be.undefined();
                expect(info!.ItemId.Id).to.be.equal(getEWSId(pimItem.unid), `Expected first occurrence id ${info!.ItemId.Id} to equal ${getEWSId(pimItem.unid)}`);

                // Validate last occurrence
                info = ewsItem.LastOccurrence;
                expect(info).to.not.be.undefined();
                const endString = '2021-05-14T09:00:00Z';
                expectedStart = new Date(endString);
                expectedEnd = new Date(expectedStart);
                expectedEnd.setHours(expectedStart.getHours() + 1); // 1 hr duration
                expect(info?.Start.getTime()).to.be.equal(expectedStart.getTime(), `Expected last occurrence start ${info!.Start.toISOString()} to equal ${expectedStart.toISOString()}`);
                expect(info?.End.getTime()).to.be.equal(expectedEnd.getTime(), `Expected last occurrence end ${info!.End.toISOString()} to equal ${expectedEnd.toISOString()}`);
                expect(info!.ItemId.Id).to.not.be.undefined();
                expect(info!.ItemId.Id).to.be.equal(getEWSId(pimItem.unid), `Expected last occurrence id ${info!.ItemId.Id} to equal ${getEWSId(pimItem.unid)}`);

                expect(ewsItem.Recurrence?.NumberedRecurrence).to.not.be.undefined();
                expect(ewsItem.Recurrence?.NumberedRecurrence?.NumberOfOccurrences).to.be.equal(9, "NumberedRecurrence occurrences is not correct");
                expect(ewsItem.Recurrence?.NumberedRecurrence?.StartDate.getTime()).to.be.equal(new Date(pimItem.start!).getTime(), "NumberedRecurrence start is not correct");

            });

            it('Every nth Week', async () => {
                // Start of first event
                const startString = '2021-05-04T09:00:00Z';

                // Every 2 weeks
                const jsCalendar = {
                    "uid": "test-uid",
                    "title": "Weely Unit Test",
                    "start": startString,
                    "duration": "PT1H00M",
                    recurrenceRules: [{
                        "@type": "RecurrenceRule",
                        "frequency": "weekly",
                        "byDay": [
                            { "@type": "NDay", day: "tu" }
                        ],
                        "interval": 2,
                        "count": 3
                    }]
                };

                const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
                const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                expect(manager).to.not.be.undefined();

                const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape) as CalendarItemType;
                expect(ewsItem.Recurrence).to.not.be.undefined();
                expect(ewsItem.ItemId?.Id).to.not.be.undefined();
                expect(ewsItem.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));
                expect(ewsItem.IsRecurring).to.be.true();

                // Validate recurrence rule
                expect(ewsItem.Recurrence?.WeeklyRecurrence?.FirstDayOfWeek).to.be.equal(DayOfWeekType.MONDAY);
                expect(ewsItem.Recurrence?.WeeklyRecurrence?.DaysOfWeek.$value).to.be.equal(DayOfWeekType.TUESDAY);
                expect(ewsItem.Recurrence?.WeeklyRecurrence?.Interval).to.be.equal(2);

                // Validate first occurrence
                let info = ewsItem.FirstOccurrence;
                expect(info).to.not.be.undefined();
                let expectedStart = new Date(startString);
                let expectedEnd = new Date(startString);
                expectedEnd.setHours(expectedEnd.getHours() + 1); // 1 hr duration
                expect(info?.Start.getTime()).to.be.equal(expectedStart.getTime(), `Expected first occurrence start ${info!.Start.toISOString()} to equal ${expectedStart.toISOString()}`);
                expect(info?.End.getTime()).to.be.equal(expectedEnd.getTime(), `Expected first occurrence end ${info!.End.toISOString()} to equal ${expectedEnd.toISOString()}`);
                expect(info!.ItemId.Id).to.not.be.undefined();
                expect(info!.ItemId.Id).to.be.equal(getEWSId(pimItem.unid), `Expected first occurrence id ${info!.ItemId.Id} to equal ${getEWSId(pimItem.unid)}`);

                // Validate last occurrence
                info = ewsItem.LastOccurrence;
                expect(info).to.not.be.undefined();
                const endString = '2021-06-01T09:00:00Z';
                expectedStart = new Date(endString);
                expectedEnd = new Date(expectedStart);
                expectedEnd.setHours(expectedStart.getHours() + 1); // 1 hr duration
                expect(info?.Start.getTime()).to.be.equal(expectedStart.getTime(), `Expected last occurrence start ${info!.Start.toISOString()} to equal ${expectedStart.toISOString()}`);
                expect(info?.End.getTime()).to.be.equal(expectedEnd.getTime(), `Expected last occurrence end ${info!.End.toISOString()} to equal ${expectedEnd.toISOString()}`);
                expect(info!.ItemId.Id).to.not.be.undefined();
                expect(info!.ItemId.Id).to.be.equal(getEWSId(pimItem.unid), `Expected last occurrence id ${info!.ItemId.Id} to equal ${getEWSId(pimItem.unid)}`);

                expect(ewsItem.Recurrence?.NumberedRecurrence).to.not.be.undefined();
                expect(ewsItem.Recurrence?.NumberedRecurrence?.NumberOfOccurrences).to.be.equal(3, "NumberedRecurrence occurrences is not correct");
                expect(ewsItem.Recurrence?.NumberedRecurrence?.StartDate.getTime()).to.be.equal(new Date(pimItem.start!).getTime(), "NumberedRecurrence start is not correct");
            });
        });

        describe('Daily recurrence', () => {

            it('Every other day', async () => {
                // Start of first event
                const startString = '2021-05-04T09:00:00Z';

                // Every other day
                const jsCalendar = {
                    "uid": "test-uid",
                    "title": "Weely Unit Test",
                    "start": startString,
                    "duration": "PT1H00M",
                    recurrenceRules: [{
                        "@type": "RecurrenceRule",
                        "frequency": "daily",
                        "interval": 2,
                        "count": 10
                    }]
                };

                const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
                const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                expect(manager).to.not.be.undefined();

                const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape) as CalendarItemType;
                expect(ewsItem.Recurrence).to.not.be.undefined();
                expect(ewsItem.ItemId?.Id).to.not.be.undefined();
                expect(ewsItem.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));
                expect(ewsItem.IsRecurring).to.be.true();

                // Validate recurrence rule
                expect(ewsItem.Recurrence?.DailyRecurrence?.Interval).to.be.equal(2);

                // Validate first occurrence
                let info = ewsItem.FirstOccurrence;
                expect(info).to.not.be.undefined();
                let expectedStart = new Date(startString);
                let expectedEnd = new Date(startString);
                expectedEnd.setHours(expectedEnd.getHours() + 1); // 1 hr duration
                expect(info?.Start.getTime()).to.be.equal(expectedStart.getTime(), `Expected first occurrence start ${info!.Start.toISOString()} to equal ${expectedStart.toISOString()}`);
                expect(info?.End.getTime()).to.be.equal(expectedEnd.getTime(), `Expected first occurrence end ${info!.End.toISOString()} to equal ${expectedEnd.toISOString()}`);
                expect(info!.ItemId.Id).to.not.be.undefined();
                expect(info!.ItemId.Id).to.be.equal(getEWSId(pimItem.unid), `Expected first occurrence id ${info!.ItemId.Id} to equal ${getEWSId(pimItem.unid)}`);

                // Validate last occurrence
                info = ewsItem.LastOccurrence;
                expect(info).to.not.be.undefined();
                const endString = '2021-05-22T09:00:00Z';
                expectedStart = new Date(endString);
                expectedEnd = new Date(expectedStart);
                expectedEnd.setHours(expectedStart.getHours() + 1); // 1 hr duration
                expect(info?.Start.getTime()).to.be.equal(expectedStart.getTime(), `Expected last occurrence start ${info!.Start.toISOString()} to equal ${expectedStart.toISOString()}`);
                expect(info?.End.getTime()).to.be.equal(expectedEnd.getTime(), `Expected last occurrence end ${info!.End.toISOString()} to equal ${expectedEnd.toISOString()}`);
                expect(info!.ItemId.Id).to.not.be.undefined();
                expect(info!.ItemId.Id).to.be.equal(getEWSId(pimItem.unid), `Expected last occurrence id ${info!.ItemId.Id} to equal ${getEWSId(pimItem.unid)}`);

                expect(ewsItem.Recurrence?.NumberedRecurrence).to.not.be.undefined();
                expect(ewsItem.Recurrence?.NumberedRecurrence?.NumberOfOccurrences).to.be.equal(10, "NumberedRecurrence occurrences is not correct");
                expect(ewsItem.Recurrence?.NumberedRecurrence?.StartDate.getTime()).to.be.equal(new Date(pimItem.start!).getTime(), "NumberedRecurrence start is not correct");
            });
        });
    });

    describe('Delete occurences', () => {

        const shape = new ItemResponseShapeType();
        shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
        const additions = [UnindexedFieldURIType.CALENDAR_RECURRENCE, UnindexedFieldURIType.CALENDAR_DELETED_OCCURRENCES];
        shape.AdditionalProperties = new NonEmptyArrayOfPathsToElementType(additions.map(type => {
            const fieldType = new PathToUnindexedFieldType();
            fieldType.FieldURI = type;
            return fieldType;
        }));

        it('Single exclude rule', async () => {
            // Start of first event
            const startString = '2021-05-04T09:00:00Z';

            // Every month on the 1st Tue of the month, except for June and July
            const jsCalendar = {
                "uid": "test-uid",
                "title": "Monthly Unit Test",
                "start": startString,
                "duration": "PT1H00M",
                recurrenceRules: [{
                    "@type": "RecurrenceRule",
                    "frequency": "monthly",
                    "byDay": [{ "@type": "NDay", day: "tu", nthOfPeriod: 1 }],
                    "count": 8
                }],
                excludedRecurrenceRules: [{
                    "@type": "RecurrenceRule",
                    "frequency": "monthly",
                    "byDay": [{ "@type": "NDay", day: "tu", nthOfPeriod: 1 }],
                    "byMonth": ["6", "7"]
                }]
            };

            const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
            const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
            expect(manager).to.not.be.undefined();

            const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape) as CalendarItemType;
            expect(ewsItem.Recurrence).to.not.be.undefined();
            expect(ewsItem.ItemId?.Id).to.not.be.undefined();
            expect(ewsItem.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));

            // Validate recurrence rule
            expect(ewsItem.Recurrence?.RelativeMonthlyRecurrence?.DayOfWeekIndex).to.be.equal(DayOfWeekIndexType.FIRST);
            expect(ewsItem.Recurrence?.RelativeMonthlyRecurrence?.DaysOfWeek).to.be.equal(DayOfWeekType.TUESDAY);
            expect(ewsItem.Recurrence?.RelativeMonthlyRecurrence?.Interval).to.be.equal(1); // Default is 1

            // Validate deleted occurrences
            expect(ewsItem.DeletedOccurrences).to.not.be.undefined();
            const expected = [new Date("2021-06-01T09:00:00Z").toISOString(), new Date("2021-07-06T09:00:00Z").toISOString()];
            expect(ewsItem.DeletedOccurrences!.DeletedOccurrence.map(occurence => occurence.Start.toISOString())).to.deepEqual(expected);

        });

        it('Multiple exclude rules', async () => {
            // Start of first event
            const startString = '2021-05-03T09:00:00Z';

            // Every weekday from 5/3 - 6/4, excluding 5/12, 5/20, and 6/4
            const jsCalendar = {
                "uid": "test-uid",
                "title": "Weekly Unit Test",
                "start": startString,
                "duration": "PT1H00M",
                recurrenceRules: [{
                    "@type": "RecurrenceRule",
                    "frequency": "weekly",
                    "byDay": [
                        { "@type": "NDay", day: "mo" },
                        { "@type": "NDay", day: "tu" },
                        { "@type": "NDay", day: "we" },
                        { "@type": "NDay", day: "th" },
                        { "@type": "NDay", day: "fr" }
                    ],
                    "count": 20
                }],
                excludedRecurrenceRules: [{
                    "@type": "RecurrenceRule",
                    "frequency": "monthly",
                    "byMonthDay": [12, 20],
                    "byMonth": ["5"],
                    "count": 2
                },
                {
                    "@type": "RecurrenceRule",
                    "frequency": "monthly",
                    "byMonthDay": [4],
                    "byMonth": ["6"],
                    "count": 1
                }]
            };

            const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
            const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
            const mailboxId = 'test@test.com';

            expect(manager).to.not.be.undefined();

            const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape, mailboxId) as CalendarItemType;
            expect(ewsItem.Recurrence).to.not.be.undefined();
            expect(ewsItem.ItemId?.Id).to.not.be.undefined();
            const itemEWSId = getEWSId(pimItem.unid, mailboxId);
            expect(ewsItem.ItemId?.Id).to.be.equal(itemEWSId);

            // Validate recurrence rule
            expect(ewsItem.Recurrence?.WeeklyRecurrence?.FirstDayOfWeek).to.be.equal(DayOfWeekType.MONDAY);
            expect(ewsItem.Recurrence?.WeeklyRecurrence?.DaysOfWeek.$value).to.be.deepEqual(`${DayOfWeekType.MONDAY} ${DayOfWeekType.TUESDAY} ${DayOfWeekType.WEDNESDAY} ${DayOfWeekType.THURSDAY} ${DayOfWeekType.FRIDAY}`);
            expect(ewsItem.Recurrence?.WeeklyRecurrence?.Interval).to.be.equal(1); // Default is 1

            // Validate deleted occurrences
            expect(ewsItem.DeletedOccurrences).to.not.be.undefined();
            const expected = [new Date("2021-05-12T09:00:00Z").toISOString(), new Date("2021-05-20T09:00:00Z").toISOString(), new Date("2021-06-04T09:00:00Z").toISOString()];
            expect(ewsItem.DeletedOccurrences!.DeletedOccurrence.map(occurence => occurence.Start.toISOString())).to.deepEqual(expected);

        });
    });

    describe('Modify occurrences', () => {

        const shape = new ItemResponseShapeType();
        shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
        const additions = [UnindexedFieldURIType.CALENDAR_RECURRENCE, UnindexedFieldURIType.CALENDAR_DELETED_OCCURRENCES, UnindexedFieldURIType.CALENDAR_MODIFIED_OCCURRENCES];
        shape.AdditionalProperties = new NonEmptyArrayOfPathsToElementType(additions.map(type => {
            const fieldType = new PathToUnindexedFieldType();
            fieldType.FieldURI = type;
            return fieldType;
        }));

        it('Modify start date', async () => {
            // Start of first event
            const startString = '2021-05-04T09:00:00';

            // Every month on the 1st Tue of the month at 9. On June 1st it starts at 10. 
            const jsCalendar = {
                "uid": "test-uid",
                "title": "Monthly Unit Test",
                "start": startString,
                "timeZone": "UTC",
                "duration": "PT1H00M",
                recurrenceRules: [{
                    "@type": "RecurrenceRule",
                    "frequency": "monthly",
                    "byDay": [{ "@type": "NDay", day: "tu", nthOfPeriod: 1 }],
                    "count": 8
                }],
                recurrenceOverrides: {
                    '2021-06-01T09:00:00': {
                        "start": '2021-06-01T10:00:00'
                    }
                }
            };

            const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
            const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
            const mailboxId = 'test@test.com';

            expect(manager).to.not.be.undefined();

            const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape, mailboxId) as CalendarItemType;
            expect(ewsItem.Recurrence).to.not.be.undefined();
            expect(ewsItem.ItemId?.Id).to.not.be.undefined();
            const itemEWSId = getEWSId(pimItem.unid, mailboxId);
            expect(ewsItem.ItemId?.Id).to.be.equal(itemEWSId);
            expect(pimItem.start).to.not.be.undefined();
            expect(pimItem.end).to.not.be.undefined();
            expect(compareDateStrings(new Date(pimItem.start!).toISOString(), new Date("2021-05-04T09:00:00Z").toISOString())).to.be.equal(0, `${pimItem.start} is not correct start time for the first event`);
            expect(compareDateStrings(new Date(pimItem.end!).toISOString(), new Date("2021-05-04T10:00:00Z").toISOString())).to.be.equal(0, `${pimItem.end} is not correct end time for the first event`);


            // Validate modify occurrences
            expect(ewsItem.ModifiedOccurrences).to.not.be.undefined();
            expect(ewsItem.ModifiedOccurrences?.Occurrence.length).to.be.equal(1);
            expect(compareDateStrings(ewsItem.ModifiedOccurrences!.Occurrence[0].Start.toISOString(), new Date("2021-06-01T10:00:00Z").toISOString())).to.be.equal(0, `${ewsItem.ModifiedOccurrences!.Occurrence[0].Start.toUTCString()} is not correct start time`);
            expect(compareDateStrings(ewsItem.ModifiedOccurrences!.Occurrence[0].End.toISOString(), new Date("2021-06-01T11:00:00Z").toISOString())).to.be.equal(0, `${ewsItem.ModifiedOccurrences!.Occurrence[0].End.toUTCString()} is not correct end time`);

        });

        it('Modify end', async () => {
            // Start of first event
            const startString = '2021-05-04T09:00:00';

            // Every month on the 1st Tue of the month for 1 hour. On July 6th it lasts 1.5 hours. 
            const jsCalendar = {
                "uid": "test-uid",
                "title": "Monthly Unit Test",
                "start": startString,
                "timeZone": "UTC",
                "duration": "PT1H00M",
                recurrenceRules: [{
                    "@type": "RecurrenceRule",
                    "frequency": "monthly",
                    "byDay": [{ "@type": "NDay", day: "tu", nthOfPeriod: 1 }],
                    "count": 8
                }],
                recurrenceOverrides: {
                    '2021-07-06T09:00:00': {
                        "duration": "PT1H30M"
                    }
                }
            };

            const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
            const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
            expect(manager).to.not.be.undefined();

            const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape) as CalendarItemType;
            expect(ewsItem.Recurrence).to.not.be.undefined();
            expect(ewsItem.ItemId?.Id).to.not.be.undefined();
            expect(ewsItem.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));
            expect(pimItem.start).to.not.be.undefined();
            expect(pimItem.end).to.not.be.undefined();
            expect(compareDateStrings(new Date(pimItem.start!).toISOString(), new Date("2021-05-04T09:00:00Z").toISOString())).to.be.equal(0, `${pimItem.start} is not correct start time for the first event`);
            expect(compareDateStrings(new Date(pimItem.end!).toISOString(), new Date("2021-05-04T10:00:00Z").toISOString())).to.be.equal(0, `${pimItem.end} is not correct end time for the first event`);


            // Validate modify occurrences
            expect(ewsItem.ModifiedOccurrences).to.not.be.undefined();
            expect(ewsItem.ModifiedOccurrences?.Occurrence.length).to.be.equal(1);
            expect(compareDateStrings(ewsItem.ModifiedOccurrences!.Occurrence[0].Start.toISOString(), new Date("2021-07-06T09:00:00Z").toISOString())).to.be.equal(0, `${ewsItem.ModifiedOccurrences!.Occurrence[0].Start.toUTCString()} is not correct start time`);
            expect(compareDateStrings(ewsItem.ModifiedOccurrences!.Occurrence[0].End.toISOString(), new Date("2021-07-06T10:30:00Z").toISOString())).to.be.equal(0, `${ewsItem.ModifiedOccurrences!.Occurrence[0].End.toUTCString()} is not correct end time`);

        });

        it('Modify start and end', async () => {
            // Start of first event
            const startString = '2021-05-04T09:00:00';

            // Every month on the 1st Tue of the month at 9 for 1 hour. On July 6th it starts at 10 and lasts 1.5 hours. 
            const jsCalendar = {
                "uid": "test-uid",
                "title": "Monthly Unit Test",
                "start": startString,
                "timeZone": "UTC",
                "duration": "PT1H00M",
                recurrenceRules: [{
                    "@type": "RecurrenceRule",
                    "frequency": "monthly",
                    "byDay": [{ "@type": "NDay", day: "tu", nthOfPeriod: 1 }],
                    "count": 8
                }],
                recurrenceOverrides: {
                    '2021-07-06T09:00:00': {
                        "start": '2021-07-06T10:00:00',
                        "duration": "PT1H30M"
                    }
                }
            };

            const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
            const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
            expect(manager).to.not.be.undefined();

            const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape) as CalendarItemType;
            expect(ewsItem.Recurrence).to.not.be.undefined();
            expect(ewsItem.ItemId?.Id).to.not.be.undefined();
            expect(ewsItem.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));
            expect(pimItem.start).to.not.be.undefined();
            expect(pimItem.end).to.not.be.undefined();
            expect(compareDateStrings(new Date(pimItem.start!).toISOString(), new Date("2021-05-04T09:00:00Z").toISOString())).to.be.equal(0, `${pimItem.start} is not correct start time for the first event`);
            expect(compareDateStrings(new Date(pimItem.end!).toISOString(), new Date("2021-05-04T10:00:00Z").toISOString())).to.be.equal(0, `${pimItem.end} is not correct end time for the first event`);


            // Validate modify occurrences
            expect(ewsItem.ModifiedOccurrences).to.not.be.undefined();
            expect(ewsItem.ModifiedOccurrences?.Occurrence.length).to.be.equal(1);
            expect(compareDateStrings(ewsItem.ModifiedOccurrences!.Occurrence[0].Start.toISOString(), new Date("2021-07-06T10:00:00Z").toISOString())).to.be.equal(0, `${ewsItem.ModifiedOccurrences!.Occurrence[0].Start.toUTCString()} is not correct start time`);
            expect(compareDateStrings(ewsItem.ModifiedOccurrences!.Occurrence[0].End.toISOString(), new Date("2021-07-06T11:30:00Z").toISOString())).to.be.equal(0, `${ewsItem.ModifiedOccurrences!.Occurrence[0].End.toUTCString()} is not correct end time`);

        });

        it('Multiple modifications', async () => {
            // Start of first event
            const startString = '2021-05-04T09:00:00';

            // Every month on the 1st Tue of the month at 9 for 1 hour. 
            // On June 1st it starts at 10 and lasts 1.5 hours. On July 6th it starts at 9 and last for 1.5 hours. 
            const jsCalendar = {
                "uid": "test-uid",
                "title": "Monthly Unit Test",
                "start": startString,
                "timeZone": "UTC",
                "duration": "PT1H00M",
                recurrenceRules: [{
                    "@type": "RecurrenceRule",
                    "frequency": "monthly",
                    "byDay": [{ "@type": "NDay", day: "tu", nthOfPeriod: 1 }],
                    "count": 8
                }],
                recurrenceOverrides: {
                    '2021-06-01T09:00:00': {
                        "start": '2021-06-01T10:00:00',
                        "duration": "PT1H30M"
                    },
                    '2021-07-06T09:00:00': {
                        "duration": "PT1H30M"
                    }
                }
            };

            const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
            const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
            expect(manager).to.not.be.undefined();

            const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape) as CalendarItemType;
            expect(ewsItem.Recurrence).to.not.be.undefined();
            expect(ewsItem.ItemId?.Id).to.not.be.undefined();
            expect(ewsItem.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));
            expect(pimItem.start).to.not.be.undefined();
            expect(pimItem.end).to.not.be.undefined();
            expect(compareDateStrings(new Date(pimItem.start!).toISOString(), new Date("2021-05-04T09:00:00Z").toISOString())).to.be.equal(0, `${pimItem.start} is not correct start time for the first event`);
            expect(compareDateStrings(new Date(pimItem.end!).toISOString(), new Date("2021-05-04T10:00:00Z").toISOString())).to.be.equal(0, `${pimItem.end} is not correct end time for the first event`);


            // Validate modify occurrences
            expect(ewsItem.ModifiedOccurrences).to.not.be.undefined();
            expect(ewsItem.ModifiedOccurrences?.Occurrence.length).to.be.equal(2);
            expect(compareDateStrings(ewsItem.ModifiedOccurrences!.Occurrence[0].Start.toISOString(), new Date("2021-06-01T10:00:00Z").toISOString())).to.be.equal(0, `${ewsItem.ModifiedOccurrences!.Occurrence[0].Start.toUTCString()} is not correct start time`);
            expect(compareDateStrings(ewsItem.ModifiedOccurrences!.Occurrence[0].End.toISOString(), new Date("2021-06-01T11:30:00Z").toISOString())).to.be.equal(0, `${ewsItem.ModifiedOccurrences!.Occurrence[0].End.toUTCString()} is not correct end time`);
            expect(compareDateStrings(ewsItem.ModifiedOccurrences!.Occurrence[1].Start.toISOString(), new Date("2021-07-06T09:00:00Z").toISOString())).to.be.equal(0, `${ewsItem.ModifiedOccurrences!.Occurrence[1].Start.toUTCString()} is not correct start time`);
            expect(compareDateStrings(ewsItem.ModifiedOccurrences!.Occurrence[1].End.toISOString(), new Date("2021-07-06T10:30:00Z").toISOString())).to.be.equal(0, `${ewsItem.ModifiedOccurrences!.Occurrence[1].End.toUTCString()} is not correct end time`);

        });

        it('Modifications dont affect start/end', async () => {
            // Start of first event
            const startString = '2021-05-04T09:00:00';

            // Every month on the 1st Tue of the month at 9 for 1 hour. 
            // On June 1st it starts at 10 and lasts 1.5 hours. On July 6th it starts at 9 and last for 1.5 hours. 
            const jsCalendar = {
                "uid": "test-uid",
                "title": "Monthly Unit Test",
                "start": startString,
                "timeZone": "UTC",
                "duration": "PT1H00M",
                "participants": {
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
                    },
                    "em9lQGZvb2GFtcGxlLmNvbQ": {
                        "@type": "Participant",
                        "name": "Zoe Zelda",
                        "email": "zoe@foobar.example.com",
                        "sendTo": {
                            "imip": "mailto:zoe@foobar.example.com"
                        },
                        "participationStatus": "accepted",
                        "roles": {
                            "owner": true,
                            "attendee": true,
                            "chair": true
                        }
                    }
                },
                recurrenceRules: [{
                    "@type": "RecurrenceRule",
                    "frequency": "monthly",
                    "byDay": [{ "@type": "NDay", day: "tu", nthOfPeriod: 1 }],
                    "count": 8
                }],
                recurrenceOverrides: {
                    '2021-06-01T09:00:00': {
                        "participants/em9lQGZvb2GFtcGxlLmNvbQ/participationStatus": "declined"
                    },
                    '2021-07-06T09:00:00': {
                        "title": "Monthly Unit Test (optional)"
                    }
                }
            };

            const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
            const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
            expect(manager).to.not.be.undefined();

            const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape) as CalendarItemType;
            expect(ewsItem.Recurrence).to.not.be.undefined();
            expect(ewsItem.ItemId?.Id).to.not.be.undefined();
            expect(ewsItem.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));
            expect(pimItem.start).to.not.be.undefined();
            expect(pimItem.end).to.not.be.undefined();
            expect(compareDateStrings(new Date(pimItem.start!).toISOString(), new Date("2021-05-04T09:00:00Z").toISOString())).to.be.equal(0, `${pimItem.start} is not correct start time for the first event`);
            expect(compareDateStrings(new Date(pimItem.end!).toISOString(), new Date("2021-05-04T10:00:00Z").toISOString())).to.be.equal(0, `${pimItem.end} is not correct end time for the first event`);


            // Validate modify occurrences. Should be created even if start/end did not change.
            expect(ewsItem.ModifiedOccurrences).to.not.be.undefined();
            expect(ewsItem.ModifiedOccurrences?.Occurrence.length).to.be.equal(2);
            expect(compareDateStrings(ewsItem.ModifiedOccurrences!.Occurrence[0].Start.toISOString(), new Date("2021-06-01T09:00:00Z").toISOString())).to.be.equal(0, `${ewsItem.ModifiedOccurrences!.Occurrence[0].Start.toUTCString()} is not correct start time`);
            expect(compareDateStrings(ewsItem.ModifiedOccurrences!.Occurrence[0].End.toISOString(), new Date("2021-06-01T10:00:00Z").toISOString())).to.be.equal(0, `${ewsItem.ModifiedOccurrences!.Occurrence[0].End.toUTCString()} is not correct end time`);
            expect(compareDateStrings(ewsItem.ModifiedOccurrences!.Occurrence[1].Start.toISOString(), new Date("2021-07-06T09:00:00Z").toISOString())).to.be.equal(0, `${ewsItem.ModifiedOccurrences!.Occurrence[1].Start.toUTCString()} is not correct start time`);
            expect(compareDateStrings(ewsItem.ModifiedOccurrences!.Occurrence[1].End.toISOString(), new Date("2021-07-06T10:00:00Z").toISOString())).to.be.equal(0, `${ewsItem.ModifiedOccurrences!.Occurrence[1].End.toUTCString()} is not correct end time`);

        });

        it('Exclude modification', async () => {
            // Start of first event
            const startString = '2021-05-04T09:00:00';

            // Every month on the 1st Tue of the month at 9 for 1 hour. 
            // The meeting does not occur on July 6th 
            const jsCalendar = {
                "uid": "test-uid",
                "title": "Monthly Unit Test",
                "start": startString,
                "timeZone": "UTC",
                "duration": "PT1H00M",
                recurrenceRules: [{
                    "@type": "RecurrenceRule",
                    "frequency": "monthly",
                    "byDay": [{ "@type": "NDay", day: "tu", nthOfPeriod: 1 }],
                    "count": 8
                }],
                recurrenceOverrides: {
                    '2021-07-06T09:00:00': {
                        "excluded": true
                    }
                }
            };

            const pimItem = PimItemFactory.newPimCalendarItem(jsCalendar, "default");
            const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
            expect(manager).to.not.be.undefined();

            const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape) as CalendarItemType;
            expect(ewsItem.Recurrence).to.not.be.undefined();
            expect(ewsItem.ItemId?.Id).to.not.be.undefined();
            expect(ewsItem.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));
            expect(pimItem.start).to.not.be.undefined();
            expect(pimItem.end).to.not.be.undefined();
            expect(compareDateStrings(new Date(pimItem.start!).toISOString(), new Date("2021-05-04T09:00:00Z").toISOString())).to.be.equal(0, `${pimItem.start} is not correct start time for the first event`);
            expect(compareDateStrings(new Date(pimItem.end!).toISOString(), new Date("2021-05-04T10:00:00Z").toISOString())).to.be.equal(0, `${pimItem.end} is not correct end time for the first event`);


            // Validate deleted occurrences
            expect(ewsItem.DeletedOccurrences).to.not.be.undefined();
            const expected = [new Date("2021-07-06T09:00:00Z").toISOString()];
            expect(ewsItem.DeletedOccurrences!.DeletedOccurrence.map(occurence => occurence.Start.toISOString())).to.deepEqual(expected);

            expect(ewsItem.ModifiedOccurrences).to.be.undefined(); // Should not have added any modifications
        });

    });

    describe('EWSCalendarManager.getAttendeeResponseType', function () {
        let pimItem: PimCalendarItem;
        let ownerParticipant: any;
        let optionalParticipant: any;
        let fyiParticipant: any;
        let attendeeParticipant: any;
        let unresponsiveParticipant: any;
    
        beforeEach(function () {
            ownerParticipant = {
                email: 'owner@hcl.com',
                type: 'Participant',
                roles: {
                    owner: true
                }
            }
    
            optionalParticipant = {
                email: 'optional@hcl.com',
                type: 'Participant',
                roles: {
                    optional: true
                },
                participationStatus: 'declined'
            }
    
            fyiParticipant = {
                email: 'fyi@hcl.com',
                type: 'Participant',
                roles: {
                    informational: true
                },
                participationStatus: 'tentative'
            }
    
            attendeeParticipant = {
                email: 'attendee@hcl.com',
                type: 'Participant',
                roles: {
                    attendee: true
                },
                participationStatus: 'accepted'
            }
    
            unresponsiveParticipant = {
                email: 'norespone@hcl.com',
                type: 'Participant',
                roles: {
                    attendee: true
                }
            }
            const pimObject: any = {
                start: new Date(),
                duration: 'P1D',
                participants: {
                    "ABC": ownerParticipant,
                    "DEF": attendeeParticipant,
                    "GHI": optionalParticipant,
                    "JKL": fyiParticipant,
                    "MNO": unresponsiveParticipant
                }
            }
            pimItem = PimItemFactory.newPimCalendarItem(pimObject, "default", PimItemFormat.DOCUMENT);
        });

        it('matching attendee', () => {
            expect(new MockEWSCalendarManager().getAttendeeResponseType(unresponsiveParticipant.email, pimItem)).to.eql(ResponseTypeType.NO_RESPONSE_RECEIVED);
        });

        it('unknown attendee', () => {

            const attendee = 'invalid@notvalid.com';
            const result = new MockEWSCalendarManager().getAttendeeResponseType(attendee, pimItem);
            expect(result).to.eql(ResponseTypeType.NO_RESPONSE_RECEIVED);
        });

        it('organizer', () => {
            expect(new MockEWSCalendarManager().getAttendeeResponseType(ownerParticipant.email, pimItem)).to.eql(ResponseTypeType.ORGANIZER);
        });

        it('accepted', () => {
            expect(new MockEWSCalendarManager().getAttendeeResponseType(attendeeParticipant.email, pimItem)).to.eql(ResponseTypeType.ACCEPT);
        });

        it('decline', () => {
            expect(new MockEWSCalendarManager().getAttendeeResponseType(optionalParticipant.email, pimItem)).to.eql(ResponseTypeType.DECLINE);
        });

        it('tentative', () => {
            expect(new MockEWSCalendarManager().getAttendeeResponseType(fyiParticipant.email, pimItem)).to.eql(ResponseTypeType.TENTATIVE);
        });
    });
});
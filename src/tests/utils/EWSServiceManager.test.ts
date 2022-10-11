import { expect, sinon } from '@loopback/testlab';
import {
    PimItem, PimTask, PimLabel, PimLabelTypes, PimTaskPriority, UserInfo, PimImportance, PimItemFormat, KeepPimManager,
    KeepPimNotebookManager, PimItemFactory, getTrimmedISODate, KeepPimBaseResults, KeepPimMessageManager, PimNote, PimCalendarItem, PimContact
} from '@hcllabs/openclientkeepcomponent';
import {
    DefaultShapeNamesType, ItemClassType, MapiPropertyTypeType, UnindexedFieldURIType, ImportanceChoicesType,
    SensitivityChoicesType, ExceptionPropertyURIType, MessageDispositionType, ResponseClassType, ResponseCodeType
} from '../../models/enum.model';
import {
    ContactItemType, ItemResponseShapeType, MessageType, ItemType, TaskType, NonEmptyArrayOfPathsToElementType,
    ExtendedPropertyType, ItemChangeType, CalendarItemType, ItemIdType, NonEmptyArrayOfItemChangeDescriptionsType, 
    SetItemFieldType
} from '../../models/mail.model';
import { PathToExceptionFieldType, PathToExtendedFieldType, PathToUnindexedFieldType } from '../../models/common.model';
import { 
    EWSServiceManager, EWSNotesManager, EWSTaskManager, EWSMessageManager, EWSContactsManager, EWSCalendarManager, 
    getEWSId 
} from '../../utils';
import { UserContext } from '../../keepcomponent';
import { stubNotebookManager, stubPimManager, getContext } from '../unitTestHelper';
import { Request } from '@loopback/rest';

describe('EWSServiceManager tests', () => {
    
    const testUser = "test.user@test.org";

    // Generic Mock subclass of EWSService manager for testing protected functions
    class MockEWSManager extends EWSServiceManager {

        createItem(
            item: ItemType, 
            userInfo: UserInfo, 
            request: Request, 
            disposition?: MessageDispositionType, 
            toLabel?: PimLabel,
            mailboxId?: string
        ): Promise<ItemType[]> {
            throw (new Error('Not yet implamented'));
        }

        async deleteItem(item: string | PimItem, userInfo: UserInfo, mailboxId?: string, hardDelete = false): Promise<void> {
            throw (new Error('Not yet implamented'));
        }

        async updateItem(pimItem: PimItem, change: ItemChangeType, userInfo: UserInfo, request: Request, toLabel?: PimLabel): Promise<ItemType | undefined> {
            throw (new Error('Not yet implamented'));
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
            throw (new Error('Not yet implamented'));
        }

        async copyItem(
            pimItem: PimItem, 
            toFolderId: string, 
            userInfo: UserInfo, 
            request: Request, 
            shape: ItemResponseShapeType,
            mailboxId?: string
        ): Promise<ItemType> {
            throw (new Error('Not yet implamented'));
        }

        protected async createNewPimItem(pimItem: PimItem, userInfo: UserInfo, mailboxId?: string): Promise<PimItem> {
            throw (new Error('Not yet implamented'));
        }

        protected fieldsForDefaultShape(): UnindexedFieldURIType[] {
            // Doesn't matter what fields we use as we're not really testing the results in the base class tests.
            return [
                ...EWSServiceManager.idOnlyFields,
                UnindexedFieldURIType.TASK_DUE_DATE,
                UnindexedFieldURIType.TASK_PERCENT_COMPLETE,
                UnindexedFieldURIType.TASK_START_DATE,
                UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS,
                UnindexedFieldURIType.ITEM_SUBJECT,
                UnindexedFieldURIType.TASK_STATUS
            ];
        }

        protected fieldsForAllPropertiesShape(): UnindexedFieldURIType[] {
            // Use a subset of the Task fields
            return [
                ...EWSServiceManager.idOnlyFields,
                UnindexedFieldURIType.TASK_DUE_DATE,
                UnindexedFieldURIType.TASK_PERCENT_COMPLETE,
                UnindexedFieldURIType.TASK_START_DATE,
                UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS,
                UnindexedFieldURIType.ITEM_SUBJECT,
                UnindexedFieldURIType.TASK_STATUS,
                UnindexedFieldURIType.ITEM_BODY,
                UnindexedFieldURIType.ITEM_CATEGORIES,
                UnindexedFieldURIType.ITEM_LAST_MODIFIED_TIME,
                UnindexedFieldURIType.ITEM_LAST_MODIFIED_NAME,
                UnindexedFieldURIType.TASK_ACTUAL_WORK,
                UnindexedFieldURIType.TASK_ASSIGNED_TIME,
                UnindexedFieldURIType.TASK_BILLING_INFORMATION,
                UnindexedFieldURIType.TASK_CHANGE_COUNT,
                UnindexedFieldURIType.TASK_COMPANIES,
                UnindexedFieldURIType.TASK_COMPLETE_DATE,
                UnindexedFieldURIType.TASK_CONTACTS
            ];
        }

        pimItemToEWSItem(
            pimITem: PimItem, 
            userInfo: UserInfo, 
            request: Request, 
            shape: ItemResponseShapeType,
            mailboxId?: string,
            targetFolderId?: string): Promise<ItemType> {
            throw (new Error('Not yet implamented'));
        }

        pimItemFromEWSItem(ewsItem: ItemType, request: Request, existing?: object): PimItem {
            throw (new Error('Not yet implamented'));
        }

        addRequestedPropertiesToEWSItem(pimItem: PimItem, toItem: ItemType, shape?: ItemResponseShapeType, mailboxId?: string): void {
            super.addRequestedPropertiesToEWSItem(pimItem, toItem, shape, mailboxId);
        }

        updateEWSItemFieldValue(item: ItemType, pimTask: PimItem, fieldId: UnindexedFieldURIType, mailboxId?: string): boolean {
            return super.updateEWSItemFieldValue(item, pimTask, fieldId, mailboxId);
        }

    }

    describe('Test getInstance functions', () => {

        it("getInstanceFromPimItem", function () {

            let pimItem: PimItem = PimItemFactory.newPimNote({});
            let manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
            expect(manager).to.be.instanceof(EWSNotesManager);

            pimItem = PimItemFactory.newPimTask({});
            manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
            expect(manager).to.be.instanceof(EWSTaskManager);

            pimItem = PimItemFactory.newPimMessage({});
            manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
            expect(manager).to.be.instanceof(EWSMessageManager);

            pimItem = PimItemFactory.newPimContact({});
            manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
            expect(manager).to.be.instanceof(EWSContactsManager);

            pimItem = PimItemFactory.newPimCalendarItem({unid: 'UNID'}, 'default', PimItemFormat.DOCUMENT);
            manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
            expect(manager).to.be.instanceOf(EWSCalendarManager);

            // Using labels
            const pimLabel = PimItemFactory.newPimLabel({}, PimItemFormat.DOCUMENT);
            pimLabel.type = PimLabelTypes.TASKS;
            manager = EWSServiceManager.getInstanceFromPimItem(pimLabel);
            expect(manager).to.be.instanceof(EWSTaskManager);

            pimLabel.type = PimLabelTypes.JOURNAL;
            manager = EWSServiceManager.getInstanceFromPimItem(pimLabel);
            expect(manager).to.be.instanceof(EWSNotesManager);

            pimLabel.type = PimLabelTypes.MAIL;
            manager = EWSServiceManager.getInstanceFromPimItem(pimLabel);
            expect(manager).to.be.instanceof(EWSMessageManager);

            pimLabel.type = PimLabelTypes.CONTACTS;
            manager = EWSServiceManager.getInstanceFromPimItem(pimLabel);
            expect(manager).to.be.instanceof(EWSContactsManager);

            pimLabel.type = PimLabelTypes.CALENDAR;
            manager = EWSServiceManager.getInstanceFromPimItem(pimLabel);
            expect(manager).to.be.instanceOf(EWSCalendarManager);
        });

        it("getInstanceFromItem", function () {

            let item: ItemType = new MessageType();
            item.ItemClass = ItemClassType.ITEM_CLASS_NOTE;
            let manager = EWSServiceManager.getInstanceFromItem(item);
            expect(manager).to.be.instanceof(EWSNotesManager);

            item = new ItemType();
            item.ItemClass = ItemClassType.ITEM_CLASS_NOTE;
            manager = EWSServiceManager.getInstanceFromItem(item);
            expect(manager).to.be.instanceof(EWSNotesManager);

            item = new TaskType();
            manager = EWSServiceManager.getInstanceFromItem(item);
            expect(manager).to.be.instanceof(EWSTaskManager);

            item = new ContactItemType();
            manager = EWSServiceManager.getInstanceFromItem(item);
            expect(manager).to.be.instanceOf(EWSContactsManager);

            item = new MessageType();
            item.ItemClass = ItemClassType.ITEM_CLASS_MESSAGE;
            manager = EWSServiceManager.getInstanceFromItem(item);
            expect(manager).to.be.instanceof(EWSMessageManager);

            item = new CalendarItemType();
            manager = EWSServiceManager.getInstanceFromItem(item);
            expect(manager).to.be.instanceOf(EWSCalendarManager);

        });
    });

    describe('Test getItem', () => {

        // Clean up the stubs
        afterEach(() => {
            sinon.restore(); // Reset stubs
        });

        it('Expected results with default shape', async function () {
            const pimNote = PimItemFactory.newPimNote({});
            pimNote.unid = 'UNID';
            pimNote.subject = 'SUBJECT';
            pimNote.body = 'BODY';
            pimNote.parentFolderIds = ['PARENT'];
            const pimManagerStub = sinon.createStubInstance(KeepPimManager);
            pimManagerStub.getPimItem.resolves(pimNote);
            stubPimManager(pimManagerStub);

            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.DEFAULT;
            const context = getContext(testUser, "password");
            const mailboxId = 'test@test.com';
            const itemEWSId = getEWSId('UNID', mailboxId);

            const note = await EWSServiceManager.getItem(itemEWSId, new UserContext(), context.request, shape);

            expect(note).to.not.be.undefined();
            if (note) {
                expect(note.ItemId?.Id).to.be.equal(itemEWSId);
                expect(note.Subject).to.be.equal('SUBJECT');
            }
        });

        it('Expected results with no shape', async () => {
            let note: MessageType | undefined = undefined;
            const pimNote = PimItemFactory.newPimNote({});
            pimNote.unid = 'UNID';
            pimNote.subject = 'SUBJECT';
            pimNote.body = 'BODY';
            pimNote.parentFolderIds = ['PARENT'];
            const pimManagerStub = sinon.createStubInstance(KeepPimManager);
            pimManagerStub.getPimItem.resolves(pimNote);
            stubPimManager(pimManagerStub);

            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.DEFAULT;
            const context = getContext(testUser, 'password');
            const itemEWSId = getEWSId('UNID');
            note = await EWSServiceManager.getItem(itemEWSId, new UserContext(), context.request);

            expect(note).to.not.be.undefined();
            if (note) {
                expect(note.ItemId?.Id).to.be.equal(itemEWSId);
                expect(note.Subject).to.be.equal('SUBJECT');
            }
        });

        it('Unsupported type', async () => {
            const pimThread = PimItemFactory.newPimThread({}, PimItemFormat.DOCUMENT);
            const pimManagerStub = sinon.createStubInstance(KeepPimManager);
            pimManagerStub.getPimItem.resolves(pimThread);
            stubPimManager(pimManagerStub);

            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.DEFAULT;
            const context = getContext(testUser, 'password');
            const itemEWSId = getEWSId('UNID');
            const note = await EWSServiceManager.getItem(itemEWSId, new UserContext(), context.request);
            expect(note).to.be.undefined();
        });

        it('Item not found', async () => {
            const pimManagerStub = sinon.createStubInstance(KeepPimManager);
            pimManagerStub.getPimItem.resolves(undefined);
            stubPimManager(pimManagerStub);

            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.DEFAULT;
            const context = getContext(testUser, 'password');
            const itemEWSId = getEWSId('UNID');
            const note = await EWSServiceManager.getItem(itemEWSId, new UserContext(), context.request);
            expect(note).to.be.undefined();
        });
    });

    describe('Test updateItems', () => {

        const context = getContext(testUser, "password");

        // Clean up the stubs
        afterEach(() => {
            sinon.restore(); // Reset stubs
        });

        it('Expected results', async () => {
            const pimNote = PimItemFactory.newPimNote({});
            pimNote.unid = 'UNID';
            pimNote.subject = 'SUBJECT';
            pimNote.body = 'BODY';
            pimNote.parentFolderIds = ['PARENT'];
            const pimManagerStub = sinon.createStubInstance(KeepPimManager);
            pimManagerStub.getPimItem.resolves(pimNote);
            stubPimManager(pimManagerStub);

            // const pimNote2 = PimItemFactory.newPimNote({});
            // pimNote2.unid = 'UNID';
            // pimNote2.parentFolderIds = ['PARENT'];
            const notebookManagerStub = sinon.createStubInstance(KeepPimNotebookManager);
            const updateBaseResults: KeepPimBaseResults = {
                unid: 'UNID',
                status: 200,
                message: 'Success'
            };
            notebookManagerStub.updateNote.resolves(updateBaseResults);
            stubNotebookManager(notebookManagerStub);

            const change = new ItemChangeType();
            const mailboxId = 'test@test.com';
            const itemEWSId = getEWSId('UNID', mailboxId);
            change.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            change.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();
            const subjectChange = new SetItemFieldType();
            subjectChange.FieldURI = new PathToUnindexedFieldType();
            subjectChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            subjectChange.Item = new ItemType();
            subjectChange.Item.Subject = 'UPDATED SUBJECT';
            change.Updates.push(subjectChange);
            const userInfo = new UserContext();

            const updatedNote = await EWSServiceManager.updateItem(change, userInfo, context.request);

            expect(updatedNote).to.not.be.undefined();
            if (updatedNote) {
                expect(updatedNote.ItemId?.Id).to.be.equal(itemEWSId);
            }

        });

        it('No itemId set on change', async () => {
            // Can't have an await in a test, so do all te work in the before and just check the results in the it
            let success = false;
            const pimNote = PimItemFactory.newPimNote({});
            pimNote.unid = 'UNID';
            pimNote.subject = 'SUBJECT';
            pimNote.body = 'BODY';
            pimNote.parentFolderIds = ['PARENT'];
            const pimManagerStub = sinon.createStubInstance(KeepPimManager);
            pimManagerStub.getPimItem.resolves(undefined);
            stubPimManager(pimManagerStub);

            const change = new ItemChangeType();
            const itemEWSId = getEWSId('UNID');
            change.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            change.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();
            const subjectChange = new SetItemFieldType();
            subjectChange.FieldURI = new PathToUnindexedFieldType();
            subjectChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            subjectChange.Item = new ItemType();
            subjectChange.Item.Subject = 'UPDATED SUBJECT';
            change.Updates.push(subjectChange);
            const userInfo = new UserContext();
            try {
                const updatedNote = await EWSServiceManager.updateItem(change, userInfo, context.request)
                success = updatedNote === undefined;
            } catch (err) {
                success = true;
            }
            expect(success).to.be.true();
        });

        it('Item not found', async () => {
            // Can't have an await in a test, so do all te work in the before and just check the results in the it
            let success = false;
            const change = new ItemChangeType();
            change.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();
            const subjectChange = new SetItemFieldType();
            subjectChange.FieldURI = new PathToUnindexedFieldType();
            subjectChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            change.Updates.push(subjectChange);
            const userInfo = new UserContext();
            try {
                const updatedNote = await EWSServiceManager.updateItem(change, userInfo, context.request);
                success = updatedNote === undefined;
            } catch (err) {
                success = true;
            }
            expect(success).to.be.true();
        });

        it('Unsupported item type', async () => {
            const pimThread = PimItemFactory.newPimThread({}, PimItemFormat.DOCUMENT);
            const pimManagerStub = sinon.createStubInstance(KeepPimManager);
            pimManagerStub.getPimItem.resolves(pimThread);
            stubPimManager(pimManagerStub);

            const change = new ItemChangeType();
            const itemEWSId = getEWSId('UNID');
            change.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            change.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();
            const subjectChange = new SetItemFieldType();
            subjectChange.FieldURI = new PathToUnindexedFieldType();
            subjectChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            subjectChange.Item = new ItemType();
            subjectChange.Item.Subject = 'UPDATED SUBJECT';
            change.Updates.push(subjectChange);
            const userInfo = new UserContext();
            
            const err: any = new Error();
            err.ResponseClass = ResponseClassType.ERROR;
            err.ResponseCode = ResponseCodeType.ERROR_ITEM_NOT_FOUND;
            err.MessageText = 'Item not found: UNID';
            await expect(EWSServiceManager.updateItem(change, userInfo, context.request)).to.be.rejectedWith(err);
        });
    });

    describe('Test moveItem', () => {
        const userInfo = new UserContext();
        const toFolderId = 'targetFolderId';
        const context = getContext(testUser, "password");
        const testId = 'UNID';
        const mailboxId = 'test@test.com';
        const testEWSId = getEWSId('UNID', mailboxId);
        const itemIdType = new ItemIdType(testEWSId, `ck-${testEWSId}`);

        beforeEach(function() {
            //Stubs
            sinon.stub(KeepPimMessageManager.getInstance(), "moveMessages").resolves({});
            sinon.stub(EWSNotesManager.getInstance(), "deleteItem").resolvesThis();
        })

        afterEach(function() {
            sinon.restore();
        })
        
        it("moveItem for a note", async function () {

            const pItem = PimItemFactory.newPimNote({"@unid": testId});
            sinon.stub(KeepPimManager.getInstance(), "getPimItem").resolves(pItem);

            // Test moveItem without itemId
            let item: ItemType = new MessageType();
            item.ItemClass = ItemClassType.ITEM_CLASS_NOTE;
            let result = await EWSServiceManager.moveItem(item, toFolderId, userInfo, context.request);
            expect(result).to.be.undefined();

            // Test moveItem with itemId
            item.ItemId = itemIdType;
            result = await EWSServiceManager.moveItem(item, toFolderId, userInfo, context.request);
            expect(result).to.be.equal(testEWSId);


            // Test moveItemWithResult without itemId
            item = new MessageType();
            item.ItemClass = ItemClassType.ITEM_CLASS_NOTE;
            let itemTypeResult = await EWSServiceManager.moveItemWithResult(item, toFolderId, userInfo, context.request);
            expect(itemTypeResult).to.be.undefined();

            // Test moveItemWithResult with itemId
            item.ItemId = itemIdType;
            itemTypeResult = await EWSServiceManager.moveItemWithResult(item, toFolderId, userInfo, context.request);
            expect(itemTypeResult instanceof MessageType).to.be.true();
            expect(itemTypeResult?.ItemClass).to.be.equal(ItemClassType.ITEM_CLASS_NOTE);

        });

        it("moveItem for a task", async function () {

            const pItem = PimItemFactory.newPimTask({"@unid": testId, "uid": testId});
            sinon.stub(KeepPimManager.getInstance(), "getPimItem").resolves(pItem);

            // Test moveItem without itemId
            let item: ItemType = new TaskType();
            let result = await EWSServiceManager.moveItem(item, toFolderId, userInfo, context.request);
            expect(result).to.be.undefined();

            // Test moveItem with itemId
            item.ItemId = itemIdType;
            result = await EWSServiceManager.moveItem(item, toFolderId, userInfo, context.request);
            expect(result).to.be.equal(testEWSId);


            // Test moveItemWithResult without itemId
            item = new TaskType();
            let itemTypeResult = await EWSServiceManager.moveItemWithResult(item, toFolderId, userInfo, context.request);
            expect(itemTypeResult).to.be.undefined();

            // Test moveItemWithResult with itemId
            item.ItemId = itemIdType;
            itemTypeResult = await EWSServiceManager.moveItemWithResult(item, toFolderId, userInfo, context.request);
            expect(itemTypeResult instanceof TaskType).to.be.true();

        });

        it("moveItem for a message", async function () {
            const pItem = PimItemFactory.newPimMessage({"@unid": testId});
            sinon.stub(KeepPimManager.getInstance(), "getPimItem").resolves(pItem);
            sinon.stub(KeepPimMessageManager.getInstance(), "getMimeMessage").resolves("A mime string");

            // Test moveItem without itemId
            let item: ItemType = new MessageType();
            item.ItemClass = ItemClassType.ITEM_CLASS_MESSAGE;
            let result = await EWSServiceManager.moveItem(item, toFolderId, userInfo, context.request);
            expect(result).to.be.undefined();

            // Test moveItem with itemId
            item.ItemId = itemIdType;
            result = await EWSServiceManager.moveItem(item, toFolderId, userInfo, context.request);
            expect(result).to.be.equal(testEWSId);


            // Test moveItemWithResult without itemId
            item = new MessageType();
            item.ItemClass = ItemClassType.ITEM_CLASS_MESSAGE;
            let itemTypeResult = await EWSServiceManager.moveItemWithResult(item, toFolderId, userInfo, context.request);
            expect(itemTypeResult).to.be.undefined();

            // Test moveItemWithResult with itemId
            item.ItemId = itemIdType;
            itemTypeResult = await EWSServiceManager.moveItemWithResult(item, toFolderId, userInfo, context.request);
            expect(itemTypeResult instanceof MessageType).to.be.true();
            expect(itemTypeResult?.ItemClass).to.be.equal(ItemClassType.ITEM_CLASS_MESSAGE);

        });

        it("moveItem for a contact", async function () {
            const pItem = PimItemFactory.newPimContact({"@unid": testId});
            sinon.stub(KeepPimManager.getInstance(), "getPimItem").resolves(pItem);

            // Test moveItem without itemId
            let item: ItemType = new ContactItemType();
            let result = await EWSServiceManager.moveItem(item, toFolderId, userInfo, context.request);
            expect(result).to.be.undefined();

            // Test moveItem with itemId
            item.ItemId = itemIdType;
            result = await EWSServiceManager.moveItem(item, toFolderId, userInfo, context.request);
            expect(result).to.be.equal(testEWSId);


            // Test moveItemWithResult without itemId
            item = new TaskType();
            let itemTypeResult = await EWSServiceManager.moveItemWithResult(item, toFolderId, userInfo, context.request);
            expect(itemTypeResult).to.be.undefined();

            // Test moveItemWithResult with itemId
            item.ItemId = itemIdType;
            itemTypeResult = await EWSServiceManager.moveItemWithResult(item, toFolderId, userInfo, context.request);
            expect(itemTypeResult instanceof ContactItemType).to.be.true();
        });

        it("moveItem for a calendar item", async function () {
            const pItem = PimItemFactory.newPimCalendarItem({"@unid": testId, "uid": testId}, "default", PimItemFormat.DOCUMENT);
            sinon.stub(KeepPimManager.getInstance(), "getPimItem").resolves(pItem);

            // Test moveItem without itemId
            let item: ItemType = new CalendarItemType();
            let result = await EWSServiceManager.moveItem(item, toFolderId, userInfo, context.request);
            expect(result).to.be.undefined();

            // Test moveItem with itemId
            item.ItemId = itemIdType;
            result = await EWSServiceManager.moveItem(item, toFolderId, userInfo, context.request);
            expect(result).to.be.equal(testEWSId);


            // Test moveItemWithResult without itemId
            item = new TaskType();
            let itemTypeResult = await EWSServiceManager.moveItemWithResult(item, toFolderId, userInfo, context.request);
            expect(itemTypeResult).to.be.undefined();

            // Test moveItemWithResult with itemId
            item.ItemId = itemIdType;
            itemTypeResult = await EWSServiceManager.moveItemWithResult(item, toFolderId, userInfo, context.request);
            expect(itemTypeResult instanceof CalendarItemType).to.be.true();
        });
    });

    describe('Test defaultPropertiesForShape', () => {

        class MockNotesManager extends EWSNotesManager {
            getFieldsForShape(shape: ItemResponseShapeType): string[] {
                return this.defaultPropertiesForShape(shape);
            }
        }

        it("test shapes", function () {
            const manager = new MockNotesManager();
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;

            // Just driving coverage.  Make sure an array is returned. We'll test contents in the subclass tests.
            let fields = manager.getFieldsForShape(shape);
            expect(fields).to.be.an.Array();

            shape.BaseShape = DefaultShapeNamesType.DEFAULT;
            fields = manager.getFieldsForShape(shape);
            expect(fields).to.be.an.Array();

            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
            fields = manager.getFieldsForShape(shape);
            expect(fields).to.be.an.Array();

            shape.BaseShape = DefaultShapeNamesType.PCX_PEOPLE_SEARCH;
            fields = manager.getFieldsForShape(shape);
            expect(fields).to.be.eql([]);
        });
    });

    describe('Test addRequestedPropertiesToEWSItem', () => {

        let pimTask: PimTask;
        let task: TaskType;
        let itemEWSId: string;
        const mailboxId = 'test@test.com';
        const parentEWSId = getEWSId('PARENT-FOLDER', mailboxId);

        // Reset pimItem before each test
        beforeEach(function () {

            const extendedProp = {
                DistinguishedPropertySetId: 'Address',
                PropertyId: 32899,
                PropertyType: 'String',
                Value: 'kermit.frog@fakemail.com'
            }

            pimTask = PimItemFactory.newPimTask({});
            pimTask.unid = 'UNID';
            pimTask.parentFolderIds = ['PARENT-FOLDER'];
            pimTask.subject = 'SUBJECT';
            pimTask.body = 'BODY';
            pimTask.addExtendedProperty(extendedProp);

            task = new TaskType();

            itemEWSId = getEWSId(pimTask.unid, mailboxId);
        });

        it("base fields for task", function () {
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.DEFAULT;

            const mockManager = new MockEWSManager();
            mockManager.addRequestedPropertiesToEWSItem(pimTask, task, shape, mailboxId);

            expect(task.ItemId?.Id).to.be.equal(itemEWSId);
            expect(task.Subject).to.be.equal(pimTask.subject);
            expect(task.ParentFolderId?.Id).to.be.equal(parentEWSId);
        });

        it("additional fields", function () {
            // Try some additional fields
            const path1 = new PathToUnindexedFieldType();
            path1.FieldURI = UnindexedFieldURIType.ITEM_BODY;

            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            shape.AdditionalProperties = new NonEmptyArrayOfPathsToElementType([path1]);

            const mockManager = new MockEWSManager();
            mockManager.addRequestedPropertiesToEWSItem(pimTask, task, shape, mailboxId);

            expect(task.ItemId?.Id).to.be.equal(itemEWSId);
            expect(task.Body?.Value).to.be.equal(pimTask.body);
            expect(task.Body?.BodyType).to.be.equal('Text');
            expect(task.Subject).to.be.undefined();
            expect(task.ParentFolderId?.Id).to.be.equal(parentEWSId);
        });


        it("extended fields", function () {
            // Try an extended properties
            const extPath = new PathToExtendedFieldType();
            extPath.PropertyId = 32899;
            extPath.PropertyType = MapiPropertyTypeType.STRING;

            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            shape.AdditionalProperties = new NonEmptyArrayOfPathsToElementType([extPath]);

            const mockManager = new MockEWSManager();
            mockManager.addRequestedPropertiesToEWSItem(pimTask, task, shape);

            if (task.ExtendedProperty) {
                const ext: ExtendedPropertyType = task.ExtendedProperty[0];
                expect(ext).is.not.undefined();
                expect(ext.Value).to.be.equal('kermit.frog@fakemail.com');
                expect(ext.ExtendedFieldURI.PropertyId).to.be.equal(32899);
                expect(ext.ExtendedFieldURI.DistinguishedPropertySetId).to.be.equal('Address');
                expect(ext.ExtendedFieldURI.PropertyType).to.be.equal('String');
            }

            // Try with requested property missing
            task = new TaskType();
            pimTask = PimItemFactory.newPimTask({uid: 'test'});
            mockManager.addRequestedPropertiesToEWSItem(pimTask, task, shape);
            expect(task.ExtendedProperty).to.be.eql([]);

            // Try requesting a different field type
            const exceptionPath = new PathToExceptionFieldType();
            exceptionPath.FieldURI = ExceptionPropertyURIType.TIMEZONE_OFFSET;
            shape.AdditionalProperties = new NonEmptyArrayOfPathsToElementType([exceptionPath]);
            mockManager.addRequestedPropertiesToEWSItem(pimTask, task, shape);
            expect(task.ExtendedProperty).to.be.eql([]);

        });

        it("all properties shape", function () {

            // Try an extended properties
            const extPath = new PathToExtendedFieldType();
            extPath.PropertyId = 32899;
            extPath.PropertyType = MapiPropertyTypeType.STRING;

            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
            shape.AdditionalProperties = new NonEmptyArrayOfPathsToElementType([extPath]);

            const mockManager = new MockEWSManager();
            mockManager.addRequestedPropertiesToEWSItem(pimTask, task, shape, mailboxId);

            expect(task.ItemId?.Id).to.be.equal(itemEWSId);
            expect(task.Body?.Value).to.be.equal(pimTask.body);
            expect(task.Body?.BodyType).to.be.equal('Text');
            expect(task.Subject).to.be.equal(pimTask.subject);
            expect(task.ParentFolderId?.Id).to.be.equal(parentEWSId);

            if (task.ExtendedProperty) {
                const ext: ExtendedPropertyType = task.ExtendedProperty[0];
                expect(ext).is.not.undefined();
                expect(ext.Value).to.be.equal('kermit.frog@fakemail.com');
                expect(ext.ExtendedFieldURI.PropertyId).to.be.equal(32899);
                expect(ext.ExtendedFieldURI.DistinguishedPropertySetId).to.be.equal('Address');
                expect(ext.ExtendedFieldURI.PropertyType).to.be.equal('String');
            }

        });

        it("no shape", function () {
            const extPath = new PathToExtendedFieldType();
            extPath.PropertyId = 32899;
            extPath.PropertyType = MapiPropertyTypeType.STRING;

            const mockManager = new MockEWSManager();
            mockManager.addRequestedPropertiesToEWSItem(pimTask, task);
            expect(task.ItemId).to.be.undefined();

            if (task.ExtendedProperty) {
                const ext: ExtendedPropertyType = task.ExtendedProperty[0];
                expect(ext).is.not.undefined();
                expect(ext.Value).to.be.equal('kermit.frog@fakemail.com');
                expect(ext.ExtendedFieldURI.PropertyId).to.be.equal(32899);
                expect(ext.ExtendedFieldURI.DistinguishedPropertySetId).to.be.equal('Address');
                expect(ext.ExtendedFieldURI.PropertyType).to.be.equal('String');
            }
        });
    });

    describe('Test updateEWSItemFieldValue', () => {
        let pimTask: PimTask;
        let task: TaskType;
        let pimNote: PimNote;
        let note: MessageType;
        let pimCalendarItem: PimCalendarItem;
        let calendarItem: CalendarItemType;
        let contact: ContactItemType;
        let pimContact: PimContact;
        let manager: MockEWSManager;
        let itemEWSId: string;
        const mailboxId = 'test@test.com';

        beforeEach(() => {
            pimTask = PimItemFactory.newPimTask({});
            pimTask.unid = 'UNID';
            task = new TaskType();
            pimNote = PimItemFactory.newPimNote({});
            pimNote.unid = 'UNID';
            note = new MessageType();
            pimCalendarItem = PimItemFactory.newPimCalendarItem({}, 'default', PimItemFormat.DOCUMENT);
            pimCalendarItem.unid = 'UNID';
            calendarItem = new CalendarItemType();
            contact = new ContactItemType();
            pimContact = PimItemFactory.newPimContact({});

            manager = new MockEWSManager();

            itemEWSId = getEWSId(pimTask.unid, mailboxId);
        });

        it("UnindexedFieldURIType.ITEM_ITEM_ID", function () {
            const handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_ITEM_ID, mailboxId);

            expect(handled).to.be.true();
            expect(task.ItemId?.Id).to.be.equal(itemEWSId);
            expect(task.ItemId?.ChangeKey).to.be.equal(`ck-${itemEWSId}`);
        });

        it("UnindexedFieldURIType.ITEM_PARENT_FOLDER_ID", function () {
            // No parent id
            let handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_PARENT_FOLDER_ID);

            expect(handled).to.be.true();
            expect(task.ParentFolderId?.Id).to.be.undefined();

            // Empty array of parent ids
            pimTask.parentFolderIds = [];
            handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_PARENT_FOLDER_ID);
            expect(handled).to.be.true();
            expect(task.ParentFolderId?.Id).to.be.undefined();

            // Set a parent id
            pimTask.parentFolderIds = ['PARENT'];
            handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_PARENT_FOLDER_ID, mailboxId);
            expect(handled).to.be.true();
            const parentFolderEWSId = getEWSId('PARENT', mailboxId);
            expect(task.ParentFolderId?.Id).to.be.equal(parentFolderEWSId);
            expect(task.ParentFolderId?.ChangeKey).to.be.equal(`ck-${parentFolderEWSId}`);
        });

        it("UnindexedFieldURIType.ITEM_ITEM_CLASS", function () {
            const handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_ITEM_CLASS);

            expect(handled).to.be.true();
            expect(task.ItemClass).to.be.equal(ItemClassType.ITEM_CLASS_TASK);
        });

        it("UnindexedFieldURIType.ITEM_MIME_CONTENT", function () {
            // TODO:  Not currently supported.  Implement when it is.
            const handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_MIME_CONTENT);

            expect(handled).to.be.true();
        });

        it("UnindexedFieldURIType.ITEM_DATE_TIME_CREATED", function () {
            pimTask.createdDate = new Date();

            const handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_DATE_TIME_CREATED);

            expect(handled).to.be.true();
            expect(task.DateTimeCreated).to.be.eql(pimTask.createdDate);
        });

        it("UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS", function () {
            pimTask.attachments = ['a1', 'a2'];

            const handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS);

            expect(handled).to.be.true();
            expect(task.HasAttachments).to.be.true();
        });

        it("UnindexedFieldURIType.ITEM_ATTACHMENTS", function () {
            // No attachments
            let handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_ATTACHMENTS);
            expect(handled).to.be.true();
            expect(task.Attachments).to.be.undefined();

            // Add some attachments
            pimTask.attachments = ['a1', 'a2'];
            handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_ATTACHMENTS);
            expect(handled).to.be.true();
            const attachments = task.Attachments;
            if (attachments) {
                let att = attachments.items[0];
                expect(att).to.not.be.undefined();
                expect(att.Name).to.be.equal('a1');
                expect(att.AttachmentId?.RootItemId).to.be.equal(pimTask.unid);
                expect(att.AttachmentId?.RootItemChangeKey).to.be.equal(`ck-${pimTask.unid}`);

                att = attachments.items[1];
                expect(att).to.not.be.undefined();
                expect(att.Name).to.be.equal('a2');
                expect(att.AttachmentId?.RootItemId).to.be.equal(pimTask.unid);
                expect(att.AttachmentId?.RootItemChangeKey).to.be.equal(`ck-${pimTask.unid}`);
            }

        });

        it("UnindexedFieldURIType.ITEM_SUBJECT", function () {
            // No subject
            let handled = manager.updateEWSItemFieldValue(note, pimNote, UnindexedFieldURIType.ITEM_SUBJECT);
            expect(handled).to.be.true();
            expect(note.Subject).to.be.equal('');

            // With a subject
            pimNote.subject = "SUBJECT";
            handled = manager.updateEWSItemFieldValue(note, pimNote, UnindexedFieldURIType.ITEM_SUBJECT);
            expect(handled).to.be.true();
            expect(note.Subject).to.be.equal(pimNote.subject);
        });

        it("UnindexedFieldURIType.ITEM_CATEGORIES", function () {
            let handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_CATEGORIES);

            expect(handled).to.be.true();
            expect(task.Categories).to.be.undefined();

            pimTask.categories = ['abc', 'def'];
            handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_CATEGORIES);

            expect(handled).to.be.true();
            expect(task.Categories).to.not.be.undefined();
            expect(task.Categories?.String).to.be.eql(['abc', 'def']);

        });

        it("UnindexedFieldURIType.ITEM_IMPORTANCE", function () {
            // Task
            // No proiority set
            let handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_IMPORTANCE);
            expect(handled).to.be.true();
            expect(task.Importance).to.be.equal(ImportanceChoicesType.NORMAL);

            // High priority for task
            pimTask.priority = PimTaskPriority.HIGH;
            handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_IMPORTANCE);
            expect(handled).to.be.true();
            expect(task.Importance).to.be.equal(ImportanceChoicesType.HIGH);

            // Normal priority for task
            pimTask.priority = PimTaskPriority.MEDIUM;
            handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_IMPORTANCE);
            expect(handled).to.be.true();
            expect(task.Importance).to.be.equal(ImportanceChoicesType.NORMAL);

            // High priority for task
            pimTask.priority = PimTaskPriority.LOW;
            handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_IMPORTANCE);
            expect(handled).to.be.true();
            expect(task.Importance).to.be.equal(ImportanceChoicesType.LOW);

            // Message
            const pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';
            const message = new MessageType();

            // No importance set
            handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.ITEM_IMPORTANCE);
            expect(handled).to.be.false();
            expect(message.Importance).to.be.undefined();

            // Importance for messages is handled in the base class

            // High importance
            pimMessage.importance = PimImportance.HIGH;
            handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.ITEM_IMPORTANCE);
            expect(handled).to.be.false();
            expect(message.Importance).to.be.undefined();

            // Calendar item
            // No importance set
            handled = manager.updateEWSItemFieldValue(calendarItem, pimCalendarItem, UnindexedFieldURIType.ITEM_IMPORTANCE);
            expect(handled).to.be.true();
            expect(calendarItem.Importance).to.be.undefined();

            // Low importance maps to undefined
            pimCalendarItem.importance = PimImportance.LOW;
            handled = manager.updateEWSItemFieldValue(calendarItem, pimCalendarItem, UnindexedFieldURIType.ITEM_IMPORTANCE);
            expect(handled).to.be.true();
            expect(calendarItem.Importance).to.be.undefined();

            // High importance
            pimCalendarItem.importance = PimImportance.HIGH;
            handled = manager.updateEWSItemFieldValue(calendarItem, pimCalendarItem, UnindexedFieldURIType.ITEM_IMPORTANCE);
            expect(handled).to.be.true();
            expect(calendarItem.Importance).to.be.equal(ImportanceChoicesType.HIGH);

            // Contact -- not supported
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.ITEM_IMPORTANCE);
            expect(handled).to.be.false();
            expect(contact.Importance).to.be.undefined();
        });

        it("UnindexedFieldURIType.ITEM_IN_REPLY_TO", function () {
            // TODO:  Not currently supported.  Implement when it is.
            const handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_IN_REPLY_TO);

            expect(handled).to.be.true();
        });

        it("UnindexedFieldURIType.ITEM_IS_DRAFT", function () {
            let handled = manager.updateEWSItemFieldValue(calendarItem, pimCalendarItem, UnindexedFieldURIType.ITEM_IS_DRAFT);
            expect(handled).to.be.true();
            expect(calendarItem.IsDraft).to.be.undefined();

            pimCalendarItem.isDraft = true;
            handled = manager.updateEWSItemFieldValue(calendarItem, pimCalendarItem, UnindexedFieldURIType.ITEM_IS_DRAFT);
            expect(handled).to.be.true();
            expect(calendarItem.IsDraft).to.be.true();

            // Try an unsupported different type
            handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_IS_DRAFT);
            expect(handled).to.be.false();
            expect(task.IsDraft).to.be.undefined();
        });

        it("UnindexedFieldURIType.ITEM_BODY", function () {
            let handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_BODY);
            expect(handled).to.be.true();
            expect(task.Body).to.be.undefined();

            pimTask.body = "BODY";
            handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_BODY);
            expect(handled).to.be.true();
            expect(task.Body?.Value).to.be.equal(pimTask.body);
            expect(task.Body?.BodyType).to.be.equal('Text');
        });

        it("UnindexedFieldURIType.ITEM_UNIQUE_BODY, ITEM_NORMALIZED_BODY, ITEM_TEXT_BODY", function () {
            pimTask.body = "BODY";

            let handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_UNIQUE_BODY);
            expect(handled).to.be.true();
            expect(task.Body).to.be.undefined();

            handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_NORMALIZED_BODY);
            expect(handled).to.be.true();
            expect(task.Body).to.be.undefined();

            handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_TEXT_BODY);
            expect(handled).to.be.true();
            expect(task.Body).to.be.undefined();
        });

        it("UnindexedFieldURIType.ITEM_SENSITIVITY", function () {
            // Do not explicitly set..isPrivate defaults to true
            let handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_SENSITIVITY);
            expect(handled).to.be.true();
            expect(task.Sensitivity).to.be.equal(SensitivityChoicesType.NORMAL);

            // Set isPrivate to true
            pimTask.isPrivate = true;
            handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_SENSITIVITY);
            expect(handled).to.be.true();
            expect(task.Sensitivity).to.be.equal(SensitivityChoicesType.CONFIDENTIAL);

            // Set isPrivate to false
            pimTask.isPrivate = false;
            handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_SENSITIVITY);
            expect(handled).to.be.true();
            expect(task.Sensitivity).to.be.equal(SensitivityChoicesType.NORMAL);

            // Try confidential
            pimTask.isConfidential = true;
            handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_SENSITIVITY);
            expect(handled).to.be.true();
            expect(task.Sensitivity).to.be.equal(SensitivityChoicesType.CONFIDENTIAL);

            // Set isConfidential to false
            pimTask.isConfidential = false;
            handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_SENSITIVITY);
            expect(handled).to.be.true();
            expect(task.Sensitivity).to.be.equal(SensitivityChoicesType.NORMAL);
        });

        it("UnindexedFieldURIType.ITEM_REMINDER_DUE_BY, ITEM_REMINDER_NEXT_TIME", function () {
            const dueDate = new Date();

            // Test with a task, using ITEM_REMINDER_DUE_BY
            let handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_REMINDER_DUE_BY);
            expect(handled).to.be.true();
            expect(task.ReminderIsSet).to.be.false();
            expect(task.ReminderDueBy).to.be.undefined();

            pimTask.alarm = 15;
            handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_REMINDER_DUE_BY);
            expect(handled).to.be.true();
            expect(task.ReminderIsSet).to.be.false();
            expect(task.ReminderDueBy).to.be.undefined();

            pimTask.due = dueDate.toISOString();
            handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_REMINDER_DUE_BY);
            expect(handled).to.be.true();
            expect(task.ReminderIsSet).to.be.true();
            expect(task.ReminderDueBy ? getTrimmedISODate(task.ReminderDueBy) : 'Fail: ReminderDueBy is not set').to.be.eql(getTrimmedISODate(dueDate));

            // Test with a task using 
            pimTask = PimItemFactory.newPimTask({});
            task = new TaskType();
            handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_REMINDER_NEXT_TIME);
            expect(handled).to.be.true();
            expect(task.ReminderIsSet).to.be.false();
            expect(task.ReminderDueBy).to.be.undefined();

            pimTask.alarm = 15;
            handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_REMINDER_NEXT_TIME);
            expect(handled).to.be.true();
            expect(task.ReminderIsSet).to.be.false();
            expect(task.ReminderDueBy).to.be.undefined();

            pimTask.due = dueDate.toISOString();
            handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_REMINDER_NEXT_TIME);
            expect(handled).to.be.true();
            expect(task.ReminderIsSet).to.be.true();
            expect(task.ReminderDueBy ? getTrimmedISODate(task.ReminderDueBy) : 'Fail: ReminderDueBy is not set').to.be.eql(getTrimmedISODate(dueDate));

            // Test with a calendarItem, using ITEM_REMINDER_DUE_BY
            handled = manager.updateEWSItemFieldValue(calendarItem, pimCalendarItem, UnindexedFieldURIType.ITEM_REMINDER_DUE_BY);
            expect(handled).to.be.true();
            expect(calendarItem.ReminderIsSet).to.be.false();
            expect(calendarItem.ReminderMinutesBeforeStart).to.be.undefined();
            expect(calendarItem.ReminderDueBy).to.be.undefined();

            pimCalendarItem.alarm = -15;
            handled = manager.updateEWSItemFieldValue(calendarItem, pimCalendarItem, UnindexedFieldURIType.ITEM_REMINDER_DUE_BY);
            expect(handled).to.be.true();
            expect(calendarItem.ReminderIsSet).to.be.true();
            expect(calendarItem.ReminderMinutesBeforeStart).to.be.equal('15');
            expect(calendarItem.ReminderDueBy).to.be.undefined();

            pimCalendarItem.start = dueDate.toISOString();
            handled = manager.updateEWSItemFieldValue(calendarItem, pimCalendarItem, UnindexedFieldURIType.ITEM_REMINDER_DUE_BY);
            expect(handled).to.be.true();
            expect(calendarItem.ReminderIsSet).to.be.true();
            expect(calendarItem.ReminderMinutesBeforeStart).to.be.equal('15');
            expect(calendarItem.ReminderDueBy).to.be.eql(dueDate);

            // Test with a calendarItem, using ITEM_REMINDER_NEXT_TIME
            calendarItem = new CalendarItemType();
            pimCalendarItem = PimItemFactory.newPimCalendarItem({}, 'default', PimItemFormat.DOCUMENT);
            pimCalendarItem.unid = 'UNID';
            handled = manager.updateEWSItemFieldValue(calendarItem, pimCalendarItem, UnindexedFieldURIType.ITEM_REMINDER_NEXT_TIME);
            expect(handled).to.be.true();
            expect(calendarItem.ReminderIsSet).to.be.false();
            expect(calendarItem.ReminderMinutesBeforeStart).to.be.undefined();
            expect(calendarItem.ReminderDueBy).to.be.undefined();

            pimCalendarItem.alarm = -15;
            handled = manager.updateEWSItemFieldValue(calendarItem, pimCalendarItem, UnindexedFieldURIType.ITEM_REMINDER_NEXT_TIME);
            expect(handled).to.be.true();
            expect(calendarItem.ReminderIsSet).to.be.true();
            expect(calendarItem.ReminderMinutesBeforeStart).to.be.equal('15');
            expect(calendarItem.ReminderDueBy).to.be.undefined();

            pimCalendarItem.start = dueDate.toISOString();
            handled = manager.updateEWSItemFieldValue(calendarItem, pimCalendarItem, UnindexedFieldURIType.ITEM_REMINDER_NEXT_TIME);
            expect(handled).to.be.true();
            expect(calendarItem.ReminderIsSet).to.be.true();
            expect(calendarItem.ReminderMinutesBeforeStart).to.be.equal('15');
            expect(calendarItem.ReminderDueBy).to.be.eql(dueDate);

            // Try different type
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.ITEM_REMINDER_NEXT_TIME);
            expect(handled).to.be.false();

            // Try some strange combinations for code coverage :-)
            handled = manager.updateEWSItemFieldValue(contact, pimTask, UnindexedFieldURIType.ITEM_REMINDER_NEXT_TIME);
            expect(handled).to.be.false();

            handled = manager.updateEWSItemFieldValue(task, pimContact, UnindexedFieldURIType.ITEM_REMINDER_NEXT_TIME);
            expect(handled).to.be.false();

        });

        it("UnindexedFieldURIType.ITEM_REMINDER_IS_SET", function () {
            // No reminder set for task
            let handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_REMINDER_IS_SET);
            expect(handled).to.be.true();
            expect(task.ReminderIsSet).to.be.false();

            // Reminder set for task
            pimTask.alarm = 15;
            handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_REMINDER_IS_SET);
            expect(handled).to.be.true();
            expect(task.ReminderIsSet).to.be.true();

            // No reminder for calendar item
            handled = manager.updateEWSItemFieldValue(calendarItem, pimCalendarItem, UnindexedFieldURIType.ITEM_REMINDER_IS_SET);
            expect(handled).to.be.true();
            expect(calendarItem.ReminderIsSet).to.be.false();

            // Reminder set for calendarItem
            pimCalendarItem.alarm = 15;
            handled = manager.updateEWSItemFieldValue(calendarItem, pimCalendarItem, UnindexedFieldURIType.ITEM_REMINDER_IS_SET);
            expect(handled).to.be.true();
            expect(calendarItem.ReminderIsSet).to.be.true();

            // Try a different type
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.ITEM_REMINDER_IS_SET);
            expect(handled).to.be.false();
            expect(contact.ReminderIsSet).to.be.undefined();
        });

        it("UnindexedFieldURIType.ITEM_REMINDER_MINUTES_BEFORE_START", function () {
            let handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_REMINDER_MINUTES_BEFORE_START);
            expect(handled).to.be.true();
            expect(task.ReminderMinutesBeforeStart).to.be.undefined();

            pimTask.alarm = -15;
            handled = manager.updateEWSItemFieldValue(task, pimTask, UnindexedFieldURIType.ITEM_REMINDER_MINUTES_BEFORE_START);
            expect(handled).to.be.true();
            expect(task.ReminderMinutesBeforeStart).to.be.equal('15');

            handled = manager.updateEWSItemFieldValue(calendarItem, pimCalendarItem, UnindexedFieldURIType.ITEM_REMINDER_MINUTES_BEFORE_START);
            expect(handled).to.be.true();
            expect(calendarItem.ReminderMinutesBeforeStart).to.be.undefined();

            pimCalendarItem.alarm = -15;
            handled = manager.updateEWSItemFieldValue(calendarItem, pimCalendarItem, UnindexedFieldURIType.ITEM_REMINDER_MINUTES_BEFORE_START);
            expect(handled).to.be.true();
            expect(calendarItem.ReminderMinutesBeforeStart).to.be.equal('15');

            // Try a different type
            handled = manager.updateEWSItemFieldValue(contact, pimContact, UnindexedFieldURIType.ITEM_REMINDER_MINUTES_BEFORE_START);
            expect(handled).to.be.false();
            expect(contact.ReminderMinutesBeforeStart).to.be.undefined();
        });

        it("UnindexedFieldURIType.ITEM_LAST_MODIFIED_TIME", function () {
            // No last modified
            let handled = manager.updateEWSItemFieldValue(note, pimNote, UnindexedFieldURIType.ITEM_LAST_MODIFIED_TIME);
            expect(handled).to.be.true();
            expect(note.LastModifiedTime).to.be.undefined();

            // With a last modified
            pimNote.lastModifiedDate = new Date();
            handled = manager.updateEWSItemFieldValue(note, pimNote, UnindexedFieldURIType.ITEM_LAST_MODIFIED_TIME);
            expect(handled).to.be.true();
            expect(note.LastModifiedTime).to.be.eql(pimNote.lastModifiedDate);
        });

        it("UnindexedFieldURIType.ITEM_EFFECTIVE_RIGHTS", function () {
            // No subject
            const handled = manager.updateEWSItemFieldValue(note, pimNote, UnindexedFieldURIType.ITEM_EFFECTIVE_RIGHTS);
            expect(handled).to.be.true();
            const rights = note.EffectiveRights;
            if (rights) {
                expect(rights.CreateAssociated).to.be.true();
                expect(rights.CreateContents).to.be.true();
                expect(rights.Delete).to.be.true();
                expect(rights.Modify).to.be.true();
                expect(rights.Read).to.be.true();
                expect(rights.ViewPrivateItems).to.be.true();
                expect(rights.CreateHierarchy).to.be.true();
            }
        });
    });    
});
import { expect, sinon } from '@loopback/testlab';
import {
    PimItem, PimLabel, UserInfo, PimItemFormat, PimItemFactory, PimNote, KeepPimNotebookManager, PimLabelTypes, 
    KeepPimLabelManager, KeepPimConstants, KeepPimMessageManager, KeepPimBaseResults
} from '@hcllabs/openclientkeepcomponent';
import {
    UnindexedFieldURIType, MessageDispositionType, BodyTypeType, DefaultShapeNamesType,
    DistinguishedPropertySetType, DictionaryURIType, MapiPropertyTypeType, ItemClassType
} from '../../models/enum.model';
import {
    ItemResponseShapeType, ItemType, ItemChangeType, ItemIdType, BodyType, NonEmptyArrayOfItemChangeDescriptionsType,
    DeleteItemFieldType, SetItemFieldType, ExtendedPropertyType, AppendToItemFieldType, MessageType, FolderIdType
} from '../../models/mail.model';
import {
    ArrayOfStringsType, PathToExtendedFieldType, PathToIndexedFieldType, PathToUnindexedFieldType
} from '../../models/common.model';
import { EWSNotesManager, getEWSId, getKeepIdPair } from '../../utils';
import { UserContext } from '../../keepcomponent';
import {
    generateTestNotesLabel,
    getContext, stubLabelsManager, stubMessageManager, stubNotebookManager
} from '../unitTestHelper';
import { Request } from '@loopback/rest';

describe('EWSNotesManager tests', () => {

    const testUser = "test.user@test.org";

    // Mock subclass of EWS notes manager for testing protected functions
    class MockEWSManager extends EWSNotesManager {

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

        async deleteItem(item: string | PimNote, userInfo: UserInfo, mailboxId?: string, hardDelete = false): Promise<void> {
            return super.deleteItem(item, userInfo, mailboxId, hardDelete);
        }

        async updateItem(
            pimItem: PimNote, 
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

        addRequestedPropertiesToEWSItem(
            pimItem: PimItem, 
            toItem: ItemType, 
            shape?: ItemResponseShapeType, 
            mailboxId?: string
        ): void {
            super.addRequestedPropertiesToEWSItem(pimItem, toItem, shape, mailboxId);
        }

        updateEWSItemFieldValue(item: ItemType, pimNote: PimNote, fieldId: UnindexedFieldURIType, mailboxId?: string): boolean {
            return super.updateEWSItemFieldValue(item, pimNote, fieldId, mailboxId);
        }

        pimItemFromEWSItem(item: MessageType, request: Request, existing?: any): PimNote {
            return super.pimItemFromEWSItem(item, request, existing);
        }

        pimItemToEWSItem(
            pimItem: PimNote, 
            userInfo: UserInfo, 
            request: Request, 
            shape: ItemResponseShapeType, 
            mailboxId?: string,
            parentFolderId?: string
        ): Promise<MessageType> {
            return super.pimItemToEWSItem(pimItem, userInfo, request, shape, mailboxId, parentFolderId);
        }
    }

    describe('Test getInstance', () => {
        it('getInstance', function () {
            const manager = EWSNotesManager.getInstance();
            expect(manager).to.be.instanceof(EWSNotesManager);

            const manager2 = EWSNotesManager.getInstance();
            expect(manager).to.be.equal(manager2);
        });
    });

    describe('Test pimItemFromEWSItem', () => {
        it('empty MessageType', () => {
            const manager = new MockEWSManager();
            const note = new MessageType();
            const context = getContext(testUser, 'password');
            const pimNote = manager.pimItemFromEWSItem(note, context.request, {});
            expect(pimNote.unid).to.be.undefined();
            expect(pimNote.createdDate).to.be.undefined();
        });

        it('common fields populated', () => {
            const created = new Date();
            const modified = new Date();
            const manager = new MockEWSManager();
            const context = getContext(testUser, 'password');
            const note = new MessageType();
            const mailboxId = 'test@test.com';
            const uid = 'uid';
            const itemEWSId = getEWSId(uid, mailboxId);
            note.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            const parentFolderEWSId = getEWSId('parent', mailboxId);
            note.ParentFolderId = new FolderIdType(parentFolderEWSId, `ck-${parentFolderEWSId}`);
            note.DateTimeCreated = created;
            note.Categories = new ArrayOfStringsType();
            note.Categories.String = ['cat1', 'cat2'];
            note.ItemClass = ItemClassType.ITEM_CLASS_NOTE;
            note.Body = new BodyType('BODY', BodyTypeType.TEXT);
            note.LastModifiedTime = modified;
            note.Subject = 'SUBJECT';
            note.IsRead = true;

            let pimNote = manager.pimItemFromEWSItem(note, context.request);
            expect(pimNote.unid).to.be.equal(uid);
            const parentIds = pimNote.parentFolderIds;
            expect(parentIds).to.be.an.Array();
            if (parentIds) {
                const [parentFolderId] = getKeepIdPair(note.ParentFolderId?.Id);
                expect(parentIds[0]).to.be.equal(parentFolderId);
            }
            expect(pimNote.diaryDate).to.be.eql(created);
            expect(pimNote.categories).to.be.an.Array();
            expect(pimNote.categories.length).to.be.equal(2);
            expect(pimNote.body).to.be.equal('BODY');
            expect(pimNote.lastModifiedDate).to.be.eql(modified);
            expect(pimNote.isRead).to.be.true();
            expect(pimNote.categories).to.be.eql(['cat1', 'cat2']);
            expect(pimNote.subject).to.be.equal(note.Subject);

            note.Categories = new ArrayOfStringsType();
            note.Categories.String = ['ABC'];
            pimNote = manager.pimItemFromEWSItem(note, context.request);
            expect(pimNote.categories).to.be.eql(['ABC']);
        });
    });

    describe('Test fieldsForDefaultShape', () => {
        it('test default fields', () => {
            const manager = EWSNotesManager.getInstance();
            const fields = manager.fieldsForDefaultShape();
            expect(fields).to.be.an.Array();
            expect(fields.length).to.be.equal(9);
            // Check a few of the fields
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_SUBJECT);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_DATE_TIME_CREATED);
        });

    });

    describe('Test fieldsForAllPropertiesShape', () => {
        it('test all properties fields', () => {
            const manager = EWSNotesManager.getInstance();
            const fields = manager.fieldsForAllPropertiesShape();
            expect(fields).to.be.an.Array();
            // Verify the size is correct
            expect(fields.length).to.be.equal(13);
            // Check a few of the fields
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_LAST_MODIFIED_TIME);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_IMPORTANCE);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_CATEGORIES);
        });
    });


    describe('Test getItems', () => {
        // Clean up the stubs
        afterEach(() => {
            sinon.restore(); // Reset stubs
        });

        it('Expected results', async () => {
            const parentLabel = generateTestNotesLabel();
            const pimNote = PimItemFactory.newPimNote({}, PimItemFormat.DOCUMENT);
            pimNote.unid = 'UNID';
            pimNote.parentFolderIds = [parentLabel.folderId];
            pimNote.body = 'BODY';
            pimNote.subject = 'SUBJECT';

            const notesManagerStub = sinon.createStubInstance(KeepPimNotebookManager);
            notesManagerStub.getNotes.resolves([pimNote]);
            stubNotebookManager(notesManagerStub);

            const manager = EWSNotesManager.getInstance();
            const userInfo = new UserContext();
            const context = getContext(testUser, 'password');
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            const mailboxId = 'test@test.com';
            const itemEWSId = getEWSId(pimNote.unid, mailboxId);

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
            expect(message.ItemId?.Id).to.be.equal(getEWSId(pimNote.unid));
        });

        it('getNotes throws an error', async () => {
            const parentLabel = generateTestNotesLabel();
            const pimNote = PimItemFactory.newPimNote({}, PimItemFormat.DOCUMENT);
            pimNote.unid = 'UNID';
            pimNote.parentFolderIds = [parentLabel.folderId];
            pimNote.body = 'BODY';
            pimNote.subject = 'SUBJECT';

            const notesManagerStub = sinon.createStubInstance(KeepPimNotebookManager);
            notesManagerStub.getNotes.throws();

            stubNotebookManager(notesManagerStub);

            const manager = EWSNotesManager.getInstance();
            const userInfo = new UserContext();
            const context = getContext(testUser, 'password');
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            const messages = await manager.getItems(userInfo, context.request, shape, 0, 10, parentLabel);
            expect(messages).to.be.an.Array();
            expect(messages.length).to.be.equal(0);
        });
    });


    describe('Test updateItem', () => {
        const context = getContext(testUser, 'password');
        const userInfo = new UserContext();
        const manager = new MockEWSManager();
        let pimNote: PimNote;
        let newNote: PimNote;
        let notesManagerStub;

        beforeEach(function () {
            pimNote = PimItemFactory.newPimNote({}, PimItemFormat.DOCUMENT);
            pimNote.subject = 'SUBJECT';
            pimNote.body = 'BODY';
            pimNote.unid = 'UNID';
            pimNote.parentFolderIds = ['PARENT'];

            newNote = PimItemFactory.newPimNote({}, PimItemFormat.DOCUMENT);
            newNote.unid = 'UNID';
            newNote.parentFolderIds = ['PARENT'];
            notesManagerStub = sinon.createStubInstance(KeepPimNotebookManager);
            const updateBaseResults: KeepPimBaseResults = {
                unid: 'UNID',
                status: 200,
                message: 'Success'
            };
            notesManagerStub.updateNote.resolves(updateBaseResults);
            stubNotebookManager(notesManagerStub);
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
            fieldURIChange.Item = new ItemType();
            fieldURIChange.Item.ItemId = new ItemIdType(pimNote.unid, `ck-${pimNote.unid}`);
            fieldURIChange.Item.Subject = 'UPDATED-SUBJECT';

            changes.ItemId = new ItemIdType(pimNote.unid, `ck-${pimNote.unid}`);
            changes.Updates.push(fieldURIChange);

            const extendedFieldChange = new SetItemFieldType();
            extendedFieldChange.ExtendedFieldURI = new PathToExtendedFieldType();
            extendedFieldChange.ExtendedFieldURI.DistinguishedPropertySetId = DistinguishedPropertySetType.ADDRESS;

            extendedFieldChange.Item = new ItemType();
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
            indextedFieldChange.Item = new ItemType();
            changes.Updates.push(indextedFieldChange);

            const mailboxId = 'test@test.com';
            const itemEWSId = getEWSId(pimNote.unid, mailboxId);
            const parentEWSId = getEWSId('PARENT-LABEL', mailboxId);

            let note = await manager.updateItem(pimNote, changes, userInfo, context.request, undefined, mailboxId);
            expect(note).to.not.be.undefined();
            expect(note?.ItemId?.Id).to.be.equal(itemEWSId);

            const label = PimItemFactory.newPimLabel({FolderId: 'PARENT-LABEL'});

            // Try with parent folderId
            note = await manager.updateItem(pimNote, changes, userInfo, context.request, label, mailboxId);
            expect(note).to.not.be.undefined();
            expect(note?.ParentFolderId?.Id).to.be.equal(parentEWSId);

            // Try with no existing parent folder
            pimNote.parentFolderIds = undefined;
            note = await manager.updateItem(pimNote, changes, userInfo, context.request, label, mailboxId);
            expect(note).to.not.be.undefined();
            expect(note?.ParentFolderId?.Id).to.be.equal(parentEWSId);

            // Try with no parent folder at all
            pimNote.parentFolderIds = undefined;
            note = await manager.updateItem(pimNote, changes, userInfo, context.request);
            expect(note).to.not.be.undefined();
            expect(note?.ParentFolderId?.Id).to.be.undefined();
        });

        it('test append field update', async () => {
            let changes = new ItemChangeType();
            changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

            // Try body
            let fieldURIChangeBody = new AppendToItemFieldType()
            fieldURIChangeBody.FieldURI = new PathToUnindexedFieldType();
            fieldURIChangeBody.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_BODY;
            fieldURIChangeBody.Item = new ItemType();
            fieldURIChangeBody.Item.ItemId = new ItemIdType(pimNote.unid, `ck-${pimNote.unid}`);
            fieldURIChangeBody.Item.Body = new BodyType('APPENDED-BODY', BodyTypeType.TEXT)

            changes.ItemId = new ItemIdType(pimNote.unid, `ck-${pimNote.unid}`);
            changes.Updates.push(fieldURIChangeBody);

            // Try subject (not supported)
            const fieldURIChangeSubject = new AppendToItemFieldType()
            fieldURIChangeSubject.FieldURI = new PathToUnindexedFieldType();
            fieldURIChangeSubject.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            fieldURIChangeSubject.Item = new ItemType();
            fieldURIChangeSubject.Item.ItemId = new ItemIdType(pimNote.unid, `ck-${pimNote.unid}`);
            fieldURIChangeSubject.Item.Subject = 'APPENDED-SUBJECT';

            changes.ItemId = new ItemIdType(pimNote.unid, `ck-${pimNote.unid}`);
            changes.Updates.push(fieldURIChangeSubject);

            // No extended fields are supported
            const extendedFieldChange = new AppendToItemFieldType();
            extendedFieldChange.ExtendedFieldURI = new PathToExtendedFieldType();
            extendedFieldChange.ExtendedFieldURI.DistinguishedPropertySetId = DistinguishedPropertySetType.ADDRESS;
            extendedFieldChange.ExtendedFieldURI.PropertyId = 32899;
            extendedFieldChange.ExtendedFieldURI.PropertyType = MapiPropertyTypeType.STRING;

            extendedFieldChange.Item = new ItemType();
            extendedFieldChange.Item.ExtendedProperty = [];
            const prop = new ExtendedPropertyType();
            prop.ExtendedFieldURI = new PathToExtendedFieldType();
            prop.ExtendedFieldURI.DistinguishedPropertySetId = DistinguishedPropertySetType.ADDRESS;
            prop.ExtendedFieldURI.PropertyId = 32899;
            prop.ExtendedFieldURI.PropertyType = MapiPropertyTypeType.STRING;
            prop.Value = 'append';
            extendedFieldChange.Item.ExtendedProperty.push(prop);
            changes.Updates.push(extendedFieldChange);

            let note = await manager.updateItem(pimNote, changes, userInfo, context.request);
            expect(note).to.not.be.undefined();
            expect(note?.ItemId?.Id).to.be.equal(getEWSId(pimNote.unid));

            // Try with no body
            changes = new ItemChangeType();
            changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

            fieldURIChangeBody = new AppendToItemFieldType()
            fieldURIChangeBody.FieldURI = new PathToUnindexedFieldType();
            fieldURIChangeBody.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_BODY;
            fieldURIChangeBody.Item = new ItemType();
            fieldURIChangeBody.Item.ItemId = new ItemIdType(pimNote.unid, `ck-${pimNote.unid}`);

            changes.ItemId = new ItemIdType(pimNote.unid, `ck-${pimNote.unid}`);
            changes.Updates.push(fieldURIChangeBody);

            note = await manager.updateItem(pimNote, changes, userInfo, context.request);
            expect(note).to.not.be.undefined();
            expect(note?.ItemId?.Id).to.be.equal(getEWSId(pimNote.unid));
        });


        it('test delete field update', async () => {
            const changes = new ItemChangeType();
            changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();
            const fieldURIChange = new DeleteItemFieldType();
            fieldURIChange.FieldURI = new PathToUnindexedFieldType();
            fieldURIChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            changes.ItemId = new ItemIdType(pimNote.unid, `ck-${pimNote.unid}`);
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

            const note = await manager.updateItem(pimNote, changes, userInfo, context.request);
            expect(note).to.not.be.undefined();
            expect(note?.ItemId?.Id).to.be.equal(getEWSId(pimNote.unid));
        });


        it('test different item types', async () => {
            const changes = new ItemChangeType();
            changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

            const fieldURIChange = new SetItemFieldType()
            fieldURIChange.FieldURI = new PathToUnindexedFieldType();
            fieldURIChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            fieldURIChange.Message = new MessageType();
            fieldURIChange.Message.ItemId = new ItemIdType(pimNote.unid, `ck-${pimNote.unid}`);
            fieldURIChange.Message.Subject = 'UPDATED-SUBJECT';

            changes.ItemId = new ItemIdType(pimNote.unid, `ck-${pimNote.unid}`);
            changes.Updates.push(fieldURIChange);

            pimNote.parentFolderIds = undefined;

            // Stub label manager to return a journal folder
            const notesLabel = PimItemFactory.newPimLabel({}, PimItemFormat.DOCUMENT);
            notesLabel.displayName = 'Contacts';
            notesLabel.view = KeepPimConstants.JOURNAL;
            notesLabel.type = PimLabelTypes.JOURNAL;

            const labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
            labelManagerStub.getLabels.resolves([notesLabel]);
            stubLabelsManager(labelManagerStub);

            const note = await manager.updateItem(pimNote, changes, userInfo, context.request);
            expect(note).to.not.be.undefined();
            expect(note?.ItemId?.Id).to.be.equal(getEWSId(pimNote.unid));
        });

        it('test no journal labels', async () => {
            const changes = new ItemChangeType();
            changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

            const fieldURIChange = new SetItemFieldType()
            fieldURIChange.FieldURI = new PathToUnindexedFieldType();
            fieldURIChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            fieldURIChange.Message = new MessageType();
            fieldURIChange.Message.ItemId = new ItemIdType(pimNote.unid, `ck-${pimNote.unid}`);
            fieldURIChange.Message.Subject = 'UPDATED-SUBJECT';

            changes.ItemId = new ItemIdType(pimNote.unid, `ck-${pimNote.unid}`);
            changes.Updates.push(fieldURIChange);

            pimNote.parentFolderIds = undefined;

            // Stub label manager to not return a journal folder
            const notesLabel = PimItemFactory.newPimLabel({}, PimItemFormat.DOCUMENT);
            notesLabel.displayName = 'Test';
            notesLabel.view = KeepPimConstants.CONTACTS;
            notesLabel.type = PimLabelTypes.CONTACTS;

            const labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
            labelManagerStub.getLabels.resolves([notesLabel]);
            stubLabelsManager(labelManagerStub);

            const note = await manager.updateItem(pimNote, changes, userInfo, context.request);
            expect(note).to.not.be.undefined();
            expect(note?.ItemId?.Id).to.be.equal(getEWSId(pimNote.unid));
        });
    });

    describe('Test createItem', () => {
        // Clean up the stubs
        afterEach(() => {
            sinon.restore(); // Reset stubs
        });

        it('test create note, no label', async () => {
            const managerStub = sinon.createStubInstance(KeepPimNotebookManager);
            managerStub.createNote.resolves('UID');
            managerStub.deleteNote.resolves({});
            stubNotebookManager(managerStub);

            const manager = new MockEWSManager();
            const userInfo = new UserContext();

            const note = new MessageType();
            note.ItemClass = ItemClassType.ITEM_CLASS_NOTE;
            note.ItemId = new ItemIdType('UID', 'ck-UID');

            const context = getContext(testUser, 'password');

            // No parent folder
            const newMessages = await manager.createItem(note, userInfo, context.request);
            expect(newMessages).to.not.be.undefined();
            expect(newMessages.length).to.be.equal(1);
            expect(newMessages[0]?.ItemId?.Id).to.be.equal(getEWSId('UID'));
        });


        it('test create note, with a label', async () => {
            const messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            messageManagerStub.moveMessages.resolves({movedIds: [{status: 200, message: 'success', unid: 'UID'}]});
            stubMessageManager(messageManagerStub);

            const managerStub = sinon.createStubInstance(KeepPimNotebookManager);
            managerStub.createNote.resolves('UID');
            managerStub.deleteNote.resolves({});
            stubNotebookManager(managerStub);

            const manager = new MockEWSManager();
            const userInfo = new UserContext();
            const toLabel = PimItemFactory.newPimLabel({}, PimItemFormat.DOCUMENT);
            toLabel.folderId = 'TO-LABEL';
            toLabel.type = PimLabelTypes.JOURNAL;

            const note = new MessageType();
            note.ItemClass = ItemClassType.ITEM_CLASS_NOTE;
            note.ItemId = new ItemIdType('UID', 'ck-UID');
            const context = getContext(testUser, 'password');

            const mailboxId = 'test@test.com';
            const itemEWSId = getEWSId('UID', mailboxId);

            const newMessages = await manager
                .createItem(note, userInfo, context.request, MessageDispositionType.SAVE_ONLY, toLabel, mailboxId);
            expect(newMessages).to.not.be.undefined();
            expect(newMessages.length).to.be.equal(1);
            expect(newMessages[0]?.ItemId?.Id).to.be.equal(itemEWSId);
        });

        it('test create note, with error thrown', async () => {
            const messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            messageManagerStub.moveMessages.throws();
            stubMessageManager(messageManagerStub);

            const managerStub = sinon.createStubInstance(KeepPimNotebookManager);
            managerStub.createNote.resolves('UID');
            managerStub.deleteNote.resolves({});
            stubNotebookManager(managerStub);

            const manager = new MockEWSManager();
            const userInfo = new UserContext();
            const toLabel = PimItemFactory.newPimLabel({}, PimItemFormat.DOCUMENT);
            toLabel.folderId = 'FOLDER-ID';
            toLabel.type = PimLabelTypes.JOURNAL;

            const note = new MessageType();
            note.ItemClass = ItemClassType.ITEM_CLASS_NOTE;
            note.ItemId = new ItemIdType('UID', 'ck-UID');
            const context = getContext(testUser, 'password');

            let success = false;
            try {
                await manager.createItem(note, userInfo, context.request, MessageDispositionType.SAVE_ONLY, toLabel);
            } catch (err) {
                success = true;
            }
            expect(success).to.be.true();
        });
    });

    describe('Test deleteItem', () => {

        it('test delete', async () => {
            const managerStub = sinon.createStubInstance(KeepPimNotebookManager);
            managerStub.deleteNote.resolves({});
            stubNotebookManager(managerStub);

            const manager = EWSNotesManager.getInstance();
            const userInfo = new UserContext();
            await manager.deleteItem('item-id)', userInfo);
        });
    });


    describe('Test pimItemToEWSItem', () => {
        // Clean up the stubs
        afterEach(() => {
            sinon.restore(); // Reset stubs
        });

        it('simple note, no parent folder', async () => {
            const pimItem = PimItemFactory.newPimNote({}, PimItemFormat.DOCUMENT);
            pimItem.unid = 'UNID';
            pimItem.subject = 'SUBJECT';
            pimItem.body = 'BODY';

            const manager = new MockEWSManager();
            const userInfo = new UserContext();
            const context = getContext(testUser, 'password');
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.DEFAULT;
            shape.IncludeMimeContent = true;

            const note = await manager.pimItemToEWSItem(pimItem, userInfo, context.request, shape);
            expect(note.ItemId?.Id).to.be.equal(getEWSId('UNID'));
            expect(note.Subject).to.be.equal(pimItem.subject);
        });

        it('simple note, with parent folder', async () => {
            const pimItem = PimItemFactory.newPimNote({}, PimItemFormat.DOCUMENT);
            pimItem.unid = 'UNID';
            pimItem.subject = 'SUBJECT';
            pimItem.body = 'BODY';

            const manager = new MockEWSManager();
            const userInfo = new UserContext();
            const context = getContext(testUser, 'password');
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            shape.IncludeMimeContent = true;
            const mailboxId = 'test@test.com';
            const itemEWSId = getEWSId('UNID', mailboxId);
            const parentEWSId = getEWSId('PARENT-FOLDER-ID', mailboxId);

            const note = await manager.pimItemToEWSItem(pimItem, userInfo, context.request, shape, mailboxId, 'PARENT-FOLDER-ID');
            expect(note.ItemId?.Id).to.be.equal(itemEWSId);
            expect(note.ParentFolderId?.Id).to.be.equal(parentEWSId);
        });
    });
});
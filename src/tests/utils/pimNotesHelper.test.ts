import { expect } from '@loopback/testlab';
import { EWSNotesManager, getEWSId } from '../../utils';
import { ItemResponseShapeType, NonEmptyArrayOfPathsToElementType, MessageType, BodyType, ItemIdType, FolderIdType } from '../../models/mail.model';
import { DefaultShapeNamesType, UnindexedFieldURIType, ItemClassType, SensitivityChoicesType, MapiPropertyTypeType, BodyTypeType } from '../../models/enum.model';
import { PimItemFormat, PimItemFactory, base64Encode } from '@hcllabs/openclientkeepcomponent';
import { PathToUnindexedFieldType, PathToExtendedFieldType, ArrayOfStringsType } from '../../models/common.model';
import { getContext } from '../unitTestHelper';
import { UserContext } from '../../keepcomponent';


class MockNotesManager extends EWSNotesManager {
    getFieldsForShape(shape: ItemResponseShapeType): string[] {
        return this.defaultPropertiesForShape(shape);
    }
}

describe('pimNotesHelper tests', () => {
    const context = getContext("testUser","testPass")
    const userInfo = new UserContext();

    describe('pimNotesHelper.pimNoteToItem', () => {

        const noteObject = {
            Body: base64Encode('Note body'),
            uid: 'NOTE-ID',
            ParentFolder: 'PARENT-FOLDER-ID',
            Subject: 'Note subject',
            size: 10,
            TimeCreated: new Date('March 2, 1996'),
            '@created': new Date('March 2, 1996'),
            '@lastmodified': new Date(),
            Categories: [
                "Red",
                "Blue"
            ],
            "AdditionalFields": {
                'xHCL-extProp_0': {
                    DistinguishedPropertySetId: 'Address',
                    PropertyId: 32899,
                    PropertyType: 'String',
                    Value: 'fozzy.bear@fakemail.com'
                },
                'xHCL-extProp_1': {
                    DistinguishedPropertySetId: 'Address',
                    PropertyId: 32896,
                    PropertyType: 'Boolean',
                    Value: true
                }
            }
        }
        const parentFolderEWSId = getEWSId('PARENT-FOLDER-ID');
        const parentFolderCK = `ck-${parentFolderEWSId}`;
        it('test ID_ONLY shape', async function () {
            const pimNote = PimItemFactory.newPimNote(noteObject, PimItemFormat.DOCUMENT);
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;

            const ewsNote = await EWSNotesManager.getInstance().pimItemToEWSItem(pimNote, userInfo, context.request, shape);
            if (ewsNote !== undefined) {
                const itemEWSId = getEWSId('NOTE-ID');
                expect(ewsNote?.ItemId?.Id).to.be.equal(itemEWSId);
                expect(ewsNote?.ItemId?.ChangeKey).to.be.equal(`ck-${itemEWSId}`);
                expect(ewsNote?.ParentFolderId?.Id).to.be.equal(parentFolderEWSId);
                expect(ewsNote?.ParentFolderId?.ChangeKey).to.be.equal(parentFolderCK);
                expect(ewsNote?.ItemClass).to.be.equal(ItemClassType.ITEM_CLASS_NOTE);
                expect(ewsNote?.Body).to.be.undefined();
                expect(ewsNote?.Categories).to.be.undefined();
                expect(ewsNote?.Size).to.be.undefined();
            }
        });

        it('test DEFAULT shape', async function () {
            const pimNote = PimItemFactory.newPimNote(noteObject, PimItemFormat.DOCUMENT);
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.DEFAULT;

            const ewsNote = await EWSNotesManager.getInstance().pimItemToEWSItem(pimNote, userInfo, context.request, shape);

            const itemEWSId = getEWSId('NOTE-ID');
            expect(ewsNote?.ItemId?.Id).to.be.equal(itemEWSId);
            expect(ewsNote?.ItemId?.ChangeKey).to.be.equal(`ck-${itemEWSId}`);
            expect(ewsNote?.ParentFolderId?.Id).to.be.equal(parentFolderEWSId);
            expect(ewsNote?.ParentFolderId?.ChangeKey).to.be.equal(parentFolderCK);
            expect(ewsNote?.ItemClass).to.be.equal(ItemClassType.ITEM_CLASS_NOTE);

            expect(ewsNote?.Body?.Value).to.be.equal('Note body');
            expect(ewsNote?.Subject).to.be.equal('Note subject');
            expect(ewsNote?.DateTimeCreated).to.be.eql(noteObject.TimeCreated);
            expect(ewsNote?.HasAttachments).to.be.false();
            expect(ewsNote?.Sensitivity).to.be.equal(SensitivityChoicesType.NORMAL);
            // expect(ewsNote.Size).to.be.equal(9); TODO:  Add size to PimNote

            expect(ewsNote?.LastModifiedTime).to.be.undefined();
            expect(ewsNote?.Categories).to.be.undefined();
            expect(ewsNote?.Size).to.be.undefined();
        });

        it('test ALL_PROPERTIES shape', async function () {
            const pimNote = PimItemFactory.newPimNote(noteObject, PimItemFormat.DOCUMENT);
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;

            const ewsNote = await EWSNotesManager.getInstance().pimItemToEWSItem(pimNote, userInfo, context.request, shape);
            const itemEWSId = getEWSId('NOTE-ID');
            expect(ewsNote?.ItemId?.Id).to.be.equal(itemEWSId);
            expect(ewsNote?.ItemId?.ChangeKey).to.be.equal(`ck-${itemEWSId}`);
            expect(ewsNote?.ParentFolderId?.Id).to.be.equal(parentFolderEWSId);
            expect(ewsNote?.ParentFolderId?.ChangeKey).to.be.equal(parentFolderCK);
            expect(ewsNote?.ItemClass).to.be.equal(ItemClassType.ITEM_CLASS_NOTE);

            expect(ewsNote?.Body?.Value).to.be.equal('Note body');
            expect(ewsNote?.Subject).to.be.equal('Note subject');
            expect(ewsNote?.DateTimeCreated).to.be.eql(noteObject.TimeCreated);
            expect(ewsNote?.HasAttachments).to.be.false();
            expect(ewsNote?.Sensitivity).to.be.equal(SensitivityChoicesType.NORMAL);
            // expect(ewsNote.Size).to.be.equal(9); TODO:  Add size to PimNote

            expect(ewsNote?.LastModifiedTime).to.be.eql(noteObject['@lastmodified']);
            expect(ewsNote?.Importance).to.be.undefined() // Not yet supported
            const catArray = ewsNote?.Categories?.String;
            expect(catArray?.length).to.be.equal(2);
            expect(catArray).to.containEql('Red');
            expect(catArray).to.containEql('Blue');
        });

        it('test ID_ONLY with Additional Properties', async function () {

            const path1 = new PathToUnindexedFieldType();
            path1.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            const path2 = new PathToUnindexedFieldType();
            path2.FieldURI = UnindexedFieldURIType.ITEM_LAST_MODIFIED_TIME;
            const path3 = new PathToUnindexedFieldType();
            path3.FieldURI = UnindexedFieldURIType.ITEM_BODY;
            const path4 = new PathToUnindexedFieldType();
            path4.FieldURI = UnindexedFieldURIType.ITEM_CATEGORIES;
            const path5 = new PathToUnindexedFieldType();
            path5.FieldURI = UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS;


            const pimNote = PimItemFactory.newPimNote(noteObject, PimItemFormat.DOCUMENT);
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;

            shape.AdditionalProperties = new NonEmptyArrayOfPathsToElementType([path1, path2, path3, path4, path5]);

            const ewsNote = await EWSNotesManager.getInstance().pimItemToEWSItem(pimNote, userInfo, context.request, shape);
            const itemEWSId = getEWSId('NOTE-ID');
            expect(ewsNote?.ItemId?.Id).to.be.equal(itemEWSId);
            expect(ewsNote?.ItemId?.ChangeKey).to.be.equal(`ck-${itemEWSId}`);
            expect(ewsNote?.ParentFolderId?.Id).to.be.equal(parentFolderEWSId);
            expect(ewsNote?.ParentFolderId?.ChangeKey).to.be.equal(parentFolderCK);
            expect(ewsNote?.ItemClass).to.be.equal(ItemClassType.ITEM_CLASS_NOTE);

            expect(ewsNote?.Body?.Value).to.be.equal('Note body');
            expect(ewsNote?.Subject).to.be.equal('Note subject');
            expect(ewsNote?.DateTimeCreated).to.be.undefined();
            expect(ewsNote?.HasAttachments).to.be.false();
            expect(ewsNote?.Sensitivity).to.be.undefined();

            expect(ewsNote?.LastModifiedTime).to.be.eql(noteObject['@lastmodified']);
            expect(ewsNote?.Importance).to.be.undefined() // Not yet supported
            const catArray = ewsNote?.Categories?.String;
            expect(catArray?.length).to.be.equal(2);
            expect(catArray).to.containEql('Red');
            expect(catArray).to.containEql('Blue');

        });

        it('test ID_ONLY with Extended Properties', async function () {

            const path1 = new PathToExtendedFieldType();
            path1.PropertyId = 32899;
            path1.PropertyType = MapiPropertyTypeType.STRING;
            const path2 = new PathToExtendedFieldType();
            path2.PropertyId = 32896;
            path2.PropertyType = MapiPropertyTypeType.BOOLEAN;

            const pimNote = PimItemFactory.newPimNote(noteObject, PimItemFormat.DOCUMENT);
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;

            shape.AdditionalProperties = new NonEmptyArrayOfPathsToElementType([path1, path2]);

            const ewsNote = await EWSNotesManager.getInstance().pimItemToEWSItem(pimNote, userInfo, context.request, shape);
            const itemEWSId = getEWSId('NOTE-ID');
            expect(ewsNote?.ItemId?.Id).to.be.equal(itemEWSId);
            expect(ewsNote?.ItemId?.ChangeKey).to.be.equal(`ck-${itemEWSId}`);
            expect(ewsNote?.ParentFolderId?.Id).to.be.equal(parentFolderEWSId);
            expect(ewsNote?.ParentFolderId?.ChangeKey).to.be.equal(parentFolderCK);
            expect(ewsNote?.ItemClass).to.be.equal(ItemClassType.ITEM_CLASS_NOTE);
            const extendedProps = ewsNote?.ExtendedProperty;
            expect(extendedProps?.length).to.be.equal(2);
        });

        it('test with parent folder passed in', async function () {
            const pimNote = PimItemFactory.newPimNote(noteObject, PimItemFormat.DOCUMENT);
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            const mailboxId = 'test@test.com';

            const ewsNote = await EWSNotesManager
                .getInstance()
                .pimItemToEWSItem(pimNote, userInfo, context.request, shape, mailboxId, 'MY-PARENT-FOLDER-ID');
            const itemEWSId = getEWSId('NOTE-ID', mailboxId);
            const passedParentFolderEWSId = getEWSId('MY-PARENT-FOLDER-ID', mailboxId);

            expect(ewsNote?.ItemId?.Id).to.be.equal(itemEWSId);
            expect(ewsNote?.ItemId?.ChangeKey).to.be.equal(`ck-${itemEWSId}`);
            expect(ewsNote?.ParentFolderId?.Id).to.be.equal(passedParentFolderEWSId);
            expect(ewsNote?.ParentFolderId?.ChangeKey).to.be.equal(`ck-${passedParentFolderEWSId}`);
            expect(ewsNote?.ItemClass).to.be.equal(ItemClassType.ITEM_CLASS_NOTE);
        });

    });

    describe('pimNotesHelper.pimNoteFromItem', () => {

        it('test all fields set', function () {
            const note = new MessageType();

            let pimNote = EWSNotesManager.getInstance().pimItemFromEWSItem(note, context.request, {});
            expect(pimNote.unid).to.be.undefined();
            expect(pimNote.categories).to.be.eql([]);

            note.Size = 2048;
            note.Subject = 'Note subject';
            note.Body = new BodyType('Note body', BodyTypeType.HTML);
            note.ItemClass = ItemClassType.ITEM_CLASS_NOTE;
            const mailboxId = 'test@test.com';
            const itemEWSId = getEWSId('NOTE-ID', mailboxId);
            note.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            const parentFolderEWSId = getEWSId('PARENT-ID', mailboxId);
            note.ParentFolderId = new FolderIdType(parentFolderEWSId, `ck-${parentFolderEWSId}`);
            note.DateTimeCreated = new Date('March 2, 1996');
            note.LastModifiedTime = new Date();
            note.Categories = new ArrayOfStringsType();
            note.Categories.String = ['Orange', 'Blue'];

            pimNote = EWSNotesManager.getInstance().pimItemFromEWSItem(note, context.request, {});
            expect(pimNote.unid).to.be.equal('NOTE-ID');
            expect(pimNote.subject).to.be.equal('Note subject');
            expect(pimNote.body).to.be.equal('Note body');
            expect(pimNote.parentFolderIds).to.be.deepEqual(['PARENT-ID']);
            expect(pimNote.diaryDate).to.be.eql(note.DateTimeCreated);
            // expect(pimNote.lastModifiedDate).to.be.eql(note.LastModifiedTime);  // not currently supported for notes
            expect(pimNote.categories).to.be.eql(['Orange', 'Blue']);

            // For code coverage
            note.Categories.String = ['Green'];
            note.IsRead = true;
            pimNote = EWSNotesManager.getInstance().pimItemFromEWSItem(note, context.request);
            expect(pimNote.unid).to.be.equal('NOTE-ID');
            expect(pimNote.categories).to.be.eql(['Green']);
        });

    });

    describe('pimNotesHelper.noteDefaultPropertiesForShape', () => {

        it('test all base shapes', function () {
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;

            let baseFields = new MockNotesManager().getFieldsForShape(shape);
            expect(baseFields.length).to.be.equal(3);
            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_ITEM_ID);
            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_PARENT_FOLDER_ID);
            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_ITEM_CLASS);

            shape.BaseShape = DefaultShapeNamesType.DEFAULT;
            baseFields = new MockNotesManager().getFieldsForShape(shape);
            expect(baseFields.length).to.be.equal(9);
            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_ITEM_ID);
            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_PARENT_FOLDER_ID);
            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_ITEM_CLASS);

            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_BODY);
            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_SUBJECT);
            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_DATE_TIME_CREATED);
            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS);
            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_SENSITIVITY);
            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_SIZE);

            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
            baseFields = new MockNotesManager().getFieldsForShape(shape);
            expect(baseFields.length).to.be.equal(13);
            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_ITEM_ID);
            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_PARENT_FOLDER_ID);
            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_ITEM_CLASS);
            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_BODY);
            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_SUBJECT);
            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_DATE_TIME_CREATED);
            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS);
            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_SENSITIVITY);
            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_SIZE);

            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_CATEGORIES);
            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_IMPORTANCE);
            expect(baseFields).to.containEql(UnindexedFieldURIType.ITEM_LAST_MODIFIED_TIME);

            // Try non-standard shape
            shape.BaseShape = DefaultShapeNamesType.PCX_PEOPLE_SEARCH;
            baseFields = new MockNotesManager().getFieldsForShape(shape);
            expect(baseFields).to.be.eql([]);

        });

    });

});
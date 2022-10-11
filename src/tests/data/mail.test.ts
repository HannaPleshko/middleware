/* eslint-disable @typescript-eslint/no-non-null-assertion */
import { expect, sinon } from '@loopback/testlab';
import { Request } from '@loopback/rest';
import util from 'util';
import { 
  ArchiveItem, CopyItem, MoveItem, CreateFolder, CreateFolderPath, CreateItem, DeleteFolder, DeleteItem, EmptyFolder,
  FindFolder,GetFolder, MoveFolder, SyncFolderHierarchy, SyncFolderItems, UpdateFolder, USER_MAILBOX_DATA,
  GetItem, GetAttachment,MarkAllItemsAsRead, MarkAsJunk, CreateAttachment, DeleteAttachment, UploadItems
} from '../../data/mail'
import { 
  base64Encode, KeepDeleteLabelResults, KeepMoveLabelResults, KeepPimAttachmentManager, KeepPimBaseResults, KeepPimConstants,
  KeepPimLabelManager, KeepPimManager, KeepPimMessageManager, PimAttachmentInfo, PimItem, PimItemFactory,
  PimItemFormat, PimLabel, PimLabelTypes, PimMessage, UserInfo
} from '@hcllabs/openclientkeepcomponent';
import { 
  ContainmentComparisonType, ContainmentModeType, CreateActionType, DefaultShapeNamesType, DisposalType, DistinguishedFolderIdNameType,
  FolderClassType, FolderQueryTraversalType, MessageDispositionType, ResponseClassType, ResponseCodeType,
  UnindexedFieldURIType 
} from '../../models/enum.model';
import {
  ArchiveItemType, BaseFolderType, BaseFolderIdType, CopyItemType, MoveItemType, CreateFolderPathType, CreateFolderType,
  CreateItemType, DeleteFolderType, DeleteItemType, DistinguishedFolderIdType, EmailAddressType, EmptyFolderType, 
  FindFolderType, FolderChangeType, FolderIdType, FolderResponseShapeType, FolderType, GetFolderType, ItemIdType, 
  ItemResponseShapeType, ItemType, MessageType, MoveFolderType, NonEmptyArrayOfAllItemsType, 
  NonEmptyArrayOfBaseFolderIdsType,NonEmptyArrayOfBaseItemIdsType, NonEmptyArrayOfFolderChangeDescriptionsType,
  NonEmptyArrayOfFolderChangesType, NonEmptyArrayOfFoldersType,OccurrenceItemIdType, SetFolderFieldType, 
  SyncFolderHierarchyType, SyncFolderItemsCreateType, SyncFolderItemsType, TargetFolderIdType, UpdateFolderType,
  MarkAllItemsAsReadType,AttachmentResponseShapeType, GetAttachmentType, GetItemType,
  NonEmptyArrayOfRequestAttachmentIdsType, RequestAttachmentIdType, SyncFolderItemsDeleteType,MarkAsJunkType,
  ItemAttachmentType, CreateAttachmentType, NonEmptyArrayOfAttachmentsType, FileAttachmentType, DeleteAttachmentType, UploadItemsType, NonEmptyArrayOfUploadItemsType, UploadItemType
} from '../../models/mail.model';
import {
  getContext, generateFolderCreateEvents, stubLabelsManager, resetUserMailBox, stubPimManager, createSinonStubInstance, 
  stubServiceManagerForContacts, generateTestInboxLabel, stubServiceManagerForMessages, generateTestTrashLabel,
  generateTestLabels, stubMessageManager
} from '../unitTestHelper';
import { DEFAULT_LABEL_EXCLUSIONS, EWSContactsManager, EWSMessageManager, EWSServiceManager, getDistinguishedFolderId,
  getEWSId, getKeepIdPair, getLabelByTarget, getUserId, pimLabelFromItem, rootFolderIdForUser 
} from '../../utils';
import { UserContext } from '../../keepcomponent';
import { PathToUnindexedFieldType } from '../../models/common.model';
import { ConstantValueType, ContainsExpressionType, OrType, RestrictionType } from '../../models/persona.model';
import { SyncStateInfo } from '../../data/SyncStateInfo';
import { SinonStub, SinonStubbedInstance } from 'sinon';

const testUser = "test.user@test.org";

describe('mail.ts tests', () => {
  describe('mail.SyncFolderHierarchy', () => {
    function generateSoapRequest(valid = true, mailboxId?: string): SyncFolderHierarchyType{
      const soapRequest = new SyncFolderHierarchyType();

      soapRequest.FolderShape = new FolderResponseShapeType();
      soapRequest.FolderShape.BaseShape = DefaultShapeNamesType.DEFAULT;

      if (valid) {
        soapRequest.SyncState = `${Date.now() - 10000}`; // Must be in the past
      } else soapRequest.SyncState = `${Date.now() + 10000}`; // make it larger that what SyncFolderHierarchy will assign
      
      if (mailboxId) {
        soapRequest.SyncFolderId = new TargetFolderIdType();
        soapRequest.SyncFolderId.DistinguishedFolderId = new DistinguishedFolderIdType();
        soapRequest.SyncFolderId.DistinguishedFolderId.Id = DistinguishedFolderIdNameType.INBOX;
        soapRequest.SyncFolderId.DistinguishedFolderId.Mailbox = new EmailAddressType();
        soapRequest.SyncFolderId.DistinguishedFolderId.Mailbox.EmailAddress = mailboxId;
      }
      return soapRequest;
    }

    let labelManagerStub: sinon.SinonStubbedInstance<KeepPimLabelManager>;

    const context = getContext(testUser, 'password');
    
    beforeEach(() => {
      labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      stubLabelsManager(labelManagerStub);
    });

    afterEach(() => {
      sinon.restore(); // Reset stubs
      resetUserMailBox(testUser); // Reset local user cache
    });

    it('passes when no folders exist', async function () {
      labelManagerStub.getLabels.resolves([]);
      const soapRequest = generateSoapRequest();
      
      const result = await SyncFolderHierarchy(soapRequest, context.request);

      expect(result.ResponseMessages.items.length).to.equal(1);
      expect(result.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.SUCCESS);
      expect(result.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.NO_ERROR);
      expect(result.ResponseMessages.items[0].SyncState).to.not.be.null();
      expect(result.ResponseMessages.items[0].Changes.items.length).to.equal(0);
    });

    it('returns an error for invalid sync state', async function () {
      const soapRequest = generateSoapRequest(false);

      const result = await SyncFolderHierarchy(soapRequest, context.request);

      expect(result.ResponseMessages.items.length).to.equal(1);
      expect(result.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.ERROR);
      expect(result.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.ERROR_INVALID_SYNC_STATE_DATA);
      expect(result.ResponseMessages.items[0].SyncState).to.not.be.null();
      expect(result.ResponseMessages.items[0].Changes.items.length).to.equal(0);
    });

    it('passes when sync state is valid, new folders added, events exist, a delegator email set', async function () {
      const delegatorMailboxId = 'delegator@test.com';
      // Setup events in the event queue. This will be use when sync state is set
      const labels = generateFolderCreateEvents(testUser, 2, delegatorMailboxId);
      // Stub the labels manager to return the two new folders
      labelManagerStub.getLabels.resolves(labels);
      const userInfo = UserContext.getUserInfo(context.request);
      const soapRequest = generateSoapRequest(true, delegatorMailboxId);

      const result = await SyncFolderHierarchy(soapRequest, context.request);

      expect(labelManagerStub.getLabels.calledWith(userInfo, true, DEFAULT_LABEL_EXCLUSIONS, delegatorMailboxId)).to.be.equal(true);
      expect(result.ResponseMessages.items.length).to.equal(1);
      expect(result.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.SUCCESS);
      expect(result.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.NO_ERROR);
      expect(result.ResponseMessages.items[0].SyncState).to.not.be.null();
      expect(result.ResponseMessages.items[0].Changes.items.length).to.equal(2);
    });

    it('passes when sync state is valid, new folders added, no events', async function () {
      const soapRequest = generateSoapRequest();

      // Stub the labels manager to return the two new folders
      const labels = [
        PimItemFactory.newPimLabel({ "FolderId": "folder-0001", "DisplayName": "Unit Test Folder One" }, PimItemFormat.DOCUMENT),
        PimItemFactory.newPimLabel({ "FolderId": "folder-0002", "DisplayName": "Unit Test Folder Two" }, PimItemFormat.DOCUMENT)
      ];
      labelManagerStub.getLabels.resolves(labels);

      const result = await SyncFolderHierarchy(soapRequest, context.request);

      expect(result.ResponseMessages.items.length).to.equal(1);
      expect(result.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.SUCCESS);
      expect(result.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.NO_ERROR);
      expect(result.ResponseMessages.items[0].SyncState).to.not.be.null();
      expect(result.ResponseMessages.items[0].Changes.items.length).to.equal(2);
    });

  });

  describe("mail.DeleteItem", () => {

    afterEach(() => {
      sinon.restore(); // Reset stubs
      resetUserMailBox(testUser); // Reset local user cache
    });
    it('deleted successfully with item', async function () {

      const pimItem = PimItemFactory.newPimContact({ unid: 'test-id' });
      const stub = sinon.createStubInstance(KeepPimManager);
      stub.getPimItem.resolves(pimItem);
      stubPimManager(stub);

      const serviceStub = createSinonStubInstance(EWSContactsManager);
      serviceStub.deleteItem.resolves();
      stubServiceManagerForContacts(serviceStub);

      const soapRequest = new DeleteItemType();
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([new ItemIdType('test-id')]);
      const context = getContext(testUser, 'password');

      const response = await DeleteItem(soapRequest, context.request);
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);

    });

    it('deleted successfully with undefined', async function () {

      const stub = sinon.createStubInstance(KeepPimManager);
      stub.getPimItem.resolves(undefined);
      stubPimManager(stub);

      const soapRequest = new DeleteItemType();
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([new ItemIdType('test-id')]);
      const context = getContext(testUser, 'password');

      const response = await DeleteItem(soapRequest, context.request);
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);

    });

    it('deleteItem Fails when 404', async function () {

      const error = { status: 404, message: "error code: 549" }
      const stub = sinon.createStubInstance(KeepPimManager);
      stub.getPimItem.rejects(error);
      stubPimManager(stub);

      const soapRequest = new DeleteItemType();
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([new ItemIdType('test-id')]);
      const context = getContext(testUser, 'password');

      const response = await DeleteItem(soapRequest, context.request);
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_ITEM_NOT_FOUND);

    });

    it('deleteItem fails when 303', async function () {

      const error = { status: 303, message: "unittest" }
      const stub = sinon.createStubInstance(KeepPimManager);
      stub.getPimItem.rejects(error);
      stubPimManager(stub);

      const soapRequest = new DeleteItemType();
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([new ItemIdType('test-id')]);
      const context = getContext(testUser, 'password');

      const response = await DeleteItem(soapRequest, context.request);
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_CANNOT_DELETE_OBJECT);
      expect(response.ResponseMessages.items[0].MessageText).to.equal('unittest');

    });

    it('When Item id type is not supported', async function () {

      const soapRequest = new DeleteItemType();
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType();
      soapRequest.ItemIds.items = [new OccurrenceItemIdType()];

      const context = getContext(testUser, "password");
      const result = await DeleteItem(soapRequest, context.request);

      expect(result.ResponseMessages.items.length).to.equal(1);
      expect(result.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.ERROR)
      expect(result.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.ERROR_INVALID_ID);
    });

  });

  describe('mail.CopyItem', function () {

    const context = getContext(testUser, "password");

    afterEach(() => {
      sinon.restore(); // Reset stubs
      resetUserMailBox(testUser); // Reset local user cache
    });

    it('to folder is not found', async function () {

      const stub = sinon.createStubInstance(KeepPimLabelManager);
      stub.getLabel.resolves();
      stubLabelsManager(stub);

      const soapRequest = new CopyItemType();
      soapRequest.ToFolderId = new TargetFolderIdType()
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([new ItemIdType('test-item-id')])

      const response = await CopyItem(soapRequest, context.request);
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_TO_FOLDER_NOT_FOUND);

    });

    it('getting label throws an error', async function () {

      const stub = sinon.createStubInstance(KeepPimLabelManager);
      stub.getLabel.rejects("getting label throws an error");
      stubLabelsManager(stub);

      const soapRequest = new CopyItemType();
      soapRequest.ToFolderId = new TargetFolderIdType()
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([new ItemIdType('test-item-id')])

      const response = await CopyItem(soapRequest, context.request);
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_TO_FOLDER_NOT_FOUND);
    });

    it('Successful test', async function () {

      const stub = sinon.createStubInstance(KeepPimLabelManager);
      stub.getLabel.resolves(generateTestInboxLabel());
      stubLabelsManager(stub);

      const pimItem = PimItemFactory.newPimMessage({ '@unid': 'test-id' });
      const stub2 = sinon.createStubInstance(KeepPimManager);
      stub2.getPimItem.resolves(pimItem);
      stubPimManager(stub2);

      const soapRequest = new CopyItemType();
      soapRequest.ToFolderId = new TargetFolderIdType();
      soapRequest.ToFolderId.FolderId = new FolderIdType("test-folder-id");
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([new ItemIdType('test-item-id')])

      const serviceStub = createSinonStubInstance(EWSMessageManager);
      const item = new ItemType();
      item.ItemId = new ItemIdType("test-item-id");
      serviceStub.copyItem.resolves(item);

      stubServiceManagerForMessages(serviceStub);

      const response = await CopyItem(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].Items.items[0].ItemId?.Id).to.be.equal('test-item-id')

    });

    it('to folder is Trash', async function () {
      // Stub label manger to return the trash label
      const stub = sinon.createStubInstance(KeepPimLabelManager);
      stub.getLabel.resolves(generateTestTrashLabel());
      stubLabelsManager(stub);

      // Stub KeepPimManager.getPimItem to return a Pim Message item
      const pimItem = PimItemFactory.newPimMessage({ '@unid': 'test-id' });
      const stub2 = sinon.createStubInstance(KeepPimManager);
      stub2.getPimItem.resolves(pimItem);
      stubPimManager(stub2);

      // Stub EWSServiceManager.getItem
      const item = new ItemType();
      item.ItemId = new ItemIdType("test-item-id");
      sinon.stub(EWSServiceManager, "getItem").callsFake((itemId: string, userInfo: UserInfo, request: Request, shape?: ItemResponseShapeType) => {
        return Promise.resolve(item);
      });

      // Stub EWSServiceManager.getInstanceFromPimItem
      const serviceStub = createSinonStubInstance(EWSMessageManager);
      serviceStub.deleteItem.resolves();
      sinon.stub(EWSServiceManager, "getInstanceFromPimItem").callsFake((item2: PimItem) => {
        return serviceStub;
      });

      // Build a test Soap request
      const soapRequest = new CopyItemType();
      soapRequest.ToFolderId = new TargetFolderIdType();
      soapRequest.ToFolderId.FolderId = new FolderIdType("test-folder-id");
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([new ItemIdType('test-item-id')])


      const response = await CopyItem(soapRequest, context.request);
      expect(response.ResponseMessages.items[0].Items.items[0].ItemId?.Id).to.be.equal('test-item-id')

    })

    it('to folder is Trash, item not found', async function () {

      const stub = sinon.createStubInstance(KeepPimLabelManager);
      stub.getLabel.resolves(generateTestTrashLabel());
      stubLabelsManager(stub);

      sinon.stub(EWSServiceManager, "getItem").callsFake((itemId: string, userInfo: UserInfo, request: Request, shape?: ItemResponseShapeType) => {
        console.info('Creating EWSServiceManager stub for getInstanceFromPimItem to return a contacts manager stub');
        return Promise.resolve(undefined);
      });

      const soapRequest = new CopyItemType();
      soapRequest.ToFolderId = new TargetFolderIdType()
      soapRequest.ToFolderId.FolderId = new FolderIdType("test-folder-id");
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([new ItemIdType('test-item-id')])

      const response = await CopyItem(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_ITEM_NOT_FOUND);
    });

    it('to folder is Trash, error thrown', async function () {

      const stub = sinon.createStubInstance(KeepPimLabelManager);
      stub.getLabel.resolves(generateTestTrashLabel());
      stubLabelsManager(stub);

      const stub2 = sinon.createStubInstance(KeepPimManager);
      stub2.getPimItem.rejects();
      stubPimManager(stub2);

      const soapRequest = new CopyItemType();
      soapRequest.ToFolderId = new TargetFolderIdType();
      soapRequest.ToFolderId.FolderId = new FolderIdType("test-folder-id");
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([new ItemIdType('test-item-id')])

      const response = await CopyItem(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_MOVE_COPY_FAILED);
    });

    it('pim item not found', async function () {

      const stub = sinon.createStubInstance(KeepPimLabelManager);
      stub.getLabel.resolves(generateTestInboxLabel());
      stubLabelsManager(stub);

      const stub2 = sinon.createStubInstance(KeepPimManager);
      stub2.getPimItem.resolves(undefined);
      stubPimManager(stub2);

      const soapRequest = new CopyItemType();
      soapRequest.ToFolderId = new TargetFolderIdType();
      soapRequest.ToFolderId.FolderId = new FolderIdType("test-folder-id");
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([new ItemIdType('test-item-id')])

      const response = await CopyItem(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_ITEM_NOT_FOUND);
    });

    it('get pim item throws an error', async function () {

      const stub = sinon.createStubInstance(KeepPimLabelManager);
      stub.getLabel.resolves(generateTestInboxLabel());
      stubLabelsManager(stub);

      const error: any = new Error('ERROR_INTERNAL_SERVER_ERROR');
      error.ResponseCode = ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR;

      const stub2 = sinon.createStubInstance(KeepPimManager);
      stub2.getPimItem.rejects(error);
      stubPimManager(stub2);

      const soapRequest = new CopyItemType();
      soapRequest.ToFolderId = new TargetFolderIdType();
      soapRequest.ToFolderId.FolderId = new FolderIdType("test-folder-id");
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([new ItemIdType('test-item-id')])

      const response = await CopyItem(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR);
    });

    it('service manager throws an error', async function () {

      const stub = sinon.createStubInstance(KeepPimLabelManager);
      stub.getLabel.resolves(generateTestInboxLabel());
      stubLabelsManager(stub);

      const pimItem = PimItemFactory.newPimMessage({ '@unid': 'test-item-id' });
      const stub2 = sinon.createStubInstance(KeepPimManager);
      stub2.getPimItem.resolves(pimItem);
      stubPimManager(stub2);

      const soapRequest = new CopyItemType();
      soapRequest.ToFolderId = new TargetFolderIdType();
      soapRequest.ToFolderId.FolderId = new FolderIdType("test-folder-id");
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([new ItemIdType('test-item-id')])

      const serviceStub = createSinonStubInstance(EWSMessageManager);
      serviceStub.copyItem.rejects("Move copy failed");
      stubServiceManagerForMessages(serviceStub);

      const response = await CopyItem(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_MOVE_COPY_FAILED);
    });

    it('service manager not found', async function () {

      const stub = sinon.createStubInstance(KeepPimLabelManager);
      stub.getLabel.resolves(generateTestInboxLabel());
      stubLabelsManager(stub);

      const pimItem = PimItemFactory.newPimMessage({ '@unid': 'test-item-id' });
      const stub2 = sinon.createStubInstance(KeepPimManager);
      stub2.getPimItem.resolves(pimItem);
      stubPimManager(stub2);

      const soapRequest = new CopyItemType();
      soapRequest.ToFolderId = new TargetFolderIdType();
      soapRequest.ToFolderId.FolderId = new FolderIdType("test-folder-id");
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([new ItemIdType('test-item-id')])

      sinon.stub(EWSServiceManager, "getInstanceFromPimItem").callsFake((item: PimItem) => {
        return undefined;
      });

      const response = await CopyItem(soapRequest, context.request);
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_MOVE_COPY_FAILED);

    });
  });


  describe('mail.MoveItem', function () {

    const context = getContext(testUser, "password");


    function stubGetItem(result: Promise<ItemType | undefined | Error>): SinonStub {
      return sinon.stub(EWSServiceManager, 'getItem').callsFake((): any => {
        return result;
      });
    }

    function stubMoveItemWithResult(result: Promise<ItemType | undefined | Error>): SinonStub {
      return sinon.stub(EWSServiceManager, 'moveItemWithResult').callsFake((): any => {
        return result;
      });
    }
    afterEach(() => {
      sinon.restore(); // Reset stubs
      resetUserMailBox(testUser); // Reset local user cache
    });

    it('to folder is not found', async function () {

      const stub = sinon.createStubInstance(KeepPimLabelManager);
      stub.getLabel.resolves();
      stubLabelsManager(stub);

      const soapRequest = new MoveItemType();
      soapRequest.ToFolderId = new TargetFolderIdType()
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([new ItemIdType('test-item-id')])

      const response = await MoveItem(soapRequest, context.request);
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_TO_FOLDER_NOT_FOUND);

    });

    it('getting label throws an error', async function () {

      const stub = sinon.createStubInstance(KeepPimLabelManager);
      stub.getLabel.rejects(new Error("getting label throws an error"));
      stubLabelsManager(stub);

      const soapRequest = new MoveItemType();
      soapRequest.ToFolderId = new TargetFolderIdType()
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([new ItemIdType('test-item-id')])

      const response = await MoveItem(soapRequest, context.request);
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_TO_FOLDER_NOT_FOUND);
    });

    it('Successful test', async function () {

      const stub = sinon.createStubInstance(KeepPimLabelManager);
      stub.getLabel.resolves(generateTestInboxLabel());
      stubLabelsManager(stub);

      const pimItem = PimItemFactory.newPimMessage({ '@unid': 'test-id' });
      const stub2 = sinon.createStubInstance(KeepPimManager);
      stub2.getPimItem.resolves(pimItem);
      stubPimManager(stub2);

      const item = new ItemType();
      item.ItemId = new ItemIdType("test-item-id");
      item.ParentFolderId = new FolderIdType("test-folder-id");

      stubGetItem(Promise.resolve(item));
      stubMoveItemWithResult(Promise.resolve(item))

      const soapRequest = new MoveItemType();
      soapRequest.ToFolderId = new TargetFolderIdType();
      soapRequest.ToFolderId.FolderId = new FolderIdType("test-folder-id");
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([new ItemIdType('')])

      const response = await MoveItem(soapRequest, context.request);
      expect(response.ResponseMessages.items[0].Items.items[0].ItemId?.Id).to.be.equal('test-item-id');
    });

    it('Successful test without parent folder', async function () {

      const stub = sinon.createStubInstance(KeepPimLabelManager);
      stub.getLabel.resolves(generateTestInboxLabel());
      stubLabelsManager(stub);

      const pimItem = PimItemFactory.newPimMessage({ '@unid': 'test-id' });
      const stub2 = sinon.createStubInstance(KeepPimManager);
      stub2.getPimItem.resolves(pimItem);
      stubPimManager(stub2);

      const item = new ItemType();
      item.ItemId = new ItemIdType("test-item-id");

      stubGetItem(Promise.resolve(item));
      stubMoveItemWithResult(Promise.resolve(item))

      const soapRequest = new MoveItemType();
      soapRequest.ToFolderId = new TargetFolderIdType();
      soapRequest.ToFolderId.FolderId = new FolderIdType("test-folder-id");
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([new ItemIdType('')])

      const response = await MoveItem(soapRequest, context.request);
      expect(response.ResponseMessages.items[0].Items.items[0].ItemId?.Id).to.be.equal('test-item-id');
    });
    
    it('passed item doesn\'t exist', async function () {

      const stub = sinon.createStubInstance(KeepPimLabelManager);
      stub.getLabel.resolves(generateTestInboxLabel());
      stubLabelsManager(stub);

      const pimItem = PimItemFactory.newPimMessage({ '@unid': 'test-id' });
      const stub2 = sinon.createStubInstance(KeepPimManager);
      stub2.getPimItem.resolves(pimItem);
      stubPimManager(stub2);

      const item = new ItemType();
      item.ItemId = new ItemIdType("test-item-id");

      stubGetItem(Promise.resolve(undefined));
      stubMoveItemWithResult(Promise.resolve(item))

      const soapRequest = new MoveItemType();
      soapRequest.ToFolderId = new TargetFolderIdType();
      soapRequest.ToFolderId.FolderId = new FolderIdType("test-folder-id");
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([new ItemIdType('test-item-id')])

      const response = await MoveItem(soapRequest, context.request);
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_MOVE_COPY_FAILED);
    });

    it('Could not retrieve moved item', async function () {

      const stub = sinon.createStubInstance(KeepPimLabelManager);
      stub.getLabel.resolves(generateTestInboxLabel());
      stubLabelsManager(stub);

      const pimItem = PimItemFactory.newPimMessage({ '@unid': 'test-id' });
      const stub2 = sinon.createStubInstance(KeepPimManager);
      stub2.getPimItem.resolves(pimItem);
      stubPimManager(stub2);

      const item = new ItemType();
      item.ItemId = new ItemIdType("test-item-id");

      stubGetItem(Promise.resolve(item));
      stubMoveItemWithResult(Promise.resolve(undefined))

      const soapRequest = new MoveItemType();
      soapRequest.ToFolderId = new TargetFolderIdType();
      soapRequest.ToFolderId.FolderId = new FolderIdType("test-folder-id");
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([new ItemIdType('test-item-id')])

      const response = await MoveItem(soapRequest, context.request);
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_MOVE_COPY_FAILED);
    });

    it('MoveItemWithResult throws an error', async function () {

      const stub = sinon.createStubInstance(KeepPimLabelManager);
      stub.getLabel.resolves(generateTestInboxLabel());
      stubLabelsManager(stub);

      const pimItem = PimItemFactory.newPimMessage({ '@unid': 'test-id' });
      const stub2 = sinon.createStubInstance(KeepPimManager);
      stub2.getPimItem.resolves(pimItem);
      stubPimManager(stub2);

      const item = new ItemType();
      item.ItemId = new ItemIdType("test-item-id");

      stubGetItem(Promise.resolve(item));
      stubMoveItemWithResult(Promise.reject(new Error))

      const soapRequest = new MoveItemType();
      soapRequest.ToFolderId = new TargetFolderIdType();
      soapRequest.ToFolderId.FolderId = new FolderIdType("test-folder-id");
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([new ItemIdType('test-item-id')])

      const response = await MoveItem(soapRequest, context.request);
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_MOVE_COPY_FAILED);
    });
  });

  describe('mail.ArchiveItem', function() {
    /*
    From LABS-2458
    Unit test to cover ArchiveItem method in mail.ts
    Test:
    - Archive folder does not exist
    - Archive folder does exist
    - OccurenceIdType passed in as ItemId to get ERROR_INVALID_ID
    - ItemId that does not exist
    - ItemId that does exist and succeeds
    - Getting a 401 from a keep call like getItem
    */
    
    const context = getContext(testUser, "password");

    function stubGetItem(result: Promise<ItemType | undefined | Error>): SinonStub {
      return sinon.stub(EWSServiceManager, 'getItem').callsFake((): any => {
        return result;
      });
    }

    function stubMoveItemWithResult(result: Promise<ItemType | undefined | Error>): SinonStub {
      return sinon.stub(EWSServiceManager, 'moveItemWithResult').callsFake((): any => {
        return result;
      });
    }

    afterEach(() => {
      sinon.restore(); // Reset stubs
      resetUserMailBox(testUser); // Reset local user cache
    });

    it('should handle ArchiveSourceFolderId that exists', async function() {
      const item = new ItemType();
      item.ItemId = new ItemIdType("test-item-id");
      item.ParentFolderId = new FolderIdType(getEWSId("test-folder-id"));

      stubGetItem(Promise.resolve(item));
      stubMoveItemWithResult(Promise.resolve(item))

      const soapRequest = new ArchiveItemType();
      const labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      stubLabelsManager(labelManagerStub);
      const labels = generateTestLabels();
      labels.push(PimItemFactory.newPimLabel({ FolderId: DistinguishedFolderIdNameType.ARCHIVE }));
      labelManagerStub.getLabels.resolves(labels);
      
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([item.ItemId]);
      soapRequest.ArchiveSourceFolderId = { FolderId: new FolderIdType(getEWSId("test-folder-id")) };

      const response = await ArchiveItem(soapRequest, context.request);
      expect(response.ResponseMessages.items[0].Items.items[0].ItemId?.Id).to.be.equal('test-item-id');
    });

    it('should handle ArchiveSourceFolderId that does not exist', async function() {

      const item = new ItemType();
      item.ItemId = new ItemIdType("test-item-id");
      item.ParentFolderId = new FolderIdType(getEWSId("test-folder-id"));

      stubGetItem(Promise.resolve(item));
      stubMoveItemWithResult(Promise.resolve(item))

      const soapRequest = new ArchiveItemType();
      const labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      stubLabelsManager(labelManagerStub);
      const labels = generateTestLabels();
      labels.push(PimItemFactory.newPimLabel({ FolderId: DistinguishedFolderIdNameType.ARCHIVE }));
      labelManagerStub.getLabels.resolves(labels);
      
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([item.ItemId]);
      soapRequest.ArchiveSourceFolderId = { FolderId: new FolderIdType("non-existed") };

      const response = await ArchiveItem(soapRequest, context.request);
      expect(response.ResponseMessages.items[0].Items.items[0].ItemId?.Id).to.be.equal('test-item-id');
    });

    it('should handle Archive folder not existing', async function () {

      const item = new ItemType();
      item.ItemId = new ItemIdType("test-item-id");
      item.ParentFolderId = new FolderIdType(getEWSId("test-folder-id"));

      const soapRequest = new ArchiveItemType();
      const labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      stubLabelsManager(labelManagerStub);
      const labels = generateTestLabels();
      labelManagerStub.getLabels.resolves(labels);
      
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([item.ItemId]);

      const response = await ArchiveItem(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR)
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_ARCHIVE_MAILBOX_NOT_ENABLED)
    });

    it('passes a ItemId that does not exists', async function() {
      const soapRequest = new ArchiveItemType();
      const labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      stubLabelsManager(labelManagerStub);
      const labels = generateTestLabels();
      labels.push(PimItemFactory.newPimLabel({ FolderId: DistinguishedFolderIdNameType.ARCHIVE }));
      labelManagerStub.getLabels.resolves(labels);
      const item = new ItemType();
      item.ItemId = new ItemIdType("test-item-id");
      item.ParentFolderId = new FolderIdType(getEWSId("test-folder-id"));

      stubGetItem(Promise.reject(item));

      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([item.ItemId]);

      const response = await ArchiveItem(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR)
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_MOVE_COPY_FAILED)
    });

    it('should handle moving existing ItemId to existing Archive folder', async function () {

      const item = new ItemType();
      item.ItemId = new ItemIdType("test-item-id");
      item.ParentFolderId = new FolderIdType(getEWSId("test-folder-id"));

      stubGetItem(Promise.resolve(item));
      stubMoveItemWithResult(Promise.resolve(item))

      const soapRequest = new ArchiveItemType();
      const labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      stubLabelsManager(labelManagerStub);
      const labels = generateTestLabels();
      labels.push(PimItemFactory.newPimLabel({ FolderId: DistinguishedFolderIdNameType.ARCHIVE }));
      labelManagerStub.getLabels.resolves(labels);
      
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([item.ItemId]);

      const response = await ArchiveItem(soapRequest, context.request);
      expect(response.ResponseMessages.items[0].Items.items[0].ItemId?.Id).to.be.equal('test-item-id');
    });

    it('passes OccurenceIdType in as ItemId to get ERROR_INVALID_ID', async function () {

      stubGetItem(Promise.resolve(undefined));
      stubMoveItemWithResult(Promise.resolve(undefined))

      const soapRequest = new ArchiveItemType();
      const labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      stubLabelsManager(labelManagerStub);
      const labels = generateTestLabels();
      labels.push(PimItemFactory.newPimLabel({ FolderId: 'test-folder-id' }));
      labels.push(PimItemFactory.newPimLabel({ FolderId: DistinguishedFolderIdNameType.ARCHIVE }));
      labelManagerStub.getLabels.resolves(labels);
      
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([new OccurrenceItemIdType()]);

      const response = await ArchiveItem(soapRequest, context.request);
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR)
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_INVALID_ID)
    });

    it('should get a null from a keep call getItem', async function () {
      const item = new ItemType();
      const itemEWSId = getEWSId("test-item-id");
      item.ItemId = new ItemIdType(itemEWSId);
      item.ParentFolderId = new FolderIdType(getEWSId("test-folder-id"));

      stubGetItem(Promise.resolve(undefined));
      stubMoveItemWithResult(Promise.resolve(item))

      const soapRequest = new ArchiveItemType();
      const labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      stubLabelsManager(labelManagerStub);
      const labels = generateTestLabels();
      labels.push(PimItemFactory.newPimLabel({ FolderId: 'test-folder-id' }));
      labels.push(PimItemFactory.newPimLabel({ FolderId: DistinguishedFolderIdNameType.ARCHIVE }));
      labelManagerStub.getLabels.resolves(labels);
      
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([item.ItemId]);

      const response = await ArchiveItem(soapRequest, context.request);
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR)
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_MOVE_COPY_FAILED)
      expect(response.ResponseMessages.items[0].MessageText).to.be.equal(`Error: Could not retrieve item to archive ${itemEWSId}`)
    });

    it('should get a null from a keep call moveItemWithResult', async function () {
      const item = new ItemType();
      const itemEWSId = getEWSId("test-item-id");
      item.ItemId = new ItemIdType(itemEWSId);
      item.ParentFolderId = new FolderIdType(getEWSId("test-folder-id"));

      stubGetItem(Promise.resolve(item));
      stubMoveItemWithResult(Promise.resolve(undefined))

      const soapRequest = new ArchiveItemType();
      const labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      stubLabelsManager(labelManagerStub);
      const labels = generateTestLabels();
      labels.push(PimItemFactory.newPimLabel({ FolderId: 'test-folder-id' }));
      labels.push(PimItemFactory.newPimLabel({ FolderId: DistinguishedFolderIdNameType.ARCHIVE }));
      labelManagerStub.getLabels.resolves(labels);
      
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([item.ItemId]);

      const response = await ArchiveItem(soapRequest, context.request);
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR)
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_MOVE_COPY_FAILED)
      expect(response.ResponseMessages.items[0].MessageText).to.be.equal(`Error: Could not retrieve archived item ${itemEWSId}`)
    });
  });

  describe('mail.GetFolder', function () {
    function generateSoapRequest(
      shape: DefaultShapeNamesType,
      empty = true,
      root = true,
      ): GetFolderType {
      const soapRequest = new GetFolderType();
      
      let items: BaseFolderIdType[] = [];
      if (!empty) {
        items = generateItems(root);
      }

      soapRequest.FolderIds = new NonEmptyArrayOfBaseFolderIdsType(items);

      soapRequest.FolderShape = new FolderResponseShapeType();
      soapRequest.FolderShape.BaseShape = shape;

      return soapRequest;
    }

    function generateItems(root: boolean): BaseFolderIdType[] {
      const items: BaseFolderIdType[] = [];

      if (root) {
        const item = new DistinguishedFolderIdType();
        item.Id = DistinguishedFolderIdNameType.ROOT;
        items.push(item);
      } else {
        const item1 = new DistinguishedFolderIdType();
        item1.Id = DistinguishedFolderIdNameType.ROOT;
        const item2 = new DistinguishedFolderIdType();
        item2.Id = DistinguishedFolderIdNameType.INBOX;
        const item3 = new FolderIdType(delegatorFolderId);
        items.push(item1, item2, item3);
      }

      return items;
    }

    let labelManagerStub: sinon.SinonStubbedInstance<KeepPimLabelManager>;

    const delegatorMailboxId = 'delegator@test.com';
    const delegatorFolderId = getEWSId('test-folder-id', delegatorMailboxId);
    const context = getContext(testUser, 'password');

    beforeEach(() => {
      labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      stubLabelsManager(labelManagerStub);
      const label = PimItemFactory.newPimLabel({ FolderId: 'test-folder-id' });
      labelManagerStub.getLabels.callsFake(
        (userInfo: UserInfo, includeUnread?: boolean | undefined, ignore?: any, mailboxId?: string | undefined): Promise<PimLabel[]> => {
          if (mailboxId === delegatorMailboxId) {
            return Promise.resolve([label]);
          } else return Promise.resolve(generateTestLabels());
        }
      );
    });

    afterEach(() => {
      sinon.restore(); // Reset stubs
      resetUserMailBox(testUser); // Reset local user cache
    });
    it('passes for a root folder', async function () {
      const soapRequest = generateSoapRequest(DefaultShapeNamesType.ID_ONLY, false, true);

      const response = await GetFolder(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
      expect(response.ResponseMessages.items[0].Folders).to.not.be.null();
    });

    it('returns an error when the soap request has no FolderId items', async function () {
      const soapRequest = generateSoapRequest(DefaultShapeNamesType.DEFAULT);

      const response = await GetFolder(soapRequest, context.request);

      expect(response.ResponseMessages.items.length).to.equal(1);
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_MISSING_ARGUMENT);
    });

    it('passes when for different folders with different mailboxes', async function () {
      const soapRequest = generateSoapRequest(DefaultShapeNamesType.DEFAULT, false, false);
      const item1FolderId = rootFolderIdForUser(getUserId(context.request)!).Id;
      const item2FolderId = getEWSId('inbox-id');

      const response = await GetFolder(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
      expect(response.ResponseMessages.items[0].Folders.items[0].FolderId?.Id).to.be.equal(item1FolderId);

      expect(response.ResponseMessages.items[1].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessages.items[1].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
      expect(response.ResponseMessages.items[1].Folders.items[0].FolderId?.Id).to.equal(item2FolderId);

      expect(response.ResponseMessages.items[2].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessages.items[2].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
      expect(response.ResponseMessages.items[2].Folders.items[0].FolderId?.Id).to.equal(delegatorFolderId);
    });

    it('returns an error for the soapRequest contains a folder that does not match in the list of labels', async function() {
      const soapRequest = generateSoapRequest(DefaultShapeNamesType.DEFAULT);
      soapRequest.FolderIds.items[0] = new FolderIdType('test-folder-id');

      const response = await GetFolder(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_FOLDER_NOT_FOUND);
    });
  });

  describe('mail.CreateFolder', function () {
    enum Parent {
      root = 'root',
      inbox = 'inbox-id',
      unknown = 'unknown',
      folder = 'test-folder-id'
    }

    function generateSoapRequest(parent: Parent, emty = false): CreateFolderType {
      const soapRequest = new CreateFolderType();

      soapRequest.ParentFolderId = new TargetFolderIdType();
      if (parent === Parent.root) {
        soapRequest.ParentFolderId.DistinguishedFolderId = new DistinguishedFolderIdType();
        soapRequest.ParentFolderId.DistinguishedFolderId.Id = DistinguishedFolderIdNameType.ROOT;
      }
      if (parent === Parent.inbox) {
        soapRequest.ParentFolderId.DistinguishedFolderId = new DistinguishedFolderIdType();
        soapRequest.ParentFolderId.DistinguishedFolderId.Id = DistinguishedFolderIdNameType.INBOX;
      } 
      if (parent === Parent.folder) {
        soapRequest.ParentFolderId.FolderId = new FolderIdType(parentFolderEWSId);
      }
      if (parent === Parent.unknown) {
        soapRequest.ParentFolderId = new TargetFolderIdType();
      }

      let items: BaseFolderType[] = [{}];
      if (!emty) {
        items = [{ FolderId: new FolderIdType(item.id), DisplayName: item.name }];
      }
      soapRequest.Folders = new NonEmptyArrayOfFoldersType(items);

      return soapRequest;
    }
    
    let labelManagerStub: sinon.SinonStubbedInstance<KeepPimLabelManager>;

    const delegatorMailboxId = 'delegator@test.com';
    const parentFolderEWSId = getEWSId(Parent.folder, delegatorMailboxId);

    const item = { id: 'test-tem-id', name: 'test-item' };
    const label = PimItemFactory.newPimLabel({ FolderId: item.id});
    label.displayName = item.name;
    const folderEWSId = getEWSId(item.id);

    const context = getContext(testUser, 'password');
    const userInfo = UserContext.getUserInfo(context.request);

    beforeEach(() => {
      labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      labelManagerStub.getLabels.resolves(generateTestLabels());
      labelManagerStub.createLabel.resolves(label);
      stubLabelsManager(labelManagerStub);
    });

    afterEach(() => {
      sinon.restore(); // Reset stubs
      resetUserMailBox(testUser); // Reset local user cache
    });

    it('passes for a delegator parent folder FolderType', async function() {
      const soapRequest = generateSoapRequest(Parent.folder);
      label.parentFolderId = Parent.folder;
      const parentLabel = PimItemFactory.newPimLabel({ FolderId: Parent.folder, View: 'test-parent-folder', Type: PimLabelTypes.MAIL });
      labelManagerStub.getLabel.resolves(parentLabel);
      const pimLabelIn = pimLabelFromItem(soapRequest.Folders.items[0], {}, parentLabel.type);
      pimLabelIn.parentFolderId = parentLabel?.folderId;

      const response = await CreateFolder(soapRequest, context.request);

      expect(labelManagerStub.createLabel.calledWith(userInfo, pimLabelIn, parentLabel.type)).to.be.equal(true);
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
      expect(response.ResponseMessages.items[0].Folders.items[0].ParentFolderId?.Id).to.be.equal(parentFolderEWSId);
      expect(response.ResponseMessages.items[0].Folders.items[0].FolderId?.Id).to.be.equal(getEWSId(item.id, delegatorMailboxId));
      expect(response.ResponseMessages.items[0].Folders.items[0].DisplayName).to.be.equal(item.name);
    });

    it('passes for a parent inbox folder and one inner folder', async function () {
      const soapRequest = generateSoapRequest(Parent.inbox);
      label.parentFolderId = Parent.inbox;
      const parentInboxEWSId = getEWSId(Parent.inbox);

      const response = await CreateFolder(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
      expect(response.ResponseMessages.items[0].Folders.items[0].ParentFolderId?.Id).to.be.equal(parentInboxEWSId);
      expect(response.ResponseMessages.items[0].Folders.items[0].FolderId?.Id).to.be.equal(folderEWSId);
      expect(response.ResponseMessages.items[0].Folders.items[0].DisplayName).to.be.equal(item.name);
    });

    it('passes for a parent root folder and one inner folder', async function () {
      const soapRequest = generateSoapRequest(Parent.root);
      label.parentFolderId = Parent.root;
      const parentInboxEWSId = getEWSId(Parent.root);

      const response = await CreateFolder(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
      expect(response.ResponseMessages.items[0].Folders.items[0].ParentFolderId?.Id).to.be.equal(parentInboxEWSId);
      expect(response.ResponseMessages.items[0].Folders.items[0].FolderId?.Id).to.be.equal(folderEWSId);
      expect(response.ResponseMessages.items[0].Folders.items[0].DisplayName).to.be.equal(item.name);
    });

    it('returns an error for a parent unknown folder', async function () {
      const soapRequest = generateSoapRequest(Parent.unknown);

      const response = await CreateFolder(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_PARENT_FOLDER_NOT_FOUND);
    });

    it('throws an error when the soap request has folder without a DisplayName or FolderId', async function () {
      const soapRequest = generateSoapRequest(Parent.inbox, true);

      const expectedError = new Error('Folder name is not set for ' + util.inspect({}, false, 5));

      await expect(CreateFolder(soapRequest, context.request)).to.be.rejectedWith(expectedError);
    });
  });

  describe('mail.DeleteFolder', function () {
    enum FOLDER_ID {
      EXISTED = '12345678ejhcdejch',
      NON_EXISTED = 'sfjrdikfvdkjvhc'
    }

    function generateSoapRequest(deleteType: DisposalType, folders = false): DeleteFolderType {
      const soapRequest = new DeleteFolderType();
      soapRequest.DeleteType = deleteType;

      if (folders) {
        soapRequest.FolderIds = new NonEmptyArrayOfBaseFolderIdsType([
          new FolderIdType(getEWSId(FOLDER_ID.EXISTED, delegatorMailboxId)), 
          new FolderIdType(getEWSId(FOLDER_ID.NON_EXISTED))
        ]);
      } else {
        const distinguishedFolders: DistinguishedFolderIdType[] = [];

        generateTestLabels().forEach(pimLabel => {
          const distFolderId = getDistinguishedFolderId(pimLabel);
          if (!distFolderId) return;
          const distinguishedFolder = new DistinguishedFolderIdType();
          distinguishedFolder.Id = distFolderId;
          distinguishedFolders.push(distinguishedFolder)
        });

        soapRequest.FolderIds = new NonEmptyArrayOfBaseFolderIdsType([
          new FolderIdType(getEWSId(FOLDER_ID.EXISTED)), 
          ...distinguishedFolders]);
      }
      return soapRequest;
    }

    let labelManagerStub: sinon.SinonStubbedInstance<KeepPimLabelManager>;
    
    const delegatorMailboxId = 'delegator@test.com';
    const expressContextStub = getContext(testUser, 'password');

    beforeEach(() => {
      labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      stubLabelsManager(labelManagerStub);
      const labels = generateTestLabels();
      labels.push(PimItemFactory.newPimLabel({ FolderId: FOLDER_ID.EXISTED }, PimItemFormat.DOCUMENT));
      labelManagerStub.getLabels.resolves(labels);
      labelManagerStub.deleteLabel.resolves({} as KeepDeleteLabelResults);
    });
    
    afterEach(() => {
      sinon.restore(); // Reset stubs
      resetUserMailBox(testUser); // Reset local user cache
    });

    it('passes when one folder is delegated and the other does not exist', async function () {
      const soapRequest = generateSoapRequest(DisposalType.HARD_DELETE, true);
      const userInfo = UserContext.getUserInfo(expressContextStub.request);

      const response = await DeleteFolder(soapRequest, expressContextStub.request);
      const [itemExisted, itemNonExisted] = response.ResponseMessages.items;

      expect(labelManagerStub.deleteLabel.calledWith(userInfo, FOLDER_ID.EXISTED, undefined, delegatorMailboxId)).to.be.true();
      expect(itemExisted.ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(itemExisted.ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);

      expect(itemNonExisted.ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(itemNonExisted.ResponseCode).to.be.equal(ResponseCodeType.ERROR_FOLDER_NOT_FOUND);

    });

    it('passes when one is a distinguished and another is not', async function () {
      const soapRequest = generateSoapRequest(DisposalType.SOFT_DELETE);

      const response = await DeleteFolder(soapRequest, expressContextStub.request);
      const [itemExisted] = response.ResponseMessages.items;
      const distinguishedItems = response.ResponseMessages.items.slice(1);

      expect(itemExisted.ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(itemExisted.ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);

      distinguishedItems.forEach(distinguishedItem => {
        expect(distinguishedItem.ResponseClass).to.be.equal(ResponseClassType.ERROR);
        expect(distinguishedItem.ResponseCode).to.be.equal(ResponseCodeType.ERROR_DELETE_DISTINGUISHED_FOLDER);
      });
    });
  });

  describe('mail.UpdateFolder', function () {
    function generateSoapRequest(delegated: boolean, newDisplayName?: string): UpdateFolderType {
      const soapRequest = new UpdateFolderType();

      soapRequest.FolderChanges = new NonEmptyArrayOfFolderChangesType();

      const folderChange = new FolderChangeType();
      const folderId = delegated ? getEWSId(FOLDER_ID.EXISTED, delegatorMailboxId) : getEWSId(FOLDER_ID.EXISTED);
      folderChange.FolderId = new FolderIdType(folderId);
      soapRequest.FolderChanges.FolderChange.push(folderChange);

      if (newDisplayName) {
        const pathToUnindexedFieldType = new PathToUnindexedFieldType();
        pathToUnindexedFieldType.FieldURI = UnindexedFieldURIType.FOLDER_DISPLAY_NAME
        const setFolderFieldType = new SetFolderFieldType();
        const folder = new FolderType();
        folder.DisplayName = newDisplayName;
        setFolderFieldType.Folder = folder;
        setFolderFieldType.FieldURI = pathToUnindexedFieldType;
        folderChange.Updates = new NonEmptyArrayOfFolderChangeDescriptionsType([setFolderFieldType]);
      } else {
        folderChange.Updates = new NonEmptyArrayOfFolderChangeDescriptionsType([]);
        const folderChangeNonExisted = new FolderChangeType();
        folderChangeNonExisted.FolderId = new FolderIdType(getEWSId(FOLDER_ID.NON_EXISTED));
        soapRequest.FolderChanges.FolderChange.push(folderChangeNonExisted);
      }

      return soapRequest;
    }

    const NEW_DISPLAY_NAME = 'new-display-name';

    enum FOLDER_ID {
      EXISTED = '12345678ejhcdejch',
      NON_EXISTED = 'sfjrdikfvdkjvhc'
    }

    const LABEL_EXISTED = PimItemFactory.newPimLabel({ FolderId: FOLDER_ID.EXISTED }, PimItemFormat.DOCUMENT);

    let labelManagerStub: sinon.SinonStubbedInstance<KeepPimLabelManager>;

    const expressContextStub = getContext(testUser, 'password');
    const userInfo = UserContext.getUserInfo(expressContextStub.request);
    const delegatorMailboxId = 'delegator@test.com';

    beforeEach(() => {
      labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      stubLabelsManager(labelManagerStub);
      const labels = generateTestLabels();
      labels.push(LABEL_EXISTED);
      labelManagerStub.getLabels.resolves(labels);
    });
    afterEach(() => {
      sinon.restore(); // Reset stubs
      resetUserMailBox(testUser); // Reset local user cache
    });

    it('passes a couple of folders where one exists and another does not', async function () {
      /* Pass a couple of folders where one exists and another does not;
      * verify ERROR_FOLDER_NOT_FOUND as the ResponseCode and ERROR ResponseClass for the one that does not exist
      * and success for the other that does exist.
      */
      const soapRequest = generateSoapRequest(false);

      const response = await UpdateFolder(soapRequest, expressContextStub.request);
      const [itemExisted, itemNonExisted] = response.ResponseMessages.items;

      expect(itemExisted.ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(itemExisted.ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
      expect(itemNonExisted.ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(itemNonExisted.ResponseCode).to.be.equal(ResponseCodeType.ERROR_FOLDER_NOT_FOUND);
    });

    it('passes a delegated folder with a FolderChange with a SetFolderFieldType for UnindexedFieldURIType.FOLDER_DISPLAY_NAME', async function () {
      /*
      * Pass a folder with a FolderChange with a SetFolderFieldType for UnindexedFieldURIType.FOLDER_DISPLAY_NAME;
      * verify KeepPimLabelManager updateLabel is called with a label having the new displayName.
      */
      const soapRequest = generateSoapRequest(true, NEW_DISPLAY_NAME);
      labelManagerStub.updateLabel.resolves({} as KeepPimBaseResults);

      await UpdateFolder(soapRequest, expressContextStub.request);

      expect(labelManagerStub.updateLabel.calledWith(userInfo, LABEL_EXISTED, delegatorMailboxId)).to.be.true();
      expect(LABEL_EXISTED.displayName).to.equal(NEW_DISPLAY_NAME);
      expect(labelManagerStub.updateLabel.calledWith(userInfo, LABEL_EXISTED)).to.be.true();
    });

    it('stubs KeepPimLabelManager updateLabel to throw an error', async function () {
      /*
      * Stub KeepPimLabelManager updateLabel to throw an error;
      * verify ResponseCode of ERROR_FOLDER_SAVE_PROPERTY_ERROR and ResponseClass of ERROR.
      */
      const soapRequest = generateSoapRequest(false, NEW_DISPLAY_NAME);
      labelManagerStub.updateLabel.rejects();

      const response = await UpdateFolder(soapRequest, expressContextStub.request);
      const [itemWithError] = response.ResponseMessages.items;

      expect(itemWithError.ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(itemWithError.ResponseCode).to.be.equal(ResponseCodeType.ERROR_FOLDER_SAVE_PROPERTY_ERROR);
    })
  });

  describe('mail.MoveFolder', function () {
    /*
    Add unit tests to mail.test.ts file to cover the MoveFolder function in mail.ts.
  These tests will need to stub KeepPimLabelManager getLabels to return a list of labels(folders).
  The labels will need to be hierarchical for various tests.
  The soap request MoveFolderType has an array of FolderIds to move and a ToFolderId under which to move the folders.
  You will need to stub KeepPimLabelManager moveLabel.
  We're not currently validating the move response other failure if an exception is thrown.
  */
    enum FOLDER_ID {
      ROOT = 'root',
      LEAF_11 = 'leaf-11',
      LEAF_12 = 'leaf-12',
      LEAF_121 = 'leaf-121',
      FOLDER_2 = 'folder-2',
      LEAF_21 = 'leaf-21',
      NON_EXISTED = 'non-existed'
    }

    function generateSoapRequest(
      sourceFolderIds: FOLDER_ID[], 
      targetFolderId: FOLDER_ID, 
      mailboxId?: string,
      sameFolder = true
      ): MoveFolderType {
      const soapRequest = new MoveFolderType();

      soapRequest.FolderIds = new NonEmptyArrayOfBaseFolderIdsType(
        sourceFolderIds.map(folderId => {
          let folderEWSId: string;
          if (!sameFolder) folderEWSId = getEWSId(folderId, delegateeMailboxId);
          else folderEWSId = getEWSId(folderId, mailboxId);
          return new FolderIdType(folderEWSId);
        })
      );

      soapRequest.ToFolderId = new TargetFolderIdType();

      if (targetFolderId === FOLDER_ID.ROOT) {
        soapRequest.ToFolderId.DistinguishedFolderId = new DistinguishedFolderIdType();
        soapRequest.ToFolderId.DistinguishedFolderId.Id = DistinguishedFolderIdNameType.ROOT;
      } else {
        const targetFolderEWSId = mailboxId ? getEWSId(targetFolderId, mailboxId) : getEWSId(targetFolderId);
        soapRequest.ToFolderId.FolderId = new FolderIdType(targetFolderEWSId);
      }
      
      return soapRequest;
    }

    let labelManagerStub: sinon.SinonStubbedInstance<KeepPimLabelManager>;

    const expressContextStub = getContext(testUser, 'password');
    const userInfo = UserContext.getUserInfo(expressContextStub.request);
    const delegatorMailboxId = 'delegator@test.com';
    const delegateeMailboxId = 'delegatee@test.com';

    beforeEach(() => {
      labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      stubLabelsManager(labelManagerStub);
      const labels = [
        //PimItemFactory.newPimLabel({ FolderId: FOLDER_ID.ROOT_1 }),
        PimItemFactory.newPimLabel({ ParentId: FOLDER_ID.ROOT, FolderId: FOLDER_ID.LEAF_11 }),
        PimItemFactory.newPimLabel({ ParentId: FOLDER_ID.ROOT, FolderId: FOLDER_ID.LEAF_12 }),
        PimItemFactory.newPimLabel({ ParentId: FOLDER_ID.LEAF_12, FolderId: FOLDER_ID.LEAF_121 }),
        PimItemFactory.newPimLabel({ FolderId: FOLDER_ID.FOLDER_2 }),
        PimItemFactory.newPimLabel({ ParentId: FOLDER_ID.FOLDER_2, FolderId: FOLDER_ID.LEAF_21 })
      ];
      labelManagerStub.getLabels.resolves(labels);
      labelManagerStub.moveLabel.resolves({} as KeepMoveLabelResults);
    });

    afterEach(() => {
      sinon.restore(); // Reset stubs
      resetUserMailBox(testUser); // Reset local user cache
    });

    it('passes when the folder with child is moved to the root', async function () {
      const soapRequest = generateSoapRequest([FOLDER_ID.FOLDER_2], FOLDER_ID.ROOT);
      const CHILD_COUNT = 1;
      const rootID = base64Encode(`id-root-${userInfo?.userId}`);

      const response = await MoveFolder(soapRequest, expressContextStub.request);
      const [item] = response.ResponseMessages.items;
      const responseFolder = item.Folders.items[0];

      // verify the response folder has a parent folder is the root folder
      expect(responseFolder.ParentFolderId?.Id).to.equal(rootID);
      expect(responseFolder.ChildFolderCount).to.equal(CHILD_COUNT);
    });

    it('passes when the delegated folder is moved to another delegated folder', async function () {
      const soapRequest = generateSoapRequest([FOLDER_ID.LEAF_11], FOLDER_ID.LEAF_21, delegatorMailboxId);
      const parentEWSId = getEWSId(FOLDER_ID.LEAF_21, delegatorMailboxId);

      const response = await MoveFolder(soapRequest, expressContextStub.request);
      const [item] = response.ResponseMessages.items;
      const responseFolder = item.Folders.items[0];

      // verify the response folder has a parent folder is the target folderid
      expect(labelManagerStub.moveLabel.calledWith(userInfo, FOLDER_ID.LEAF_21, [FOLDER_ID.LEAF_11], delegatorMailboxId)).to.true();
      expect(responseFolder.ParentFolderId?.Id).to.equal(parentEWSId);
    });

    it('passes a couple of folders where one exists and another does not', async function () {
      const soapRequest = generateSoapRequest([FOLDER_ID.LEAF_21, FOLDER_ID.NON_EXISTED], FOLDER_ID.LEAF_11);

      await MoveFolder(soapRequest, expressContextStub.request);

      const response = await MoveFolder(soapRequest, expressContextStub.request);
      const [item, itemNonExisted] = response.ResponseMessages.items;

      // verify ERROR_FOLDER_NOT_FOUND as the ResponseCode and ERROR ResponseClass for the one that does not exist and success for the other that does exist.
      expect(item.ResponseClass).to.equal(ResponseClassType.SUCCESS);
      expect(item.ResponseCode).to.equal(ResponseCodeType.NO_ERROR);

      expect(itemNonExisted.ResponseClass).to.equal(ResponseClassType.ERROR);
      expect(itemNonExisted.ResponseCode).to.equal(ResponseCodeType.ERROR_FOLDER_NOT_FOUND);
    });

    it('passes a couple of folders to move to folder that does not exist', async function () {
      const soapRequest = generateSoapRequest([FOLDER_ID.LEAF_11, FOLDER_ID.LEAF_12], FOLDER_ID.NON_EXISTED);

      const response = await MoveFolder(soapRequest, expressContextStub.request);
      const [itemLeaf1, itemLeaf2] = response.ResponseMessages.items;

      // verify for each folder that does not exist  as the ResponseCode  and ERROR ResponseClass
      expect(itemLeaf1.ResponseClass).to.equal(ResponseClassType.ERROR);
      expect(itemLeaf1.ResponseCode).to.equal(ResponseCodeType.ERROR_TO_FOLDER_NOT_FOUND);
      expect(itemLeaf2.ResponseClass).to.equal(ResponseClassType.ERROR);
      expect(itemLeaf2.ResponseCode).to.equal(ResponseCodeType.ERROR_TO_FOLDER_NOT_FOUND);
    });

    it('stubs KeepPimLabelManager.moveLabel to throw an error', async function () {
      const soapRequest = generateSoapRequest([FOLDER_ID.LEAF_21], FOLDER_ID.ROOT);
      labelManagerStub.moveLabel.rejects();

      const response = await MoveFolder(soapRequest, expressContextStub.request);

      // verify the ResponseCode is ERROR_MOVE_COPY_FAILED and ERROR ResponseClass.
      const [itemWithError] = response.ResponseMessages.items;
      expect(itemWithError.ResponseClass).to.equal(ResponseClassType.ERROR);
      expect(itemWithError.ResponseCode).to.equal(ResponseCodeType.ERROR_MOVE_COPY_FAILED);
    });

    it('throws an error when a delegator folder move to a delegatee folder', async function() {
      const soapRequest = generateSoapRequest([FOLDER_ID.LEAF_21], FOLDER_ID.LEAF_11, delegatorMailboxId, false);
      let receivedError: Error;
      
      await MoveFolder(soapRequest, expressContextStub.request).catch(err => receivedError = err);
      const expectedMessage = `Folder cannot be moved from ${delegateeMailboxId} to ${delegatorMailboxId} mailbox`;

      expect(receivedError!.message).to.equal(expectedMessage);
    });
  });

  describe('mail.CreateItem', function() {
    /*
      From: LABS-2131
      Add unit tests to mail.test.ts file to cover the CreateItem function in mail.ts.
      These tests will need to stub KeepPimLabelManager getLabels to return a list of labels(folders).
      The soap request CreateItemType has a SaveFolderItemId for where to save the item once created an
      array of Items to create and a MessageDisposition for whether to save and/or send a message.
    */
      enum FOLDER_ID {
        ROOT_1 = 'root-1',
        LEAF_11 = 'leaf-11',
        LEAF_12 = 'leaf-12',
        LEAF_121 = 'leaf-121',
        ROOT_2 = 'root-2',
        LEAF_21 = 'leaf-21',
        NON_EXISTED = 'non-existed'
      }
      function generateSoapRequest(folderId: FOLDER_ID): CreateItemType {
        const soapRequest = new CreateItemType();
        const item = new MessageType();
        item.ItemId = new ItemIdType(FOLDER_ID.LEAF_11);
        soapRequest.Items = new NonEmptyArrayOfAllItemsType([item]);
        soapRequest.SavedItemFolderId = new TargetFolderIdType();
        soapRequest.SavedItemFolderId.FolderId = new FolderIdType(folderId);
        soapRequest.MessageDisposition = MessageDispositionType.SAVE_ONLY;
        return soapRequest;
      }
      let labelManagerStub: sinon.SinonStubbedInstance<KeepPimLabelManager>;
      let messageManagerStub: sinon.SinonStubbedInstance<EWSMessageManager>;
      const expressContextStub = getContext(testUser, 'password');
      beforeEach(() => {
        labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
        stubLabelsManager(labelManagerStub);
        labelManagerStub.getLabel.resolves(PimItemFactory.newPimLabel({ FolderId: FOLDER_ID.ROOT_1 }));
        
        messageManagerStub = createSinonStubInstance(EWSMessageManager);
        sinon.stub(EWSServiceManager, 'getInstanceFromItem').callsFake((item: MessageType): any => {
          return messageManagerStub;
        });
      });
      afterEach(() => {
        sinon.restore(); // Reset stubs
        resetUserMailBox(testUser); // Reset local user cache
      });
      it('passes a SavedItemFolderId that does not exist', async function() {
        const soapRequest = generateSoapRequest(FOLDER_ID.NON_EXISTED);
        labelManagerStub.getLabel.rejects({ status: 404 });

        const response = await CreateItem(soapRequest, expressContextStub.request);

        //verify in the response for each item a ResponseCode of ERROR_FOLDER_NOT_FOUND and ResponseClass of ERROR.
        expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR)
        expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_FOLDER_NOT_FOUND)
      });
      it('passes an item and a valid SavedItemFolderId matching an existing folder', async function() {
        const soapRequest = generateSoapRequest(FOLDER_ID.ROOT_1);
        messageManagerStub.createItem.resolves([new MessageType()]);

        const response = await CreateItem(soapRequest, expressContextStub.request);

        //verify the response has the new item and ResponseClass of SUCCESS
        expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
        expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
      });
      it('stubs to rejectsWith an error with a field having a status of 403', async function() {
        const soapRequest = generateSoapRequest(FOLDER_ID.ROOT_1);
        messageManagerStub.createItem.rejects({ status: 403 });

        const response = await CreateItem(soapRequest, expressContextStub.request);

        //verify a ResponseCode of ERROR_CREATE_ITEM_ACCESS_DENIED and ResponseClass of ERROR
        expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
        expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_CREATE_ITEM_ACCESS_DENIED);
      });
  });

  describe('mail.EmptyFolder', function () {
    /*
    From: LABS-2125
    Add unit tests to mail.test.ts file to cover the EmptyFolder function in mail.ts.
    These tests will need to stub KeepPimLabelManager getLabels to return a list of labels(folders). The labels will need to be hierarchical for various tests.
    The soap request EmptyFolderType has an array of FolderIds to empty, whether to include subfolders or just non folders and
    whether it's a soft or hard delete.
    You will need to stub KeepPimMessageManager getMailMessages and deleteMessage
    */
    enum FOLDER_ID {
      ROOT_1 = 'root-1',
      LEAF_11 = 'leaf-11',
      LEAF_111 = 'leaf-111',
      LEAF_12 = 'leaf-12',
      LEAF_121 = 'leaf-121',
      LEAF_122 = 'leaf-122',
      ROOT_2 = 'root-2',
      LEAF_21 = 'leaf-21',
      NON_EXISTED = 'non-existed'
    }
    const COUNT_OF_MESSAGES = 5;
  
    function generateSoapRequest(
      folderIds: FOLDER_ID[], 
      deleteType: DisposalType, 
      deleteSubFolders = false,
      mailboxId? : string
      ): EmptyFolderType {
      const soapRequest = new EmptyFolderType();
      soapRequest.FolderIds = new NonEmptyArrayOfBaseFolderIdsType(
        folderIds.map(folderId => {
          const folderEWSId = getEWSId(folderId, mailboxId);
          return new FolderIdType(folderEWSId);
        })
      );
      soapRequest.DeleteType = deleteType;
      soapRequest.DeleteSubFolders = deleteSubFolders;
      return soapRequest;
    }
  
    let labelManagerStub: sinon.SinonStubbedInstance<KeepPimLabelManager>;
    let messageManagerStub: sinon.SinonStubbedInstance<KeepPimMessageManager>;
  
    const expressContextStub = getContext(testUser, 'password');
    const userInfo = UserContext.getUserInfo(expressContextStub.request);
    const delegatorMailboxId = 'delegator@test.com';

    beforeEach(() => {
      labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      stubLabelsManager(labelManagerStub);
  
      messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
      stubMessageManager(messageManagerStub);
      // Stub KeepPimMessageManager getMailMessages to return {{COUNT_OF_MESSAGES}} of messages
      let skip = 0;
      messageManagerStub.getMailMessages.callsFake( (_userInfo: UserInfo, labelId?: string, docs?: boolean): Promise<PimMessage[]> => {
        const count = 3;
        if ((skip)>=COUNT_OF_MESSAGES) {
          skip = 0;
          return Promise.resolve([]);
        }
        const countReturned = Math.min(COUNT_OF_MESSAGES - skip, count);
        skip += count;
        return  Promise.resolve(Array(countReturned).fill(PimItemFactory.newPimMessage({})));
      });
      // messageManagerStub.getMailMessages.resolves(Array(COUNT_OF_MESSAGES).fill(PimItemFactory.newPimMessage({})));

      messageManagerStub.deleteMessage.resolves({});
  
      const labels = [
        PimItemFactory.newPimLabel({ FolderId: FOLDER_ID.ROOT_1 }),
        PimItemFactory.newPimLabel({ ParentId: FOLDER_ID.ROOT_1, FolderId: FOLDER_ID.LEAF_11 }),
        PimItemFactory.newPimLabel({ ParentId: FOLDER_ID.LEAF_11, FolderId: FOLDER_ID.LEAF_111 }),
        PimItemFactory.newPimLabel({ ParentId: FOLDER_ID.ROOT_1, FolderId: FOLDER_ID.LEAF_12 }),
        PimItemFactory.newPimLabel({ ParentId: FOLDER_ID.LEAF_12, FolderId: FOLDER_ID.LEAF_121 }),
        PimItemFactory.newPimLabel({ ParentId: FOLDER_ID.LEAF_12, FolderId: FOLDER_ID.LEAF_122 }),
        PimItemFactory.newPimLabel({ FolderId: FOLDER_ID.ROOT_2 }),
        PimItemFactory.newPimLabel({ ParentId: FOLDER_ID.ROOT_2, FolderId: FOLDER_ID.LEAF_21 })
      ];
      labelManagerStub.getLabels.resolves(labels);
      labelManagerStub.deleteLabel.resolves({} as KeepDeleteLabelResults);
    });
    afterEach(() => {
      sinon.restore(); // Reset stubs
      resetUserMailBox(testUser); // Reset local user cache
    });
  
    it('passes in a folder to empty', async function () {
      const soapRequest = generateSoapRequest([FOLDER_ID.LEAF_21], DisposalType.HARD_DELETE);

      await EmptyFolder(soapRequest, expressContextStub.request);
  
      // verify KeepPimMessageManager.getMailMessages is called.
      expect(messageManagerStub.getMailMessages.called).to.true();
      // verify deleteMessage is called n times where n is the number {{COUNT_OF_MESSAGES}} of messages  returned from getMailMessages
      expect(messageManagerStub.deleteMessage.callCount).to.equal(COUNT_OF_MESSAGES);
    });
  
    it('passes in a delegated folder to empty that has subfolders with include subfolders true (HARD_DELETE)', async function () {
      const soapRequest = generateSoapRequest([FOLDER_ID.ROOT_1], DisposalType.HARD_DELETE, true, delegatorMailboxId);
  
      await EmptyFolder(soapRequest, expressContextStub.request);
  
      // verify KeepPimLabelManager deleteLabel is called for each subfolder
      expect(labelManagerStub.deleteLabel.calledWith(userInfo, FOLDER_ID.LEAF_11, undefined, delegatorMailboxId)).to.true();
      expect(labelManagerStub.deleteLabel.calledWith(userInfo, FOLDER_ID.LEAF_111, undefined, delegatorMailboxId)).to.true();
      expect(labelManagerStub.deleteLabel.calledWith(userInfo, FOLDER_ID.LEAF_12, undefined, delegatorMailboxId)).to.true();
      expect(labelManagerStub.deleteLabel.calledWith(userInfo, FOLDER_ID.LEAF_121, undefined, delegatorMailboxId)).to.true();
      expect(labelManagerStub.deleteLabel.calledWith(userInfo, FOLDER_ID.LEAF_122, undefined, delegatorMailboxId)).to.true();
  
      //Verify deleteMessage called COUNT_OF_MESSAGES * COUNT OF ALL FOLDERS - because it is called for all subfolders.
      const COUNT_OF_SUBFOLDERS = labelManagerStub.deleteLabel.callCount;
      expect(messageManagerStub.deleteMessage.callCount).to.equal(COUNT_OF_MESSAGES * (1 + COUNT_OF_SUBFOLDERS));
  
      // verify KeepPimLabelManager deleteMessage is NOT called for subfolder of ROOT_2 folder
      expect(labelManagerStub.deleteLabel.calledWith(userInfo, FOLDER_ID.LEAF_21)).to.false();
    });
  
    it('pass in a folder to empty that has subfolders with include subfolders true (SOFT_DELETE)', async function () {
      const soapRequest = generateSoapRequest([FOLDER_ID.ROOT_1], DisposalType.SOFT_DELETE, true);
  
      await EmptyFolder(soapRequest, expressContextStub.request);
  
      // verify KeepPimLabelManager deleteLabel is called for each subfolder
      expect(labelManagerStub.deleteLabel.calledWith(userInfo, FOLDER_ID.LEAF_11)).to.true();
      expect(labelManagerStub.deleteLabel.calledWith(userInfo, FOLDER_ID.LEAF_111)).to.true();
      expect(labelManagerStub.deleteLabel.calledWith(userInfo, FOLDER_ID.LEAF_12)).to.true();
      expect(labelManagerStub.deleteLabel.calledWith(userInfo, FOLDER_ID.LEAF_121)).to.true();
      expect(labelManagerStub.deleteLabel.calledWith(userInfo, FOLDER_ID.LEAF_122)).to.true();
  
      //Verify deleteMessage called COUNT_OF_MESSAGES * COUNT OF ALL FOLDERS - because it is called for all subfolders.
      const COUNT_OF_SUBFOLDERS = labelManagerStub.deleteLabel.callCount;
      expect(messageManagerStub.deleteMessage.callCount).to.equal(COUNT_OF_MESSAGES * (1 + COUNT_OF_SUBFOLDERS));
  
      // verify KeepPimLabelManager deleteMessage is NOT called for subfolder of ROOT_2 folder
      expect(labelManagerStub.deleteLabel.calledWith(userInfo, FOLDER_ID.LEAF_21)).to.false();
    });
  
    it('pass in a folder to empty that has subfolders with include subfolders false', async function () {
      const soapRequest = generateSoapRequest([FOLDER_ID.ROOT_1], DisposalType.SOFT_DELETE);
  
      await EmptyFolder(soapRequest, expressContextStub.request);
  
      // verify KeepPimLabelManager deleteLabel is not called.
      expect(labelManagerStub.deleteLabel.notCalled).to.true();
  
      //Verify deleteMessage called COUNT_OF_MESSAGES times
      expect(messageManagerStub.deleteMessage.callCount).to.equal(COUNT_OF_MESSAGES);
    });
  
    it('pass a couple of folders where one exists and another does not', async function () {
      const soapRequest = generateSoapRequest([FOLDER_ID.LEAF_11, FOLDER_ID.NON_EXISTED], DisposalType.SOFT_DELETE, true);
  
      const response = await EmptyFolder(soapRequest, expressContextStub.request);
      const [item, itemNonExisted] = response.ResponseMessages.items;
  
      // verify ERROR_FOLDER_NOT_FOUND as the ResponseCode and ERROR ResponseClass for the one that does not exist and success for the other that does exist.
      expect(item.ResponseClass).to.equal(ResponseClassType.SUCCESS);
      expect(item.ResponseCode).to.equal(ResponseCodeType.NO_ERROR);
  
      expect(itemNonExisted.ResponseClass).to.equal(ResponseClassType.ERROR);
      expect(itemNonExisted.ResponseCode).to.equal(ResponseCodeType.ERROR_FOLDER_NOT_FOUND);
    });
  
    it('stub getMailMessages to throw an exception for a folder', async function () {
      const soapRequest = generateSoapRequest([FOLDER_ID.LEAF_21], DisposalType.HARD_DELETE);
      messageManagerStub.getMailMessages.rejects();
  
      const response = await EmptyFolder(soapRequest, expressContextStub.request);
  
      // verify ERROR_CANNOT_EMPTY_FOLDER ResponseCode and ERROR ResponseClass
      const [itemWithError] = response.ResponseMessages.items;
      expect(itemWithError.ResponseClass).to.equal(ResponseClassType.ERROR);
      expect(itemWithError.ResponseCode).to.equal(ResponseCodeType.ERROR_CANNOT_EMPTY_FOLDER);
    });
  });

  describe('mail.CreateFolderPath', function () {
    function generateSoapRequest(
      distinguishedFolderId: DistinguishedFolderIdNameType,
      folders: BaseFolderType[],
      folderId?: string
    ): CreateFolderPathType {
      const soapRequest = new CreateFolderPathType();
      soapRequest.ParentFolderId = new TargetFolderIdType();

      if (folderId) {
        soapRequest.ParentFolderId.FolderId = new FolderIdType(getEWSId(folderId));
      } else {
        soapRequest.ParentFolderId.DistinguishedFolderId = new DistinguishedFolderIdType();
        soapRequest.ParentFolderId.DistinguishedFolderId.Id = distinguishedFolderId;
      }
      soapRequest.RelativeFolderPath = new NonEmptyArrayOfFoldersType(folders);

      return soapRequest;
    }

    let labelManagerStub: sinon.SinonStubbedInstance<KeepPimLabelManager>;

    const FOLDER_1 = 'folder1';
    const folder1 = { FolderId: new FolderIdType(getEWSId(FOLDER_1)), DisplayName: 'Folder1' };
    const folder1Label = PimItemFactory.newPimLabel({
      FolderId: FOLDER_1,
      DisplayName: folder1.DisplayName
    });

    const FOLDER_2 = 'folder2';
    const folder2 = { DisplayName: 'Folder2', ParentId: folder1.FolderId };
    const folder2Label = PimItemFactory.newPimLabel({
      DisplayName: folder2.DisplayName,
      FolderId: FOLDER_2,
      ParentId: FOLDER_1
    });

    const context = getContext(testUser, 'password');
    const userInfo = UserContext.getUserInfo(context.request);
    const parentRootId = base64Encode(`id-root-${userInfo.userId}`);

    beforeEach(() => {
      labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      labelManagerStub.getLabels.resolves([generateTestInboxLabel()]);
      stubLabelsManager(labelManagerStub);
    });

    afterEach(() => {
      sinon.restore(); // Reset stubs
      resetUserMailBox(testUser); // Reset local user cache
    });

    it('passes for a delegated parent inbox folder and one inner folder with DisplayName', async function () {
      const delegatorMailboxId = 'delegator@test.com';
      const delegatedId = 'delegatedId';
      const delegatedEWSId = getEWSId(delegatedId, delegatorMailboxId);
      const delegatedName = 'DelegatedFolder';
      const delegatedFolder = { DisplayName: delegatedName };
      const parentEWSId = getEWSId(DistinguishedFolderIdNameType.INBOX, delegatorMailboxId);

      const soapRequest = generateSoapRequest(DistinguishedFolderIdNameType.INBOX, [delegatedFolder]);
      soapRequest.ParentFolderId.DistinguishedFolderId!.Mailbox = new EmailAddressType();
      soapRequest.ParentFolderId.DistinguishedFolderId!.Mailbox.EmailAddress = delegatorMailboxId;

      const pimLabelIn = pimLabelFromItem(soapRequest.RelativeFolderPath.items[0], {});
      const root = await getLabelByTarget(soapRequest.ParentFolderId, userInfo);
      pimLabelIn.parentFolderId = root?.folderId;

      labelManagerStub.createLabel.resolves(
        PimItemFactory.newPimLabel({
          FolderId: delegatedId,
          DisplayName: delegatedName,
          ParentId: DistinguishedFolderIdNameType.INBOX
        })
      );

      const response = await CreateFolderPath(soapRequest, context.request);

      expect(labelManagerStub.createLabel.calledWith(userInfo, pimLabelIn, undefined, delegatorMailboxId)).to.be.equal(true);
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
      expect(response.ResponseMessages.items[0].Folders.items[0].FolderId?.Id).to.be.equal(delegatedEWSId);
      expect(response.ResponseMessages.items[0].Folders.items[0].DisplayName).to.be.equal(delegatedFolder.DisplayName);
      expect(response.ResponseMessages.items[0].Folders.items[0].ParentFolderId?.Id).to.be.equal(parentEWSId);
    });

    it('passes for a root parent folder and one inner folder with FolderId and DisplayName', async function () {
      const soapRequest = generateSoapRequest(DistinguishedFolderIdNameType.ROOT, [folder1]);
      labelManagerStub.createLabel.resolves(folder1Label);

      const response = await CreateFolderPath(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
      expect(response.ResponseMessages.items[0].Folders.items[0].FolderId?.Id).to.be.equal(folder1.FolderId.Id);
      expect(response.ResponseMessages.items[0].Folders.items[0].DisplayName).to.be.equal(folder1.DisplayName);
      expect(response.ResponseMessages.items[0].Folders.items[0].ParentFolderId?.Id).to.be.equal(parentRootId);
    });

    it('returns an error when parent folder is unknown', async function () {
      const soapRequest = generateSoapRequest(DistinguishedFolderIdNameType.INBOX, [folder1], 'unknown');

      const response = await CreateFolderPath(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_PARENT_FOLDER_NOT_FOUND);
    });

    it('returns an error for inbox parent folder and one inner folder without DisplayName', async function () {
      const invalidFolder = { FolderId: new FolderIdType(getEWSId('invalid'))};
      const soapRequest = generateSoapRequest(DistinguishedFolderIdNameType.INBOX, [invalidFolder]);

      const response = await CreateFolderPath(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_INVALID_FOLDER_TYPE_FOR_OPERATION);
    });

    it('passes for hierarchy two folders deep in the default root folder', async function() {
      const soapRequest = generateSoapRequest(DistinguishedFolderIdNameType.ROOT, [folder1, folder2]);

      labelManagerStub.createLabel.callsFake((uInfo: UserInfo, pimLabelIn: any): any => {
        return pimLabelIn.parentFolderId === FOLDER_1 ? folder2Label : folder1Label;
      });

      const response = await CreateFolderPath(soapRequest, context.request);
      const [receivedFolder1, receivedFolder2] = response.ResponseMessages.items;

      expect(receivedFolder1.ResponseClass).to.equal(ResponseClassType.SUCCESS);
      expect(receivedFolder1.ResponseCode).to.equal(ResponseCodeType.NO_ERROR);
      expect(receivedFolder1.Folders.items[0].DisplayName).to.equal(folder1.DisplayName);
      expect(receivedFolder1.Folders.items[0].FolderId?.Id).to.equal(folder1.FolderId.Id);
      expect(receivedFolder1.Folders.items[0].ParentFolderId?.Id).to.equal(parentRootId);

      expect(receivedFolder2.ResponseClass).to.equal(ResponseClassType.SUCCESS);
      expect(receivedFolder2.ResponseCode).to.equal(ResponseCodeType.NO_ERROR);
      expect(receivedFolder2.Folders.items[0].DisplayName).to.equal(folder2.DisplayName);
      expect(receivedFolder2.Folders.items[0].FolderId?.Id).to.equal(getEWSId(FOLDER_2));
      expect(receivedFolder2.Folders.items[0].ParentFolderId?.Id).to.equal(folder1.FolderId.Id);
    });

    it('returns an error when createLabel is stubbed to reject', async function () {
      const soapRequest = generateSoapRequest(DistinguishedFolderIdNameType.INBOX, [folder1]);
      labelManagerStub.createLabel.rejects();

      const response = await CreateFolderPath(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_FOLDER_SAVE_FAILED);
    });
  });

  describe('mailTest.FindFolder', function () {
    /*
    From: LABS-2123
    Add unit tests to mail.test.ts file to cover the FindFolder function in mail.ts.
    These tests will need to stub KeepPimLabelManager getLabels to return a list of labels(folders).
    The labels will need to be hierarchical for various tests.
    The soap request FindFolderType has an array of ParentFolderIds and has a Traversal (Deep, Shallow or SoftDeleted)
    */

    enum FOLDER_ID {
      ROOT_1 = 'root-1',
      LEAF_11 = 'leaf-11',
      LEAF_12 = 'leaf-12',
      LEAF_121 = 'leaf-121',
      ROOT_2 = 'root-2',
      LEAF_21 = 'leaf-21',
      NON_EXISTED = 'non-existed'
    }
    const COUNT_PARENT_FOLDERS_IN_ROOT_1 = 2;
    const COUNT_FOLDERS_IN_ROOT_1 = 3;

    let labelManagerStub: sinon.SinonStubbedInstance<KeepPimLabelManager>;

    const context = getContext(testUser, 'password');

    function generateSoapRequest(sourceFolderIds: FOLDER_ID[], traversal: FolderQueryTraversalType, mailboxId?: string): FindFolderType {
      const soapRequest = new FindFolderType();
      soapRequest.Traversal = traversal;
      soapRequest.FolderShape = new FolderResponseShapeType();
      soapRequest.FolderShape.BaseShape = DefaultShapeNamesType.DEFAULT;
      soapRequest.ParentFolderIds = new NonEmptyArrayOfBaseFolderIdsType(
        sourceFolderIds.map(folderId => {
          const folderEWSId = getEWSId(folderId, mailboxId);
          return new FolderIdType(folderEWSId);
        })
      );
      return soapRequest;
    }

    function addFolderClassTypeRestriction(soapRequest: FindFolderType, constant: FolderClassType): FindFolderType {
      soapRequest.Restriction = new RestrictionType();
      soapRequest.Restriction.Or = new OrType();
      const restriction = new ContainsExpressionType();
      restriction.FieldURI = new PathToUnindexedFieldType();
      restriction.FieldURI.FieldURI = UnindexedFieldURIType.FOLDER_FOLDER_CLASS;
      restriction.ContainmentMode = ContainmentModeType.FULL_STRING;
      restriction.ContainmentComparison = ContainmentComparisonType.EXACT;
      restriction.Constant = new ConstantValueType();
      restriction.Constant.Value = constant;
      soapRequest.Restriction.Or.items[0] = restriction;
      return soapRequest;
    }

    const delegatorMailboxId = 'delegator@test.com';

    beforeEach(() => {
      labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      stubLabelsManager(labelManagerStub);
      const labels = [
        PimItemFactory.newPimLabel({ FolderId: FOLDER_ID.ROOT_1 }),
        PimItemFactory.newPimLabel({ ParentId: FOLDER_ID.ROOT_1, FolderId: FOLDER_ID.LEAF_11 }),
        PimItemFactory.newPimLabel({ ParentId: FOLDER_ID.ROOT_1, FolderId: FOLDER_ID.LEAF_12 }),
        PimItemFactory.newPimLabel({ ParentId: FOLDER_ID.LEAF_12, FolderId: FOLDER_ID.LEAF_121, Type: PimLabelTypes.TASKS }),
        PimItemFactory.newPimLabel({ FolderId: FOLDER_ID.ROOT_2 }),
        PimItemFactory.newPimLabel({ ParentId: FOLDER_ID.ROOT_2, FolderId: FOLDER_ID.LEAF_21 })
      ];
      labelManagerStub.getLabels.resolves(labels);
    });

    afterEach(() => {
      sinon.restore(); // Reset stubs
      resetUserMailBox(testUser); // Reset local user cache
    });

    it('when pass in a set of delegated parent folders for shallow traversal', async function () {
      const soapRequest = generateSoapRequest([FOLDER_ID.ROOT_1], FolderQueryTraversalType.SHALLOW, delegatorMailboxId);
      const userInfo = UserContext.getUserInfo(context.request);
      const response = await FindFolder(soapRequest, context.request);

      // Verify direct children are returned for shallow traversal
      expect(labelManagerStub.getLabels.calledWith(userInfo, true, DEFAULT_LABEL_EXCLUSIONS, delegatorMailboxId)).to.be.equal(true);
      expect(response.ResponseMessages.items[0].RootFolder?.IndexedPagingOffset).to.be.equal(COUNT_PARENT_FOLDERS_IN_ROOT_1);
      expect(response.ResponseMessages.items[0].RootFolder?.TotalItemsInView).to.be.equal(COUNT_PARENT_FOLDERS_IN_ROOT_1);
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
    });

    it('when pass in a set of parent folders for deep traversal', async function () {
      const soapRequest = generateSoapRequest([FOLDER_ID.ROOT_1], FolderQueryTraversalType.DEEP, delegatorMailboxId);

      const response = await FindFolder(soapRequest, context.request);
      const receivedItems = response.ResponseMessages.items[0].RootFolder?.Folders?.items;
      expect(receivedItems).to.not.be.undefined(); 

      // Verify all levels are returned for Deep traversal 
      expect(response.ResponseMessages.items[0].RootFolder?.IndexedPagingOffset).to.be.equal(COUNT_FOLDERS_IN_ROOT_1);
      expect(response.ResponseMessages.items[0].RootFolder?.TotalItemsInView).to.be.equal(COUNT_FOLDERS_IN_ROOT_1);
      receivedItems!.forEach(receivedItem => {
        const [, receivedMailboxId] = getKeepIdPair(receivedItem.FolderId!.Id!);
        expect(receivedMailboxId).to.be.equal(delegatorMailboxId);
      });
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
    });

    it('when pass in a parent folder having no children', async function () {
      const soapRequest = generateSoapRequest([FOLDER_ID.LEAF_11], FolderQueryTraversalType.SHALLOW);

      const response = await FindFolder(soapRequest, context.request);

      // verify TotalItemsInView and IndexedPagingOffset
      expect(response.ResponseMessages.items[0].RootFolder?.IndexedPagingOffset).to.be.equal(0);
      expect(response.ResponseMessages.items[0].RootFolder?.TotalItemsInView).to.be.equal(0);
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
    });

    it('when pass in a parent folder that does not exist in folders returned from getLabels', async function () {
      const soapRequest = generateSoapRequest([FOLDER_ID.NON_EXISTED], FolderQueryTraversalType.SHALLOW);

      const response = await FindFolder(soapRequest, context.request);

      // verify ERROR_FOLDER_NOT_FOUND   
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_FOLDER_NOT_FOUND);
    });

    it('when child folders have a mix of FolderClass and soapRequest with a Restriction', async function () {
      const soapRequest = generateSoapRequest([FOLDER_ID.ROOT_1], FolderQueryTraversalType.DEEP);
      addFolderClassTypeRestriction(soapRequest, FolderClassType.MESSAGE);
      const COUNT_FOLDER_CLASS_TYPE_TASKS = 1;
      const COUNT_RESTRICTED_FOLDERS = COUNT_FOLDERS_IN_ROOT_1 - COUNT_FOLDER_CLASS_TYPE_TASKS;

      const response = await FindFolder(soapRequest, context.request);

      // Verify the matching child folders are returned 
      expect(response.ResponseMessages.items[0].RootFolder?.IndexedPagingOffset).to.be.equal(COUNT_RESTRICTED_FOLDERS);
      expect(response.ResponseMessages.items[0].RootFolder?.TotalItemsInView).to.be.equal(COUNT_RESTRICTED_FOLDERS);
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
    });
  });

  describe('mail.SyncFolderItems', function () {
    /*
    From: LABS-2128
    Add unit tests to mail.test.ts file to cover the SyncFolderItems function in mail.ts.
    These tests will need to stub KeepPimLabelManager getLabels to return a list of labels(folders).
    The soap request SyncFolderItemsType has a SyncFolderId for the id of the folder to synchronize and ItemShape among other fields.
    */
    enum FOLDER_ID {
      NON_EXISTED = 'non-existed'
    }

    function generateSoapRequest(syncFolderId: DistinguishedFolderIdNameType | FOLDER_ID, isInvalidSyncState?: boolean): SyncFolderItemsType {
      const soapRequest = new SyncFolderItemsType();
      soapRequest.SyncFolderId = new TargetFolderIdType();
      if (Object.values(FOLDER_ID).includes(syncFolderId as FOLDER_ID)) {
        soapRequest.SyncFolderId.FolderId = new FolderIdType(syncFolderId);
      } else {
        soapRequest.SyncFolderId.DistinguishedFolderId = new DistinguishedFolderIdType();
        soapRequest.SyncFolderId.DistinguishedFolderId.Id = syncFolderId as DistinguishedFolderIdNameType
      }
  
      soapRequest.ItemShape = new ItemResponseShapeType();
      soapRequest.ItemShape.IncludeMimeContent = false;

      if (isInvalidSyncState) {
        const nowDate = new Date();
        const futureDate = new Date(nowDate.getTime() + 1000 * 1 * 60 * 60 * 24).getTime(); // 24 hours forward
        const syncStateInfo = new SyncStateInfo(futureDate.toString(), 0);
        soapRequest.SyncState = syncStateInfo.getBase64SyncState();
      }

      return soapRequest;
    }

    let labelManagerStub: sinon.SinonStubbedInstance<KeepPimLabelManager>;

    const expressContextStub = getContext(testUser, 'password');

    beforeEach(() => {
      labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      stubLabelsManager(labelManagerStub);

      const labels = generateTestLabels();
      labelManagerStub.getLabels.resolves(labels);
    });

    afterEach(() => {
      sinon.restore(); // Reset stubs
      resetUserMailBox(testUser); // Reset local user cache
    });

    it('pass the root folder id for SyncFolderId', async function () {
      const soapRequest = generateSoapRequest(DistinguishedFolderIdNameType.ROOT);

      const response = await SyncFolderItems(soapRequest, expressContextStub.request);
      const [item] = response.ResponseMessages.items;

      // verify empty changes in response and IncludesLastItemInRange is true.
      expect(item?.Changes?.items?.length).to.equal(0);
      expect(item.IncludesLastItemInRange).to.be.true();
    });

    it('pass a folder that does not exist in KeepPimLabelManager getLabels stub', async function () {
      const soapRequest = generateSoapRequest(FOLDER_ID.NON_EXISTED);

      const response = await SyncFolderItems(soapRequest, expressContextStub.request);
      const [itemWithError] = response.ResponseMessages.items;

      // verify ResponseCode of ERROR_SYNC_FOLDER_NOT_FOUND and ResponseClass of ERROR.
      expect(itemWithError.ResponseClass).to.equal(ResponseClassType.ERROR);
      expect(itemWithError.ResponseCode).to.equal(ResponseCodeType.ERROR_SYNC_FOLDER_NOT_FOUND);
    });

    it('pass a syncState value > the current datetimestamp', async function () {
      const soapRequest = generateSoapRequest(DistinguishedFolderIdNameType.INBOX, true);

      const response = await SyncFolderItems(soapRequest, expressContextStub.request);
      const [itemWithError] = response.ResponseMessages.items;

      // verify ResponseCode of ERROR_INVALID_SYNC_STATE_DATA and ResponseClass of ERROR.
      expect(itemWithError.ResponseClass).to.equal(ResponseClassType.ERROR);
      expect(itemWithError.ResponseCode).to.equal(ResponseCodeType.ERROR_INVALID_SYNC_STATE_DATA);
    });

    it('pass a the INBOX folder id, DistinguishedFolderIdNameType.INBOX, and stub EWSMessageManager getItems to return an array of ItemTypes', async function () {
      const COUNT_OF_MESSAGE_ITEMS = 5
      const soapRequest = generateSoapRequest(DistinguishedFolderIdNameType.INBOX);

      const ewsMessageManagerStub = sinon.createStubInstance(EWSMessageManager);
      sinon.stub(EWSMessageManager, 'getInstance').callsFake((): any => {
        console.info('Creating stub for EWSMessageManager');
        return ewsMessageManagerStub;
      });

      // create five messageItems
      const messageItems = Array(COUNT_OF_MESSAGE_ITEMS).fill(undefined).map((_el, ind) => {
        const msgItem = new MessageType();
        msgItem.ItemId = new ItemIdType(`id-${ind}`);
        return msgItem
      });

      ewsMessageManagerStub.getItems.resolves(messageItems);

      const response = await SyncFolderItems(soapRequest, expressContextStub.request);
      const [item] = response.ResponseMessages.items;
      const changesItems = item.Changes.items;

      // verify the response has CreateItemType for each ItemType
      expect(changesItems.length).to.equal(COUNT_OF_MESSAGE_ITEMS);
      changesItems.forEach((changesItem) => {
        expect(changesItem instanceof SyncFolderItemsCreateType).to.be.true();
      })
    });

  });

  describe('mail.addChanges', function () {
    /*
      From: LABS-2130
      Add unit tests to mail.test.ts file to cover the addChanges function in mail.ts.
    */
    function generateSoapRequest(): SyncFolderItemsType {
      const soapRequest = new SyncFolderItemsType();
      soapRequest.SyncFolderId = new TargetFolderIdType();

      soapRequest.SyncFolderId.DistinguishedFolderId = new DistinguishedFolderIdType();
      soapRequest.SyncFolderId.DistinguishedFolderId.Id = DistinguishedFolderIdNameType.INBOX;
      soapRequest.SyncFolderId.DistinguishedFolderId.Mailbox = new EmailAddressType();
      soapRequest.SyncFolderId.DistinguishedFolderId.Mailbox.EmailAddress = delegatorMailboxId;

      soapRequest.ItemShape = new ItemResponseShapeType();
      soapRequest.ItemShape.IncludeMimeContent = false;
      soapRequest.SyncState = 'syncState';

      return soapRequest;
    }
    
    function generateItem(id: string): ItemType {
      const item = new MessageType();
      item.ItemId = new ItemIdType(getEWSId(id, delegatorMailboxId));
      return item;
    }

    let labelManagerStub: sinon.SinonStubbedInstance<KeepPimLabelManager>;
    let ewsMessageManagerStub: sinon.SinonStubbedInstance<EWSServiceManager>;

    const delegatorMailboxId = 'delegator@test.com';
    const COUNT_OF_MESSAGE_ITEMS = 3;
    const inboxEWSId = getEWSId('inbox-id', delegatorMailboxId);
    const context = getContext(testUser, 'password');

    beforeEach(() => {
      resetUserMailBox(testUser);
      labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      stubLabelsManager(labelManagerStub);

      const labels = generateTestLabels();
      labelManagerStub.getLabels.resolves(labels);

      ewsMessageManagerStub = sinon.createStubInstance(EWSMessageManager);
      sinon.stub(EWSMessageManager, 'getInstance').callsFake((): any => {
        return ewsMessageManagerStub;
      });
    });

    afterEach(() => {
      sinon.restore(); // Reset stubs
      resetUserMailBox(testUser); // Reset local user cache
    });

    it('passes for previousItems that are not in pimItems', async function () {
      const soapRequest = generateSoapRequest();
      const previousItems = Array(COUNT_OF_MESSAGE_ITEMS).fill(undefined).map((el, ind) => generateItem(`prev-id-${ind}`));
      const pimItems = Array(COUNT_OF_MESSAGE_ITEMS).fill(undefined).map((el, ind) => generateItem(`pim-id-${ind}`)); 
      USER_MAILBOX_DATA[testUser].LAST_SYNCFOLDER_ITEMS[inboxEWSId] = previousItems;
      ewsMessageManagerStub.getItems.resolves(pimItems);

      const response = await SyncFolderItems(soapRequest, context.request);

      const items = response.ResponseMessages.items[0].Changes.items;
      const checkedPreviousItems = [];
      const itemsDeleteType = items.filter(item => item instanceof SyncFolderItemsDeleteType);
      items.forEach(item => {
        if (item instanceof SyncFolderItemsDeleteType) {
          previousItems.forEach(previousItem => {
            if (previousItem.ItemId?.Id === item.ItemId.Id) checkedPreviousItems.push(previousItem);
          });
        }
      });

      //verify SyncFolderItemsDeleteType added for each in previousItems not in pimItems
      expect(itemsDeleteType.length).to.be.equal(checkedPreviousItems.length);
      expect(checkedPreviousItems.length).to.be.equal(previousItems.length);
    });

    it('passes for pimItems are not in previousItems and some matching those in previousItems', async function() {
      const soapRequest = generateSoapRequest();
      const COUNT_SAME_ITEMS = 1;
      const sameItem = generateItem('same-id');
      const previousItems = Array(COUNT_OF_MESSAGE_ITEMS).fill(undefined).map((el, ind) => generateItem(`prev-id-${ind}`));
      const pimItems = Array(COUNT_OF_MESSAGE_ITEMS).fill(undefined).map((el, ind) => generateItem(`pim-id-${ind}`));
      previousItems.push(sameItem);
      pimItems.push(sameItem);
      USER_MAILBOX_DATA[testUser].LAST_SYNCFOLDER_ITEMS[inboxEWSId] = previousItems;
      ewsMessageManagerStub.getItems.resolves(pimItems);

      const response = await SyncFolderItems(soapRequest, context.request);
      const items = response.ResponseMessages.items[0].Changes.items;

      const itemsCreateType = items.filter(item => item instanceof SyncFolderItemsCreateType);
      const checkedItems = items.filter(item => {
        if (item instanceof SyncFolderItemsCreateType) {
          if (item.Item?.ItemId !== sameItem.ItemId?.Id) return item;
        }
      });

      //verify SyncFolderItemsCreateType for each new item not in previousItems
      expect(itemsCreateType.length).to.be.equal(checkedItems.length);
      expect(checkedItems.length).to.be.equal(pimItems.length - COUNT_SAME_ITEMS);
    });
  });

  describe('mail.GetItem', function() {
    /*
    From: LABS-2129
    Add unit tests to mail.test.ts file to cover the GetItem function in mail.ts.
    The soap request GetItemType has an array of ItemIds and a ItemShape for the structure to return.
    - Pass An array with a mix of ItemIdType and an OccurrenceIdType; verify ResponseCode of ERROR_INVALID_ID and ResponseClass of ERROR for the OccurrenceIdTypes.
    - Stub EWSServiceManager.getItem to return undefined; verify ResponseCode of ERROR_ITEM_NOT_FOUND and ResponseClass of ERROR
    - Stub EWSServiceManager.getItem to return a MessageType; verify ResponseClass is SUCCESS and the item is returned in the msg.Items field of the response.
    */

    enum FOLDER_ID {
      ROOT_1 = 'root-1',
      LEAF_11 = 'leaf-11',
      LEAF_12 = 'leaf-12',
      LEAF_121 = 'leaf-121',
      ROOT_2 = 'root-2',
      LEAF_21 = 'leaf-21',
      NON_EXISTED = 'non-existed'
    }
    
    function generateSoapRequest(): GetItemType {
      const soapRequest = new GetItemType();
    
      soapRequest.ItemShape = new ItemResponseShapeType();
      soapRequest.ItemShape.BaseShape = DefaultShapeNamesType.DEFAULT;
      soapRequest.ItemShape.IncludeMimeContent = false;

      return soapRequest;
    }
    
    const expressContextStub = getContext(testUser, 'password')
    
    afterEach(() => {
      sinon.restore(); // Reset stubs
      resetUserMailBox(testUser); // Reset local user cache
    });
    
    it('passes an array with a mix of ItemIdType and an OccurrenceIdType', async function () {
      const soapRequest = generateSoapRequest();
      const item1 = new ItemIdType(FOLDER_ID.ROOT_1);
      const item2 = new OccurrenceItemIdType();
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([item1, item2]);   

      sinon.stub(EWSServiceManager, 'getItem').callsFake((): any => {
        return new ItemType();
      });

      const response = await GetItem(soapRequest, expressContextStub.request);
      
      expect(response.ResponseMessages.items[1]. ResponseClass).to.be.equal(ResponseClassType.ERROR)
      expect(response.ResponseMessages.items[1].ResponseCode).to.be.equal(ResponseCodeType.ERROR_INVALID_ID)
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS)
      expect(response.ResponseMessages.items.length).to.be.equal(2);
    })

    it('stubs EWSServiceManager.getItem to return undefined', async function () {
      const soapRequest = generateSoapRequest();
      const item = new ItemIdType(FOLDER_ID.ROOT_1)
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([item]);

      sinon.stub(EWSServiceManager, `getItem`).callsFake((): any => {
        return undefined;
      });

      const response = await GetItem(soapRequest, expressContextStub.request);

      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR)
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_ITEM_NOT_FOUND)
    })

    it('stubs EWSServiceManager.getItem to return a MessageType', async function () {
      const soapRequest = generateSoapRequest();
      const item = new ItemIdType(FOLDER_ID.ROOT_1)
      soapRequest.ItemIds = new NonEmptyArrayOfBaseItemIdsType([item]);
      soapRequest.ItemShape.IncludeMimeContent = true;

      sinon.stub(EWSServiceManager, `getItem`).callsFake((): any => {
        return new MessageType();
      });

      const response = await GetItem(soapRequest, expressContextStub.request);

      // verify ResponseClass is SUCCESS
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS)
      // verify the item is returned in the msg.Items field of the response
      expect(response.ResponseMessages.items.length).to.not.be.equal(0);
    })
  });

  describe('mail.MarkAllItemsAsRead', function() {
      
    enum FOLDER_ID {
      INBOX = 'inbox-id',
      NON_EXISTED = 'non-existed'
    }

  const COUNT_OF_MESSAGES = 5;

  function generateSoapRequest(folderId: FOLDER_ID, readFlag: boolean, suppressReadReceipt: boolean): MarkAllItemsAsReadType {
    const soapRequest = new MarkAllItemsAsReadType();
    soapRequest.ReadFlag = readFlag;
    soapRequest.SuppressReadReceipts = suppressReadReceipt;
    const item = new FolderIdType(getEWSId(folderId));
    soapRequest.FolderIds = new NonEmptyArrayOfBaseFolderIdsType([item]);
    return soapRequest;
  }

  let labelManagerStub: sinon.SinonStubbedInstance<KeepPimLabelManager>;
  let messageManagerStub: sinon.SinonStubbedInstance<KeepPimMessageManager>;
  const expressContextStub = getContext(testUser, 'password');

  beforeEach(() => {
    labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
    stubLabelsManager(labelManagerStub);
    const labels = generateTestLabels();
    labelManagerStub.getLabels.resolves(labels);

    messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
    stubMessageManager(messageManagerStub);
    messageManagerStub.getMailMessages.resolves(Array(COUNT_OF_MESSAGES).fill(PimItemFactory.newPimMessage({})));
  });
  
  afterEach(() => {
    sinon.restore(); // Reset stubs
    resetUserMailBox(testUser); // Reset local user cache
  });

  it('passes a FolderId that does not exist', async function() {
    const soapRequest = generateSoapRequest(FOLDER_ID.NON_EXISTED, true,true);
    const response = await MarkAllItemsAsRead(soapRequest, expressContextStub.request);

    //verify in the response for each item a ResponseCode of ERROR_FOLDER_NOT_FOUND and ResponseClass of ERROR.
    expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR)
    expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_FOLDER_NOT_FOUND)
  });

  it('passes an item and a valid FolderId matching an existing folder', async function() {
    messageManagerStub.getMailMessages.resolves(Array(COUNT_OF_MESSAGES).fill(PimItemFactory.newPimMessage({})));

    const soapRequest = generateSoapRequest(FOLDER_ID.INBOX, true, true);
    const response = await MarkAllItemsAsRead(soapRequest, expressContextStub.request);

    //verify the response has the new item and ResponseClass of SUCCESS
    expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
    expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
  });
  
  it('stub getMailMessages to throw an exception for a folder', async function () {
    messageManagerStub.getMailMessages.rejects();

    const soapRequest = generateSoapRequest(FOLDER_ID.INBOX, true, true);
    const response = await MarkAllItemsAsRead(soapRequest, expressContextStub.request);

    expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
    expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR);
  });
  });


  describe('mail.GetAttachment', function() {
    /*
  From LABS-2461
  Unit test to cover GetAttachment method in mail.ts  
  */

    enum ATTACHMENTS {
      CASE_1 = '##test1',
      CASE_2 = '123##',
      CASE_3 = '123##456',
      CASE_4 = 'testAttach'
    }
    
    function generateSoapRequest(): GetAttachmentType {
      const soapRequest = new GetAttachmentType();
      
      soapRequest.AttachmentIds = new NonEmptyArrayOfRequestAttachmentIdsType();

      return soapRequest;
    }
    
    const expressContextStub = getContext(testUser, 'password')
    
    afterEach(() => {
      sinon.restore(); // Reset stubs
      resetUserMailBox(testUser); // Reset local user cache
    });
    
    it('passes that can not find file attachment and parentId is null', async function () {
      const soapRequest = generateSoapRequest();

      const arrAttach = new RequestAttachmentIdType();
      arrAttach.Id = base64Encode(ATTACHMENTS.CASE_1);
      soapRequest.AttachmentIds.AttachmentId.push(arrAttach);

      const response = await GetAttachment(soapRequest, expressContextStub.request);

      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_CANNOT_FIND_FILE_ATTACHMENT)
      expect(response.ResponseMessages.items[0]. ResponseClass).to.be.equal(ResponseClassType.ERROR)
    })

    it('passes that can not find file attachment and attachment name is null', async function () {
      const soapRequest = generateSoapRequest();
      const arrAttach = new RequestAttachmentIdType();
      arrAttach.Id = base64Encode(ATTACHMENTS.CASE_2);
      soapRequest.AttachmentIds.AttachmentId.push(arrAttach);

      const response = await GetAttachment(soapRequest, expressContextStub.request);

      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_CANNOT_FIND_FILE_ATTACHMENT)
      expect(response.ResponseMessages.items[0]. ResponseClass).to.be.equal(ResponseClassType.ERROR)
    })

    it('passes that can not open file attachment', async function () {
      const soapRequest = generateSoapRequest();
      const arrAttach = new RequestAttachmentIdType();
      arrAttach.Id = base64Encode(ATTACHMENTS.CASE_3);
      soapRequest.AttachmentIds.AttachmentId.push(arrAttach);
      soapRequest.AttachmentShape = new AttachmentResponseShapeType();
      soapRequest.AttachmentShape.IncludeMimeContent = true;

      const attachInfo = new PimAttachmentInfo();
      const keepPimAttachMngrStub = sinon.createStubInstance(KeepPimAttachmentManager);
      sinon.stub(KeepPimAttachmentManager, 'getInstance').callsFake((): any => {
        return keepPimAttachMngrStub;
      });
      keepPimAttachMngrStub.getAttachment.resolves(attachInfo);

      const response = await GetAttachment(soapRequest, expressContextStub.request);

      
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_CANNOT_OPEN_FILE_ATTACHMENT);
      expect(response.ResponseMessages.items[0]. ResponseClass).to.be.equal(ResponseClassType.ERROR)
    })

    it('passes that parent id , attachment name and attachment buffer are correct', async function () {
      const soapRequest = generateSoapRequest();
      const arrAttach = new RequestAttachmentIdType();
      const attachInfo = new PimAttachmentInfo();
      const keepPimAttachMngrStub = sinon.createStubInstance(KeepPimAttachmentManager);

      arrAttach.Id = base64Encode(ATTACHMENTS.CASE_3);

      const bufAttach = Buffer.from(ATTACHMENTS.CASE_4, "utf-8");
      attachInfo.attachmentBuffer = bufAttach;

      soapRequest.AttachmentIds.AttachmentId.push(arrAttach);
      soapRequest.AttachmentShape = new AttachmentResponseShapeType();
      soapRequest.AttachmentShape.IncludeMimeContent = true;


      sinon.stub(KeepPimAttachmentManager, 'getInstance').callsFake((): any => {
        return keepPimAttachMngrStub;
      });
      keepPimAttachMngrStub.getAttachment.resolves(attachInfo);

      const response = await GetAttachment(soapRequest, expressContextStub.request);

      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
      expect(response.ResponseMessages.items[0]. ResponseClass).to.be.equal(ResponseClassType.SUCCESS)
    })
  });

  describe('mail.MarkAsJunk', function() {
    /*
    From LABS-2457
    Unit test to cover MarkAsJunk method in mail.ts
    */
    const FOLDER_ID = "test-folder-id";

    const context = getContext(testUser, 'password');
    
    const testItemId = new ItemIdType('test-id');
    const testItemIds = new NonEmptyArrayOfBaseItemIdsType([testItemId]);
    const testItem = new ItemType();
    testItem.ParentFolderId = new FolderIdType(getEWSId(FOLDER_ID));
    testItem.ItemId = testItemId;

    function generateSoapRequest(
        itemIds: NonEmptyArrayOfBaseItemIdsType,
        isJunk = true,
        moveItem = true,
    ): MarkAsJunkType {
        const soapRequest = new MarkAsJunkType();

        soapRequest.ItemIds = itemIds;
        soapRequest.IsJunk = isJunk;
        soapRequest.MoveItem = moveItem;

        return soapRequest;
    }

    function getLabelsManagerStub(labels?: PimLabel[]): SinonStubbedInstance<KeepPimLabelManager> {
      const stub = sinon.createStubInstance(KeepPimLabelManager);

      if (labels !== undefined) {
        stub.getLabels.resolves(labels)
        return stub;
      }

      const junkLabel = PimItemFactory.newPimLabel({ FolderId: 'junkemail' });
      junkLabel.view = KeepPimConstants.JUNKMAIL;

      stub.getLabels.resolves([junkLabel])
      return stub;
    }

    function stubGetItem(result: Promise<ItemType | undefined | Error>): SinonStub {
      return sinon.stub(EWSServiceManager, 'getItem').callsFake((): any => {
        return result;
      });
    }

    function stubMoveItemWithResult(result: Promise<ItemType | undefined | Error>): SinonStub {
      return sinon.stub(EWSServiceManager, 'moveItemWithResult').callsFake((): any => {
        return result;
      });
    }

    afterEach(() => {
      sinon.restore(); // Reset stubs
      resetUserMailBox(testUser); // Reset local user cache
    });

    it('should succeed if the JunkEmail folder exists', async function () {
      const labelsManagerStub = getLabelsManagerStub();
      stubLabelsManager(labelsManagerStub);

      stubGetItem(Promise.resolve(testItem));
      stubMoveItemWithResult(Promise.resolve(testItem))

      const soapRequest = generateSoapRequest(testItemIds);

      const response = await MarkAsJunk(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].MovedItemId).to.be.equal(testItemId)
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
    });

    it('should handle the absence of JunkEmail folder', async function () {
      const labelsManagerStub = getLabelsManagerStub([]);
      stubLabelsManager(labelsManagerStub);

      const soapRequest = generateSoapRequest(testItemIds);

      const response = await MarkAsJunk(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_FOLDER_NOT_FOUND);
    });

    it('should succeed if IsJunk === true && MoveItem === false', async function () {
      const labelsManagerStub = getLabelsManagerStub();
      stubLabelsManager(labelsManagerStub);

      stubGetItem(Promise.resolve(testItem));
      stubMoveItemWithResult(Promise.resolve(testItem))
      
      const soapRequest = generateSoapRequest(testItemIds, true, false);

      const response = await MarkAsJunk(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].MovedItemId).to.be.equal(undefined)
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
    });

    it('should succeed if IsJunk === false && MoveItem === true', async function () {
      const inboxLabel = generateTestInboxLabel();
      const labelsManagerStub = getLabelsManagerStub([inboxLabel]);
      stubLabelsManager(labelsManagerStub);

      stubGetItem(Promise.resolve(testItem));
      stubMoveItemWithResult(Promise.resolve(testItem))
      
      const soapRequest = generateSoapRequest(testItemIds, false, true);

      const response = await MarkAsJunk(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].MovedItemId).to.be.equal(testItemId)
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
    });

    it('should succeed if IsJunk === false && MoveItem === false', async function () {
      const inboxLabel = generateTestInboxLabel();
      const labelsManagerStub = getLabelsManagerStub([inboxLabel]);
      stubLabelsManager(labelsManagerStub);

      stubGetItem(Promise.resolve(testItem));
      stubMoveItemWithResult(Promise.resolve(testItem))
      
      const soapRequest = generateSoapRequest(testItemIds, false, false);

      const response = await MarkAsJunk(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].MovedItemId).to.be.equal(undefined)
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
    });

    it('should handle getting an invalid item', async function () {
      const labelsManagerStub = getLabelsManagerStub();
      stubLabelsManager(labelsManagerStub);
      
      const invalidItemIds = new NonEmptyArrayOfBaseItemIdsType([new OccurrenceItemIdType()]);
      const soapRequest = generateSoapRequest(invalidItemIds);

      const response = await MarkAsJunk(soapRequest, context.request);
      
      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_INVALID_ID);
    });

    it('should handle getting a non-existent item', async function () {
      const labelsManagerStub = getLabelsManagerStub();
      stubLabelsManager(labelsManagerStub);

      stubGetItem(Promise.resolve(undefined))

      const soapRequest = generateSoapRequest(testItemIds);

      const response = await MarkAsJunk(soapRequest, context.request);

      expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_MOVE_COPY_FAILED);
    });

    it('should handle getting a 401 error from EWSServiceManager.getItem call', async function () {
      const labelsManagerStub = getLabelsManagerStub();
      stubLabelsManager(labelsManagerStub);

      const error = { status: 401 };
      stubGetItem(Promise.reject(error))

      const soapRequest = generateSoapRequest(testItemIds);

      await expect(MarkAsJunk(soapRequest, context.request)).to.be.rejectedWith(error);
    });
  });

  describe('mail.CreateAttachment', function() {
    /*
    From LABS-2460
    Unit test to cover CreateAttachment method in mail.ts
    */
    
      enum ATTACHMENTS {
        ID = '123456',
        CHANGEKEY = '789',
        ATTACH_ID = '123##456',
        ATTACH_NAME = 'testAttach',
        ITEM_ID = '##test-item-id',
        CONTENT = 'TEST_ATTACHMENT'
      }
    
      function generateSoapRequest(): CreateAttachmentType {
        const soapRequest = new CreateAttachmentType();
        
        soapRequest.ParentItemId = new ItemIdType(ATTACHMENTS.ID, ATTACHMENTS.CHANGEKEY);
        soapRequest.Attachments = new NonEmptyArrayOfAttachmentsType();
        
        return soapRequest;
      }
      
      const expressContextStub = getContext(testUser, 'password')
      
      afterEach(() => {
        sinon.restore(); // Reset stubs
        resetUserMailBox(testUser); // Reset local user cache
      });
      
      it('passes that can not open file attachment', async function () {
        const soapRequest = generateSoapRequest();
        const item = new ItemType();
        const item1 = new FileAttachmentType();    
        const attachInfo = new PimAttachmentInfo();
    
        item.ItemId = new ItemIdType(ATTACHMENTS.ITEM_ID);
        sinon.stub(EWSServiceManager, "getItem").callsFake((itemId: string, userInfo: UserInfo, request: Request, shape?: ItemResponseShapeType) => {
          return Promise.resolve(item);
        });
    
        item1.Name = ATTACHMENTS.ATTACH_NAME;
        item1.Content = ATTACHMENTS.CONTENT;
        soapRequest.Attachments.push(item1);
    
        const keepPimAttachMngrStub = sinon.createStubInstance(KeepPimAttachmentManager);
        sinon.stub(KeepPimAttachmentManager, 'getInstance').callsFake((): any => {
          return keepPimAttachMngrStub;
        });
        keepPimAttachMngrStub.getAttachment.resolves(attachInfo);
    
        const response = await CreateAttachment(soapRequest, expressContextStub.request);
    
        expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_CANNOT_OPEN_FILE_ATTACHMENT);
        expect(response.ResponseMessages.items[0]. ResponseClass).to.be.equal(ResponseClassType.ERROR);
      });

      it('passes that attachment has not attachment name or content', async function () {
        const soapRequest = generateSoapRequest();
        const item = new ItemType();
        const item1 = new FileAttachmentType();    
        const attachInfo = new PimAttachmentInfo();
    
        item.ItemId = new ItemIdType(ATTACHMENTS.ITEM_ID);
        sinon.stub(EWSServiceManager, "getItem").callsFake((itemId: string, userInfo: UserInfo, request: Request, shape?: ItemResponseShapeType) => {
          return Promise.resolve(item);
        });
    
        soapRequest.Attachments.push(item1);
    
        const keepPimAttachMngrStub = sinon.createStubInstance(KeepPimAttachmentManager);
        sinon.stub(KeepPimAttachmentManager, 'getInstance').callsFake((): any => {
          return keepPimAttachMngrStub;
        });
        keepPimAttachMngrStub.getAttachment.resolves(attachInfo);
    
        const response = await CreateAttachment(soapRequest, expressContextStub.request);
    
        expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_INVALID_ITEM_FOR_OPERATION_CREATE_ITEM_ATTACHMENT);
        expect(response.ResponseMessages.items[0]. ResponseClass).to.be.equal(ResponseClassType.ERROR);
      });
    
      it('passes CreateAttachment and GetAttachment and get SUCCESS', async function () {
        const soapRequest = generateSoapRequest();
        const item = new ItemType();
        const item1 = new FileAttachmentType();    
        const attachInfo = new PimAttachmentInfo();
        const arrAttach = new RequestAttachmentIdType();
    
        item.ParentFolderId = new FolderIdType (ATTACHMENTS.ID, ATTACHMENTS.CHANGEKEY);
        item.ItemId = new ItemIdType(ATTACHMENTS.ATTACH_ID);
        sinon.stub(EWSServiceManager, "getItem").callsFake((itemId: string, userInfo: UserInfo, request: Request, shape?: ItemResponseShapeType) => {
          return Promise.resolve(item);
        });
    
        item1.Name = ATTACHMENTS.ATTACH_NAME;
        item1.Content = base64Encode(ATTACHMENTS.CONTENT);
        soapRequest.Attachments.push(item1);
      
        arrAttach.Id = base64Encode(ATTACHMENTS.ATTACH_ID);
        const bufAttach = Buffer.from(ATTACHMENTS.ATTACH_NAME, "utf-8");
        attachInfo.attachmentBuffer = bufAttach; 
    
        const keepPimAttachMngrStub = sinon.createStubInstance(KeepPimAttachmentManager);
        sinon.stub(KeepPimAttachmentManager, 'getInstance').callsFake((): any => {
          return keepPimAttachMngrStub;
        });
        keepPimAttachMngrStub.getAttachment.resolves(attachInfo);
    
        const response = await CreateAttachment(soapRequest, expressContextStub.request);
    
        expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
        expect(response.ResponseMessages.items[0]. ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      });
    
      it('passes that item not found', async function () {
        const soapRequest = generateSoapRequest();
        const item = new ItemType();
        const item1 = new FileAttachmentType();    
        const attachInfo = new PimAttachmentInfo();
        const arrAttach = new RequestAttachmentIdType();
    
        item.ParentFolderId = new FolderIdType (ATTACHMENTS.ID, ATTACHMENTS.CHANGEKEY);
        item.ItemId = new ItemIdType(ATTACHMENTS.ATTACH_ID);
        
        sinon.stub(EWSServiceManager, "getItem").callsFake(() => {
          return Promise.reject({ResponseCode: ResponseCodeType.ERROR_ITEM_NOT_FOUND});
        });
    
        item1.Name = ATTACHMENTS.ATTACH_NAME;
        item1.Content = base64Encode(ATTACHMENTS.CONTENT);
        soapRequest.Attachments.push(item1);
    
        arrAttach.Id = base64Encode(ATTACHMENTS.ATTACH_ID);
        const bufAttach = Buffer.from(ATTACHMENTS.ATTACH_NAME, "utf-8");
        attachInfo.attachmentBuffer = bufAttach;
      
        const response = await CreateAttachment(soapRequest, expressContextStub.request);
    
        expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_ITEM_NOT_FOUND);
        expect(response.ResponseMessages.items[0]. ResponseClass).to.be.equal(ResponseClassType.ERROR);
      });
    
      it('passes throw ERROR_INVALID_ITEM_FOR_OPERATION_CREATE_ITEM_ATTACHMENT on call createAttachment', async function () {
        const soapRequest = generateSoapRequest();
        const item = new ItemType();
        const item1 = new FileAttachmentType();    
        const attachInfo = new PimAttachmentInfo();
        const arrAttach = new RequestAttachmentIdType();
    
        item.ParentFolderId = new FolderIdType (ATTACHMENTS.ID, ATTACHMENTS.CHANGEKEY);
        item.ItemId = new ItemIdType(ATTACHMENTS.ITEM_ID);
        
        sinon.stub(EWSServiceManager, "getItem").callsFake((itemId: string, userInfo: UserInfo, request: Request, shape?: ItemResponseShapeType) => {
          return Promise.resolve(item);  
        });
    
        item1.Name = ATTACHMENTS.ATTACH_NAME;
        item1.Content = base64Encode(ATTACHMENTS.CONTENT);
        soapRequest.Attachments.push(item1);
    
        arrAttach.Id = base64Encode(ATTACHMENTS.ATTACH_ID);
        const bufAttach = Buffer.from(ATTACHMENTS.ATTACH_NAME, "utf-8");
        attachInfo.attachmentBuffer = bufAttach;
      
        const keepPimAttachMngrStub = sinon.createStubInstance(KeepPimAttachmentManager);
        sinon.stub(KeepPimAttachmentManager, 'getInstance').callsFake((): any => {
          return keepPimAttachMngrStub;      
        });
        keepPimAttachMngrStub.createAttachment.rejects({status: 500});
    
        const response = await CreateAttachment(soapRequest, expressContextStub.request);
    
        expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_INVALID_ITEM_FOR_OPERATION_CREATE_ITEM_ATTACHMENT);
        expect(response.ResponseMessages.items[0]. ResponseClass).to.be.equal(ResponseClassType.ERROR);
      });
    
      it('stubs att as not FileAttachmentType and throw ERROR_INVALID_ITEM_FOR_OPERATION_CREATE_ITEM_ATTACHMENT', async function () {
        const soapRequest = generateSoapRequest();
        const item = new ItemType();
        const item1 = new ItemAttachmentType();    
        const attachInfo = new PimAttachmentInfo();
        const arrAttach = new RequestAttachmentIdType();
    
        item.ParentFolderId = new FolderIdType (ATTACHMENTS.ID, ATTACHMENTS.CHANGEKEY);
        item.ItemId = new ItemIdType(ATTACHMENTS.ITEM_ID);
        
        sinon.stub(EWSServiceManager, "getItem").callsFake((itemId: string, userInfo: UserInfo, request: Request, shape?: ItemResponseShapeType) => {
          return Promise.resolve(item);  
        });
    
        item1.Name = ATTACHMENTS.ATTACH_NAME;
        soapRequest.Attachments.push(item1);
    
        arrAttach.Id = base64Encode(ATTACHMENTS.ATTACH_ID);
        const bufAttach = Buffer.from(ATTACHMENTS.ATTACH_NAME, "utf-8");
        attachInfo.attachmentBuffer = bufAttach;
      
    
        const keepPimAttachMngrStub = sinon.createStubInstance(KeepPimAttachmentManager);
        sinon.stub(KeepPimAttachmentManager, 'getInstance').callsFake((): any => {
          return keepPimAttachMngrStub;      
        });
        keepPimAttachMngrStub.getAttachment.resolves(attachInfo);
    
        const response = await CreateAttachment(soapRequest, expressContextStub.request);
    
        expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_INVALID_ITEM_FOR_OPERATION_CREATE_ITEM_ATTACHMENT);
        expect(response.ResponseMessages.items[0]. ResponseClass).to.be.equal(ResponseClassType.ERROR);
      });
    });
    
  describe('mail.DeleteAttachment', function() {
  /*
  From LABS-2462
  Unit test to cover DeleteAttachment method in mail.ts
  */
      
    enum ATTACHMENTS {
      ID = '123456',
      CHANGEKEY = '789',
      ATTACH_ID = '123##456',
      ATTACH_WITHOUT_NAME= '123##',
      ATTACH_NAME = 'testAttach',
      ITEM_ID = '##test-item-id',
      CONTENT = 'TEST_ATTACHMENT'  
    }
      
    function generateSoapRequest(): DeleteAttachmentType {
      const soapRequest = new DeleteAttachmentType();
      
      soapRequest.AttachmentIds = new NonEmptyArrayOfRequestAttachmentIdsType();
          
      return soapRequest;
    }
        
    const expressContextStub = getContext(testUser, 'password')
        
    afterEach(() => {
      sinon.restore(); // Reset stubs
      resetUserMailBox(testUser); // Reset local user cache
    });
        
        it('passes that can not find file attachment and parentId is null', async function () {
          const soapRequest = generateSoapRequest();
      
          const arrAttach = new RequestAttachmentIdType();
          arrAttach.Id = base64Encode(ATTACHMENTS.ID);
          soapRequest.AttachmentIds.AttachmentId.push(arrAttach);
      
          const response = await DeleteAttachment(soapRequest, expressContextStub.request);
      
          expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_CANNOT_FIND_FILE_ATTACHMENT)
          expect(response.ResponseMessages.items[0]. ResponseClass).to.be.equal(ResponseClassType.ERROR)
        })
      
        it('passes that can not find file attachment and attachment name is null', async function () {
          const soapRequest = generateSoapRequest();
          const arrAttach = new RequestAttachmentIdType();
          arrAttach.Id = base64Encode(ATTACHMENTS.ATTACH_WITHOUT_NAME);
          soapRequest.AttachmentIds.AttachmentId.push(arrAttach);
      
          const response = await DeleteAttachment(soapRequest, expressContextStub.request);
      
          expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_CANNOT_FIND_FILE_ATTACHMENT)
          expect(response.ResponseMessages.items[0]. ResponseClass).to.be.equal(ResponseClassType.ERROR)
        })
      
        it('passes no error and return SUCCESS', async function () {
          const soapRequest = generateSoapRequest();
          const arrAttach = new RequestAttachmentIdType();
          const attachInfo = new PimAttachmentInfo();
          
          arrAttach.Id = base64Encode(ATTACHMENTS.ATTACH_ID);
          soapRequest.AttachmentIds.AttachmentId.push(arrAttach);
      
          const keepPimAttachMngrStub = sinon.createStubInstance(KeepPimAttachmentManager);
          sinon.stub(KeepPimAttachmentManager, 'getInstance').callsFake((): any => {
            return keepPimAttachMngrStub;
          });
          keepPimAttachMngrStub.deleteAttachment.resolves(attachInfo);
    
          const item = new ItemType();
          item.ParentFolderId = new FolderIdType (ATTACHMENTS.ID, ATTACHMENTS.CHANGEKEY);
          item.ItemId = new ItemIdType(ATTACHMENTS.ITEM_ID);
    
          sinon.stub(EWSServiceManager, "getItem").callsFake((itemId: string, userInfo: UserInfo, request: Request, shape?: ItemResponseShapeType) => {
            return Promise.resolve(item);  
          });
      
          const response = await DeleteAttachment(soapRequest, expressContextStub.request);
          
          expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
          expect(response.ResponseMessages.items[0]. ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
        });
    
        it('rejects deleteAttachment with error 404 and throw ERROR_CANNOT_FIND_FILE_ATTACHMENT', async function () {
          const soapRequest = generateSoapRequest();
          const arrAttach = new RequestAttachmentIdType();
                
          arrAttach.Id = base64Encode(ATTACHMENTS.ATTACH_ID);
          soapRequest.AttachmentIds.AttachmentId.push(arrAttach);
      
          const keepPimAttachMngrStub = sinon.createStubInstance(KeepPimAttachmentManager);
          sinon.stub(KeepPimAttachmentManager, 'getInstance').callsFake((): any => {
            return keepPimAttachMngrStub;
          });
          keepPimAttachMngrStub.deleteAttachment.rejects({status: 404});
    
          const response = await DeleteAttachment(soapRequest, expressContextStub.request);
          
          expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_CANNOT_FIND_FILE_ATTACHMENT);
          expect(response.ResponseMessages.items[0]. ResponseClass).to.be.equal(ResponseClassType.ERROR);
        })
    
        it('stubs EWSServiceManager.getItem and throw ERROR_CANNOT_FIND_FILE_ATTACHMENT', async function () {
          const soapRequest = generateSoapRequest();
          const arrAttach = new RequestAttachmentIdType();
          const attachInfo = new PimAttachmentInfo();
          const error = { ResponseCode: ResponseCodeType.ERROR_CANNOT_FIND_FILE_ATTACHMENT }
          
          arrAttach.Id = base64Encode(ATTACHMENTS.ATTACH_ID);
          soapRequest.AttachmentIds.AttachmentId.push(arrAttach);
      
          const keepPimAttachMngrStub = sinon.createStubInstance(KeepPimAttachmentManager);
          sinon.stub(KeepPimAttachmentManager, 'getInstance').callsFake((): any => {
            return keepPimAttachMngrStub;
          });
          keepPimAttachMngrStub.deleteAttachment.resolves(attachInfo);
    
          const item = new ItemType();
          item.ParentFolderId = new FolderIdType (ATTACHMENTS.ID, ATTACHMENTS.CHANGEKEY);
        
          sinon.stub(EWSServiceManager, "getItem").callsFake((itemId: string, userInfo: UserInfo, request: Request, shape?: any) => {
            return Promise.reject(error); 
          });
      
          const response = await DeleteAttachment(soapRequest, expressContextStub.request);
          
          expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_CANNOT_FIND_FILE_ATTACHMENT);
          expect(response.ResponseMessages.items[0]. ResponseClass).to.be.equal(ResponseClassType.ERROR);
    });
  });

  describe('mail.UploadItems', function() {
  /*
  From LABS-3015
  Unit test to cover UploadItems method in mail.ts
  */
    enum UPLOADS {
      ITEM_ID = 'test-item-id',
      FOLDER_ID = 'test-folder-id',
      CONTENT = 'AQAAAAg=',
      DELEGATOR_MAILBOX_ID = 'delegator@test.com' 
    }
    
    const itemEWSId = getEWSId(UPLOADS.ITEM_ID, UPLOADS.DELEGATOR_MAILBOX_ID);

    function generateSoapRequest(): UploadItemsType {
      const soapRequest = new UploadItemsType();
      const item = new UploadItemType();
      const parentEWSId = getEWSId(UPLOADS.FOLDER_ID, UPLOADS.DELEGATOR_MAILBOX_ID);

      item.CreateAction = CreateActionType.CREATE_NEW;
      item.Data = base64Encode(UPLOADS.CONTENT);
      item.IsAssociated = false;
      item.ItemId = new ItemIdType(itemEWSId);
      item.ParentFolderId = new FolderIdType(parentEWSId);

      soapRequest.Items = new NonEmptyArrayOfUploadItemsType();
      soapRequest.Items.Item = [item];
      
      return soapRequest;
    }

    const expressContextStub = getContext(testUser, 'password');

    let labelManagerStub: sinon.SinonStubbedInstance<KeepPimLabelManager>;
    let messageManagerStub: sinon.SinonStubbedInstance<KeepPimMessageManager>;
    
    beforeEach(() => {
      labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);

      stubMessageManager(messageManagerStub);
      messageManagerStub.updateMimeMessage.resolves(UPLOADS.ITEM_ID);
  
      stubLabelsManager(labelManagerStub);
      labelManagerStub.getLabels.resolves([PimItemFactory.newPimLabel({ FolderId: UPLOADS.FOLDER_ID})]);
      
      const serviceStub = createSinonStubInstance(EWSMessageManager);
      const item = new ItemType();
      item.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
      serviceStub.createItem.resolves([item]);

      stubServiceManagerForMessages(serviceStub);
    });  

    afterEach(() => {
      sinon.restore(); // Reset stubs
      resetUserMailBox(testUser); // Reset local user cache
    });
        
        it('passes no error and return SUCCESS', async function () {
          const soapRequest = generateSoapRequest();
          
          const response = await UploadItems(soapRequest, expressContextStub.request)
          
          expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
          expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
          expect(response.ResponseMessages.items[0].ItemId.Id).to.be.equal(itemEWSId);
        });

        it('passes with error if item is associated with a folder', async function () {
          const soapRequest = generateSoapRequest();
          soapRequest.Items.Item[0].IsAssociated = true;
          
          const response = await UploadItems(soapRequest, expressContextStub.request);
          
          expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR);
          expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
          expect(response.ResponseMessages.items[0].MessageText).to.be.equal(`UploadItems for Folders is not implemented`);
        }); 

        it('passes with error when update request did not include an item id', async function () {
          const soapRequest = generateSoapRequest();
          soapRequest.Items.Item[0].CreateAction = CreateActionType.UPDATE;
          soapRequest.Items.Item[0].ItemId = undefined;
          
          const response = await UploadItems(soapRequest, expressContextStub.request);
          
          expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.ERROR_INVALID_ID_EMPTY);
          expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.ERROR);
          expect(response.ResponseMessages.items[0].MessageText).to.be.equal(`UploadItems update request did not include an item id.`);
        });
        
        it('should update item with update action type', async function () {
          const soapRequest = generateSoapRequest();
          soapRequest.Items.Item[0].CreateAction = CreateActionType.UPDATE;
          
          const response = await UploadItems(soapRequest, expressContextStub.request);
          
          expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
          expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
        });

        it('throws an error when the pimItemId is undefined with UPDATE action type', async function () {
          const soapRequest = generateSoapRequest();
          soapRequest.Items.Item[0].CreateAction = CreateActionType.UPDATE;
          soapRequest.Items.Item[0].ItemId!.Id = '';
    
          const expectedError = new Error(`Invalid PIM item id ${soapRequest.Items.Item[0].ItemId?.Id}`);

          await expect(UploadItems(soapRequest, expressContextStub.request)).to.be.rejectedWith(expectedError);
        });

        it('throws an error when the pimItemId is undefinde with UPDATE_OR_CREATE action type', async function () {
          const soapRequest = generateSoapRequest();
          soapRequest.Items.Item[0].CreateAction = CreateActionType.UPDATE_OR_CREATE;
          soapRequest.Items.Item[0].ItemId!.Id = '';
    
          const expectedError = new Error(`Invalid PIM item id ${soapRequest.Items.Item[0].ItemId?.Id}`);

          await expect(UploadItems(soapRequest, expressContextStub.request)).to.be.rejectedWith(expectedError);
        });

        it('should create new item when ItemId is undefined', async function () {
          const soapRequest = generateSoapRequest();
          soapRequest.Items.Item[0].CreateAction = CreateActionType.UPDATE_OR_CREATE;
          soapRequest.Items.Item[0].ItemId = undefined;
          
          const response = await UploadItems(soapRequest, expressContextStub.request);
          
          expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
          expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
          expect(response.ResponseMessages.items[0].ItemId.Id).to.be.equal(itemEWSId);
        });

        it('should update item with UPDATE_OR_CREATE action type', async function () {
          const soapRequest = generateSoapRequest();
          soapRequest.Items.Item[0].CreateAction = CreateActionType.UPDATE_OR_CREATE;
          
          const response = await UploadItems(soapRequest, expressContextStub.request);
          
          expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
          expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
          expect(response.ResponseMessages.items[0].ItemId.Id).to.be.equal(itemEWSId);
          expect(response.ResponseMessages.items[0].ItemId.ChangeKey).to.be.equal(`ck-${itemEWSId}`);
        });

        it('should create item when item was not find', async function () {
          const soapRequest = generateSoapRequest();
          soapRequest.Items.Item[0].CreateAction = CreateActionType.UPDATE_OR_CREATE;
          messageManagerStub.updateMimeMessage.rejects({status: 404});
          
          const response = await UploadItems(soapRequest, expressContextStub.request);
          
          expect(response.ResponseMessages.items[0].ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
          expect(response.ResponseMessages.items[0].ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
          expect(response.ResponseMessages.items[0].ItemId.Id).to.be.equal(itemEWSId);
        })
  });
});
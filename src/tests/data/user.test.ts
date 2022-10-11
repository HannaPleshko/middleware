/* eslint-disable @typescript-eslint/no-non-null-assertion */
import { 
  base64Encode, base64Decode, KeepPimLabelManager, KeepPimManager, PimOOO, 
  KeepPimCalendarManager } from '@hcllabs/openclientkeepcomponent';
import { expect, sinon } from '@loopback/testlab';
import { 
  CreateUserConfigurationResponse, GetUserConfigurationResponse, UpdateUserConfigurationResponse, GetUserOofSettings,
   DeleteUserConfigurationResponse, SetUserOofSettings} from '../../data/user';
import { 
  DistinguishedFolderIdNameType, ExternalAudience, OofState, ResponseClassType, ResponseCodeType, 
  UserConfigurationPropertyTypeEnum } from '../../models/enum.model';
import { DistinguishedFolderIdType, ItemIdType, UserConfigurationNameType } from '../../models/mail.model';
import { 
  CreateUserConfigurationType, GetUserConfigurationType, UpdateUserConfigurationType, UserConfigurationDictionaryType,
  UserConfigurationType, EmailAddress, GetUserOofSettingsRequest, ReplyBody, 
  DeleteUserConfigurationType, SetUserOofSettingsRequest, UserOofSettings, Duration} from '../../models/user.model';
import { 
  getContext, resetUserMailBox,  generateTestDraftsLabel, stubLabelsManager,stubPimManager, 
  stubCalendarManager } from '../unitTestHelper';

const testUser = "test.user@test.org";

describe('user.ts tests', () => {
  describe('user.GetUserConfigurationResponse', function() {
    /*
    From LABS-2558
    Add unit tests for GetUserConfiguration
    */
    const BINARY_DATA = 'BinaryData';
    const XML_DATA = '<div><p>XmlData</p></div>';
    const CONFIG_NAME = 'TestConfig';

    const TEST_CONFIG = {
      Dictionary: 'Dictionary',
      XmlData: base64Encode(XML_DATA),
      BinaryData: base64Encode(BINARY_DATA),
      Id: 'Id',
    }

    function generateSoapRequest (type: UserConfigurationPropertyTypeEnum, distinguished = true): GetUserConfigurationType {
      const soapRequest = new GetUserConfigurationType();

      soapRequest.UserConfigurationName = new UserConfigurationNameType();
      soapRequest.UserConfigurationName.Name = CONFIG_NAME;

      soapRequest.UserConfigurationName.DistinguishedFolderId = new DistinguishedFolderIdType();
      soapRequest.UserConfigurationName.DistinguishedFolderId.Id = DistinguishedFolderIdNameType.DRAFTS;

      soapRequest.UserConfigurationProperties[0] = type;

      return soapRequest;
    }

    let labelManagerStub: sinon.SinonStubbedInstance<KeepPimLabelManager>;

    const expressContextStub = getContext(testUser, 'password');
    
    beforeEach(() => {
      labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      labelManagerStub.getLabels.resolves([generateTestDraftsLabel()]);
      stubLabelsManager(labelManagerStub);
    });

    afterEach(() => {
      sinon.restore(); 
      resetUserMailBox(testUser);
    });

    describe('asks the user configuration for different types', function() {
      beforeEach(() => {
        const testDraftsLabel = generateTestDraftsLabel();
        testDraftsLabel.setAdditionalProperty(CONFIG_NAME, TEST_CONFIG);
        labelManagerStub.getLabels.resolves([testDraftsLabel]);
      });

      it('asks for the Dictionary type', async function() {
        const mockSoapRequest = generateSoapRequest(UserConfigurationPropertyTypeEnum.DICTIONARY);
    
        const expectedResponse = await GetUserConfigurationResponse(mockSoapRequest, expressContextStub.request);
        
        expect(expectedResponse.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.SUCCESS);
        expect(expectedResponse.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.NO_ERROR);
        expect(expectedResponse.ResponseMessages.items[0].UserConfiguration?.UserConfigurationName.Name).to.equal(CONFIG_NAME);
        expect(expectedResponse.ResponseMessages.items[0].UserConfiguration?.Dictionary).to.equal(TEST_CONFIG.Dictionary);
      });
    
      it('asks for the XmlData type', async function() {
        const mockSoapRequest = generateSoapRequest(UserConfigurationPropertyTypeEnum.XML_DATA);
    
        const expectedResponse = await GetUserConfigurationResponse(mockSoapRequest, expressContextStub.request);
        const responseXMLData = expectedResponse.ResponseMessages.items[0].UserConfiguration?.XmlData;

        expect(expectedResponse.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.SUCCESS);
        expect(expectedResponse.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.NO_ERROR);
        expect(expectedResponse.ResponseMessages.items[0].UserConfiguration?.UserConfigurationName.Name).to.equal(CONFIG_NAME);
        expect(base64Decode(responseXMLData as string)).to.equal(XML_DATA);
      });
    
      it('asks for the BinaryData type', async function() {
        const mockSoapRequest = generateSoapRequest(UserConfigurationPropertyTypeEnum.BINARY_DATA);
    
        const expectedResponse = await GetUserConfigurationResponse(mockSoapRequest, expressContextStub.request);
        const responseBinaryData = expectedResponse.ResponseMessages.items[0].UserConfiguration?.BinaryData;
        
        expect(expectedResponse.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.SUCCESS);
        expect(expectedResponse.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.NO_ERROR);
        expect(expectedResponse.ResponseMessages.items[0].UserConfiguration?.UserConfigurationName.Name).to.equal(CONFIG_NAME);
        expect(base64Decode(responseBinaryData as string)).to.equal(BINARY_DATA);
      });
    
      it('asks for ItemId type', async function() {
        const mockSoapRequest = generateSoapRequest(UserConfigurationPropertyTypeEnum.ID);
    
        const expectedResponse = await GetUserConfigurationResponse(mockSoapRequest, expressContextStub.request);
        
        expect(expectedResponse.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.SUCCESS);
        expect(expectedResponse.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.NO_ERROR);
        expect(expectedResponse.ResponseMessages.items[0].UserConfiguration?.UserConfigurationName.Name).to.equal(CONFIG_NAME);
        expect(expectedResponse.ResponseMessages.items[0].UserConfiguration?.ItemId).to.equal(TEST_CONFIG.Id);
      });
    
      it('asks for the All type', async function() {
        const mockSoapRequest = generateSoapRequest(UserConfigurationPropertyTypeEnum.ALL);
    
        const expectedResponse = await GetUserConfigurationResponse(mockSoapRequest, expressContextStub.request);
        const responseXMLData = expectedResponse.ResponseMessages.items[0].UserConfiguration?.XmlData;
        const responseBinaryData = expectedResponse.ResponseMessages.items[0].UserConfiguration?.BinaryData;
        
        expect(expectedResponse.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.SUCCESS);
        expect(expectedResponse.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.NO_ERROR);
        expect(expectedResponse.ResponseMessages.items[0].UserConfiguration?.UserConfigurationName.Name).to.equal(CONFIG_NAME);
        expect(expectedResponse.ResponseMessages.items[0].UserConfiguration?.Dictionary).to.equal(TEST_CONFIG.Dictionary);
        expect(base64Decode(responseXMLData as string)).to.equal(XML_DATA);
        expect(base64Decode(responseBinaryData as string)).to.equal(BINARY_DATA);
        expect(expectedResponse.ResponseMessages.items[0].UserConfiguration?.ItemId).to.equal(TEST_CONFIG.Id);
      });
    });

    it('gets userconfiguration for a folder that does not have a userconfiguration', async function() {
      const mockSoapRequest = generateSoapRequest(UserConfigurationPropertyTypeEnum.ALL);

      const expectedResponse = await GetUserConfigurationResponse(mockSoapRequest, expressContextStub.request);

      expect(expectedResponse.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.SUCCESS);
      expect(expectedResponse.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.NO_ERROR);
      expect(expectedResponse.ResponseMessages.items[0].UserConfiguration?.UserConfigurationName.Name).to.equal(CONFIG_NAME);
    });

    it('gets userconfiguration for a folder for a folder that does not exist', async function() {
      const mockSoapRequest = generateSoapRequest(UserConfigurationPropertyTypeEnum.ALL);
      const mockError = { status: 404 };
      labelManagerStub.getLabels.rejects(mockError);

      const expectedResponse = await GetUserConfigurationResponse(mockSoapRequest, expressContextStub.request);

      expect(expectedResponse.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.ERROR);
      expect(expectedResponse.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.ERROR_FOLDER_NOT_FOUND);
    });
  });

  describe('user.CreateUserConfigurationResponse', function() {
    function generateSoapRequest (): CreateUserConfigurationType {
      const soapRequest = new CreateUserConfigurationType();

      soapRequest.UserConfiguration = new UserConfigurationType();
      soapRequest.UserConfiguration.UserConfigurationName = new UserConfigurationNameType();
      soapRequest.UserConfiguration.UserConfigurationName.Name = 'TestConfig';

      soapRequest.UserConfiguration.UserConfigurationName.DistinguishedFolderId = new DistinguishedFolderIdType();
      soapRequest.UserConfiguration.UserConfigurationName.DistinguishedFolderId.Id = DistinguishedFolderIdNameType.DRAFTS;

      soapRequest.UserConfiguration.XmlData = base64Encode('<div><p>XmlData</p></div>');
      soapRequest.UserConfiguration.Dictionary = new UserConfigurationDictionaryType();
      soapRequest.UserConfiguration.ItemId = new ItemIdType('test-id');
      soapRequest.UserConfiguration.BinaryData = base64Encode('BinaryData');

      return soapRequest;
    }

    let labelManagerStub: sinon.SinonStubbedInstance<KeepPimLabelManager>;
    let pimManagerSub: sinon.SinonStubbedInstance<KeepPimManager>;

    const expressContextStub = getContext(testUser, 'password');
    
    beforeEach(() => {
      labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      labelManagerStub.getLabels.resolves([generateTestDraftsLabel()]);
      stubLabelsManager(labelManagerStub);

      pimManagerSub = sinon.createStubInstance(KeepPimManager);
      pimManagerSub.updatePimItem.resolves();
      stubPimManager(pimManagerSub);
    });

    afterEach(() => {
      sinon.restore();
      resetUserMailBox(testUser);
    });

    it('returns an error when a folder does not exist', async function() {
      const mockSoapRequest = generateSoapRequest();
      const mockError = { status: 404 };
      labelManagerStub.getLabels.rejects(mockError);

      const expectedResponse = await CreateUserConfigurationResponse(mockSoapRequest, expressContextStub.request);

      expect(expectedResponse.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.ERROR);
      expect(expectedResponse.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.ERROR_FOLDER_NOT_FOUND);
    });

    it('returns an error when receiving a folder was rejected', async function() {
      const mockSoapRequest = generateSoapRequest();
      const mockError = { message: 'test error' };
      labelManagerStub.getLabels.rejects(mockError);

      const expectedResponse = await CreateUserConfigurationResponse(mockSoapRequest, expressContextStub.request);

      expect(expectedResponse.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.ERROR);
      expect(expectedResponse.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR);
      expect(expectedResponse.ResponseMessages.items[0].MessageText).to.equal(mockError.message);
    });

    it('returns an error when pimItem was not updated', async function() {
      const mockSoapRequest = generateSoapRequest();
      const mockError = { message: 'test error' };
      pimManagerSub.updatePimItem.rejects(mockError);

      const expectedResponse = await CreateUserConfigurationResponse(mockSoapRequest, expressContextStub.request);

      expect(expectedResponse.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.ERROR);
      expect(expectedResponse.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR);
      expect(expectedResponse.ResponseMessages.items[0].MessageText).to.equal(mockError.message);
    });

    it('creates user configuration binaryData, dictionary, xmlData, ItemId types for a folder', async function() {
      const mockSoapRequest = generateSoapRequest();

      const expectedResponse = await CreateUserConfigurationResponse(mockSoapRequest, expressContextStub.request);

      expect(expectedResponse.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.SUCCESS);
      expect(expectedResponse.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.NO_ERROR);
    });
  });

  describe('user.UpdateUserConfigurationResponse', function() {
    function generateSoapRequest (): UpdateUserConfigurationType {
      const soapRequest = new UpdateUserConfigurationType();

      soapRequest.UserConfiguration = new UserConfigurationType();
      soapRequest.UserConfiguration.UserConfigurationName = new UserConfigurationNameType();
      soapRequest.UserConfiguration.UserConfigurationName.Name = 'TestConfig';

      soapRequest.UserConfiguration.UserConfigurationName.DistinguishedFolderId = new DistinguishedFolderIdType();
      soapRequest.UserConfiguration.UserConfigurationName.DistinguishedFolderId.Id = DistinguishedFolderIdNameType.DRAFTS;

      soapRequest.UserConfiguration.XmlData = base64Encode('<div><p>new XmlData</p></div>');
      soapRequest.UserConfiguration.Dictionary = new UserConfigurationDictionaryType();
      soapRequest.UserConfiguration.ItemId = new ItemIdType('new-test-id');
      soapRequest.UserConfiguration.BinaryData = base64Encode('new BinaryData');

      return soapRequest;
    }
    
    let labelManagerStub: sinon.SinonStubbedInstance<KeepPimLabelManager>;
    let pimManagerSub: sinon.SinonStubbedInstance<KeepPimManager>;

    const expressContextStub = getContext(testUser, 'password');

    beforeEach(() => {
      labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      labelManagerStub.getLabels.resolves([generateTestDraftsLabel()]);
      stubLabelsManager(labelManagerStub);

      pimManagerSub = sinon.createStubInstance(KeepPimManager);
      pimManagerSub.updatePimItem.resolves();
      stubPimManager(pimManagerSub);
    });

    afterEach(() => {
      sinon.restore();
      resetUserMailBox(testUser);
    });

    it('returns an error when receiving a folder was rejected', async function() {
      const mockSoapRequest = generateSoapRequest();
      const mockError = { message: 'test error' };
      labelManagerStub.getLabels.rejects(mockError);

      const expectedResponse = await UpdateUserConfigurationResponse(mockSoapRequest, expressContextStub.request);

      expect(expectedResponse.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.ERROR);
      expect(expectedResponse.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR);
      expect(expectedResponse.ResponseMessages.items[0].MessageText).to.equal(mockError.message);
    });

    it('returns an error when a folder does not exist', async function() {
      const mockSoapRequest = generateSoapRequest();
      const mockError = { status: 404 };
      labelManagerStub.getLabels.rejects(mockError);

      const expectedResponse = await UpdateUserConfigurationResponse(mockSoapRequest, expressContextStub.request);

      expect(expectedResponse.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.ERROR);
      expect(expectedResponse.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.ERROR_FOLDER_NOT_FOUND);
    });

    it('returns an error when pimItem was not updated', async function() {
      const mockSoapRequest = generateSoapRequest();
      const mockError = { message: 'test error' };
      pimManagerSub.updatePimItem.rejects(mockError);

      const expectedResponse = await UpdateUserConfigurationResponse(mockSoapRequest, expressContextStub.request);

      expect(expectedResponse.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.ERROR);
      expect(expectedResponse.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR);
      expect(expectedResponse.ResponseMessages.items[0].MessageText).to.equal(mockError.message);
    });

    it('updates user configuration binaryData, dictionary, xmlData, ItemId types for a folder', async function() {
      const mockSoapRequest = generateSoapRequest();

      const expectedResponse = await UpdateUserConfigurationResponse(mockSoapRequest, expressContextStub.request);

      expect(expectedResponse.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.SUCCESS);
      expect(expectedResponse.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.NO_ERROR);
    });
  });

  describe('user.GetUserOofSettings', function() {
    /*
    From LABS-2463
    Unit test to cover GetUserOofSettings method in user.ts
    */
    
    const testUser1 = "test.user1@test.org";

    function generateSoapRequest(): GetUserOofSettingsRequest {
      const soapRequest = new GetUserOofSettingsRequest();
      soapRequest.Mailbox = new EmailAddress();
      soapRequest.Mailbox.Address = testUser;
      return soapRequest;
    }
    
    let keepPimCalendarManagerStub: sinon.SinonStubbedInstance<KeepPimCalendarManager>;
    let settings: PimOOO;
    const expressContextStub = getContext(testUser, 'password');
    
    beforeEach(() => {
      keepPimCalendarManagerStub = sinon.createStubInstance(KeepPimCalendarManager);
      stubCalendarManager(keepPimCalendarManagerStub);
    });

    afterEach(() => {
      sinon.restore(); // Reset stubs
      resetUserMailBox(testUser); // Reset local user cache
    });
    
    it('passes the authenticated user does not have access to the OOO settings', async function () {
      const soapRequest = generateSoapRequest();
      soapRequest.Mailbox.Address = testUser1;
      
      const response = await GetUserOofSettings(soapRequest, expressContextStub.request);
      
      expect(response.ResponseMessage.ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessage.ResponseCode).to.be.equal(ResponseCodeType.ERROR_ACCESS_DENIED);
    });

    it('passes failed to retrieve user out-of-office settings', async function () {
      const soapRequest = generateSoapRequest();
      keepPimCalendarManagerStub.getOOO.rejects();
                
      const response = await GetUserOofSettings(soapRequest, expressContextStub.request);
      
      expect(response.ResponseMessage.ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessage.ResponseCode).to.be.equal(ResponseCodeType.ERROR_SERVICE_UNAVAILABLE);
    });

    it('passes OOO is disabled', async function () {
      const soapRequest = generateSoapRequest();
      keepPimCalendarManagerStub.getOOO.resolves(undefined);
                
      const response = await GetUserOofSettings(soapRequest, expressContextStub.request);
      
      expect(response.ResponseMessage.ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessage.ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
      expect(response.OofSettings?.OofState).to.be.equal(OofState.DISABLED);
      expect(response.OofSettings?.ExternalAudience).to.be.equal(ExternalAudience.NONE);
    });

    it('passes settings OOOState is SCHEDULED for ExternalAudience is ALL', async function () {
      const soapRequest = generateSoapRequest();
      settings = new PimOOO(
        {
          "StartDateTime": "2021-08-23T15:35:12Z",
          "EndDateTime": "2021-08-24T15:35:12Z",
          "ExcludeInternet": false,
          "Enabled": true,
          "GeneralSubject": "Test User is out of the office",
          "GeneralMessage": ""
        }
      );  
      const reply = new ReplyBody(settings.replyMessage);
      keepPimCalendarManagerStub.getOOO.resolves(settings);
                
      const response = await GetUserOofSettings(soapRequest, expressContextStub.request);
      
      expect(response.ResponseMessage.ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessage.ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
      expect(response.AllowExternalOof).to.be.equal(ExternalAudience.ALL);
      expect(response.OofSettings?.ExternalAudience).to.be.equal(ExternalAudience.ALL);
      expect(response.OofSettings?.OofState).to.be.equal(OofState.SCHEDULED);
      expect(response.OofSettings!.Duration!.StartTime.toISOString()).to.be.equal(settings.startDate!.toISOString());
      expect(response.OofSettings!.Duration!.EndTime.toISOString()).to.be.equal(settings.endDate!.toISOString());
      expect(response.OofSettings?.ExternalReply?.Message).to.be.equal(reply.Message);
      expect(response.OofSettings?.InternalReply?.Message).to.be.equal(reply.Message);
    });

    it('passes settings OOOState is ENABLED for ExternalAudience is NONE', async function () {
      const soapRequest = generateSoapRequest();
      settings = new PimOOO(
        {
          "ExcludeInternet": true,
          "Enabled": true,
          "GeneralSubject": "Test User is out of the office",
          "GeneralMessage": ""
        }
      );  
      const reply = new ReplyBody(settings.replyMessage);
      keepPimCalendarManagerStub.getOOO.resolves(settings);
                
      const response = await GetUserOofSettings(soapRequest, expressContextStub.request);
      
      expect(response.ResponseMessage.ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessage.ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
      expect(response.AllowExternalOof).to.be.equal(ExternalAudience.NONE);
      expect(response.OofSettings?.ExternalAudience).to.be.equal(ExternalAudience.NONE);
      expect(response.OofSettings?.OofState).to.be.equal(OofState.ENABLED);
      expect(response.OofSettings?.ExternalReply?.Message).to.be.equal(reply.Message);
      expect(response.OofSettings?.InternalReply?.Message).to.be.equal(reply.Message);
    });
  });

  describe('user.DeleteUserConfigurationResponse', function() {
    const BINARY_DATA = 'BinaryData';
    const XML_DATA = '<div><p>XmlData</p></div>';
    const CONFIG_NAME = 'TestConfig';

    const TEST_CONFIG = {
      Dictionary: 'Dictionary',
      XmlData: base64Encode(XML_DATA),
      BinaryData: base64Encode(BINARY_DATA),
      Id: 'Id',
    };

    function generateSoapRequest(): DeleteUserConfigurationType {
      const soapRequest = new DeleteUserConfigurationType();

      soapRequest.UserConfigurationName = new UserConfigurationNameType();
      soapRequest.UserConfigurationName.Name = CONFIG_NAME;

      soapRequest.UserConfigurationName.DistinguishedFolderId = new DistinguishedFolderIdType();
      soapRequest.UserConfigurationName.DistinguishedFolderId.Id = DistinguishedFolderIdNameType.DRAFTS;

      return soapRequest;
    }

    let labelManagerStub: sinon.SinonStubbedInstance<KeepPimLabelManager>;

    const expressContextStub = getContext(testUser, 'password');

    beforeEach(() => {
      labelManagerStub = sinon.createStubInstance(KeepPimLabelManager);
      stubLabelsManager(labelManagerStub);
    });

    afterEach(() => {
      sinon.restore();
      resetUserMailBox(testUser);
    });

    it('returns an error when receiving a folder was rejected', async function() {
      const mockSoapRequest = generateSoapRequest();
      const mockError = { message: 'test error' };
      labelManagerStub.getLabels.rejects(mockError);

      const expectedResponse = await DeleteUserConfigurationResponse(mockSoapRequest, expressContextStub.request);

      expect(expectedResponse.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.ERROR);
      expect(expectedResponse.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR);
      expect(expectedResponse.ResponseMessages.items[0].MessageText).to.equal(mockError.message);
    });

    it('returns an error when a folder does not exist', async function() {
      const mockSoapRequest = generateSoapRequest();
      const mockError = { status: 404 };
      labelManagerStub.getLabels.rejects(mockError);

      const expectedResponse = await DeleteUserConfigurationResponse(mockSoapRequest, expressContextStub.request);

      expect(expectedResponse.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.ERROR);
      expect(expectedResponse.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.ERROR_FOLDER_NOT_FOUND);
    });

    it('passes if the user configuration with asked name is not set', async function() {
      const mockSoapRequest = generateSoapRequest();
      labelManagerStub.getLabels.resolves([generateTestDraftsLabel()]);

      const expectedResponse = await DeleteUserConfigurationResponse(mockSoapRequest, expressContextStub.request);

      expect(expectedResponse.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.SUCCESS);
      expect(expectedResponse.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.NO_ERROR);
    });
    
    describe('user configuration with asked name is set', function() {
      let pimManagerSub: sinon.SinonStubbedInstance<KeepPimManager>;

      beforeEach(() => {
        const testDraftsLabel = generateTestDraftsLabel();
        testDraftsLabel.setAdditionalProperty(CONFIG_NAME, TEST_CONFIG);
        labelManagerStub.getLabels.resolves([testDraftsLabel]);

        pimManagerSub = sinon.createStubInstance(KeepPimManager);
        stubPimManager(pimManagerSub);
      });

      it('returns an error when pimItem was not updated', async function() {
        const mockSoapRequest = generateSoapRequest();
        const mockError = { message: 'test error' };
        pimManagerSub.updatePimItem.rejects(mockError);
    
        const expectedResponse = await DeleteUserConfigurationResponse(mockSoapRequest, expressContextStub.request);
    
        expect(expectedResponse.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.ERROR);
        expect(expectedResponse.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.ERROR_INTERNAL_SERVER_ERROR);
        expect(expectedResponse.ResponseMessages.items[0].MessageText).to.equal(mockError.message);
      });

      it('passes when the user configuration is deleted', async function() {
        const mockSoapRequest = generateSoapRequest();
        pimManagerSub.updatePimItem.resolves();
    
        const expectedResponse = await DeleteUserConfigurationResponse(mockSoapRequest, expressContextStub.request);
    
        expect(expectedResponse.ResponseMessages.items[0].ResponseClass).to.equal(ResponseClassType.SUCCESS);
        expect(expectedResponse.ResponseMessages.items[0].ResponseCode).to.equal(ResponseCodeType.NO_ERROR);
      });
    });
  });

  describe('user.SetUserOofSettings', function() {
    /*
    From LABS-2464
    Unit test to cover SetUserOofSettings method in user.ts
    */
    
    const testUser1 = "test.user1@test.org";

    function generateSoapRequest(email: string ): SetUserOofSettingsRequest {
      const soapRequest = new SetUserOofSettingsRequest();
      soapRequest.Mailbox = new EmailAddress();
      soapRequest.Mailbox.Address = email;
      soapRequest.UserOofSettings = new UserOofSettings;
      return soapRequest;
    }
    
    let keepPimCalendarManagerStub: sinon.SinonStubbedInstance<KeepPimCalendarManager>;
    const expressContextStub = getContext(testUser, 'password');
    
    beforeEach(() => {
      keepPimCalendarManagerStub = sinon.createStubInstance(KeepPimCalendarManager);
      stubCalendarManager(keepPimCalendarManagerStub);
    });

    afterEach(() => {
      sinon.restore(); // Reset stubs
      resetUserMailBox(testUser); // Reset local user cache
    });
    
    it('passes the authenticated user does not have access to the OOO settings', async function () {
      const soapRequest = generateSoapRequest(testUser1);
      
      const response = await SetUserOofSettings(soapRequest, expressContextStub.request);
      
      expect(response.ResponseMessage?.ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessage?.ResponseCode).to.be.equal(ResponseCodeType.ERROR_ACCESS_DENIED);
    });

    it('passes failed to update user out-of-office settings', async function () {
      const soapRequest = generateSoapRequest(testUser);
      const mockError = { message: 'test error' };
      keepPimCalendarManagerStub.updateOOO.rejects(mockError);
                
      const response = await SetUserOofSettings(soapRequest, expressContextStub.request);
      
      expect(response.ResponseMessage?.ResponseClass).to.be.equal(ResponseClassType.ERROR);
      expect(response.ResponseMessage?.ResponseCode).to.be.equal(ResponseCodeType.ERROR_SERVICE_UNAVAILABLE);
      expect(response.ResponseMessage?.MessageText).to.be.equal(`Failed to update user out-of-office settings`);
    });

    it('passes set OOO settings to Disabled', async function () {
      const soapRequest = generateSoapRequest(testUser);
      soapRequest.UserOofSettings.OofState = OofState.DISABLED;
      keepPimCalendarManagerStub.updateOOO.resolves();
                
      const response = await SetUserOofSettings(soapRequest, expressContextStub.request);
      
      expect(response.ResponseMessage?.ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessage?.ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);    
    });

    it('passes set OOO settings ENABLED with All options', async function () {
      const soapRequest = generateSoapRequest(testUser);
      soapRequest.UserOofSettings.OofState = OofState.ENABLED;
      soapRequest.UserOofSettings.ExternalAudience = ExternalAudience.ALL;
      soapRequest.UserOofSettings.Duration = new Duration;
      soapRequest.UserOofSettings.Duration.StartTime = new Date("2021-08-31T08:00:00Z");
      soapRequest.UserOofSettings.Duration.EndTime = new Date("2021-09-01T21:00:00Z");
      soapRequest.UserOofSettings.InternalReply = new ReplyBody("Test User is out of the office");
      soapRequest.UserOofSettings.ExternalReply = new ReplyBody("Test User is out of the office");
      keepPimCalendarManagerStub.updateOOO.resolves();
      
      const response = await SetUserOofSettings(soapRequest, expressContextStub.request);
      
      expect(response.ResponseMessage?.ResponseClass).to.be.equal(ResponseClassType.SUCCESS);
      expect(response.ResponseMessage?.ResponseCode).to.be.equal(ResponseCodeType.NO_ERROR);
    }); 
  });
});
/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import assert from 'assert';
import {
    ResponseClassType, ResponseCodeType, LocationSourceType,
    MailboxTypeType, MapiPropertyTypeType, DistinguishedPropertySetType, DefaultShapeNamesType
} from '../models/enum.model';
import { FindPeopleResponseMessageType, GetPersonaResponseMessageType , ResolveNamesType, ResolveNamesResponseType, ResolveNamesResponseMessageType, ResolutionType, ArrayOfResolutionType} from "../models/persona.model";
import { EmailAddressType, BaseFolderIdType} from '../models/mail.model';
import { filterPimContactsByFolderIds, createContactItemFromPimContact } from "../utils/pimHelper";
//import parseAddress from 'email-addresses';
import { Request } from '@loopback/rest';
import { UserContext } from '../keepcomponent';
import { Logger } from '../utils';
import { isDevelopment, KeepPimContactsManager, PimContact, UserManager} from '@hcllabs/openclientkeepcomponent';


/**
 * Resolve names
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns resolve names result
 */
export async function ResolveNames(soapRequest: ResolveNamesType, request: Request): Promise<ResolveNamesResponseType> {
    assert(soapRequest instanceof ResolveNamesType);

    const userInfo = UserContext.getUserInfo(request);

    const response = new ResolveNamesResponseType();
    const msg = new ResolveNamesResponseMessageType();

        //const parsed = parseAddress.parseOneAddress(soapRequest.UnresolvedEntry) as parseAddress.ParsedMailbox;
        const unresolvedEntry = soapRequest.UnresolvedEntry;
        if (unresolvedEntry) {

            let pimItems: PimContact[] = [];
            try{
                let uEntry = unresolvedEntry;
                if (isDevelopment() && unresolvedEntry.endsWith('ngrok.io')) {
                    uEntry = UserManager.getInstance().getDevelopmentUserName(unresolvedEntry);
                }

                pimItems = await KeepPimContactsManager.getInstance().lookupContacts(userInfo, uEntry);

                //delete after fixing (LABS-2900) Keep issue for addresslookup folder id 
                //delete begin
                if(soapRequest.ParentFolderIds && soapRequest.ParentFolderIds.items){
                    if(pimItems?.length){ 
                        for(const pimItem of pimItems){
                             if(!(pimItem?.parentFolderIds?.length)){
                                    const foundPimItem = await KeepPimContactsManager.getInstance().getContact(pimItem.unid, userInfo);
                                    if(foundPimItem?.parentFolderIds?.length) pimItem.parentFolderIds = foundPimItem.parentFolderIds;
                                }
                        }
                    }
                }   
                //delete end
            } catch (err) {
                Logger.getInstance().error(`Error getting contacts: ${err}`)
                //throw err;
            }

            if(soapRequest.ParentFolderIds && soapRequest.ParentFolderIds.items){
                const folderIds: BaseFolderIdType[] = soapRequest.ParentFolderIds.items;
                pimItems = filterPimContactsByFolderIds(pimItems , folderIds);
            }
        
            if(pimItems && pimItems.length>1){
                msg.ResponseClass = ResponseClassType.WARNING;
                msg.ResponseCode = ResponseCodeType.ERROR_NAME_RESOLUTION_MULTIPLE_RESULTS;
                msg.MessageText = 'Multiple results are found.';   
            }else if(pimItems && pimItems.length>0){    
                const pimItem = pimItems[0];

                const resolution = new ResolutionType();
                //set Maibox
                resolution.Mailbox = new EmailAddressType();
                resolution.Mailbox.EmailAddress = pimItem.primaryEmail ?? '';
                resolution.Mailbox.Name = pimItem.fullName ? pimItem.fullName[0] : '';
                resolution.Mailbox.RoutingType = 'SMTP';
                resolution.Mailbox.MailboxType = MailboxTypeType.MAILBOX;

                
                //set Contact
                const contactDataShape: DefaultShapeNamesType = soapRequest.ContactDataShape ? soapRequest.ContactDataShape : DefaultShapeNamesType.DEFAULT;
                const returnFullContactData: boolean = soapRequest.ReturnFullContactData?.toString().toLowerCase().trim() === 'true';
                resolution.Contact = await createContactItemFromPimContact(pimItem, userInfo, request, returnFullContactData, contactDataShape);            

                msg.ResolutionSet = new ArrayOfResolutionType();
                msg.ResolutionSet.Resolution.push(resolution);
            } else {
                msg.ResponseClass = ResponseClassType.ERROR;
                msg.ResponseCode = ResponseCodeType.ERROR_NAME_RESOLUTION_NO_RESULTS;
                msg.MessageText = 'No results were found.';        
            }

        } else {
            msg.ResponseClass = ResponseClassType.ERROR;
            msg.ResponseCode = ResponseCodeType.ERROR_NAME_RESOLUTION_NO_RESULTS;
            msg.MessageText = 'No results were found.';
        }



    response.ResponseMessages.push(msg);
    return response;
}

// mock response for FindPeople
export const mockFindPeopleResponse: FindPeopleResponseMessageType = {
    ResponseClass: ResponseClassType.SUCCESS,
    ResponseCode: ResponseCodeType.NO_ERROR,
    TotalNumberOfPeopleInView: 10,
    FirstLoadedRowIndex: 0,
    FirstMatchingRowIndex: 1,
    People: {
        Persona: [
            {
                PersonaId: {
                    Id: 'AAUQAKinPDThxzhKgcfMnleMw5A='
                },
                DisplayName: 'Isaiah Langer',
                Title: 'Sales Rep',
                RelevanceScore: 2147483647
            },
            {
                PersonaId: {
                    Id: 'AAUQANX74GJCjBhIte3yNglKfWs='
                },
                DisplayName: 'Johanna Lorenz',
                Title: 'Senior Engineer',
                RelevanceScore: 2147483647
            },
            {
                PersonaId: {
                    Id: 'AAUQAKHb5J73MhFPnUK+plAqjTA='
                },
                DisplayName: 'Lee Gu',
                Title: 'Director',
                RelevanceScore: 2147483647
            },
            {
                PersonaId: {
                    Id: 'AAUQABTBBlVcpw5JuidYJ6tfO+Y='
                },
                DisplayName: 'Lidia Holloway',
                Title: 'Product Manager',
                RelevanceScore: 2147483647
            },
            {
                PersonaId: {
                    Id: 'AAUQAGOh/ljw4f1EnH1ragAAJkA='
                },
                DisplayName: 'liquan liu',
                RelevanceScore: 2147483647
            },
            {
                PersonaId: {
                    Id: 'AAUQAKmeVYlw3L9Mlrk0CnYSHeI='
                },
                DisplayName: 'Lynne Robbins',
                Title: 'Planner',
                RelevanceScore: 2147483647
            }
        ]
    }
};

// mock response for GetPersona
export const mockGetPersonaResponse: GetPersonaResponseMessageType = {
    ResponseClass: ResponseClassType.SUCCESS,
    ResponseCode: ResponseCodeType.NO_ERROR,
    Persona: {
        PersonaId: {
            Id: 'AAUQAGOh/ljw4f1EnH1ragAAJkA='
        },
        PersonaType: 'Person',
        CreationTime: new Date('2012-06-01T17:00:34Z'),
        DisplayName: 'Brian Johnson',
        DisplayNameFirstLast: 'Brian Johnson',
        DisplayNameLastFirst: 'Johnson Brian',
        FileAs: 'Johnson, Brian',
        FileAsId: 'None',
        GivenName: 'Brian',
        Surname: 'Johnsoon',
        CompanyName: 'Contoso',
        RelevanceScore: 4255550110,
        Attributions: {
            Attribution: [{
                Id: '0',
                SourceId: {
                    Id: 'AAMkA =',
                    ChangeKey: 'EQAAABY+'
                },
                DisplayName: 'Outlook',
                IsWritable: true,
                IsQuickContact: false,
                IsHidden: false,
                FolderId: {
                    Id: 'AAMkA=',
                    ChangeKey: 'AQAAAA=='
                }
            }]
        },
        DisplayNames: {
            StringAttributedValue: [{
                Value: 'Brian Johnson',
                Attributions: {
                    Attribution: ['att1', 'att2']
                }
            }]
        },
        FileAses: {
            StringAttributedValue: [{
                Value: 'Johnson, Brian',
                Attributions: {
                    Attribution: ['0']
                }
            }]
        },
        FileAsIds: {
            StringAttributedValue: [{
                Value: 'None',
                Attributions: {
                    Attribution: ['0']
                }
            }]
        },
        GivenNames: {
            StringAttributedValue: [{
                Value: 'Brian',
                Attributions: {
                    Attribution: ['0']
                }
            }]
        },
        Surnames: {
            StringAttributedValue: [{
                Value: 'Johnson',
                Attributions: {
                    Attribution: ['0']
                }
            }]
        },
        MobilePhones: {
            PhoneNumberAttributedValue: [{
                Value: {
                    Number: '(425)555-0110',
                    Type: 'Mobile'
                },
                Attributions: {
                    Attribution: ['0']
                }
            }]
        },
        CompanyNames: {
            StringAttributedValue: [{
                Value: 'Contoso',
                Attributions: {
                    Attribution: ['0']
                }
            }]
        },
        ItemLinkIds: {
            StringArrayAttributedValue: [
                {
                    Values: {
                        Value: ['val1', 'val2']
                    },
                    Attributions: {
                        Attribution: ['attr1', 'attr2']
                    }
                }
            ]
        },
        ExtendedProperties: {
            ExtendedPropertyAttributedValue: [{
                Value: {
                    Values: {
                        Value: ['val1', 'val2']
                    },
                    ExtendedFieldURI: {
                        PropertyId: 12435134,
                        PropertyType: MapiPropertyTypeType.APPLICATION_TIME,
                        DistinguishedPropertySetId: DistinguishedPropertySetType.CALENDAR_ASSISTANT
                    }
                },
                Attributions: {
                    Attribution: []
                }
            }]
        },
        BusinessPhoneNumbers: {
            PhoneNumberAttributedValue: [{
                Value: {
                    Number: '23413241234',
                    Type: 'Business'
                },
                Attributions: {
                    Attribution: []
                }
            }]
        },
        Emails1: {
            EmailAddressAttributedValue: [{
                Value: {
                    Name: 'leviastan@llqdev.onmicrosoft.com',
                    EmailAddress: 'leviastan@llqdev.onmicrosoft.com',
                    RoutingType: 'SMTP',
                    MailboxType: MailboxTypeType.MAILBOX,
                },
                Attributions: {
                    Attribution: []
                }
            }]
        },
        ImAddresses: {
            StringAttributedValue: [{
                Value: 'sip:leviastan@llqdev.onmicrosoft.com',
                Attributions: {
                    Attribution: []
                }
            }]
        },
        BusinessAddresses: {
            PostalAddressAttributedValue: [{
                Value: {
                    Country: 'United States',
                    Type: 'Business',
                    LocationSource: LocationSourceType.CONTACT
                },
                Attributions: {
                    Attribution: []
                }
            }]
        }
    }
};

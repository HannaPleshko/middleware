import { expect, sinon } from '@loopback/testlab';
import {
    PimItem, PimMessage, PimLabel, UserInfo, base64Decode, PimImportance, KeepPimMessageManager, KeepPimConstants,
    PimItemFormat, PimNoticeTypes, PimItemFactory, KeepPimManager, KeepPimCalendarManager
} from '@hcllabs/openclientkeepcomponent';
import {
    UnindexedFieldURIType, MessageDispositionType, BodyTypeType, ImportanceChoicesType, DefaultShapeNamesType,
    DistinguishedPropertySetType, MapiPropertyTypeType, MapiPropertyIds, DictionaryURIType, MeetingRequestTypeType,
    CalendarItemTypeType
} from '../../models/enum.model';
import {
    ItemResponseShapeType, MessageType, ItemType, ItemChangeType, ItemIdType, BodyType, SingleRecipientType,
    EmailAddressType, ArrayOfRecipientsType, NonEmptyArrayOfItemChangeDescriptionsType, DeleteItemFieldType, SetItemFieldType,
    ContactItemType, ExtendedPropertyType, AppendToItemFieldType, SharingMessageType, RoleMemberItemType, MeetingMessageType,
    MeetingRequestMessageType, MeetingResponseMessageType, MeetingCancellationMessageType, MimeContentType,
    NonEmptyArrayOfPathsToElementType,
    FolderIdType,
    CalendarItemType,
    AcceptItemType,
    DeclineItemType,
    TentativelyAcceptItemType,
    CancelCalendarItemType
} from '../../models/mail.model';
import {
    PathToExtendedFieldType, PathToIndexedFieldType, PathToUnindexedFieldType
} from '../../models/common.model';
import {
    EWSMessageManager, getEWSId, getKeepIdPair
} from '../../utils';
import { UserContext } from '../../keepcomponent';
import {
    generateTestCalendarItems,
    generateTestInboxLabel,
    getContext, stubCalendarManager, stubMessageManager, stubPimManager
} from '../unitTestHelper';
import { Request } from '@loopback/rest';
import { simpleParser, ParsedMail, Attachment, Headers, HeaderLines } from 'mailparser';

describe('EWSMessageManager tests', () => {

    // Sample MIME content
    const sampleMime = 'VXNlci1BZ2VudDogTWljcm9zb2Z0LU1hY091dGxvb2svMTYuMzguMjAwNjE0MDEKRGF0ZTogRnJpLCAxMCBKdWwgMjAyMCAxMTo0OToyMiAtMDQwMApTdWJqZWN0OiBIaQpGcm9tOiAiam9obi5kb2VAMjdlODMzNWFiYmQzLm5ncm9rLmlvIiA8am9obi5kb2VAMjdlODMzNWFiYmQzLm5ncm9rLmlvPgpUbzogRGF2aWQgS2VubmVkeSA8ZGF2aWQua2VubmVkeUBoY2wuY29tPgpNZXNzYWdlLUlEOiA8N0Q4QjY0QTgtOUJCNS00OTNDLTk5NTItMkMxREU4ODZGODI0QDI3ZTgzMzVhYmJkMy5uZ3Jvay5pbz4KUmVmZXJlbmNlczogPE9GMDM4RkI2MEIuMTg2QkZENUEtT04wMDI1ODczRC4wMDU1Q0MyQy00MzI1ODczRC4wMDU2MDM5NEBMb2NhbERvbWFpbj4gPE9GMjNFQjQ2QkUuMEU1NTA2MTUtT04wMDI1ODczRC4wMDU1RUEzMy00MzI1ODczRC4wMDU2MUFFQkBMb2NhbERvbWFpbj4KVGhyZWFkLVRvcGljOiBIaQpUaHJlYWQtSW5kZXg6IDU1Ck1pbWUtdmVyc2lvbjogMS4wCkNvbnRlbnQtdHlwZTogbXVsdGlwYXJ0L2FsdGVybmF0aXZlOwoJYm91bmRhcnk9IkJfMzY3NzIyNjU2Ml8xNzkxODczMDI0IgoKPiBUaGlzIG1lc3NhZ2UgaXMgaW4gTUlNRSBmb3JtYXQuIFNpbmNlIHlvdXIgbWFpbCByZWFkZXIgZG9lcyBub3QgdW5kZXJzdGFuZAp0aGlzIGZvcm1hdCwgc29tZSBvciBhbGwgb2YgdGhpcyBtZXNzYWdlIG1heSBub3QgYmUgbGVnaWJsZS4KCi0tQl8zNjc3MjI2NTYyXzE3OTE4NzMwMjQKQ29udGVudC10eXBlOiB0ZXh0L3BsYWluOwoJY2hhcnNldD0iVVRGLTgiCkNvbnRlbnQtdHJhbnNmZXItZW5jb2Rpbmc6IDdiaXQKCkhlbGxvIFRhbnlhCgoKLS1CXzM2NzcyMjY1NjJfMTc5MTg3MzAyNApDb250ZW50LXR5cGU6IHRleHQvaHRtbDsKCWNoYXJzZXQ9IlVURi04IgpDb250ZW50LXRyYW5zZmVyLWVuY29kaW5nOiBxdW90ZWQtcHJpbnRhYmxlCgo8aHRtbCB4bWxuczpvPTNEInVybjpzY2hlbWFzLW1pY3Jvc29mdC1jb206b2ZmaWNlOm9mZmljZSIgeG1sbnM6dz0zRCJ1cm46c2NoZW1hPQpzLW1pY3Jvc29mdC1jb206b2ZmaWNlOndvcmQiIHhtbG5zOm09M0QiaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2UvMjA9CjA0LzEyL29tbWwiIHhtbG5zPTNEImh0dHA6Ly93d3cudzMub3JnL1RSL1JFQy1odG1sNDAiPjxoZWFkPjxtZXRhIGh0dHAtZXF1aXY9M0RDPQpvbnRlbnQtVHlwZSBjb250ZW50PTNEInRleHQvaHRtbDsgY2hhcnNldD0zRHV0Zi04Ij48bWV0YSBuYW1lPTNER2VuZXJhdG9yIGNvbnRlbnQ9M0Q9CiJNaWNyb3NvZnQgV29yZCAxNSAoZmlsdGVyZWQgbWVkaXVtKSI+PHN0eWxlPjwhLS0KLyogRm9udCBEZWZpbml0aW9ucyAqLwpAZm9udC1mYWNlCgl7Zm9udC1mYW1pbHk6IkNhbWJyaWEgTWF0aCI7CglwYW5vc2UtMToyIDQgNSAzIDUgNCA2IDMgMiA0O30KQGZvbnQtZmFjZQoJe2ZvbnQtZmFtaWx5OkNhbGlicmk7CglwYW5vc2UtMToyIDE1IDUgMiAyIDIgNCAzIDIgNDt9Ci8qIFN0eWxlIERlZmluaXRpb25zICovCnAuTXNvTm9ybWFsLCBsaS5Nc29Ob3JtYWwsIGRpdi5Nc29Ob3JtYWwKCXttYXJnaW46MGluOwoJbWFyZ2luLWJvdHRvbTouMDAwMXB0OwoJZm9udC1zaXplOjExLjBwdDsKCWZvbnQtZmFtaWx5OiJDYWxpYnJpIixzYW5zLXNlcmlmO30Kc3Bhbi5FbWFpbFN0eWxlMTcKCXttc28tc3R5bGUtdHlwZTpwZXJzb25hbC1jb21wb3NlOwoJZm9udC1mYW1pbHk6IkNhbGlicmkiLHNhbnMtc2VyaWY7Cgljb2xvcjp3aW5kb3d0ZXh0O30KLk1zb0NocERlZmF1bHQKCXttc28tc3R5bGUtdHlwZTpleHBvcnQtb25seTsKCWZvbnQtZmFtaWx5OiJDYWxpYnJpIixzYW5zLXNlcmlmO30KQHBhZ2UgV29yZFNlY3Rpb24xCgl7c2l6ZTo4LjVpbiAxMS4waW47CgltYXJnaW46MS4waW4gMS4waW4gMS4waW4gMS4waW47fQpkaXYuV29yZFNlY3Rpb24xCgl7cGFnZTpXb3JkU2VjdGlvbjE7fQotLT48L3N0eWxlPjwvaGVhZD48Ym9keSBsYW5nPTNERU4tVVMgbGluaz0zRCIjMDU2M0MxIiB2bGluaz0zRCIjOTU0RjcyIj48ZGl2IGNsYXM9CnM9M0RXb3JkU2VjdGlvbjE+PHAgY2xhc3M9M0RNc29Ob3JtYWw+SGVsbG8gVGFueWE8bzpwPjwvbzpwPjwvcD48L2Rpdj48L2JvZHk+PC9oPQp0bWw+CgotLUJfMzY3NzIyNjU2Ml8xNzkxODczMDI0LS0KCg==';
    const testUser = "test.user@test.org";
    // Sample iCal response
    const iCalDeclineResponse = "BEGIN:VCALENDAR\r\nMETHOD:REPLY\r\nPRODID:Microsoft Exchange Server 2010\r\nVERSION:2.0\r\nBEGIN:VTIMEZONE\r\nTZID:Eastern Standard Time\r\nBEGIN:STANDARD\r\nDTSTART:16010101T020000\r\nTZOFFSETFROM:-0400\r\nTZOFFSETTO:-0500\r\nRRULE:FREQ=YEARLY;INTERVAL=1;BYDAY=1SU;BYMONTH=11\r\nEND:STANDARD\r\nBEGIN:DAYLIGHT\r\nDTSTART:16010101T020000\r\nTZOFFSETFROM:-0500\r\nTZOFFSETTO:-0400\r\nRRULE:FREQ=YEARLY;INTERVAL=1;BYDAY=2SU;BYMONTH=3\r\nEND:DAYLIGHT\r\nEND:VTIMEZONE\r\nBEGIN:VEVENT\r\nATTENDEE;PARTSTAT=DECLINED;CN=Roger Wilson:mailto:roger.wilson@hcl.com\r\nCOMMENT;LANGUAGE=en-US:\\n::DISCLAIMER::\\n________________________________\\n\r\n The contents of this e-mail and any attachment(s) are confidential and int\r\n ended for the named recipient(s) only. E-mail transmission is not guarante\r\n ed to be secure or error-free as information could be intercepted\\, corrup\r\n ted\\, lost\\, destroyed\\, arrive late or incomplete\\, or may contain viruse\r\n s in transmission. The e mail and its contents (with or without referred e\r\n rrors) shall therefore not attach any liability on the originator or HCL o\r\n r its affiliates. Views or opinions\\, if any\\, presented in this email are\r\n  solely those of the author and may not necessarily reflect the views or o\r\n pinions of HCL or its affiliates. Any form of reproduction\\, dissemination\r\n \\, copying\\, disclosure\\, modification\\, distribution and / or publication\r\n  of this message without the prior written consent of authorized represent\r\n ative of HCL is strictly prohibited. If you have received this email in er\r\n ror please delete it and notify the sender immediately. Before opening any\r\n  email and/or attachments\\, please check them for viruses and other defect\r\n s.\\n________________________________\\n\r\nUID:584ED245E02216FF002586B2006BCAE0-Lotus_Notes_Generated\r\nSUMMARY;LANGUAGE=en-US:Declined: test inviation...please counter\r\nDTSTART;TZID=Eastern Standard Time:20210415T130000\r\nDTEND;TZID=Eastern Standard Time:20210415T140000\r\nCLASS:PUBLIC\r\nPRIORITY:5\r\nDTSTAMP:20210409T193905Z\r\nTRANSP:OPAQUE\r\nSTATUS:CONFIRMED\r\nSEQUENCE:0\r\nLOCATION;LANGUAGE=en-US:test\r\nX-MICROSOFT-CDO-APPT-SEQUENCE:0\r\nX-MICROSOFT-CDO-OWNERAPPTID:0\r\nX-MICROSOFT-CDO-BUSYSTATUS:BUSY\r\nX-MICROSOFT-CDO-INTENDEDSTATUS:BUSY\r\nX-MICROSOFT-CDO-ALLDAYEVENT:FALSE\r\nX-MICROSOFT-CDO-IMPORTANCE:1\r\nX-MICROSOFT-CDO-INSTTYPE:0\r\nX-MICROSOFT-DONOTFORWARDMEETING:FALSE\r\nX-MICROSOFT-DISALLOW-COUNTER:FALSE\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n";
    // Sample iCal invitation
    const iCalInvitation = "BEGIN:VCALENDAR\r\nX-LOTUS-CHARSET:UTF-8\r\nVERSION:2.0\r\nPRODID:-//Lotus Development Corporation//NONSGML Notes 9.0.1//EN_S\r\nMETHOD:REQUEST\r\nBEGIN:VTIMEZONE\r\nTZID:America/New_York\r\nBEGIN:STANDARD\r\nDTSTART:19501105T020000\r\nTZOFFSETFROM:-0400\r\nTZOFFSETTO:-0500\r\nRRULE:FREQ=YEARLY;BYMINUTE=0;BYHOUR=2;BYDAY=1SU;BYMONTH=11\r\nEND:STANDARD\r\nBEGIN:DAYLIGHT\r\nDTSTART:19500312T020000\r\nTZOFFSETFROM:-0500\r\nTZOFFSETTO:-0400\r\nRRULE:FREQ=YEARLY;BYMINUTE=0;BYHOUR=2;BYDAY=2SU;BYMONTH=3\r\nEND:DAYLIGHT\r\nEND:VTIMEZONE\r\nBEGIN:VEVENT\r\nDTSTART;TZID=\"America/New_York\":20210225T190000\r\nDTEND;TZID=\"America/New_York\":20210225T200000\r\nTRANSP:OPAQUE\r\nDTSTAMP:20210225T154905Z\r\nSEQUENCE:0\r\nATTENDEE;ROLE=CHAIR;PARTSTAT=ACCEPTED;CN=\"Davek Mail/ProjectKeep\"\r\n ;RSVP=FALSE:mailto:david.kennedy@pnp-hcl.com\r\nATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE\r\n :mailto:rustyg.mail@quattro.rocks\r\nCLASS:PUBLIC\r\nDESCRIPTION;ALTREP=\"CID:<FFFF__=8FBB0C14DFC55C888f9e8a93df938690918c8FB@>\":\r\nSUMMARY:test\r\nORGANIZER;CN=\"Davek Mail/ProjectKeep\";SENT-BY=\"mailto\r\n :david.kennedy@pnp-hcl.com\"\r\n :mailto:Davek_Mail/ProjectKeep@irisa-nrpc01.ir3.wdc01.isc4sb.com\r\nUID:CA0CBE8538C6B44D002586870056DA18-Lotus_Notes_Generated\r\nX-LOTUS-BROADCAST:FALSE\r\nX-LOTUS-UPDATE-SEQ:1\r\nX-LOTUS-UPDATE-WISL:$S:1;$L:1;$B:1;$R:1;$E:1;$W:1;$O:1;$M:1;RequiredAttendees:1;INetRequiredNames:1;AltRequiredNames:1;StorageRequiredNames:1;OptionalAttendees:1;INetOptionalNames:1;AltOptionalNames:1;StorageOptionalNames:1;ApptUNIDURL:1;STUnyteConferenceURL:1;STUnyteConferenceID:1;SametimeType:1;WhiteBoardContent:1;STRoomName:1;$ECPAllowedItems:1\r\nX-LOTUS-NOTESVERSION:2\r\nX-LOTUS-NOTICETYPE:I\r\nX-LOTUS-APPTTYPE:3\r\nX-LOTUS-CHILD-UID:CA0CBE8538C6B44D002586870056DA18\r\nCOMMENT;ALTREP=CID:<FFFF__=0ABB0C8CDFE5A1D78f9e8a93df938690918c0AB@>:Go Buckeyes!\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n\r\n";

    // Generic Mock subclass of EWSService manager for testing protected functions
    class MockEWSManager extends EWSMessageManager {

        async createItem(
            item: ItemType, 
            userInfo: UserInfo, 
            request: Request, 
            disposition?: MessageDispositionType, 
            toLabel?: PimLabel, 
            mailboxId?: string
        ): Promise<ItemType[]> {
            return super.createItem(item, userInfo, request, disposition, toLabel, mailboxId);
        }

        async deleteItem(item: string | PimMessage, userInfo: UserInfo, mailboxId?: string, hardDelete = false): Promise<void> {
            return super.deleteItem(item, userInfo, mailboxId, hardDelete);
        }

        async updateItem(
            pimItem: PimMessage, 
            change: ItemChangeType, 
            userInfo: UserInfo, 
            request: Request, 
            toLabel?: PimLabel,
            mailboxId?: string
        ): Promise<ItemType | undefined> {
            return super.updateItem(pimItem, change, userInfo, request, toLabel, mailboxId);
        }

        async getItems(userInfo: UserInfo, request: Request, shape: ItemResponseShapeType, startIndex?: number, count?: number, fromLabel?: PimLabel): Promise<ItemType[]> {
            return super.getItems(userInfo, request, shape, startIndex, count, fromLabel);
        }

        addRequestedPropertiesToEWSItem(pimItem: PimItem, toItem: ItemType, shape?: ItemResponseShapeType): void {
            super.addRequestedPropertiesToEWSItem(pimItem, toItem, shape);
        }

        updateEWSItemFieldValue(
            item: ItemType, 
            pimMessage: PimMessage, 
            fieldId: UnindexedFieldURIType, 
            mailboxId?: string
        ): boolean {
            return super.updateEWSItemFieldValue(item, pimMessage, fieldId, mailboxId);
        }

        addRequestedPropertiesToEWSMessage(
            pimMessage: PimMessage | undefined, 
            message: MessageType, 
            mimeContent: string, 
            encoding = 'base64', 
            shape?: ItemResponseShapeType,
            mailboxId?: string
        ): Promise<void> {
            return super.addRequestedPropertiesToEWSMessage(pimMessage, message, mimeContent, encoding, shape, mailboxId);
        }

        async updateEWSItemFieldValueFromMime(
            message: MessageType, 
            parsed: ParsedMail, 
            pimMessage: PimMessage | undefined, 
            fieldIdentifier: UnindexedFieldURIType,
            mailboxId?: string
        ): Promise<boolean> {
            return super.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, fieldIdentifier, mailboxId);
        }

        pimItemFromEWSItem(item: MessageType, request: Request, existing?: object): PimMessage {
            return super.pimItemFromEWSItem(item, request, existing);
        }

        pimItemToEWSItem(
            pimItem: PimMessage, 
            userInfo: UserInfo, 
            request: Request, 
            shape: ItemResponseShapeType, 
            mailboxId?: string,
            parentFolderId?: string
        ): Promise<MessageType> {
            return super.pimItemToEWSItem(pimItem, userInfo, request, shape, mailboxId, parentFolderId);
        }

        buildMimeContentForMeetingMessage(pimItem: PimMessage): MimeContentType {
            return super.buildMimeContentForMeetingMessage(pimItem);
        }

        addRequestedPropertiesToEWSMeetingMessage(
            pimMessage: PimMessage, 
            meeting: MeetingMessageType, 
            shape: ItemResponseShapeType, 
            request: Request,
            mailboxId?: string
        ): void {
            return super.addRequestedPropertiesToEWSMeetingMessage(pimMessage, meeting, shape, request, mailboxId);
        }

        updateEWSItemFieldValueForMeeting(
            message: MeetingMessageType, 
            pimMessage: PimMessage, 
            fieldIdentifier: UnindexedFieldURIType, 
            request: Request, 
            iCalObject: any,
            mailboxId?: string
        ): boolean {
            return super.updateEWSItemFieldValueForMeeting(message, pimMessage, fieldIdentifier, request, iCalObject, mailboxId);
        }

    }

    describe('Test getInstance', () => {

        it('getInstance', function () {

            const manager = EWSMessageManager.getInstance();
            expect(manager).to.be.instanceof(EWSMessageManager);

            const manager2 = EWSMessageManager.getInstance();
            expect(manager).to.be.equal(manager2);
        });
    });

    describe('Test pimItemFromEWSItem', () => {
        it('empty MessageType', () => {
            const manager = new MockEWSManager();
            const message = new MessageType();
            const context = getContext(testUser, 'password');
            const pimMessage = manager.pimItemFromEWSItem(message, context.request);
            expect(pimMessage.unid).to.be.undefined();
            expect(pimMessage.createdDate).to.be.undefined();
        });

        it('common fields populated', () => {
            const message = new MessageType();
            const mailboxId = 'test@test.com';
            const itemEWSId = getEWSId('uid', mailboxId);
            message.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            const parentEWSId = getEWSId('parent', mailboxId);
            message.ParentFolderId = new FolderIdType(parentEWSId, `ck-${parentEWSId}`);
            const created = new Date();
            message.DateTimeCreated = created;
            const modified = new Date();
            message.LastModifiedTime = modified;
            const sent = new Date();
            message.DateTimeSent = sent;

            const body = new BodyType('BODY', BodyTypeType.TEXT);
            message.Body = body;

            message.Subject = 'SUBJECT';
            message.Size = 2048;
            message.IsReadReceiptRequested = true;

            const manager = new MockEWSManager();
            const context = getContext(testUser, 'password');

            let pimMessage = manager.pimItemFromEWSItem(message, context.request);
            const [pimId] = getKeepIdPair(message.ItemId.Id);
            expect(pimMessage.unid).to.be.equal(pimId);
            const parentIds = pimMessage.parentFolderIds;
            expect(parentIds).to.be.an.Array();
            if (parentIds) {
                const [parenFolderId] = getKeepIdPair(message.ParentFolderId.Id);
                expect(parentIds[0]).to.be.equal(parenFolderId);
            }
            expect(pimMessage.createdDate).to.be.eql(created);
            expect(pimMessage.lastModifiedDate).to.be.eql(modified);
            expect(pimMessage.sentDate).to.be.eql(sent);

            expect(pimMessage.body).to.be.equal(message.Body.Value);
            expect(pimMessage.bodyType).to.be.equal('text/plain; charset=utf-8');
            expect(pimMessage.subject).to.be.equal(message.Subject);
            expect(pimMessage.size).to.be.equal(message.Size);
            expect(pimMessage.returnReceipt).to.be.true();


            // Try with HTML body type
            message.Body.BodyType = BodyTypeType.HTML;
            pimMessage = manager.pimItemFromEWSItem(message, context.request);
            expect(pimMessage.bodyType).to.be.equal('text/html; charset=utf-8');
        });

        it('test from address', () => {
            const message = new MessageType();
            const itemEWSId = getEWSId('uid');
            message.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            const manager = new MockEWSManager();
            const context = getContext(testUser, 'password');

            // No From set
            let pimMessage = manager.pimItemFromEWSItem(message, context.request);
            expect(pimMessage.from).to.be.equal(testUser);

            // Set up a mailbox with no email address, but a name
            message.From = new SingleRecipientType();
            message.From.Mailbox = new EmailAddressType();
            message.From.Mailbox.Name = 'Regina Phalange';
            pimMessage = manager.pimItemFromEWSItem(message, context.request);
            expect(pimMessage.from).to.be.equal('Regina Phalange');

            // Try special case
            message.From = new SingleRecipientType();
            message.From.Mailbox = new EmailAddressType();
            message.From.Mailbox.Name = 'RustyG Mail';
            pimMessage = manager.pimItemFromEWSItem(message, context.request);
            expect(pimMessage.from).to.be.equal('rustyg.mail@mail.quattro.rocks');

            message.From = new SingleRecipientType();
            message.From.Mailbox = new EmailAddressType();
            message.From.Mailbox.EmailAddress = 'test@test.com';
            pimMessage = manager.pimItemFromEWSItem(message, context.request);
            expect(pimMessage.from).to.be.equal('test@test.com');

            // No name or email
            message.From = new SingleRecipientType();
            message.From.Mailbox = new EmailAddressType();
            pimMessage = manager.pimItemFromEWSItem(message, context.request);
            expect(pimMessage.from).to.be.equal(testUser);
        });

        it('test to recipients', () => {
            const message = new MessageType();
            const itemEWSId = getEWSId('uid');
            message.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            const manager = new MockEWSManager();
            const context = getContext(testUser, 'password');

            message.ToRecipients = new ArrayOfRecipientsType();
            let mailbox = new EmailAddressType();
            mailbox.Name = 'Regina Phalange';
            message.ToRecipients.Mailbox.push(mailbox);

            let pimMessage = manager.pimItemFromEWSItem(message, context.request);
            expect(pimMessage.to).to.be.eql(['Regina Phalange']);

            message.ToRecipients = new ArrayOfRecipientsType();
            mailbox = new EmailAddressType();
            mailbox.Name = 'RustyG Mail';
            message.ToRecipients.Mailbox.push(mailbox);

            pimMessage = manager.pimItemFromEWSItem(message, context.request);
            expect(pimMessage.to).to.be.eql(['rustyg.mail@mail.quattro.rocks']);

            message.ToRecipients = new ArrayOfRecipientsType();
            mailbox = new EmailAddressType();
            mailbox.EmailAddress = 'test@test.com';
            message.ToRecipients.Mailbox.push(mailbox);

            pimMessage = manager.pimItemFromEWSItem(message, context.request);
            expect(pimMessage.to).to.be.eql(['test@test.com']);

            // No name or email
            message.ToRecipients = new ArrayOfRecipientsType();
            mailbox = new EmailAddressType();
            message.ToRecipients.Mailbox.push(mailbox);

            pimMessage = manager.pimItemFromEWSItem(message, context.request);
            expect(pimMessage.to).to.be.undefined();
        });

        it('test bcc recipients', () => {
            const message = new MessageType();
            const itemEWSId = getEWSId('uid');
            message.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            const manager = new MockEWSManager();
            const context = getContext(testUser, 'password');

            message.BccRecipients = new ArrayOfRecipientsType();
            let mailbox = new EmailAddressType();
            mailbox.Name = 'Regina Phalange';
            message.BccRecipients.Mailbox.push(mailbox);

            let pimMessage = manager.pimItemFromEWSItem(message, context.request);
            expect(pimMessage.bcc).to.be.eql(['Regina Phalange']);

            message.BccRecipients = new ArrayOfRecipientsType();
            mailbox = new EmailAddressType();
            mailbox.Name = 'RustyG Mail';
            message.BccRecipients.Mailbox.push(mailbox);

            pimMessage = manager.pimItemFromEWSItem(message, context.request);
            expect(pimMessage.bcc).to.be.eql(['rustyg.mail@mail.quattro.rocks']);

            message.BccRecipients = new ArrayOfRecipientsType();
            mailbox = new EmailAddressType();
            mailbox.EmailAddress = 'test@test.com';
            message.BccRecipients.Mailbox.push(mailbox);

            pimMessage = manager.pimItemFromEWSItem(message, context.request);
            expect(pimMessage.bcc).to.be.eql(['test@test.com']);

            // No name or email
            message.BccRecipients = new ArrayOfRecipientsType();
            mailbox = new EmailAddressType();
            message.BccRecipients.Mailbox.push(mailbox);

            pimMessage = manager.pimItemFromEWSItem(message, context.request);
            expect(pimMessage.bcc).to.be.undefined();
        });

        it('test cc recipients', () => {
            const message = new MessageType();
            const itemEWSId = getEWSId('uid', 'test@test.com');
            message.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            const manager = new MockEWSManager();
            const context = getContext(testUser, 'password');

            message.CcRecipients = new ArrayOfRecipientsType();
            let mailbox = new EmailAddressType();
            mailbox.Name = 'Regina Phalange';
            message.CcRecipients.Mailbox.push(mailbox);

            let pimMessage = manager.pimItemFromEWSItem(message, context.request);
            expect(pimMessage.cc).to.be.eql(['Regina Phalange']);

            message.CcRecipients = new ArrayOfRecipientsType();
            mailbox = new EmailAddressType();
            mailbox.Name = 'RustyG Mail';
            message.CcRecipients.Mailbox.push(mailbox);

            pimMessage = manager.pimItemFromEWSItem(message, context.request);
            expect(pimMessage.cc).to.be.eql(['rustyg.mail@mail.quattro.rocks']);

            message.CcRecipients = new ArrayOfRecipientsType();
            mailbox = new EmailAddressType();
            mailbox.EmailAddress = 'test@test.com';
            message.CcRecipients.Mailbox.push(mailbox);

            pimMessage = manager.pimItemFromEWSItem(message, context.request);
            expect(pimMessage.cc).to.be.eql(['test@test.com']);

            // No name or email
            message.CcRecipients = new ArrayOfRecipientsType();
            mailbox = new EmailAddressType();
            message.CcRecipients.Mailbox.push(mailbox);

            pimMessage = manager.pimItemFromEWSItem(message, context.request);
            expect(pimMessage.cc).to.be.undefined();
        });

        it('test importance', () => {
            const message = new MessageType();
            const itemEWSId = getEWSId('uid');
            message.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            message.Importance = ImportanceChoicesType.HIGH;

            const manager = new MockEWSManager();
            const context = getContext(testUser, 'password');

            let pimMessage = manager.pimItemFromEWSItem(message, context.request);
            expect(pimMessage.importance).to.be.equal(PimImportance.HIGH);

            message.Importance = ImportanceChoicesType.LOW;
            pimMessage = manager.pimItemFromEWSItem(message, context.request);
            expect(pimMessage.importance).to.be.equal(PimImportance.LOW);

            message.Importance = ImportanceChoicesType.NORMAL;
            pimMessage = manager.pimItemFromEWSItem(message, context.request);
            expect(pimMessage.importance).to.be.equal(PimImportance.NONE);

        });

    });

    describe('Test fieldsForDefaultShape', () => {

        it('test default fields', () => {
            const manager = new EWSMessageManager();
            const fields = manager.fieldsForDefaultShape();
            expect(fields).to.be.an.Array();
            // Verify size is correct
            expect(fields.length).to.be.equal(14);
            // Check a few of the fields
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_DATE_TIME_CREATED);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_SENSITIVITY);
            expect(fields).to.containEql(UnindexedFieldURIType.MESSAGE_IS_DELIVERY_RECEIPT_REQUESTED);
        });

    });

    describe('Test fieldsForAllPropertiesShape', () => {
        it('test all properties fields', () => {
            const manager = new EWSMessageManager();
            const fields = manager.fieldsForAllPropertiesShape();
            expect(fields).to.be.an.Array();
            // Verify the size is correct
            expect(fields.length).to.be.equal(51);
            // Check a few of the fields
            expect(fields).to.containEql(UnindexedFieldURIType.MESSAGE_CONVERSATION_INDEX);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_EFFECTIVE_RIGHTS);
            expect(fields).to.containEql(UnindexedFieldURIType.MESSAGE_REFERENCES);
        });
    });


    describe('Test getItems', () => {

        // Clean up the stubs
        afterEach(() => {
            sinon.restore(); // Reset stubs
        });

        it('Expected results', async () => {
            const parentLabel = generateTestInboxLabel();
            const pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';
            pimMessage.subject = 'SUBJECT';
            pimMessage.body = 'BODY';
            pimMessage.parentFolderIds = [parentLabel.folderId];
            const messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            messageManagerStub.getMailMessages.resolves([pimMessage]);
            messageManagerStub.getMimeMessage.resolves('mime content');
            stubMessageManager(messageManagerStub);

            const manager = EWSMessageManager.getInstance();
            const userInfo = new UserContext();
            const context = getContext(testUser, 'password');
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            const mailboxId = 'test@test.com';
            let messages = await manager.getItems(userInfo, context.request, shape, 0, 10, parentLabel, mailboxId);

            expect(messages).to.be.an.Array();
            expect(messages.length).to.be.equal(1);
            let message = messages[0];
            const itemEWSId = getEWSId(pimMessage.unid, mailboxId);

            expect(message.ItemId?.Id).to.be.equal(itemEWSId);

            // No labelId
            messages = await manager.getItems(userInfo, context.request, shape);

            expect(messages).to.be.an.Array();
            expect(messages.length).to.be.equal(1);
            message = messages[0];
            expect(message.ItemId?.Id).to.be.equal(getEWSId(pimMessage.unid));
        });

        it('getMailMessages throws an error', async () => {
            const parentLabel = generateTestInboxLabel();
            const pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';
            pimMessage.subject = 'SUBJECT';
            pimMessage.body = 'BODY';
            pimMessage.parentFolderIds = [parentLabel.folderId];
            const messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            messageManagerStub.getMailMessages.throws();
            messageManagerStub.getMimeMessage.resolves('mime content');
            stubMessageManager(messageManagerStub);

            const manager = EWSMessageManager.getInstance();
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

        // Clean up the stubs
        afterEach(() => {
            sinon.restore(); // Reset stubs
        });

        it('test set field update', async () => {
            const pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';
            pimMessage.subject = 'SUBJECT';
            pimMessage.body = 'BODY';
            pimMessage.parentFolderIds = ['PARENT'];
            const messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            messageManagerStub.getMailMessages.throws();
            messageManagerStub.getMimeMessage.resolves('mime content');
            stubMessageManager(messageManagerStub);

            const manager = new MockEWSManager();
            const userInfo = new UserContext();

            const changes = new ItemChangeType();
            changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

            const fieldURIChange = new SetItemFieldType()
            fieldURIChange.FieldURI = new PathToUnindexedFieldType();
            fieldURIChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            fieldURIChange.Item = new MessageType();
            const mailboxId = 'test@test.com';
            const itemEWSId = getEWSId(pimMessage.unid, mailboxId);
            fieldURIChange.Item.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            fieldURIChange.Item.Subject = 'UPDATED-SUBJECT';

            changes.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            changes.Updates.push(fieldURIChange);

            const extendedFieldChange = new SetItemFieldType();
            extendedFieldChange.ExtendedFieldURI = new PathToExtendedFieldType();
            extendedFieldChange.ExtendedFieldURI.DistinguishedPropertySetId = DistinguishedPropertySetType.COMMON;
            extendedFieldChange.ExtendedFieldURI.PropertyId = MapiPropertyIds.REMINDER_SET;
            extendedFieldChange.ExtendedFieldURI.PropertyType = MapiPropertyTypeType.BOOLEAN;

            extendedFieldChange.Item = new MessageType();
            extendedFieldChange.Item.ExtendedProperty = [];
            const prop = new ExtendedPropertyType();
            prop.ExtendedFieldURI = new PathToExtendedFieldType();
            prop.ExtendedFieldURI.DistinguishedPropertySetId = DistinguishedPropertySetType.COMMON;
            prop.ExtendedFieldURI.PropertyId = MapiPropertyIds.REMINDER_SET;
            prop.ExtendedFieldURI.PropertyType = MapiPropertyTypeType.BOOLEAN;
            prop.Value = 'true';
            extendedFieldChange.Item.ExtendedProperty.push(prop);
            changes.Updates.push(extendedFieldChange);

            const indextedFieldChange = new SetItemFieldType();
            indextedFieldChange.IndexedFieldURI = new PathToIndexedFieldType();
            indextedFieldChange.IndexedFieldURI.FieldURI = DictionaryURIType.CONTACTS_EMAIL_ADDRESS;
            indextedFieldChange.Item = new MessageType();
            changes.Updates.push(indextedFieldChange);

            let message = await manager.updateItem(pimMessage, changes, userInfo, context.request, undefined, mailboxId);
            expect(message).to.not.be.undefined();
            expect(message?.ItemId?.Id).to.be.equal(itemEWSId);

            const label = PimItemFactory.newPimLabel({FolderId: 'PARENT-LABEL'});

            // Try with parent folderId
            message = await manager.updateItem(pimMessage, changes, userInfo, context.request, label, mailboxId);
            expect(message).to.not.be.undefined();
            const parentEWSId = getEWSId('PARENT-LABEL', mailboxId);
            expect(message?.ParentFolderId?.Id).to.be.equal(parentEWSId);

            // Try with no existing parent folder
            pimMessage.parentFolderIds = undefined;
            message = await manager.updateItem(pimMessage, changes, userInfo, context.request, label, mailboxId);
            expect(message).to.not.be.undefined();
            expect(message?.ParentFolderId?.Id).to.be.equal(parentEWSId);

            // Try with no parent folder at all
            pimMessage.parentFolderIds = undefined;
            message = await manager.updateItem(pimMessage, changes, userInfo, context.request);
            expect(message).to.not.be.undefined();
            expect(message?.ParentFolderId?.Id).to.be.undefined();
        });

        it('test append field update', async () => {
            const pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';
            pimMessage.subject = 'SUBJECT';
            pimMessage.body = 'BODY';
            pimMessage.parentFolderIds = ['PARENT'];
            const messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            messageManagerStub.getMailMessages.throws();
            messageManagerStub.getMimeMessage.resolves('mime content');
            stubMessageManager(messageManagerStub);

            const manager = new MockEWSManager();
            const userInfo = new UserContext();

            let changes = new ItemChangeType();
            changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

            // Try body
            let fieldURIChangeBody = new AppendToItemFieldType()
            fieldURIChangeBody.FieldURI = new PathToUnindexedFieldType();
            fieldURIChangeBody.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_BODY;
            fieldURIChangeBody.Item = new MessageType();
            const itemEWSId = getEWSId(pimMessage.unid);
            fieldURIChangeBody.Item.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            fieldURIChangeBody.Item.Body = new BodyType('APPENDED-BODY', BodyTypeType.TEXT)

            changes.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            changes.Updates.push(fieldURIChangeBody);


            // Try subject (not supported)
            const fieldURIChangeSubject = new AppendToItemFieldType()
            fieldURIChangeSubject.FieldURI = new PathToUnindexedFieldType();
            fieldURIChangeSubject.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            fieldURIChangeSubject.Item = new MessageType();
            fieldURIChangeSubject.Item.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            fieldURIChangeSubject.Item.Subject = 'APPENDED-SUBJECT';

            changes.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            changes.Updates.push(fieldURIChangeSubject);

            // No extended fields are supported
            const extendedFieldChange = new AppendToItemFieldType();
            extendedFieldChange.ExtendedFieldURI = new PathToExtendedFieldType();
            extendedFieldChange.ExtendedFieldURI.DistinguishedPropertySetId = DistinguishedPropertySetType.COMMON;
            extendedFieldChange.ExtendedFieldURI.PropertyId = MapiPropertyIds.EMAIL1_DISPLAY_NAME;
            extendedFieldChange.ExtendedFieldURI.PropertyType = MapiPropertyTypeType.STRING;

            extendedFieldChange.Item = new MessageType();
            extendedFieldChange.Item.ExtendedProperty = [];
            const prop = new ExtendedPropertyType();
            prop.ExtendedFieldURI = new PathToExtendedFieldType();
            prop.ExtendedFieldURI.DistinguishedPropertySetId = DistinguishedPropertySetType.COMMON;
            prop.ExtendedFieldURI.PropertyId = MapiPropertyIds.EMAIL1_DISPLAY_NAME;
            prop.ExtendedFieldURI.PropertyType = MapiPropertyTypeType.STRING;
            prop.Value = 'append';
            extendedFieldChange.Item.ExtendedProperty.push(prop);
            changes.Updates.push(extendedFieldChange);

            let message = await manager.updateItem(pimMessage, changes, userInfo, context.request);
            expect(message).to.not.be.undefined();
            expect(message?.ItemId?.Id).to.be.equal(itemEWSId);

            // Try with no body
            changes = new ItemChangeType();
            changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

            fieldURIChangeBody = new AppendToItemFieldType()
            fieldURIChangeBody.FieldURI = new PathToUnindexedFieldType();
            fieldURIChangeBody.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_BODY;
            fieldURIChangeBody.Item = new MessageType();
            fieldURIChangeBody.Item.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);

            changes.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            changes.Updates.push(fieldURIChangeBody);

            message = await manager.updateItem(pimMessage, changes, userInfo, context.request);
            expect(message).to.not.be.undefined();
            expect(message?.ItemId?.Id).to.be.equal(itemEWSId);
        });


        it('test delete field update', async () => {
            const pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';
            pimMessage.subject = 'SUBJECT';
            pimMessage.body = 'BODY';
            pimMessage.parentFolderIds = ['PARENT'];
            const messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            messageManagerStub.getMailMessages.throws();
            messageManagerStub.getMimeMessage.resolves('mime content');
            stubMessageManager(messageManagerStub);

            const manager = new MockEWSManager();
            const userInfo = new UserContext();

            const changes = new ItemChangeType();
            changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();
            const fieldURIChange = new DeleteItemFieldType();
            fieldURIChange.FieldURI = new PathToUnindexedFieldType();
            fieldURIChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            const mailboxId = 'test@test.com';
            const itemEWSId = getEWSId(pimMessage.unid, mailboxId);
            changes.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            changes.Updates.push(fieldURIChange);

            const extendedFieldChange = new DeleteItemFieldType();
            extendedFieldChange.ExtendedFieldURI = new PathToExtendedFieldType();
            extendedFieldChange.ExtendedFieldURI.DistinguishedPropertySetId = DistinguishedPropertySetType.COMMON;
            extendedFieldChange.ExtendedFieldURI.PropertyId = MapiPropertyIds.REMINDER_SET;
            extendedFieldChange.ExtendedFieldURI.PropertyType = MapiPropertyTypeType.BOOLEAN;
            changes.Updates.push(extendedFieldChange);

            const indextedFieldChange = new DeleteItemFieldType();
            indextedFieldChange.IndexedFieldURI = new PathToIndexedFieldType();
            indextedFieldChange.IndexedFieldURI.FieldURI = DictionaryURIType.CONTACTS_EMAIL_ADDRESS;
            changes.Updates.push(indextedFieldChange);

            const message = await manager.updateItem(pimMessage, changes, userInfo, context.request, undefined, mailboxId);
            expect(message).to.not.be.undefined();
            expect(message?.ItemId?.Id).to.be.equal(itemEWSId);
        });

        it('test different item types', async () => {
            const pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';
            pimMessage.subject = 'SUBJECT';
            pimMessage.body = 'BODY';
            pimMessage.parentFolderIds = ['PARENT'];
            const messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            messageManagerStub.getMailMessages.throws();
            messageManagerStub.getMimeMessage.resolves('mime content');
            stubMessageManager(messageManagerStub);

            const manager = new MockEWSManager();
            const userInfo = new UserContext();

            // Message type
            let changes = new ItemChangeType();
            changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

            let fieldURIChange = new SetItemFieldType()
            fieldURIChange.FieldURI = new PathToUnindexedFieldType();
            fieldURIChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            fieldURIChange.Message = new MessageType();
            const itemEWSId = getEWSId(pimMessage.unid);
            fieldURIChange.Message.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            fieldURIChange.Message.Subject = 'UPDATED-SUBJECT';

            changes.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            changes.Updates.push(fieldURIChange);

            let message = await manager.updateItem(pimMessage, changes, userInfo, context.request);
            expect(message).to.not.be.undefined();
            expect(message?.ItemId?.Id).to.be.equal(itemEWSId);

            // SharingMessage type
            changes = new ItemChangeType();
            changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

            fieldURIChange = new SetItemFieldType()
            fieldURIChange.FieldURI = new PathToUnindexedFieldType();
            fieldURIChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            fieldURIChange.SharingMessage = new SharingMessageType();
            fieldURIChange.SharingMessage.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            fieldURIChange.SharingMessage.Subject = 'UPDATED-SUBJECT';

            changes.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            changes.Updates.push(fieldURIChange);

            message = await manager.updateItem(pimMessage, changes, userInfo, context.request);
            expect(message).to.not.be.undefined();
            expect(message?.ItemId?.Id).to.be.equal(itemEWSId);

            // RoleMember type??
            changes = new ItemChangeType();
            changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

            fieldURIChange = new SetItemFieldType()
            fieldURIChange.FieldURI = new PathToUnindexedFieldType();
            fieldURIChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            fieldURIChange.RoleMember = new RoleMemberItemType();
            fieldURIChange.RoleMember.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            fieldURIChange.RoleMember.Subject = 'UPDATED-SUBJECT';

            changes.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            changes.Updates.push(fieldURIChange);

            message = await manager.updateItem(pimMessage, changes, userInfo, context.request);
            expect(message).to.not.be.undefined();
            expect(message?.ItemId?.Id).to.be.equal(itemEWSId);

            // MeetingMessage type
            changes = new ItemChangeType();
            changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

            fieldURIChange = new SetItemFieldType()
            fieldURIChange.FieldURI = new PathToUnindexedFieldType();
            fieldURIChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            fieldURIChange.MeetingMessage = new MeetingMessageType();
            fieldURIChange.MeetingMessage.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            fieldURIChange.MeetingMessage.Subject = 'UPDATED-SUBJECT';

            changes.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            changes.Updates.push(fieldURIChange);

            message = await manager.updateItem(pimMessage, changes, userInfo, context.request);
            expect(message).to.not.be.undefined();
            expect(message?.ItemId?.Id).to.be.equal(itemEWSId);

            // MeetingRequest type
            changes = new ItemChangeType();
            changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

            fieldURIChange = new SetItemFieldType()
            fieldURIChange.FieldURI = new PathToUnindexedFieldType();
            fieldURIChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            fieldURIChange.MeetingRequest = new MeetingRequestMessageType();
            fieldURIChange.MeetingRequest.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            fieldURIChange.MeetingRequest.Subject = 'UPDATED-SUBJECT';

            changes.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            changes.Updates.push(fieldURIChange);

            message = await manager.updateItem(pimMessage, changes, userInfo, context.request);
            expect(message).to.not.be.undefined();
            expect(message?.ItemId?.Id).to.be.equal(itemEWSId);

            // MeetingResponse type
            changes = new ItemChangeType();
            changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

            fieldURIChange = new SetItemFieldType()
            fieldURIChange.FieldURI = new PathToUnindexedFieldType();
            fieldURIChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            fieldURIChange.MeetingResponse = new MeetingResponseMessageType();
            fieldURIChange.MeetingResponse.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            fieldURIChange.MeetingResponse.Subject = 'UPDATED-SUBJECT';

            changes.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            changes.Updates.push(fieldURIChange);

            message = await manager.updateItem(pimMessage, changes, userInfo, context.request);
            expect(message).to.not.be.undefined();
            expect(message?.ItemId?.Id).to.be.equal(itemEWSId);

            // MeetingCancellation type
            changes = new ItemChangeType();
            changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

            fieldURIChange = new SetItemFieldType()
            fieldURIChange.FieldURI = new PathToUnindexedFieldType();
            fieldURIChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            fieldURIChange.MeetingCancellation = new MeetingCancellationMessageType();
            fieldURIChange.MeetingCancellation.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            fieldURIChange.MeetingCancellation.Subject = 'UPDATED-SUBJECT';

            changes.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            changes.Updates.push(fieldURIChange);

            message = await manager.updateItem(pimMessage, changes, userInfo, context.request);
            expect(message).to.not.be.undefined();
            expect(message?.ItemId?.Id).to.be.equal(itemEWSId);

            // SharingMessage type
            changes = new ItemChangeType();
            changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

            fieldURIChange = new SetItemFieldType()
            fieldURIChange.FieldURI = new PathToUnindexedFieldType();
            fieldURIChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            fieldURIChange.MeetingCancellation = new MeetingCancellationMessageType();
            fieldURIChange.MeetingCancellation.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            fieldURIChange.MeetingCancellation.Subject = 'UPDATED-SUBJECT';

            changes.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            changes.Updates.push(fieldURIChange);

            message = await manager.updateItem(pimMessage, changes, userInfo, context.request);
            expect(message).to.not.be.undefined();
            expect(message?.ItemId?.Id).to.be.equal(itemEWSId);

            // Non-meeting type
            changes = new ItemChangeType();
            changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

            fieldURIChange = new SetItemFieldType()
            fieldURIChange.FieldURI = new PathToUnindexedFieldType();
            fieldURIChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            fieldURIChange.Contact = new ContactItemType();
            fieldURIChange.Contact.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            fieldURIChange.Contact.Subject = 'UPDATED-SUBJECT';

            changes.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            changes.Updates.push(fieldURIChange);

            message = await manager.updateItem(pimMessage, changes, userInfo, context.request);
            expect(message).to.not.be.undefined();
            expect(message?.ItemId?.Id).to.be.equal(itemEWSId);

        });

        describe('test setting follow up flag', () => {
            let pimMessage: PimMessage;
            let manager: MockEWSManager;
            let userInfo: UserContext;

            const propertyTag = '0x1090';

            beforeEach(() => {
                pimMessage = PimItemFactory.newPimMessage({});
                pimMessage.unid = 'UNID';
                pimMessage.subject = 'SUBJECT';
                pimMessage.body = 'BODY';
                pimMessage.parentFolderIds = ['PARENT'];
                const messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
                messageManagerStub.getMailMessages.throws();
                messageManagerStub.getMimeMessage.resolves('mime content');
                stubMessageManager(messageManagerStub);

                manager = new MockEWSManager();
                userInfo = new UserContext();

            });

            afterEach(() => {
                sinon.restore(); 
            });

            it('set follow up flag on', async () => {
                const changes = new ItemChangeType();
                changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

                /* 
                <t:SetItemField>
                    <t:ExtendedFieldURI PropertyTag="0x1090" PropertyType="Integer"/>
                    <t:Message>
                        <t:ExtendedProperty>
                            <t:ExtendedFieldURI PropertyTag="0x1090" PropertyType="Integer"/>
                            <t:Value>2</t:Value>
                        </t:ExtendedProperty>
                    </t:Message>
                </t:SetItemField>
                */
                const fieldURIChange = new SetItemFieldType()
                const extFieldType = new PathToExtendedFieldType();
                extFieldType.PropertyTag = propertyTag;
                extFieldType.PropertyType = MapiPropertyTypeType.INTEGER;
                fieldURIChange.ExtendedFieldURI = extFieldType;
                fieldURIChange.Item = new MessageType();
                const propertyType = new ExtendedPropertyType();
                propertyType.ExtendedFieldURI = extFieldType;
                propertyType.Value = "2";
                fieldURIChange.Item.ExtendedProperty = [propertyType];
                const mailboxId = 'test@test.com';

                const itemEWSId = getEWSId(pimMessage.unid, mailboxId);
                changes.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
                changes.Updates.push(fieldURIChange);

                const message = await manager.updateItem(pimMessage, changes, userInfo, context.request, undefined, mailboxId);
                expect(message).to.not.be.undefined();
                expect(message?.ItemId?.Id).to.be.equal(itemEWSId);
                expect(pimMessage.isFlaggedForFollowUp).to.be.true();

            });

            it('set follow up flag off with value', async () => {
                pimMessage.isFlaggedForFollowUp = true; // Start with this set on

                const changes = new ItemChangeType();
                changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

                /* 
                <t:SetItemField>
                    <t:ExtendedFieldURI PropertyTag="0x1090" PropertyType="Integer"/>
                    <t:Message>
                        <t:ExtendedProperty>
                            <t:ExtendedFieldURI PropertyTag="0x1090" PropertyType="Integer"/>
                            <t:Value>1</t:Value>
                        </t:ExtendedProperty>
                    </t:Message>
                </t:SetItemField>
                */
                const fieldURIChange = new SetItemFieldType()
                const extFieldType = new PathToExtendedFieldType();
                extFieldType.PropertyTag = propertyTag;
                extFieldType.PropertyType = MapiPropertyTypeType.INTEGER;
                fieldURIChange.ExtendedFieldURI = extFieldType;
                fieldURIChange.Item = new MessageType();
                const propertyType = new ExtendedPropertyType();
                propertyType.ExtendedFieldURI = extFieldType;
                propertyType.Value = "1";
                fieldURIChange.Item.ExtendedProperty = [propertyType];

                const itemEWSId = getEWSId(pimMessage.unid);
                changes.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
                changes.Updates.push(fieldURIChange);

                const message = await manager.updateItem(pimMessage, changes, userInfo, context.request);
                expect(message).to.not.be.undefined();
                expect(message?.ItemId?.Id).to.be.equal(itemEWSId);
                expect(pimMessage.isFlaggedForFollowUp).to.be.false();

            });

            it('set follow up flag off with delete field', async () => {
                pimMessage.isFlaggedForFollowUp = true; // Start with this set on

                const changes = new ItemChangeType();
                changes.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

                /* 
                <t:DeleteItemField>
                    <t:ExtendedFieldURI PropertyTag="0x1090" PropertyType="Integer"/>
                </t:DeleteItemField>
                */
                const deleteItem = new DeleteItemFieldType(); 
                const extFieldType = new PathToExtendedFieldType();
                extFieldType.PropertyTag = propertyTag;
                extFieldType.PropertyType = MapiPropertyTypeType.INTEGER;
                deleteItem.ExtendedFieldURI = extFieldType;

                const mailboxId = 'test@test.com';
                const itemEWSId = getEWSId(pimMessage.unid, mailboxId);
                changes.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
                changes.Updates.push(deleteItem);

                const message = await manager.updateItem(pimMessage, changes, userInfo, context.request, undefined, mailboxId);
                expect(message).to.not.be.undefined();
                expect(message?.ItemId?.Id).to.be.equal(itemEWSId);
                expect(pimMessage.isFlaggedForFollowUp).to.be.false();

            });
        });
    });

    describe('Test createItem', () => {

        // Clean up the stubs
        afterEach(() => {
            sinon.restore(); // Reset stubs
        });

        it('test create mime message, no label', async () => {
            const messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            messageManagerStub.createMimeMessage.resolves('UID');
            messageManagerStub.moveMessages.resolves({});
            messageManagerStub.deleteMessage.resolves({});
            messageManagerStub.moveMessages.resolves();

            stubMessageManager(messageManagerStub);

            const manager = new MockEWSManager();
            const userInfo = new UserContext();

            const message = new MessageType();
            const itemEWSId = getEWSId('UID');
            message.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            message.MimeContent = new MimeContentType(sampleMime, 'utf8');

            const context = getContext(testUser, 'password');

            const newMessages = await manager.createItem(message, userInfo, context.request);
            expect(newMessages).to.not.be.undefined();
            expect(newMessages.length).to.be.equal(1);
            expect(newMessages[0]?.ItemId?.Id).to.be.equal(itemEWSId);

        });

        it('test create mime message, with a label', async () => {

            const messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            messageManagerStub.createMimeMessage.resolves('UID');
            messageManagerStub.moveMessages.resolves({movedIds: [{status: 200, message: 'success', unid: 'UID'}]});
            messageManagerStub.deleteMessage.resolves({});

            stubMessageManager(messageManagerStub);

            const manager = new MockEWSManager();
            const userInfo = new UserContext();
            const toLabel = PimItemFactory.newPimLabel({}, PimItemFormat.DOCUMENT);
            toLabel.view = KeepPimConstants.INBOX;
            toLabel.unid = 'TO-LABEL';

            const message = new MessageType();
            const mailboxId = 'test@test.com';
            const itemEWSId = getEWSId('UID', mailboxId);
            message.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            message.MimeContent = new MimeContentType(sampleMime, 'utf8');
            message.IsReadReceiptRequested = true;
            message.IsDeliveryReceiptRequested = true;

            const context = getContext(testUser, 'password');

            const newMessages = await manager.createItem(message, userInfo, context.request, MessageDispositionType.SAVE_ONLY, toLabel, mailboxId);
            expect(newMessages).to.not.be.undefined();
            expect(newMessages.length).to.be.equal(1);
            expect(newMessages[0]?.ItemId?.Id).to.be.equal(itemEWSId);
        });

        it('create mime message with label with no unid', async () => {

            const messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            messageManagerStub.createMimeMessage.resolves('UID');
            messageManagerStub.moveMessages.resolves({});
            messageManagerStub.deleteMessage.resolves({});
            messageManagerStub.moveMessages.resolves();

            stubMessageManager(messageManagerStub);

            const manager = new MockEWSManager();
            const userInfo = new UserContext();
            const toLabel = PimItemFactory.newPimLabel({}, PimItemFormat.DOCUMENT);
            toLabel.view = KeepPimConstants.INBOX;

            const message = new MessageType();
            const itemEWSId = getEWSId('UID');
            message.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            message.MimeContent = new MimeContentType(sampleMime, 'utf8');

            const context = getContext(testUser, 'password');

            let success = false;
            try {
                await manager.createItem(message, userInfo, context.request, MessageDispositionType.SAVE_ONLY, toLabel);
            } catch (err) {
                success = true;
            }
            expect(success).to.be.true();
        });

        it('create mime message to SENT folder', async () => {
            const messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            messageManagerStub.createMimeMessage.resolves('UID');
            messageManagerStub.moveMessages.resolves({});
            messageManagerStub.deleteMessage.resolves({});
            messageManagerStub.moveMessages.resolves();

            stubMessageManager(messageManagerStub);

            const manager = new MockEWSManager();
            const userInfo = new UserContext();
            const toLabel = PimItemFactory.newPimLabel({}, PimItemFormat.DOCUMENT);
            toLabel.view = KeepPimConstants.SENT;
            toLabel.unid = 'TO-LABEL';

            const message = new MessageType();
            const itemEWSId = getEWSId('UID');
            message.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            message.MimeContent = new MimeContentType(sampleMime, 'utf8');

            const context = getContext(testUser, 'password');

            const newMessage = await manager.createItem(message, userInfo, context.request, MessageDispositionType.SAVE_ONLY, toLabel);
            expect(newMessage).to.not.be.undefined();

        });

        it('create mime message to DRAFTS folder', async () => {
            const messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            messageManagerStub.createMimeMessage.resolves('UID');
            messageManagerStub.moveMessages.resolves({});
            messageManagerStub.deleteMessage.resolves({});
            messageManagerStub.moveMessages.resolves();

            stubMessageManager(messageManagerStub);

            const manager = new MockEWSManager();
            const userInfo = new UserContext();
            const toLabel = PimItemFactory.newPimLabel({}, PimItemFormat.DOCUMENT);
            toLabel.view = KeepPimConstants.DRAFTS;
            toLabel.unid = 'TO-LABEL';

            const message = new MessageType();
            const itemEWSId = getEWSId('UID');
            message.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            message.MimeContent = new MimeContentType(sampleMime, 'utf8');

            const context = getContext(testUser, 'password');

            const newMessage = await manager.createItem(message, userInfo, context.request, MessageDispositionType.SAVE_ONLY, toLabel);
            expect(newMessage).to.not.be.undefined();

        });

        it('test create message', async () => {

            const messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            messageManagerStub.createMessage.resolves('UID');
            messageManagerStub.moveMessages.resolves({});
            messageManagerStub.deleteMessage.resolves({});

            stubMessageManager(messageManagerStub);

            const manager = new MockEWSManager();
            const userInfo = new UserContext();

            const message = new MessageType();
            const itemEWSId = getEWSId('UID');
            message.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            const context = getContext(testUser, 'password');

            const newMessages = await manager.createItem(message, userInfo, context.request);
            expect(newMessages).to.not.be.undefined();
            expect(newMessages.length).to.be.equal(1);
            expect(newMessages[0]?.ItemId?.Id).to.be.equal(itemEWSId);
        });

        it('create mime message createMimeMessage throws error', async () => {
            const messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            messageManagerStub.createMimeMessage.throws(new Error('error'));
            messageManagerStub.moveMessages.resolves({});
            messageManagerStub.deleteMessage.resolves({});
            messageManagerStub.moveMessages.resolves();

            stubMessageManager(messageManagerStub);

            const manager = new MockEWSManager();
            const userInfo = new UserContext();
            const toLabel = PimItemFactory.newPimLabel({}, PimItemFormat.DOCUMENT);
            toLabel.view = KeepPimConstants.DRAFTS;
            toLabel.unid = 'TO-LABEL';

            const message = new MessageType();
            const mailboxId = 'test@test.com';
            const itemEWSId = getEWSId('UID', mailboxId);
            message.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);
            message.MimeContent = new MimeContentType(sampleMime, 'utf8');

            const context = getContext(testUser, 'password');

            let success = false;
            try {
                await manager.createItem(message, userInfo, context.request, MessageDispositionType.SAVE_ONLY, toLabel, mailboxId);
            } catch (err) {
                success = true
            }
            expect(success).to.be.true();

        });

        describe('test creating the calendar response', () => {
            const manager = new MockEWSManager();
            const userInfo = new UserContext();
            const context = getContext(testUser, 'password');
            const messageEWSId = getEWSId('message-id');
            const calendarEWSId = getEWSId('cal-item-id');

            const testData = [
                { itemType: 'accept', item: new AcceptItemType() },
                { itemType: 'decline', item: new DeclineItemType() },
                { itemType: 'tentatively accept', item: new TentativelyAcceptItemType() },
                { itemType: 'cancel calendar', item: new CancelCalendarItemType() }
            ];

            afterEach(() => {
                sinon.restore(); 
            });

            testData.forEach(({ itemType, item }) => {
                it(`test ${itemType} item type of message`, async () => {
                    const messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
                    messageManagerStub.createMessage.resolves('message-id');
                    messageManagerStub.moveMessages.resolves({});
                    messageManagerStub.deleteMessage.resolves({});
                    stubMessageManager(messageManagerStub);

                    const calendarManagerStub = sinon.createStubInstance(KeepPimCalendarManager);
                    calendarManagerStub.createCalendarResponse.resolves({});
                    stubCalendarManager(calendarManagerStub);

                    const pimCalendarItem = PimItemFactory.newPimCalendarItem({},  "default", PimItemFormat.DOCUMENT);
                    pimCalendarItem.unid = 'cal-item-id';
                    const pimManagerStub = sinon.createStubInstance(KeepPimManager);
                    pimManagerStub.getPimItem.resolves(pimCalendarItem);
                    stubPimManager(pimManagerStub);

                    const message = item;

                    message.ItemId = new ItemIdType(messageEWSId);
                    message.ReferenceItemId = new ItemIdType(calendarEWSId);
        
                    const newItems = await manager.createItem(message, userInfo, context.request);
                    expect(newItems).to.not.be.undefined();
                    expect(newItems.length).to.be.equal(1);
                    expect(newItems[0]?.ItemId?.Id).to.be.equal(calendarEWSId);
                })

                it(`test ${itemType} item type of message but no reference id provided`, async () => {
                    const message = item;

                    message.ItemId = new ItemIdType(messageEWSId);
                    message.ReferenceItemId = undefined;
        
                    const newItems = await manager.createItem(message, userInfo, context.request);
                    expect(newItems).to.not.be.undefined();
                    expect(newItems.length).to.be.equal(0);
                })
            })
        });
    });
    /*   
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
                   change.ItemId = new ItemIdType('UNID', 'ck-UNID');
                   change.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();
                   const subjectChange = new SetItemFieldType();
                   subjectChange.FieldURI = new PathToUnindexedFieldType();
                   subjectChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
                   subjectChange.Item = new ItemType();
                   subjectChange.Item.Subject = 'UPDATED SUBJECT';
                   change.Updates.push(subjectChange);
                   const userInfo = new UserContext();
                   try {
                       const updatedNote = await EWSServiceManager.updateItem(change, userInfo)
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
                       const updatedNote = await EWSServiceManager.updateItem(change, userInfo)
                       success = updatedNote === undefined;
                   } catch (err) {
                       success = true;
                   }
                   expect(success).to.be.true();
               });
       
               it('Unsupported item type', async () => {
                   // Can't have an await in a test, so do all te work in the before and just check the results in the it
                   const pimThread = PimItemFactory.newPimThread({});
                   const pimManagerStub = sinon.createStubInstance(KeepPimManager);
                   pimManagerStub.getPimItem.resolves(pimThread);
                   stubPimManager(pimManagerStub);
       
                   const change = new ItemChangeType();
                   change.ItemId = new ItemIdType('UNID', 'ck-UNID');
                   change.Updates = new NonEmptyArrayOfItemChangeDescriptionsType();
                   const subjectChange = new SetItemFieldType();
                   subjectChange.FieldURI = new PathToUnindexedFieldType();
                   subjectChange.FieldURI.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
                   subjectChange.Item = new ItemType();
                   subjectChange.Item.Subject = 'UPDATED SUBJECT';
                   change.Updates.push(subjectChange);
                   const userInfo = new UserContext();
                   const updatedNote = await EWSServiceManager.updateItem(change, userInfo)
       
                   expect(updatedNote).to.be.undefined();
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
       */

    describe('Test addRequestedPropertiesToEWSMessage', () => {

        let pimMessage: PimMessage;
        let message: MessageType;

        // Reset pimItem before each test
        beforeEach(function () {

            const extendedProp = {
                DistinguishedPropertySetId: 'Address',
                PropertyId: 32899,
                PropertyType: 'String',
                Value: 'kermit.frog@fakemail.com'
            }

            pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';
            pimMessage.parentFolderIds = ['PARENT-FOLDER'];
            pimMessage.subject = 'SUBJECT';
            pimMessage.body = 'BODY';
            pimMessage.addExtendedProperty(extendedProp);

            message = new MessageType();

        });

        it("base fields for message", async function () {
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;

            const mockManager = new MockEWSManager();
            const mimeString = base64Decode(sampleMime);
            const mailboxId = 'test@test.com';

            await mockManager.addRequestedPropertiesToEWSMessage(pimMessage, message, mimeString, 'utf8', shape, mailboxId);
            
            const parentEWSId = getEWSId('PARENT-FOLDER', mailboxId);
            const itemEWSId = getEWSId(pimMessage.unid, mailboxId);

            expect(message.ItemId?.Id).to.be.equal(itemEWSId);
            expect(message.Subject).to.be.equal('Hi');
            expect(message.ParentFolderId?.Id).to.be.equal(parentEWSId);

            // Try with a size set on the pimMessage
            pimMessage.size = 1024;
            await mockManager.addRequestedPropertiesToEWSMessage(pimMessage, message, mimeString, 'utf8', shape, mailboxId);
            expect(message.ItemId?.Id).to.be.equal(itemEWSId);
            expect(message.Subject).to.be.equal('Hi');
            expect(message.Size).to.be.equal(1024);

        });

        it("base fields for message, no shape", async function () {
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;

            const mockManager = new MockEWSManager();
            const mimeString = base64Decode(sampleMime);
            await mockManager.addRequestedPropertiesToEWSMessage(pimMessage, message, mimeString, 'utf8');
            expect(message.ItemId).to.be.undefined();
            expect(message.Subject).to.be.undefined();
        });

        it("base shape with additional properties", async function () {
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            const path1 = new PathToUnindexedFieldType();
            path1.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            const path2 = new PathToUnindexedFieldType();
            path2.FieldURI = UnindexedFieldURIType.ITEM_SIZE;
            const path3 = new PathToUnindexedFieldType();
            path3.FieldURI = UnindexedFieldURIType.ITEM_EFFECTIVE_RIGHTS;
            const path4 = new PathToIndexedFieldType();
            path4.FieldURI = DictionaryURIType.ITEM_INTERNET_MESSAGE_HEADER;

            shape.AdditionalProperties = new NonEmptyArrayOfPathsToElementType([path1, path2, path3, path4]);

            const mockManager = new MockEWSManager();
            const mimeString = base64Decode(sampleMime);
            const mailboxId = 'test@test.com';

            await mockManager.addRequestedPropertiesToEWSMessage(pimMessage, message, mimeString, 'utf8', shape, mailboxId);

            const parentEWSId = getEWSId('PARENT-FOLDER', mailboxId);
            const itemEWSId = getEWSId(pimMessage.unid, mailboxId);

            expect(message.ItemId?.Id).to.be.equal(itemEWSId);
            expect(message.Subject).to.be.equal('Hi');
            expect(message.ParentFolderId?.Id).to.be.equal(parentEWSId);

            // Try with a size set on the pimMessage
            pimMessage.size = 1024;
            await mockManager.addRequestedPropertiesToEWSMessage(pimMessage, message, mimeString, 'utf8', shape, mailboxId);
            expect(message.ItemId?.Id).to.be.equal(itemEWSId);
            expect(message.Subject).to.be.equal('Hi');
            expect(message.Size).to.be.equal(1024);
        });

        it("extended fields", async function () {
            // Try an extended properties
            const extPath = new PathToExtendedFieldType();
            extPath.PropertyId = 32899;
            extPath.PropertyType = MapiPropertyTypeType.STRING;
            // Path that doesn't exist
            const extPath2 = new PathToExtendedFieldType();
            extPath2.PropertyId = 44333;
            extPath2.PropertyType = MapiPropertyTypeType.STRING;


            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            shape.AdditionalProperties = new NonEmptyArrayOfPathsToElementType([extPath, extPath2]);

            const mockManager = new MockEWSManager();
            const mimeString = base64Decode(sampleMime);
            await mockManager.addRequestedPropertiesToEWSMessage(pimMessage, message, mimeString, 'utf8', shape);
            expect(message.ItemId?.Id).to.be.equal(getEWSId(pimMessage.unid));
            expect(message.Subject).to.be.undefined();
            expect(message.ParentFolderId?.Id).to.be.equal(getEWSId('PARENT-FOLDER'));

            if (message.ExtendedProperty) {
                expect(message.ExtendedProperty.length).to.be.equal(1);
                const ext: ExtendedPropertyType = message.ExtendedProperty[0];
                expect(ext).is.not.undefined();
                expect(ext.Value).to.be.equal('kermit.frog@fakemail.com');
                expect(ext.ExtendedFieldURI.PropertyId).to.be.equal(32899);
                expect(ext.ExtendedFieldURI.DistinguishedPropertySetId).to.be.equal('Address');
                expect(ext.ExtendedFieldURI.PropertyType).to.be.equal('String');
            }
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
            mockManager.addRequestedPropertiesToEWSItem(pimMessage, message, shape);

            expect(message.ItemId?.Id).to.be.equal(getEWSId(pimMessage.unid));
            expect(message.Body?.Value).to.be.equal(pimMessage.body);
            expect(message.Body?.BodyType).to.be.equal('Text');
            expect(message.Subject).to.be.equal(pimMessage.subject);
            expect(message.ParentFolderId?.Id).to.be.equal(getEWSId('PARENT-FOLDER'));

            if (message.ExtendedProperty) {
                const ext: ExtendedPropertyType = message.ExtendedProperty[0];
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
            mockManager.addRequestedPropertiesToEWSItem(pimMessage, message);
            expect(message.ItemId).to.be.undefined();

            if (message.ExtendedProperty) {
                const ext: ExtendedPropertyType = message.ExtendedProperty[0];
                expect(ext).is.not.undefined();
                expect(ext.Value).to.be.equal('kermit.frog@fakemail.com');
                expect(ext.ExtendedFieldURI.PropertyId).to.be.equal(32899);
                expect(ext.ExtendedFieldURI.DistinguishedPropertySetId).to.be.equal('Address');
                expect(ext.ExtendedFieldURI.PropertyType).to.be.equal('String');
            }
        });
    });

    describe('Test updateEWSItemFieldValue', () => {

        it("Test unhandled items", function () {
            const pimMessage = PimItemFactory.newPimMessage({});
            const message = new MessageType();
            const manager = new MockEWSManager();

            const unhandledFields = [
                UnindexedFieldURIType.ITEM_IS_ASSOCIATED,
                UnindexedFieldURIType.ITEM_RESPONSE_OBJECTS,
                UnindexedFieldURIType.CONTACTS_BUSINESS_HOME_PAGE,
                UnindexedFieldURIType.CONVERSATION_CONVERSATION_ID,
                UnindexedFieldURIType.MESSAGE_CONVERSATION_TOPIC,
                UnindexedFieldURIType.ITEM_CULTURE,
                UnindexedFieldURIType.ITEM_DISPLAY_CC,
                UnindexedFieldURIType.ITEM_DISPLAY_TO,
                UnindexedFieldURIType.MESSAGE_INTERNET_MESSAGE_ID,
                UnindexedFieldURIType.ITEM_INTERNET_MESSAGE_HEADERS,
                UnindexedFieldURIType.ITEM_IS_FROM_ME,
                UnindexedFieldURIType.ITEM_IS_RESEND,
                UnindexedFieldURIType.MESSAGE_IS_RESPONSE_REQUESTED,
                UnindexedFieldURIType.ITEM_IS_SUBMITTED,
                UnindexedFieldURIType.ITEM_IS_UNMODIFIED,
                UnindexedFieldURIType.MESSAGE_REFERENCES,
                UnindexedFieldURIType.MESSAGE_RECEIVED_BY,
                UnindexedFieldURIType.MESSAGE_RECEIVED_REPRESENTING,
                UnindexedFieldURIType.ITEM_REMINDER_DUE_BY,
                UnindexedFieldURIType.ITEM_REMINDER_IS_SET,
                UnindexedFieldURIType.ITEM_REMINDER_MINUTES_BEFORE_START,
                UnindexedFieldURIType.ITEM_WEB_CLIENT_EDIT_FORM_QUERY_STRING,
                UnindexedFieldURIType.ITEM_WEB_CLIENT_READ_FORM_QUERY_STRING
            ];

            unhandledFields.forEach(field => {
                const handled = manager.updateEWSItemFieldValue(message, pimMessage, field);
                expect(handled).to.be.false();
            });
        });

        it("UnindexedFieldURIType.ITEM_DATE_TIME_SENT", function () {
            const pimMessage = PimItemFactory.newPimMessage({});
            const message = new MessageType();
            const manager = new MockEWSManager();

            // No sent date
            let handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.ITEM_DATE_TIME_SENT);
            expect(handled).to.be.true();
            expect(message.DateTimeSent).to.be.undefined();

            pimMessage.sentDate = new Date();
            handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.ITEM_DATE_TIME_SENT);
            expect(handled).to.be.true();
            expect(message.DateTimeSent).to.be.eql(pimMessage.sentDate);
        });

        it("UnindexedFieldURIType.MESSAGE_FROM", function () {
            const pimMessage = PimItemFactory.newPimMessage({});
            const message = new MessageType();
            const manager = new MockEWSManager();

            // No from
            let handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.MESSAGE_FROM);
            expect(handled).to.be.true();
            expect(message.From).to.be.undefined();

            pimMessage.from = testUser;
            handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.MESSAGE_FROM);
            expect(handled).to.be.true();
            expect(message.From?.Mailbox.EmailAddress).to.be.equal(testUser);
        });

        it('UnindexedFieldURIType.MESSAGE_BCC_RECIPIENTS', function () {
            const pimMessage = PimItemFactory.newPimMessage({});
            const message = new MessageType();
            const manager = new MockEWSManager();

            // No from
            let handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.MESSAGE_BCC_RECIPIENTS);
            expect(handled).to.be.true();
            expect(message.BccRecipients).to.be.undefined();

            pimMessage.bcc = [testUser];
            handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.MESSAGE_BCC_RECIPIENTS);
            expect(handled).to.be.true();
            expect(message.BccRecipients?.Mailbox[0].EmailAddress).to.be.equal(testUser);
        });

        it('UnindexedFieldURIType.MESSAGE_CC_RECIPIENTS', function () {
            const pimMessage = PimItemFactory.newPimMessage({});
            const message = new MessageType();
            const manager = new MockEWSManager();

            // No from
            let handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.MESSAGE_CC_RECIPIENTS);
            expect(handled).to.be.true();
            expect(message.CcRecipients).to.be.undefined();

            pimMessage.cc = [testUser];
            handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.MESSAGE_CC_RECIPIENTS);
            expect(handled).to.be.true();
            expect(message.CcRecipients?.Mailbox[0].EmailAddress).to.be.equal(testUser);
        });

        it('UnindexedFieldURIType.MESSAGE_TO_RECIPIENTS', function () {
            const pimMessage = PimItemFactory.newPimMessage({});
            const message = new MessageType();
            const manager = new MockEWSManager();

            // No from
            let handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.MESSAGE_TO_RECIPIENTS);
            expect(handled).to.be.true();
            expect(message.ToRecipients).to.be.undefined();

            pimMessage.to = [testUser];
            handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.MESSAGE_TO_RECIPIENTS);
            expect(handled).to.be.true();
            expect(message.ToRecipients?.Mailbox[0].EmailAddress).to.be.equal(testUser);
        });

        it('UnindexedFieldURIType.ITEM_IMMPORTANCE', function () {
            const pimMessage = PimItemFactory.newPimMessage({});
            let message = new MessageType();
            const manager = new MockEWSManager();

            // No from
            let handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.ITEM_IMPORTANCE);
            expect(handled).to.be.true();
            expect(message.Importance).to.be.undefined();

            pimMessage.importance = PimImportance.HIGH;
            handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.ITEM_IMPORTANCE);
            expect(handled).to.be.true();
            expect(message.Importance).to.be.equal(ImportanceChoicesType.HIGH);

            pimMessage.importance = PimImportance.LOW;
            handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.ITEM_IMPORTANCE);
            expect(handled).to.be.true();
            expect(message.Importance).to.be.equal(ImportanceChoicesType.LOW);

            message = new MessageType();
            pimMessage.importance = PimImportance.MEDIUM;
            handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.ITEM_IMPORTANCE);
            expect(handled).to.be.true();
            expect(message.Importance).to.be.undefined();
        });

        it('UnindexedFieldURIType.ITEM_SIZE', function () {
            const pimMessage = PimItemFactory.newPimMessage({});
            const message = new MessageType();
            const manager = new MockEWSManager();
            pimMessage.size = 1024;

            // No from
            const handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.ITEM_SIZE);
            expect(handled).to.be.true();
            expect(message.Size).to.be.equal(1024);
        });

        it('UnindexedFieldURIType.MESSAGE_IS_READ', function () {
            const pimMessage = PimItemFactory.newPimMessage({});
            const message = new MessageType();
            const manager = new MockEWSManager();
            pimMessage.isRead = true;

            // No from
            const handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.MESSAGE_IS_READ);
            expect(handled).to.be.true();
            expect(message.IsRead).to.be.true();
        });

        it("UnindexedFieldURIType.ITEM_DATE_TIME_RECEIVED", function () {
            const pimMessage = PimItemFactory.newPimMessage({});
            const message = new MessageType();
            const manager = new MockEWSManager();

            // No sent date
            let handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.ITEM_DATE_TIME_RECEIVED);
            expect(handled).to.be.true();
            expect(message.DateTimeReceived).to.be.undefined();

            pimMessage.receivedDate = new Date();
            handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.ITEM_DATE_TIME_RECEIVED);
            expect(handled).to.be.true();
            expect(message.DateTimeReceived).to.be.eql(pimMessage.receivedDate);
        });


        it('UnindexedFieldURIType.MESSAGE_IS_READ_RECEIPT_REQUESTED', function () {
            const pimMessage = PimItemFactory.newPimMessage({});
            const message = new MessageType();
            const manager = new MockEWSManager();

            let handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.MESSAGE_IS_READ_RECEIPT_REQUESTED);
            expect(handled).to.be.true();
            expect(message.IsReadReceiptRequested).to.be.undefined();

            pimMessage.returnReceipt = true;
            handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.MESSAGE_IS_READ_RECEIPT_REQUESTED);
            expect(handled).to.be.true();
            expect(message.IsReadReceiptRequested).to.be.true();
        });


        it('UnindexedFieldURIType.MESSAGE_IS_DELIVERY_RECEIPT_REQUESTED', function () {
            const pimMessage = PimItemFactory.newPimMessage({});
            const message = new MessageType();
            const manager = new MockEWSManager();

            let handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.MESSAGE_IS_DELIVERY_RECEIPT_REQUESTED);
            expect(handled).to.be.true();
            expect(message.IsDeliveryReceiptRequested).to.be.undefined();

            pimMessage.deliveryReceipt = true;
            handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.MESSAGE_IS_DELIVERY_RECEIPT_REQUESTED);
            expect(handled).to.be.true();
            expect(message.IsDeliveryReceiptRequested).to.be.true();
        });

        it('UnindexedFieldURIType.MESSAGE_REPLY_TO', function () {
            const pimMessage = PimItemFactory.newPimMessage({});
            const message = new MessageType();
            const manager = new MockEWSManager();

            // No from
            let handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.MESSAGE_REPLY_TO);
            expect(handled).to.be.true();
            expect(message.ReplyTo).to.be.undefined();

            pimMessage.replyTo = [testUser];
            handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.MESSAGE_REPLY_TO);
            expect(handled).to.be.true();
            expect(message.ReplyTo?.Mailbox[0].EmailAddress).to.be.equal(testUser);
        });

        it('UnindexedFieldURIType.MESSAGE_SENDER', function () {
            const pimMessage = PimItemFactory.newPimMessage({});
            const message = new MessageType();
            const manager = new MockEWSManager();

            // No from
            let handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.MESSAGE_SENDER);
            expect(handled).to.be.true();
            expect(message.Sender?.Mailbox).to.be.undefined();

            pimMessage.from = testUser;
            handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.MESSAGE_SENDER);
            expect(handled).to.be.true();
            expect(message.Sender?.Mailbox.EmailAddress).to.be.equal(testUser);
        });

        it('Unhandled fields', () => {
            const pimMessage = PimItemFactory.newPimMessage({});
            const message = new MessageType();
            const manager = new MockEWSManager();

            pimMessage.unid = 'uid';
            const mailboxId = 'test@test.com';
            const itemEWSId = getEWSId(pimMessage.unid, mailboxId);
            let handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.ITEM_ITEM_ID, mailboxId);
            expect(handled).to.be.true();
            expect(message.ItemId?.Id).to.be.equal(itemEWSId);

            handled = manager.updateEWSItemFieldValue(message, pimMessage, UnindexedFieldURIType.TASK_BILLING_INFORMATION, mailboxId);
            expect(handled).to.be.false();
        });

    });

    describe('Test buildMimeContentFromMeetingMessage', () => {
        it('meeting request, no iCal', async () => {
            const pimItem = PimItemFactory.newPimMessage({ 'NoticeType': 'I' }, PimItemFormat.DOCUMENT);
            pimItem.unid = 'UNID';
            pimItem.from = testUser;
            pimItem.to = ['user@hcl.com'];
            pimItem.cc = ['ccuser@hcl.com'];
            pimItem.bcc = ['bccuser@hcl.com'];
            pimItem.replyTo = [testUser];
            pimItem.createdDate = new Date();
            pimItem.subject = 'SUBJECT';
            pimItem.icalMessageId = 'ICAL-ID';
            pimItem.body = 'BODY';

            const manager = new MockEWSManager();

            const mime = manager.buildMimeContentForMeetingMessage(pimItem);
            expect(mime).to.not.be.undefined();

            const mimeValue = mime.Value;
            const messageBuffer = Buffer.from(mimeValue, 'base64');
            const parsed = await simpleParser(messageBuffer);
            expect(parsed.to?.text).to.be.equal('user@hcl.com');
            expect(parsed.bcc?.text).to.be.equal('bccuser@hcl.com');
            expect(parsed.cc?.text).to.be.equal('ccuser@hcl.com');
            expect(parsed.subject).to.be.equal(pimItem.subject);
        });

        it('meeting request, no info', () => {
            const pimItem = PimItemFactory.newPimMessage({ 'NoticeType': 'I' }, PimItemFormat.DOCUMENT);
            pimItem.unid = 'UNID';

            const manager = new MockEWSManager();

            const mime = manager.buildMimeContentForMeetingMessage(pimItem);
            expect(mime).to.not.be.undefined();
        });

        it('meeting request with iCal', async () => {
            const pimItem = PimItemFactory.newPimMessage({ 'NoticeType': 'I' }, PimItemFormat.DOCUMENT);
            pimItem.unid = 'UNID';
            pimItem.from = testUser;
            pimItem.to = ['user@hcl.com'];
            pimItem.cc = ['ccuser@hcl.com'];
            pimItem.bcc = ['bccuser@hcl.com'];
            pimItem.replyTo = [testUser];
            pimItem.createdDate = new Date();
            pimItem.subject = 'SUBJECT';
            pimItem.icalMessageId = 'ICAL-ID';
            pimItem.icalStream = iCalInvitation;

            const manager = new MockEWSManager();

            const mime = manager.buildMimeContentForMeetingMessage(pimItem);
            expect(mime).to.not.be.undefined();

            const mimeValue = mime.Value;
            const messageBuffer = Buffer.from(mimeValue, 'base64');
            const parsed = await simpleParser(messageBuffer);
            expect(parsed.to?.text).to.be.equal('user@hcl.com');
            expect(parsed.bcc?.text).to.be.equal('bccuser@hcl.com');
            expect(parsed.cc?.text).to.be.equal('ccuser@hcl.com');
            expect(parsed.subject).to.be.equal(pimItem.subject);
            const body = parsed.html.toString();
            expect(body.includes('Buckeyes')).to.be.true();
        });

        it('meeting response, no iCal', async () => {
            // User declined
            const pimItem = PimItemFactory.newPimMessage({ 'NoticeType': 'R' }, PimItemFormat.DOCUMENT);
            pimItem.unid = 'UNID';
            pimItem.from = testUser;
            pimItem.to = ['user@hcl.com'];
            pimItem.cc = ['ccuser@hcl.com'];
            pimItem.bcc = ['bccuser@hcl.com'];
            pimItem.replyTo = [testUser];
            pimItem.createdDate = new Date();
            pimItem.subject = 'SUBJECT';
            pimItem.icalMessageId = 'ICAL-ID';
            pimItem.body = 'BODY';

            const manager = new MockEWSManager();

            const mime = manager.buildMimeContentForMeetingMessage(pimItem);
            expect(mime).to.not.be.undefined();

            const mimeValue = mime.Value;
            const messageBuffer = Buffer.from(mimeValue, 'base64');
            const parsed = await simpleParser(messageBuffer);
            expect(parsed.to?.text).to.be.equal('user@hcl.com');
            expect(parsed.bcc?.text).to.be.equal('bccuser@hcl.com');
            expect(parsed.cc?.text).to.be.equal('ccuser@hcl.com');
            expect(parsed.subject).to.be.equal(pimItem.subject);
        });

        it('meeting response, no info', () => {
            // Decline response
            const pimItem = PimItemFactory.newPimMessage({ 'NoticeType': 'R' }, PimItemFormat.DOCUMENT);
            pimItem.unid = 'UNID';

            const manager = new MockEWSManager();

            const mime = manager.buildMimeContentForMeetingMessage(pimItem);
            expect(mime).to.not.be.undefined();
        });

        it('meeting response with iCal', async () => {
            const pimItem = PimItemFactory.newPimMessage({ 'NoticeType': 'R' }, PimItemFormat.DOCUMENT);
            pimItem.unid = 'UNID';
            pimItem.from = testUser;
            pimItem.to = ['user@hcl.com'];
            pimItem.cc = ['ccuser@hcl.com'];
            pimItem.bcc = ['bccuser@hcl.com'];
            pimItem.replyTo = [testUser];
            pimItem.createdDate = new Date();
            pimItem.subject = 'SUBJECT';
            pimItem.icalMessageId = 'ICAL-ID';
            pimItem.icalStream = iCalDeclineResponse;

            const manager = new MockEWSManager();

            const mime = manager.buildMimeContentForMeetingMessage(pimItem);
            expect(mime).to.not.be.undefined();

            const mimeValue = mime.Value;
            const messageBuffer = Buffer.from(mimeValue, 'base64');
            const parsed = await simpleParser(messageBuffer);
            expect(parsed.to?.text).to.be.equal('user@hcl.com');
            expect(parsed.bcc?.text).to.be.equal('bccuser@hcl.com');
            expect(parsed.cc?.text).to.be.equal('ccuser@hcl.com');
            expect(parsed.subject).to.be.equal(pimItem.subject);
            const body = parsed.html.toString();
            expect(body.includes('DISCLAIMER')).to.be.true();
        });


    });
    describe('Test pimItemToEWSMessage', () => {

        // Clean up the stubs
        afterEach(() => {
            sinon.restore(); // Reset stubs
        });

        it('meeting request', async () => {
            const pimItem = PimItemFactory.newPimMessage({ 'NoticeType': 'I' }, PimItemFormat.DOCUMENT);
            pimItem.unid = 'UNID';
            pimItem.from = testUser;
            pimItem.to = ['user@hcl.com'];
            pimItem.cc = ['ccuser@hcl.com'];
            pimItem.bcc = ['bccuser@hcl.com'];
            pimItem.replyTo = [testUser];
            pimItem.createdDate = new Date();
            pimItem.subject = 'SUBJECT';
            pimItem.icalMessageId = 'ICAL-ID';
            pimItem.body = 'BODY';


            const manager = new MockEWSManager();
            const userInfo = new UserContext();
            const context = getContext(testUser, 'password');
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            shape.IncludeMimeContent = true;

            const message = await manager.pimItemToEWSItem(pimItem, userInfo, context.request, shape);
            expect(message.ItemId?.Id).to.be.equal(getEWSId('UNID'));

        });

        it('meeting response', async () => {
            const pimItem = PimItemFactory.newPimMessage({ 'NoticeType': 'A' }, PimItemFormat.DOCUMENT);
            pimItem.unid = 'UNID';

            const manager = new MockEWSManager();
            const userInfo = new UserContext();
            const context = getContext(testUser, 'password');
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;

            const message = await manager.pimItemToEWSItem(pimItem, userInfo, context.request, shape);
            expect(message.ItemId?.Id).to.be.equal(getEWSId('UNID'));

        });

        it('mime message', async () => {

            const messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            messageManagerStub.getMimeMessage.resolves(base64Decode(sampleMime));
            stubMessageManager(messageManagerStub);

            const pimItem = PimItemFactory.newPimMessage({}, PimItemFormat.DOCUMENT);
            pimItem.unid = 'UNID';
            pimItem.from = testUser;

            const manager = new MockEWSManager();
            const userInfo = new UserContext();
            const context = getContext(testUser, 'password');
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;

            const mailboxId = 'test@test.com';

            let message = await manager.pimItemToEWSItem(pimItem, userInfo, context.request, shape, mailboxId, 'PARENT-FOLDER');

            const itemEWSId = getEWSId('UNID', mailboxId);
            const parentEWSId = getEWSId('PARENT-FOLDER', mailboxId);

            expect(message.ItemId?.Id).to.be.equal(itemEWSId);
            expect(message.ParentFolderId?.Id).to.be.equal(parentEWSId);

            shape.IncludeMimeContent = true;
            message = await manager.pimItemToEWSItem(pimItem, userInfo, context.request, shape, mailboxId, 'PARENT-FOLDER');
            expect(message.ItemId?.Id).to.be.equal(itemEWSId);
            expect(message.ParentFolderId?.Id).to.be.equal(parentEWSId);
            expect(message.MimeContent?.Value).to.be.equal(sampleMime);
        });

        it('mime message error path', async () => {

            const messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            messageManagerStub.getMimeMessage.resolves(undefined);
            stubMessageManager(messageManagerStub);

            const pimItem = PimItemFactory.newPimMessage({}, PimItemFormat.DOCUMENT);
            pimItem.unid = 'UNID';
            pimItem.from = testUser;

            const manager = new MockEWSManager();
            const userInfo = new UserContext();
            const context = getContext(testUser, 'password');
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;

            let success = false;
            try {
                await manager.pimItemToEWSItem(pimItem, userInfo, context.request, shape, 'test@test.com', 'PARENT-FOLDER');
            } catch (err) {
                success = true;
            }
            expect(success).to.be.true();
        });

    });

    describe('test updateEWSItemFieldValueFromMime', () => {

        let parsed: ParsedMail;
        const messageBuffer = Buffer.from(sampleMime, 'base64');

        class MyAttachmentClass implements Attachment {
            type: 'attachment';
            /**
            * MIME type of the message.
            */
            contentType: string;
            /**
             * Content disposition type for the attachment,
             * most probably `'attachment'`.
             */
            contentDisposition: string;
            /**
             * File name of the attachment.
             */
            filename?: string;
            /**
             * A Map value that holds MIME headers for the attachment node.
             */
            headers: Headers;
            /**
             * An array of raw header lines for the attachment node.
             */
            headerLines: HeaderLines;
            /**
             * A MD5 hash of the message content.
             */
            checksum: string;
            /**
             * Message size in bytes.
             */
            size: number;
            /**
             * The header value from `Content-ID`.
             */
            contentId?: string;
            /**
             * `contentId` without `<` and `>`.
             */
            cid?: string;   // e.g. '5.1321281380971@localhost'

            /**
            * A Buffer that contains the attachment contents.
            */
            content: Buffer;
            /**
             * If true then this attachment should not be offered for download
             * (at least not in the main attachments list).
             */
            related: boolean;
        }

        beforeEach(async function () {
            parsed = await simpleParser(messageBuffer);
        });

        it('test ITEM_HAS_ATTACHMENTS', async () => {
            const manager = new MockEWSManager();
            const message = new MessageType();
            const pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';

            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS);
            expect(message.HasAttachments).to.be.false();

            // Mime overrides pimItem setting
            pimMessage.attachments = ['abc'];
            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS);
            expect(message.HasAttachments).to.be.false();

            parsed.attachments = [];
            const attachment = new MyAttachmentClass();
            attachment.contentType = 'message/rfc822';
            attachment.content = Buffer.from('abc');
            parsed.attachments.push(attachment);

            pimMessage.attachments = [];
            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS);
            expect(message.HasAttachments).to.be.true();
        });

        it('test ITEM_ATTACHMENTS', async () => {
            const manager = new MockEWSManager();
            let message = new MessageType();
            const pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';

            parsed.attachments = [];
            const attachment = new MyAttachmentClass();
            attachment.contentType = 'message/rfc822';
            attachment.content = Buffer.from('abc');
            parsed.attachments.push(attachment);

            // Attachment with disposition of inline will be skipped
            const attachment2 = new MyAttachmentClass();
            attachment2.contentDisposition = 'inline';
            attachment2.contentType = 'message/rfc822';
            attachment2.content = Buffer.from('abc');
            parsed.attachments.push(attachment2);

            // File attachment
            const attachment3 = new MyAttachmentClass();
            attachment3.contentDisposition = 'inLine';
            attachment3.contentType = 'text/plain';
            attachment3.content = Buffer.from('abc');
            attachment3.filename = 'file.txt'
            parsed.attachments.push(attachment3);

            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.ITEM_ATTACHMENTS);
            expect(message.Attachments?.items).to.be.an.Array();
            expect(message.Attachments?.items.length).to.be.equal(2);

            // No attachments
            parsed.attachments = [];
            message = new MessageType();
            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.ITEM_ATTACHMENTS);
            expect(message.Attachments).to.be.undefined();
        });

        it('test ITEM_SUBJECT', async () => {
            const manager = new MockEWSManager();
            const message = new MessageType();
            const pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';

            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.ITEM_SUBJECT);
            expect(message.Subject).to.be.equal('Hi');
        });

        it('test ITEM_IN_REPLY_TO', async () => {
            const manager = new MockEWSManager();
            const message = new MessageType();
            const pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';
            parsed.inReplyTo = testUser;

            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.ITEM_IN_REPLY_TO);
            expect(message.InReplyTo).to.be.equal(testUser);
        });

        it('test ITEM_BODY, ITEM_UNIQUE_BODY, ITEM_NORMALIZED_BODY, ITEM_TEXT_BODY', async () => {
            const manager = new MockEWSManager();
            let message = new MessageType();
            const pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';

            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.ITEM_BODY);
            expect(message.Body?.Value).to.be.equal(parsed.html);

            message = new MessageType();
            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.ITEM_UNIQUE_BODY);
            expect(message.Body?.Value).to.be.equal(parsed.html);

            message = new MessageType();
            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.ITEM_NORMALIZED_BODY);
            expect(message.Body?.Value).to.be.equal(parsed.html);

            message = new MessageType();
            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.ITEM_TEXT_BODY);
            expect(message.Body?.Value).to.be.equal(parsed.html);

            parsed.html = false;
            message = new MessageType();
            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.ITEM_TEXT_BODY);
            expect(message.Body?.Value).to.be.equal(parsed.text);
        });

        it('test MESSAGE_FROM', async () => {
            const manager = new MockEWSManager();
            let message = new MessageType();
            const pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';

            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.MESSAGE_FROM);
            expect(message.From?.Mailbox.EmailAddress).to.be.equal('john.doe@27e8335abbd3.ngrok.io');

            // Try an address with only a name
            if (parsed.from) {
                parsed.from.value = [{ 'address': '', 'name': 'RustyG Mail' }]
            }
            message = new MessageType();
            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.MESSAGE_FROM);
            expect(message.From?.Mailbox.EmailAddress).to.be.equal('rustyg.mail@mail.quattro.rocks');
        });

        it('test ITEM_DATE_TIME_SENT', async () => {
            const manager = new MockEWSManager();
            let message = new MessageType();
            const pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';

            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.ITEM_DATE_TIME_SENT);
            expect(message.DateTimeSent).to.be.eql(parsed.date);

            // Try with no date
            parsed.date = undefined;
            message = new MessageType();
            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.ITEM_DATE_TIME_SENT);
            expect(message.DateTimeSent).to.be.undefined();

        });

        it('test MESSAGE_BCC_RECIPIENTS', async () => {
            const manager = new MockEWSManager();
            const message = new MessageType();
            const pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';
            parsed.bcc = {
                value: [
                    { 'address': testUser, 'name': 'Test User' },
                    { 'address': 'test@hcl.com', 'name': 'Hcl User' }
                ],
                html: 'html',
                text: 'text'
            }

            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.MESSAGE_BCC_RECIPIENTS);
            const bcc = message.BccRecipients?.Mailbox;
            expect(bcc).to.not.be.undefined();
            if (bcc) {
                const r1 = bcc[0];
                expect(r1.EmailAddress).to.be.equal(testUser);
                // FIXME:  Update this when we remove the digital week getname() hack
                expect(r1.Name).to.be.equal(testUser);
                const r2 = bcc[1];
                expect(r2.EmailAddress).to.be.equal('test@hcl.com');
                // FIXME:  Update this when we remove the digital week getName() hack
                expect(r1.Name).to.be.equal(testUser);
            }
        });

        it('test MESSAGE_CC_RECIPIENTS', async () => {
            const manager = new MockEWSManager();
            const message = new MessageType();
            const pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';
            parsed.cc = {
                value: [
                    { 'address': testUser, 'name': 'Test User' },
                    { 'address': 'test@hcl.com', 'name': 'Hcl User' }
                ],
                html: 'html',
                text: 'text'
            }

            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.MESSAGE_CC_RECIPIENTS);
            const cc = message.CcRecipients?.Mailbox;
            expect(cc).to.not.be.undefined();
            if (cc) {
                const r1 = cc[0];
                expect(r1.EmailAddress).to.be.equal(testUser);
                // FIXME:  Update this when we remove the digital week getname() hack
                expect(r1.Name).to.be.equal(testUser);
                const r2 = cc[1];
                expect(r2.EmailAddress).to.be.equal('test@hcl.com');
                // FIXME:  Update this when we remove the digital week getName() hack
                expect(r1.Name).to.be.equal(testUser);
            }
        });

        it('test MESSAGE_CONVERSATION_TOPIC', async () => {
            const manager = new MockEWSManager();
            const message = new MessageType();
            const pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';

            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.MESSAGE_CONVERSATION_TOPIC);
            expect(message.ConversationTopic).to.be.equal('Hi');
        });

        it('test MESSAGE_INTERNET_MESSAGE_ID', async () => {
            const manager = new MockEWSManager();
            const message = new MessageType();
            const pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';

            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.MESSAGE_INTERNET_MESSAGE_ID);
            expect(message.InternetMessageId).to.be.equal(parsed.messageId);
        });

        it('test MESSAGE_REFERENCES', async () => {
            const manager = new MockEWSManager();
            const message = new MessageType();
            const pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';

            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.MESSAGE_REFERENCES);
            expect(message.References).to.be.equal(parsed.references?.join(' '));
        });

        it('test MESSAGE_REPLY_TO', async () => {
            const manager = new MockEWSManager();
            const message = new MessageType();
            const pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';
            parsed.replyTo = {
                value: [
                    { 'address': testUser, 'name': 'Test User' },
                    { 'address': 'test@hcl.com', 'name': 'Hcl User' }
                ],
                html: 'html',
                text: 'text'
            }

            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.MESSAGE_REPLY_TO);
            const replyTo = message.ReplyTo?.Mailbox;
            expect(replyTo).to.not.be.undefined();
            if (replyTo) {
                const r1 = replyTo[0];
                expect(r1.EmailAddress).to.be.equal(testUser);
                // FIXME:  Update this when we remove the digital week getname() hack
                expect(r1.Name).to.be.equal(testUser);
                const r2 = replyTo[1];
                expect(r2.EmailAddress).to.be.equal('test@hcl.com');
                // FIXME:  Update this when we remove the digital week getName() hack
                expect(r1.Name).to.be.equal(testUser);
            }
        });

        it('test MESSAGE_SENDER', async () => {
            const manager = new MockEWSManager();
            const message = new MessageType();
            const pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';

            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.MESSAGE_SENDER);
            expect(message.Sender?.Mailbox.EmailAddress).to.be.equal('john.doe@27e8335abbd3.ngrok.io');
        });

        it('test MESSAGE_TO_RECIPIENTS', async () => {
            const manager = new MockEWSManager();
            const message = new MessageType();
            const pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';
            parsed.to = {
                value: [
                    { 'address': testUser, 'name': 'Test User' },
                    { 'address': 'test@hcl.com', 'name': 'Hcl User' }
                ],
                html: 'html',
                text: 'text'
            }

            await manager.updateEWSItemFieldValueFromMime(message, parsed, pimMessage, UnindexedFieldURIType.MESSAGE_TO_RECIPIENTS);
            const to = message.ToRecipients?.Mailbox;
            if (to) {
                const r1 = to[0];
                expect(r1.EmailAddress).to.be.equal(testUser);
                // FIXME:  Update this when we remove the digital week getname() hack
                expect(r1.Name).to.be.equal(testUser);
                const r2 = to[1];
                expect(r2.EmailAddress).to.be.equal('test@hcl.com');
                // FIXME:  Update this when we remove the digital week getName() hack
                expect(r1.Name).to.be.equal(testUser);
            }
        });
    });

    describe('Test addRequestedPropertiesToEWSMeetingMessage', () => {

        it('Test with iCal content', () => {

            const pimMessage = PimItemFactory.newPimMessage({});
            let message = new MeetingRequestMessageType();
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            const context = getContext(testUser, 'password');

            const extendedProp = {
                DistinguishedPropertySetId: 'Address',
                PropertyId: 32899,
                PropertyType: 'String',
                Value: 'kermit.frog@fakemail.com'
            }
            pimMessage.icalMessageId = 'MSG-ID';
            pimMessage.icalStream = iCalDeclineResponse;
            pimMessage.addExtendedProperty(extendedProp);
            pimMessage.from = testUser;
            pimMessage.unid = 'UNID';
            pimMessage.parentFolderIds = ['PARENT'];
            pimMessage.subject = 'SUBJECT';

            const path1 = new PathToUnindexedFieldType();
            path1.FieldURI = UnindexedFieldURIType.ITEM_SUBJECT;
            const path2 = new PathToUnindexedFieldType();
            path2.FieldURI = UnindexedFieldURIType.ITEM_SIZE;
            const path3 = new PathToUnindexedFieldType();
            path3.FieldURI = UnindexedFieldURIType.CALENDAR_IS_ORGANIZER;

            const extPath = new PathToExtendedFieldType();
            extPath.PropertyId = 32899;
            extPath.PropertyType = MapiPropertyTypeType.STRING;

            shape.AdditionalProperties = new NonEmptyArrayOfPathsToElementType([path1, path2, path3, extPath]);

            const mailboxId = 'test@test.com';

            const manager = new MockEWSManager();
            manager.addRequestedPropertiesToEWSMeetingMessage(pimMessage, message, shape, context.request, mailboxId);

            const itemEWSId = getEWSId(pimMessage.unid, mailboxId);
            const parentEWSId = getEWSId(pimMessage.parentFolderIds[0], mailboxId);

            expect(message.ItemId?.Id).to.be.equal(itemEWSId);
            expect(message.ParentFolderId?.Id).to.be.equal(parentEWSId);
            expect(message.Subject).to.be.equal(pimMessage.subject);

            // All properties
            message = new MeetingRequestMessageType();
            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
            manager.addRequestedPropertiesToEWSMeetingMessage(pimMessage, message, shape, context.request, mailboxId);
            expect(message.ItemId?.Id).to.be.equal(itemEWSId);
            expect(message.ParentFolderId?.Id).to.be.equal(parentEWSId);
            expect(message.Subject).to.be.equal(pimMessage.subject);

            // No additional properties
            message = new MeetingRequestMessageType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            shape.AdditionalProperties = undefined;
            manager.addRequestedPropertiesToEWSMeetingMessage(pimMessage, message, shape, context.request, mailboxId);
            expect(message.ItemId?.Id).to.be.equal(itemEWSId);
            expect(message.ParentFolderId?.Id).to.be.equal(parentEWSId);
            expect(message.Subject).to.be.undefined();

            // No additional properties, but ALL_PROPERTIES shape
            message = new MeetingRequestMessageType();
            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
            manager.addRequestedPropertiesToEWSMeetingMessage(pimMessage, message, shape, context.request, mailboxId);
            expect(message.ItemId?.Id).to.be.equal(itemEWSId);
            expect(message.ParentFolderId?.Id).to.be.equal(parentEWSId);
            expect(message.Subject).to.be.equal(pimMessage.subject);

            if (message.ExtendedProperty) {
                expect(message.ExtendedProperty.length).to.be.equal(1);
                const ext: ExtendedPropertyType = message.ExtendedProperty[0];
                expect(ext).is.not.undefined();
                expect(ext.Value).to.be.equal('kermit.frog@fakemail.com');
                expect(ext.ExtendedFieldURI.PropertyId).to.be.equal(32899);
                expect(ext.ExtendedFieldURI.DistinguishedPropertySetId).to.be.equal('Address');
                expect(ext.ExtendedFieldURI.PropertyType).to.be.equal('String');
            }
        });
    });

    describe('Test updateEWSItemFieldValueForMeeting', () => {

        let pimMessage: PimMessage;
        let meetingRequest: MeetingRequestMessageType;
        let meetingResponse: MeetingResponseMessageType;
        const context = getContext(testUser, 'password');
        let manager: MockEWSManager;

        beforeEach(function () {

            const extendedProp = {
                DistinguishedPropertySetId: 'Address',
                PropertyId: 32899,
                PropertyType: 'String',
                Value: 'kermit.frog@fakemail.com'
            }

            pimMessage = PimItemFactory.newPimMessage({});
            pimMessage.unid = 'UNID';
            pimMessage.parentFolderIds = ['PARENT-FOLDER'];
            pimMessage.subject = 'SUBJECT';
            pimMessage.body = 'BODY';
            pimMessage.addExtendedProperty(extendedProp);
            pimMessage.from = testUser;

            meetingRequest = new MeetingRequestMessageType();
            meetingResponse = new MeetingResponseMessageType();
            manager = new MockEWSManager();

        });

        it('Test UnindexedFieldURIType.CALENDAR_IS_ORGANIZER', () => {
            let handled = manager.updateEWSItemFieldValueForMeeting(meetingRequest, pimMessage, UnindexedFieldURIType.CALENDAR_IS_ORGANIZER, context.request, {});
            expect(handled).to.be.true();
            expect(meetingRequest.IsOrganizer).to.be.true();

            pimMessage.from = 'joe@test.com';
            meetingRequest = new MeetingRequestMessageType();
            handled = manager.updateEWSItemFieldValueForMeeting(meetingRequest, pimMessage, UnindexedFieldURIType.CALENDAR_IS_ORGANIZER, context.request, {});
            expect(handled).to.be.true();
            expect(meetingRequest.IsOrganizer).to.be.false();

        });

        it('Test UnindexedFieldURIType.MEETING_IS_DELEGATED', () => {
            pimMessage.noticeType = PimNoticeTypes.REQUEST_DELEGATED;

            let handled = manager.updateEWSItemFieldValueForMeeting(meetingRequest, pimMessage, UnindexedFieldURIType.MEETING_IS_DELEGATED, context.request, {});
            expect(handled).to.be.true();
            expect(meetingRequest.IsDelegated).to.be.true();

            pimMessage.noticeType = PimNoticeTypes.INVITATION_REQUEST;
            meetingRequest = new MeetingRequestMessageType();
            handled = manager.updateEWSItemFieldValueForMeeting(meetingRequest, pimMessage, UnindexedFieldURIType.MEETING_IS_DELEGATED, context.request, {});
            expect(handled).to.be.true();
            expect(meetingRequest.IsDelegated).to.be.undefined();
        });

        it('Test UnindexedFieldURIType.MEETING_ASSOCIATED_CALENDAR_ITEM_ID', () => {
            let handled = manager.updateEWSItemFieldValueForMeeting(
                meetingResponse, 
                pimMessage, 
                UnindexedFieldURIType.MEETING_ASSOCIATED_CALENDAR_ITEM_ID, 
                context.request, 
                {}
            );
            expect(handled).to.be.true();
            expect(meetingResponse.AssociatedCalendarItemId).to.be.undefined();
            
            const mailboxId = 'test@test.com';

            pimMessage.referencedCalendarItemUnid = 'CAL-ITEM';
            handled = manager.updateEWSItemFieldValueForMeeting(
                meetingResponse, 
                pimMessage, 
                UnindexedFieldURIType.MEETING_ASSOCIATED_CALENDAR_ITEM_ID, 
                context.request, 
                {},
                mailboxId
            );
            const itemEWSId = getEWSId('CAL-ITEM', mailboxId);

            expect(handled).to.be.true();
            expect(meetingResponse.AssociatedCalendarItemId?.Id).to.be.equal(itemEWSId);
        });

        it('Test UnindexedFieldURIType.CALENDAR_START', () => {
            const start = new Date();
            let handled = manager.updateEWSItemFieldValueForMeeting(meetingResponse, pimMessage, UnindexedFieldURIType.CALENDAR_START, context.request, {});
            expect(handled).to.be.true();
            expect(meetingResponse.Start).to.be.undefined();

            handled = manager.updateEWSItemFieldValueForMeeting(meetingRequest, pimMessage, UnindexedFieldURIType.CALENDAR_START, context.request, {});
            expect(handled).to.be.true();
            expect(meetingRequest.Start).to.be.undefined();

            handled = manager.updateEWSItemFieldValueForMeeting(meetingResponse, pimMessage, UnindexedFieldURIType.CALENDAR_START, context.request, { 'Start': start });
            expect(handled).to.be.true();
            expect(meetingResponse.Start).to.be.eql(start);

            handled = manager.updateEWSItemFieldValueForMeeting(meetingRequest, pimMessage, UnindexedFieldURIType.CALENDAR_START, context.request, { 'Start': start });
            expect(handled).to.be.true();
            expect(meetingRequest.Start).to.be.eql(start);
        });

        it('Test UnindexedFieldURIType.CALENDAR_END', () => {
            const end = new Date();
            let handled = manager.updateEWSItemFieldValueForMeeting(meetingResponse, pimMessage, UnindexedFieldURIType.CALENDAR_END, context.request, {});
            expect(handled).to.be.true();
            expect(meetingResponse.End).to.be.undefined();

            handled = manager.updateEWSItemFieldValueForMeeting(meetingRequest, pimMessage, UnindexedFieldURIType.CALENDAR_END, context.request, {});
            expect(handled).to.be.true();
            expect(meetingRequest.End).to.be.undefined();

            handled = manager.updateEWSItemFieldValueForMeeting(meetingResponse, pimMessage, UnindexedFieldURIType.CALENDAR_END, context.request, { 'End': end });
            expect(handled).to.be.true();
            expect(meetingResponse.End).to.be.eql(end);

            handled = manager.updateEWSItemFieldValueForMeeting(meetingRequest, pimMessage, UnindexedFieldURIType.CALENDAR_END, context.request, { 'End': end });
            expect(handled).to.be.true();
            expect(meetingRequest.End).to.be.eql(end);
        });

        it('Test UnindexedFieldURIType.MEETING_PROPOSED_START', () => {
            const start = new Date();
            let handled = manager.updateEWSItemFieldValueForMeeting(meetingResponse, pimMessage, UnindexedFieldURIType.MEETING_PROPOSED_START, context.request, {});
            expect(handled).to.be.true();
            expect(meetingResponse.ProposedStart).to.be.undefined();

            pimMessage.noticeType = PimNoticeTypes.COUNTER_PROPOSAL_REQUEST;
            handled = manager.updateEWSItemFieldValueForMeeting(meetingRequest, pimMessage, UnindexedFieldURIType.MEETING_PROPOSED_START, context.request, {});
            expect(handled).to.be.true();
            expect(meetingResponse.ProposedStart).to.be.undefined();

            pimMessage.newStartDate = start;
            handled = manager.updateEWSItemFieldValueForMeeting(meetingResponse, pimMessage, UnindexedFieldURIType.MEETING_PROPOSED_START, context.request, {});
            expect(handled).to.be.true();
            expect(meetingResponse.ProposedStart).to.be.eql(start);

        });

        it('Test UnindexedFieldURIType.MEETING_PROPOSED_END', () => {
            const end = new Date();
            let handled = manager.updateEWSItemFieldValueForMeeting(meetingResponse, pimMessage, UnindexedFieldURIType.MEETING_PROPOSED_END, context.request, {});
            expect(handled).to.be.true();
            expect(meetingResponse.ProposedEnd).to.be.undefined();

            pimMessage.noticeType = PimNoticeTypes.COUNTER_PROPOSAL_REQUEST;
            handled = manager.updateEWSItemFieldValueForMeeting(meetingRequest, pimMessage, UnindexedFieldURIType.MEETING_PROPOSED_END, context.request, {});
            expect(handled).to.be.true();
            expect(meetingResponse.ProposedEnd).to.be.undefined();

            pimMessage.newEndDate = end;
            handled = manager.updateEWSItemFieldValueForMeeting(meetingResponse, pimMessage, UnindexedFieldURIType.MEETING_PROPOSED_END, context.request, {});
            expect(handled).to.be.true();
            expect(meetingResponse.ProposedEnd).to.be.eql(end);
        });

        it('Test UnindexedFieldURIType.MEETING_REQUEST_MEETING_REQUEST_TYPE', () => {
            let handled = manager.updateEWSItemFieldValueForMeeting(meetingRequest, pimMessage, UnindexedFieldURIType.MEETING_REQUEST_MEETING_REQUEST_TYPE, context.request, {});
            expect(handled).to.be.true();
            expect(meetingRequest.MeetingRequestType).to.be.equal(MeetingRequestTypeType.NEW_MEETING_REQUEST);

            handled = manager.updateEWSItemFieldValueForMeeting(meetingResponse, pimMessage, UnindexedFieldURIType.MEETING_REQUEST_MEETING_REQUEST_TYPE, context.request, {});
            expect(handled).to.be.true();
        });

        it('Test UnindexedFieldURIType.CALENDAR_IS_MEETING', () => {
            let handled = manager.updateEWSItemFieldValueForMeeting(meetingRequest, pimMessage, UnindexedFieldURIType.CALENDAR_IS_MEETING, context.request, {});
            expect(handled).to.be.true();
            expect(meetingRequest.IsMeeting).to.be.true();

            handled = manager.updateEWSItemFieldValueForMeeting(meetingResponse, pimMessage, UnindexedFieldURIType.CALENDAR_IS_MEETING, context.request, {});
            expect(handled).to.be.true();
        });

        it('Test UnindexedFieldURIType.CALENDAR_IS_CANCELLED', () => {
            let handled = manager.updateEWSItemFieldValueForMeeting(meetingRequest, pimMessage, UnindexedFieldURIType.CALENDAR_IS_CANCELLED, context.request, {});
            expect(handled).to.be.true();
            expect(meetingRequest.IsCancelled).to.be.false();

            handled = manager.updateEWSItemFieldValueForMeeting(meetingResponse, pimMessage, UnindexedFieldURIType.CALENDAR_IS_CANCELLED, context.request, {});
            expect(handled).to.be.true();
        });

        it('Test UnindexedFieldURIType.CALENDAR_CALENDAR_ITEM_TYPE', () => {
            let handled = manager.updateEWSItemFieldValueForMeeting(meetingRequest, pimMessage, UnindexedFieldURIType.CALENDAR_CALENDAR_ITEM_TYPE, context.request, {});
            expect(handled).to.be.true();
            expect(meetingRequest.CalendarItemType).to.be.equal(CalendarItemTypeType.SINGLE);

            handled = manager.updateEWSItemFieldValueForMeeting(meetingResponse, pimMessage, UnindexedFieldURIType.CALENDAR_CALENDAR_ITEM_TYPE, context.request, {});
            expect(handled).to.be.true();
        });

        it('Test UnindexedFieldURIType.ITEM_MIME_CONTENT', () => {
            pimMessage.cc = ['test1@hcl.com', 'test2@hcl.com'];
            pimMessage.bcc = ['bcc@hcl.com'];
            pimMessage.replyTo = ['replyto@hcl.com'];
            pimMessage.to = ['to1@hcl.com', 'to2@hcl.com'];
            const handled = manager.updateEWSItemFieldValueForMeeting(meetingRequest, pimMessage, UnindexedFieldURIType.ITEM_MIME_CONTENT, context.request, {});
            expect(handled).to.be.true();
            expect(meetingRequest.MimeContent).to.not.be.undefined();
        });

    });

    describe('Test getCancelCalendarItemResponse', () => {

        const context = getContext(testUser, 'password');
        const userInfo = UserContext.getUserInfo(context.request);
        const manager = new MockEWSManager();

        afterEach(() => {
            sinon.restore(); 
        });

        it('returns ews calendar item if the referenced calendar item exists', async () => {
            const pimCalendarItem = generateTestCalendarItems(KeepPimConstants.DEFAULT_CALENDAR_NAME, 1)[0];
            pimCalendarItem.unid = 'test-unid';
            const pimManagerStub = sinon.createStubInstance(KeepPimManager);
            pimManagerStub.getPimItem.resolves(pimCalendarItem);
            stubPimManager(pimManagerStub);

            const mailboxId = 'test@test.com';
            const referencedItemId = getEWSId(pimCalendarItem.unid, mailboxId);

            const result = await manager.getCancelCalendarItemResponse(referencedItemId, userInfo, context.request)

            expect(result).to.be.an.Array();
            expect(result).to.have.length(1);

            if (result) {
                expect(result[0]).to.be.instanceof(CalendarItemType);
                expect(result[0].ItemId?.Id).to.be.equal(referencedItemId);
            }
        });

        it('returns undefined if the referenced calendar item doesnt exist', async () => {
            const pimManagerStub = sinon.createStubInstance(KeepPimManager);
            pimManagerStub.getPimItem.resolves(undefined);
            stubPimManager(pimManagerStub);

            const itemId = 'test-unid'
            const mailboxId = 'test@test.com';
            const referencedItemId = getEWSId(itemId, mailboxId);

            const result = await manager.getCancelCalendarItemResponse(referencedItemId, userInfo, context.request)

            expect(result).to.be.undefined()
        });

        it('throws an error if an invalid reference id was provided', async () => {
            const invalidReferencedItemId = '';

            const expectedError = new Error(`Error cancelling calendar item. referencedItemId, ${invalidReferencedItemId}, is an invalid id`);
            
            await expect(manager.getCancelCalendarItemResponse(invalidReferencedItemId, userInfo, context.request)).to.be.rejectedWith(expectedError);
        });
    });
});
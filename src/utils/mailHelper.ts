import { Request } from '@loopback/rest';
import { simpleParser } from 'mailparser';
import { getEmail } from '.';
import { UserMailboxData, USER_MAILBOX_DATA } from '../data/mail';
import { BodyTypeType } from '../models/enum.model';
import {
    ArrayOfRecipientsType, BodyType, EmailAddressType, FileAttachmentType, ItemAttachmentType, MessageType,
    NonEmptyArrayOfAttachmentsType, SingleRecipientType} from '../models/mail.model';
import { Logger } from './logger';
import { v4 as uuidv4 } from 'uuid';
import util from 'util';

/**
 * Returns a user id from a SOAP request. 
 * @param request The SOAP request.
 * @returns The userid or undefined if not found
 */
export function getUserId(request: Request): string | undefined {
    /* Outlook includes a X-User-Identity or X-Anchormailbox header that contains the email address.
    However since we now require authorization, we will always use the authorization header to determine the userid. 
     */

    let userId: string | undefined;
    // Attempt to get the userid from the authorization token
    const authInfo = request.header("authorization")?.split(" ");
    if (authInfo !== undefined) {
        if (authInfo.length > 1) {
            if (authInfo[0] === "Basic") {
                const buffer = Buffer.from(authInfo[1], "base64");
                const unencoded = buffer.toString("utf8");
                const cred = unencoded.split(":");
                userId = cred[0];
            }
            else if (authInfo[0] === "Bearer") {
                // TBD
                throw new Error("Bearer token is not supported");
            }
        }
    }

    return userId;
}

/**
* Get user mailbox data.
* @param request The HTTP request
* @param folders An array of folders from Domino. If not specified, mock folder data will be used.
* @returns user mailbox data
*/
export function getUserMailData(request: Request): UserMailboxData {
    const userId = getUserId(request);
    if (!userId) {
        throw new Error('User identity not found');
    }

    USER_MAILBOX_DATA[userId] = USER_MAILBOX_DATA[userId] ?? new UserMailboxData(userId);
    return USER_MAILBOX_DATA[userId];
}

/**
 * Populate a MessageType object with the settings parsed from its mime content.
 * @param userData Used to access the current user's email address for the from field.
 * @param item The item we're populating.
 * @param mimeContent Mime content to parse and use to populate the fields in item.
 * @param encoding Encoding of the mime content.
 * @returns The updated item.
 */
export async function populateItemFromMime(userData: UserMailboxData, item: MessageType, mimeContent: string, encoding = 'base64'): Promise<MessageType> {
    // Parse mime content
    const messageBuffer = Buffer.from(mimeContent, encoding);
    const parsed = await simpleParser(messageBuffer);
    Logger.getInstance().debug(`Mime content: ${util.inspect(parsed, false, 5)}`);

    // Don't override it if it is already set...may have been set by messageToItem
    if (item.Size === undefined) {
        item.Size = messageBuffer.length;
    }

    item.Subject = parsed.subject;
    item.InReplyTo = parsed.inReplyTo;
    item.InternetMessageId = parsed.messageId;
    if (parsed.date) {
        item.DateTimeSent = parsed.date;
    }

    if (parsed.headers.has('thread-topic')) {
        item.ConversationTopic = parsed.headers.get('thread-topic')?.toString();
    }
    
    if (parsed.to && parsed.to.value && Array.isArray(parsed.to.value)) {
        item.ToRecipients = getMimeRecipients(parsed.to.value);
    }

    if (parsed.cc && parsed.cc.value && Array.isArray(parsed.cc.value)) {
        item.CcRecipients = getMimeRecipients(parsed.cc.value);
    }

    if (parsed.bcc && parsed.bcc.value && Array.isArray(parsed.bcc.value)) {
        item.BccRecipients = getMimeRecipients(parsed.bcc.value);
    }

    if (parsed.replyTo && parsed.replyTo.value && Array.isArray(parsed.replyTo.value)) {
        item.ReplyTo = getMimeRecipients(parsed.replyTo.value);
    }

    if (parsed.from && parsed.from.value) {
        let val: any = {};
        if (Array.isArray(parsed.from.value)) {
            val = parsed.from.value[0];
        } else {
            val = parsed.from.value;
        }
        const recipient = new SingleRecipientType();
        recipient.Mailbox = getEmail(val.address.length > 0 ? val.address : val.name);
        if (!recipient.Mailbox.EmailAddress) {
            recipient.Mailbox.EmailAddress = val.address;
        }
        if (val.name && val.name.length > 0) {
            recipient.Mailbox.Name = val.name;
        }
        // FIXME: Hack to avoid domino name showing for digital week
        if (recipient.Mailbox.EmailAddress) {
            recipient.Mailbox.Name = recipient.Mailbox.EmailAddress;
        }

        item.From = recipient;
        item.Sender = recipient;
    }

    // ConversationIndex must come from the mime or be set in the item.  If there is an itemId we could look it up from the server IF it's an existing item
    // if (!item.ConversationIndex) {
    //     if (!item.InReplyTo) {
    //         item.ConversationIndex = uuidv4();
    //     } else {
    //         for (const _item of Object.values(userData.ITEMS)) {
    //             if (_item instanceof MessageType && _item.InternetMessageId === item.InReplyTo) {
    //                 item.ConversationIndex = _item.ConversationIndex;
    //                 break;
    //             }
    //         }
    //     }
    // }
    if (parsed.attachments && parsed.attachments.length) {
        item.HasAttachments = true;
        item.Attachments = new NonEmptyArrayOfAttachmentsType();

        for (const att of parsed.attachments) {
            if (att.contentDisposition === 'inline') {
                continue;
            }

            if (att.contentType === 'message/rfc822') {
                const parsedItem = await simpleParser(att.content);
                // for (const _item of Object.values(userData.ITEMS)) {
                //     if (_item instanceof MessageType && _item.InternetMessageId === parsedItem.messageId) {

                const attchment = new ItemAttachmentType();
                item.Attachments.push(attchment);

                // FIXME: We should lookup the attachment by InternetMessageId and store the unid as the ContentId
                attchment.ContentId = uuidv4();
                attchment.ContentType = att.contentType;
                attchment.Name = parsedItem.subject;

                //         UserMailboxData.ATTACHMENTS_CONTENTS[attchment.ContentId] = deepcoy(_item);
                //         break;
                //     }
                // }
            } else {
                const attchment = new FileAttachmentType();
                item.Attachments.push(attchment);

                attchment.ContentId = uuidv4();
                attchment.ContentType = att.contentType;
                attchment.Size = att.size;
                attchment.Name = att.filename;

                // UserMailboxData.ATTACHMENTS_CONTENTS[attchment.ContentId] = att.content.toString('base64');
            }
        }
    } else {
        item.HasAttachments = false;
    }
    if (parsed.html) {
        item.Body = new BodyType(parsed.html, BodyTypeType.HTML);
    } else {
        item.Body = new BodyType(parsed.text ?? '', BodyTypeType.TEXT);
    }

    return item;
}

// internal functions
export function getMimeRecipients(parseArray: any[]): ArrayOfRecipientsType | undefined {
    let recipients: ArrayOfRecipientsType | undefined;
    if (parseArray.length > 0) {
        recipients = new ArrayOfRecipientsType();
        const addressTypes: EmailAddressType[] = [];
        for (const val of parseArray) {
            const emailType = getEmail(val.address.length > 0 ? val.address : val.name);
            if (!emailType.EmailAddress) {
                emailType.EmailAddress = val.address;
            }
            if (val.name && val.name.length > 0) {
                emailType.Name = val.name;
            }
            // FIXME: Hack for digital week demo to avoid domino names.
            if (emailType.EmailAddress) {
                emailType.Name = emailType.EmailAddress;
            }
            addressTypes.push(emailType);
        }
        recipients.Mailbox = addressTypes;
    }
    return recipients;
}
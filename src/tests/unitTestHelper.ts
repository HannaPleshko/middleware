import { sinon, ExpressContextStub, stubExpressContext } from '@loopback/testlab';
import { FolderEvent, UserMailboxData, USER_MAILBOX_DATA } from '../data/mail';
import {
    KeepPimLabelManager, KeepPimManager, KeepPimMessageManager, KeepPimNotebookManager, PimLabel, base64Encode,
    PimItemFactory, PimItemFormat, KeepPimContactsManager, KeepPimTasksManager, KeepPimCalendarManager, PimLabelTypes, KeepPimConstants, PimCalendarItem, PimItem
} from '@hcllabs/openclientkeepcomponent';
import { FolderIdType, FolderType, ItemType, SyncFolderHierarchyCreateType } from '../models/mail.model';
import { EWSContactsManager, EWSMessageManager, EWSServiceManager, EWSTaskManager, getEWSId } from '../utils';
import { createStubInstance, StubbableType, SinonStubbedInstance, SinonStubbedMember } from 'sinon';

/**
 * Returns a context that can be use to create Loopback Reqest and Response objects using in unit tests. 
 * @param user The test user
 * @param password The test user's password
 * @returns A context object that can be used to get a Request (context.request) or a Response (context.response).
 */
export function getContext(user: string, password: string): ExpressContextStub {
    const authorizationString = `${user}:${password}`;
    return stubExpressContext({
        method: 'POST',
        url: '/EWS/Exchange.asmx',
        headers: { "authorization": `Basic ${base64Encode(authorizationString)}` }
    });
}

/**
 * Returns a user mail box. This will create one if one does not exist for the user. 
 * @param user The test user. 
 */
export function createUserMailBox(user: string): UserMailboxData {
    if (USER_MAILBOX_DATA[user] === undefined) {
        USER_MAILBOX_DATA[user] = new UserMailboxData(user);
    }

    return USER_MAILBOX_DATA[user];
}

/**
 * Resets a user mailbox to the initial state. 
 * @param user The test user
 */
export function resetUserMailBox(user: string): void {
    USER_MAILBOX_DATA[user] = new UserMailboxData(user);
}

/**
 * Generate folder create events. 
 * @param user The test user id
 * @param amount The number of folder create events to generate. The default is one. Each folder will have a folder id called "folder-n", where n it the number of the folder starting at 1.
 * @returns An array of pim labels that represent the folders added to the event queue. 
 */
export function generateFolderCreateEvents(user: string, amount = 1, delegate?: string): PimLabel[] {
    const folders: PimLabel[] = [];
    const userData = createUserMailBox(user);
    for (let i = 0; i < amount; i++) {
        const unid = `folder-${i}`;
        const change = new SyncFolderHierarchyCreateType();
        change.Folder = new FolderType();
        change.Folder.DisplayName = `Unit Test Folder ${i}`;
        folders.push(
            PimItemFactory.newPimLabel(
                { 
                    FolderId: unid,
                    DisplayName: change.Folder.DisplayName 
                }, 
            PimItemFormat.DOCUMENT)
        );
        const folderEWSId = getEWSId(unid, delegate);
        change.Folder.FolderId = new FolderIdType(folderEWSId, `ck-${folderEWSId}`);
        userData.FOLDER_EVENTS[folderEWSId] = new FolderEvent(Date.now(), change);
    }

    return folders;
}

/**
 * Stub out the KeepPimLabelManager to avoid calls to Keep. 
 * @param stub A stubbed instance of KeepPimLabelsManager that will be returned when code get an instance of KeepPimLabelManager.
 */
export function stubLabelsManager(stub: sinon.SinonStubbedInstance<KeepPimLabelManager>): void {
    sinon.stub(KeepPimLabelManager, 'getInstance').callsFake(() => {
        console.info("Creating stub for KeepPimLabelManager");
        return stub;
    });
}

/**
 * Stub out the KeepPimNoteboookManager to avoid calls to Keep. 
 * @param stub A stubbed instance of KeepPimNotebelManager that will be returned when code get an instance of KeepPimNotebookManager.
 */
export function stubNotebookManager(stub: sinon.SinonStubbedInstance<KeepPimNotebookManager>): void {
    sinon.stub(KeepPimNotebookManager, 'getInstance').callsFake(() => {
        console.info('Creating stub for KeepPimNotebookManager');
        return stub;
    });
}

/**
 * Stub out the KeepPimManager to avoid calls to Keep. 
 * @param stub A stubbed instance of KeepPimManager that will be returned when code get an instance of KeepPimManager.
 */
export function stubPimManager(stub: sinon.SinonStubbedInstance<KeepPimManager>): void {
    sinon.stub(KeepPimManager, 'getInstance').callsFake(() => {
        console.info('Creating stub for KeepPimManager');
        return stub;
    });
}

/**
 * Stub out the KeepPimMessageManager to avoid calls to Keep. 
 * @param stub A stubbed instance of KeepPimNotebelManager that will be returned when code get an instance of KeepPimMessageManager.
 */
export function stubMessageManager(stub: sinon.SinonStubbedInstance<KeepPimMessageManager>): void {
    sinon.stub(KeepPimMessageManager, 'getInstance').callsFake(() => {
        console.info('Creating stub for KeepPimMessageManager');
        return stub;
    });
}

/**
 * Stub out the KeepPimContactsManager to avoid calls to Keep. 
 * @param stub A stubbed instance of KeepPimContactsManager that will be returned when code get an instance of KeepPimContactsManager.
 */
export function stubContactsManager(stub: sinon.SinonStubbedInstance<KeepPimContactsManager>): void {
    sinon.stub(KeepPimContactsManager, 'getInstance').callsFake(() => {
        console.info('Creating stub for KeepPimContactsManager');
        return stub;
    });
}

/**
 * Stub out the KeepPimTasksManager to avoid calls to Keep. 
 * @param stub A stubbed instance of KeepPimTasksManager that will be returned when code get an instance of KeepPimTasksManager.
 */
export function stubTasksManager(stub: sinon.SinonStubbedInstance<KeepPimTasksManager>): void {
    sinon.stub(KeepPimTasksManager, 'getInstance').callsFake(() => {
        console.info('Creating stub for KeepPimTasksManager');
        return stub;
    });
}

/**
 * Stub out the KeepPimCalendarManager to avoid calls to Keep. 
 * @param stub A stubbed instance of KeepPimCalendarManager that will be returned when code get an instance of KeepPimCalendarManager.
 */
export function stubCalendarManager(stub: sinon.SinonStubbedInstance<KeepPimCalendarManager>): void {
    sinon.stub(KeepPimCalendarManager, 'getInstance').callsFake(() => {
        console.info('Creating stub for KeepPimCalendarManager');
        return stub;
    });
}

/**
 * Label helpers
 */

/**
 * Generate a test inbox label
 * @returns A pim label for the inbox
 */
export function generateTestInboxLabel(): PimLabel {
    const label = PimItemFactory.newPimLabel({ FolderId: 'inbox-id' });
    label.displayName = 'Inbox';
    label.type = PimLabelTypes.MAIL;
    label.view = KeepPimConstants.INBOX;
    return label;
}

/**
 * Generate a test sent label
 * @returns A pim label for the sent view
 */
export function generateTestSentLabel(): PimLabel {
    const label = PimItemFactory.newPimLabel({ FolderId: 'sent-id' });
    label.displayName = 'Sent';
    label.type = PimLabelTypes.MAIL;
    label.view = KeepPimConstants.SENT;
    return label;
}

/**
 * Generate a test drafts label
 * @returns A pim label for the draft view
 */
export function generateTestDraftsLabel(): PimLabel {
    const label = PimItemFactory.newPimLabel({ FolderId: 'drafts-id' });
    label.displayName = 'Drafts';
    label.type = PimLabelTypes.MAIL;
    label.view = KeepPimConstants.DRAFTS;
    return label;
}

/**
 * Generate a test trash label
 * @returns A pim label for the trash folder
 */
export function generateTestTrashLabel(): PimLabel {
    const label = PimItemFactory.newPimLabel({ FolderId: 'trash-id' });
    label.displayName = 'Trash';
    label.type = PimLabelTypes.MAIL;
    label.view = KeepPimConstants.TRASH;
    return label;
}

/**
 * Generate a test calendar label
 * @param secondary True to generate a secondary calendar folder
 * @returns A pim label for the calendar folder
 */
export function generateTestCalendarLabel(secondary = false): PimLabel {
    let label: PimLabel;
    if (secondary) {
        label = PimItemFactory.newPimLabel({ FolderId: 'calendar-second-id' });
        label.displayName = 'Calendar 2';
        label.type = PimLabelTypes.CALENDAR;
        label.view = `(NotesCalendar)\\${label.displayName}`;
    }
    else {
        label = PimItemFactory.newPimLabel({ FolderId: 'calendar-id' });
        label.displayName = 'Calendar';
        label.type = PimLabelTypes.CALENDAR;
        label.view = KeepPimConstants.CALENDAR;
    }
    return label;
}

/**
 * Generate a test tasks label
 * @param secondary True to generate a secondary tasks folder
 * @returns A pim label for the tasks folder
 */
export function generateTestTasksLabel(secondary = false): PimLabel {
    let label: PimLabel;
    if (secondary) {
        label = PimItemFactory.newPimLabel({ FolderId: 'tasks-second-id' });
        label.displayName = 'Tasks 2';
        label.type = PimLabelTypes.TASKS;
        label.view = `(NotesTasks)\\${label.displayName}`;
    }
    else {
        label = PimItemFactory.newPimLabel({ FolderId: 'tasks-id' });
        label.displayName = 'Tasks';
        label.type = PimLabelTypes.TASKS;
        label.view = KeepPimConstants.TASKS;
    }
    return label;
}

/**
 * Generate a test contacts label
 * @param secondary True to generate a secondary contacts folder
 * @returns A pim label for the contacts folder
 */
export function generateTestContactsLabel(secondary = false): PimLabel {
    let label: PimLabel;
    if (secondary) {
        label = PimItemFactory.newPimLabel({ FolderId: 'contacts-second-id' });
        label.displayName = 'Contacts 2';
        label.type = PimLabelTypes.CONTACTS;
        label.view = `(NotesContacts)\\${label.displayName}`;
    }
    else {
        label = PimItemFactory.newPimLabel({ FolderId: 'contacts-id' });
        label.displayName = 'Contacts';
        label.type = PimLabelTypes.CONTACTS;
        label.view = KeepPimConstants.CONTACTS;
    }
    return label;
}

/**
 * Generate a test journal label
 * @param secondary True to generate a secondary journal folder
 * @returns A pim label for the journal folder
 */
export function generateTestNotesLabel(secondary = false): PimLabel {
    let label: PimLabel;
    if (secondary) {
        label = PimItemFactory.newPimLabel({ FolderId: 'journal-second-id' });
        label.displayName = 'Notes 2';
        label.type = PimLabelTypes.JOURNAL;
        label.view = `(NotesJournal)\\${label.displayName}`;
    }
    else {
        label = PimItemFactory.newPimLabel({ FolderId: 'journal-id' });
        label.displayName = 'Notes';
        label.type = PimLabelTypes.JOURNAL;
        label.view = KeepPimConstants.JOURNAL;
    }
    return label;
}

/**
 * Generate test labels.
 * @param includeSecondary True to generate secondary labels for Calendar, Contacts, Tasks, and Notes.
 * @returns 
 */
export function generateTestLabels(includeSecondary = false): PimLabel[] {
    const labels: PimLabel[] = [];

    labels.push(generateTestInboxLabel());
    labels.push(generateTestSentLabel());
    labels.push(generateTestDraftsLabel());
    labels.push(generateTestTrashLabel());
    labels.push(generateTestCalendarLabel());
    labels.push(generateTestTasksLabel());
    labels.push(generateTestContactsLabel());
    labels.push(generateTestNotesLabel());

    if (includeSecondary) {
        labels.push(generateTestCalendarLabel(true));
        labels.push(generateTestTasksLabel(true));
        labels.push(generateTestContactsLabel(true));
        labels.push(generateTestNotesLabel(true));
    }

    return labels;
}

/**
 * Generate test calendar items. 
 * @param calName The name of the calendar the items belong to.
 * @param count The number of items to return. The default is 2. 
 * @returns An array of PimCalendarItems. 
 */
export function generateTestCalendarItems(calName: string, count = 2): PimCalendarItem[] {

    const items: PimCalendarItem[] = [];

    for (let index = 0; index < count; index++) {
        const pimItem = PimItemFactory.newPimCalendarItem({}, calName, PimItemFormat.PRIMITIVE);
        pimItem.unid = `UNID-${index}`;
        pimItem.parentFolderIds = [`PARENT-${index}`];
        pimItem.subject = `SUBJECT-${index}`;
        pimItem.body = `BODY-${index}`;
        items.push(pimItem);
    }

    return items;
}

/** 
* EWS service manager stubs 
*/
/**
* This is a work around for a typescript issue with Sinon. It is documented here: https://github.com/sinonjs/sinon/issues/1963
* Use it to create stubbed instances of classes with private members. 
*/
export type StubbedClass<T> = SinonStubbedInstance<T> & T;
export function createSinonStubInstance<T>(
    constructor: StubbableType<T>,
    overrides?: { [K in keyof T]?: SinonStubbedMember<T[K]> },
): StubbedClass<T> {
    const stub = createStubInstance<T>(constructor, overrides);
    return stub as unknown as StubbedClass<T>;
}
/**
* Stub the getInstanceFromPimItem and getInstanceFromItem for EWSServiceManager to return a stubbed instance of a contacts service manager. 
* @param stub Optional stubbed instance of one of the EWS contacts manager. Use createSinonStubInstance to create this stub. Omit to simulate 
* error where service manager is not returned. 
*/
export function stubServiceManagerForContacts(stub?: StubbedClass<EWSContactsManager>): void {
    sinon.stub(EWSServiceManager, "getInstanceFromPimItem").callsFake((pimItem: PimItem) => {
        console.info('Creating EWSServiceManager stub for getInstanceFromPimItem to return a contacts manager stub');
        return stub;
    });
    sinon.stub(EWSServiceManager, "getInstanceFromItem").callsFake((ewsItem: ItemType) => {
        console.info('Creating EWSServiceManager stub for getInstanceFromItem to return a contacts manager stub');
        return stub;
    });
}
/**
* Stub the getInstanceFromPimItem and getInstanceFromItem for EWSServiceManager to return a stubbed instance of a task service manager. 
* @param stub Optional stubbed instance of one of the EWS task manager. Use createSinonStubInstance to create this stub. Omit to simulate 
* error where service manager is not returned. 
*/
export function stubServiceManagerForTask(stub?: StubbedClass<EWSTaskManager>): void {
    sinon.stub(EWSServiceManager, "getInstanceFromPimItem").callsFake((pimItem: PimItem) => {
        console.info('Creating EWSServiceManager stub for getInstanceFromPimItem to return a task manager stub');
        return stub;
    });
    sinon.stub(EWSServiceManager, "getInstanceFromItem").callsFake((ewsItem: ItemType) => {
        console.info('Creating EWSServiceManager stub for getInstanceFromItem to return a task manager stub');
        return stub;
    });
}
/**
* Stub the getInstanceFromPimItem and getInstanceFromItem for EWSServiceManager to return a stubbed instance of a message service manager. 
* @param stub Optional stubbed instance of one of the EWS messages manager. Use createSinonStubInstance to create this stub. Omit to simulate 
* error where service manager is not returned. 
*/
export function stubServiceManagerForMessages(stub?: StubbedClass<EWSMessageManager>): void {
    sinon.stub(EWSServiceManager, "getInstanceFromPimItem").callsFake((pimItem: PimItem) => {
        console.info('Creating EWSServiceManager stub for getInstanceFromPimItem to return a message manager stub');
        return stub;
    });
    sinon.stub(EWSServiceManager, "getInstanceFromItem").callsFake((ewsItem: ItemType) => {
        console.info('Creating EWSServiceManager stub for getInstanceFromItem to return a message manager stub');
        return stub;
    });
}
/**
* Stub the getInstanceFromPimItem and getInstanceFromItem for EWSServiceManager to return a stubbed instance of a calendar service manager. 
* @param stub Optional stubbed instance of one of the EWS calendar manager. Use createSinonStubInstance to create this stub. Omit to simulate 
* error where service manager is not returned. 
*/
export function stubServiceManagerForCalendar(stub?: StubbedClass<EWSMessageManager>): void {
    sinon.stub(EWSServiceManager, "getInstanceFromPimItem").callsFake((pimItem: PimItem) => {
        console.info('Creating EWSServiceManager stub for getInstanceFromPimItem to return a calendar manager stub');
        return stub;
    });
    sinon.stub(EWSServiceManager, "getInstanceFromItem").callsFake((ewsItem: ItemType) => {
        console.info('Creating EWSServiceManager stub for getInstanceFromItem to return a calendar manager stub');
        return stub;
    });
}

/**
* Compare two ISO date strings.
* @param date1 First ISO date string
* @param date2 Second ISO date string
* @returns 0 if the dates are equal, 1 if the first string is greater than the second, or -1 if the second string is greater than the first.
*/
export function compareDateStrings(date1: string, date2: string): number {
    const first = new Date(date1);
    const second = new Date(date2);

    const results = first.getTime() - second.getTime();
    if (results < 0) {
        return -1; // Second date is greater than the first
    }
    if (results > 0) {
        return 1; // First date is greater than the second
    }
    else {
        return 0; // They are equal
    }
}

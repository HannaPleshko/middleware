/* eslint-disable @typescript-eslint/no-non-null-assertion */
import { expect } from '@loopback/testlab';
import { 
    PimItemFormat, PimTaskPriority, PimItemFactory, PimTask, getTrimmedISODate, KeepPimTasksManager, KeepPimConstants,
    KeepPimMessageManager, KeepMoveMessagesResults 
} from '@hcllabs/openclientkeepcomponent';
import { ArrayOfStringsType, PathToExtendedFieldType, PathToUnindexedFieldType } from '../../models/common.model';
import {
    BodyTypeType, DayOfWeekIndexType, DayOfWeekType, DefaultShapeNamesType, DistinguishedPropertySetType, ImportanceChoicesType, 
    MapiPropertyIds, MapiPropertyTypeType, MonthNamesType, SensitivityChoicesType, TaskStatusType, UnindexedFieldURIType
} from '../../models/enum.model';
import { 
    BodyType, ExtendedPropertyType, ItemResponseShapeType, NonEmptyArrayOfPathsToElementType, TaskType 
} from '../../models/mail.model';
import { EWSServiceManager, EWSTaskManager, getEWSId } from '../../utils';
import { getContext, stubMessageManager, stubTasksManager } from '../unitTestHelper';
import { UserContext } from '../../keepcomponent';
import sinon from 'sinon';

describe('EWSTaskManager tests', function () {
    describe('EWSTaskManager.pimItemFromEWSItem', function () {
        const context = getContext("testUser", "testPass");

        it('item is subject', function () {
            const item = new TaskType();
            item.Subject = "Hi";
            // Provide an existing object to seed values
            const pimItem = EWSTaskManager.getInstance().pimItemFromEWSItem(item, context.request);
            expect(pimItem.title).to.equal("Hi");
            expect(pimItem.subject).to.equal("Hi");
        });

        it('item is body', function () {
            const item = new TaskType();
            item.Body = new BodyType("Hello Team", BodyTypeType.TEXT, false);
            let pimItem = EWSTaskManager.getInstance().pimItemFromEWSItem(item, context.request);
            expect(pimItem.description).to.equal("Hello Team");
            expect(pimItem.body).to.equal("Hello Team");

            // Try with BodyType of HTML
            item.Body = new BodyType("Hello Team", BodyTypeType.HTML, false);
            pimItem = EWSTaskManager.getInstance().pimItemFromEWSItem(item, context.request);
            expect(pimItem.description).to.equal("Hello Team");
            expect(pimItem.body).to.equal("Hello Team");

        });

        it('item is category', function () {
            const item = new TaskType();
            item.Categories = new ArrayOfStringsType();
            item.Categories.String = ["list1", "list2"];
            let pimItem = EWSTaskManager.getInstance().pimItemFromEWSItem(item, context.request);
            expect(pimItem.categories).to.be.deepEqual(["list1", "list2"]);

            // try a single category
            item.Categories.String = ['orange'];
            pimItem = EWSTaskManager.getInstance().pimItemFromEWSItem(item, context.request);
            expect(pimItem.categories).to.be.deepEqual(['orange']);
        });

        it('item is reminderSet', function () {
            const item = new TaskType();
            item.ReminderIsSet = true;
            item.ReminderMinutesBeforeStart = '15';
            item.StartDate = new Date('2021-01-05T19:40:10Z');
            const pimItem = EWSTaskManager.getInstance().pimItemFromEWSItem(item, context.request);

            expect(pimItem.alarm).to.be.equal(-15);
        });

        it('item is reminderSet with extended properties', function () {
            const item = new TaskType();

            const prop = new ExtendedPropertyType();

            prop.ExtendedFieldURI = new PathToExtendedFieldType();
            prop.ExtendedFieldURI.DistinguishedPropertySetId = DistinguishedPropertySetType.COMMON;
            prop.ExtendedFieldURI.PropertyId = MapiPropertyIds.REMINDER_SET;
            prop.ExtendedFieldURI.PropertyType = MapiPropertyTypeType.BOOLEAN;
            prop.Value = "true";

            const prop2 = new ExtendedPropertyType();
            prop2.ExtendedFieldURI = new PathToExtendedFieldType();
            prop2.ExtendedFieldURI.DistinguishedPropertySetId = DistinguishedPropertySetType.COMMON;
            prop2.ExtendedFieldURI.PropertyId = MapiPropertyIds.REMINDER_TIME;
            prop2.ExtendedFieldURI.PropertyType = MapiPropertyTypeType.SYSTEM_TIME;
            prop2.Value = "2021-01-05T18:40:10Z";

            item.ExtendedProperty = [prop, prop2];
            item.StartDate = new Date("2021-01-05T19:40:10Z");
            let pimItem = EWSTaskManager.getInstance().pimItemFromEWSItem(item, context.request);

            expect(pimItem.alarm).to.be.equal(-60);
            expect(pimItem.extendedProperties.length).to.equal(2);

            // With Due date but no start date and a future alarm
            item.StartDate = undefined;
            item.DueDate = new Date('2021-01-05T19:40:10Z');
            prop2.Value = "2021-01-05T20:40:10Z";
            item.ExtendedProperty = [prop, prop2];
            pimItem = EWSTaskManager.getInstance().pimItemFromEWSItem(item, context.request);
            expect(pimItem.alarm).to.be.equal(60);

            // With neither a start date nor a due date
            item.DueDate = undefined;
            pimItem = EWSTaskManager.getInstance().pimItemFromEWSItem(item, context.request);
            expect(pimItem.alarm).to.be.equal(0);
        });

        describe('item is importance', function () {
            const item = new TaskType();

            it('importance is High', function () {
                item.Importance = ImportanceChoicesType.HIGH;
                const pimItem = EWSTaskManager.getInstance().pimItemFromEWSItem(item, context.request);
                expect(pimItem.priority).to.be.equal(PimTaskPriority.HIGH);
            })

            it('importance is Low', function () {
                item.Importance = ImportanceChoicesType.LOW;
                const pimItem = EWSTaskManager.getInstance().pimItemFromEWSItem(item, context.request);
                expect(pimItem.priority).to.be.equal(PimTaskPriority.LOW);
            })

            it('importance is Normal', function () {
                item.Importance = ImportanceChoicesType.NORMAL;
                const pimItem = EWSTaskManager.getInstance().pimItemFromEWSItem(item, context.request);
                expect(pimItem.priority).to.be.equal(PimTaskPriority.MEDIUM);
            });
        });

        it('item is Sensitivity', function () {
            const item = new TaskType();
            item.Sensitivity = SensitivityChoicesType.CONFIDENTIAL;
            const pimItem = EWSTaskManager.getInstance().pimItemFromEWSItem(item, context.request);
            expect(pimItem.isConfidential).to.be.equal(true);
        });

        describe('item instanceOf taskType', function () {
            const item = new TaskType();
            it('item startDate', function () {
                item.StartDate = new Date("1988-06-25T23:00:03Z");
                const pimItem = EWSTaskManager.getInstance().pimItemFromEWSItem(item, context.request);
                expect(pimItem.start).to.not.be.undefined();
                expect(new Date(pimItem.start!)).to.be.deepEqual(item.StartDate);
            });

            it('item dueDate', function () {
                item.DueDate = new Date("1990-06-30T23:00:03Z");
                let pimItem = EWSTaskManager.getInstance().pimItemFromEWSItem(item, context.request);
                expect(pimItem.due).to.not.be.undefined();
                expect(new Date(pimItem.due!)).to.be.deepEqual(item.DueDate);

                // With reminder
                item.ReminderIsSet = true;
                pimItem = EWSTaskManager.getInstance().pimItemFromEWSItem(item, context.request);
                expect(pimItem.due).to.not.be.undefined();
                expect(new Date(pimItem.due!)).to.be.deepEqual(item.DueDate);
            });
        });

        describe('To check values of dueState', function () {
            const item = new TaskType();

            it('dueState completed', function () {
                item.Status = TaskStatusType.COMPLETED;
                item.CompleteDate = new Date('2021-06-30T23:00:03Z');
                let pimItem = EWSTaskManager.getInstance().pimItemFromEWSItem(item, context.request);
                expect(pimItem.isComplete).to.be.equal(true);
                expect(pimItem.completedDate ? getTrimmedISODate(pimItem.completedDate) : undefined).to.be.equal(getTrimmedISODate(item.CompleteDate));

                // Test with no complete date
                item.CompleteDate = undefined;
                pimItem = EWSTaskManager.getInstance().pimItemFromEWSItem(item, context.request);
                expect(pimItem.isComplete).to.be.equal(true);
                expect(pimItem.completedDate ? pimItem.completedDate.toDateString() : undefined).to.be.eql(new Date().toDateString());
            });

            it('dueState is inProgress or waiting on others', function () {
                item.Status = TaskStatusType.IN_PROGRESS;
                item.Status = TaskStatusType.WAITING_ON_OTHERS;
                const pimItem = EWSTaskManager.getInstance().pimItemFromEWSItem(item, context.request);
                expect(pimItem.isInProgress).to.be.equal(true);
            });

            it('dueState is notStarted or deferred', function () {
                item.Status = TaskStatusType.NOT_STARTED;
                item.Status = TaskStatusType.DEFERRED;
                const pimItem = EWSTaskManager.getInstance().pimItemFromEWSItem(item, context.request);
                expect(pimItem.isNotStarted).to.be.equal(true);
            });
        });

        it('item is statusDescription', function () {
            const item = new TaskType();
            item.StatusDescription = "Done";
            const pimItem = EWSTaskManager.getInstance().pimItemFromEWSItem(item, context.request);
            expect(pimItem.statusLabel).to.equal("Done");
        });

    });

    describe('EWSTaskManager.pimItemToEWSItem', function () {
        let pimTask: PimTask;
        const testDescription = "<p>validate shape is body</p>";
        const testCreatedDate = new Date("1988-02-06T22:00:03Z");
        const context = getContext("testUser", "testPass");
        const request = context.request;
        const userInfo = new UserContext();
        const mailboxId = 'test@test.com';

        beforeEach(function () {
            pimTask = PimItemFactory.newPimTask({}, PimItemFormat.DOCUMENT);
            pimTask.unid = "testId";
            pimTask.parentFolderIds = ["pFolder"];
            pimTask.description = testDescription;
            pimTask.categories = [
                "Item1",
                "Item2",
            ];
            pimTask.createdDate = testCreatedDate;
            pimTask.due = "1990-09-06T23:00:03Z";
            pimTask.alarm = 10;
            pimTask.priority = PimTaskPriority.HIGH;
            pimTask.statusLabel = "status";
            pimTask.isConfidential = true;
            pimTask.attachments = [
                "attachment1",
                "attachment2",
            ];
        });

        it('shape id only', async function () {
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            const task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, request, shape, mailboxId);
            expect(task?.ItemId?.Id).to.be.equal(getEWSId("testId", mailboxId));
            expect(task?.ParentFolderId?.Id).to.be.equal(getEWSId("pFolder", mailboxId));
        })

        it('shape is body', async function () {
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.DEFAULT;
            let task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, request, shape, mailboxId);
            expect(task?.Body?.Value).to.be.undefined();
            expect(task?.Body?.BodyType).to.be.undefined();

            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
            task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, request, shape, mailboxId);
            expect(task?.Body?.Value).to.be.equal(testDescription);
            expect(task?.Body?.BodyType).to.be.equal(BodyTypeType.HTML);
        });

        it('shape is categories', async function () {
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
            const task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, request, shape);
            expect(task?.Categories?.String).to.be.deepEqual(["Item1", "Item2"]);
        });

        it('shape is Date', async function () {
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
            const task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, request, shape);
            expect(task?.DateTimeCreated ? getTrimmedISODate(task.DateTimeCreated) : undefined).to.be.equal(pimTask.createdDate ? getTrimmedISODate(pimTask.createdDate) : undefined);
            expect(task?.DateTimeCreated ? getTrimmedISODate(task.DateTimeCreated) : undefined).to.be.equal(getTrimmedISODate(testCreatedDate));
        });

        it('shape is due', async function () {
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
            let task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, request, shape);
            expect(task?.DueDate ? getTrimmedISODate(task.DueDate) : undefined).to.be.equal(pimTask.due ? getTrimmedISODate(pimTask.due) : undefined);

            // Test with empty due date
            pimTask.due = undefined;
            task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, request, shape);
            expect(task?.DueDate).to.be.undefined();
        });

        it('shape is start', async function () {
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
            const task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, request, shape);
            expect(task?.StartDate ? getTrimmedISODate(task.StartDate) : undefined).to.be.equal(pimTask.start ? getTrimmedISODate(pimTask.start) : undefined);

        });


        it('shape is reminder', async function () {
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
            const task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, request, shape);
            expect(task?.ReminderMinutesBeforeStart).to.be.equal('' + (pimTask.alarm ? (pimTask.alarm < 0 ? `${Math.abs(pimTask.alarm)}` : "0") : undefined));
            expect(pimTask.due).to.not.be.undefined();
            expect(task?.ReminderDueBy ? getTrimmedISODate(task.ReminderDueBy) : undefined).to.be.equal(getTrimmedISODate(pimTask.due!));
        });

        it('Passing in a folderId', async function () {
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
            const task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, request, shape, undefined, 'PARENT-FOLDER');
            const parentFolderEWSId = getEWSId('PARENT-FOLDER');
            expect(task?.ParentFolderId?.Id).to.be.equal(parentFolderEWSId);
            expect(task?.ParentFolderId?.ChangeKey).to.be.equal(`ck-${parentFolderEWSId}`);
        });

        describe('shape is priority', function () {

            it('importance is high', async function () {
                const shape = new ItemResponseShapeType();
                shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
                const task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, request, shape);
                expect(task?.Importance).to.be.equal(ImportanceChoicesType.HIGH);
            });

            it('importance is low', async function () {
                pimTask.priority = PimTaskPriority.LOW;
                const shape = new ItemResponseShapeType();
                shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
                const task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, request, shape);
                expect(task?.Importance).to.be.equal(ImportanceChoicesType.LOW);
            });

            it('importance is Medium', async function () {
                pimTask.priority = PimTaskPriority.MEDIUM;
                const shape = new ItemResponseShapeType();
                shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
                const task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, request, shape);
                expect(task?.Importance).to.be.equal(ImportanceChoicesType.NORMAL);
            });
            it('importance is None', async function () {
                pimTask.priority = PimTaskPriority.NONE;
                const shape = new ItemResponseShapeType();
                shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
                const task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, request, shape);
                expect(task?.Importance).to.be.equal(ImportanceChoicesType.NORMAL);
            });
        });

        it('shape is isComplete', async function () {
            pimTask.isComplete = true;
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
            const task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, request, shape);
            expect(task?.Status).to.be.equal(TaskStatusType.COMPLETED);
        });

        it('shape is isInProgress', async function () {
            pimTask.isInProgress = true;
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
            const task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, context.request, shape);
            expect(task?.Status).to.be.equal(TaskStatusType.IN_PROGRESS);
        });

        it('shape is isNotStarted', async function () {
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
            const task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, request, shape);
            expect(task?.Status).to.be.equal(TaskStatusType.NOT_STARTED);
            expect(task?.PercentComplete).to.be.equal(0);
        });

        it('shape is statusLabel', async function () {
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
            const task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, request, shape);
            expect(task?.StatusDescription).to.be.equal(pimTask.statusLabel);
        });

        it('shape is isConfidential confidential', async function () {
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
            const task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, request, shape);
            expect(task?.Sensitivity).to.be.equal(SensitivityChoicesType.CONFIDENTIAL); //true
        });
        it('shape is isConfidential Normal', async function () {
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.DEFAULT;
            let task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, request, shape);
            expect(task?.Sensitivity).to.be.undefined();

            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
            task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, request, shape);
            expect(task?.Sensitivity).to.be.equal(SensitivityChoicesType.CONFIDENTIAL);

        });

        it('shape is attachments', async function () {
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
            const task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, request, shape);
            expect(task?.HasAttachments).to.be.equal(true); //true
        });

        it('shape is effectiveRights', async function () {
            const shape = new ItemResponseShapeType();
            shape.BaseShape = DefaultShapeNamesType.ALL_PROPERTIES;
            const task = await EWSTaskManager.getInstance().pimItemToEWSItem(pimTask, userInfo, request, shape);
            expect(task?.EffectiveRights?.CreateAssociated).to.be.equal(true);
            expect(task?.EffectiveRights?.CreateContents).to.be.equal(true);
            expect(task?.EffectiveRights?.CreateHierarchy).to.be.equal(true);
            expect(task?.EffectiveRights?.CreateHierarchy).to.be.equal(true);
            expect(task?.EffectiveRights?.Delete).to.be.equal(true);
            expect(task?.EffectiveRights?.Modify).to.be.equal(true);
            expect(task?.EffectiveRights?.Read).to.be.equal(true);
        });
    });

    describe('Test getInstance', () => {
        it('getInstance', function () {
            const manager = EWSTaskManager.getInstance();
            expect(manager).to.be.instanceof(EWSTaskManager);

            const manager2 = EWSTaskManager.getInstance();
            expect(manager).to.be.equal(manager2);
        });
    });

    describe('Test fieldsForDefaultShape', () => {
        it('test default fields', () => {
            const manager = EWSTaskManager.getInstance();
            const fields = manager.fieldsForDefaultShape();
            expect(fields).to.be.an.Array();
            expect(fields.length).to.be.equal(9);
            expect(fields.length).to.be.equal(9);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_ITEM_ID);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_PARENT_FOLDER_ID);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_ITEM_CLASS);

            expect(fields).to.containEql(UnindexedFieldURIType.TASK_DUE_DATE);
            expect(fields).to.containEql(UnindexedFieldURIType.TASK_PERCENT_COMPLETE);
            expect(fields).to.containEql(UnindexedFieldURIType.TASK_START_DATE);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_SUBJECT);
            expect(fields).to.containEql(UnindexedFieldURIType.TASK_STATUS);
        });

    });

    describe('Test fieldsForAllPropertiesShape', () => {
        it('test all properties fields', () => {
            const manager = EWSTaskManager.getInstance();
            const fields = manager.fieldsForAllPropertiesShape();
            expect(fields).to.be.an.Array();
            // Verify the size is correct
            expect(fields.length).to.be.equal(54);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_ITEM_ID);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_PARENT_FOLDER_ID);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_ITEM_CLASS);

            expect(fields).to.containEql(UnindexedFieldURIType.TASK_DUE_DATE);
            expect(fields).to.containEql(UnindexedFieldURIType.TASK_PERCENT_COMPLETE);
            expect(fields).to.containEql(UnindexedFieldURIType.TASK_START_DATE);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_SUBJECT);
            expect(fields).to.containEql(UnindexedFieldURIType.TASK_STATUS);

            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_BODY);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_CATEGORIES);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_LAST_MODIFIED_TIME);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_LAST_MODIFIED_NAME);
            expect(fields).to.containEql(UnindexedFieldURIType.TASK_ACTUAL_WORK);
            expect(fields).to.containEql(UnindexedFieldURIType.TASK_ASSIGNED_TIME);
            expect(fields).to.containEql(UnindexedFieldURIType.TASK_BILLING_INFORMATION);
            expect(fields).to.containEql(UnindexedFieldURIType.TASK_CHANGE_COUNT);
            expect(fields).to.containEql(UnindexedFieldURIType.TASK_COMPANIES);
            expect(fields).to.containEql(UnindexedFieldURIType.TASK_COMPLETE_DATE);
            expect(fields).to.containEql(UnindexedFieldURIType.TASK_CONTACTS);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_CONVERSATION_ID);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_CULTURE);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_DATE_TIME_CREATED);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_DATE_TIME_RECEIVED);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_DATE_TIME_SENT);
            expect(fields).to.containEql(UnindexedFieldURIType.TASK_DELEGATION_STATE);
            expect(fields).to.containEql(UnindexedFieldURIType.TASK_DELEGATOR);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_DISPLAY_CC);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_DISPLAY_TO);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_EFFECTIVE_RIGHTS);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_IMPORTANCE);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_IN_REPLY_TO);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_INTERNET_MESSAGE_HEADERS);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_IS_ASSOCIATED);
            expect(fields).to.containEql(UnindexedFieldURIType.TASK_IS_COMPLETE);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_IS_DRAFT);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_IS_FROM_ME);
            expect(fields).to.containEql(UnindexedFieldURIType.TASK_IS_RECURRING);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_IS_RESEND);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_IS_SUBMITTED);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_IS_UNMODIFIED);
            expect(fields).to.containEql(UnindexedFieldURIType.TASK_MILEAGE);
            expect(fields).to.containEql(UnindexedFieldURIType.TASK_OWNER);
            expect(fields).to.containEql(UnindexedFieldURIType.TASK_RECURRENCE);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_REMINDER_DUE_BY);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_REMINDER_IS_SET);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_REMINDER_MINUTES_BEFORE_START);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_RESPONSE_OBJECTS);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_SENSITIVITY);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_SIZE);
            expect(fields).to.containEql(UnindexedFieldURIType.TASK_STATUS_DESCRIPTION);
            expect(fields).to.containEql(UnindexedFieldURIType.TASK_TOTAL_WORK);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_WEB_CLIENT_EDIT_FORM_QUERY_STRING);
            expect(fields).to.containEql(UnindexedFieldURIType.ITEM_WEB_CLIENT_READ_FORM_QUERY_STRING);
        });
    });

    describe('Test recurrence', () => {
        const context = getContext("testUser", "testPass");
        const userInfo = new UserContext();

        const shape = new ItemResponseShapeType();
        shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
        const additions = [UnindexedFieldURIType.TASK_RECURRENCE, UnindexedFieldURIType.TASK_IS_RECURRING];
        shape.AdditionalProperties = new NonEmptyArrayOfPathsToElementType(additions.map(type => {
            const fieldType = new PathToUnindexedFieldType();
            fieldType.FieldURI = type;
            return fieldType;
        }));

        describe('Yearly recurrences', () => {
            it('Yearly on day of month', async () => {
                // Start of first event
                const startString = '2021-05-11T09:00:00Z';
                const dueString = '2021-05-12T09:00:00Z';

                // Yearly on May 11th
                const jsTask = {
                    "@type": "jstask",
                    "uid": "unit-test-id",
                    "title": "Unit Test",
                    "description": "This is an item created by unit tests.",
                    "start": startString,
                    "due": dueString,
                    "priority": 2,
                    "progress": "needs-action",
                    "privacy": "public",
                    recurrenceRules: [{
                        "@type": "RecurrenceRule",
                        "frequency": "yearly",
                        "byMonthDay": [11],
                        "byMonth": ['5'],
                        "count": 2
                    }]
                };

                const pimItem = PimItemFactory.newPimTask(jsTask);
                const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                expect(manager).to.not.be.undefined();

                const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape) as TaskType;
                expect(ewsItem.Recurrence).to.not.be.undefined();
                expect(ewsItem.ItemId?.Id).to.not.be.undefined();
                expect(ewsItem.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));
                expect(ewsItem.IsRecurring).to.be.true();

                // Validate recurrence rule
                expect(ewsItem.Recurrence?.AbsoluteYearlyRecurrence?.DayOfMonth).to.be.equal(11);
                expect(ewsItem.Recurrence?.AbsoluteYearlyRecurrence?.Month).to.be.equal(MonthNamesType.MAY);

            });

            it('Yearly on nth day of week in a month', async () => {
                const dueString = '2021-04-08T09:00:00Z';
                const endString = '2022-04-14T09:00:00Z';

                const jsTasks = [
                    {
                        "@type": "jstask",
                        "uid": "unit-test-id",
                        "title": "Unit Test",
                        "description": "This is an item created by unit tests.",
                        "due": dueString,
                        "priority": 2,
                        "progress": "needs-action",
                        "privacy": "public",
                        recurrenceRules: [{
                            "@type": "RecurrenceRule",
                            "frequency": "yearly",
                            "byMonth": ['4'],
                            "byDay": [{ "@type": "NDay", day: "th", nthOfPeriod: 2 }],
                            "until": endString
                        }]
                    },
                    {
                        "@type": "jstask",
                        "uid": "unit-test-id",
                        "title": "Unit Test",
                        "description": "This is an item created by unit tests.",
                        "due": dueString,
                        "priority": 2,
                        "progress": "needs-action",
                        "privacy": "public",
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

                for (const jsTask of jsTasks) {
                    const pimItem = PimItemFactory.newPimTask(jsTask);
                    const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                    expect(manager).to.not.be.undefined();

                    const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape) as TaskType;
                    expect(ewsItem.Recurrence).to.not.be.undefined();
                    expect(ewsItem.ItemId?.Id).to.not.be.undefined();
                    expect(ewsItem.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));
                    expect(ewsItem.IsRecurring).to.be.true();

                    // Validate recurrence rule
                    expect(ewsItem.Recurrence?.RelativeYearlyRecurrence?.DayOfWeekIndex).to.be.equal(DayOfWeekIndexType.SECOND);
                    expect(ewsItem.Recurrence?.RelativeYearlyRecurrence?.DaysOfWeek).to.be.equal(DayOfWeekType.THURSDAY);
                    expect(ewsItem.Recurrence?.RelativeYearlyRecurrence?.Month).to.be.equal(MonthNamesType.APRIL);
                }
            });
        });

        describe('Monthly recurrence', () => {

            it('Monthly, on day of month', async () => {
                // Start of first event
                const startString = '2021-05-11T09:00:00Z';
                const dueString = '2021-05-12T09:00:00Z';

                // Every 2 Months on the 11th of the month
                const jsTask = {
                    "@type": "jstask",
                    "uid": "unit-test-id",
                    "title": "Unit Test",
                    "description": "This is an item created by unit tests.",
                    "start": startString,
                    "due": dueString,
                    "priority": 2,
                    "progress": "needs-action",
                    "privacy": "public",
                    recurrenceRules: [{
                        "@type": "RecurrenceRule",
                        "frequency": "monthly",
                        "byMonthDay": [11],
                        "interval": 2,
                        "count": 3
                    }]
                };

                const pimItem = PimItemFactory.newPimTask(jsTask);
                const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                expect(manager).to.not.be.undefined();

                const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape) as TaskType;
                expect(ewsItem.Recurrence).to.not.be.undefined();
                expect(ewsItem.ItemId?.Id).to.not.be.undefined();
                expect(ewsItem.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));
                expect(ewsItem.IsRecurring).to.be.true();

                // Validate recurrence rule
                expect(ewsItem.Recurrence?.AbsoluteMonthlyRecurrence?.DayOfMonth).to.be.equal(11);
                expect(ewsItem.Recurrence?.AbsoluteMonthlyRecurrence?.Interval).to.be.equal(2);

            });

            it('Monthly, on nth day of week', async () => {
                // Start of first event
                const dueString = '2021-05-04T09:00:00Z';

                // Every Month on the 1st Tue of the month
                const jsTasks = [
                    {
                        "@type": "jstask",
                        "uid": "unit-test-id",
                        "title": "Unit Test",
                        "description": "This is an item created by unit tests.",
                        "due": dueString,
                        "priority": 2,
                        "progress": "needs-action",
                        "privacy": "public",
                        recurrenceRules: [{
                            "@type": "RecurrenceRule",
                            "frequency": "monthly",
                            "byDay": [{ "@type": "NDay", day: "tu", nthOfPeriod: 1 }],
                            "count": 3
                        }]
                    },
                    {
                        "@type": "jstask",
                        "uid": "unit-test-id",
                        "title": "Unit Test",
                        "description": "This is an item created by unit tests.",
                        "due": dueString,
                        "priority": 2,
                        "progress": "needs-action",
                        "privacy": "public",
                        recurrenceRules: [{
                            "@type": "RecurrenceRule",
                            "frequency": "monthly",
                            "byDay": [{ "@type": "NDay", day: "tu" }],
                            "bySetPosition": [1],
                            "count": 3
                        }]
                    }
                ];

                for (const jsTask of jsTasks) {
                    const pimItem = PimItemFactory.newPimTask(jsTask);
                    const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                    expect(manager).to.not.be.undefined();

                    const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape) as TaskType;
                    expect(ewsItem.Recurrence).to.not.be.undefined();
                    expect(ewsItem.ItemId?.Id).to.not.be.undefined();
                    expect(ewsItem.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));
                    expect(ewsItem.IsRecurring).to.be.true();

                    // Validate recurrence rule
                    expect(ewsItem.Recurrence?.RelativeMonthlyRecurrence?.DayOfWeekIndex).to.be.equal(DayOfWeekIndexType.FIRST);
                    expect(ewsItem.Recurrence?.RelativeMonthlyRecurrence?.DaysOfWeek).to.be.equal(DayOfWeekType.TUESDAY);
                    expect(ewsItem.Recurrence?.RelativeMonthlyRecurrence?.Interval).to.be.equal(1); // Default is 1
                }
            });
        });

        describe('Weekly recurrence', () => {

            it('Weekly on day of the week', async () => {
                // Start of first event
                const startString = '2021-05-04T09:00:00Z';
                const dueString = '2021-05-05T09:00:00Z';

                // Every work day for 9 occurences
                const jsTask = {
                    "@type": "jstask",
                    "uid": "unit-test-id",
                    "title": "Unit Test",
                    "description": "This is an item created by unit tests.",
                    "start": startString,
                    "due": dueString,
                    "priority": 2,
                    "progress": "needs-action",
                    "privacy": "public",
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

                const pimItem = PimItemFactory.newPimTask(jsTask);
                const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                expect(manager).to.not.be.undefined();

                const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape) as TaskType;
                expect(ewsItem.Recurrence).to.not.be.undefined();
                expect(ewsItem.ItemId?.Id).to.not.be.undefined();
                expect(ewsItem.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));
                expect(ewsItem.IsRecurring).to.be.true();

                // Validate recurrence rule
                expect(ewsItem.Recurrence?.WeeklyRecurrence?.FirstDayOfWeek).to.be.equal(DayOfWeekType.MONDAY);
                expect(ewsItem.Recurrence?.WeeklyRecurrence?.DaysOfWeek.$value).to.be.deepEqual(`${DayOfWeekType.MONDAY} ${DayOfWeekType.TUESDAY} ${DayOfWeekType.WEDNESDAY} ${DayOfWeekType.THURSDAY} ${DayOfWeekType.FRIDAY}`);
                expect(ewsItem.Recurrence?.WeeklyRecurrence?.Interval).to.be.equal(1); // Default is 1
            });

            it('Every nth Week', async () => {
                // Start of first event
                const dueString = '2021-05-04T09:00:00Z';

                // Every 2 weeks
                const jsTask = {
                    "@type": "jstask",
                    "uid": "unit-test-id",
                    "title": "Unit Test",
                    "description": "This is an item created by unit tests.",
                    "due": dueString,
                    "priority": 2,
                    "progress": "needs-action",
                    "privacy": "public",
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

                const pimItem = PimItemFactory.newPimTask(jsTask);
                const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                expect(manager).to.not.be.undefined();

                const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape) as TaskType;
                expect(ewsItem.Recurrence).to.not.be.undefined();
                expect(ewsItem.ItemId?.Id).to.not.be.undefined();
                expect(ewsItem.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));
                expect(ewsItem.IsRecurring).to.be.true();

                // Validate recurrence rule
                expect(ewsItem.Recurrence?.WeeklyRecurrence?.FirstDayOfWeek).to.be.equal(DayOfWeekType.MONDAY);
                expect(ewsItem.Recurrence?.WeeklyRecurrence?.DaysOfWeek.$value).to.be.equal(DayOfWeekType.TUESDAY);
                expect(ewsItem.Recurrence?.WeeklyRecurrence?.Interval).to.be.equal(2);

            });
        });

        describe('Daily recurrence', () => {

            it('Every other day', async () => {
                // Start of first event
                const startString = '2021-05-04T09:00:00Z';
                const dueString = '2021-05-05T09:00:00Z';

                // Every other day
                const jsTask = {
                    "@type": "jstask",
                    "uid": "unit-test-id",
                    "title": "Unit Test",
                    "description": "This is an item created by unit tests.",
                    "start": startString,
                    "due": dueString,
                    "priority": 2,
                    "progress": "needs-action",
                    "privacy": "public",
                    recurrenceRules: [{
                        "@type": "RecurrenceRule",
                        "frequency": "daily",
                        "interval": 2,
                        "count": 10
                    }]
                };

                const pimItem = PimItemFactory.newPimTask(jsTask);
                const manager = EWSServiceManager.getInstanceFromPimItem(pimItem);
                expect(manager).to.not.be.undefined();

                const ewsItem = await manager!.pimItemToEWSItem(pimItem, userInfo, context.request, shape) as TaskType;
                expect(ewsItem.Recurrence).to.not.be.undefined();
                expect(ewsItem.ItemId?.Id).to.not.be.undefined();
                expect(ewsItem.ItemId?.Id).to.be.equal(getEWSId(pimItem.unid));
                expect(ewsItem.IsRecurring).to.be.true();

                // Validate recurrence rule
                expect(ewsItem.Recurrence?.DailyRecurrence?.Interval).to.be.equal(2);

            });
        });
    });

    describe('Test createItem', () => {
        let taskManagerStub: sinon.SinonStubbedInstance<KeepPimTasksManager>;
        let messageManagerStub: sinon.SinonStubbedInstance<KeepPimMessageManager>;

        const userInfo = new UserContext();
        const testUser = 'test.user@test.org';
        const context = getContext(testUser, 'password');
        const mailboxId = 'test@test.com';
        const unid = 'UID';
        const itemEWSId = getEWSId(unid, mailboxId);

        let movemessageResults: KeepMoveMessagesResults;

        beforeEach(() => {
            taskManagerStub = sinon.createStubInstance(KeepPimTasksManager);
            taskManagerStub.createTask.resolves(unid);
            taskManagerStub.deleteTask.resolves();
            stubTasksManager(taskManagerStub);
            messageManagerStub = sinon.createStubInstance(KeepPimMessageManager);
            stubMessageManager(messageManagerStub);
        });

        afterEach(() => {
            sinon.restore(); 
        });

        it('creates task, no label', async () => {
            const task = new TaskType();
            
            const receivedTasks = await EWSTaskManager
                .getInstance()
                .createItem(task, userInfo, context.request, undefined, undefined, mailboxId);

            expect(receivedTasks).to.not.be.undefined();
            expect(receivedTasks.length).to.be.equal(1);
            expect(receivedTasks[0]?.ItemId?.Id).to.be.equal(itemEWSId);
        });

        it('creates task with a label', async () => {
            const toLabel = PimItemFactory.newPimLabel({}, PimItemFormat.DOCUMENT);
            toLabel.view = KeepPimConstants.INBOX;
            toLabel.unid = 'to-label-uid';
            movemessageResults = {
                movedIds: [{
                    status: 200,
                    message: 'ok',
                    unid,
                }]
            };
            messageManagerStub.moveMessages.resolves(movemessageResults);
            const task = new TaskType();

            const receivedTasks = await EWSTaskManager
                .getInstance()
                .createItem(task, userInfo, context.request, undefined, toLabel, mailboxId);

            expect(receivedTasks).to.not.be.undefined();
            expect(receivedTasks.length).to.be.equal(1);
            expect(receivedTasks[0]?.ItemId?.Id).to.be.equal(itemEWSId);
        });

        it('throws an error after move failed', async () => {
            const toLabel = PimItemFactory.newPimLabel({}, PimItemFormat.DOCUMENT);
            toLabel.view = KeepPimConstants.INBOX;
            toLabel.unid = 'to-label-uid';
            movemessageResults = {
                movedIds: [{
                    status: 404,
                    message: 'fail',
                }]
            };
            messageManagerStub.moveMessages.resolves(movemessageResults);
            const task = new TaskType();

            const expectedError = `Move of item ${unid} failed`;

            await expect(EWSTaskManager
                .getInstance()
                .createItem(task, userInfo, context.request, undefined, toLabel)
            ).to.be.rejectedWith(expectedError);
        });
    });
});
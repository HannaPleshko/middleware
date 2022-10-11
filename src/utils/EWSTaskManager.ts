import {
    PimTask, PimLabel, KeepPimTasksManager, PimItemFactory, KeepPimConstants, UserInfo, PimTaskPriority, 
    KeepPimMessageManager
} from '@hcllabs/openclientkeepcomponent';
import { EWSServiceManager } from '.';
import { UserContext } from '../keepcomponent';
import { Logger } from './logger';
import {
    DefaultShapeNamesType, UnindexedFieldURIType, ImportanceChoicesType, MessageDispositionType, SensitivityChoicesType,
    BodyTypeType, TaskStatusType, ExtendedPropertyKeyType, DistinguishedPropertySetType, MapiPropertyIds
} from '../models/enum.model';
import {
    ItemResponseShapeType, TaskType, ItemIdType, FolderIdType, ItemChangeType, ItemType, SetItemFieldType, 
    AppendToItemFieldType, DeleteItemFieldType
} from '../models/mail.model';
import { getEWSId } from './pimHelper';
import {
    addExtendedPropertiesToPIM, getValueForExtendedFieldURI, identifiersForPathToExtendedFieldType, addExtendedPropertyToPIM
} from '../utils';
import * as util from 'util';
import { fromString } from 'html-to-text';
import { Request } from '@loopback/rest';

// EWSServiceManager subclass implemented for Task related operations
export class EWSTaskManager extends EWSServiceManager {

    // Singleton instance
    private static instance: EWSTaskManager;

    public static getInstance(): EWSTaskManager {
        if (!EWSTaskManager.instance) {
            this.instance = new EWSTaskManager();
        }
        return this.instance;
    }

    // Fields included for DEFAULT shape for tasks
    private static defaultFields = [
        ...EWSServiceManager.idOnlyFields,
        UnindexedFieldURIType.TASK_DUE_DATE,
        UnindexedFieldURIType.TASK_PERCENT_COMPLETE,
        UnindexedFieldURIType.TASK_START_DATE,
        UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS,
        UnindexedFieldURIType.ITEM_SUBJECT,
        UnindexedFieldURIType.TASK_STATUS
    ];

    // Fields included for ALL_PROPERTIES shape for tasks
    private static allPropertiesFields = [
        ...EWSTaskManager.defaultFields,
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
        UnindexedFieldURIType.TASK_CONTACTS,
        UnindexedFieldURIType.ITEM_CONVERSATION_ID,
        UnindexedFieldURIType.ITEM_CULTURE,
        UnindexedFieldURIType.ITEM_DATE_TIME_CREATED,
        UnindexedFieldURIType.ITEM_DATE_TIME_RECEIVED,
        UnindexedFieldURIType.ITEM_DATE_TIME_SENT,
        UnindexedFieldURIType.TASK_DELEGATION_STATE,
        UnindexedFieldURIType.TASK_DELEGATOR,
        UnindexedFieldURIType.ITEM_DISPLAY_CC,
        UnindexedFieldURIType.ITEM_DISPLAY_TO,
        UnindexedFieldURIType.ITEM_EFFECTIVE_RIGHTS,
        UnindexedFieldURIType.ITEM_IMPORTANCE,
        UnindexedFieldURIType.ITEM_IN_REPLY_TO,
        UnindexedFieldURIType.ITEM_INTERNET_MESSAGE_HEADERS,
        UnindexedFieldURIType.ITEM_IS_ASSOCIATED,
        UnindexedFieldURIType.TASK_IS_COMPLETE,
        UnindexedFieldURIType.ITEM_IS_DRAFT,
        UnindexedFieldURIType.ITEM_IS_FROM_ME,
        UnindexedFieldURIType.TASK_IS_RECURRING,
        UnindexedFieldURIType.ITEM_IS_RESEND,
        UnindexedFieldURIType.ITEM_IS_SUBMITTED,
        UnindexedFieldURIType.ITEM_IS_UNMODIFIED,
        UnindexedFieldURIType.TASK_MILEAGE,
        UnindexedFieldURIType.TASK_OWNER,
        UnindexedFieldURIType.TASK_RECURRENCE,
        UnindexedFieldURIType.ITEM_REMINDER_DUE_BY,
        UnindexedFieldURIType.ITEM_REMINDER_IS_SET,
        UnindexedFieldURIType.ITEM_REMINDER_MINUTES_BEFORE_START,
        UnindexedFieldURIType.ITEM_RESPONSE_OBJECTS,
        UnindexedFieldURIType.ITEM_SENSITIVITY,
        UnindexedFieldURIType.ITEM_SIZE,
        UnindexedFieldURIType.TASK_STATUS_DESCRIPTION,
        UnindexedFieldURIType.TASK_TOTAL_WORK,
        UnindexedFieldURIType.ITEM_WEB_CLIENT_EDIT_FORM_QUERY_STRING,
        UnindexedFieldURIType.ITEM_WEB_CLIENT_READ_FORM_QUERY_STRING
    ];

    async getItems(
        userInfo: UserInfo, 
        request: Request, 
        shape: ItemResponseShapeType, 
        startIndex?: number, 
        count?: number, 
        fromLabel?: PimLabel,
        mailboxId?: string
    ): Promise<TaskType[]> {
        const taskItems: TaskType[] = [];
        let pimItems: PimTask[] | undefined = undefined;
        try {
            const needDocuments = (
                shape.BaseShape === DefaultShapeNamesType.ID_ONLY 
                && (shape.AdditionalProperties === undefined || shape.AdditionalProperties.items.length === 0)
                ) ? false : true;

            pimItems = await KeepPimTasksManager
                .getInstance()
                .getTasks(userInfo, needDocuments, startIndex, count, fromLabel?.unid, mailboxId);
        } catch (err) {
            Logger.getInstance().debug("Error retrieving PIM task entries: " + err);
            // If we throw the err here the client will continue in a loop with SyncFolderHierarchy and SyncFolderItems asking for messages.
        }
        if (pimItems) {
            for (const pimItem of pimItems) {
                const taskItem = await this.pimItemToEWSItem(pimItem, userInfo, request, shape, mailboxId, fromLabel?.unid);
                if (taskItem) {
                    taskItems.push(taskItem);
                } else {
                    Logger.getInstance().error(`An EWS task type could not be created for PIM task ${pimItem.unid}`);
                }
            }
        }
        return taskItems;
    }

    async updateItem(
        pimTask: PimTask, 
        change: ItemChangeType, 
        userInfo: UserInfo, 
        request: Request, 
        toLabel?: PimLabel,
        mailboxid?: string
    ): Promise<TaskType | undefined> {

        for (const fieldUpdate of change.Updates.items) {
            let newItem: ItemType | undefined = undefined;
            if (fieldUpdate instanceof SetItemFieldType || fieldUpdate instanceof AppendToItemFieldType) {
                // Should only have to worry about Item and Task field updates for tasks
                newItem = fieldUpdate.Item ?? fieldUpdate.Task;
            }

            if (fieldUpdate instanceof SetItemFieldType) {
                if (newItem === undefined) {
                    Logger.getInstance().error(`No new item set for update field: ${util.inspect(newItem, false, 5)}`);
                    continue;
                }

                if (fieldUpdate.FieldURI) {
                    const field = fieldUpdate.FieldURI.FieldURI.split(':')[1];
                    this.updatePimItemFieldValue(pimTask, fieldUpdate.FieldURI.FieldURI, (newItem as any)[field]);
                } else if (fieldUpdate.ExtendedFieldURI && newItem.ExtendedProperty) {
                    const identifiers = identifiersForPathToExtendedFieldType(fieldUpdate.ExtendedFieldURI);
                    const newValue = getValueForExtendedFieldURI(newItem, identifiers);

                    const current = pimTask.findExtendedProperty(identifiers);
                    if (current) {
                        current.Value = newValue;
                        pimTask.updateExtendedProperty(identifiers, current);
                    }
                    else {
                        addExtendedPropertyToPIM(pimTask, newItem.ExtendedProperty[0]);
                    }
                } else if (fieldUpdate.IndexedFieldURI) {
                    Logger.getInstance().warn(`Unhandled SetItemField request for notes for field:  ${fieldUpdate.IndexedFieldURI}`);
                }
            }
            else if (fieldUpdate instanceof AppendToItemFieldType) {
                /*
                    The following properties are supported for the append action for notes:
                    - item:Body
                */
                if (fieldUpdate.FieldURI?.FieldURI === UnindexedFieldURIType.ITEM_BODY && newItem) {
                    pimTask.body = `${pimTask.body}${newItem?.Body ?? ""}`
                } else {
                    Logger.getInstance().warn(`Unhandled AppendToItemField request for notes for field:  ${fieldUpdate.FieldURI?.FieldURI}`);
                }
            }
            else if (fieldUpdate instanceof DeleteItemFieldType) {
                if (fieldUpdate.FieldURI) {
                    this.updatePimItemFieldValue(pimTask, fieldUpdate.FieldURI.FieldURI, undefined);
                }
                else if (fieldUpdate.ExtendedFieldURI) {
                    pimTask.deleteExtendedProperty(identifiersForPathToExtendedFieldType(fieldUpdate.ExtendedFieldURI));
                }
                else if (fieldUpdate.IndexedFieldURI) {
                    // Only handling these for contacts currently
                    Logger.getInstance().warn(`Unhandled DeleteItemField request for notes for field:  ${fieldUpdate.IndexedFieldURI}`);
                }
            }
        }

        // Set the target folder if passed in
        if (toLabel) {
            pimTask.parentFolderIds = [toLabel.folderId];  // May be possible we've lost other parent ids set by another client.  
            // Even if stored as an extra property for the client, it could have been updated on the server since stored on the client.
            // To shrink the window, we'd need to request for the item from the server and update it right away
        }

        // The pimTask should now be updated with the new information.  Send it to Keep.
        await KeepPimTasksManager.getInstance().updateTask(pimTask, userInfo, mailboxid);

        const shape = new ItemResponseShapeType();
        shape.BaseShape = DefaultShapeNamesType.ID_ONLY;
        return this.pimItemToEWSItem(pimTask, userInfo, request, shape, mailboxid);
    }

    async createItem(
        task: TaskType, 
        userInfo: UserInfo, 
        request: Request, 
        disposition?: MessageDispositionType,
        toLabel?: PimLabel,
        mailboxId?: string
    ): Promise<TaskType[]> {

        const pimTask = this.pimItemFromEWSItem(task, request);

        // Make the Keep API call to create the task
        const unid = await KeepPimTasksManager.getInstance().createTask(pimTask, userInfo, mailboxId);
        const targetItem = new TaskType();
        const itemEWSId = getEWSId(unid, mailboxId);
        targetItem.ItemId = new ItemIdType(itemEWSId, `ck-${itemEWSId}`);

        // TODO:  Uncomment this line when we can set the parent folder to the base tasks folder
        //        See LABS-866
        // If no toLabel is passed in, we will create the contact in the top level contacts folder
        // if (toLabel === undefined) {
        //     const labels = await KeepPimLabelManager.getInstance().getLabels(userInfo);
        //     toLabel = labels.find(label => { return label.view === KeepPimConstants.TASKS });
        // }

        // The Keep API does not support creating a task IN a folder, so we must now move the 
        // task to the desired folder
        // TODO:  If this contact is being created in the top level tasks folder, do not attempt to 
        //        move the task to the top level folder as this will fail.
        //        See LABS-866
        if (toLabel !== undefined && toLabel.view !== KeepPimConstants.TASKS) {
            try {
                //Move the item in keep.  
                const moveItems: string[] = [];
                moveItems.push(unid);

                const results = await KeepPimMessageManager
                    .getInstance()
                    .moveMessages(userInfo, toLabel.folderId, moveItems, undefined, undefined, undefined, mailboxId);

                if (results.movedIds && results.movedIds.length > 0) {
                    const result = results.movedIds[0];
                    if (result.unid !== unid || result.status !== 200) {
                        Logger.getInstance().error(`move for task ${unid} to folder ${toLabel.folderId} failed: ${util.inspect(results, false, 5)}`);
                        throw new Error(`Move of item ${unid} failed`);
                    }
                }
            } catch (err) {
                // If there was an error, delete the note on the server since it didn't make it to the correct location
                Logger.getInstance().debug(`error moving item, ${unid},  to desired folder ${toLabel.unid}: ${err}`);
                await this.deleteItem(unid, userInfo);
                throw err;
            }
        }
        return [targetItem];
    }

    /**
     * Creates a new pim item. 
     * @param pimItem The task item to create.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param mailboxId SMTP mailbox delegator or delegatee address.
     * @returns The new Pim task item created with Keep.
     * @throws An error if the create fails
     */
    protected async createNewPimItem(pimItem: PimTask, userInfo: UserInfo, mailboxId?: string): Promise<PimTask> {
        const unid = await KeepPimTasksManager.getInstance().createTask(pimItem, userInfo, mailboxId);
        const newItem = await KeepPimTasksManager.getInstance().getTask(unid, userInfo, mailboxId);
        if (newItem === undefined) {
            // Try to delete item since it may have been created
            try {
                await KeepPimTasksManager.getInstance().deleteTask(unid, userInfo, undefined, mailboxId);
            }
            catch {
                // Ignore errors
            }
            throw new Error(`Unable to retrieve task ${unid} after create`);
        }

        return newItem;
    }

    /**
     * This function will make a Keep API request to delete a task.
     * @param item The pim task or unid of the task to delete.
     * @param userInfo The user's credentials to be passed to Keep.
     * @param hardDelete Indicator if the task should be deleted from trash as well
     */
    async deleteItem(item: string | PimTask, userInfo: UserInfo, mailboxId?: string, hardDelete = false): Promise<void> {
        const unid = (typeof item === 'string') ? item : item.unid;
        await KeepPimTasksManager.getInstance().deleteTask(unid, userInfo, hardDelete, mailboxId);
    }

    /**
      * Getter for the list of fields included in the default shape for tasks
      * @returns Array of fields for the default shape
      */
    fieldsForDefaultShape(): UnindexedFieldURIType[] {
        return EWSTaskManager.defaultFields;
    }

    /**
     * Getter for the list of fields included in the all properties shape for tasks
     * @returns Array of fields for the all properties shape
     */
    fieldsForAllPropertiesShape(): UnindexedFieldURIType[] {
        return EWSTaskManager.allPropertiesFields;
    }

    async pimItemToEWSItem(
        pimTask: PimTask, 
        userInfo: UserInfo, 
        request: Request, 
        shape: ItemResponseShapeType, 
        mailboxId?: string,
        targetParentFolderId?: string
    ): Promise<TaskType> {
        const task = new TaskType();

        // Add all requested properties based on the shape
        this.addRequestedPropertiesToEWSItem(pimTask, task, shape, mailboxId);

        if (targetParentFolderId) {
            const parentFolderEWSId = getEWSId(targetParentFolderId, mailboxId);
            task.ParentFolderId = new FolderIdType(parentFolderEWSId, `ck-${parentFolderEWSId}`);
        }
        return task;
    }

    updateEWSItemFieldValue(task: TaskType, pimTask: PimTask, fieldIdentifier: UnindexedFieldURIType, mailboxId?: string): boolean {

        let handled = true;

        switch (fieldIdentifier) {
            case UnindexedFieldURIType.TASK_DUE_DATE:
                if (pimTask.due) {
                    task.DueDate = new Date(pimTask.due);
                }
                break;
            case UnindexedFieldURIType.TASK_PERCENT_COMPLETE:
                if (pimTask.isComplete) {
                    task.PercentComplete = 100;
                } else {
                    // Keep doesn't tell us a percentage, so set to 0 if not complete.
                    // Should we store this in AdditionalProperties if the client sets it?
                    task.PercentComplete = 0;
                }
                break;
            case UnindexedFieldURIType.TASK_START_DATE:
                if (pimTask.start) {
                    task.StartDate = new Date(pimTask.start);
                }
                break;
            case UnindexedFieldURIType.TASK_STATUS:
                if (pimTask.isComplete) {
                    task.Status = TaskStatusType.COMPLETED;
                }
                else if (pimTask.isInProgress) {
                    task.Status = TaskStatusType.IN_PROGRESS;
                }
                else if (pimTask.isNotStarted) {
                    task.Status = TaskStatusType.NOT_STARTED;
                }
                break;
            case UnindexedFieldURIType.TASK_ACTUAL_WORK:
            case UnindexedFieldURIType.TASK_ASSIGNED_TIME:
            case UnindexedFieldURIType.TASK_BILLING_INFORMATION:
            case UnindexedFieldURIType.TASK_CHANGE_COUNT:
            case UnindexedFieldURIType.TASK_COMPANIES:
            case UnindexedFieldURIType.TASK_CONTACTS:
            case UnindexedFieldURIType.TASK_DELEGATION_STATE:
            case UnindexedFieldURIType.TASK_DELEGATOR:
            case UnindexedFieldURIType.TASK_MILEAGE:
            case UnindexedFieldURIType.TASK_OWNER:
            case UnindexedFieldURIType.TASK_TOTAL_WORK:
                Logger.getInstance().info(`Field ${fieldIdentifier} not yet supported for Tasks`);
                break;
            case UnindexedFieldURIType.TASK_RECURRENCE: {
                    // Outlook items do not show if the recurrence field is set but undefined.
                    const recurrence = this.getEWSRecurrence(pimTask);
                    if (recurrence) {
                        task.Recurrence = recurrence;
                    }
                }
                break;
            case UnindexedFieldURIType.TASK_IS_RECURRING:
                task.IsRecurring = (pimTask.recurrenceRules ?? []).length > 0;
                break;
            case UnindexedFieldURIType.TASK_COMPLETE_DATE:
                if (pimTask.isComplete) {
                    task.CompleteDate = pimTask.completedDate;
                }
                break;
            case UnindexedFieldURIType.TASK_IS_COMPLETE:
                task.IsComplete = pimTask.isComplete;
                break;
            case UnindexedFieldURIType.TASK_STATUS_DESCRIPTION:
                task.StatusDescription = pimTask.statusLabel;
                break;
            // The following fields that apply to tasks are handled in common code.
            // case UnindexedFieldURIType.ITEM_HAS_ATTACHMENTS:
            // case UnindexedFieldURIType.ITEM_SUBJECT:
            // case UnindexedFieldURIType.ITEM_BODY:
            // case UnindexedFieldURIType.ITEM_CATEGORIES:
            // case UnindexedFieldURIType.ITEM_LAST_MODIFIED_TIME:
            // case UnindexedFieldURIType.ITEM_LAST_MODIFIED_NAME:
            // case UnindexedFieldURIType.ITEM_CONVERSATION_ID:
            // case UnindexedFieldURIType.ITEM_CULTURE:
            // case UnindexedFieldURIType.ITEM_DATE_TIME_CREATED:
            // case UnindexedFieldURIType.ITEM_DATE_TIME_RECEIVED:
            // case UnindexedFieldURIType.ITEM_DATE_TIME_SENT:
            // case UnindexedFieldURIType.ITEM_DISPLAY_CC:
            // case UnindexedFieldURIType.ITEM_DISPLAY_TO:
            // case UnindexedFieldURIType.ITEM_EFFECTIVE_RIGHTS:
            // case UnindexedFieldURIType.ITEM_IMPORTANCE:
            // case UnindexedFieldURIType.ITEM_IN_REPLY_TO:
            // case UnindexedFieldURIType.ITEM_INTERNET_MESSAGE_HEADERS:
            // case UnindexedFieldURIType.ITEM_IS_ASSOCIATED:
            // case UnindexedFieldURIType.ITEM_IS_DRAFT:
            // case UnindexedFieldURIType.ITEM_IS_FROM_ME:
            // case UnindexedFieldURIType.ITEM_IS_RESEND:
            // case UnindexedFieldURIType.ITEM_IS_SUBMITTED:
            // case UnindexedFieldURIType.ITEM_IS_UNMODIFIED:
            // case UnindexedFieldURIType.ITEM_REMINDER_DUE_BY:
            // case UnindexedFieldURIType.ITEM_REMINDER_IS_SET:
            // case UnindexedFieldURIType.ITEM_REMINDER_MINUTES_BEFORE_START:
            // case UnindexedFieldURIType.ITEM_RESPONSE_OBJECTS:
            // case UnindexedFieldURIType.ITEM_SENSITIVITY:
            // case UnindexedFieldURIType.ITEM_SIZE:
            // case UnindexedFieldURIType.ITEM_WEB_CLIENT_EDIT_FORM_QUERY_STRING:
            // case UnindexedFieldURIType.ITEM_WEB_CLIENT_READ_FORM_QUERY_STRING:
            default:
                handled = false;
        }

        if (handled === false) {
            // Common fields handled in superclass
            handled = super.updateEWSItemFieldValue(task, pimTask, fieldIdentifier, mailboxId);
        }

        if (handled === false) {
            Logger.getInstance().warn(`Unrecognized field ${fieldIdentifier} not processed for Task`);
        }

        return handled;
    }

    /**
     * Update a value for an EWS field in a Keep PIM task. 
     * @param pimItem The pim task that will be updated.
     * @param fieldIdentifier The EWS unindexed field identifier
     * @param newValue The new value to set. The type is based on what fieldIdentifier is set to. To delete a field, pass in undefined. 
     * @returns true if field was handled
     */
    protected updatePimItemFieldValue(pimTask: PimTask, fieldIdentifier: UnindexedFieldURIType, newValue: any | undefined): boolean {
        if (!super.updatePimItemFieldValue(pimTask, fieldIdentifier, newValue)) {
            if (fieldIdentifier === UnindexedFieldURIType.TASK_COMPLETE_DATE) {
                if (newValue === undefined) {
                    pimTask.completedDate = undefined;
                }
                else {
                    pimTask.completedDate = new Date(newValue);
                }
            }
            else if (fieldIdentifier === UnindexedFieldURIType.TASK_DUE_DATE) {
                pimTask.due = newValue instanceof Date ? newValue.toISOString() : newValue;
            }
            else if (fieldIdentifier === UnindexedFieldURIType.TASK_IS_COMPLETE) {
                // TODO:  Should we clear the completed date if this is set to false?  Maybe this should
                //        be done in the pimTask isComplete setter?
                pimTask.isComplete = newValue ?? false;
            }
            else if (fieldIdentifier === UnindexedFieldURIType.TASK_START_DATE) {
                pimTask.start = newValue instanceof Date ? newValue.toISOString() : newValue;
            }
            else if (fieldIdentifier === UnindexedFieldURIType.TASK_PERCENT_COMPLETE) {
                // TODO:  Should we do more here?
                if (newValue === undefined) {
                    pimTask.isComplete = false;
                }
                else if (newValue === 100) {
                    pimTask.isComplete = true;
                }
            }
            else if (fieldIdentifier === UnindexedFieldURIType.TASK_STATUS) {
                // NOTE:  Each set below updates an underlying task state value in the pimTask, so state is maintained properly
                if (newValue === undefined) {
                    pimTask.isNotStarted = false;
                }
                else if (newValue === TaskStatusType.NOT_STARTED) {
                    pimTask.isNotStarted = true;
                }
                else if (newValue === TaskStatusType.IN_PROGRESS) {
                    pimTask.isInProgress = true;
                }
                else if (newValue === TaskStatusType.COMPLETED) {
                    pimTask.isComplete = true;
                } else {
                    Logger.getInstance().warn(`TASK_STATUS value of ${fieldIdentifier} not supported`);
                }
                // Other possible values are "WaitingOnOhters" and "Deferred", which do not map to Keep values
            }
            else if (fieldIdentifier === UnindexedFieldURIType.TASK_STATUS_DESCRIPTION) {
                pimTask.statusLabel = newValue ?? "";
            }
            else {
                Logger.getInstance().error(`Unsupported field ${fieldIdentifier} for PIM task ${pimTask.unid}`);
                return false;
            }

            /*
            The following item fields are not currently supported:
    
            TASK_ACTUAL_WORK = 'task:ActualWork',
            TASK_ASSIGNED_TIME = 'task:AssignedTime',
            TASK_BILLING_INFORMATION = 'task:BillingInformation',
            TASK_CHANGE_COUNT = 'task:ChangeCount',
            TASK_COMPANIES = 'task:Companies',
            TASK_CONTACTS = 'task:Contacts',
            TASK_DELEGATION_STATE = 'task:DelegationState',
            TASK_DELEGATOR = 'task:Delegator',
            TASK_IS_ASSIGNMENT_EDITABLE = 'task:IsAssignmentEditable',
            TASK_IS_RECURRING = 'task:IsRecurring',
            TASK_IS_TEAM_TASK = 'task:IsTeamTask',
            TASK_MILEAGE = 'task:Mileage',
            TASK_OWNER = 'task:Owner',
            TASK_PERCENT_COMPLETE = 'task:PercentComplete',
            TASK_RECURRENCE = 'task:Recurrence',
            TASK_TOTAL_WORK = 'task:TotalWork',
            */

        }
        return true;
    }

    pimItemFromEWSItem(item: TaskType, request: Request, existing?: object): PimTask {
        const rtn = PimItemFactory.newPimTask();

        if (item.Subject) {
            rtn.title = item.Subject;
        }

        if (item.Body) {
            if (item.Body.BodyType === BodyTypeType.HTML) {
                rtn.description = fromString(item.Body.Value);
            }
            else {
                rtn.description = item.Body.Value;
            }
        }

        // TODO: Handle recurring tasks (IsRecurring?: boolean, Recurrence?: TaskRecurrenceType)

        if (item.Categories && item.Categories.String.length > 0) {
            rtn.categories = item.Categories.String;
        }

        if (item.ReminderIsSet) {
            rtn.alarm = item.ReminderMinutesBeforeStart ? -Number.parseInt(item.ReminderMinutesBeforeStart) : 0;
        }

        switch (item.Importance) {
            case ImportanceChoicesType.HIGH:
                rtn.priority = PimTaskPriority.HIGH;
                break;
            case ImportanceChoicesType.LOW:
                rtn.priority = PimTaskPriority.LOW;
                break;
            default: // Normal
                rtn.priority = PimTaskPriority.MEDIUM;
                break;
        }

        if (item.Sensitivity === SensitivityChoicesType.CONFIDENTIAL) {
            rtn.isConfidential = true;
        }
        else {
            rtn.isConfidential = false;
        }

        // EWS Task do not have a privacy setting, so mark all as public. 
        rtn.isPrivate = true;

        if (item.StartDate) {
            rtn.start = item.StartDate.toISOString();
            // } else {
            //     rtn.start = [new Date()];
        }

        if (item.DueDate) {
            rtn.due = item.DueDate.toISOString();
            // if (item.ReminderIsSet) {
            //     taskObject.ReminderDueBy = item.DueDate.toISOString();
            // }
        }

        // The alarm may also be set in the extended field URI 
        let reminderIdentifiers: any = {};
        reminderIdentifiers[ExtendedPropertyKeyType.DISTINGUISHED_PROPERTY_SETID] = DistinguishedPropertySetType.COMMON;
        reminderIdentifiers[ExtendedPropertyKeyType.PROPERTY_ID] = MapiPropertyIds.REMINDER_SET;
        const alarm = getValueForExtendedFieldURI(item, reminderIdentifiers); // Alarm is set
        if (alarm) {
            // taskObject["$Alarm"] = 1;
            reminderIdentifiers = {};
            reminderIdentifiers[ExtendedPropertyKeyType.DISTINGUISHED_PROPERTY_SETID] = DistinguishedPropertySetType.COMMON;
            reminderIdentifiers[ExtendedPropertyKeyType.PROPERTY_ID] = MapiPropertyIds.REMINDER_TIME;
            const alarmTime: Date = getValueForExtendedFieldURI(item, reminderIdentifiers); // Alarm time
            const compareDate = item.StartDate ?? item.DueDate;
            if (compareDate) {
                const diff = compareDate.getTime() - alarmTime.getTime();
                const offset = Math.floor(diff / 60000); // Alarm offset is in minutes
                // taskObject["$AlarmOffset"] = (offset < 0) ? offset : -offset;
                rtn.alarm = -offset;
            } else {
                // Setting alarm since an alarm was set, there just wasn't an offset
                rtn.alarm = 0;
            }
        }

        /*
            Values for DueState:
            9 = Completed
            0 = Overdue
            1 = In Progress
            2 = No Started
            8 = Rejected
        */
        if (item.IsComplete || item.Status === TaskStatusType.COMPLETED) {
            if (item.CompleteDate) {
                rtn.completedDate = item.CompleteDate;
            }
            rtn.isComplete = true;
        }
        else if (item.Status === TaskStatusType.IN_PROGRESS || item.Status === TaskStatusType.WAITING_ON_OTHERS) {
            rtn.isInProgress = true;
        }
        else if (item.Status === TaskStatusType.NOT_STARTED || item.Status === TaskStatusType.DEFERRED) {
            rtn.isNotStarted = true;
        }

        if (item.StatusDescription) {
            rtn.statusLabel = item.StatusDescription;
        }

        // "TaskType" has the following meainging. TODO: Support assignments
        // "1" : your own task
        // "2" : task you created, assigned to someone else
        rtn.taskType = "1";

        // From is expected, if not set then us the current user. 
        const name = UserContext.getSubjectFromRequest(request);
        if (name) {
            rtn.from = name;
            rtn.altChair = name;
        }
        else {
            Logger.getInstance().error(`Task does not contain From`);
        }

        // Copy any extended fields to the PimMessage to preserve them.
        addExtendedPropertiesToPIM(item, rtn);

        return rtn;
    }
}
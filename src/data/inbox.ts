/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import { Request } from '@loopback/rest';
import { 
    KeepPimRuleManager, PimRule, KeepPimLabelManager, PimRuleSeparatorConstants, PimRuleOperatorConstants, 
    PimRuleConditionConstants, PimRuleActionConstants 
} from '@hcllabs/openclientkeepcomponent';
import { UserContext } from '../keepcomponent';
import {
    ResponseClassType, ResponseCodeType, RuleFieldURIType, RuleValidationErrorCodeType, ImportanceChoicesType
} from '../models/enum.model';
import { 
    ArrayOfRuleOperationErrorsType, RuleOperationErrorType, ArrayOfRulesType, GetInboxRulesRequestType, 
    GetInboxRulesResponseType, RuleActionsType, RulePredicatesType, RuleType, UpdateInboxRulesRequestType, 
    RuleValidationErrorType, UpdateInboxRulesResponseType, DeleteRuleOperationType, SetRuleOperationType, 
    CreateRuleOperationType, ArrayOfRuleValidationErrorsType 
} from '../models/inbox.model';
import { FolderIdType, TargetFolderIdType, ArrayOfEmailAddressesType, EmailAddressType } from '../models/mail.model';
import { ArrayOfStringsType } from '../models/common.model';
import { findLabelByTarget, Logger, getEWSId } from '../utils';

/**
 * Get inbox rules.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns inbox rules
 */
export async function GetInboxRules(soapRequest: GetInboxRulesRequestType, request: Request): Promise<GetInboxRulesResponseType> {
    const userInfo = UserContext.getUserInfo(request);

    const response = new GetInboxRulesResponseType();
    response.ResponseClass = ResponseClassType.SUCCESS;
    response.ResponseCode =  ResponseCodeType.NO_ERROR;
    response.OutlookRuleBlobExists = true;

    if (soapRequest.MailboxSmtpAddress !== userInfo.userId) { // Only return response for current user for now. 
        response.ResponseClass = ResponseClassType.ERROR;
        response.ResponseCode = ResponseCodeType.ERROR_ACCESS_DENIED;
        response.MessageText = `The authenticated user does not have access to rules for ${soapRequest.MailboxSmtpAddress}`;
        return response;
    }

    let pimRules: PimRule[] = [];
    try {
        pimRules = await KeepPimRuleManager.getInstance().getRules(userInfo);
    } catch (err) {
        Logger.getInstance().error(`GetInboxRules is unable to get PIM rules from the server: ${err}`);
        throw err;
    }

    if (pimRules && pimRules.length > 0) {
        response.InboxRules = new ArrayOfRulesType();
        response.InboxRules.Rule = [];
        for (const pimRule of pimRules) {
            const ewsRule = await pimRuleToRule(request, pimRule);
            if (ewsRule) {
                response.InboxRules.Rule.push(ewsRule);
            }
        }
    }
    return response;
}


/**
 * Update the inbox rules.
 * @param soapRequest The SOAP request
 * @param request The HTTP request
 * @returns Any rule operation errors
 */
export async function UpdateInboxRules(soapRequest: UpdateInboxRulesRequestType, request: Request): Promise<UpdateInboxRulesResponseType> {
    const userInfo = UserContext.getUserInfo(request);

    const response = new UpdateInboxRulesResponseType();
    response.ResponseClass = ResponseClassType.SUCCESS;
    response.ResponseCode =  ResponseCodeType.NO_ERROR;

    if (soapRequest.MailboxSmtpAddress !== userInfo.userId) { // Only return response for current user for now. 
        response.ResponseClass = ResponseClassType.ERROR;
        response.ResponseCode = ResponseCodeType.ERROR_ACCESS_DENIED;
        response.MessageText = `The authenticated user does not have access to update rules for ${soapRequest.MailboxSmtpAddress}`;
        return response;
    }

    //Create the new rules
    const opErrors: RuleOperationErrorType[] = [];
    // for (const ruleOperation of soapRequest.Operations) {
    for (let index = 0; index < soapRequest.Operations.items.length; index++) {
        const ruleOperation = soapRequest.Operations.items[index];
        let opError: RuleOperationErrorType | undefined;
        if (ruleOperation instanceof SetRuleOperationType) {
            opError = await updateRuleOperation(request, ruleOperation);
        } else if (ruleOperation instanceof CreateRuleOperationType) {
            opError = await createRuleOperation(request, ruleOperation);
        } else if (ruleOperation instanceof DeleteRuleOperationType) {
            opError = await deleteRuleOperation(request, ruleOperation);
        }
        if (opError) {
            opError.OperationIndex = index;
            opErrors.push(opError);
        }
    }
    
    if (opErrors.length > 0) {
        response.RuleOperationErrors = new ArrayOfRuleOperationErrorsType();
        response.RuleOperationErrors.RuleOperationError = opErrors;
    }

    return response;
}


async function deleteRuleOperation(request: Request, ruleOperation: DeleteRuleOperationType): Promise<RuleOperationErrorType | undefined> {
    Logger.getInstance().debug(`Deleting rule operation ${ruleOperation.RuleId}`);
    let opError: RuleOperationErrorType | undefined;
    const userInfo = UserContext.getUserInfo(request);

    if (ruleOperation.RuleId) {
        try {
            await KeepPimRuleManager.getInstance().deleteRule(userInfo, ruleOperation.RuleId);
        } catch (err) {
            opError = new RuleOperationErrorType();
            opError.ValidationErrors = new ArrayOfRuleValidationErrorsType();
            const valErrors: RuleValidationErrorType[] = [];
            const valError = new RuleValidationErrorType();
            valError.FieldURI = RuleFieldURIType.RULE_ID;
            valError.ErrorCode = (err.status && err.status === 404) ? RuleValidationErrorCodeType.RULE_NOT_FOUND : RuleValidationErrorCodeType.UNEXPECTED_ERROR;
            valError.ErrorMessage = err;
            valError.FieldValue = 'invalid';
            valErrors.push(valError);
            opError.ValidationErrors.Error = valErrors;
        }
    } else {
        opError = new RuleOperationErrorType();
        opError.ValidationErrors = new ArrayOfRuleValidationErrorsType();
        const valErrors: RuleValidationErrorType[] = [];
        const valError = new RuleValidationErrorType();
        valError.FieldURI = RuleFieldURIType.RULE_ID;
        valError.ErrorCode = RuleValidationErrorCodeType.UNSUPPORTED_RULE;
        valError.ErrorMessage = 'Error deleting a rule.  A rule id is required';
        valError.FieldValue = '32 character id generated on the server when the rule is created.';
        valErrors.push(valError);
        opError.ValidationErrors.Error = valErrors;
    }

    return Promise.resolve(opError);
}

async function updateRuleOperation(request: Request, ruleOperation: SetRuleOperationType): Promise<RuleOperationErrorType | undefined> {
    Logger.getInstance().debug(`Setting rule operation ${ruleOperation.Rule.DisplayName}`);
    let opError: RuleOperationErrorType | undefined;
    const userInfo = UserContext.getUserInfo(request);

    if (ruleOperation.Rule.RuleId) {
        try {
            await KeepPimRuleManager.getInstance().updateRule(userInfo, ruleOperation.Rule.RuleId, await getRuleStructure(request, ruleOperation));
        } catch (err) {
            opError = new RuleOperationErrorType();
            opError.ValidationErrors = new ArrayOfRuleValidationErrorsType();
            const valErrors: RuleValidationErrorType[] = [];
            const valError = new RuleValidationErrorType();
            valError.FieldURI = RuleFieldURIType.IS_NOT_SUPPORTED;
            valError.ErrorCode = RuleValidationErrorCodeType.UNSUPPORTED_RULE;
            valError.ErrorMessage = err;
            valError.FieldValue = 'invalid';
            valErrors.push(valError);
            opError.ValidationErrors.Error = valErrors;
        }
    } else {
        opError = new RuleOperationErrorType();
        opError.ValidationErrors = new ArrayOfRuleValidationErrorsType();
        const valErrors: RuleValidationErrorType[] = [];
        const valError = new RuleValidationErrorType();
        valError.FieldURI = RuleFieldURIType.RULE_ID;
        valError.ErrorCode = RuleValidationErrorCodeType.UNSUPPORTED_RULE;
        valError.ErrorMessage = 'A rule id is required';
        valError.FieldValue = '32 character id generated on the server when the rule is created.';
        valErrors.push(valError);
        opError.ValidationErrors.Error = valErrors;
    }

    return Promise.resolve(opError);
}

async function createRuleOperation(request: Request, ruleOperation: CreateRuleOperationType): Promise<RuleOperationErrorType | undefined> {
    Logger.getInstance().debug(`Creating rule operation ${ruleOperation.Rule.DisplayName}`);
    let opError: RuleOperationErrorType | undefined;
    const userInfo = UserContext.getUserInfo(request);
    
    try {
        await KeepPimRuleManager.getInstance().createRule(userInfo, await getRuleStructure(request, ruleOperation));
    } catch (err) {
        opError = new RuleOperationErrorType();
        opError.ValidationErrors = new ArrayOfRuleValidationErrorsType();
        const valErrors: RuleValidationErrorType[] = [];
        const valError = new RuleValidationErrorType();
        valError.FieldURI = RuleFieldURIType.IS_NOT_SUPPORTED;
        valError.ErrorCode = RuleValidationErrorCodeType.UNEXPECTED_ERROR;
        valError.ErrorMessage = err;
        valError.FieldValue = 'invalid';
        valErrors.push(valError);
        opError.ValidationErrors.Error = valErrors;
    }

    return Promise.resolve(opError);
}


async function getRuleStructure(request: Request, ruleOperation: SetRuleOperationType | CreateRuleOperationType): Promise<any> {
        // EWS Rule fields
        // ===================
        // RuleId?: string;
        // DisplayName: string;
        // Priority: number;
        // IsEnabled: boolean;
        // IsNotSupported?: boolean;
        // IsInError?: boolean;
        // Conditions?: RulePredicatesType;
        // Exceptions?: RulePredicatesType;
        // Actions?: RuleActionsType;

        // PimRule create structure
        // ==========================
        // {
        //     "ActionList": [
        //       " move to folder ($JunkMail)"
        //     ],
        //     "ConditionList": [
        //       "   Sender contains @junk"
        //     ],
        //     "Enable": "1",
        //     "TokActionList": [
        //       "1¦1¦($JunkMail)"
        //     ],
        //     "tokConditionList": [
        //       "1¦1¦@junk¦0"
        //     ]
        //   }

        const rule = ruleOperation.Rule;
        const ruleStructure: any = {};
        // We're not going to allow the update of the RuleId
        // RuleId?: string;

        ruleStructure.DisplayName = rule.DisplayName;
        ruleStructure.Priority = rule.Priority;
        ruleStructure.Enable = rule.IsEnabled ? "1" : "0";
        if (rule.IsNotSupported) {
            ruleStructure.IsNotSupported = rule.IsNotSupported;
        }
        if (rule.IsInError) {
            ruleStructure.IsInError = rule.IsInError;
        }

        addPimRuleConditions(rule.Conditions, ruleStructure);
        addPimRuleExceptions(rule.Exceptions, ruleStructure);
        await addPimRuleActions(request, rule.Actions, ruleStructure);

        return Promise.resolve(ruleStructure);
    }

// Convert from an EWS Condition to a PimCondition
function addPimRuleConditions(ewsConditions: RulePredicatesType | undefined, ruleStructure: any): void {

    // Pim Condition example
    // =======================
    //     "ConditionList": [
    //       "   Sender contains @junk"
    //     ],

    // EWS rule predicate candidates
    // =============================
    // Categories?: ArrayOfStringsType;
    // ContainsBodyStrings?: ArrayOfStringsType;
    // ContainsHeaderStrings?: ArrayOfStringsType;
    // ContainsRecipientStrings?: ArrayOfStringsType;
    // ContainsSenderStrings?: ArrayOfStringsType;
    // ContainsSubjectOrBodyStrings?: ArrayOfStringsType;
    // ContainsSubjectStrings?: ArrayOfStringsType;
    // FlaggedForAction?: FlaggedForActionType;
    // FromAddresses?: ArrayOfEmailAddressesType;
    // FromConnectedAccounts?: ArrayOfStringsType;
    // HasAttachments?: boolean;
    // Importance?: ImportanceChoicesType;
    // IsApprovalRequest?: boolean;
    // IsAutomaticForward?: boolean;
    // IsAutomaticReply?: boolean;
    // IsEncrypted?: boolean;
    // IsMeetingRequest?: boolean;
    // IsMeetingResponse?: boolean;
    // IsNDR?: boolean;
    // IsPermissionControlled?: boolean;
    // IsReadReceipt?: boolean;
    // IsSigned?: boolean;
    // IsVoicemail?: boolean;
    // ItemClasses?: ArrayOfStringsType;
    // MessageClassifications?: ArrayOfStringsType;
    // NotSentToMe?: boolean;
    // SentCcMe?: boolean;
    // SentOnlyToMe?: boolean;
    // SentToAddresses?: ArrayOfEmailAddressesType;
    // SentToMe?: boolean;
    // SentToOrCcMe?: boolean;
    // Sensitivity?: SensitivityChoicesType;
    // WithinDateRange?: RulePredicateDateRangeType;
    // WithinSizeRange?: RulePredicateSizeRangeType;


    // Column Mappings for tokConditionList
    // ====================================
    //     Condition | Logic | StringValue | Operator
    // Condition =
    //     Case "1"
    //         FieldString = GetString(COND_SENDER)
        
    //     Case "2"
    //         FieldString =  GetString(COND_SUBJECT)
        
    //     Case "3"
    //         FieldString = GetString(COND_BODY)
        
    //     Case "4"
    //         FieldString = GetString(COND_IMPORTANCE)
        
    //     Case "5"
    //         FieldString = GetString(COND_DELPRIORITY)
        
    //     Case "6"
    //         FieldString =GetString(COND_TO)
        
    //     Case "7"
    //         FieldString=GetString(COND_CC)
        
    //     Case "8"
    //         FieldString = GetString(COND_TOORCC)
        
    //     Case "9"
    //         FieldString = GetString(COND_BODYORSUBJECT)
        
    //     Case "A"
    //         FieldString=GetString(COND_BCC)
        
    //     Case "B"
    //         FieldString = GetString(COND_INTERNETDOMAIN)
        
    //     Case "C"
    //         FieldString = GetString(COND_SIZE)
        
    //     Case "D"
    //         FieldString = GetString(COND_ALLDOCS)
        
    //     Case "E"
    //         FieldString = GetString(COND_ATTCHNAMES)
        
    //     Case "F"
    //         FieldString = GetString(COND_ATTCHNUMBER)
        
    //     Case "G"
    //         FieldString = GetString(COND_FORM)
        
    //     Case "H"
    //         FieldString = GetString(COND_RECIPIENTCOUNT)
        
    //     Case "I"
    //         FieldString = GetString(COND_ANYRECIPIENT)
        
    //     Case "J"
    //         FieldString = GetString(COND_CUSTOMIZE)
        
    //     Case "K"
    //         FieldString = GetString(COND_BLACKLIST)
        
    //     Case "L"
    //         FieldString = GetString(COND_WHITELIST)
    // Logic =
    //     Case "1"
    //         If (strCond = "C") Or (strCond = "H")  Or (strCond = "F") Then
    //             logicstring = GetString (LOGIC_NUMLESS)           
    //         Else
    //             logicstring = GetString(LOGIC_CONTAINS)           
    //         End If
    //     Case "2"
    //         If (strCond = "C") Or (strCond = "H")  Or (strCond = "F")Then
    //             logicstring = GetString (LOGIC_NUMGREATER)           
    //         Else
    //             logicstring = GetString (LOGIC_CONTAINSNOT)
    //         End If
    //     Case "3"
    //         logicstring = GetString(LOGIC_IS)
    //     Case "4"
    //         logicstring = GetString (LOGIC_ISNOT)
    // StringValue = <token>
    // Operator = "0" for AND, "1" for OR
    let conditionList: string[] | undefined;
    let tokConditionList: string[] | undefined;
    if (ewsConditions) {
        conditionList = [];
        tokConditionList = [];
        if (ewsConditions.ContainsSubjectStrings) {
            for (let i = 0; i < ewsConditions.ContainsSubjectStrings.String.length; i++) {
                const prefix = i > 0 ? 'OR ' : (conditionList.length > 0 ? 'AND ' : '');
                conditionList.push( `${prefix}Subject contains ${ewsConditions.ContainsSubjectStrings.String[i]} `);
                tokConditionList.push(`2¦1¦${ewsConditions.ContainsSubjectStrings.String[i]}¦0`);  // Subject | contains | <token> | AND
            }
        }

        if (ewsConditions.ContainsSubjectOrBodyStrings) {
            for (let i = 0; i < ewsConditions.ContainsSubjectOrBodyStrings.String.length; i++) {
                const prefix = i > 0 ? 'OR ' : (conditionList.length > 0 ? 'AND ' : '');
                conditionList.push( `${prefix}Body or Subject contains ${ewsConditions.ContainsSubjectOrBodyStrings.String[i]} `);
                tokConditionList.push(`9¦1¦${ewsConditions.ContainsSubjectOrBodyStrings.String[i]}¦0`);  // Body or Subject | contains | <token> | AND
            }
        }

        if (ewsConditions.ContainsSenderStrings) {
            // const sString = ewsConditions.ContainsSenderStrings.String.length > 1 ? ewsConditions.ContainsSenderStrings.String.join(' ') : ewsConditions.ContainsSenderStrings.String[0];
            // conditionList.push( `${conditionList.length > 0 ? 'AND ' : ''}Sender contains ${sString} `);
            // tokConditionList.push(`1¦1¦${sString}¦0`);  // Sender | contains | <token> | AND
            for (let i = 0; i < ewsConditions.ContainsSenderStrings.String.length; i++) {
                const prefix = i > 0 ? 'OR ' : (conditionList.length > 0 ? 'AND ' : '');
                conditionList.push( `${prefix}Sender contains ${ewsConditions.ContainsSenderStrings.String[i]} `);
                tokConditionList.push(`1¦1¦${ewsConditions.ContainsSenderStrings.String[i]}¦0`);  // Sender | contains | <token> | AND
            }
        }

        if (ewsConditions.FromAddresses) {
            for (let i = 0; i < ewsConditions.FromAddresses.Address.length; i++) {
                const prefix = i > 0 ? 'OR ' : (conditionList.length > 0 ? 'AND ' : '');
                const operator = i > 0 ? 1 : 0;  // 1 === or, 0 === and
                conditionList.push( `${prefix}Sender is ${ewsConditions.FromAddresses.Address[i].EmailAddress} `);
                tokConditionList.push(`1¦3¦${ewsConditions.FromAddresses.Address[i].EmailAddress}¦${operator}`);  // Sender | is | <token> | <operator>
            }
        }
    }

    if (conditionList) {
        ruleStructure.ConditionList = conditionList ;
    }
    if (tokConditionList) {
        ruleStructure.tokConditionList = tokConditionList;
    }
}

// Convert from an EWS Exception to a PimException
function addPimRuleExceptions(ewsExceptions: RulePredicatesType | undefined, ruleStructure: any): undefined {
    // if (ruleExceptions) {
    //     ruleStructure.RuleExceptions = ruleExceptions;
    // }

    return undefined;
}

// Convert from an EWS rule action to a PimAction
async function addPimRuleActions(request: Request, ewsActions: RuleActionsType | undefined, ruleStructure: any): Promise<void> {
    const userInfo = UserContext.getUserInfo(request);
    let actionList: string[] | undefined;
    let tokActionList: string[] | undefined;
    if (ewsActions) {
        actionList = [];
        tokActionList = [];
        // EWS Actions
        // ==================
        // AssignCategories?: ArrayOfStringsType;
        // CopyToFolder?: TargetFolderIdType;
        // Delete?: boolean;
        // ForwardAsAttachmentToRecipients?: ArrayOfEmailAddressesType;
        // ForwardToRecipients?: ArrayOfEmailAddressesType;
        // MarkImportance?: ImportanceChoicesType;
        // MarkAsRead?: boolean;
        // MoveToFolder?: TargetFolderIdType;
        // PermanentDelete?: boolean;
        // RedirectToRecipients?: ArrayOfEmailAddressesType;
        // SendSMSAlertToRecipients?: ArrayOfEmailAddressesType;
        // ServerReplyWithMessage?: ItemIdType;
        // StopProcessingRules?: boolean;



        // TokActionList columns
        // =======================
        //Build Action Token phrase with this format =======>> Action | Logic | StringValue | Operator
        //         Select Case ACT_sAction
        //         Case "1"            'MovetoFolder
        //             note.tokactionlist = CStr(ACT_sAction + STR_TOKENSEPARATOR + "1" + STR_TOKENSEPARATOR + ACT_sString )
        //         Case "2"            'Importance
        //             note.tokactionlist = CStr(ACT_sAction + STR_TOKENSEPARATOR + "1" + STR_TOKENSEPARATOR + ACT_sImportance )
        //         Case "3"            'Delete
        //             note.tokactionlist = CStr(ACT_sAction + STR_TOKENSEPARATOR + "1" + STR_TOKENSEPARATOR)
        //         Case "A"
        //             Call itmTokActionList.AppendToTextList( ACT_sAction + STR_TOKENSEPARATOR + ACT_sExpDate + STR_TOKENSEPARATOR + ACT_sExpNumber)
        //         Case "B"
        //             Call itmTokActionList.AppendToTextList( ACT_sAction + STR_TOKENSEPARATOR + ACT_sChoice + STR_TOKENSEPARATOR+ ArrayToString(ACT_vAddress, ":") )
        //         Case "8"
        //             Call itmTokActionList.AppendToTextList( ACT_sAction + STR_TOKENSEPARATOR + ACT_sBehavior )
        //         Case Else
        //             Call itmTokActionList.AppendToTextList( ACT_sAction + STR_TOKENSEPARATOR + ACT_sImportance + STR_TOKENSEPARATOR + ACT_sFolder)   
        
        
        // Where ACT_sAction is (ActionString below):
        // ActionString
        //     Case "1"
        //         ActionString = GetSTring (ACT_MOVETOFOLDER)+ ACT_sFolder
        //     Case "2"   
        //         Select Case ACT_sImportance
        //         Case "1"
        //             ActionString = GetString(ACT_CHANGEIMP) + GetString(ACT_HIGH)
        //         Case "2"
        //             ActionString = GetString(ACT_CHANGEIMP) + GetString(ACT_NORMAL)
        //         Case "3"
        //             ActionString = GetString(ACT_CHANGEIMP) + GetString(ACT_LOW)
        //         End Select
        //     Case "3"
        //         ActionString = GetString(ACT_DELETE)
        //     Case "4"
        //         ActionString =GetString(ACT_COPYTOFOLDER) + ACT_sFolder
        //     Case "5"
        //         ActionString = GetString(ACT_JOURNAL)
        //     Case "6"
        //         ActionString = GetString(ACT_MOVETODB) + ACT_sFolder
        //     Case "7"
        //         ActionString = GetString(ACT_DELETE)
        //     Case "8"
        //         ActionString = GetString(ACT_DONOTDELIVER) + BehaviorString(ACT_sBehavior)
        //     Case "9"
        //         ActionString = GetString(ACT_ROUTINGSTATE) + BehaviorString(ACT_sBehavior)
        //     Case "A"   
        //         ActionString = GetString( ACT_SETEXPIREDATE ) +| | + ACT_sExpNumber + | | + BehaviorString(ACT_sExpDate)
        //     Case "B"
        //         Select Case ACT_sChoice
        //         Case "1"
        //             ActionString = GetString(ACT_COPYTO_1) + | | + ArrayToString(ACT_vAddress, ", ")
        //         Case "2"
        //             ActionString = GetString(ACT_COPYTO_2) + | | + ArrayToString(ACT_vAddress, ", ")
        //         Case "3"
        //             ActionString = GetString(ACT_COPYTO_3) + | | + ArrayToString(ACT_vAddress, ", ")
        //         Case "4"
        //             ActionString = GetString(ACT_COPYTO_4) + | | + ArrayToString(ACT_vAddress, ", ")
        //         End Select
        //     Case "C"
        //         ActionString = GetString(ACT_STOP_PROCESSING)
        //     Case Else
        //         ActionString = ""
        if (ewsActions.CopyToFolder) {
            Logger.getInstance().debug(`Processing rule action CopyToFolder`);
            const currentPimLabels = await KeepPimLabelManager.getInstance().getLabels(userInfo);
            const copyPimLabel = findLabelByTarget(currentPimLabels, ewsActions.CopyToFolder);
            if (copyPimLabel) {
                addAction(` copy to folder ${copyPimLabel.view} `, actionList);
                tokActionList.push(`4¦1¦${copyPimLabel.view}`);
            }
        }

        if (ewsActions.MoveToFolder) {
            Logger.getInstance().debug(`Processing rule action MoveToFolder`);
            const currentPimLabels = await KeepPimLabelManager.getInstance().getLabels(userInfo);
            const movePimLabel = findLabelByTarget(currentPimLabels, ewsActions.MoveToFolder);
            if (movePimLabel) {
                addAction( ` move to folder ${movePimLabel.view} `, actionList);
                tokActionList.push(`1¦1¦${movePimLabel.view}`);
            }
        }

        if (ewsActions.ForwardToRecipients) {
            Logger.getInstance().debug(`Processing rule action ForwardToRecipients`);
        //     "ActionList": " send Full Copy to davek1121@gmail.com",
        //     "TokActionList": "B¦1¦davek1121@gmail.com",
          
        //    "ActionList": "send Copy of Headers to David Kennedy/USA/PNPHCL",
        //    "TokActionList": "B¦4¦David Kennedy/USA/PNPHCL",
            const recips = ewsActions.ForwardToRecipients.Address.map(emailAddress => {return emailAddress.EmailAddress}).join(", ");
            addAction( ` ${PimRuleActionConstants.SEND_FULL_COPY} ${recips}`, actionList);
            tokActionList.push(`B¦1¦${recips}`);

        }

        if (ewsActions.Delete) {
            Logger.getInstance().debug(`Processing rule action Delete`);
            // "ActionList": " don't Accept Message",
            // "TokActionList": "3¦1¦",
            addAction( ` ${PimRuleActionConstants.DELETE} `, actionList);
            tokActionList.push(`3¦1¦`);
        }

        if (ewsActions.StopProcessingRules) {
            Logger.getInstance().debug(`Processing rule action StopProcessingRules`);
            // "ActionList": " Stop Processing further Rules",
            // "TokActionList": "C¦1¦",
            addAction( ` ${PimRuleActionConstants.STOP_PROCESSING} `, actionList);
            tokActionList.push(`C¦1¦`);

        }

        if (ewsActions.MarkImportance) {
            Logger.getInstance().debug(`Processing rule action Importance`);
            // "ActionList": " change Importance to High",
            // "TokActionList": "2¦1¦",
            // "ActionList": " change Importance to Low",
            // "TokActionList": "2¦3¦",

            let importanceValue = 2;  // Importance = Normal
            let importanceText = "Normal";
            if (ewsActions.MarkImportance === ImportanceChoicesType.HIGH) {
                importanceValue = 1;
                importanceText = "High";
            } else if (ewsActions.MarkImportance === ImportanceChoicesType.LOW) {
                importanceValue = 3;
                importanceText = "Low";
            }
            addAction( ` ${PimRuleActionConstants.CHANGE_IMPORTANCE} ${importanceText} `, actionList);
            tokActionList.push(`2¦${importanceValue}¦`);
        }
    }

    if (actionList) {
        ruleStructure.ActionList = actionList;
    }

    if (tokActionList) {
        ruleStructure.TokActionList = tokActionList;
    }
}

function addAction(action: string, actionList: string[]): void {
    const prefix = actionList.length > 0 ? ` ${PimRuleSeparatorConstants.AND} ` : "";
    actionList.push( prefix + action);
}

async function pimRuleToRule(request: Request, pimRule: PimRule): Promise<RuleType | undefined> {
    const rt = new RuleType();
    rt.RuleId = pimRule.unid;
    rt.DisplayName = pimRule.displayName ?? pimRule.unid;
    rt.Priority = pimRule.ruleIndex;
    rt.IsEnabled = pimRule.ruleEnabled;
    
    const conditions = getRuleConditions(pimRule);
    if (conditions) {
        rt.Conditions = conditions;
    }

    const exceptions = getRuleExceptions(pimRule);
    if (exceptions) {
        rt.Exceptions = exceptions;
    }

    const actions = await getRuleActions(request, pimRule);
    if (actions) {
        rt.Actions = actions;
    }

    return Promise.resolve(rt.Conditions && rt.Actions ? rt : undefined);
}

function getRuleExceptions(pimRule: PimRule): RulePredicatesType | undefined {
    return undefined;
}

function getRuleConditions(pimRule: PimRule): RulePredicatesType | undefined {
    return getRulePredicatesType(pimRule.rawRuleCondition);
}

function getRulePredicatesType(predicatePhrase: string, predicatesType?: RulePredicatesType): RulePredicatesType | undefined {
    // Candidate Predicate fields
    // ==============================
    // Categories?: ArrayOfStringsType;
    // ContainsBodyStrings?: ArrayOfStringsType;
    // ContainsHeaderStrings?: ArrayOfStringsType;
    // ContainsRecipientStrings?: ArrayOfStringsType;
    // ContainsSenderStrings?: ArrayOfStringsType;
    // ContainsSubjectOrBodyStrings?: ArrayOfStringsType;
    // ContainsSubjectStrings?: ArrayOfStringsType;
    // FlaggedForAction?: FlaggedForActionType;
    // FromAddresses?: ArrayOfEmailAddressesType;
    // FromConnectedAccounts?: ArrayOfStringsType;
    // HasAttachments?: boolean;
    // Importance?: ImportanceChoicesType;
    // IsApprovalRequest?: boolean;
    // IsAutomaticForward?: boolean;
    // IsAutomaticReply?: boolean;
    // IsEncrypted?: boolean;
    // IsMeetingRequest?: boolean;
    // IsMeetingResponse?: boolean;
    // IsNDR?: boolean;
    // IsPermissionControlled?: boolean;
    // IsReadReceipt?: boolean;
    // IsSigned?: boolean;
    // IsVoicemail?: boolean;
    // ItemClasses?: ArrayOfStringsType;
    // MessageClassifications?: ArrayOfStringsType;
    // NotSentToMe?: boolean;
    // SentCcMe?: boolean;
    // SentOnlyToMe?: boolean;
    // SentToAddresses?: ArrayOfEmailAddressesType;
    // SentToMe?: boolean;
    // SentToOrCcMe?: boolean;
    // Sensitivity?: SensitivityChoicesType;
    // WithinDateRange?: RulePredicateDateRangeType;
    // WithinSizeRange?: RulePredicateSizeRangeType;

    let nextPhrase: string | undefined;
    let operator = PimRuleSeparatorConstants.AND;
    if (predicatePhrase) {
        let phrase = predicatePhrase.trim();

        //It's the beginning....strip off leading and trailing [ and ]
        if (phrase.startsWith('[') && phrase.length > 1) {
            phrase = phrase.substring(1);
            if (phrase.endsWith(']')) {
                phrase = phrase.substring(0, phrase.length - 1).trim();
            }
        }

        // Strip a when
        let token: string = PimRuleSeparatorConstants.WHEN;
        if (phrase.toLowerCase().startsWith(token.toLowerCase())) {
            phrase = phrase.substring(token.length).trim();
        }

        // Strip leading 'AND'
        token = PimRuleSeparatorConstants.AND;
        if (phrase.toLowerCase().startsWith(token.toLowerCase())) {
            phrase = phrase.substring(token.length).trim();
            operator = PimRuleSeparatorConstants.AND;
        }
        
        // Strip leading 'OR'
        token = PimRuleSeparatorConstants.OR;
        if (phrase.toLowerCase().startsWith(token.toLowerCase())) {
            phrase = phrase.substring(token.length).trim();
            operator = PimRuleSeparatorConstants.OR;
        }

        // Look for the next phrase
        token = PimRuleSeparatorConstants.SEPARATOR;
        let tokenIndex = phrase.indexOf(token);
        if (tokenIndex >= 0) {
            nextPhrase = phrase.substring(tokenIndex + token.length).trim();
            phrase = phrase.substring(0, tokenIndex).trim();
        }
        
        // Process Subject
        if (phrase.toLowerCase().startsWith(PimRuleConditionConstants.SUBJECT.toLowerCase())) {
            const remainder = phrase.substring(PimRuleConditionConstants.SUBJECT.length).trim();
            if (remainder.toLowerCase().startsWith(PimRuleOperatorConstants.CONTAINS.toLowerCase())) {
                const containsStrings = remainder.substring(PimRuleOperatorConstants.CONTAINS.length).trim();
                // extract an array of the strings for containment
                if (containsStrings) {
                    predicatesType = predicatesType ?? new RulePredicatesType();
                    if (predicatesType.ContainsSubjectStrings) {
                        if (operator === PimRuleSeparatorConstants.OR) {
                            // We accept multiple conditions of the same type with the or operator, 
                            // not the 'and' operator.
                            predicatesType.ContainsSubjectStrings.String.push(containsStrings);
                        } else {
                            return undefined; // Ignored rule
                        }
                    } else {
                        predicatesType.ContainsSubjectStrings = new ArrayOfStringsType();
                        predicatesType.ContainsSubjectStrings.String = [containsStrings];
                    }
                }
            } else {
                return undefined; // We don't accept "is"....ignore the rule.
            }

        } else if (phrase.toLowerCase().startsWith(PimRuleConditionConstants.BODY_OR_SUBJECT.toLowerCase())) {
            const remainder = phrase.substring(PimRuleConditionConstants.BODY_OR_SUBJECT.length).trim();
            token = PimRuleOperatorConstants.CONTAINS;
            tokenIndex = remainder.toLowerCase().indexOf(token.toLowerCase());
            if (tokenIndex >= 0) {
                const containsStrings = remainder.substring(tokenIndex + token.length).trim();
                if (containsStrings) {
                    predicatesType = predicatesType ?? new RulePredicatesType();
                    if (predicatesType.ContainsSubjectOrBodyStrings) {
                        if (operator === PimRuleSeparatorConstants.OR) {
                            // We accept multiple conditions of the same type with the or operator, 
                            // not the 'and' operator.
                            predicatesType.ContainsSubjectOrBodyStrings.String.push(containsStrings);
                        } else {
                            return undefined; // Ignored rule
                        }
                    } else {
                        predicatesType.ContainsSubjectOrBodyStrings = new ArrayOfStringsType();
                        predicatesType.ContainsSubjectOrBodyStrings.String = [containsStrings];
                    }
                }
            } else {
                return undefined; // We don't accept "is"....ignore the rule.
            }
        } else if (phrase.toLowerCase().startsWith(PimRuleConditionConstants.SENDER.toLowerCase())) {
            const remainder = phrase.substring(PimRuleConditionConstants.SENDER.length).trim();
            if (remainder.toLowerCase().startsWith(PimRuleOperatorConstants.CONTAINS.toLowerCase())) {
                const containsStrings = remainder.substring(PimRuleOperatorConstants.CONTAINS.length).trim();
                if (containsStrings) {
                    predicatesType = predicatesType ?? new RulePredicatesType();
                    if (predicatesType.ContainsSenderStrings) {
                        if (operator === PimRuleSeparatorConstants.OR) {
                            // We accept multiple conditions of the same type with the or operator, 
                            // not the 'and' operator.
                            predicatesType.ContainsSenderStrings.String.push(containsStrings);
                        } else {
                            return undefined; // Ignored rule
                        }
                    } else {
                        predicatesType.ContainsSenderStrings = new ArrayOfStringsType();
                        predicatesType.ContainsSenderStrings.String = [containsStrings];
                    }
                }
            } else if (remainder.toLowerCase().startsWith(PimRuleOperatorConstants.IS.toLowerCase())) {
                const isStrings = remainder.substring(PimRuleOperatorConstants.IS.length).trim();
                if (isStrings) {
                    predicatesType = predicatesType ?? new RulePredicatesType();
                    if (predicatesType.FromAddresses && predicatesType.FromAddresses.Address.length > 0) {
                        if (operator === PimRuleSeparatorConstants.OR) {
                            // We accept multiple conditions of the same type with the or operator, 
                            // not the 'and' operator.
                            const emailAddress = new EmailAddressType();
                            emailAddress.EmailAddress = isStrings;
                            predicatesType.FromAddresses.Address.push(emailAddress);
                        } else {
                            return undefined; // Ignored rule
                        }
                    } else {
                        predicatesType.FromAddresses = new ArrayOfEmailAddressesType();
                        predicatesType.FromAddresses.Address = [];
                        const emailAddress = new EmailAddressType();
                        emailAddress.EmailAddress = isStrings;
                        predicatesType.FromAddresses.Address.push(emailAddress);
                    }
                }
            } else {
                return undefined; // We don't accept "is"....ignore the rule.
            }
        } else {
            return undefined;
        }
    }

    // let ewsSensitivity: SensitivityChoicesType;
    // switch (pimRule.sensitivity) {
    //     case "Confidential":
    //         ewsSensitivity = SensitivityChoicesType.CONFIDENTIAL;
    //         break;
    //     case "Personal":
    //         ewsSensitivity = SensitivityChoicesType.PERSONAL;
    //         break;
    //     default:
    //         ewsSensitivity = SensitivityChoicesType.NORMAL;
    // }
    // predicateType.Sensitivity = ewsSensitivity;  //SensitivityChoicesType



    // predicateType.FlaggedForAction = FlaggedForActionType.FORWARD;
    // predicateType.FromAddresses = new ArrayOfEmailAddressesType();
    // const addresses: EmailAddressType[] = [];
    // let emailAddress = new EmailAddressType();
    // emailAddress.Name = 'Tom';
    // emailAddress.EmailAddress = 'tom@tc.com';
    // addresses.push(emailAddress);

    // emailAddress = new EmailAddressType();
    // emailAddress.Name = 'Mike';
    // emailAddress.EmailAddress = 'mike@tc.com';
    // addresses.push(emailAddress);

    // predicateType.FromAddresses.Address = addresses;
    // predicateType.ContainsSubjectStrings = {
    //         String: ['Interesting']
    //     };
    // rt.Conditions = predicateType;

    if (nextPhrase) {
        predicatesType = getRulePredicatesType(nextPhrase, predicatesType);
    }
    return predicatesType;
}

async function getRuleActions(request: Request, pimRule: PimRule): Promise<RuleActionsType | undefined> {
    return getRuleActionsType(request, pimRule.rawRuleAction);
}
    
async function getRuleActionsType(request: Request, ruleActionsPhrase: string, actionsType?: RuleActionsType | undefined): Promise<RuleActionsType | undefined> {
    const userInfo = UserContext.getUserInfo(request);

    // Candidate Action fields
    // ==============================
    // AssignCategories?: ArrayOfStringsType;
    // CopyToFolder?: TargetFolderIdType;
    // Delete?: boolean;
    // ForwardAsAttachmentToRecipients?: ArrayOfEmailAddressesType;
    // ForwardToRecipients?: ArrayOfEmailAddressesType;
    // MarkImportance?: ImportanceChoicesType;
    // MarkAsRead?: boolean;
    // MoveToFolder?: TargetFolderIdType;
    // PermanentDelete?: boolean;
    // RedirectToRecipients?: ArrayOfEmailAddressesType;
    // SendSMSAlertToRecipients?: ArrayOfEmailAddressesType;
    // ServerReplyWithMessage?: ItemIdType;
    // StopProcessingRules?: boolean;

    let nextPhrase: string | undefined;
    if (ruleActionsPhrase) {
        let phrase = ruleActionsPhrase.trim();

        //It's the beginning....strip off leading and trailing [ and ]
        if (phrase.startsWith('[') && phrase.length > 1) {
            phrase = phrase.substring(1);
            if (phrase.endsWith(']')) {
                phrase = phrase.substring(0, phrase.length - 1).trim();
            }
        }

        // Strip THEN
        let token: string = PimRuleSeparatorConstants.THEN;
        if (phrase.toLowerCase().startsWith(token.toLowerCase())) {
            phrase = phrase.substring(token.length).trim();
        }
    
        // Strip leading 'AND'
        token = PimRuleSeparatorConstants.AND;
        if (phrase.toLowerCase().startsWith(token.toLowerCase())) {
            phrase = phrase.substring(token.length).trim();
        }

        // Look for the next phrase
        token = PimRuleSeparatorConstants.SEPARATOR;
        const tokenIndex = phrase.indexOf(token);
        if (tokenIndex >= 0) {
            nextPhrase = phrase.substring(tokenIndex + token.length).trim();
            phrase = phrase.substring(0, tokenIndex).trim();
        }

        if (phrase.toLowerCase().startsWith(PimRuleActionConstants.MOVE_TO_FOLDER.toLowerCase())) {
            const remainder = phrase.substring(PimRuleActionConstants.MOVE_TO_FOLDER.length).trim();
            // Should be the view to move to.....but we need the folderid given the view
            const currentPimLabels = await KeepPimLabelManager.getInstance().getLabels(userInfo);
            const pimLabel = currentPimLabels.find(pLabel => pLabel.view === remainder);
            if (pimLabel) {
                actionsType = actionsType ?? new RuleActionsType();
                actionsType.MoveToFolder = new TargetFolderIdType();
                const folderEWSId = getEWSId(pimLabel.unid);
                actionsType.MoveToFolder.FolderId = new FolderIdType(folderEWSId, `ck-${folderEWSId}`);
            }
        } else if (phrase.toLowerCase().startsWith(PimRuleActionConstants.COPY_TO_FOLDER.toLowerCase())) {
            const remainder = phrase.substring(PimRuleActionConstants.COPY_TO_FOLDER.length).trim();
            // Should be the view to copy to.....but we need the folderid given the view
            const currentPimLabels = await KeepPimLabelManager.getInstance().getLabels(userInfo);
            const pimLabel = currentPimLabels.find(pLabel => pLabel.view === remainder);
            if (pimLabel) {
                actionsType = actionsType ?? new RuleActionsType();
                actionsType.CopyToFolder = new TargetFolderIdType();
                const folderEWSId = getEWSId(pimLabel.unid);
                actionsType.CopyToFolder.FolderId = new FolderIdType(folderEWSId, `ck-${folderEWSId}`);
            }
        } else if (phrase.toLowerCase().startsWith(PimRuleActionConstants.DELETE.toLowerCase())) {
            actionsType = actionsType ?? new RuleActionsType();
            actionsType.Delete = true;
        } else if (phrase.toLowerCase().startsWith(PimRuleActionConstants.STOP_PROCESSING.toLowerCase())) {
            actionsType = actionsType ?? new RuleActionsType();
            actionsType.StopProcessingRules = true;
        } else if (phrase.toLowerCase().startsWith(PimRuleActionConstants.CHANGE_IMPORTANCE.toLowerCase())) {
            const remainder = phrase.substring(PimRuleActionConstants.CHANGE_IMPORTANCE.length).trim();
            let newImportance: ImportanceChoicesType | undefined;
            switch (remainder.toLowerCase()) {
                case "high": {
                    newImportance = ImportanceChoicesType.HIGH;
                    break;
                }
                case "normal": {
                    newImportance = ImportanceChoicesType.NORMAL;
                    break;
                }
                case "low": {
                    newImportance = ImportanceChoicesType.LOW;
                    break;
                }
            }
            if (newImportance) {
                actionsType = actionsType ?? new RuleActionsType();
                actionsType.MarkImportance = newImportance;
            }
        } else if (phrase.toLowerCase().startsWith(PimRuleActionConstants.SEND_FULL_COPY.toLowerCase())) {
            const remainder = phrase.substring(PimRuleActionConstants.SEND_FULL_COPY.length).trim();
            actionsType = actionsType ?? new RuleActionsType();
            nextPhrase = processSendCopyRecipients(remainder, nextPhrase, actionsType);
        } else if (phrase.toLowerCase().startsWith(PimRuleActionConstants.SEND_HEADERS_COPY.toLowerCase())) {
            const remainder = phrase.substring(PimRuleActionConstants.SEND_HEADERS_COPY.length).trim();
            actionsType = actionsType ?? new RuleActionsType();
            nextPhrase = processSendCopyRecipients(remainder, nextPhrase, actionsType);
        } else {
            // If it got here and there is a phrase, we didn't understand the rule and should exit out
            return undefined;
        }
    }
    
    if (nextPhrase) {
        actionsType = await getRuleActionsType(request, nextPhrase, actionsType);
    }
    return Promise.resolve(actionsType);
}

function processSendCopyRecipients(recipient: string, nextPhrase: string | undefined, actionsType: RuleActionsType): string | undefined {
    if (recipient.length > 0) {
        const recips = new ArrayOfEmailAddressesType();
        const addresses: EmailAddressType[] = [];
        let emailAddress = new EmailAddressType();
        emailAddress.EmailAddress = recipient;
        addresses.push(emailAddress);            
        recips.Address = addresses;
        actionsType.ForwardToRecipients = recips;

        // Check if the next phrase is really the beginning of a new action or another token for this action
        while (nextPhrase && nextPhrase.length > 0 && !isNewActionPhrase(nextPhrase)) {
            // Assume a token to be applied to the email addresses
            // Look for the next phrase
            const token = PimRuleSeparatorConstants.SEPARATOR;
            const tokenIndex = nextPhrase.indexOf(token);
            if (tokenIndex >= 0) {
                recipient = nextPhrase.substring(0, tokenIndex).trim();
                nextPhrase = nextPhrase.substring(tokenIndex + token.length).trim();
            } else {
                recipient = nextPhrase;
                nextPhrase = undefined;
            }
            if (recipient) {
                emailAddress = new EmailAddressType();
                emailAddress.EmailAddress = recipient;
                addresses.push(emailAddress);            
            }
        }
    }
    return nextPhrase;
}

function isNewActionPhrase(ruleActionsPhrase: string): boolean {
    let phrase = ruleActionsPhrase.trim();

    // Look for the next phrase
    let token = PimRuleSeparatorConstants.SEPARATOR;
    const tokenIndex = phrase.indexOf(token);
    if (tokenIndex >= 0) {
        phrase = phrase.substring(0, tokenIndex).trim();
    }
    
    // Strip THEN
    token = PimRuleSeparatorConstants.THEN;
    if (phrase.toLowerCase().startsWith(token.toLowerCase())) {
        phrase = phrase.substring(token.length).trim();
    }

    // Strip leading 'AND'
    token = PimRuleSeparatorConstants.AND;
    if (phrase.toLowerCase().startsWith(token.toLowerCase())) {
        phrase = phrase.substring(token.length).trim();
    }

    for (const action of Object.values(PimRuleActionConstants)) {
        if (phrase.toLowerCase().startsWith(action.toLowerCase())) {
            return true;
        }
    }
    return false;
}

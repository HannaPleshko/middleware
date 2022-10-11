/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import { ResponseMessageType, BaseRequestType, GenericArray, BaseResponseMessageType, PathChoiceType, NonEmptyArrayOfFieldOrdersType } from './common.model';
import {
    BodyTypeType, LocationSourceType, DefaultShapeNamesType, IndexBasePointType,
    ContainmentModeType, ContainmentComparisonType, ResolveNamesSearchScopeType
} from './enum.model';
import {
    NonEmptyArrayOfPathsToElementType, NonEmptyArrayOfBaseFolderIdsType,
    FolderIdType, ArrayOfFolderIdType, TargetFolderIdType, ExtendedPropertyType,
    EmailAddressType, ArrayOfEmailAddressesType, ItemIdType, ContactItemType,
} from './mail.model';

/**
 * Defines a response to a ResolveNames request.
 */
export class ResolveNamesResponseType extends BaseResponseMessageType<ResolveNamesResponseMessageType> {
}

/**
 * Contains the status and result of a ResolveNames operation request.
 */
export class ResolveNamesResponseMessageType extends ResponseMessageType {
    ResolutionSet: ArrayOfResolutionType;
}

/**
 * Contains an array of resolutions for an ambiguous name.
 */
export class ArrayOfResolutionType {
    Resolution: ResolutionType[] = [];
    IndexedPagingOffset?: number;
    NumeratorOffset?: number;
    AbsoluteDenominator?: number;
    IncludesLastItemInRange?: boolean;
    TotalItemsInView?: number;
}

/**
 * Contains a single resolved entity.
 */
export class ResolutionType {
    Mailbox: EmailAddressType;
    Contact?: ContactItemType;
}

/**
 * Contains the response to a FindPeople request.
 */
export class FindPeopleResponseMessageType extends ResponseMessageType {
    People?: ArrayOfPeopleType;
    TotalNumberOfPeopleInView?: number;
    FirstMatchingRowIndex?: number;
    FirstLoadedRowIndex?: number;
    TransactionId?: string;
}

/**
 * Specifies an array of persona data returned as the result of a FindPeople request.
 */
export class ArrayOfPeopleType {
    Persona?: PersonaType[];
}

/**
 * Contains the response data resulting from a GetPersona request.
 */
export class GetPersonaResponseMessageType extends ResponseMessageType {
    Persona: PersonaType;
}

/**
 * Specifies a set of persona data returned by a FindPeople/GetPersona request.
 */
export class PersonaType {
    PersonaId: ItemIdType;
    PersonaType?: string;
    PersonaObjectStatus?: string;
    CreationTime?: Date;
    Bodies?: ArrayOfBodyContentAttributedValuesType;
    DisplayNameFirstLastSortKey?: string;
    DisplayNameLastFirstSortKey?: string;
    CompanyNameSortKey?: string;
    HomeCitySortKey?: string;
    WorkCitySortKey?: string;
    DisplayNameFirstLastHeader?: string;
    DisplayNameLastFirstHeader?: string;
    DisplayName?: string;
    DisplayNameFirstLast?: string;
    DisplayNameLastFirst?: string;
    FileAs?: string;
    FileAsId?: string;
    DisplayNamePrefix?: string;
    GivenName?: string;
    MiddleName?: string;
    Surname?: string;
    Generation?: string;
    Nickname?: string;
    YomiCompanyName?: string;
    YomiFirstName?: string;
    YomiLastName?: string;
    Yitle?: string;
    Department?: string;
    CompanyName?: string;
    Location?: string;
    EmailAddress?: EmailAddressType;
    EmailAddresses?: ArrayOfEmailAddressesType;
    PhoneNumber?: PersonaPhoneNumberType;
    ImAddress?: string;
    HomeCity?: string;
    WorkCity?: string;
    RelevanceScore?: number;
    FolderIds?: ArrayOfFolderIdType;
    Attributions?: ArrayOfPersonaAttributionsType;
    DisplayNames?: ArrayOfStringAttributedValuesType;
    FileAses?: ArrayOfStringAttributedValuesType;
    FileAsIds?: ArrayOfStringAttributedValuesType;
    DisplayNamePrefixes?: ArrayOfStringAttributedValuesType;
    GivenNames?: ArrayOfStringAttributedValuesType;
    MiddleNames?: ArrayOfStringAttributedValuesType;
    Surnames?: ArrayOfStringAttributedValuesType;
    Generations?: ArrayOfStringAttributedValuesType;
    Nicknames?: ArrayOfStringAttributedValuesType;
    Initials?: ArrayOfStringAttributedValuesType;
    YomiCompanyNames?: ArrayOfStringAttributedValuesType;
    YomiFirstNames?: ArrayOfStringAttributedValuesType;
    YomiLastNames?: ArrayOfStringAttributedValuesType;
    BusinessPhoneNumbers?: ArrayOfPhoneNumberAttributedValuesType;
    BusinessPhoneNumbers2?: ArrayOfPhoneNumberAttributedValuesType;
    HomePhones?: ArrayOfPhoneNumberAttributedValuesType;
    HomePhones2?: ArrayOfPhoneNumberAttributedValuesType;
    MobilePhones?: ArrayOfPhoneNumberAttributedValuesType;
    MobilePhones2?: ArrayOfPhoneNumberAttributedValuesType;
    AssistantPhoneNumbers?: ArrayOfPhoneNumberAttributedValuesType;
    CallbackPhones?: ArrayOfPhoneNumberAttributedValuesType;
    CarPhones?: ArrayOfPhoneNumberAttributedValuesType;
    HomeFaxes?: ArrayOfPhoneNumberAttributedValuesType;
    OrganizationMainPhones?: ArrayOfPhoneNumberAttributedValuesType;
    OtherFaxes?: ArrayOfPhoneNumberAttributedValuesType;
    OtherTelephones?: ArrayOfPhoneNumberAttributedValuesType;
    OtherPhones2?: ArrayOfPhoneNumberAttributedValuesType;
    Pagers?: ArrayOfPhoneNumberAttributedValuesType;
    RadioPhones?: ArrayOfPhoneNumberAttributedValuesType;
    TelexNumbers?: ArrayOfPhoneNumberAttributedValuesType;
    WorkFaxes?: ArrayOfPhoneNumberAttributedValuesType;
    Emails1?: ArrayOfEmailAddressAttributedValuesType;
    Emails2?: ArrayOfEmailAddressAttributedValuesType;
    Emails3?: ArrayOfEmailAddressAttributedValuesType;
    BusinessHomePages?: ArrayOfStringAttributedValuesType;
    PersonalHomePages?: ArrayOfStringAttributedValuesType;
    OfficeLocations?: ArrayOfStringAttributedValuesType;
    ImAddresses?: ArrayOfStringAttributedValuesType;
    ImAddresses2?: ArrayOfStringAttributedValuesType;
    ImAddresses3?: ArrayOfStringAttributedValuesType;
    BusinessAddresses?: ArrayOfPostalAddressAttributedValuesType;
    HomeAddresses?: ArrayOfPostalAddressAttributedValuesType;
    OtherAddresses?: ArrayOfPostalAddressAttributedValuesType;
    Title?: string;
    Titles?: ArrayOfStringAttributedValuesType;
    Departments?: ArrayOfStringAttributedValuesType;
    CompanyNames?: ArrayOfStringAttributedValuesType;
    Managers?: ArrayOfStringAttributedValuesType;
    AssistantNames?: ArrayOfStringAttributedValuesType;
    Professions?: ArrayOfStringAttributedValuesType;
    SpouseNames?: ArrayOfStringAttributedValuesType;
    Children?: ArrayOfStringArrayAttributedValuesType;
    Schools?: ArrayOfStringAttributedValuesType;
    Hobbies?: ArrayOfStringAttributedValuesType;
    WeddingAnniversaries?: ArrayOfStringAttributedValuesType;
    Birthdays?: ArrayOfStringAttributedValuesType;
    Locations?: ArrayOfStringAttributedValuesType;
    InlineLinks?: ArrayOfStringAttributedValuesType;
    ItemLinkIds?: ArrayOfStringArrayAttributedValuesType;
    HasActiveDeals?: string;
    IsBusinessContact?: string;
    AttributedHasActiveDeals?: ArrayOfStringAttributedValuesType;
    AttributedIsBusinessContact?: ArrayOfStringAttributedValuesType;
    SourceMailboxGuids?: ArrayOfStringAttributedValuesType;
    LastContactedDate?: Date;
    ExtendedProperties?: ArrayOfExtendedPropertyAttributedValueType;
    ExternalDirectoryObjectId?: string;
    MapiEntryId?: string;
    MapiEmailAddress?: string;
    MapiAddressType?: string;
    MapiSearchKey?: string;
    MapiTransmittableDisplayName?: string;
    MapiSendRichInfo?: boolean;
    TTYTDDPhoneNumbers?: ArrayOfPhoneNumberAttributedValuesType;
}

/**
 * Specifies an array of body contents.
 */
export class ArrayOfBodyContentAttributedValuesType {
    BodyContentAttributedValue?: BodyContentAttributedValueType[];
}

/**
 * Specifies the body content of an item.
 */
export class BodyContentAttributedValueType {
    Value: BodyContentType;
    Attributions: ArrayOfValueAttributionsType;
}

/**
 * Represents persona phone number.
 */
export class PersonaPhoneNumberType {
    Number: string;
    Type: string;
}

/**
 * Represents an array of persona attributions.
 */
export class ArrayOfPersonaAttributionsType {
    Attribution: PersonaAttributionType[];
}

/**
 * Represents an array of strings with attributed values.
 */
export class ArrayOfStringAttributedValuesType {
    StringAttributedValue?: StringAttributedValueType[];
}

/**
 * Represents an array of string arrays with attributed values.
 */
export class ArrayOfStringArrayAttributedValuesType {
    StringArrayAttributedValue?: StringArrayAttributedValueType[];
}

/**
 * Represents an array of phone numbers with attributed values.
 */
export class ArrayOfPhoneNumberAttributedValuesType {
    PhoneNumberAttributedValue?: PhoneNumberAttributedValueType[];
}

/**
 * Represents an array of email addresses with attributed values.
 */
export class ArrayOfEmailAddressAttributedValuesType {
    EmailAddressAttributedValue?: EmailAddressAttributedValueType[];
}

/**
 * Represents an array of postal addresses with attributed values.
 */
export class ArrayOfPostalAddressAttributedValuesType {
    PostalAddressAttributedValue?: PostalAddressAttributedValueType[];
}

/**
 * Represents an array of extended properties with attributed values.
 */
export class ArrayOfExtendedPropertyAttributedValueType {
    ExtendedPropertyAttributedValue?: ExtendedPropertyAttributedValueType[];
}

/**
 * Represents persona attribution.
 */
export class PersonaAttributionType {
    Id: string;
    SourceId: ItemIdType;
    DisplayName: string;
    IsWritable?: boolean;
    IsQuickContact?: boolean;
    IsHidden?: boolean;
    FolderId?: FolderIdType;
}

/**
 * Represents string value with attributions.
 */
export class StringAttributedValueType {
    Value: string;
    Attributions: ArrayOfValueAttributionsType;
}

/**
 * Represents string array value with attributions.
 */
export class StringArrayAttributedValueType {
    Values: ArrayOfStringValueType;
    Attributions: ArrayOfValueAttributionsType;
}

/**
 * Represents phone number with attributions.
 */
export class PhoneNumberAttributedValueType {
    Value: PersonaPhoneNumberType;
    Attributions: ArrayOfValueAttributionsType;
}

/**
 * Represents email address with attributions.
 */
export class EmailAddressAttributedValueType {
    Value: EmailAddressType;
    Attributions: ArrayOfValueAttributionsType;
}

/**
 * Represents postal address with attributions.
 */
export class PostalAddressAttributedValueType {
    Value: PersonaPostalAddressType;
    Attributions: ArrayOfValueAttributionsType;
}

/**
 * Represents extended property with attributions.
 */
export class ExtendedPropertyAttributedValueType {
    Value: ExtendedPropertyType;
    Attributions: ArrayOfValueAttributionsType;
}

/**
 * Represents body content type.
 */
export class BodyContentType {
    Value: string;
    BodyType: BodyTypeType;
}

/**
 * Represents an array of string vlaues.
 */
export class ArrayOfStringValueType {
    Value: string[];
}

/**
 * Represents an array of attributions.
 */
export class ArrayOfValueAttributionsType {
    Attribution: string[];
}

/**
 * Represents a persona postal address.
 */
export class PersonaPostalAddressType {
    Street?: string;
    City?: string;
    State?: string;
    Country?: string;
    PostalCode?: string;
    PostOfficeBox?: string;
    Type?: string;
    Latitude?: number;
    Longitude?: number;
    Accuracy?: number;
    Altitude?: number;
    AltitudeAccuracy?: number;
    FormattedAddress?: string;
    LocationUri?: string;
    LocationSource?: LocationSourceType;
}

////// Request Models

/**
 * Contains the request to resolve name.
 */
export class ResolveNamesType extends BaseRequestType {
    ParentFolderIds?: NonEmptyArrayOfBaseFolderIdsType;
    UnresolvedEntry: string;
    ReturnFullContactData: boolean;
    SearchScope: ResolveNamesSearchScopeType;
    ContactDataShape: DefaultShapeNamesType;
}

/**
 * Contains the request to get a persona.
 */
export class GetPersonaType extends BaseRequestType {
    PersonaId?: ItemIdType;
    EmailAddress?: EmailAddressType;
    ParentFolderId?: TargetFolderIdType;
    ItemLinkId?: string;
    AdditionalProperties?: NonEmptyArrayOfPathsToElementType;
}

/**
 * Contains the request to find people.
 */
export class FindPeopleType extends BaseRequestType {
    PersonaShape?: PersonaResponseShapeType;
    IndexedPageItemView: IndexedPageViewType;
    Restriction?: RestrictionType;
    AggregationRestriction?: RestrictionType;
    SortOrder?: NonEmptyArrayOfFieldOrdersType;
    ParentFolderId?: TargetFolderIdType;
    QueryString?: string;
    SearchPeopleSuggestionIndex?: boolean;
    TopicQueryString?: string;
    Context?: ArrayOfContextProperty;
    QuerySources: ArrayOfPeopleQuerySource;
    ReturnFlattenedResults?: boolean;
}

/**
 * Represents an array of context properties.
 */
export class ArrayOfContextProperty {
    ContextProperty: ContextPropertyType[];
}

/**
 * Represents a context property.
 */
export class ContextPropertyType {
    Key: string;
    Value: string;
}

/**
 * Represents an array of people query sources.
 */
export class ArrayOfPeopleQuerySource {
    Source: string[];
}

/**
 * Specifies the set of persona properties to be returned from a FindPeople request.
 */
export class PersonaResponseShapeType {
    BaseShape: DefaultShapeNamesType;
    AdditionalProperties?: NonEmptyArrayOfPathsToElementType;
}

/**
 * Base paging type.
 */
export abstract class BasePagingType {
    MaxEntriesReturned?: number;
}

/**
 * Describes how paged conversation or item information is returned.
 */
export class IndexedPageViewType extends BasePagingType {
    Offset: number;
    BasePoint: IndexBasePointType;
}

/**
 * Constant value used in search expression.
 */
export class ConstantValueType {
    Value: string;
}

/**
 * Excludes value used in search expression.
 */
export class ExcludesValueType {
    Value: string;
}

/**
 * Represents the substituted element within a restriction. This element is not used in an XML instance document.
 */
export interface SearchExpressionType { }

/**
 * Field URI or constant type.
 */
export class FieldURIOrConstantType extends PathChoiceType {
    Constant: ConstantValueType;
}

/**
 * Represents a search expression that returns true if the supplied property exists on an item.
 */
export class ExistsType extends PathChoiceType implements SearchExpressionType { }

/**
 * Performs a bitwise mask of the properties.
 */
export class ExcludesType extends PathChoiceType implements SearchExpressionType {
    Bitmask: ExcludesValueType;
}

/**
 * Represents a search expression that determines whether a given property contains the supplied constant string value.
 */
export class ContainsExpressionType extends PathChoiceType implements SearchExpressionType {
    Constant: ConstantValueType;
    ContainmentMode: ContainmentModeType;
    ContainmentComparison: ContainmentComparisonType;
}

/**
 * Base search expression type to perform string comparisons.
 */
export abstract class TwoOperandExpressionType extends PathChoiceType implements SearchExpressionType {
    FieldURIOrConstant: FieldURIOrConstantType;
}
/**
 * Represents a search expression that compares a property with either a constant value or another property
 * and evaluates to true if they are equal.
 */
export class IsEqualToType extends TwoOperandExpressionType { }

/**
 * Represents a search expression that compares a property with either a constant value or another property
 * and returns true if the values are not the same.
 */
export class IsNotEqualToType extends TwoOperandExpressionType { }

/**
 * Represents a search expression that compares a property with either a constant value or another property
 * and returns true if the first property is greater than the value or property.
 */
export class IsGreaterThanType extends TwoOperandExpressionType { }

/**
 * Represents a search expression that compares a property with either a constant value or another property
 * and returns true if the first property is greater than or equal to the value or property.
 */
export class IsGreaterThanOrEqualToType extends TwoOperandExpressionType { }

/**
 * Represents a search expression that compares a property with either a constant value or another property
 * and returns true if the first property is less than the value or property.
 */
export class IsLessThanType extends TwoOperandExpressionType { }

/**
 * Represents a search expression that compares a property with either a constant value or another property
 * and returns true if the first property is less than or equal to the value or property.
 */
export class IsLessThanOrEqualToType extends TwoOperandExpressionType { }

/**
 * Represents a search expression that performs a logical OR operation on the search expression it contains.
 * The Or element will return true if any of its children return true.
 */
export class OrType extends GenericArray<SearchExpressionType> { }

/**
 * Represents a search expression that enables you to perform a Boolean AND operation between two or more search expressions.
 */
export class AndType extends GenericArray<SearchExpressionType> { }

/**
 * Represents a search expression that enables you to perform a near search with distance.
 */
export class NearType extends GenericArray<SearchExpressionType> {
    Distance: bigint;
    Ordered: boolean;
}

/**
 * Search expression choice.
 */
export abstract class SearchExpressionChoiceType {
    Not?: NotType;
    Exists?: ExistsType;
    Excludes?: ExcludesType;
    ContainsExpression?: ContainsExpressionType;
    IsEqualTo?: IsEqualToType;
    IsNotEqualTo?: IsNotEqualToType;
    IsGreaterThan?: IsGreaterThanType;
    IsGreaterThanOrEqualTo?: IsGreaterThanOrEqualToType;
    IsLessThan?: IsLessThanType;
    IsLessThanOrEqualTo?: IsLessThanOrEqualToType;
    Or?: OrType;
    And?: AndType;
    Near?: NearType;
}

/**
 * Represents a search expression that negates the Boolean value of the search expression it contains.
 */
export class NotType extends SearchExpressionChoiceType implements SearchExpressionType { }

/**
 * Represents the restriction or query that is used to filter items or folders.
 */
export class RestrictionType extends SearchExpressionChoiceType { }

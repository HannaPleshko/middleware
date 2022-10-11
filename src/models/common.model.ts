/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import { ResponseClassType, ResponseCodeType, DayOfWeekType, 
    UnindexedFieldURIType, DictionaryURIType, DistinguishedPropertySetType,
    MapiPropertyTypeType, ExceptionPropertyURIType, SortDirectionType } from './enum.model';

export const AttributesKey = '$attributes';
export const ValueKey = '$value';
export const XSITypeKey = '$xsiType';

/**
 * The server version.
 */
export class ServerVersionInfo {
    MajorVersion = 15;
    MinorVersion = 20;
    MajorBuildNumber = 2835;
    MinorBuildNumber = 22;
    Version = 'V2018_01_08';
}

/**
 * Within xs:choice, there could be elements with arbitary element names while refer to same type.
 * In such case there could be conflict and we need a way to specify element names.
 */
export interface NamedElement {
    readonly $qname: string;
}

/**
 * An array containing items which have dynamic sub-class types.
 */
export class GenericArray<T> {
    constructor(public items: T[] = []) { }
    push(item: T): number {
        if (!this.items) {
            this.items = [];
        }
        return this.items.push(item);
    }
}

/**
 * Represents an xs:list array, which will be rendered by join them via space, and parsed by splitting space.
 * E.g. A xs:list array of string ['value1', 'value2'] correponds to one string 'value1 value2' in xml.
 */
export abstract class XSList<T> extends Array {
    constructor(items?: T[]) {
        super();
        if (items) {
            this.push(...items);
        }
    }
    push(...items: T[]): number {
        if (items) {
            items.forEach(item => {
                if (item) {
                    const split = (item + '').split(' ');
                    super.push(...split);
                }
            })
        }
        return this.length;
    }
    get [ValueKey](): string {
        return this.join(' ');
    }
}

export class DaysOfWeekType extends XSList<DayOfWeekType> { }

/**
 * Array of string values.
 */
export class ArrayOfStringsType {
    String: string[] = [];
}

/**
 * Represents SOAP fault error.
 */
export class SOAPFault extends Error {
    Fault: { faultcode: ResponseCodeType; faultstring: string; statusCode: number }
    constructor(faultcode: ResponseCodeType, faultstring: string, statusCode = 500) {
        super(faultstring);
        this.Fault = {
            faultcode,
            faultstring,
            statusCode,
        }
    }
}

/**
 * Base request type.
 */
export abstract class BaseRequestType {
    SOAPHeaders: any;
}

/**
 * Simple value type with attribute.
 */
export abstract class AttributeSimpleType<T> {
    constructor(public Value: T) {
    }
    set [ValueKey](_value: T) {
        this.Value = _value;
    }
    get [ValueKey](): T {
        return this.Value;
    }
}

/**
 * Provides additional error response information.
 */
export class Value extends AttributeSimpleType<string> {
    constructor(public Name: string, _value: string) {
        super(_value);
    }
}

/**
 * Provides descriptive information about the response status for a single entity within a request.
 * By default it sets to success status.
 */
export class ResponseMessageType {
    ResponseClass: ResponseClassType = ResponseClassType.SUCCESS;
    ResponseCode: ResponseCodeType = ResponseCodeType.NO_ERROR;
    MessageText?: string;
    DescriptiveLinkKey?: number;
    MessageXml?: any;
}

/**
 * Base response messages type.
 */
export abstract class BaseResponseMessageType<T> {
    ResponseMessages = new GenericArray<T>();
}

/**
 * Defines how items are sorted.
 */
export class NonEmptyArrayOfFieldOrdersType {
    FieldOrder: FieldOrderType[];
}

/**
 * Base path to element type.
 */
export abstract class BasePathToElementType { }

/**
 * Identifies frequently referenced properties by URI.
 */
export class PathToUnindexedFieldType extends BasePathToElementType {
    FieldURI: UnindexedFieldURIType;
}

/**
 * Identifies frequently referenced dictionary properties by URI.
 */
export class PathToIndexedFieldType extends BasePathToElementType {
    FieldURI: DictionaryURIType;
    FieldIndex: string;
}

/**
 * Identifies extended MAPI properties to get, set, or create.
 */
export class PathToExtendedFieldType extends BasePathToElementType {
    DistinguishedPropertySetId?: DistinguishedPropertySetType;
    PropertySetId?: string;
    PropertyTag?: string;
    PropertyName?: string;
    PropertyId?: number;
    PropertyType: MapiPropertyTypeType;
}

/**
 * Specifies an offending property path in an error
 */
export class PathToExceptionFieldType extends BasePathToElementType {
    FieldURI: ExceptionPropertyURIType;
}

/**
 * Represents path choice. Only one of FieldURI/IndexedFieldURI/ExtendedFieldURI could be present.
 */
export abstract class PathChoiceType {
    FieldURI?: PathToUnindexedFieldType;
    IndexedFieldURI?: PathToIndexedFieldType;
    ExtendedFieldURI?: PathToExtendedFieldType;
}

/**
 * Represents a single field by which to sort results and indicates the direction for the sort.
 * One or more of these elements may be included. FieldOrder elements are applied in the order
 * specified for sorting.
 */
export class FieldOrderType extends PathChoiceType {
    Order: SortDirectionType;
}



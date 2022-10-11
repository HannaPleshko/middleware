/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import isClass from 'is-class';
import * as commonModel from './common.model';
import * as userModel from './user.model';
import * as mailModel from './mail.model';
import * as personaModel from './persona.model';
import * as clientModel from './client.model';
import * as inboxModel from './inbox.model';
import * as autodiscoverModel from './autodiscover.model';
import { SOAP, ExchangeSoap } from '../soap';
import { Logger } from '../utils';

/**
 * Factory to create model object. This is used when parsing incoming soap request.
 * After soap xml parsed, the generated JS object will have correct model type.
 *
 * @param type The type name of model object to create.
 *             It will be same as defined in XSD schema (wsdl/messages.xsd and wsdl/types.xsd).
 * @param models The models
 * @returns model created, or an empty object if not recognized.
 */
function createObject(type: string, models: any[]): any {
    for (const model of models) {
        const modelType = model[type];
        if (modelType && isClass(modelType)) {
            return new modelType();
        }
    }
    Logger.getInstance().warn('Type not recognized: ' + type);
    return {};
}

/**
 * Create object factory.
 * @param soap The SOAP definition
 * @returns Factory function to create object.
 */
export function objectFactory(soap: SOAP): Function {
    const models = soap === ExchangeSoap ? [commonModel, userModel, mailModel, personaModel, clientModel, inboxModel] : [autodiscoverModel];
    return (type: string): any => createObject(type, models);
}

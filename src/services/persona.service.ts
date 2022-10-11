/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import assert from 'assert';
import { BindingScope, bind } from '@loopback/core';

import { ExchangeSoap } from '../soap';
import {
    FindPeopleType, FindPeopleResponseMessageType, GetPersonaType, GetPersonaResponseMessageType,
    ResolveNamesType, ResolveNamesResponseType
} from '../models/persona.model';
import { mockFindPeopleResponse, mockGetPersonaResponse, ResolveNames } from '../data/persona';
import { Request } from '@loopback/rest';

/**
 * The persona service.
 */
@bind({ scope: BindingScope.SINGLETON, tags: ExchangeSoap.service })
export class PersonaService {

    /**
     * Find people.
     * @param soapRequest The soap request, which is type of FindPeopleType
     * @returns Result of people found
     */
    async FindPeople(soapRequest: FindPeopleType): Promise<FindPeopleResponseMessageType> {
        assert(soapRequest instanceof FindPeopleType);
        return mockFindPeopleResponse;
    }

    /**
     * Get persona.
     * @param soapRequest The soap request, which is type of GetPersonaType
     * @returns Persona retrieved
     */
    async GetPersona(soapRequest: GetPersonaType): Promise<GetPersonaResponseMessageType> {
        assert(soapRequest instanceof GetPersonaType);
        return mockGetPersonaResponse;
    }

    /**
     * Resolve name.
     * @param soapRequest The soap request, which is type of ResolveNamesType
     * @param request The HTTP request
     * @returns resolve result
     */
    async ResolveNames(soapRequest: ResolveNamesType, request: Request): Promise<ResolveNamesResponseType> {
        assert(soapRequest instanceof ResolveNamesType);

        return ResolveNames(soapRequest, request);
    }
}
/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import assert from 'assert';
import { bind, BindingScope } from '@loopback/core';

import { ExchangeSoap } from '../soap';
import {
    GetUserPhotoType, GetUserPhotoResponseMessageType,
    SetUserPhotoType, SetUserPhotoResponseMessageType,
    GetUserOofSettingsRequest, GetUserOofSettingsResponseMessageType,
    SetUserOofSettingsRequest, SetUserOofSettingsResponseMessageType,
    GetUserAvailabilityRequestType, GetUserAvailabilityResponseMessageType,
    GetUserConfigurationResponseType, GetUserConfigurationType,
    CreateUserConfigurationResponseType, CreateUserConfigurationType,
    UpdateUserConfigurationResponseType, UpdateUserConfigurationType,
    DeleteUserConfigurationResponseType, DeleteUserConfigurationType
} from '../models/user.model';

import {
    mockSetUserPhotoResponse,
    SetUserOofSettings, GetUserOofSettings,
    mockGetUserAvailabilityResponse, GetUserConfigurationResponse, CreateUserConfigurationResponse, UpdateUserConfigurationResponse, DeleteUserConfigurationResponse, GetUserPhotoResponse
} from '../data/user';
import { Request } from '@loopback/rest';

/**
 * The user service.
 */
@bind({ scope: BindingScope.SINGLETON, tags: ExchangeSoap.service })
export class UserService {

    /**
     * Get user photo.
     * @param soapRequest The soap request, which is type of GetUserPhotoType
     * @returns User photo
     */
    async GetUserPhoto(soapRequest: GetUserPhotoType, request: Request): Promise<GetUserPhotoResponseMessageType> {
        assert(soapRequest instanceof GetUserPhotoType);
        return GetUserPhotoResponse(soapRequest, request);
    }

    /**
     * Set user photo.
     * @param soapRequest The soap request, which is type of SetUserPhotoType
     * @returns Operation result
     */
    async SetUserPhoto(soapRequest: SetUserPhotoType): Promise<SetUserPhotoResponseMessageType> {
        assert(soapRequest instanceof SetUserPhotoType);
        return mockSetUserPhotoResponse;
    }

    /**
     * Get user OOO settings.
     * @param soapRequest The soap request, which is type of GetUserOofSettingsRequest
     * @param request The request being processed
     * @returns User OOF settings
     */
    async GetUserOofSettings(soapRequest: GetUserOofSettingsRequest, request: Request): Promise<GetUserOofSettingsResponseMessageType> {
        assert(soapRequest instanceof GetUserOofSettingsRequest);
        return GetUserOofSettings(soapRequest, request);
    }

    /**
     * Set user OOF settings.
     * @param soapRequest The soap request, which is type of SetUserOofSettingsRequest
     * @returns Operation result
     */
    async SetUserOofSettings(soapRequest: SetUserOofSettingsRequest, request: Request): Promise<SetUserOofSettingsResponseMessageType> {
        assert(soapRequest instanceof SetUserOofSettingsRequest);
        return SetUserOofSettings(soapRequest, request);
    }

    /**
     * Get user availability.
     * @param soapRequest The soap request, which is type of GetUserAvailabilityRequestType
     * @returns User availability
     */
    async GetUserAvailability(soapRequest: GetUserAvailabilityRequestType): Promise<GetUserAvailabilityResponseMessageType> {
        assert(soapRequest instanceof GetUserAvailabilityRequestType);
        return mockGetUserAvailabilityResponse;
    }

    /**
     * Get user configuration.
     * @param soapRequest The soap request, which is type of GetUserConfigurationType
     * @param request The HTTP request
     * @returns User configuration
     */
    async GetUserConfiguration(soapRequest: GetUserConfigurationType, request: Request): Promise<GetUserConfigurationResponseType> {
        assert(soapRequest instanceof GetUserConfigurationType);

        return GetUserConfigurationResponse(soapRequest, request);
    }

    /**
     * Create user configuration
     * @param soapRequest The soap request, which is type of CreateUserConfigurationType
     * @param request The HTTP request
     * @returns Create user configuration response
     */
    async CreateUserConfiguration(soapRequest: CreateUserConfigurationType, request: Request): Promise<CreateUserConfigurationResponseType> {
        assert(soapRequest instanceof CreateUserConfigurationType);

        return CreateUserConfigurationResponse(soapRequest, request);
    }

    /**
     * Update user configuation
     * @param soapRequest The soap request, which is type of UpdateUserConfiguationType
     * @param request The HTTP request
     * @returns Update user configuration response
     */
    async UpdateUserConfiguration(soapRequest: UpdateUserConfigurationType, request: Request): Promise<UpdateUserConfigurationResponseType> {
        assert(soapRequest instanceof UpdateUserConfigurationType);

        return UpdateUserConfigurationResponse(soapRequest, request);
    }

    /**
     * Delete user configuation
     * @param soapRequest The soap request, which is type of DeleteUserConfigurationType
     * @param request The HTTP request
     * @returns Delete user configuration response
     */
    async DeleteUserConfiguration(soapRequest: DeleteUserConfigurationType, request: Request): Promise<DeleteUserConfigurationResponseType> {
        assert(soapRequest instanceof DeleteUserConfigurationType);

        return DeleteUserConfigurationResponse(soapRequest, request);
    }
}


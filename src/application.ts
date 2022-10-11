/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */
import { BootMixin } from '@loopback/boot';
import {RepositoryMixin} from '@loopback/repository';
import { ApplicationConfig } from '@loopback/core';
import { RestApplication, RestBindings } from '@loopback/rest';

/**
 * The EWS middleware application.
 */
export class EWSMiddlewareApplication extends BootMixin(RepositoryMixin(RestApplication)) {

    /**
     * Creates a new instance of EWSMiddlewareApplication.
     * @param options The application config options
     */
    constructor(options: ApplicationConfig = {}) {
        super(options);

        this.bind(RestBindings.REQUEST_BODY_PARSER_OPTIONS).to({
            text: {
                type: function (): boolean { return true; },
                limit: options.rest.maxRequestBodySize
            },
        });

        this.projectRoot = __dirname;
    }
}

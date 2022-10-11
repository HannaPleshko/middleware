/*
 * Copyright (c) 2020 Topcoder, Inc. All rights reserved.
 */

/**
 * SOAP definition.
 */
export interface SOAP {
    /**
     * Service name.
     */
    service: string;
    /**
     * Port name.
     */
    port: string;
    /**
     * URL path.
     */
    path: string;
    /**
     * WSDL file location.
     */
    wsdl: string;
    /**
     * Namespace.
     */
    namespace: string;
}

/**
 * Exchange SOAP definition.
 */
export const ExchangeSoap: SOAP = {
    service: 'ExchangeService',
    port: 'ExchangeServicePort',
    path: '/EWS/Exchange.asmx',
    wsdl: '../wsdl/ews/Services.wsdl',
    namespace: 'http://schemas.microsoft.com/exchange/services/2006/types',
}

/**
 * Autodiscover SOAP definition.
 */
export const AutodiscoverSoap: SOAP = {
    service: 'AutodiscoverService',
    port: 'AutodiscoverServicePort',
    path: '/autodiscover/autodiscover.svc',
    wsdl: '../wsdl/autodiscover/Services.wsdl',
    namespace: 'http://schemas.microsoft.com/exchange/2010/Autodiscover',
}
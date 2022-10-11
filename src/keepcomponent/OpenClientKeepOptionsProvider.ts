import { OpenClientKeepOptions } from '@hcllabs/openclientkeepcomponent';
import { appContext } from '..';


export class OpenClientKeepOptionsProvider implements OpenClientKeepOptions {

    getKeepBaseUrl(): string {
        return appContext.options.rest.keepBaseUrl;
    }
}


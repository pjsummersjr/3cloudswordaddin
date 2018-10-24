import {Authenticator} from '@microsoft/office-js-helpers';

/**
 * 
 */
export interface IAddInConfig {
    clientId: string;               //This is your client app id, so this needs to change if you create a new one
    resource: string;               //This is your API resource - need to change this if you destination server/API changes
    baseUrl: string; //if you change tenants, you need to change this
    authorizeUrl: string;
    responseType: string;
    nonce: boolean;
    state: boolean;
}
/**
 * Base property interface for any Office Add-In
 */
export interface IOfficeAddinProps {
    isOfficeInitialized:boolean;
    config: IAddInConfig;
    authenticator: Authenticator;
}


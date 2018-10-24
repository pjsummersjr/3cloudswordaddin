import {Authenticator} from '@microsoft/office-js-helpers';
import * as React from 'react';
import * as q from 'q';

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

export interface IOfficeAddinState {}

export class OfficeAddInComponent extends React.Component<IOfficeAddinProps, any> {

    protected provider: string = null;

    loadContent = (resourceUrl: string): Q.Promise<any> => {
        let self = this;

        let deferred = q.defer();

        self.props.authenticator.authenticate(this.provider, false).then(
            function(response: any){
                console.debug(`Requesting data from ${resourceUrl}`);
                let accessToken = response.access_token;
                let requestConfig = {
                    method:"GET",
                    headers: {
                        "Authorization": `Bearer ${accessToken}`,
                        "Content-Type": `application/json`                    }
                }
                fetch(resourceUrl, requestConfig)
                .then((response) => 
                {
                    console.log(`Resolving promise`);
                    deferred.resolve(response.json())
                }) 
            }
        )
        .catch(function(error:any){
            console.error(`Error caught in code: ${error}`);
            return deferred.reject(error);
        });
        return deferred.promise;
    }
}


import {IAddInConfig} from '../components/OfficeAddin';

let tenantConfig = {
    tenantName: 'paulsumm.onmicrosoft.com',
    apiurl: 'pjsummersjr2.ngrok.io'
}

let BaseConfig: IAddInConfig = {
    clientId: "3bca74a9-cd38-477b-9681-ebd8abfea27f",               
    resource: "ca825ba9-2e55-4d7b-bff1-45bfe153f7ab",
    baseUrl: `https://login.windows.net/${tenantConfig.tenantName}`,
    authorizeUrl: '/oauth2/authorize',
    responseType: 'token',
    nonce: true,
    state: true
}

export { BaseConfig, tenantConfig }
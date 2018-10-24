import * as React from 'react';
import * as ReactDOM from "react-dom";
import SimpleService from './services/SimpleService';
import WordAddin from './components/WordAddin';
import * as OfficeHelpers from '@microsoft/office-js-helpers';
//import {Authenticator} from '@microsoft/office-js-helpers';

import {BaseConfig, tenantConfig } from './config/app.config';

import './index.scss';
import { IAddInConfig } from './components/OfficeAddin';

let HelloWorld = () => {
  let service: SimpleService = new SimpleService();
  let message: string = service.getSimpleData()[0];
  return (<h1>Hello there {message}! Welcome to the React World with Webpack!</h1>);
}

const render = (isOfficeInitialized: boolean, authProvider: OfficeHelpers.Authenticator) => {
  
  ReactDOM.render(
    <WordAddin isOfficeInitialized={isOfficeInitialized} config={BaseConfig} authenticator={authProvider}/>,
    document.getElementById("root")
  );
}

Office.initialize = () => {
  /* OfficeHelpers JS required code */      
  if(OfficeHelpers.Authenticator.isAuthDialog()) return;

  let authenticator = new OfficeHelpers.Authenticator();
  authenticator.endpoints.registerAzureADAuth(BaseConfig.clientId, tenantConfig.tenantName, BaseConfig);
  
  /* End OfficeHelper JS code */
  
  render(true, authenticator);
}

render(false, null);



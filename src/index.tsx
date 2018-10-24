import * as React from 'react';
import * as ReactDOM from "react-dom";
import SimpleService from './services/SimpleService';
import WordAddin from './components/WordAddin';
import * as OfficeHelpers from '@microsoft/office-js-helpers';

import {BaseConfig, tenantConfig } from './config/app.config';

import './index.scss';


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
/**
 * We call this so something loads in the page if the page is loaded outside of an Office Add-In.
 */
render(false, null);



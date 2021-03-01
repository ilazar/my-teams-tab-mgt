// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import React from 'react';
import './App.css';
import * as microsoftTeams from "@microsoft/teams-js";
import { Providers, TeamsProvider } from '@microsoft/mgt';
import { BrowserRouter as Router, Route } from "react-router-dom";

import Privacy from "./Privacy";
import TermsOfUse from "./TermsOfUse";
import Tab from "./Tab";

// Initialize the Microsoft Teams SDK
microsoftTeams.initialize();

TeamsProvider.microsoftTeamsLib = microsoftTeams;
Providers.globalProvider = new TeamsProvider ({
    clientId: 'c85f08e1-aa95-4cba-8230-ab0a2ac623d1',
    authPopupUrl: '/auth.html',
    scopes: ['contacts.read', 'user.read'],
})
/**
 * The main app which handles the initialization and routing
 * of the app.
 */
function App() {

  // Display the app home page hosted in Teams
  return (
    <Router>
      <Route exact path="/privacy" component={Privacy} />
      <Route exact path="/termsofuse" component={TermsOfUse} />
      <Route exact path="/tab" component={Tab} />
    </Router>
  );
}

export default App;

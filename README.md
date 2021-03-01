# Accessing Outlook Contacts from React Apps

See [Use the Microsoft Graph Toolkit in Team Tabs](https://docs.microsoft.com/en-us/graph/toolkit/get-started/build-a-microsoft-teams-tab).

## Install Microsoft Graph Toolkit libs

```
npm i @microsoft/mgt-react  --save
npm i @microsoft/mgt-element --save
npm i @microsoft/mgt --save
npm i @microsoft/teams-js --save
```

## Create an Azure Active Directory app

- Goto [Ajure Portal](https://portal.azure.com/#home)
- Open [Ajure Active Directory](https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/Overview)
- Open App registrations (from left menu)
- Create a new registration (e.g. MyApp)
- In the context of MyApp
  - Open Authentication (from left menu)
    - in implicit grant section, check Access tokens (used for implicit flows) and ID tokens (used for implicit and hybrid flows)
  - Open API permissions (from left menu)
    - add 'User.Read' and 'Contacts.Read' permissions for Microsoft Graph as delegated permissions
  - Open the Overview (from left menu)
    - copy the Application (client) ID

## Configure app client ID

Add the following code in ```index.js```

```javascript
import * as microsoftTeams from "@microsoft/teams-js";
import { Providers, TeamsProvider } from '@microsoft/mgt';

// Initialize the Microsoft Teams SDK
microsoftTeams.initialize();

TeamsProvider.microsoftTeamsLib = microsoftTeams;
Providers.globalProvider = new TeamsProvider ({
    clientId: 'c85f08e1-aa95-4cba-8230-ab0a2ac623d1',
    authPopupUrl: '/auth.html',
    scopes: ['contacts.read', 'user.read'],
})
```

## Add the auth page

Add the following ```auth.html``` page as public page:

```html
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <link rel="icon" href="%PUBLIC_URL%/favicon.ico" />
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/11.0.0/css/fabric.min.css" />
    <title>Microsoft Teams Tab Auth</title>
  </head>
  <body>
    <script src="https://unpkg.com/@microsoft/teams-js/dist/MicrosoftTeams.min.js" crossorigin="anonymous"></script>
    <script src="https://unpkg.com/@microsoft/mgt/dist/bundle/mgt-loader.js"></script>
    
    <script>
      mgt.TeamsProvider.handleAuth();
    </script>
  </body>
</html>
```

## Login then import contacts

Add the following code in your ```Tab.js```

```javascript
import React, { useEffect, useState } from 'react';
import { Providers, ProviderState } from '@microsoft/mgt';
import { Login, Get } from "@microsoft/mgt-react";

function useIsSignedIn() {
  const [isSignedIn, setIsSignedIn] = useState(false);
  useEffect(() => {
    const updateState = () => {
      const provider = Providers.globalProvider;
      setIsSignedIn(provider && provider.state === ProviderState.SignedIn);
    };
    Providers.onProviderUpdated(updateState);
    updateState();
    return () => {
      Providers.removeProviderUpdatedListener(updateState);
    }
  }, []);
  return { isSignedIn };
}

function Tab() {
  const [isImporting, setImporting] = useState(false);
  const { isSignedIn } = useIsSignedIn();
  useEffect(() => setImporting(true), [isSignedIn]);
  const log = event => () => console.log(event);
  const handleDataChange = e => {
    setImporting(false);
    if (e.detail.error) {
      console.warn('Failed to fetch contacts', e.detail.error);
    } else {
      const contacts = e.detail.response.value;
      console.log('Fetched contacts');
      contacts.forEach(console.log);
    }
  };
  return (
    <div>
      <div>Outlook Contacts</div>
      <Login
        loginInitiated={log('loginInitiated')}
        loginCompleted={log('loginCompleted')}
        loginFailed={log('loginFailed')}
        logoutCompleted={log('logoutCompleted')}
      />
      {isSignedIn && (
        <>
          <Get
            resource="me/contacts"
            maxPages={10}
            dataChange={handleDataChange}
          />
          {isImporting && <div>Importing...</div>}
        </>
      )}
    </div>
  );
}

export default Tab;
```

### Details about the MGT components

- ```Login``` component (see https://docs.microsoft.com/en-us/graph/toolkit/components/login)
  - When user is not signed in, it shows a ```Sign in``` label and an icon.
  - When user triggers sign in, a popup is shown allowing the user to authenticate, then authorize the applcation
    to access user contacts.
  - When user is signed in, it shows the username.
  - When the user clicks on username, a popup is opened allowing the user to sign out.

- The custom hook ```useIsSignedIn``` determine whether the user is signed in
  - the MGT ```Providers``` keep internally the sign in state
  
- ```Get``` component (see https://docs.microsoft.com/en-us/graph/toolkit/components/get)
  - Fetch any Microsoft Graph resource if the user is signed in
  - For outlook contacts, the resource is ```me/contacts``` (see https://docs.microsoft.com/en-us/graph/api/user-list-contacts?view=graph-rest-1.0&tabs=http)
  - Default ```maxPages``` is 3. Setting ```maxPages``` to 0 means all pages.
  - ```dataChange``` callback can be used to get the result (```detail.error``` or an array of contacts ```detail.response.value```)
  - This component provides also a ```refresh([force = false])``` method to refresh the fetched contacts.
  - Caching the fetched content can also be configured.

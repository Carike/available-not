// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// <msalConfig>
const msalConfig = {
  auth: {
    clientId: '[ENTER YOUR CLIENT ID]',
    redirectUri: 'http://localhost:8081'
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
    forceRefresh: false
  }
};

const loginRequest = {
  scopes: [
    'openid',
    'profile',
    'user.read',
    'calendars.read',
    'presence.read'
  ]
}
// </msalConfig>

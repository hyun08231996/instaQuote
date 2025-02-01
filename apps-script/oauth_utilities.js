// Defines the level of access the OAuth token has to Google services and
// resources. Scope list and Documentation:
// https://developers.google.com/identity/protocols/oauth2/scopes
// Reference for the code below : https://medium.com/@jessecodescode/connect-an-apps-script-to-google-cloud-platform-via-oauth-2-0-tokens-june-2023-2c104f672c93
const scope = [
'https://www.googleapis.com/auth/script.external_request',
'https://www.googleapis.com/auth/cloud-platform',
'https://www.googleapis.com/auth/devstorage.read_write',
'https://mail.google.com/',
'https://www.googleapis.com/auth/userinfo.email',
'https://www.googleapis.com/auth/drive'
].join(' ');

// Retrieves web app URL of the current Google Apps Script project.
const webAppUrl = ScriptApp.getService().getUrl();

// Check that webAppUrl ends in /exec or /dev. Then generates callback URL for use
// as a redirect after authorization of script's access to Google account.
function generateCallbackUrl(webAppUrl) {
  // if (webAppUrl.endsWith("/exec")) {
  //   return (webAppUrl.slice(0, -4)) + 'usercallback'
  // } if (webAppUrl.endsWith("/dev")) {
  //   return (webAppUrl.slice(0, -3)) + 'usercallback'
  // } else {
  //   throw new Error('Problem with generating callback url. Web app URL issue.');
  // }
  if (webAppUrl.includes("/exec")) {
    return (webAppUrl.split("/exec")[0]) + '/usercallback'
  } if (webAppUrl.includes("/dev")) {
    return (webAppUrl.split("/dev")[0]) + '/usercallback'
  } else {
    throw new Error('Problem with generating callback url. Web app URL issue.');
  }
}
const callbackUrl = generateCallbackUrl(webAppUrl)

// Gets a property store that all users can access, but only within this script.
const propertyStore = PropertiesService.getScriptProperties();

// Gets a copy of all key-value pairs in the current Properties store.
const scriptProps = propertyStore.getProperties();

// Gets the cache instance scoped to the script. Allows you to insert,
// retrieve, and remove items from a cache.
const cache = CacheService.getScriptCache();

// Function to build HTML webpage for user to initiate OAuth flow.
// Checks that refresh token not present then proceeds to make state token
// and authorization url for use in HTML output. State token is used to
// maintain the state and security of OAuth flow. Authorization url takes
// required parameters documented here:
// https://developers.google.com/identity/protocols/oauth2/web-server#creatingclient
// If refresh token present, generates HTML stating authorization completed.

function buildOAuthAuthPage() {
  if (!scriptProps.refresh_token) {
    const stateToken = ScriptApp.newStateToken()
      .withMethod('googleCallback')
      .withArgument('name', 'value')
      .withTimeout(360)
      .createToken();

    const params = {
      'state': stateToken,
      'scope': scope,
      'client_id': scriptProps.client_id,
      'redirect_uri': callbackUrl,
      'response_type': 'code',
      'access_type': 'offline',
      'approval_prompt': 'force',
    };

    // Construct query string for authorization URL
    const queryString = Object.keys(params)
      .map(function (key) {
        return key + '=' + encodeURIComponent(params[key]);
      })
      .join('&');

    const authUrl = 'https://accounts.google.com/o/oauth2/auth?' + queryString;

    return HtmlService.createHtmlOutput('<div style="max-width: 1000px;       \
    margin: 0 auto; padding-top: 75px;"><h2>Step 1: Callback URL:</h2>        \
    <p>Add this callback to the OAuth Credential as an Authorized redirect    \
    URI:</p><div style="background-color: lightblue;padding: 15px;            \
    margin-bottom: 100px;"><code>' + callbackUrl + '</code></div><h2>         \
    Step 2: Authorization URL:</h2><p>After waiting ~5 minutes, visit         \
    this link to authorize the application:</p><div style="background-color:  \
    lightblue;padding: 15px;"><code>' + authUrl + '</code></div></div>');
  } else {
    // Return HTML output for already authorized case.
    return HtmlService.createHtmlOutput('Already authorized. Yippee!');
  }
}

// This function is called as a callback when the user completes the OAuth
// authorization process and is redirected back to the script.
// Receives authorization code as a parameter from Google. Performs a
// fetch with POST method and required parameters documented here:
// https://developers.google.com/identity/protocols/oauth2/web-server#httprest_3
// If no error in the response, function puts access token in cache and refresh
// token in property store. Then returns HTML success page.

function googleCallback(params) {
  const url = 'https://oauth2.googleapis.com/token';

  const config = {
    "method": "post",
    "payload": {
      'code': params.parameter.code,
      'redirect_uri': callbackUrl,
      'client_id': scriptProps.client_id,
      'client_secret': scriptProps.client_secret,
      'grant_type': 'authorization_code',
    }
  };

  try {
    const response = JSON.parse(UrlFetchApp.fetch(url, config));
    cache.put("access_token", response.access_token);
    propertyStore.setProperty('refresh_token', response.refresh_token);
    return HtmlService.createHtmlOutput('<h1>Success!</h1><p>You did it! There is \
    now a refresh token in the property store and an access token in the cache.</p>');
  } catch (response) {
      throw new Error(response);
  }
}

// Checks if the token is present, it returns the access token.
// If the token is not found, it calls the refreshAccessToken() function to
// obtain a new access token then returns access token.

function getAccessToken() {
  let accessToken = cache.get('access_token');
  if (!accessToken) {
    accessToken = refreshAccessToken();
  } else {
    return accessToken;
  }
}

// Makes a request to the Google OAuth token endpoint to refresh the
// access token using the refresh token. Performs a fetch with POST
// method and required parameters documented here:
// https://developers.google.com/identity/protocols/oauth2/web-server#offline
// Does not prompt user for permission.

function refreshAccessToken() {
  const url = 'https://oauth2.googleapis.com/token';
  const config = {
    "method": "post",
    "payload": {
      'client_id': scriptProps.client_id,
      'client_secret': scriptProps.client_secret,
      'refresh_token': scriptProps.refresh_token,
      'grant_type': 'refresh_token',
    }
  };

  try {
    const response = JSON.parse(UrlFetchApp.fetch(url, config));
    cache.put("access_token", response.access_token, response.expires_in);
    return response.access_token;
  } catch (response) {
      throw new Error(response);
  }
}

// Checks for access token in cache by logging to execution log.
function logAccessToken() {
  Logger.log(cache.get('access_token'))
}

// Deletes access token in cache.
function deleteAccessToken() {
  cache.remove('access_token')
  Logger.log('Manually removed access_token from cache.')
}
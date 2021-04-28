// Requires:
//  - Library with script ID 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
//  - Client ID & Client Secret from Github:
//    (set redirect URL to https://script.google.com/macros/d/{THIS_SCRIPTS_ID}/usercallback)
//  - OAuth scope https://www.googleapis.com/auth/script.external_request

const DOCUMENT_OAUTH_KEY = "document_oauth_key";

/* A bare-bones Github Gists API client. */
function GithubGistClient() {
  // Use of OAuth2 requires https://www.googleapis.com/auth/script.external_request
  this.oauthService = OAuth2.createService('githubGistClient')
    .setAuthorizationBaseUrl('https://github.com/login/oauth/authorize')
    .setTokenUrl('https://github.com/login/oauth/access_token')
    .setClientId(/* it's a secret ;) */)
    .setClientSecret(/* it's a secret ;) */)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setCache(CacheService.getUserCache())
    .setScope('gist');
}

/*
@returns {boolean} Whether the script can get the user's Github OAuth token.
*/
GithubGistClient.prototype.hasAccess = function() {
  return this.oauthService.hasAccess();
}

/*
@returns {string} URL to send the user to for authentication.
*/
GithubGistClient.prototype.getAuthorizationUrl = function() {
  return this.oauthService.getAuthorizationUrl();
}

/*
Logs the current user out.
*/
GithubGistClient.prototype.logout = function() {
  return this.oauthService.reset();
}

GithubGistClient.prototype.logTokenUseInThisSheet = function() {
  PropertiesService.getDocumentProperties().setProperty(DOCUMENT_OAUTH_KEY, this.oauthService.getAccessToken());
}

GithubGistClient.prototype.tokenHasBeenUsedInThisSheet = function() {
  return PropertiesService.getDocumentProperties().getProperty(DOCUMENT_OAUTH_KEY) == this.oauthService.getAccessToken();
}

/*
@returns {HtmlOutput} HTML to display to the user upon successful or failed auth.
*/
GithubGistClient.prototype.handleCallback = function(request) {
  var template = HtmlService.createTemplateFromFile('githubGistClientAuthMessage');
  if (this.oauthService.handleCallback(request)) {
    template.error = false;
    template.message = 'Success! You can close this tab.';
  } else {
    template.error = true;
    template.message = 'Denied. Please close this tab and try again.';
  }
  return template.evaluate();
}

/* 
Creates a new Github Gist.
Does not do *anything* to handle errors. Most common exception will be if the user revokes access in Github.

@param {string} content - Content of new Gist.
@param {string} filename - Filename for new Gist.
@param {string} language - Language the new Gist is written in.
@param {string} description - Description of the new Gist.
@param {boolean} public - Whether the new Gist should be publicly viewable.
@returns {string} URL of new gist.
*/
GithubGistClient.prototype.newGist = function(content, filename, language, description, public) {
  var body = {
    "description": description,
    "files": {
      filename: {
        "content": content,
        "filename": filename,
        "language": language
      }
    },
    "public": public,
  };

  var response = this.makeRequest("post", /* resource = */ "", body);

  return response.files[filename].raw_url;
};

/* 
Makes authenticated HTTP request to Github client.
@param {string} method - HTTP method.
@param {string} resource - path within Github Gists API to make a request to, starting with "/".
@returns {string} URL of new commit.
*/
GithubGistClient.prototype.makeRequest = function(method, resource, data) {
   var headers = {
    "Authorization" : "token " + this.oauthService.getAccessToken(),
    "accept": "application/vnd.github.v3+json",
  };
  
  var options = {'headers': headers, method: method};
  
  if (data) {
    options['contentType'] = 'application/json';
    options['payload'] = JSON.stringify(data);
  }
  // Requires https://www.googleapis.com/auth/script.external_request
  var response = UrlFetchApp.fetch("https://api.github.com/gists" + resource, options);
  return JSON.parse(response);
}

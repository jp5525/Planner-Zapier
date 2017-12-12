
const getAccessToken = (z, bundle) => {
  
  const promise = z.request(`https://login.microsoftonline.com/common/oauth2/v2.0/token`, {
    method: 'POST',
    body: {
      client_id: '{{process.env.CLIENT_ID}}',
      code: '{{bundle.inputData.code}}',
      scope: "offline_access user.read.all  Group.ReadWrite.All",
      redirect_uri: '{{bundle.inputData.redirect_uri}}',
      grant_type: 'authorization_code',
      client_secret: '{{process.env.CLIENT_SECRET}}',
    },
    headers: {
      'content-type': 'application/x-www-form-urlencoded'
    }
  });

  // Needs to return at minimum, `access_token`, and if your app also does refresh, then `refresh_token` too
  return promise.then((response) => {
    if (response.status !== 200) {
      throw new Error('Unable to fetch access token: ' + response.content);
    }

    const result = JSON.parse(response.content);
    
    return {
      access_token: result.access_token,
      refresh_token: result.refresh_token
    };
  });
};

const refreshAccessToken = (z, bundle) => {
  const promise = z.request(`https://login.microsoftonline.com/common/oauth2/v2.0/token`, {
    method: 'POST',
    body: {
      refresh_token: '{{bundle.authData.refresh_token}}',
      client_id: '{{process.env.CLIENT_ID}}',
      client_secret: '{{process.env.CLIENT_SECRET}}',
      grant_type: 'refresh_token',
      scope: "offline_access user.read.all Group.ReadWrite.All",
      redirect_uri: '{{bundle.inputData.redirect_uri}}',
     
    },
    headers: {
      'content-type': 'application/x-www-form-urlencoded'
    }
  });

  // Needs to return `access_token`. If the refresh token stays constant, can skip it. If it changes, can
  // return it here to update the user's auth on Zapier.
  return promise.then((response) => {
    if (response.status !== 200) {
      console.log(response.content);
      throw new Error('Unable to fetch access token: ' + response.content);
    }

    const result = JSON.parse(response.content);
    return {
      access_token: result.access_token
    };
  });
};

const testAuth = (z /*, bundle*/) => {
  // Normally you want to make a request to an endpoint that is either specifically designed to test auth, or one that
  // every user will have access to, such as an account or profile endpoint like /me.
  const promise = z.request({
    url: `https://graph.microsoft.com/v1.0/me`,
  });

  // This method can return any truthy value to indicate the credentials are valid.
  // Raise an error to show
  return promise.then((response) => {
    if (response.status === 401) {
      throw new Error('The access token you supplied is not valid');
    }
    return response;
  });
};

module.exports = {
  type: 'oauth2',
  oauth2Config: {
    // Step 1 of the OAuth flow; specify where to send the user to authenticate with your API.
    // Zapier generates the state and redirect_uri, you are responsible for providing the rest.
    // Note: can also be a function that returns a string
    authorizeUrl:{
        url: `https://login.microsoftonline.com/common/oauth2/v2.0/authorize`,
        params: {
            client_id: '{{process.env.CLIENT_ID}}',
            response_type: 'code',
            redirect_uri: '{{bundle.inputData.redirect_uri}}',
            response_mode: 'query',
            scope: 'offline_access user.read.all  Group.ReadWrite.All',
            state: '{{bundle.inputData.state}}',
            
        }
    },
    // Step 2 of the OAuth flow; Exchange a code for an access token.
    // This could also use the request shorthand.
    getAccessToken: getAccessToken,
    // (Optional) If the access token expires after a pre-defined amount of time, you can implement
    // this method to tell Zapier how to refresh it.
    refreshAccessToken: refreshAccessToken,
    // If you want Zapier to automatically invoke `refreshAccessToken` on a 401 response, set to true
    autoRefresh: true
    // If there is a specific scope you want to limit your Zapier app to, you can define it here.
    // Will get passed along to the authorizeUrl
    // scope: 'read,write'
  },
  
  // The test method allows Zapier to verify that the access token is valid. We'll execute this
  // method after the OAuth flow is complete to ensure everything is setup properly.
  test: testAuth
};
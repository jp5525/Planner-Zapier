
const {requestWithAccessToken} = require('./util/withAccessToken');

const testAuth = (z, bundle) => {

  // Normally you want to make a request to an endpoint that is either specifically designed to test auth, or one that
  // every user will have access to, such as an account or profile endpoint like /me.
  
  const promise = requestWithAccessToken({url: `https://graph.microsoft.com/v1.0/groups`}, z, bundle);

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
    type:"custom",
    fields: [
      { key: 'appId', label: 'App ID', required: true, helpText: "The App Id of the app registered in Azure Entra Admin Center" },
      { key: 'clientSecret', label: 'Client Secret', required: true, helpText: "The value of the client secrect ( NOT the ID ) app registered in Azure Entra Admin Center" },
      { key: 'tenant', label: 'Tenant', required: true, helpText: "The directory tenant that you want to request permission from. The value can be in GUID or a friendly name format." }
    ],
    test:testAuth
};
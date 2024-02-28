const {requestWithAccessToken} = require('../withAccessToken');

const findUser = (z, bundle, user) => 
  requestWithAccessToken({url: `https://graph.microsoft.com/v1.0/users?\$filter=displayName eq '${user}' or mail eq '${user}' or id eq '${user}'`}, z, bundle)
    .then(response => JSON.parse(response.content).value[0]);

module.exports = { findUser }
  
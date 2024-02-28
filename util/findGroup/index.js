const {requestWithAccessToken} = require('../withAccessToken');

const findGroup = (z, bundle, group) =>
    requestWithAccessToken({url: `https://graph.microsoft.com/v1.0/groups?\$filter=displayName eq '${group}'`}, z, bundle)
        .then(response => {return JSON.parse(response.content).value[0]});
      
module.exports = { findGroup }
  
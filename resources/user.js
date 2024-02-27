const {requestWithAccessToken} = require('../util/withAccessToken');

// get a list of users
const listUsers = (z, bundle) => 
  requestWithAccessToken({url: 'https://graph.microsoft.com/v1.0/users'}, z, bundle)
    .then(response => JSON.parse(response.content).value);

module.exports = {
  key: 'user',
  noun: 'User',

  
  list: {
    display: {
      label: 'New User',
      description: 'Lists the users.'
    },
    operation: {
      perform: listUsers
    }
  },
  
};

const {requestWithAccessToken} = require('../util/withAccessToken');

// get a list of groups
const listGroups = (z, bundle) => 
  requestWithAccessToken({url: 'https://graph.microsoft.com/v1.0/groups'}, z, bundle)
    .then(response => {return JSON.parse(response.content).value});

module.exports = {
  key: 'group',
  noun: 'Group',

  
  list: {
    display: {
      label: 'New Group',
      description: 'Lists the groups.'
    },
    operation: {
      perform: listGroups
    }
  },

  

};


// get a list of groups
const listGroups = (z) => {
  const responsePromise = z.request({
    url: 'https://graph.microsoft.com/v1.0/groups',
   
  });
  return responsePromise
    .then(response => {return JSON.parse(response.content).value});
};


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

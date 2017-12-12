
// get a list of users
const listUsers = (z) => {
  const responsePromise = z.request({
    url: 'https://graph.microsoft.com/v1.0/users'
  });
  return responsePromise
    .then(response => JSON.parse(response.content).value);
};



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

const {requestWithAccessToken} = require('../util/withAccessToken');
const {findGroup} = require("../util/findGroup")

// get a list of plans
const listPlans = (z, bundle, groupId) =>
  requestWithAccessToken({url: `https://graph.microsoft.com/v1.0/groups/${groupId ?? bundle.inputData.Group }/planner/plans`}, z, bundle)
    .then(response => JSON.parse(response.content).value);


// create a plan
const createPlan = async (z, bundle) => 
  requestWithAccessToken({
    method: 'POST',
    url: 'https://graph.microsoft.com/v1.0/planner/plans',
    body: {
      title: bundle.inputData.name,
      container:{
        containerId: (await findGroup(z, bundle, bundle.inputData.Group)).id,
        type: "group"
        
      }
    }
  }, z, bundle)
  .then(response => JSON.parse(response.content));


module.exports = {
  key: 'plan',
  noun: 'Plan',

  list: {
    display: {
      label: 'New Plan',
      description: 'Lists the plans.'
    },
    operation: {
      inputFields:[
        {key: 'Group', required: true, label: "Get plans belonging to this group", dynamic: 'groupList.id.displayName'},
      ],
      perform: listPlans
    }
  },

  create: {
    display: {
      label: 'Create Plan',
      description: 'Creates a new plan.'
    },
    operation: {
      inputFields: [
        {key: 'Group', required: true, label: "The group the plan will belong to", dynamic: 'groupList.id.displayName'},
        {key: 'name', required: true}
      ],
      perform: createPlan
    },
  }
  
  
};

// get a list of plans
const listPlans = (z, bundle) => {
  const responsePromise = z.request({
    url: `https://graph.microsoft.com/v1.0/groups/${bundle.inputData.Group}/planner/plans`,
  });
  
  return responsePromise
    .then(response => JSON.parse(response.content).value);
};

// create a plan
const createPlan = (z, bundle) => {

  const responsePromise = z.request({
    method: 'POST',
    url: 'https://graph.microsoft.com/v1.0/planner/plans',
    body: {
      title: bundle.inputData.name, // json by default
      owner: bundle.inputData.Group
    }
  });
  return responsePromise
    .then(response => JSON.parse(response.content));
};

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
  },

  
};

const {requestWithAccessToken} = require('../util/withAccessToken');
const { findPlan } = require("../util/findPlan")
const { findGroup } = require("../util/findGroup")

// get a list of buckets
const listBuckets = (z, bundle) => 
  requestWithAccessToken(
    {url: `https://graph.microsoft.com/v1.0/planner/plans/${bundle.inputData.Plan}/buckets/` },
    z,
    bundle
  )
  .then(response => JSON.parse(response.content).value);

// create a bucket
const createBucket = async (z, bundle) => {
  const {
    inputData:{ Group, Plan, name}
  } = bundle;
  const {id: groupId }= await findGroup(z, bundle, Group);
  const {id: planId} =  await findPlan(z, bundle, Plan, groupId);

  return requestWithAccessToken(
    {
      method: 'POST',
      url: 'https://graph.microsoft.com/v1.0/planner/buckets',
      body: { name,  planId } 
    },
    z,
    bundle
  )
  .then(response => JSON.parse(response.content));

}
module.exports = {
  key: 'bucket',
  noun: 'Bucket',

  list: {
    display: {
      label: 'New Bucket',
      description: 'Lists the buckets.'
    },
    operation: {
      inputFields:[
        {key: 'Group', required: true, label: "Get Plans belonging to this group", dynamic: 'groupList.id.displayName'},
        {key: 'Plan', required: true, label: "Get Buckets belonging to this Plan", dynamic: 'planList.id.title'}
      ],
      perform: listBuckets
    }
  },

  create: {
    display: {
      label: 'Create Bucket',
      description: 'Creates a new bucket.'
    },

    operation: {
      inputFields: [
        {key: 'Group', required: true, label: "Get plans belonging to this group", dynamic: 'groupList.id.displayName'},
        {key: 'Plan', required: true, label: "The plan to add this bucket to", dynamic: 'planList.id.title'},
        {key: 'name', required: true}
      ],
      perform: createBucket
    },
  },

};

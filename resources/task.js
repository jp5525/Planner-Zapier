const { requestWithAccessToken } = require('../util/withAccessToken');
const { findPlan } = require("../util/findPlan")
const { findGroup } = require("../util/findGroup")
const { findBucket } = require("../util/findBucket")
const { list:{ operation:{perform: getUsers}}} = require("./user")

// get a list of tasks
const listTasks = async(z, bundle) => {
  
  const {id: planId} = (await findPlan(z, bundle, bundle.inputData.Plan, bundle.groupId));
  return requestWithAccessToken({url: `https://graph.microsoft.com/v1.0/planner/plans/${planId}/tasks`}, z, bundle)
    .then(response => JSON.parse(response.content).value)
}
// create a task
const createTask = async (z, bundle) => {
  const {authData} = bundle
  let a = ''+bundle.inputData.Assignee;
  let assignees = {};
  const users = await getUsers(z, bundle);

  if(bundle.inputData.Assignee != undefined)
    assignees = a.split(',').reduce((all, curr, index) =>{
      const {id : userId} = users.find(({displayName, mail, id}) => [displayName, mail, id].includes(curr) )
      all[userId] = {
        "@odata.type": "microsoft.graph.plannerAssignment",
        "orderHint": " !"+" !".repeat(index)
      }
      return all;
    }, {})

  
  const {id: groupId }= await findGroup(z, bundle, bundle.inputData.Group);
  const {id: planId} =  await findPlan(z, bundle, bundle.inputData.Plan, groupId);
  const {id: bucketId} =  await findBucket(z, {authData, inputData:{Plan:planId}}, bundle.inputData.Bucket)
  const body = {
    planId,
    bucketId,
    title: bundle.inputData.Title,
    startDateTime: bundle.inputData.StartDate,
    dueDateTime: bundle.inputData.DueDate,
    assignments: assignees
  };

  const responsePromise = requestWithAccessToken({
    method: 'POST',
    url: 'https://graph.microsoft.com/v1.0/planner/tasks',
    body
  }, z, bundle);
  
  return responsePromise
    .then(response => JSON.parse(response.content));
};

//Needs testing
const searchTask = async(z, bundle) =>{
  const {inputData: {Plan, Title, Description, Group}} = bundle;
  const {id: groupId} = await findGroup(z, bundle, Group);
  const {id: planId} = await findPlan(z, bundle, Plan, groupId);
  let foundTask;
  //get the tasks
  const tasksRaw = await requestWithAccessToken({method: "GET", url:`https://graph.microsoft.com/v1.0/planner/plans/${planId}/tasks/`}, z, bundle);
  //filter by title
  const tasks = tasksRaw.json.value.filter(x=> Title? x.title.includes(Title) : true );
  // get details and filter by description
  if(Description){
    const tasksWithDescription = tasks.filter(x=>x.hasDescription);
    const taskDetails = await Promise.all(
      tasksWithDescription.map(x=>
        requestWithAccessToken({
          method: "GET",
          url:`https://graph.microsoft.com/v1.0/planner/tasks/${x.id}/details`
        }, z, bundle)
        )
      )
    const taskDetailsWithMatchingDescription = taskDetails.map(x=>x.json).filter(x=>x.description.includes(Description))
    if(taskDetailsWithMatchingDescription.length >0){
      const details = taskDetailsWithMatchingDescription[0];
      const task  = tasks.find(x=>x.id == details.id);
      foundTask = {...task, details};
    }
    else{
      //If we didnt find task with a matching title and description
      return []
    }
    
  }
  else{
    if(tasks.length >0){
      const task = tasks[0];
      const details = await requestWithAccessToken({
        method: "GET",
        url:`https://graph.microsoft.com/v1.0/planner/tasks/${task.id}/details`
      }, z, bundle)
      .then(res=>res.json);

      foundTask = {...task, details}
    }
    else {
      //If we didnt find a task with a mathing Title
      return []
    }
  }

  //Get conversations on task
  const conversationThreadRaw = await requestWithAccessToken({
    method: "GET",
    url:`https://graph.microsoft.com/v1.0/groups/${groupId}/threads/${foundTask.conversationThreadId}`
  }, z, bundle);

  if (conversationThreadRaw.json.error){
    return [{...foundTask, groupId: inputData.Group}]
  }
  else{
    const posts =  await requestWithAccessToken({
      method: "GET",
      url:`https://graph.microsoft.com/v1.0/groups/${groupId}/threads/${foundTask.conversationThreadId}/posts`
    }, z, bundle)

    const conversationThread = {...conversationThreadRaw.json, posts: posts.json.value};
    return [{...foundTask, conversationThread, groupId }]
  }
  
}


module.exports = {
  key: 'task',
  noun: 'Task',

  search:{
    display: {
      label: 'Find Task',
      description: 'Finds an existing Task by name.'
    },
    operation: {
      inputFields: [
        {key: 'Group', required: true, label: "Get plans belonging to this group", dynamic: 'groupList.id.displayName'},
        {key: 'Plan', required: true, label: "Get Tasks belonging to this Plan", dynamic: 'planList.id.title'},
        {key: 'Title', required: false, type: 'string' },
        {key: 'Description', required: false, type: 'string' }
      ],
      perform: searchTask,
      
    }
  },

  list: {
    display: {
      label: 'New Task',
      description: 'Lists the tasks.'
    },
    operation: {
      inputFields:[
        {key: 'Group', required: true, label: "Get plans belonging to this group", dynamic: 'groupList.id.displayName'},
        {key: 'Plan', required: true, label: "Get Tasks belonging to this Plan", dynamic: 'planList.id.title'}    
      ],
      perform: listTasks
    }
  },

  create: 
    {
    display: {
      label: 'Create Task',
      description: 'Creates a new task.'
    },
    operation: {
      inputFields:[
        {key: 'Group', required: true, label: "Get plans belonging to this group", dynamic: 'groupList.id.displayName'},
        {key: 'Plan', required: true, label: "Get Buckets belonging to this Plan", dynamic: 'planList.id.title'},
        {key: 'Bucket', required: true, label: "Add Tasks to to this Bucket", dynamic: 'bucketList.id.name'},
        {key: 'Title', type: 'string', required: true, helpText: 'The name of the task'},
        {key: 'Assignee', required: false, label: "The user to assign the task to", dynamic: 'userList.id.displayName', list: true},
        {key: 'StartDate', required: false, label: "The date work should start on.", type:"datetime"},
        {key: 'DueDate', required: false, label: "The date work should be completed by", type:"datetime"}
      ],
      perform: createTask
    },
  }
    

  
};

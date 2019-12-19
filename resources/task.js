
// get a list of tasks
const listTasks = (z, bundle) => {
  const responsePromise = z.request({
    url: `https://graph.microsoft.com/v1.0/planner/plans/${bundle.inputData.Plan}/tasks`
  });
  return responsePromise
    .then(response => JSON.parse(response.content).value);
};

// create a task
const createTask = (z, bundle) => {

  let a = ''+bundle.inputData.Assignee;
  let assignees = {};
  if(bundle.inputData.Assignee != undefined)
    assignees = a.split(',').reduce((all, curr, index) =>{
      all[curr] = {
        "@odata.type": "microsoft.graph.plannerAssignment",
        "orderHint": " !"+" !".repeat(index)
      }
      return all;
    }, {})

 

  const responsePromise = z.request({
    method: 'POST',
    url: 'https://graph.microsoft.com/v1.0/planner/tasks',
    body: {
      planId: bundle.inputData.Plan,
      bucketId: bundle.inputData.Bucket,
      title: bundle.inputData.Title,
      startDateTime: bundle.inputData.StartDate,
      dueDateTime: bundle.inputData.DueDate,
      assignments: assignees
    }
  });
  
  return responsePromise
    .then(response => JSON.parse(response.content));
};

const searchTask = (z, bundle) =>{
  let allDetails;
  const {inputData} = bundle;
  //get the tasks
  return z.request({
    method: "GET",
    url:`https://graph.microsoft.com/v1.0/planner/plans/${bundle.inputData.Plan}/tasks/`
  })
  //filter by title
  .then(response =>{
    if(inputData.Title){
      return response.json.value.filter(x=>x.title.includes(inputData.Title))
    }
    else{
      return response.json.value
    }
  })
  // get details and filter by description
  .then(tasks=>{
    if(inputData.Description){
      const tasksWithDescription = tasks.filter(x=>x.hasDescription);
      return Promise.all(tasksWithDescription.map(x=>
        z.request({
          method: "GET",
          url:`https://graph.microsoft.com/v1.0/planner/tasks/${x.id}/details`
        }
      )))
      .then(res => {
        allDetails = res.map(x=>x.json);
        return allDetails.filter(x=>x.description.includes(inputData.Description));
      })
      .then(res=>{
        if(res.length >0){
          const details = res[0];
          const task  = tasks.find(x=>x.id == details.id);
          return {...task, details};
        }
        else if(tasks.length > 0){
          const task = tasks[0];
          const details = allDetails.find(x=>x.id == task.id)
          return {...task, details};
        }
        else return {}
      })
    }
    else{
      if(tasks.length >0){
        const task = tasks[0];
        return z.request({
          method: "GET",
          url:`https://graph.microsoft.com/v1.0/planner/tasks/${task.id}/details`
        })
        .then(res=>({...task, details: res.json}))
      }
      else return {}
    }
  })
  //get the task conversation
  .then(task =>{
    return z.request({
      method: "GET",
      url:`https://graph.microsoft.com/v1.0/groups/${bundle.inputData.Group}/threads/${task.conversationThreadId}`
    })
    .then(response => {
      if (response.json.error)
        return [{...task, groupId: inputData.Group}]
      else{
        //get the posts on the conversation
        return z.request({
          method: "GET",
          url:`https://graph.microsoft.com/v1.0/groups/${bundle.inputData.Group}/threads/${task.conversationThreadId}/posts`
        })
        .then(res=>{
          const conversationThread = {...response.json, posts: res.json.value}
          return [{...task, conversationThread, groupId:inputData.Group}]
        })
      }
    })
  })
  
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
        { key: 'Title', required: false, type: 'string' },
        { key: 'Description', required: false, type: 'string' }
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

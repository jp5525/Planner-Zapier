
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

module.exports = {
  key: 'task',
  noun: 'Task',


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

  create: {
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
  },

  
};

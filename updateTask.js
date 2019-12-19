const updateTask = (z, bundle) =>{
    return updateTaskFields(z, bundle)
    .then(task=>{
      return updateDetails(z, bundle)
      .then(details=>({...task, details}))
    })
  
  }
  
  const updateTaskFields = (z, bundle) =>{
    const { Task, Task_Etag, DueDateTime: dueDateTime, PercentComplete:percentComplete, StartDateTime:startDateTime, Title:title} = bundle.inputData
    const body={
      ...(dueDateTime && {dueDateTime}),
      ...(startDateTime && {startDateTime}),
      ...(percentComplete && {percentComplete}),
      ...(title && {title}),
    }
    if(Object.keys(body).length > 0){
      return z.request({
        method: "PATCH",
        url:`https://graph.microsoft.com/v1.0//planner/tasks/${Task}`,
        headers:{"If-Match": Task_Etag},
        body
      })
      .catch(response=>new Promise((res, rej)=>res({fieldsMessage: "did not update fields"})))
    }
    else{
      return new Promise((res, rej)=>res({fieldsMessage: "did not update fields"}))
    }
  }
  
  const updateDetails = (z, bundle) => {
    const { Task, Task_Details_Etag:Etag, Description:description} = bundle.inputData
    if(description){
      return z.request({
        method: "PATCH",
        url:`https://graph.microsoft.com/v1.0//planner/tasks/${Task}/details`,
        headers:{"If-Match": Etag},
        body:{description}
      })
      .catch(response =>new Promise((res, rej)=>res({detailsMessage: "did not update details"})))
    }
    else{
      return new Promise((res, rej) =>res({detailsMessage: "did not update details"}))
    }
  }
module.exports = {
    key: "task",
    noun: "task",
    display: {
        label: 'Update Task',
        description: 'Update a new task.'
      },
      operation: {
        inputFields:[
          {key: 'Group', required: false, label: "Get plans belonging to this group", dynamic: 'groupList.id.displayName'},
          {key: 'Plan', required: false, label: "Get Buckets belonging to this Plan", dynamic: 'planList.id.title'},
          {key: 'Bucket', required: false, label: "Get Tasks belonging to this Bucket", dynamic: 'bucketList.id.name'},
          {key: 'Task', required: true, label: "The Task to comment on", dynamic: 'taskList.id.title'},
          {key: 'Task_Etag', required: false, label: "The etag value of the task. This is required to update the fields below"},
          {key: 'Title', type: 'string', required: false, helpText: 'The title of the Task'},
          {key: 'StartDate', required: false, label: "The date work should start on.", type:"datetime"},
          {key: 'DueDate', required: false, label: "The date work should be completed by", type:"datetime"},
          {key: 'PercentComplete', required: false, label: "The the percent of the task that has been complete", type:"integer"},
          {key: 'Task_Details_Etag', required: false, label: "The etag value of the task details. This is required to update the fields below"},
          {key: 'Description', type: 'string', required: false, helpText: 'The description of the task'},
          
        ],
        perform: updateTask
      },
}
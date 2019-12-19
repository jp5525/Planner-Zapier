// create a comment
const createConversation = (z, bundle) => {
  let conversationThreadId, eTag, title;
  //get the task to comment on
  const responsePromise = z.request({
    method: 'GET',
    url: `https://graph.microsoft.com/v1.0/planner/tasks/${bundle.inputData.Task}`,
  })
  .then(response => z.JSON.parse(response.content))
  //add comment to task
  .then(json => {
    conversationThreadId = json.conversationThreadId;
    eTag = json["@odata.etag"];
    title = json.title;
    if(conversationThreadId){   //in the case that there is already conversation thread on the task
      return z.request({
        method: 'POST',
        url: `https://graph.microsoft.com/v1.0/groups/${bundle.inputData.Group}/threads/${conversationThreadId}/reply`,
        body:{
          post:{
            body:{
              content: bundle.inputData.Content,
              contentType: "text"
            }
          }
        }
      })
    }
    else{   //if there is not already a conversation thread then create one and add a comment
      return z.request({
        method: 'POST',
        url: `https://graph.microsoft.com/v1.0/groups/${bundle.inputData.Group}/conversations`,
        body:{
          topic: `Comments on task "${title}"`,
          threads:[
            {
              posts:[
                {
                  body:{
                    content: bundle.inputData.Content,
                    contentType: "text"
                  }
                }
              ]
            }
          ]
        }
      })
      .then(response => z.JSON.parse(response.content))
      .then(json =>{
        const {threads} = json;
        const [{id}] = threads;
        conversationThreadId = id;
        z.console.log({id, threads});
        return z.request({
          method: 'PATCH',
          url: `https://graph.microsoft.com/v1.0/planner/tasks/${bundle.inputData.Task}`,
          body:{
            "conversationThreadId": id
          },
          headers:{"If-Match": eTag}
        })

      })
    }
  })
  .then(res =>{ 
    if(res.status >= 200 && res.status < 300)
      return {success: true, conversationThreadId}
  })
 
  
  return responsePromise
};

module.exports = {
  key: 'comment',
  noun: 'comment',


  create: {
    display: {
      label: 'Create comment',
      description: 'Add a new comment to a task.'
    },
    operation: {
      inputFields: [
        {key: 'Group', required: true, label: "Get plans belonging to this group", dynamic: 'groupList.id.displayName'},
        {key: 'Plan', required: true, label: "Get Tasks belonging to this Plan", dynamic: 'planList.id.title'},
        {key: 'Task', required: true, label: "The Task to comment on", dynamic: 'taskList.id.title'},
        {key: 'Content', required: true, label: "The text of your comment" }
      ],
      perform: createConversation
    },
  },

  sample: {
    id: 'AAQkADRhMDFiZWYxLTc3YTQtNDY1NC1hNDA5LTM3NjFjNTc0ZTQ2ZQAQAMX4dRAYzHtIiTqIFiCbRSs=',
    topic: 'Comments on task \"Test\"'
  },

  outputFields: [
    {key: 'id', label: 'ID'},
    {key: 'topic', label: 'Topic'}
  ]
};

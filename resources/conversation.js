const {requestWithAccessToken} = require('../util/withAccessToken');
const { findGroup } = require("../util/findGroup")
const { findTask } = require("../util/findTask")
// create a comment
const createConversation = async (z, bundle) => {
  //get the task to comment on
  const { inputData:{ Task, Group, Content}} = bundle;

  const { id: groupId} = await findGroup(z, bundle, Group);
  const task = await findTask(z, {...bundle, groupId}, Task);
  const {conversationThreadId, title} = task
  const eTag = task["@odata.etag"];
  
  if(conversationThreadId){   //in the case that there is already conversation thread on the task
      return await requestWithAccessToken({
        method: 'POST',
        url: `https://graph.microsoft.com/v1.0/groups/${groupId}/threads/${conversationThreadId}/reply`,
        body:{
          post:{
            body:{
              content: Content,
              contentType: "text"
            }
          }
        }
      }, z, bundle)
  }
  else{   //if there is not already a conversation thread then create one and add a comment
    const conversationRaw = await requestWithAccessToken({
      method: 'POST',
      url: `https://graph.microsoft.com/v1.0/groups/${groupId}/conversations`,
      body:{
        topic: `Comments on task "${title}"`,
        threads:[
          {
            posts:[
              {
                body:{
                  content: Content,
                  contentType: "text"
                }
              }
            ]
          }
        ]
      }
    }, z, bundle)
    const {threads: [conversationThreadId]} = z.JSON.parse(conversationRaw.content)   
    await requestWithAccessToken({
      method: 'PATCH',
      url: `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}`,
      body:{ conversationThreadId },
      headers:{"If-Match": eTag}
    }, z, bundle)

    if(conversationRaw.status >= 200 && conversationRaw.status < 300){
      return {success: true, conversationThreadId}
    }
    
  }
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

# Planner-Zapier
A Zapier integration for Microsoft planner. 

If this is something you would find helpful you can get access to the app via the invite link 
https://zapier.com/platform/public-invite/1754/ede4dcddff02f9fec29ac3bc6b3829d8/

## Setting up Office 365
To set this up in Office 365 you will need to create an app in Microsoft Entra. This connector is designed to use app-only access. This artical goes over how to setup the app.

https://learn.microsoft.com/en-us/graph/auth-v2-service?tabs=http

The app will need these premission
* Directory.ReadWrite.All
* Group.ReadWrite.All
* Tasks.ReadWrite.All
* User.Read

You will use the AppId, client secret, and the tenant id to setup the connection in Zapier.

## Actions
___
Most of the parameters for these actions can take either the ID or the human readable display names  


**Create a Task**
* Group (id, displayName)
* Plan (id, title)
* Bucket (id, name)
* Title
* Assignee (id, email, name)
* StartDate
* DueDate

**Update a Task**
* Group (id, displayName)
* Plan (id, title)
* Bucket (id, name)
* Task (id, title )
* Task_ETag (This is the etag that is currently on the task. You will need to pull by doing a search on for the tasks)
* Title
* StartDate
* DueDate
* PercentComplete
* Task_Details_Etag  (This is the etag that is currently on the task detail. You will need to pull by doing a search on for the tasks)
* Description

**Create Bucket**
* Group (id, displayName)
* Plan (id, title)
* name

**Create Plan**
* Group (id, displayName)
* name

## Searchs
___

**Find a Task**
* Group (id, displayName)
* Plan (id, title)
* Title
* Description

## Triggers
___
**New Task**

**New Bucket**

**New Plan**

**New Group**



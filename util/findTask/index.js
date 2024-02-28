const { list:{ operation:{ perform: listTasks} }} = require("../../resources/task")

const findTask = (z, bundle, task) => listTasks(z, bundle)
    .then(tasks=> tasks.find(
        ({title, id}) => [title, id].includes(task))
    )
      
module.exports = { findTask }
  

const { list:{ operation:{ perform: listPlans}}} =  require("../../resources/plan");

const findPlan = (z, bundle, plan, groupId) => listPlans(z, bundle, groupId)
    .then(plans=> plans.find(
        ({title, id}) => [title, id].includes(plan)
      )
    )
      
module.exports = { findPlan }
  
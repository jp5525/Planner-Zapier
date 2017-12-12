const UserResource = require('./resources/user');
const GroupResource = require('./resources/group');
const PlanResource = require('./resources/plan');
const BucketResource = require('./resources/bucket');
const TaskResource = require('./resources/task');
// We can roll up all our behaviors in an App.
const authentication = require('./authentication');
const addBearerHeader = (request, z, bundle) => {
  if (bundle.authData && bundle.authData.access_token) {
    request.headers.Authorization = `Bearer ${bundle.authData.access_token}`;
  }
  return request;
};

const App = {
  // This is just shorthand to reference the installed dependencies you have. Zapier will
  // need to know these before we can upload
  version: require('./package.json').version,
  platformVersion: require('zapier-platform-core').version,

  authentication: authentication,
  // beforeRequest & afterResponse are optional hooks into the provided HTTP client
  beforeRequest: [
    addBearerHeader
  ],

  afterResponse: [
  ],

  // If you want to define optional resources to simplify creation of triggers, searches, creates - do that here!
  resources: {
    [UserResource.key]: UserResource,
    [GroupResource.key]: GroupResource,
    [PlanResource.key]: PlanResource,
    [BucketResource.key]: BucketResource,
    [TaskResource.key]: TaskResource,
  },

  // If you want your trigger to show up, you better include it here!
  triggers: {
  },

  // If you want your searches to show up, you better include it here!
  searches: {
  },

  // If you want your creates to show up, you better include it here!
  creates: {
  }
};

// Finally, export the app.
module.exports = App;
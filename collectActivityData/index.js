// Import Modules
const groupBy = require('lodash/groupBy')
const GRAPH = require('../modules/ms-graph')
const STORAGE = require('../modules/ms-storage')

// Create the module objects
const msgraph = new GRAPH()
const storage = new STORAGE()

const main = async function (context, timer) {

  // authenticate
  await msgraph.authenticate()

  // Export AD Users
  const usersData = await msgraph.getAllUsers()
  context.log(`Total Users: ${usersData.length}`)

  const groupedUserData = groupBy(usersData, 'userPrincipalName')

  // Get Teams Activity Report Location
  const teamsData = await msgraph.getTeamsUserActivityUserDetail()

  const augmentedTeamsData = teamsData
    // Filter leavers from teams data
    .filter(reportItem => groupedUserData[reportItem.userPrincipalName])
    // Enrich Teams Data with user data
    .map(reportItem => 
      Object.assign(
        reportItem,
        groupedUserData[reportItem.userPrincipalName][0]
      )
  );

  // Active Teams Data
  console.log(`Active Teams Data: ${augmentedTeamsData.length}`);


  // Find the active users without a teams account.
  const teamsUserArray = augmentedTeamsData.map(
    ({ userPrincipalName }) => userPrincipalName
  );

  // Remove the Active users with out teams accounts.
  const activeTeamsUserData = usersData.filter(user => teamsUserArray.includes(user.userPrincipalName));
  console.log(`Active Users with Teams account: ${activeTeamsUserData.length}`);

  // Write user data to Blob
  const teamsBlobId = await storage.uploadBlob('teams-json', 'teams_activity_data.json', JSON.stringify({ value: augmentedTeamsData }))

  // Log Blob Upload Response Id
  context.log('Teams JSON Blob Upload Id: ' + teamsBlobId)

  // End
  context.done();

};

module.exports = main

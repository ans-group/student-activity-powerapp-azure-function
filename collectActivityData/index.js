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
  const groupedUserData = groupBy(usersData, 'userPrincipalName')

  // Get Teams Activity Report Location
  const teamsData = await msgraph.getTeamsUserActivityUserDetail()

  const augmentedTeamsData = teamsData.map(reportItem => Object.assign(reportItem, (groupedUserData[reportItem.userPrincipalName] && groupedUserData[reportItem.userPrincipalName][0]) || {}))

  // Write user data to Blob
  const userBlobId = await storage.uploadBlob('user-json', 'user_data.json', JSON.stringify({ value: usersData }))

  // Write user data to Blob
  const teamsBlobId = await storage.uploadBlob('teams-json', 'teams_activity_data.json', JSON.stringify({ value: augmentedTeamsData }))

  // Log Blob Upload Response Id
  context.log('User JSON Blob Upload Id: ' + userBlobId)
  context.log('Teams JSON Blob Upload Id: ' + teamsBlobId)

  // End
  context.done();

};

module.exports = main

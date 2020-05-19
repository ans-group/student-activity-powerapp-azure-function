require('dotenv').config()
const request = require('request-promise')

const TENANT_ID = process.env.TENANT_ID
const CLIENT_ID = process.env.CLIENT_ID
const CLIENT_SECRET = process.env.CLIENT_SECRET

const loginUrl = 'https://login.microsoftonline.com'
const graphUrl = 'https://graph.microsoft.com'

module.exports = class GRAPH {
  // Get Token Function
  async authenticate () {
    // Configure Request Options
    const options = {
      method: 'POST',
      uri: `${loginUrl}/${TENANT_ID}/oauth2/v2.0/token`,
      form: {
        client_id: CLIENT_ID,
        scope: 'https://graph.microsoft.com/.default',
        client_secret: CLIENT_SECRET,
        grant_type: 'client_credentials'
      },
      headers: {
        'content-type': 'application/x-www-form-urlencoded'
      },
      json: true
    }

    try {
      // Send Request
      const response = await request(options)

      // Return Response
      this.token = response.access_token
    } catch (err) {
      throw err.error.error.message
    };
  };

  // Get Member Graph Users
  async getAllUsers () {
    try {
      if (!this.token) throw new Error('Please authenticate first')

      const uri = `${graphUrl}/v1.0/users?$filter=userType eq 'Member'&$select=displayName,givenName,surname,mail,usageLocation,id,userPrincipalName`
      const data = await this._fetchAll(uri)

      // Return Response
      return data
    } catch (err) {
      throw new Error(err)
    };
  };

  // Get Graph Group Members
  async getGroupMembersById (groupId) {
    try {
      if (!this.token) throw new Error('Please authenticate first')

      const uri = `${graphUrl}/v1.0/groups/${groupId}/members?$select=displayName,givenName,surname,mail,usageLocation,id,userPrincipalName`
      const data = await this._fetchAll(uri)

      // Return Response
      return data
    } catch (err) {
      throw new Error(err)
    };
  };

  // Get Teams User Activity
  async getTeamsUserActivityUserDetail () {
    try {
      if (!this.token) throw new Error('Please authenticate first')

      const uri = `${graphUrl}/beta/reports/getTeamsUserActivityUserDetail(period='D7')?$format=application/json`
      const data = await this._fetchAll(uri)

      // Return data
      return data
    } catch (err) {
      throw new Error(err)
    }
  };

  async _fetchAll (uri, data = [], nextLink) {
    try {
      const options = {
        method: 'GET',
        uri: nextLink || uri,
        headers: {
          Authorization: `Bearer ${this.token}`,
          Host: 'graph.microsoft.com'
        },
        json: true
      }
      const response = await request(options)
      const newData = data.concat(response.value)
      return response['@odata.nextLink'] && response.value && response.value.length ? await this._fetchAll(uri, newData, response['@odata.nextLink']) : newData
    } catch (err) {
      throw new Error(err)
    }
  };
  
};

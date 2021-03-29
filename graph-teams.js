var graph = require('@microsoft/microsoft-graph-client');
const { response } = require('./app');
require('isomorphic-fetch');


module.exports = {
//<GET TEAMS>
  getTeams: async function(accessToken) {
    const client = getAuthenticatedClient(accessToken);
    const teams = await client
      .api('/me/joinedTeams')
      // Get just the properties used by the app
      .select('id,displayName')
      .get();
    return teams;
  },

//</GET TEAMS>


// <CreateTeamSnippet>
   createTeam: async function(accessToken, formData) {
    const client = getAuthenticatedClient(accessToken);

    // Build a Graph team
    const newTeam = {
      "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
      displayName: formData.displayName,
      body: {
        contentType: 'text',
        content: formData.body
      }
    };
    // POST /teams
    await client
      .api('/teams')
      .post(newTeam);
  },
// </CreateTeamSnippet>

//<DeleteTeamSnippet>

  deleteTeam: async function(accessToken,id){

    const client = getAuthenticatedClient(accessToken);
    const teams = await client
    //INSERT group id in the URL
    .api('/groups/135ecca9-bd35-4fd3-be16-037e5ef26d1b')
    .delete();
    return teams;
  },

//</DeleteTeamSnippet>


//<GET Teams Channels>

getChannels: async function(accessToken) {
  const client = getAuthenticatedClient(accessToken);
  const channels = await client
  //set groupId in the url
    .api('/teams/e5c34a29-2b47-431e-b3bc-f9934f347581/channels')
    // Get just the properties used by the app
    .get();
  return channels;
},
//</GET Teams Channels>

// <CreateChannelSnippet>
createChannel: async function(accessToken, formData) {
  const client = getAuthenticatedClient(accessToken);

  // Build a Graph channel
  const newChannel = {
    displayName: formData.displayName,
    body: {
      contentType: 'text',
      content: formData.body
    }
  };
  // POST /channel
  await client
    //set /teams/groupId/channels
    .api('/teams/e5c34a29-2b47-431e-b3bc-f9934f347581/channels')
    .post(newChannel);
},
// </CreatChannelSnippet>

//<DeleteChannelSnippet>

deleteChannel: async function(accessToken){

  const client = getAuthenticatedClient(accessToken);
  const channels = await client
  //INSERT group id and channel id in the URL
  .api('/teams/e5c34a29-2b47-431e-b3bc-f9934f347581/channels/19%3ae6e347a95b6f4059a0eb52ce64c7dcf4%40')
  .delete();
  return channels;
},
//</DeleteChannelSnippet>

//<GetChannelMembers>


getChannelMembers: async function(accessToken) {
  const client = getAuthenticatedClient(accessToken);
  const members = await client
    .api('/teams/{{TeamId}}/channels/{{ChannelId}}')
    // Get just the properties used by the app
    .get();
  return members;
},

//</GetChannelMembers>

};

//Authentication Function
function getAuthenticatedClient(accessToken) {
  // Initialize Graph client
  const client = graph.Client.init({
    // Use the provided access token to authenticate
    // requests
    authProvider: (done) => {
      done(null, accessToken);
    }
  });

  return client;
}
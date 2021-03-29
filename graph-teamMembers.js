var graph = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

module.exports = {
  //groups/{{TeamId}}/members
  // <GetTeamMembersSnippet>

  getTeamMembers: async function (accessToken) {
    const client = getAuthenticatedClient(accessToken);
    const teamMembers = await client
    //set /groups/groupId/members
      .api('/groups/0796ed98-d3b0-49df-9c58-31786640382e/members')
      .get();
    return teamMembers;
  },
  // </GetTeamMembersSnippet>

  // <CreateMemberSnippet>
  createMember: async function (accessToken, formData) {
    const client = getAuthenticatedClient(accessToken);
    // Build a Graph members
    const newmember =
    {
      "@odata.type": "#microsoft.graph.aadUserConversationMember",
      "roles": ["guest"],
      "user@odata.bind": "https://graph.microsoft.com/v1.0/users('bf3719ef-af75-45d0-8017-9bab3d2d3a26')" //set user-id in the url

    };
    // POST /member
    await client
    // /teams/groupId/members
      .api('/teams/e5c34a29-2b47-431e-b3bc-f9934f347581/members')
      .post(newmember);
  },
  // </CreateMemberSnippet>

  //<DELETEMEMBER>

  deleteMember: async function (accessToken, id) {

    const client = getAuthenticatedClient(accessToken);
    const teams = await client

      //set groupId and membership-id
      .api('/teams/e5c34a29-2b47-431e-b3bc-f9934f347581/members/ZTVjMzRhMjktMmI0Ny00MzFlLWIzYmMtZjk5MzRmMzQ3NTgxIyNiZjM3MTllZi1hZjc1LTQ1ZDAtODAxNy05YmFiM2QyZDNhMjY=')
      .delete();
    return teams;
  },
  //<DELETEMEMBER>
};

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
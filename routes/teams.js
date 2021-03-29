const router = require('express-promise-router')();
const graph = require('../graph-teams');
const { body, validationResult } = require('express-validator');
const validator = require('validator');
const { compile } = require('morgan');
const { use } = require('.');


/* GET /teams */
router.get('/',
  async function (req, res) {
    if (!req.session.userId) {
      // Redirect unauthenticated requests to home page
      res.redirect('/')
    } else {
      const params = {
        active: { teams: true }
      };
      // Get the user
      const user = req.app.locals.users[req.session.userId];

      // Get the access token
      var accessToken;
      try {
        accessToken = await getAccessToken(req.session.userId, req.app.locals.msalClient);
      } catch (err) {
        req.flash('error_msg', {
          message: 'Could not get access token. Try signing out and signing in again.',
          debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
        });
        return;
      }

      if (accessToken && accessToken.length > 0) {
        try {
          // Get the teams
          const teams = await graph.getTeams(accessToken)
          params.teams = teams.value;
        } catch (err) {
          req.flash('error_msg', {
            message: 'Could not fetch teams',
            debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
          });
        }
      }
      else {
        req.flash('error_msg', 'Could not get an access token');
      }
      res.render('teams', params);
    }
  }
);

router.get('/newteam',
  function (req, res) {
    if (!req.session.userId) {
      // Redirect unauthenticated requests to home page
      res.redirect('/')
    } else {
      res.locals.newTeam = {};
      res.render('createTeam');
    }
  }
);

// <PostTeamFormSnippet>
/* POST /team/new */
router.post('/newteam', [
  body('display-name').customSanitizer(value => {
    return value;
  }),
  body('tm-body').escape()
], async function (req, res) {
  if (!req.session.userId) {
    // Redirect unauthenticated requests to home page
    res.redirect('/')
  } else {
    // Build an object from the form values
    const formData = {
      displayName: req.body['display-name'],
      body: req.body['tm-body']
    };

    // Check if there are any errors with the form values
    const formErrors = validationResult(req);
    if (!formErrors.isEmpty()) {

      let invalidFields = '';
      formErrors.errors.forEach(error => {
        invalidFields += `${error.param.slice(3, error.param.length)},`
      });
    }

    // Get the access token
    var accessToken;
    try {
      accessToken = await getAccessToken(req.session.userId, req.app.locals.msalClient);
    } catch (err) {
      req.flash('error_msg', {
        message: 'Could not get access token. Try signing out and signing in again.',
        debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
      });
      return;
    }

    // Get the user
    const user = req.app.locals.users[req.session.userId];

    // Create the team
    try {
      await graph.createTeam(accessToken, formData);
    } catch (error) {
      req.flash('error_msg', {
        message: 'Could not create team',
        debug: JSON.stringify(error, Object.getOwnPropertyNames(error))
      });
    }

    // Redirect back to the teams page
    return res.redirect('/teams');
  }
});

//<DELETE-TEAM>
router.get('/delete',
  async function (req, res) {
    console.log(req.params.id)
    if (!req.session.userId) {
      // Redirect unauthenticated requests to home page
      res.redirect('/')
    } else {
      const params = {
        active: { teams: true }
      };
      // Get the user
      const user = req.app.locals.users[req.session.userId];

      // Get the access token
      var accessToken;
      try {
        accessToken = await getAccessToken(req.session.userId, req.app.locals.msalClient);
      } catch (err) {
        req.flash('error_msg', {
          message: 'Could not get access token. Try signing out and signing in again.',
          debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
        });
        return;
      }

      if (accessToken && accessToken.length > 0) {
        try {
          // Delete the teams
          const teams = await graph.deleteTeam(accessToken)
          params.teams = teams.value;


        } catch (err) {
          req.flash('error_msg', {
            message: 'Could not fetch teams',
            debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
          });
        }
      }
      else {
        req.flash('error_msg', 'Could not get an access token');
      }
      res.render('teams', params);
    }
  }
);
//</DELETE-TEAM>

/* GET CHANNELS */

router.get('/channels',
  async function (req, res) {
    if (!req.session.userId) {
      // Redirect unauthenticated requests to home page
      res.redirect('/')
    } else {
      const params = {
        active: { channels: true }
      };
      // Get the user
      const user = req.app.locals.users[req.session.userId];
      // Get the access token
      var accessToken;
      try {
        accessToken = await getAccessToken(req.session.userId, req.app.locals.msalClient);
      } catch (err) {
        req.flash('error_msg', {
          message: 'Could not get access token. Try signing out and signing in again.',
          debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
        });
        return;
      }

      if (accessToken && accessToken.length > 0) {
        try {
          // Get the channels
          const channels = await graph.getChannels(accessToken)
          params.channels = channels.value;

          console.log(channels)
        } catch (err) {
          req.flash('error_msg', {
            message: 'Could not fetch channels',
            debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
          });
        }
      }
      else {
        req.flash('error_msg', 'Could not get an access token');
      }
      res.render('channels', params);
    }
  }
);
//<DELETE-CHANNEL>
router.get('/removeChannel',
  async function (req, res) {
    console.log(req.params.id)
    if (!req.session.userId) {
      // Redirect unauthenticated requests to home page
      res.redirect('/')
    } else {
      const params = {
        active: { channel: true }
      };
      // Get the user
      const user = req.app.locals.users[req.session.userId];

      // Get the access token
      var accessToken;
      try {
        accessToken = await getAccessToken(req.session.userId, req.app.locals.msalClient);
      } catch (err) {
        req.flash('error_msg', {
          message: 'Could not get access token. Try signing out and signing in again.',
          debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
        });
        return;
      }

      if (accessToken && accessToken.length > 0) {
        try {
          // Delete the teams
          const channels = await graph.deleteChannel(accessToken)
          params.channels = channels.value;


        } catch (err) {
          req.flash('error_msg', {
            message: 'Could not fetch channels',
            debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
          });
        }
      }
      else {
        req.flash('error_msg', 'Could not get an access token');
      }
      res.render('channels', params);
    }
  }
);
//</DELETE-CHANNEL>


router.get('/newchannel',
  function (req, res) {
    if (!req.session.userId) {
      // Redirect unauthenticated requests to home page
      res.redirect('/')
    } else {
      res.locals.newChannel = {};
      res.render('createChannel');
    }
  }
);

// <PostTeamFormSnippet>
/* POST /team/new */
router.post('/newchannel', [
  body('display-name').customSanitizer(value => {
    return value;
  }),
  body('ch-body').escape()
], async function (req, res) {
  if (!req.session.userId) {
    // Redirect unauthenticated requests to home page
    res.redirect('/')
  } else {
    // Build an object from the form values
    const formData = {
      displayName: req.body['display-name'],
      body: req.body['ch-body']
    };

    // Check if there are any errors with the form values
    const formErrors = validationResult(req);
    if (!formErrors.isEmpty()) {

      let invalidFields = '';
      formErrors.errors.forEach(error => {
        invalidFields += `${error.param.slice(3, error.param.length)},`
      });
    }

    // Get the access token
    var accessToken;
    try {
      accessToken = await getAccessToken(req.session.userId, req.app.locals.msalClient);
    } catch (err) {
      req.flash('error_msg', {
        message: 'Could not get access token. Try signing out and signing in again.',
        debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
      });
      return;
    }

    // Get the user
    const user = req.app.locals.users[req.session.userId];

    // Create the team
    try {
      await graph.createChannel(accessToken, formData);
    } catch (error) {
      req.flash('error_msg', {
        message: 'Could not create team',
        debug: JSON.stringify(error, Object.getOwnPropertyNames(error))
      });
    }

    // Redirect back to the teams page
    return res.redirect('/teams');
  }
});


//<GetChannelMembers>>

router.get('/',
  async function (req, res) {
    if (!req.session.userId) {
      // Redirect unauthenticated requests to home page
      res.redirect('/')
    } else {
      const params = {
        active: { channelsMembers: true }
      };
      // Get the user
      const user = req.app.locals.users[req.session.userId];

      // Get the access token
      var accessToken;
      try {
        accessToken = await getAccessToken(req.session.userId, req.app.locals.msalClient);
      } catch (err) {
        req.flash('error_msg', {
          message: 'Could not get access token. Try signing out and signing in again.',
          debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
        });
        return;
      }

      if (accessToken && accessToken.length > 0) {
        try {
          // Get the teams
          const channelsMembers = await graph.getChannelMembers(accessToken)
          params.channelsMembers = channelsMembers.value;
        } catch (err) {
          req.flash('error_msg', {
            message: 'Could not fetch members',
            debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
          });
        }
      }
      else {
        req.flash('error_msg', 'Could not get an access token');
      }
      res.render('channels', params);
    }
  }
);


// </PostEventFormSnippet>

async function getAccessToken(userId, msalClient) {
  // Look up the user's account in the cache
  try {
    const accounts = await msalClient
      .getTokenCache()
      .getAllAccounts();

    const userAccount = accounts.find(a => a.homeAccountId === userId);

    // Get the token silently
    const response = await msalClient.acquireTokenSilent({
      scopes: process.env.OAUTH_SCOPES.split(','),
      redirectUri: process.env.OAUTH_REDIRECT_URI,
      account: userAccount
    });

    return response.accessToken;
  } catch (err) {
    console.log(JSON.stringify(err, Object.getOwnPropertyNames(err)));
  }
}

module.exports = router;

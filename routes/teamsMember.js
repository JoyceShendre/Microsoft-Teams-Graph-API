const router = require('express-promise-router')();
const graph = require('../graph-teamMembers');
const { body, validationResult } = require('express-validator');
const validator = require('validator');
const { compile } = require('morgan');


/* GET /teamMembers */
router.get('/',
  async function (req, res) {
    if (!req.session.userId) {
      // Redirect unauthenticated requests to home page
      res.redirect('/')
    } else {
      const params = {
        active: { teamMembers: true }
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
          // Get the teamMembers
          const teamMembers = await graph.getTeamMembers(accessToken)
          params.teamMembers = teamMembers.value;
          console.log(teamMembers)


        } catch (err) {
          req.flash('error_msg', {
            message: 'Could not fetch teamMembers',
            debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
          });
        }
      }
      else {
        req.flash('error_msg', 'Could not get an access token');
      }
      res.render('members', params);
    }
  }
);


/* POST /team/new */
router.get('/newmember',
  function (req, res) {
    if (!req.session.userId) {
      // Redirect unauthenticated requests to home page
      res.redirect('/')
    } else {
      res.locals.newmember = {};
      res.render('createmember');
    }
  }
);
/* POST /team/new */
router.post('/newmember', [
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
      displayName: req.body['display-name']
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
      await graph.createMember(accessToken, formData);
    } catch (error) {
      req.flash('error_msg', {
        message: 'Could not create member',
        debug: JSON.stringify(error, Object.getOwnPropertyNames(error))
      });
    }

    // Redirect back to the teamMembers page
    return res.redirect('/teamMembers');
  }
});

router.get('/remove',
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
          const teams = await graph.deleteMember(accessToken)
          params.teams = teams.value;


        } catch (err) {
          req.flash('error_msg', {
            message: 'Could not fetch member',
            debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
          });
        }
      }
      else {
        req.flash('error_msg', 'Could not get an access token');
      }
      res.render('members', params);
    }
  }
);
//</DELETE-TEAM>

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

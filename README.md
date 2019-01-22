# passport-microsoftgraph-contacts

## Install

    $ npm install passport-microsoftgraph-contacts

#### Configure Strategy

The Microsoft authentication strategy authenticates users using a Microsoft account and OAuth 2.0 tokens.  The strategy requires a `verify` callback,
which accepts these credentials and calls `done` providing a user, as well as
`options` specifying a client ID, client secret, and callback URL.

    passport.use(new MicrosoftStrategy({
        clientID: MICROSOFT_CLIENT_ID,
        clientSecret: MICROSOFT_CLIENT_SECRET,
        callbackURL: "http://www.example.com/auth/microsoft/callback"
      },
      function(accessToken, refreshToken, profile, done) {

      }
    ));

#### Authenticate Requests

Use `passport.authenticate()`, specifying the `'microsoft'` strategy, to
authenticate requests.

For example, as route middleware in an [Express](http://expressjs.com/)
application:

    app.get('/auth/outlook',
      passport.authenticate('windowslive', { scope: [      
      'openid',
      'profile',
      'offline_access',
      'https://outlook.office.com/contacts.read'] 
    }));

    app.get('/auth/outlook/callback', 
      passport.authenticate('windowslive', { failureRedirect: '/login' }),
      function(req, res) {
        // Successful authentication, redirect home.
        res.redirect('/');
      });

## License

[The MIT License](http://opensource.org/licenses/MIT)


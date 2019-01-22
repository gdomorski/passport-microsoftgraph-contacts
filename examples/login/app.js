var express = require('express')
  , passport = require('passport')
  , MicrosoftGraphStrategy = require('passport-microsoftgraph-contacts').Strategy
  , logger = require('morgan')
  , cookieParser = require('cookie-parser')
  , bodyParser = require('body-parser')
  , session = require('express-session');

  var MICROSOFT_GRAPH_CLIENT_ID = "--insert-client-id-here--"
  var MICROSOFT_GRAPH_CLIENT_SECRET = "--insert-password-here--";


// Passport session setup.
//   To support persistent login sessions, Passport needs to be able to
//   serialize users into and deserialize users out of the session.  Typically,
//   this will be as simple as storing the user ID when serializing, and finding
//   the user by ID when deserializing.  However, since this example does not
//   have a database of user records, the complete Microsoft Graph profile is
//   serialized and deserialized.
passport.serializeUser(function(user, done) {
  done(null, user);
});

passport.deserializeUser(function(obj, done) {
  done(null, obj);
});

var app = express()

// Use the MicrosoftGraphStrategy within Passport.
//   Strategies in Passport require a `verify` function, which accept
//   credentials (in this case, an accessToken, refreshToken, and Microsoft Graph
//   profile), and invoke a callback with a user object.
passport.use(new MicrosoftGraphStrategy({
    clientID: MICROSOFT_GRAPH_CLIENT_ID,
    clientSecret: MICROSOFT_GRAPH_CLIENT_SECRET,
    callbackURL: "http://localhost:3000/auth/outlook/callback",
    passReqToCallback: true,
  },
  function(request, token, tokenSecret, profile, done) {
    request.session.outlook = { queryCode: request.query.code }
    return done(null, profile)
  }
));




// configure Express
app.set('views', __dirname + '/views');
app.set('view engine', 'ejs');
app.use(logger('dev'));
app.use(bodyParser());
app.use(cookieParser());
app.use(session({
    secret: 'blah'
}));

// Initialize Passport!  Also use passport.session() middleware, to support
// persistent login sessions (recommended).
app.use(passport.initialize());
app.use(passport.session());
app.use(express.static(__dirname + '/public'));


app.get('/', function(req, res){
  res.render('index', { user: req.user });
});

app.get('/account', ensureAuthenticated, function(req, res){
  res.render('account', { user: req.user });
});

app.get('/login', function(req, res){
  res.render('login', { user: req.user });
});

// GET /auth/windowslive
//   Use passport.authenticate() as route middleware to authenticate the
//   request.  The first step in Microsoft Graph authentication will involve
//   redirecting the user to live.com.  After authorization, Microsoft Graph
//   will redirect the user back to this application at
//   /auth/windowslive/callback
app.get('/auth/outlook',
  passport.authenticate('windowslive', {
    scope: [
    'openid',
    'profile',
    'offline_access',
    'https://outlook.office.com/contacts.read'
    ],
  }),
  function(req, res){
    // The request will be redirected to Microsoft Graph for authentication, so
    // this function will not be called.
  });

// GET /auth/windowslive/callback
//   Use passport.authenticate() as route middleware to authenticate the
//   request.  If authentication fails, the user will be redirected back to the
//   login page.  Otherwise, the primary route function function will be called,
//   which, in this example, will redirect the user to the home page.
app.get('/auth/outlook/callback', 
  passport.authenticate('windowslive', { failureRedirect: '/login' }),
  function(req, res) {
    res.redirect('/');
  });

app.get('/logout', function(req, res){
  req.logout();
  res.redirect('/');
});

app.listen(3000);


// Simple route middleware to ensure user is authenticated.
//   Use this route middleware on any resource that needs to be protected.  If
//   the request is authenticated (typically via a persistent login session),
//   the request will proceed.  Otherwise, the user will be redirected to the
//   login page.
function ensureAuthenticated(req, res, next) {
  if (req.isAuthenticated()) { return next(); }
  res.redirect('/login')
}

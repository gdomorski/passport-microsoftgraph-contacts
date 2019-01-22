/**
 * Module dependencies.
 */
var util = require('util')
  , OAuth2Strategy = require('passport-oauth2')
  , Profile = require('./profile')
  , InternalOAuthError = require('passport-oauth2').InternalOAuthError
  , LiveConnectAPIError = require('./errors/liveconnectapierror');


/**
 * `Strategy` constructor.
 *
 * The Windows Live authentication strategy authenticates requests by delegating
 * to Windows Live using the OAuth 2.0 protocol.
 *
 * Applications must supply a `verify` callback which accepts an `accessToken`,
 * `refreshToken` and service-specific `profile`, and then calls the `done`
 * callback supplying a `user`, which should be set to `false` if the
 * credentials are not valid.  If an exception occured, `err` should be set.
 *
 * Options:
 *   - `clientID`      your Windows Live application's client ID
 *   - `clientSecret`  your Windows Live application's client secret
 *   - `callbackURL`   URL to which Windows Live will redirect the user after granting authorization
 *
 * Examples:
 *
 *     passport.use(new WindowsLiveStrategy({
 *         clientID: '123-456-789',
 *         clientSecret: 'shhh-its-a-secret'
 *         callbackURL: 'https://www.example.net/auth/windowslive/callback'
 *       },
 *       function(accessToken, refreshToken, profile, done) {
 *         User.findOrCreate(..., function (err, user) {
 *           done(err, user);
 *         });
 *       }
 *     ));
 *
 * @param {Object} options
 * @param {Function} verify
 * @api public
 */
function Strategy(options, verify) {
  options = options || {};
  options.authorizationURL = options.authorizationURL || 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize';
  options.tokenURL = options.tokenURL || 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
  
  OAuth2Strategy.call(this, options, verify);
  this.name = 'microsoft';
  this._userProfileURL = options.userProfileURL || 'https://graph.microsoft.com/v1.0/me/';
}

/**
 * Inherit from `OAuth2Strategy`.
 */
util.inherits(Strategy, OAuth2Strategy);


/**
 * Retrieve user profile from Windows Live.
 *
 * This function constructs a normalized profile, with the following properties:
 *
 *   - `provider`         always set to `windowslive`
 *   - `id`               the user's Windows Live ID
 *   - `displayName`      the user's full name
 *
 * @param {String} accessToken
 * @param {Function} done
 * @api protected
 */
Strategy.prototype.userProfile = function(accessToken, done) {
  var oauth2 = this._oauth2;

  oauth2.get(this._userProfileURL, accessToken, function (err, body, res) {
    var json,
      json2,
      contacts;

    if (err) {
      if (err.data) {
        try {
          json = JSON.parse(err.data);
        } catch (_) {}
      }
        
      if (json && json.error) {
        return done(new LiveConnectAPIError(json.error.message, json.error.code));
      }
      return done(new InternalOAuthError('Failed to fetch user profile', err));
    }
    
    try {
      json = JSON.parse(body);

      oauth2.get('https://graph.microsoft.com/v1.0/me/contacts', accessToken, function (err, body, res) {
        if (err) {
          if (err.data) {
            try {
              json2 = JSON.parse(err.data);
            } catch (_) {}
          }

          if (json2 && json2.error) {
            return done(new LiveConnectAPIError(json2.error.message, json2.error.code));
          }
          return done(new InternalOAuthError('Failed to fetch user profile', err));
        }

        try {
          contacts = JSON.parse(body);
        } catch (ex) {
          return done(new Error('Failed to parse user contacts'));
        }

        json.contacts = contacts.data;

        var profile = Profile.parse(json);
        profile.provider  = 'microsoft';
        profile._raw = body;
        profile._json = json;

        done(null, profile);
      });
    } catch (ex) {
      return done(new Error('Failed to parse user profile'));
    }
  });



};

/**
 * Expose `Strategy`.
 */
module.exports = Strategy;

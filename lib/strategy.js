const util = require('util')
  , OAuth2Strategy = require('passport-oauth2')
  , request = require('request')
  , InternalOAuthError = require('passport-oauth2').InternalOAuthError;

function Strategy(options, verify) {
  options = options || {};
  options.authorizationURL = options.authorizationURL || 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize';
  options.tokenURL = options.tokenURL || 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
  
  OAuth2Strategy.call(this, options, verify);
  this.name = options.name || 'windowslive';
  this._userProfileURL = options.userProfileURL || 'https://outlook.office.com/api/v2.0/me';
  this._userContactsURL = options.userContactsURL || 'https://outlook.office.com/api/v2.0/me/contacts'

  /**
   * Overwrite `_oauth2.get` to use `request.get` allowing for custom headers.
   */
  this._oauth2.get = (url, accessToken) => {
    return new Promise(function (resolve, reject) {
      request.get({url: url,
        headers: {
          'Authorization': 'Bearer ' + accessToken,
          'Accept': 'application/json; odata.metadata=none'
        }
      }, function (error, res, body) {
        if (!error && res.statusCode == 200) {
          resolve({body, res});
        } else {
          reject(error);
        }
      });
    });
    }
  }


/**
 * Inherit from `OAuth2Strategy`.
 */
util.inherits(Strategy, OAuth2Strategy);


const parseContacts = contacts => {
  let arrayOfContacts = [];
  if(contacts.length){
    contacts.forEach(person => {
      person.EmailAddresses.forEach(eachEmail => {
        arrayOfContacts.push({ name: person.DisplayName, email: eachEmail.Address })
      })
    })
    return arrayOfContacts;
  }
}

Strategy.prototype.userProfile = async function (accessToken, done) {

  let profile;
  try {
    let {body: profileContent, res: profileRes} = await this._oauth2.get(this._userProfileURL, accessToken)
    let {body: userContacts, res: contactsRes} = await this._oauth2.get(this._userContactsURL, accessToken)

    if (profileRes && profileRes.statusCode === 404) {
      return done(new OutlookAPIError(res.headers['x-caserrorcode'], res.statusCode));
    }

    let json;

    try {
      json = JSON.parse(profileContent);
      listOfContacts = JSON.parse(userContacts)
    } catch (ex) {
      return done(new Error('Failed to parse user profile'));
    }

    profile = json

    profile.contacts = parseContacts(listOfContacts.value)
    profile.provider = 'windowslive';

  } catch(err) {
    if (err.data) {
      try {
        let errorResp = JSON.parse(err.data);
      } catch (_) {}
    }
    if (json && json.error) {
      return done(new OutlookAPIError(json.error.message, json.error.code));
    }
    return done(new InternalOAuthError('Failed to fetch user profile', err));
  }

    done(null, profile);
  }


module.exports = Strategy;

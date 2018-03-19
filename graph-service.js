//==============================================================================
// Provides access to the Microsoft Graph API.
//==============================================================================
// Copyright (c) 2018 Frank Hellwig
//==============================================================================

'use strict';

//------------------------------------------------------------------------------
// Dependencies
//------------------------------------------------------------------------------

const ClientCredentials = require('client-credentials');
const HttpsService = require('https-service');
const async = require('async');
const util = require('util');

//------------------------------------------------------------------------------
// Initialization
//------------------------------------------------------------------------------

const NEW_ENDPOINT = 'https://graph.microsoft.com';
const OLD_ENDPOINT = 'https://graph.windows.net';

//https://graph.windows.net/{tenant_id}/{resource_path}?{api_version}[odata_query_parameters]

//------------------------------------------------------------------------------
// Public
//------------------------------------------------------------------------------

class GraphService extends HttpsService {
  constructor(tenant, clientId, clientSecret, apiVersion) {
    const endpoint = apiVersion ? OLD_ENDPOINT : NEW_ENDPOINT;
    super(endpoint);
    this.apiVersion = apiVersion;
    this.endpoint = endpoint;
    this.tenant = tenant;
    if (arguments.length > 2) {
      this.cred = new ClientCredentials(tenant, clientId, clientSecret);
    }
  }

  all(path, query) {
    return new Promise((resolve, reject) => {
      let results = [];
      async.doWhilst(
        callback => {
          this.get(path, query)
            .then(response => {
              if (response.type !== HttpsService.JSON_MEDIA_TYPE) {
                return callback(
                  new Error('Expected ' + HttpsService.JSON_MEDIA_TYPE + ' as the type.')
                );
              }
              if (!response.data || !Array.isArray(response.data.value)) {
                return callback(new Error('Expected an array body.value property.'));
              }
              results = results.concat(response.data.value);
              path = response.data['@odata.nextLink'] || response.data['odata.nextLink'] || null;
              query = null;
              callback(null, response);
            })
            .catch(err => {
              callback(err);
            });
        },
        _ => {
          return path !== null;
        },
        (err, response) => {
          if (err) return reject(err);
          response.data = results;
          resolve(response);
        }
      );
    });
  }

  request(method, path, headers, data, callback) {
    if (!this.cred) {
      throw new Error(
        'No client credentials. Either create a ' +
          'GraphService object with three arguments or provide your ' +
          'own access token and call the authorizedRequest method.'
      );
    }
    return this.cred.getAccessToken(this.endpoint, token => {
      return this.authorizedRequest(token, method, path, headers, data);
    });
  }

  authorizedRequest(token, method, path, headers, data) {
    headers = headers || {};
    headers['authorization'] = 'Bearer ' + token;
    let raw = path.endsWith('$value');
    if (!raw) {
      headers['accept'] = HttpsService.JSON_MEDIA_TYPE;
    }
    if (this.apiVersion) {
      let buf = ['/'];
      buf.push(this.tenant);
      if (!path.startsWith('/')) {
        buf.push('/');
      }
      buf.push(path);
      if (path.indexOf('?') < 0) {
        buf.push('?api-version=');
      } else {
        buf.push('&api-version=');
      }
      buf.push(this.apiVersion);
      path = buf.join('');
    }
    return super.request(method, path, headers, data);
  }
}

//------------------------------------------------------------------------------
// Exports
//------------------------------------------------------------------------------

module.exports = GraphService;

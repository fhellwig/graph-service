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
    let results = [];
    return this._all(path, query, results);
  }

  _all(path, query, results) {
    return this.get(path, query).then(response => {
      if (response.type !== HttpsService.JSON_MEDIA_TYPE) {
        throw new Error(`Expected ${HttpsService.JSON_MEDIA_TYPE} as the type.`);
      }
      const data = response.data;
      if (!data || !Array.isArray(data.value)) {
        throw new Error('Expected an array value property.');
      }
      results = results.concat(data.value);
      path = data['@odata.nextLink'] || data['odata.nextLink'] || null;
      if (path) {
        return this._all(path, null, results);
      } else {
        response.data = results;
        return response;
      }
    });
  }

  request(method, path, headers, data) {
    if (!this.cred) {
      throw new Error(
        'No client credentials. Either create a ' +
          'GraphService object with three arguments or provide your ' +
          'own access token and call the authorizedRequest method.'
      );
    }
    return this.cred.getAccessToken(this.endpoint).then(token => {
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

//==============================================================================
// Provides access to the Microsoft Graph API.
//==============================================================================
// Copyright (c) 2018 Frank Hellwig
//==============================================================================

'use strict';

const HttpsService = require('https-service');

const ENDPOINT = 'https://graph.microsoft.com';

class GraphService extends HttpsService {
  constructor(credentials, version) {
    super(ENDPOINT);
    if (typeof credentials !== 'object' || typeof credentials.getAccessToken !== 'function') {
      throw new Error('The credentials must be an object providing the getAccessToken method.');
    }
    this.credentials = credentials;
    this.version = version || null;
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
    return this.credentials.getAccessToken(ENDPOINT).then(token => {
      if (this.version) {
        path = this.version + path;
      }
      headers = headers || {};
      headers['authorization'] = 'Bearer ' + token;
      if (!path.endsWith('$value')) {
        headers['accept'] = HttpsService.JSON_MEDIA_TYPE;
      }
      return super.request(method, path, headers, data);
    });
  }
}

module.exports = GraphService;

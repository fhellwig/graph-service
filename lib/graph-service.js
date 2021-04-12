//==============================================================================
// Provides access to the Microsoft Graph API.
//==============================================================================
// Copyright (c) 2018 Frank Hellwig
//==============================================================================

'use strict';

const HttpsService = require('https-service');

const ENDPOINT = 'https://graph.microsoft.com';
const VERSION = 'v1.0';

class GraphService extends HttpsService {
  constructor(credentials, version = VERSION) {
    super(ENDPOINT);
    if (typeof credentials === 'string') {
      this.credentials = null;
      this.accessToken = credentials;
    } else if (
      typeof credentials === 'object' &&
      typeof credentials.getAccessToken === 'function'
    ) {
      this.credentials = credentials;
      this.accessToken = null;
    } else {
      throw new Error(
        'The credentials must be a string (an access token) or an object providing the getAccessToken method.'
      );
    }
    if (typeof version === 'string' && version.length > 0) {
      this.version = version;
    } else {
      throw new Error(
        'The version (if specified) must be a non-empty string such as v1.0 (the default value) or beta.'
      );
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
        // If the pagination url includes the hostname it is removed
        let match = `https://${this.host}/${this.version}`;
        if (path.startsWith(match)) {
          path = path.split(match)[1];
        }
        return this._all(path, null, results);
      } else {
        response.data = results;
        return response;
      }
    });
  }

  _request(token, method, path, headers, data) {
    // Don't add the version for @odata.nextLink paths as they are absolute URIs.
    if (path.startsWith('/')) {
      path = `/${this.version}${path}`;
    }
    headers = headers || {};
    headers['authorization'] = 'Bearer ' + token;
    if (!path.endsWith('$value')) {
      headers['accept'] = HttpsService.JSON_MEDIA_TYPE;
    }
    return super.request(method, path, headers, data);
  }

  request(method, path, headers, data) {
    if (this.credentials === null) {
      return this._request(this.accessToken, method, path, headers, data);
    } else {
      return this.credentials.getAccessToken(ENDPOINT).then(token => {
        return this._request(token, method, path, headers, data);
      });
    }
  }
}

module.exports = GraphService;

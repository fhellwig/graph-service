//==============================================================================
// Provides access to the Microsoft Graph API.
//==============================================================================
// Copyright (c) 2016 Frank Hellwig
//==============================================================================

'use strict';

//------------------------------------------------------------------------------
// Dependencies
//------------------------------------------------------------------------------

const ClientCredentials = require('client-credentials');
const HttpsService = require('https-service');
const async = require('async');

//------------------------------------------------------------------------------
// Initialization
//------------------------------------------------------------------------------

const RESOURCE = 'https://graph.microsoft.com';

//------------------------------------------------------------------------------
// Public
//------------------------------------------------------------------------------

class GraphService extends HttpsService {
    constructor(tenant, clientId, clientSecret) {
        super(RESOURCE);
        if (arguments.length === 3) {
            this.cred = new ClientCredentials(tenant, clientId, clientSecret);
        }
    }

    all(path, query, callback) {
        if (typeof query === 'function') {
            callback = query;
            query = null;
        }
        let results = [];
        async.doWhilst(
            callback => {
                this.get(path, query, (err, body, type, headers) => {
                    if (err) {
                        return callback(err);
                    }
                    if (type !== HttpsService.JSON_MEDIA_TYPE) {
                        return callback(new Error('Expected ' + HttpsService.JSON_MEDIA_TYPE + ' as the type.'));
                    }
                    if (!body || !util.isArray(body.value)) {
                        return callback(new Error('Expected an array body.value property.'));
                    }
                    results = results.concat(body.value);
                    path = body['@odata.nextLink'] || null;
                    query = null;
                    callback(null, results, type, headers);
                });
            },
            _ => {
                return path !== null;
            },
            callback
        );
    }

    request(method, path, headers, data, callback) {
        if (!this.cred) {
            throw new Error('No client credentials. Either create a ' +
                'GraphService object with three arguments or provide your ' +
                'own access token and call the authorizedRequest method.');
        }
        this.cred.getAccessToken(RESOURCE, (err, token) => {
            this.authorizedRequest(token, method, path, headers, data, callback);
        });
    }

    authorizedRequest(token, method, path, headers, data, callback) {
        headers = headers || {};
        headers['authorization'] = 'Bearer ' + token;
        let raw = path.endsWith('$value');
        if (!raw) {
            headers['accept'] = HttpsService.JSON_MEDIA_TYPE;
        }
        super.request(method, path, headers, data, callback);
    }
}

//------------------------------------------------------------------------------
// Exports
//------------------------------------------------------------------------------

module.exports = GraphService;

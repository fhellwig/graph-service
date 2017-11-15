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
                    path = body['@odata.nextLink'] || body['odata.nextLink'] || null;
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
        this.cred.getAccessToken(this.endpoint, (err, token) => {
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
        super.request(method, path, headers, data, callback);
    }
}

//------------------------------------------------------------------------------
// Exports
//------------------------------------------------------------------------------

module.exports = GraphService;

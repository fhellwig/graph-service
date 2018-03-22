# graph-service

A service for accessing the Microsoft Graph API.

Version 2.1.0

Exports the `GraphService` class allowing you to access the Microsoft Graph API at `https://graph.microsoft.com`. The `GraphService` class is subclassed from the [`HttpsService`](https://github.com/fhellwig/https-service) class. The `GraphService` class constructor requires a credentials object having a `getAccessToken` method to obtain the bearer token that is sent with each request. You can use a [`ClientCredentials`](https://github.com/fhellwig/client-credentials) instance or provide your own instance as long as it provides the `getAccessToken` method that returns a promise resolved with an access token.

The default Graph API version is `v1.0`. This can be overridded in the constructor (e.g., `beta`).

## 1. Installation

Install this package and, optionally, the [`client-credentials`](https://github.com/fhellwig/client-credentials) package.

```bash
$ npm install --save graph-service
$ npm install --save client-credentials
```

## 2. Usage

```javascript
const GraphService = require('graph-service');
const ClientCredentials = require('client-credentials');

const tenant = 'my-company.com';
const clientId = '0b13aa29-ca6b-42e8-a083-89e5bccdf141';
const clientSecret = 'lsl2isRe99Flsj32elwe89234ljhasd8239jsad2sl='

const credentials = new ClientCredentials(tenant, clientId, clientSecret);

const service = new GraphService(credentials)

service.all('/users').then(response => {
    console.log(response.data);
});
```

## 3. API

Since the `GraphService` class subclasses the [HttpsService](https://github.com/fhellwig/https-service) class, the API is identical to that class, with two exceptions. First, the constructor requires a credentials object (an object that provides the `getAccessToken` method). Second, the `all` method performs repeated `GET` requests, accumulating the results.

### 3.1 constructor

```javascript
GraphService(credentials, version)
```

Creates a new `GraphService` instance using the specified `credentials` object. This normally is an instance of the [`ClientCredentials`](https://github.com/fhellwig/client-credentials) class. It can also be an object you create, as long as it provides the `getAccessToken(resource)` method where the `resource` is always set to the `graph.microsoft.com` endpoint. This method must return a promise that is resolved with a valid access token. For example, if you have your own token from a user who has already authenticated with Azure AD, then you can create a simple object that returns this token in a promise. Note that creating a `GraphService` instance is not expensive (no network traffic takes place) so you can create a new instance for every new user request without any significant performance impact.

The version parameter defaults to the string `v1.0` and is prepended to all paths. For example, calling `service.get('/users')` will send the request to the `/v1.0/users` resource. You must include the initial slash on all paths since the path is created using the `/{version}{path}` construction.

### 3.1 all (method)

```javascript
all(path [, query])
```

Sends repeated `GET` requests to a resource that returns a list. This method accumulates the results from the `value` property and follows the `@odata.nextLink` property. Returns a promise that is resolved with the response from the last `HttpsService` GET request and the `data` property set to the concatenation of all of the retrieved objects.

## 4. License

The MIT License (MIT)

Copyright (c) 2018 Frank Hellwig

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

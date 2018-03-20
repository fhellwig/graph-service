# graph-service

A service for accessing the Microsoft Graph API.

Version 2.0.0

Exports the GraphService class. This class is subclassed from the `HttpsService` class (see [https-service](https://github.com/fhellwig/https-service)) and uses the `ClientCredentials` class (see [client-credentials](https://github.com/fhellwig/client-credentials)).

## 1. Installation

```bash
$ npm install --save graph-service
```

## 2. Usage

```javascript
const GraphService = require('graph-service');

const api = new GraphService('my-company.com', 'client-id', 'client-secret');

api.all('/v1.0/users').then(response => {
    console.log(response.data);
});
```

### 2.1 Endpoints

By default, this utility uses the `https://graph.microsoft.com` endpoint. You can specify the older `https://graph.windows.net` endpoint by adding a version number as the last argument in the constructor:

```javascript
const api = new GraphService('my-company.com', 'client-id', 'client-secret', '1.6');
```

This will use the older endpoint and add the `api-version` query parameter to all paths.

## 3. API

Since the `GraphService` class subclasses the [HttpsService](https://github.com/fhellwig/https-service) class, the API is identical as for that class, with three exceptions. First, the constructor accepts the same arguments as the [ClientCredentials](https://github.com/fhellwig/client-credentials) class. Second, there is an additional `authorizedRequest` method that accepts a token provided by the client instead of using the client credentials created in the constructor. Third, the `all` method performs repeated `GET` requests, accumulating the results.

### 3.1 constructor

```javascript
GraphService([tenant, clientId, clientSecret])
```

Creates a new `GraphService` instance for the specified `tenant`. The `clientId` and the `clientSecret` must be for an AAD application that has access rights to the Microsoft Graph resource (`https://graph.microsoft.com`). These three parameters are optional. If you have your own token, and wish to only call the `me` API resources, then you can call the constructor with no arguments. In that case, you must use the `authorizedRequest` method instead of the `request` method, passing in your own access token as the first parameter.

### 3.2 authorizedRequest

```javascript
authorizedRequest(token, method, path, headers, data, callback)
```

Sets the `Authorization` header with the bearer token and then calls the `request` method.

### 3.3 all

```javascript
all(path [, query])
```

Sends repeated `GET` requests to a resource that returns a list. This method accumulates the results from the `value` property and follows the `@odata.nextLink` property. Returns a promise that is resolved with the last response and the data set to all of the retrieved objects.

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

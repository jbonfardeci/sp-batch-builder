"use strict";
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (Object.hasOwnProperty.call(mod, k)) result[k] = mod[k];
    result["default"] = mod;
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
var $ = __importStar(require("jquery"));
var RestBatchResult = /** @class */ (function () {
    function RestBatchResult() {
        this.status = '';
        this.result = null;
    }
    return RestBatchResult;
}());
exports.RestBatchResult = RestBatchResult;
var BatchRequest = /** @class */ (function () {
    function BatchRequest(endpoint, payload, headers, verb, binary) {
        if (verb === void 0) { verb = 'POST'; }
        if (binary === void 0) { binary = false; }
        this.endpoint = endpoint;
        this.payload = payload;
        this.headers = headers;
        this.verb = verb;
        this.binary = binary;
        this.resultToken = '';
    }
    return BatchRequest;
}());
exports.BatchRequest = BatchRequest;
/**
 * Build and execute Batch requests for the SharePoint REST API.
 * Adapted and extended from https://github.com/SteveCurran/sp-rest-batch-execution/blob/master/RestBatchExecutor.js
 */
var SpBatchBuilder = /** @class */ (function () {
    function SpBatchBuilder(appWebUrl) {
        this.appWebUrl = appWebUrl;
        this.changeRequests = [];
        this.getRequests = [];
        this.resultsIndex = [];
    }
    // #region public methods.
    /**
     * Format list items endpoint.
     * @param siteUrl
     * @param listGuid
     * @param itemId
     */
    SpBatchBuilder.prototype.createListItemsUrl = function (siteUrl, listGuid, itemId) {
        siteUrl = /\/$/.test(siteUrl) ? siteUrl : siteUrl + '/';
        return siteUrl + "_api/web/lists(guid'" + listGuid + "')/items" + (itemId ? "(" + itemId + ")" : '');
    };
    /**
     * Add a GET request to be executed.
     * @param endpoint
     * @param headers
     */
    SpBatchBuilder.prototype.get = function (endpoint, headers) {
        var batchRequest = new BatchRequest(endpoint, null, headers, 'GET');
        this.loadRequest(batchRequest);
        return this;
    };
    /**
     * Add an INSERT request to be executed.
     * @param siteUrl
     * @param listGuid
     * @param payload
     * @param type
     */
    SpBatchBuilder.prototype.insert = function (siteUrl, listGuid, payload, type) {
        var endpoint = this.createListItemsUrl(siteUrl, listGuid);
        var data = $.extend(payload, { __metadata: { type: type } });
        var batchRequest = new BatchRequest(endpoint, data, null, 'POST');
        this.loadChangeRequest(batchRequest);
        return this;
    };
    /**
     * Add an UPDATE request to be executed.
     * @param siteUrl
     * @param listGuid
     * @param payload
     * @param type
     * @param etag
     */
    SpBatchBuilder.prototype.update = function (siteUrl, listGuid, payload, type, etag) {
        if (etag === void 0) { etag = '*'; }
        var endpoint = this.createListItemsUrl(siteUrl, listGuid, payload.Id);
        var data = $.extend(payload, { __metadata: { type: type } });
        var batchRequest = new BatchRequest(endpoint, data, { 'If-Match': etag }, 'MERGE');
        this.loadChangeRequest(batchRequest);
        return this;
    };
    /**
     * Add a DELETE request to be executed.
     * @param siteUrl
     * @param listGuid
     * @param itemId
     * @param etag
     */
    SpBatchBuilder.prototype.delete = function (siteUrl, listGuid, itemId, etag) {
        if (etag === void 0) { etag = '*'; }
        var endpoint = this.createListItemsUrl(siteUrl, listGuid, itemId);
        var batchRequest = new BatchRequest(endpoint, null, { 'If-Match': etag }, 'DELETE');
        this.loadChangeRequest(batchRequest);
        return this;
    };
    /**
     * Load a list item change request (POST, MEGRE, DELETE) into the batch collection to be sent to the server.
     * @param request
     */
    SpBatchBuilder.prototype.loadChangeRequest = function (request) {
        request.resultToken = this.getUniqueId();
        this.changeRequests.push($.extend({}, request));
        return request.resultToken;
    };
    /**
     * Load a GET list item request into the batch collection to be sent to the server.
     * @param request
     */
    SpBatchBuilder.prototype.loadRequest = function (request) {
        request.resultToken = this.getUniqueId();
        this.getRequests.push($.extend({}, request));
        return request.resultToken;
    };
    /**
     * Execute AJAX request.
     */
    SpBatchBuilder.prototype.executeAsync = function () {
        var dfd = $.Deferred();
        var payload = this.buildBatch();
        this.executeJQueryAsync(payload).done(function (result) {
            dfd.resolve(result);
        }).fail(function (err) {
            dfd.reject(err);
        });
        return dfd.promise();
    };
    // #endregion
    // #region private methods.
    /**
     * Get the ASP.NET form degist authentication token.
     * If doesn't exist on the .aspx page (or other) get a new one from the API.
     */
    SpBatchBuilder.prototype.getFormDigest = function () {
        var d = $.Deferred();
        var digest = document.querySelector('#__REQUESTDIGEST');
        if (!!(digest || { value: undefined }).value) {
            d.resolve(digest.value);
            return d.promise();
        }
        $.ajax({
            'url': this.appWebUrl + '_api/contextinfo',
            'method': 'POST',
            'headers': { 'Accept': 'application/json;odata=verbose' }
        }).done(function (digest) {
            d.resolve(digest.d.GetContextWebInformation.FormDigestValue);
        });
        return d.promise();
    };
    /**
     * Send the Batch body to be processed by the REST API.
     * @param batchBody
     */
    SpBatchBuilder.prototype.executeJQueryAsync = function (batchBody) {
        var self = this;
        var dfd = $.Deferred();
        var batchUrl = this.appWebUrl + "_api/$batch";
        this.getFormDigest().done(ajax);
        function ajax(digest) {
            var hdrs = {
                'accept': 'application/json;odata=verbose',
                'content-Type': 'multipart/mixed; boundary=batch_8890ae8a-f656-475b-a47b-d46e194fa574',
                'X-RequestDigest': digest
            };
            $.ajax({
                'url': batchUrl,
                'type': 'POST',
                'data': batchBody,
                'headers': hdrs,
                'success': function (data) {
                    var results = self.buildResults(data);
                    self.clearRequests();
                    dfd.resolve(results);
                },
                'error': function (err) {
                    self.clearRequests();
                    dfd.reject(err);
                }
            });
        }
        return dfd.promise();
    };
    SpBatchBuilder.prototype.getBatchRequestHeaders = function (headers, batchCommand) {
        var isAccept = false;
        if (headers) {
            $.each(Object.keys(headers), function (k, v) {
                batchCommand.push(v + ": " + headers[v]);
                if (!isAccept) {
                    isAccept = (v.toUpperCase() === "ACCEPT");
                }
            });
        }
        if (!isAccept) {
            batchCommand.push('accept:application/json;odata=verbose');
        }
    };
    /**
     * Build the batch body command.
     */
    SpBatchBuilder.prototype.buildBatch = function () {
        var self = this;
        var batchCommand = [];
        var batchBody;
        $.each(this.changeRequests, function (k, v) {
            self.buildBatchChangeRequest(batchCommand, v, k);
            self.resultsIndex.push(v.resultToken);
        });
        batchCommand.push("--changeset_f9c96a07-641a-4897-90ed-d285d2dbfc2e--");
        $.each(this.getRequests, function (k, v) {
            self.buildBatchGetRequest(batchCommand, v, k);
            self.resultsIndex.push(v.resultToken);
        });
        batchBody = batchCommand.join('\r\n');
        //embed all requests into one batch
        batchCommand = new Array();
        batchCommand.push("--batch_8890ae8a-f656-475b-a47b-d46e194fa574");
        batchCommand.push("Content-Type: multipart/mixed; boundary=changeset_f9c96a07-641a-4897-90ed-d285d2dbfc2e");
        batchCommand.push('Content-Length: ' + batchBody.length);
        batchCommand.push('Content-Transfer-Encoding: binary');
        batchCommand.push('');
        batchCommand.push(batchBody);
        batchCommand.push('');
        batchCommand.push("--batch_8890ae8a-f656-475b-a47b-d46e194fa574--");
        batchBody = batchCommand.join('\r\n');
        return batchBody;
    };
    SpBatchBuilder.prototype.buildBatchChangeRequest = function (batchCommand, request, batchIndex) {
        batchCommand.push("--changeset_f9c96a07-641a-4897-90ed-d285d2dbfc2e");
        batchCommand.push("Content-Type: application/http");
        batchCommand.push("Content-Transfer-Encoding: binary");
        batchCommand.push("Content-ID: " + (batchIndex + 1));
        batchCommand.push(request.binary ? "processData: false" : "processData: true");
        batchCommand.push('');
        batchCommand.push(request.verb.toUpperCase() + " " + request.endpoint + " HTTP/1.1");
        this.getBatchRequestHeaders(request.headers, batchCommand);
        if (!request.binary && request.payload) {
            batchCommand.push("Content-Type: application/json;odata=verbose");
        }
        if (request.binary && request.payload) {
            batchCommand.push("Content-Length :" + request.payload.byteLength);
        }
        batchCommand.push('');
        if (request.payload) {
            batchCommand.push(request.binary ? request.payload : JSON.stringify(request.payload));
            batchCommand.push('');
        }
    };
    SpBatchBuilder.prototype.buildBatchGetRequest = function (batchCommand, request, batchIndex) {
        batchCommand.push("--batch_8890ae8a-f656-475b-a47b-d46e194fa574");
        batchCommand.push('Content-Type: application/http');
        batchCommand.push('Content-Transfer-Encoding: binary');
        batchCommand.push("Content-ID: " + (batchIndex + 1));
        batchCommand.push('');
        batchCommand.push('GET ' + request.endpoint + ' HTTP/1.1');
        this.getBatchRequestHeaders(request.headers, batchCommand);
        batchCommand.push('');
    };
    SpBatchBuilder.prototype.buildResults = function (responseBody) {
        var self = this;
        var responseBoundary = responseBody.substring(0, 52);
        var resultTemp = responseBody.split(responseBoundary);
        var resultData = [];
        $.each(resultTemp, function (k, v) {
            if (v.indexOf('\r\nContent-Type: application/http\r\nContent-Transfer-Encoding: binary') == 0) {
                var responseTemp = v.split('\r\n');
                var batchResult = new RestBatchResult();
                //grab just the http status code
                batchResult.status = responseTemp[4].substr(9, 3);
                //based on the status pull the result from response
                batchResult.result = self.getResult(batchResult.status, responseTemp);
                //assign return token to result
                var result = { id: self.resultsIndex[k - 1], result: batchResult };
                resultData.push(result);
            }
        });
        return resultData;
    };
    SpBatchBuilder.prototype.getResult = function (status, response) {
        switch (status) {
            case "400":
            case "404":
            case "500":
            case "200":
                return this.parseJSON(response[7]);
            case "204":
            case "201":
                return this.parseJSON(response[9]);
            default:
                return this.parseJSON(response[4]);
        }
    };
    SpBatchBuilder.prototype.getUniqueId = function () {
        return (this.randomNum() + this.randomNum() + this.randomNum() + this.randomNum() + this.randomNum() + this.randomNum() + this.randomNum() + this.randomNum());
    };
    SpBatchBuilder.prototype.randomNum = function () {
        return (((1 + Math.random()) * 0x10000) | 0).toString(16).substring(1);
    };
    SpBatchBuilder.prototype.clearRequests = function () {
        while (!!this.changeRequests.length) {
            this.changeRequests.pop();
        }
        while (!!this.getRequests.length) {
            this.getRequests.pop();
        }
        while (!!this.resultsIndex.length) {
            this.resultsIndex.pop();
        }
    };
    ;
    SpBatchBuilder.prototype.parseJSON = function (jsonString) {
        try {
            var o = JSON.parse(jsonString);
            // Handle non-exception-throwing cases:
            // Neither JSON.parse(false) or JSON.parse(1234) throw errors, hence the type-checking,
            // but... JSON.parse(null) returns 'null', and typeof null === "object",
            // so we must check for that, too.
            if (o && typeof o === 'object' && o !== null) {
                return o;
            }
        }
        catch (e) { }
        return jsonString;
    };
    return SpBatchBuilder;
}());
exports.SpBatchBuilder = SpBatchBuilder;
//# sourceMappingURL=SpRestBatchBuilder.js.map
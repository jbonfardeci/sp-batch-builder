import * as $ from 'jquery';

export interface IBatchResult {
  id: number;
  result: RestBatchResult
}

export class RestBatchResult {
  public status: string = '';
  public result: any = null;
}

export class BatchRequest {
  public resultToken: string = '';

  constructor(public endpoint: string,
    public payload: any,
    public headers: any,
    public verb: string = 'POST',
    public binary: boolean = false) {

  }
}

/**
 * Build and execute Batch requests for the SharePoint REST API.
 * Adapted and extended from https://github.com/SteveCurran/sp-rest-batch-execution/blob/master/RestBatchExecutor.js
 */
export class SpRestBatchBuilder {

  private changeRequests: any[] = [];
  private getRequests: any[] = [];
  private resultsIndex: any[] = [];

  constructor(public appWebUrl: string){

  }

  // #region public methods.

  /**
   * Format list items endpoint.
   * @param siteUrl 
   * @param listGuid 
   * @param itemId 
   */
  public createListItemsUrl(siteUrl: string, listGuid: string, itemId?: number): string {
    siteUrl = /\/$/.test(siteUrl) ? siteUrl : siteUrl + '/';
    return `${siteUrl}_api/web/lists(guid'${listGuid}')/items` + (itemId ? `(${itemId})` : '');
  }

  /**
   * Add a GET request to be executed.
   * @param endpoint 
   * @param headers 
   */
  public get(endpoint: string, headers?: any): SpRestBatchBuilder {
    const batchRequest = new BatchRequest(endpoint, null, headers, 'GET');
    this.loadRequest(batchRequest);
    return this;
  }

  /**
   * Add an INSERT request to be executed.
   * @param siteUrl 
   * @param listGuid 
   * @param payload 
   * @param type 
   */
  public insert(siteUrl: string, listGuid: string, payload: any, type: string): SpRestBatchBuilder{
    const endpoint = this.createListItemsUrl(siteUrl, listGuid);
    const data = $.extend(payload, { __metadata: { type: type } });
    const batchRequest = new BatchRequest(endpoint, data, null, 'POST');
    this.loadChangeRequest(batchRequest);
    return this;
  }

  /**
   * Add an UPDATE request to be executed.
   * @param siteUrl 
   * @param listGuid 
   * @param payload 
   * @param type 
   * @param etag 
   */
  public update(siteUrl: string, listGuid: string, payload: any, type: string, etag: string = '*'): SpRestBatchBuilder {
    const endpoint = this.createListItemsUrl(siteUrl, listGuid, payload.Id);
    const data = $.extend(payload, { __metadata: { type: type } });
    const batchRequest = new BatchRequest(endpoint, data, {'If-Match': etag}, 'MERGE');
    this.loadChangeRequest(batchRequest);
    return this;
  }

  /**
   * Add a DELETE request to be executed.
   * @param siteUrl 
   * @param listGuid 
   * @param itemId 
   * @param etag 
   */
  public delete(siteUrl: string, listGuid: string, itemId: number, etag: string = '*'): SpRestBatchBuilder {
    const endpoint = this.createListItemsUrl(siteUrl, listGuid, itemId);
    const batchRequest = new BatchRequest(endpoint, null, {'If-Match': etag}, 'DELETE');
    this.loadChangeRequest(batchRequest);
    return this;
  }

  /**
   * Load a list item change request (POST, MEGRE, DELETE) into the batch collection to be sent to the server.
   * @param request
   */
  public loadChangeRequest(request: BatchRequest): string {
    request.resultToken = this.getUniqueId();
    this.changeRequests.push($.extend({}, request));
    return request.resultToken;
  }

  /**
   * Load a GET list item request into the batch collection to be sent to the server.
   * @param request
   */
  public loadRequest(request: BatchRequest): string {
    request.resultToken = this.getUniqueId();
    this.getRequests.push($.extend({}, request));
    return request.resultToken;
  }

  /**
   * Execute AJAX request.
   */
  public executeAsync(): JQueryPromise<any> {
    const dfd = $.Deferred();
    const payload = this.buildBatch();

    this.executeJQueryAsync(payload).done(function (result) {
      dfd.resolve(result);
    }).fail(function (err) {
      dfd.reject(err);
    });

    return dfd.promise();
  }

  // #endregion


  // #region private methods.

  /**
   * Get the ASP.NET form degist authentication token.
   * If doesn't exist on the .aspx page (or other) get a new one from the API.
   */
  private getFormDigest(): JQueryPromise<string> {
    const d = $.Deferred();
    const digest: HTMLInputElement = <HTMLInputElement>document.querySelector('#__REQUESTDIGEST');

    if(!!(digest || {value: undefined}).value) {
      d.resolve(digest.value);
      return d.promise();
    }

    $.ajax({
      'url': this.appWebUrl + '_api/contextinfo',
      'method': 'POST',
      'headers': { 'Accept': 'application/json;odata=verbose' }
    }).done((digest) => {
      d.resolve(digest.d.GetContextWebInformation.FormDigestValue);
    });

    return d.promise();
  }

  /**
   * Send the Batch body to be processed by the REST API.
   * @param batchBody
   */
  private executeJQueryAsync(batchBody: string): JQueryPromise<IBatchResult[]> {
    const self = this;
    const dfd = $.Deferred();
    const batchUrl = this.appWebUrl + "_api/$batch";

    this.getFormDigest().done(ajax);

    function ajax(digest: string) {
      let hdrs: any = {
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
          const results = self.buildResults(data);
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
  }

  private getBatchRequestHeaders(headers: any, batchCommand: string[]): void {
    let isAccept = false;
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
  }

  /**
   * Build the batch body command.
   */
  private buildBatch(): string {
    const self = this;
    let batchCommand: string[] = [];
    let batchBody: string;

    $.each(this.changeRequests, (k, v) => {
      self.buildBatchChangeRequest(batchCommand, v, k);
      self.resultsIndex.push(v.resultToken);
    });

    batchCommand.push("--changeset_f9c96a07-641a-4897-90ed-d285d2dbfc2e--");

    $.each(this.getRequests, (k, v) => {
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
  }

  private buildBatchChangeRequest(batchCommand: string[], request: any, batchIndex: number): void {
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
  }

  private buildBatchGetRequest(batchCommand: string[], request: any, batchIndex: number): void {
    batchCommand.push("--batch_8890ae8a-f656-475b-a47b-d46e194fa574");
    batchCommand.push('Content-Type: application/http');
    batchCommand.push('Content-Transfer-Encoding: binary');
    batchCommand.push("Content-ID: " + (batchIndex + 1));
    batchCommand.push('');
    batchCommand.push('GET ' + request.endpoint + ' HTTP/1.1');
    this.getBatchRequestHeaders(request.headers, batchCommand);
    batchCommand.push('');
  }

  private buildResults(responseBody: string): IBatchResult[] {
    const self = this;
    const responseBoundary = responseBody.substring(0, 52);
    const resultTemp = responseBody.split(responseBoundary);
    const resultData: any[] = [];

    $.each(resultTemp, function (k: number, v) {
      if (v.indexOf('\r\nContent-Type: application/http\r\nContent-Transfer-Encoding: binary') == 0) {
        const responseTemp = v.split('\r\n');
        const batchResult = new RestBatchResult();

        //grab just the http status code
        batchResult.status = responseTemp[4].substr(9, 3);

        //based on the status pull the result from response
        batchResult.result = self.getResult(batchResult.status, responseTemp);

        //assign return token to result
        const result: IBatchResult = { id: self.resultsIndex[k - 1], result: batchResult };
        resultData.push(result);
      }
    });

    return resultData;
  }

  private getResult(status: string, response: any): string {
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
  }

  private getUniqueId(): string {
    return (this.randomNum() + this.randomNum() + this.randomNum() + this.randomNum() + this.randomNum() + this.randomNum() + this.randomNum() + this.randomNum());
  }

  private randomNum(): string {
    return (((1 + Math.random()) * 0x10000) | 0).toString(16).substring(1);
  }

  private clearRequests(): void {
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

  private parseJSON(jsonString: string): any {
    try  {
      const o = JSON.parse(jsonString);

      // Handle non-exception-throwing cases:
      // Neither JSON.parse(false) or JSON.parse(1234) throw errors, hence the type-checking,
      // but... JSON.parse(null) returns 'null', and typeof null === "object",
      // so we must check for that, too.
      if (o && typeof o === 'object' && o !== null) {
        return o;
      }
    } catch (e) {}

    return jsonString;
  }

  // #endregion

}

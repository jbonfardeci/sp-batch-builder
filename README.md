# SharePoint Batch Builder
A utility to simplify building batch requests in SharePoint.

This utility was adapted and extended from https://github.com/SteveCurran/sp-rest-batch-execution/blob/master/RestBatchExecutor.js

Batch requests allow you to send multiple create/read/update/delete operations all with one request. While this SharePoint REST API 
greatly reduces network chatter, build a Batch request is not so straight forward.

A Batch request is sent in the body of a POST request even though you can send GET, POST, MERGE, and DELETE requests together.

From the batch body example below, it's easy to tell building a batch request is hard:

```(text)
--batch_8890ae8a-f656-475b-a47b-d46e194fa574
Content-Type: multipart/mixed; boundary=changeset_f9c96a07-641a-4897-90ed-d285d2dbfc2e
Content-Length: 1762
Content-Transfer-Encoding: binary

--changeset_f9c96a07-641a-4897-90ed-d285d2dbfc2e
Content-Type: application/http
Content-Transfer-Encoding: binary
Content-ID: 1
processData: true

POST https://<my-sp-site-url>/_api/web/lists(guid'<my-list-guid>')/items HTTP/1.1
accept:application/json;odata=verbose
Content-Type: application/json;odata=verbose

{"Title":"My Title 1","__metadata":{"type":"SP.Data.<SomeType>ListItem"}}

--changeset_f9c96a07-641a-4897-90ed-d285d2dbfc2e
Content-Type: application/http
Content-Transfer-Encoding: binary
Content-ID: 2
processData: true

POST https://<my-sp-site-url>/_api/web/lists(guid'<my-list-guid>')/items HTTP/1.1
accept:application/json;odata=verbose
Content-Type: application/json;odata=verbose

{"Title":"My Title 2","__metadata":{"type":"SP.Data.<SomeType>ListItem"}}

--changeset_f9c96a07-641a-4897-90ed-d285d2dbfc2e
Content-Type: application/http
Content-Transfer-Encoding: binary
Content-ID: 3
processData: true

DELETE https://<my-sp-site-url>/_api/web/lists(guid'<my-list-guid>')/items(25) HTTP/1.1
If-Match: "1"
accept:application/json;odata=verbose

--changeset_f9c96a07-641a-4897-90ed-d285d2dbfc2e
Content-Type: application/http
Content-Transfer-Encoding: binary
Content-ID: 4
processData: true

DELETE https://<my-sp-site-url>/_api/web/lists(guid'<my-list-guid>')/items(1) HTTP/1.1
If-Match: "2"
accept:application/json;odata=verbose

--changeset_f9c96a07-641a-4897-90ed-d285d2dbfc2e--

--batch_8890ae8a-f656-475b-a47b-d46e194fa574--
```

The Batch Builder utility greatly simplifies building a batch request. For example, sending 2 insert, 2 update, and 2 delete requests:

```(JavaScript)

const siteUrl = 'https://my-sharepoint-site.com/sites/my-site;
const listGuid = 'my-list-guid';
const batchExec = new SpRestBatchBuilder(siteUrl);
const camlListName = 'camlListName';
const listItemType = `SP.Data.${camlListName}ListItem`;

// New list items to insert.
const toInsert = [{Title: 'My Title 1'}, {Title: 'My Title 2'}];

// Existing list item values to update.
const toUpdate = [{Id: 1, Title: 'My Title 3', etag: '*'}, {Id: 2, Title: 'My Title 4', etag: '*'}];

// Existing list items to delete.
const toDelete = [{Id: 1, etag: '*'}, {Id: 2, etag: '*'}];

toInsert.forEach((item) => {
    batchExec.insert(siteUrl, listGuid, item, listItemType);
});

toUpdate.forEach((item) => {
    batchExec.update(siteUrl, listGuid, item, listItemType, item.etag);
});

toDelete.forEach((item) => {
    batchExec.delete(siteUrl, listGuid, item.Id, item.etag);
});

batchExec.executeAsync().done((result) => {
    console.info(result);
});

```
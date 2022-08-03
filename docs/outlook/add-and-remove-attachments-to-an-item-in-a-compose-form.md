---
title: Outlook アドインで添付ファイルを追加および削除する
description: さまざまな添付ファイル API を使用して、ユーザーが作成しているアイテムに添付されているファイルまたは Outlook アイテムを管理します。
ms.date: 08/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 23a1ce1a64d308f0ea51152726bf4d99d7a6300b
ms.sourcegitcommit: 143ab022c9ff6ba65bf20b34b5b3a5836d36744c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/03/2022
ms.locfileid: "67177687"
---
# <a name="manage-an-items-attachments-in-a-compose-form-in-outlook"></a>Outlook の作成フォームでアイテムの添付ファイルを管理する

Office JavaScript API には、ユーザーが作成するときにアイテムの添付ファイルを管理するために使用できるいくつかの API が用意されています。

## <a name="attach-a-file-or-outlook-item"></a>ファイルまたは Outlook アイテムを添付する

添付ファイルの種類に適したメソッドを使用して、ファイルまたは Outlook アイテムを作成フォームに添付できます。

- [addFileAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods): ファイルを添付する
- [addFileAttachmentFromBase64Async: base64](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) 文字列を使用してファイルを添付する
- [addItemAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods): Outlook アイテムを添付する

これらは非同期メソッドです。つまり、アクションの完了を待たずに実行を行うことができます。 追加する添付ファイルの元の場所とサイズによっては、非同期呼び出しの完了に時間がかかる場合があります。

完了するアクションに依存するタスクがある場合は、それらのタスクをコールバック関数で実行する必要があります。 このコールバック関数は省略可能であり、添付ファイルのアップロードが完了したときに呼び出されます。 コールバック関数は [、AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトを出力パラメーターとして受け取り、添付ファイルの追加から状態、エラー、および戻り値を提供します。 コールバックがその他のパラメーターを必要とする場合、オプションの `options.asyncContext` パラメーターでそれを指定することができます。 `options.asyncContext` は、コールバック関数が想定する任意の型にすることができます。

たとえば、1 つ以上のキーと値のペアを含む JSON オブジェクトとして定義 `options.asyncContext` できます。 Office アドインの非同期プログラミングの Office アドイン プラットフォームでは、省略可能なパラメーターを非同期メソッドに渡す方法について詳しく説明 [します](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-to-asynchronous-methods)。次の例では、パラメーターを使用 `asyncContext` してコールバック関数に 2 つの引数を渡す方法を示します。

```js
const options = { asyncContext: { var1: 1, var2: 2}};

Office.context.mailbox.item.addFileAttachmentAsync('https://contoso.com/rtm/icon.png', 'icon.png', options, callback);
```

オブジェクトのプロパティを使用して `status` 、コールバック関数で非同期メソッド呼び出しの成功または `error` エラーを `AsyncResult` 確認できます。 添付が正常に完了した場合は、プロパティを `AsyncResult.value` 使用して添付ファイル ID を取得できます。 添付ファイル ID は整数であり、後で添付ファイルを削除するときに使用できます。

> [!NOTE]
> 添付ファイル ID は同じセッション内でのみ有効であり、セッション間で同じ添付ファイルにマップすることは保証されません。 セッションが終了した場合の例には、ユーザーがアドインを閉じた場合や、ユーザーがインライン フォームで作成を開始し、その後インライン フォームをポップアップして別のウィンドウで続行する場合などがあります。

### <a name="attach-a-file"></a>ファイルを添付する

メソッドを使用 `addFileAttachmentAsync` し、ファイルの URI を指定することで、作成フォームのメッセージまたは予定にファイルを添付できます。 メソッドを `addFileAttachmentFromBase64Async` 使用することもできますが、base64 文字列を入力として指定することもできます。 If the file is protected, you can include an appropriate identity or authentication token as a URI query string parameter. Exchange will make a call to the URI to get the attachment, and the web service which protects the file will need to use the token as a means of authentication.

次の JavaScript の例は、Web サーバーのファイル picture.png を作成中のメッセージまたは予定に添付する新規作成アドインです。 コールバック関数はパラメーターとして受け取り `asyncResult` 、結果の状態を確認し、メソッドが成功した場合に添付ファイル ID を取得します。

```js
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add the specified file attachment to the item
        // being composed.
        // When the attachment finishes uploading, the
        // callback function is invoked and gets the attachment ID.
        // You can optionally pass any object that you would
        // access in the callback function as an argument to
        // the asyncContext parameter.
        Office.context.mailbox.item.addFileAttachmentAsync(
            `https://webserver/picture.png`,
            'picture.png',
            { asyncContext: null },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                } else {
                    // Get the ID of the attached file.
                    const attachmentID = asyncResult.value;
                    write('ID of added attachment: ' + attachmentID);
                }
            });
    });
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

インライン base64 イメージを構成するメッセージの本文に追加するには、メソッドを使用してイメージを挿入する前に、まずメソッドを `Office.context.mailbox.item.body.getAsync` 使用して現在のメッセージ本文を `addFileAttachmentFromBase64Async` 取得する必要があります。 それ以外の場合、イメージは挿入後にメッセージにレンダリングされません。 ガイダンスについては、次の JavaScript の例を参照してください。これは、メッセージ本文の先頭にインライン base64 イメージを追加します。

```js
const mailItem = Office.context.mailbox.item;
const base64String =
  "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAnUExURQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAN0S+bUAAAAMdFJOUwAQIDBAUI+fr7/P7yEupu8AAAAJcEhZcwAADsMAAA7DAcdvqGQAAAF8SURBVGhD7dfLdoMwDEVR6Cspzf9/b20QYOthS5Zn0Z2kVdY6O2WULrFYLBaLxd5ur4mDZD14b8ogWS/dtxV+dmx9ysA2QUj9TQRWv5D7HyKwuIW9n0vc8tkpHP0W4BOg3wQ8wtlvA+PC1e8Ao8Ld7wFjQtHvAiNC2e8DdqHqKwCrUPc1gE1AfRVgEXBfB+gF0lcCWoH2tYBOYPpqQCNwfT3QF9i+AegJfN8CtAWhbwJagtS3AbIg9o2AJMh9M5C+SVGBvx6zAfmT0r+Bv8JMwP4kyFPir+cswF5KL3WLv14zAFBCLf56Tw9cparFX4upgaJUtPhrOS1QlY5W+vWTXrGgBFB/b72ev3/0igUdQPppP/nfowfKUUEFcP207y/yxKmgAYQ+PywoAFOfCH3A2MdCFzD3kdADBvq10AGG+pXQBgb7pdAEhvuF0AIc/VtoAK7+JciAs38KIuDugyAC/v4hiMCE/i7IwLRBsh68N2WQjMVisVgs9i5bln8LGScNcCrONQAAAABJRU5ErkJggg==";

// Get the current body of the message.
mailItem.body.getAsync(Office.CoercionType.Html, (bodyResult) => {
  if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
    // Insert the base64 image to the beginning of the message body.
    const options = { isInline: true, asyncContext: bodyResult.value };
    mailItem.addFileAttachmentFromBase64Async(base64String, "sample.png", options, (attachResult) => {
      if (attachResult.status === Office.AsyncResultStatus.Succeeded) {
        let body = attachResult.asyncContext;
        body = body.replace("<p class=MsoNormal>", `<p class=MsoNormal><img src="cid:sample.png">`);
        mailItem.body.setAsync(body, { coercionType: Office.CoercionType.Html }, (setResult) => {
          if (setResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Inline base64 image added to message.");
          } else {
            console.log(setResult.error.message);
          }
        });
      } else {
        console.log(attachResult.error.message);
      }
    });
  } else {
    console.log(bodyResult.error.message);
  }
});
```

### <a name="attach-an-outlook-item"></a>Outlook アイテムを添付する

アイテムの Exchange Web Services (EWS) ID を指定し、メソッドを使用 `addItemAttachmentAsync` して、作成フォームのメッセージまたは予定に Outlook アイテム (電子メール、予定表、連絡先アイテムなど) を添付できます。 [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) メソッドを使用して EWS 操作 [FindItem](/exchange/client-developer/web-service-reference/finditem-operation) にアクセスすることで、ユーザーのメールボックス内の電子メール、予定表、連絡先、またはタスク アイテムの EWS ID を取得できます。 閲覧フォームの既存のアイテムでは、 [item.itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) プロパティでも EWS ID が取得できます。

次の JavaScript 関数は、 `addItemAttachment`上記の最初の例を拡張し、構成されている電子メールまたは予定にアイテムを添付ファイルとして追加します。 この関数は、添付するアイテムの EWS ID を引数として受け取ります。 アタッチが成功すると、同じセッションでその添付ファイルを削除するなど、さらに処理するための添付ファイル ID が取得されます。

```js
// Adds the specified item as an attachment to the composed item.
// ID is the EWS ID of the item to be attached.
function addItemAttachment(itemId) {
    // When the attachment finishes uploading, the
    // callback function is invoked. Here, the callback
    // function uses only asyncResult as a parameter,
    // and if the attaching succeeds, gets the attachment ID.
    // You can optionally pass any other object you wish to
    // access in the callback function as an argument to
    // the asyncContext parameter.
    Office.context.mailbox.item.addItemAttachmentAsync(
        itemId,
        'Welcome email',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            } else {
                const attachmentID = asyncResult.value;
                write('ID of added attachment: ' + attachmentID);
            }
        });
}
```

> [!NOTE]
> 作成アドインを使用して、Outlook on the webまたはモバイル デバイスで定期的な予定のインスタンスをアタッチできます。 ただし、サポートされている Outlook デスクトップ クライアントでは、インスタンスをアタッチしようとすると、定期的な系列 (親予定) がアタッチされます。

## <a name="get-attachments"></a>添付ファイルを取得する

作成モードで添付ファイルを取得する API は、 [要件セット 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8) から入手できます。

- [getAttachmentsAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [getAttachmentContentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)

[getAttachmentsAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) メソッドを使用して、構成されているメッセージまたは予定の添付ファイルを取得できます。

添付ファイルのコンテンツを取得するには、 [getAttachmentContentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) メソッドを使用します。 サポートされている形式は、 [AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat) 列挙型に一覧表示されます。

出力パラメーター オブジェクトを使用して状態とエラーを確認するコールバック関数を指定する `AsyncResult` 必要があります。 省略可能 `asyncContext` なパラメーターを使用して、追加のパラメーターをコールバック関数に渡すこともできます。

次の JavaScript の例では、添付ファイルを取得し、サポートされている添付ファイル形式ごとに個別の処理を設定できます。

```js
const item = Office.context.mailbox.item;
const options = {asyncContext: {currentItem: item}};
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (let i = 0 ; i < result.value.length ; i++) {
      result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      // Handle file attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      // Handle email item attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      // Handle .icalender attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Handle cloud attachment.
      break;
    default:
      // Handle attachment formats that are not supported.
  }
}
```

## <a name="remove-an-attachment"></a>添付ファイルを削除する

[removeAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) メソッドを使用するときに、対応する添付ファイル ID を指定することで、作成フォームのメッセージまたは予定アイテムからファイルまたはアイテムの添付ファイルを削除できます。

> [!IMPORTANT]
> 要件セット 1.7 以前を使用している場合は、同じアドインが同じセッションに追加した添付ファイルのみを削除する必要があります。

`addFileAttachmentAsync`、`addItemAttachmentAsync`およびメソッドと`getAttachmentsAsync`同様に、`removeAttachmentAsync`非同期メソッドです。 出力パラメーター オブジェクトを使用して状態とエラーを確認するコールバック関数を指定する `AsyncResult` 必要があります。 省略可能 `asyncContext` なパラメーターを使用して、追加のパラメーターをコールバック関数に渡すこともできます。

次の JavaScript 関数は、 `removeAttachment`上記の例を引き続き拡張し、構成されている電子メールまたは予定から指定された添付ファイルを削除します。 この関数は、削除する添付ファイルの ID を引数として受け取ります。 添付ファイルの ID は、成功した `addFileAttachmentAsync``addFileAttachmentFromBase64Async`、または`addItemAttachmentAsync`メソッド呼び出しの後に取得し、後続`removeAttachmentAsync`のメソッド呼び出しで使用できます。 また、(要件セット 1.8 で導入された) を呼び出 `getAttachmentsAsync` して、そのアドイン セッションの添付ファイルとその ID を取得することもできます。

```js
// Removes the specified attachment from the composed item.
function removeAttachment(attachmentId) {
    // When the attachment is removed, the callback function is invoked.
    // Here, the callback function uses an asyncResult parameter and
    // gets the ID of the removed attachment if the removal succeeds.
    // You can optionally pass any object you wish to access in the
    // callback function as an argument to the asyncContext parameter.
    Office.context.mailbox.item.removeAttachmentAsync(
        attachmentId,
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                write(asyncResult.error.message);
            } else {
                write('Removed attachment with the ID: ' + asyncResult.value);
            }
        });
}
```

## <a name="see-also"></a>関連項目

- [新規作成フォーム用の Outlook アドインを作成する](compose-scenario.md)
- [Office アドインにおける非同期プログラミング](../develop/asynchronous-programming-in-office-add-ins.md)

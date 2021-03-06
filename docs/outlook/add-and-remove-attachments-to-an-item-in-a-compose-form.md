---
title: Outlook アドインで添付ファイルを追加および削除する
description: さまざまな添付ファイル API を使用して、ユーザーが作成しているアイテムに添付されているファイルまたは Outlook アイテムを管理できます。
ms.date: 02/24/2021
localization_priority: Normal
ms.openlocfilehash: da426813e865f5607ec3e2c65252e8a406d889e2
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505501"
---
# <a name="manage-an-items-attachments-in-a-compose-form-in-outlook"></a>Outlook の作成フォームでアイテムの添付ファイルを管理する

JavaScript API Officeには、ユーザーが作成するときにアイテムの添付ファイルを管理するために使用できるいくつかの API が提供されています。

## <a name="attach-a-file-or-outlook-item"></a>ファイルまたは Outlook アイテムを添付する

添付ファイルの種類に適したメソッドを使用して、ファイルまたは Outlook アイテムを作成フォームに添付できます。

- [addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): ファイルを添付する
- [addFileAttachmentFromBase64Async : base64](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)文字列を使用してファイルを添付する
- [addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Outlook アイテムの添付

これらは非同期メソッドです。つまり、アクションが完了するのを待たずに実行を続けできます。 追加する添付ファイルの元の場所とサイズによっては、非同期呼び出しの完了に時間がかかる場合があります。

アクションの完了に依存するようなタスクがある場合、それらのタスクはコールバック メソッドで実行する必要があります。 このコールバック メソッドはオプションで、添付ファイルのアップロードが完了すると呼び出されます。 コールバック メソッドは、状態、エラー、そして添付ファイル追加によって返される値を提供する出力パラメーターとして、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトを使用します。 コールバックがその他のパラメーターを必要とする場合、オプションの `options.asyncContext` パラメーターでそれを指定することができます。 `options.asyncContext` は、コールバック メソッドが予期する任意の種類となることができます。

たとえば、`options.asyncContext` は 1 つ以上の「キーと値のペア」を含む JSON オブジェクトとして定義することができます。非同期メソッドにオプション パラメーターを渡すさらに多くの例は、「[Office アドインにおける非同期プログラミング](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)」の Office アドイン プラットフォームで見いだすことができます。以下の例は、コールバック メソッドに引数を 2 つ渡すための `asyncContext` パラメーターの使用法を示しています。

```js
var options = { asyncContext: { var1: 1, var2: 2}};

Office.context.mailbox.item.addFileAttachmentAsync('https://contoso.com/rtm/icon.png', 'icon.png', options, callback);
```

非同期メソッド呼び出しの正常終了またはエラーの確認は、コールバック メソッドにおいて `AsyncResult` オブジェクトの `status` と `error` のプロパティを使用して行うことができます。添付が成功裏に完了した場合、`AsyncResult.value` プロパティを使用して添付ファイル ID を取得することができます。添付ファイル ID は整数で、その後、添付ファイルを削除する場合に使用できます。

> [!NOTE]
> 添付ファイル ID は同じセッション内でのみ有効であり、セッション間で同じ添付ファイルにマップされる保証はありません。 セッションが終了した場合の例としては、ユーザーがアドインを閉じる場合、またはユーザーがインライン フォームで作成を開始し、その後インライン フォームをポップアウトして別のウィンドウで続行する場合があります。

### <a name="attach-a-file"></a>ファイルの添付

作成フォームのメッセージまたは予定にファイルを添付するには、メソッドを使用してファイル `addFileAttachmentAsync` の URI を指定します。 メソッドを使用できますが `addFileAttachmentFromBase64Async` 、base64 文字列を入力として指定できます。 If the file is protected, you can include an appropriate identity or authentication token as a URI query string parameter. Exchange will make a call to the URI to get the attachment, and the web service which protects the file will need to use the token as a means of authentication.

次の JavaScript 例は、picture.png ファイルを web サーバーから取得して新規作成中のメッセージあるいは予定に添付する新規作成アドインです。コールバック メソッドはパラメーターとして `asyncResult` を使用し、結果の状態を確認し、メソッドが成功した場合に添付ファイル ID を取得します。

```js
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add the specified file attachment to the item
        // being composed.
        // When the attachment finishes uploading, the
        // callback method is invoked and gets the attachment ID.
        // You can optionally pass any object that you would
        // access in the callback method as an argument to
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
                    var attachmentID = asyncResult.value;
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

### <a name="attach-an-outlook-item"></a>Outlook アイテムを添付する

アイテムの Exchange Web Services (EWS) ID を指定し、メソッドを使用して、作成フォーム内のメッセージまたは予定に Outlook アイテム (電子メール、予定表、連絡先アイテムなど) を添付できます。 `addItemAttachmentAsync` [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)メソッドを使用して EWS 操作[FindItem](/exchange/client-developer/web-service-reference/finditem-operation)にアクセスすると、ユーザーのメールボックス内の電子メール、予定表、連絡先、またはタスク アイテムの EWS ID を取得できます。 閲覧フォームの既存のアイテムでは、 [item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティでも EWS ID が取得できます。

次の JavaScript 関数は、上記の最初の例を拡張し、構成されている電子メールまたは予定に添付ファイルとしてアイテム `addItemAttachment` を追加します。 この関数は、添付するアイテムの EWS ID を引数として受け取ります。 接続が成功した場合は、同じセッションで添付ファイルを削除するなど、さらに処理するための添付ファイル ID を取得します。

```js
// Adds the specified item as an attachment to the composed item.
// ID is the EWS ID of the item to be attached.
function addItemAttachment(itemId) {
    // When the attachment finishes uploading, the
    // callback method is invoked. Here, the callback
    // method uses only asyncResult as a parameter,
    // and if the attaching succeeds, gets the attachment ID.
    // You can optionally pass any other object you wish to
    // access in the callback method as an argument to
    // the asyncContext parameter.
    Office.context.mailbox.item.addItemAttachmentAsync(
        itemId,
        'Welcome email',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            } else {
                var attachmentID = asyncResult.value;
                write('ID of added attachment: ' + attachmentID);
            }
        });
}
```

> [!NOTE]
> 作成アドインを使用して、Outlook on the web またはモバイル デバイスで定期的な予定のインスタンスを添付できます。 ただし、サポートしている Outlook デスクトップ クライアントでは、インスタンスを接続しようとすると、定期的な系列 (親予定) が添付されます。

## <a name="get-attachments"></a>添付ファイルを取得する

作成モードで添付ファイルを取得する API は、要件セット [1.8 から利用できます](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)。

- [getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [getAttachmentContentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

[getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)メソッドを使用して、構成されているメッセージまたは予定の添付ファイルを取得できます。

添付ファイルのコンテンツを取得するには [、getAttachmentContentAsync メソッドを使用](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) できます。 サポートされている形式は [、AttachmentContentFormat 列挙型に一覧表示](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat) されます。

output パラメーター オブジェクトを使用して、状態とエラーを確認するコールバック メソッドを `AsyncResult` 指定する必要があります。 省略可能なパラメーターを使用して、コールバック メソッドに追加のパラメーターを渡 `asyncContext` することもできます。

次の JavaScript の例では、添付ファイルを取得し、サポートされている添付ファイル形式ごとに個別の処理を設定できます。

```js
var item = Office.context.mailbox.item;
var options = {asyncContext: {currentItem: item}};
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
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

## <a name="remove-an-attachment"></a>添付ファイルの削除

[removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)メソッドを使用する場合は、対応する添付ファイル ID を指定して、作成フォームのメッセージまたは予定アイテムからファイルまたはアイテムの添付ファイルを削除できます。

> [!IMPORTANT]
> 要件セット 1.7 以前を使用している場合は、同じアドインが同じセッションで追加した添付ファイルのみを削除する必要があります。

、 および `addFileAttachmentAsync` メソッド `addItemAttachmentAsync` と `getAttachmentsAsync` 同様に、 `removeAttachmentAsync` 非同期メソッドです。 output パラメーター オブジェクトを使用して、状態とエラーを確認するコールバック メソッドを `AsyncResult` 指定する必要があります。 省略可能なパラメーターを使用して、コールバック メソッドに追加のパラメーターを渡 `asyncContext` することもできます。

次の JavaScript 関数は、上記の例を引き続き拡張し、構成されている電子メールまたは予定から指定された添付ファイル `removeAttachment` を削除します。 この関数は、削除する添付ファイルの ID を引数として受け取ります。 添付ファイルの ID は、成功した 、またはメソッドの呼び出し後に取得し、後続のメソッド呼び出 `addFileAttachmentAsync` `addFileAttachmentFromBase64Async` `addItemAttachmentAsync` しで `removeAttachmentAsync` 使用できます。 (要件セット 1.8 で導入) を呼び出して、そのアドイン セッションの添付ファイルとその `getAttachmentsAsync` ID を取得することもできます。

```js
// Removes the specified attachment from the composed item.
function removeAttachment(attachmentId) {
    // When the attachment is removed, the callback method is invoked.
    // Here, the callback method uses an asyncResult parameter and
    // gets the ID of the removed attachment if the removal succeeds.
    // You can optionally pass any object you wish to access in the
    // callback method as an argument to the asyncContext parameter.
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

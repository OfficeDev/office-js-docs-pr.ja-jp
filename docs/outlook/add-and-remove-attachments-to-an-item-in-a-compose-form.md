---
title: Outlook アドインで添付ファイルを追加および削除する
description: さまざまな添付ファイル Api を使用して、ユーザーが作成しているアイテムに添付されたファイルまたは Outlook アイテムを管理できます。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 977b8fa814a251c76aabc64345762a3a9556a60b
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166600"
---
# <a name="manage-an-items-attachments-in-a-compose-form-in-outlook"></a>Outlook で新規作成フォーム内のアイテムの添付ファイルを管理する

JavaScript API for Office には、ユーザーの作成時にアイテムの添付ファイルを管理するために使用できるいくつかの Api が用意されています。

## <a name="attach-a-file-or-outlook-item"></a>ファイルまたは Outlook アイテムを添付する

添付ファイルの種類に適した方法を使用して、ファイルまたは Outlook アイテムを新規作成フォームに添付できます。

- [Addfileattachmentasync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): ファイルの添付
- [addFileAttachmentFromBase64Async](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): base64 文字列を使用してファイルを添付します。
- [Additemattachmentasync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Outlook アイテムを添付する

これらの非同期メソッドは、アクションが完了するまで待機しなくても実行が可能であることを意味します。 追加する添付ファイルの元の場所とサイズによっては、非同期呼び出しの完了に時間がかかる場合があります。

アクションの完了に依存するようなタスクがある場合、それらのタスクはコールバック メソッドで実行する必要があります。 このコールバックメソッドはオプションであり、添付ファイルのアップロードが完了したときに呼び出されます。 コールバック メソッドは、状態、エラー、そして添付ファイル追加によって返される値を提供する出力パラメーターとして、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトを使用します。 コールバックがその他のパラメーターを必要とする場合、オプションの `options.asyncContext` パラメーターでそれを指定することができます。 `options.asyncContext` は、コールバック メソッドが予期する任意の種類となることができます。

たとえば、`options.asyncContext` は 1 つ以上の「キーと値のペア」を含む JSON オブジェクトとして定義することができます。非同期メソッドにオプション パラメーターを渡すさらに多くの例は、「[Office アドインにおける非同期プログラミング](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)」の Office アドイン プラットフォームで見いだすことができます。以下の例は、コールバック メソッドに引数を 2 つ渡すための `asyncContext` パラメーターの使用法を示しています。

```js
var options = { asyncContext: { var1: 1, var2: 2}};

Office.context.mailbox.item.addFileAttachmentAsync('https://contoso.com/rtm/icon.png', 'icon.png', options, callback);
```

非同期メソッド呼び出しの正常終了またはエラーの確認は、コールバック メソッドにおいて `AsyncResult` オブジェクトの `status` と `error` のプロパティを使用して行うことができます。添付が成功裏に完了した場合、`AsyncResult.value` プロパティを使用して添付ファイル ID を取得することができます。添付ファイル ID は整数で、その後、添付ファイルを削除する場合に使用できます。

> [!NOTE]
> ベスト プラクティスとしては、この添付ファイル ID による添付ファイルの削除は、同じアドインが同じセッション内でその添付ファイルを追加した場合のみ使用すべきです。 Outlook on the web およびモバイルデバイスでは、添付ファイル ID は同じセッション内でのみ有効です。 ユーザーがアドインを閉じたとき、またはユーザーがインライン フォームで作成を開始した後インライン フォームから出て別のウィンドウで作業を続行したとき、セッションは終了します。

### <a name="attach-a-file"></a>ファイルを添付する

`addFileAttachmentAsync`メソッドを使用してファイルの URI を指定することによって、新規作成フォーム内のメッセージまたは予定にファイルを添付することができます。 `addFileAttachmentFromBase64Async`メソッドを使用することもできますが、入力として base64 文字列を指定することもできます。 If the file is protected, you can include an appropriate identity or authentication token as a URI query string parameter. Exchange will make a call to the URI to get the attachment, and the web service which protects the file will need to use the token as a means of authentication.

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
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
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

新規作成フォームで、メッセージまたは予定に Outlook アイテム (メール、予定表、連絡先のアイテムなど) を添付するには、そのアイテムの Exchange Web Service (EWS) の ID を指定して `addItemAttachmentAsync` メソッドを使用します。[mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッドを使用し、EWS 操作である [FindItem](/exchange/client-developer/web-service-reference/finditem-operation) にアクセスすることにより、ユーザーのメールボックス内の電子メール、予定表、連絡先、タスク アイテムの EWS ID を取得することができます。[item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティはまた、閲覧フォーム内の既存アイテムの EWS ID も提供します。

次の JavaScript 関数`addItemAttachment`は、上の最初の例を拡張し、作成中の電子メールまたは予定の添付ファイルとしてアイテムを追加します。 この関数は、添付するアイテムの EWS ID を引数として受け取ります。 接続が成功した場合は、同じセッションでその添付ファイルを削除するなど、追加の処理のための添付ファイル ID を取得します。

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
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                var attachmentID = asyncResult.value;
                write('ID of added attachment: ' + attachmentID);
            }
        });
}
```

> [!NOTE]
> 新規作成アドインを使用すると、Outlook on the web またはモバイルデバイスで定期的な予定のインスタンスを添付することができます。 ただし、サポート Outlook リッチ クライアントでは、インスタンスを 1 つ添付しようとしても、定期的な一連の予定 (マスター予定) が添付されます。

## <a name="get-attachments"></a>添付ファイルを取得する

[GetAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)メソッドを使用すると、構成されているメッセージまたは予定の添付ファイルを取得できます。

添付ファイルのコンテンツを取得するには、 [Getattachmentcontentasync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)メソッドを使用します。 サポートされている形式は、 [Attachmentcontentformat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)列挙体に一覧表示されています。

`AsyncResult`出力パラメータオブジェクトを使用して、ステータスとエラーを確認するコールバックメソッドを提供する必要があります。 オプション`asyncContext`のパラメーターを使用して、コールバックメソッドに追加のパラメーターを渡すこともできます。

次の JavaScript の例では、添付ファイルを取得し、サポートされている添付ファイルの形式ごとに個別の処理をセットアップできるようにします。

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

## <a name="remove-an-attachment"></a>添付ファイルを削除する

新規作成フォームのメッセージまたは予定アイテムからファイルまたはアイテムの添付ファイルを削除するには、対応する添付ファイル ID を指定し、 [Removeattachmentasync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)メソッドを使用します。 同じアドインが同じセッションで追加した添付ファイルのみを削除する必要があります。 `addFileAttachmentAsync`およびメソッドと`addItemAttachmentAsync`同様に、 `removeAttachmentAsync`は非同期メソッドです。 `AsyncResult`出力パラメータオブジェクトを使用して、ステータスとエラーを確認するコールバックメソッドを提供する必要があります。 オプション`asyncContext`のパラメーターを使用して、コールバックメソッドに追加のパラメーターを渡すこともできます。

次の JavaScript 関数`removeAttachment`は、上記の例を引き続き拡張し、作成されている電子メールまたは予定から指定された添付ファイルを削除します。 この関数は、削除する添付ファイルの ID を引数として受け取ります。 添付ファイルの ID を取得するには、成功`addFileAttachmentAsync` `addFileAttachmentFromBase64Async`した、 `addItemAttachmentAsync`またはメソッドが呼び出された後に`removeAttachmentAsync` 、以降のメソッド呼び出しに対して保存します。

```js
// Removes the specified attachment from the composed item.
// ID is the Exchange identifier of the attachment to be
// removed.
function removeAttachment(attachmentId) {
    // When the attachment is removed, the
    // callback method is invoked. Here, the callback
    // method uses an asyncResult parameter and gets
    // the ID of the removed attachment if the removal
    // succeeds.
    // You can optionally pass any object you wish to
    // access in the callback method as an argument to
    // the asyncContext parameter.
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

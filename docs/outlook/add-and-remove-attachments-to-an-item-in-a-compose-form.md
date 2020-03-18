---
title: Outlook アドインで添付ファイルを追加および削除する
description: さまざまな添付ファイル Api を使用して、ユーザーが作成しているアイテムに添付されたファイルまたは Outlook アイテムを管理できます。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: bb966ff80bae37fbaa781b5a428f6e26391aa9f4
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720884"
---
# <a name="manage-an-items-attachments-in-a-compose-form-in-outlook"></a><span data-ttu-id="d509f-103">Outlook で新規作成フォーム内のアイテムの添付ファイルを管理する</span><span class="sxs-lookup"><span data-stu-id="d509f-103">Manage an item's attachments in a compose form in Outlook</span></span>

<span data-ttu-id="d509f-104">Office JavaScript API には、ユーザーの作成時にアイテムの添付ファイルを管理するために使用できるいくつかの Api が用意されています。</span><span class="sxs-lookup"><span data-stu-id="d509f-104">The Office JavaScript API provides several APIs you can use to manage an item's attachments when the user is composing.</span></span>

## <a name="attach-a-file-or-outlook-item"></a><span data-ttu-id="d509f-105">ファイルまたは Outlook アイテムを添付する</span><span class="sxs-lookup"><span data-stu-id="d509f-105">Attach a file or Outlook item</span></span>

<span data-ttu-id="d509f-106">添付ファイルの種類に適した方法を使用して、ファイルまたは Outlook アイテムを新規作成フォームに添付できます。</span><span class="sxs-lookup"><span data-stu-id="d509f-106">You can attach a file or Outlook item to a compose form by using the method that's appropriate for the type of attachment.</span></span>

- <span data-ttu-id="d509f-107">[Addfileattachmentasync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): ファイルの添付</span><span class="sxs-lookup"><span data-stu-id="d509f-107">[addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Attach a file</span></span>
- <span data-ttu-id="d509f-108">[addFileAttachmentFromBase64Async](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): base64 文字列を使用してファイルを添付します。</span><span class="sxs-lookup"><span data-stu-id="d509f-108">[addFileAttachmentFromBase64Async](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Attach a file using its base64 string</span></span>
- <span data-ttu-id="d509f-109">[Additemattachmentasync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Outlook アイテムを添付する</span><span class="sxs-lookup"><span data-stu-id="d509f-109">[addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Attach an Outlook item</span></span>

<span data-ttu-id="d509f-110">これらの非同期メソッドは、アクションが完了するまで待機しなくても実行が可能であることを意味します。</span><span class="sxs-lookup"><span data-stu-id="d509f-110">These are asynchronous methods, which means execution can go on without waiting for the action to complete.</span></span> <span data-ttu-id="d509f-111">追加する添付ファイルの元の場所とサイズによっては、非同期呼び出しの完了に時間がかかる場合があります。</span><span class="sxs-lookup"><span data-stu-id="d509f-111">Depending on the original location and size of the attachment being added, the asynchronous call may take a while to complete.</span></span>

<span data-ttu-id="d509f-112">アクションの完了に依存するようなタスクがある場合、それらのタスクはコールバック メソッドで実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d509f-112">If there are tasks that depend on the action to complete, you should carry out those tasks in a callback method.</span></span> <span data-ttu-id="d509f-113">このコールバックメソッドはオプションであり、添付ファイルのアップロードが完了したときに呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d509f-113">This callback method is optional and is invoked when the attachment upload has completed.</span></span> <span data-ttu-id="d509f-114">コールバック メソッドは、状態、エラー、そして添付ファイル追加によって返される値を提供する出力パラメーターとして、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトを使用します。</span><span class="sxs-lookup"><span data-stu-id="d509f-114">The callback method takes an [AsyncResult](/javascript/api/office/office.asyncresult) object as an output parameter that provides any status, error, and returned value from adding the attachment.</span></span> <span data-ttu-id="d509f-115">コールバックがその他のパラメーターを必要とする場合、オプションの `options.asyncContext` パラメーターでそれを指定することができます。</span><span class="sxs-lookup"><span data-stu-id="d509f-115">If the callback requires any extra parameters, you can specify them in the optional `options.asyncContext` parameter.</span></span> <span data-ttu-id="d509f-116">`options.asyncContext` は、コールバック メソッドが予期する任意の種類となることができます。</span><span class="sxs-lookup"><span data-stu-id="d509f-116">`options.asyncContext` can be of any type that your callback method expects.</span></span>

<span data-ttu-id="d509f-p103">たとえば、`options.asyncContext` は 1 つ以上の「キーと値のペア」を含む JSON オブジェクトとして定義することができます。非同期メソッドにオプション パラメーターを渡すさらに多くの例は、「[Office アドインにおける非同期プログラミング](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)」の Office アドイン プラットフォームで見いだすことができます。以下の例は、コールバック メソッドに引数を 2 つ渡すための `asyncContext` パラメーターの使用法を示しています。</span><span class="sxs-lookup"><span data-stu-id="d509f-p103">For example, you can define `options.asyncContext` as a JSON object that contains one or more key-value pairs. You can find more examples about passing optional parameters to asynchronous methods in the Office Add-ins platform in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods). The following example shows how to use the `asyncContext` parameter to pass 2 arguments to a callback method:</span></span>

```js
var options = { asyncContext: { var1: 1, var2: 2}};

Office.context.mailbox.item.addFileAttachmentAsync('https://contoso.com/rtm/icon.png', 'icon.png', options, callback);
```

<span data-ttu-id="d509f-p104">非同期メソッド呼び出しの正常終了またはエラーの確認は、コールバック メソッドにおいて `AsyncResult` オブジェクトの `status` と `error` のプロパティを使用して行うことができます。添付が成功裏に完了した場合、`AsyncResult.value` プロパティを使用して添付ファイル ID を取得することができます。添付ファイル ID は整数で、その後、添付ファイルを削除する場合に使用できます。</span><span class="sxs-lookup"><span data-stu-id="d509f-p104">You can check for success or error of an asynchronous method call in the callback method using the `status` and `error` properties of the `AsyncResult` object. If the attaching completes successfully, you can use the `AsyncResult.value` property to get the attachment ID. The attachment ID is an integer which you can subsequently use to remove the attachment.</span></span>

> [!NOTE]
> <span data-ttu-id="d509f-122">ベスト プラクティスとしては、この添付ファイル ID による添付ファイルの削除は、同じアドインが同じセッション内でその添付ファイルを追加した場合のみ使用すべきです。</span><span class="sxs-lookup"><span data-stu-id="d509f-122">As a best practice, you should use the attachment ID to remove an attachment only if the same add-in has added that attachment in the same session.</span></span> <span data-ttu-id="d509f-123">Outlook on the web およびモバイルデバイスでは、添付ファイル ID は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="d509f-123">In Outlook on the web and mobile devices, the attachment ID is valid only within the same session.</span></span> <span data-ttu-id="d509f-124">ユーザーがアドインを閉じたとき、またはユーザーがインライン フォームで作成を開始した後インライン フォームから出て別のウィンドウで作業を続行したとき、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="d509f-124">A session is over when the user closes the add-in, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

### <a name="attach-a-file"></a><span data-ttu-id="d509f-125">ファイルを添付する</span><span class="sxs-lookup"><span data-stu-id="d509f-125">Attach a file</span></span>

<span data-ttu-id="d509f-126">`addFileAttachmentAsync`メソッドを使用してファイルの URI を指定することによって、新規作成フォーム内のメッセージまたは予定にファイルを添付することができます。</span><span class="sxs-lookup"><span data-stu-id="d509f-126">You can attach a file to a message or appointment in a compose form by using the `addFileAttachmentAsync` method and specifying the URI of the file.</span></span> <span data-ttu-id="d509f-127">`addFileAttachmentFromBase64Async`メソッドを使用することもできますが、入力として base64 文字列を指定することもできます。</span><span class="sxs-lookup"><span data-stu-id="d509f-127">You can also use the `addFileAttachmentFromBase64Async` method but specify the base64 string as input.</span></span> <span data-ttu-id="d509f-128">If the file is protected, you can include an appropriate identity or authentication token as a URI query string parameter.</span><span class="sxs-lookup"><span data-stu-id="d509f-128">If the file is protected, you can include an appropriate identity or authentication token as a URI query string parameter.</span></span> <span data-ttu-id="d509f-129">Exchange will make a call to the URI to get the attachment, and the web service which protects the file will need to use the token as a means of authentication.</span><span class="sxs-lookup"><span data-stu-id="d509f-129">Exchange will make a call to the URI to get the attachment, and the web service which protects the file will need to use the token as a means of authentication.</span></span>

<span data-ttu-id="d509f-p107">次の JavaScript 例は、picture.png ファイルを web サーバーから取得して新規作成中のメッセージあるいは予定に添付する新規作成アドインです。コールバック メソッドはパラメーターとして `asyncResult` を使用し、結果の状態を確認し、メソッドが成功した場合に添付ファイル ID を取得します。</span><span class="sxs-lookup"><span data-stu-id="d509f-p107">The following JavaScript example is a compose add-in that attaches a file, picture.png, from a web server to the message or appointment being composed. The callback method takes `asyncResult` as a parameter, checks for the result status, and gets the attachment ID if the method succeeds.</span></span>

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

### <a name="attach-an-outlook-item"></a><span data-ttu-id="d509f-132">Outlook アイテムを添付する</span><span class="sxs-lookup"><span data-stu-id="d509f-132">Attach an Outlook item</span></span>

<span data-ttu-id="d509f-p108">新規作成フォームで、メッセージまたは予定に Outlook アイテム (メール、予定表、連絡先のアイテムなど) を添付するには、そのアイテムの Exchange Web Service (EWS) の ID を指定して `addItemAttachmentAsync` メソッドを使用します。[mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッドを使用し、EWS 操作である [FindItem](/exchange/client-developer/web-service-reference/finditem-operation) にアクセスすることにより、ユーザーのメールボックス内の電子メール、予定表、連絡先、タスク アイテムの EWS ID を取得することができます。[item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティはまた、閲覧フォーム内の既存アイテムの EWS ID も提供します。</span><span class="sxs-lookup"><span data-stu-id="d509f-p108">You can attach an Outlook item (for example, email, calendar, or contact item) to a message or appointment in a compose form by specifying the Exchange Web Services (EWS) ID of the item and using the `addItemAttachmentAsync` method. You can get the EWS ID of an email, calendar, contact or task item in the user's mailbox by using the [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method and accessing the EWS operation [FindItem](/exchange/client-developer/web-service-reference/finditem-operation). The [item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property also provides the EWS ID of an existing item in a read form.</span></span>

<span data-ttu-id="d509f-136">次の JavaScript 関数`addItemAttachment`は、上の最初の例を拡張し、作成中の電子メールまたは予定の添付ファイルとしてアイテムを追加します。</span><span class="sxs-lookup"><span data-stu-id="d509f-136">The following JavaScript function, `addItemAttachment`, extends the first example above, and adds an item as an attachment to the email or appointment that is being composed.</span></span> <span data-ttu-id="d509f-137">この関数は、添付するアイテムの EWS ID を引数として受け取ります。</span><span class="sxs-lookup"><span data-stu-id="d509f-137">The function takes as an argument the EWS ID of the item that is to be attached.</span></span> <span data-ttu-id="d509f-138">接続が成功した場合は、同じセッションでその添付ファイルを削除するなど、追加の処理のための添付ファイル ID を取得します。</span><span class="sxs-lookup"><span data-stu-id="d509f-138">If attaching succeeds, it gets the attachment ID for further processing, including removing that attachment in the same session.</span></span>

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
> <span data-ttu-id="d509f-139">新規作成アドインを使用すると、Outlook on the web またはモバイルデバイスで定期的な予定のインスタンスを添付することができます。</span><span class="sxs-lookup"><span data-stu-id="d509f-139">You can use a compose add-in to attach an instance of a recurring appointment in Outlook on the web or mobile devices.</span></span> <span data-ttu-id="d509f-140">ただし、サポート Outlook リッチ クライアントでは、インスタンスを 1 つ添付しようとしても、定期的な一連の予定 (マスター予定) が添付されます。</span><span class="sxs-lookup"><span data-stu-id="d509f-140">However, in a supporting Outlook rich client, attempting to attach an instance would result in attaching the recurring series (the master appointment).</span></span>

## <a name="get-attachments"></a><span data-ttu-id="d509f-141">添付ファイルを取得する</span><span class="sxs-lookup"><span data-stu-id="d509f-141">Get attachments</span></span>

<span data-ttu-id="d509f-142">[GetAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)メソッドを使用すると、構成されているメッセージまたは予定の添付ファイルを取得できます。</span><span class="sxs-lookup"><span data-stu-id="d509f-142">You can use the [getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method to get the attachments of the message or appointment being composed.</span></span>

<span data-ttu-id="d509f-143">添付ファイルのコンテンツを取得するには、 [Getattachmentcontentasync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="d509f-143">To get an attachment's content, you can use the [getAttachmentContentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="d509f-144">サポートされている形式は、 [Attachmentcontentformat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)列挙体に一覧表示されています。</span><span class="sxs-lookup"><span data-stu-id="d509f-144">The supported formats are listed in the [AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat) enum.</span></span>

<span data-ttu-id="d509f-145">`AsyncResult`出力パラメータオブジェクトを使用して、ステータスとエラーを確認するコールバックメソッドを提供する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d509f-145">You should provide a callback method to check for the status and any error by using the `AsyncResult` output parameter object.</span></span> <span data-ttu-id="d509f-146">オプション`asyncContext`のパラメーターを使用して、コールバックメソッドに追加のパラメーターを渡すこともできます。</span><span class="sxs-lookup"><span data-stu-id="d509f-146">You can also pass any additional parameters to the callback method by using the optional `asyncContext` parameter.</span></span>

<span data-ttu-id="d509f-147">次の JavaScript の例では、添付ファイルを取得し、サポートされている添付ファイルの形式ごとに個別の処理をセットアップできるようにします。</span><span class="sxs-lookup"><span data-stu-id="d509f-147">The following JavaScript example gets the attachments and allows you to set up distinct handling for each supported attachment format.</span></span>

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

## <a name="remove-an-attachment"></a><span data-ttu-id="d509f-148">添付ファイルを削除する</span><span class="sxs-lookup"><span data-stu-id="d509f-148">Remove an attachment</span></span>

<span data-ttu-id="d509f-149">新規作成フォームのメッセージまたは予定アイテムからファイルまたはアイテムの添付ファイルを削除するには、対応する添付ファイル ID を指定し、 [Removeattachmentasync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="d509f-149">You can remove a file or item attachment from a message or appointment item in a compose form by specifying the corresponding attachment ID and using the [removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="d509f-150">同じアドインが同じセッションで追加した添付ファイルのみを削除する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d509f-150">You should only remove attachments that the same add-in has added in the same session.</span></span> <span data-ttu-id="d509f-151">`addFileAttachmentAsync`およびメソッドと`addItemAttachmentAsync`同様に、 `removeAttachmentAsync`は非同期メソッドです。</span><span class="sxs-lookup"><span data-stu-id="d509f-151">Similar to the `addFileAttachmentAsync` and `addItemAttachmentAsync` methods, `removeAttachmentAsync` is an asynchronous method.</span></span> <span data-ttu-id="d509f-152">`AsyncResult`出力パラメータオブジェクトを使用して、ステータスとエラーを確認するコールバックメソッドを提供する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d509f-152">You should provide a callback method to check for the status and any error by using the `AsyncResult` output parameter object.</span></span> <span data-ttu-id="d509f-153">オプション`asyncContext`のパラメーターを使用して、コールバックメソッドに追加のパラメーターを渡すこともできます。</span><span class="sxs-lookup"><span data-stu-id="d509f-153">You can also pass any additional parameters to the callback method by using the optional `asyncContext` parameter.</span></span>

<span data-ttu-id="d509f-154">次の JavaScript 関数`removeAttachment`は、上記の例を引き続き拡張し、作成されている電子メールまたは予定から指定された添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="d509f-154">The following JavaScript function, `removeAttachment`, continues to extend the examples above, and removes the specified attachment from the email or appointment that is being composed.</span></span> <span data-ttu-id="d509f-155">この関数は、削除する添付ファイルの ID を引数として受け取ります。</span><span class="sxs-lookup"><span data-stu-id="d509f-155">The function takes as an argument the ID of the attachment to be removed.</span></span> <span data-ttu-id="d509f-156">添付ファイルの ID を取得するには、成功`addFileAttachmentAsync` `addFileAttachmentFromBase64Async`した、 `addItemAttachmentAsync`またはメソッドが呼び出された後に`removeAttachmentAsync` 、以降のメソッド呼び出しに対して保存します。</span><span class="sxs-lookup"><span data-stu-id="d509f-156">You can obtain the ID of an attachment after a successful `addFileAttachmentAsync`, `addFileAttachmentFromBase64Async`, or `addItemAttachmentAsync` method call, and store it for a subsequent `removeAttachmentAsync` method call.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="d509f-157">関連項目</span><span class="sxs-lookup"><span data-stu-id="d509f-157">See also</span></span>

- [<span data-ttu-id="d509f-158">新規作成フォーム用の Outlook アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="d509f-158">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)
- [<span data-ttu-id="d509f-159">Office アドインにおける非同期プログラミング</span><span class="sxs-lookup"><span data-stu-id="d509f-159">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)

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
# <a name="manage-an-items-attachments-in-a-compose-form-in-outlook"></a><span data-ttu-id="1827a-103">Outlook の作成フォームでアイテムの添付ファイルを管理する</span><span class="sxs-lookup"><span data-stu-id="1827a-103">Manage an item's attachments in a compose form in Outlook</span></span>

<span data-ttu-id="1827a-104">JavaScript API Officeには、ユーザーが作成するときにアイテムの添付ファイルを管理するために使用できるいくつかの API が提供されています。</span><span class="sxs-lookup"><span data-stu-id="1827a-104">The Office JavaScript API provides several APIs you can use to manage an item's attachments when the user is composing.</span></span>

## <a name="attach-a-file-or-outlook-item"></a><span data-ttu-id="1827a-105">ファイルまたは Outlook アイテムを添付する</span><span class="sxs-lookup"><span data-stu-id="1827a-105">Attach a file or Outlook item</span></span>

<span data-ttu-id="1827a-106">添付ファイルの種類に適したメソッドを使用して、ファイルまたは Outlook アイテムを作成フォームに添付できます。</span><span class="sxs-lookup"><span data-stu-id="1827a-106">You can attach a file or Outlook item to a compose form by using the method that's appropriate for the type of attachment.</span></span>

- <span data-ttu-id="1827a-107">[addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): ファイルを添付する</span><span class="sxs-lookup"><span data-stu-id="1827a-107">[addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Attach a file</span></span>
- <span data-ttu-id="1827a-108">[addFileAttachmentFromBase64Async : base64](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)文字列を使用してファイルを添付する</span><span class="sxs-lookup"><span data-stu-id="1827a-108">[addFileAttachmentFromBase64Async](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Attach a file using its base64 string</span></span>
- <span data-ttu-id="1827a-109">[addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Outlook アイテムの添付</span><span class="sxs-lookup"><span data-stu-id="1827a-109">[addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Attach an Outlook item</span></span>

<span data-ttu-id="1827a-110">これらは非同期メソッドです。つまり、アクションが完了するのを待たずに実行を続けできます。</span><span class="sxs-lookup"><span data-stu-id="1827a-110">These are asynchronous methods, which means execution can go on without waiting for the action to complete.</span></span> <span data-ttu-id="1827a-111">追加する添付ファイルの元の場所とサイズによっては、非同期呼び出しの完了に時間がかかる場合があります。</span><span class="sxs-lookup"><span data-stu-id="1827a-111">Depending on the original location and size of the attachment being added, the asynchronous call may take a while to complete.</span></span>

<span data-ttu-id="1827a-112">アクションの完了に依存するようなタスクがある場合、それらのタスクはコールバック メソッドで実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1827a-112">If there are tasks that depend on the action to complete, you should carry out those tasks in a callback method.</span></span> <span data-ttu-id="1827a-113">このコールバック メソッドはオプションで、添付ファイルのアップロードが完了すると呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1827a-113">This callback method is optional and is invoked when the attachment upload has completed.</span></span> <span data-ttu-id="1827a-114">コールバック メソッドは、状態、エラー、そして添付ファイル追加によって返される値を提供する出力パラメーターとして、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトを使用します。</span><span class="sxs-lookup"><span data-stu-id="1827a-114">The callback method takes an [AsyncResult](/javascript/api/office/office.asyncresult) object as an output parameter that provides any status, error, and returned value from adding the attachment.</span></span> <span data-ttu-id="1827a-115">コールバックがその他のパラメーターを必要とする場合、オプションの `options.asyncContext` パラメーターでそれを指定することができます。</span><span class="sxs-lookup"><span data-stu-id="1827a-115">If the callback requires any extra parameters, you can specify them in the optional `options.asyncContext` parameter.</span></span> <span data-ttu-id="1827a-116">`options.asyncContext` は、コールバック メソッドが予期する任意の種類となることができます。</span><span class="sxs-lookup"><span data-stu-id="1827a-116">`options.asyncContext` can be of any type that your callback method expects.</span></span>

<span data-ttu-id="1827a-p103">たとえば、`options.asyncContext` は 1 つ以上の「キーと値のペア」を含む JSON オブジェクトとして定義することができます。非同期メソッドにオプション パラメーターを渡すさらに多くの例は、「[Office アドインにおける非同期プログラミング](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)」の Office アドイン プラットフォームで見いだすことができます。以下の例は、コールバック メソッドに引数を 2 つ渡すための `asyncContext` パラメーターの使用法を示しています。</span><span class="sxs-lookup"><span data-stu-id="1827a-p103">For example, you can define `options.asyncContext` as a JSON object that contains one or more key-value pairs. You can find more examples about passing optional parameters to asynchronous methods in the Office Add-ins platform in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods). The following example shows how to use the `asyncContext` parameter to pass 2 arguments to a callback method:</span></span>

```js
var options = { asyncContext: { var1: 1, var2: 2}};

Office.context.mailbox.item.addFileAttachmentAsync('https://contoso.com/rtm/icon.png', 'icon.png', options, callback);
```

<span data-ttu-id="1827a-p104">非同期メソッド呼び出しの正常終了またはエラーの確認は、コールバック メソッドにおいて `AsyncResult` オブジェクトの `status` と `error` のプロパティを使用して行うことができます。添付が成功裏に完了した場合、`AsyncResult.value` プロパティを使用して添付ファイル ID を取得することができます。添付ファイル ID は整数で、その後、添付ファイルを削除する場合に使用できます。</span><span class="sxs-lookup"><span data-stu-id="1827a-p104">You can check for success or error of an asynchronous method call in the callback method using the `status` and `error` properties of the `AsyncResult` object. If the attaching completes successfully, you can use the `AsyncResult.value` property to get the attachment ID. The attachment ID is an integer which you can subsequently use to remove the attachment.</span></span>

> [!NOTE]
> <span data-ttu-id="1827a-122">添付ファイル ID は同じセッション内でのみ有効であり、セッション間で同じ添付ファイルにマップされる保証はありません。</span><span class="sxs-lookup"><span data-stu-id="1827a-122">The attachment ID is valid only within the same session and isn't guaranteed to map to the same attachment across sessions.</span></span> <span data-ttu-id="1827a-123">セッションが終了した場合の例としては、ユーザーがアドインを閉じる場合、またはユーザーがインライン フォームで作成を開始し、その後インライン フォームをポップアウトして別のウィンドウで続行する場合があります。</span><span class="sxs-lookup"><span data-stu-id="1827a-123">Examples of when a session is over include when the user closes the add-in, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

### <a name="attach-a-file"></a><span data-ttu-id="1827a-124">ファイルの添付</span><span class="sxs-lookup"><span data-stu-id="1827a-124">Attach a file</span></span>

<span data-ttu-id="1827a-125">作成フォームのメッセージまたは予定にファイルを添付するには、メソッドを使用してファイル `addFileAttachmentAsync` の URI を指定します。</span><span class="sxs-lookup"><span data-stu-id="1827a-125">You can attach a file to a message or appointment in a compose form by using the `addFileAttachmentAsync` method and specifying the URI of the file.</span></span> <span data-ttu-id="1827a-126">メソッドを使用できますが `addFileAttachmentFromBase64Async` 、base64 文字列を入力として指定できます。</span><span class="sxs-lookup"><span data-stu-id="1827a-126">You can also use the `addFileAttachmentFromBase64Async` method but specify the base64 string as input.</span></span> <span data-ttu-id="1827a-127">If the file is protected, you can include an appropriate identity or authentication token as a URI query string parameter.</span><span class="sxs-lookup"><span data-stu-id="1827a-127">If the file is protected, you can include an appropriate identity or authentication token as a URI query string parameter.</span></span> <span data-ttu-id="1827a-128">Exchange will make a call to the URI to get the attachment, and the web service which protects the file will need to use the token as a means of authentication.</span><span class="sxs-lookup"><span data-stu-id="1827a-128">Exchange will make a call to the URI to get the attachment, and the web service which protects the file will need to use the token as a means of authentication.</span></span>

<span data-ttu-id="1827a-p107">次の JavaScript 例は、picture.png ファイルを web サーバーから取得して新規作成中のメッセージあるいは予定に添付する新規作成アドインです。コールバック メソッドはパラメーターとして `asyncResult` を使用し、結果の状態を確認し、メソッドが成功した場合に添付ファイル ID を取得します。</span><span class="sxs-lookup"><span data-stu-id="1827a-p107">The following JavaScript example is a compose add-in that attaches a file, picture.png, from a web server to the message or appointment being composed. The callback method takes `asyncResult` as a parameter, checks for the result status, and gets the attachment ID if the method succeeds.</span></span>

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

### <a name="attach-an-outlook-item"></a><span data-ttu-id="1827a-131">Outlook アイテムを添付する</span><span class="sxs-lookup"><span data-stu-id="1827a-131">Attach an Outlook item</span></span>

<span data-ttu-id="1827a-132">アイテムの Exchange Web Services (EWS) ID を指定し、メソッドを使用して、作成フォーム内のメッセージまたは予定に Outlook アイテム (電子メール、予定表、連絡先アイテムなど) を添付できます。 `addItemAttachmentAsync`</span><span class="sxs-lookup"><span data-stu-id="1827a-132">You can attach an Outlook item (for example, email, calendar, or contact item) to a message or appointment in a compose form by specifying the Exchange Web Services (EWS) ID of the item and using the `addItemAttachmentAsync` method.</span></span> <span data-ttu-id="1827a-133">[mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)メソッドを使用して EWS 操作[FindItem](/exchange/client-developer/web-service-reference/finditem-operation)にアクセスすると、ユーザーのメールボックス内の電子メール、予定表、連絡先、またはタスク アイテムの EWS ID を取得できます。</span><span class="sxs-lookup"><span data-stu-id="1827a-133">You can get the EWS ID of an email, calendar, contact, or task item in the user's mailbox by using the [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method and accessing the EWS operation [FindItem](/exchange/client-developer/web-service-reference/finditem-operation).</span></span> <span data-ttu-id="1827a-134">閲覧フォームの既存のアイテムでは、 [item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティでも EWS ID が取得できます。</span><span class="sxs-lookup"><span data-stu-id="1827a-134">The [item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property also provides the EWS ID of an existing item in a read form.</span></span>

<span data-ttu-id="1827a-135">次の JavaScript 関数は、上記の最初の例を拡張し、構成されている電子メールまたは予定に添付ファイルとしてアイテム `addItemAttachment` を追加します。</span><span class="sxs-lookup"><span data-stu-id="1827a-135">The following JavaScript function, `addItemAttachment`, extends the first example above, and adds an item as an attachment to the email or appointment that is being composed.</span></span> <span data-ttu-id="1827a-136">この関数は、添付するアイテムの EWS ID を引数として受け取ります。</span><span class="sxs-lookup"><span data-stu-id="1827a-136">The function takes as an argument the EWS ID of the item that is to be attached.</span></span> <span data-ttu-id="1827a-137">接続が成功した場合は、同じセッションで添付ファイルを削除するなど、さらに処理するための添付ファイル ID を取得します。</span><span class="sxs-lookup"><span data-stu-id="1827a-137">If attaching succeeds, it gets the attachment ID for further processing, including removing that attachment in the same session.</span></span>

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
> <span data-ttu-id="1827a-138">作成アドインを使用して、Outlook on the web またはモバイル デバイスで定期的な予定のインスタンスを添付できます。</span><span class="sxs-lookup"><span data-stu-id="1827a-138">You can use a compose add-in to attach an instance of a recurring appointment in Outlook on the web or on mobile devices.</span></span> <span data-ttu-id="1827a-139">ただし、サポートしている Outlook デスクトップ クライアントでは、インスタンスを接続しようとすると、定期的な系列 (親予定) が添付されます。</span><span class="sxs-lookup"><span data-stu-id="1827a-139">However, in a supporting Outlook desktop client, attempting to attach an instance would result in attaching the recurring series (the parent appointment).</span></span>

## <a name="get-attachments"></a><span data-ttu-id="1827a-140">添付ファイルを取得する</span><span class="sxs-lookup"><span data-stu-id="1827a-140">Get attachments</span></span>

<span data-ttu-id="1827a-141">作成モードで添付ファイルを取得する API は、要件セット [1.8 から利用できます](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)。</span><span class="sxs-lookup"><span data-stu-id="1827a-141">APIs to get attachments in compose mode are available from [requirement set 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

- [<span data-ttu-id="1827a-142">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="1827a-142">getAttachmentsAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [<span data-ttu-id="1827a-143">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="1827a-143">getAttachmentContentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

<span data-ttu-id="1827a-144">[getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)メソッドを使用して、構成されているメッセージまたは予定の添付ファイルを取得できます。</span><span class="sxs-lookup"><span data-stu-id="1827a-144">You can use the [getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method to get the attachments of the message or appointment being composed.</span></span>

<span data-ttu-id="1827a-145">添付ファイルのコンテンツを取得するには [、getAttachmentContentAsync メソッドを使用](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) できます。</span><span class="sxs-lookup"><span data-stu-id="1827a-145">To get an attachment's content, you can use the [getAttachmentContentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="1827a-146">サポートされている形式は [、AttachmentContentFormat 列挙型に一覧表示](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat) されます。</span><span class="sxs-lookup"><span data-stu-id="1827a-146">The supported formats are listed in the [AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat) enum.</span></span>

<span data-ttu-id="1827a-147">output パラメーター オブジェクトを使用して、状態とエラーを確認するコールバック メソッドを `AsyncResult` 指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1827a-147">You should provide a callback method to check for the status and any error by using the `AsyncResult` output parameter object.</span></span> <span data-ttu-id="1827a-148">省略可能なパラメーターを使用して、コールバック メソッドに追加のパラメーターを渡 `asyncContext` することもできます。</span><span class="sxs-lookup"><span data-stu-id="1827a-148">You can also pass any additional parameters to the callback method by using the optional `asyncContext` parameter.</span></span>

<span data-ttu-id="1827a-149">次の JavaScript の例では、添付ファイルを取得し、サポートされている添付ファイル形式ごとに個別の処理を設定できます。</span><span class="sxs-lookup"><span data-stu-id="1827a-149">The following JavaScript example gets the attachments and allows you to set up distinct handling for each supported attachment format.</span></span>

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

## <a name="remove-an-attachment"></a><span data-ttu-id="1827a-150">添付ファイルの削除</span><span class="sxs-lookup"><span data-stu-id="1827a-150">Remove an attachment</span></span>

<span data-ttu-id="1827a-151">[removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)メソッドを使用する場合は、対応する添付ファイル ID を指定して、作成フォームのメッセージまたは予定アイテムからファイルまたはアイテムの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="1827a-151">You can remove a file or item attachment from a message or appointment item in a compose form by specifying the corresponding attachment ID when using the [removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1827a-152">要件セット 1.7 以前を使用している場合は、同じアドインが同じセッションで追加した添付ファイルのみを削除する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1827a-152">If you're using requirement set 1.7 or earlier, you should only remove attachments that the same add-in has added in the same session.</span></span>

<span data-ttu-id="1827a-153">、 および `addFileAttachmentAsync` メソッド `addItemAttachmentAsync` と `getAttachmentsAsync` 同様に、 `removeAttachmentAsync` 非同期メソッドです。</span><span class="sxs-lookup"><span data-stu-id="1827a-153">Similar to the `addFileAttachmentAsync`, `addItemAttachmentAsync`, and `getAttachmentsAsync` methods, `removeAttachmentAsync` is an asynchronous method.</span></span> <span data-ttu-id="1827a-154">output パラメーター オブジェクトを使用して、状態とエラーを確認するコールバック メソッドを `AsyncResult` 指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1827a-154">You should provide a callback method to check for the status and any error by using the `AsyncResult` output parameter object.</span></span> <span data-ttu-id="1827a-155">省略可能なパラメーターを使用して、コールバック メソッドに追加のパラメーターを渡 `asyncContext` することもできます。</span><span class="sxs-lookup"><span data-stu-id="1827a-155">You can also pass any additional parameters to the callback method by using the optional `asyncContext` parameter.</span></span>

<span data-ttu-id="1827a-156">次の JavaScript 関数は、上記の例を引き続き拡張し、構成されている電子メールまたは予定から指定された添付ファイル `removeAttachment` を削除します。</span><span class="sxs-lookup"><span data-stu-id="1827a-156">The following JavaScript function, `removeAttachment`, continues to extend the examples above, and removes the specified attachment from the email or appointment that is being composed.</span></span> <span data-ttu-id="1827a-157">この関数は、削除する添付ファイルの ID を引数として受け取ります。</span><span class="sxs-lookup"><span data-stu-id="1827a-157">The function takes as an argument the ID of the attachment to be removed.</span></span> <span data-ttu-id="1827a-158">添付ファイルの ID は、成功した 、またはメソッドの呼び出し後に取得し、後続のメソッド呼び出 `addFileAttachmentAsync` `addFileAttachmentFromBase64Async` `addItemAttachmentAsync` しで `removeAttachmentAsync` 使用できます。</span><span class="sxs-lookup"><span data-stu-id="1827a-158">You can obtain the ID of an attachment after a successful `addFileAttachmentAsync`, `addFileAttachmentFromBase64Async`, or `addItemAttachmentAsync` method call, and use it in a subsequent `removeAttachmentAsync` method call.</span></span> <span data-ttu-id="1827a-159">(要件セット 1.8 で導入) を呼び出して、そのアドイン セッションの添付ファイルとその `getAttachmentsAsync` ID を取得することもできます。</span><span class="sxs-lookup"><span data-stu-id="1827a-159">You can also call `getAttachmentsAsync` (introduced in requirement set 1.8) to get the attachments and their IDs for that add-in session.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="1827a-160">関連項目</span><span class="sxs-lookup"><span data-stu-id="1827a-160">See also</span></span>

- [<span data-ttu-id="1827a-161">新規作成フォーム用の Outlook アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="1827a-161">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)
- [<span data-ttu-id="1827a-162">Office アドインにおける非同期プログラミング</span><span class="sxs-lookup"><span data-stu-id="1827a-162">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)

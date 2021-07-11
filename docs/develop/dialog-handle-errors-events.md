---
title: Office ダイアログ ボックスでのエラーとイベントの処理
description: '[エラー] ダイアログ ボックスを開いて使用するときにエラーをトラップして処理するOffice説明します。'
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: be1fb8bcd30b47ac6399657d928d3cad7f857f39
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349897"
---
# <a name="handling-errors-and-events-in-the-office-dialog-box"></a><span data-ttu-id="72f27-103">Office ダイアログ ボックスでのエラーとイベントの処理</span><span class="sxs-lookup"><span data-stu-id="72f27-103">Handling errors and events in the Office dialog box</span></span>

<span data-ttu-id="72f27-104">この記事では、ダイアログ ボックスを開く際にエラーをトラップして処理する方法と、ダイアログ ボックス内で発生するエラーについて説明します。</span><span class="sxs-lookup"><span data-stu-id="72f27-104">This article describes how to trap and handle errors when opening the dialog box and errors that happen inside the dialog box.</span></span>

> [!NOTE]
> <span data-ttu-id="72f27-105">この記事では、「Office アドインで Office ダイアログ API を使用する」の説明に従って、Office ダイアログ[API](dialog-api-in-office-add-ins.md)の使用の基本について理解している必要があります。</span><span class="sxs-lookup"><span data-stu-id="72f27-105">This article presupposes that you are familiar with the basics of using the Office dialog API as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).</span></span>
> 
> <span data-ttu-id="72f27-106">詳細については、「[ベスト プラクティスとルール」を参照Office API を参照してください](dialog-best-practices.md)。</span><span class="sxs-lookup"><span data-stu-id="72f27-106">See also [Best practices and rules for the Office dialog API](dialog-best-practices.md).</span></span>

<span data-ttu-id="72f27-107">コードでイベントの 2 つのカテゴリを処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="72f27-107">Your code should handle two categories of events:</span></span>

- <span data-ttu-id="72f27-108">ダイアログ ボックスを作成できないために `displayDialogAsync` の呼び出しによって返されるエラー。</span><span class="sxs-lookup"><span data-stu-id="72f27-108">Errors returned by the call of `displayDialogAsync` because the dialog box cannot be created.</span></span>
- <span data-ttu-id="72f27-109">ダイアログ ボックスのエラー、その他のイベント。</span><span class="sxs-lookup"><span data-stu-id="72f27-109">Errors, and other events, in the dialog box.</span></span>

## <a name="errors-from-displaydialogasync"></a><span data-ttu-id="72f27-110">displayDialogAsync のエラー</span><span class="sxs-lookup"><span data-stu-id="72f27-110">Errors from displayDialogAsync</span></span>

<span data-ttu-id="72f27-111">プラットフォームとシステムの一般的なエラーに加えて、4 つのエラーは呼び出しに固有です `displayDialogAsync` 。</span><span class="sxs-lookup"><span data-stu-id="72f27-111">In addition to general platform and system errors, four errors are specific to calling `displayDialogAsync`.</span></span>

|<span data-ttu-id="72f27-112">コード番号</span><span class="sxs-lookup"><span data-stu-id="72f27-112">Code number</span></span>|<span data-ttu-id="72f27-113">意味</span><span class="sxs-lookup"><span data-stu-id="72f27-113">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="72f27-114">12004</span><span class="sxs-lookup"><span data-stu-id="72f27-114">12004</span></span>|<span data-ttu-id="72f27-p101">`displayDialogAsync` に渡される URL のドメインは信頼されていません。ドメインは、ホスト ページと同じドメインにある必要があります (プロトコルとポート番号を含む)。</span><span class="sxs-lookup"><span data-stu-id="72f27-p101">The domain of the URL passed to `displayDialogAsync` is not trusted. The domain must be the same domain as the host page (including protocol and port number).</span></span>|
|<span data-ttu-id="72f27-117">12005</span><span class="sxs-lookup"><span data-stu-id="72f27-117">12005</span></span>|<span data-ttu-id="72f27-118">`displayDialogAsync` に渡される URL には HTTP プロトコルを使用します。</span><span class="sxs-lookup"><span data-stu-id="72f27-118">The URL passed to `displayDialogAsync` uses the HTTP protocol.</span></span> <span data-ttu-id="72f27-119">HTTPS が必要です。</span><span class="sxs-lookup"><span data-stu-id="72f27-119">HTTPS is required.</span></span> <span data-ttu-id="72f27-120">(一部のバージョンの Office 12005 で返されるエラー メッセージ テキストは、12004 で返されるのと同じです)。</span><span class="sxs-lookup"><span data-stu-id="72f27-120">(In some versions of Office, the error message text returned with 12005 is the same one returned for 12004.)</span></span>|
|<span data-ttu-id="72f27-121"><span id="12007">12007</span></span><span class="sxs-lookup"><span data-stu-id="72f27-121"><span id="12007">12007</span></span></span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|<span data-ttu-id="72f27-p103">ダイアログ ボックスは、このホスト ウィンドウで既に開いています。作業ウィンドウなどのホスト ウィンドウで一度に開けるダイアログ ボックスは 1 つだけです。</span><span class="sxs-lookup"><span data-stu-id="72f27-p103">A dialog box is already opened from this host window. A host window, such as a task pane, can only have one dialog box open at a time.</span></span>|
|<span data-ttu-id="72f27-124">12009</span><span class="sxs-lookup"><span data-stu-id="72f27-124">12009</span></span>|<span data-ttu-id="72f27-125">ダイアログ ボックスを無視するようにユーザーが選択しました。</span><span class="sxs-lookup"><span data-stu-id="72f27-125">The user chose to ignore the dialog box.</span></span> <span data-ttu-id="72f27-126">このエラーは、ユーザー Office on the webダイアログ ボックスの表示を許可しない場合がある場合に発生する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="72f27-126">This error can occur in Office on the web, where users may choose not to allow an add-in to present a dialog box.</span></span> <span data-ttu-id="72f27-127">詳細については、「ポップアップ ブロック[を使用したポップアップ ブロックの処理」を参照Office on the web。](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web)</span><span class="sxs-lookup"><span data-stu-id="72f27-127">For more information, see [Handling pop-up blockers with Office on the web](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web).</span></span>|

<span data-ttu-id="72f27-128">呼 `displayDialogAsync` び出された場合 [、AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトをコールバック関数に渡します。</span><span class="sxs-lookup"><span data-stu-id="72f27-128">When `displayDialogAsync` is called, it passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to its callback function.</span></span> <span data-ttu-id="72f27-129">呼び出しが成功すると、ダイアログ ボックスが開き、オブジェクトのプロパティ `value` `AsyncResult` が [Dialog](/javascript/api/office/office.dialog) オブジェクトになります。</span><span class="sxs-lookup"><span data-stu-id="72f27-129">When the call is successful, the dialog box is opened, and the `value` property of the `AsyncResult` object is a [Dialog](/javascript/api/office/office.dialog) object.</span></span> <span data-ttu-id="72f27-130">この例については、「ダイアログ ボックスから [ホスト ページに情報を送信する」を参照してください](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page)。</span><span class="sxs-lookup"><span data-stu-id="72f27-130">For an example of this, see [Send information from the dialog box to the host page](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page).</span></span> <span data-ttu-id="72f27-131">呼び出しが失敗すると、ダイアログ ボックスは作成されません。オブジェクトのプロパティはに設定され、オブジェクトの `displayDialogAsync` `status` `AsyncResult` `Office.AsyncResultStatus.Failed` `error` プロパティが設定されます。</span><span class="sxs-lookup"><span data-stu-id="72f27-131">When the call to `displayDialogAsync` fails, the dialog box is not created, the `status` property of the `AsyncResult` object is set to `Office.AsyncResultStatus.Failed`, and the `error` property of the object is populated.</span></span> <span data-ttu-id="72f27-132">エラーが発生した場合は、常にテストし `status` 、応答するコールバックを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="72f27-132">You should always provide a callback that tests the `status` and responds when it's an error.</span></span> <span data-ttu-id="72f27-133">コード番号に関係なくエラー メッセージを報告する例については、次のコードを参照してください。</span><span class="sxs-lookup"><span data-stu-id="72f27-133">For an example that reports the error message regardless of its code number, see the following code.</span></span> <span data-ttu-id="72f27-134">(この `showNotification` 記事で定義されていない関数は、エラーを表示またはログに記録します。</span><span class="sxs-lookup"><span data-stu-id="72f27-134">(The `showNotification` function, not defined in this article, either displays or logs the error.</span></span> <span data-ttu-id="72f27-135">アドイン内でこの関数を実装する方法の例については、「Office ダイアログ API の例 」[を参照してください](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。</span><span class="sxs-lookup"><span data-stu-id="72f27-135">For an example of how you can implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).)</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        showNotification(asyncResult.error.code = ": " + asyncResult.error.message);
    } else {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
});
```

## <a name="errors-and-events-in-the-dialog-box"></a><span data-ttu-id="72f27-136">ダイアログ ボックスのエラーとイベント</span><span class="sxs-lookup"><span data-stu-id="72f27-136">Errors and events in the dialog box</span></span>

<span data-ttu-id="72f27-137">ダイアログ ボックス内の 3 つのエラーとイベントは、ホスト `DialogEventReceived` ページでイベントを発生します。</span><span class="sxs-lookup"><span data-stu-id="72f27-137">Three errors and events in the dialog box will raise a `DialogEventReceived` event in the host page.</span></span> <span data-ttu-id="72f27-138">ホスト ページの種類を確認するには、「ホスト ページからダイアログ ボックスを開 [く」を参照してください](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)。</span><span class="sxs-lookup"><span data-stu-id="72f27-138">For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span></span>

|<span data-ttu-id="72f27-139">コード番号</span><span class="sxs-lookup"><span data-stu-id="72f27-139">Code number</span></span>|<span data-ttu-id="72f27-140">意味</span><span class="sxs-lookup"><span data-stu-id="72f27-140">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="72f27-141">12002</span><span class="sxs-lookup"><span data-stu-id="72f27-141">12002</span></span>|<span data-ttu-id="72f27-142">以下のいずれか:</span><span class="sxs-lookup"><span data-stu-id="72f27-142">One of the following:</span></span><br> <span data-ttu-id="72f27-143">- `displayDialogAsync` に渡された URL にページが存在しない。</span><span class="sxs-lookup"><span data-stu-id="72f27-143">- No page exists at the URL that was passed to `displayDialogAsync`.</span></span><br> <span data-ttu-id="72f27-144">- 読み込まれたが、ダイアログ ボックスが見つからまたは読み込めないページにリダイレクトされたページ、または構文が無効な URL にリダイレクト `displayDialogAsync` されたページ。</span><span class="sxs-lookup"><span data-stu-id="72f27-144">- The page that was passed to `displayDialogAsync` loaded, but the dialog box was then redirected to a page that it cannot find or load, or it has been directed to a URL with invalid syntax.</span></span>|
|<span data-ttu-id="72f27-145">12003</span><span class="sxs-lookup"><span data-stu-id="72f27-145">12003</span></span>|<span data-ttu-id="72f27-p107">ダイアログ ボックスが HTTP プロトコルを使用している URL を指していました。HTTPS が必要です。</span><span class="sxs-lookup"><span data-stu-id="72f27-p107">The dialog box was directed to a URL with the HTTP protocol. HTTPS is required.</span></span>|
|<span data-ttu-id="72f27-148">12006</span><span class="sxs-lookup"><span data-stu-id="72f27-148">12006</span></span>|<span data-ttu-id="72f27-149">ダイアログ ボックスが閉じられました。通常、ユーザーが [閉じる] ボタン **X を選択\*\*\*\*したためです**。</span><span class="sxs-lookup"><span data-stu-id="72f27-149">The dialog box was closed, usually because the user chose the **Close** button **X**.</span></span>|

<span data-ttu-id="72f27-p108">コードで、呼び出し内の `DialogEventReceived` イベントのハンドラーを `displayDialogAsync` に割り当てることができます。次に簡単な例を示します。</span><span class="sxs-lookup"><span data-stu-id="72f27-p108">Your code can assign a handler for the `DialogEventReceived` event in the call to `displayDialogAsync`. The following is a simple example.</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

<span data-ttu-id="72f27-152">エラー コードごとにカスタム エラー メッセージを作成するイベントのハンドラーの例については、 `DialogEventReceived` 次の例を参照してください。</span><span class="sxs-lookup"><span data-stu-id="72f27-152">For an example of a handler for the `DialogEventReceived` event that creates custom error messages for each error code, see the following example.</span></span>

```js
function processDialogEvent(arg) {
    switch (arg.error) {
        case 12002:
            showNotification("The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.");
            break;
        case 12003:
            showNotification("The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.");            break;
        case 12006:
            showNotification("Dialog closed.");
            break;
        default:
            showNotification("Unknown error in dialog box.");
            break;
    }
}
```

<span data-ttu-id="72f27-153">この方法でエラーを処理するサンプル アドインについては、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="72f27-153">For a sample add-in that handles errors in this way, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

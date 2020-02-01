---
title: Office ダイアログボックスでのエラーとイベントの処理
description: Office ダイアログボックスを開いて使用するときに発生するエラーをトラップして処理する方法について説明します。
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: a35131a46dc9f5edc18df37495abe5d8c2c5ad2a
ms.sourcegitcommit: 4c9e02dac6f8030efc7415e699370753ec9415c8
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/01/2020
ms.locfileid: "41650109"
---
# <a name="handling-errors-and-events-in-the-office-dialog-box"></a><span data-ttu-id="f4f93-103">Office ダイアログボックスでのエラーとイベントの処理</span><span class="sxs-lookup"><span data-stu-id="f4f93-103">Handling errors and events in the Office dialog box</span></span>

<span data-ttu-id="f4f93-104">この記事では、ダイアログボックスを開くときにエラーをトラップして処理する方法と、ダイアログボックス内で発生するエラーについて説明します。</span><span class="sxs-lookup"><span data-stu-id="f4f93-104">This article describes how to trap and handle errors when opening the dialog box and errors that happen inside the dialog box.</span></span>

> [!NOTE]
> <span data-ttu-id="f4f93-105">この記事では、「office[アドインで office ダイアログ api を使用](dialog-api-in-office-add-ins.md)する」で説明されている OFFICE ダイアログ api の使用についての基本事項を presupposes しています。</span><span class="sxs-lookup"><span data-stu-id="f4f93-105">This article presupposes that you are familiar with the basics of using the Office dialog API as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).</span></span>
> 
> <span data-ttu-id="f4f93-106">「 [Office ダイアログ API のベストプラクティスとルール](dialog-best-practices.md)」も参照してください。</span><span class="sxs-lookup"><span data-stu-id="f4f93-106">See also [Best practices and rules for the Office dialog API](dialog-best-practices.md).</span></span>

<span data-ttu-id="f4f93-107">コードでイベントの 2 つのカテゴリを処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f4f93-107">Your code should handle two categories of events:</span></span>

- <span data-ttu-id="f4f93-108">ダイアログ ボックスを作成できないために `displayDialogAsync` の呼び出しによって返されるエラー。</span><span class="sxs-lookup"><span data-stu-id="f4f93-108">Errors returned by the call of `displayDialogAsync` because the dialog box cannot be created.</span></span>
- <span data-ttu-id="f4f93-109">ダイアログボックス内のエラーおよびその他のイベント。</span><span class="sxs-lookup"><span data-stu-id="f4f93-109">Errors, and other events, in the dialog box.</span></span>

## <a name="errors-from-displaydialogasync"></a><span data-ttu-id="f4f93-110">displayDialogAsync のエラー</span><span class="sxs-lookup"><span data-stu-id="f4f93-110">Errors from displayDialogAsync</span></span>

<span data-ttu-id="f4f93-111">一般的なプラットフォームおよびシステムのエラーに加えて、4つのエラー `displayDialogAsync`が呼び出しに固有のものです。</span><span class="sxs-lookup"><span data-stu-id="f4f93-111">In addition to general platform and system errors, four errors are specific to calling `displayDialogAsync`.</span></span>

|<span data-ttu-id="f4f93-112">コード番号</span><span class="sxs-lookup"><span data-stu-id="f4f93-112">Code number</span></span>|<span data-ttu-id="f4f93-113">意味</span><span class="sxs-lookup"><span data-stu-id="f4f93-113">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="f4f93-114">12004</span><span class="sxs-lookup"><span data-stu-id="f4f93-114">12004</span></span>|<span data-ttu-id="f4f93-p101">`displayDialogAsync` に渡される URL のドメインは信頼されていません。ドメインは、ホスト ページと同じドメインにある必要があります (プロトコルとポート番号を含む)。</span><span class="sxs-lookup"><span data-stu-id="f4f93-p101">The domain of the URL passed to `displayDialogAsync` is not trusted. The domain must be the same domain as the host page (including protocol and port number).</span></span>|
|<span data-ttu-id="f4f93-117">12005</span><span class="sxs-lookup"><span data-stu-id="f4f93-117">12005</span></span>|<span data-ttu-id="f4f93-118">`displayDialogAsync` に渡される URL には HTTP プロトコルを使用します。</span><span class="sxs-lookup"><span data-stu-id="f4f93-118">The URL passed to `displayDialogAsync` uses the HTTP protocol.</span></span> <span data-ttu-id="f4f93-119">HTTPS が必要です。</span><span class="sxs-lookup"><span data-stu-id="f4f93-119">HTTPS is required.</span></span> <span data-ttu-id="f4f93-120">(一部のバージョンの Office では、12005で返されるエラーメッセージのテキストは、12004で返されるものと同じです。)</span><span class="sxs-lookup"><span data-stu-id="f4f93-120">(In some versions of Office, the error message text returned with 12005 is the same one returned for 12004.)</span></span>|
|<span data-ttu-id="f4f93-121"><span id="12007">12007</span></span><span class="sxs-lookup"><span data-stu-id="f4f93-121"><span id="12007">12007</span></span></span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|<span data-ttu-id="f4f93-p103">ダイアログ ボックスは、このホスト ウィンドウで既に開いています。作業ウィンドウなどのホスト ウィンドウで一度に開けるダイアログ ボックスは 1 つだけです。</span><span class="sxs-lookup"><span data-stu-id="f4f93-p103">A dialog box is already opened from this host window. A host window, such as a task pane, can only have one dialog box open at a time.</span></span>|
|<span data-ttu-id="f4f93-124">12009</span><span class="sxs-lookup"><span data-stu-id="f4f93-124">12009</span></span>|<span data-ttu-id="f4f93-125">ダイアログ ボックスを無視するようにユーザーが選択しました。</span><span class="sxs-lookup"><span data-stu-id="f4f93-125">The user chose to ignore the dialog box.</span></span> <span data-ttu-id="f4f93-126">このエラーは、web 上の Office で発生する可能性があります。ユーザーは、アドインによるダイアログボックスの表示を許可しないことを選択できます。</span><span class="sxs-lookup"><span data-stu-id="f4f93-126">This error can occur in Office on the web, where users may choose not to allow an add-in to present a dialog box.</span></span> <span data-ttu-id="f4f93-127">詳細については、「 [web 上の Office を使用したポップアップブロックの処理](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f4f93-127">For more information, see [Handling pop-up blockers with Office on the web](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web).</span></span>|

<span data-ttu-id="f4f93-128">が`displayDialogAsync`呼び出されると、 [AsyncResult](/javascript/api/office/office.asyncresult)オブジェクトをコールバック関数に渡します。</span><span class="sxs-lookup"><span data-stu-id="f4f93-128">When `displayDialogAsync` is called, it passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to its callback function.</span></span> <span data-ttu-id="f4f93-129">呼び出しが成功すると、ダイアログボックスが開き、 `value` `AsyncResult`オブジェクトのプロパティは[dialog](/javascript/api/office/office.dialog)オブジェクトになります。</span><span class="sxs-lookup"><span data-stu-id="f4f93-129">When the call is successful, the dialog box is opened, and the `value` property of the `AsyncResult` object is a [Dialog](/javascript/api/office/office.dialog) object.</span></span> <span data-ttu-id="f4f93-130">この例については、「[送信情報をダイアログボックスからホストページに送信する](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f4f93-130">For an example of this, see [Send information from the dialog box to the host page](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page).</span></span> <span data-ttu-id="f4f93-131">呼び出しが失敗する`displayDialogAsync`と、ダイアログボックスは作成されず、 `status` `AsyncResult`オブジェクトのプロパティがに`Office.AsyncResultStatus.Failed`設定され、オブジェクト`error`のプロパティが設定されます。</span><span class="sxs-lookup"><span data-stu-id="f4f93-131">When the call to `displayDialogAsync` fails, the dialog box is not created, the `status` property of the `AsyncResult` object is set to `Office.AsyncResultStatus.Failed`, and the `error` property of the object is populated.</span></span> <span data-ttu-id="f4f93-132">をテストして、 `status`エラーが発生したときに応答するコールバックを常に提供する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f4f93-132">You should always provide a callback that tests the `status` and responds when it's an error.</span></span> <span data-ttu-id="f4f93-133">コード番号に関係なくエラーメッセージを報告する例については、次のコードを参照してください。</span><span class="sxs-lookup"><span data-stu-id="f4f93-133">For an example that reports the error message regardless of its code number, see the following code.</span></span> <span data-ttu-id="f4f93-134">(こ`showNotification`の記事で定義されていない関数は、エラーを表示またはログ記録します。</span><span class="sxs-lookup"><span data-stu-id="f4f93-134">(The `showNotification` function, not defined in this article, either displays or logs the error.</span></span> <span data-ttu-id="f4f93-135">アドイン内でこの関数を実装する方法の例については、「 [Office アドインダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f4f93-135">For an example of how you can implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).)</span></span>

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

## <a name="errors-and-events-in-the-dialog-box"></a><span data-ttu-id="f4f93-136">ダイアログボックスのエラーとイベント</span><span class="sxs-lookup"><span data-stu-id="f4f93-136">Errors and events in the dialog box</span></span>

<span data-ttu-id="f4f93-137">ダイアログボックス内の3つのエラーとイベントは`DialogEventReceived` 、ホストページでイベントを発生させます。</span><span class="sxs-lookup"><span data-stu-id="f4f93-137">Three errors and events in the dialog box will raise a `DialogEventReceived` event in the host page.</span></span> <span data-ttu-id="f4f93-138">ホストページについての通知については、「[ホストページからダイアログボックスを開く](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f4f93-138">For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span></span>

|<span data-ttu-id="f4f93-139">コード番号</span><span class="sxs-lookup"><span data-stu-id="f4f93-139">Code number</span></span>|<span data-ttu-id="f4f93-140">意味</span><span class="sxs-lookup"><span data-stu-id="f4f93-140">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="f4f93-141">12002</span><span class="sxs-lookup"><span data-stu-id="f4f93-141">12002</span></span>|<span data-ttu-id="f4f93-142">以下のいずれか:</span><span class="sxs-lookup"><span data-stu-id="f4f93-142">One of the following:</span></span><br> <span data-ttu-id="f4f93-143">- `displayDialogAsync` に渡された URL にページが存在しない。</span><span class="sxs-lookup"><span data-stu-id="f4f93-143">- No page exists at the URL that was passed to `displayDialogAsync`.</span></span><br> <span data-ttu-id="f4f93-144">-読み込みに`displayDialogAsync`渡されたページ。ただし、ダイアログボックスは、検出または読み込みできないページにリダイレクトされたか、または無効な構文の URL に転送されています。</span><span class="sxs-lookup"><span data-stu-id="f4f93-144">- The page that was passed to `displayDialogAsync` loaded, but the dialog box was then redirected to a page that it cannot find or load, or it has been directed to a URL with invalid syntax.</span></span>|
|<span data-ttu-id="f4f93-145">12003</span><span class="sxs-lookup"><span data-stu-id="f4f93-145">12003</span></span>|<span data-ttu-id="f4f93-p107">ダイアログ ボックスが HTTP プロトコルを使用している URL を指していました。HTTPS が必要です。</span><span class="sxs-lookup"><span data-stu-id="f4f93-p107">The dialog box was directed to a URL with the HTTP protocol. HTTPS is required.</span></span>|
|<span data-ttu-id="f4f93-148">12006</span><span class="sxs-lookup"><span data-stu-id="f4f93-148">12006</span></span>|<span data-ttu-id="f4f93-149">ダイアログボックスが閉じられました。通常は、ユーザーが [**閉じる**] ボタン**X**を選択したためです。</span><span class="sxs-lookup"><span data-stu-id="f4f93-149">The dialog box was closed, usually because the user chose the **Close** button **X**.</span></span>|

<span data-ttu-id="f4f93-p108">コードで、呼び出し内の `DialogEventReceived` イベントのハンドラーを `displayDialogAsync` に割り当てることができます。次に簡単な例を示します。</span><span class="sxs-lookup"><span data-stu-id="f4f93-p108">Your code can assign a handler for the `DialogEventReceived` event in the call to `displayDialogAsync`. The following is a simple example:</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

<span data-ttu-id="f4f93-152">各エラー コードのカスタム エラー メッセージを作成する `DialogEventReceived` イベントのハンドラーの例を、次に示します。</span><span class="sxs-lookup"><span data-stu-id="f4f93-152">For an example of a handler for the `DialogEventReceived` event that creates custom error messages for each error code, see the following example:</span></span>

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

<span data-ttu-id="f4f93-153">この方法でエラーを処理するサンプル アドインについては、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f4f93-153">For a sample add-in that handles errors in this way, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

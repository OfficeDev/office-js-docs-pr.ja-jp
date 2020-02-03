---
title: Office アドインで Office ダイアログ API を使用する
description: Office アドインのダイアログボックス作成の基本について説明します。
ms.date: 01/29/2020
localization_priority: Priority
ms.openlocfilehash: 13badafd0d3a6bb3fdf656b5caf93c9f514921d9
ms.sourcegitcommit: 4c9e02dac6f8030efc7415e699370753ec9415c8
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/01/2020
ms.locfileid: "41650006"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a><span data-ttu-id="3cd48-103">Office アドインで Office ダイアログ API を使用する</span><span class="sxs-lookup"><span data-stu-id="3cd48-103">Use the Office dialog API in Office Add-ins</span></span>

<span data-ttu-id="3cd48-104">[Office ダイアログ API](/javascript/api/office/office.ui) を使用して、Office アドインでダイアログ ボックスを開くことができます。</span><span class="sxs-lookup"><span data-stu-id="3cd48-104">You can use the [Office dialog API](/javascript/api/office/office.ui) to open dialog boxes in your Office Add-in.</span></span> <span data-ttu-id="3cd48-105">この記事では、Office アドインでダイアログ API を使用するためのガイダンスを提供します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-105">This article provides guidance for using the dialog API in your Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="3cd48-p102">ダイアログ API の現在のサポート状態に関する詳細は、「[ダイアログ API の要件セット](/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets)」を参照してください。現在、ダイアログ API は Word、Excel、PowerPoint、および Outlook でサポートされています。</span><span class="sxs-lookup"><span data-stu-id="3cd48-p102">For information about where the Dialog API is currently supported, see [Dialog API requirement sets](/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets). The Dialog API is currently supported for Word, Excel, PowerPoint, and Outlook.</span></span>

<span data-ttu-id="3cd48-108">ダイアログ API の主要なシナリオは、Google や Facebook、Microsoft Graph などのリソースで認証を有効にすることです。</span><span class="sxs-lookup"><span data-stu-id="3cd48-108">A primary scenario for the Dialog API is to enable authentication with a resource such as Google, Facebook, or Microsoft Graph.</span></span> <span data-ttu-id="3cd48-109">詳細については、この記事をよく読んだ*後*で「[Office Dialog API を使用して認証する](auth-with-office-dialog-api.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3cd48-109">For more information, see [Authenticate with the Office dialog API](auth-with-office-dialog-api.md) *after* you are familiar with this article.</span></span>

<span data-ttu-id="3cd48-110">作業ウィンドウ アドイン、コンテンツ アドイン、[アドイン コマンド](../design/add-in-commands.md)からダイアログ ボックスを開いて、次の操作を実行することを検討してください。</span><span class="sxs-lookup"><span data-stu-id="3cd48-110">Consider opening a dialog box from a task pane or content add-in or [add-in command](../design/add-in-commands.md) to do the following:</span></span>

- <span data-ttu-id="3cd48-111">作業ウィンドウに直接開くことができないサインイン ページを表示する。</span><span class="sxs-lookup"><span data-stu-id="3cd48-111">Display sign in pages that cannot be opened directly in a task pane.</span></span>
- <span data-ttu-id="3cd48-112">アドインでの作業用に画面領域を広げる (あるいは全画面表示)。</span><span class="sxs-lookup"><span data-stu-id="3cd48-112">Provide more screen space, or even a full screen, for some tasks in your add-in.</span></span>
- <span data-ttu-id="3cd48-113">ビデオが作業ウィンドウに限定されている場合に、小さすぎるビデオをホストする。</span><span class="sxs-lookup"><span data-stu-id="3cd48-113">Host a video that would be too small if confined to a task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="3cd48-114">UI 要素を重ねて表示することはお勧めできないため、シナリオで必要な場合を除き、作業ウィンドウでダイアログ ボックスを開かないようにします。</span><span class="sxs-lookup"><span data-stu-id="3cd48-114">Because overlapping UI elements are discouraged, avoid opening a dialog box from a task pane unless your scenario requires it.</span></span> <span data-ttu-id="3cd48-115">作業ウィンドウの表示領域の使用方法を検討するときには、作業ウィンドウはタブ表示できることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="3cd48-115">When you consider how to use the surface area of a task pane, note that task panes can be tabbed.</span></span> <span data-ttu-id="3cd48-116">例については、[Excel アドイン JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) のサンプルを参照してください。</span><span class="sxs-lookup"><span data-stu-id="3cd48-116">For an example, see the [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) sample.</span></span>

<span data-ttu-id="3cd48-117">次の画像は、ダイアログ ボックスの例を示します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-117">The following image shows an example of a dialog box.</span></span>

![アドイン コマンド](../images/auth-o-dialog-open.png)

<span data-ttu-id="3cd48-119">ダイアログ ボックスが常に画面の中央に開くことに注意してください。</span><span class="sxs-lookup"><span data-stu-id="3cd48-119">Note that the dialog box always opens in the center of the screen.</span></span> <span data-ttu-id="3cd48-120">ユーザーはダイアログ ボックスの移動とサイズ変更ができます。</span><span class="sxs-lookup"><span data-stu-id="3cd48-120">The user can move and resize it.</span></span> <span data-ttu-id="3cd48-121">ウィンドウは*モードレス*です。ホスト Office アプリケーションのドキュメントの操作と、作業ウィンドウのページ (存在する場合) の操作の両方を続行できます。</span><span class="sxs-lookup"><span data-stu-id="3cd48-121">The window is *nonmodal*--a user can continue to interact with both the document in the host Office application and with the page in the task pane, if there is one.</span></span>

## <a name="open-a-dialog-box-from-a-host-page"></a><span data-ttu-id="3cd48-122">ホスト ページからダイアログ ボックスを開く</span><span class="sxs-lookup"><span data-stu-id="3cd48-122">Open a dialog box from a host page</span></span>

<span data-ttu-id="3cd48-123">Office JavaScript API には、[Dialog](/javascript/api/office/office.dialog) オブジェクトと [Office.context.ui 名前空間](/javascript/api/office/office.ui)の 2 つの関数が含まれます。</span><span class="sxs-lookup"><span data-stu-id="3cd48-123">The Office JavaScript APIs include a [Dialog](/javascript/api/office/office.dialog) object and two functions in the [Office.context.ui namespace](/javascript/api/office/office.ui).</span></span>

<span data-ttu-id="3cd48-124">ダイアログ ボックスを開くには、コード (通常は作業ウィンドウ内のページ) で [displayDialogAsync](/javascript/api/office/office.ui) メソッドを呼び出して、開くリソースの URL を渡します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-124">To open a dialog box, your code, typically a page in a task pane, calls the [displayDialogAsync](/javascript/api/office/office.ui) method and passes to it the URL of the resource that you want to open.</span></span> <span data-ttu-id="3cd48-125">このメソッドを呼び出すページは、「ホスト ページ」と呼ばれます。</span><span class="sxs-lookup"><span data-stu-id="3cd48-125">The page on which this method is called is known as the "host page".</span></span> <span data-ttu-id="3cd48-126">たとえば、作業ウィンドウの index.html にあるスクリプトでこのメソッドを呼び出した場合は、index.html がメソッドが開いたダイアログ ボックスのホスト ページです。</span><span class="sxs-lookup"><span data-stu-id="3cd48-126">For example, if you call this method in script on index.html in a task pane, then index.html is the host page of the dialog box that the method opens.</span></span>

<span data-ttu-id="3cd48-127">ダイアログ ボックスで開かれるリソースは通常ページですが、MVC アプリケーションのコントローラー メソッド、ルート、Web サービス メソッド、またはその他のリソースの場合もあります。</span><span class="sxs-lookup"><span data-stu-id="3cd48-127">The resource that is opened in the dialog box is usually a page, but it can be a controller method in an MVC application, a route, a web service method, or any other resource.</span></span> <span data-ttu-id="3cd48-128">この記事では、'ページ' または 'Web サイト' とは、ダイアログ ボックス内のリソースを意味します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-128">In this article, 'page' or 'website' refers to the resource in the dialog box.</span></span> <span data-ttu-id="3cd48-129">次のコードは簡単な例を示しています。</span><span class="sxs-lookup"><span data-stu-id="3cd48-129">The following code is a simple example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - <span data-ttu-id="3cd48-130">URL には HTTP**S** プロトコルを使用します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-130">The URL uses the HTTP**S** protocol.</span></span> <span data-ttu-id="3cd48-131">これは、読み込まれる最初のページだけでなく、ダイアログ ボックスに読み込まれるすべてのページに対して必須です。</span><span class="sxs-lookup"><span data-stu-id="3cd48-131">This is mandatory for all pages loaded in a dialog box, not just the first page loaded.</span></span>
> - <span data-ttu-id="3cd48-132">ダイアログ ボックスのドメインはホスト ページのドメインと同じです。ホスト ページは、作業ウィンドウ内のページまたはアドイン コマンドの[関数ファイル](/office/dev/add-ins/reference/manifest/functionfile)にすることができます。</span><span class="sxs-lookup"><span data-stu-id="3cd48-132">The dialog box's domain is the same as the domain of the host page, which can be the page in a task pane or the [function file](/office/dev/add-ins/reference/manifest/functionfile) of an add-in command.</span></span> <span data-ttu-id="3cd48-133">ページ、コントローラーのメソッド、または `displayDialogAsync` メソッドに渡されるその他のリソースは、ホスト ページと同じドメインにある必要があります。</span><span class="sxs-lookup"><span data-stu-id="3cd48-133">This is required: the page, controller method, or other resource that is passed to the `displayDialogAsync` method must be in the same domain as the host page.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3cd48-134">ダイアログ ボックスで開くホスト ページとリソースのフル ドメインは、同じである必要があります。</span><span class="sxs-lookup"><span data-stu-id="3cd48-134">The host page and the resource that opens in the dialog box must have the same full domain.</span></span> <span data-ttu-id="3cd48-135">`displayDialogAsync` にアドインのドメインのサブドメインを渡そうとすると、正常に動作しません。</span><span class="sxs-lookup"><span data-stu-id="3cd48-135">If you attempt to pass `displayDialogAsync` a subdomain of the add-in's domain, it will not work.</span></span> <span data-ttu-id="3cd48-136">サブドメインを含む、フル ドメインが一致している必要があります。</span><span class="sxs-lookup"><span data-stu-id="3cd48-136">The full domain, including any subdomain, must match.</span></span>

<span data-ttu-id="3cd48-137">最初のページ (または他のリソース) が読み込まれると、ユーザーはリンクまたは他の UI を使用して HTTPS を使用する任意の Web サイト (または他のリソース) に移動できます。</span><span class="sxs-lookup"><span data-stu-id="3cd48-137">After the first page (or other resource) is loaded, a user can use links or other UI to navigate to any website (or other resource) that uses HTTPS.</span></span> <span data-ttu-id="3cd48-138">また、すぐに別のサイトにリダイレクトするように最初のページを設計することもできます。</span><span class="sxs-lookup"><span data-stu-id="3cd48-138">You can also design the first page to immediately redirect to another site.</span></span>

<span data-ttu-id="3cd48-139">既定では、ダイアログ ボックスのサイズはデバイス画面の高さと幅の 80% ですが、次の例に示すように、メソッドに構成オブジェクトを渡すことによってさまざまな割合を設定できます。</span><span class="sxs-lookup"><span data-stu-id="3cd48-139">By default, the dialog box will occupy 80% of the height and width of the device screen, but you can set different percentages by passing a configuration object to the method, as shown in the following example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

<span data-ttu-id="3cd48-140">これを実行するサンプル アドインについては、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3cd48-140">For a sample add-in that does this, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="3cd48-p112">全画面表示で効率的に操作するには、両方の値を 100% に設定します。(最大有効値は 99.5% であり、最大有効値にしても、ウィンドウは移動とサイズ変更が可能です。)</span><span class="sxs-lookup"><span data-stu-id="3cd48-p112">Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)</span></span>

> [!NOTE]
> <span data-ttu-id="3cd48-p113">ホスト ウィンドウから開くことができるのは、1 つのダイアログ ボックスのみです。別のダイアログ ボックスを開こうとすると、エラーが発生します。たとえば、ユーザーが作業ウィンドウからダイアログ ボックスを開いた場合には、作業ウィンドウの別のページから 2 番目のダイアログ ボックスを開くことができません。ただし、[アドイン コマンド](../design/add-in-commands.md)からダイアログ ボックスを開く場合は、選択するたびにコマンドによって新しい (ただし非表示の) HTML ファイルが開かれます。これにより、新しい (非表示) ホスト ウィンドウが作成されるため、これらの各ウィンドウは独自のダイアログ ボックスを起動できます。詳細については、「[displayDialogAsync のエラー](dialog-handle-errors-events.md#errors-from-displaydialogasync)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3cd48-p113">You can open only one dialog box from a host window. An attempt to open another dialog box generates an error. For example, if a user opens a dialog box from a task pane, she cannot open a second dialog box, from a different page in the task pane. However, when a dialog box is opened from an [add-in command](../design/add-in-commands.md), the command opens a new (but unseen) HTML file each time it is selected. This creates a new (unseen) host window, so each such window can launch its own dialog box. For more information, see [Errors from displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).</span></span>

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a><span data-ttu-id="3cd48-149">Office on the web のパフォーマンス オプションを利用する</span><span class="sxs-lookup"><span data-stu-id="3cd48-149">Take advantage of a performance option in Office on the web</span></span>

<span data-ttu-id="3cd48-150">`displayInIframe` プロパティは、`displayDialogAsync` に渡すことのできる構成オブジェクトの追加のプロパティです。</span><span class="sxs-lookup"><span data-stu-id="3cd48-150">The `displayInIframe` property is an additional property in the configuration object that you can pass to `displayDialogAsync`.</span></span> <span data-ttu-id="3cd48-151">このプロパティを `true` に設定し、Office on the web で開いたドキュメントでアドインを実行している場合、ダイアログ ボックスは浮動の iframe で開き、独立したウィンドウでは開きません (この方が速く開きます)。</span><span class="sxs-lookup"><span data-stu-id="3cd48-151">When this property is set to `true`, and the add-in is running in a document opened in Office on the web, the dialog box will open as a floating iframe rather than an independent window, which makes it open faster.</span></span> <span data-ttu-id="3cd48-152">例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-152">The following is an example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

<span data-ttu-id="3cd48-153">既定値は `false` です。これはプロパティを完全に省略した場合と同じ状態です。</span><span class="sxs-lookup"><span data-stu-id="3cd48-153">The default value is `false`, which is the same as omitting the property entirely.</span></span> <span data-ttu-id="3cd48-154">アドインが Office on the web で実行されていない場合、`displayInIframe` は無視されます。</span><span class="sxs-lookup"><span data-stu-id="3cd48-154">If the add-in is not running in Office on the web, the `displayInIframe` is ignored.</span></span>

> [!NOTE]
> <span data-ttu-id="3cd48-155">どの時点であれ、iframe で開けないページにダイアログ ボックスがリダイレクトされることになる場合は、`displayInIframe: true` を使用すべきでは**ありません**。</span><span class="sxs-lookup"><span data-stu-id="3cd48-155">You should **not** use `displayInIframe: true` if the dialog box will at any point redirect to a page that cannot be opened in an iframe.</span></span> <span data-ttu-id="3cd48-156">たとえば、Google や Microsoft アカウントなどの多くの一般的な Web サービスのサインイン ページは iframe で開くことができません。</span><span class="sxs-lookup"><span data-stu-id="3cd48-156">For example, the sign in pages of many popular web services, such as Google and Microsoft Account, cannot be opened in an iframe.</span></span>

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a><span data-ttu-id="3cd48-157">ダイアログ ボックスからホスト ページに情報を送信する</span><span class="sxs-lookup"><span data-stu-id="3cd48-157">Send information from the dialog box to the host page</span></span>

<span data-ttu-id="3cd48-158">ダイアログ ボックスは、以下の場合を除いて、作業ウィンドウのホスト ページと通信できません。</span><span class="sxs-lookup"><span data-stu-id="3cd48-158">The dialog box cannot communicate with the host page in the task pane unless:</span></span>

- <span data-ttu-id="3cd48-159">ダイアログ ボックスの現在のページがホスト ページと同じドメインにある。</span><span class="sxs-lookup"><span data-stu-id="3cd48-159">The current page in the dialog box is in the same domain as the host page.</span></span>
- <span data-ttu-id="3cd48-p117">Office JavaScript ライブラリがページに読み込まれている。(Office JavaScript ライブラリを使用するすべてのページと同様に、ページのスクリプトは `Office.initialize` プロパティにメソッドを割り当てる必要があります (空のメソッドでもかまいません)。詳細については、「[アドインの初期化](understanding-the-javascript-api-for-office.md#initializing-your-add-in)」を参照してください。)</span><span class="sxs-lookup"><span data-stu-id="3cd48-p117">The Office JavaScript library is loaded in the page. (Like any page that uses the Office JavaScript library, script for the page must assign a method to the `Office.initialize` property, although it can be an empty method. For details, see [Initializing your add-in](understanding-the-javascript-api-for-office.md#initializing-your-add-in).)</span></span>

<span data-ttu-id="3cd48-163">ダイアログ ボックスのコードは、[messageParent](/javascript/api/office/office.ui#messageparent-message-) 関数を使用して、ブール値または文字列メッセージのいずれかをホスト ページに送信します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-163">Code in the dialog box uses the [messageParent](/javascript/api/office/office.ui#messageparent-message-) function to send either a Boolean value or a string message to the host page.</span></span> <span data-ttu-id="3cd48-164">文字列には、単語、文、XML BLOB、文字列に変換された JSON、文字列にシリアル化できるすべてのものを指定できます。</span><span class="sxs-lookup"><span data-stu-id="3cd48-164">The string can be a word, sentence, XML blob, stringified JSON, or anything else that can be serialized to a string.</span></span> <span data-ttu-id="3cd48-165">例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-165">The following is an example:</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!NOTE]
> - <span data-ttu-id="3cd48-p119">`messageParent` 関数は、ダイアログ ボックスで呼び出すことができる、*ただ* 2 つの Office API のうちの 1 つです。もう 1 つは `Office.context.requirements.isSetSupported` です。詳細は、「[Office のホストと API の要件を指定する](specify-office-hosts-and-api-requirements.md)」を参照してださい。</span><span class="sxs-lookup"><span data-stu-id="3cd48-p119">The `messageParent` function is one of *only* two Office APIs that can be called in the dialog box. The other is `Office.context.requirements.isSetSupported`. For information about it, see [Specify Office hosts and API requirements](specify-office-hosts-and-api-requirements.md).</span></span>
> - <span data-ttu-id="3cd48-169">`messageParent` 関数を呼び出せるのは、ホスト ページと同じドメイン (プロトコルとポートを含む) を持つページ上のみです。</span><span class="sxs-lookup"><span data-stu-id="3cd48-169">The `messageParent` function can only be called on a page with the same domain (including protocol and port) as the host page.</span></span>

<span data-ttu-id="3cd48-170">次の例では、`googleProfile` は文字列に変換されたバージョンのユーザーの Google プロファイルです。</span><span class="sxs-lookup"><span data-stu-id="3cd48-170">In the next example, `googleProfile` is a stringified version of the user's Google profile.</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

<span data-ttu-id="3cd48-p120">ホスト ページは、メッセージを受信するように構成する必要があります。これを構成するには、`displayDialogAsync` の元の呼び出しにコールバック パラメーターを追加します。コールバックはハンドラーを `DialogMessageReceived` イベントに割り当てます。次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-p120">The host page must be configured to receive the message. You do this by adding a callback parameter to the original call of `displayDialogAsync`. The callback assigns a handler to the `DialogMessageReceived` event. The following is an example:</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20},
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);
```

> [!NOTE]
> - <span data-ttu-id="3cd48-175">Office は [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトをコールバックに渡します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-175">Office passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to the callback.</span></span> <span data-ttu-id="3cd48-176">Office はダイアログ ボックスを開こうとした結果を表します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-176">It represents the result of the attempt to open the dialog box.</span></span> <span data-ttu-id="3cd48-177">ただし、ダイアログ ボックスでのイベントの結果は表しません。</span><span class="sxs-lookup"><span data-stu-id="3cd48-177">It does not represent the outcome of any events in the dialog box.</span></span> <span data-ttu-id="3cd48-178">この違いの詳細については、「[エラーとイベントの処理](dialog-handle-errors-events.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3cd48-178">For more on this distinction, see [Handle errors and events](dialog-handle-errors-events.md).</span></span>
> - <span data-ttu-id="3cd48-179">`asyncResult` の `value` プロパティは [Dialog](/javascript/api/office/office.dialog) オブジェクトに設置されます。このオブジェクトはダイアログ ボックスの実行コンテキストではなく、ホスト ページに存在します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-179">The `value` property of the `asyncResult` is set to a [Dialog](/javascript/api/office/office.dialog) object, which exists in the host page, not in the dialog box's execution context.</span></span>
> - <span data-ttu-id="3cd48-p122">`processMessage` はイベントを処理する関数です。任意の名前を指定できます。</span><span class="sxs-lookup"><span data-stu-id="3cd48-p122">The `processMessage` is the function that handles the event. You can give it any name you want.</span></span>
> - <span data-ttu-id="3cd48-182">`dialog` 変数は、`processMessage` でも参照されるため、コールバックよりも広い範囲で宣言されます。</span><span class="sxs-lookup"><span data-stu-id="3cd48-182">The `dialog` variable is declared at a wider scope than the callback because it is also referenced in `processMessage`.</span></span>

<span data-ttu-id="3cd48-183">`DialogMessageReceived` イベントのハンドラーの簡単な例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-183">The following is a simple example of a handler for the `DialogMessageReceived` event:</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - <span data-ttu-id="3cd48-184">Office は `arg` オブジェクトをハンドラーに渡します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-184">Office passes the `arg` object to the handler.</span></span> <span data-ttu-id="3cd48-185">その `message` プロパティは、ダイアログ ボックスの `messageParent` の呼び出しで送信されるブール値または文字列です。</span><span class="sxs-lookup"><span data-stu-id="3cd48-185">Its `message` property is the Boolean or string sent by the call of `messageParent` in the dialog box.</span></span> <span data-ttu-id="3cd48-186">この例では、Microsoft アカウントまたは Google などのサービスからのユーザーのプロファイルの文字列に変換された表記です。このため、`JSON.parse` を含むオブジェクトに逆シリアル化されます。</span><span class="sxs-lookup"><span data-stu-id="3cd48-186">In this example, it is a stringified representation of a user's profile from a service such as Microsoft Account or Google, so it is deserialized back to an object with `JSON.parse`.</span></span>
> - <span data-ttu-id="3cd48-p124">`showUserName` 実装は表示されません。作業ウィンドウ上に個人用のウェルカム メッセージが表示される場合があります。</span><span class="sxs-lookup"><span data-stu-id="3cd48-p124">The `showUserName` implementation is not shown. It might display a personalized welcome message on the task pane.</span></span>

<span data-ttu-id="3cd48-189">ダイアログ ボックスのユーザー操作が完了すると、次の例に示すようにメッセージ ハンドラーはダイアログ ボックスを閉じます。</span><span class="sxs-lookup"><span data-stu-id="3cd48-189">When the user interaction with the dialog box is completed, your message handler should close the dialog box, as shown in this example.</span></span>

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
> - <span data-ttu-id="3cd48-190">`dialog` オブジェクトは `displayDialogAsync` の呼び出しによって返されるものと同じである必要があります。</span><span class="sxs-lookup"><span data-stu-id="3cd48-190">The `dialog` object must be the same one that is returned by the call of `displayDialogAsync`.</span></span>
> - <span data-ttu-id="3cd48-191">`dialog.close` の呼び出しは、直ちにダイアログ ボックスを閉じるよう Office に指示します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-191">The call of `dialog.close` tells Office to immediately close the dialog box.</span></span>

<span data-ttu-id="3cd48-192">これらの手法を使用するサンプル アドインについては、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3cd48-192">For a sample add-in that uses these techniques, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="3cd48-p125">メッセージを受信した後、アドインで作業ウィンドウの別のページを開く必要がある場合は、ハンドラーの最後の行として `window.location.replace` メソッド (または `window.location.href`) を使用できます。次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-p125">If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example:</span></span>

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

<span data-ttu-id="3cd48-195">これを実行するアドインの例については、「[Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)」 (PowerPoint アドインで Microsoft Graph を使用した Excel グラフの挿入) のサンプルを参照してください。</span><span class="sxs-lookup"><span data-stu-id="3cd48-195">For an example of an add-in that does this, see the [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) sample.</span></span>

### <a name="conditional-messaging"></a><span data-ttu-id="3cd48-196">条件付きのメッセージング</span><span class="sxs-lookup"><span data-stu-id="3cd48-196">Conditional messaging</span></span>

<span data-ttu-id="3cd48-p126">ダイアログ ボックスから複数の `messageParent` 呼び出しを送信できますが、`DialogMessageReceived` イベントのホスト ページにあるハンドラーは 1 つのみのため、ハンドラーは条件ロジックを使用してさまざまなメッセージを区別する必要があります。たとえば、ユーザーに対して Microsoft アカウントまたは Google などの ID プロバイダーにサインインするよう求めるダイアログ ボックスが表示されると、ダイアログ ボックスはユーザーのプロファイルをメッセージとして送信します。認証が失敗した場合、次の例のように、ダイアログ ボックスはホスト ページにエラー情報を送信します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-p126">Because you can send multiple `messageParent` calls from the dialog box, but you have only one handler in the host page for the `DialogMessageReceived` event, the handler must use conditional logic to distinguish different messages. For example, if the dialog box prompts a user to sign in to an identity provider such as Microsoft Account or Google, it sends the user's profile as a message. If authentication fails, the dialog box sends error information to the host page, as in the following example:</span></span>

```js
if (loginSuccess) {
    var userProfile = getProfile();
    var messageObject = {messageType: "signinSuccess", profile: userProfile};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
} else {
    var errorDetails = getError();
    var messageObject = {messageType: "signinFailure", error: errorDetails};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

> [!NOTE]
> - <span data-ttu-id="3cd48-200">`loginSuccess` 変数は、ID プロバイダーからの HTTP 応答を読み取ることによって初期化されます。</span><span class="sxs-lookup"><span data-stu-id="3cd48-200">The `loginSuccess` variable would be initialized by reading the HTTP response from the identity provider.</span></span>
> - <span data-ttu-id="3cd48-p127">`getProfile` 関数と `getError` 関数の実装は表示されません。両方の関数はそれぞれ、クエリ パラメーターまたは HTTP 応答の本文からデータを取得します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-p127">The the implementation of the `getProfile` and `getError` functions are not not shown. They each get data from a query parameter or from the body of the HTTP response.</span></span>
> - <span data-ttu-id="3cd48-p128">サインインが成功したかどうかに応じて、さまざまな種類の匿名のオブジェクトが送信されます。両方の関数に `messageType` プロパティがありますが、一方には `profile` プロパティ、もう一方には `error` プロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="3cd48-p128">Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.</span></span>

<span data-ttu-id="3cd48-p129">次の例に示すように、ホスト ページのハンドラー コードは分岐に `messageType` プロパティの値を使用します。`showUserName` 関数は上記の例と同じであり、`showNotification` 関数はホスト ページの UI にエラーを表示することに注意してください。</span><span class="sxs-lookup"><span data-stu-id="3cd48-p129">The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "signinSuccess") {
        dialog.close();
        showUserName(messageFromDialog.profile.name);
        window.location.replace("/newPage.html");
    } else {
        dialog.close();
        showNotification("Unable to authenticate user: " + messageFromDialog.error);
    }
}
```

> [!NOTE]
> <span data-ttu-id="3cd48-207">`showNotification`の実装は、この記事のサンプル コードでは表示されません。</span><span class="sxs-lookup"><span data-stu-id="3cd48-207">The `showNotification` implementation is not shown in the sample code provided by this article.</span></span> <span data-ttu-id="3cd48-208">アドインでこの関数を実装する方法の例は、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3cd48-208">For an example of how you might implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

## <a name="pass-information-to-the-dialog-box"></a><span data-ttu-id="3cd48-209">情報をダイアログ ボックスに渡す</span><span class="sxs-lookup"><span data-stu-id="3cd48-209">Pass information to the dialog box</span></span>

<span data-ttu-id="3cd48-p131">ホスト ページがダイアログ ボックスに情報を渡す必要がある場合もあります。これは主に 2 つの方法で実行することができます。</span><span class="sxs-lookup"><span data-stu-id="3cd48-p131">Sometimes the host page needs to pass information to the dialog box. You can do this in two primary ways:</span></span>

- <span data-ttu-id="3cd48-212">`displayDialogAsync` に渡される URL にクエリ パラメーターを追加します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-212">Add query parameters to the URL that is passed to `displayDialogAsync`.</span></span>
- <span data-ttu-id="3cd48-213">ホスト ウィンドウとダイアログ ボックスの両方にアクセス可能な場所に情報を格納します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-213">Store the information somewhere that is accessible to both the host window and dialog box.</span></span> <span data-ttu-id="3cd48-214">2 つのウィンドウは共通のセッション ストレージを共有しませんが、ポート番号 (存在する場合) を含む*ドメインが同じである場合*は、共通の[ローカル ストレージ](https://www.w3schools.com/html/html5_webstorage.asp)を共有します。\*</span><span class="sxs-lookup"><span data-stu-id="3cd48-214">The two windows do not share a common session storage, but *if they have the same domain* (including port number, if any), they share a common [Local Storage](https://www.w3schools.com/html/html5_webstorage.asp).\*</span></span>

> [!NOTE]
> <span data-ttu-id="3cd48-215">\* トークン処理の戦略に影響を与えるバグがあります。</span><span class="sxs-lookup"><span data-stu-id="3cd48-215">\* There is a bug that will effect your strategy for token handling.</span></span> <span data-ttu-id="3cd48-216">Safari または Microsoft Edge ブラウザーの **Office on the web** でアドインを実行している場合、ダイアログ ボックスとタスク ウィンドウは同じローカル ストレージを共有しないため、これらの間の通信に使用できません。</span><span class="sxs-lookup"><span data-stu-id="3cd48-216">If the add-in is running in **Office on the web** in either the Safari or Edge browser, the dialog box and task pane do not share the same Local Storage, so it cannot be used to communicate between them.</span></span>

### <a name="use-local-storage"></a><span data-ttu-id="3cd48-217">ローカル ストレージの使用</span><span class="sxs-lookup"><span data-stu-id="3cd48-217">Use local storage</span></span>

<span data-ttu-id="3cd48-218">ローカル ストレージを使用するには、次の例に示すように、`displayDialogAsync` 呼び出しの前に、コードはホスト ページで `window.localStorage` オブジェクトの `setItem` メソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-218">To use local storage, your code calls the `setItem` method of the `window.localStorage` object in the host page before the `displayDialogAsync` call, as in the following example:</span></span>

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

<span data-ttu-id="3cd48-219">ダイアログ ボックス内のコードは、次の例に示すように、必要に応じて項目を読み取ります。</span><span class="sxs-lookup"><span data-stu-id="3cd48-219">Code in the dialog box reads the item when it's needed, as in the following example:</span></span>

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

### <a name="use-query-parameters"></a><span data-ttu-id="3cd48-220">クエリ パラメーターの使用</span><span class="sxs-lookup"><span data-stu-id="3cd48-220">Use query parameters</span></span>

<span data-ttu-id="3cd48-221">次の例では、クエリ パラメーターを使用してデータを渡す方法を示します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-221">The following example shows how to pass data with a query parameter:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

<span data-ttu-id="3cd48-222">この手法を使用するサンプルについては、「[PowerPoint アドインで Microsoft Graph を使用した Excel グラフの挿入](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3cd48-222">For a sample that uses this technique, see [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span></span>

<span data-ttu-id="3cd48-223">ダイアログ ボックス内のコードは、URL を解析し、パラメーター値を読み取ることができます。</span><span class="sxs-lookup"><span data-stu-id="3cd48-223">Code in your dialog box can parse the URL and read the parameter value.</span></span>

> [!NOTE]
> <span data-ttu-id="3cd48-p134">Office は、`displayDialogAsync` に渡される URL に `_host_info` というクエリ パラメーターを自動的に追加します (カスタム クエリ パラメーターが存在する場合は、その後に追加されます。ダイアログ ボックスが移動する先の後続の URL には追加されません)。Microsoft は、将来、この値の内容を変更したり、完全に削除したりする可能性があるため、コードでこの値の内容を読み取らないでください。ダイアログ ボックスのセッション ストレージには、同じ値が追加されます。この場合も、*コードではこの値に対する読み取りも書き込みも行わないでください*。</span><span class="sxs-lookup"><span data-stu-id="3cd48-p134">Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`. (It is appended after your custom query parameters, if any. It is not appended to any subsequent URLs that the dialog box navigates to.) Microsoft may change the content of this value, or remove it entirely, in the future, so your code should not read it. The same value is added to the dialog box's session storage. Again, *your code should neither read nor write to this value*.</span></span>

## <a name="closing-the-dialog-box"></a><span data-ttu-id="3cd48-229">ダイアログ ボックスを閉じる</span><span class="sxs-lookup"><span data-stu-id="3cd48-229">Closing the dialog box</span></span>

<span data-ttu-id="3cd48-p135">ダイアログ ボックスを閉じるボタンをダイアログ ボックス内に実装できます。これを実行するには、ボタンのクリック イベント ハンドラーは `messageParent` を使用して、ボタンがクリックされたことをホスト ページに通知する必要があります。次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="3cd48-p135">You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example:</span></span>

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

<span data-ttu-id="3cd48-233">`DialogMessageReceived` のホスト ページ ハンドラーは、この例のように `dialog.close` を呼び出します </span><span class="sxs-lookup"><span data-stu-id="3cd48-233">The host page handler for `DialogMessageReceived` would call `dialog.close`, as in this example.</span></span> <span data-ttu-id="3cd48-234">(`dialog` オブジェクトを初期化する方法を示す前述の例を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="3cd48-234">(See previous examples that show how the `dialog` object is initialized.)</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

<span data-ttu-id="3cd48-235">独自の終了ダイアログ UI がない場合でも、エンド ユーザーは右上隅にある **X** を選択してダイアログ ボックスを閉じることができます。</span><span class="sxs-lookup"><span data-stu-id="3cd48-235">Even when you don't have your own close-dialog UI, an end user can close the dialog box by choosing the **X** in the upper-right corner.</span></span> <span data-ttu-id="3cd48-236">この操作により `DialogEventReceived` イベントがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="3cd48-236">This action triggers the `DialogEventReceived` event.</span></span> <span data-ttu-id="3cd48-237">イベントがトリガーされたときに、ホスト ウィンドウに通知する必要がある場合、ホスト ウィンドウはこのイベントのハンドラーを宣言する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3cd48-237">If your host pane needs to know when this happens, it should declare a handler for this event.</span></span> <span data-ttu-id="3cd48-238">詳細については、「[ダイアログ ボックスでのエラーとイベント](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box)」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="3cd48-238">See the section [Errors and events in the dialog box](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box) for details.</span></span>

## <a name="advanced-topics-and-special-scenarios"></a><span data-ttu-id="3cd48-239">高度なトピックと特殊なシナリオ</span><span class="sxs-lookup"><span data-stu-id="3cd48-239">Advanced topics and special scenarios</span></span>

### <a name="use-the-dialog-api-to-show-a-video"></a><span data-ttu-id="3cd48-240">ダイアログ API を使用してビデオを表示する</span><span class="sxs-lookup"><span data-stu-id="3cd48-240">Use the Dialog API to show a video</span></span>

<span data-ttu-id="3cd48-241">「[Office ダイアログ ボックスを使用してビデオを表示する](dialog-video.md)」を参照してください 。</span><span class="sxs-lookup"><span data-stu-id="3cd48-241">See [Use the Office dialog box to show a video](dialog-video.md).</span></span>

### <a name="use-the-dialog-apis-in-an-authentication-flow"></a><span data-ttu-id="3cd48-242">認証フローでダイアログ API を使用する</span><span class="sxs-lookup"><span data-stu-id="3cd48-242">Use the Dialog APIs in an authentication flow</span></span>

<span data-ttu-id="3cd48-243">「[Office Dialog API を使用して認証する](auth-with-office-dialog-api.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3cd48-243">See [Authenticate with the Office dialog API](auth-with-office-dialog-api.md).</span></span>

### <a name="using-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a><span data-ttu-id="3cd48-244">単一ページ アプリケーションとクライアント側ルーティングで Office ダイアログ API を使用する</span><span class="sxs-lookup"><span data-stu-id="3cd48-244">Using the Office dialog API with single-page applications and client-side routing</span></span>

<span data-ttu-id="3cd48-245">Office ダイアログ API を使用する場合は、SPA およびクライアント側のルーティングを慎重に行う必要があります。</span><span class="sxs-lookup"><span data-stu-id="3cd48-245">SPAs and client-side routing need to be handled with care when you are using the Office dialog API.</span></span> <span data-ttu-id="3cd48-246">「[SPA で Office ダイアログ API を使用する場合のベスト プラクティス](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3cd48-246">Please see [Best practices for using the Office dialog API in an SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).</span></span>

### <a name="error-and-event-handling"></a><span data-ttu-id="3cd48-247">エラーとイベントの処理</span><span class="sxs-lookup"><span data-stu-id="3cd48-247">Error and event handling</span></span>

<span data-ttu-id="3cd48-248">詳細については、「[Office ダイアログ ボックスでのエラーとイベントの処理](dialog-handle-errors-events.md)」を参照し ます。</span><span class="sxs-lookup"><span data-stu-id="3cd48-248">See [Handling errors and events in the Office dialog box](dialog-handle-errors-events.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="3cd48-249">次の手順</span><span class="sxs-lookup"><span data-stu-id="3cd48-249">Next steps</span></span>

<span data-ttu-id="3cd48-250">Office ダイアログ API に関するヒントとヘスと プラクティスの詳細については、「[Office ダイアログ API のベスト プラクティスとルール](dialog-best-practices.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3cd48-250">Learn about gotchas and best practices for the Office dialog API in [Best practices and rules for the Office dialog API](dialog-best-practices.md).</span></span>
---
title: Office アドインで Office ダイアログ API を使用する
description: Office アドインのダイアログボックス作成の基本について説明します。
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 5220d4876d0a8de9c731d2879f0bcb5e669066cd
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626464"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a><span data-ttu-id="78716-103">Office アドインで Office ダイアログ API を使用する</span><span class="sxs-lookup"><span data-stu-id="78716-103">Use the Office dialog API in Office Add-ins</span></span>

<span data-ttu-id="78716-104">[Office ダイアログ API](/javascript/api/office/office.ui) を使用して、Office アドインでダイアログ ボックスを開くことができます。</span><span class="sxs-lookup"><span data-stu-id="78716-104">You can use the [Office dialog API](/javascript/api/office/office.ui) to open dialog boxes in your Office Add-in.</span></span> <span data-ttu-id="78716-105">この記事では、Office アドインでダイアログ API を使用するためのガイダンスを提供します。</span><span class="sxs-lookup"><span data-stu-id="78716-105">This article provides guidance for using the dialog API in your Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="78716-p102">ダイアログ API の現在のサポート状態に関する詳細は、「[ダイアログ API の要件セット](../reference/requirement-sets/dialog-api-requirement-sets.md)」を参照してください。現在、ダイアログ API は Word、Excel、PowerPoint、および Outlook でサポートされています。</span><span class="sxs-lookup"><span data-stu-id="78716-p102">For information about where the Dialog API is currently supported, see [Dialog API requirement sets](../reference/requirement-sets/dialog-api-requirement-sets.md). The Dialog API is currently supported for Word, Excel, PowerPoint, and Outlook.</span></span>

<span data-ttu-id="78716-108">ダイアログ API の主要なシナリオは、Google や Facebook、Microsoft Graph などのリソースで認証を有効にすることです。</span><span class="sxs-lookup"><span data-stu-id="78716-108">A primary scenario for the Dialog API is to enable authentication with a resource such as Google, Facebook, or Microsoft Graph.</span></span> <span data-ttu-id="78716-109">詳細については、この記事をよく読んだ*後*で「[Office Dialog API を使用して認証する](auth-with-office-dialog-api.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="78716-109">For more information, see [Authenticate with the Office dialog API](auth-with-office-dialog-api.md) *after* you are familiar with this article.</span></span>

<span data-ttu-id="78716-110">作業ウィンドウ アドイン、コンテンツ アドイン、[アドイン コマンド](../design/add-in-commands.md)からダイアログ ボックスを開いて、次の操作を実行することを検討してください。</span><span class="sxs-lookup"><span data-stu-id="78716-110">Consider opening a dialog box from a task pane or content add-in or [add-in command](../design/add-in-commands.md) to do the following:</span></span>

- <span data-ttu-id="78716-111">作業ウィンドウに直接開くことができないサインイン ページを表示する。</span><span class="sxs-lookup"><span data-stu-id="78716-111">Display sign in pages that cannot be opened directly in a task pane.</span></span>
- <span data-ttu-id="78716-112">アドインでの作業用に画面領域を広げる (あるいは全画面表示)。</span><span class="sxs-lookup"><span data-stu-id="78716-112">Provide more screen space, or even a full screen, for some tasks in your add-in.</span></span>
- <span data-ttu-id="78716-113">ビデオが作業ウィンドウに限定されている場合に、小さすぎるビデオをホストする。</span><span class="sxs-lookup"><span data-stu-id="78716-113">Host a video that would be too small if confined to a task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="78716-114">UI 要素を重ねて表示することはお勧めできないため、シナリオで必要な場合を除き、作業ウィンドウでダイアログ ボックスを開かないようにします。</span><span class="sxs-lookup"><span data-stu-id="78716-114">Because overlapping UI elements are discouraged, avoid opening a dialog box from a task pane unless your scenario requires it.</span></span> <span data-ttu-id="78716-115">作業ウィンドウの表示領域の使用方法を検討するときには、作業ウィンドウはタブ表示できることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="78716-115">When you consider how to use the surface area of a task pane, note that task panes can be tabbed.</span></span> <span data-ttu-id="78716-116">例については、[Excel アドイン JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) のサンプルを参照してください。</span><span class="sxs-lookup"><span data-stu-id="78716-116">For an example, see the [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) sample.</span></span>

<span data-ttu-id="78716-117">次の画像は、ダイアログ ボックスの例を示します。</span><span class="sxs-lookup"><span data-stu-id="78716-117">The following image shows an example of a dialog box.</span></span>

![アドイン コマンド](../images/auth-o-dialog-open.png)

<span data-ttu-id="78716-119">ダイアログ ボックスが常に画面の中央に開くことに注意してください。</span><span class="sxs-lookup"><span data-stu-id="78716-119">Note that the dialog box always opens in the center of the screen.</span></span> <span data-ttu-id="78716-120">ユーザーはダイアログ ボックスの移動とサイズ変更ができます。</span><span class="sxs-lookup"><span data-stu-id="78716-120">The user can move and resize it.</span></span> <span data-ttu-id="78716-121">ウィンドウが非 *モーダル*である--ユーザーは、Office アプリケーション内のドキュメントと、作業ウィンドウのページがある場合は、そのページを引き続き操作できます。</span><span class="sxs-lookup"><span data-stu-id="78716-121">The window is *nonmodal*--a user can continue to interact with both the document in the Office application and with the page in the task pane, if there is one.</span></span>

## <a name="open-a-dialog-box-from-a-host-page"></a><span data-ttu-id="78716-122">ホスト ページからダイアログ ボックスを開く</span><span class="sxs-lookup"><span data-stu-id="78716-122">Open a dialog box from a host page</span></span>

<span data-ttu-id="78716-123">Office JavaScript API には、[Dialog](/javascript/api/office/office.dialog) オブジェクトと [Office.context.ui 名前空間](/javascript/api/office/office.ui)の 2 つの関数が含まれます。</span><span class="sxs-lookup"><span data-stu-id="78716-123">The Office JavaScript APIs include a [Dialog](/javascript/api/office/office.dialog) object and two functions in the [Office.context.ui namespace](/javascript/api/office/office.ui).</span></span>

<span data-ttu-id="78716-124">ダイアログ ボックスを開くには、コード (通常は作業ウィンドウ内のページ) で [displayDialogAsync](/javascript/api/office/office.ui) メソッドを呼び出して、開くリソースの URL を渡します。</span><span class="sxs-lookup"><span data-stu-id="78716-124">To open a dialog box, your code, typically a page in a task pane, calls the [displayDialogAsync](/javascript/api/office/office.ui) method and passes to it the URL of the resource that you want to open.</span></span> <span data-ttu-id="78716-125">このメソッドを呼び出すページは、「ホスト ページ」と呼ばれます。</span><span class="sxs-lookup"><span data-stu-id="78716-125">The page on which this method is called is known as the "host page".</span></span> <span data-ttu-id="78716-126">たとえば、作業ウィンドウの index.html にあるスクリプトでこのメソッドを呼び出した場合は、index.html がメソッドが開いたダイアログ ボックスのホスト ページです。</span><span class="sxs-lookup"><span data-stu-id="78716-126">For example, if you call this method in script on index.html in a task pane, then index.html is the host page of the dialog box that the method opens.</span></span>

<span data-ttu-id="78716-127">ダイアログ ボックスで開かれるリソースは通常ページですが、MVC アプリケーションのコントローラー メソッド、ルート、Web サービス メソッド、またはその他のリソースの場合もあります。</span><span class="sxs-lookup"><span data-stu-id="78716-127">The resource that is opened in the dialog box is usually a page, but it can be a controller method in an MVC application, a route, a web service method, or any other resource.</span></span> <span data-ttu-id="78716-128">この記事では、'ページ' または 'Web サイト' とは、ダイアログ ボックス内のリソースを意味します。</span><span class="sxs-lookup"><span data-stu-id="78716-128">In this article, 'page' or 'website' refers to the resource in the dialog box.</span></span> <span data-ttu-id="78716-129">次のコードは簡単な例を示しています。</span><span class="sxs-lookup"><span data-stu-id="78716-129">The following code is a simple example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - <span data-ttu-id="78716-130">URL には HTTP**S** プロトコルを使用します。</span><span class="sxs-lookup"><span data-stu-id="78716-130">The URL uses the HTTP**S** protocol.</span></span> <span data-ttu-id="78716-131">これは、読み込まれる最初のページだけでなく、ダイアログ ボックスに読み込まれるすべてのページに対して必須です。</span><span class="sxs-lookup"><span data-stu-id="78716-131">This is mandatory for all pages loaded in a dialog box, not just the first page loaded.</span></span>
> - <span data-ttu-id="78716-132">ダイアログ ボックスのドメインはホスト ページのドメインと同じです。ホスト ページは、作業ウィンドウ内のページまたはアドイン コマンドの[関数ファイル](../reference/manifest/functionfile.md)にすることができます。</span><span class="sxs-lookup"><span data-stu-id="78716-132">The dialog box's domain is the same as the domain of the host page, which can be the page in a task pane or the [function file](../reference/manifest/functionfile.md) of an add-in command.</span></span> <span data-ttu-id="78716-133">ページ、コントローラーのメソッド、または `displayDialogAsync` メソッドに渡されるその他のリソースは、ホスト ページと同じドメインにある必要があります。</span><span class="sxs-lookup"><span data-stu-id="78716-133">This is required: the page, controller method, or other resource that is passed to the `displayDialogAsync` method must be in the same domain as the host page.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="78716-134">ダイアログ ボックスで開くホスト ページとリソースのフル ドメインは、同じである必要があります。</span><span class="sxs-lookup"><span data-stu-id="78716-134">The host page and the resource that opens in the dialog box must have the same full domain.</span></span> <span data-ttu-id="78716-135">`displayDialogAsync` にアドインのドメインのサブドメインを渡そうとすると、正常に動作しません。</span><span class="sxs-lookup"><span data-stu-id="78716-135">If you attempt to pass `displayDialogAsync` a subdomain of the add-in's domain, it will not work.</span></span> <span data-ttu-id="78716-136">サブドメインを含む、フル ドメインが一致している必要があります。</span><span class="sxs-lookup"><span data-stu-id="78716-136">The full domain, including any subdomain, must match.</span></span>

<span data-ttu-id="78716-137">最初のページ (または他のリソース) が読み込まれると、ユーザーはリンクまたは他の UI を使用して HTTPS を使用する任意の Web サイト (または他のリソース) に移動できます。</span><span class="sxs-lookup"><span data-stu-id="78716-137">After the first page (or other resource) is loaded, a user can use links or other UI to navigate to any website (or other resource) that uses HTTPS.</span></span> <span data-ttu-id="78716-138">また、すぐに別のサイトにリダイレクトするように最初のページを設計することもできます。</span><span class="sxs-lookup"><span data-stu-id="78716-138">You can also design the first page to immediately redirect to another site.</span></span>

<span data-ttu-id="78716-139">既定では、ダイアログ ボックスのサイズはデバイス画面の高さと幅の 80% ですが、次の例に示すように、メソッドに構成オブジェクトを渡すことによってさまざまな割合を設定できます。</span><span class="sxs-lookup"><span data-stu-id="78716-139">By default, the dialog box will occupy 80% of the height and width of the device screen, but you can set different percentages by passing a configuration object to the method, as shown in the following example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

<span data-ttu-id="78716-140">これを実行するサンプル アドインについては、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="78716-140">For a sample add-in that does this, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="78716-p112">全画面表示で効率的に操作するには、両方の値を 100% に設定します。(最大有効値は 99.5% であり、最大有効値にしても、ウィンドウは移動とサイズ変更が可能です。)</span><span class="sxs-lookup"><span data-stu-id="78716-p112">Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)</span></span>

> [!NOTE]
> <span data-ttu-id="78716-p113">ホスト ウィンドウから開くことができるのは、1 つのダイアログ ボックスのみです。別のダイアログ ボックスを開こうとすると、エラーが発生します。たとえば、ユーザーが作業ウィンドウからダイアログ ボックスを開いた場合には、作業ウィンドウの別のページから 2 番目のダイアログ ボックスを開くことができません。ただし、[アドイン コマンド](../design/add-in-commands.md)からダイアログ ボックスを開く場合は、選択するたびにコマンドによって新しい (ただし非表示の) HTML ファイルが開かれます。これにより、新しい (非表示) ホスト ウィンドウが作成されるため、これらの各ウィンドウは独自のダイアログ ボックスを起動できます。詳細については、「[displayDialogAsync のエラー](dialog-handle-errors-events.md#errors-from-displaydialogasync)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="78716-p113">You can open only one dialog box from a host window. An attempt to open another dialog box generates an error. For example, if a user opens a dialog box from a task pane, she cannot open a second dialog box, from a different page in the task pane. However, when a dialog box is opened from an [add-in command](../design/add-in-commands.md), the command opens a new (but unseen) HTML file each time it is selected. This creates a new (unseen) host window, so each such window can launch its own dialog box. For more information, see [Errors from displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).</span></span>

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a><span data-ttu-id="78716-149">Office on the web のパフォーマンス オプションを利用する</span><span class="sxs-lookup"><span data-stu-id="78716-149">Take advantage of a performance option in Office on the web</span></span>

<span data-ttu-id="78716-150">`displayInIframe` プロパティは、`displayDialogAsync` に渡すことのできる構成オブジェクトの追加のプロパティです。</span><span class="sxs-lookup"><span data-stu-id="78716-150">The `displayInIframe` property is an additional property in the configuration object that you can pass to `displayDialogAsync`.</span></span> <span data-ttu-id="78716-151">このプロパティを `true` に設定し、Office on the web で開いたドキュメントでアドインを実行している場合、ダイアログ ボックスは浮動の iframe で開き、独立したウィンドウでは開きません (この方が速く開きます)。</span><span class="sxs-lookup"><span data-stu-id="78716-151">When this property is set to `true`, and the add-in is running in a document opened in Office on the web, the dialog box will open as a floating iframe rather than an independent window, which makes it open faster.</span></span> <span data-ttu-id="78716-152">例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="78716-152">The following is an example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

<span data-ttu-id="78716-153">既定値は `false` です。これはプロパティを完全に省略した場合と同じ状態です。</span><span class="sxs-lookup"><span data-stu-id="78716-153">The default value is `false`, which is the same as omitting the property entirely.</span></span> <span data-ttu-id="78716-154">アドインが Office on the web で実行されていない場合、`displayInIframe` は無視されます。</span><span class="sxs-lookup"><span data-stu-id="78716-154">If the add-in is not running in Office on the web, the `displayInIframe` is ignored.</span></span>

> [!NOTE]
> <span data-ttu-id="78716-155">どの時点であれ、iframe で開けないページにダイアログ ボックスがリダイレクトされることになる場合は、`displayInIframe: true` を使用すべきでは**ありません**。</span><span class="sxs-lookup"><span data-stu-id="78716-155">You should **not** use `displayInIframe: true` if the dialog box will at any point redirect to a page that cannot be opened in an iframe.</span></span> <span data-ttu-id="78716-156">たとえば、Google や Microsoft アカウントなど、多くの一般的な web サービスのサインインページを iframe で開くことはできません。</span><span class="sxs-lookup"><span data-stu-id="78716-156">For example, the sign in pages of many popular web services, such as Google and Microsoft account, cannot be opened in an iframe.</span></span>

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a><span data-ttu-id="78716-157">ダイアログ ボックスからホスト ページに情報を送信する</span><span class="sxs-lookup"><span data-stu-id="78716-157">Send information from the dialog box to the host page</span></span>

<span data-ttu-id="78716-158">ダイアログ ボックスは、以下の場合を除いて、作業ウィンドウのホスト ページと通信できません。</span><span class="sxs-lookup"><span data-stu-id="78716-158">The dialog box cannot communicate with the host page in the task pane unless:</span></span>

- <span data-ttu-id="78716-159">ダイアログ ボックスの現在のページがホスト ページと同じドメインにある。</span><span class="sxs-lookup"><span data-stu-id="78716-159">The current page in the dialog box is in the same domain as the host page.</span></span>
- <span data-ttu-id="78716-p117">Office JavaScript API ライブラリがページにロードされます。(Office JavaScript API ライブラリを使用するページと同様に、ページのスクリプトはメソッドをプロパティに割り当てる必要がありますが、空のメソッドにする `Office.initialize` こともできます。詳細については、「 [Office アドインを初期化する](initialize-add-in.md)」を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="78716-p117">The Office JavaScript API library is loaded in the page. (Like any page that uses the Office JavaScript API library, script for the page must assign a method to the `Office.initialize` property, although it can be an empty method. For details, see [Initialize your Office Add-in](initialize-add-in.md).)</span></span>

<span data-ttu-id="78716-163">ダイアログ ボックスのコードは、[messageParent](/javascript/api/office/office.ui#messageparent-message-) 関数を使用して、ブール値または文字列メッセージのいずれかをホスト ページに送信します。</span><span class="sxs-lookup"><span data-stu-id="78716-163">Code in the dialog box uses the [messageParent](/javascript/api/office/office.ui#messageparent-message-) function to send either a Boolean value or a string message to the host page.</span></span> <span data-ttu-id="78716-164">文字列には、単語、文、XML BLOB、文字列に変換された JSON、文字列にシリアル化できるすべてのものを指定できます。</span><span class="sxs-lookup"><span data-stu-id="78716-164">The string can be a word, sentence, XML blob, stringified JSON, or anything else that can be serialized to a string.</span></span> <span data-ttu-id="78716-165">例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="78716-165">The following is an example:</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!IMPORTANT]
> - <span data-ttu-id="78716-166">`messageParent` 関数を呼び出せるのは、ホスト ページと同じドメイン (プロトコルとポートを含む) を持つページ上のみです。</span><span class="sxs-lookup"><span data-stu-id="78716-166">The `messageParent` function can only be called on a page with the same domain (including protocol and port) as the host page.</span></span>
> - <span data-ttu-id="78716-167">この `messageParent` 関数は、ダイアログ*only*ボックスで呼び出すことができる2つの Office JS api のうちの1つです。</span><span class="sxs-lookup"><span data-stu-id="78716-167">The `messageParent` function is one of *only* two Office JS APIs that can be called in the dialog box.</span></span> 
> - <span data-ttu-id="78716-168">ダイアログボックスで呼び出すことができるその他の JS API は、 `Office.context.requirements.isSetSupported` です。</span><span class="sxs-lookup"><span data-stu-id="78716-168">The other JS API that can be called in the dialog box is `Office.context.requirements.isSetSupported`.</span></span> <span data-ttu-id="78716-169">詳細については、「 [Office アプリケーションと API 要件を指定する](specify-office-hosts-and-api-requirements.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="78716-169">For information about it, see [Specify Office applications and API requirements](specify-office-hosts-and-api-requirements.md).</span></span> <span data-ttu-id="78716-170">ただし、ダイアログボックスでは、この API は Outlook 2016 1 での購入時 (つまり、MSI バージョン) ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="78716-170">However, in the dialog box, this API isn't supported in Outlook 2016 one-time purchase (that is, the MSI version).</span></span>


<span data-ttu-id="78716-171">次の例では、`googleProfile` は文字列に変換されたバージョンのユーザーの Google プロファイルです。</span><span class="sxs-lookup"><span data-stu-id="78716-171">In the next example, `googleProfile` is a stringified version of the user's Google profile.</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

<span data-ttu-id="78716-p120">ホスト ページは、メッセージを受信するように構成する必要があります。これを構成するには、`displayDialogAsync` の元の呼び出しにコールバック パラメーターを追加します。コールバックはハンドラーを `DialogMessageReceived` イベントに割り当てます。次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="78716-p120">The host page must be configured to receive the message. You do this by adding a callback parameter to the original call of `displayDialogAsync`. The callback assigns a handler to the `DialogMessageReceived` event. The following is an example:</span></span>

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
> - <span data-ttu-id="78716-176">Office は [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトをコールバックに渡します。</span><span class="sxs-lookup"><span data-stu-id="78716-176">Office passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to the callback.</span></span> <span data-ttu-id="78716-177">Office はダイアログ ボックスを開こうとした結果を表します。</span><span class="sxs-lookup"><span data-stu-id="78716-177">It represents the result of the attempt to open the dialog box.</span></span> <span data-ttu-id="78716-178">ただし、ダイアログ ボックスでのイベントの結果は表しません。</span><span class="sxs-lookup"><span data-stu-id="78716-178">It does not represent the outcome of any events in the dialog box.</span></span> <span data-ttu-id="78716-179">この違いの詳細については、「[エラーとイベントの処理](dialog-handle-errors-events.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="78716-179">For more on this distinction, see [Handle errors and events](dialog-handle-errors-events.md).</span></span>
> - <span data-ttu-id="78716-180">`asyncResult` の `value` プロパティは [Dialog](/javascript/api/office/office.dialog) オブジェクトに設置されます。このオブジェクトはダイアログ ボックスの実行コンテキストではなく、ホスト ページに存在します。</span><span class="sxs-lookup"><span data-stu-id="78716-180">The `value` property of the `asyncResult` is set to a [Dialog](/javascript/api/office/office.dialog) object, which exists in the host page, not in the dialog box's execution context.</span></span>
> - <span data-ttu-id="78716-p122">`processMessage` はイベントを処理する関数です。任意の名前を指定できます。</span><span class="sxs-lookup"><span data-stu-id="78716-p122">The `processMessage` is the function that handles the event. You can give it any name you want.</span></span>
> - <span data-ttu-id="78716-183">`dialog` 変数は、`processMessage` でも参照されるため、コールバックよりも広い範囲で宣言されます。</span><span class="sxs-lookup"><span data-stu-id="78716-183">The `dialog` variable is declared at a wider scope than the callback because it is also referenced in `processMessage`.</span></span>

<span data-ttu-id="78716-184">`DialogMessageReceived` イベントのハンドラーの簡単な例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="78716-184">The following is a simple example of a handler for the `DialogMessageReceived` event:</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - <span data-ttu-id="78716-185">Office は `arg` オブジェクトをハンドラーに渡します。</span><span class="sxs-lookup"><span data-stu-id="78716-185">Office passes the `arg` object to the handler.</span></span> <span data-ttu-id="78716-186">その `message` プロパティは、ダイアログ ボックスの `messageParent` の呼び出しで送信されるブール値または文字列です。</span><span class="sxs-lookup"><span data-stu-id="78716-186">Its `message` property is the Boolean or string sent by the call of `messageParent` in the dialog box.</span></span> <span data-ttu-id="78716-187">この例では、Microsoft アカウントや Google などのサービスからのユーザーのプロファイルを文字列で表現しています。これは、を使用してオブジェクトに逆シリアル化され `JSON.parse` ます。</span><span class="sxs-lookup"><span data-stu-id="78716-187">In this example, it is a stringified representation of a user's profile from a service such as Microsoft account or Google, so it is deserialized back to an object with `JSON.parse`.</span></span>
> - <span data-ttu-id="78716-p124">`showUserName` 実装は表示されません。作業ウィンドウ上に個人用のウェルカム メッセージが表示される場合があります。</span><span class="sxs-lookup"><span data-stu-id="78716-p124">The `showUserName` implementation is not shown. It might display a personalized welcome message on the task pane.</span></span>

<span data-ttu-id="78716-190">ダイアログ ボックスのユーザー操作が完了すると、次の例に示すようにメッセージ ハンドラーはダイアログ ボックスを閉じます。</span><span class="sxs-lookup"><span data-stu-id="78716-190">When the user interaction with the dialog box is completed, your message handler should close the dialog box, as shown in this example.</span></span>

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
> - <span data-ttu-id="78716-191">`dialog` オブジェクトは `displayDialogAsync` の呼び出しによって返されるものと同じである必要があります。</span><span class="sxs-lookup"><span data-stu-id="78716-191">The `dialog` object must be the same one that is returned by the call of `displayDialogAsync`.</span></span>
> - <span data-ttu-id="78716-192">`dialog.close` の呼び出しは、直ちにダイアログ ボックスを閉じるよう Office に指示します。</span><span class="sxs-lookup"><span data-stu-id="78716-192">The call of `dialog.close` tells Office to immediately close the dialog box.</span></span>

<span data-ttu-id="78716-193">これらの手法を使用するサンプル アドインについては、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="78716-193">For a sample add-in that uses these techniques, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="78716-p125">メッセージを受信した後、アドインで作業ウィンドウの別のページを開く必要がある場合は、ハンドラーの最後の行として `window.location.replace` メソッド (または `window.location.href`) を使用できます。次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="78716-p125">If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example:</span></span>

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

<span data-ttu-id="78716-196">これを実行するアドインの例については、「[Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)」 (PowerPoint アドインで Microsoft Graph を使用した Excel グラフの挿入) のサンプルを参照してください。</span><span class="sxs-lookup"><span data-stu-id="78716-196">For an example of an add-in that does this, see the [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) sample.</span></span>

### <a name="conditional-messaging"></a><span data-ttu-id="78716-197">条件付きのメッセージング</span><span class="sxs-lookup"><span data-stu-id="78716-197">Conditional messaging</span></span>

<span data-ttu-id="78716-198">ダイアログ ボックスから複数の `messageParent` 呼び出しを送信できますが、`DialogMessageReceived` イベントのホスト ページにあるハンドラーは 1 つのみのため、ハンドラーは条件ロジックを使用してさまざまなメッセージを区別する必要があります。</span><span class="sxs-lookup"><span data-stu-id="78716-198">Because you can send multiple `messageParent` calls from the dialog box, but you have only one handler in the host page for the `DialogMessageReceived` event, the handler must use conditional logic to distinguish different messages.</span></span> <span data-ttu-id="78716-199">たとえば、ユーザーが Microsoft アカウントや Google などの id プロバイダーにサインインするように求めるダイアログボックスが表示された場合、ユーザーのプロファイルがメッセージとして送信されます。</span><span class="sxs-lookup"><span data-stu-id="78716-199">For example, if the dialog box prompts a user to sign in to an identity provider such as Microsoft account or Google, it sends the user's profile as a message.</span></span> <span data-ttu-id="78716-200">認証が失敗した場合、次の例のように、ダイアログ ボックスはホスト ページにエラー情報を送信します。</span><span class="sxs-lookup"><span data-stu-id="78716-200">If authentication fails, the dialog box sends error information to the host page, as in the following example:</span></span>

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
> - <span data-ttu-id="78716-201">`loginSuccess` 変数は、ID プロバイダーからの HTTP 応答を読み取ることによって初期化されます。</span><span class="sxs-lookup"><span data-stu-id="78716-201">The `loginSuccess` variable would be initialized by reading the HTTP response from the identity provider.</span></span>
> - <span data-ttu-id="78716-p127">`getProfile` 関数と `getError` 関数の実装は表示されません。両方の関数はそれぞれ、クエリ パラメーターまたは HTTP 応答の本文からデータを取得します。</span><span class="sxs-lookup"><span data-stu-id="78716-p127">The the implementation of the `getProfile` and `getError` functions are not not shown. They each get data from a query parameter or from the body of the HTTP response.</span></span>
> - <span data-ttu-id="78716-p128">サインインが成功したかどうかに応じて、さまざまな種類の匿名のオブジェクトが送信されます。両方の関数に `messageType` プロパティがありますが、一方には `profile` プロパティ、もう一方には `error` プロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="78716-p128">Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.</span></span>

<span data-ttu-id="78716-p129">次の例に示すように、ホスト ページのハンドラー コードは分岐に `messageType` プロパティの値を使用します。`showUserName` 関数は上記の例と同じであり、`showNotification` 関数はホスト ページの UI にエラーを表示することに注意してください。</span><span class="sxs-lookup"><span data-stu-id="78716-p129">The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.</span></span>

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
> <span data-ttu-id="78716-208">`showNotification`の実装は、この記事のサンプル コードでは表示されません。</span><span class="sxs-lookup"><span data-stu-id="78716-208">The `showNotification` implementation is not shown in the sample code provided by this article.</span></span> <span data-ttu-id="78716-209">アドインでこの関数を実装する方法の例は、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="78716-209">For an example of how you might implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

## <a name="pass-information-to-the-dialog-box"></a><span data-ttu-id="78716-210">情報をダイアログ ボックスに渡す</span><span class="sxs-lookup"><span data-stu-id="78716-210">Pass information to the dialog box</span></span>

<span data-ttu-id="78716-211">アドインは、 [messageChild](/javascript/api/office/office.dialog#messagechild-message-)を使用して、[ホストページ](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)からダイアログボックスにメッセージを送信できます。</span><span class="sxs-lookup"><span data-stu-id="78716-211">Your add-in can send messages from the [host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) to a dialog box using [Dialog.messageChild](/javascript/api/office/office.dialog#messagechild-message-).</span></span>

### <a name="use-messagechild-from-the-host-page"></a><span data-ttu-id="78716-212">`messageChild()`ホストページからの使用</span><span class="sxs-lookup"><span data-stu-id="78716-212">Use `messageChild()` from the host page</span></span>

<span data-ttu-id="78716-213">Office ダイアログ API を呼び出してダイアログボックスを開くと、 [dialog](/javascript/api/office/office.dialog) オブジェクトが返されます。</span><span class="sxs-lookup"><span data-stu-id="78716-213">When you call the Office dialog API to open a dialog box, a [Dialog](/javascript/api/office/office.dialog) object is returned.</span></span> <span data-ttu-id="78716-214">オブジェクトは他のメソッドによって参照されるため、 [Displaydialogasync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) メソッドよりもスコープが大きい変数に割り当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="78716-214">It should be assigned to a variable that has greater scope than the [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) method because the object will be referenced by other methods.</span></span> <span data-ttu-id="78716-215">例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="78716-215">The following is an example:</span></span>

```javascript
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);

function processMessage(arg) {
    dialog.close();

  // message processing code goes here;

}
```

<span data-ttu-id="78716-216">この `Dialog` オブジェクトには、文字列データを含む任意の文字列をダイアログボックスに送信する [messageChild](/javascript/api/office/office.dialog#messagechild-message-) メソッドがあります。</span><span class="sxs-lookup"><span data-stu-id="78716-216">This `Dialog` object has a [messageChild](/javascript/api/office/office.dialog#messagechild-message-) method that sends any string, including stringified data, to the dialog box.</span></span> <span data-ttu-id="78716-217">これにより `DialogParentMessageReceived` 、ダイアログボックスでイベントが発生します。</span><span class="sxs-lookup"><span data-stu-id="78716-217">This raises a `DialogParentMessageReceived` event in the dialog box.</span></span> <span data-ttu-id="78716-218">コードでは、次のセクションに示すように、このイベントを処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="78716-218">Your code should handle this event, as shown in the next section.</span></span>

<span data-ttu-id="78716-219">ダイアログの UI が現在アクティブなワークシートに関連付けられていて、他のワークシートを基準としたワークシートの位置を示すシナリオを考えてみます。</span><span class="sxs-lookup"><span data-stu-id="78716-219">Consider a scenario in which the UI of the dialog is related to the currently active worksheet and that worksheet's position relative to the other worksheets.</span></span> <span data-ttu-id="78716-220">次の例では、 `sheetPropertiesChanged` Excel ワークシートのプロパティをダイアログボックスに送信します。</span><span class="sxs-lookup"><span data-stu-id="78716-220">In the following example, `sheetPropertiesChanged` sends Excel worksheet properties to the dialog box.</span></span> <span data-ttu-id="78716-221">この例では、現在のワークシートに "My Sheet" という名前が付けられ、ブックの2番目のシートになります。</span><span class="sxs-lookup"><span data-stu-id="78716-221">In this case, the current worksheet is named "My Sheet" and it's the second sheet in the workbook.</span></span> <span data-ttu-id="78716-222">データは、オブジェクトと文字列にカプセル化され、に渡すことができ `messageChild` ます。</span><span class="sxs-lookup"><span data-stu-id="78716-222">The data is encapsulated in an object and stringified so that it can be passed to `messageChild`.</span></span>

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

### <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a><span data-ttu-id="78716-223">ダイアログボックスで DialogParentMessageReceived を処理する</span><span class="sxs-lookup"><span data-stu-id="78716-223">Handle DialogParentMessageReceived in the dialog box</span></span>

<span data-ttu-id="78716-224">ダイアログボックスの JavaScript で、イベントのハンドラーを `DialogParentMessageReceived` [UI. Addhandler async](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) メソッドに登録します。</span><span class="sxs-lookup"><span data-stu-id="78716-224">In the dialog box's JavaScript, register a handler for the `DialogParentMessageReceived` event with the [UI.addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) method.</span></span> <span data-ttu-id="78716-225">これは通常、次のように、 [tialize メソッドまたは Office.iniメソッド](initialize-add-in.md)で実行されます。</span><span class="sxs-lookup"><span data-stu-id="78716-225">This is typically done in the [Office.onReady or Office.initialize methods](initialize-add-in.md), as shown in the following.</span></span> <span data-ttu-id="78716-226">(より堅牢な例は次のとおりです。)</span><span class="sxs-lookup"><span data-stu-id="78716-226">(A more robust example is below.)</span></span>

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

<span data-ttu-id="78716-227">その後、ハンドラーを定義し `onMessageFromParent` ます。</span><span class="sxs-lookup"><span data-stu-id="78716-227">Then, define the `onMessageFromParent` handler.</span></span> <span data-ttu-id="78716-228">次のコードでは、前のセクションの例を続行します。</span><span class="sxs-lookup"><span data-stu-id="78716-228">The following code continues the example from the preceding section.</span></span> <span data-ttu-id="78716-229">Office によってハンドラーに引数が渡さ `message` れ、引数オブジェクトのプロパティにホストページの文字列が含まれていることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="78716-229">Note that Office passes an argument to the handler and that the `message` property of the argument object contains the string from the host page.</span></span> <span data-ttu-id="78716-230">この例では、メッセージはオブジェクトに再変換、jQuery を使用して、新しいワークシート名に一致するダイアログのトップの見出しを設定しています。</span><span class="sxs-lookup"><span data-stu-id="78716-230">In this example, the message is reconverted to an object and jQuery is used to set the top heading of the dialog to match the new worksheet name.</span></span>

```javascript
function onMessageFromParent(event) {
    var messageFromParent = JSON.parse(event.message);
    $('h1').text(messageFromParent.name);
}
```

<span data-ttu-id="78716-231">ハンドラーが適切に登録されていることを確認することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="78716-231">It is a best practice to verify that your handler is properly registered.</span></span> <span data-ttu-id="78716-232">これを行うには、コールバックをメソッドに渡し `addHandlerAsync` ます。</span><span class="sxs-lookup"><span data-stu-id="78716-232">You can do this by passing a callback to the `addHandlerAsync` method.</span></span> <span data-ttu-id="78716-233">これは、ハンドラーの登録が完了したときに実行されます。</span><span class="sxs-lookup"><span data-stu-id="78716-233">This runs when the attempt to register the handler completes.</span></span> <span data-ttu-id="78716-234">ハンドラーが正常に登録されなかった場合は、ハンドラーを使用して、エラーを記録または表示します。</span><span class="sxs-lookup"><span data-stu-id="78716-234">Use the handler to log or show an error if the handler was not successfully registered.</span></span> <span data-ttu-id="78716-235">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="78716-235">The following is an example.</span></span> <span data-ttu-id="78716-236">ここで `reportError` は、エラーを記録または表示する関数であることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="78716-236">Note that `reportError` is a function, not defined here, that logs or displays the error.</span></span>

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent,
            onRegisterMessageComplete);
    });

function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reportError(asyncResult.error.message);
    }
}
```

### <a name="conditional-messaging-from-parent-page-to-dialog-box"></a><span data-ttu-id="78716-237">親ページからダイアログボックスへの条件付きメッセージング</span><span class="sxs-lookup"><span data-stu-id="78716-237">Conditional messaging from parent page to dialog box</span></span>

<span data-ttu-id="78716-238">ホストページから複数の呼び出しを行うことはできます `messageChild` が、イベントのダイアログボックスにはハンドラーが1つしかないため、 `DialogParentMessageReceived` ハンドラーは異なるメッセージを区別するために条件付きロジックを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="78716-238">Because you can make multiple `messageChild` calls from the host page, but you have only one handler in the dialog box for the `DialogParentMessageReceived` event, the handler must use conditional logic to distinguish different messages.</span></span> <span data-ttu-id="78716-239">[条件付き](#conditional-messaging)メッセージの説明に従って、ダイアログボックスがホストページにメッセージを送信しているときに、条件付きメッセージを構造化する方法で、これを正確に行うことができます。</span><span class="sxs-lookup"><span data-stu-id="78716-239">You can do this in a way that is precisely parallel to how you would structure conditional messaging when the dialog box is sending a message to the host page as described in [Conditional messaging](#conditional-messaging).</span></span>

> [!NOTE]
> <span data-ttu-id="78716-240">状況によって1.2 は、この api は、表示され `messageChild` ている [必要](../reference/requirement-sets/dialog-api-requirement-sets.md)があります。</span><span class="sxs-lookup"><span data-stu-id="78716-240">In some situations, the `messageChild` API, which is a part of the [DialogApi 1.2 requirement set](../reference/requirement-sets/dialog-api-requirement-sets.md),  may not be supported.</span></span> <span data-ttu-id="78716-241">別の方法として、ダイアログ [ボックスへのメッセージをホストページからダイアログボックスに渡す](parent-to-dialog.md)方法もあります。</span><span class="sxs-lookup"><span data-stu-id="78716-241">Some alternative ways for parent-to-dialog-box messaging are described in [Alternative ways of passing messages to a dialog box from its host page](parent-to-dialog.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="78716-242">設定されている場合は、アドインマニフェストのセクションでは、[参照] [api の1.2 要件セット](../reference/requirement-sets/dialog-api-requirement-sets.md) を指定できません `<Requirements>` 。</span><span class="sxs-lookup"><span data-stu-id="78716-242">The [DialogApi 1.2 requirement set](../reference/requirement-sets/dialog-api-requirement-sets.md) cannot be specified in the `<Requirements>` section of an add-in manifest.</span></span> <span data-ttu-id="78716-243">[Issetsupported](specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)メソッドを使用して、実行時に、操作1.2 のサポートをチェックする必要があります。</span><span class="sxs-lookup"><span data-stu-id="78716-243">You will have to check for support for DialogApi 1.2 at runtime using the [isSetSupported](specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code) method.</span></span> <span data-ttu-id="78716-244">マニフェスト要件のサポートは開発中です。</span><span class="sxs-lookup"><span data-stu-id="78716-244">Support for manifest requirements is under development.</span></span>

## <a name="closing-the-dialog-box"></a><span data-ttu-id="78716-245">ダイアログ ボックスを閉じる</span><span class="sxs-lookup"><span data-stu-id="78716-245">Closing the dialog box</span></span>

<span data-ttu-id="78716-p140">ダイアログ ボックスを閉じるボタンをダイアログ ボックス内に実装できます。これを実行するには、ボタンのクリック イベント ハンドラーは `messageParent` を使用して、ボタンがクリックされたことをホスト ページに通知する必要があります。次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="78716-p140">You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example:</span></span>

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

<span data-ttu-id="78716-249">`DialogMessageReceived` のホスト ページ ハンドラーは、この例のように `dialog.close` を呼び出します </span><span class="sxs-lookup"><span data-stu-id="78716-249">The host page handler for `DialogMessageReceived` would call `dialog.close`, as in this example.</span></span> <span data-ttu-id="78716-250">(`dialog` オブジェクトを初期化する方法を示す前述の例を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="78716-250">(See previous examples that show how the `dialog` object is initialized.)</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

<span data-ttu-id="78716-251">独自の終了ダイアログ UI がない場合でも、エンド ユーザーは右上隅にある **X** を選択してダイアログ ボックスを閉じることができます。</span><span class="sxs-lookup"><span data-stu-id="78716-251">Even when you don't have your own close-dialog UI, an end user can close the dialog box by choosing the **X** in the upper-right corner.</span></span> <span data-ttu-id="78716-252">この操作により `DialogEventReceived` イベントがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="78716-252">This action triggers the `DialogEventReceived` event.</span></span> <span data-ttu-id="78716-253">イベントがトリガーされたときに、ホスト ウィンドウに通知する必要がある場合、ホスト ウィンドウはこのイベントのハンドラーを宣言する必要があります。</span><span class="sxs-lookup"><span data-stu-id="78716-253">If your host pane needs to know when this happens, it should declare a handler for this event.</span></span> <span data-ttu-id="78716-254">詳細については、「[ダイアログ ボックスでのエラーとイベント](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box)」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="78716-254">See the section [Errors and events in the dialog box](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box) for details.</span></span>

## <a name="advanced-topics-and-special-scenarios"></a><span data-ttu-id="78716-255">高度なトピックと特殊なシナリオ</span><span class="sxs-lookup"><span data-stu-id="78716-255">Advanced topics and special scenarios</span></span>

### <a name="use-the-dialog-api-to-show-a-video"></a><span data-ttu-id="78716-256">ダイアログ API を使用してビデオを表示する</span><span class="sxs-lookup"><span data-stu-id="78716-256">Use the Dialog API to show a video</span></span>

<span data-ttu-id="78716-257">「[Office ダイアログ ボックスを使用してビデオを表示する](dialog-video.md)」を参照してください 。</span><span class="sxs-lookup"><span data-stu-id="78716-257">See [Use the Office dialog box to show a video](dialog-video.md).</span></span>

### <a name="use-the-dialog-apis-in-an-authentication-flow"></a><span data-ttu-id="78716-258">認証フローでダイアログ API を使用する</span><span class="sxs-lookup"><span data-stu-id="78716-258">Use the Dialog APIs in an authentication flow</span></span>

<span data-ttu-id="78716-259">「[Office Dialog API を使用して認証する](auth-with-office-dialog-api.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="78716-259">See [Authenticate with the Office dialog API](auth-with-office-dialog-api.md).</span></span>

### <a name="using-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a><span data-ttu-id="78716-260">単一ページ アプリケーションとクライアント側ルーティングで Office ダイアログ API を使用する</span><span class="sxs-lookup"><span data-stu-id="78716-260">Using the Office dialog API with single-page applications and client-side routing</span></span>

<span data-ttu-id="78716-261">Office ダイアログ API を使用する場合は、SPA およびクライアント側のルーティングを慎重に行う必要があります。</span><span class="sxs-lookup"><span data-stu-id="78716-261">SPAs and client-side routing need to be handled with care when you are using the Office dialog API.</span></span> <span data-ttu-id="78716-262">「[SPA で Office ダイアログ API を使用する場合のベスト プラクティス](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="78716-262">Please see [Best practices for using the Office dialog API in an SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).</span></span>

### <a name="error-and-event-handling"></a><span data-ttu-id="78716-263">エラーとイベントの処理</span><span class="sxs-lookup"><span data-stu-id="78716-263">Error and event handling</span></span>

<span data-ttu-id="78716-264">詳細については、「[Office ダイアログ ボックスでのエラーとイベントの処理](dialog-handle-errors-events.md)」を参照し ます。</span><span class="sxs-lookup"><span data-stu-id="78716-264">See [Handling errors and events in the Office dialog box](dialog-handle-errors-events.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="78716-265">次の手順</span><span class="sxs-lookup"><span data-stu-id="78716-265">Next steps</span></span>

<span data-ttu-id="78716-266">Office ダイアログ API に関するヒントとヘスと プラクティスの詳細については、「[Office ダイアログ API のベスト プラクティスとルール](dialog-best-practices.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="78716-266">Learn about gotchas and best practices for the Office dialog API in [Best practices and rules for the Office dialog API](dialog-best-practices.md).</span></span>

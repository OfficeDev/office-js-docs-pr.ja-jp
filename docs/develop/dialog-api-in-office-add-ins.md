---
title: Office アドインでダイアログ API を使用する
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 7d92609cbb4542d56a5bc5a28310b1d39d9c634b
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506239"
---
# <a name="use-the-dialog-api-in-your-office-add-ins"></a><span data-ttu-id="3152c-102">Office アドインでダイアログ API を使用する</span><span class="sxs-lookup"><span data-stu-id="3152c-102">Use the Dialog API in your Office Add-ins</span></span>

<span data-ttu-id="3152c-p101">[ダイアログ API](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) を使用して、Office アドインでダイアログ ボックスを開くことができます。この記事では、Office アドインでダイアログ API を使用するためのガイダンスを提供します。</span><span class="sxs-lookup"><span data-stu-id="3152c-p101">You can use the [Dialog API](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) to open dialog boxes in your Office Add-in. This article provides guidance for using the Dialog API in your Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="3152c-p102">ダイアログ API の現在のサポート状況の詳細については、「[ダイアログ API の要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets?view=office-js)」を参照してください。現在、ダイアログ API は Word、Excel、PowerPoint、Outlook でサポートされています。</span><span class="sxs-lookup"><span data-stu-id="3152c-p102">For information about where the Dialog API is currently supported, see [Dialog API requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets?view=office-js). The Dialog API is currently supported for Word, Excel, PowerPoint, and Outlook.</span></span>

> <span data-ttu-id="3152c-p103">ダイアログ API の主要なシナリオは、Google や Facebook などのリソースを使用して認証を有効にすることです。アドインで Office ユーザーのデータまたは Office 365 や OneDrive など Microsoft Graph を使用してアクセスできるリソースのデータが必要な場合は、可能な限り、シングル サインオン API を使用することをお勧めします。シングル サインオンの API を使用する場合、ダイアログ API は必要はありません。詳細については、「[Office アドインのシングル サインオンを有効化する](sso-in-office-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3152c-p103">A primary scenario for the Dialog APIs is to enable authentication with a resource such as Google or Facebook. If your add-in requires data about the Office user or their resources accessible through Microsoft Graph, such as Office 365 or OneDrive, we recommend that you use the single sign-on API whenever you can. If you use the APIs for single sign-on, then you will not need the Dialog API. For details, see [Enable single sign-on for Office Add-ins](sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="3152c-111">作業ウィンドウ、コンテンツ アドイン、または[アドイン コマンド](../design/add-in-commands.md)からダイアログ ボックスを開いて、次の操作を実行することを検討してください。</span><span class="sxs-lookup"><span data-stu-id="3152c-111">Consider opening a dialog box from a task pane or content add-in or [add-in command](../design/add-in-commands.md) to do the following:</span></span>

- <span data-ttu-id="3152c-112">作業ウィンドウで直接開くことができないサインイン ページを表示する。</span><span class="sxs-lookup"><span data-stu-id="3152c-112">Display sign in pages that cannot be opened directly in a task pane.</span></span>
- <span data-ttu-id="3152c-113">アドインでの作業用に画面領域を広げるか全画面表示にする。</span><span class="sxs-lookup"><span data-stu-id="3152c-113">Provide more screen space, or even a full screen, for some tasks in your add-in.</span></span>
- <span data-ttu-id="3152c-114">再生が作業ウィンドウに限定されている場合に、小さすぎるビデオをホストする。</span><span class="sxs-lookup"><span data-stu-id="3152c-114">Host a video that would be too small if confined to a task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="3152c-p104">UI 要素を重ねて表示することはお勧めできないため、シナリオで必要な場合を除き、作業ウィンドウでダイアログを開かないようにします。作業ウィンドウの表示領域の使用方法を検討するときには、作業ウィンドウはタブ表示できることに注意してください。例については、「[Excel アドイン JavaScriptSalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)」のサンプルを参照してください。</span><span class="sxs-lookup"><span data-stu-id="3152c-p104">Because overlapping UI elements are discouraged, avoid opening a dialog from a task pane unless your scenario requires it. When you consider how to use the surface area of a task pane, note that task panes can be tabbed. For an example, see the [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) sample.</span></span>

<span data-ttu-id="3152c-118">次の画像は、ダイアログ ボックスの例を示します。</span><span class="sxs-lookup"><span data-stu-id="3152c-118">The following image shows an example of a dialog box.</span></span>

![アドイン コマンド](../images/auth-o-dialog-open.png)

<span data-ttu-id="3152c-p105">ダイアログ ボックスが常に画面の中央に開くことに注意してください。ユーザーはダイアログ ボックスの移動とサイズ変更ができます。ウィンドウは*モードレス*です。Office のホスト アプリケーションのドキュメントの操作と、作業ウィンドウのホスト ページ (存在する場合) の操作の両方を続行できます。</span><span class="sxs-lookup"><span data-stu-id="3152c-p105">Note that the dialog box always opens in the center of the screen. The user can move and resize it. The window is *nonmodal*--a user can continue to interact with both the document in the host Office application and with the host page in the task pane, if there is one.</span></span>

## <a name="dialog-api-scenarios"></a><span data-ttu-id="3152c-123">ダイアログ API のシナリオ</span><span class="sxs-lookup"><span data-stu-id="3152c-123">Dialog API scenarios</span></span>

<span data-ttu-id="3152c-124">Office JavaScript API は、[Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js) オブジェクトと [Office.context.ui 名前空間](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js)の 2 つの関数を使用する以下のシナリオをサポートしています。</span><span class="sxs-lookup"><span data-stu-id="3152c-124">The Office JavaScript APIs support the following scenarios with a [Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js) object and two functions in the [Office.context.ui namespace](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js).</span></span>

### <a name="open-a-dialog-box"></a><span data-ttu-id="3152c-125">ダイアログ ボックスを開く</span><span class="sxs-lookup"><span data-stu-id="3152c-125">Open a dialog box</span></span>

<span data-ttu-id="3152c-p106">ダイアログ ボックスを開くには、作業ウィンドウのコードで [displayDialogAsync](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) メソッドを呼び出して、開くリソースの URL を渡します。これは、通常はページですが、MVC アプリケーションのコントローラー メソッド、ルート、Web サービス メソッド、またはその他のリソースの場合もあります。この記事では、'ページ' または 'Web サイト' とは、ダイアログ内のリソースを意味します。以下のコードは簡単な例を示しています。</span><span class="sxs-lookup"><span data-stu-id="3152c-p106">To open a dialog box, your code in the task pane calls the [displayDialogAsync](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) method and passes to it the URL of the resource that you want to open. This is usually a page, but it can be a controller method in an MVC application, a route, a web service method, or any other resource. In this article, 'page' or 'website' refers to the resource in the dialog. The following code is a simple example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - <span data-ttu-id="3152c-p107">URL には HTTP**S** プロトコルを使用します。これは、読み込まれる最初のページだけでなく、ダイアログ ボックスに読み込まれるすべてのページに対して必須です。</span><span class="sxs-lookup"><span data-stu-id="3152c-p107">The URL uses the HTTP**S** protocol. This is mandatory for all pages loaded in a dialog box, not just the first page loaded.</span></span>
> - <span data-ttu-id="3152c-p108">ダイアログ リソースのドメインは、ホスト ページのドメインと同じで、作業ウィンドウ内のページまたはアドイン コマンドの [関数のファイル](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/functionfile?view=office-js) となります。ページ、コント ローラー メソッド、または `displayDialogAsync` メソッドに渡されるその他のリソースは、ホスト ページと同じドメインである必要があり、必須です。</span><span class="sxs-lookup"><span data-stu-id="3152c-p108">The dialog resource's domain is the same as the domain of the host page, which can be the page in a task pane or the [function file](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/functionfile?view=office-js) of an add-in command. This is required: the page, controller method, or other resource that is passed to the `displayDialogAsync` method must be in the same domain as the host page.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3152c-p109">ダイアログのホスト ページとリソースには、同じフル ドメインがなくてはいけません。`displayDialogAsync` にアドインのドメインのサブドメインを渡そうとすると、正常に作動しません。サブドメインを含め、フル ドメインが一致する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3152c-p109">The host page and the resources of the dialog must have the same full domain. If you attempt to pass `displayDialogAsync` a subdomain of the add-in's domain, it will not work. The full domain, including any subdomain, must match.</span></span>

<span data-ttu-id="3152c-p110">最初のページ (または他のリソース) が読み込まれると、ユーザーは HTTPS を使用する任意の Web サイト (または他のリソース) に移動できます。また、すぐに別のサイトにリダイレクトするように最初のページを設計することもできます。</span><span class="sxs-lookup"><span data-stu-id="3152c-p110">After the first page (or other resource) is loaded, a user can go to any website (or other resource) that uses HTTPS. You can also design the first page to immediately redirect to another site.</span></span>

<span data-ttu-id="3152c-139">既定では、ダイアログ ボックスのサイズはデバイス画面の高さと幅の 80% ですが、以下の例に示すように、メソッドに構成オブジェクトを渡すことによってさまざまな割合を設定することができます。</span><span class="sxs-lookup"><span data-stu-id="3152c-139">By default, the dialog box will occupy 80% of the height and width of the device screen, but you can set different percentages by passing a configuration object to the method, as shown in the following example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

<span data-ttu-id="3152c-140">これを実行するサンプル アドインについては、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3152c-140">For a sample add-in that does this, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="3152c-p111">全画面表示で効率的に操作するには、両方の値を 100% に設定します。(最大有効値は 99.5% であり、最大有効値にしても、ウィンドウは移動とサイズ変更が可能です。)</span><span class="sxs-lookup"><span data-stu-id="3152c-p111">Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)</span></span>

> [!NOTE]
> <span data-ttu-id="3152c-p112">ホスト ウィンドウから開くことができるのは、1 つのダイアログ ボックスのみです。別のダイアログ ボックスを開こうとすると、エラーが発生します。たとえば、ユーザーが作業ウィンドウからダイアログ ボックスを開いた場合、作業ウィンドウの別のページから 2 番目のダイアログ ボックスを開くことができません。ただし、[アドイン コマンド](../design/add-in-commands.md)からダイアログ ボックスを開く場合、選択するたびにコマンドによって新しい (ただし非表示の) HTML ファイルが開かれます。これにより、新しい (非表示) ホスト ウィンドウが作成されるため、これらの各ウィンドウは独自のダイアログ ボックスを起動できます。詳細については、「[displayDialogAsync のエラー](#errors-from-displaydialogasync)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3152c-p112">You can open only one dialog box from a host window. An attempt to open another dialog box generates an error. For example, if a user opens a dialog box from a task pane, she cannot open a second dialog box, from a different page in the task pane. However, when a dialog box is opened from an [add-in command](../design/add-in-commands.md), the command opens a new (but unseen) HTML file each time it is selected. This creates a new (unseen) host window, so each such window can launch its own dialog box. For more information, see [Errors from displayDialogAsync](#errors-from-displaydialogasync).</span></span>

### <a name="take-advantage-of-a-performance-option-in-office-online"></a><span data-ttu-id="3152c-149">Office Online のパフォーマンス オプションを利用する</span><span class="sxs-lookup"><span data-stu-id="3152c-149">Take advantage of a performance option in Office Online</span></span>

<span data-ttu-id="3152c-p113">`displayInIframe` プロパティは、`displayDialogAsync` に渡すことのできる構成オブジェクトの追加のプロパティです。このプロパティを `true` に設定し、Office Online で開いたドキュメントでアドインを実行している場合、ダイアログ ボックスは独立したウィンドウではなく、浮動化した iframe として開きます。この方法によってより速く開きます。以下に例を示します。</span><span class="sxs-lookup"><span data-stu-id="3152c-p113">The `displayInIframe` property is an additional property in the configuration object that you can pass to `displayDialogAsync`. When this property is set to `true`, and the add-in is running in a document opened in Office Online, the dialog box will open as a floating iframe rather than an independent window, which makes it open faster. The following is an example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

<span data-ttu-id="3152c-p114">規定値は `false` です。これは、プロパティを完全に省略することと同じです。アドインが Office Online で実行されていない場合、`displayInIframe` は無視されます。</span><span class="sxs-lookup"><span data-stu-id="3152c-p114">The default value is `false`, which is the same as omitting the property entirely. If the add-in is not running in Office Online, the `displayInIframe` is ignored.</span></span>

> [!NOTE]
> <span data-ttu-id="3152c-p115">どの時点であっても、iframe で開けないページにダイアログがリダイレクトされることになる場合は、 を**使用**しないでください`displayInIframe: true`。たとえば、Google や Microsoft アカウントなどの多くの一般的な Web サービスのサインイン ページは、iframe で開くことができません。</span><span class="sxs-lookup"><span data-stu-id="3152c-p115">You should **not** use `displayInIframe: true` if the dialog will at any point redirect to a page that cannot be opened in an iframe. For example, the sign in pages of many popular web services, such as Google and Microsoft Account, cannot be opened in an iframe.</span></span>

### <a name="send-information-from-the-dialog-box-to-the-host-page"></a><span data-ttu-id="3152c-157">ダイアログ ボックスからホスト ページに情報を送信する</span><span class="sxs-lookup"><span data-stu-id="3152c-157">Send information from the dialog box to the host page</span></span>

<span data-ttu-id="3152c-158">ダイアログ ボックスは、以下の場合を除いて、作業ウィンドウのホスト ページと通信できません。</span><span class="sxs-lookup"><span data-stu-id="3152c-158">The dialog box cannot communicate with the host page in the task pane unless:</span></span>

- <span data-ttu-id="3152c-159">ダイアログ ボックスの現在のページがホスト ページと同じドメインにある。</span><span class="sxs-lookup"><span data-stu-id="3152c-159">The current page in the dialog box is in the same domain as the host page.</span></span>
- <span data-ttu-id="3152c-p116">Office JavaScript ライブラリがページに読み込まれている。(Office JavaScript ライブラリを使用するすべてのページと同様に、ページのスクリプトは `Office.initialize` プロパティにメソッドを割り当てる必要があり、これは空のメソッドでもかまいません。詳細については、「[アドインの初期化](understanding-the-javascript-api-for-office.md#initializing-your-add-in)」を参照してください。)</span><span class="sxs-lookup"><span data-stu-id="3152c-p116">The Office JavaScript library is loaded in the page. (Like any page that uses the Office JavaScript library, script for the page must assign a method to the `Office.initialize` property, although it can be an empty method. For details, see [Initializing your add-in](understanding-the-javascript-api-for-office.md#initializing-your-add-in).)</span></span>

<span data-ttu-id="3152c-p117">ダイアログ ページのコードは、`messageParent` 関数を使用して、ブール値または文字列メッセージのいずれかをホスト ページに送信します。文字列には、単語、文、XML BLOB、文字列に変換された JSON、または文字列にシリアル化できるすべてのものを指定できます。以下に例を示します。</span><span class="sxs-lookup"><span data-stu-id="3152c-p117">Code in the dialog page uses the `messageParent` function to send either a Boolean value or a string message to the host page. The string can be a word, sentence, XML blob, stringified JSON, or anything else that can be serialized to a string. The following is an example:</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!NOTE]
> - <span data-ttu-id="3152c-p118">`messageParent` 関数は、ダイアログ ボックスで呼び出すことができる、2 つ*だけ*の Office API のうちの 1 つです。もう 1 つは `Office.context.requirements.isSetSupported` です。詳細は、「[Office のホストと API の要件を指定する](specify-office-hosts-and-api-requirements.md)」を参照してださい。</span><span class="sxs-lookup"><span data-stu-id="3152c-p118">The `messageParent` function is one of *only* two Office APIs that can be called in the dialog box. The other is `Office.context.requirements.isSetSupported`. For information about it, see [Specify Office hosts and API requirements](specify-office-hosts-and-api-requirements.md).</span></span>
> - <span data-ttu-id="3152c-169">`messageParent` 関数を呼び出せるのは、ホスト ページと同じドメイン (プロトコルとポートを含む) を持つページ上のみです。</span><span class="sxs-lookup"><span data-stu-id="3152c-169">The `messageParent` function can only be called on a page with the same domain (including protocol and port) as the host page.</span></span>

<span data-ttu-id="3152c-170">以下の例では、`googleProfile` は文字列に変換されたバージョンのユーザーの Google プロファイルです。</span><span class="sxs-lookup"><span data-stu-id="3152c-170">In the next example, `googleProfile` is a stringified version of the user's Google profile.</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

<span data-ttu-id="3152c-p119">ホスト ページは、メッセージを受信するように構成する必要があります。これを構成するには、`displayDialogAsync` の元の呼び出しにコールバック パラメータを追加します。コールバックはハンドラを `DialogMessageReceived` イベントに割り当てます。以下に例を示します。</span><span class="sxs-lookup"><span data-stu-id="3152c-p119">The host page must be configured to receive the message. You do this by adding a callback parameter to the original call of `displayDialogAsync`. The callback assigns a handler to the `DialogMessageReceived` event. The following is an example:</span></span>

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
> - <span data-ttu-id="3152c-p120">Office は [AsyncResult]() オブジェクトをコールバックに渡します。このオブジェクトは、ダイアログ ボックスを開こうとした結果を表します。ただし、ダイアログ ボックスでのイベントの結果は表しません。この違いの詳細については、「[エラーとイベントの処理](#handle-errors-and-events)」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="3152c-p120">Office passes an [AsyncResult]() object to the callback. It represents the result of the attempt to open the dialog box. It does not represent the outcome of any events in the dialog box. For more on this distinction, see the section [Handle errors and events](#handle-errors-and-events).</span></span>
> - <span data-ttu-id="3152c-179">`asyncResult` の `value` プロパティは [Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js) オブジェクトに設定されます。このオブジェクトはダイアログ ボックスの実行コンテキストではなく、ホスト ページに存在します。</span><span class="sxs-lookup"><span data-stu-id="3152c-179">The `value` property of the `asyncResult` is set to a [Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js) object, which exists in the host page, not in the dialog box's execution context.</span></span>
> - <span data-ttu-id="3152c-p121">`processMessage` はイベントを処理する関数です。任意の名前を指定できます。</span><span class="sxs-lookup"><span data-stu-id="3152c-p121">The `processMessage` is the function that handles the event. You can give it any name you want.</span></span>
> - <span data-ttu-id="3152c-182">`dialog` 変数は、`processMessage` でも参照されるため、コールバックよりも広い範囲で宣言されます。</span><span class="sxs-lookup"><span data-stu-id="3152c-182">The `dialog` variable is declared at a wider scope than the callback because it is also referenced in `processMessage`.</span></span>

<span data-ttu-id="3152c-183">`DialogMessageReceived` イベントのハンドラの簡単な例を以下に示します。</span><span class="sxs-lookup"><span data-stu-id="3152c-183">The following is a simple example of a handler for the `DialogMessageReceived` event:</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - <span data-ttu-id="3152c-p122">Office は `arg` オブジェクトをハンドラに渡します。その `message` プロパティは、ダイアログの `messageParent` の呼び出しで送信されるブール値または文字列です。この例では、Microsoft アカウントまたは Google などのサービスからのユーザーのプロファイルの文字列に変換された表記です。このため、`JSON.parse` を使用すればオブジェクトに逆シリアル化されます。</span><span class="sxs-lookup"><span data-stu-id="3152c-p122">Office passes the `arg` object to the handler. Its `message` property is the Boolean or string sent by the call of `messageParent` in the dialog. In this example, it is a stringified representation of a user's profile from a service such as Microsoft Account or Google, so it is deserialized back to an object with `JSON.parse`.</span></span>
> - <span data-ttu-id="3152c-p123">`showUserName` 実装は表示されません。作業ウィンドウ上に個人用のウェルカム メッセージが表示される場合があります。</span><span class="sxs-lookup"><span data-stu-id="3152c-p123">The `showUserName` implementation is not shown. It might display a personalized welcome message on the task pane.</span></span>

<span data-ttu-id="3152c-189">ダイアログ ボックスのユーザー操作が完了すると、以下の例に示すようにメッセージ ハンドラはダイアログ ボックスを閉じます。</span><span class="sxs-lookup"><span data-stu-id="3152c-189">When the user interaction with the dialog box is completed, your message handler should close the dialog box, as shown in this example.</span></span>

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
> - <span data-ttu-id="3152c-190">`dialog` オブジェクトは `displayDialogAsync` の呼び出しによって返されるものと同じである必要があります。</span><span class="sxs-lookup"><span data-stu-id="3152c-190">The `dialog` object must be the same one that is returned by the call of `displayDialogAsync`.</span></span>
> - <span data-ttu-id="3152c-191">`dialog.close` の呼び出しは、直ちにダイアログ ボックスを閉じるよう Office に指示します。</span><span class="sxs-lookup"><span data-stu-id="3152c-191">The call of `dialog.close` tells Office to immediately close the dialog box.</span></span>

<span data-ttu-id="3152c-192">これらの手法を使用するサンプル アドインについては、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3152c-192">For a sample add-in that uses these techniques, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="3152c-p124">メッセージを受信した後、アドインで作業ウィンドウの別のページを開く必要がある場合、ハンドラの最後の行として `window.location.replace` メソッド (または `window.location.href`) を使用できます。以下に例を示します。</span><span class="sxs-lookup"><span data-stu-id="3152c-p124">If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example:</span></span>

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

<span data-ttu-id="3152c-195">これを実行するアドインの例については、「[PowerPoint アドインで Microsoft Graph を使用した Excel グラフの挿入](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)」のサンプルを参照してください。</span><span class="sxs-lookup"><span data-stu-id="3152c-195">For an example of an add-in that does this, see the [Insert Excel charts using Microsoft Graph in a PowerPoint Add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) sample.</span></span>

#### <a name="conditional-messaging"></a><span data-ttu-id="3152c-196">条件付きのメッセージング</span><span class="sxs-lookup"><span data-stu-id="3152c-196">Conditional messaging</span></span>
<span data-ttu-id="3152c-p125">ダイアログ ボックスから複数の `messageParent` 呼び出しを送信できますが、`DialogMessageReceived` イベントのホスト ページにあるハンドラは 1 つのみのため、ハンドラーは条件ロジックを使用してさまざまなメッセージを区別する必要があります。たとえば、ユーザーに対して Microsoft アカウントまたは Google などの ID プロバイダーにサインインするよう求めるダイアログ ボックスが表示されると、ダイアログ ボックスはユーザーのプロファイルをメッセージとして送信します。認証が失敗した場合、以下の例のように、ダイアログ ボックスはホスト ページにエラー情報を送信します。</span><span class="sxs-lookup"><span data-stu-id="3152c-p125">Because you can send multiple `messageParent` calls from the dialog box, but you have only one handler in the host page for the `DialogMessageReceived` event, the handler must use conditional logic to distinguish different messages. For example, if the dialog box prompts a user to sign in to an identity provider such as Microsoft Account or Google, it sends the user's profile as a message. If authentication fails, the dialog box sends error information to the host page, as in the following example:</span></span>

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
> - <span data-ttu-id="3152c-200">`loginSuccess` 変数は、ID プロバイダーからの HTTP 応答を読み取ることによって初期化されます。</span><span class="sxs-lookup"><span data-stu-id="3152c-200">The `loginSuccess` variable would be initialized by reading the HTTP response from the identity provider.</span></span>
> - <span data-ttu-id="3152c-p126">`getProfile` 関数と `getError` 関数の実装は表示されません。これらの関数はそれぞれ、クエリ パラメーターまたは HTTP 応答の本文からデータを取得します。</span><span class="sxs-lookup"><span data-stu-id="3152c-p126">The the implementation of the `getProfile` and `getError` functions are not not shown. They each get data from a query parameter or from the body of the HTTP response.</span></span>
> - <span data-ttu-id="3152c-p127">サインインが成功したかどうかに応じて、さまざまな種類の匿名のオブジェクトが送信されます。両方の関数に `messageType` プロパティがありますが、一方には `profile` プロパティ、もう一方には `error` プロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="3152c-p127">Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.</span></span>

<span data-ttu-id="3152c-205">条件付きメッセージを使用するサンプルについては、以下を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3152c-205">For samples that use conditional messaging, see:</span></span>
- [<span data-ttu-id="3152c-206">Auth0 サービスを使用してソーシャル ログインを簡略化する Office アドイン</span><span class="sxs-lookup"><span data-stu-id="3152c-206">Office Add-in that uses the Auth0 Service to Simplify Social Login</span></span>](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [<span data-ttu-id="3152c-207">OAuth.io サービスを使用して大手のオンライン サービスへのアクセスを簡略化する Office アドイン</span><span class="sxs-lookup"><span data-stu-id="3152c-207">Office Add-in that uses the OAuth.io Service to Simplify Access to Popular Online Services</span></span>](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

<span data-ttu-id="3152c-p128">以下の例に示すように、ホスト ページのハンドラ コードは分岐に `messageType` プロパティの値を使用します。`showUserName` 関数は上記の例と同じであり、`showNotification` 関数はホスト ページの UI にエラーを表示することに注意してください。</span><span class="sxs-lookup"><span data-stu-id="3152c-p128">The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.</span></span>

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

### <a name="closing-the-dialog-box"></a><span data-ttu-id="3152c-210">ダイアログ ボックスを閉じる</span><span class="sxs-lookup"><span data-stu-id="3152c-210">Closing the dialog box</span></span>

<span data-ttu-id="3152c-p129">ダイアログ ボックスを閉じるボタンをダイアログ ボックス内に実装できます。これを実行するには、ボタンのクリック イベント ハンドラは `messageParent` を使用して、ボタンがクリックされたことをホスト ページに通知する必要があります。以下に例を示します。</span><span class="sxs-lookup"><span data-stu-id="3152c-p129">You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example:</span></span>

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};            
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

<span data-ttu-id="3152c-p130">`DialogMessageReceived` のホスト ページ ハンドラは、この例のように `dialog.close` を呼び出します (ダイアログ オブジェクトを初期化する方法を示す、前述の例を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="3152c-p130">The host page handler for `DialogMessageReceived` would call `dialog.close`, as in this example. (See previous examples that show how the dialog object is initialized.)</span></span>


```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

<span data-ttu-id="3152c-216">この手法を使用するサンプルについては、「[Office アドインの UX 設計パターン](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)」リポジトリの[ダイアログ ナビゲーション設計パターン](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3152c-216">For a sample that uses this technique, see the [dialog navigation design pattern](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation) in the [UX design patterns for Office Add-ins](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) repo.</span></span>

<span data-ttu-id="3152c-p131">独自の終了ダイアログ UI がない場合でも、エンド ユーザーは右上隅にある **X** を選択してダイアログ ボックスを閉じることができます。この操作により `DialogEventReceived` イベントがトリガーされます。イベントがトリガーされたときにホスト ウィンドウに通知する必要がある場合、ホスト ウィンドウはこのイベントのハンドラを宣言する必要があります。詳細については、「[ダイアログ ウィンドウでのエラーとイベント](#errors-and-events-in-the-dialog-window)」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="3152c-p131">Even when you don't have your own close dialog UI, an end user can close the dialog box by choosing the **X** in the upper-right corner. This action triggers the `DialogEventReceived` event. If your host pane needs to know when this happens, it should declare a handler for this event. See the section [Errors and events in the dialog window](#errors-and-events-in-the-dialog-window) for details.</span></span>

## <a name="handle-errors-and-events"></a><span data-ttu-id="3152c-221">エラーとイベントを処理する</span><span class="sxs-lookup"><span data-stu-id="3152c-221">Handle errors and events</span></span>

<span data-ttu-id="3152c-222">コードでイベントの 2 つのカテゴリを処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3152c-222">Your code should handle two categories of events:</span></span>

- <span data-ttu-id="3152c-223">ダイアログ ボックスを作成できないために `displayDialogAsync` の呼び出しによって返されるエラー。</span><span class="sxs-lookup"><span data-stu-id="3152c-223">Errors returned by the call of `displayDialogAsync` because the dialog box cannot be created.</span></span>
- <span data-ttu-id="3152c-224">ダイアログ ウィンドウでのエラーと他のイベント。</span><span class="sxs-lookup"><span data-stu-id="3152c-224">Errors, and other events, in the dialog window.</span></span>

### <a name="errors-from-displaydialogasync"></a><span data-ttu-id="3152c-225">displayDialogAsync のエラー</span><span class="sxs-lookup"><span data-stu-id="3152c-225">Errors from displayDialogAsync</span></span>

<span data-ttu-id="3152c-226">一般的なプラットフォーム エラーおよびシステム エラーの他に、`displayDialogAsync` の呼び出しに特有の以下の 3 つのエラーがあります。</span><span class="sxs-lookup"><span data-stu-id="3152c-226">In addition to general platform and system errors, three errors are specific to calling `displayDialogAsync`.</span></span>

|<span data-ttu-id="3152c-227">コード番号</span><span class="sxs-lookup"><span data-stu-id="3152c-227">Code number</span></span>|<span data-ttu-id="3152c-228">意味</span><span class="sxs-lookup"><span data-stu-id="3152c-228">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="3152c-229">12004</span><span class="sxs-lookup"><span data-stu-id="3152c-229">12004</span></span>|<span data-ttu-id="3152c-p132">`displayDialogAsync` に渡される URL のドメインは信頼されていません。ドメインは、ホスト ページと同じドメインである必要があります (プロトコルとポート番号を含む)。</span><span class="sxs-lookup"><span data-stu-id="3152c-p132">The domain of the URL passed to `displayDialogAsync` is not trusted. The domain must be the same domain as the host page (including protocol and port number).</span></span>|
|<span data-ttu-id="3152c-232">12005</span><span class="sxs-lookup"><span data-stu-id="3152c-232">12005</span></span>|<span data-ttu-id="3152c-p133">`displayDialogAsync` に渡される URL には HTTP プロトコルを使用します。HTTPS が必須です。(Office の一部のバージョンでは、12004 に返されるのと同じエラー メッセージが 12005 でも返されます。)</span><span class="sxs-lookup"><span data-stu-id="3152c-p133">The URL passed to `displayDialogAsync` uses the HTTP protocol. HTTPS is required. (In some versions of Office, the error message returned with 12005 is the same one returned for 12004.)</span></span>|
|<span data-ttu-id="3152c-236"><span id="12007">12007</span></span><span class="sxs-lookup"><span data-stu-id="3152c-236"><span id="12007">12007</span></span></span>|<span data-ttu-id="3152c-p134">ダイアログ ボックスは、このホスト ウィンドウで既に開いています。作業ウィンドウなどのホスト ウィンドウで一度に開けるダイアログ ボックスは 1 つだけです。</span><span class="sxs-lookup"><span data-stu-id="3152c-p134">A dialog box is already opened from this host window. A host window, such as a task pane, can only have one dialog box open at a time.</span></span>|

<span data-ttu-id="3152c-p135">`displayDialogAsync` が呼び出されると、常に [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js) オブジェクトがコールバック関数に渡されます。呼び出しが成功した場合 (つまり、ダイアログ ウィンドウが開いた場合)、`AsyncResult` オブジェクトの `value` プロパティは [Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js) オブジェクトです。この例は、「[ダイアログ ボックスからホスト ページに情報を送信する](#send-information-from-the-dialog-box-to-the-host-page)」セクションで参照できます。`displayDialogAsync` への呼び出しが失敗した場合は、ウィンドウは作成されず、`AsyncResult` オブジェクトの `status` プロパティが `Office.AsyncResultStatus.Failed` に設定され、そのオブジェクトの `error` プロパティが追加されます。`status`をテストして、エラーが発生したときに応答するコールバックを常に設定しておく必要があります。コード番号に関係なくエラー メッセージのみを報告する例については、以下のコードを参照してください。</span><span class="sxs-lookup"><span data-stu-id="3152c-p135">When  is called, it always passes an AsyncResult object to its callback function. When the call is successful - that is, the dialog window is opened - the  property of the  object is a Dialog object. An example of this is in the section Send information from the dialog box to the host page. When the call to  fails, the window is not created, the  property of the  object is set to "failed", and the  property of the object is populated. You should always have a callback that tests the  and responds when it's an error. For an example that simply reports the error message regardless of its code number, see the following code:</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        showNotification(asynceResult.error.code = ": " + asyncResult.error.message);
    } else {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
});
```

### <a name="errors-and-events-in-the-dialog-window"></a><span data-ttu-id="3152c-245">ダイアログ ウィンドウでのエラーとイベント</span><span class="sxs-lookup"><span data-stu-id="3152c-245">Errors and events in the dialog window</span></span>

<span data-ttu-id="3152c-246">ダイアログ ボックス内の、コード番号で知られている 3 つのエラーとイベントによって、ホスト ページで `DialogEventReceived` イベントがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="3152c-246">Three errors and events, known by their code numbers, in the dialog box will trigger a `DialogEventReceived` event in the host page.</span></span>

|<span data-ttu-id="3152c-247">コード番号</span><span class="sxs-lookup"><span data-stu-id="3152c-247">Code number</span></span>|<span data-ttu-id="3152c-248">意味</span><span class="sxs-lookup"><span data-stu-id="3152c-248">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="3152c-249">12002</span><span class="sxs-lookup"><span data-stu-id="3152c-249">12002</span></span>|<span data-ttu-id="3152c-250">以下が考えられます:</span><span class="sxs-lookup"><span data-stu-id="3152c-250">One of the following:</span></span><br> <span data-ttu-id="3152c-251">`displayDialogAsync` に渡された URL にページが存在しない。</span><span class="sxs-lookup"><span data-stu-id="3152c-251">- No page exists at the URL that was passed to `displayDialogAsync`.</span></span><br> <span data-ttu-id="3152c-252">`displayDialogAsync` に渡されたページが読み込まれたが、ダイアログ ボックスが見つからないか読み込むことができないページを指していたか、またはダイアログ ボックスが無効な構文を含む URL を指している。</span><span class="sxs-lookup"><span data-stu-id="3152c-252">- The page that was passed to `displayDialogAsync` loaded, but the dialog box was directed to a page that it cannot find or load, or it has been directed to a URL with invalid syntax.</span></span>|
|<span data-ttu-id="3152c-253">12003</span><span class="sxs-lookup"><span data-stu-id="3152c-253">12003</span></span>|<span data-ttu-id="3152c-p136">ダイアログ ボックスが HTTP プロトコルを使用している URL を指していました。HTTPS が必須です。</span><span class="sxs-lookup"><span data-stu-id="3152c-p136">The dialog box was directed to a URL with the HTTP protocol. HTTPS is required.</span></span>|
|<span data-ttu-id="3152c-256">12006</span><span class="sxs-lookup"><span data-stu-id="3152c-256">12006</span></span>|<span data-ttu-id="3152c-257">ダイアログ ボックスが閉じられました。通常は、ユーザーが **X** ボタンを選択したためです。</span><span class="sxs-lookup"><span data-stu-id="3152c-257">The dialog box was closed, usually because the user chooses the **X** button.</span></span>|

<span data-ttu-id="3152c-p137">コードで、呼び出し内の `DialogEventReceived` イベントのハンドラを `displayDialogAsync` に割り当てることができます。以下に簡単な例を示します。</span><span class="sxs-lookup"><span data-stu-id="3152c-p137">Your code can assign a handler for the `DialogEventReceived` event in the call to `displayDialogAsync`. The following is a simple example:</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

<span data-ttu-id="3152c-260">各エラー コードのカスタム エラー メッセージを作成する `DialogEventReceived` イベントのハンドラーの例を、以下に示します。</span><span class="sxs-lookup"><span data-stu-id="3152c-260">For an example of a handler for the `DialogEventReceived` event that creates custom error messages for each error code, see the following example:</span></span>

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

<span data-ttu-id="3152c-261">この方法でエラーを処理するサンプル アドインについては、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3152c-261">For a sample add-in that handles errors in this way, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>


## <a name="pass-information-to-the-dialog-box"></a><span data-ttu-id="3152c-262">情報をダイアログ ボックスに渡す</span><span class="sxs-lookup"><span data-stu-id="3152c-262">Pass information to the dialog box</span></span>

<span data-ttu-id="3152c-p138">ホスト ページがダイアログ ボックスに情報を渡す必要がある場合もあります。これは主に 2 つの方法で実行することができます。</span><span class="sxs-lookup"><span data-stu-id="3152c-p138">Sometimes the host page needs to pass information to the dialog box. You can do this in two primary ways:</span></span>

- <span data-ttu-id="3152c-265">`displayDialogAsync` に渡される URL にクエリ パラメーターを追加します。</span><span class="sxs-lookup"><span data-stu-id="3152c-265">Add query parameters to the URL that is passed to `displayDialogAsync`.</span></span>
- <span data-ttu-id="3152c-p139">ホスト ウィンドウとダイアログ ボックスの両方にアクセス可能な場所に情報を格納します。2 つのウィンドウは共通のセッション ストレージを共有しませんが、ポート番号 (存在する場合) を含む*ドメインが同じである場合*は、共通の[ローカル ストレージ](https://www.w3schools.com/html/html5_webstorage.asp)を共有します。</span><span class="sxs-lookup"><span data-stu-id="3152c-p139">Store the information somewhere that is accessible to both the host window and dialog box. The two windows do not share a common session storage, but *if they have the same domain* (including port number, if any),  they share a common [local storage](https://www.w3schools.com/html/html5_webstorage.asp).</span></span>

### <a name="use-local-storage"></a><span data-ttu-id="3152c-268">ローカル ストレージの使用</span><span class="sxs-lookup"><span data-stu-id="3152c-268">Use local storage</span></span>

<span data-ttu-id="3152c-269">ローカル ストレージを使用するため、以下の例に示すように、`displayDialogAsync` 呼び出しの前に、コードはホスト ページで `window.localStorage` オブジェクトの `setItem` メソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="3152c-269">To use local storage, your code calls the `setItem` method of the `window.localStorage` object in the host page before the `displayDialogAsync` call, as in the following example:</span></span>

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

<span data-ttu-id="3152c-270">ダイアログ ウィンドウ内のコードは、以下の例に示すように、必要に応じて項目を読み取ります。</span><span class="sxs-lookup"><span data-stu-id="3152c-270">Code in the dialog window reads the item when it's needed, as in the following example:</span></span>

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

<span data-ttu-id="3152c-271">この方法でローカル ストレージを使用するサンプル アドインについては、以下を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3152c-271">For sample add-ins that uses local storage in this way, see:</span></span>

- [<span data-ttu-id="3152c-272">Auth0 サービスを使用してソーシャル ログインを簡略化する Office アドイン</span><span class="sxs-lookup"><span data-stu-id="3152c-272">Office Add-in that uses the Auth0 Service to Simplify Social Login</span></span>](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [<span data-ttu-id="3152c-273">OAuth.io サービスを使用して大手のオンライン サービスへのアクセスを簡略化する Office アドイン</span><span class="sxs-lookup"><span data-stu-id="3152c-273">Office Add-in that uses the OAuth.io Service to Simplify Access to Popular Online Services</span></span>](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

### <a name="use-query-parameters"></a><span data-ttu-id="3152c-274">クエリ パラメーターの使用</span><span class="sxs-lookup"><span data-stu-id="3152c-274">Use query parameters</span></span>

<span data-ttu-id="3152c-275">次の例は、クエリ パラメーターを使用してデータを渡す方法を示します。</span><span class="sxs-lookup"><span data-stu-id="3152c-275">The following example shows how to pass data with a query parameter:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

<span data-ttu-id="3152c-276">この手法を使用するサンプルについては、「[PowerPoint アドインで Microsoft Graph を使用した Excel グラフの挿入](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3152c-276">For a sample that uses this technique, see [Insert Excel charts using Microsoft Graph in a PowerPoint Add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span></span>

<span data-ttu-id="3152c-277">ダイアログ ウィンドウ内のコードは、URL を解析し、パラメータ値を読み取ります。</span><span class="sxs-lookup"><span data-stu-id="3152c-277">Code in your dialog window can parse the URL and read the parameter value.</span></span>

> [!NOTE]
> <span data-ttu-id="3152c-p140">Office は、`displayDialogAsync` に渡される URL に `_host_info` というクエリ パラメーターを自動的に追加します (カスタム クエリ パラメーターが存在する場合は、その後に追加されます。ダイアログ ボックスが移動する先の後続の URL には追加されません)。Microsoft は、将来、この値の内容を変更したり、完全に削除したりする可能性があるため、コードでこの値の内容を読み取らないでください。ダイアログ ボックスのセッション ストレージには、同じ値が追加されます。この場合も、*コードではこの値に対する読み取りも書き込みも行わないでください*。</span><span class="sxs-lookup"><span data-stu-id="3152c-p140">Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`. (It is appended after your custom query parameters, if any. It is not appended to any subsequent URLs that the dialog box navigates to.) Microsoft may change the content of this value, or remove it entirely, in the future, so your code should not read it. The same value is added to the dialog box's session storage. Again, *your code should neither read nor write to this value*.</span></span>

## <a name="use-the-dialog-apis-to-show-a-video"></a><span data-ttu-id="3152c-283">ダイアログ API を使用してビデオを表示する</span><span class="sxs-lookup"><span data-stu-id="3152c-283">Use the Dialog APIs to show a video</span></span>

<span data-ttu-id="3152c-284">ダイアログ ボックスでビデオを表示するには</span><span class="sxs-lookup"><span data-stu-id="3152c-284">To show a video in a dialog box:</span></span>

1.  <span data-ttu-id="3152c-p141">コンテンツが iframe のみのページを作成します。iframe の `src` 属性はオンライン ビデオを指しています。ビデオの URL のプロトコルは HTTP**S** である必要があります。この記事では、このページを "video.dialogbox.html" と呼びます。マークアップの例を以下に示します。</span><span class="sxs-lookup"><span data-stu-id="3152c-p141">Create a page whose only content is an iframe. The `src` attribute of the iframe points to an online video. The protocol of the video's URL must be HTTP**S**. In this article we'll call this page "video.dialogbox.html". The following is an example of the markup:</span></span>

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2.  <span data-ttu-id="3152c-290">video.dialogbox.html ページは、ホスト ページと同じドメインにある必要があります。</span><span class="sxs-lookup"><span data-stu-id="3152c-290">The video.dialogbox.html page must be in the same domain as the host page.</span></span>
3.  <span data-ttu-id="3152c-291">ホスト ページで `displayDialogAsync` の呼び出しを使用して、video.dialogbox.html を開きます。</span><span class="sxs-lookup"><span data-stu-id="3152c-291">Use a call of `displayDialogAsync` in the host page to open video.dialogbox.html.</span></span>
4.  <span data-ttu-id="3152c-p142">ユーザーがダイアログ ボックスを閉じたときに、アドインに通知する必要がある場合は、`DialogEventReceived` イベントのハンドラを登録して、12006 イベントを処理します。詳細については、「[ダイアログ ウィンドウでのエラーとイベント](#errors-and-events-in-the-dialog-window)」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="3152c-p142">If your add-in needs to know when the user closes the dialog box, register a handler for the `DialogEventReceived` event and handle the 12006 event. For details, see the section [Errors and events in the dialog window](#errors-and-events-in-the-dialog-window).</span></span>

<span data-ttu-id="3152c-294">ダイアログ ボックスにビデオを表示するサンプルについては、「[Office アドインの UX 設計パターン](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)」リポジトリの[ビデオ プレースマット設計パターン](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3152c-294">For a sample that shows a video in a dialog box, see the [video placemat design pattern](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat) in the [UX design patterns for Office Add-ins](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) repo.</span></span>

![アドイン ダイアログ ボックスに表示されるビデオのスクリーン ショット](../images/video-placemats-dialog-open.png)

## <a name="use-the-dialog-apis-in-an-authentication-flow"></a><span data-ttu-id="3152c-296">認証フローでダイアログ API を使用する</span><span class="sxs-lookup"><span data-stu-id="3152c-296">Use the Dialog APIs in an authentication flow</span></span>

<span data-ttu-id="3152c-297">ダイアログ API の主要なシナリオは、Microsoft アカウント、Office 365、Google、Facebook など、iframe でサインイン ページが開くことを許可しないリソースまたは ID プロバイダーを使用して認証を有効にすることです。</span><span class="sxs-lookup"><span data-stu-id="3152c-297">A primary scenario for the Dialog APIs is to enable authentication with a resource or identity provider that does not allow its sign-in page to open in an Iframe, such as Microsoft Account, Office 365, Google, and Facebook.</span></span>

> [!NOTE]
> <span data-ttu-id="3152c-p143">このシナリオでダイアログ API を使用する際には、`displayDialogAsync` への呼び出しで `displayInIframe: true` オプションを使用*しないでください*。このオプションの詳細については、この記事で先述したように [Office Online のパフォーマンス オプションを利用](#take-advantage-of-a-performance-option-in-office-online)してください。</span><span class="sxs-lookup"><span data-stu-id="3152c-p143">When you are using the Dialog APIs for this scenario, do *not* use the `displayInIframe: true` option in the call to `displayDialogAsync`. See [Take advantage of a performance option in Office Online](#take-advantage-of-a-performance-option-in-office-online) previously in this article for details about this option.</span></span>

<span data-ttu-id="3152c-300">シンプルで標準的な認証フローを、以下に示します。</span><span class="sxs-lookup"><span data-stu-id="3152c-300">The following is a simple and typical authentication flow:</span></span>

1. <span data-ttu-id="3152c-p144">ダイアログ ボックスで開く最初のページは、アドインのドメイン (つまりホスト ウィンドウのドメイン) でホストされるローカル ページ (または他のリソース) です。このページには、"*NAME-OF-PROVIDER* にサインインが可能なページにリダイレクトしていますので、お待ちください。" という簡単な UI を含めることができます。「[情報をダイアログ ボックスに渡す](#pass-information-to-the-dialog-box)」に記載されているように、このページのコードは、ダイアログ ボックスに渡される情報を使用して、ID プロバイダーのサインイン ページの URL を構築します。</span><span class="sxs-lookup"><span data-stu-id="3152c-p144">The first page that opens in the dialog box is a local page (or other resource) that is hosted in the add-in's domain; that is, the host window's domain. This page can have a simple UI that says "Please wait, we are redirecting you to the page where you can sign in to *NAME-OF-PROVIDER*." Code in this page constructs the URL of the identity provider's sign-in page by using information that is passed to the dialog box as described in [Pass information to the dialog box](#pass-information-to-the-dialog-box).</span></span>
2. <span data-ttu-id="3152c-p145">次に、ダイアログ ウィンドウをサインイン ページにリダイレクトします。URL には、ユーザーがサインインしたらダイアログ ウィンドウを特定のページにリダイレクトするように ID プロバイダーに指示するクエリ パラメーターが含まれています。この記事では、このページを "redirectPage.html" と呼びます。(*このページはホスト ウィンドウと同じドメイン内のページにする必要があります*。これは、ダイアログ ウィンドウがサインイン試行の結果を渡す唯一の方法が `messageParent` の呼び出しを使用することであるためです。この呼び出しは、ホスト ウィンドウと同じドメインのページでしか行うことができません。)</span><span class="sxs-lookup"><span data-stu-id="3152c-p145">The dialog window then redirects to the sign-in page. The URL includes a query parameter that tells the identity provider to redirect the dialog window, after the user signs in, to a specific page. In this article, we'll call this page "redirectPage.html". (*This must be a page in the same domain as the host window*, because the only way for the dialog window to pass the results of the sign-in attempt is with a call of `messageParent`, which can only be called on a page with the same domain as the host window.)</span></span>
2. <span data-ttu-id="3152c-p146">ID プロバイダーのサービスは、ダイアログ ウィンドウからの着信 GET 要求を処理します。ユーザーが既にログオンしている場合は、直ちにウィンドウを redirectPage.html にリダイレクトして、ユーザー データをクエリ パラメーターとして含めます。ユーザーがまだサインインしていない場合は、プロバイダーのサインイン ページがウィンドウに表示され、ユーザーがサインインします。ほとんどのプロバイダーでは、ユーザーが正常にサインインできない場合、プロバイダーはダイアログ ウィンドウにエラー ページを表示して、redirectPage.html にはリダイレクトしません。ユーザーは隅にある **X** を選択して、ウィンドウを閉じる必要があります。ユーザーが正常にサインインした場合は、ダイアログ ウィンドウが redirectPage.html にリダイレクトされ、ユーザー データがクエリ パラメーターとして含まれます。</span><span class="sxs-lookup"><span data-stu-id="3152c-p146">The identity provider's service processes the incoming GET request from the dialog window. If the user is already logged on, it immediately redirects the window to redirectPage.html and includes user data as a query parameter. If the user is not already signed in, the provider's sign-in page appears in the window, and the user signs in. For most providers, if the user cannot sign in successfully, the provider shows an error page in the dialog window and does not redirect to redirectPage.html. The user must close the window by selecting the **X** in the corner. If the user successfully signs in, the dialog window is redirected to redirectPage.html and user data is included as a query parameter.</span></span>
3. <span data-ttu-id="3152c-314">edirectPage.html ページが開くと、`messageParent` を呼び出して、成功または失敗をホスト ページに報告し、また必要に応じて、ユーザー データまたはエラー データも報告します。</span><span class="sxs-lookup"><span data-stu-id="3152c-314">When the redirectPage.html page opens, it calls `messageParent` to report the success or failure to the host page and optionally also report user data or error data.</span></span>
4. <span data-ttu-id="3152c-315">`DialogMessageReceived` イベントがホスト ページで発生し、そのハンドラがダイアログ ウィンドウを閉じ、メッセージの他の処理を必要に応じて実行します。</span><span class="sxs-lookup"><span data-stu-id="3152c-315">The `DialogMessageReceived` event fires in the host page and its handler closes the dialog window and optionally does other processing of the message.</span></span>

<span data-ttu-id="3152c-316">このパターンを使用するサンプル アドインについては、以下を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3152c-316">For sample add-ins that use this pattern, see:</span></span>

- <span data-ttu-id="3152c-p147">[PowerPoint アドインで Microsoft Graph を使用した Excel グラフの挿入](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart):ダイアログ ウィンドウで最初に開かれるリソースは、独自のビューがないコントローラーのメソッドです。これは次に、Office 365 のサインイン ページにリダイレクトします。</span><span class="sxs-lookup"><span data-stu-id="3152c-p147">[Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart): The resource that is initially opened in the dialog window is a controller method that has no view of its own. It redirects to the Office 365 sign in page.</span></span>
- <span data-ttu-id="3152c-319">[Office アドイン Office 365 のクライアント認証 AngularJS 用](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth):ダイアログ ウィンドウで最初に開かれるリソースは、ページです。</span><span class="sxs-lookup"><span data-stu-id="3152c-319">[Office Add-in Office 365 Client Authentication for AngularJS](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth): The resource that is initially opened in the dialog window is a page.</span></span>

#### <a name="support-multiple-identity-providers"></a><span data-ttu-id="3152c-320">複数の ID プロバイダーのサポート</span><span class="sxs-lookup"><span data-stu-id="3152c-320">Support multiple identity providers</span></span>

<span data-ttu-id="3152c-p148">アドインによってユーザーが Microsoft アカウント、Google、Facebook などのプロバイダーを選択できる場合は、ユーザーがプロバイダーを選択するための UI を提供するローカルの最初のページ (前述のセクションを参照) が必要です。選択すると、サインイン URL とその URL へのリダイレクトの構築がトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="3152c-p148">If your add-in gives the user a choice of providers, such as Microsoft Account, Google, or Facebook, you need a local first page (see preceding section) that provides a UI for the user to select a provider. Selection triggers the construction of the sign-in URL and redirection to it.</span></span>

<span data-ttu-id="3152c-323">このパターンを使用するサンプルについては、「[Auth0 サービスを使用してソーシャル ログインを簡略化する Office アドイン](https://github.com/OfficeDev/Office-Add-in-Auth0)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3152c-323">For a sample that uses this pattern, see [Office Add-in that uses the Auth0 Service to Simplify Social Login](https://github.com/OfficeDev/Office-Add-in-Auth0).</span></span>

#### <a name="authorization-of-the-add-in-to-an-external-resource"></a><span data-ttu-id="3152c-324">外部リソースへのアドインの承認</span><span class="sxs-lookup"><span data-stu-id="3152c-324">Authorization of the add-in to an external resource</span></span>

<span data-ttu-id="3152c-p149">最新の Web では、Web アプリケーションはユーザーと同等の重要なセキュリティ プリンシパルであり、アプリケーションは Office 365、Google+、Facebook、LinkedIn などのオンライン リソースに対する独自の ID とアクセス許可を持っています。アプリケーションは、展開前にリソース プロバイダーに登録されます。登録には以下が含まれています。</span><span class="sxs-lookup"><span data-stu-id="3152c-p149">In the modern web, web applications are security principals just as users are, and the application has its own identity and permissions to an online resource such as Office 365, Google Plus, Facebook, or LinkedIn. The application is registered with the resource provider before it is deployed. The registration includes:</span></span>

- <span data-ttu-id="3152c-328">アプリケーションが必要とする、ユーザーのリソースへのアクセス許可の一覧。</span><span class="sxs-lookup"><span data-stu-id="3152c-328">A list of the permissions that the application needs to a user's resources.</span></span>
- <span data-ttu-id="3152c-329">アプリケーションがサービスにアクセスするときに、リソース サービスがアクセス トークンを返す宛先の URL。</span><span class="sxs-lookup"><span data-stu-id="3152c-329">A URL to which the resource service should return an access token when the application accesses the service.</span></span>  

<span data-ttu-id="3152c-p150">リソース サービスのユーザーのデータにアクセスするアプリケーションでユーザーが関数を呼び出すと、ユーザーはサービスにサインインするように求められ、アプリケーションが必要とするユーザーのリソースへのアクセス許可をアプリケーションに付与するように求められます。次に、サービスはサインイン ウィンドウを既に登録済みの URL にリダイレクトし、アクセス トークンを渡します。アプリケーションはアクセス トークンを使用して、ユーザーのリソースにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="3152c-p150">When a user invokes a function in the application that accesses the user's data in the resource service, they are prompted to sign in to the service and then prompted to grant the application the permissions it needs to the user's resources. The service then redirects the sign-in window to the previously registered URL and passes the access token. The application uses the access token to access the user's resources.</span></span>

<span data-ttu-id="3152c-p151">ユーザーのサインイン用に示されているフローと類似したフローを使用すると、ダイアログ API を使用してこのプロセスを管理できます。違いは以下の点のみです。</span><span class="sxs-lookup"><span data-stu-id="3152c-p151">You can use the Dialog APIs to manage this process by using a flow that is similar to the one described for users to sign in. The only differences are:</span></span>

- <span data-ttu-id="3152c-335">ユーザーがアプリケーションが必要とするアクセス許可をアプリケーションに付与したことがない場合は、サインインすると、ユーザーに対してこれを実行するよう求めるメッセージがダイアログ ボックスに表示されます。</span><span class="sxs-lookup"><span data-stu-id="3152c-335">If the user hasn't previously granted the application the permissions it needs, she is prompted to do so in the dialog box after signing in.</span></span>
- <span data-ttu-id="3152c-p152">ダイアログ ウィンドウは、`messageParent` を使用して文字列に変換されたアクセス トークンを送信するか、またはホスト ウィンドウがアクセス トークンを取得できる場所にアクセス トークンを格納することで、アクセス トークンをホスト ウィンドウに送信します。トークンには制限時間がありますが、制限時間内であれば追加のメッセージを表示することなく、ホスト ウィンドウはトークンを使用して、ユーザーのリソースに直接アクセスできます。</span><span class="sxs-lookup"><span data-stu-id="3152c-p152">The dialog window sends the access token to the host window either by using `messageParent` to send the stringified access token or by storing the access token where the host window can retrieve it. The token has a time limit, but while it lasts, the host window can use it to directly access the user's resources without any further prompting.</span></span>

<span data-ttu-id="3152c-338">これを実現するため、次のサンプルはダイアログ API を使用します。</span><span class="sxs-lookup"><span data-stu-id="3152c-338">The following samples use the Dialog APIs for this purpose:</span></span>
- <span data-ttu-id="3152c-339">[PowerPoint アドインで Microsoft Graph を使用した Excel グラフの挿入](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) - データベースにアクセス トークンを格納します。</span><span class="sxs-lookup"><span data-stu-id="3152c-339">[Insert Excel charts using Microsoft Graph in a PowerPoint Add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) - Stores the access token in a database.</span></span>
- [<span data-ttu-id="3152c-340">OAuth.io サービスを使用して大手のオンライン サービスへのアクセスを簡略化する Office アドイン</span><span class="sxs-lookup"><span data-stu-id="3152c-340">Office Add-in that uses the OAuth.io Service to Simplify Access to Popular Online Services</span></span>](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

<span data-ttu-id="3152c-341">アドインにおける認証と承認の詳細については、以下を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3152c-341">For more information about authentication and authorization in add-ins, see:</span></span>
- [<span data-ttu-id="3152c-342">Office アドインで外部サービスを承認する</span><span class="sxs-lookup"><span data-stu-id="3152c-342">Authorize external services in your Office Add-in</span></span>](auth-external-add-ins.md)
- [<span data-ttu-id="3152c-343">Office の JavaScript API ヘルパーのライブラリ</span><span class="sxs-lookup"><span data-stu-id="3152c-343">Office JavaScript API Helpers library</span></span>](https://github.com/OfficeDev/office-js-helpers)


## <a name="use-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a><span data-ttu-id="3152c-344">単一ページ アプリケーションとクライアント側ルーティングで Office ダイアログ API を使用する</span><span class="sxs-lookup"><span data-stu-id="3152c-344">Use the Office Dialog API with single-page applications and client-side routing</span></span>

<span data-ttu-id="3152c-345">単一ページ アプリケーションが通常使用するように、アドインがクライアント側ルーティングを使用している場合は、HTML の完了ページと個別ページの URL の代わりに、ルートの URL を [displayDialogAsync](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) メソッドに渡すこともできます。</span><span class="sxs-lookup"><span data-stu-id="3152c-345">If your add-in uses client-side routing, as single-page applications typically do, you have the option to pass the URL of a route to the [displayDialogAsync](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) method, instead of the URL of a complete and separate HTML page.</span></span>

> [!IMPORTANT]
><span data-ttu-id="3152c-p153">ダイアログ ボックスは、独自の実行コンテキストを含む新しいウィンドウ内にあります。ルートを渡すと、ダイアログ ウィンドウで、この新しいコンテキストに対して基本ページとそのすべての初期化、およびブートストラップ コードを再度実行し、すべての変数が初期値に設定されます。この手法により、ダイアログ ウィンドウで、アプリケーションの 2 番目のインスタンスが起動します。ダイアログ ウィンドウ内の変数を変更するコードは、同じ変数の作業ウィンドウのバージョンは変更しません。同様に、ダイアログ ウィンドウには、それ自体にセッション ストレージがあり、作業ウィンドウからコードでそこにアクセスすることはできません。</span><span class="sxs-lookup"><span data-stu-id="3152c-p153">The dialog box is in a new window with its own execution context. If you pass a route, your base page and all its initialization and bootstrapping code run again in this new context, and any variables are set to their initial values in the dialog window. So this technique launches a second instance of your application in the dialog window. Code that changes variables in the dialog window does not change the task pane version of the same variables. Similarly, the dialog window has its own session storage, which is not accessible from code in the task pane.</span></span>

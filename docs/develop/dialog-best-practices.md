---
title: Office ダイアログ API のベスト プラクティスとルール
description: 単一ページ アプリケーション (SPA) のベスト プラクティスなどOfficeダイアログ API のルールとベスト プラクティスを提供します。
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 99129636cf722f98cef36c272f2e00e8a9321ccf
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349911"
---
# <a name="best-practices-and-rules-for-the-office-dialog-api"></a><span data-ttu-id="e27e2-103">Office ダイアログ API のベスト プラクティスとルール</span><span class="sxs-lookup"><span data-stu-id="e27e2-103">Best practices and rules for the Office dialog API</span></span>

<span data-ttu-id="e27e2-104">この記事では、ダイアログの UI を設計し、API を単一ページ アプリケーション (SPA) で使用するためのベスト プラクティスを含む、Office ダイアログ API のルール、gotchas、およびベスト プラクティスについて説明します。</span><span class="sxs-lookup"><span data-stu-id="e27e2-104">This article provides rules, gotchas, and best practices for the Office dialog API, including best practices for designing the UI of a dialog and using the API with in a single-page application (SPA)</span></span>

> [!NOTE]
> <span data-ttu-id="e27e2-105">この記事では、「Office アドインで Office ダイアログ API を使用する」の説明に従って、Office ダイアログ[API](dialog-api-in-office-add-ins.md)の使用の基本について理解している必要があります。</span><span class="sxs-lookup"><span data-stu-id="e27e2-105">This article presupposes that you are familiar with the basics of using the Office dialog API as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).</span></span>
> 
> <span data-ttu-id="e27e2-106">「エラーと[イベントの処理とエラーの処理」Officeを参照してください](dialog-handle-errors-events.md)。</span><span class="sxs-lookup"><span data-stu-id="e27e2-106">See also [Handling errors and events with the Office dialog box](dialog-handle-errors-events.md).</span></span>

## <a name="rules-and-gotchas"></a><span data-ttu-id="e27e2-107">ルールと注意事項</span><span class="sxs-lookup"><span data-stu-id="e27e2-107">Rules and gotchas</span></span>

- <span data-ttu-id="e27e2-108">ダイアログ ボックスは HTTP ではなく HTTPS URL にのみ移動できます。</span><span class="sxs-lookup"><span data-stu-id="e27e2-108">The dialog box can only navigate to HTTPS URLs, not HTTP.</span></span>
- <span data-ttu-id="e27e2-109">[displayDialogAsync](/javascript/api/office/office.ui)メソッドに渡される URL は、アドイン自体とまったく同じドメインにある必要があります。</span><span class="sxs-lookup"><span data-stu-id="e27e2-109">The URL passed to the [displayDialogAsync](/javascript/api/office/office.ui) method must be in the exact same domain as the add-in itself.</span></span> <span data-ttu-id="e27e2-110">サブドメインにすることはできません。</span><span class="sxs-lookup"><span data-stu-id="e27e2-110">It cannot be a subdomain.</span></span> <span data-ttu-id="e27e2-111">ただし、そのページに渡されたページは、別のドメインのページにリダイレクトできます。</span><span class="sxs-lookup"><span data-stu-id="e27e2-111">But the page that is passed to it can redirect to a page in another domain.</span></span>
- <span data-ttu-id="e27e2-112">アドイン コマンドの作業ウィンドウまたは UI レス関数ファイルを使用[](../reference/manifest/functionfile.md)できるホスト ウィンドウでは、一度に開くことができるダイアログ ボックスは 1 つのみです。</span><span class="sxs-lookup"><span data-stu-id="e27e2-112">A host window, which can be a task pane or the UI-less [function file](../reference/manifest/functionfile.md) of an add-in command, can have only one dialog box open at a time.</span></span>
- <span data-ttu-id="e27e2-113">ダイアログ ボックスOffice 2 つの API のみを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="e27e2-113">Only two Office APIs can be called in the dialog box:</span></span>
  - <span data-ttu-id="e27e2-114">[messageParent](/javascript/api/office/office.ui#messageparent-message-)関数。</span><span class="sxs-lookup"><span data-stu-id="e27e2-114">The [messageParent](/javascript/api/office/office.ui#messageparent-message-) function.</span></span>
  - <span data-ttu-id="e27e2-115">`Office.context.requirements.isSetSupported`(詳細については、「アプリケーションと[API 要件Office指定する」を参照してください](specify-office-hosts-and-api-requirements.md)。</span><span class="sxs-lookup"><span data-stu-id="e27e2-115">`Office.context.requirements.isSetSupported` (For more information, see [Specify Office applications and API requirements](specify-office-hosts-and-api-requirements.md).)</span></span>
- <span data-ttu-id="e27e2-116">[messageParent 関数](/javascript/api/office/office.ui#messageparent-message-)は、アドイン自体とまったく同じドメイン内のページからのみ呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e27e2-116">The [messageParent](/javascript/api/office/office.ui#messageparent-message-) function can only be called from a page in the exact same domain as the add-in itself.</span></span>

## <a name="best-practices"></a><span data-ttu-id="e27e2-117">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="e27e2-117">Best practices</span></span>

### <a name="avoid-overusing-dialog-boxes"></a><span data-ttu-id="e27e2-118">ダイアログ ボックスの使い過ぎを回避する</span><span class="sxs-lookup"><span data-stu-id="e27e2-118">Avoid overusing dialog boxes</span></span>

<span data-ttu-id="e27e2-119">UI 要素を重ねて表示することはお勧めできないため、シナリオで必要な場合を除き、作業ウィンドウでダイアログ ボックスを開かないようにします。</span><span class="sxs-lookup"><span data-stu-id="e27e2-119">Because overlapping UI elements are discouraged, avoid opening a dialog box from a task pane unless your scenario requires it.</span></span> <span data-ttu-id="e27e2-120">作業ウィンドウの表示領域の使用方法を検討するときには、作業ウィンドウはタブ表示できることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="e27e2-120">When you consider how to use the surface area of a task pane, note that task panes can be tabbed.</span></span> <span data-ttu-id="e27e2-121">タブ付き作業ウィンドウの例については、「Excel [JavaScript SalesTracker サンプル」を参照](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)してください。</span><span class="sxs-lookup"><span data-stu-id="e27e2-121">For an example of a tabbed task pane, see the [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) sample.</span></span>

### <a name="designing-a-dialog-box-ui"></a><span data-ttu-id="e27e2-122">ダイアログ ボックス UI の設計</span><span class="sxs-lookup"><span data-stu-id="e27e2-122">Designing a dialog box UI</span></span>

<span data-ttu-id="e27e2-123">ダイアログ ボックス設計のベスト プラクティスについては、「ダイアログ ボックス」を[参照Officeアドインを参照してください](../design/dialog-boxes.md)。</span><span class="sxs-lookup"><span data-stu-id="e27e2-123">For best practices in dialog box design, see [Dialog boxes in Office Add-ins](../design/dialog-boxes.md).</span></span>

### <a name="handling-pop-up-blockers-with-office-on-the-web"></a><span data-ttu-id="e27e2-124">Office on the web を使用したポップアップ ブロックの処理</span><span class="sxs-lookup"><span data-stu-id="e27e2-124">Handling pop-up blockers with Office on the web</span></span>

<span data-ttu-id="e27e2-125">ブラウザーを使用している間にダイアログ Office on the webを表示しようとすると、ブラウザーのポップアップ ブロッカーがダイアログ ボックスをブロックする可能性があります。</span><span class="sxs-lookup"><span data-stu-id="e27e2-125">Attempting to display a dialog box while using Office on the web may cause the browser's pop-up blocker to block the dialog box.</span></span> <span data-ttu-id="e27e2-126">Office on the webには、アドインのダイアログ ボックスがブラウザーのポップアップ ブロッカーの例外になる機能があります。</span><span class="sxs-lookup"><span data-stu-id="e27e2-126">Office on the web has a feature that enables your add-in's dialog boxes to be an exception to the browser's pop-up blocker.</span></span> <span data-ttu-id="e27e2-127">コードがメソッドを呼び `displayDialogAsync` 出す場合、Office on the web次のようなプロンプトが開きます。</span><span class="sxs-lookup"><span data-stu-id="e27e2-127">When your code calls the `displayDialogAsync` method, then Office on the web will open a prompt similar to the following:</span></span>

![ブラウザー内のポップアップ ブロックを回避するためにアドインが生成できる簡単な説明と [許可] ボタンと [無視] ボタンを含むプロンプトを示すスクリーンショット。](../images/dialog-prompt-before-open.png)

<span data-ttu-id="e27e2-129">ユーザーが [許可] を **選択すると**、[Office] ダイアログ ボックスが開きます。</span><span class="sxs-lookup"><span data-stu-id="e27e2-129">If the user chooses **Allow**, the Office dialog box opens.</span></span> <span data-ttu-id="e27e2-130">ユーザーが [無視] を **選択** すると、プロンプトが閉じOfficeダイアログ ボックスが開かれません。</span><span class="sxs-lookup"><span data-stu-id="e27e2-130">If the user chooses **Ignore**, the prompt closes and the Office dialog box does not open.</span></span> <span data-ttu-id="e27e2-131">代わりに、メソッド `displayDialogAsync` はエラー 12009 を返します。</span><span class="sxs-lookup"><span data-stu-id="e27e2-131">Instead, the `displayDialogAsync` method returns error 12009.</span></span> <span data-ttu-id="e27e2-132">コードは、このエラーをキャッチし、ダイアログを必要としない代替エクスペリエンスを提供するか、アドインがダイアログを許可する必要があるというメッセージをユーザーに表示する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e27e2-132">Your code should catch this error and either provide an alternate experience that does not require a dialog, or display a message to the user advising that the add-in requires them to allow the dialog.</span></span> <span data-ttu-id="e27e2-133">(12009 の詳細については [、「displayDialogAsync からのエラー」を参照](dialog-handle-errors-events.md#errors-from-displaydialogasync)してください)。</span><span class="sxs-lookup"><span data-stu-id="e27e2-133">(For more about 12009, see [Errors from displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).)</span></span>

<span data-ttu-id="e27e2-134">何らかの理由でこの機能をオフにする場合は、コードをオプトアウトする必要があります。この要求は、メソッドに渡 [される DialogOptions](/javascript/api/office/office.dialogoptions) オブジェクトを使用して行 `displayDialogAsync` います。</span><span class="sxs-lookup"><span data-stu-id="e27e2-134">If, for any reason, you want to turn off this feature, then your code must opt out. It makes this request with the [DialogOptions](/javascript/api/office/office.dialogoptions) object that is passed to the `displayDialogAsync` method.</span></span> <span data-ttu-id="e27e2-135">具体的には、オブジェクトに . を含める必要があります `promptBeforeOpen: false` 。</span><span class="sxs-lookup"><span data-stu-id="e27e2-135">Specifically, the object should include `promptBeforeOpen: false`.</span></span> <span data-ttu-id="e27e2-136">このオプションが false に設定されている場合、Office on the web アドインがダイアログを開くことを許可するように求めるメッセージは表示され、Office ダイアログは開かれません。</span><span class="sxs-lookup"><span data-stu-id="e27e2-136">When this option is set to false, Office on the web will not prompt the user to allow the add-in open a dialog, and the Office dialog will not open.</span></span>

### <a name="do-not-use-the-_host_info-value"></a><span data-ttu-id="e27e2-137">ホスト情報の値 \_ を \_ 使用しない</span><span class="sxs-lookup"><span data-stu-id="e27e2-137">Do not use the \_host\_info value</span></span>

<span data-ttu-id="e27e2-138">Office は、`_host_info` に渡される URL に `displayDialogAsync` というクエリ パラメーターを自動的に追加します (カスタム クエリ パラメーターが存在する場合は、その後に追加されます。</span><span class="sxs-lookup"><span data-stu-id="e27e2-138">Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`.</span></span> <span data-ttu-id="e27e2-139">カスタム クエリ パラメーターがある場合は、その後に追加されます。</span><span class="sxs-lookup"><span data-stu-id="e27e2-139">It is appended after your custom query parameters, if any.</span></span> <span data-ttu-id="e27e2-140">ダイアログ ボックスが移動する後続の URL には追加されません。</span><span class="sxs-lookup"><span data-stu-id="e27e2-140">It is not appended to any subsequent URLs that the dialog box navigates to.</span></span> <span data-ttu-id="e27e2-141">Microsoft は、この値の内容を変更したり、完全に削除したりする場合があります。そのため、コードで読み取る必要はありません。</span><span class="sxs-lookup"><span data-stu-id="e27e2-141">Microsoft may change the content of this value, or remove it entirely, so your code should not read it.</span></span> <span data-ttu-id="e27e2-142">同じ値がダイアログ ボックスのセッション ストレージ [(Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) プロパティ) に追加されます。</span><span class="sxs-lookup"><span data-stu-id="e27e2-142">The same value is added to the dialog box's session storage (that is, the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property).</span></span> <span data-ttu-id="e27e2-143">この場合も、*コードではこの値に対する読み取りも書き込みも行わないでください*。</span><span class="sxs-lookup"><span data-stu-id="e27e2-143">Again, *your code should neither read nor write to this value*.</span></span>

### <a name="opening-another-dialog-immediately-after-closing-one"></a><span data-ttu-id="e27e2-144">1 つを閉じるとすぐに別のダイアログを開く</span><span class="sxs-lookup"><span data-stu-id="e27e2-144">Opening another dialog immediately after closing one</span></span>

<span data-ttu-id="e27e2-145">特定のホスト ページから複数のダイアログを開く必要はないので、別のダイアログを開く前に、開いているダイアログで [Dialog.close](/javascript/api/office/office.dialog#close__) を呼び `displayDialogAsync` 出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="e27e2-145">You can't have more than one dialog open from a given host page, so your code should call [Dialog.close](/javascript/api/office/office.dialog#close__) on an open dialog before it calls `displayDialogAsync` to open another dialog.</span></span> <span data-ttu-id="e27e2-146">メソッド `close` は非同期です。</span><span class="sxs-lookup"><span data-stu-id="e27e2-146">The `close` method is asynchronous.</span></span> <span data-ttu-id="e27e2-147">このため、呼び出しの直後に呼び出した場合、2 番目のダイアログを開Officeが完全に閉じない `displayDialogAsync` `close` 可能性があります。</span><span class="sxs-lookup"><span data-stu-id="e27e2-147">For this reason, if you call `displayDialogAsync` immediately after a call of `close`, the first dialog may not have completely closed when Office attempts to open the second.</span></span> <span data-ttu-id="e27e2-148">この場合、Office [12007](dialog-handle-errors-events.md#12007)エラーが返されます。"このアドインには既にアクティブなダイアログが含まれるため、操作は失敗しました。</span><span class="sxs-lookup"><span data-stu-id="e27e2-148">If that happens, Office will return a [12007](dialog-handle-errors-events.md#12007) error: "The operation failed because this add-in already has an active dialog."</span></span>

<span data-ttu-id="e27e2-149">メソッドはコールバック パラメーターを受け入れないので、Promise オブジェクトを返すので、キーワードまたはメソッドで待つ `close` `await` `then` 必要はありません。</span><span class="sxs-lookup"><span data-stu-id="e27e2-149">The `close` method doesn't accept a callback parameter, and it doesn't return a Promise object so it cannot be awaited with either the `await` keyword or with a `then` method.</span></span> <span data-ttu-id="e27e2-150">このため、ダイアログを閉じる直後に新しいダイアログを開く必要がある場合は、メソッドで新しいダイアログを開くコードをカプセル化し、戻り値の呼び出しが発生した場合にメソッドを再帰的に呼び出すメソッドを設計する必要がある場合は、次の方法をお勧めします。 `displayDialogAsync` `12007`</span><span class="sxs-lookup"><span data-stu-id="e27e2-150">For this reason, we suggest the following technique when you need to open a new dialog immediately after closing a dialog: encapsulate the code to open the new dialog in a method and design the method to recursively call itself if the call of `displayDialogAsync` returns `12007`.</span></span> <span data-ttu-id="e27e2-151">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="e27e2-151">The following is an example.</span></span>

```javascript
function openFirstDialog() {
  Office.context.ui.displayDialogAsync("https://MyDomain/firstDialog.html", { width: 50, height: 50},
     (result) => {
      if(result.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = result.value;
        dialog.close();
        openSecondDialog();
      }
      else {
         // Handle errors
      }
    }
  );
}
 
function openSecondDialog() {
  Office.context.ui.displayDialogAsync("https://MyDomain/secondDialog.html", { width: 50, height: 50},
    (result) => {
      if(result.status === Office.AsyncResultStatus.Failed) {
        if (result.error.code === 12007) {
          openSecondDialog(); // Recursive call
        }
        else {
         // Handle other errors
        }
      }
    }
  );
}
```

<span data-ttu-id="e27e2-152">または [、setTimeout](https://www.w3schools.com/jsref/met_win_settimeout.asp) メソッドを使用して 2 番目のダイアログを開く前に、コードを強制的に一時停止できます。</span><span class="sxs-lookup"><span data-stu-id="e27e2-152">Alternatively, you could force the code to pause before it tries to open the second dialog by using the [setTimeout](https://www.w3schools.com/jsref/met_win_settimeout.asp) method.</span></span> <span data-ttu-id="e27e2-153">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="e27e2-153">The following is an example.</span></span>

```javascript
function openFirstDialog() {
  Office.context.ui.displayDialogAsync("https://MyDomain/firstDialog.html", { width: 50, height: 50},
     (result) => {
      if(result.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = result.value;
        dialog.close();
        setTimeout(() => { 
          Office.context.ui.displayDialogAsync("https://MyDomain/secondDialog.html", { width: 50, height: 50},
             (result) => { /* callback body */ }
          );
        }, 1000);
      }
      else {
         // Handle errors
      }
    }
  );
}
```

### <a name="best-practices-for-using-the-office-dialog-api-in-an-spa"></a><span data-ttu-id="e27e2-154">SPA で [Office] ダイアログ API を使用するためのベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="e27e2-154">Best practices for using the Office dialog API in an SPA</span></span>

<span data-ttu-id="e27e2-155">アドインがクライアント側ルーティングを使用する場合は、通常、単一ページ アプリケーション (SPA) のように、別の HTML ページの URL ではなく、ルートの URL を [displayDialogAsync](/javascript/api/office/office.ui) メソッドに渡すオプションがあります。</span><span class="sxs-lookup"><span data-stu-id="e27e2-155">If your add-in uses client-side routing, as single-page applications (SPAs) typically do, you have the option to pass the URL of a route to the [displayDialogAsync](/javascript/api/office/office.ui) method instead of the URL of a separate HTML page.</span></span> <span data-ttu-id="e27e2-156">*以下に示す理由により、これを行うのはお勧めしません。*</span><span class="sxs-lookup"><span data-stu-id="e27e2-156">*We recommend against doing so for the reasons given below.*</span></span>

> [!NOTE]
> <span data-ttu-id="e27e2-157">この記事は、Express *ベース* の Web アプリケーションなど、サーバー側のルーティングには関係ありません。</span><span class="sxs-lookup"><span data-stu-id="e27e2-157">This article is not relevant to *server-side* routing, such as in an Express-based web application.</span></span>

#### <a name="problems-with-spas-and-the-office-dialog-api"></a><span data-ttu-id="e27e2-158">SPA とダイアログ API のOffice問題</span><span class="sxs-lookup"><span data-stu-id="e27e2-158">Problems with SPAs and the Office dialog API</span></span>

<span data-ttu-id="e27e2-159">[Office] ダイアログ ボックスは、JavaScript エンジンの独自のインスタンスを持つ新しいウィンドウに表示され、それ故に、完全な実行コンテキストになります。</span><span class="sxs-lookup"><span data-stu-id="e27e2-159">The Office dialog box is in a new window with its own instance of the JavaScript engine, and hence it's own complete execution context.</span></span> <span data-ttu-id="e27e2-160">ルートを渡した場合、基本ページとその初期化コードとブートストラップ コードはすべて、この新しいコンテキストで再び実行され、すべての変数はダイアログ ボックスの初期値に設定されます。</span><span class="sxs-lookup"><span data-stu-id="e27e2-160">If you pass a route, your base page and all its initialization and bootstrapping code run again in this new context, and any variables are set to their initial values in the dialog box.</span></span> <span data-ttu-id="e27e2-161">したがって、この手法では、アプリケーションの 2 番目のインスタンスがボックス ウィンドウでダウンロードおよび起動され、SPA の目的の一部が打ち負かされます。</span><span class="sxs-lookup"><span data-stu-id="e27e2-161">So this technique downloads and launches a second instance of your application in the  box window, which partially defeats the purpose of an SPA.</span></span> <span data-ttu-id="e27e2-162">さらに、ダイアログ ボックス ウィンドウで変数を変更するコードでは、同じ変数の作業ウィンドウ バージョンは変更されません。</span><span class="sxs-lookup"><span data-stu-id="e27e2-162">In addition, code that changes variables in the dialog box window does not change the task pane version of the same variables.</span></span> <span data-ttu-id="e27e2-163">同様に、ダイアログ ボックス ウィンドウには独自のセッション ストレージ [(Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) プロパティ) があります。これは作業ウィンドウ内のコードからアクセスできません。</span><span class="sxs-lookup"><span data-stu-id="e27e2-163">Similarly, the dialog box window has its own session storage (the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property), which is not accessible from code in the task pane.</span></span> <span data-ttu-id="e27e2-164">ダイアログ ボックスと、呼び出されたホスト ページは、サーバーに対して 2 つの異 `displayDialogAsync` なるクライアントのように見えます。</span><span class="sxs-lookup"><span data-stu-id="e27e2-164">The dialog box and the host page on which `displayDialogAsync` was called look like two different clients to your server.</span></span> <span data-ttu-id="e27e2-165">(ホスト ページの種類を確認するには、「ホスト ページからダイアログ ボックスを開 [く」を参照](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)してください。</span><span class="sxs-lookup"><span data-stu-id="e27e2-165">(For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).)</span></span>

<span data-ttu-id="e27e2-166">したがって、メソッドにルートを渡した場合、SPA は実際には使用しないので、同じ SPA のインスタンスが 2 `displayDialogAsync` *つ必要になります*。</span><span class="sxs-lookup"><span data-stu-id="e27e2-166">So, if you passed a route to the `displayDialogAsync` method, you wouldn't really have an SPA; you'd have *two instances of the same SPA*.</span></span> <span data-ttu-id="e27e2-167">さらに、作業ウィンドウ インスタンス内のコードの多くが、そのインスタンスでは使用されません。ダイアログ ボックス インスタンス内のコードの多くは、そのインスタンスでは使用されません。</span><span class="sxs-lookup"><span data-stu-id="e27e2-167">Moreover, much of the code in the task pane instance would never be used in that instance and much of the code in the dialog box instance would never be used in that instance.</span></span> <span data-ttu-id="e27e2-168">同じバンドルに 2 つの SPA があるようなものです。</span><span class="sxs-lookup"><span data-stu-id="e27e2-168">It would be like having two SPAs in the same bundle.</span></span>

#### <a name="microsoft-recommendations"></a><span data-ttu-id="e27e2-169">Microsoft の推奨事項</span><span class="sxs-lookup"><span data-stu-id="e27e2-169">Microsoft recommendations</span></span>

<span data-ttu-id="e27e2-170">クライアント側ルートをメソッドに渡す代わりに、次のいずれかを `displayDialogAsync` 実行することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="e27e2-170">Instead of passing a client-side route to the `displayDialogAsync` method, we recommend that you do one of the following:</span></span>

* <span data-ttu-id="e27e2-171">ダイアログ ボックスで実行するコードが十分に複雑な場合は、2 つの異なる SPA を明示的に作成します。つまり、同じドメインの異なるフォルダーに 2 つの SPA があります。</span><span class="sxs-lookup"><span data-stu-id="e27e2-171">If the code that you want to run in the dialog box is sufficiently complex, create two different SPAs explicitly; that is, have two SPAs in different folders of the same domain.</span></span> <span data-ttu-id="e27e2-172">1 つの SPA はダイアログ ボックスで実行され、もう 1 つはダイアログ ボックスのホスト ページで呼び出 `displayDialogAsync` されました。</span><span class="sxs-lookup"><span data-stu-id="e27e2-172">One SPA runs in the dialog box and the other in the dialog box's host page where `displayDialogAsync` was called.</span></span> 
* <span data-ttu-id="e27e2-173">ほとんどのシナリオでは、ダイアログ ボックスで必要なのは単純なロジックのみです。</span><span class="sxs-lookup"><span data-stu-id="e27e2-173">In most scenarios, only simple logic is needed in the dialog box.</span></span> <span data-ttu-id="e27e2-174">このような場合、SPA のドメインに JavaScript が埋め込まれているか参照されている単一の HTML ページをホストすることで、プロジェクトが大幅に簡略化されます。</span><span class="sxs-lookup"><span data-stu-id="e27e2-174">In such cases, your project will be greatly simplified by hosting a single HTML page, with embedded or referenced JavaScript, in the domain of your SPA.</span></span> <span data-ttu-id="e27e2-175">ページの URL を `displayDialogAsync` メソッドに渡します。</span><span class="sxs-lookup"><span data-stu-id="e27e2-175">Pass the URL of the page to the `displayDialogAsync` method.</span></span> <span data-ttu-id="e27e2-176">つまり、単一ページ アプリの文字通りの考え方から離れつきます。このダイアログ API を使用している場合、SPA のインスタンスは実際には 1 つOffice必要があります。</span><span class="sxs-lookup"><span data-stu-id="e27e2-176">While this means that you are deviating from the literal idea of a single-page app; you don't really have a single instance of an SPA when you are using the Office dialog API.</span></span>

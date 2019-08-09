---
ms.date: 07/09/2019
description: Excel のカスタム関数を使用してユーザーを認証します。
title: カスタム関数の認証
localization_priority: Priority
ms.openlocfilehash: f746947122da7ef3d54a0dd3b4f90dd059e5830f
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268139"
---
# <a name="authentication-for-custom-functions"></a><span data-ttu-id="28ad4-103">カスタム関数の認証</span><span class="sxs-lookup"><span data-stu-id="28ad4-103">Authentication for custom functions</span></span>

<span data-ttu-id="28ad4-104">一部のシナリオでは、保護されたリソースにアクセスするために、ユーザーを認証する必要があります。</span><span class="sxs-lookup"><span data-stu-id="28ad4-104">In some scenarios your custom function will need to authenticate the user in order to access protected resources.</span></span> <span data-ttu-id="28ad4-105">カスタム関数は特定の認証方法を必要としませんが、カスタム関数は、アドインの作業ウィンドウと他の UI 要素とは別のランタイムで実行されます。</span><span class="sxs-lookup"><span data-stu-id="28ad4-105">While custom functions don't require a specific method of authentication, you should be aware that custom functions run in a separate runtime from the task pane and other UI elements of your add-in.</span></span> <span data-ttu-id="28ad4-106">このため、`OfficeRuntime.storage` オブジェクトとダイアログ API を使用して 2 つのランタイム間でデータを受け渡しする必要があります。</span><span class="sxs-lookup"><span data-stu-id="28ad4-106">Because of this, you'll need to pass data back and forth between the two runtimes using the `OfficeRuntime.storage` object and the Dialog API.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="officeruntimestorage-object"></a><span data-ttu-id="28ad4-107">OfficeRuntime.storage オブジェクト</span><span class="sxs-lookup"><span data-stu-id="28ad4-107">OfficeRuntime.storage object</span></span>

<span data-ttu-id="28ad4-108">カスタム関数ランタイムには、通常はデータを格納するグローバルウィンドウに使用できる `localStorage` オブジェクトがありません。</span><span class="sxs-lookup"><span data-stu-id="28ad4-108">The custom functions runtime doesn't have a `localStorage` object available on the global window, where you might typically store data.</span></span> <span data-ttu-id="28ad4-109">代わりに、データを設定して取得するためのストレージ[OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage)を使用して、カスタム関数と作業ウィンドウの間でデータを共有する必要があります。</span><span class="sxs-lookup"><span data-stu-id="28ad4-109">Instead, you should share data between custom functions and task panes by using [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) to set and get data.</span></span>

<span data-ttu-id="28ad4-110">また、`storage`オブジェクトを使用すると便利です。セキュリティサンドボックス環境を使用するため、他のアドインがデータにアクセスすることができません。</span><span class="sxs-lookup"><span data-stu-id="28ad4-110">Additionally, there is a benefit to using the `storage` object; it uses a secure sandbox environment so that your data cannot be accessed by other add-ins.</span></span>

### <a name="suggested-usage"></a><span data-ttu-id="28ad4-111">おすすめの使用法</span><span class="sxs-lookup"><span data-stu-id="28ad4-111">Suggested usage</span></span>

<span data-ttu-id="28ad4-112">作業ウィンドウまたはカスタム関数から認証する必要がある場合は、アクセストークンが既に取得されているかどうか `storage` を確認します。</span><span class="sxs-lookup"><span data-stu-id="28ad4-112">When you need to authenticate either from the task pane or a custom function, check `storage` to see if the access token was already acquired.</span></span> <span data-ttu-id="28ad4-113">取得されていない場合は、ダイアログ API を使用してユーザーを認証し、アクセストークンを取得して、後で使用するために `storage` に保存します。</span><span class="sxs-lookup"><span data-stu-id="28ad4-113">If not, use the dialog API to authenticate the user, retrieve the access token, and then store the token in `storage` for future use.</span></span>

## <a name="dialog-api"></a><span data-ttu-id="28ad4-114">ダイアログ API</span><span class="sxs-lookup"><span data-stu-id="28ad4-114">Dialog API example</span></span>

<span data-ttu-id="28ad4-115">トークンが存在しない場合は、ユーザーにサインインを求めるダイアログ API を表示する必要があります。</span><span class="sxs-lookup"><span data-stu-id="28ad4-115">If a token doesn't exist, you should use the Dialog API to ask the user to sign in.</span></span> <span data-ttu-id="28ad4-116">ユーザーが資格情報を入力すると、結果のアクセストークンが `storage` に保存されます。</span><span class="sxs-lookup"><span data-stu-id="28ad4-116">After a user enters their credentials, the resulting access token can be stored in `storage`.</span></span>

> [!NOTE]
> <span data-ttu-id="28ad4-117">カスタム関数のランタイムは、作業ウィンドウで使用されるブラウザー エンジン ランタイムのダイアログ オブジェクトとは少し異なるダイアログ オブジェクトを使用します。</span><span class="sxs-lookup"><span data-stu-id="28ad4-117">The custom functions runtime uses a Dialog object that is slightly different from the Dialog object in the browser engine runtime used by task panes.</span></span> <span data-ttu-id="28ad4-118">いずれも "ダイアログ API" と呼ばれていますが、カスタム関数のランタイムでユーザーを認証するために `OfficeRuntime.Dialog` を使用します。</span><span class="sxs-lookup"><span data-stu-id="28ad4-118">They're both referred to as the "Dialog API", but use `OfficeRuntime.Dialog` to authenticate users in the custom functions runtime.</span></span>

<span data-ttu-id="28ad4-119">`Dialog` オブジェクトを使用する方法の詳細については、「[カスタム関数ダイアログ](/office/dev/add-ins/excel/custom-functions-dialog)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="28ad4-119">For information on how to use the `Dialog` object, see [Custom Functions dialog](/office/dev/add-ins/excel/custom-functions-dialog).</span></span>

<span data-ttu-id="28ad4-120">認証プロセス全体を構想するときには、アドインの作業ウィンドウと UI 要素、アドインのカスタム関数部分が、`OfficeRuntime.storage` を通じて相互に通信できる個別のエンティティとして考えてみることをおすすめします。</span><span class="sxs-lookup"><span data-stu-id="28ad4-120">When envisioning the entire authentication process as a whole, it might be helpful to think of the task pane and UI elements of your add-in and the custom functions part of your add-in as separate entities which can communicate with each other through `OfficeRuntime.storage`.</span></span>

<span data-ttu-id="28ad4-121">この基本的な手順を次の図に示します。</span><span class="sxs-lookup"><span data-stu-id="28ad4-121">The following diagram outlines this basic process.</span></span> <span data-ttu-id="28ad4-122">点線は、ユーザーが個別の操作を実行している間に、カスタム関数とアドインの作業ウィンドウがアドインの一部であることを示しています。</span><span class="sxs-lookup"><span data-stu-id="28ad4-122">Note that the dotted line indicates that while they perform separate actions, custom functions and your add-in's task pane are both part of your add-in as a whole.</span></span>

1. <span data-ttu-id="28ad4-123">Excel ワークブックのセルからカスタム関数を発行します。</span><span class="sxs-lookup"><span data-stu-id="28ad4-123">You issue a custom function call from a cell in an Excel workbook.</span></span>
2. <span data-ttu-id="28ad4-124">カスタム関数は、ユーザーの資格情報を web サイトに渡すために `Dialog` を使用します。</span><span class="sxs-lookup"><span data-stu-id="28ad4-124">The custom function uses `Dialog` to pass your user credentials to a website.</span></span>
3. <span data-ttu-id="28ad4-125">その後、web サイトは、アクセストークンをカスタム関数に返します。</span><span class="sxs-lookup"><span data-stu-id="28ad4-125">This website then returns an access token to the custom function.</span></span>
4. <span data-ttu-id="28ad4-126">このアクセストークンは、カスタム関数によって `storage` に設定されます。</span><span class="sxs-lookup"><span data-stu-id="28ad4-126">Your custom function then sets this access token to the `storage`.</span></span>
5. <span data-ttu-id="28ad4-127">アドインの作業ウィンドウは、`storage` からトークンにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="28ad4-127">Your add-in's task pane accesses the token from `storage`.</span></span>

<span data-ttu-id="28ad4-128">![アクセス トークンを取得するためのダイアログ API を使用したカスタム関数の図と、OfficeRuntime.storage API を通してトークンを作業ウィンドウと共有します。](../images/authentication-diagram.png "認証図。")</span><span class="sxs-lookup"><span data-stu-id="28ad4-128">![Diagram of custom function using dialog API to get access token, and then share token with task pane through the OfficeRuntime.storage API.](../images/authentication-diagram.png "Authentication diagram.")</span></span>

## <a name="storing-the-token"></a><span data-ttu-id="28ad4-129">トークンの格納</span><span class="sxs-lookup"><span data-stu-id="28ad4-129">Storing the token</span></span>

<span data-ttu-id="28ad4-130">次の例は、[カスタム関数の OfficeRuntime.storage を使用 ](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage)したコードサンプルです。</span><span class="sxs-lookup"><span data-stu-id="28ad4-130">The following examples are from the [Using OfficeRuntime.storage in custom functions](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) code sample.</span></span> <span data-ttu-id="28ad4-131">カスタム関数と作業ウィンドウ間のデータ共有の例については、このコードサンプルを参照してください。</span><span class="sxs-lookup"><span data-stu-id="28ad4-131">Refer to this code sample for a complete example of sharing data between custom functions and the task pane.</span></span>

<span data-ttu-id="28ad4-132">カスタム関数が認証されたら、アクセストークンを受け取り、`storage`に保存する必要があります。</span><span class="sxs-lookup"><span data-stu-id="28ad4-132">If the custom function authenticates, then it receives the access token and will need to store it in `storage`.</span></span> <span data-ttu-id="28ad4-133">次のコードサンプルは、`storage.setItem`メソッドを呼び出して値を格納する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="28ad4-133">The following code sample shows how to call the `storage.setItem` method to store a value.</span></span> <span data-ttu-id="28ad4-134">`storeValue` 関数は、ユーザーからの値を格納するためのカスタム関数です。</span><span class="sxs-lookup"><span data-stu-id="28ad4-134">The `storeValue` function is a custom function that for example purposes stores a value from the user.</span></span> <span data-ttu-id="28ad4-135">必要なトークン値を格納するように変更できます。</span><span class="sxs-lookup"><span data-stu-id="28ad4-135">You can modify this to store any token value you need.</span></span>

```js
/**
 * Stores a key-value pair into OfficeRuntime.storage.
 * @customfunction
 * @param {string} key Key of item to put into storage.
 * @param {*} value Value of item to put into storage.
 */
function storeValue(key, value) {
  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

<span data-ttu-id="28ad4-136">作業ウィンドウにアクセストークンが必要な場合は、`storage`から トークンを取得できます。</span><span class="sxs-lookup"><span data-stu-id="28ad4-136">When the task pane needs the access token, it can retrieve the token from `storage`.</span></span> <span data-ttu-id="28ad4-137">次のコードサンプルは、`storage.getItem`メソッドを使用してトークンを取得する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="28ad4-137">The following code sample shows how to use the `storage.getItem` method to retrieve the token.</span></span>

```js
/**
 * Read a token from storage.
 * @customfunction GETTOKEN
 */
function receiveTokenFromCustomFunction() {
  var key = "token";
  var tokenSendStatus = document.getElementById('tokenSendStatus');
  OfficeRuntime.storage.getItem(key).then(function (result) {
     tokenSendStatus.value = "Success: Item with key '" + key + "' read from storage.";
     document.getElementById('tokenTextBox2').value = result;
  }, function (error) {
     tokenSendStatus.value = "Error: Unable to read item with key '" + key + "' from storage. " + error;
  });
}
```

## <a name="general-guidance"></a><span data-ttu-id="28ad4-138">一般的なガイダンス</span><span class="sxs-lookup"><span data-stu-id="28ad4-138">General Guidance</span></span>

<span data-ttu-id="28ad4-139">Office アドインは web ベースで、あらゆる web 認証技術を使用できます。</span><span class="sxs-lookup"><span data-stu-id="28ad4-139">Office Add-ins are web-based and you can use any web authentication technique.</span></span> <span data-ttu-id="28ad4-140">カスタム関数を使用して独自の認証を実装するのに、特定のパターンやメソッドはありません。</span><span class="sxs-lookup"><span data-stu-id="28ad4-140">There is no particular pattern or method you must follow to implement your own authentication with custom functions.</span></span> <span data-ttu-id="28ad4-141">さまざまな認証パターンに関するドキュメントを参照してください。 [この記事では、外部サービスによる認証について説明します。](/office/dev/add-ins/develop/auth-external-add-ins)</span><span class="sxs-lookup"><span data-stu-id="28ad4-141">You may wish to consult the documentation about various authentication patterns, starting with [this article about authorizing via external services](/office/dev/add-ins/develop/auth-external-add-ins).</span></span>  

<span data-ttu-id="28ad4-142">カスタム関数を開発するときに、次の場所にデータを格納しないようにします。</span><span class="sxs-lookup"><span data-stu-id="28ad4-142">Avoid using the following locations to store data when developing custom functions:</span></span>  

- <span data-ttu-id="28ad4-143">`localStorage`: カスタム関数はグローバル `window` オブジェクトへのアクセス権がないため、`localStorage` に保存されているデータにはアクセスできません。</span><span class="sxs-lookup"><span data-stu-id="28ad4-143">`localStorage`: Custom functions do not have access to the global `window` object and therefore have no access to data stored in `localStorage`.</span></span>
- <span data-ttu-id="28ad4-144">`Office.context.document.settings`: この場所は安全ではないため、アドインを使用しているユーザーが情報を抽出できます。</span><span class="sxs-lookup"><span data-stu-id="28ad4-144">`Office.context.document.settings`:  This location is not secure and information can be extracted by anyone using the add-in.</span></span>

## <a name="next-steps"></a><span data-ttu-id="28ad4-145">次の手順</span><span class="sxs-lookup"><span data-stu-id="28ad4-145">Next steps</span></span>
<span data-ttu-id="28ad4-146">[カスタム関数のダイアログ API ](custom-functions-dialog.md) について説明します。</span><span class="sxs-lookup"><span data-stu-id="28ad4-146">Learn about the [dialog API for custom functions](custom-functions-dialog.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="28ad4-147">関連項目</span><span class="sxs-lookup"><span data-stu-id="28ad4-147">See also</span></span>

* [<span data-ttu-id="28ad4-148">カスタム関数のアーキテクチャ</span><span class="sxs-lookup"><span data-stu-id="28ad4-148">Custom functions architecture</span></span>](custom-functions-architecture.md)
* [<span data-ttu-id="28ad4-149">カスタム関数でデータを受信して​​処理する</span><span class="sxs-lookup"><span data-stu-id="28ad4-149">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="28ad4-150">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="28ad4-150">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="28ad4-151">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="28ad4-151">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)

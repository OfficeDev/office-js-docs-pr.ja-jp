---
ms.date: 05/17/2020
description: 作業ウィンドウを使用しない Excel でカスタム関数を使用してユーザーを認証します。
title: UI レスのカスタム関数の認証
localization_priority: Normal
ms.openlocfilehash: b4ff234f71ed2a36cc311e45f47498d19380b862
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609339"
---
# <a name="authentication-for-ui-less-custom-functions"></a><span data-ttu-id="e2dc8-103">UI レスのカスタム関数の認証</span><span class="sxs-lookup"><span data-stu-id="e2dc8-103">Authentication for UI-less custom functions</span></span>

<span data-ttu-id="e2dc8-104">一部のシナリオでは、作業ウィンドウやその他のユーザーインターフェイス要素を使用しないカスタム関数 (UI レスカスタム関数) は、保護されたリソースにアクセスするためにユーザーを認証する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-104">In some scenarios your custom function that does not use a task pane or other user interface elements (UI-less custom function) will need to authenticate the user in order to access protected resources.</span></span> <span data-ttu-id="e2dc8-105">UI を使用しないカスタム関数は、JavaScript のみのランタイムで実行されることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-105">Be aware that UI-less custom functions run in a JavaScript-only runtime.</span></span> <span data-ttu-id="e2dc8-106">そのため、JavaScript のみのランタイムと、オブジェクトとダイアログ API を使用してほとんどのアドインで使用される一般的なブラウザーエンジンランタイムとの間でデータをやり取りする必要があり `OfficeRuntime.storage` ます。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-106">Because of this, you'll need to pass data back and forth between the JavaScript-only runtime and the typical browser engine runtime used by most add-ins using the `OfficeRuntime.storage` object and the Dialog API.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

## <a name="officeruntimestorage-object"></a><span data-ttu-id="e2dc8-107">OfficeRuntime.storage オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e2dc8-107">OfficeRuntime.storage object</span></span>

<span data-ttu-id="e2dc8-108">UI に含まれないカスタム関数で使用される JavaScript 専用のランタイムには、 `localStorage` 通常、データを格納するグローバルウィンドウで使用できるオブジェクトがありません。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-108">The JavaScript-only runtime used by UI-less custom functions doesn't have a `localStorage` object available on the global window, where you typically store data.</span></span> <span data-ttu-id="e2dc8-109">その代わりに、" [Officeruntime](/javascript/api/office-runtime/officeruntime.storage) " を使用して UI レスのカスタム関数と作業ウィンドウ間でデータを共有する必要があります。データを設定および取得するためのストレージです。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-109">Instead, you should share data between UI-less custom functions and task panes by using [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) to set and get data.</span></span>

### <a name="suggested-usage"></a><span data-ttu-id="e2dc8-110">おすすめの使用法</span><span class="sxs-lookup"><span data-stu-id="e2dc8-110">Suggested usage</span></span>

<span data-ttu-id="e2dc8-111">UI を使用しないカスタム関数から認証する必要がある場合は、 `storage` アクセストークンが既に取得されているかどうかを確認します。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-111">When you need to authenticate from a UI-less custom function, check `storage` to see if the access token was already acquired.</span></span> <span data-ttu-id="e2dc8-112">取得されていない場合は、ダイアログ API を使用してユーザーを認証し、アクセストークンを取得して、後で使用するために `storage` に保存します。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-112">If not, use the dialog API to authenticate the user, retrieve the access token, and then store the token in `storage` for future use.</span></span>

## <a name="dialog-api"></a><span data-ttu-id="e2dc8-113">ダイアログ API</span><span class="sxs-lookup"><span data-stu-id="e2dc8-113">Dialog API</span></span>

<span data-ttu-id="e2dc8-114">トークンが存在しない場合は、ユーザーにサインインを求めるダイアログ API を表示する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-114">If a token doesn't exist, you should use the Dialog API to ask the user to sign in.</span></span> <span data-ttu-id="e2dc8-115">ユーザーが資格情報を入力すると、結果のアクセストークンが `storage` に保存されます。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-115">After a user enters their credentials, the resulting access token can be stored in `storage`.</span></span>

> [!NOTE]
> <span data-ttu-id="e2dc8-116">JavaScript のみのランタイムは、作業ウィンドウで使用されるブラウザーエンジンランタイムの Dialog オブジェクトとは少し異なるダイアログオブジェクトを使用します。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-116">The JavaScript-only runtime uses a Dialog object that is slightly different from the Dialog object in the browser engine runtime used by task panes.</span></span> <span data-ttu-id="e2dc8-117">これらはどちらも "Dialog API" と呼ばれていますが、 `OfficeRuntime.Dialog` JavaScript のみのランタイムでユーザーを認証するために使用します。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-117">They're both referred to as the "Dialog API", but use `OfficeRuntime.Dialog` to authenticate users in the JavaScript-only runtime.</span></span>

<span data-ttu-id="e2dc8-118">この基本的な手順を次の図に示します。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-118">The following diagram outlines this basic process.</span></span> <span data-ttu-id="e2dc8-119">点線は、UI を使用しないカスタム関数とアドインの作業ウィンドウがどちらもアドインの一部であることを示していますが、個別のランタイムを使用しています。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-119">The dotted line indicates that UI-less custom functions and your add-in's task pane are both part of your add-in as a whole, though they use separate runtimes.</span></span>

1. <span data-ttu-id="e2dc8-120">Excel ブックのセルから UI を使用しないカスタム関数呼び出しを発行します。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-120">You issue a UI-less custom function call from a cell in an Excel workbook.</span></span>
2. <span data-ttu-id="e2dc8-121">UI を使用しないカスタム関数は、 `Dialog` ユーザーの資格情報を web サイトに渡すために使用します。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-121">The UI-less custom function uses `Dialog` to pass your user credentials to a website.</span></span>
3. <span data-ttu-id="e2dc8-122">この web サイトは、UI なしのカスタム関数へのアクセストークンを返します。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-122">This website then returns an access token to the UI-less custom function.</span></span>
4. <span data-ttu-id="e2dc8-123">UI を使用しないカスタム関数は、このアクセストークンをに設定し `storage` ます。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-123">Your UI-less custom function then sets this access token to the `storage`.</span></span>
5. <span data-ttu-id="e2dc8-124">アドインの作業ウィンドウは、`storage` からトークンにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-124">Your add-in's task pane accesses the token from `storage`.</span></span>

<span data-ttu-id="e2dc8-125">![アクセストークンを取得するためにダイアログ API を使用したカスタム関数の図。次に、この方法で、作業ウィンドウを使用してトークンを保存します。ストレージ API を使用します。](../images/authentication-diagram.png "認証の図。")</span><span class="sxs-lookup"><span data-stu-id="e2dc8-125">![Diagram of custom function using dialog API to get access token, and then share token with task pane through the OfficeRuntime.storage API.](../images/authentication-diagram.png "Authentication diagram.")</span></span>

## <a name="storing-the-token"></a><span data-ttu-id="e2dc8-126">トークンの格納</span><span class="sxs-lookup"><span data-stu-id="e2dc8-126">Storing the token</span></span>

<span data-ttu-id="e2dc8-127">次の例は、[カスタム関数の OfficeRuntime.storage を使用 ](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage)したコードサンプルです。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-127">The following examples are from the [Using OfficeRuntime.storage in custom functions](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) code sample.</span></span> <span data-ttu-id="e2dc8-128">UI を使用しないカスタム関数と作業ウィンドウ間でデータを共有する完全な例については、以下のコードサンプルを参照してください。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-128">Refer to this code sample for a complete example of sharing data between UI-less custom functions and the task pane.</span></span>

<span data-ttu-id="e2dc8-129">UI を省略したカスタム関数が認証された場合は、アクセストークンを受け取り、それをに格納する必要があり `storage` ます。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-129">If the UI-less custom function authenticates, then it receives the access token and will need to store it in `storage`.</span></span> <span data-ttu-id="e2dc8-130">次のコードサンプルは、`storage.setItem`メソッドを呼び出して値を格納する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-130">The following code sample shows how to call the `storage.setItem` method to store a value.</span></span> <span data-ttu-id="e2dc8-131">この `storeValue` 関数は、ユーザーの値を格納するなど、UI を使用しないカスタム関数です。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-131">The `storeValue` function is a UI-less custom function that for example purposes stores a value from the user.</span></span> <span data-ttu-id="e2dc8-132">必要なトークン値を格納するように変更できます。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-132">You can modify this to store any token value you need.</span></span>

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

<span data-ttu-id="e2dc8-133">作業ウィンドウにアクセストークンが必要な場合は、`storage`から トークンを取得できます。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-133">When the task pane needs the access token, it can retrieve the token from `storage`.</span></span> <span data-ttu-id="e2dc8-134">次のコードサンプルは、`storage.getItem`メソッドを使用してトークンを取得する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-134">The following code sample shows how to use the `storage.getItem` method to retrieve the token.</span></span>

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

## <a name="general-guidance"></a><span data-ttu-id="e2dc8-135">一般的なガイダンス</span><span class="sxs-lookup"><span data-stu-id="e2dc8-135">General guidance</span></span>

<span data-ttu-id="e2dc8-136">Office アドインは web ベースで、あらゆる web 認証技術を使用できます。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-136">Office Add-ins are web-based and you can use any web authentication technique.</span></span> <span data-ttu-id="e2dc8-137">UI を使用しないカスタム関数を使用して独自の認証を実装するために従う必要のある特定のパターンやメソッドはありません。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-137">There is no particular pattern or method you must follow to implement your own authentication with UI-less custom functions.</span></span> <span data-ttu-id="e2dc8-138">さまざまな認証パターンに関するドキュメントを参照してください。 [この記事では、外部サービスによる認証について説明します。](../develop/auth-external-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="e2dc8-138">You may wish to consult the documentation about various authentication patterns, starting with [this article about authorizing via external services](../develop/auth-external-add-ins.md).</span></span>  

<span data-ttu-id="e2dc8-139">カスタム関数を開発するときに、次の場所にデータを格納しないようにします。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-139">Avoid using the following locations to store data when developing custom functions:</span></span>  

- <span data-ttu-id="e2dc8-140">`localStorage`: UI を持たないカスタム関数は、グローバルオブジェクトにアクセスできない `window` ため、に格納されているデータにはアクセスできません `localStorage` 。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-140">`localStorage`: UI-less custom functions do not have access to the global `window` object and therefore have no access to data stored in `localStorage`.</span></span>
- <span data-ttu-id="e2dc8-141">`Office.context.document.settings`: この場所は安全ではないため、アドインを使用しているユーザーが情報を抽出できます。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-141">`Office.context.document.settings`:  This location is not secure and information can be extracted by anyone using the add-in.</span></span>

## <a name="dialog-box-api-example"></a><span data-ttu-id="e2dc8-142">ダイアログボックス API の例</span><span class="sxs-lookup"><span data-stu-id="e2dc8-142">Dialog box API example</span></span>

<span data-ttu-id="e2dc8-143">次のコードサンプルでは、関数は `getTokenViaDialog` API の関数を使用して `Dialog` `displayWebDialogOptions` ダイアログボックスを表示します。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-143">In the following code sample, the function `getTokenViaDialog` uses the `Dialog` API's `displayWebDialogOptions` function to display a dialog box.</span></span> <span data-ttu-id="e2dc8-144">このサンプルは、オブジェクトの機能を示すために提供されており `Dialog` 、認証方法を示すものではありません。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-144">This sample is provided to show the capabilities of the `Dialog` object, not demonstrate how to authenticate.</span></span>

```JavaScript
/**
 * Function retrieves a cached token or opens a dialog box if there is no saved token. Note that this is not a sufficient example of authentication but is intended to show the capabilities of the Dialog object.
 * @param {string} url URL for a stored token.
 */
function getTokenViaDialog(url) {
  return new Promise (function (resolve, reject) {
    if (_dialogOpen) {
      // Can only have one dialog box open at once. Wait for previous dialog box's token.
      let timeout = 5;
      let count = 0;
      var intervalId = setInterval(function () {
        count++;
        if(_cachedToken) {
          resolve(_cachedToken);
          clearInterval(intervalId);
        }
        if(count >= timeout) {
          reject("Timeout while waiting for token");
          clearInterval(intervalId);
        }
      }, 1000);
    } else {
      _dialogOpen = true;
      OfficeRuntime.displayWebDialog(url, {
        height: '50%',
        width: '50%',
        onMessage: function (message, dialog) {
          _cachedToken = message;
          resolve(message);
          dialog.close();
          return;
        },
        onRuntimeError: function(error, dialog) {
          reject(error);
        },
      }).catch(function (e) {
        reject(e);
      });
    }
  });
}
```

## <a name="next-steps"></a><span data-ttu-id="e2dc8-145">次の手順</span><span class="sxs-lookup"><span data-stu-id="e2dc8-145">Next steps</span></span>
<span data-ttu-id="e2dc8-146">[UI のないカスタム関数をデバッグ](custom-functions-debugging.md)する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="e2dc8-146">Learn how to [debug UI-less custom functions](custom-functions-debugging.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="e2dc8-147">関連項目</span><span class="sxs-lookup"><span data-stu-id="e2dc8-147">See also</span></span>

* [<span data-ttu-id="e2dc8-148">UI レス Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="e2dc8-148">Runtime for UI-less Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="e2dc8-149">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="e2dc8-149">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)

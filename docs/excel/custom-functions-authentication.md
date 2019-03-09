---
ms.date: 03/06/2019
description: Excel でカスタム関数を使用してユーザーを認証します。
title: カスタム関数の認証
ms.openlocfilehash: 4358d9f570ef8b31db98b1886c01ff4a89a6b1be
ms.sourcegitcommit: 8e7b7b0cfb68b91a3a95585d094cf5f5ffd00178
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/09/2019
ms.locfileid: "30512854"
---
# <a name="authentication"></a><span data-ttu-id="48fa6-103">認証</span><span class="sxs-lookup"><span data-stu-id="48fa6-103">Authentication</span></span>

<span data-ttu-id="48fa6-104">一部のシナリオでは、カスタム関数は、保護されたリソースにアクセスするためにユーザーを認証する必要があります。</span><span class="sxs-lookup"><span data-stu-id="48fa6-104">In some scenarios your custom function will need to authenticate the user in order to access protected resources.</span></span> <span data-ttu-id="48fa6-105">カスタム関数では、特定の認証方法を使用する必要はありませんが、カスタム関数は、アドインの作業ウィンドウや他の UI 要素とは別のランタイムで実行されることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="48fa6-105">While custom functions don't require a specific method of authentication, you should be aware that custom functions run in a separate runtime from the task pane and other UI elements of your add-in.</span></span> <span data-ttu-id="48fa6-106">そのため、 `AsyncStorage`オブジェクトとダイアログ API を使用して、2つのランタイム間でデータをやり取りする必要があります。</span><span class="sxs-lookup"><span data-stu-id="48fa6-106">Because of this, you'll need to pass data back and forth between the two runtimes using the `AsyncStorage` object and the Dialog API.</span></span>
  
## <a name="asyncstorage-object"></a><span data-ttu-id="48fa6-107">asyncstorage オブジェクト</span><span class="sxs-lookup"><span data-stu-id="48fa6-107">AsyncStorage object</span></span>

<span data-ttu-id="48fa6-108">カスタム関数ランタイムには、通常`localStorage` 、データを格納するグローバルウィンドウで使用できるオブジェクトがありません。</span><span class="sxs-lookup"><span data-stu-id="48fa6-108">The custom functions runtime doesn't have a `localStorage` object available on the global window, where you might typically store data.</span></span> <span data-ttu-id="48fa6-109">代わりに、データを設定して取得するために、 [officeruntime](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.asyncstorage)を使用して、カスタム関数と作業ウィンドウ間でデータを共有する必要があります。</span><span class="sxs-lookup"><span data-stu-id="48fa6-109">Instead, you should share data between custom functions and task panes by using [OfficeRuntime.AsyncStorage](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.asyncstorage) to set and get data.</span></span>

<span data-ttu-id="48fa6-110">さらに、を使用`AsyncStorage`するメリットがあります。セキュリティで保護されたサンドボックス環境を使用して、他のアドインがデータにアクセスできないようにします。</span><span class="sxs-lookup"><span data-stu-id="48fa6-110">Additionally, there is a benefit to using `AsyncStorage`; it uses a secure sandbox environment so that your data cannot be accessed by other add-ins.</span></span>

### <a name="suggested-usage"></a><span data-ttu-id="48fa6-111">推奨される使用法</span><span class="sxs-lookup"><span data-stu-id="48fa6-111">Suggested usage</span></span>

<span data-ttu-id="48fa6-112">作業ウィンドウまたはカスタム関数から認証を受ける必要がある場合は、 `AsyncStorage`アクセストークンが既に取得されているかどうかを確認します。</span><span class="sxs-lookup"><span data-stu-id="48fa6-112">When you need to authenticate either from the task pane or a custom function, check `AsyncStorage` to see if the access token was already acquired.</span></span> <span data-ttu-id="48fa6-113">表示されない場合は、ダイアログ API を使用してユーザーを認証し、アクセストークンを取得し`AsyncStorage`てから、トークンを後で使用するために格納します。</span><span class="sxs-lookup"><span data-stu-id="48fa6-113">If not, use the dialog API to authenticate the user, retrieve the access token, and then store the token in `AsyncStorage` for future use.</span></span>

## <a name="dialog-api"></a><span data-ttu-id="48fa6-114">ダイアログ API</span><span class="sxs-lookup"><span data-stu-id="48fa6-114">Dialog API</span></span>

<span data-ttu-id="48fa6-115">トークンが存在しない場合は、ダイアログ API を使用して、ユーザーにサインインを要求する必要があります。</span><span class="sxs-lookup"><span data-stu-id="48fa6-115">If a token doesn't exist, you should use the Dialog API to ask the user to sign in.</span></span> <span data-ttu-id="48fa6-116">ユーザーが資格情報を入力すると、作成されたアクセストークン`AsyncStorage`がに保存されます。</span><span class="sxs-lookup"><span data-stu-id="48fa6-116">After a user enters their credentials, the resulting access token can be stored in `AsyncStorage`.</span></span>

> [!NOTE]
> <span data-ttu-id="48fa6-117">カスタム関数ランタイムは、作業ウィンドウで使用されるブラウザーエンジンランタイムの dialog オブジェクトとは少し異なるダイアログオブジェクトを使用します。</span><span class="sxs-lookup"><span data-stu-id="48fa6-117">The custom functions runtime uses a Dialog object that is slightly different from the Dialog object in the browser engine runtime used by task panes.</span></span> <span data-ttu-id="48fa6-118">これらはどちらも "Dialog API" と呼ばれています`Officeruntime.Dialog`が、カスタム関数ランタイムでユーザーを認証するために使用します。</span><span class="sxs-lookup"><span data-stu-id="48fa6-118">They're both referred to as the "Dialog API", but use `Officeruntime.Dialog` to authenticate users in the custom functions runtime.</span></span>

<span data-ttu-id="48fa6-119">の`OfficeRuntime.Dialog`使用方法については、「[カスタム関数ランタイム](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime?view=office-js#displaying-a-dialog-box)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="48fa6-119">For information on how to use the `OfficeRuntime.Dialog`, see [Custom Functions runtime](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime?view=office-js#displaying-a-dialog-box).</span></span>

<span data-ttu-id="48fa6-120">全体として認証プロセス全体を構想する場合は、アドインの作業ウィンドウと UI 要素、およびアドインのカスタム関数の部分を、を通じて`AsyncStorage`相互に通信できる個別のエンティティと考えることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="48fa6-120">When envisioning the entire authentication process as a whole, it might be helpful to think of the task pane and UI elements of your add-in and the custom functions part of your add-in as separate entities which can communicate with each other through `AsyncStorage`.</span></span>

<span data-ttu-id="48fa6-121">次の図は、この基本的なプロセスの概要を示しています。</span><span class="sxs-lookup"><span data-stu-id="48fa6-121">The following diagram outlines this basic process.</span></span> <span data-ttu-id="48fa6-122">点線では、個別の操作を実行する一方で、カスタム関数とアドインの作業ウィンドウは、どちらもアドインの一部であることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="48fa6-122">Note that the dotted line indicates that while they perform separate actions, custom functions and your add-in's task pane are both part of your add-in as a whole.</span></span>

1. <span data-ttu-id="48fa6-123">Excel ブックのセルからカスタム関数呼び出しを発行します。</span><span class="sxs-lookup"><span data-stu-id="48fa6-123">You issue a custom function call from a cell in an Excel workbook.</span></span>
2. <span data-ttu-id="48fa6-124">カスタム関数を使用`Officeruntime.Dialog`して、ユーザーの資格情報を web サイトに渡します。</span><span class="sxs-lookup"><span data-stu-id="48fa6-124">The custom function uses `Officeruntime.Dialog` to pass your user credentials to a website.</span></span>
3. <span data-ttu-id="48fa6-125">その後、この web サイトは、カスタム関数へのアクセストークンを返します。</span><span class="sxs-lookup"><span data-stu-id="48fa6-125">This website then returns an access token to the custom function.</span></span>
4. <span data-ttu-id="48fa6-126">次に、カスタム関数は、 `AsyncStorage`このアクセストークンをに設定します。</span><span class="sxs-lookup"><span data-stu-id="48fa6-126">Your custom function then sets this access token to the `AsyncStorage`.</span></span>
5. <span data-ttu-id="48fa6-127">アドインの作業ウィンドウで、から`AsyncStorage`トークンにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="48fa6-127">Your add-in's task pane accesses the token from `AsyncStorage`.</span></span>

<span data-ttu-id="48fa6-128">![カスタム関数、officeruntime、および共同作業ウィンドウの図](../images/Authdiagram.png "認証の図。")</span><span class="sxs-lookup"><span data-stu-id="48fa6-128">![Diagram of custom functions, OfficeRuntime, and task panes working together.](../images/Authdiagram.png "Authentication diagram.")</span></span>

## <a name="storing-the-token"></a><span data-ttu-id="48fa6-129">トークンの保存</span><span class="sxs-lookup"><span data-stu-id="48fa6-129">Storing the token</span></span>

<span data-ttu-id="48fa6-130">次の例は、 [「カスタム関数の asyncstorage を使用する」](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage)のコードサンプルのものです。</span><span class="sxs-lookup"><span data-stu-id="48fa6-130">The following examples are from the [Using AsyncStorage in custom functions](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) code sample.</span></span> <span data-ttu-id="48fa6-131">カスタム関数と作業ウィンドウとの間でデータを共有する完全な例については、次のコードサンプルを参照してください。</span><span class="sxs-lookup"><span data-stu-id="48fa6-131">Refer to this code sample for a complete example of sharing data between custom functions and the task pane.</span></span>

<span data-ttu-id="48fa6-132">カスタム関数が認証された場合は、アクセストークンを受け取り、それをに`AsyncStorage`格納する必要があります。</span><span class="sxs-lookup"><span data-stu-id="48fa6-132">If the custom function authenticates, then it receives the access token and will need to store it in `AsyncStorage`.</span></span> <span data-ttu-id="48fa6-133">次のコードサンプルは、メソッドを呼び出し`AsyncStorage.setItem`て値を格納する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="48fa6-133">The following code sample shows how to call the `AsyncStorage.setItem` method to store a value.</span></span> <span data-ttu-id="48fa6-134">この`StoreValue`関数は、たとえば、ユーザーの値を格納するためのカスタム関数です。</span><span class="sxs-lookup"><span data-stu-id="48fa6-134">The `StoreValue` function is a custom function that for example purposes stores a value from the user.</span></span> <span data-ttu-id="48fa6-135">必要なトークン値を格納するように変更することができます。</span><span class="sxs-lookup"><span data-stu-id="48fa6-135">You can modify this to store any token value you need.</span></span>

```javascript
function StoreValue(key, value) {
  return OfficeRuntime.AsyncStorage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to AsyncStorage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to AsyncStorage. " + error;
  });
}
```

<span data-ttu-id="48fa6-136">作業ウィンドウでアクセストークンが必要になると、そのトークンを取得`AsyncStorage`することができます。</span><span class="sxs-lookup"><span data-stu-id="48fa6-136">When the task pane needs the access token, it can retrieve the token from `AsyncStorage`.</span></span> <span data-ttu-id="48fa6-137">次のコードサンプルは、 `AsyncStorage.getItem`メソッドを使用してトークンを取得する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="48fa6-137">The following code sample shows how to use the `AsyncStorage.getItem` method to retrieve the token.</span></span>

```javascript
function ReceiveTokenFromCustomFunction() {
   var key = "token";
   var tokenSendStatus = document.getElementById('tokenSendStatus');
   OfficeRuntime.AsyncStorage.getItem(key).then(function (result) {
      tokenSendStatus.value = "Success: Item with key '" + key + "' read from AsyncStorage.";
      document.getElementById('tokenTextBox2').value = result;
   }, function (error) {
      tokenSendStatus.value = "Error: Unable to read item with key '" + key + "' from AsyncStorage. " + error;
   });
}
```

## <a name="general-guidance"></a><span data-ttu-id="48fa6-138">一般的なガイダンス</span><span class="sxs-lookup"><span data-stu-id="48fa6-138">General guidance</span></span>

<span data-ttu-id="48fa6-139">Office アドインは web ベースであり、任意の web 認証方法を使用できます。</span><span class="sxs-lookup"><span data-stu-id="48fa6-139">Office Add-ins are web-based and you can use any web authentication technique.</span></span> <span data-ttu-id="48fa6-140">カスタム関数を使用して独自の認証を実装するために従う必要のある特定のパターンやメソッドはありません。</span><span class="sxs-lookup"><span data-stu-id="48fa6-140">There is no particular pattern or method you must follow to implement your own authentication with custom functions.</span></span> <span data-ttu-id="48fa6-141">[外部サービスによる承認については、この記事](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins?view=office-js)から始まるさまざまな認証パターンに関するドキュメントを参照してください。</span><span class="sxs-lookup"><span data-stu-id="48fa6-141">You may wish to consult the documentation about various authentication patterns, starting with [this article about authorizing via external services](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins?view=office-js).</span></span>  

<span data-ttu-id="48fa6-142">カスタム関数を開発する際に、次の場所を使用してデータを保存しないようにします。</span><span class="sxs-lookup"><span data-stu-id="48fa6-142">Avoid using the following locations to store data when developing custom functions:</span></span>  

- <span data-ttu-id="48fa6-143">`localStorage`: カスタム関数にはグローバル`window`オブジェクトへのアクセス権がないため、に`localStorage`格納されているデータにアクセスできません。</span><span class="sxs-lookup"><span data-stu-id="48fa6-143">`localStorage`: Custom functions do not have access to the global `window` object and therefore have no access to data     stored in `localStorage`.</span></span>
- <span data-ttu-id="48fa6-144">`Office.context.document.settings`: この場所はセキュリティで保護されていないため、アドインを使用するすべてのユーザーが情報を抽出できます。</span><span class="sxs-lookup"><span data-stu-id="48fa6-144">`Office.context.document.settings`:  This location is not secure and information can be extracted by anyone using the     add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="48fa6-145">関連項目</span><span class="sxs-lookup"><span data-stu-id="48fa6-145">See also</span></span>

* [<span data-ttu-id="48fa6-146">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="48fa6-146">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="48fa6-147">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="48fa6-147">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="48fa6-148">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="48fa6-148">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="48fa6-149">チュートリアル: Excel でカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="48fa6-149">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)

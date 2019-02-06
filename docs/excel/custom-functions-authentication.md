---
ms.date: 1/29/2019
description: Excel でカスタム関数を使用してユーザーを認証します。
title: ユーザー定義関数での認証
ms.openlocfilehash: 0e42dbc93cb545660a8dbaae5bdb48724f3b7376
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/05/2019
ms.locfileid: "29745418"
---
# <a name="authentication"></a><span data-ttu-id="51f83-103">認証</span><span class="sxs-lookup"><span data-stu-id="51f83-103">Authentication</span></span>

<span data-ttu-id="51f83-104">保護されたリソースを場合によっては、ユーザー定義関数にアクセスするためにユーザーを認証する必要があります。</span><span class="sxs-lookup"><span data-stu-id="51f83-104">In some scenarios your custom function will need to authenticate the user in order to access protected resources.</span></span> <span data-ttu-id="51f83-105">カスタム関数は、特定の認証方法を必要としない、作業ウィンドウと、アドインの場合は、他の UI 要素から別の実行時にユーザー定義関数を実行するに注意する必要があります。</span><span class="sxs-lookup"><span data-stu-id="51f83-105">While custom functions doesn't require a specific method of authentication, you should be aware that custom functions runs in a separate runtime from the task pane and other UI elements of your add-in.</span></span> <span data-ttu-id="51f83-106">使用して 2 つのランタイムの間で前後にデータを渡す必要があります、このため、`AsyncStorage`オブジェクトとダイアログ ボックス API です。</span><span class="sxs-lookup"><span data-stu-id="51f83-106">Because of this, you'll need to pass data back and forth between the two runtimes using the `AsyncStorage` object and the Dialog API.</span></span>
  
## <a name="asyncstorage-object"></a><span data-ttu-id="51f83-107">AsyncStorage オブジェクト</span><span class="sxs-lookup"><span data-stu-id="51f83-107">AsyncStorage object</span></span>

<span data-ttu-id="51f83-108">ユーザー定義関数の実行時の有効期限がない、`localStorage`される可能性があります通常データを格納するグローバル ウィンドウで、使用可能なオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="51f83-108">The custom functions runtime doesn't have a `localStorage` object available on the global window, where you might typically store data.</span></span> <span data-ttu-id="51f83-109">代わりに、設定およびデータを取得する[OfficeRuntime.AsyncStorage](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.asyncstorage)を使用して、独自の機能と作業ウィンドウ間でデータを共有する必要があります。</span><span class="sxs-lookup"><span data-stu-id="51f83-109">Instead, you should share data between custom functions and task panes, by using [OfficeRuntime.AsyncStorage](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.asyncstorage) to set and get data.</span></span> 

<span data-ttu-id="51f83-110">使用するメリットがあるさらに、 `AsyncStorage`。その他のアドインを使用して、データにアクセスできないようにセキュリティで保護されたサンド ボックス環境を使用します。</span><span class="sxs-lookup"><span data-stu-id="51f83-110">Additionally, there is a benefit to using `AsyncStorage`; it uses a secure sandbox environment so that your data cannot be accessed by other add-ins.</span></span>  

### <a name="suggested-usage"></a><span data-ttu-id="51f83-111">推奨される使用方法</span><span class="sxs-lookup"><span data-stu-id="51f83-111">Suggested usage</span></span>

<span data-ttu-id="51f83-112">作業ウィンドウまたはカスタム関数のいずれかを認証する場合は、アクセス トークンがすでに取得したかどうかを参照してくださいに AsyncStorage を確認してください。</span><span class="sxs-lookup"><span data-stu-id="51f83-112">When you need to authenticate either from the task pane or a custom function, check AsyncStorage to see if the access token was already acquired.</span></span> <span data-ttu-id="51f83-113">それ以外の場合は、ダイアログ ボックス API を使用して、ユーザーを認証し、アクセス トークンを取得し、AsyncStorage で後で使用できるトークンを格納します。</span><span class="sxs-lookup"><span data-stu-id="51f83-113">If not, use the dialog API to authenticate the user, retrieve the access token, and then store the token in AsyncStorage for future use.</span></span>

## <a name="dialog-api"></a><span data-ttu-id="51f83-114">ダイアログ API</span><span class="sxs-lookup"><span data-stu-id="51f83-114">Dialog API</span></span>

<span data-ttu-id="51f83-115">トークンが存在しない場合は、サインインするユーザーを確認するダイアログ ボックス API を使用してください。</span><span class="sxs-lookup"><span data-stu-id="51f83-115">If a token doesn't exist, you should use the Dialog API to ask the user to sign in.</span></span> <span data-ttu-id="51f83-116">作成されたアクセス トークンを格納できるユーザーが各自の資格情報を入力した後`AsyncStorage`。</span><span class="sxs-lookup"><span data-stu-id="51f83-116">After a user enters their credentials, the resulting access token can be stored in `AsyncStorage`.</span></span>

> [!NOTE]
> <span data-ttu-id="51f83-117">ユーザー定義関数の実行時では、作業ウィンドウで使用される実行時にダイアログ オブジェクトとは少し異なりますダイアログ オブジェクトを使用します。</span><span class="sxs-lookup"><span data-stu-id="51f83-117">The custom functions runtime uses a Dialog object that is slightly different from the Dialog object in the runtime used by task panes.</span></span> <span data-ttu-id="51f83-118">いる両方と呼ばれる「ダイアログ API」、これらの使用が`Officeruntime.Dialog`ユーザー定義関数の実行時にユーザー認証します。</span><span class="sxs-lookup"><span data-stu-id="51f83-118">They're both referred to as the "Dialog API", but use `Officeruntime.Dialog` to authenticate users in the custom functions runtime.</span></span>

<span data-ttu-id="51f83-119">使用する方法については、 `OfficeRuntime.Dialog`、[実行時のカスタム関数](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime?view=office-js#displaying-a-dialog-box)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="51f83-119">For information on how to use the `OfficeRuntime.Dialog`, see [Custom Functions runtime](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime?view=office-js#displaying-a-dialog-box).</span></span>

<span data-ttu-id="51f83-120">全体として全体の認証プロセスを予見するには場合があります] 作業ウィンドウと、アドインの UI 要素と考えるとよいとカスタムを通じて相互に通信できる、個別のエンティティとして、アドインの一部の機能`AsyncStorage`。</span><span class="sxs-lookup"><span data-stu-id="51f83-120">When envisioning the entire authentication process as a whole, it might be helpful to think of the task pane and UI elements of your add-in and the custom functions portions of your add-in as separate entities which can communicate with each other through `AsyncStorage`.</span></span>

<span data-ttu-id="51f83-121">次の図では、この基本的な手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="51f83-121">The following diagram outlines this basic process.</span></span> <span data-ttu-id="51f83-122">点線に個別の操作を実行すると、ユーザー定義関数、アドインの作業ウィンドウは、アドインを全体としての両方の部分を示すことに注意してください。</span><span class="sxs-lookup"><span data-stu-id="51f83-122">Note that the dotted line indicates that while they perform separate actions, custom functions and your add-in's task pane are both parts of your add-in as a whole.</span></span>

1. <span data-ttu-id="51f83-123">Excel ブック内のセルからユーザー定義関数の呼び出しを発行するとします。</span><span class="sxs-lookup"><span data-stu-id="51f83-123">You issue a custom function call from a cell in an Excel workbook.</span></span>
2. <span data-ttu-id="51f83-124">ユーザー定義関数を使用して`Officeruntime.Dialog`web サイトにユーザーの資格情報を渡すことです。</span><span class="sxs-lookup"><span data-stu-id="51f83-124">The custom function uses `Officeruntime.Dialog` to pass your user credentials to a website.</span></span>
3. <span data-ttu-id="51f83-125">この web サイトは、アクセス トークンをユーザー定義関数に戻ります。</span><span class="sxs-lookup"><span data-stu-id="51f83-125">This website then returns an access token to the custom function.</span></span>
4. <span data-ttu-id="51f83-126">ユーザー定義関数が、このアクセス トークンを設定、 `AsyncStorage`。</span><span class="sxs-lookup"><span data-stu-id="51f83-126">Your custom function then sets this access token to the `AsyncStorage`.</span></span>
5. <span data-ttu-id="51f83-127">アドインの作業ウィンドウからのトークンにアクセスする`AsyncStorage`。</span><span class="sxs-lookup"><span data-stu-id="51f83-127">Your add-in's task pane accesses the token from `AsyncStorage`.</span></span>

<span data-ttu-id="51f83-128">![ユーザー定義関数、OfficeRuntime、および共同作業の作業ウィンドウのダイアグラム]。(../images/Authdiagram.png "認証のダイアグラム")。</span><span class="sxs-lookup"><span data-stu-id="51f83-128">![Diagram of custom functions, OfficeRuntime, and task panes working together.](../images/Authdiagram.png "Authentication diagram.")</span></span>

## <a name="general-guidance"></a><span data-ttu-id="51f83-129">一般的なガイダンス</span><span class="sxs-lookup"><span data-stu-id="51f83-129">General guidance</span></span>

<span data-ttu-id="51f83-130">Office アドインでは、web ベースおよび web のすべての認証テクニックを使用することができます。</span><span class="sxs-lookup"><span data-stu-id="51f83-130">Office Add-ins are web-based and you can use any web authentication technique.</span></span> <span data-ttu-id="51f83-131">特定のパターンまたはカスタム関数を使用して、独自の認証を実装するメソッドはありません。</span><span class="sxs-lookup"><span data-stu-id="51f83-131">There is no particular pattern or method you must follow to implement your own authentication with custom functions.</span></span> <span data-ttu-id="51f83-132">できます各種の認証パターンについてのマニュアルを参照する[外部サービスを使用して付与するには、この資料](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins?view=office-js)で始まります。</span><span class="sxs-lookup"><span data-stu-id="51f83-132">You may wish to consult the documentation about various authentication patterns, starting with [this article about authorizing via external services](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins?view=office-js).</span></span>  

<span data-ttu-id="51f83-133">次の場所を使用してカスタム機能を開発するときにデータを格納するを回避するには。</span><span class="sxs-lookup"><span data-stu-id="51f83-133">Avoid using the following locations to store data when developing custom functions:</span></span>  

- <span data-ttu-id="51f83-134">`localStorage`: ユーザー定義関数では、グローバルへのアクセスを必要はありません`window`オブジェクト、に格納されているデータへのアクセスはそのためありません`localStorage`。</span><span class="sxs-lookup"><span data-stu-id="51f83-134">`localStorage`: Custom functions do not have access to the global `window` object and therefore have no access to data     stored in `localStorage`.</span></span>
- <span data-ttu-id="51f83-135">`Office.context.document.settings`: この場所は安全ではありませんし、アドインを使用するすべてのユーザーが情報を抽出することができます。</span><span class="sxs-lookup"><span data-stu-id="51f83-135">`Office.context.document.settings`:  This location is not secure and information can be extracted by anyone using the     add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="51f83-136">関連項目</span><span class="sxs-lookup"><span data-stu-id="51f83-136">See also</span></span>

* [<span data-ttu-id="51f83-137">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="51f83-137">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="51f83-138">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="51f83-138">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="51f83-139">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="51f83-139">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="51f83-140">チュートリアル: Excel でカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="51f83-140">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)

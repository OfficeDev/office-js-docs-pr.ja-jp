---
ms.date: 01/29/2019
description: Excel でカスタム関数を使用してユーザーを認証します。
title: カスタム関数の認証
ms.openlocfilehash: 260f15c39758b82a2145474f543c3c9ff5edd132
ms.sourcegitcommit: 70ef38a290c18a1d1a380fd02b263470207a5dc6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/15/2019
ms.locfileid: "30052736"
---
# <a name="authentication"></a><span data-ttu-id="3d1c0-103">認証</span><span class="sxs-lookup"><span data-stu-id="3d1c0-103">Authentication</span></span>

<span data-ttu-id="3d1c0-104">一部のシナリオでは、カスタム関数は、保護されたリソースにアクセスするためにユーザーを認証する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-104">In some scenarios your custom function will need to authenticate the user in order to access protected resources.</span></span> <span data-ttu-id="3d1c0-105">カスタム関数は、特定の認証方法を必要としませんが、カスタム関数は、アドインの作業ウィンドウや他の UI 要素とは別のランタイムで実行されることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-105">While custom functions doesn't require a specific method of authentication, you should be aware that custom functions runs in a separate runtime from the task pane and other UI elements of your add-in.</span></span> <span data-ttu-id="3d1c0-106">そのため、 `AsyncStorage`オブジェクトとダイアログ API を使用して、2つのランタイム間でデータをやり取りする必要があります。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-106">Because of this, you'll need to pass data back and forth between the two runtimes using the `AsyncStorage` object and the Dialog API.</span></span>
  
## <a name="asyncstorage-object"></a><span data-ttu-id="3d1c0-107">asyncstorage オブジェクト</span><span class="sxs-lookup"><span data-stu-id="3d1c0-107">AsyncStorage object</span></span>

<span data-ttu-id="3d1c0-108">カスタム関数ランタイムには、通常`localStorage` 、データを格納するグローバルウィンドウで使用できるオブジェクトがありません。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-108">The custom functions runtime doesn't have a `localStorage` object available on the global window, where you might typically store data.</span></span> <span data-ttu-id="3d1c0-109">代わりに、データを設定して取得するために、 [officeruntime](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.asyncstorage)を使用して、カスタム関数と作業ウィンドウ間でデータを共有する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-109">Instead, you should share data between custom functions and task panes, by using [OfficeRuntime.AsyncStorage](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.asyncstorage) to set and get data.</span></span> 

<span data-ttu-id="3d1c0-110">さらに、を使用`AsyncStorage`するメリットがあります。セキュリティで保護されたサンドボックス環境を使用して、他のアドインがデータにアクセスできないようにします。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-110">Additionally, there is a benefit to using `AsyncStorage`; it uses a secure sandbox environment so that your data cannot be accessed by other add-ins.</span></span>  

### <a name="suggested-usage"></a><span data-ttu-id="3d1c0-111">推奨される使用法</span><span class="sxs-lookup"><span data-stu-id="3d1c0-111">Suggested usage</span></span>

<span data-ttu-id="3d1c0-112">作業ウィンドウまたはカスタム関数から認証する必要がある場合は、asyncstorage でアクセストークンが既に取得されているかどうかを確認します。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-112">When you need to authenticate either from the task pane or a custom function, check AsyncStorage to see if the access token was already acquired.</span></span> <span data-ttu-id="3d1c0-113">それ以外の場合は、ダイアログ API を使用してユーザーを認証し、アクセストークンを取得してから、トークンを asyncstorage に保存しておきます。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-113">If not, use the dialog API to authenticate the user, retrieve the access token, and then store the token in AsyncStorage for future use.</span></span>

## <a name="dialog-api"></a><span data-ttu-id="3d1c0-114">ダイアログ API</span><span class="sxs-lookup"><span data-stu-id="3d1c0-114">Dialog API</span></span>

<span data-ttu-id="3d1c0-115">トークンが存在しない場合は、ダイアログ API を使用して、ユーザーにサインインを要求する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-115">If a token doesn't exist, you should use the Dialog API to ask the user to sign in.</span></span> <span data-ttu-id="3d1c0-116">ユーザーが資格情報を入力すると、作成されたアクセストークン`AsyncStorage`がに保存されます。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-116">After a user enters their credentials, the resulting access token can be stored in `AsyncStorage`.</span></span>

> [!NOTE]
> <span data-ttu-id="3d1c0-117">カスタム関数ランタイムは、作業ウィンドウで使用されるランタイムの dialog オブジェクトとは少し異なるダイアログオブジェクトを使用します。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-117">The custom functions runtime uses a Dialog object that is slightly different from the Dialog object in the runtime used by task panes.</span></span> <span data-ttu-id="3d1c0-118">これらはどちらも "Dialog API" と呼ばれています`Officeruntime.Dialog`が、カスタム関数ランタイムでユーザーを認証するために使用します。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-118">They're both referred to as the "Dialog API", but use `Officeruntime.Dialog` to authenticate users in the custom functions runtime.</span></span>

<span data-ttu-id="3d1c0-119">の`OfficeRuntime.Dialog`使用方法については、「[カスタム関数ランタイム](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime?view=office-js#displaying-a-dialog-box)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-119">For information on how to use the `OfficeRuntime.Dialog`, see [Custom Functions runtime](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime?view=office-js#displaying-a-dialog-box).</span></span>

<span data-ttu-id="3d1c0-120">全体として認証プロセス全体を構想する場合は、アドインの作業ウィンドウと UI 要素、およびアドインのカスタム関数部分を、を通じて`AsyncStorage`相互に通信できる個別のエンティティと考えることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-120">When envisioning the entire authentication process as a whole, it might be helpful to think of the task pane and UI elements of your add-in and the custom functions portions of your add-in as separate entities which can communicate with each other through `AsyncStorage`.</span></span>

<span data-ttu-id="3d1c0-121">次の図は、この基本的なプロセスの概要を示しています。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-121">The following diagram outlines this basic process.</span></span> <span data-ttu-id="3d1c0-122">点線では、個別のアクションを実行する一方で、カスタム関数とアドインの作業ウィンドウは、どちらもアドインの一部であることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-122">Note that the dotted line indicates that while they perform separate actions, custom functions and your add-in's task pane are both parts of your add-in as a whole.</span></span>

1. <span data-ttu-id="3d1c0-123">Excel ブックのセルからカスタム関数呼び出しを発行します。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-123">You issue a custom function call from a cell in an Excel workbook.</span></span>
2. <span data-ttu-id="3d1c0-124">カスタム関数を使用`Officeruntime.Dialog`して、ユーザーの資格情報を web サイトに渡します。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-124">The custom function uses `Officeruntime.Dialog` to pass your user credentials to a website.</span></span>
3. <span data-ttu-id="3d1c0-125">その後、この web サイトは、カスタム関数へのアクセストークンを返します。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-125">This website then returns an access token to the custom function.</span></span>
4. <span data-ttu-id="3d1c0-126">次に、カスタム関数は、 `AsyncStorage`このアクセストークンをに設定します。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-126">Your custom function then sets this access token to the `AsyncStorage`.</span></span>
5. <span data-ttu-id="3d1c0-127">アドインの作業ウィンドウで、から`AsyncStorage`トークンにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-127">Your add-in's task pane accesses the token from `AsyncStorage`.</span></span>

<span data-ttu-id="3d1c0-128">![カスタム関数、officeruntime、および共同作業ウィンドウの図](../images/Authdiagram.png "認証の図。")</span><span class="sxs-lookup"><span data-stu-id="3d1c0-128">![Diagram of custom functions, OfficeRuntime, and task panes working together.](../images/Authdiagram.png "Authentication diagram.")</span></span>

## <a name="general-guidance"></a><span data-ttu-id="3d1c0-129">一般的なガイダンス</span><span class="sxs-lookup"><span data-stu-id="3d1c0-129">General guidance</span></span>

<span data-ttu-id="3d1c0-130">Office アドインは web ベースであり、任意の web 認証方法を使用できます。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-130">Office Add-ins are web-based and you can use any web authentication technique.</span></span> <span data-ttu-id="3d1c0-131">カスタム関数を使用して独自の認証を実装するために従う必要のある特定のパターンやメソッドはありません。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-131">There is no particular pattern or method you must follow to implement your own authentication with custom functions.</span></span> <span data-ttu-id="3d1c0-132">[外部サービスによる承認については、この記事](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins?view=office-js)から始まるさまざまな認証パターンに関するドキュメントを参照してください。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-132">You may wish to consult the documentation about various authentication patterns, starting with [this article about authorizing via external services](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins?view=office-js).</span></span>  

<span data-ttu-id="3d1c0-133">カスタム関数を開発する際に、次の場所を使用してデータを保存しないようにします。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-133">Avoid using the following locations to store data when developing custom functions:</span></span>  

- <span data-ttu-id="3d1c0-134">`localStorage`: カスタム関数にはグローバル`window`オブジェクトへのアクセス権がないため、に`localStorage`格納されているデータにアクセスできません。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-134">`localStorage`: Custom functions do not have access to the global `window` object and therefore have no access to data     stored in `localStorage`.</span></span>
- <span data-ttu-id="3d1c0-135">`Office.context.document.settings`: この場所はセキュリティで保護されていないため、アドインを使用するすべてのユーザーが情報を抽出できます。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-135">`Office.context.document.settings`:  This location is not secure and information can be extracted by anyone using the     add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="3d1c0-136">関連項目</span><span class="sxs-lookup"><span data-stu-id="3d1c0-136">See also</span></span>

* [<span data-ttu-id="3d1c0-137">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="3d1c0-137">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="3d1c0-138">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="3d1c0-138">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="3d1c0-139">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="3d1c0-139">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="3d1c0-140">チュートリアル: Excel でカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="3d1c0-140">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)

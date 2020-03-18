---
ms.date: 02/13/2020
description: カスタム関数、リボン ボタン、作業ウィンドウのコードを同じ JavaScript ランタイムで実行して、さまざまなアドインでシナリオを調整する方法について説明します。
title: 共有の JavaScript ランタイムでアドイン コードを実行する (プレビュー)
localization_priority: Priority
ms.openlocfilehash: 774990a9452d450bd5c4d968027bc64ebee858af
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719533"
---
# <a name="overview-run-your-add-in-code-in-a-shared-javascript-runtime-preview"></a><span data-ttu-id="77e2e-103">概要: 共有の JavaScript ランタイムでアドイン コードを実行する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="77e2e-103">Overview: Run your add-in code in a shared JavaScript runtime (preview)</span></span>

[!include[Running custom functions in shared JavaScript runtime note](../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="77e2e-104">Windows または Mac で Excel を実行する場合、アドインは、リボン ボタン、カスタム関数、作業ウィンドウのコードを別の JavaScript ランタイム環境で実行します。</span><span class="sxs-lookup"><span data-stu-id="77e2e-104">When running Excel on Windows or Mac, your add-in will run code for ribbon buttons, custom functions, and the task pane in separate JavaScript runtime environments.</span></span> <span data-ttu-id="77e2e-105">これにより、グローバル データを簡単に共有できない、カスタム関数からすべての CORS 機能にアクセスできないなどの制限が発生します。</span><span class="sxs-lookup"><span data-stu-id="77e2e-105">This creates limitations such as not being able to easily share global data, and not being able to access all CORS functionality from a custom function.</span></span>

<span data-ttu-id="77e2e-106">ただし、Excel アドインを構成すれば、同じ JavaScript ランタイム (共有ランタイムとも呼ばれる) でコードを共有できるようになります。</span><span class="sxs-lookup"><span data-stu-id="77e2e-106">However, you can configure your Excel add-in to share code in the same JavaScript runtime (also referred to as a shared runtime).</span></span> <span data-ttu-id="77e2e-107">これにより、アドイン間での調整が容易になり、アドインのすべての部分から、作業ウィンドウの DOM や CORS にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="77e2e-107">This enables better coordination across your add-in and access to the task pane DOM and CORS from all parts of your add-in.</span></span>

<span data-ttu-id="77e2e-108">共有ランタイムを構成すると、次のシナリオが可能になります。</span><span class="sxs-lookup"><span data-stu-id="77e2e-108">Configuring a shared runtime enables the following scenarios:</span></span>

- <span data-ttu-id="77e2e-109">アドインに、リボン、作業ウィンドウ、カスタム関数のすべてがアクセスできる共有の DOM が含まれます。</span><span class="sxs-lookup"><span data-stu-id="77e2e-109">Your add-in will have a shared DOM that the ribbon, task pane, and custom functions can all access.</span></span>
- <span data-ttu-id="77e2e-110">カスタム関数で CORS がすべてサポートされます。</span><span class="sxs-lookup"><span data-stu-id="77e2e-110">Your custom functions will have full CORS support.</span></span>
- <span data-ttu-id="77e2e-111">カスタム関数で、Office.js API を呼び出して、スプレッドシート ドキュメントのデータを読み取ることができます。</span><span class="sxs-lookup"><span data-stu-id="77e2e-111">Your custom functions can call Office.js APIs to read spreadsheet document data.</span></span>
- <span data-ttu-id="77e2e-112">ドキュメントを開いてすぐに、アドインでコードを実行できます。</span><span class="sxs-lookup"><span data-stu-id="77e2e-112">Your add-in can run code as soon as the document is opened.</span></span>
- <span data-ttu-id="77e2e-113">作業ウィンドウが閉じられた後でも、アドインでコードの実行を続けられます。</span><span class="sxs-lookup"><span data-stu-id="77e2e-113">Your add-in can continue running code after the task pane is closed.</span></span>

<span data-ttu-id="77e2e-114">共有ランタイムを使用して作業ウィンドウでカスタム関数を実行すると、「[Office アドインで使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」で説明されているように、別のプラットフォームのブラウザー インスタンスで実行されます。また、Excel アドインのリボンに表示するボタンはすべて、同じ共有ランタイムで実行されます。</span><span class="sxs-lookup"><span data-stu-id="77e2e-114">When you run custom functions in a shared runtime with the task pane, it will run in a browser instance on different platforms as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Additionally, any buttons that your Excel add-in displays on the ribbon will run in the same shared runtime.</span></span> <span data-ttu-id="77e2e-115">次の図は、カスタム関数、リボン UI、作業ウィンドウのコードがすべて同じ JavaScript ランタイム内で実行される様子を示しています。</span><span class="sxs-lookup"><span data-stu-id="77e2e-115">The following image shows how custom functions, the ribbon UI, and the task pane code will all run in the same JavaScript runtime.</span></span>

![Excel でカスタム関数をリボン ボタンと作業ウィンドウと一緒に共有ランタイムで実行](../images/custom-functions-in-browser-runtime.png)

## <a name="differences-when-running-custom-functions-in-a-shared-runtime"></a><span data-ttu-id="77e2e-117">共有ランタイムでカスタム関数を実行するときの違い</span><span class="sxs-lookup"><span data-stu-id="77e2e-117">Differences when running custom functions in a shared runtime</span></span>

<span data-ttu-id="77e2e-118">Excel アドイン プロジェクトを構成して、共有ランタイムでカスタム関数を実行する場合、カスタム関数のランタイムを使用するのとは異なる点がいくつかあります。</span><span class="sxs-lookup"><span data-stu-id="77e2e-118">When you configure your Excel add-in project to run custom functions in a shared runtime, there are a few differences from using the custom function runtime.</span></span>

### <a name="storage"></a><span data-ttu-id="77e2e-119">ストレージ</span><span class="sxs-lookup"><span data-stu-id="77e2e-119">Storage</span></span>

<span data-ttu-id="77e2e-120">作業ウィンドウ、カスタム関数、リボン UI の間でデータを共有するための**ストレージ** API を使用する必要がなくなりました。</span><span class="sxs-lookup"><span data-stu-id="77e2e-120">You no longer need to use the **Storage** API to share data between the task pane, custom functions or ribbon UI.</span></span> <span data-ttu-id="77e2e-121">**ウィンドウ** オブジェクトにグローバル変数を入力するか、お好みの状態管理アプローチを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="77e2e-121">You can put global variables in the **window** object, or use your own preferred state management approach.</span></span>

### <a name="authentication"></a><span data-ttu-id="77e2e-122">認証</span><span class="sxs-lookup"><span data-stu-id="77e2e-122">Authentication</span></span>

<span data-ttu-id="77e2e-123">認証の一環としてトークンを受け取る場合、作業ウィンドウ、カスタム関数、リボン UI 間でそのトークンを共有するために **ストレージ** API を使用する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="77e2e-123">When you receive tokens as part of authentication, you don't need to use the **Storage** API to share them between the task pane, custom functions and ribbon UI.</span></span> <span data-ttu-id="77e2e-124">お好みのストレージ方法で `localStorage` などの保存場所で共有することができます。</span><span class="sxs-lookup"><span data-stu-id="77e2e-124">You can use your own preferred storage technique and storage location to share them, such as `localStorage`.</span></span>

### <a name="dialog-api"></a><span data-ttu-id="77e2e-125">ダイアログ API</span><span class="sxs-lookup"><span data-stu-id="77e2e-125">Dialog API</span></span>

<span data-ttu-id="77e2e-126">**OfficeRuntime.Dialog** API を使ってカスタム関数からのダイアログを表示する必要はなくなります。</span><span class="sxs-lookup"><span data-stu-id="77e2e-126">You no longer need to use the **OfficeRuntime.Dialog** API to display a dialog from a custom function.</span></span> <span data-ttu-id="77e2e-127">カスタム関数、リボン ボタン、作業ウィンドウに対して、同じ[ダイアログ API](../develop/dialog-api-in-office-add-ins.md) を使うことができます。</span><span class="sxs-lookup"><span data-stu-id="77e2e-127">You can use the same [Dialog API](../develop/dialog-api-in-office-add-ins.md) for custom functions, ribbon buttons, and the task pane.</span></span>

### <a name="debugging"></a><span data-ttu-id="77e2e-128">デバッグ</span><span class="sxs-lookup"><span data-stu-id="77e2e-128">Debugging</span></span>

<span data-ttu-id="77e2e-129">共有ランタイムを使用している場合、この時点では、Windows の Excel でカスタム関数をデバッグするために Visual Studio Code を使用することはできません。</span><span class="sxs-lookup"><span data-stu-id="77e2e-129">When using a shared runtime, you can't use Visual Studio Code to debug custom functions in Excel on Windows at this time.</span></span> <span data-ttu-id="77e2e-130">開発者ツールを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="77e2e-130">You'll need to use developer tools.</span></span> <span data-ttu-id="77e2e-131">さらに詳しい情報については、「[Windows 10 で開発者ツールを使用してアドインをデバッグする](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="77e2e-131">For more information, see [Debug add-ins using developer tools on Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).</span></span>

## <a name="get-started"></a><span data-ttu-id="77e2e-132">使用を開始する</span><span class="sxs-lookup"><span data-stu-id="77e2e-132">Get Started</span></span>

<span data-ttu-id="77e2e-133">共有ランタイムでカスタム関数を実行するように Excel のアドイン プロジェクトを構成する方法については、「[共有の JavaScript ランタイムを使用するように Excel アドインを構成する (プレビュー)](configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="77e2e-133">To configure your Excel add-in project to run custom functions in a shared runtime, see [Configure your Excel add-in to use a shared JavaScript runtime (preview)](configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="give-us-feedback"></a><span data-ttu-id="77e2e-134">ご意見をお寄せください</span><span class="sxs-lookup"><span data-stu-id="77e2e-134">Give us feedback</span></span>

<span data-ttu-id="77e2e-135">この機能について、ご意見をお待ちしております。</span><span class="sxs-lookup"><span data-stu-id="77e2e-135">We'd love to hear your feedback on this feature.</span></span> <span data-ttu-id="77e2e-136">バグや問題が発生したり、この機能について要求がございましたら、[office-js repo](https://github.com/OfficeDev/office-js) で GitHub に関する問題を作成してお知らせください。</span><span class="sxs-lookup"><span data-stu-id="77e2e-136">If you find any bugs, issues, or have requests on this feature, please let us know by creating a GitHub issue in the [office-js repo](https://github.com/OfficeDev/office-js).</span></span>

## <a name="see-also"></a><span data-ttu-id="77e2e-137">関連項目</span><span class="sxs-lookup"><span data-stu-id="77e2e-137">See also</span></span>

<span data-ttu-id="77e2e-138">共有ランタイムの関連記事の一覧</span><span class="sxs-lookup"><span data-stu-id="77e2e-138">List of related articles for shared runtime</span></span>
- [<span data-ttu-id="77e2e-139">チュートリアル: Excel カスタム関数と作業ウィンドウの間でデータとイベントを共有する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="77e2e-139">Tutorial: Share data and events between Excel custom functions and the task pane (preview)</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [<span data-ttu-id="77e2e-140">カスタム関数から Excel API を呼び出す (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="77e2e-140">Call Excel APIs from your custom function (preview)</span></span>](call-excel-apis-from-custom-function.md)
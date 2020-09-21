---
ms.date: 08/13/2020
description: カスタム関数、リボン ボタン、作業ウィンドウのコードを同じ JavaScript ランタイムで実行して、さまざまなアドインでシナリオを調整する方法について説明します。
title: 共有の JavaScript ランタイムでアドイン コードを実行する
localization_priority: Priority
ms.openlocfilehash: 04932bcf292686fd9d0abf2ff99c19f062f21456
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819547"
---
# <a name="overview-run-your-add-in-code-in-a-shared-javascript-runtimes"></a><span data-ttu-id="13025-103">概要: 共有の JavaScript ランタイムでアドイン コードを実行する</span><span class="sxs-lookup"><span data-stu-id="13025-103">Overview: Run your add-in code in a shared JavaScript runtimes</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="13025-104">Windows または Mac で Excel を実行する場合、アドインは、リボン ボタン、カスタム関数、作業ウィンドウのコードを別の JavaScript ランタイム環境で実行します。</span><span class="sxs-lookup"><span data-stu-id="13025-104">When running Excel on Windows or Mac, your add-in will run code for ribbon buttons, custom functions, and the task pane in separate JavaScript runtime environments.</span></span> <span data-ttu-id="13025-105">これにより、グローバル データを簡単に共有できない、カスタム関数からすべての CORS 機能にアクセスできないなどの制限が発生します。</span><span class="sxs-lookup"><span data-stu-id="13025-105">This creates limitations such as not being able to easily share global data, and not being able to access all CORS functionality from a custom function.</span></span>

<span data-ttu-id="13025-106">ただし、Excel アドインを構成すれば、同じ JavaScript ランタイム (共有ランタイムとも呼ばれる) でコードを共有できるようになります。</span><span class="sxs-lookup"><span data-stu-id="13025-106">However, you can configure your Excel add-in to share code in the same JavaScript runtime (also referred to as a shared runtime).</span></span> <span data-ttu-id="13025-107">これにより、アドイン間での調整が容易になり、アドインのすべての部分から、作業ウィンドウの DOM や CORS にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="13025-107">This enables better coordination across your add-in and access to the task pane DOM and CORS from all parts of your add-in.</span></span>

<span data-ttu-id="13025-108">共有ランタイムを構成すると、次のシナリオが可能になります。</span><span class="sxs-lookup"><span data-stu-id="13025-108">Configuring a shared runtime enables the following scenarios:</span></span>

- <span data-ttu-id="13025-109">アドインに、リボン、作業ウィンドウ、カスタム関数のすべてがアクセスできる共有の DOM が含まれます。</span><span class="sxs-lookup"><span data-stu-id="13025-109">Your add-in will have a shared DOM that the ribbon, task pane, and custom functions can all access.</span></span>
- <span data-ttu-id="13025-110">カスタム関数で CORS がすべてサポートされます。</span><span class="sxs-lookup"><span data-stu-id="13025-110">Your custom functions will have full CORS support.</span></span>
- <span data-ttu-id="13025-111">カスタム関数で、Office.js API を呼び出して、スプレッドシート ドキュメントのデータを読み取ることができます。</span><span class="sxs-lookup"><span data-stu-id="13025-111">Your custom functions can call Office.js APIs to read spreadsheet document data.</span></span>
- <span data-ttu-id="13025-112">ドキュメントを開いてすぐに、アドインでコードを実行できます。</span><span class="sxs-lookup"><span data-stu-id="13025-112">Your add-in can run code as soon as the document is opened.</span></span>
- <span data-ttu-id="13025-113">作業ウィンドウが閉じられた後でも、アドインでコードの実行を続けられます。</span><span class="sxs-lookup"><span data-stu-id="13025-113">Your add-in can continue running code after the task pane is closed.</span></span>

<span data-ttu-id="13025-114">共有ランタイムを使用して作業ウィンドウでカスタム関数を実行すると、「[Office アドインで使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」で説明されているように、アドインが Microsoft Internet Explorer 11 のブラウザー インスタンスで実行されます。また、Excel アドインのリボンに表示するボタンはすべて、同じ共有ランタイムで実行されます。</span><span class="sxs-lookup"><span data-stu-id="13025-114">When you run custom functions in a shared runtime with the task pane, your add-in will run in a Microsoft Internet Explorer 11 browser instance, as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Additionally, any buttons that your Excel add-in displays on the ribbon will run in the same shared runtime.</span></span> <span data-ttu-id="13025-115">次の図は、カスタム関数、リボン UI、作業ウィンドウのコードがすべて同じ JavaScript ランタイム内で実行される様子を示しています。</span><span class="sxs-lookup"><span data-stu-id="13025-115">The following image shows how custom functions, the ribbon UI, and the task pane code will all run in the same JavaScript runtime.</span></span>

![Excel のリボン ボタンと作業ウィンドウを備えた共有ランタイムで実行されるカスタム関数](../images/custom-functions-in-browser-runtime.png)

## <a name="set-up-a-shared-runtime"></a><span data-ttu-id="13025-117">共有ランタイムを設定する</span><span class="sxs-lookup"><span data-stu-id="13025-117">Set up a shared runtime</span></span>

<span data-ttu-id="13025-118">共有ランタイムを使用するようにカスタム関数を設定する方法については、[共有ランタイムの設定に関する記事](configure-your-add-in-to-use-a-shared-runtime.md) をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="13025-118">See the [configuring a shared runtime article](configure-your-add-in-to-use-a-shared-runtime.md) to learn how to set up your custom functions to use a shared runtime.</span></span>

### <a name="debugging"></a><span data-ttu-id="13025-119">デバッグ</span><span class="sxs-lookup"><span data-stu-id="13025-119">Debugging</span></span>

<span data-ttu-id="13025-120">共有ランタイムを使用している場合、この時点では、Windows の Excel でカスタム関数をデバッグするために Visual Studio Code を使用することはできません。</span><span class="sxs-lookup"><span data-stu-id="13025-120">When using a shared runtime, you can't use Visual Studio Code to debug custom functions in Excel on Windows at this time.</span></span> <span data-ttu-id="13025-121">代わりに開発者ツールを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="13025-121">You'll need to use developer tools instead.</span></span> <span data-ttu-id="13025-122">さらに詳しい情報については、「[Windows 10 で開発者ツールを使用してアドインをデバッグする](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="13025-122">For more information, see [Debug add-ins using developer tools on Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).</span></span>

## <a name="give-us-feedback"></a><span data-ttu-id="13025-123">ご意見ご感想をお寄せください</span><span class="sxs-lookup"><span data-stu-id="13025-123">Give us feedback</span></span>

<span data-ttu-id="13025-124">この機能について、ご意見をお待ちしております。</span><span class="sxs-lookup"><span data-stu-id="13025-124">We'd love to hear your feedback on this feature.</span></span> <span data-ttu-id="13025-125">バグや問題が発生したり、この機能について要求がございましたら、[office-js repo](https://github.com/OfficeDev/office-js) で GitHub に関する問題を作成してお知らせください。</span><span class="sxs-lookup"><span data-stu-id="13025-125">If you find any bugs, issues, or have requests on this feature, please let us know by creating a GitHub issue in the [office-js repo](https://github.com/OfficeDev/office-js).</span></span>

## <a name="see-also"></a><span data-ttu-id="13025-126">関連項目</span><span class="sxs-lookup"><span data-stu-id="13025-126">See also</span></span>

- [<span data-ttu-id="13025-127">チュートリアル: Excel カスタム関数と作業ウィンドウの間でデータとイベントを共有する</span><span class="sxs-lookup"><span data-stu-id="13025-127">Tutorial: Share data and events between Excel custom functions and the task pane</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [<span data-ttu-id="13025-128">カスタム関数から Excel API を呼び出す</span><span class="sxs-lookup"><span data-stu-id="13025-128">Call Excel APIs from your custom function</span></span>](call-excel-apis-from-custom-function.md)

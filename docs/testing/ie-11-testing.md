---
title: Internet Explorer 11 テスト
description: 11 でOfficeアドインをテストInternet Explorerします。
ms.date: 05/19/2021
localization_priority: Normal
ms.openlocfilehash: de256ee8b0633f18d3188c5bbfae52cb24ff2c35
ms.sourcegitcommit: 0d3bf72f8ddd1b287bf95f832b7ecb9d9fa62a24
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/02/2021
ms.locfileid: "52727935"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a><span data-ttu-id="bd308-103">11 でOfficeアドインをテストInternet Explorerする</span><span class="sxs-lookup"><span data-stu-id="bd308-103">Test your Office Add-in on Internet Explorer 11</span></span>

<span data-ttu-id="bd308-104">AppSource を使用してアドインを販売する予定がある場合、または以前のバージョンの Windows および Office をサポートする予定の場合、アドインは Internet Explorer 11 (IE11) に基づく埋め込み可能なブラウザー コントロールで動作する必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd308-104">If you plan to market your add-in through AppSource or you plan to support older versions of Windows and Office, your add-in must work in the embeddable browser control that is based on Internet Explorer 11 (IE11).</span></span> <span data-ttu-id="bd308-105">コマンド ラインを使用して、アドインで使用される最新のランタイムから、このテスト用の Internet Explorer 11 ランタイムに切り替えます。</span><span class="sxs-lookup"><span data-stu-id="bd308-105">You can use a command line to switch from more modern runtimes used by add-ins to the Internet Explorer 11 runtime for this testing.</span></span> <span data-ttu-id="bd308-106">Windows および Office のバージョンで Internet Explorer 11 Web ビュー コントロールを使用する方法については、「Office アドインで使用されるブラウザー」を[参照](../concepts/browsers-used-by-office-web-add-ins.md)してください。</span><span class="sxs-lookup"><span data-stu-id="bd308-106">For information about which versions of Windows and Office use the Internet Explorer 11 web view control, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bd308-107">Internet Explorer 11はES5以降のJavaScriptバージョンをサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="bd308-107">Internet Explorer 11 does not support JavaScript versions later than ES5.</span></span> <span data-ttu-id="bd308-108">ECMAScript 2015 以降の構文と機能を使用する場合は、次の 2 つのオプションがあります。</span><span class="sxs-lookup"><span data-stu-id="bd308-108">If you want to use the syntax and features of ECMAScript 2015 or later, you have two options:</span></span>
>
> - <span data-ttu-id="bd308-109">ECMAScript 2015 (ES6 とも呼ばれる) 以降の JavaScript または TypeScript でコードを記述し、バベルや[tsc](https://www.typescriptlang.org/index.html)などの[](https://babeljs.io/)コンパイラを使用してコードを ES5 JavaScript にコンパイルします。</span><span class="sxs-lookup"><span data-stu-id="bd308-109">Write your code in ECMAScript 2015 (also called ES6) or later JavaScript, or in TypeScript, and then compile your code to ES5 JavaScript using a compiler such as [babel](https://babeljs.io/) or [tsc](https://www.typescriptlang.org/index.html).</span></span>
> - <span data-ttu-id="bd308-110">ECMAScript 2015 以降の JavaScript で記述します[](https://en.wikipedia.org/wiki/Polyfill_(programming))が、IE でコードを実行できる[core-js](https://github.com/zloirock/core-js)などのポリフィル ライブラリも読み込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd308-110">Write in ECMAScript 2015 or later JavaScript, but also load a [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) library such as [core-js](https://github.com/zloirock/core-js) that enables IE to run your code.</span></span>
>
> <span data-ttu-id="bd308-111">これらのオプションの詳細については [、「Support Internet Explorer 11」を参照してください](../develop/support-ie-11.md)。</span><span class="sxs-lookup"><span data-stu-id="bd308-111">For more information about these options, see [Support Internet Explorer 11](../develop/support-ie-11.md).</span></span>
>
> <span data-ttu-id="bd308-112">また、Internet Explorer 11 は、メディア、録音、および位置情報などの HTML 5 機能の一部をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="bd308-112">Also, Internet Explorer 11 does not support some HTML5 features such as media, recording, and location.</span></span>

> [!NOTE]
> <span data-ttu-id="bd308-113">Internet Explorer 11 ブラウザーでアドインをテストするには、Office on the webでInternet Explorerを開き、アドインを[サイドロードします](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="bd308-113">To test your add-in on the Internet Explorer 11 browser, open Office on the web in Internet Explorer and [sideload the add-in](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="bd308-114">前提条件</span><span class="sxs-lookup"><span data-stu-id="bd308-114">Prerequisites</span></span>

- <span data-ttu-id="bd308-115">[Node.js](https://nodejs.org/) (最新 [LTS](https://nodejs.org/about/releases) バージョン)</span><span class="sxs-lookup"><span data-stu-id="bd308-115">[Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>

<span data-ttu-id="bd308-116">これらの手順では、以前に Yo ジェネレーター プロジェクトをOffice前提とします。</span><span class="sxs-lookup"><span data-stu-id="bd308-116">These instructions assume you have set up a Yo Office generator project before.</span></span> <span data-ttu-id="bd308-117">前にこれを行ったことがない場合は、クイック スタート (アドイン用など) を読[Excel検討してください](../quickstarts/excel-quickstart-jquery.md)。</span><span class="sxs-lookup"><span data-stu-id="bd308-117">If you haven't done this before, consider reading a quick start, such as [this one for Excel add-ins](../quickstarts/excel-quickstart-jquery.md).</span></span>

## <a name="switching-to-the-internet-explorer-11-webview"></a><span data-ttu-id="bd308-118">11 webview Internet Explorer切り替える</span><span class="sxs-lookup"><span data-stu-id="bd308-118">Switching to the Internet Explorer 11 webview</span></span>

1. <span data-ttu-id="bd308-119">Yo ジェネレーター プロジェクトOffice作成します。</span><span class="sxs-lookup"><span data-stu-id="bd308-119">Create a Yo Office generator project.</span></span> <span data-ttu-id="bd308-120">選択するプロジェクトの種類は関係ありませんが、このツールは、すべてのプロジェクトの種類で動作します。</span><span class="sxs-lookup"><span data-stu-id="bd308-120">It doesn't matter what kind of project you select, this tooling will work with all project types.</span></span>

    > [!NOTE]
    > <span data-ttu-id="bd308-121">既存のプロジェクトを持ち、新しいプロジェクトを作成せずにこのツールを追加する場合は、この手順をスキップして次の手順に進みます。</span><span class="sxs-lookup"><span data-stu-id="bd308-121">If you have an existing project and want to add this tooling without creating a new project, skip this step and move to the next step.</span></span> 

1. <span data-ttu-id="bd308-122">プロジェクトのルート フォルダーで、コマンド ラインで次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="bd308-122">In the root folder of your project, run the following in the command line.</span></span> <span data-ttu-id="bd308-123">この例では、プロジェクトのマニフェスト ファイルがルートにあると仮定します。</span><span class="sxs-lookup"><span data-stu-id="bd308-123">This example assumes that your project's manifest file is in the root.</span></span> <span data-ttu-id="bd308-124">指定されていない場合は、マニフェスト ファイルへの相対パスを指定します。</span><span class="sxs-lookup"><span data-stu-id="bd308-124">If it isn't, specify the relative path to the manifest file.</span></span> <span data-ttu-id="bd308-125">コマンド ラインに、Web ビューの種類が IE に設定されているというメッセージが表示されます。</span><span class="sxs-lookup"><span data-stu-id="bd308-125">You should see a message in the command line that the web view type is now set to IE.</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

> [!TIP]
> <span data-ttu-id="bd308-126">このコマンドを使用する必要はありません。ただし、11 ランタイムに関連する問題の大部分をデバッグInternet Explorer必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd308-126">It isn't necessary to use this command, but it should help debug the majority of issues related to the Internet Explorer 11 runtime.</span></span> <span data-ttu-id="bd308-127">完全な堅牢性を得る場合は、Windows 7、8.1、および 10 とさまざまなバージョンの Office のさまざまな組み合わせのコンピューターを使用してテストする必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd308-127">For complete robustness, you should test using computers with various combinations of Windows 7, 8.1, and 10 and various versions of Office.</span></span> <span data-ttu-id="bd308-128">詳細については、「Office アドインで使用されるブラウザー」および「How to revert [to](../concepts/browsers-used-by-office-web-add-ins.md) earlier version of Office」 を[参照してください](https://support.microsoft.com/topic/how-to-revert-to-an-earlier-version-of-office-2bd5c457-a917-d57e-35a1-f709e3dda841)。</span><span class="sxs-lookup"><span data-stu-id="bd308-128">For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md) and [How to revert to an earlier version of Office](https://support.microsoft.com/topic/how-to-revert-to-an-earlier-version-of-office-2bd5c457-a917-d57e-35a1-f709e3dda841).</span></span>

### <a name="command-options"></a><span data-ttu-id="bd308-129">コマンド オプション</span><span class="sxs-lookup"><span data-stu-id="bd308-129">Command options</span></span>

<span data-ttu-id="bd308-130">この `office-addin-dev-settings webview` コマンドは、引数として多数のランタイムを受け取る場合があります。</span><span class="sxs-lookup"><span data-stu-id="bd308-130">The `office-addin-dev-settings webview` command can also take a number of runtimes as arguments:</span></span>

- <span data-ttu-id="bd308-131">すなわち</span><span class="sxs-lookup"><span data-stu-id="bd308-131">ie</span></span>
- <span data-ttu-id="bd308-132">エッジ</span><span class="sxs-lookup"><span data-stu-id="bd308-132">edge</span></span>
- <span data-ttu-id="bd308-133">default</span><span class="sxs-lookup"><span data-stu-id="bd308-133">default</span></span>

## <a name="see-also"></a><span data-ttu-id="bd308-134">関連項目</span><span class="sxs-lookup"><span data-stu-id="bd308-134">See also</span></span>

* [<span data-ttu-id="bd308-135">Office アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="bd308-135">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)
* [<span data-ttu-id="bd308-136">テスト用に Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="bd308-136">Sideload Office Add-ins for testing</span></span>](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [<span data-ttu-id="bd308-137">Windows 10 で開発者ツールを使用してアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="bd308-137">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
* [<span data-ttu-id="bd308-138">作業ウィンドウからデバッガーをアタッチする</span><span class="sxs-lookup"><span data-stu-id="bd308-138">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)

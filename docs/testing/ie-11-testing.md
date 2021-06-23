---
title: Internet Explorer 11 テスト
description: 11 でOfficeアドインをテストInternet Explorerします。
ms.date: 06/18/2021
localization_priority: Normal
ms.openlocfilehash: fa9550884a24feffdd750171f3a7e08648f9432f
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076407"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a><span data-ttu-id="ebc0e-103">11 でOfficeアドインをテストInternet Explorerする</span><span class="sxs-lookup"><span data-stu-id="ebc0e-103">Test your Office Add-in on Internet Explorer 11</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ebc0e-104">**Internet ExplorerアドインOffice引き続き使用する**</span><span class="sxs-lookup"><span data-stu-id="ebc0e-104">**Internet Explorer still used in Office Add-ins**</span></span>
>
> <span data-ttu-id="ebc0e-105">Microsoft は、アドインのサポートInternet Explorer終了していますが、これはアドインのOffice大きな影響を及ぼします。Office アドインで使用されるブラウザーで説明したように、プラットフォームと Office バージョンの一部の組み合わせ (Office 2019 までのすべての一時購入バージョンを含む) は、Internet Explorer 11 に付属する webview[](../concepts/browsers-used-by-office-web-add-ins.md)コントロールを引き続き使用してアドインをホストします。さらに、これらの組み合わせのサポートは、AppSource にInternet Explorerアドインに対して引き続き[必要です](/office/dev/store/submit-to-appsource-via-partner-center)。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-105">Microsoft is ending support for Internet Explorer, but this doesn't significantly affect Office Add-ins. Some combinations of platforms and Office versions, including all one-time-purchase versions through Office 2019, will continue to use the webview control that comes with Internet Explorer 11 to host add-ins, as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Moreover, support for these combinations, and hence for Internet Explorer, is still required for add-ins submitted to [AppSource](/office/dev/store/submit-to-appsource-via-partner-center).</span></span> <span data-ttu-id="ebc0e-106">次の *2 つの点が変化* しています。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-106">Two things *are* changing:</span></span>
>
> - <span data-ttu-id="ebc0e-107">AppSource は、ブラウザーとしてアプリケーションを使用してOffice on the webアドインInternet Explorerテストしなくなりました。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-107">AppSource no longer tests add-ins in Office on the web using Internet Explorer as the browser.</span></span> <span data-ttu-id="ebc0e-108">ただし、AppSource は引き続き、プラットフォームとデスクトップ バージョンの組み合Office *使用* するデスクトップ バージョンの組み合わせをテストInternet Explorer。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-108">But AppSource still tests for combinations of platform and Office *desktop* versions that use Internet Explorer.</span></span>
> - <span data-ttu-id="ebc0e-109">2021 [Script Lab](../overview/explore-with-script-lab.md)ツールは、2021 Internet Explorerで作業を停止します。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-109">The [Script Lab tool](../overview/explore-with-script-lab.md) will stop working in Internet Explorer sometime in 2021.</span></span>

<span data-ttu-id="ebc0e-110">AppSource を使用してアドインを販売する予定がある場合、または以前のバージョンの Windows および Office をサポートする予定の場合、アドインは Internet Explorer 11 (IE11) に基づく埋め込み可能なブラウザー コントロールで動作する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-110">If you plan to market your add-in through AppSource or you plan to support older versions of Windows and Office, your add-in must work in the embeddable browser control that is based on Internet Explorer 11 (IE11).</span></span> <span data-ttu-id="ebc0e-111">コマンド ラインを使用して、アドインで使用される最新のランタイムから、このテスト用の Internet Explorer 11 ランタイムに切り替えます。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-111">You can use a command line to switch from more modern runtimes used by add-ins to the Internet Explorer 11 runtime for this testing.</span></span> <span data-ttu-id="ebc0e-112">Windows および Office のバージョンで Internet Explorer 11 Web ビュー コントロールを使用する方法については、「Office アドインで使用されるブラウザー」を[参照](../concepts/browsers-used-by-office-web-add-ins.md)してください。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-112">For information about which versions of Windows and Office use the Internet Explorer 11 web view control, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ebc0e-113">Internet Explorer 11はES5以降のJavaScriptバージョンをサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-113">Internet Explorer 11 does not support JavaScript versions later than ES5.</span></span> <span data-ttu-id="ebc0e-114">ECMAScript 2015 以降の構文と機能を使用する場合は、次の 2 つのオプションがあります。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-114">If you want to use the syntax and features of ECMAScript 2015 or later, you have two options:</span></span>
>
> - <span data-ttu-id="ebc0e-115">ECMAScript 2015 (ES6 とも呼ばれる) 以降の JavaScript または TypeScript でコードを記述し、バベルや[tsc](https://www.typescriptlang.org/index.html)などの[](https://babeljs.io/)コンパイラを使用してコードを ES5 JavaScript にコンパイルします。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-115">Write your code in ECMAScript 2015 (also called ES6) or later JavaScript, or in TypeScript, and then compile your code to ES5 JavaScript using a compiler such as [babel](https://babeljs.io/) or [tsc](https://www.typescriptlang.org/index.html).</span></span>
> - <span data-ttu-id="ebc0e-116">ECMAScript 2015 以降の JavaScript で記述します[](https://en.wikipedia.org/wiki/Polyfill_(programming))が、IE でコードを実行できる[core-js](https://github.com/zloirock/core-js)などのポリフィル ライブラリも読み込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-116">Write in ECMAScript 2015 or later JavaScript, but also load a [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) library such as [core-js](https://github.com/zloirock/core-js) that enables IE to run your code.</span></span>
>
> <span data-ttu-id="ebc0e-117">これらのオプションの詳細については [、「Support Internet Explorer 11」を参照してください](../develop/support-ie-11.md)。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-117">For more information about these options, see [Support Internet Explorer 11](../develop/support-ie-11.md).</span></span>
>
> <span data-ttu-id="ebc0e-118">また、Internet Explorer 11 は、メディア、録音、および位置情報などの HTML 5 機能の一部をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-118">Also, Internet Explorer 11 does not support some HTML5 features such as media, recording, and location.</span></span>

> [!NOTE]
> <span data-ttu-id="ebc0e-119">Internet Explorer 11 ブラウザーでアドインをテストするには、Office on the webでInternet Explorerを開き、アドインを[サイドロードします](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-119">To test your add-in on the Internet Explorer 11 browser, open Office on the web in Internet Explorer and [sideload the add-in](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="ebc0e-120">前提条件</span><span class="sxs-lookup"><span data-stu-id="ebc0e-120">Prerequisites</span></span>

- <span data-ttu-id="ebc0e-121">[Node.js](https://nodejs.org/) (最新 [LTS](https://nodejs.org/about/releases) バージョン)</span><span class="sxs-lookup"><span data-stu-id="ebc0e-121">[Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>

<span data-ttu-id="ebc0e-122">これらの手順では、以前に Yo ジェネレーター プロジェクトをOffice前提とします。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-122">These instructions assume you have set up a Yo Office generator project before.</span></span> <span data-ttu-id="ebc0e-123">前にこれを行ったことがない場合は、クイック スタート (アドイン用など) を読[Excel検討してください](../quickstarts/excel-quickstart-jquery.md)。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-123">If you haven't done this before, consider reading a quick start, such as [this one for Excel add-ins](../quickstarts/excel-quickstart-jquery.md).</span></span>

## <a name="switching-to-the-internet-explorer-11-webview"></a><span data-ttu-id="ebc0e-124">11 webview Internet Explorer切り替える</span><span class="sxs-lookup"><span data-stu-id="ebc0e-124">Switching to the Internet Explorer 11 webview</span></span>

1. <span data-ttu-id="ebc0e-125">Yo ジェネレーター プロジェクトOffice作成します。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-125">Create a Yo Office generator project.</span></span> <span data-ttu-id="ebc0e-126">選択するプロジェクトの種類は関係ありませんが、このツールは、すべてのプロジェクトの種類で動作します。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-126">It doesn't matter what kind of project you select, this tooling will work with all project types.</span></span>

    > [!NOTE]
    > <span data-ttu-id="ebc0e-127">既存のプロジェクトを持ち、新しいプロジェクトを作成せずにこのツールを追加する場合は、この手順をスキップして次の手順に進みます。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-127">If you have an existing project and want to add this tooling without creating a new project, skip this step and move to the next step.</span></span> 

1. <span data-ttu-id="ebc0e-128">プロジェクトのルート フォルダーで、コマンド ラインで次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-128">In the root folder of your project, run the following in the command line.</span></span> <span data-ttu-id="ebc0e-129">この例では、プロジェクトのマニフェスト ファイルがルートにあると仮定します。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-129">This example assumes that your project's manifest file is in the root.</span></span> <span data-ttu-id="ebc0e-130">指定されていない場合は、マニフェスト ファイルへの相対パスを指定します。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-130">If it isn't, specify the relative path to the manifest file.</span></span> <span data-ttu-id="ebc0e-131">コマンド ラインに、Web ビューの種類が IE に設定されているというメッセージが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-131">You should see a message in the command line that the web view type is now set to IE.</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

> [!TIP]
> <span data-ttu-id="ebc0e-132">このコマンドを使用する必要はありません。ただし、11 ランタイムに関連する問題の大部分をデバッグInternet Explorer必要があります。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-132">It isn't necessary to use this command, but it should help debug the majority of issues related to the Internet Explorer 11 runtime.</span></span> <span data-ttu-id="ebc0e-133">完全な堅牢性を得る場合は、Windows 7、8.1、および 10 とさまざまなバージョンの Office のさまざまな組み合わせのコンピューターを使用してテストする必要があります。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-133">For complete robustness, you should test using computers with various combinations of Windows 7, 8.1, and 10 and various versions of Office.</span></span> <span data-ttu-id="ebc0e-134">詳細については、「Office アドインで使用されるブラウザー」および「How to revert [to](../concepts/browsers-used-by-office-web-add-ins.md) earlier version of Office」 を[参照してください](https://support.microsoft.com/topic/how-to-revert-to-an-earlier-version-of-office-2bd5c457-a917-d57e-35a1-f709e3dda841)。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-134">For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md) and [How to revert to an earlier version of Office](https://support.microsoft.com/topic/how-to-revert-to-an-earlier-version-of-office-2bd5c457-a917-d57e-35a1-f709e3dda841).</span></span>

### <a name="command-options"></a><span data-ttu-id="ebc0e-135">コマンド オプション</span><span class="sxs-lookup"><span data-stu-id="ebc0e-135">Command options</span></span>

<span data-ttu-id="ebc0e-136">この `office-addin-dev-settings webview` コマンドは、引数として多数のランタイムを受け取る場合があります。</span><span class="sxs-lookup"><span data-stu-id="ebc0e-136">The `office-addin-dev-settings webview` command can also take a number of runtimes as arguments:</span></span>

- <span data-ttu-id="ebc0e-137">すなわち</span><span class="sxs-lookup"><span data-stu-id="ebc0e-137">ie</span></span>
- <span data-ttu-id="ebc0e-138">エッジ</span><span class="sxs-lookup"><span data-stu-id="ebc0e-138">edge</span></span>
- <span data-ttu-id="ebc0e-139">default</span><span class="sxs-lookup"><span data-stu-id="ebc0e-139">default</span></span>

## <a name="see-also"></a><span data-ttu-id="ebc0e-140">関連項目</span><span class="sxs-lookup"><span data-stu-id="ebc0e-140">See also</span></span>

* [<span data-ttu-id="ebc0e-141">Office アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="ebc0e-141">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)
* [<span data-ttu-id="ebc0e-142">テスト用に Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="ebc0e-142">Sideload Office Add-ins for testing</span></span>](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [<span data-ttu-id="ebc0e-143">Windows 10 で開発者ツールを使用してアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="ebc0e-143">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
* [<span data-ttu-id="ebc0e-144">作業ウィンドウからデバッガーをアタッチする</span><span class="sxs-lookup"><span data-stu-id="ebc0e-144">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)

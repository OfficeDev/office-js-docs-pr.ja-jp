---
title: Windows 10 で F12 開発者ツールを使用してアドインをデバッグする
description: ''
ms.date: 10/16/2018
ms.openlocfilehash: 3df245fcd651ec227e0a32d53da186ee332beb8f
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579843"
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a><span data-ttu-id="49a90-102">Windows 10 で F12 開発者ツールを使用してアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="49a90-102">Debug add-ins using F12 developer tools on Windows 10</span></span>

<span data-ttu-id="49a90-103">Windows 10 に含まれている F12 開発者ツールにより、web ページのデバッグ、テスト、および高速化ができます。</span><span class="sxs-lookup"><span data-stu-id="49a90-103">The F12 developer tools included in Windows 10 help you debug, test, and speed up your webpages.</span></span> <span data-ttu-id="49a90-104">それらを使用すれば、Visual Studio などの IDE を使用していない場合や、アドインを IDE の外部で実行中に問題を調査する必要がある場合に、Office アドインの開発とデバッグを行うこともできます。</span><span class="sxs-lookup"><span data-stu-id="49a90-104">The F12 developer tools included in Windows 10 help you debug, test, and speed up your webpages. You can also use them to develop and debug Office Add-ins, if you are not using an IDE like Visual Studio, or if you need to investigate a problem while running your add-in outside the IDE. You can start the F12 developer tools after your add-in is running.</span></span> <span data-ttu-id="49a90-105">この記事では、Windows 10 で F12 開発者ツールのデバッガー ツールを使用して、ご利用の Office アドインをテストする方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="49a90-105">This article describes how to use the Debugger tool from the F12 developer tools in Windows 10 to test your Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="49a90-106">この記事の手順を使用して、実行関数を使用する Outlook アドインをデバッグすることはできません。</span><span class="sxs-lookup"><span data-stu-id="49a90-106">The instructions in this article cannot be used to debug an Outlook add-in that uses Execute Functions.</span></span> <span data-ttu-id="49a90-107">実行関数を使用する Outlook アドインのデバッグには、スクリプト モードの Visual Studio またはその他のスクリプト デバッガーにアタッチすることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="49a90-107">To debug an Outlook add-in that uses Execute Functions, we recommend that you attach to Visual Studio in script mode or to some other script debugger.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="49a90-108">前提条件</span><span class="sxs-lookup"><span data-stu-id="49a90-108">Prerequisites</span></span>

<span data-ttu-id="49a90-109">以下のソフトウェアが必要です。</span><span class="sxs-lookup"><span data-stu-id="49a90-109">You need the following software:</span></span>

- <span data-ttu-id="49a90-110">Windows 10 に含まれる F12 開発者ツール。</span><span class="sxs-lookup"><span data-stu-id="49a90-110">The F12 developer tools, which are included in Windows 10.</span></span> 
    
- <span data-ttu-id="49a90-111">アドインをホストする Office クライアント アプリケーション。 </span><span class="sxs-lookup"><span data-stu-id="49a90-111">The Office client application that hosts your add-in.</span></span> 
    
- <span data-ttu-id="49a90-112">アドイン。 </span><span class="sxs-lookup"><span data-stu-id="49a90-112">Your add-in.</span></span> 

## <a name="using-the-debugger"></a><span data-ttu-id="49a90-113">デバッガーの使用</span><span class="sxs-lookup"><span data-stu-id="49a90-113">Using the Debugger</span></span>

<span data-ttu-id="49a90-114">Windows 10 の F12 開発者ツールからデバッガーを使用して、AppSource からのアドインやその他の場所から追加したアドインをテストすることができます。</span><span class="sxs-lookup"><span data-stu-id="49a90-114">This article shows how you how to use the Debugger tool from the F12 developer tools in Windows 10 to test your Office Add-in. You can test add-ins from AppSource or add-ins that you have added from other locations. The F12 tools display in a separate window and do not use Visual Studio.</span></span> <span data-ttu-id="49a90-115">アドインの実行後、F12 開発者ツールを起動できます。</span><span class="sxs-lookup"><span data-stu-id="49a90-115">You can start the F12 developer tools after your add-in is running.</span></span> <span data-ttu-id="49a90-116">F12 ツールは個別のウィンドウに表示され、Visual Studio を使用しません。</span><span class="sxs-lookup"><span data-stu-id="49a90-116">The F12 tools display in a separate window and do not use Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="49a90-p104">デバッガーは、Windows 10 および Internet Explorer 上の F12 開発者ツールの一部です。Windows の以前のバージョンにはデバッガーは含まれません。</span><span class="sxs-lookup"><span data-stu-id="49a90-p104">The Debugger is part of the F12 developer tools in Windows 10 and Internet Explorer. Earlier versions of Windows do not include the Debugger.</span></span> 

<span data-ttu-id="49a90-119">次の例では、AppSource から Word と無料のアドインを使用します。</span><span class="sxs-lookup"><span data-stu-id="49a90-119">This example uses Word and a free add-in from AppSource.</span></span>

1. <span data-ttu-id="49a90-120">Word を起動し、空白の文書を選択します。</span><span class="sxs-lookup"><span data-stu-id="49a90-120">Open Word and choose a blank document.</span></span> 
    
2. <span data-ttu-id="49a90-121">アドイン グループの [**挿入**] タブで、[**ストア**]、[**QR4Office**] アドインの順に選択します </span><span class="sxs-lookup"><span data-stu-id="49a90-121">On the Insert tab, in the Add-ins group, choose Store and select the QR4Office add-in. (You can load any add-in from the Store or your add-in catalog.)</span></span> <span data-ttu-id="49a90-122">(ストアやアドイン カタログから、任意のアドインを読み込むことができます)。</span><span class="sxs-lookup"><span data-stu-id="49a90-122">On the  Insert tab, in the Add-ins group, Store and select the QR4Office add-in. (You can load any add-in from the Store or your add-in catalog.)</span></span>
    
3. <span data-ttu-id="49a90-123">ご利用の Office のバージョンに対応する F12 開発者ツールを起動します。</span><span class="sxs-lookup"><span data-stu-id="49a90-123">Launch the F12 development tools that corresponds to your version of Office:</span></span>
    
   - <span data-ttu-id="49a90-124">32 ビット版の Office の場合は、C:\Windows\System32\F12\IEChooser.exe を使用します</span><span class="sxs-lookup"><span data-stu-id="49a90-124">For the 32-bit version of Office, use C:\Windows\System32\F12\F12Chooser.exe</span></span>
    
   - <span data-ttu-id="49a90-125">64 ビット版の Office の場合は、C:\Windows\SysWOW64\F12\IEChooser.exe を使用します</span><span class="sxs-lookup"><span data-stu-id="49a90-125">For the 64-bit version of Office, use C:\Windows\SysWOW64\F12\F12Chooser.exe</span></span>
    
   <span data-ttu-id="49a90-126">IEChooser を起動すると、[デバッグするターゲットの選択] という名前の別ウィンドウに、デバッグ可能なアプリケーションが表示されます。</span><span class="sxs-lookup"><span data-stu-id="49a90-126">When you launch F12Chooser, a separate window named "Choose target to debug" displays the possible applications to debug.</span></span> <span data-ttu-id="49a90-127">関心があるアプリケーションを選択します。</span><span class="sxs-lookup"><span data-stu-id="49a90-127">Select the application that you are interested in.</span></span> <span data-ttu-id="49a90-128">独自のアドインを記述している場合、アドインを展開した Web サイトを選択します。これは、localhost の URL である可能性があります。</span><span class="sxs-lookup"><span data-stu-id="49a90-128">If you are writing your own add-in, select the website where you have the add-in deployed, which might be a localhost URL.</span></span> 
    
   <span data-ttu-id="49a90-129">たとえば、**home.html** を選択します。</span><span class="sxs-lookup"><span data-stu-id="49a90-129">For example, select **home.html**.</span></span> 
    
   ![バブルのアドインをポイントする IEChooser 画面](../images/choose-target-to-debug.png)

4. <span data-ttu-id="49a90-131">F12 ウィンドウで、デバッグするファイルを選択します。</span><span class="sxs-lookup"><span data-stu-id="49a90-131">In the F12 window, select the file you want to debug.</span></span>
    
   <span data-ttu-id="49a90-132">F12 ウィンドウのファイルを選択するには、**スクリプト** (左側) ウィンドウの上にあるフォルダー アイコンを選びます。</span><span class="sxs-lookup"><span data-stu-id="49a90-132">To select the file, choose the folder icon above the  **script** (left) pane.</span></span> <span data-ttu-id="49a90-133">ドロップダウン リストに表示される利用可能なファイルのリストから [**Home.js**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="49a90-133">From the list of available files shown in the dropdown list, select **Home.js**.</span></span>
    
5. <span data-ttu-id="49a90-134">ブレークポイントを設定します。</span><span class="sxs-lookup"><span data-stu-id="49a90-134">Set the breakpoint.</span></span>
    
   <span data-ttu-id="49a90-135">**Home.js** にブレークポイントを設定するために、`textChanged` 関数内の行 144 を選択します。</span><span class="sxs-lookup"><span data-stu-id="49a90-135">To set the breakpoint in home.js, choose line 144, which is in the  textChanged function.</span></span> <span data-ttu-id="49a90-136">その行の左側と **[呼び出し履歴] と [ブレークポイント]** (右下) ウィンドウの対応する行に赤い点が表示されます。</span><span class="sxs-lookup"><span data-stu-id="49a90-136">You will see a red dot to the left of the line and a corresponding line in the **Callstack and Breakpoints** (bottom right) pane.</span></span> <span data-ttu-id="49a90-137">ブレークポイントを設定するその他の方法については、「[デバッガーを使用して実行中の JavaScript を検査する](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="49a90-137">For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span></span> 
    
   ![home.js ファイルのブレーキポイントを含むデバッガー](../images/debugger-home-js-02.png)

6. <span data-ttu-id="49a90-139">アドインを実行して、ブレークポイントをトリガーします。</span><span class="sxs-lookup"><span data-stu-id="49a90-139">Run your add-in to trigger the breakpoint.</span></span>
    
   <span data-ttu-id="49a90-140">Word で、[**QR4Office**] ウィンドウの上部にある [URL] テキスト ボックスを選択して、テキストを入力してみます。</span><span class="sxs-lookup"><span data-stu-id="49a90-140">In Word, choose the URL textbox in the upper part of the **QR4Office** pane and attempt to enter some text.</span></span> <span data-ttu-id="49a90-141">デバッガー内の **[呼び出し履歴] と [ブレークポイント]** ウィンドウで、ブレークポイントがトリガーされ、さまざまな情報が表示されることがわかります。</span><span class="sxs-lookup"><span data-stu-id="49a90-141">In the Debugger, in the  **Callstack and Breakpoints** pane, you'll see that the breakpoint has triggered and shows various information.</span></span> <span data-ttu-id="49a90-142">結果を確認するには、デバッガーの更新が必要な場合があります。</span><span class="sxs-lookup"><span data-stu-id="49a90-142">You might need to refresh the F12 tool to see the results.</span></span>
    
   ![トリガーされたブレークポイントの結果を含むデバッガー](../images/debugger-home-js-01.png)


## <a name="see-also"></a><span data-ttu-id="49a90-144">関連項目</span><span class="sxs-lookup"><span data-stu-id="49a90-144">See also</span></span>

- <span data-ttu-id="49a90-145">[デバッガーを使用して実行中の JavaScript を検査する](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="49a90-145">[Inspect running JavaScript with the Debugger](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span></span>
- <span data-ttu-id="49a90-146">
  [F12 開発者ツールの使用](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="49a90-146">[Using the F12 developer tools](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span></span>

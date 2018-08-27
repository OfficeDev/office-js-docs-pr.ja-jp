---
title: Windows 10 で F12 開発者ツールを使用してアドインをデバッグする
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 226773962fb1777a3a1f0e09445721ae2b8b5f5b
ms.sourcegitcommit: e1c92ba882e6eb03a165867c6021a6aa742aa310
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/20/2018
ms.locfileid: "22925606"
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a><span data-ttu-id="ed308-102">Windows 10 で F12 開発者ツールを使用してアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="ed308-102">Debug add-ins using F12 developer tools on Windows 10</span></span>

<span data-ttu-id="ed308-p101">Windows 10 に含まれている F12 開発者ツールにより、web ページのデバッグ、テスト、および高速化ができます。それらを使用すれば、Visual Studio などの IDE を使用していない場合や、アドインを IDE の外部で実行中に問題を調査する必要がある場合に、Office アドインの開発とデバッグを行うことも可能です。アドインの実行後、F12 開発者ツールを起動できます。</span><span class="sxs-lookup"><span data-stu-id="ed308-p101">The F12 developer tools included in Windows 10 help you debug, test, and speed up your webpages. You can also use them to develop and debug Office Add-ins, if you are not using an IDE like Visual Studio, or if you need to investigate a problem while running your add-in outside the IDE. You can start the F12 developer tools after your add-in is running.</span></span>

<span data-ttu-id="ed308-p102">この記事では、Windows 10 で F12 開発者ツールのデバッガー ツールを使用して、Office アドインをテストする方法を説明します。AppSource からのアドイン、また他の場所から追加したアドインもテストできます。F12 ツールは独自のウィンドウに表示され、Visual Studio を使用しません。</span><span class="sxs-lookup"><span data-stu-id="ed308-p102">This article shows how you how to use the Debugger tool from the F12 developer tools in Windows 10 to test your Office Add-in. You can test add-ins from AppSource or add-ins that you have added from other locations. The F12 tools display in a separate window and do not use Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="ed308-p103">デバッガーは、Windows 10 および Internet Explorer 上の F12 開発者ツールの一部です。Windows の以前のバージョンにはデバッガーは含まれません。</span><span class="sxs-lookup"><span data-stu-id="ed308-p103">The Debugger is part of the F12 developer tools in Windows 10 and Internet Explorer. Earlier versions of Windows do not include the Debugger.</span></span> 

## <a name="prerequisites"></a><span data-ttu-id="ed308-111">前提条件</span><span class="sxs-lookup"><span data-stu-id="ed308-111">Prerequisites</span></span>

<span data-ttu-id="ed308-112">以下のソフトウェアが必要です。</span><span class="sxs-lookup"><span data-stu-id="ed308-112">You need the following software:</span></span>

- <span data-ttu-id="ed308-113">Windows 10 に含まれる F12 開発者ツール。</span><span class="sxs-lookup"><span data-stu-id="ed308-113">The F12 developer tools, which are included in Windows 10.</span></span> 
    
- <span data-ttu-id="ed308-114">アドインをホストする Office クライアント アプリケーション。</span><span class="sxs-lookup"><span data-stu-id="ed308-114">The Office client application that hosts your add-in.</span></span> 
    
- <span data-ttu-id="ed308-115">アドイン。</span><span class="sxs-lookup"><span data-stu-id="ed308-115">Your add-in.</span></span> 

## <a name="using-the-debugger"></a><span data-ttu-id="ed308-116">デバッガーの使用</span><span class="sxs-lookup"><span data-stu-id="ed308-116">Using the Debugger</span></span>

<span data-ttu-id="ed308-117">次の例では、AppSource から Word と無料のアドインを使用します。</span><span class="sxs-lookup"><span data-stu-id="ed308-117">This example uses Word and a free add-in from AppSource.</span></span>

1. <span data-ttu-id="ed308-118">Word を起動し、空白の文書を選択します。</span><span class="sxs-lookup"><span data-stu-id="ed308-118">Open Word and choose a blank document.</span></span> 
    
2. <span data-ttu-id="ed308-p104">アドイン グループの **[挿入]** タブで、**[ストア]** をクリックし、QR4Office アドインを選択します。(ストアやアドイン カタログから、追加のアドインを読み込むことができます。)</span><span class="sxs-lookup"><span data-stu-id="ed308-p104">On the **Insert** tab, in the Add-ins group, choose **Store** and select the QR4Office add-in. (You can load any add-in from the Store or your add-in catalog.)</span></span>
    
3. <span data-ttu-id="ed308-121">Office のバージョンに対応する F12 開発者ツールを起動します。</span><span class="sxs-lookup"><span data-stu-id="ed308-121">Launch the F12 development tools that corresponds to your version of Office:</span></span>
    
   - <span data-ttu-id="ed308-122">32 ビット版の Office の場合は、C:\Windows\System32\F12\IEChooser.exe を使用します。</span><span class="sxs-lookup"><span data-stu-id="ed308-122">For the 32-bit version of Office, use C:\Windows\System32\F12\F12Chooser.exe</span></span>
    
   - <span data-ttu-id="ed308-123">64 ビット版の Office の場合は、C:\Windows\SysWOW64\F12\IEChooser.exe を使用します。</span><span class="sxs-lookup"><span data-stu-id="ed308-123">For the 64-bit version of Office, use C:\Windows\SysWOW64\F12\F12Chooser.exe</span></span>
    
   <span data-ttu-id="ed308-124">IEChooser を起動すると、[デバッグするターゲットの選択] という名前の別ウィンドウに、デバッグの対象になりうるアプリケーションが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ed308-124">When you launch F12Chooser, a separate window named "Choose target to debug" displays the possible applications to debug.</span></span> <span data-ttu-id="ed308-125">デバッグするアプリケーションを選択します。</span><span class="sxs-lookup"><span data-stu-id="ed308-125">Select the application that you are interested in.</span></span> <span data-ttu-id="ed308-126">独自のアドインを記述している場合、アドインを展開した Web サイトを選択します。ローカル ホストの URL も選択できます。</span><span class="sxs-lookup"><span data-stu-id="ed308-126">If you are writing your own add-in, select the website where you have the add-in deployed, which might be a localhost URL.</span></span> 
    
   <span data-ttu-id="ed308-127">たとえば、**home.html** を選択します。</span><span class="sxs-lookup"><span data-stu-id="ed308-127">For example, select **home.html**.</span></span> 
    
   ![IEChooser 画面、バブル アドインを示す](../images/choose-target-to-debug.png)

4. <span data-ttu-id="ed308-129">F12 ウィンドウで、デバッグするファイルを選択します。</span><span class="sxs-lookup"><span data-stu-id="ed308-129">In the F12 window, select the file you want to debug.</span></span>
    
   <span data-ttu-id="ed308-p106">ファイルを選択するには、**スクリプト** (左側) ウィンドウの上にあるフォルダー アイコンを選択します。ドロップダウン リストに利用可能なファイルが表示されます。home.js を選択します。</span><span class="sxs-lookup"><span data-stu-id="ed308-p106">To select the file, choose the folder icon above the  **script** (left) pane. The dropdown list shows the available files. Select home.js.</span></span>
    
5. <span data-ttu-id="ed308-133">ブレークポイントを設定します。</span><span class="sxs-lookup"><span data-stu-id="ed308-133">Set the breakpoint.</span></span>
    
   <span data-ttu-id="ed308-p107">home.js にブレークポイントを設定するには、_textChanged_ 関数内の行 144 を選択します。その行の左側と **[コールスタックとブレークポイント]** (右下) ウィンドウの対応する行に赤い点が表示されます。ブレークポイントを設定するその他の方法については、「[デバッガーを使用して実行中の JavaScript を検査する](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ed308-p107">To set the breakpoint in home.js, choose line 144, which is in the  _textChanged_ function. You will see a red dot to the left of the line and a corresponding line in the **Callstack and Breakpoints** (bottom right) pane. For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span></span> 
    
   ![home.js ファイルのブレークポイントを含むデバッガー](../images/debugger-home-js-02.png)

6. <span data-ttu-id="ed308-138">アドインを実行して、ブレークポイントをトリガーします。</span><span class="sxs-lookup"><span data-stu-id="ed308-138">Run your add-in to trigger the breakpoint.</span></span>
    
   <span data-ttu-id="ed308-p108">[QR4Office] ウィンドウの上部にある [URL] テキスト ボックスを選択して、テキストを変更します。デバッガー内の **[コールスタックとブレークポイント]** ウィンドウで、ブレークポイントがトリガーされ、さまざまな情報を示していることがわかります。結果を確認するには、F12 ツールの更新が必要な場合があります。</span><span class="sxs-lookup"><span data-stu-id="ed308-p108">Choose the URL textbox in the upper part of the QR4Office pane to change the text. In the Debugger, in the **Callstack and Breakpoints** pane, you'll see that the breakpoint has triggered and shows various information. You might need to refresh the F12 tool to see the results.</span></span>
    
   ![トリガーされるブレーキポイントの結果を持つデバッガー](../images/debugger-home-js-01.png)


## <a name="see-also"></a><span data-ttu-id="ed308-143">関連項目</span><span class="sxs-lookup"><span data-stu-id="ed308-143">See also</span></span>

- <span data-ttu-id="ed308-144">[デバッガーを使用して実行中の JavaScript を検査する](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="ed308-144">[Inspect running JavaScript with the Debugger](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span></span>
- <span data-ttu-id="ed308-145">[F12 開発者ツールの使用](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="ed308-145">[Using the F12 developer tools](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span></span>
    

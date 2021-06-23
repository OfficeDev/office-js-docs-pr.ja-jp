---
title: Windows 10 で開発者ツールを使用してアドインをデバッグする
description: Windows 10 で Microsoft Edge 開発者ツールを使用してアドインをデバッグする
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 41e7f2c8efb6406948c30522b56424ed7f9aa400
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076533"
---
# <a name="debug-add-ins-using-developer-tools-on-windows-10"></a><span data-ttu-id="eba7d-103">Windows 10 で開発者ツールを使用してアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="eba7d-103">Debug add-ins using developer tools on Windows 10</span></span>

<span data-ttu-id="eba7d-104">Windows 10 のアドインのデバッグに役立つ IDE の外部の開発者ツールがあります。</span><span class="sxs-lookup"><span data-stu-id="eba7d-104">There are developer tools outside of IDEs available to help you debug your add-ins on Windows 10.</span></span> <span data-ttu-id="eba7d-105">これは、IDE の外部でアドインを実行しているときに問題を調査する必要がある場合に便利です。</span><span class="sxs-lookup"><span data-stu-id="eba7d-105">These are useful when you need to investigate a problem while running your add-in outside the IDE.</span></span>

<span data-ttu-id="eba7d-106">使用するツールは、アドインが Microsoft Edge または Internet Explorer のどちらで実行されているかによって異なります。</span><span class="sxs-lookup"><span data-stu-id="eba7d-106">The tool that you use depends on whether the add-in is running in Microsoft Edge or Internet Explorer.</span></span> <span data-ttu-id="eba7d-107">これは、Windows 10 のバージョンとコンピューターにインストールされている Office のバージョンによって決まります。</span><span class="sxs-lookup"><span data-stu-id="eba7d-107">This is determined by the version of Windows 10 and the version of Office that are installed on the computer.</span></span> <span data-ttu-id="eba7d-108">開発用コンピューターで使用されているブラウザーを確認するには、「[Office アドインによって使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="eba7d-108">To determine which browser is being used on your development computer, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).</span></span>

> [!NOTE]
> <span data-ttu-id="eba7d-109">この記事の手順を使用して、実行関数を使用する Outlook アドインをデバッグすることはできません。</span><span class="sxs-lookup"><span data-stu-id="eba7d-109">The instructions in this article cannot be used to debug an Outlook add-in that uses Execute Functions.</span></span> <span data-ttu-id="eba7d-110">実行関数を使用する Outlook アドインのデバッグには、スクリプト モードの Visual Studio またはその他のスクリプト デバッガーにアタッチすることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="eba7d-110">To debug an Outlook add-in that uses Execute Functions, we recommend that you attach to Visual Studio in script mode or to some other script debugger.</span></span>

## <a name="when-the-add-in-is-running-in-microsoft-edge"></a><span data-ttu-id="eba7d-111">アドインが Microsoft Edge で実行されている場合</span><span class="sxs-lookup"><span data-stu-id="eba7d-111">When the add-in is running in Microsoft Edge</span></span>

[!include[Enable debugging on Microsoft Edge DevTools](../includes/enable-debugging-on-edge-devtools.md)]

### <a name="debug-using-microsoft-edge-devtools"></a><span data-ttu-id="eba7d-112">Microsoft Edge DevTools を使用してデバッグする</span><span class="sxs-lookup"><span data-stu-id="eba7d-112">Debug using Microsoft Edge DevTools</span></span>

<span data-ttu-id="eba7d-113">アドインが Microsoft Edge で実行されている場合は、[Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab) を使用できます。</span><span class="sxs-lookup"><span data-stu-id="eba7d-113">When the add-in is running in Microsoft Edge, you can use the [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab).</span></span>

1. <span data-ttu-id="eba7d-114">アドインを実行します。</span><span class="sxs-lookup"><span data-stu-id="eba7d-114">Run the add-in.</span></span>

2. <span data-ttu-id="eba7d-115">Microsoft Edge DevTools を実行します。</span><span class="sxs-lookup"><span data-stu-id="eba7d-115">Run the Microsoft Edge DevTools.</span></span>

3. <span data-ttu-id="eba7d-116">ツールで、**[ローカル]** タブを開きます。アドインの名前が一覧表示されます。</span><span class="sxs-lookup"><span data-stu-id="eba7d-116">In the tools, open the **Local** tab. Your add-in will be listed by its name.</span></span>

4. <span data-ttu-id="eba7d-117">アドイン名をクリックして、ツールで開きます。</span><span class="sxs-lookup"><span data-stu-id="eba7d-117">Click the add-in name to open it in the tools.</span></span>

5. <span data-ttu-id="eba7d-118">**[デバッガー]** タブを開きます。</span><span class="sxs-lookup"><span data-stu-id="eba7d-118">Open the **Debugger** tab.</span></span> 

6. <span data-ttu-id="eba7d-119">**[スクリプト]** (左側) ウィンドウの上にあるフォルダー アイコンを選択します。</span><span class="sxs-lookup"><span data-stu-id="eba7d-119">Choose the folder icon above the **script** (left) pane.</span></span> <span data-ttu-id="eba7d-120">ドロップダウン リストに表示される利用可能なファイルのリストから、デバッグする JavaScript ファイルを選択します。</span><span class="sxs-lookup"><span data-stu-id="eba7d-120">From the list of available files shown in the dropdown list, select the JavaScript file that you want to debug.</span></span>

7. <span data-ttu-id="eba7d-121">ブレークポイントを設定するには、行を選択します。</span><span class="sxs-lookup"><span data-stu-id="eba7d-121">To set a breakpoint, select the line.</span></span> <span data-ttu-id="eba7d-122">その行の左側と **[呼び出し履歴]** (右下) ウィンドウの対応する行に赤い点が表示されます。</span><span class="sxs-lookup"><span data-stu-id="eba7d-122">You will see a red dot to the left of the line and a corresponding line in the **Call stack** (bottom right) pane.</span></span>

8. <span data-ttu-id="eba7d-123">必要に応じてアドインの関数を実行して、ブレークポイントをトリガーします。</span><span class="sxs-lookup"><span data-stu-id="eba7d-123">Execute functions in the add-in as needed to trigger the breakpoint.</span></span>

## <a name="when-the-add-in-is-running-in-internet-explorer"></a><span data-ttu-id="eba7d-124">アドインが Internet Explorer で実行されている場合</span><span class="sxs-lookup"><span data-stu-id="eba7d-124">When the add-in is running in Internet Explorer</span></span>

<span data-ttu-id="eba7d-125">Internet Explorer でアドインを実行している場合は、Windows 10 の F12 開発者ツールのデバッガーを使用して、アドインをテストできます。</span><span class="sxs-lookup"><span data-stu-id="eba7d-125">When the add-in is running in Internet Explorer, you can use the debugger from the F12 developer tools in Windows 10 to test your add-in.</span></span> <span data-ttu-id="eba7d-126">アドインの実行後、F12 開発者ツールを起動できます。</span><span class="sxs-lookup"><span data-stu-id="eba7d-126">You can start the F12 developer tools after the add-in is running.</span></span> <span data-ttu-id="eba7d-127">F12 ツールは個別のウィンドウに表示され、Visual Studio を使用しません。</span><span class="sxs-lookup"><span data-stu-id="eba7d-127">The F12 tools are displayed in a separate window and do not use Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="eba7d-p107">デバッガーは、Windows 10 および Internet Explorer 上の F12 開発者ツールの一部です。Windows の以前のバージョンにはデバッガーは含まれません。</span><span class="sxs-lookup"><span data-stu-id="eba7d-p107">The Debugger is part of the F12 developer tools in Windows 10 and Internet Explorer. Earlier versions of Windows do not include the Debugger.</span></span> 

<span data-ttu-id="eba7d-130">次の例では、AppSource から Word と無料のアドインを使用します。</span><span class="sxs-lookup"><span data-stu-id="eba7d-130">This example uses Word and a free add-in from AppSource.</span></span>

1. <span data-ttu-id="eba7d-131">Word を起動し、空白の文書を選択します。</span><span class="sxs-lookup"><span data-stu-id="eba7d-131">Open Word and choose a blank document.</span></span> 
    
2. <span data-ttu-id="eba7d-132">アドイン グループの [**挿入**] タブで、[**ストア**]、[**QR4Office**] アドインの順に選択します </span><span class="sxs-lookup"><span data-stu-id="eba7d-132">On the **Insert** tab, in the Add-ins group, choose **Store** and select the **QR4Office** Add-in.</span></span> <span data-ttu-id="eba7d-133">(ストアやアドイン カタログから、任意のアドインを読み込むことができます)。</span><span class="sxs-lookup"><span data-stu-id="eba7d-133">(You can load any add-in from the Store or your add-in catalog.)</span></span>
    
3. <span data-ttu-id="eba7d-134">ご利用の Office のバージョンに対応する F12 開発者ツールを起動します。</span><span class="sxs-lookup"><span data-stu-id="eba7d-134">Launch the F12 development tools that corresponds to your version of Office:</span></span>
    
   - <span data-ttu-id="eba7d-135">32 ビット版の Office の場合は、C:\Windows\System32\F12\IEChooser.exe を使用します</span><span class="sxs-lookup"><span data-stu-id="eba7d-135">For the 32-bit version of Office, use C:\Windows\System32\F12\IEChooser.exe</span></span>
    
   - <span data-ttu-id="eba7d-136">64 ビット版の Office の場合は、C:\Windows\SysWOW64\F12\IEChooser.exe を使用します</span><span class="sxs-lookup"><span data-stu-id="eba7d-136">For the 64-bit version of Office, use C:\Windows\SysWOW64\F12\IEChooser.exe</span></span>
    
   <span data-ttu-id="eba7d-137">IEChooser を起動すると、[デバッグするターゲットの選択] という名前の別ウィンドウに、デバッグ可能なアプリケーションが表示されます。</span><span class="sxs-lookup"><span data-stu-id="eba7d-137">When you launch IEChooser, a separate window named "Choose target to debug" displays the possible applications to debug.</span></span> <span data-ttu-id="eba7d-138">関心があるアプリケーションを選択します。</span><span class="sxs-lookup"><span data-stu-id="eba7d-138">Select the application that you are interested in.</span></span> <span data-ttu-id="eba7d-139">独自のアドインを記述している場合、アドインを展開した Web サイトを選択します。これは、localhost の URL である可能性があります。</span><span class="sxs-lookup"><span data-stu-id="eba7d-139">If you are writing your own add-in, select the website where you have the add-in deployed, which might be a localhost URL.</span></span> 
    
   <span data-ttu-id="eba7d-140">たとえば、**home.html** を選択します。</span><span class="sxs-lookup"><span data-stu-id="eba7d-140">For example, select **home.html**.</span></span> 
    
   ![[IEChooser] 画面で、バブル アドインをポイントします。](../images/choose-target-to-debug.png)

4. <span data-ttu-id="eba7d-142">F12 ウィンドウで、デバッグするファイルを選択します。</span><span class="sxs-lookup"><span data-stu-id="eba7d-142">In the F12 window, select the file you want to debug.</span></span>
    
   <span data-ttu-id="eba7d-143">F12 ウィンドウのファイルを選択するには、**スクリプト** (左側) ウィンドウの上にあるフォルダー アイコンを選びます。</span><span class="sxs-lookup"><span data-stu-id="eba7d-143">To select the file in the F12 window, choose the folder icon above the **script** (left) pane.</span></span> <span data-ttu-id="eba7d-144">ドロップダウン リストに表示される利用可能なファイルのリストから [**Home.js**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="eba7d-144">From the list of available files shown in the dropdown list, select **Home.js**.</span></span>
    
5. <span data-ttu-id="eba7d-145">ブレークポイントを設定します。</span><span class="sxs-lookup"><span data-stu-id="eba7d-145">Set the breakpoint.</span></span>
    
   <span data-ttu-id="eba7d-146">**Home.js** にブレークポイントを設定するために、`textChanged` 関数内の行 144 を選択します。</span><span class="sxs-lookup"><span data-stu-id="eba7d-146">To set the breakpoint in **Home.js**, choose line 144, which is in the  `textChanged` function.</span></span> <span data-ttu-id="eba7d-147">その行の左側と **[呼び出し履歴] と [ブレークポイント]** (右下) ウィンドウの対応する行に赤い点が表示されます。</span><span class="sxs-lookup"><span data-stu-id="eba7d-147">You will see a red dot to the left of the line and a corresponding line in the **Call stack and Breakpoints** (bottom right) pane.</span></span> <span data-ttu-id="eba7d-148">ブレークポイントを設定するその他の方法については、「[デバッガーを使用して実行中の JavaScript を検査する](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="eba7d-148">For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span></span> 
    
   ![ファイル内のブレークポイントを持つデバッガーhome.jsします。](../images/debugger-home-js-02.png)

6. <span data-ttu-id="eba7d-150">アドインを実行して、ブレークポイントをトリガーします。</span><span class="sxs-lookup"><span data-stu-id="eba7d-150">Run your add-in to trigger the breakpoint.</span></span>
    
   <span data-ttu-id="eba7d-151">Word で、[**QR4Office**] ウィンドウの上部にある [URL] テキスト ボックスを選択して、テキストを入力してみます。</span><span class="sxs-lookup"><span data-stu-id="eba7d-151">In Word, choose the URL textbox in the upper part of the **QR4Office** pane and attempt to enter some text.</span></span> <span data-ttu-id="eba7d-152">デバッガー内の **[呼び出し履歴] と [ブレークポイント]** ウィンドウで、ブレークポイントがトリガーされ、さまざまな情報が表示されることがわかります。</span><span class="sxs-lookup"><span data-stu-id="eba7d-152">In the Debugger, in the **Call stack and Breakpoints** pane, you'll see that the breakpoint has triggered and shows various information.</span></span> <span data-ttu-id="eba7d-153">結果を確認するには、デバッガーの更新が必要な場合があります。</span><span class="sxs-lookup"><span data-stu-id="eba7d-153">You might need to refresh the Debugger to see the results.</span></span>
    
   ![トリガーされたブレークポイントからの結果を含むデバッガー。](../images/debugger-home-js-01.png)


## <a name="see-also"></a><span data-ttu-id="eba7d-155">関連項目</span><span class="sxs-lookup"><span data-stu-id="eba7d-155">See also</span></span>

- <span data-ttu-id="eba7d-156">[デバッガーを使用して実行中の JavaScript を検査する](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="eba7d-156">[Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span></span>
- <span data-ttu-id="eba7d-157">[F12 開発者ツールの使用](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="eba7d-157">[Using the F12 developer tools](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span></span>

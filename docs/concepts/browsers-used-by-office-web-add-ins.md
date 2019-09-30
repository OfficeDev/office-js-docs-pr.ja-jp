---
title: Office アドインによって使用されるブラウザー
description: Office アドインによって使用されるブラウザーをオペレーティング システムおよび Office バージョンが決定する方法を指定します。
ms.date: 09/25/2019
localization_priority: Priority
ms.openlocfilehash: b5d7198e556f020bccdf7ba1e0a0fcffa3a9171b
ms.sourcegitcommit: c8914ce0f48a0c19bbfc3276a80d090bb7ce68e1
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/26/2019
ms.locfileid: "37235296"
---
# <a name="browsers-used-by-office-add-ins"></a><span data-ttu-id="39293-103">Office アドインによって使用されるブラウザー</span><span class="sxs-lookup"><span data-stu-id="39293-103">Web viewers used by Office Add-ins</span></span>

<span data-ttu-id="39293-104">Office アドインは、Office on the web での実行時に iFrame を使用して表示され、デスクトップおよびモバイル クライアント用に Office に埋め込まれたブラウザー コントロールを使用して表示される Web アプリケーションです。</span><span class="sxs-lookup"><span data-stu-id="39293-104">Office add-ins are web applications that are displayed using iFrames when running in Office on the web and using embedded browser controls in Office for desktop and mobile clients.</span></span> <span data-ttu-id="39293-105">アドインには JavaScript を実行するための JavaScript エンジンも必要です。</span><span class="sxs-lookup"><span data-stu-id="39293-105">Add-ins also need a JavaScript engine to run the JavaScript.</span></span> <span data-ttu-id="39293-106">埋め込まれたブラウザーおよびエンジン、どちらもユーザーのコンピュータにインストールされているブラウザによって提供されます。</span><span class="sxs-lookup"><span data-stu-id="39293-106">Both the embedded browser and the engine are supplied by a browser installed on the user’s computer.</span></span>

<span data-ttu-id="39293-107">どのブラウザが使用されているかは、以下によります。</span><span class="sxs-lookup"><span data-stu-id="39293-107">Which browser is used depends on:</span></span>

- <span data-ttu-id="39293-108">コンピュータのオペレーティングシステム。</span><span class="sxs-lookup"><span data-stu-id="39293-108">The computer’s operating system.</span></span>
- <span data-ttu-id="39293-109">アドインが Office on the web、Office 365、または登録のない Office 2013 以降で実行されているかどうか。</span><span class="sxs-lookup"><span data-stu-id="39293-109">Whether the add-in is running in Office Online, Office 365, or non-subscription Office 2013 or later.</span></span>

<span data-ttu-id="39293-110">次の表は、さまざまなプラットフォームとオペレーティングシステムに使用されているブラウザを示しています。</span><span class="sxs-lookup"><span data-stu-id="39293-110">The following table shows which browser is used for the various platforms and operating systems.</span></span>

|<span data-ttu-id="39293-111">**OS / Platform**</span><span class="sxs-lookup"><span data-stu-id="39293-111">**OS / Platform**</span></span>|<span data-ttu-id="39293-112">**Browser**</span><span class="sxs-lookup"><span data-stu-id="39293-112">**Browser**</span></span>|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|<span data-ttu-id="39293-113">Office on the web</span><span class="sxs-lookup"><span data-stu-id="39293-113">Office on the web</span></span>|<span data-ttu-id="39293-114">Office が開かれているブラウザー。</span><span class="sxs-lookup"><span data-stu-id="39293-114">The browser in which Office Online is opened.</span></span>|
|<span data-ttu-id="39293-115">Mac</span><span class="sxs-lookup"><span data-stu-id="39293-115">Mac</span></span>|<span data-ttu-id="39293-116">Safari</span><span class="sxs-lookup"><span data-stu-id="39293-116">Safari</span></span>|
|<span data-ttu-id="39293-117">iOS</span><span class="sxs-lookup"><span data-stu-id="39293-117">iOS</span></span>|<span data-ttu-id="39293-118">Safari</span><span class="sxs-lookup"><span data-stu-id="39293-118">Safari</span></span>|
|<span data-ttu-id="39293-119">Android</span><span class="sxs-lookup"><span data-stu-id="39293-119">Android</span></span>|<span data-ttu-id="39293-120">Chrome</span><span class="sxs-lookup"><span data-stu-id="39293-120">Chrome</span></span>|
|<span data-ttu-id="39293-121">Windows / 非登録 Office 2013以降</span><span class="sxs-lookup"><span data-stu-id="39293-121">Windows / non-subscription Office 2013 or later</span></span>|<span data-ttu-id="39293-122">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="39293-122">Internet Explorer 11</span></span>|
|<span data-ttu-id="39293-123">Windows 10 バージョン</span><span class="sxs-lookup"><span data-stu-id="39293-123">Windows 10 ver.</span></span> <span data-ttu-id="39293-124">< 1903 / Office 365</span><span class="sxs-lookup"><span data-stu-id="39293-124">< 1903 / Office 365</span></span>|<span data-ttu-id="39293-125">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="39293-125">Internet Explorer 11</span></span>|
|<span data-ttu-id="39293-126">Windows 10 バージョン</span><span class="sxs-lookup"><span data-stu-id="39293-126">Windows 10 ver.</span></span> <span data-ttu-id="39293-127">>= 1903 / Office 365 ver < 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="39293-127">>= 1903 / Office 365 ver < 16.0.11629</span></span>|<span data-ttu-id="39293-128">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="39293-128">Internet Explorer 11</span></span>|
|<span data-ttu-id="39293-129">Windows 10 バージョン</span><span class="sxs-lookup"><span data-stu-id="39293-129">Windows 10 ver.</span></span> <span data-ttu-id="39293-130">>= 1903 / Office 365 ver >= 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="39293-130">>= 1903 / Office 365 ver >= 16.0.11629</span></span>|<span data-ttu-id="39293-131">Microsoft Edge\*</span><span class="sxs-lookup"><span data-stu-id="39293-131">Microsoft Edge\*</span></span>|

<span data-ttu-id="39293-132">\*Microsoft Edge が使用されている場合、Windows 10 ナレーター (「スクリーン リーダー」と呼ばれることもあります) は、作業ウィンドウで開いているページの `<title>` タグを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="39293-132">\* When Edge is being used, the Windows 10 Narrator (sometimes called a "screen reader") reads the `<title>` tag in the page that opens in the task pane.</span></span> <span data-ttu-id="39293-133">Internet Explorer 11 が使用されている場合、ナレーターはアドイン マニフェストの `<DisplayName>` の値から提供される作業ウィンドウのタイトル バーを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="39293-133">When Internet Explorer 11 is being used, the Narrator reads the title bar of the task pane, which comes from the `<DisplayName>` value in the add-in's manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="39293-134">Internet Explorer 11はES5以降のJavaScriptバージョンをサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="39293-134">Internet Explorer 11 does not support JavaScript versions later than ES5.</span></span> <span data-ttu-id="39293-135">アドインのユーザーが Internet Explorer 11 を使用するプラットフォームを使用している場合、ECMAScript 2015 以降の構文と機能を使用するには、JavaScript を ES 5 にトランスパイルするか、ポリフィルを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="39293-135">If any of your add-in's users have platforms that use Internet Explorer 11, then to use the syntax and features of ECMAScript 2015 or later, you will need to either transpile your JavaScript to ES5 or use a polyfill.</span></span> <span data-ttu-id="39293-136">また、Internet Explorer 11 は、メディア、録音、および位置情報などの HTML 5 機能の一部をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="39293-136">Also, Internet Explorer 11 does not support some HTML 5 features such as media, recording, and location.</span></span>

> [!NOTE]
> <span data-ttu-id="39293-137">これらが一般に利用可能になるまで、Windows バージョン 1903 以降を入手するには Windows Insider である必要があり、また、Office バージョン 16.0.11629 以降を入手するには Office Insider である必要があります。</span><span class="sxs-lookup"><span data-stu-id="39293-137">Until they are generally available, you need to be a Windows Insider to get a Windows version 1903 or greater, and you need to be an Office Insider to get Office version 16.0.11629 or greater.</span></span>
>
> <span data-ttu-id="39293-138">Windows インサイダーに参加するには</span><span class="sxs-lookup"><span data-stu-id="39293-138">To join Windows Insiders:</span></span>
> 
> 1. <span data-ttu-id="39293-139">[Windows インサイダー](https://insider.windows.com)に移動し、リンクをクリックしてWindows インサイダーに参加してください。</span><span class="sxs-lookup"><span data-stu-id="39293-139">Go to [Windows Insider](https://insider.windows.com) and click the link to join Windows Insiders.</span></span>
> 2. <span data-ttu-id="39293-140">Windowsのプレビュービルドを有効にするためのWindowsの設定の使用方法についての説明が記載されたページに移動します。</span><span class="sxs-lookup"><span data-stu-id="39293-140">You will be taken to a page with instructions about how to use Windows Settings to enable preview builds of Windows.</span></span> <span data-ttu-id="39293-141">指示に従います。</span><span class="sxs-lookup"><span data-stu-id="39293-141">Follow the instructions.</span></span> <span data-ttu-id="39293-142">更新頻度を選択する際は、一番速いオプションを選択してください。</span><span class="sxs-lookup"><span data-stu-id="39293-142">When you select the pace of updates, choose the fastest option.</span></span>
>
> <span data-ttu-id="39293-143">Office インサイダーに参加するには</span><span class="sxs-lookup"><span data-stu-id="39293-143">To join Office Insiders:</span></span>
> 
> 1. <span data-ttu-id="39293-144">[Office Insiderになりましょう](https://insider.office.com/join)に移動してください。</span><span class="sxs-lookup"><span data-stu-id="39293-144">Go to [Get started as an Office Insider](https://insider.office.com/join).</span></span>
> 2. <span data-ttu-id="39293-145">そのページの指示に従って参加してください。</span><span class="sxs-lookup"><span data-stu-id="39293-145">Follow the instruction on that page to join.</span></span> <span data-ttu-id="39293-146">チャンネルを指定するように求められたら、[インサイダー]を選択します。</span><span class="sxs-lookup"><span data-stu-id="39293-146">When asked to specify a channel, select Insider.</span></span>

## <a name="troubleshooting-microsoft-edge-issues"></a><span data-ttu-id="39293-147">Microsoft Edge の問題のトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="39293-147">Troubleshooting Microsoft Edge Issues</span></span>

### <a name="scroll-bar-does-not-appear-in-task-pane"></a><span data-ttu-id="39293-148">作業ウィンドウにスクロール バーが表示されない</span><span class="sxs-lookup"><span data-stu-id="39293-148">Scroll bar does not appear in task pane</span></span>

<span data-ttu-id="39293-149">既定では、Microsoft Edge のスクロール バーはホバーするまで非表示になっています。</span><span class="sxs-lookup"><span data-stu-id="39293-149">By default, scroll bars in Microsoft Edge are hidden until hovered over.</span></span> <span data-ttu-id="39293-150">スクロールバーが常に表示されるようにするには、作業ウィンドウのページの`<body>`要素に適用される CSS スタイルに [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/-ms-overflow-style) プロパティを含め、`scrollbar`に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="39293-150">To ensure that the scroll bar is always visible, the CSS styling that applies to the `<body>` element of the pages in the task pane should include the [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/-ms-overflow-style) property and it should be set to `scrollbar`.</span></span> 

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a><span data-ttu-id="39293-151">Microsoft Edge DevTools を使用してデバッグすると、アドインがクラッシュまたは再読み込みされる</span><span class="sxs-lookup"><span data-stu-id="39293-151">When debugging with the Microsoft Edge DevTools, the add-in crashes or reloads</span></span>

<span data-ttu-id="39293-152">[Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) にブレークポイントを設定すると、アドインがハングしていると Office に判断される可能性があります。</span><span class="sxs-lookup"><span data-stu-id="39293-152">Setting breakpoints in the [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) can cause Office to think that the add-in is hung.</span></span> <span data-ttu-id="39293-153">これが発生すると、アドインが自動的に再読み込みされます。</span><span class="sxs-lookup"><span data-stu-id="39293-153">It will automatically reload the add-in when this happens.</span></span> <span data-ttu-id="39293-154">これを防ぐには、開発用コンピューターに以下のレジストリ キーと値を追加します: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`。</span><span class="sxs-lookup"><span data-stu-id="39293-154">To prevent this, add the following Registry key and value to the development computer: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`.</span></span>

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a><span data-ttu-id="39293-155">アドインを開こうとすると、"アドイン エラー localhost からこのアドインを開けません" というエラーが表示される</span><span class="sxs-lookup"><span data-stu-id="39293-155">When the add-in tries to open, get "ADD-IN ERROR We can't open this add-in from the localhost" error</span></span>

<span data-ttu-id="39293-156">既知の原因の1つとして、Microsoft Edge では開発用コンピューター上では localhost にループバックの除外を与える必要があることが挙げられます。</span><span class="sxs-lookup"><span data-stu-id="39293-156">One known cause is that Microsoft Edge requires that localhost be given a loopback exemption on the development computer.</span></span> <span data-ttu-id="39293-157">[Cannot open add-in from localhost (localhostからアドインを開くことができません)](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost)の指示に従ってください。</span><span class="sxs-lookup"><span data-stu-id="39293-157">Follow the instructions at [Cannot open add-in from localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).</span></span>


## <a name="see-also"></a><span data-ttu-id="39293-158">関連項目</span><span class="sxs-lookup"><span data-stu-id="39293-158">See also</span></span>

- [<span data-ttu-id="39293-159">Officeアドインを実行するための要件</span><span class="sxs-lookup"><span data-stu-id="39293-159">Requirements for Running Office Add-ins</span></span>](requirements-for-running-office-add-ins.md)

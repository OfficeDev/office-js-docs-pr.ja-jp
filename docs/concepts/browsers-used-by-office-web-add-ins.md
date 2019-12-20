---
title: Office アドインによって使用されるブラウザー
description: Office アドインによって使用されるブラウザーをオペレーティング システムおよび Office バージョンが決定する方法を指定します。
ms.date: 12/13/2019
localization_priority: Priority
ms.openlocfilehash: 3709157449634dfb49805e2247e47debe60f468f
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40813985"
---
# <a name="browsers-used-by-office-add-ins"></a><span data-ttu-id="1fa34-103">Office アドインによって使用されるブラウザー</span><span class="sxs-lookup"><span data-stu-id="1fa34-103">Browsers used by Office Add-ins</span></span>

<span data-ttu-id="1fa34-104">Office アドインは、Office on the web での実行時に iFrame を使用して表示され、デスクトップおよびモバイル クライアント用に Office に埋め込まれたブラウザー コントロールを使用して表示される Web アプリケーションです。</span><span class="sxs-lookup"><span data-stu-id="1fa34-104">Office add-ins are web applications that are displayed using iFrames when running in Office on the web and using embedded browser controls in Office for desktop and mobile clients.</span></span> <span data-ttu-id="1fa34-105">アドインには JavaScript を実行するための JavaScript エンジンも必要です。</span><span class="sxs-lookup"><span data-stu-id="1fa34-105">Add-ins also need a JavaScript engine to run the JavaScript.</span></span> <span data-ttu-id="1fa34-106">埋め込まれたブラウザーおよびエンジン、どちらもユーザーのコンピュータにインストールされているブラウザによって提供されます。</span><span class="sxs-lookup"><span data-stu-id="1fa34-106">Both the embedded browser and the engine are supplied by a browser installed on the user’s computer.</span></span>

<span data-ttu-id="1fa34-107">どのブラウザが使用されているかは、以下によります。</span><span class="sxs-lookup"><span data-stu-id="1fa34-107">Which browser is used depends on:</span></span>

- <span data-ttu-id="1fa34-108">コンピュータのオペレーティングシステム。</span><span class="sxs-lookup"><span data-stu-id="1fa34-108">The computer’s operating system.</span></span>
- <span data-ttu-id="1fa34-109">アドインが Office on the web、Office 365、または登録のない Office 2013 以降で実行されているかどうか。</span><span class="sxs-lookup"><span data-stu-id="1fa34-109">Whether the add-in is running in Office on the web, Office 365, or non-subscription Office 2013 or later.</span></span>

<span data-ttu-id="1fa34-110">次の表は、さまざまなプラットフォームとオペレーティングシステムに使用されているブラウザを示しています。</span><span class="sxs-lookup"><span data-stu-id="1fa34-110">The following table shows which browser is used for the various platforms and operating systems.</span></span>

|<span data-ttu-id="1fa34-111">**OS / Platform**</span><span class="sxs-lookup"><span data-stu-id="1fa34-111">**OS / Platform**</span></span>|<span data-ttu-id="1fa34-112">**Browser**</span><span class="sxs-lookup"><span data-stu-id="1fa34-112">**Browser**</span></span>|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|<span data-ttu-id="1fa34-113">Office on the web</span><span class="sxs-lookup"><span data-stu-id="1fa34-113">Office on the web</span></span>|<span data-ttu-id="1fa34-114">Office が開かれているブラウザー。</span><span class="sxs-lookup"><span data-stu-id="1fa34-114">The browser in which Office is opened.</span></span>|
|<span data-ttu-id="1fa34-115">Mac</span><span class="sxs-lookup"><span data-stu-id="1fa34-115">Mac</span></span>|<span data-ttu-id="1fa34-116">Safari</span><span class="sxs-lookup"><span data-stu-id="1fa34-116">Safari</span></span>|
|<span data-ttu-id="1fa34-117">iOS</span><span class="sxs-lookup"><span data-stu-id="1fa34-117">iOS</span></span>|<span data-ttu-id="1fa34-118">Safari</span><span class="sxs-lookup"><span data-stu-id="1fa34-118">Safari</span></span>|
|<span data-ttu-id="1fa34-119">Android</span><span class="sxs-lookup"><span data-stu-id="1fa34-119">Android</span></span>|<span data-ttu-id="1fa34-120">Chrome</span><span class="sxs-lookup"><span data-stu-id="1fa34-120">Chrome</span></span>|
|<span data-ttu-id="1fa34-121">Windows / 非登録 Office 2013以降</span><span class="sxs-lookup"><span data-stu-id="1fa34-121">Windows / non-subscription Office 2013 or later</span></span>|<span data-ttu-id="1fa34-122">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="1fa34-122">Internet Explorer 11</span></span>|
|<span data-ttu-id="1fa34-123">Windows 10 バージョン</span><span class="sxs-lookup"><span data-stu-id="1fa34-123">Windows 10 ver.</span></span> <span data-ttu-id="1fa34-124">< 1903 / Office 365</span><span class="sxs-lookup"><span data-stu-id="1fa34-124">< 1903 / Office 365</span></span>|<span data-ttu-id="1fa34-125">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="1fa34-125">Internet Explorer 11</span></span>|
|<span data-ttu-id="1fa34-126">Windows 10 バージョン</span><span class="sxs-lookup"><span data-stu-id="1fa34-126">Windows 10 ver.</span></span> <span data-ttu-id="1fa34-127">>= 1903 / Office 365 ver < 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="1fa34-127">>= 1903 / Office 365 ver < 16.0.11629</span></span>|<span data-ttu-id="1fa34-128">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="1fa34-128">Internet Explorer 11</span></span>|
|<span data-ttu-id="1fa34-129">Windows 10 バージョン</span><span class="sxs-lookup"><span data-stu-id="1fa34-129">Windows 10 ver.</span></span> <span data-ttu-id="1fa34-130">>= 1903 / Office 365 ver >= 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="1fa34-130">>= 1903 / Office 365 ver >= 16.0.11629</span></span>|<span data-ttu-id="1fa34-131">Microsoft Edge\*</span><span class="sxs-lookup"><span data-stu-id="1fa34-131">Microsoft Edge\*</span></span>|

<span data-ttu-id="1fa34-132">\*Microsoft Edge が使用されている場合、Windows 10 ナレーター (「スクリーン リーダー」と呼ばれることもあります) は、作業ウィンドウで開いているページの `<title>` タグを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="1fa34-132">\* When Microsoft Edge is being used, the Windows 10 Narrator (sometimes called a "screen reader") reads the `<title>` tag in the page that opens in the task pane.</span></span> <span data-ttu-id="1fa34-133">Internet Explorer 11 が使用されている場合、ナレーターはアドイン マニフェストの `<DisplayName>` の値から提供される作業ウィンドウのタイトル バーを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="1fa34-133">When Internet Explorer 11 is being used, the Narrator reads the title bar of the task pane, which comes from the `<DisplayName>` value in the add-in's manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1fa34-134">Internet Explorer 11はES5以降のJavaScriptバージョンをサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="1fa34-134">Internet Explorer 11 does not support JavaScript versions later than ES5.</span></span> <span data-ttu-id="1fa34-135">アドインのユーザーが Internet Explorer 11 を使用するプラットフォームを使用している場合、ECMAScript 2015 以降の構文と機能を使用するには、JavaScript を ES 5 にトランスパイルするか、ポリフィルを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1fa34-135">If any of your add-in's users have platforms that use Internet Explorer 11, then to use the syntax and features of ECMAScript 2015 or later, you will need to either transpile your JavaScript to ES5 or use a polyfill.</span></span> <span data-ttu-id="1fa34-136">また、Internet Explorer 11 は、メディア、録音、および位置情報などの HTML 5 機能の一部をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="1fa34-136">Also, Internet Explorer 11 does not support some HTML5 features such as media, recording, and location.</span></span>

## <a name="troubleshooting-microsoft-edge-issues"></a><span data-ttu-id="1fa34-137">Microsoft Edge の問題のトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="1fa34-137">Troubleshooting Microsoft Edge Issues</span></span>

### <a name="scroll-bar-does-not-appear-in-task-pane"></a><span data-ttu-id="1fa34-138">作業ウィンドウにスクロール バーが表示されない</span><span class="sxs-lookup"><span data-stu-id="1fa34-138">Scroll bar does not appear in task pane</span></span>

<span data-ttu-id="1fa34-139">既定では、Microsoft Edge のスクロール バーはホバーするまで非表示になっています。</span><span class="sxs-lookup"><span data-stu-id="1fa34-139">By default, scroll bars in Microsoft Edge are hidden until hovered over.</span></span> <span data-ttu-id="1fa34-140">スクロールバーが常に表示されるようにするには、作業ウィンドウのページの`<body>`要素に適用される CSS スタイルに [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/-ms-overflow-style) プロパティを含め、`scrollbar`に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1fa34-140">To ensure that the scroll bar is always visible, the CSS styling that applies to the `<body>` element of the pages in the task pane should include the [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/-ms-overflow-style) property and it should be set to `scrollbar`.</span></span> 

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a><span data-ttu-id="1fa34-141">Microsoft Edge DevTools を使用してデバッグすると、アドインがクラッシュまたは再読み込みされる</span><span class="sxs-lookup"><span data-stu-id="1fa34-141">When debugging with the Microsoft Edge DevTools, the add-in crashes or reloads</span></span>

<span data-ttu-id="1fa34-142">[Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) にブレークポイントを設定すると、アドインがハングしていると Office に判断される可能性があります。</span><span class="sxs-lookup"><span data-stu-id="1fa34-142">Setting breakpoints in the [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) can cause Office to think that the add-in is hung.</span></span> <span data-ttu-id="1fa34-143">これが発生すると、アドインが自動的に再読み込みされます。</span><span class="sxs-lookup"><span data-stu-id="1fa34-143">It will automatically reload the add-in when this happens.</span></span> <span data-ttu-id="1fa34-144">これを防ぐには、開発用コンピューターに以下のレジストリ キーと値を追加します: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`。</span><span class="sxs-lookup"><span data-stu-id="1fa34-144">To prevent this, add the following Registry key and value to the development computer: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`.</span></span>

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a><span data-ttu-id="1fa34-145">アドインを開こうとすると、"アドイン エラー localhost からこのアドインを開けません" というエラーが表示される</span><span class="sxs-lookup"><span data-stu-id="1fa34-145">When the add-in tries to open, get "ADD-IN ERROR We can't open this add-in from the localhost" error</span></span>

<span data-ttu-id="1fa34-146">既知の原因の1つとして、Microsoft Edge では開発用コンピューター上では localhost にループバックの除外を与える必要があることが挙げられます。</span><span class="sxs-lookup"><span data-stu-id="1fa34-146">One known cause is that Microsoft Edge requires that localhost be given a loopback exemption on the development computer.</span></span> <span data-ttu-id="1fa34-147">[Cannot open add-in from localhost (localhostからアドインを開くことができません)](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost)の指示に従ってください。</span><span class="sxs-lookup"><span data-stu-id="1fa34-147">Follow the instructions at [Cannot open add-in from localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).</span></span>


## <a name="see-also"></a><span data-ttu-id="1fa34-148">関連項目</span><span class="sxs-lookup"><span data-stu-id="1fa34-148">See also</span></span>

- [<span data-ttu-id="1fa34-149">Officeアドインを実行するための要件</span><span class="sxs-lookup"><span data-stu-id="1fa34-149">Requirements for Running Office Add-ins</span></span>](requirements-for-running-office-add-ins.md)

---
title: Office アドインによって使用されるブラウザー
description: Office アドインによって使用されるブラウザーをオペレーティング システムおよび Office バージョンが決定する方法を指定します。
ms.date: 05/28/2019
localization_priority: Priority
ms.openlocfilehash: 92218bb012ae9031ebfc429606885a0ec0ea85b3
ms.sourcegitcommit: b299b8a5dfffb6102cb14b431bdde4861abfb47f
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/30/2019
ms.locfileid: "34592130"
---
# <a name="browsers-used-by-office-add-ins"></a><span data-ttu-id="11ab9-103">Office アドインによって使用されるブラウザー</span><span class="sxs-lookup"><span data-stu-id="11ab9-103">Web viewers used by Office Add-ins</span></span>

<span data-ttu-id="11ab9-104">Office アドインは、Office Online で実行しているときに iFrame を使用して表示され、デスクトップおよびモバイル クライアント用に Office に埋め込まれたブラウザー コントロールを使用して表示される Web アプリケーションです。</span><span class="sxs-lookup"><span data-stu-id="11ab9-104">Office add-ins are web applications that are displayed using iFrames when running in Office Online and using embedded browser controls in Office for desktop and mobile clients.</span></span> <span data-ttu-id="11ab9-105">アドインには JavaScript を実行するための JavaScript エンジンも必要です。</span><span class="sxs-lookup"><span data-stu-id="11ab9-105">Add-ins also need a JavaScript engine to run the JavaScript.</span></span> <span data-ttu-id="11ab9-106">埋め込まれたブラウザーおよびエンジン、どちらもユーザーのコンピュータにインストールされているブラウザによって提供されます。</span><span class="sxs-lookup"><span data-stu-id="11ab9-106">Both the embedded browser and the engine are supplied by a browser installed on the user’s computer.</span></span>

<span data-ttu-id="11ab9-107">どのブラウザが使用されているかは、以下によります。</span><span class="sxs-lookup"><span data-stu-id="11ab9-107">Which browser is used depends on:</span></span>

- <span data-ttu-id="11ab9-108">コンピュータのオペレーティングシステム。</span><span class="sxs-lookup"><span data-stu-id="11ab9-108">The computer’s operating system.</span></span>
- <span data-ttu-id="11ab9-109">アドインがOffice Online、Office 365、または登録のないOffice 2013以降で実行されているかどうか。</span><span class="sxs-lookup"><span data-stu-id="11ab9-109">Whether the add-in is running in Office Online, Office 365, or non-subscription Office 2013 or later.</span></span>

<span data-ttu-id="11ab9-110">次の表は、さまざまなプラットフォームとオペレーティングシステムに使用されているブラウザを示しています。</span><span class="sxs-lookup"><span data-stu-id="11ab9-110">The following table shows which browser is used for the various platforms and operating systems.</span></span>

|<span data-ttu-id="11ab9-111">**OS / Platform**</span><span class="sxs-lookup"><span data-stu-id="11ab9-111">**OS / Platform**</span></span>|<span data-ttu-id="11ab9-112">**Browser**</span><span class="sxs-lookup"><span data-stu-id="11ab9-112">**Browser**</span></span>|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|<span data-ttu-id="11ab9-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="11ab9-113">Office Online</span></span>|<span data-ttu-id="11ab9-114">Office Onlineが開かれているブラウザ。</span><span class="sxs-lookup"><span data-stu-id="11ab9-114">The browser in which Office Online is opened.</span></span>|
|<span data-ttu-id="11ab9-115">Mac</span><span class="sxs-lookup"><span data-stu-id="11ab9-115">Mac</span></span>|<span data-ttu-id="11ab9-116">Safari</span><span class="sxs-lookup"><span data-stu-id="11ab9-116">Safari</span></span>|
|<span data-ttu-id="11ab9-117">iOS</span><span class="sxs-lookup"><span data-stu-id="11ab9-117">iOS</span></span>|<span data-ttu-id="11ab9-118">Safari</span><span class="sxs-lookup"><span data-stu-id="11ab9-118">Safari</span></span>|
|<span data-ttu-id="11ab9-119">Android</span><span class="sxs-lookup"><span data-stu-id="11ab9-119">Android</span></span>|<span data-ttu-id="11ab9-120">Chrome</span><span class="sxs-lookup"><span data-stu-id="11ab9-120">Chrome</span></span>|
|<span data-ttu-id="11ab9-121">Windows / 非登録 Office 2013以降</span><span class="sxs-lookup"><span data-stu-id="11ab9-121">Windows / non-subscription Office 2013 or later</span></span>|<span data-ttu-id="11ab9-122">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="11ab9-122">Internet Explorer 11</span></span>|
|<span data-ttu-id="11ab9-123">Windows 10 バージョン</span><span class="sxs-lookup"><span data-stu-id="11ab9-123">Windows 10 ver.</span></span> <span data-ttu-id="11ab9-124">< 1903 / Office 365</span><span class="sxs-lookup"><span data-stu-id="11ab9-124">< 1903 / Office 365</span></span>|<span data-ttu-id="11ab9-125">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="11ab9-125">Internet Explorer 11</span></span>|
|<span data-ttu-id="11ab9-126">Windows 10 バージョン</span><span class="sxs-lookup"><span data-stu-id="11ab9-126">Windows 10 ver.</span></span> <span data-ttu-id="11ab9-127">>= 1903 / Office 365 ver < 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="11ab9-127">>= 1903 / Office 365 ver < 16.0.11629</span></span>|<span data-ttu-id="11ab9-128">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="11ab9-128">Internet Explorer 11</span></span>|
|<span data-ttu-id="11ab9-129">Windows 10 バージョン</span><span class="sxs-lookup"><span data-stu-id="11ab9-129">Windows 10 ver.</span></span> <span data-ttu-id="11ab9-130">>= 1903 / Office 365 ver >= 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="11ab9-130">>= 1903 / Office 365 ver >= 16.0.11629</span></span>|<span data-ttu-id="11ab9-131">Microsoft Edge\*</span><span class="sxs-lookup"><span data-stu-id="11ab9-131">Microsoft Edge\*</span></span>|

<span data-ttu-id="11ab9-132">\*Microsoft Edge が使用されている場合、Windows 10 ナレーター (「スクリーン リーダー」と呼ばれることもあります) は、作業ウィンドウで開いているページの `<title>` タグを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="11ab9-132">\* When Edge is being used, the Windows 10 Narrator (sometimes called a "screen reader") reads the `<title>` tag in the page that opens in the task pane.</span></span> <span data-ttu-id="11ab9-133">Internet Explorer 11 が使用されている場合、ナレーターはアドイン マニフェストの `<DisplayName>` の値から提供される作業ウィンドウのタイトル バーを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="11ab9-133">When Internet Explorer 11 is being used, the Narrator reads the title bar of the task pane, which comes from the `<DisplayName>` value in the add-in's manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="11ab9-134">Internet Explorer 11はES5以降のJavaScriptバージョンをサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="11ab9-134">Internet Explorer 11 does not support JavaScript versions later than ES5.</span></span> <span data-ttu-id="11ab9-135">アドインのユーザーが Internet Explorer 11 を使用するプラットフォームを使用している場合、ECMAScript 2015 以降の構文と機能を使用するには、JavaScript を ES 5 にトランスパイルするか、ポリフィルを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="11ab9-135">If any of your add-in's users have platforms that use Internet Explorer 11, then to use the syntax and features of ECMAScript 2015 or later, you will need to either transpile your JavaScript to ES5 or use a polyfill.</span></span> <span data-ttu-id="11ab9-136">また、Internet Explorer 11 は、メディア、録音、および位置情報などの HTML 5 機能の一部をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="11ab9-136">Also, Internet Explorer 11 does not support some HTML 5 features such as media, recording, and location.</span></span>

> [!NOTE]
> <span data-ttu-id="11ab9-137">これらが一般に利用可能になるまで、Windows バージョン 1903 以降を入手するには Windows Insider である必要があり、また、Office バージョン 16.0.11629 以降を入手するには Office Insider である必要があります。</span><span class="sxs-lookup"><span data-stu-id="11ab9-137">Until they are generally available, you need to be a Windows Insider to get a Windows version 1903 or greater, and you need to be an Office Insider to get Office version 16.0.11629 or greater.</span></span>
>
> <span data-ttu-id="11ab9-138">Windows インサイダーに参加するには</span><span class="sxs-lookup"><span data-stu-id="11ab9-138">To join Windows Insiders:</span></span>
> 
> 1. <span data-ttu-id="11ab9-139">[Windows インサイダー](https://insider.windows.com)に移動し、リンクをクリックしてWindows インサイダーに参加してください。</span><span class="sxs-lookup"><span data-stu-id="11ab9-139">Go to [Windows Insider](https://insider.windows.com) and click the link to join Windows Insiders.</span></span>
> 2. <span data-ttu-id="11ab9-140">Windowsのプレビュービルドを有効にするためのWindowsの設定の使用方法についての説明が記載されたページに移動します。</span><span class="sxs-lookup"><span data-stu-id="11ab9-140">You will be taken to a page with instructions about how to use Windows Settings to enable preview builds of Windows.</span></span> <span data-ttu-id="11ab9-141">指示に従います。</span><span class="sxs-lookup"><span data-stu-id="11ab9-141">Follow the instructions.</span></span> <span data-ttu-id="11ab9-142">更新頻度を選択する際は、一番速いオプションを選択してください。</span><span class="sxs-lookup"><span data-stu-id="11ab9-142">When you select the pace of updates, choose the fastest option.</span></span>
>
> <span data-ttu-id="11ab9-143">Office インサイダーに参加するには</span><span class="sxs-lookup"><span data-stu-id="11ab9-143">To join Office Insiders:</span></span>
> 
> 1. <span data-ttu-id="11ab9-144">[Office Insiderになりましょう](https://insider.office.com/join)に移動してください。</span><span class="sxs-lookup"><span data-stu-id="11ab9-144">Go to [Get started as an Office Insider](https://insider.office.com/join).</span></span>
> 2. <span data-ttu-id="11ab9-145">そのページの指示に従って参加してください。</span><span class="sxs-lookup"><span data-stu-id="11ab9-145">Follow the instruction on that page to join.</span></span> <span data-ttu-id="11ab9-146">チャンネルを指定するように求められたら、[インサイダー]を選択します。</span><span class="sxs-lookup"><span data-stu-id="11ab9-146">When asked to specify a channel, select Insider.</span></span>

## <a name="see-also"></a><span data-ttu-id="11ab9-147">関連項目</span><span class="sxs-lookup"><span data-stu-id="11ab9-147">See also</span></span>

- [<span data-ttu-id="11ab9-148">Officeアドインを実行するための要件</span><span class="sxs-lookup"><span data-stu-id="11ab9-148">Requirements for Running Office Add-ins</span></span>](requirements-for-running-office-add-ins.md)

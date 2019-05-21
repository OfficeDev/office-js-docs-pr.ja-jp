---
title: Officeアドインで使用されるWebビューア
description: ''
ms.date: 05/03/2019
localization_priority: Priority
ms.openlocfilehash: 6cb0d6e97dd559727b6a1e140d8417e1146e479a
ms.sourcegitcommit: 944cbb5c6ce055f6db1833182b24d490d1dce01d
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/14/2019
ms.locfileid: "33992127"
---
# <a name="web-viewers-used-by-office-add-ins"></a><span data-ttu-id="8d7ce-102">Officeアドインで使用されるWebビューア</span><span class="sxs-lookup"><span data-stu-id="8d7ce-102">Web viewers used by Office Add-ins</span></span>

<span data-ttu-id="8d7ce-103">OfficeアドインはWebアプリケーションなので、WebアプリケーションのHTMLページを表示するためのWebページビューアと、JavaScriptを実行するためのJavaScriptエンジンが必要です。</span><span class="sxs-lookup"><span data-stu-id="8d7ce-103">Since Office Add-ins are web applications, they need a web page viewer to display the HTML pages of the web application and a JavaScript engine to run the JavaScript.</span></span> <span data-ttu-id="8d7ce-104">どちらもユーザーのコンピュータにインストールされているブラウザによって提供されます。</span><span class="sxs-lookup"><span data-stu-id="8d7ce-104">Both are supplied by a browser installed on the user’s computer.</span></span>

<span data-ttu-id="8d7ce-105">どのブラウザが使用されているかは、以下によります。</span><span class="sxs-lookup"><span data-stu-id="8d7ce-105">Which browser is used depends on:</span></span>

- <span data-ttu-id="8d7ce-106">コンピュータのオペレーティングシステム。</span><span class="sxs-lookup"><span data-stu-id="8d7ce-106">The computer’s operating system.</span></span>
- <span data-ttu-id="8d7ce-107">アドインがOffice Online、Office 365、または登録のないOffice 2013以降で実行されているかどうか。</span><span class="sxs-lookup"><span data-stu-id="8d7ce-107">Whether the add-in is running in Office Online, Office 365, or non-subscription Office 2013 or later.</span></span>

<span data-ttu-id="8d7ce-108">次の表は、さまざまなプラットフォームとオペレーティングシステムに使用されているブラウザを示しています。</span><span class="sxs-lookup"><span data-stu-id="8d7ce-108">The following table shows which browser is used for the various platforms and operating systems.</span></span>

|<span data-ttu-id="8d7ce-109">**OS / Platform**</span><span class="sxs-lookup"><span data-stu-id="8d7ce-109">**OS / Platform**</span></span>|<span data-ttu-id="8d7ce-110">**Browser**</span><span class="sxs-lookup"><span data-stu-id="8d7ce-110">**Browser**</span></span>|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|<span data-ttu-id="8d7ce-111">Office Online</span><span class="sxs-lookup"><span data-stu-id="8d7ce-111">Office Online</span></span>|<span data-ttu-id="8d7ce-112">Office Onlineが開かれているブラウザ。</span><span class="sxs-lookup"><span data-stu-id="8d7ce-112">The browser in which Office Online is opened.</span></span>|
|<span data-ttu-id="8d7ce-113">Mac</span><span class="sxs-lookup"><span data-stu-id="8d7ce-113">Mac</span></span>|<span data-ttu-id="8d7ce-114">Safari</span><span class="sxs-lookup"><span data-stu-id="8d7ce-114">Safari</span></span>|
|<span data-ttu-id="8d7ce-115">iOS</span><span class="sxs-lookup"><span data-stu-id="8d7ce-115">iOS</span></span>|<span data-ttu-id="8d7ce-116">Safari</span><span class="sxs-lookup"><span data-stu-id="8d7ce-116">Safari</span></span>|
|<span data-ttu-id="8d7ce-117">Android</span><span class="sxs-lookup"><span data-stu-id="8d7ce-117">Android</span></span>|<span data-ttu-id="8d7ce-118">Chrome</span><span class="sxs-lookup"><span data-stu-id="8d7ce-118">Chrome</span></span>|
|<span data-ttu-id="8d7ce-119">Windows / 非登録 Office 2013以降</span><span class="sxs-lookup"><span data-stu-id="8d7ce-119">Windows / non-subscription Office 2013 or later</span></span>|<span data-ttu-id="8d7ce-120">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="8d7ce-120">Internet Explorer 11</span></span>|
|<span data-ttu-id="8d7ce-121">Windows 10 バージョン</span><span class="sxs-lookup"><span data-stu-id="8d7ce-121">Windows 10 ver.</span></span> <span data-ttu-id="8d7ce-122">< 1903 / Office 365</span><span class="sxs-lookup"><span data-stu-id="8d7ce-122">< 1903 / Office 365</span></span>|<span data-ttu-id="8d7ce-123">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="8d7ce-123">Internet Explorer 11</span></span>|
|<span data-ttu-id="8d7ce-124">Windows 10 バージョン</span><span class="sxs-lookup"><span data-stu-id="8d7ce-124">Windows 10 ver.</span></span> <span data-ttu-id="8d7ce-125">>= 1903 / Office 365 ver < 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="8d7ce-125">>= 1903 / Office 365 ver < 16.0.11629</span></span>|<span data-ttu-id="8d7ce-126">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="8d7ce-126">Internet Explorer 11</span></span>|
|<span data-ttu-id="8d7ce-127">Windows 10 バージョン</span><span class="sxs-lookup"><span data-stu-id="8d7ce-127">Windows 10 ver.</span></span> <span data-ttu-id="8d7ce-128">>= 1903 / Office 365 ver >= 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="8d7ce-128">>= 1903 / Office 365 ver >= 16.0.11629</span></span>|<span data-ttu-id="8d7ce-129">Microsoft Edge\*</span><span class="sxs-lookup"><span data-stu-id="8d7ce-129">Microsoft Edge\*</span></span>|

<span data-ttu-id="8d7ce-130">\*Microsoft Edge が使用されている場合、Windows 10 ナレーター (「スクリーン リーダー」と呼ばれることもあります) は、作業ウィンドウで開いているページの `<title>` タグを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="8d7ce-130">\* When Microsoft Edge is being used, the Windows 10 Narrator (sometimes called a "screen reader") reads the `<title>` tag in the page that opens in the task pane.</span></span> <span data-ttu-id="8d7ce-131">Internet Explorer 11 が使用されている場合、ナレーターはアドイン マニフェストの `<DisplayName>` の値から提供される作業ウィンドウのタイトル バーを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="8d7ce-131">When Internet Explorer 11 is being used, the Narrator reads the title bar of the task pane, which comes from the `<DisplayName>` value in the add-in's manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8d7ce-132">Internet Explorer 11はES5以降のJavaScriptバージョンをサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="8d7ce-132">Internet Explorer 11 does not support JavaScript versions later than ES5.</span></span> <span data-ttu-id="8d7ce-133">アドインのユーザーが Internet Explorer 11 を使用するプラットフォームを使用している場合、ECMAScript 2015 以降の構文と機能を使用するには、JavaScript を ES 5 にトランスパイルするか、ポリフィルを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8d7ce-133">If any of your add-in's users have platforms that use Internet Explorer 11, then to use the syntax and features of ECMAScript 2015 or later, you will need to either transpile your JavaScript to ES5 or use a polyfill.</span></span> <span data-ttu-id="8d7ce-134">また、Internet Explorer 11 は、メディア、録音、および位置情報などの HTML 5 機能の一部をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="8d7ce-134">Also, Internet Explorer 11 does not support some HTML5 features such as media, recording, and location.</span></span>

> [!NOTE]
> <span data-ttu-id="8d7ce-135">これらが一般に利用可能になるまで、Windows バージョン 1903 以降を入手するには Windows Insider である必要があり、また、Office バージョン 16.0.11629 以降を入手するには Office Insider である必要があります。</span><span class="sxs-lookup"><span data-stu-id="8d7ce-135">Until they are generally available, you need to be a Windows Insider to get a Windows version 1903 or greater, and you need to be an Office Insider to get Office version 16.0.11629 or greater.</span></span>
>
> <span data-ttu-id="8d7ce-136">Windows インサイダーに参加するには</span><span class="sxs-lookup"><span data-stu-id="8d7ce-136">To join Windows Insiders:</span></span>
> 
> 1. <span data-ttu-id="8d7ce-137">[Windows インサイダー](https://insider.windows.com)に移動し、リンクをクリックしてWindows インサイダーに参加してください。</span><span class="sxs-lookup"><span data-stu-id="8d7ce-137">Go to [Windows Insider](https://insider.windows.com) and click the link to join Windows Insiders.</span></span>
> 2. <span data-ttu-id="8d7ce-138">Windowsのプレビュービルドを有効にするためのWindowsの設定の使用方法についての説明が記載されたページに移動します。</span><span class="sxs-lookup"><span data-stu-id="8d7ce-138">You will be taken to a page with instructions about how to use Windows Settings to enable preview builds of Windows.</span></span> <span data-ttu-id="8d7ce-139">指示に従います。</span><span class="sxs-lookup"><span data-stu-id="8d7ce-139">Follow the instructions on the page.</span></span> <span data-ttu-id="8d7ce-140">更新頻度を選択する際は、一番速いオプションを選択してください。</span><span class="sxs-lookup"><span data-stu-id="8d7ce-140">When you select the pace of updates, choose the fastest option.</span></span>
>
> <span data-ttu-id="8d7ce-141">Office インサイダーに参加するには</span><span class="sxs-lookup"><span data-stu-id="8d7ce-141">To join Office Insiders:</span></span>
> 
> 1. <span data-ttu-id="8d7ce-142">[Office Insiderになりましょう](https://insider.office.com/join)に移動してください。</span><span class="sxs-lookup"><span data-stu-id="8d7ce-142">Go to [Get started as an Office Insider](https://insider.office.com/join).</span></span>
> 2. <span data-ttu-id="8d7ce-143">そのページの指示に従って参加してください。</span><span class="sxs-lookup"><span data-stu-id="8d7ce-143">Follow the instruction on that page to join.</span></span> <span data-ttu-id="8d7ce-144">チャンネルを指定するように求められたら、[インサイダー]を選択します。</span><span class="sxs-lookup"><span data-stu-id="8d7ce-144">When asked to specify a channel, select Insider.</span></span>

## <a name="see-also"></a><span data-ttu-id="8d7ce-145">関連項目</span><span class="sxs-lookup"><span data-stu-id="8d7ce-145">See also</span></span>

- [<span data-ttu-id="8d7ce-146">Officeアドインを実行するための要件</span><span class="sxs-lookup"><span data-stu-id="8d7ce-146">Requirements for running Office Add-ins</span></span>](requirements-for-running-office-add-ins.md)

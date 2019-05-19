---
title: Officeアドインで使用されるWebビューア
description: ''
ms.date: 05/03/2019
localization_priority: Priority
ms.openlocfilehash: 632f62cbc02917d9e28ab260f3710498156194db
ms.sourcegitcommit: 47b792755e655043d3db2f1fdb9a1eeb7453c636
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2019
ms.locfileid: "33630406"
---
# <a name="web-viewers-used-by-office-add-ins"></a><span data-ttu-id="2c075-102">Officeアドインで使用されるWebビューア</span><span class="sxs-lookup"><span data-stu-id="2c075-102">Web viewers used by Office Add-ins</span></span>

<span data-ttu-id="2c075-103">OfficeアドインはWebアプリケーションなので、WebアプリケーションのHTMLページを表示するためのWebページビューアと、JavaScriptを実行するためのJavaScriptエンジンが必要です。</span><span class="sxs-lookup"><span data-stu-id="2c075-103">Since Office Add-ins are web applications, they need a web page viewer to display the HTML pages of the web application and a JavaScript engine to run the JavaScript.</span></span> <span data-ttu-id="2c075-104">どちらもユーザーのコンピュータにインストールされているブラウザによって提供されます。</span><span class="sxs-lookup"><span data-stu-id="2c075-104">Both are supplied by a browser installed on the user’s computer.</span></span>

<span data-ttu-id="2c075-105">どのブラウザが使用されているかは、以下によります。</span><span class="sxs-lookup"><span data-stu-id="2c075-105">Which browser is used depends on:</span></span>

- <span data-ttu-id="2c075-106">コンピュータのオペレーティングシステム。</span><span class="sxs-lookup"><span data-stu-id="2c075-106">The computer’s operating system.</span></span>
- <span data-ttu-id="2c075-107">アドインがOffice Online、Office 365、または登録のないOffice 2013以降で実行されているかどうか。</span><span class="sxs-lookup"><span data-stu-id="2c075-107">Whether the add-in is running in Office Online, Office 365, or non-subscription Office 2013 or later.</span></span>

<span data-ttu-id="2c075-108">次の表は、さまざまなプラットフォームとオペレーティングシステムに使用されているブラウザを示しています。</span><span class="sxs-lookup"><span data-stu-id="2c075-108">The following table shows which browser is used for the various platforms and operating systems.</span></span>

|<span data-ttu-id="2c075-109">**OS / Platform**</span><span class="sxs-lookup"><span data-stu-id="2c075-109">**OS / Platform**</span></span>|<span data-ttu-id="2c075-110">**Browser**</span><span class="sxs-lookup"><span data-stu-id="2c075-110">**Browser**</span></span>|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|<span data-ttu-id="2c075-111">Office Online</span><span class="sxs-lookup"><span data-stu-id="2c075-111">Office Online</span></span>|<span data-ttu-id="2c075-112">Office Onlineが開かれているブラウザ。</span><span class="sxs-lookup"><span data-stu-id="2c075-112">The browser in which Office Online is opened.</span></span>|
|<span data-ttu-id="2c075-113">Mac</span><span class="sxs-lookup"><span data-stu-id="2c075-113">Mac</span></span>|<span data-ttu-id="2c075-114">Safari</span><span class="sxs-lookup"><span data-stu-id="2c075-114">Safari</span></span>|
|<span data-ttu-id="2c075-115">iOS</span><span class="sxs-lookup"><span data-stu-id="2c075-115">iOS</span></span>|<span data-ttu-id="2c075-116">Safari</span><span class="sxs-lookup"><span data-stu-id="2c075-116">Safari</span></span>|
|<span data-ttu-id="2c075-117">Android</span><span class="sxs-lookup"><span data-stu-id="2c075-117">Android</span></span>|<span data-ttu-id="2c075-118">Chrome</span><span class="sxs-lookup"><span data-stu-id="2c075-118">Chrome</span></span>|
|<span data-ttu-id="2c075-119">Windows / 非登録 Office 2013以降</span><span class="sxs-lookup"><span data-stu-id="2c075-119">Windows / non-subscription Office 2013 or later</span></span>|<span data-ttu-id="2c075-120">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="2c075-120">Internet Explorer 11</span></span>|
|<span data-ttu-id="2c075-121">Windows 10 バージョン</span><span class="sxs-lookup"><span data-stu-id="2c075-121">Windows 10 ver.</span></span> <span data-ttu-id="2c075-122">< 1903 / Office 365</span><span class="sxs-lookup"><span data-stu-id="2c075-122">< 1903 / Office 365</span></span>|<span data-ttu-id="2c075-123">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="2c075-123">Internet Explorer 11</span></span>|
|<span data-ttu-id="2c075-124">Windows 10 バージョン</span><span class="sxs-lookup"><span data-stu-id="2c075-124">Windows 10 ver.</span></span> <span data-ttu-id="2c075-125">>= 1903 / Office 365 ver < 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="2c075-125">>= 1903 / Office 365 ver < 16.0.11629</span></span>|<span data-ttu-id="2c075-126">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="2c075-126">Internet Explorer 11</span></span>|
|<span data-ttu-id="2c075-127">Windows 10 バージョン</span><span class="sxs-lookup"><span data-stu-id="2c075-127">Windows 10 ver.</span></span> <span data-ttu-id="2c075-128">>= 1903 / Office 365 ver >= 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="2c075-128">>= 1903 / Office 365 ver >= 16.0.11629</span></span>|<span data-ttu-id="2c075-129">Edge\*</span><span class="sxs-lookup"><span data-stu-id="2c075-129">Edge\*</span></span>|

<span data-ttu-id="2c075-130">\*Edgeが使用されている場合、Windows 10ナレータ（ "スクリーンリーダー"と呼ばれることもあります）は、作業ペインに表示されるページの`<title>`タグを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="2c075-130">\* When Edge is being used, the Windows 10 Narrator (sometimes called a "screen reader") reads the `<title>` tag in the page that opens in the task pane.</span></span> <span data-ttu-id="2c075-131">Internet Explorer 11が使用されている場合、ナレータは作業ペインのタイトルバーを読み取ります。これはアドインのマニフェストの`<DisplayName>`の値から取得されます。</span><span class="sxs-lookup"><span data-stu-id="2c075-131">When Internet Explorer 11 is being used, the Narrator reads the title bar of the task pane, which comes from the `<DisplayName>` value in the add-in's manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2c075-132">Internet Explorer 11はES5以降のJavaScriptバージョンをサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="2c075-132">Internet Explorer 11 does not support JavaScript versions later than ES5.</span></span> <span data-ttu-id="2c075-133">アドインのユーザーがInternet Explorer 11を使用するプラットフォームを使用している場合に、ECMAScript 2015以降の構文と機能を使用するには、JavaScriptをES 5に変換するか、ポリフィルを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2c075-133">If any of your add-in's users have platforms that use Internet Explorer 11, then to use the syntax and features of ECMAScript 2015 or later, you will need to either transpile your JavaScript to ES5 or use a polyfill.</span></span> <span data-ttu-id="2c075-134">また、Internet Explorer 11は、メディア、録音、および位置情報などのHTML 5機能の一部をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="2c075-134">Also, Internet Explorer 11 does not support some HTML 5 features such as media, recording, and location.</span></span>

> [!NOTE]
> <span data-ttu-id="2c075-135">これらが一般に利用可能になるまで、Windowsバージョン1903以降を入手するためWindowsインサイダーである必要があり、また、Officeバージョン16.0.11629以降を入手するためOfficeインサイダーである必要があります。</span><span class="sxs-lookup"><span data-stu-id="2c075-135">Until they are generally available, you need to be a Windows Insider to get a Windows version 1903 or greater, and you need to be an Office Insider to get Office version 16.0.11629 or greater.</span></span>
>
> <span data-ttu-id="2c075-136">Windows インサイダーに参加するには</span><span class="sxs-lookup"><span data-stu-id="2c075-136">To join Windows Insiders:</span></span>
> 
> 1. <span data-ttu-id="2c075-137">[Windows インサイダー](https://insider.windows.com)に移動し、リンクをクリックしてWindows インサイダーに参加してください。</span><span class="sxs-lookup"><span data-stu-id="2c075-137">Go to [Windows Insider](https://insider.windows.com) and click the link to join Windows Insiders.</span></span>
> 2. <span data-ttu-id="2c075-138">Windowsのプレビュービルドを有効にするためのWindowsの設定の使用方法についての説明が記載されたページに移動します。</span><span class="sxs-lookup"><span data-stu-id="2c075-138">You will be taken to a page with instructions about how to use Windows Settings to enable preview builds of Windows.</span></span> <span data-ttu-id="2c075-139">指示に従います。</span><span class="sxs-lookup"><span data-stu-id="2c075-139">Follow the instructions on the page.</span></span> <span data-ttu-id="2c075-140">更新頻度を選択する際は、一番速いオプションを選択してください。</span><span class="sxs-lookup"><span data-stu-id="2c075-140">When you select the pace of updates, choose the fastest option.</span></span>
>
> <span data-ttu-id="2c075-141">Office インサイダーに参加するには</span><span class="sxs-lookup"><span data-stu-id="2c075-141">To join Office Insiders:</span></span>
> 
> 1. <span data-ttu-id="2c075-142">[Office Insiderになりましょう](https://insider.office.com/join)に移動してください。</span><span class="sxs-lookup"><span data-stu-id="2c075-142">Go to [Get started as an Office Insider](https://insider.office.com/join).</span></span>
> 2. <span data-ttu-id="2c075-143">そのページの指示に従って参加してください。</span><span class="sxs-lookup"><span data-stu-id="2c075-143">Follow the instruction on that page to join.</span></span> <span data-ttu-id="2c075-144">チャンネルを指定するように求められたら、[インサイダー]を選択します。</span><span class="sxs-lookup"><span data-stu-id="2c075-144">When asked to specify a channel, select Insider.</span></span>

## <a name="see-also"></a><span data-ttu-id="2c075-145">関連項目</span><span class="sxs-lookup"><span data-stu-id="2c075-145">See also</span></span>

- [<span data-ttu-id="2c075-146">Officeアドインを実行するための要件</span><span class="sxs-lookup"><span data-stu-id="2c075-146">Requirements for running Office Add-ins</span></span>](requirements-for-running-office-add-ins.md)

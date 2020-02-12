---
title: Office アドインを実行するための要件
description: ''
ms.date: 07/01/2019
localization_priority: Normal
ms.openlocfilehash: 61108b713299f91e3f60993d17885daf022254f8
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950370"
---
# <a name="requirements-for-running-office-add-ins"></a><span data-ttu-id="a4c90-102">Office アドインを実行するための要件</span><span class="sxs-lookup"><span data-stu-id="a4c90-102">Requirements for running Office Add-ins</span></span>

<span data-ttu-id="a4c90-103">この記事では、Office アドインを実行するためのソフトウェアとデバイスの要件について説明します。</span><span class="sxs-lookup"><span data-stu-id="a4c90-103">This article describes the software and device requirements for running Office Add-ins.</span></span>

> [!NOTE]
> <span data-ttu-id="a4c90-p101">AppSource にアドインを[公開](../publish/publish.md)し、Office エクスペリエンスで利用できるようにする予定がある場合は、[AppSource の検証ポリシー](/office/dev/store/validation-policies)に準拠していることを確認してください。たとえば、検証に合格するには、定義したメソッドをサポートするすべてのプラットフォームでアドインが動作する必要があります (詳細については、[セクション 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) と [Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)のページを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="a4c90-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

<span data-ttu-id="a4c90-106">現時点での Office アドインのサポート状況について、概要は「[Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a4c90-106">For a high-level view of where Office Add-ins are currently supported, see [Office Add-in host and platform availability](../overview/office-add-in-availability.md).</span></span>

## <a name="server-requirements"></a><span data-ttu-id="a4c90-107">サーバーの要件</span><span class="sxs-lookup"><span data-stu-id="a4c90-107">Server requirements</span></span>

<span data-ttu-id="a4c90-108">Office アドインをインストールおよび実行できるようにするには、まずアドインの UI とコードのマニフェストと Web ページ ファイルを、適切なサーバーの場所に展開する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a4c90-108">To be able to install and run any Office Add-in, you first need to deploy the manifest and webpage files for the UI and code of your add-in to the appropriate server locations.</span></span>

<span data-ttu-id="a4c90-109">すべての種類のアドイン (コンテンツ、Outlook、作業ウィンドウの、アドインとアドイン コマンド) で、アドインの Web ページ ファイルを Web サーバーや [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md) などの Web ホスティング サービスに展開する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a4c90-109">For all types of add-ins (content, Outlook, and task pane add-ins and add-in commands), you need to deploy your add-in's webpage files to a web server, or web hosting service, such as [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md).</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

> [!TIP]
> <span data-ttu-id="a4c90-110">Visual Studio でアドインを開発およびデバッグする際、Visual Studio は IIS Express を使用してアドインの Web ページ ファイルをローカルで展開および実行するので、追加の Web サーバーは必要ありません。</span><span class="sxs-lookup"><span data-stu-id="a4c90-110">When you develop and debug an add-in in Visual Studio, Visual Studio deploys and runs your add-in's webpage files locally with IIS Express, and doesn't require an additional web server.</span></span>

<span data-ttu-id="a4c90-111">サポートされている Office ホスト アプリケーション (Excel、PowerPoint、Project、または Word) のコンテンツ アドインと作業ウィンドウ アドインでは、アドインの XML マニフェスト ファイルをアップロードするために、SharePoint の[アプリ カタログ](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)も必要になります。</span><span class="sxs-lookup"><span data-stu-id="a4c90-111">For content and task pane add-ins, in the supported Office host applications - Excel, PowerPoint, Project, or Word - you also need an [app catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) on SharePoint to upload the add-in's XML manifest file.</span></span>

<span data-ttu-id="a4c90-p102">Outlook アドインをテストおよび実行するには、ユーザーの Outlook 電子メール アカウントが、Office 365、Exchange Online、またはオンプレミスのインストールから使用できる Exchange 2013 以降のバージョン上に存在する必要があります。ユーザーまたは管理者は、サーバー上に Outlook アドインのマニフェスト ファイルをインストールします。</span><span class="sxs-lookup"><span data-stu-id="a4c90-p102">To test and run an Outlook add-in, the user's Outlook email account must reside on Exchange 2013 or later, which is available through Office 365, Exchange Online, or through an on-premises installation. The user or administrator installs manifest files for Outlook add-ins on that server.</span></span>

> [!NOTE]
> <span data-ttu-id="a4c90-114">Outlook の POP および IMAP 電子メール アカウントは、Office アドインをサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="a4c90-114">POP and IMAP email accounts in Outlook don't support Office Add-ins.</span></span>

## <a name="client-requirements-windows-desktop-and-tablet"></a><span data-ttu-id="a4c90-115">クライアントの要件: Windows デスクトップおよびタブレット</span><span class="sxs-lookup"><span data-stu-id="a4c90-115">Client requirements: Windows desktop and tablet</span></span>

<span data-ttu-id="a4c90-116">Windows ベースのデスクトップ、ノート PC、または タブレット デバイス上で実行されるサポート対象の Office デスクトップ クライアントまたは Web クライアント向けの Office アドインを開発するには、以下のソフトウェアが必要です。</span><span class="sxs-lookup"><span data-stu-id="a4c90-116">The following software is required for developing an Office Add-in for the supported Office desktop clients or web clients that run on Windows-based desktop, laptop, or tablet devices:</span></span>


- <span data-ttu-id="a4c90-117">Windows x86 および x64 デスクトップおよび Surface Pro などのタブレット:</span><span class="sxs-lookup"><span data-stu-id="a4c90-117">For Windows x86 and x64 desktops, and tablets such as Surface Pro:</span></span>
    - <span data-ttu-id="a4c90-118">Windows 7 以降のバージョンで実行している Office 2013 以降のバージョンの、32 ビットまたは 64 ビット バージョン。</span><span class="sxs-lookup"><span data-stu-id="a4c90-118">The 32- or 64-bit version of Office 2013 or a later version, running on Windows 7 or a later version.</span></span>
    - <span data-ttu-id="a4c90-p103">Excel 2013、Outlook 2013、PowerPoint 2013、Project Professional 2013、Project 2013 SP1、Word 2013、またはそれ以降の Office クライアントのバージョン (特にこれらの Office デスクトップ クライアントを対象として Office アドインをテストまたは実行する場合)。Office デスクトップ クライアントはオンプレミスでインストールすることも、クイック実行によってクライアント コンピューターにインストールすることもできます。</span><span class="sxs-lookup"><span data-stu-id="a4c90-p103">Excel 2013, Outlook 2013, PowerPoint 2013, Project Professional 2013, Project 2013 SP1, Word 2013, or a later version of the Office client, if you are testing or running an Office Add-in specifically for one of these Office desktop clients. Office desktop clients can be installed on premises or via Click-to-Run on the client computer.</span></span>

  <span data-ttu-id="a4c90-121">有効な Office 365 サブスクリプションがあり、Office クライアント へのアクセス権がない場合は、[最新バージョンの Office をダウンロードしてインストールする](https://support.office.com/article/download-and-install-or-reinstall-office-365-or-office-2019-on-a-pc-or-mac-4414eaaf-0478-48be-9c42-23adc4716658)ことができます。</span><span class="sxs-lookup"><span data-stu-id="a4c90-121">If you have a valid Office 365 subscription and you do not have access to the Office client, you can [download and install the latest version of Office](https://support.office.com/article/download-and-install-or-reinstall-office-365-or-office-2019-on-a-pc-or-mac-4414eaaf-0478-48be-9c42-23adc4716658).</span></span>

- <span data-ttu-id="a4c90-122">Internet Explorer 11 または Microsoft Edge (Windows および Office のバージョンによる) がインストールされている必要がありますが、既定のブラウザーである必要はありません。</span><span class="sxs-lookup"><span data-stu-id="a4c90-122">Internet Explorer 11 or Microsoft Edge (depending on the Windows and Office versions) must be installed but doesn't have to be the default browser.</span></span> <span data-ttu-id="a4c90-123">Office アドインをサポートするために、ホストとして動作する Office のクライアントは、Internet Explorer 11 または Microsoft Edge に組み込まれているブラウザー コンポーネントを使用します。</span><span class="sxs-lookup"><span data-stu-id="a4c90-123">To support Office Add-ins, the Office client that acts as host uses browser components that are part of Internet Explorer 11 or Microsoft Edge.</span></span> <span data-ttu-id="a4c90-124">詳細については、「[Office アドインによって使用されるブラウザー](browsers-used-by-office-web-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a4c90-124">See [Browsers used by Office Add-ins](browsers-used-by-office-web-add-ins.md) for more details.</span></span>

  > [!NOTE]
  > <span data-ttu-id="a4c90-125">Office Web アドインが機能するためには、Internet Explorer のセキュリティ強化の構成 (ESC) がオフになっている必要があります。</span><span class="sxs-lookup"><span data-stu-id="a4c90-125">Internet Explorer's Enhanced Security Configuration (ESC) must be turned off for Office Web Add-ins to work.</span></span> <span data-ttu-id="a4c90-126">アドインを開発する際に Windows Server コンピューターをクライアントとして使用する場合は、Windows Server では既定で ESC がオンになっていることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="a4c90-126">If you are using a Windows Server computer as your client when developing add-ins, note that ESC is turned on by default in Windows Server.</span></span>

- <span data-ttu-id="a4c90-127">既定のブラウザーとして次のいずれか: Internet Explorer 11、または Microsoft Edge、Chrome、Firefox、Safari (Mac OS) の最新バージョンのうちいずれか。</span><span class="sxs-lookup"><span data-stu-id="a4c90-127">One of the following as the default browser: Internet Explorer 11, or the latest version of Microsoft Edge, Chrome, Firefox, or Safari (Mac OS).</span></span>
- <span data-ttu-id="a4c90-128">メモ帳などの HTML および JavaScript エディター、[Visual Studio および Microsoft Developer Tools](https://www.visualstudio.com/features/office-tools-vs)、またはサードパーティの Web 開発ツール。</span><span class="sxs-lookup"><span data-stu-id="a4c90-128">An HTML and JavaScript editor such as Notepad, [Visual Studio and the Microsoft Developer Tools](https://www.visualstudio.com/features/office-tools-vs), or a third-party web development tool.</span></span>

## <a name="client-requirements-os-x-desktop"></a><span data-ttu-id="a4c90-129">クライアントの要件: OS X デスクトップ</span><span class="sxs-lookup"><span data-stu-id="a4c90-129">Client requirements: OS X desktop</span></span>

<span data-ttu-id="a4c90-130">Mac 上の Outlook は、Office 365 に付属していて、Outlook アドインをサポートします。Mac 上の Outlook で Outlook アドインを実行するための要件は、Mac 上の Outlook そのものの要件と同じです。オペレーティング システムは、少なくとも OS X v10.10 "Yosemite" である必要があります。</span><span class="sxs-lookup"><span data-stu-id="a4c90-130">Outlook on Mac, which is distributed as part of Office 365, supports Outlook add-ins. Running Outlook add-ins in Outlook on Mac has the same requirements as Outlook on Mac itself: the operating system must be at least OS X v10.10 "Yosemite".</span></span> <span data-ttu-id="a4c90-131">Mac 上の Outlook はレイアウト エンジンとして WebKit を使用して、アドイン ページを表示するので、追加のブラウザーの依存関係はありません。</span><span class="sxs-lookup"><span data-stu-id="a4c90-131">Because Outlook on Mac uses WebKit as a layout engine to render the add-in pages, there is no additional browser dependency.</span></span>

<span data-ttu-id="a4c90-132">次は、Office アドインをサポートする Mac 上の Office の最小クライアント バージョンです。</span><span class="sxs-lookup"><span data-stu-id="a4c90-132">The following are the minimum client versions of Office on Mac that support Office Add-ins.</span></span>

- <span data-ttu-id="a4c90-133">Word バージョン 15.18 (160109)</span><span class="sxs-lookup"><span data-stu-id="a4c90-133">Word version 15.18 (160109)</span></span>
- <span data-ttu-id="a4c90-134">Excel バージョン 15.19 (160206)</span><span class="sxs-lookup"><span data-stu-id="a4c90-134">Excel version 15.19 (160206)</span></span>
- <span data-ttu-id="a4c90-135">PowerPoint バージョン 15.24 (160614)</span><span class="sxs-lookup"><span data-stu-id="a4c90-135">PowerPoint version 15.24 (160614)</span></span>

## <a name="client-requirements-browser-support-for-office-web-clients-and-sharepoint"></a><span data-ttu-id="a4c90-136">クライアントの要件: Office Web クライアントと SharePoint のブラウザー サポート</span><span class="sxs-lookup"><span data-stu-id="a4c90-136">Client requirements: Browser support for Office web clients and SharePoint</span></span>

<span data-ttu-id="a4c90-137">Internet Explorer 11、または Microsoft Edge、Chrome、Firefox、Safari (Mac OS) の最新バージョンなど、ECMAScript 5.1、HTML5、CSS3 をサポートする任意のブラウザー。</span><span class="sxs-lookup"><span data-stu-id="a4c90-137">Any browser that supports ECMAScript 5.1, HTML5, and CSS3, such as Internet Explorer 11, or the latest version of Microsoft Edge, Chrome, Firefox, or Safari (Mac OS).</span></span>


## <a name="client-requirements-non-windows-smartphone-and-tablet"></a><span data-ttu-id="a4c90-138">クライアントの要件: Windows 以外のスマートフォンおよびタブレット</span><span class="sxs-lookup"><span data-stu-id="a4c90-138">Client requirements: non-Windows smartphone and tablet</span></span>

<span data-ttu-id="a4c90-139">特に、スマートフォンや Windows 以外のタブレット デバイス上のブラウザーで動作する Outlook の場合、Outlook アドインをテストおよび実行するのに以下のソフトウェアが必要です。</span><span class="sxs-lookup"><span data-stu-id="a4c90-139">Specifically for Outlook running in a browser on smartphones and non-Windows tablet devices, the following software is required for testing and running Outlook add-ins.</span></span>


| <span data-ttu-id="a4c90-140">ホスト アプリケーション</span><span class="sxs-lookup"><span data-stu-id="a4c90-140">Host application</span></span> | <span data-ttu-id="a4c90-141">デバイス</span><span class="sxs-lookup"><span data-stu-id="a4c90-141">Device</span></span> | <span data-ttu-id="a4c90-142">オペレーティング システム</span><span class="sxs-lookup"><span data-stu-id="a4c90-142">Operating system</span></span> | <span data-ttu-id="a4c90-143">Exchange アカウント</span><span class="sxs-lookup"><span data-stu-id="a4c90-143">Exchange account</span></span> | <span data-ttu-id="a4c90-144">モバイル ブラウザー</span><span class="sxs-lookup"><span data-stu-id="a4c90-144">Mobile browser</span></span> |
|:-----|:-----|:-----|:-----|:-----|
|<span data-ttu-id="a4c90-145">Android 上の Outlook</span><span class="sxs-lookup"><span data-stu-id="a4c90-145">Outlook on Android</span></span>|<span data-ttu-id="a4c90-146">Android のタブレットとスマートフォン</span><span class="sxs-lookup"><span data-stu-id="a4c90-146">Android tablets and smartphones</span></span>|<span data-ttu-id="a4c90-147">Android 4.4 KitKat 以降</span><span class="sxs-lookup"><span data-stu-id="a4c90-147">Android 4.4 KitKat later</span></span>|<span data-ttu-id="a4c90-148">Office 365 for Business または Exchange Online の最新の更新プログラムが対象</span><span class="sxs-lookup"><span data-stu-id="a4c90-148">On the latest update of Office 365 for business or Exchange Online</span></span>|<span data-ttu-id="a4c90-149">Android 用のネイティブ アプリ、ブラウザーは適用外</span><span class="sxs-lookup"><span data-stu-id="a4c90-149">Native app for Android, browser not applicable</span></span>|
|<span data-ttu-id="a4c90-150">iOS 上の Outlook</span><span class="sxs-lookup"><span data-stu-id="a4c90-150">Outlook on iOS</span></span>|<span data-ttu-id="a4c90-151">iPad のタブレット、iPhone のスマート フォン</span><span class="sxs-lookup"><span data-stu-id="a4c90-151">iPad tablets, iPhone smartphones</span></span>|<span data-ttu-id="a4c90-152">iOS 11 以降</span><span class="sxs-lookup"><span data-stu-id="a4c90-152">iOS 11 or later</span></span>|<span data-ttu-id="a4c90-153">Office 365 for Business または Exchange Online の最新の更新プログラムが対象</span><span class="sxs-lookup"><span data-stu-id="a4c90-153">On the latest update of Office 365 for business or Exchange Online</span></span>|<span data-ttu-id="a4c90-154">iOS 用のネイティブ アプリ、ブラウザーは適用外</span><span class="sxs-lookup"><span data-stu-id="a4c90-154">Native app for iOS, browser not applicable</span></span>|
|<span data-ttu-id="a4c90-155">Outlook on the web</span><span class="sxs-lookup"><span data-stu-id="a4c90-155">Outlook on the web</span></span>|<span data-ttu-id="a4c90-156">iPhone 4 以降、iPad 2 以降、iPod Touch 4 以降</span><span class="sxs-lookup"><span data-stu-id="a4c90-156">iPhone 4 or later, iPad 2 or later, iPod Touch 4 or later</span></span>|<span data-ttu-id="a4c90-157">iOS 5 以降</span><span class="sxs-lookup"><span data-stu-id="a4c90-157">iOS 5 or later</span></span>|<span data-ttu-id="a4c90-158">Office 365、Exchange Online、または Exchange Server 2013 以降のオンプレミスが対象</span><span class="sxs-lookup"><span data-stu-id="a4c90-158">On Office 365, Exchange Online, or on premises on Exchange Server 2013 or later</span></span>|<span data-ttu-id="a4c90-159">Safari</span><span class="sxs-lookup"><span data-stu-id="a4c90-159">Safari</span></span>|

> [!NOTE]
> <span data-ttu-id="a4c90-160">ネイティブ アプリの OWA for Android、OWA for iPad、および OWA for iPhone は[廃止](https://support.office.com/article/Microsoft-OWA-mobile-apps-are-being-retired-076ec122-4576-4900-bc26-937f84d25a4b)され、Outlook アドインのテストには不要になり、利用もできなくなりました。</span><span class="sxs-lookup"><span data-stu-id="a4c90-160">The native apps OWA for Android, OWA for iPad, and OWA for iPhone have been [deprecated](https://support.office.com/article/Microsoft-OWA-mobile-apps-are-being-retired-076ec122-4576-4900-bc26-937f84d25a4b) and are no longer required or available for testing Outlook add-ins.</span></span>


## <a name="see-also"></a><span data-ttu-id="a4c90-161">関連項目</span><span class="sxs-lookup"><span data-stu-id="a4c90-161">See also</span></span>

- [<span data-ttu-id="a4c90-162">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="a4c90-162">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="a4c90-163">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="a4c90-163">Office Add-in host and platform availability</span></span>](../overview/office-add-in-availability.md)
- [<span data-ttu-id="a4c90-164">Office アドインによって使用されるブラウザー</span><span class="sxs-lookup"><span data-stu-id="a4c90-164">Browsers used by Office Add-ins</span></span>](browsers-used-by-office-web-add-ins.md)

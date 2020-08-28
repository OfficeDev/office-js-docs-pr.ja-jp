---
title: Office アドインを実行するための要件
description: エンドユーザーが Office アドインを実行するために必要なクライアントおよびサーバーの要件について説明します。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 3c4ed1116fa83de694f390cfc26fde9a01858460
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293010"
---
# <a name="requirements-for-running-office-add-ins"></a><span data-ttu-id="9c330-103">Office アドインを実行するための要件</span><span class="sxs-lookup"><span data-stu-id="9c330-103">Requirements for running Office Add-ins</span></span>

<span data-ttu-id="9c330-104">この記事では、Office アドインを実行するためのソフトウェアとデバイスの要件について説明します。</span><span class="sxs-lookup"><span data-stu-id="9c330-104">This article describes the software and device requirements for running Office Add-ins.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

<span data-ttu-id="9c330-105">現在、Office アドインが現在サポートされている場所の概要については、「office [クライアントアプリケーションおよび Office アドインのプラットフォームの可用性](../overview/office-add-in-availability.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9c330-105">For a high-level view of where Office Add-ins are currently supported, see [Office client application and platform availability for Office Add-ins](../overview/office-add-in-availability.md).</span></span>

## <a name="server-requirements"></a><span data-ttu-id="9c330-106">サーバーの要件</span><span class="sxs-lookup"><span data-stu-id="9c330-106">Server requirements</span></span>

<span data-ttu-id="9c330-107">Office アドインをインストールおよび実行できるようにするには、まずアドインの UI とコードのマニフェストと Web ページ ファイルを、適切なサーバーの場所に展開する必要があります。</span><span class="sxs-lookup"><span data-stu-id="9c330-107">To be able to install and run any Office Add-in, you first need to deploy the manifest and webpage files for the UI and code of your add-in to the appropriate server locations.</span></span>

<span data-ttu-id="9c330-108">すべての種類のアドイン (コンテンツ、Outlook、作業ウィンドウの、アドインとアドイン コマンド) で、アドインの Web ページ ファイルを Web サーバーや [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md) などの Web ホスティング サービスに展開する必要があります。</span><span class="sxs-lookup"><span data-stu-id="9c330-108">For all types of add-ins (content, Outlook, and task pane add-ins and add-in commands), you need to deploy your add-in's webpage files to a web server, or web hosting service, such as [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md).</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

> [!TIP]
> <span data-ttu-id="9c330-109">Visual Studio でアドインを開発およびデバッグする際、Visual Studio は IIS Express を使用してアドインの Web ページ ファイルをローカルで展開および実行するので、追加の Web サーバーは必要ありません。</span><span class="sxs-lookup"><span data-stu-id="9c330-109">When you develop and debug an add-in in Visual Studio, Visual Studio deploys and runs your add-in's webpage files locally with IIS Express, and doesn't require an additional web server.</span></span>

<span data-ttu-id="9c330-110">コンテンツアドインと作業ウィンドウアドインについては、サポートされている Office クライアントアプリケーション (Excel、PowerPoint、Project、または Word) で、SharePoint の [アプリカタログ](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) を使用してアドインの XML マニフェストファイルをアップロードするか、 [一元展開](../publish/centralized-deployment.md)を使用してアドインを展開する必要があります。</span><span class="sxs-lookup"><span data-stu-id="9c330-110">For content and task pane add-ins, in the supported Office client applications - Excel, PowerPoint, Project, or Word - you also need either an [app catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) on SharePoint to upload the add-in's XML manifest file, or you need to deploy the add-in using [Centralized Deployment](../publish/centralized-deployment.md).</span></span>

<span data-ttu-id="9c330-111">Outlook アドインをテストして実行するには、ユーザーの Outlook メールアカウントが、Microsoft 365、Exchange Online、またはオンプレミスのインストールによって利用可能な Exchange 2013 以降に存在する必要があります。</span><span class="sxs-lookup"><span data-stu-id="9c330-111">To test and run an Outlook add-in, the user's Outlook email account must reside on Exchange 2013 or later, which is available through Microsoft 365, Exchange Online, or through an on-premises installation.</span></span> <span data-ttu-id="9c330-112">ユーザーまたは管理者は、サーバー上に Outlook アドインのマニフェスト ファイルをインストールします。</span><span class="sxs-lookup"><span data-stu-id="9c330-112">The user or administrator installs manifest files for Outlook add-ins on that server.</span></span>

> [!NOTE]
> <span data-ttu-id="9c330-113">Outlook の POP および IMAP 電子メール アカウントは、Office アドインをサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="9c330-113">POP and IMAP email accounts in Outlook don't support Office Add-ins.</span></span>

## <a name="client-requirements-windows-desktop-and-tablet"></a><span data-ttu-id="9c330-114">クライアントの要件: Windows デスクトップおよびタブレット</span><span class="sxs-lookup"><span data-stu-id="9c330-114">Client requirements: Windows desktop and tablet</span></span>

<span data-ttu-id="9c330-115">Windows ベースのデスクトップ、ノート PC、または タブレット デバイス上で実行されるサポート対象の Office デスクトップ クライアントまたは Web クライアント向けの Office アドインを開発するには、以下のソフトウェアが必要です。</span><span class="sxs-lookup"><span data-stu-id="9c330-115">The following software is required for developing an Office Add-in for the supported Office desktop clients or web clients that run on Windows-based desktop, laptop, or tablet devices:</span></span>


- <span data-ttu-id="9c330-116">Windows x86 および x64 デスクトップおよび Surface Pro などのタブレット:</span><span class="sxs-lookup"><span data-stu-id="9c330-116">For Windows x86 and x64 desktops, and tablets such as Surface Pro:</span></span>
    - <span data-ttu-id="9c330-117">Windows 7 以降のバージョンで実行している Office 2013 以降のバージョンの、32 ビットまたは 64 ビット バージョン。</span><span class="sxs-lookup"><span data-stu-id="9c330-117">The 32- or 64-bit version of Office 2013 or a later version, running on Windows 7 or a later version.</span></span>
    - <span data-ttu-id="9c330-p102">Excel 2013、Outlook 2013、PowerPoint 2013、Project Professional 2013、Project 2013 SP1、Word 2013、またはそれ以降の Office クライアントのバージョン (特にこれらの Office デスクトップ クライアントを対象として Office アドインをテストまたは実行する場合)。Office デスクトップ クライアントはオンプレミスでインストールすることも、クイック実行によってクライアント コンピューターにインストールすることもできます。</span><span class="sxs-lookup"><span data-stu-id="9c330-p102">Excel 2013, Outlook 2013, PowerPoint 2013, Project Professional 2013, Project 2013 SP1, Word 2013, or a later version of the Office client, if you are testing or running an Office Add-in specifically for one of these Office desktop clients. Office desktop clients can be installed on premises or via Click-to-Run on the client computer.</span></span>

  <span data-ttu-id="9c330-120">有効な Microsoft 365 サブスクリプションがあり、Office クライアントにアクセスできない場合は、 [最新バージョンの office をダウンロードしてインストール](https://support.office.com/article/download-and-install-or-reinstall-office-365-or-office-2019-on-a-pc-or-mac-4414eaaf-0478-48be-9c42-23adc4716658)することができます。</span><span class="sxs-lookup"><span data-stu-id="9c330-120">If you have a valid Microsoft 365 subscription and you do not have access to the Office client, you can [download and install the latest version of Office](https://support.office.com/article/download-and-install-or-reinstall-office-365-or-office-2019-on-a-pc-or-mac-4414eaaf-0478-48be-9c42-23adc4716658).</span></span>

- <span data-ttu-id="9c330-121">Internet Explorer 11 または Microsoft Edge (Windows および Office のバージョンによる) がインストールされている必要がありますが、既定のブラウザーである必要はありません。</span><span class="sxs-lookup"><span data-stu-id="9c330-121">Internet Explorer 11 or Microsoft Edge (depending on the Windows and Office versions) must be installed but doesn't have to be the default browser.</span></span> <span data-ttu-id="9c330-122">Office アドインをサポートするために、ホストとして動作する Office のクライアントは、Internet Explorer 11 または Microsoft Edge に組み込まれているブラウザー コンポーネントを使用します。</span><span class="sxs-lookup"><span data-stu-id="9c330-122">To support Office Add-ins, the Office client that acts as host uses browser components that are part of Internet Explorer 11 or Microsoft Edge.</span></span> <span data-ttu-id="9c330-123">詳細については、「[Office アドインによって使用されるブラウザー](browsers-used-by-office-web-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9c330-123">See [Browsers used by Office Add-ins](browsers-used-by-office-web-add-ins.md) for more details.</span></span>

  > [!NOTE]
  > <span data-ttu-id="9c330-124">Office Web アドインが機能するためには、Internet Explorer のセキュリティ強化の構成 (ESC) がオフになっている必要があります。</span><span class="sxs-lookup"><span data-stu-id="9c330-124">Internet Explorer's Enhanced Security Configuration (ESC) must be turned off for Office Web Add-ins to work.</span></span> <span data-ttu-id="9c330-125">アドインを開発する際に Windows Server コンピューターをクライアントとして使用する場合は、Windows Server では既定で ESC がオンになっていることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="9c330-125">If you are using a Windows Server computer as your client when developing add-ins, note that ESC is turned on by default in Windows Server.</span></span>

- <span data-ttu-id="9c330-126">既定のブラウザーとして次のいずれか: Internet Explorer 11、または Microsoft Edge、Chrome、Firefox、Safari (Mac OS) の最新バージョンのうちいずれか。</span><span class="sxs-lookup"><span data-stu-id="9c330-126">One of the following as the default browser: Internet Explorer 11, or the latest version of Microsoft Edge, Chrome, Firefox, or Safari (Mac OS).</span></span>
- <span data-ttu-id="9c330-127">メモ帳などの HTML および JavaScript エディター、[Visual Studio および Microsoft Developer Tools](https://www.visualstudio.com/features/office-tools-vs)、またはサードパーティの Web 開発ツール。</span><span class="sxs-lookup"><span data-stu-id="9c330-127">An HTML and JavaScript editor such as Notepad, [Visual Studio and the Microsoft Developer Tools](https://www.visualstudio.com/features/office-tools-vs), or a third-party web development tool.</span></span>

## <a name="client-requirements-os-x-desktop"></a><span data-ttu-id="9c330-128">クライアントの要件: OS X デスクトップ</span><span class="sxs-lookup"><span data-stu-id="9c330-128">Client requirements: OS X desktop</span></span>

<span data-ttu-id="9c330-129">Microsoft 365 の一部として配布される outlook on Mac は、Outlook アドインをサポートします。 Mac 上の outlook アドインを実行すると、outlook For mac 自体と同じ要件があります。オペレーティングシステムは、少なくとも OS X v 10.10 "ヨーク Semite" である必要があります。</span><span class="sxs-lookup"><span data-stu-id="9c330-129">Outlook on Mac, which is distributed as part of Microsoft 365, supports Outlook add-ins. Running Outlook add-ins in Outlook on Mac has the same requirements as Outlook on Mac itself: the operating system must be at least OS X v10.10 "Yosemite".</span></span> <span data-ttu-id="9c330-130">Mac 上の Outlook はレイアウト エンジンとして WebKit を使用して、アドイン ページを表示するので、追加のブラウザーの依存関係はありません。</span><span class="sxs-lookup"><span data-stu-id="9c330-130">Because Outlook on Mac uses WebKit as a layout engine to render the add-in pages, there is no additional browser dependency.</span></span>

<span data-ttu-id="9c330-131">次は、Office アドインをサポートする Mac 上の Office の最小クライアント バージョンです。</span><span class="sxs-lookup"><span data-stu-id="9c330-131">The following are the minimum client versions of Office on Mac that support Office Add-ins.</span></span>

- <span data-ttu-id="9c330-132">Word バージョン 15.18 (160109)</span><span class="sxs-lookup"><span data-stu-id="9c330-132">Word version 15.18 (160109)</span></span>
- <span data-ttu-id="9c330-133">Excel バージョン 15.19 (160206)</span><span class="sxs-lookup"><span data-stu-id="9c330-133">Excel version 15.19 (160206)</span></span>
- <span data-ttu-id="9c330-134">PowerPoint バージョン 15.24 (160614)</span><span class="sxs-lookup"><span data-stu-id="9c330-134">PowerPoint version 15.24 (160614)</span></span>

## <a name="client-requirements-browser-support-for-office-web-clients-and-sharepoint"></a><span data-ttu-id="9c330-135">クライアントの要件: Office Web クライアントと SharePoint のブラウザー サポート</span><span class="sxs-lookup"><span data-stu-id="9c330-135">Client requirements: Browser support for Office web clients and SharePoint</span></span>

<span data-ttu-id="9c330-136">Internet Explorer 11、または Microsoft Edge、Chrome、Firefox、Safari (Mac OS) の最新バージョンなど、ECMAScript 5.1、HTML5、CSS3 をサポートする任意のブラウザー。</span><span class="sxs-lookup"><span data-stu-id="9c330-136">Any browser that supports ECMAScript 5.1, HTML5, and CSS3, such as Internet Explorer 11, or the latest version of Microsoft Edge, Chrome, Firefox, or Safari (Mac OS).</span></span>


## <a name="client-requirements-non-windows-smartphone-and-tablet"></a><span data-ttu-id="9c330-137">クライアントの要件: Windows 以外のスマートフォンおよびタブレット</span><span class="sxs-lookup"><span data-stu-id="9c330-137">Client requirements: non-Windows smartphone and tablet</span></span>

<span data-ttu-id="9c330-138">特に、スマートフォンや Windows 以外のタブレット デバイス上のブラウザーで動作する Outlook の場合、Outlook アドインをテストおよび実行するのに以下のソフトウェアが必要です。</span><span class="sxs-lookup"><span data-stu-id="9c330-138">Specifically for Outlook running in a browser on smartphones and non-Windows tablet devices, the following software is required for testing and running Outlook add-ins.</span></span>


| <span data-ttu-id="9c330-139">Office アプリケーション</span><span class="sxs-lookup"><span data-stu-id="9c330-139">Office application</span></span> | <span data-ttu-id="9c330-140">デバイス</span><span class="sxs-lookup"><span data-stu-id="9c330-140">Device</span></span> | <span data-ttu-id="9c330-141">オペレーティング システム</span><span class="sxs-lookup"><span data-stu-id="9c330-141">Operating system</span></span> | <span data-ttu-id="9c330-142">Exchange アカウント</span><span class="sxs-lookup"><span data-stu-id="9c330-142">Exchange account</span></span> | <span data-ttu-id="9c330-143">モバイル ブラウザー</span><span class="sxs-lookup"><span data-stu-id="9c330-143">Mobile browser</span></span> |
|:-----|:-----|:-----|:-----|:-----|
|<span data-ttu-id="9c330-144">Android 上の Outlook</span><span class="sxs-lookup"><span data-stu-id="9c330-144">Outlook on Android</span></span>|<span data-ttu-id="9c330-145">Android のタブレットとスマートフォン</span><span class="sxs-lookup"><span data-stu-id="9c330-145">Android tablets and smartphones</span></span>|<span data-ttu-id="9c330-146">Android 4.4 KitKat 以降</span><span class="sxs-lookup"><span data-stu-id="9c330-146">Android 4.4 KitKat later</span></span>|<span data-ttu-id="9c330-147">Microsoft 365 Apps for business または Exchange Online の最新の更新プログラム</span><span class="sxs-lookup"><span data-stu-id="9c330-147">On the latest update of Microsoft 365 Apps for business or Exchange Online</span></span>|<span data-ttu-id="9c330-148">Android 用のネイティブ アプリ、ブラウザーは適用外</span><span class="sxs-lookup"><span data-stu-id="9c330-148">Native app for Android, browser not applicable</span></span>|
|<span data-ttu-id="9c330-149">iOS 上の Outlook</span><span class="sxs-lookup"><span data-stu-id="9c330-149">Outlook on iOS</span></span>|<span data-ttu-id="9c330-150">iPad のタブレット、iPhone のスマート フォン</span><span class="sxs-lookup"><span data-stu-id="9c330-150">iPad tablets, iPhone smartphones</span></span>|<span data-ttu-id="9c330-151">iOS 11 以降</span><span class="sxs-lookup"><span data-stu-id="9c330-151">iOS 11 or later</span></span>|<span data-ttu-id="9c330-152">Microsoft 365 Apps for business または Exchange Online の最新の更新プログラム</span><span class="sxs-lookup"><span data-stu-id="9c330-152">On the latest update of Microsoft 365 Apps for business or Exchange Online</span></span>|<span data-ttu-id="9c330-153">iOS 用のネイティブ アプリ、ブラウザーは適用外</span><span class="sxs-lookup"><span data-stu-id="9c330-153">Native app for iOS, browser not applicable</span></span>|
|<span data-ttu-id="9c330-154">Outlook on the web</span><span class="sxs-lookup"><span data-stu-id="9c330-154">Outlook on the web</span></span>|<span data-ttu-id="9c330-155">iPhone 4 以降、iPad 2 以降、iPod Touch 4 以降</span><span class="sxs-lookup"><span data-stu-id="9c330-155">iPhone 4 or later, iPad 2 or later, iPod Touch 4 or later</span></span>|<span data-ttu-id="9c330-156">iOS 5 以降</span><span class="sxs-lookup"><span data-stu-id="9c330-156">iOS 5 or later</span></span>|<span data-ttu-id="9c330-157">Microsoft 365、Exchange Online、または Exchange Server 2013 以降のオンプレミス</span><span class="sxs-lookup"><span data-stu-id="9c330-157">On Microsoft 365, Exchange Online, or on premises on Exchange Server 2013 or later</span></span>|<span data-ttu-id="9c330-158">Safari</span><span class="sxs-lookup"><span data-stu-id="9c330-158">Safari</span></span>|

> [!NOTE]
> <span data-ttu-id="9c330-159">ネイティブ アプリの OWA for Android、OWA for iPad、および OWA for iPhone は[廃止](https://support.office.com/article/Microsoft-OWA-mobile-apps-are-being-retired-076ec122-4576-4900-bc26-937f84d25a4b)され、Outlook アドインのテストには不要になり、利用もできなくなりました。</span><span class="sxs-lookup"><span data-stu-id="9c330-159">The native apps OWA for Android, OWA for iPad, and OWA for iPhone have been [deprecated](https://support.office.com/article/Microsoft-OWA-mobile-apps-are-being-retired-076ec122-4576-4900-bc26-937f84d25a4b) and are no longer required or available for testing Outlook add-ins.</span></span>


## <a name="see-also"></a><span data-ttu-id="9c330-160">関連項目</span><span class="sxs-lookup"><span data-stu-id="9c330-160">See also</span></span>

- [<span data-ttu-id="9c330-161">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="9c330-161">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="9c330-162">Office アドインの office クライアントアプリケーションとプラットフォームの可用性</span><span class="sxs-lookup"><span data-stu-id="9c330-162">Office client application and platform availability for Office Add-ins</span></span>](../overview/office-add-in-availability.md)
- [<span data-ttu-id="9c330-163">Office アドインによって使用されるブラウザー</span><span class="sxs-lookup"><span data-stu-id="9c330-163">Browsers used by Office Add-ins</span></span>](browsers-used-by-office-web-add-ins.md)

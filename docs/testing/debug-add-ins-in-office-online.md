---
title: Office on the web でアドインをデバッグする
description: Office on the web を使用してアドインをテストおよびデバッグする方法。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: f7ef3fa3d6389629e28b428b9bdbe3b128896b1f
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094492"
---
# <a name="debug-add-ins-in-office-on-the-web"></a><span data-ttu-id="60856-103">Office on the web でアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="60856-103">Debug add-ins in Office on the web</span></span>

<span data-ttu-id="60856-104">Windows、Office 2013、または Office 2016 デスクトップ クライアントを実行していないコンピューター (たとえば、Mac で開発を行っている場合) でアドインの作成とデバッグを行えます。この記事では、Office Online を使用してアドインのテストとデバッグを行う方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="60856-104">You can build and debug add-ins on a computer that isn't running Windows or the Office desktop client&mdash;for example, if you're developing on a Mac.</span></span> <span data-ttu-id="60856-105">この記事では、Office on the web を使用してアドインをテストおよびデバッグする方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="60856-105">This article describes how to use Office on the web to test and debug your add-ins.</span></span> 

## <a name="prerequisites"></a><span data-ttu-id="60856-106">前提条件</span><span class="sxs-lookup"><span data-stu-id="60856-106">Prerequisites</span></span>

<span data-ttu-id="60856-107">開始するには</span><span class="sxs-lookup"><span data-stu-id="60856-107">To get started:</span></span>

- <span data-ttu-id="60856-108">Microsoft 365 開発者アカウントを持っていない場合、または SharePoint サイトにアクセスできない場合は、Microsoft 開発者アカウントを取得します。</span><span class="sxs-lookup"><span data-stu-id="60856-108">Get a Microsoft 365 developer account if you don't already have one or have access to a SharePoint site.</span></span>

  > [!NOTE]
  > <span data-ttu-id="60856-p102">90日更新プログラムの Microsoft 365 開発者向けサブスクリプションを無料で入手するには、 [microsoft 365 developer プログラム](https://developer.microsoft.com/office/dev-program)にご参加ください。Microsoft 365 開発者プログラムに参加し、サブスクリプションを構成する方法の詳細な手順については、 [microsoft 365 開発者向けプログラムのドキュメント](/office/developer-program/office-365-developer-program)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="60856-p102">To get a free, 90-day renewable Microsoft 365 developer subscription, join our [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program). See the [Microsoft 365 developer program documentation](/office/developer-program/office-365-developer-program) for step-by-step instructions about how to join the Microsoft 365 developer program and configure your subscription.</span></span>

- <span data-ttu-id="60856-p103">SharePoint Online でアプリカタログを設定します。アプリカタログは、Office アドインのドキュメントライブラリをホストする SharePoint Online の専用サイトコレクションです。独自の SharePoint サイトがある場合は、アプリカタログドキュメントライブラリをセットアップすることができます。詳細については、「 [SharePoint のアプリカタログに作業ウィンドウアドインとコンテンツアドインを発行する](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="60856-p103">Set up an app catalog on SharePoint Online. An app catalog is a dedicated site collection in SharePoint Online that hosts document libraries for Office Add-ins. If you have your own SharePoint site, you can set up an app catalog document library. For more information, see [Publish task pane and content add-ins to an app catalog on SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span></span>


## <a name="debug-your-add-in-from-excel-or-word-on-the-web"></a><span data-ttu-id="60856-114">Excel または Word on the web からアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="60856-114">Debug your add-in from Excel or Word on the web</span></span>

<span data-ttu-id="60856-115">Word on the web を使用してアドインをデバッグするには: </span><span class="sxs-lookup"><span data-stu-id="60856-115">To debug your add-in by using Office on the web:</span></span>

1. <span data-ttu-id="60856-116">SSL をサポートするサーバーにアドインを展開します。</span><span class="sxs-lookup"><span data-stu-id="60856-116">Deploy your add-in to a server that supports SSL.</span></span>

    > [!NOTE]
    > <span data-ttu-id="60856-117">[Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)を使用して、アドインを作成し、ホストすることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="60856-117">We recommend that you use the [Yeoman generator](https://github.com/OfficeDev/generator-office) to create and host your add-in.</span></span>

2. <span data-ttu-id="60856-p104">[アドイン マニフェスト ファイル](../develop/add-in-manifests.md)で、相対 URI ではなく絶対 URI を含めるように **SourceLocation** 要素の値を更新します。たとえば次のようにします。</span><span class="sxs-lookup"><span data-stu-id="60856-p104">In your [add-in manifest file](../develop/add-in-manifests.md), update the **SourceLocation** element value to include an absolute, rather than a relative, URI. For example:</span></span>

    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```

3. <span data-ttu-id="60856-120">SharePoint のアプリ カタログにある Office アドイン ライブラリにマニフェストをアップロードします。</span><span class="sxs-lookup"><span data-stu-id="60856-120">Upload the manifest to the Office Add-ins library in the app catalog on SharePoint.</span></span>

4. <span data-ttu-id="60856-121">Microsoft 365 のアプリ起動ツールから web 上で Excel または Word を起動し、新しいドキュメントを開きます。</span><span class="sxs-lookup"><span data-stu-id="60856-121">Launch Excel or Word on the web from the app launcher in Microsoft 365, and open a new document.</span></span>

5. <span data-ttu-id="60856-122">[挿入] タブで、**[個人用アドイン]** または **[Office アドイン]** をクリックし、アプリにアドインを挿入してテストします。</span><span class="sxs-lookup"><span data-stu-id="60856-122">On the Insert tab, choose **My Add-ins** or **Office Add-ins** to insert your add-in and test it in the app.</span></span>

6. <span data-ttu-id="60856-123">お気に入りのブラウザーのツール デバッガーを使用してアドインをデバッグします。</span><span class="sxs-lookup"><span data-stu-id="60856-123">Use your favorite browser tool debugger to debug your add-in.</span></span>

## <a name="potential-issues"></a><span data-ttu-id="60856-124">潜在的な問題</span><span class="sxs-lookup"><span data-stu-id="60856-124">Potential issues</span></span>

<span data-ttu-id="60856-125">以下は、デバッグ時に発生する可能性がある問題です。</span><span class="sxs-lookup"><span data-stu-id="60856-125">The following are some issues that you might encounter as you debug:</span></span>

- <span data-ttu-id="60856-126">表示される JavaScript エラーのいくつかは Office on the web に起因している可能性があります。</span><span class="sxs-lookup"><span data-stu-id="60856-126">Some JavaScript errors that you see might originate from Office on the web.</span></span>

- <span data-ttu-id="60856-127">ブラウザーに無効な証明書エラーが表示されることがありますが、このエラーはバイパスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="60856-127">The browser might show an invalid certificate error that you will need to bypass.</span></span> <span data-ttu-id="60856-128">これを行うプロセスは、ブラウザおよびこの変更を定期的に行うさまざまなブラウザの UI によって異なります。</span><span class="sxs-lookup"><span data-stu-id="60856-128">The process for doing this varies with the browser and the various browsers' UIs for doing this change periodically.</span></span> <span data-ttu-id="60856-129">詳細については、ブラウザーのヘルプを検索するか、オンラインで検索してください。</span><span class="sxs-lookup"><span data-stu-id="60856-129">You should search the browser's help or search online for instructions.</span></span> <span data-ttu-id="60856-130">(たとえば、「Microsoft Edge の無効な証明書警告」を検索します。) ほとんどのブラウザーには、警告ページにリンクがあり、このリンクをクリックするとアドイン ページにアクセスされます。</span><span class="sxs-lookup"><span data-stu-id="60856-130">(For example, search for "Microsoft Edge invalid certificate warning".) Most browsers will have a link on the warning page that enables you to click through to the add-in page.</span></span> <span data-ttu-id="60856-131">たとえば、Microsoft Edge には「Web ページへ移動 (推奨しません)」というリンクがあります。</span><span class="sxs-lookup"><span data-stu-id="60856-131">For example, Microsoft Edge has a link "Go on to the webpage (Not recommended)".</span></span> <span data-ttu-id="60856-132">ただし、通常はアドインが再び読み込まれるたびに、このリンクを経由する必要があります。</span><span class="sxs-lookup"><span data-stu-id="60856-132">But you will usually have to go through this link every time the add-in reloads.</span></span> <span data-ttu-id="60856-133">継続的なバイパスについては、お勧めのヘルプを参照してください。</span><span class="sxs-lookup"><span data-stu-id="60856-133">For a longer lasting bypass, see the help as suggested.</span></span>

- <span data-ttu-id="60856-134">コードにブレークポイントを設定すると、保存できないというエラーが Office on the web からスローされることがあります。</span><span class="sxs-lookup"><span data-stu-id="60856-134">If you set breakpoints in your code, Office on the web might throw an error indicating that it is unable to save.</span></span>

## <a name="see-also"></a><span data-ttu-id="60856-135">関連項目</span><span class="sxs-lookup"><span data-stu-id="60856-135">See also</span></span>

- [<span data-ttu-id="60856-136">Office アドイン開発のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="60856-136">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="60856-137">AppSource の検証ポリシー</span><span class="sxs-lookup"><span data-stu-id="60856-137">AppSource validation policies</span></span>](/legal/marketplace/certification-policies)  
- [<span data-ttu-id="60856-138">効率的な AppSource アプリおよびアドインを作成する</span><span class="sxs-lookup"><span data-stu-id="60856-138">Create effective AppSource apps and add-ins</span></span>](/office/dev/store/create-effective-office-store-listings)  
- [<span data-ttu-id="60856-139">Office アドインでのユーザー エラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="60856-139">Troubleshoot user errors with Office Add-ins</span></span>](testing-and-troubleshooting.md)

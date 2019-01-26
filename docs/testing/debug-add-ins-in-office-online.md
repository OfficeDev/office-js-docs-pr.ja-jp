---
title: Office Online でアドインをデバッグする
description: Office Online を使用してアドインのテストとデバッグを行う方法
ms.date: 03/14/2018
localization_priority: Priority
ms.openlocfilehash: 6252a713444f7ec8bf955c3283a650f72cbcbed1
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29386555"
---
# <a name="debug-add-ins-in-office-online"></a><span data-ttu-id="b4cf9-103">Office Online でアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="b4cf9-103">Debug add-ins in Office Online</span></span>


<span data-ttu-id="b4cf9-104">Windows、Office 2013、または Office 2016 デスクトップ クライアントを実行していないコンピューター (たとえば、Mac で開発を行っている場合) でアドインの作成とデバッグを行えます。この記事では、Office Online を使用してアドインのテストとデバッグを行う方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="b4cf9-104">You can build and debug add-ins on a computer that isn't running Windows or the Office desktop client&mdash;for example, if you're developing on a Mac.</span></span> <span data-ttu-id="b4cf9-105">この記事では、Office Online を使用してアドインのテストとデバッグを行う方法を説明します。</span><span class="sxs-lookup"><span data-stu-id="b4cf9-105">This article describes how to use Office Online to test and debug your add-ins.</span></span> 

## <a name="prerequisites"></a><span data-ttu-id="b4cf9-106">前提条件</span><span class="sxs-lookup"><span data-stu-id="b4cf9-106">Prerequisites</span></span>

<span data-ttu-id="b4cf9-107">開始するには</span><span class="sxs-lookup"><span data-stu-id="b4cf9-107">To get started:</span></span>

- <span data-ttu-id="b4cf9-108">Office 365 の開発者アカウントをまだお持ちでない場合はこれを取得します。または SharePoint サイトにアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="b4cf9-108">Get an Office 365 developer account if you don't already have one or have access to a SharePoint site.</span></span>
    
  > [!NOTE]
  > <span data-ttu-id="b4cf9-109">無料の Office 365 開発者サブスクリプションにサインアップするには、[Office 365 開発者プログラム](https://developer.microsoft.com/office/dev-program)にご参加ください。</span><span class="sxs-lookup"><span data-stu-id="b4cf9-109">To sign up for a free Office 365 developer subscription, join our [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span> <span data-ttu-id="b4cf9-110">Office 365 開発者プログラムに参加し、サブスクリプションにサインアップして構成する方法についての詳しい手順については、[Office 365 開発者プログラムのドキュメント](https://docs.microsoft.com/office/developer-program/office-365-developer-program)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b4cf9-110">See the [Office 365 Developer Program documentation](https://docs.microsoft.com/office/developer-program/office-365-developer-program) for step-by-step instructions about how to join the Office 365 Developer Program and sign up and configure your subscription.</span></span>
     
- <span data-ttu-id="b4cf9-p103">Office 365 (SharePoint Online) 上でアドイン カタログをセットアップするアドイン カタログとは、Office アドイン用のドキュメント ライブラリをホストする SharePoint Online の専用サイト コレクションです。独自の SharePoint サイトを所有している場合は、アドイン カタログのドキュメント ライブラリをセットアップすることができます。詳細については、「[作業ウィンドウ アドインとコンテンツ アドインを SharePoint カタログに発行する](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="b4cf9-p103">Set up an add-in catalog on Office 365 (SharePoint Online). An add-in catalog is a dedicated site collection in SharePoint Online that hosts document libraries for Office Add-ins. If you have your own SharePoint site, you can set up an add-in catalog document library. For more information, see [Publish task pane and content add-ins to an add-in catalog on SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span></span>
    

## <a name="debug-your-add-in-from-excel-online-or-word-online"></a><span data-ttu-id="b4cf9-114">Excel Online または Word Online からアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="b4cf9-114">Debug your add-in from Excel Online or Word Online</span></span>

<span data-ttu-id="b4cf9-115">Office Online を使用してアドインをデバッグするには、</span><span class="sxs-lookup"><span data-stu-id="b4cf9-115">To debug your add-in by using Office Online:</span></span>

1. <span data-ttu-id="b4cf9-116">SSL をサポートするサーバーにアドインを展開します。</span><span class="sxs-lookup"><span data-stu-id="b4cf9-116">Deploy your add-in to a server that supports SSL.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="b4cf9-117">[Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)を使用して、アドインを作成し、ホストすることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="b4cf9-117">We recommend that you use the [Yeoman generator](https://github.com/OfficeDev/generator-office) to create and host your add-in.</span></span>
     
2. <span data-ttu-id="b4cf9-p104">[アドイン マニフェスト ファイル](../develop/add-in-manifests.md)で、相対 URI ではなく絶対 URI を含めるように **SourceLocation** 要素の値を更新します。たとえば次のようにします。</span><span class="sxs-lookup"><span data-stu-id="b4cf9-p104">In your [add-in manifest file](../develop/add-in-manifests.md), update the **SourceLocation** element value to include an absolute, rather than a relative, URI. For example:</span></span>
      
    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```
    
3. <span data-ttu-id="b4cf9-120">SharePoint のアドイン カタログにある Office アドイン ライブラリにマニフェストをアップロードします。</span><span class="sxs-lookup"><span data-stu-id="b4cf9-120">Upload the manifest to the Office Add-ins library in the add-in catalog on SharePoint.</span></span>
    
4. <span data-ttu-id="b4cf9-121">Office 365 のアプリ起動ツールから Excel Online または Word Online を起動し、新しいドキュメントを開きます。</span><span class="sxs-lookup"><span data-stu-id="b4cf9-121">Launch Excel Online or Word Online from the app launcher in Office 365, and open a new document.</span></span>
    
5. <span data-ttu-id="b4cf9-122">[挿入] タブで、 **[個人用アドイン]** または **[Office アドイン]** をクリックし、アプリにアドインを挿入してテストします。</span><span class="sxs-lookup"><span data-stu-id="b4cf9-122">On the Insert tab, choose  **My Add-ins** or **Office Add-ins** to insert your add-in and test it in the app.</span></span>
    
6. <span data-ttu-id="b4cf9-123">お気に入りのブラウザーのツール デバッガーを使用してアドインをデバッグします。</span><span class="sxs-lookup"><span data-stu-id="b4cf9-123">Use your favorite browser tool debugger to debug your add-in.</span></span>

## <a name="potential-issues"></a><span data-ttu-id="b4cf9-124">潜在的な問題</span><span class="sxs-lookup"><span data-stu-id="b4cf9-124">Potential issues</span></span>    

<span data-ttu-id="b4cf9-125">以下は、デバッグ時に発生する可能性がある問題です。</span><span class="sxs-lookup"><span data-stu-id="b4cf9-125">The following are some issues that you might encounter as you debug:</span></span>
    
- <span data-ttu-id="b4cf9-126">表示される JavaScript エラーのいくつかは Office Online に起因している可能性があります。</span><span class="sxs-lookup"><span data-stu-id="b4cf9-126">Some JavaScript errors that you see might originate from Office Online.</span></span>
      
- <span data-ttu-id="b4cf9-127">ブラウザーが、バイパスが必要になる、無効な証明書エラーを表示することがあります。</span><span class="sxs-lookup"><span data-stu-id="b4cf9-127">The browser might show an invalid certificate error that you will need to bypass.</span></span>
      
- <span data-ttu-id="b4cf9-128">コードにブレークポイントを設定する場合、Office Online から、保存できないというエラーがスローされることがあります。</span><span class="sxs-lookup"><span data-stu-id="b4cf9-128">If you set breakpoints in your code, Office Online might throw an error indicating that it is unable to save.</span></span>

## <a name="see-also"></a><span data-ttu-id="b4cf9-129">関連項目</span><span class="sxs-lookup"><span data-stu-id="b4cf9-129">See also</span></span>

- [<span data-ttu-id="b4cf9-130">Office アドイン開発のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="b4cf9-130">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="b4cf9-131">AppSource の検証ポリシー</span><span class="sxs-lookup"><span data-stu-id="b4cf9-131">AppSource validation policies</span></span>](https://docs.microsoft.com/office/dev/store/validation-policies)  
- [<span data-ttu-id="b4cf9-132">効率的な AppSource アプリおよびアドインを作成する</span><span class="sxs-lookup"><span data-stu-id="b4cf9-132">Create effective AppSource apps and add-ins</span></span>](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings)  
- [<span data-ttu-id="b4cf9-133">Office アドインでのユーザー エラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="b4cf9-133">Troubleshoot user errors with Office Add-ins</span></span>](testing-and-troubleshooting.md)
    

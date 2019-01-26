---
title: Microsoft Azure で Office アドインをホストする | Microsoft Docs
description: アドイン Web アプリを Azure に展開して、Office クライアント アプリケーションでテストのためにアドインをサイドロードする方法について説明します。
ms.date: 01/25/2018
localization_priority: Priority
ms.openlocfilehash: ce1cea8078c1842f4ce8cc57b8702c30393d8be8
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29386835"
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a><span data-ttu-id="c9e8d-103">Microsoft Azure で Office アドインをホストする</span><span class="sxs-lookup"><span data-stu-id="c9e8d-103">Host an Office Add-in on Microsoft Azure</span></span>

<span data-ttu-id="c9e8d-p101">最も簡単な Office アドインは、XML マニフェスト ファイルと HTML ページで成り立っています。XML マニフェスト ファイルには、アドインの特性 (アドインの名前や実行可能な Office クライアント アプリケーションの種類、アドインの HTML ページの URL など) を記述します。HTML ページには、ユーザーが Office クライアント アプリケーションにアドインをインストールして実行したときに操作する、Web アプリが含まれています。Office アドインの Web アプリは、Azure を含む、あらゆる Web ホスティング プラットフォームでホストできます。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-p101">The simplest Office Add-in is made up of an XML manifest file and an HTML page. The XML manifest file describes the add-in's characteristics, such as its name, what Office client applications it can run in, and the URL for the add-in's HTML page. The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application. You can host the web app of an Office Add-in on any web hosting platform, including Azure.</span></span>

<span data-ttu-id="c9e8d-108">この記事では、アドイン Web アプリを Azure に展開して、Office クライアント アプリケーションでテストのために[アドインをサイドロード](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-108">This article describes how to deploy an add-in web app to Azure and [sideload the add-in](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) for testing in an Office client application.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="c9e8d-109">前提条件</span><span class="sxs-lookup"><span data-stu-id="c9e8d-109">Prerequisites</span></span> 

1. <span data-ttu-id="c9e8d-110">[Visual Studio 2017](https://www.visualstudio.com/downloads) をインストールします。このとき、**Azure 開発**ワークロードを含めるようにします。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-110">Install [Visual Studio 2017](https://www.visualstudio.com/downloads) and choose to include the **Azure development** workload.</span></span>

    > [!NOTE]
    > <span data-ttu-id="c9e8d-111">既に Visual Studio 2017 がインストールされている場合は、[Visual Studio インストーラー](https://docs.microsoft.com/visualstudio/install/modify-visual-studio)を使用して、**Azure 開発**ワークロードがインストールされていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-111">If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](https://docs.microsoft.com/visualstudio/install/modify-visual-studio) to ensure that the **Azure development** workload is installed.</span></span> 

2. <span data-ttu-id="c9e8d-112">Office をインストールする。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-112">Install Office.</span></span> 
    
    > [!NOTE]
    > <span data-ttu-id="c9e8d-113">まだ Office を所持していない場合は、[1 か月間無料試用版の登録](https://products.office.com/en-US/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735)が可能です。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-113">If you don't already have Office, you can [register for a free 1-month trial](https://products.office.com/en-US/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).</span></span>

3.  <span data-ttu-id="c9e8d-114">Azure サブスクリプションを取得します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-114">Obtain an Azure subscription.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="c9e8d-115">まだ Azure サブスクリプションを所持していない場合、このサブスクリプションは [Visual Studio サブスクリプションの一部として取得](https://azure.microsoft.com/ja-JP/pricing/member-offers/visual-studio-subscriptions/)できます。また、[無料試用版の登録](https://azure.microsoft.com/pricing/free-trial)も可能です。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-115">If don't already have an Azure subscription, you can [get one as part of your MSDN subscription](https://azure.microsoft.com/ja-JP/pricing/member-offers/visual-studio-subscriptions/) or [register for a free trial](https://azure.microsoft.com/pricing/free-trial).</span></span> 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a><span data-ttu-id="c9e8d-116">手順 1: アドインの XML マニフェスト ファイルをホストするための共有フォルダーを作成する</span><span class="sxs-lookup"><span data-stu-id="c9e8d-116">Step 1: Create a shared folder to host your add-in XML manifest file</span></span>

1. <span data-ttu-id="c9e8d-117">開発用のコンピューターでエクスプローラーを開きます。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-117">Open File Explorer on your development computer.</span></span>
    
2. <span data-ttu-id="c9e8d-118">C:\ ドライブを右クリックして、**[新規]** > **[フォルダー]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-118">Right-click the C:\ drive and then choose **New** > **Folder**.</span></span>
    
3. <span data-ttu-id="c9e8d-119">新規フォルダーに「AddinManifests」という名前を付けます。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-119">Name the new folder AddinManifests.</span></span>
    
4. <span data-ttu-id="c9e8d-120">[AddinManifests] フォルダーを右クリックして、**[共有相手]** > **[特定の人]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-120">Right-click the AddinManifests folder and then choose **Share with** > **Specific people**.</span></span>
    
5. <span data-ttu-id="c9e8d-121">**[ファイル共有]** で、ドロップダウンの矢印をクリックして、**[すべてのユーザー]** > **[追加]** > **[共有]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-121">In **File Sharing**, choose the drop-down arrow and then choose **Everyone** > **Add** > **Share**.</span></span>
    
> [!NOTE]
> <span data-ttu-id="c9e8d-p102">このチュートリアルでは、信頼できるカタログとしてローカルのファイル共有を使用します。アドインの XML マニフェスト ファイルは、この場所に保存することになります。現実のシナリオでは、[SharePoint カタログに XML マニフェスト ファイルを展開](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)するか、[AppSource にアドインを発行](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)することもできます。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-p102">In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file. In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store).</span></span>

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a><span data-ttu-id="c9e8d-124">手順 2: 信頼できるアドイン カタログにファイル共有を追加する</span><span class="sxs-lookup"><span data-stu-id="c9e8d-124">Step 2: Add the file share to the Trusted Add-ins catalog</span></span>

1.  <span data-ttu-id="c9e8d-125">Word を起動してドキュメントを作成します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-125">Start Word and create a document.</span></span>

    > [!NOTE]
    > <span data-ttu-id="c9e8d-126">この例では Word を使用していますが、Office アドインをサポートしている任意の Office アプリケーションを使用できます (Excel、Outlook、PowerPoint、Project など)。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-126">Although this example uses Word, you can use any Office application that supports Office Add-ins such as Excel, Outlook, PowerPoint, or Project.</span></span>
    
2.  <span data-ttu-id="c9e8d-127">**[ファイル]**  >  **[オプション]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-127">Choose **File** > **Options**.</span></span>
    
3.  <span data-ttu-id="c9e8d-128">**[Word オプション]** ダイアログ ボックスで、**[セキュリティ センター]** をクリックして、**[セキュリティ センターの設定]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-128">In the **Word Options** dialog box, choose **Trust Center** and then choose **Trust Center Settings**.</span></span> 
    
4.  <span data-ttu-id="c9e8d-p103">**[セキュリティ センター]** ダイアログ ボックスで、**[信頼できるアドイン カタログ]** をクリックします。**[カタログの URL]** として、前の手順で作成したファイル共有の汎用名前付け規則 (UNC) パス (たとえば、\\\YourMachineName\AddinManifests) を入力して、**[カタログの追加]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-p103">In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**. Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**.</span></span> 
    
5. <span data-ttu-id="c9e8d-131">**[メニューに表示する]** チェック ボックスをオンにします。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-131">Select the check box for **Show in Menu**.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="c9e8d-132">信頼できるアドイン カタログとして指定されている共有にアドインの XML マニフェスト ファイルを保存すると、そのアドインは、ユーザーがリボンの **[挿入]** タブから **[個人用アドイン]** をクリックしたときに、**[Office アドイン]** ダイアログ ボックスの **[共有フォルダー]** に表示されるようになります。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-132">When you store an add-in XML manifest file on a share that is specified as a trusted web add-in catalog, the add-in appears under **Shared Folder** in the **Office Add-ins** dialog box when the user navigates to the **Insert** tab in the ribbon and chooses **My Add-ins**.</span></span>

6. <span data-ttu-id="c9e8d-133">Word を終了します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-133">Close Word.</span></span>

## <a name="step-3-create-a-web-app-in-azure"></a><span data-ttu-id="c9e8d-134">手順 3: Azure に Web アプリを作成する</span><span class="sxs-lookup"><span data-stu-id="c9e8d-134">Step 3: Create a web app in Azure</span></span>

<span data-ttu-id="c9e8d-135">Azure に空の Web アプリを作成するには、[Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) または [Azure ポータル](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal)のどちらかを使用します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-135">Create an empty web app in Azure either by using [Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) or by using the [Azure portal](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal).</span></span>

### <a name="using-visual-studio-2017"></a><span data-ttu-id="c9e8d-136">Visual Studio 2017 を使用する場合</span><span class="sxs-lookup"><span data-stu-id="c9e8d-136">Using Visual Studio 2017</span></span>

<span data-ttu-id="c9e8d-137">Visual Studio 2017 を使用して Web アプリを作成するには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-137">To create the web app using Visual Studio 2017, complete the following steps.</span></span>

1. <span data-ttu-id="c9e8d-p104">Visual Studio の **[表示]** メニューで、**[サーバー エクスプローラー]** をクリックします。**[Azure]** を右クリックして、**[Microsoft Azure サブスクリプションへの接続]** をクリックします。Azure サブスクリプションに接続するための指示に従います。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-p104">In Visual Studio, in the **View** menu, choose **Server Explorer**. Right-click **Azure** and choose **Connect to Microsoft Azure subscription**. Follow the instructions for connecting to your Azure subscription.</span></span>
    
2. <span data-ttu-id="c9e8d-141">Visual Studio の **[サーバー エクスプローラー]** で、**[Azure]** を展開し、**[App Service]** を右クリックして **[新しい App Service の作成]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-141">In Visual Studio, in **Server Explorer**, expand **Azure**, right-click **App Service**, and then choose **Create New App Service**.</span></span>
    
3. <span data-ttu-id="c9e8d-142">**[App Service の作成]** ダイアログ ボックスで、次の情報を指定します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-142">In the **Create App Service** dialog box, provide this information:</span></span>
    
      - <span data-ttu-id="c9e8d-p105">サイトの一意の **[Web アプリの名前]** を入力します。Azure は、サイト名が azurewebsites.net ドメイン全体で一意であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-p105">Enter a unique **Web App Name** for your site. Azure verifies that the site name is unique across the azurewebsites.net domain.</span></span>

      - <span data-ttu-id="c9e8d-145">このサイトの作成に使用する **[サブスクリプション]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-145">Choose the **Subscription** to use for creating this site.</span></span>

      - <span data-ttu-id="c9e8d-p106">サイトの **[リソース グループ]** を選択します。新しいグループを作成する場合は、そのグループに名前を指定する必要もあります。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-p106">Choose the **Resource Group** for your site. If you create a new group, you also need to name it.</span></span>
    
      - <span data-ttu-id="c9e8d-p107">このサイトの作成に使用する **[App Service プラン]** を選択します。新しいプランを作成する場合は、そのプランに名前を指定する必要もあります。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-p107">Choose the **App Service Plan** to use for creating this site. If you create a new plan, you also need to name it.</span></span>
       
      - <span data-ttu-id="c9e8d-150">**[作成]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-150">Choose **Create**.</span></span>

    <span data-ttu-id="c9e8d-151">新しい Web アプリが、**[サーバー エクスプローラー]** の **[Azure]** >> **[App Service]** >> (選択したリソース グループ) に表示されます。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-151">The new web app appears in **Server Explorer** under **Azure** >> **App Service** >> (the chosen resouce group).</span></span>
    
4. <span data-ttu-id="c9e8d-p108">新しい Web アプリを右クリックして、**[ブラウザーで表示]** をクリックします。ブラウザーが開いて、「App Service アプリが作成されました」というメッセージを示す Web ページが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-p108">Right-click the new web app and then choose **View in Browser**. Your browser opens and displays a webpage with the message "Your App Service app has been created."</span></span>
    
5. <span data-ttu-id="c9e8d-154">ブラウザーのアドレス バーで、HTTPS を使用するように Web アプリの URL を変更してから **Enter** キーを押して、HTTPS プロトコルが有効であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-154">In the browser address bar, change the URL for the web app so that it uses HTTPS and press **Enter** to confirm that the HTTPS protocol is enabled.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="c9e8d-155">Azure の Web サイトは自動的に HTTPS エンドポイントを提供します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-155">Azure websites automatically provide an HTTPS endpoint.</span></span>
    
### <a name="using-the-azure-portal"></a><span data-ttu-id="c9e8d-156">Azure ポータルを使用する場合</span><span class="sxs-lookup"><span data-stu-id="c9e8d-156">Using the Azure portal</span></span>

<span data-ttu-id="c9e8d-157">Azure ポータルを使用して Web アプリケーションを作成するには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-157">To create the web app using the Azure portal, complete the following steps.</span></span>

1. <span data-ttu-id="c9e8d-158">Azure の資格情報を使用して、[Azure ポータル](https://portal.azure.com/)にログオンします。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-158">Log on to the [Azure portal](https://portal.azure.com/) using your Azure credentials.</span></span>
    
2. <span data-ttu-id="c9e8d-159">**[新規]** > **[Web + モバイル]** > **[Web アプリ]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-159">Choose **New** > **Web + Mobile** > **Web App**.</span></span> 

3. <span data-ttu-id="c9e8d-160">**[Web アプリの作成]** ダイアログ ボックスで、次の情報を指定します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-160">In the **Web App Create** dialog box, provide this information:</span></span>
    
      - <span data-ttu-id="c9e8d-p109">サイトの一意の **[アプリ名]** を入力します。Azure は、サイト名が azureweb apps.net ドメイン全体で一意であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-p109">Enter a unique **App name** for your site. Azure verifies that the site name is unique across the azureweb apps.net domain.</span></span>

      - <span data-ttu-id="c9e8d-163">このサイトの作成に使用する **[サブスクリプション]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-163">Choose the **Subscription** to use for creating this site.</span></span>

      - <span data-ttu-id="c9e8d-p110">サイトの **[リソース グループ]** を選択します。新しいグループを作成する場合は、そのグループに名前を指定する必要もあります。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-p110">Choose the **Resource Group** for your site. If you create a new group, you also need to name it.</span></span>

      - <span data-ttu-id="c9e8d-166">サイトの **[OS]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-166">Choose the **OS** for your site.</span></span>
    
      - <span data-ttu-id="c9e8d-p111">このサイトの作成用に使用する **[App Service プラン]** を選択します。新しいプランを作成する場合は、そのプランに名前を指定する必要もあります。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-p111">Choose the **App Service plan** to use for creating this site. If you create a new plan, you also need to name it.</span></span>
       
      - <span data-ttu-id="c9e8d-169">**[作成]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-169">Choose **Create**.</span></span>

4. <span data-ttu-id="c9e8d-170">**[通知]** (Azure ポータルの上辺に配置されているベル アイコン) をクリックし、**[デプロイメントが成功しました]** の通知をクリックして Azure ポータルでサイトの **[概要]** ページを開きます。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-170">Choose **Notifications** (the bell icon that is located along the top edge of the Azure portal) and then choose the **Deployments succeeded** notification to open the site's **Overview** page in the Azure portal.</span></span>

    > [!NOTE]
    > <span data-ttu-id="c9e8d-171">この通知は、サイトのデプロイが完了した時点で **[デプロイは進行中です]** から **[デプロイメントが成功しました]** に変化します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-171">The notification will change from **Deployment in progress** to **Deployments succeeded** when the site deployment completes.</span></span>

5. <span data-ttu-id="c9e8d-p112">Azure ポータルのサイトの **[概要]** ページにある **[要点]** セクションで、**[URL]** の下に表示されている URL を選択します。ブラウザーが開いて、「App Service アプリが作成されました」というメッセージを示す Web ページが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-p112">In the **Essentials** section of the site's **Overview** page in the Azure portal, choose the URL that is displayed under **URL**. Your browser opens and displays a webpage with the message "Your App Service app has been created."</span></span> 
    
6. <span data-ttu-id="c9e8d-174">ブラウザーのアドレス バーで、HTTPS を使用するように Web アプリの URL を変更してから **Enter** キーを押して、HTTPS プロトコルが有効であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-174">In the browser address bar, change the URL for the web app so that it uses HTTPS and press **Enter** to confirm that the HTTPS protocol is enabled.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="c9e8d-175">Azure の Web サイトは自動的に HTTPS エンドポイントを提供します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-175">Azure websites automatically provide an HTTPS endpoint.</span></span>

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a><span data-ttu-id="c9e8d-176">手順 4: Visual Studio で Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="c9e8d-176">Step 4: Create an Office Add-in in Visual Studio</span></span>

1. <span data-ttu-id="c9e8d-177">管理者として Visual Studio を起動します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-177">Start Visual Studio as an administrator.</span></span>
    
2. <span data-ttu-id="c9e8d-178">**[ファイル]** > **[新規作成]** > **[プロジェクト]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-178">Choose **File** > **New** > **Project**.</span></span>
    
3. <span data-ttu-id="c9e8d-179">**[テンプレート]** の **[Visual C#]** (または **[Visual Basic]**) を展開し、**[Office/SharePoint]** を展開して **[アドイン]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-179">Under **Templates**, expand **Visual C#** (or **Visual Basic**), expand **Office/SharePoint**, and then choose **Add-ins**.</span></span>
    
4. <span data-ttu-id="c9e8d-180">**[Word Web アドイン]** を選択してから、**[OK]** をクリックして、既定の設定を使用します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-180">Choose **Word Web Add-in**, and then choose **OK** to accept the default settings.</span></span>
       
<span data-ttu-id="c9e8d-181">Visual Studio により、基本的な Word アドインが作成されます。このアドインは、Web プロジェクトに一切変更を加えることなく、そのままの状態で発行できます。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-181">Visual Studio creates a basic Word add-in that you'll be able to publish as-is, without making any changes to its web project.</span></span>

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a><span data-ttu-id="c9e8d-182">手順 5: Azure に Office アドイン Web アプリを発行する</span><span class="sxs-lookup"><span data-stu-id="c9e8d-182">Step 5: Publish your Office Add-in web app to Azure</span></span>

1. <span data-ttu-id="c9e8d-183">Visual Studio でアドイン プロジェクトを開き、**[ソリューション エクスプローラー]** のソリューション ノードを展開して、ソリューションの両方のプロジェクトが表示されるようにします。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-183">With your add-in project open in Visual Studio, expand the solution node in **Solution Explorer** so that you see both projects for the solution.</span></span>
    
2. <span data-ttu-id="c9e8d-p113">Web プロジェクトを右クリックして、**[発行]** をクリックします。Web プロジェクトには Office アドイン Web アプリのファイルが含まれているため、このプロジェクトを Azure に発行することになります。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-p113">Right-click the web project and then choose **Publish**. The web project contains Office Add-in web app files so this is the project that you publish to Azure.</span></span>
    
3. <span data-ttu-id="c9e8d-186">**[発行]** タブで、次の操作を実行します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-186">On the **Publish** tab:</span></span>

      - <span data-ttu-id="c9e8d-187">**[Microsoft Azure App Service]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-187">Choose **Microsoft Azure App Service**.</span></span>
      
      - <span data-ttu-id="c9e8d-188">**[既存のものを選択]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-188">Choose **Select Existing**.</span></span>

      - <span data-ttu-id="c9e8d-189">**[発行]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-189">Choose **Publish**.</span></span> 

6. <span data-ttu-id="c9e8d-190">**[App Service]** ダイアログ ボックスで、「[手順 3: Azure に Web アプリを作成する](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure)」で作成した Web アプリを見つけて、そのアプリを選択してから **[OK]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-190">In the **App Service** dialog box, find and choose the web app that you created in [Step 3: Create a web app in Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure) and then choose **OK**.</span></span> 

    <span data-ttu-id="c9e8d-p114">Visual Studio により、Office アドインの Web プロジェクトが Azure Web アプリに発行されます。Visual Studio による Web プロジェクトの発行が完了すると、ブラウザーが開いて、「App Service アプリが作成されました」というテキストを示す Web ページが表示されます。これは、Web アプリの現在の既定のページです。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-p114">Visual Studio publishes the web project for your Office Add-in to your Azure web app. When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created." This is the current default page for the web app.</span></span>

7. <span data-ttu-id="c9e8d-194">アドインの Web ページを確認するには、そのアドインが HTTPS を使用するように URL を変更して、アドインの HTML ページのパスを指定します (たとえば https://YourDomain.azurewebsites.net/Home.html))。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-194">To see the webpage for your add-in, change the URL so that it uses HTTPS and specifies the path of your add-in's HTML page (for example: https://YourDomain.azurewebsites.net/Home.html).</span></span> <span data-ttu-id="c9e8d-195">こうすることで、アドインの Web アプリが Azure でホストされるようになったことを確認します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-195">This confirms that your add-in's web app is now hosted on Azure.</span></span> <span data-ttu-id="c9e8d-196">ルート URL (たとえば https://YourDomain.azurewebsites.net)) をコピーします。この URL は、アドイン マニフェスト ファイルの編集時に必要になります。これについては、この記事で説明します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-196">Copy the root URL (for example: https://YourDomain.azurewebsites.net); you'll need it when you edit the add-in manifest file later in this article.</span></span>
    
## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a><span data-ttu-id="c9e8d-197">手順 6: アドインの XML マニフェスト ファイルを編集して展開する</span><span class="sxs-lookup"><span data-stu-id="c9e8d-197">Step 6: Edit and deploy the add-in XML manifest file</span></span>

1. <span data-ttu-id="c9e8d-198">Visual Studio の **[ソリューション エクスプローラー]** でサンプルの Office アドインを開いて、ソリューションを展開し、両方のプロジェクトが表示されるようにします。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-198">In Visual Studio with the sample Office Add-in open in **Solution Explorer**, expand the solution so that both projects show.</span></span>
    
2. <span data-ttu-id="c9e8d-p116">Office アドイン プロジェクト (たとえば、WordWebAddIn) を展開し、マニフェスト フォルダーを右クリックして **[開く]** をクリックします。アドインの XML マニフェスト ファイルが開きます。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-p116">Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**. The add-in XML manifest file opens.</span></span>
    
3. <span data-ttu-id="c9e8d-201">XML マニフェスト ファイルで、"~remoteAppUrl" というインスタンスをすべて検索して、Azure のアドイン Web アプリのルート URL に置換します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-201">In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure.</span></span> <span data-ttu-id="c9e8d-202">この URL は、前の手順で Azure にアドイン Web アプリを発行した後にコピーしたものです (たとえば https://YourDomain.azurewebsites.net))。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-202">This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span></span> 
    
4. <span data-ttu-id="c9e8d-p118">**[ファイル]** をクリックして、**[すべてを保存]** をクリックします。アドインの XML マニフェスト ファイルを閉じます。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-p118">Choose **File** and then choose **Save All**. Close the add-in XML manifest file.</span></span>
    
5. <span data-ttu-id="c9e8d-205">**[ソリューション エクスプローラー]** に戻って、マニフェストのフォルダーを右クリックして、**[エクスプローラーでフォルダーを開く]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-205">Back in **Solution Explorer**, right-click the manifest folder and choose **Open Folder In File Explorer**.</span></span>
    
6. <span data-ttu-id="c9e8d-206">アドインの XML マニフェスト ファイル (たとえば、WordWebAddIn.xml) をコピーします。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-206">Copy the add-in XML manifest file (for example, WordWebAddIn.xml).</span></span> 
    
7. <span data-ttu-id="c9e8d-207">「[手順 1: 共有フォルダーを作成する](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file)」で作成したネットワーク ファイル共有を参照して、そのフォルダー内にマニフェスト ファイルを貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-207">Browse to the network file share that you created in [Step 1: Create a shared folder](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) and paste the manifest file into the folder.</span></span>

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a><span data-ttu-id="c9e8d-208">手順 7: Office クライアント アプリケーションにアプリを挿入し、実行する</span><span class="sxs-lookup"><span data-stu-id="c9e8d-208">Step 7: Insert and run the add-in in the Office client application</span></span>

1. <span data-ttu-id="c9e8d-209">Word を起動してドキュメントを作成します。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-209">Start Word 2016 and create a document.</span></span>
    
2. <span data-ttu-id="c9e8d-210">リボンで、**[挿入]** > **[個人用アドイン]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-210">On the ribbon, choose **Insert** > **My Add-ins**.</span></span> 
    
3. <span data-ttu-id="c9e8d-p119">**[Office アドイン]** ダイアログ ボックスで、**[共有フォルダー]** をクリックします。Word により、信頼できるアドイン カタログとしてリストしたフォルダー (「[手順 2: 信頼できるアドイン カタログにファイル共有を追加する](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)」で指定したもの) がスキャンされ、アドインがダイアログ ボックスに表示されます。サンプル アドインのアイコンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-p119">In the **Office Add-ins** dialog box, choose **SHARED FOLDER**. Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box. You should see an icon for your sample add-in.</span></span>
    
4. <span data-ttu-id="c9e8d-p120">アドインを選択して、**[追加]** をクリックします。リボンに、そのアドインの **[作業ウィンドウの表示]** ボタンが追加されます。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-p120">Choose the icon for your add-in and then choose **Add**. A **Show Taskpane** button for your add-in is added to the ribbon.</span></span> 

5. <span data-ttu-id="c9e8d-p121">**[ホーム]** タブのリボンで、**[作業ウィンドウの表示]** ボタンをクリックします。現在のドキュメントの右側の作業ウィンドウ内でアドインが開きます。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-p121">On the ribbon of the **Home** tab, choose the **Show Taskpane** button. The add-in opens in a task pane to the right of the current document.</span></span>
    
6. <span data-ttu-id="c9e8d-p122">アドインが動作していることを確認するために、ドキュメント内のテキストを選択して、作業ウィンド内の **[Highlight!]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="c9e8d-p122">Verify that the add-in works by selecting some text in the document and choosing the **Highlight!** button in the task pane.</span></span> 

## <a name="see-also"></a><span data-ttu-id="c9e8d-220">関連項目</span><span class="sxs-lookup"><span data-stu-id="c9e8d-220">See also</span></span>

- [<span data-ttu-id="c9e8d-221">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="c9e8d-221">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="c9e8d-222">発行のための準備として Visual Studio を使用してアドインをパッケージ化する</span><span class="sxs-lookup"><span data-stu-id="c9e8d-222">Package your add-in using Visual Studio to prepare for publishing</span></span>](../publish/package-your-add-in-using-visual-studio.md)
    

---
title: Microsoft Azure で Office アドインをホストする | Microsoft Docs
description: アドイン Web アプリを Azure に展開して、Office クライアント アプリケーションでテストのためにアドインをサイドロードする方法について説明します。
ms.date: 10/16/2019
localization_priority: Normal
ms.openlocfilehash: 3488217da5aafe108ed9d38c1c4cfe415424d41a
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950657"
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a><span data-ttu-id="5e297-103">Microsoft Azure で Office アドインをホストする</span><span class="sxs-lookup"><span data-stu-id="5e297-103">Host an Office Add-in on Microsoft Azure</span></span>

<span data-ttu-id="5e297-p101">最も簡単な Office アドインは、XML マニフェスト ファイルと HTML ページで成り立っています。XML マニフェスト ファイルには、アドインの特性 (アドインの名前や実行可能な Office クライアント アプリケーションの種類、アドインの HTML ページの URL など) を記述します。HTML ページには、ユーザーが Office クライアント アプリケーションにアドインをインストールして実行したときに操作する、Web アプリが含まれています。Office アドインの Web アプリは、Azure を含む、あらゆる Web ホスティング プラットフォームでホストできます。</span><span class="sxs-lookup"><span data-stu-id="5e297-p101">The simplest Office Add-in is made up of an XML manifest file and an HTML page. The XML manifest file describes the add-in's characteristics, such as its name, what Office client applications it can run in, and the URL for the add-in's HTML page. The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application. You can host the web app of an Office Add-in on any web hosting platform, including Azure.</span></span>

<span data-ttu-id="5e297-108">この記事では、アドイン Web アプリを Azure に展開して、Office クライアント アプリケーションでテストのために[アドインをサイドロード](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="5e297-108">This article describes how to deploy an add-in web app to Azure and [sideload the add-in](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) for testing in an Office client application.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="5e297-109">前提条件</span><span class="sxs-lookup"><span data-stu-id="5e297-109">Prerequisites</span></span> 

1. <span data-ttu-id="5e297-110">[Visual Studio 2019](https://www.visualstudio.com/downloads) をインストールし、**Azure 開発**ワークロードを含めるよう選択します。</span><span class="sxs-lookup"><span data-stu-id="5e297-110">Install [Visual Studio 2019](https://www.visualstudio.com/downloads) and choose to include the **Azure development** workload.</span></span>

    > [!NOTE]
    > <span data-ttu-id="5e297-111">既に Visual Studio 2019 がインストールされている場合は、[Visual Studio インストーラーを使用](/visualstudio/install/modify-visual-studio)して、**Azure 開発**ワークロードがインストールされていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="5e297-111">If you've previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Azure development** workload is installed.</span></span> 

2. <span data-ttu-id="5e297-112">Office をインストールする。</span><span class="sxs-lookup"><span data-stu-id="5e297-112">Install Office.</span></span>

    > [!NOTE]
    > <span data-ttu-id="5e297-113">まだ Office を所持していない場合は、[1 か月間無料試用版の登録](https://products.office.com/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735)が可能です。</span><span class="sxs-lookup"><span data-stu-id="5e297-113">If you don't already have Office, you can [register for a free 1-month trial](https://products.office.com/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).</span></span>

3. <span data-ttu-id="5e297-114">Azure サブスクリプションを取得します。</span><span class="sxs-lookup"><span data-stu-id="5e297-114">Obtain an Azure subscription.</span></span>

    > [!NOTE]
    > <span data-ttu-id="5e297-115">まだ Azure サブスクリプションを所持していない場合、このサブスクリプションは [Visual Studio サブスクリプションの一部として取得](https://azure.microsoft.com/pricing/member-offers/visual-studio-subscriptions/)できます。また、[無料試用版の登録](https://azure.microsoft.com/pricing/free-trial)も可能です。</span><span class="sxs-lookup"><span data-stu-id="5e297-115">If don't already have an Azure subscription, you can [get one as part of your Visual Studio subscription](https://azure.microsoft.com/pricing/member-offers/visual-studio-subscriptions/) or [register for a free trial](https://azure.microsoft.com/pricing/free-trial).</span></span> 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a><span data-ttu-id="5e297-116">手順 1: アドインの XML マニフェスト ファイルをホストするための共有フォルダーを作成する</span><span class="sxs-lookup"><span data-stu-id="5e297-116">Step 1: Create a shared folder to host your add-in XML manifest file</span></span>

1. <span data-ttu-id="5e297-117">開発用のコンピューターでエクスプローラーを開きます。</span><span class="sxs-lookup"><span data-stu-id="5e297-117">Open File Explorer on your development computer.</span></span>

2. <span data-ttu-id="5e297-118">C:\ ドライブを右クリックして、**[新規]** > **[フォルダー]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="5e297-118">Right-click the C:\ drive and then choose **New** > **Folder**.</span></span>

3. <span data-ttu-id="5e297-119">新規フォルダーに「AddinManifests」という名前を付けます。</span><span class="sxs-lookup"><span data-stu-id="5e297-119">Name the new folder AddinManifests.</span></span>

4. <span data-ttu-id="5e297-120">[AddinManifests] フォルダーを右クリックして、**[共有相手]** > **[特定の人]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="5e297-120">Right-click the AddinManifests folder and then choose **Share with** > **Specific people**.</span></span>

5. <span data-ttu-id="5e297-121">**[ファイル共有]** で、ドロップダウンの矢印をクリックして、**[すべてのユーザー]** > **[追加]** > **[共有]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="5e297-121">In **File Sharing**, choose the drop-down arrow and then choose **Everyone** > **Add** > **Share**.</span></span>

> [!NOTE]
> <span data-ttu-id="5e297-p102">このチュートリアルでは、信頼できるカタログとしてローカルのファイル共有を使用します。アドインの XML マニフェスト ファイルは、この場所に保存することになります。現実のシナリオでは、[SharePoint カタログに XML マニフェスト ファイルを展開](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)するか、[AppSource にアドインを発行](/office/dev/store/submit-to-appsource-via-partner-center)することもできます。</span><span class="sxs-lookup"><span data-stu-id="5e297-p102">In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file. In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](/office/dev/store/submit-to-appsource-via-partner-center).</span></span>

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a><span data-ttu-id="5e297-124">手順 2: 信頼できるアドイン カタログにファイル共有を追加する</span><span class="sxs-lookup"><span data-stu-id="5e297-124">Step 2: Add the file share to the Trusted Add-ins catalog</span></span>

1. <span data-ttu-id="5e297-125">Word を起動してドキュメントを作成します。</span><span class="sxs-lookup"><span data-stu-id="5e297-125">Start Word and create a document.</span></span>

    > [!NOTE]
    > <span data-ttu-id="5e297-126">この例では Word を使用していますが、Office アドインをサポートしている任意の Office アプリケーションを使用できます (Excel、Outlook、PowerPoint、Project など)。</span><span class="sxs-lookup"><span data-stu-id="5e297-126">Although this example uses Word, you can use any Office application that supports Office Add-ins such as Excel, Outlook, PowerPoint, or Project.</span></span>

2. <span data-ttu-id="5e297-127">**[ファイル]**  >  **[オプション]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="5e297-127">Choose **File** > **Options**.</span></span>

3. <span data-ttu-id="5e297-128">**[Word オプション]** ダイアログ ボックスで、**[セキュリティ センター]** をクリックして、**[セキュリティ センターの設定]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="5e297-128">In the **Word Options** dialog box, choose **Trust Center** and then choose **Trust Center Settings**.</span></span>

4. <span data-ttu-id="5e297-p103">**[セキュリティ センター]** ダイアログ ボックスで、**[信頼できるアドイン カタログ]** をクリックします。**[カタログの URL]** として、前の手順で作成したファイル共有の汎用名前付け規則 (UNC) パス (たとえば、\\\YourMachineName\AddinManifests) を入力して、**[カタログの追加]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="5e297-p103">In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**. Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**.</span></span> 

5. <span data-ttu-id="5e297-131">**[メニューに表示する]** チェック ボックスをオンにします。</span><span class="sxs-lookup"><span data-stu-id="5e297-131">Select the check box for **Show in Menu**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="5e297-132">信頼できるアドイン カタログとして指定されている共有にアドインの XML マニフェスト ファイルを保存すると、そのアドインは、ユーザーがリボンの **[挿入]** タブから **[個人用アドイン]** をクリックしたときに、**[Office アドイン]** ダイアログ ボックスの **[共有フォルダー]** に表示されるようになります。</span><span class="sxs-lookup"><span data-stu-id="5e297-132">When you store an add-in XML manifest file on a share that is specified as a trusted web add-in catalog, the add-in appears under **Shared Folder** in the **Office Add-ins** dialog box when the user navigates to the **Insert** tab in the ribbon and chooses **My Add-ins**.</span></span>

6. <span data-ttu-id="5e297-133">Word を終了します。</span><span class="sxs-lookup"><span data-stu-id="5e297-133">Close Word.</span></span>

## <a name="step-3-create-a-web-app-in-azure-using-the-azure-portal"></a><span data-ttu-id="5e297-134">手順 3: Azure ポータルを使用して Azure で Web アプリを作成する</span><span class="sxs-lookup"><span data-stu-id="5e297-134">Step 3: Create a web app in Azure using the Azure portal</span></span>

<span data-ttu-id="5e297-135">Azure ポータルを使用して Web アプリケーションを作成するには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="5e297-135">To create the web app using the Azure portal, complete the following steps.</span></span>

1. <span data-ttu-id="5e297-136">Azure の資格情報を使用して、[Azure ポータル](https://portal.azure.com/)にログオンします。</span><span class="sxs-lookup"><span data-stu-id="5e297-136">Log on to the [Azure portal](https://portal.azure.com/) using your Azure credentials.</span></span>

2. <span data-ttu-id="5e297-137">[**Azure サービス**] で [**Web アプリ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="5e297-137">Under **Azure Services** select **Web Apps**.</span></span>

3. <span data-ttu-id="5e297-138">[**App Service**] ページで、[**追加**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="5e297-138">On the **App Service** page, select **Add**.</span></span> <span data-ttu-id="5e297-139">この情報を提供してください:</span><span class="sxs-lookup"><span data-stu-id="5e297-139">Provide this information:</span></span>

      - <span data-ttu-id="5e297-140">このサイトの作成に使用する **[サブスクリプション]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="5e297-140">Choose the **Subscription** to use for creating this site.</span></span>
      
      - <span data-ttu-id="5e297-p105">サイトの **[リソース グループ]** を選択します。新しいグループを作成する場合は、そのグループに名前を指定する必要もあります。</span><span class="sxs-lookup"><span data-stu-id="5e297-p105">Choose the **Resource Group** for your site. If you create a new group, you also need to name it.</span></span>
      
      - <span data-ttu-id="5e297-143">サイトの一意の **[アプリ名]** を入力します。</span><span class="sxs-lookup"><span data-stu-id="5e297-143">Enter a unique **App name** for your site.</span></span> <span data-ttu-id="5e297-144">Azure は、サイト名が azureweb apps.net ドメイン全体で一意であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="5e297-144">Azure verifies that the site name is unique across the azureweb apps.net domain.</span></span>

      - <span data-ttu-id="5e297-145">コードを使用して発行するか、Docker コンテナを使用して発行するかを選択します。</span><span class="sxs-lookup"><span data-stu-id="5e297-145">Choose whether to publish using code or a docker container.</span></span>

      - <span data-ttu-id="5e297-146">**ランタイム スタック**を指定します。</span><span class="sxs-lookup"><span data-stu-id="5e297-146">Specify a **Runtime stack**.</span></span>

      - <span data-ttu-id="5e297-147">サイトの **OS** を選択します。</span><span class="sxs-lookup"><span data-stu-id="5e297-147">Choose the **OS** for your site.</span></span>

      - <span data-ttu-id="5e297-148">**地域**を選択します。</span><span class="sxs-lookup"><span data-stu-id="5e297-148">Choose a **Region**.</span></span>

      - <span data-ttu-id="5e297-149">このサイトの作成に使用する [**App Service プラン**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="5e297-149">Choose the **App Service plan** to use for creating this site.</span></span>

      - <span data-ttu-id="5e297-150">[**作成**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="5e297-150">Choose **Create**.</span></span>

4. <span data-ttu-id="5e297-151">次のページでは、展開が進行中であること、完了したことが通知されます。</span><span class="sxs-lookup"><span data-stu-id="5e297-151">The next page will let you know that your deployment is underway and when it completes.</span></span> <span data-ttu-id="5e297-152">完了したら、[**リソースに移動**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="5e297-152">When it is completed, select **Go to resource**.</span></span>  

5. <span data-ttu-id="5e297-153">[**概要**] セクションで、[**URL**] の下に表示される URL を選択します。</span><span class="sxs-lookup"><span data-stu-id="5e297-153">In the **Overview** section, choose the URL that is displayed under **URL**.</span></span> <span data-ttu-id="5e297-154">ブラウザが開き、"App Service アプリが起動し、実行中です" というメッセージを含む Web ページが表示されます。</span><span class="sxs-lookup"><span data-stu-id="5e297-154">Your browser opens and displays a webpage with the message "Your App Service app is up and running."</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="5e297-155">Azure の Web サイトは自動的に HTTPS エンドポイントを提供します。</span><span class="sxs-lookup"><span data-stu-id="5e297-155">Azure websites automatically provide an HTTPS endpoint.</span></span>

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a><span data-ttu-id="5e297-156">手順 4: Visual Studio で Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="5e297-156">Step 4: Create an Office Add-in in Visual Studio</span></span>

1. <span data-ttu-id="5e297-157">管理者として Visual Studio を起動します。</span><span class="sxs-lookup"><span data-stu-id="5e297-157">Start Visual Studio as an administrator.</span></span>

2. <span data-ttu-id="5e297-158">[**新規プロジェクトの作成**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="5e297-158">Choose **Create a new project**.</span></span>

3. <span data-ttu-id="5e297-159">検索ボックスを使用して、**アドイン**と入力します。</span><span class="sxs-lookup"><span data-stu-id="5e297-159">Using the search box, enter **add-in**.</span></span>

4. <span data-ttu-id="5e297-160">プロジェクト タイプとして **Word Web アドイン**を選択し、[**次へ**] を選択して規定の設定を使用します。</span><span class="sxs-lookup"><span data-stu-id="5e297-160">Choose **Word Web Add-in** as the project type, and then choose **Next** to accept the default settings.</span></span>

<span data-ttu-id="5e297-161">Visual Studio は、Web プロジェクトに変更を加えることなくそのまま発行できる、基本的な Word アドインを作成します。</span><span class="sxs-lookup"><span data-stu-id="5e297-161">Visual Studio creates a basic Word add-in that you'll be able to publish as-is, without making any changes to its web project.</span></span> <span data-ttu-id="5e297-162">Excel などの異なる Office ホスト タイプのアドインを作成するには、この手順を繰り返して、目的の Office ホストのプロジェクト タイプを選択します。</span><span class="sxs-lookup"><span data-stu-id="5e297-162">To make an add-in for a different Office host type, such as Excel, repeat the steps and choose a project type with your desired Office host.</span></span>

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a><span data-ttu-id="5e297-163">手順 5: Azure に Office アドイン Web アプリを発行する</span><span class="sxs-lookup"><span data-stu-id="5e297-163">Step 5: Publish your Office Add-in web app to Azure</span></span>

1. <span data-ttu-id="5e297-164">アドイン プロジェクトを Visual Studio で開いた状態で、**ソリューション エクスプローラー**でソリューション ノードを展開し、**App Service** を選択します。</span><span class="sxs-lookup"><span data-stu-id="5e297-164">With your add-in project open in Visual Studio, expand the solution node in **Solution Explorer**, then select **App Service**.</span></span>

2. <span data-ttu-id="5e297-p110">Web プロジェクトを右クリックして、**[発行]** をクリックします。Web プロジェクトには Office アドイン Web アプリのファイルが含まれているため、このプロジェクトを Azure に発行することになります。</span><span class="sxs-lookup"><span data-stu-id="5e297-p110">Right-click the web project and then choose **Publish**. The web project contains Office Add-in web app files so this is the project that you publish to Azure.</span></span>

3. <span data-ttu-id="5e297-167">**[発行]** タブで、次の操作を実行します。</span><span class="sxs-lookup"><span data-stu-id="5e297-167">On the **Publish** tab:</span></span>

      - <span data-ttu-id="5e297-168">**[Microsoft Azure App Service]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="5e297-168">Choose **Microsoft Azure App Service**.</span></span>

      - <span data-ttu-id="5e297-169">**[既存のものを選択]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="5e297-169">Choose **Select Existing**.</span></span>

      - <span data-ttu-id="5e297-170">**[発行]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="5e297-170">Choose **Publish**.</span></span>

4. <span data-ttu-id="5e297-p111">Visual Studio により、Office アドインの Web プロジェクトが Azure Web アプリに発行されます。Visual Studio による Web プロジェクトの発行が完了すると、ブラウザーが開いて、「App Service アプリが作成されました」というテキストを示す Web ページが表示されます。これは、Web アプリの現在の既定のページです。</span><span class="sxs-lookup"><span data-stu-id="5e297-p111">Visual Studio publishes the web project for your Office Add-in to your Azure web app. When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created." This is the current default page for the web app.</span></span>

5. <span data-ttu-id="5e297-174">ルート URL (たとえば https://YourDomain.azurewebsites.net)) をコピーします。この URL は、アドイン マニフェスト ファイルの編集時に必要になります。これについては、この記事で説明します。</span><span class="sxs-lookup"><span data-stu-id="5e297-174">Copy the root URL (for example: https://YourDomain.azurewebsites.net); you'll need it when you edit the add-in manifest file later in this article.</span></span>

## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a><span data-ttu-id="5e297-175">手順 6: アドインの XML マニフェスト ファイルを編集して展開する</span><span class="sxs-lookup"><span data-stu-id="5e297-175">Step 6: Edit and deploy the add-in XML manifest file</span></span>

1. <span data-ttu-id="5e297-176">Visual Studio の **[ソリューション エクスプローラー]** でサンプルの Office アドインを開いて、ソリューションを展開し、両方のプロジェクトが表示されるようにします。</span><span class="sxs-lookup"><span data-stu-id="5e297-176">In Visual Studio with the sample Office Add-in open in **Solution Explorer**, expand the solution so that both projects show.</span></span>

2. <span data-ttu-id="5e297-p112">Office アドイン プロジェクト (たとえば、WordWebAddIn) を展開し、マニフェスト フォルダーを右クリックして **[開く]** をクリックします。アドインの XML マニフェスト ファイルが開きます。</span><span class="sxs-lookup"><span data-stu-id="5e297-p112">Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**. The add-in XML manifest file opens.</span></span>

3. <span data-ttu-id="5e297-p113">XML マニフェスト ファイルで、"~remoteAppUrl" というインスタンスをすべて検索して、Azure のアドイン Web アプリのルート URL に置換します。この URL は、前の手順で Azure にアドイン Web アプリを発行した後にコピーしたものです (たとえば、https://YourDomain.azurewebsites.net)。</span><span class="sxs-lookup"><span data-stu-id="5e297-p113">In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure. This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span></span> 

4. <span data-ttu-id="5e297-181">[**ファイル**] をクリックして、[**すべてを保存**] をクリックします。</span><span class="sxs-lookup"><span data-stu-id="5e297-181">Choose **File** and then choose **Save All**.</span></span> <span data-ttu-id="5e297-182">次に、アドイン XML マニフェスト ファイル (WordWebAddIn.xml など) をコピーします。</span><span class="sxs-lookup"><span data-stu-id="5e297-182">Next, Copy the add-in XML manifest file (for example, WordWebAddIn.xml).</span></span>

5. <span data-ttu-id="5e297-183">**ファイル エクスプローラー** プログラムを使用して、「[手順 1: 共有フォルダーを作成する](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file)」で作成したネットワーク ファイル共有を参照し、マニフェスト ファイルをそのフォルダー内に貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="5e297-183">Using the **File Explorer** program, browse to the network file share that you created in [Step 1: Create a shared folder](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) and paste the manifest file into the folder.</span></span>

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a><span data-ttu-id="5e297-184">手順 7: Office クライアント アプリケーションにアプリを挿入し、実行する</span><span class="sxs-lookup"><span data-stu-id="5e297-184">Step 7: Insert and run the add-in in the Office client application</span></span>

1. <span data-ttu-id="5e297-185">Word を起動してドキュメントを作成します。</span><span class="sxs-lookup"><span data-stu-id="5e297-185">Start Word and create a document.</span></span>

2. <span data-ttu-id="5e297-186">リボンで、**[挿入]** > **[個人用アドイン]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="5e297-186">On the ribbon, choose **Insert** > **My Add-ins**.</span></span>

3. <span data-ttu-id="5e297-p115">**[Office アドイン]** ダイアログ ボックスで、**[共有フォルダー]** をクリックします。Word により、信頼できるアドイン カタログとしてリストしたフォルダー (「[手順 2: 信頼できるアドイン カタログにファイル共有を追加する](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)」で指定したもの) がスキャンされ、アドインがダイアログ ボックスに表示されます。サンプル アドインのアイコンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="5e297-p115">In the **Office Add-ins** dialog box, choose **SHARED FOLDER**. Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box. You should see an icon for your sample add-in.</span></span>

4. <span data-ttu-id="5e297-p116">アドインを選択して、**[追加]** をクリックします。リボンに、そのアドインの **[作業ウィンドウの表示]** ボタンが追加されます。</span><span class="sxs-lookup"><span data-stu-id="5e297-p116">Choose the icon for your add-in and then choose **Add**. A **Show Taskpane** button for your add-in is added to the ribbon.</span></span>

5. <span data-ttu-id="5e297-p117">**[ホーム]** タブのリボンで、**[作業ウィンドウの表示]** ボタンをクリックします。現在のドキュメントの右側の作業ウィンドウ内でアドインが開きます。</span><span class="sxs-lookup"><span data-stu-id="5e297-p117">On the ribbon of the **Home** tab, choose the **Show Taskpane** button. The add-in opens in a task pane to the right of the current document.</span></span>

6. <span data-ttu-id="5e297-p118">アドインが動作していることを確認するために、ドキュメント内のテキストを選択して、作業ウィンド内の **[Highlight!]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="5e297-p118">Verify that the add-in works by selecting some text in the document and choosing the **Highlight!** button in the task pane.</span></span>

## <a name="see-also"></a><span data-ttu-id="5e297-196">関連項目</span><span class="sxs-lookup"><span data-stu-id="5e297-196">See also</span></span>

- [<span data-ttu-id="5e297-197">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="5e297-197">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="5e297-198">Visual Studio を使用してアドインを発行する</span><span class="sxs-lookup"><span data-stu-id="5e297-198">Publish your add-in using Visual Studio</span></span>](../publish/package-your-add-in-using-visual-studio.md)

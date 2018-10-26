---
title: 発行のための準備として Visual Studio を使用してアドインをパッケージ化する| Microsoft Docs
description: この記事では、Visual Studio 2017 を使用して、Web プロジェクトを展開し、アドインをパッケージ化する方法について説明します。
ms.date: 01/25/2018
ms.openlocfilehash: 3515f88e41bc5f0af62a3b043beae5177f3291ac
ms.sourcegitcommit: c400a220783b03a739449e2d3ff00bbffe5ec7c1
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/20/2018
ms.locfileid: "25681764"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a><span data-ttu-id="3107d-103">発行のための準備として Visual Studio を使用してアドインをパッケージ化する</span><span class="sxs-lookup"><span data-stu-id="3107d-103">Package your add-in using Visual Studio to prepare for publishing</span></span>

<span data-ttu-id="3107d-p101">Office アドイン パッケージには、アドインの発行に使用する XML [マニフェスト ファイル](../develop/add-in-manifests.md)が含まれています。プロジェクトの web アプリケーション ファイルを個別に発行する必要があります。この記事では、Visual Studio 2017 を使用して、Web プロジェクトを展開し、アドインをパッケージ化する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="3107d-p101">Your Office Add-in package contains an XML [manifest file](../develop/add-in-manifests.md) that you'll use to publish the add-in. You'll have to publish the web application files of your project separately. This article describes how to deploy your web project and package your add-in by using Visual Studio 2015.</span></span>

## <a name="to-deploy-your-web-project-using-visual-studio-2017"></a><span data-ttu-id="3107d-107">Visual Studio 2017 を使用して Web プロジェクトを展開するには</span><span class="sxs-lookup"><span data-stu-id="3107d-107">To deploy your web project using Visual Studio 2015</span></span>

<span data-ttu-id="3107d-108">Visual Studio 2017 を使用して Web プロジェクトを展開する次の手順を完了します。</span><span class="sxs-lookup"><span data-stu-id="3107d-108">Complete the following steps to deploy your web project using Visual Studio 2015.</span></span>

1. <span data-ttu-id="3107d-109">**[ソリューション エクスプローラー]** で、アドイン プロジェクトのショートカット メニューを開き、**[発行]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="3107d-109">In  **Solution Explorer**, open the shortcut menu for the add-in project, and then choose  **Publish**.</span></span>
    
    <span data-ttu-id="3107d-110">[**アドインの発行**] ページが表示されます。</span><span class="sxs-lookup"><span data-stu-id="3107d-110">The  **Publish your add-in** page appears.</span></span>
    
2. <span data-ttu-id="3107d-111">**[現在のプロファイル]** ドロップダウン リストで、プロファイルを選択するか、**[新規…]** を選択して新しいプロファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="3107d-111">In the  **Current profile** drop-down list, select a profile or choose **New ...** to create a new profile.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="3107d-112">発行プロファイルでは、展開先のサーバー、サーバーへのログオンに必要な資格情報、展開するデータベース、およびその他の展開オプションを指定します。</span><span class="sxs-lookup"><span data-stu-id="3107d-112">A publish profile specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.</span></span>

    <span data-ttu-id="3107d-p102">[**新規...**] を選択した場合、[**発行プロファイルの作成**] ページとともにウィザードが表示されます。このウィザードを使用して、Microsoft Azure などの Web サイトをホストするプロバイダーから発行プロファイルをインポートするか、新しいプロファイルを作成するかして、次の手順でサーバー、資格情報、その他の設定を追加することができます。</span><span class="sxs-lookup"><span data-stu-id="3107d-p102">If you choose  **New ...**, the  **Create publishing profile** wizard appears. You can use this wizard to import a publishing profile from a web site hosting provider such as Microsoft Azure or create a new profile and add your server, credentials, and other settings in the next procedure.</span></span>
    
    <span data-ttu-id="3107d-115">発行プロファイルのインポートまたは新しい発行プロファイルの作成の詳細については、「[発行プロファイルの作成](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3107d-115">For more information about importing publishing profiles or creating new publishing profiles, see [Creating a Publish Profile](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).</span></span>
    
3. <span data-ttu-id="3107d-116">[**アドインを発行する**] ページで、[**Web プロジェクトの展開**] リンクを選択します。</span><span class="sxs-lookup"><span data-stu-id="3107d-116">In the  **Publish your add-in** page, choose the **Deploy your web project** link.</span></span>
    
    <span data-ttu-id="3107d-p103">**[Web を発行する]** ダイアログ ボックスが表示されます。このウィザードの使用方法については、「[方法: Visual Studio でワンクリック発行を使用して Web プロジェクトを展開する](https://msdn.microsoft.com/library/dd465337.aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3107d-p103">The  **Publish Web** dialog box appears. For more information about using this wizard, see [How to: Deploy a Web Project using On-Click Publishing in Visual Studio](https://msdn.microsoft.com/library/dd465337.aspx).</span></span>
    

## <a name="to-package-your-add-in-using-visual-studio-2017"></a><span data-ttu-id="3107d-119">Visual Studio 2017 を使用してアドインをパッケージ化するには</span><span class="sxs-lookup"><span data-stu-id="3107d-119">To package your add-in using Visual Studio 2015</span></span>

<span data-ttu-id="3107d-120">Visual Studio 2017 を使用してアドインをパッケージ化する次の手順を完了します。</span><span class="sxs-lookup"><span data-stu-id="3107d-120">Complete the following steps to package your add-in using Visual Studio 2015.</span></span>

1. <span data-ttu-id="3107d-121">**[アドインの発行]** ページで、**[アドインのパッケージ]** ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="3107d-121">In the **Publish your add-in** page, choose the **Package the add-in** link.</span></span>
    
    <span data-ttu-id="3107d-122">ウィザードが **[アドインのパッケージ]** ページと共に表示されます。</span><span class="sxs-lookup"><span data-stu-id="3107d-122">A wizard appears with the **Package the add-in** page.</span></span>
    
2. <span data-ttu-id="3107d-123">**[Web サイトがホストされている場所]** ボックスで、アドインのコンテンツ ファイルをホストする Web サイトの URL を入力して、**[完了]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="3107d-123">In the **Where is your website hosted?** dropdown list, select or enter the HTTPS URL of the website that will host the content files of your add-in, and then choose **Finish**.</span></span>
    
    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="3107d-124">Azure の Web サイトは自動的に HTTPS エンドポイントを提供します。</span><span class="sxs-lookup"><span data-stu-id="3107d-124">Azure websites automatically provide an HTTPS endpoint.</span></span>

    <span data-ttu-id="3107d-125">Visual Studio は、アドインの発行に必要なファイルを生成して、発行の出力フォルダーを開きます。</span><span class="sxs-lookup"><span data-stu-id="3107d-125">Visual Studio generates the files that you need to publish your add-in and then opens the publish output folder.</span></span>
    
<span data-ttu-id="3107d-126">AppSource にアドインの提出を予定している場合は、**[検証チェックの実行]** ボタンをクリックして、アドインの受け入れが阻害される問題点を識別します。</span><span class="sxs-lookup"><span data-stu-id="3107d-126">If you plan to submit your add-in to AppSource, you can choose the **Perform a validation check** link to identify any issues that will prevent your add-in from being accepted.</span></span> <span data-ttu-id="3107d-127">アドインをストアに提出する前に、すべての問題に対処してください。</span><span class="sxs-lookup"><span data-stu-id="3107d-127">You should address all issues before you submit your add-in to the store.</span></span>

<span data-ttu-id="3107d-p105">XML マニフェストを適切な場所にアップロードして[アドインを発行](../publish/publish.md)できるようになりました。XML マニフェストは、`app.publish`フォルダーの`OfficeAppManifests`にあります。たとえば、次のようになります。</span><span class="sxs-lookup"><span data-stu-id="3107d-p105">You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:</span></span>

 `%UserProfile%\Documents\Visual Studio 2017\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a><span data-ttu-id="3107d-131">関連項目</span><span class="sxs-lookup"><span data-stu-id="3107d-131">See also</span></span>

- [<span data-ttu-id="3107d-132">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="3107d-132">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="3107d-133">AppSource と Office 内でソリューションを使用できるようにする</span><span class="sxs-lookup"><span data-stu-id="3107d-133">Make your solutions available in AppSource and within Office</span></span>](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)
    

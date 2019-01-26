---
title: 発行のための準備として Visual Studio を使用してアドインをパッケージ化する | Microsoft Docs
description: Visual Studio 2017 を使用して Web プロジェクトを展開しアドインをパッケージ化する方法です。
ms.date: 01/25/2018
localization_priority: Priority
ms.openlocfilehash: a135e8e72703c3de60290a9eb7b2e03c63449124
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29386436"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a><span data-ttu-id="a1b08-103">発行のための準備として Visual Studio を使用してアドインをパッケージ化する</span><span class="sxs-lookup"><span data-stu-id="a1b08-103">Package your add-in using Visual Studio to prepare for publishing</span></span>

<span data-ttu-id="a1b08-104">Office アドイン パッケージには、アドインの発行に使用する XML [マニフェスト ファイル](../develop/add-in-manifests.md)が含まれています。</span><span class="sxs-lookup"><span data-stu-id="a1b08-104">Your Office Add-in package contains an XML [manifest file](../develop/add-in-manifests.md) that you'll use to publish the add-in.</span></span> <span data-ttu-id="a1b08-105">プロジェクトの Web アプリケーション ファイルは個別に発行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a1b08-105">You'll have to publish the web application files of your project separately.</span></span> <span data-ttu-id="a1b08-106">この記事では、Visual Studio 2017 を使用して、Web プロジェクトを展開し、アドインをパッケージ化する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="a1b08-106">This article describes how to deploy your web project and package your add-in by using Visual Studio 2017.</span></span>

## <a name="to-deploy-your-web-project-using-visual-studio-2017"></a><span data-ttu-id="a1b08-107">Visual Studio 2017 を使用して Web プロジェクトを展開するには</span><span class="sxs-lookup"><span data-stu-id="a1b08-107">To deploy your web project using Visual Studio 2017</span></span>

<span data-ttu-id="a1b08-108">次に示す、Visual Studio 2017 を使用して Web プロジェクトを展開する手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="a1b08-108">Complete the following steps to deploy your web project using Visual Studio 2017.</span></span>

1. <span data-ttu-id="a1b08-109">**[ソリューション エクスプローラー]** で、アドイン プロジェクトのショートカット メニューを開き、**[発行]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a1b08-109">In  **Solution Explorer**, open the shortcut menu for the add-in project, and then choose  **Publish**.</span></span>
    
    <span data-ttu-id="a1b08-110">[**アドインの発行**] ページが表示されます。</span><span class="sxs-lookup"><span data-stu-id="a1b08-110">The  **Publish your add-in** page appears.</span></span>
    
2. <span data-ttu-id="a1b08-111">**[現在のプロファイル]** ドロップダウン リストで、プロファイルを選択するか、**[新規…]** を選択して新しいプロファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="a1b08-111">In the  **Current profile** drop-down list, select a profile or choose **New ...** to create a new profile.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="a1b08-112">発行プロファイルでは、展開先のサーバー、サーバーへのログオンに必要な資格情報、展開するデータベース、およびその他の展開オプションを指定します。</span><span class="sxs-lookup"><span data-stu-id="a1b08-112">A publish profile specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.</span></span>

    <span data-ttu-id="a1b08-113">**[新規...]** を選択すると、ウィザードが表示され、その **[発行プロファイルの作成]** ページが表示されます。</span><span class="sxs-lookup"><span data-stu-id="a1b08-113">If you choose  **New ...**, a wizard appears with the **Create publishing profile** page.</span></span> <span data-ttu-id="a1b08-114">このウィザードを使用して、Microsoft Azure などの Web サイトをホストするプロバイダーから発行プロファイルをインポートするか、新しいプロファイルを作成するかして、次の手順でサーバー、資格情報、その他の設定を追加することができます。</span><span class="sxs-lookup"><span data-stu-id="a1b08-114">You can use this wizard to import a publishing profile from a web site hosting provider such as Microsoft Azure or create a new profile and add your server, credentials, and other settings in the next procedure.</span></span>
    
    <span data-ttu-id="a1b08-115">発行プロファイルのインポートまたは新しい発行プロファイルの作成の詳細については、「[発行プロファイルの作成](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a1b08-115">For more information about importing publishing profiles or creating new publishing profiles, see [Creating a Publish Profile](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).</span></span>
    
3. <span data-ttu-id="a1b08-116">**[アドインを発行する]** ページで、**[Web プロジェクトの配置]** リンクを選択します。</span><span class="sxs-lookup"><span data-stu-id="a1b08-116">On the **Publish your add-in** page, choose the **Deploy your web project** link.</span></span>
    
    <span data-ttu-id="a1b08-117">**[発行]** ダイアログ ボックスが表示されます。</span><span class="sxs-lookup"><span data-stu-id="a1b08-117">The  **Publish** dialog box appears.</span></span> <span data-ttu-id="a1b08-118">このウィザードの使用法の詳細については、「[手順: Visual Studio でワンクリック発行を使用して Web プロジェクトを展開する](https://msdn.microsoft.com/library/dd465337.aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a1b08-118">For more information about using this wizard, see [How to: Deploy a Web Project using On-Click Publishing in Visual Studio](https://msdn.microsoft.com/library/dd465337.aspx).</span></span>
    

## <a name="to-package-your-add-in-using-visual-studio-2017"></a><span data-ttu-id="a1b08-119">Visual Studio 2017 を使用してアドインをパッケージ化するには</span><span class="sxs-lookup"><span data-stu-id="a1b08-119">To package your add-in using Visual Studio 2017</span></span>

<span data-ttu-id="a1b08-120">次に示す、Visual Studio 2017 を使用してアドインをパッケージ化する手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="a1b08-120">Complete the following steps to package your add-in using Visual Studio 2017.</span></span>

1. <span data-ttu-id="a1b08-121">**[アドインの発行]** ページで、**[アドインのパッケージ]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a1b08-121">In the **Publish your add-in** page, choose the **Package the add-in** button.</span></span>
    
    <span data-ttu-id="a1b08-122">ウィザードが表示され、その **[アドインのパッケージ]** ページが表示されます。</span><span class="sxs-lookup"><span data-stu-id="a1b08-122">A wizard appears with the **Package the add-in** page.</span></span>
    
2. <span data-ttu-id="a1b08-123">**[Web サイトがホストされている場所]** ドロップダウン リストで、アドインのコンテンツ ファイルをホストする Web サイトの URL を選択するか入力して、**[完了]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a1b08-123">In the **Where is your website hosted?** box, enter the URL of the website that will host the content files of your add-in, and then choose **Finish**.</span></span>
    
    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="a1b08-124">Azure の Web サイトは自動的に HTTPS エンドポイントを提供します。</span><span class="sxs-lookup"><span data-stu-id="a1b08-124">Azure websites automatically provide an HTTPS endpoint.</span></span>

    <span data-ttu-id="a1b08-125">Visual Studio は、アドインの発行に必要なファイルを生成して、発行の出力フォルダーを開きます。</span><span class="sxs-lookup"><span data-stu-id="a1b08-125">Visual Studio generates the files that you need to publish your add-in and then opens the publish output folder.</span></span>
    
<span data-ttu-id="a1b08-126">AppSource にアドインを提出する予定がある場合は、**[検証チェックを実行する]** をクリックして、アドインが受け入れられなくなる問題点を特定します。</span><span class="sxs-lookup"><span data-stu-id="a1b08-126">If you plan to submit your add-in to AppSource, you can choose the **Perform a validation check** link to identify any issues that will prevent your add-in from being accepted.</span></span> <span data-ttu-id="a1b08-127">アドインをストアに提出する前に、すべての問題を解決してください。</span><span class="sxs-lookup"><span data-stu-id="a1b08-127">You should address all issues before you submit your add-in to the store.</span></span>

<span data-ttu-id="a1b08-p105">XML マニフェストを適切な場所にアップロードして[アドインを発行](../publish/publish.md)できるようになりました。XML マニフェストは、`app.publish` フォルダーの `OfficeAppManifests` にあります。たとえば、次のようになります。</span><span class="sxs-lookup"><span data-stu-id="a1b08-p105">You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:</span></span>

 `%UserProfile%\Documents\Visual Studio 2017\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a><span data-ttu-id="a1b08-131">関連項目</span><span class="sxs-lookup"><span data-stu-id="a1b08-131">See also</span></span>

- [<span data-ttu-id="a1b08-132">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="a1b08-132">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="a1b08-133">AppSource と Office 内でソリューションを使用できるようにする</span><span class="sxs-lookup"><span data-stu-id="a1b08-133">Make your solutions available in AppSource and within Office</span></span>](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)
    

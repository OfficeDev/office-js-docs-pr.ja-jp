---
title: Visual Studio を使用してアドインを発行する
description: Visual Studio 2019 を使用して Web プロジェクトを展開し、アドインをパッケージ化する方法。
ms.date: 12/02/2019
localization_priority: Priority
ms.openlocfilehash: 5da7fc643eb517f777325658d01889f3e51906bd
ms.sourcegitcommit: 44f1a4a3e1ae3c33d7d5fabcee14b84af94e03da
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/03/2019
ms.locfileid: "39670196"
---
# <a name="publish-your-add-in-using-visual-studio"></a><span data-ttu-id="d9b7d-103">Visual Studio を使用してアドインを発行する</span><span class="sxs-lookup"><span data-stu-id="d9b7d-103">Package your add-in using Visual Studio</span></span>

<span data-ttu-id="d9b7d-104">Office アドイン パッケージには、アドインの発行に使用する XML [マニフェスト ファイル](../develop/add-in-manifests.md)が含まれています。</span><span class="sxs-lookup"><span data-stu-id="d9b7d-104">Your Office Add-in package contains an XML [manifest file](../develop/add-in-manifests.md) that you'll use to publish the add-in.</span></span> <span data-ttu-id="d9b7d-105">プロジェクトの Web アプリケーション ファイルは個別に発行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d9b7d-105">You'll have to publish the web application files of your project separately.</span></span> <span data-ttu-id="d9b7d-106">この記事では、Visual Studio 2019 を使用して Web プロジェクトを展開し、アドインをパッケージ化する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="d9b7d-106">This article describes how to deploy your web project and package your add-in by using Visual Studio 2019.</span></span>

> [!NOTE]
> <span data-ttu-id="d9b7d-107">Yeoman ジェネレーターを使用して作成し、Visual Studio Code またはその他のエディターで開発した Office アドインの発行については、「[Visual Studio Code で開発したアドインの発行](publish-add-in-vs-code.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d9b7d-107">For information about publishing an Office Add-in that you created using the Yeoman generator and developed with Visual Studio Code or any other editor, see [Publish an add-in developed with Visual Studio Code](publish-add-in-vs-code.md).</span></span>

## <a name="to-deploy-your-web-project-using-visual-studio-2019"></a><span data-ttu-id="d9b7d-108">Visual Studio 2019 を使用して Web プロジェクトを展開するには</span><span class="sxs-lookup"><span data-stu-id="d9b7d-108">To deploy your web project using Visual Studio 2019</span></span>

<span data-ttu-id="d9b7d-109">Visual Studio 2019 を使用して Web プロジェクトを展開するには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="d9b7d-109">Complete the following steps to deploy your web project using Visual Studio 2019.</span></span>

1. <span data-ttu-id="d9b7d-110">[**ビルド**] タブから、[**公開 [アドインの名前]**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="d9b7d-110">From the **Build** tab, choose **Publish [Name of your add-in]**.</span></span>

2. <span data-ttu-id="d9b7d-111">[**発行先の選択**] ウィンドウで、優先されるターゲットに公開するオプションのいずれかを選択します。</span><span class="sxs-lookup"><span data-stu-id="d9b7d-111">In the **Pick a publish target** window, choose one of the options to publish to your preferred target.</span></span> <span data-ttu-id="d9b7d-112">各発行ターゲットでは、Azure Virtual Machine やフォルダーの場所など、開始するための詳細な情報を含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="d9b7d-112">Each publish target requires you to include more information to get started, such as an Azure Virtual Machine or folder location.</span></span> <span data-ttu-id="d9b7d-113">公開場所を指定し、必要な情報をすべて入力したら、[**公開**] を選択します</span><span class="sxs-lookup"><span data-stu-id="d9b7d-113">Once you have specified a publish location and filled in all of the information required, select **Publish**</span></span>

    > [!NOTE]
    > <span data-ttu-id="d9b7d-114">公開ターゲットを選択すると、展開先のサーバー、サーバーへのログオンに必要な資格情報、展開するデータベース、およびその他の展開オプションが指定されます。</span><span class="sxs-lookup"><span data-stu-id="d9b7d-114">Picking a publish target specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.</span></span>

3. <span data-ttu-id="d9b7d-115">各発行ターゲット オプションの展開手順の詳細については、「[First look at deployment in Visual Studio (Visual Studioでの展開の最初の画面)](/visualstudio/deployment/deploying-applications-services-and-components?view=vs-2019)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d9b7d-115">For more information about deployment steps for each publish target option, see [First look at deployment in Visual Studio](/visualstudio/deployment/deploying-applications-services-and-components?view=vs-2019).</span></span>

## <a name="to-package-and-publish-your-add-in-using-iis-ftp-or-web-deploy-using-visual-studio-2019"></a><span data-ttu-id="d9b7d-116">IIS、FTP、または Visual Studio 2019 を使用したWeb 配置を使用してアドインをパッケージ化して公開するには</span><span class="sxs-lookup"><span data-stu-id="d9b7d-116">To package and publish your add-in using IIS, FTP, or Web Deploy using Visual Studio 2019</span></span>

<span data-ttu-id="d9b7d-117">Visual Studio 2019 を使用してアドインをパッケージ化するには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="d9b7d-117">Complete the following steps to package your add-in using Visual Studio 2019.</span></span>

1. <span data-ttu-id="d9b7d-118">[**ビルド**] タブから、[**公開 [アドインの名前]**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="d9b7d-118">From the **Build** tab, choose **Publish [Name of your add-in]**.</span></span>
2. <span data-ttu-id="d9b7d-119">[**発行先の選択**]ウィンドウで **IIS、FTPなど**を選択し、[**構成**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="d9b7d-119">In the **Pick a publish target** window, choose **IIS, FTP, etc**, and select **Configure**.</span></span> <span data-ttu-id="d9b7d-120">次に、[**発行**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="d9b7d-120">Next, select **Publish**.</span></span>
3. <span data-ttu-id="d9b7d-121">プロセスをガイドするウィザードが表示されます。</span><span class="sxs-lookup"><span data-stu-id="d9b7d-121">A wizard appears that will help guide you through the process.</span></span> <span data-ttu-id="d9b7d-122">公開方法が Web 配置などの優先される方法であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="d9b7d-122">Ensure the publish method is your preferred method, such as Web Deploy.</span></span>
4. <span data-ttu-id="d9b7d-123">[**接続先 URL**] ボックスに、アドインのコンテンツ ファイルをホストする Web サイトの URL を入力し、[**次へ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="d9b7d-123">In the **Destination URL** box, enter the URL of the website that will host the content files of your add-in, and then select **Next**.</span></span> <span data-ttu-id="d9b7d-124">アドインを AppSource に提出する場合には、[**接続の検証**] ボタンを選択し、アドインの受け入れを妨げている問題を特定できます。</span><span class="sxs-lookup"><span data-stu-id="d9b7d-124">If you plan to submit your add-in to AppSource, you can choose the **Validate Connection** button to identify any issues that will prevent your add-in from being accepted.</span></span> <span data-ttu-id="d9b7d-125">アドインをストアに提出する前に、すべての問題に対処する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d9b7d-125">You should address all issues before you submit your add-in to the store.</span></span>
5. <span data-ttu-id="d9b7d-126">**ファイル発行オプション**を含む必要な設定を確認し、[**保存**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="d9b7d-126">Confirm any settings desired including **File Publish Options** and select **Save**.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="d9b7d-127">Azure の Web サイトは自動的に HTTPS エンドポイントを提供します。</span><span class="sxs-lookup"><span data-stu-id="d9b7d-127">Azure websites automatically provide an HTTPS endpoint.</span></span>

<span data-ttu-id="d9b7d-p106">XML マニフェストを適切な場所にアップロードして[アドインを発行](../publish/publish.md)できるようになりました。XML マニフェストは、`app.publish` フォルダーの `OfficeAppManifests` にあります。たとえば、次のようになります。</span><span class="sxs-lookup"><span data-stu-id="d9b7d-p106">You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:</span></span>

 `%UserProfile%\Documents\Visual Studio 2019\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`

## <a name="see-also"></a><span data-ttu-id="d9b7d-131">関連項目</span><span class="sxs-lookup"><span data-stu-id="d9b7d-131">See also</span></span>

- [<span data-ttu-id="d9b7d-132">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="d9b7d-132">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="d9b7d-133">AppSource と Office 内でソリューションを使用できるようにする</span><span class="sxs-lookup"><span data-stu-id="d9b7d-133">Make your solutions available in AppSource and within Office</span></span>](/office/dev/store/submit-to-the-office-store)

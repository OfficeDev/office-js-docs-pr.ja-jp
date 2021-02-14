---
title: 作業ウィンドウ アドインとコンテンツ アドインを SharePoint アプリ カタログに発行する
description: 組織内のユーザーが Office アドインにアクセスできるようにするために、管理者は組織のアプリ カタログに Office アドインのマニフェスト ファイルをアップロードできます。
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 737448a498741ec0327939dc9e562fc04d78a8e5
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234178"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-app-catalog"></a><span data-ttu-id="3d665-103">作業ウィンドウ アドインとコンテンツ アドインを SharePoint アプリ カタログに発行する</span><span class="sxs-lookup"><span data-stu-id="3d665-103">Publish task pane and content add-ins to a SharePoint app catalog</span></span>

<span data-ttu-id="3d665-p101">アプリ カタログは、Office アドインと SharePoint アドインのドキュメント ライブラリをホストする SharePoint Web アプリケーションまたは SharePoint Online テナンシーの専用サイト コレクションです。組織内のユーザーが Office アドインにアクセスできるようにするために、管理者は組織のアプリ カタログに Office アドインのマニフェスト ファイルをアップロードできます。管理者がアプリ カタログを信頼できるカタログとして登録すると、ユーザーは Office クライアント アプリケーションで挿入 UI からアドインを挿入できます。</span><span class="sxs-lookup"><span data-stu-id="3d665-p101">An app catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for Office and SharePoint Add-ins. To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the app catalog for their organization. When an administrator registers an app catalog as a trusted catalog, users can insert the add-in from the insertion UI in an Office client application.</span></span>

> [!IMPORTANT]
> - <span data-ttu-id="3d665-106">SharePoint のアプリ カタログでは、アドイン コマンドなど、[アドイン マニフェスト](../develop/add-in-manifests.md)の `VersionOverrides` ノードで実装されるアドイン機能がサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="3d665-106">App catalogs on SharePoint do not support add-in features that are implemented in the `VersionOverrides` node of the [add-in manifest](../develop/add-in-manifests.md), such as add-in commands.</span></span>
> - <span data-ttu-id="3d665-107">クラウド環境またはハイブリッド環境をターゲットにしている場合は [、Microsoft 365](../publish/centralized-deployment.md) 管理センターから一元展開を使用してアドインを発行することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="3d665-107">If you’re targeting a cloud or hybrid environment, we recommend that you [use Centralized Deployment via the Microsoft 365 admin center](../publish/centralized-deployment.md) to publish your add-ins.</span></span>
> - <span data-ttu-id="3d665-108">SharePoint のアプリ カタログは Office on Mac ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="3d665-108">App catalogs on SharePoint are not supported in Office on Mac.</span></span> <span data-ttu-id="3d665-109">Office アドインを Mac クライアントに展開するには、そのアドインを [AppSource](/office/dev/store/submit-to-the-office-store) に提出する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3d665-109">To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](/office/dev/store/submit-to-the-office-store).</span></span>

## <a name="create-an-app-catalog"></a><span data-ttu-id="3d665-110">アプリ カタログを作成する</span><span class="sxs-lookup"><span data-stu-id="3d665-110">Create an app catalog</span></span>

<span data-ttu-id="3d665-111">次のいずれかのセクションの手順を実行して、オンプレミスの SharePoint Server または Microsoft 365 でアプリ カタログを作成します。</span><span class="sxs-lookup"><span data-stu-id="3d665-111">Complete the steps in one of the following sections to create an app catalog with on-premises SharePoint Server or on Microsoft 365.</span></span>

### <a name="to-create-an-app-catalog-for-on-premises-sharepoint-server"></a><span data-ttu-id="3d665-112">オンプレミス SharePoint サーバーでアプリ カタログを作成する</span><span class="sxs-lookup"><span data-stu-id="3d665-112">To create an app catalog for on-premises SharePoint Server</span></span>

<span data-ttu-id="3d665-113">SharePoint アプリ カタログを作成するには、[web アプリケーションのアプリ カタログサイトを作成する](/sharepoint/administration/manage-the-app-catalog)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3d665-113">To create the SharePoint app catalog, follow the instructions at [Configure the App Catalog site for a web application](/sharepoint/administration/manage-the-app-catalog).</span></span>

<span data-ttu-id="3d665-114">アプリ カタログを作成したら [Office アドインを発行する](#publish-an-office-add-in) 手順に従います。</span><span class="sxs-lookup"><span data-stu-id="3d665-114">Once you have created the app catalog follow the steps to [publish an Office Add-in](#publish-an-office-add-in).</span></span>

### <a name="to-create-an-app-catalog-on-microsoft-365"></a><span data-ttu-id="3d665-115">Microsoft 365 でアプリ カタログを作成するには</span><span class="sxs-lookup"><span data-stu-id="3d665-115">To create an app catalog on Microsoft 365</span></span>

<span data-ttu-id="3d665-116">SharePoint アプリ カタログを作成するには、「アプリ カタログ サイト コレクションを作成 [する」の手順に従います](/sharepoint/use-app-catalog#step-1-create-the-app-catalog-site-collection)。</span><span class="sxs-lookup"><span data-stu-id="3d665-116">To create the SharePoint app catalog, follow the instructions at [Create the App Catalog site collection](/sharepoint/use-app-catalog#step-1-create-the-app-catalog-site-collection).</span></span> <span data-ttu-id="3d665-117">アプリ カタログを作成したら、次のセクションの手順に従って、新しいアドインOffice発行します。</span><span class="sxs-lookup"><span data-stu-id="3d665-117">Once you have created the app catalog, follow the steps in the next section to publish an Office Add-in.</span></span>

## <a name="publish-an-office-add-in"></a><span data-ttu-id="3d665-118">Office アドインの発行</span><span class="sxs-lookup"><span data-stu-id="3d665-118">Publish an Office Add-in</span></span>

<span data-ttu-id="3d665-119">Microsoft 365 またはオンプレミスの SharePoint Server のアプリ カタログに Office アドインを発行するには、次のいずれかのセクションの手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="3d665-119">Complete the steps in one of the following sections to publish an Office Add-in to an app catalog on Microsoft 365 or on-premises SharePoint Server.</span></span>

### <a name="to-publish-an-office-add-in-to-a-sharepoint-app-catalog-on-microsoft-365"></a><span data-ttu-id="3d665-120">Microsoft 365 Office SharePoint アプリ カタログにアドインを発行するには</span><span class="sxs-lookup"><span data-stu-id="3d665-120">To publish an Office Add-in to a SharePoint app catalog on Microsoft 365</span></span>

1. <span data-ttu-id="3d665-121">[新しい SharePoint 管理センターの [アクティブなサイト] ページ](https://admin.microsoft.com/sharepoint?page=siteManagement&modern=true)に移動し、組織の[管理者権限](/sharepoint/sharepoint-admin-role)が付与されているアカウントでサインインします。</span><span class="sxs-lookup"><span data-stu-id="3d665-121">Go to the [Active sites page of the new SharePoint admin center](https://admin.microsoft.com/sharepoint?page=siteManagement&modern=true) and sign in with an account that has [admin permissions](/sharepoint/sharepoint-admin-role) for your organization.</span></span>

    > [!NOTE]
    > <span data-ttu-id="3d665-122">Microsoft 365 Germany を使用している場合は [、Microsoft 365](https://go.microsoft.com/fwlink/p/?linkid=848041)管理センターにサインインし、SharePoint 管理センターに移動して、[その他の機能] ページを開きます。</span><span class="sxs-lookup"><span data-stu-id="3d665-122">If you have Microsoft 365 Germany, [sign in to the Microsoft 365 admin center](https://go.microsoft.com/fwlink/p/?linkid=848041), then browse to the SharePoint admin center and open the More features page.</span></span> <br><span data-ttu-id="3d665-123">21Vianet (中国) が運営する Microsoft 365 を使用している場合は [、Microsoft 365](https://go.microsoft.com/fwlink/p/?linkid=850627)管理センターにサインインし、SharePoint 管理センターに移動して、[その他の機能] ページを開きます。</span><span class="sxs-lookup"><span data-stu-id="3d665-123">If you have Microsoft 365 operated by 21Vianet (China), [sign in to the Microsoft 365 admin center](https://go.microsoft.com/fwlink/p/?linkid=850627), then browse to the SharePoint admin center and open the More features page.</span></span>

1. <span data-ttu-id="3d665-124">URL 列で URL を選択して、アプリ カタログ サイトを開きます。</span><span class="sxs-lookup"><span data-stu-id="3d665-124">Open the app catalog site by selecting its URL in the URL column.</span></span>

    > [!NOTE]
    > <span data-ttu-id="3d665-125">前のセクションでアプリ カタログ サイトを作成したばかりである場合、サイトのセットアップが完了するには数分かかる場合があります。</span><span class="sxs-lookup"><span data-stu-id="3d665-125">If you just created the app catalog site in the previous section, it can take a few minutes for the site to finish setting up.</span></span>

1. <span data-ttu-id="3d665-126">[**Office 用アプリを配信する**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="3d665-126">Choose **Distribute apps for Office**.</span></span>
1. <span data-ttu-id="3d665-127">[**Office 用アプリ**] ページで、[**新規**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="3d665-127">In the **Apps for Office** page, choose **New**.</span></span>
1. <span data-ttu-id="3d665-128">[**ドキュメントの追加**] ダイアログで、[**ファイルの選択**] ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="3d665-128">In the **Add a document** dialog, select the **Choose Files** button.</span></span>
1. <span data-ttu-id="3d665-129">アップロードする [マニフェスト](../develop/add-in-manifests.md) ファイルを見つけて指定し、[**開く**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="3d665-129">Locate and specify the [manifest](../develop/add-in-manifests.md) file to upload and choose **Open**.</span></span>
1. <span data-ttu-id="3d665-130">[**ドキュメントの追加**] ダイアログで、[**OK**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="3d665-130">In the **Add a document** dialog, choose **OK**.</span></span>

### <a name="to-publish-an-add-in-to-an-app-catalog-with-on-premises-sharepoint-server"></a><span data-ttu-id="3d665-131">オンプレミスの SharePoint サーバーでアプリ カタログにアドインを発行する</span><span class="sxs-lookup"><span data-stu-id="3d665-131">To publish an add-in to an app catalog with on-premises SharePoint Server</span></span>

1. <span data-ttu-id="3d665-132">**[サーバーの全体管理]** ページを開きます。</span><span class="sxs-lookup"><span data-stu-id="3d665-132">Open the **Central Administration** page.</span></span>
1. <span data-ttu-id="3d665-133">左側の作業ウィンドウで、[**アプリ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="3d665-133">In the left task pane, choose **Apps**.</span></span>
1. <span data-ttu-id="3d665-134">**[アプリ]** ページの **[アプリの管理]** で **[アプリ カタログの管理]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="3d665-134">On the **Apps** page, under **App Management**, choose **Manage App Catalog**.</span></span>
1. <span data-ttu-id="3d665-135">**[アプリ カタログの管理]** ページの **Web アプリケーション** セレクターで正しい Web アプリケーションが選択されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="3d665-135">On the **Manage App Catalog** page, make sure you have the right web application selected in the **Web Application** Selector.</span></span>
1. <span data-ttu-id="3d665-136">**サイト URL** の下にある URL を選び、アプリ カタログサイトを開きます。</span><span class="sxs-lookup"><span data-stu-id="3d665-136">Choose the URL under the **Site URL** to open the app catalog site.</span></span>
1. <span data-ttu-id="3d665-137">[**Office 用アプリを配信する**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="3d665-137">Choose **Distribute apps for Office**.</span></span>
1. <span data-ttu-id="3d665-138">[**Office 用アプリ**] ページで、[**新規**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="3d665-138">In the **Apps for Office** page, choose **New**.</span></span>
1. <span data-ttu-id="3d665-139">[**ドキュメントの追加**] ダイアログで、[**ファイルの選択**] ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="3d665-139">In the **Add a document** dialog, select the **Choose Files** button.</span></span>
1. <span data-ttu-id="3d665-140">アップロードする [マニフェスト](../develop/add-in-manifests.md) ファイルを見つけて指定し、[**開く**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="3d665-140">Locate and specify the [manifest](../develop/add-in-manifests.md) file to upload and choose **Open**.</span></span>
1. <span data-ttu-id="3d665-141">[**ドキュメントの追加**] ダイアログで、[**OK**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="3d665-141">In the **Add a document** dialog, choose **OK**.</span></span>

## <a name="insert-office-add-ins-from-the-app-catalog"></a><span data-ttu-id="3d665-142">アプリ カタログから Office アドインを挿入する</span><span class="sxs-lookup"><span data-stu-id="3d665-142">Insert Office Add-ins from the app catalog</span></span>

<span data-ttu-id="3d665-143">オンライン Office アプリケーションの場合は、次の手順を実行してアプリ カタログから Office アドインを見つけることができます。</span><span class="sxs-lookup"><span data-stu-id="3d665-143">For online Office applications, you can find Office Add-ins from the app catalog by completing the following steps.</span></span>

1. <span data-ttu-id="3d665-144">オンライン Office アプリケーション (Excel、PowerPoint、または Word) を開きます。</span><span class="sxs-lookup"><span data-stu-id="3d665-144">Open the online Office application (Excel, PowerPoint, or Word).</span></span>
1. <span data-ttu-id="3d665-145">文書を作成または開く。</span><span class="sxs-lookup"><span data-stu-id="3d665-145">Create or open a document.</span></span>
1. <span data-ttu-id="3d665-146">**[挿入]** > **[アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="3d665-146">Choose **Insert** > **Add-ins**.</span></span>
1. <span data-ttu-id="3d665-147">[Office アドイン] ダイアログの **[自分の所属組織]** タブを選択します。Office アドインのリストが表示されます。</span><span class="sxs-lookup"><span data-stu-id="3d665-147">In the Office Add-ins dialog, choose the **MY ORGANIZATION** tab.  The Office Add-ins are listed.</span></span>
1. <span data-ttu-id="3d665-148">Office アドインを選択し、 **追加** を選択します。</span><span class="sxs-lookup"><span data-stu-id="3d665-148">Choose an Office Add-in and then choose **Add**.</span></span>

<span data-ttu-id="3d665-149">デスクトップの Office アプリケーションの場合は、次の手順を実行してアプリ カタログから Office アドインを見つけることができます。</span><span class="sxs-lookup"><span data-stu-id="3d665-149">For Office applications on the desktop, you can find Office Add-ins from the app catalog by completing the following steps.</span></span>

1. <span data-ttu-id="3d665-150">デスクトップ Office アプリケーション (Excel、Word、または PowerPoint) を開きます。</span><span class="sxs-lookup"><span data-stu-id="3d665-150">Open the desktop Office application (Excel, Word, or PowerPoint)</span></span>
1. <span data-ttu-id="3d665-151">**[ファイル]**  >  **[オプション]**  >  **[セキュリティ センター]**  >  **[セキュリティ センターの設定]**  >  **[信頼できるアドイン カタログ]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="3d665-151">Choose **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.</span></span>
1. <span data-ttu-id="3d665-152">[**カタログ URL** ] ボックスに SharePoint アプリ カタログの URL を入力し、**[カタログの追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="3d665-152">Enter the URL of the SharePoint app catalog in the **Catalog Url** box and choose **Add catalog**.</span></span>
    <span data-ttu-id="3d665-153">短い URL を使用します。</span><span class="sxs-lookup"><span data-stu-id="3d665-153">Use the shorter form of the URL.</span></span> <span data-ttu-id="3d665-154">たとえば、SharePoint アプリ カタログの URL が次のような場合:</span><span class="sxs-lookup"><span data-stu-id="3d665-154">For example, if the URL of the SharePoint app catalog is:</span></span>
    - `https://<domain>/sites/<AddinCatalogSiteCollection>/AgaveCatalog`

    <span data-ttu-id="3d665-155">親サイト コレクションの URL のみを指定します。</span><span class="sxs-lookup"><span data-stu-id="3d665-155">Specify just the URL of the parent site collection:</span></span>
    - `https://<domain>/sites/<AddinCatalogSiteCollection>`
1. <span data-ttu-id="3d665-156">Office アプリケーションを閉じてから、もう一度開きます。</span><span class="sxs-lookup"><span data-stu-id="3d665-156">Close and reopen the Office application.</span></span>
1. <span data-ttu-id="3d665-157">**[挿入]** > **[アドインの取得]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="3d665-157">Choose **Insert** > **Get Add-ins**.</span></span>
1. <span data-ttu-id="3d665-158">[Office アドイン] ダイアログの **[自分の所属組織]** タブを選択します。Office アドインのリストが表示されます。</span><span class="sxs-lookup"><span data-stu-id="3d665-158">In the Office Add-ins dialog, choose the **MY ORGANIZATION** tab.  The Office Add-ins are listed.</span></span>
1. <span data-ttu-id="3d665-159">Office アドインを選択し、 **追加** を選択します。</span><span class="sxs-lookup"><span data-stu-id="3d665-159">Choose an Office Add-in and then choose **Add**.</span></span>

<span data-ttu-id="3d665-160">または、管理者はグループ ポリシーを使用して SharePoint のアプリ カタログを指定できます。</span><span class="sxs-lookup"><span data-stu-id="3d665-160">Alternatively, an administrator can specify an app catalog on SharePoint by using Group Policy.</span></span> <span data-ttu-id="3d665-161">関連するポリシー設定は [、Microsoft 365 Apps、Office 2019、および Office 2016](https://www.microsoft.com/download/details.aspx?id=49030) の管理用テンプレート ファイル (ADMX/ADML) で使用できます。ユーザーの構成\ポリシー\管理用テンプレート \**\Microsoft Office 2016\セキュリティ設定\セキュリティ センター\** 信頼済みカタログにあります。</span><span class="sxs-lookup"><span data-stu-id="3d665-161">The relevant policy settings are available in the [Administrative Template files (ADMX/ADML) for Microsoft 365 Apps, Office 2019, and Office 2016](https://www.microsoft.com/download/details.aspx?id=49030) and be found under **User Configuration\Policies\Administrative Templates\Microsoft Office 2016\Security Settings\Trust Center\Trusted Catalogs**.</span></span>

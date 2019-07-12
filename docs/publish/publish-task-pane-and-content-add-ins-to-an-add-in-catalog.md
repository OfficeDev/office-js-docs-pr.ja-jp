---
title: 作業ウィンドウ アドインとコンテンツ アドインを SharePoint アプリ カタログに発行する
description: 組織内のユーザーが Office アドインにアクセスできるようにするために、管理者は組織のアプリ カタログに Office アドインのマニフェスト ファイルをアップロードできます。
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 106dfd2b1610be92f1b53dc1644ff3f8c60c0543
ms.sourcegitcommit: 9c5a836d4464e49846c9795bf44cfe23e9fc8fbe
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2019
ms.locfileid: "35617031"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-app-catalog"></a><span data-ttu-id="94d0a-103">作業ウィンドウ アドインとコンテンツ アドインを SharePoint アプリ カタログに発行する</span><span class="sxs-lookup"><span data-stu-id="94d0a-103">Publish task pane and content add-ins to a SharePoint catalog</span></span>

<span data-ttu-id="94d0a-p101">アプリ カタログは、Office アドインと SharePoint アドインのドキュメント ライブラリをホストする SharePoint Web アプリケーションまたは SharePoint Online テナンシーの専用サイト コレクションです。組織内のユーザーが Office アドインにアクセスできるようにするために、管理者は組織のアプリ カタログに Office アドインのマニフェスト ファイルをアップロードできます。管理者がアプリ カタログを信頼できるカタログとして登録すると、ユーザーは Office クライアント アプリケーションで挿入 UI からアドインを挿入できます。</span><span class="sxs-lookup"><span data-stu-id="94d0a-p101">An add-in catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for Office and SharePoint Add-ins. To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the add-in catalog for their organization. When an administrator registers an add-in catalog as a trusted catalog, users can insert the add-in from the insertion UI in an Office client application.</span></span>

> [!IMPORTANT]
> - <span data-ttu-id="94d0a-106">SharePoint のアプリ カタログでは、アドイン コマンドなど、[アドイン マニフェスト](../develop/add-in-manifests.md)の `VersionOverrides` ノードで実装されるアドイン機能がサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="94d0a-106">Add-in catalogs on SharePoint do not support add-in features that are implemented in the `VersionOverrides` node of the [add-in manifest](../develop/add-in-manifests.md), such as add-in commands.</span></span>
> - <span data-ttu-id="94d0a-107">クラウド環境またはハイブリッド環境をターゲットにしている場合は、アドインの発行に [Office 365 管理センターからの一元展開を使用する](../publish/centralized-deployment.md)ことをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="94d0a-107">If you’re targeting a cloud or hybrid environment, we recommend that you [use Centralized Deployment via the Office 365 admin center](../publish/centralized-deployment.md) to publish your add-ins.</span></span>
> - <span data-ttu-id="94d0a-108">SharePoint のアプリ カタログは Office on Mac ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="94d0a-108">App catalogs on SharePoint are not supported in Office on Mac.</span></span> <span data-ttu-id="94d0a-109">Office アドインを Mac クライアントに展開するには、そのアドインを [AppSource](/office/dev/store/submit-to-the-office-store) に提出する必要があります。</span><span class="sxs-lookup"><span data-stu-id="94d0a-109">To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](/office/dev/store/submit-to-the-office-store).</span></span>

## <a name="create-an-app-catalog"></a><span data-ttu-id="94d0a-110">アプリ カタログを作成する</span><span class="sxs-lookup"><span data-stu-id="94d0a-110">Create app catalog site</span></span>

<span data-ttu-id="94d0a-111">次のいずれかのセクションの手順を完了し、オンプレミス SharePoint サーバーまたは Office 365 を使用して、アプリ カタログを作成します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-111">Complete the steps in one of the following sections to set up an add-in catalog on SharePoint or on Office 365.</span></span>

### <a name="to-create-an-app-catalog-for-on-premises-sharepoint-server"></a><span data-ttu-id="94d0a-112">オンプレミス SharePoint サーバーでアプリ カタログを作成する</span><span class="sxs-lookup"><span data-stu-id="94d0a-112">To create an app catalog for on-premises SharePoint Server</span></span>

<span data-ttu-id="94d0a-113">SharePoint アプリ カタログを作成するには、[web アプリケーションのアプリ カタログサイトを作成する](/sharepoint/administration/manage-the-app-catalog)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="94d0a-113">To create the SharePoint app catalog, follow the instructions at [Configure the App Catalog site for a web application](/sharepoint/administration/manage-the-app-catalog).</span></span>

<span data-ttu-id="94d0a-114">アプリ カタログを作成したら [Office アドインを発行する](#publish-an-office-add-in) 手順に従います。</span><span class="sxs-lookup"><span data-stu-id="94d0a-114">Once you have created the app catalog follow the steps to [publish an Office Add-in](#publish-an-office-add-in).</span></span>

### <a name="to-create-an-app-catalog-on-office-365"></a><span data-ttu-id="94d0a-115">Office 365 でアプリ カタログを作成する</span><span class="sxs-lookup"><span data-stu-id="94d0a-115">To create an app catalog on Office 365</span></span>

1. <span data-ttu-id="94d0a-116">Microsoft 365 管理センターに移動します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-116">Go to the Microsoft 365 admin center.</span></span> <span data-ttu-id="94d0a-117">管理センターの検索方法については、「[Microsoft 365 管理センターについて](/office365/admin/admin-overview/about-the-admin-center)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="94d0a-117">For information on how to find the admin center, see [About the Microsoft 365 admin center](/office365/admin/admin-overview/about-the-admin-center).</span></span>

2. <span data-ttu-id="94d0a-118">Microsoft 365 管理センターのページで [**管理センター**] のリストを展開し、[**SharePoint**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-118">On the Microsoft 365 admin center page, expand the list of **Admin centers**, and then choose **SharePoint**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="94d0a-119">カタログを作成するには、従来の SharePoint 管理センターを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="94d0a-119">You need to use the Classic SharePoint admin center to create the catalog.</span></span> <span data-ttu-id="94d0a-120">新しい SharePoint 管理センターにいる場合は、左側のウィンドウで「**従来の SharePoint 管理センター**」を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-120">If you are in the new SharePoint admin center, choose **Classic SharePoint admin center** in the left pane.</span></span>

3. <span data-ttu-id="94d0a-121">左側の作業ウィンドウで、[**アプリ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-121">In the left task pane, choose  **Apps**.</span></span>

4. <span data-ttu-id="94d0a-122">[**アプリ**] ページで、[**アプリ カタログ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-122">On the **apps** page, select **App Catalog**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="94d0a-123">アプリ カタログが既に作成されてこのページに表示されている場合は、残りの手順をスキップしてこの記事の次のセクションに移動し、カタログにアドインを発行できます。</span><span class="sxs-lookup"><span data-stu-id="94d0a-123">If an app catalog is already created and appears on this page, then you can skip the rest of these steps and go to the next section of this article to publish your add-in to the catalog.</span></span>

5. <span data-ttu-id="94d0a-124">[**アプリ カタログ サイト**] ページで、[**OK**] を選択して既定のオプションを受け入れ、新しいアプリ カタログ サイトを作成します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-124">On the **App Catalog Site** page, select **OK** to accept the default option and create a new app catalog site.</span></span>

6. <span data-ttu-id="94d0a-125">[**アプリ カタログ サイト コレクションを作成する**] ページで、アプリ カタログ サイトのタイトルを指定します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-125">On the  **Create Add-in Catalog Site Collection** page, specify the title of your Add-in Catalog site.</span></span>

7. <span data-ttu-id="94d0a-126">[**Web サイトのアドレス**]を指定します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-126">Specify the web site address.</span></span>

8. <span data-ttu-id="94d0a-127">[**管理者**]を指定します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-127">Specify an **Administrator**.</span></span>

9. <span data-ttu-id="94d0a-128">[**サーバー リソース クォータ**] に 0 (ゼロ) を設定します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-128">Set the Server Resource Quota to 0 (zero), and then select OK.</span></span> <span data-ttu-id="94d0a-129">(サーバー リソース クォータは、パフォーマンスが低いサンドボックス ソリューションのスロットルに関連していますが、このアプリ カタログ サイトにはサンドボックス ソリューションをインストールしません。)</span><span class="sxs-lookup"><span data-stu-id="94d0a-129">(The server resource quota is related to throttling poorly performing sandboxed solutions, but you won't be installing any sandboxed solutions on your add-in catalog site.)</span></span>

10. <span data-ttu-id="94d0a-130">[**OK**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-130">Choose **OK**.</span></span>

## <a name="publish-an-office-add-in"></a><span data-ttu-id="94d0a-131">Office アドインの発行</span><span class="sxs-lookup"><span data-stu-id="94d0a-131">Publish an Office Add-in</span></span>

<span data-ttu-id="94d0a-132">次のいずれかのセクションの手順を完了し、Office アドインを Office 365 またはオンプレミス SharePoint サーバーのアプリ カタログに発行します。 </span><span class="sxs-lookup"><span data-stu-id="94d0a-132">Complete the steps in one of the following sections to publish an Office Add-in to an app catalog on Office 365 or on-premises SharePoint Server.</span></span>

### <a name="to-publish-an-office-add-in-to-a-sharepoint-app-catalog-on-office-365"></a><span data-ttu-id="94d0a-133">Office アドインを Office 365 の SharePoint アプリ カタログに発行する</span><span class="sxs-lookup"><span data-stu-id="94d0a-133">To publish an Office add-in to a SharePoint app catalog on Office 365</span></span>

1. <span data-ttu-id="94d0a-134">Microsoft 365 管理センターに移動します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-134">Go to the Microsoft 365 admin center.</span></span> <span data-ttu-id="94d0a-135">管理センターの検索方法については、「[Microsoft 365 管理センターについて](/office365/admin/admin-overview/about-the-admin-center)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="94d0a-135">For information on how to find the admin center, see [About the Microsoft 365 admin center](/office365/admin/admin-overview/about-the-admin-center).</span></span>
2. <span data-ttu-id="94d0a-136">Microsoft 365 管理センターのページで [**管理センター**] のリストを展開し、[**SharePoint**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-136">On the Microsoft 365 admin center page, expand the list of **Admin centers**, and then choose **SharePoint**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="94d0a-137">カタログを作成するには、従来の SharePoint 管理センターを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="94d0a-137">You need to use the Classic SharePoint admin center to create the catalog.</span></span> <span data-ttu-id="94d0a-138">新しい SharePoint 管理センターにいる場合は、左側のウィンドウで「**従来の SharePoint 管理センター**」を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-138">If you are in the new SharePoint admin center, choose **Classic SharePoint admin center** in the left pane.</span></span>
3. <span data-ttu-id="94d0a-139">左側の作業ウィンドウで、[**アプリ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-139">In the left task pane, choose  **Apps**.</span></span>
4. <span data-ttu-id="94d0a-140">[**アプリ**] ページで、[**アプリ カタログ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-140">On the **apps** page, select **App Catalog**.</span></span>
5. <span data-ttu-id="94d0a-141">[**Office 用アプリを配信する**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-141">Choose **Distribute apps for Office**.</span></span>
6. <span data-ttu-id="94d0a-142">[**Office 用アプリ**] ページで、[**新規**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-142">In the **Apps for Office** page, choose **New**.</span></span>
7. <span data-ttu-id="94d0a-143">[**ドキュメントの追加**] ダイアログで、[**ファイルの選択**] ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="94d0a-143">In the **Add a document** dialog, select the **Choose Files** button.</span></span>
8. <span data-ttu-id="94d0a-144">アップロードする[マニフェスト](../develop/add-in-manifests.md) ファイルを見つけて指定し、[**開く**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-144">Locate and specify the [manifest](../develop/add-in-manifests.md) file to upload and choose **Open**.</span></span>
9. <span data-ttu-id="94d0a-145">[**ドキュメントの追加**] ダイアログで、[**OK**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-145">In the **Add a document** dialog box, choose **OK**.</span></span>

### <a name="to-publish-an-add-in-to-an-app-catalog-with-on-premises-sharepoint-server"></a><span data-ttu-id="94d0a-146">オンプレミスの SharePoint サーバーでアプリ カタログにアドインを発行する</span><span class="sxs-lookup"><span data-stu-id="94d0a-146">To publish an add-in to an app catalog with on-premises SharePoint Server</span></span>

1. <span data-ttu-id="94d0a-147">**[サーバーの全体管理]** ページを開きます。</span><span class="sxs-lookup"><span data-stu-id="94d0a-147">Open the SharePoint Central Administration main page.</span></span>
2. <span data-ttu-id="94d0a-148">左側の作業ウィンドウで、[**アプリ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-148">In the left task pane, choose  **Apps**.</span></span>
3. <span data-ttu-id="94d0a-149">**[アプリ]** ページの **[アプリの管理]** で **[アプリ カタログの管理]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-149">On the  **Apps** page, under **App Management**, choose  **Manage App Catalog**.</span></span>
4. <span data-ttu-id="94d0a-150">**[アプリ カタログの管理]** ページの \*\* Web アプリケーション \*\* セレクターで正しい Web アプリケーションが選択されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-150">On the  **Manage App Catalog** page, make sure you have the right web application selected in the **Web Application Selector**.</span></span>
5. <span data-ttu-id="94d0a-151">**サイト URL**の下にある URL を選び、アプリ カタログサイトを開きます。</span><span class="sxs-lookup"><span data-stu-id="94d0a-151">Choose the URL under the **Site URL** to open the app catalog site.</span></span>
6. <span data-ttu-id="94d0a-152">[**Office 用アプリを配信する**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-152">Choose **Distribute apps for Office**.</span></span>
7. <span data-ttu-id="94d0a-153">[**Office 用アプリ**] ページで、[**新規**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-153">In the **Apps for Office** page, choose **New**.</span></span>
8. <span data-ttu-id="94d0a-154">[**ドキュメントの追加**] ダイアログで、[**ファイルの選択**] ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="94d0a-154">In the **Add a document** dialog, select the **Choose Files** button.</span></span>
9. <span data-ttu-id="94d0a-155">アップロードする[マニフェスト](../develop/add-in-manifests.md) ファイルを見つけて指定し、[**開く**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-155">Locate and specify the [manifest](../develop/add-in-manifests.md) file to upload and choose **Open**.</span></span>
10. <span data-ttu-id="94d0a-156">[**ドキュメントの追加**] ダイアログで、[**OK**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-156">In the **Add a document** dialog box, choose **OK**.</span></span>

## <a name="insert-office-add-ins-from-the-app-catalog"></a><span data-ttu-id="94d0a-157">アプリ カタログから Office アドインを挿入する</span><span class="sxs-lookup"><span data-stu-id="94d0a-157">Insert Office Add-ins from the app catalog</span></span>

<span data-ttu-id="94d0a-158">オンライン Office アプリケーションの場合は、次の手順を実行してアプリ カタログから Office アドインを見つけることができます。</span><span class="sxs-lookup"><span data-stu-id="94d0a-158">For online Office applications, you can find Office Add-ins from the app catalog by completing the following steps.</span></span>

1. <span data-ttu-id="94d0a-159">オンライン Office アプリケーション (Excel、PowerPoint、または Word) を開きます。</span><span class="sxs-lookup"><span data-stu-id="94d0a-159">Open the online Office application (Excel, PowerPoint, or Word).</span></span>
2. <span data-ttu-id="94d0a-160">文書を作成または開く。</span><span class="sxs-lookup"><span data-stu-id="94d0a-160">Create or open a document.</span></span>
3. <span data-ttu-id="94d0a-161">**[挿入]** > **[アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-161">Choose **Insert** > **Add-ins**.</span></span>
4. <span data-ttu-id="94d0a-162">[Office アドイン] ダイアログの **[自分の所属組織]** タブを選択します。Office アドインのリストが表示されます。</span><span class="sxs-lookup"><span data-stu-id="94d0a-162">In the Office Add-ins dialog, choose the **MY ORGANIZATION** tab.  The Office Add-ins are listed.</span></span>
5. <span data-ttu-id="94d0a-163">Office アドインを選択し、 **追加** を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-163">Choose an Office Add-in and then choose **Add**.</span></span>

<span data-ttu-id="94d0a-164">デスクトップの Office アプリケーションの場合は、次の手順を実行してアプリ カタログから Office アドインを見つけることができます。</span><span class="sxs-lookup"><span data-stu-id="94d0a-164">For Office applications on the desktop, you can find Office Add-ins from the app catalog by completing the following steps.</span></span>

1. <span data-ttu-id="94d0a-165">デスクトップ Office アプリケーション (Excel、Word、または PowerPoint) を開きます。</span><span class="sxs-lookup"><span data-stu-id="94d0a-165">Open the desktop Office application (Excel, Word, or PowerPoint)</span></span>
2. <span data-ttu-id="94d0a-166">**[ファイル]**  >  **[オプション]**  >  **[セキュリティ センター]**  >  **[セキュリティ センターの設定]**  >  **[信頼できるアドイン カタログ]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-166">Choose **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.</span></span>
3. <span data-ttu-id="94d0a-167">[**カタログ URL** ] ボックスに SharePoint アプリ カタログの URL を入力し、**[カタログの追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-167">Enter the URL of the SharePoint app catalog in the **Catalog Url** box and choose **Add catalog**.</span></span>
    <span data-ttu-id="94d0a-168">短い URL を使用します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-168">Use the shorter form of the URL.</span></span> <span data-ttu-id="94d0a-169">たとえば、SharePoint アプリ カタログの URL が次のような場合:</span><span class="sxs-lookup"><span data-stu-id="94d0a-169">For example, if the URL of the Office Add-ins catalog is:</span></span>
    - `https://<domain>/sites/<AddinCatalogSiteCollection>/AgaveCatalog`
    
    <span data-ttu-id="94d0a-170">親サイト コレクションの URL のみを指定します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-170">Specify just the URL of the parent site collection:</span></span>
    - `https://<domain>/sites/<AddinCatalogSiteCollection>`
4. <span data-ttu-id="94d0a-171">Office アプリケーションを閉じてから、もう一度開きます。</span><span class="sxs-lookup"><span data-stu-id="94d0a-171">Close and reopen the Office application.</span></span> 
5. <span data-ttu-id="94d0a-172">**[挿入]** > **[アドインの取得]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-172">Choose **Insert** > **Get Add-ins**.</span></span>
4. <span data-ttu-id="94d0a-173">[Office アドイン] ダイアログの **[自分の所属組織]** タブを選択します。Office アドインのリストが表示されます。</span><span class="sxs-lookup"><span data-stu-id="94d0a-173">In the Office Add-ins dialog, choose the **MY ORGANIZATION** tab.  The Office Add-ins are listed.</span></span>
5. <span data-ttu-id="94d0a-174">Office アドインを選択し、 **追加**を選択します。</span><span class="sxs-lookup"><span data-stu-id="94d0a-174">Choose an Office Add-in and then choose **Add**.</span></span>

<span data-ttu-id="94d0a-175">または、管理者はグループ ポリシーを使用して SharePoint のアプリ カタログを指定できます。</span><span class="sxs-lookup"><span data-stu-id="94d0a-175">Alternatively, an administrator can specify an Office Add-in catalog on SharePoint by using group policy.</span></span> <span data-ttu-id="94d0a-176">関連するポリシー設定を「[365 ProPlus、Office 2019、Office 2016 の管理用テンプレート ファイル (ADMX/ADML)](https://www.microsoft.com/download/details.aspx?id=49030) 」で利用できますし、**ユーザー構成\ポリシー\管理用テンプレート\Microsoft Office 2016\セキュリティ設定\セキュリティ センター\信頼できるカタログ**の下にも見つけられます。</span><span class="sxs-lookup"><span data-stu-id="94d0a-176">The relevant policy settings are available in the [Administrative Template files (ADMX/ADML) for Office 365 ProPlus, Office 2019, and Office 2016](https://www.microsoft.com/download/details.aspx?id=49030) and be found under **User Configuration\Policies\Administrative Templates\Microsoft Office 2016\Security Settings\Trust Center\Trusted Catalogs**.</span></span>

---
title: 作業ウィンドウ アドインとコンテンツ アドインを SharePoint カタログに発行する
description: 組織内のユーザーが Office アドインにアクセスできるようにするために、管理者は組織のアドイン カタログに Office アドインのマニフェスト ファイルをアップロードできます。
ms.date: 05/22/2019
localization_priority: Priority
ms.openlocfilehash: bffbf3e83a2e6d8d0c63252c27ba54826611f78b
ms.sourcegitcommit: adaee1329ae9bb69e49bde7f54a4c0444c9ba642
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/24/2019
ms.locfileid: "34432244"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-catalog"></a><span data-ttu-id="be2a2-103">作業ウィンドウ アドインとコンテンツ アドインを SharePoint カタログに発行する</span><span class="sxs-lookup"><span data-stu-id="be2a2-103">Publish task pane and content add-ins to a SharePoint catalog</span></span>

<span data-ttu-id="be2a2-p101">アドイン カタログは、Office アドインと SharePoint アドインのドキュメント ライブラリをホストする SharePoint Web アプリケーションまたは SharePoint Online テナンシーの専用サイト コレクションです。組織内のユーザーが Office アドインにアクセスできるようにするために、管理者は組織のアドイン カタログに Office アドインのマニフェスト ファイルをアップロードできます。管理者がアドイン カタログを信頼できるカタログとして登録すると、ユーザーは Office クライアント アプリケーションで挿入 UI からアドインを挿入できます。</span><span class="sxs-lookup"><span data-stu-id="be2a2-p101">An add-in catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for Office and SharePoint Add-ins. To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the add-in catalog for their organization. When an administrator registers an add-in catalog as a trusted catalog, users can insert the add-in from the insertion UI in an Office client application.</span></span>

> [!IMPORTANT]
> - <span data-ttu-id="be2a2-106">SharePoint のアドイン カタログでは、アドイン コマンドなど、[アドイン マニフェスト](../develop/add-in-manifests.md)の `VersionOverrides` ノードで実装されるアドイン機能がサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="be2a2-106">Add-in catalogs on SharePoint do not support add-in features that are implemented in the `VersionOverrides` node of the [add-in manifest](../develop/add-in-manifests.md), such as add-in commands.</span></span>
> - <span data-ttu-id="be2a2-107">クラウド環境またはハイブリッド環境をターゲットにしている場合は、アドインの発行に [Office 365 管理センターからの一元展開を使用する](../publish/centralized-deployment.md)ことをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="be2a2-107">If you’re targeting a cloud or hybrid environment, we recommend that you [use Centralized Deployment via the Office 365 admin center](../publish/centralized-deployment.md) to publish your add-ins.</span></span>
> - <span data-ttu-id="be2a2-108">SharePoint カタログは Office for Mac ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="be2a2-108">SharePoint catalogs are not supported for Office for Mac.</span></span> <span data-ttu-id="be2a2-109">Office アドインを Mac クライアントに展開するには、そのアドインを [AppSource](/office/dev/store/submit-to-the-office-store) に提出する必要があります。</span><span class="sxs-lookup"><span data-stu-id="be2a2-109">To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](/office/dev/store/submit-to-the-office-store).</span></span>   

## <a name="create-an-add-in-catalog"></a><span data-ttu-id="be2a2-110">アドイン カタログの作成</span><span class="sxs-lookup"><span data-stu-id="be2a2-110">Create an add-in catalog</span></span>

<span data-ttu-id="be2a2-111">次のいずれかのセクションに示す手順を完了して、SharePoint または Office 365 にアドイン カタログを作成します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-111">Complete the steps in one of the following sections to set up an add-in catalog on SharePoint or on Office 365.</span></span>

### <a name="to-create-an-add-in-catalog-for-on-premises-sharepoint"></a><span data-ttu-id="be2a2-112">オンプレミス SharePoint 上でアドイン カタログを作成するには</span><span class="sxs-lookup"><span data-stu-id="be2a2-112">To set up an add-in catalog for on-premises SharePoint</span></span>

> [!NOTE]
> <span data-ttu-id="be2a2-113">オンプレミス SharePoint の UI はアドインを**アプリ**として参照します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-113">The UI in on-premises SharePoint still refers to add-ins as **apps**.</span></span>

1. <span data-ttu-id="be2a2-114">**サーバーの全体管理 Web サイト**を参照します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-114">Browse to the  **Central Administration Site**.</span></span>

2. <span data-ttu-id="be2a2-115">左側の作業ウィンドウで、**[アプリ]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-115">In the left task pane, choose  **Apps**.</span></span>

3. <span data-ttu-id="be2a2-116">**[アプリ]** ページの **[アプリの管理]** で **[アプリ カタログの管理]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-116">On the  **Apps** page, under **App Management**, choose  **Manage App Catalog**.</span></span>

4. <span data-ttu-id="be2a2-117">**[アプリ カタログの管理]** ページの **Web アプリケーション セレクター** で正しい Web アプリケーションが選択されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-117">On the  **Manage App Catalog** page, make sure you have the right web application selected in the **Web Application Selector**.</span></span>

5. <span data-ttu-id="be2a2-118">**[サイトの設定の表示]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-118">Choose  **View site settings**.</span></span>

6. <span data-ttu-id="be2a2-119">[ **サイトの設定**] ページで、[ **サイト コレクション管理者**] を選択してサイト コレクション管理者を指定してから、[ **OK**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-119">On the  **Site Settings** page, choose **Site collection administrators** to specify the site collection administrators, and then choose **OK**.</span></span>

7. <span data-ttu-id="be2a2-120">ユーザーにサイト アクセス許可を付与するには、[ **サイトの権限**] を選択してから、[ **アクセス許可の付与**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-120">To grant site permissions to users, choose  **Site Permissions**, and then choose  **Grant Permissions**.</span></span>

8. <span data-ttu-id="be2a2-121">[ **アプリ カタログ サイトの共有**] ダイアログ ボックスで、1 人以上のサイト ユーザーを指定して、それらに適切なアクセス許可を設定し、必要に応じて他のオプションを設定してから、[  **共有**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-121">In the  **Share 'App Catalog Site'** dialog box, specify one or more site users, set the appropriate permissions for them, optionally set other options, and then choose **Share**.</span></span>

9. <span data-ttu-id="be2a2-122">Office アドインのアドイン カタログにアドインを追加する場合は、**[Office 用アプリ]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-122">To add an add-in to the Office Add-ins add-in catalog, choose **Apps for Office**.</span></span>

### <a name="to-create-an-app-catalog-on-office-365"></a><span data-ttu-id="be2a2-123">Office 365 でアプリ カタログを作成するには</span><span class="sxs-lookup"><span data-stu-id="be2a2-123">To create an app catalog on Office 365</span></span>

<span data-ttu-id="be2a2-124">SharePoint ではカタログを「アプリ」カタログと名付けますが、Office アドインをアプリ カタログに登録できます。</span><span class="sxs-lookup"><span data-stu-id="be2a2-124">Even though SharePoint names the catalog an "app" catalog, you can register Office Add-ins in the app catalog.</span></span>

1. <span data-ttu-id="be2a2-125">Microsoft 365 管理センターに移動します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-125">Go to the Microsoft 365 admin center.</span></span> <span data-ttu-id="be2a2-126">管理センターの検索方法については、「[Microsoft 365 管理センターについて](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="be2a2-126">For information on how to find the admin center, see [About the Microsoft 365 admin center](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).</span></span>

2. <span data-ttu-id="be2a2-127">Microsoft 365 管理センターのページで [**管理センター**] のリストを展開し、[**SharePoint**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-127">On the Microsoft 365 admin center page, expand the list of **Admin centers**, and then choose **SharePoint**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="be2a2-128">カタログを作成するには、従来の SharePoint 管理センターを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="be2a2-128">You need to use the Classic SharePoint admin center to create the catalog.</span></span> <span data-ttu-id="be2a2-129">新しい SharePoint 管理センターにいる場合は、左側のウィンドウで「**従来の SharePoint 管理センター**」を選択します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-129">If you are in the new SharePoint admin center, choose **Classic SharePoint admin center** in the left pane.</span></span>

3. <span data-ttu-id="be2a2-130">左側の作業ウィンドウで、[**アプリ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-130">In the left task pane, choose  **Apps**.</span></span>

4. <span data-ttu-id="be2a2-131">[**アプリ**] ページで、[**アプリ カタログ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-131">On the **apps** page, select **App Catalog**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="be2a2-132">アプリ カタログが既に作成されてこのページに表示されている場合は、残りの手順をスキップしてこの記事の次のセクションに移動し、カタログにアドインを発行できます。</span><span class="sxs-lookup"><span data-stu-id="be2a2-132">If an app catalog is already created and appears on this page, then you can skip the rest of these steps and go to the next section of this article to publish your add-in to the catalog.</span></span>

5. <span data-ttu-id="be2a2-133">[**アプリ カタログ サイト**] ページで、[**OK**] を選択して既定のオプションを受け入れ、新しいアドイン カタログ サイトを作成します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-133">On the  **Add-in Catalog Site** page, choose **OK** to accept the default option and create a new add-in catalog site.</span></span>

6. <span data-ttu-id="be2a2-134">[**アプリ カタログ サイト コレクションを作成する**] ページで、アプリ カタログ サイトのタイトルを指定します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-134">On the  **Create Add-in Catalog Site Collection** page, specify the title of your Add-in Catalog site.</span></span>

7. <span data-ttu-id="be2a2-135">[**Web サイトのアドレス**]を指定します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-135">Specify the web site address.</span></span>

8. <span data-ttu-id="be2a2-136">[**管理者**]を指定します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-136">Specify an **Administrator**.</span></span>

9. <span data-ttu-id="be2a2-137">[**サーバー リソース クォータ**] に 0 (ゼロ) を設定します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-137">Set the Server Resource Quota to 0 (zero), and then select OK.</span></span> <span data-ttu-id="be2a2-138">(サーバー リソース クォータは、パフォーマンスが低いサンドボックス ソリューションのスロットルに関連していますが、このアプリ カタログ サイトにはサンドボックス ソリューションをインストールしません。)</span><span class="sxs-lookup"><span data-stu-id="be2a2-138">(The server resource quota is related to throttling poorly performing sandboxed solutions, but you won't be installing any sandboxed solutions on your add-in catalog site.)</span></span>

10. <span data-ttu-id="be2a2-139">[**OK**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-139">Choose **OK**.</span></span>

<span data-ttu-id="be2a2-140">アプリ カタログが作成されました。</span><span class="sxs-lookup"><span data-stu-id="be2a2-140">The app catalog is now created.</span></span>

## <a name="publish-an-add-in-to-an-app-catalog"></a><span data-ttu-id="be2a2-141">アプリ カタログへのアドインの発行</span><span class="sxs-lookup"><span data-stu-id="be2a2-141">Publish an add-in to an add-in catalog</span></span>

<span data-ttu-id="be2a2-142">既存のアプリ カタログにアドインを発行するには、次に示す手順を完了します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-142">To publish an add-in to an add-in catalog, complete the following steps.</span></span>

1. <span data-ttu-id="be2a2-143">Microsoft 365 管理センターに移動します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-143">Go to the Microsoft 365 admin center.</span></span> <span data-ttu-id="be2a2-144">管理センターの検索方法については、「[Microsoft 365 管理センターについて](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="be2a2-144">For information on how to find the admin center, see [About the Microsoft 365 admin center](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).</span></span>
2. <span data-ttu-id="be2a2-145">Microsoft 365 管理センターのページで [**管理センター**] のリストを展開し、[**SharePoint**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-145">On the Microsoft 365 admin center page, expand the list of **Admin centers**, and then choose **SharePoint**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="be2a2-146">カタログを作成するには、従来の SharePoint 管理センターを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="be2a2-146">You need to use the Classic SharePoint admin center to create the catalog.</span></span> <span data-ttu-id="be2a2-147">新しい SharePoint 管理センターにいる場合は、左側のウィンドウで「**従来の SharePoint 管理センター**」を選択します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-147">If you are in the new SharePoint admin center, choose **Classic SharePoint admin center** in the left pane.</span></span>
3. <span data-ttu-id="be2a2-148">左側の作業ウィンドウで、[**アプリ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-148">In the left task pane, choose  **Apps**.</span></span>
4. <span data-ttu-id="be2a2-149">[**アプリ**] ページで、[**アプリ カタログ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-149">On the **apps** page, select **App Catalog**.</span></span>
5. <span data-ttu-id="be2a2-150">[**Office 用アプリを配信する**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-150">Choose **Distribute apps for Office**.</span></span>
6. <span data-ttu-id="be2a2-151">[**Office 用アプリ**] ページで、[**新規**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-151">In the **Apps for Office** page, choose **New**.</span></span>
7. <span data-ttu-id="be2a2-152">[**ドキュメントの追加**] ダイアログで、[**ファイルの選択**] ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="be2a2-152">In the **Add a document** dialog, select the **Choose Files** button.</span></span>
8. <span data-ttu-id="be2a2-153">アップロードする[マニフェスト](../develop/add-in-manifests.md) ファイルを見つけて指定し、[**開く**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-153">Locate and specify the [manifest](../develop/add-in-manifests.md) file to upload and choose **Open**.</span></span>
9. <span data-ttu-id="be2a2-154">[**ドキュメントの追加**] ダイアログで、[**OK**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-154">In the **Add a document** dialog box, choose **OK**.</span></span>

    <span data-ttu-id="be2a2-p108">このカタログのコンテンツおよび作業ウィンドウのアドインが **[Office アドイン]** ダイアログ ボックスから使用できるようになりました。これらにアクセスするには、**[挿入]** タブで **[個人用アドイン]** を選択して、**[自分の所属組織]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-p108">Content and task pane add-ins in this catalog are now available from the  **Office Add-ins** dialog box. To access them, choose **My Add-ins** on the **Insert** tab, and then choose **MY ORGANIZATION**.</span></span>

## <a name="end-user-experience-with-the-add-in-catalog"></a><span data-ttu-id="be2a2-157">アドイン カタログのエンド ユーザー エクスペリエンス</span><span class="sxs-lookup"><span data-stu-id="be2a2-157">End user experience with the add-in catalog</span></span>

<span data-ttu-id="be2a2-158">エンド ユーザーは、次に示す手順を実行することで Office アプリケーションのアドイン カタログにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="be2a2-158">End users can access the add-in catalog in an Office application by completing the following steps:</span></span>

1. <span data-ttu-id="be2a2-159">Office アプリケーションで、**[ファイル]**  >  **[オプション]**  >  **[セキュリティ センター]**  >  **[セキュリティ センターの設定]**  >  **[信頼できるアドイン カタログ]** の順に移動します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-159">In the Office application, go to  **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.</span></span>

2. <span data-ttu-id="be2a2-160">アドイン カタログの _親 SharePoint サイト コレクション_の URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-160">Specify the URL of the  _parent SharePoint site collection_ of the add-in catalog.</span></span> 

    <span data-ttu-id="be2a2-161">たとえば、Office アドイン カタログの URL が次のような場合:</span><span class="sxs-lookup"><span data-stu-id="be2a2-161">For example, if the URL of the Office Add-ins catalog is:</span></span>

    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog`

    <span data-ttu-id="be2a2-162">親サイト コレクションの URL のみを指定します。</span><span class="sxs-lookup"><span data-stu-id="be2a2-162">Specify just the URL of the parent site collection:</span></span>

    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_`

3. <span data-ttu-id="be2a2-p109">Office アプリケーションを閉じ、もう一度開きます。アドイン カタログが **[Office アドイン]** ダイアログ ボックスに表示されます。</span><span class="sxs-lookup"><span data-stu-id="be2a2-p109">Close and reopen the Office application. The add-in catalog will be available in the **Office Add-ins** dialog box.</span></span>

<span data-ttu-id="be2a2-165">または、管理者はグループ ポリシーを使用して SharePoint の Office アドイン カタログを指定できます。</span><span class="sxs-lookup"><span data-stu-id="be2a2-165">Alternatively, an administrator can specify an Office Add-in catalog on SharePoint by using group policy.</span></span> <span data-ttu-id="be2a2-166">詳細については、「[グループ ポリシーを使用して、ユーザーが Office アドインをインストールおよび使用する方法を管理する](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office)」のセクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="be2a2-166">For details, see the section [Using Group Policy to manage how users can install and use Office Add-ins](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).</span></span>

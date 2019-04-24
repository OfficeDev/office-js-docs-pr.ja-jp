---
title: 作業ウィンドウ アドインとコンテンツ アドインを SharePoint カタログに発行する
description: 組織内のユーザーが Office アドインにアクセスできるようにするために、管理者は組織のアドイン カタログに Office アドインのマニフェスト ファイルをアップロードできます。
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: ed4f9778e4cd7dccba00d2e8c019bd4441b70eeb
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450801"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-catalog"></a><span data-ttu-id="a43d5-103">作業ウィンドウ アドインとコンテンツ アドインを SharePoint カタログに発行する</span><span class="sxs-lookup"><span data-stu-id="a43d5-103">Publish task pane and content add-ins to a SharePoint catalog</span></span>

<span data-ttu-id="a43d5-p101">アドイン カタログは、Office アドインと SharePoint アドインのドキュメント ライブラリをホストする SharePoint Web アプリケーションまたは SharePoint Online テナンシーの専用サイト コレクションです。組織内のユーザーが Office アドインにアクセスできるようにするために、管理者は組織のアドイン カタログに Office アドインのマニフェスト ファイルをアップロードできます。管理者がアドイン カタログを信頼できるカタログとして登録すると、ユーザーは Office クライアント アプリケーションで挿入 UI からアドインを挿入できます。</span><span class="sxs-lookup"><span data-stu-id="a43d5-p101">An add-in catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for Office and SharePoint Add-ins. To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the add-in catalog for their organization. When an administrator registers an add-in catalog as a trusted catalog, users can insert the add-in from the insertion UI in an Office client application.</span></span>

> [!IMPORTANT]
> - <span data-ttu-id="a43d5-106">SharePoint のアドイン カタログでは、アドイン コマンドなど、[アドイン マニフェスト](../develop/add-in-manifests.md)の `VersionOverrides` ノードで実装されるアドイン機能がサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="a43d5-106">Add-in catalogs on SharePoint do not support add-in features that are implemented in the `VersionOverrides` node of the [add-in manifest](../develop/add-in-manifests.md), such as add-in commands.</span></span>
> - <span data-ttu-id="a43d5-107">クラウド環境またはハイブリッド環境をターゲットにしている場合は、アドインの発行に [Office 365 管理センターからの一元展開を使用する](../publish/centralized-deployment.md)ことをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="a43d5-107">If you’re targeting a cloud or hybrid environment, we recommend that you [use Centralized Deployment via the Office 365 admin center](../publish/centralized-deployment.md) to publish your add-ins.</span></span>
> - <span data-ttu-id="a43d5-108">SharePoint カタログは Office for Mac ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="a43d5-108">SharePoint catalogs are not supported for Office for Mac.</span></span> <span data-ttu-id="a43d5-109">Office アドインを Mac クライアントに展開するには、そのアドインを [AppSource](/office/dev/store/submit-to-the-office-store) に提出する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a43d5-109">To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](/office/dev/store/submit-to-the-office-store).</span></span>   

## <a name="set-up-an-add-in-catalog"></a><span data-ttu-id="a43d5-110">アドイン カタログのセットアップ</span><span class="sxs-lookup"><span data-stu-id="a43d5-110">Set up an add-in catalog</span></span>

<span data-ttu-id="a43d5-111">次のいずれかのセクションに示す手順を完了して、SharePoint または Office 365 にアドイン カタログをセットアップします。</span><span class="sxs-lookup"><span data-stu-id="a43d5-111">Complete the steps in one of the following sections to set up an add-in catalog on SharePoint or on Office 365.</span></span>

### <a name="to-set-up-an-add-in-catalog-for-on-premises-sharepoint"></a><span data-ttu-id="a43d5-112">オンプレミス SharePoint 上でアドイン カタログをセットアップするには</span><span class="sxs-lookup"><span data-stu-id="a43d5-112">To set up an add-in catalog for on-premises SharePoint</span></span>

> [!NOTE]
> <span data-ttu-id="a43d5-113">オンプレミス SharePoint の UI はアドインを**アプリ**として参照します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-113">The UI in on-premises SharePoint still refers to add-ins as **apps**.</span></span>

1. <span data-ttu-id="a43d5-114">**サーバーの全体管理 Web サイト**を参照します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-114">Browse to the  **Central Administration Site**.</span></span>

2. <span data-ttu-id="a43d5-115">左側の作業ウィンドウで、**[アプリ]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-115">In the left task pane, choose  **Apps**.</span></span>

3. <span data-ttu-id="a43d5-116">**[アプリ]** ページの **[アプリの管理]** で **[アプリ カタログの管理]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-116">On the  **Apps** page, under **App Management**, choose  **Manage App Catalog**.</span></span>

4. <span data-ttu-id="a43d5-117">**[アプリ カタログの管理]** ページの **Web アプリケーション セレクター** で正しい Web アプリケーションが選択されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-117">On the  **Manage App Catalog** page, make sure you have the right web application selected in the **Web Application Selector**.</span></span>

5. <span data-ttu-id="a43d5-118">**[サイトの設定の表示]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-118">Choose  **View site settings**.</span></span>

6. <span data-ttu-id="a43d5-119">[ **サイトの設定**] ページで、[ **サイト コレクション管理者**] を選択してサイト コレクション管理者を指定してから、[ **OK**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-119">On the  **Site Settings** page, choose **Site collection administrators** to specify the site collection administrators, and then choose **OK**.</span></span>

7. <span data-ttu-id="a43d5-120">ユーザーにサイト アクセス許可を付与するには、[ **サイトの権限**] を選択してから、[ **アクセス許可の付与**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-120">To grant site permissions to users, choose  **Site Permissions**, and then choose  **Grant Permissions**.</span></span>

8. <span data-ttu-id="a43d5-121">[ **アプリ カタログ サイトの共有**] ダイアログ ボックスで、1 人以上のサイト ユーザーを指定して、それらに適切なアクセス許可を設定し、必要に応じて他のオプションを設定してから、[  **共有**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-121">In the  **Share 'App Catalog Site'** dialog box, specify one or more site users, set the appropriate permissions for them, optionally set other options, and then choose **Share**.</span></span>

9. <span data-ttu-id="a43d5-122">Office アドインのアドイン カタログにアドインを追加する場合は、**[Office 用アプリ]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-122">To add an add-in to the Office Add-ins add-in catalog, choose **Apps for Office**.</span></span>

### <a name="to-set-up-an-add-in-catalog-on-office-365"></a><span data-ttu-id="a43d5-123">Office 365 でアドイン カタログをセットアップするには</span><span class="sxs-lookup"><span data-stu-id="a43d5-123">To set up an add-in catalog on Office 365</span></span>

1. <span data-ttu-id="a43d5-124">[Office 365 管理センター] ページで、 **[管理]**、 **[SharePoint]** の順にクリックします。</span><span class="sxs-lookup"><span data-stu-id="a43d5-124">On the Office 365 admin center page, choose  **Admin**, and then choose  **SharePoint**.</span></span>

2. <span data-ttu-id="a43d5-125">左側の作業ウィンドウで、[ **アドイン**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-125">In the left task pane, choose  **add-ins**.</span></span>

3. <span data-ttu-id="a43d5-126">[ **アドイン**] ページで、[ **アドイン カタログ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-126">On the  **add-ins** page, choose **Add-in Catalog**.</span></span>

4. <span data-ttu-id="a43d5-127">[ **アドイン カタログ サイト**] ページで、[ **OK**] を選択して既定のオプションを受け入れ、新しいアドイン カタログ サイトを作成します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-127">On the  **Add-in Catalog Site** page, choose **OK** to accept the default option and create a new add-in catalog site.</span></span>

5. <span data-ttu-id="a43d5-128">[ **アドイン カタログ サイト コレクションの作成**] ページで、アドイン カタログ サイトのタイトルを指定します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-128">On the  **Create Add-in Catalog Site Collection** page, specify the title of your Add-in Catalog site.</span></span>

6. <span data-ttu-id="a43d5-129">Web サイト アドレスを指定します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-129">Specify the web site address.</span></span>

7. <span data-ttu-id="a43d5-p103">[ **記憶域のクォータ**] を可能な限り小さい値に設定します (現在は 110)。このサイト コレクションにはアドイン パッケージだけをインストールしますが、パッケージは非常に小さなものです。</span><span class="sxs-lookup"><span data-stu-id="a43d5-p103">Set the  **Storage Quota** to the lowest possible value (currently 110). You will only be installing add-in packages on this site collection and they are very small.</span></span>

8. <span data-ttu-id="a43d5-p104">[ **サーバー リソース クォータ**] を 0 (ゼロ) に設定します。(サーバー リソース クォータは、パフォーマンスが低いサンドボックス ソリューションのスロットルに関連していますが、このアドインのカタログ サイトにはサンドボックス ソリューションをインストールしません。)</span><span class="sxs-lookup"><span data-stu-id="a43d5-p104">Set the  **Server Resource Quota** to 0 (zero). (The server resource quota is related to throttling poorly performing sandboxed solutions, but you won't be installing any sandboxed solutions on your add-in catalog site.)</span></span>

9. <span data-ttu-id="a43d5-134">**[OK]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="a43d5-134">Choose  **OK**.</span></span>

10. <span data-ttu-id="a43d5-p105">アドイン カタログ サイトにアドインを追加するために、前の手順で作成したサイトを参照します。左側のナビゲーション ウィンドウで、**[Office アドイン]** を選択して Office アドイン マニフェスト ファイルをアップロードし、**[新規アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-p105">To add an add-in to the Add-in Catalog Site, browse to the site you have just created. In the left navigation pane, choose  **Office Add-ins**, and then, to upload an Office Add-in manifest file, choose  **new add-in**.</span></span>

## <a name="publish-an-add-in-to-an-add-in-catalog"></a><span data-ttu-id="a43d5-137">アドイン カタログへのアドインの発行</span><span class="sxs-lookup"><span data-stu-id="a43d5-137">Publish an add-in to an add-in catalog</span></span>

<span data-ttu-id="a43d5-138">アドイン カタログにアドインを発行するには、次に示す手順を完了します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-138">To publish an add-in to an add-in catalog, complete the following steps.</span></span>

1. <span data-ttu-id="a43d5-139">アドイン カタログを参照します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-139">Browse to the add-in catalog:</span></span>

    - <span data-ttu-id="a43d5-140">SharePoint サーバーの全体管理メイン ページを開きます。</span><span class="sxs-lookup"><span data-stu-id="a43d5-140">Open the SharePoint Central Administration main page.</span></span>

    - <span data-ttu-id="a43d5-141">**[アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-141">Select  **Add-ins**.</span></span>

    - <span data-ttu-id="a43d5-142">**[アドイン カタログの管理]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-142">Select  **Manage Add-in Catalog**.</span></span>

    - <span data-ttu-id="a43d5-143">表示されたリンクを選択し、左側のナビゲーション バーで **[Office アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-143">Choose the link provided, and then choose  **Office Add-ins** on the left navigation bar.</span></span>

2. <span data-ttu-id="a43d5-144">**[新しいアイテムの追加]** リンクを選択します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-144">Choose the  **Click to add new item** link.</span></span>

3. <span data-ttu-id="a43d5-145">**[参照]** を選択し、アップロードする [[マニフェスト]](../develop/add-in-manifests.md) を指定します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-145">Choose  **Browse**, and then specify the [manifest](../develop/add-in-manifests.md) to upload.</span></span>

    <span data-ttu-id="a43d5-p106">このカタログのコンテンツおよび作業ウィンドウのアドインが **[Office アドイン]** ダイアログ ボックスから使用できるようになりました。これらにアクセスするには、**[挿入]** タブで **[個人用アドイン]** を選択して、**[自分の所属組織]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-p106">Content and task pane add-ins in this catalog are now available from the  **Office Add-ins** dialog box. To access them, choose **My Add-ins** on the **Insert** tab, and then choose **MY ORGANIZATION**.</span></span>

## <a name="end-user-experience-with-the-add-in-catalog"></a><span data-ttu-id="a43d5-148">アドイン カタログのエンド ユーザー エクスペリエンス</span><span class="sxs-lookup"><span data-stu-id="a43d5-148">End user experience with the add-in catalog</span></span>

<span data-ttu-id="a43d5-149">エンド ユーザーは、次に示す手順を実行することで Office アプリケーションのアドイン カタログにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="a43d5-149">End users can access the add-in catalog in an Office application by completing the following steps:</span></span>

1. <span data-ttu-id="a43d5-150">Office アプリケーションで、**[ファイル]**  >  **[オプション]**  >  **[セキュリティ センター]**  >  **[セキュリティ センターの設定]**  >  **[信頼できるアドイン カタログ]** の順に移動します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-150">In the Office application, go to  **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.</span></span>

2. <span data-ttu-id="a43d5-151">アドイン カタログの _親 SharePoint サイト コレクション_の URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-151">Specify the URL of the  _parent SharePoint site collection_ of the add-in catalog.</span></span> 

    <span data-ttu-id="a43d5-152">たとえば、Office アドイン カタログの URL が次のような場合:</span><span class="sxs-lookup"><span data-stu-id="a43d5-152">For example, if the URL of the Office Add-ins catalog is:</span></span>

    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog`

    <span data-ttu-id="a43d5-153">親サイト コレクションの URL のみを指定します。</span><span class="sxs-lookup"><span data-stu-id="a43d5-153">Specify just the URL of the parent site collection:</span></span>

    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_`

3. <span data-ttu-id="a43d5-p107">Office アプリケーションを閉じ、もう一度開きます。アドイン カタログが **[Office アドイン]** ダイアログ ボックスに表示されます。</span><span class="sxs-lookup"><span data-stu-id="a43d5-p107">Close and reopen the Office application. The add-in catalog will be available in the **Office Add-ins** dialog box.</span></span>

<span data-ttu-id="a43d5-156">または、管理者はグループ ポリシーを使用して SharePoint の Office アドイン カタログを指定できます。</span><span class="sxs-lookup"><span data-stu-id="a43d5-156">Alternatively, an administrator can specify an Office Add-in catalog on SharePoint by using group policy.</span></span> <span data-ttu-id="a43d5-157">詳細については、「[グループ ポリシーを使用して、ユーザーが Office アドインをインストールおよび使用する方法を管理する](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office)」のセクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="a43d5-157">For details, see the section [Using Group Policy to manage how users can install and use Office Add-ins](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).</span></span>

---
title: Office アドインを展開し、発行する | Microsoft Docs
description: テスト目的またはユーザーに配布する目的で Office アドインを展開するための方法とオプション。
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: eeaf4b61948952ff7e536f3e1a6b38dc46adb93e
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871704"
---
# <a name="deploy-and-publish-your-office-add-in"></a><span data-ttu-id="c7e8a-103">Office アドインを展開し、発行する</span><span class="sxs-lookup"><span data-stu-id="c7e8a-103">Deploy and publish your Office Add-in</span></span>

<span data-ttu-id="c7e8a-104">さまざまな方法を利用し、テスト目的またはユーザーに配布する目的で、Office アドインを展開できます。</span><span class="sxs-lookup"><span data-stu-id="c7e8a-104">You can use one of several methods to deploy your Office Add-in for testing or distribution to users.</span></span>

|<span data-ttu-id="c7e8a-105">**メソッド**</span><span class="sxs-lookup"><span data-stu-id="c7e8a-105">**Method**</span></span>|<span data-ttu-id="c7e8a-106">**Use...**</span><span class="sxs-lookup"><span data-stu-id="c7e8a-106">**Use...**</span></span>|
|:---------|:------------|
|[<span data-ttu-id="c7e8a-107">サイドロード</span><span class="sxs-lookup"><span data-stu-id="c7e8a-107">Sideloading</span></span>](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|<span data-ttu-id="c7e8a-108">開発プロセスの一環として、Windows、Office Online、iPad、Mac で実行するアドインをテストします。</span><span class="sxs-lookup"><span data-stu-id="c7e8a-108">As part of your development process, to test your add-in running on Windows, Office Online, iPad, or Mac.</span></span>|
|[<span data-ttu-id="c7e8a-109">一元展開</span><span class="sxs-lookup"><span data-stu-id="c7e8a-109">Centralized Deployment</span></span>](centralized-deployment.md)|<span data-ttu-id="c7e8a-110">クラウド環境またはハイブリッド環境で、Office 365 管理センターを使用して組織内のユーザーにアドインを配布します。</span><span class="sxs-lookup"><span data-stu-id="c7e8a-110">In a cloud or hybrid deployment, to distribute your add-in to users in your organization by using the Office 365 admin center.</span></span>|
|[<span data-ttu-id="c7e8a-111">SharePoint カタログ</span><span class="sxs-lookup"><span data-stu-id="c7e8a-111">SharePoint catalog</span></span>](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|<span data-ttu-id="c7e8a-112">オンプレミス環境で、組織内のユーザーにアドインを配布します。</span><span class="sxs-lookup"><span data-stu-id="c7e8a-112">In an on-premises environment, to distribute your add-in to users in your organization.</span></span>|
|[<span data-ttu-id="c7e8a-113">AppSource</span><span class="sxs-lookup"><span data-stu-id="c7e8a-113">AppSource</span></span>](/office/dev/store/submit-to-the-office-store)|<span data-ttu-id="c7e8a-114">ユーザーに配布する目的でアドインを公開します。</span><span class="sxs-lookup"><span data-stu-id="c7e8a-114">To distribute your add-in publicly to users.</span></span>|
|[<span data-ttu-id="c7e8a-115">Exchange サーバー</span><span class="sxs-lookup"><span data-stu-id="c7e8a-115">Exchange server</span></span>](#outlook-add-in-deployment)|<span data-ttu-id="c7e8a-116">オンプレミス環境またはオンライン環境で、ユーザーに Outlook アドインを配布します。</span><span class="sxs-lookup"><span data-stu-id="c7e8a-116">In an on-premises or online environment, to distribute Outlook add-ins to users.</span></span>|
|[<span data-ttu-id="c7e8a-117">ネットワーク共有</span><span class="sxs-lookup"><span data-stu-id="c7e8a-117">Network share</span></span>](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|<span data-ttu-id="c7e8a-118">アドインをホストさせようとしているネットワーク上の Windows コンピューターで、共有フォルダー カタログとして使用するフォルダーの親フォルダーまたはドライブ文字に移動します。</span><span class="sxs-lookup"><span data-stu-id="c7e8a-118">On a Windows computer on a network where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>|

> [!NOTE]
> <span data-ttu-id="c7e8a-p101">AppSource にアドインを[公開](../publish/publish.md)し、Office エクスペリエンスで利用できるようにする予定がある場合は、[AppSource の検証ポリシー](/office/dev/store/validation-policies)に準拠していることを確認してください。たとえば、検証に合格するには、定義したメソッドをサポートするすべてのプラットフォームでアドインが動作する必要があります (詳細については、[セクション 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) と [Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)のページを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="c7e8a-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="deployment-options-by-office-host"></a><span data-ttu-id="c7e8a-121">Office のホストごとの展開オプション</span><span class="sxs-lookup"><span data-stu-id="c7e8a-121">Deployment options by Office host</span></span>

<span data-ttu-id="c7e8a-122">選択可能な展開オプションは、対象の Office ホストや作成するアドインの種類によって異なります。</span><span class="sxs-lookup"><span data-stu-id="c7e8a-122">The deployment options that are available depend on the Office host that you're targeting and the type of add-in you create.</span></span>

### <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a><span data-ttu-id="c7e8a-123">Word、Excel、PowerPoint のアドインの展開オプション</span><span class="sxs-lookup"><span data-stu-id="c7e8a-123">Deployment options for Word, Excel, and PowerPoint add-ins</span></span>

| <span data-ttu-id="c7e8a-124">拡張点</span><span class="sxs-lookup"><span data-stu-id="c7e8a-124">Extension point</span></span> | <span data-ttu-id="c7e8a-125">サイドロード</span><span class="sxs-lookup"><span data-stu-id="c7e8a-125">Sideloading</span></span> | <span data-ttu-id="c7e8a-126">Office 365 管理センター</span><span class="sxs-lookup"><span data-stu-id="c7e8a-126">Office 365 admin center</span></span> |<span data-ttu-id="c7e8a-127">AppSource</span><span class="sxs-lookup"><span data-stu-id="c7e8a-127">AppSource</span></span>   | <span data-ttu-id="c7e8a-128">SharePoint カタログ\*</span><span class="sxs-lookup"><span data-stu-id="c7e8a-128">SharePoint catalog\*</span></span> |
|:----------------|:-----------:|:-----------------------:|:----------:|:--------------------:|
| <span data-ttu-id="c7e8a-129">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c7e8a-129">Content</span></span>         | <span data-ttu-id="c7e8a-130">X</span><span class="sxs-lookup"><span data-stu-id="c7e8a-130">X</span></span>           | <span data-ttu-id="c7e8a-131">X</span><span class="sxs-lookup"><span data-stu-id="c7e8a-131">X</span></span>                       | <span data-ttu-id="c7e8a-132">X</span><span class="sxs-lookup"><span data-stu-id="c7e8a-132">X</span></span>          | <span data-ttu-id="c7e8a-133">X</span><span class="sxs-lookup"><span data-stu-id="c7e8a-133">X</span></span>                    |
| <span data-ttu-id="c7e8a-134">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c7e8a-134">Task pane</span></span>       | <span data-ttu-id="c7e8a-135">X</span><span class="sxs-lookup"><span data-stu-id="c7e8a-135">X</span></span>           | <span data-ttu-id="c7e8a-136">X</span><span class="sxs-lookup"><span data-stu-id="c7e8a-136">X</span></span>                       | <span data-ttu-id="c7e8a-137">X</span><span class="sxs-lookup"><span data-stu-id="c7e8a-137">X</span></span>          | <span data-ttu-id="c7e8a-138">X</span><span class="sxs-lookup"><span data-stu-id="c7e8a-138">X</span></span>                    |
| <span data-ttu-id="c7e8a-139">コマンド</span><span class="sxs-lookup"><span data-stu-id="c7e8a-139">Command</span></span>         | <span data-ttu-id="c7e8a-140">X</span><span class="sxs-lookup"><span data-stu-id="c7e8a-140">X</span></span>           | <span data-ttu-id="c7e8a-141">X</span><span class="sxs-lookup"><span data-stu-id="c7e8a-141">X</span></span>                       | <span data-ttu-id="c7e8a-142">X</span><span class="sxs-lookup"><span data-stu-id="c7e8a-142">X</span></span>          |                      |

<span data-ttu-id="c7e8a-143">&#42; SharePoint カタログは、Office for Mac をサポートしません。</span><span class="sxs-lookup"><span data-stu-id="c7e8a-143">&#42; SharePoint catalogs do not support Office for Mac.</span></span>

### <a name="deployment-options-for-outlook-add-ins"></a><span data-ttu-id="c7e8a-144">Outlook アドインの展開オプション</span><span class="sxs-lookup"><span data-stu-id="c7e8a-144">Deployment options for Outlook add-ins</span></span>

| <span data-ttu-id="c7e8a-145">拡張点</span><span class="sxs-lookup"><span data-stu-id="c7e8a-145">Extension point</span></span> | <span data-ttu-id="c7e8a-146">サイドロード</span><span class="sxs-lookup"><span data-stu-id="c7e8a-146">Sideloading</span></span> | <span data-ttu-id="c7e8a-147">Exchange サーバー</span><span class="sxs-lookup"><span data-stu-id="c7e8a-147">Exchange server</span></span> | <span data-ttu-id="c7e8a-148">AppSource</span><span class="sxs-lookup"><span data-stu-id="c7e8a-148">AppSource</span></span>    |
|:----------------|:-----------:|:---------------:|:------------:|
| <span data-ttu-id="c7e8a-149">メール アプリ</span><span class="sxs-lookup"><span data-stu-id="c7e8a-149">Mail app</span></span>        | <span data-ttu-id="c7e8a-150">X</span><span class="sxs-lookup"><span data-stu-id="c7e8a-150">X</span></span>           | <span data-ttu-id="c7e8a-151">X</span><span class="sxs-lookup"><span data-stu-id="c7e8a-151">X</span></span>               | <span data-ttu-id="c7e8a-152">X</span><span class="sxs-lookup"><span data-stu-id="c7e8a-152">X</span></span>            |
| <span data-ttu-id="c7e8a-153">コマンド</span><span class="sxs-lookup"><span data-stu-id="c7e8a-153">Command</span></span>         | <span data-ttu-id="c7e8a-154">X</span><span class="sxs-lookup"><span data-stu-id="c7e8a-154">X</span></span>           | <span data-ttu-id="c7e8a-155">X</span><span class="sxs-lookup"><span data-stu-id="c7e8a-155">X</span></span>               | <span data-ttu-id="c7e8a-156">X</span><span class="sxs-lookup"><span data-stu-id="c7e8a-156">X</span></span>            |

## <a name="deployment-methods"></a><span data-ttu-id="c7e8a-157">展開方法</span><span class="sxs-lookup"><span data-stu-id="c7e8a-157">Deployment methods</span></span>

<span data-ttu-id="c7e8a-158">次からの各セクションでは、組織内のユーザーに Office アドインを配布する際に最も一般的に使用される展開方法についての追加情報を示します。</span><span class="sxs-lookup"><span data-stu-id="c7e8a-158">The following sections provide additional information about the deployment methods that are most commonly used to distribute Office Add-ins to users within an organization.</span></span>

<span data-ttu-id="c7e8a-159">エンド ユーザーがアドインを取得、挿入、実行する方法については、「[Office アドインの使用を開始する](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c7e8a-159">For information about how end users acquire, insert, and run add-ins, see [Start using your Office Add-in](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE).</span></span>

### <a name="centralized-deployment-via-the-office-365-admin-center"></a><span data-ttu-id="c7e8a-160">Office 365 管理センターからの一元展開</span><span class="sxs-lookup"><span data-stu-id="c7e8a-160">Centralized Deployment via the Office 365 admin center</span></span> 

<span data-ttu-id="c7e8a-p102">Office 365 管理センターを使用すると、管理者は組織内のユーザーとグループに Office アドインを簡単に展開できるようになります。管理センターを介して展開されたアドインは、ユーザーがすぐに Office アプリケーションで利用できるようになります。クライアントの構成は必要ありません。一元展開は、内部アドインの展開に使用することも、ISV が提供するアドインの展開に使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="c7e8a-p102">The Office 365 admin center makes it easy for an administrator to deploy Office Add-ins to users and groups in their organization. Add-ins deployed via the admin center are available to users in their Office applications right away, with no client configuration required. You can use Centralized Deployment to deploy internal add-ins as well as add-ins provided by ISVs.</span></span>

<span data-ttu-id="c7e8a-164">詳細については、「[Office 365 管理センターからの一元展開を使用した Office アドインの発行](centralized-deployment.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c7e8a-164">For more information, see [Publish Office Add-ins using Centralized Deployment via the Office 365 admin center](centralized-deployment.md).</span></span>

### <a name="sharepoint-catalog-deployment"></a><span data-ttu-id="c7e8a-165">SharePoint カタログの展開</span><span class="sxs-lookup"><span data-stu-id="c7e8a-165">SharePoint catalog deployment</span></span>

<span data-ttu-id="c7e8a-p103">SharePoint アドイン カタログは、Word、Excel、PowerPoint のアドインをホストするために作成できる特別なサイト コレクションです。SharePoint カタログは、マニフェストの `VersionOverrides` ノードに実装されている新しいアドイン機能 (アドイン コマンドを含む) をサポートしていないため、可能な場合は管理センター経由の一元展開を実行することをお勧めします。SharePoint カタログによって展開したアドイン コマンドは、既定では作業ウィンドウで開かれます。</span><span class="sxs-lookup"><span data-stu-id="c7e8a-p103">A SharePoint add-in catalog is a special site collection that you can create to host Word, Excel, and PowerPoint add-ins. Because SharePoint catalogs don't support new add-in features implemented in the `VersionOverrides` node of the manifest, including add-in commands, we recommend that you use Centralized Deployment via the admin center if possible. Add-in commands deployed via a SharePoint catalog open in a task pane by default.</span></span>

<span data-ttu-id="c7e8a-p104">オンプレミス環境でアドインを展開する場合は、SharePoint カタログを使用します。詳細については、「[作業ウィンドウ アドインとコンテンツ アドインを SharePoint カタログに発行する](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c7e8a-p104">If you are deploying add-ins in an on-premises environment, use a SharePoint catalog. For details, see [Publish task pane and content add-ins to a SharePoint catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span></span>

> [!NOTE]
> <span data-ttu-id="c7e8a-170">SharePoint カタログは、Office for Mac をサポートしません。</span><span class="sxs-lookup"><span data-stu-id="c7e8a-170">SharePoint catalogs do not support Office for Mac.</span></span> <span data-ttu-id="c7e8a-171">Office アドインを Mac クライアントに展開するには、そのアドインを [AppSource](/office/dev/store/submit-to-the-office-store) に提出する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c7e8a-171">To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](/office/dev/store/submit-to-the-office-store).</span></span>

### <a name="outlook-add-in-deployment"></a><span data-ttu-id="c7e8a-172">Outlook アドインの展開</span><span class="sxs-lookup"><span data-stu-id="c7e8a-172">Outlook add-in deployment</span></span>

<span data-ttu-id="c7e8a-173">Azure AD の ID サービスを使用しないオンプレミス環境およびオンライン環境では、Exchange サーバー経由で Outlook アドインを展開することができます。</span><span class="sxs-lookup"><span data-stu-id="c7e8a-173">For on-premises and online environments that do not use the Azure AD identity service, you can deploy Outlook add-ins via the Exchange server.</span></span>

<span data-ttu-id="c7e8a-174">Outlook アドインの展開には以下が必要です。</span><span class="sxs-lookup"><span data-stu-id="c7e8a-174">Outlook add-in deployment requires:</span></span>

- <span data-ttu-id="c7e8a-175">Office 365、Exchange Online、または Exchange Server 2013 以降</span><span class="sxs-lookup"><span data-stu-id="c7e8a-175">Office 365, Exchange Online, or Exchange Server 2013 or later</span></span>
- <span data-ttu-id="c7e8a-176">Outlook 2013 以降</span><span class="sxs-lookup"><span data-stu-id="c7e8a-176">Outlook 2013 or later</span></span>

<span data-ttu-id="c7e8a-p106">アドインをテナントに割り当てるには、Exchange 管理センターを使用して、ファイルまたは URL から直接マニフェストをアップロードするか、または AppSource からアドインを追加します。アドインを個々のユーザーに割り当てるには、Exchange PowerShell を使用する必要があります。詳細については、TechNet の「[組織の Outlook アドインをインストールまたは削除する](https://technet.microsoft.com/library/jj943752(v=exchg.150).aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c7e8a-p106">To assign add-ins to tenants, you use the Exchange admin center to upload a manifest directly, either from a file or a URL, or add an add-in from AppSource. To assign add-ins to individual users, you must use Exchange PowerShell. For details, see [Install or remove Outlook add-ins for your organization](https://technet.microsoft.com/library/jj943752(v=exchg.150).aspx) on TechNet.</span></span>

## <a name="see-also"></a><span data-ttu-id="c7e8a-180">関連項目</span><span class="sxs-lookup"><span data-stu-id="c7e8a-180">See also</span></span>

- [<span data-ttu-id="c7e8a-181">テスト用に Outlook アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="c7e8a-181">Sideload Outlook add-ins for testing</span></span>](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- <span data-ttu-id="c7e8a-182">[AppSource に提出する][AppSource]</span><span class="sxs-lookup"><span data-stu-id="c7e8a-182">[Submit to AppSource][AppSource]</span></span>
- [<span data-ttu-id="c7e8a-183">Office アドインの設計ガイドライン</span><span class="sxs-lookup"><span data-stu-id="c7e8a-183">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="c7e8a-184">効果的な AppSource 登録リストを作成する</span><span class="sxs-lookup"><span data-stu-id="c7e8a-184">Create effective AppSource listings</span></span>](/office/dev/store/create-effective-office-store-listings)
- [<span data-ttu-id="c7e8a-185">Office アドインでのユーザー エラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="c7e8a-185">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)

[AppSource]: https://docs.microsoft.com/office/dev/store/submit-to-the-office-store
[Office Add-in host and platform availability]: ../overview/office-add-in-availability

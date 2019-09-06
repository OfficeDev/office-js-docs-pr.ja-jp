---
title: Office アドインを展開し、発行する | Microsoft Docs
description: テスト目的またはユーザーに配布する目的で Office アドインを展開するための方法とオプション。
ms.date: 09/05/2019
localization_priority: Priority
ms.openlocfilehash: c47f8743edeed1fd366d948d781c97da1c97958a
ms.sourcegitcommit: d34aa0b282cc76ffff579da2a7945efd12fb7340
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/05/2019
ms.locfileid: "36769555"
---
# <a name="deploy-and-publish-your-office-add-in"></a><span data-ttu-id="bfc9b-103">Office アドインを展開し、発行する</span><span class="sxs-lookup"><span data-stu-id="bfc9b-103">Deploy and publish your Office Add-in</span></span>

<span data-ttu-id="bfc9b-104">さまざまな方法を利用し、テスト目的またはユーザーに配布する目的で、Office アドインを展開できます。</span><span class="sxs-lookup"><span data-stu-id="bfc9b-104">You can use one of several methods to deploy your Office Add-in for testing or distribution to users.</span></span>

|<span data-ttu-id="bfc9b-105">**メソッド**</span><span class="sxs-lookup"><span data-stu-id="bfc9b-105">**Method**</span></span>|<span data-ttu-id="bfc9b-106">**Use...**</span><span class="sxs-lookup"><span data-stu-id="bfc9b-106">**Use...**</span></span>|
|:---------|:------------|
|[<span data-ttu-id="bfc9b-107">サイドロード</span><span class="sxs-lookup"><span data-stu-id="bfc9b-107">Sideloading</span></span>](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing)|<span data-ttu-id="bfc9b-108">開発プロセスの一環として、Windows、iPad、Mac、またはブラウザーで実行するアドインをテストします。</span><span class="sxs-lookup"><span data-stu-id="bfc9b-108">As part of your development process, to test your add-in running on Windows, Office Online, iPad, or Mac.</span></span>|
|[<span data-ttu-id="bfc9b-109">一元展開</span><span class="sxs-lookup"><span data-stu-id="bfc9b-109">Centralized Deployment</span></span>](centralized-deployment.md)|<span data-ttu-id="bfc9b-110">クラウド環境またはハイブリッド環境で、Office 365 管理センターを使用して組織内のユーザーにアドインを配布します。</span><span class="sxs-lookup"><span data-stu-id="bfc9b-110">In a cloud or hybrid deployment, to distribute your add-in to users in your organization by using the Office 365 admin center.</span></span>|
|[<span data-ttu-id="bfc9b-111">SharePoint カタログ</span><span class="sxs-lookup"><span data-stu-id="bfc9b-111">SharePoint catalog</span></span>](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|<span data-ttu-id="bfc9b-112">オンプレミス環境で、組織内のユーザーにアドインを配布します。</span><span class="sxs-lookup"><span data-stu-id="bfc9b-112">In an on-premises environment, to distribute your add-in to users in your organization.</span></span>|
|[<span data-ttu-id="bfc9b-113">AppSource</span><span class="sxs-lookup"><span data-stu-id="bfc9b-113">AppSource</span></span>](/office/dev/store/submit-to-the-office-store)|<span data-ttu-id="bfc9b-114">ユーザーに配布する目的でアドインを公開します。</span><span class="sxs-lookup"><span data-stu-id="bfc9b-114">To distribute your add-in publicly to users.</span></span>|
|[<span data-ttu-id="bfc9b-115">Exchange サーバー</span><span class="sxs-lookup"><span data-stu-id="bfc9b-115">Exchange server</span></span>](#outlook-add-in-deployment)|<span data-ttu-id="bfc9b-116">オンプレミス環境またはオンライン環境で、ユーザーに Outlook アドインを配布します。</span><span class="sxs-lookup"><span data-stu-id="bfc9b-116">In an on-premises or online environment, to distribute Outlook add-ins to users.</span></span>|
|[<span data-ttu-id="bfc9b-117">ネットワーク共有</span><span class="sxs-lookup"><span data-stu-id="bfc9b-117">Network share</span></span>](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|<span data-ttu-id="bfc9b-118">アドインをホストさせようとしているネットワーク上の Windows コンピューターで、共有フォルダー カタログとして使用するフォルダーの親フォルダーまたはドライブ文字に移動します。</span><span class="sxs-lookup"><span data-stu-id="bfc9b-118">On a Windows computer on a network where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>|

> [!NOTE]
> <span data-ttu-id="bfc9b-p101">AppSource にアドインを[公開](../publish/publish.md)し、Office エクスペリエンスで利用できるようにする予定がある場合は、[AppSource の検証ポリシー](/office/dev/store/validation-policies)に準拠していることを確認してください。たとえば、検証に合格するには、定義したメソッドをサポートするすべてのプラットフォームでアドインが動作する必要があります (詳細については、[セクション 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) と [Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)のページを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="bfc9b-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="deployment-options-by-office-host"></a><span data-ttu-id="bfc9b-121">Office のホストごとの展開オプション</span><span class="sxs-lookup"><span data-stu-id="bfc9b-121">Deployment options by Office host</span></span>

<span data-ttu-id="bfc9b-122">選択可能な展開オプションは、対象の Office ホストや作成するアドインの種類によって異なります。</span><span class="sxs-lookup"><span data-stu-id="bfc9b-122">The deployment options that are available depend on the Office host that you're targeting and the type of add-in you create.</span></span>

### <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a><span data-ttu-id="bfc9b-123">Word、Excel、PowerPoint のアドインの展開オプション</span><span class="sxs-lookup"><span data-stu-id="bfc9b-123">Deployment options for Word, Excel, and PowerPoint add-ins</span></span>

| <span data-ttu-id="bfc9b-124">拡張点</span><span class="sxs-lookup"><span data-stu-id="bfc9b-124">Extension point</span></span> | <span data-ttu-id="bfc9b-125">サイドロード</span><span class="sxs-lookup"><span data-stu-id="bfc9b-125">Sideloading</span></span> | <span data-ttu-id="bfc9b-126">Office 365 管理センター</span><span class="sxs-lookup"><span data-stu-id="bfc9b-126">Office 365 admin center</span></span> |<span data-ttu-id="bfc9b-127">AppSource</span><span class="sxs-lookup"><span data-stu-id="bfc9b-127">AppSource</span></span>   | <span data-ttu-id="bfc9b-128">SharePoint カタログ\*</span><span class="sxs-lookup"><span data-stu-id="bfc9b-128">SharePoint catalog\*</span></span> |
|:----------------|:-----------:|:-----------------------:|:----------:|:--------------------:|
| <span data-ttu-id="bfc9b-129">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="bfc9b-129">Content</span></span>         | <span data-ttu-id="bfc9b-130">X</span><span class="sxs-lookup"><span data-stu-id="bfc9b-130">X</span></span>           | <span data-ttu-id="bfc9b-131">X</span><span class="sxs-lookup"><span data-stu-id="bfc9b-131">X</span></span>                       | <span data-ttu-id="bfc9b-132">X</span><span class="sxs-lookup"><span data-stu-id="bfc9b-132">X</span></span>          | <span data-ttu-id="bfc9b-133">X</span><span class="sxs-lookup"><span data-stu-id="bfc9b-133">X</span></span>                    |
| <span data-ttu-id="bfc9b-134">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc9b-134">Task pane</span></span>       | <span data-ttu-id="bfc9b-135">X</span><span class="sxs-lookup"><span data-stu-id="bfc9b-135">X</span></span>           | <span data-ttu-id="bfc9b-136">X</span><span class="sxs-lookup"><span data-stu-id="bfc9b-136">X</span></span>                       | <span data-ttu-id="bfc9b-137">X</span><span class="sxs-lookup"><span data-stu-id="bfc9b-137">X</span></span>          | <span data-ttu-id="bfc9b-138">X</span><span class="sxs-lookup"><span data-stu-id="bfc9b-138">X</span></span>                    |
| <span data-ttu-id="bfc9b-139">コマンド</span><span class="sxs-lookup"><span data-stu-id="bfc9b-139">Command</span></span>         | <span data-ttu-id="bfc9b-140">X</span><span class="sxs-lookup"><span data-stu-id="bfc9b-140">X</span></span>           | <span data-ttu-id="bfc9b-141">X</span><span class="sxs-lookup"><span data-stu-id="bfc9b-141">X</span></span>                       | <span data-ttu-id="bfc9b-142">X</span><span class="sxs-lookup"><span data-stu-id="bfc9b-142">X</span></span>          |                      |

<span data-ttu-id="bfc9b-143">&#42; SharePoint カタログは Office on Mac をサポートしません。</span><span class="sxs-lookup"><span data-stu-id="bfc9b-143">&#42; SharePoint catalogs do not support Office for Mac.</span></span>

### <a name="deployment-options-for-outlook-add-ins"></a><span data-ttu-id="bfc9b-144">Outlook アドインの展開オプション</span><span class="sxs-lookup"><span data-stu-id="bfc9b-144">Deployment options for Outlook add-ins</span></span>

| <span data-ttu-id="bfc9b-145">拡張点</span><span class="sxs-lookup"><span data-stu-id="bfc9b-145">Extension point</span></span> | <span data-ttu-id="bfc9b-146">サイドロード</span><span class="sxs-lookup"><span data-stu-id="bfc9b-146">Sideloading</span></span> | <span data-ttu-id="bfc9b-147">Exchange サーバー</span><span class="sxs-lookup"><span data-stu-id="bfc9b-147">Exchange server</span></span> | <span data-ttu-id="bfc9b-148">AppSource</span><span class="sxs-lookup"><span data-stu-id="bfc9b-148">AppSource</span></span>    |
|:----------------|:-----------:|:---------------:|:------------:|
| <span data-ttu-id="bfc9b-149">メール アプリ</span><span class="sxs-lookup"><span data-stu-id="bfc9b-149">Mail app</span></span>        | <span data-ttu-id="bfc9b-150">X</span><span class="sxs-lookup"><span data-stu-id="bfc9b-150">X</span></span>           | <span data-ttu-id="bfc9b-151">X</span><span class="sxs-lookup"><span data-stu-id="bfc9b-151">X</span></span>               | <span data-ttu-id="bfc9b-152">X</span><span class="sxs-lookup"><span data-stu-id="bfc9b-152">X</span></span>            |
| <span data-ttu-id="bfc9b-153">コマンド</span><span class="sxs-lookup"><span data-stu-id="bfc9b-153">Command</span></span>         | <span data-ttu-id="bfc9b-154">X</span><span class="sxs-lookup"><span data-stu-id="bfc9b-154">X</span></span>           | <span data-ttu-id="bfc9b-155">X</span><span class="sxs-lookup"><span data-stu-id="bfc9b-155">X</span></span>               | <span data-ttu-id="bfc9b-156">X</span><span class="sxs-lookup"><span data-stu-id="bfc9b-156">X</span></span>            |

## <a name="deployment-methods"></a><span data-ttu-id="bfc9b-157">展開方法</span><span class="sxs-lookup"><span data-stu-id="bfc9b-157">Deployment methods</span></span>

<span data-ttu-id="bfc9b-158">次からの各セクションでは、組織内のユーザーに Office アドインを配布する際に最も一般的に使用される展開方法についての追加情報を示します。</span><span class="sxs-lookup"><span data-stu-id="bfc9b-158">The following sections provide additional information about the deployment methods that are most commonly used to distribute Office Add-ins to users within an organization.</span></span>

<span data-ttu-id="bfc9b-159">エンド ユーザーがアドインを取得、挿入、実行する方法については、「[Office アドインの使用を開始する](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bfc9b-159">For information about how end users acquire, insert, and run add-ins, see [Start using your Office Add-in](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE).</span></span>

### <a name="centralized-deployment-via-the-office-365-admin-center"></a><span data-ttu-id="bfc9b-160">Office 365 管理センターからの一元展開</span><span class="sxs-lookup"><span data-stu-id="bfc9b-160">Centralized Deployment via the Office 365 admin center</span></span> 

<span data-ttu-id="bfc9b-p102">Office 365 管理センターを使用すると、管理者は組織内のユーザーとグループに Office アドインを簡単に展開できるようになります。管理センターを介して展開されたアドインは、ユーザーがすぐに Office アプリケーションで利用できるようになります。クライアントの構成は必要ありません。一元展開は、内部アドインの展開に使用することも、ISV が提供するアドインの展開に使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="bfc9b-p102">The Office 365 admin center makes it easy for an administrator to deploy Office Add-ins to users and groups in their organization. Add-ins deployed via the admin center are available to users in their Office applications right away, with no client configuration required. You can use Centralized Deployment to deploy internal add-ins as well as add-ins provided by ISVs.</span></span>

<span data-ttu-id="bfc9b-164">詳細については、「[Office 365 管理センターからの一元展開を使用した Office アドインの発行](centralized-deployment.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bfc9b-164">For more information, see [Publish Office Add-ins using Centralized Deployment via the Office 365 admin center](centralized-deployment.md).</span></span>

### <a name="sharepoint-app-catalog-deployment"></a><span data-ttu-id="bfc9b-165">SharePoint アプリ カタログの展開</span><span class="sxs-lookup"><span data-stu-id="bfc9b-165">SharePoint catalog deployment</span></span>

<span data-ttu-id="bfc9b-p103">SharePoint アプリ カタログは、Word、Excel、PowerPoint のアドインをホストするために作成できる特別なサイト コレクションです。SharePoint カタログは、マニフェストの `VersionOverrides` ノードに実装されている新しいアドイン機能 (アドイン コマンドを含む) をサポートしていないため、可能な場合は管理センター経由の一元展開を実行することをお勧めします。SharePoint カタログによって展開したアドイン コマンドは、既定では作業ウィンドウで開かれます。</span><span class="sxs-lookup"><span data-stu-id="bfc9b-p103">A SharePoint add-in catalog is a special site collection that you can create to host Word, Excel, and PowerPoint add-ins. Because SharePoint catalogs don't support new add-in features implemented in the `VersionOverrides` node of the manifest, including add-in commands, we recommend that you use Centralized Deployment via the admin center if possible. Add-in commands deployed via a SharePoint catalog open in a task pane by default.</span></span>

<span data-ttu-id="bfc9b-p104">オンプレミス環境でアドインを展開する場合は、SharePoint カタログを使用します。詳細については、「[作業ウィンドウ アドインとコンテンツ アドインを SharePoint カタログに発行する](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bfc9b-p104">If you are deploying add-ins in an on-premises environment, use a SharePoint catalog. For details, see [Publish task pane and content add-ins to a SharePoint catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span></span>

> [!NOTE]
> <span data-ttu-id="bfc9b-170">SharePoint カタログは Office on Mac をサポートしません。</span><span class="sxs-lookup"><span data-stu-id="bfc9b-170">SharePoint catalogs do not support Office for Mac.</span></span> <span data-ttu-id="bfc9b-171">Office アドインを Mac クライアントに展開するには、そのアドインを [AppSource](/office/dev/store/submit-to-the-office-store) に提出する必要があります。</span><span class="sxs-lookup"><span data-stu-id="bfc9b-171">To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](/office/dev/store/submit-to-the-office-store).</span></span>

### <a name="outlook-add-in-deployment"></a><span data-ttu-id="bfc9b-172">Outlook アドインの展開</span><span class="sxs-lookup"><span data-stu-id="bfc9b-172">Outlook add-in deployment</span></span>

<span data-ttu-id="bfc9b-173">Azure AD の ID サービスを使用しないオンプレミス環境およびオンライン環境では、Exchange サーバー経由で Outlook アドインを展開することができます。</span><span class="sxs-lookup"><span data-stu-id="bfc9b-173">For on-premises and online environments that do not use the Azure AD identity service, you can deploy Outlook add-ins via the Exchange server.</span></span>

<span data-ttu-id="bfc9b-174">Outlook アドインの展開には以下が必要です。</span><span class="sxs-lookup"><span data-stu-id="bfc9b-174">Outlook add-in deployment requires:</span></span>

- <span data-ttu-id="bfc9b-175">Office 365、Exchange Online、または Exchange Server 2013 以降</span><span class="sxs-lookup"><span data-stu-id="bfc9b-175">Office 365, Exchange Online, or Exchange Server 2013 or later</span></span>
- <span data-ttu-id="bfc9b-176">Outlook 2013 以降</span><span class="sxs-lookup"><span data-stu-id="bfc9b-176">Outlook 2013 or later</span></span>

<span data-ttu-id="bfc9b-p106">アドインをテナントに割り当てるには、Exchange 管理センターを使用して、ファイルまたは URL から直接マニフェストをアップロードするか、または AppSource からアドインを追加します。アドインを個々のユーザーに割り当てるには、Exchange PowerShell を使用する必要があります。詳細については、TechNet の「[組織の Outlook アドインをインストールまたは削除する](https://technet.microsoft.com/library/jj943752(v=exchg.150).aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bfc9b-p106">To assign add-ins to tenants, you use the Exchange admin center to upload a manifest directly, either from a file or a URL, or add an add-in from AppSource. To assign add-ins to individual users, you must use Exchange PowerShell. For details, see [Install or remove Outlook add-ins for your organization](https://technet.microsoft.com/library/jj943752(v=exchg.150).aspx) on TechNet.</span></span>

## <a name="see-also"></a><span data-ttu-id="bfc9b-180">関連項目</span><span class="sxs-lookup"><span data-stu-id="bfc9b-180">See also</span></span>

- [<span data-ttu-id="bfc9b-181">テスト用に Outlook アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="bfc9b-181">Sideload Outlook add-ins for testing</span></span>](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- <span data-ttu-id="bfc9b-182">[AppSource に提出する][AppSource]</span><span class="sxs-lookup"><span data-stu-id="bfc9b-182">[Submit to AppSource][AppSource]</span></span>
- [<span data-ttu-id="bfc9b-183">Office アドインの設計ガイドライン</span><span class="sxs-lookup"><span data-stu-id="bfc9b-183">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="bfc9b-184">効果的な AppSource 登録リストを作成する</span><span class="sxs-lookup"><span data-stu-id="bfc9b-184">Create effective AppSource listings</span></span>](/office/dev/store/create-effective-office-store-listings)
- [<span data-ttu-id="bfc9b-185">Office アドインでのユーザー エラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="bfc9b-185">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)

[AppSource]: /office/dev/store/submit-to-the-office-store
[Office Add-in host and platform availability]: ../overview/office-add-in-availability

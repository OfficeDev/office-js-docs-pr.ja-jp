---
title: Outlook アドインに関するプライバシー、アクセス許可、セキュリティ
description: Outlook アドインで、プライバシー、アクセス許可、セキュリティを管理する方法について説明します。
ms.date: 08/03/2020
localization_priority: Priority
ms.openlocfilehash: 9807cbb2346d6fc067f3894c9f5d265f83dccdc3
ms.sourcegitcommit: a3b743598025466bad19177e0ba9ca94ea66d490
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/04/2020
ms.locfileid: "46547536"
---
# <a name="privacy-permissions-and-security-for-outlook-add-ins"></a><span data-ttu-id="71d6a-103">Outlook アドインに関するプライバシー、アクセス許可、セキュリティ</span><span class="sxs-lookup"><span data-stu-id="71d6a-103">Privacy, permissions, and security for Outlook add-ins</span></span>

<span data-ttu-id="71d6a-104">エンドユーザー、開発者、および管理者は、Outlook アドインのセキュリティ モデルの階層化されたアクセス許可レベルを使用して、プライバシーとパフォーマンスを制御することができます。</span><span class="sxs-lookup"><span data-stu-id="71d6a-104">End users, developers, and administrators can use the tiered permission levels of the security model for Outlook add-ins to control privacy and performance.</span></span>

<span data-ttu-id="71d6a-105">この記事では、Outlook アドインで要求可能なアクセス許可について説明し、次のような観点からセキュリティ モデルを調べます。</span><span class="sxs-lookup"><span data-stu-id="71d6a-105">This article describes the possible permissions that Outlook add-ins can request, and examines the security model from the following perspectives:</span></span>

- <span data-ttu-id="71d6a-106">**AppSource**: アドインの整合性</span><span class="sxs-lookup"><span data-stu-id="71d6a-106">**AppSource**: add-in integrity</span></span>
    
- <span data-ttu-id="71d6a-107">**エンド ユーザー**: プライバシーとパフォーマンスの問題</span><span class="sxs-lookup"><span data-stu-id="71d6a-107">**End-users**: privacy and performance concerns</span></span>
    
- <span data-ttu-id="71d6a-108">**開発者**: アクセス許可の選択とリソース使用量の制限</span><span class="sxs-lookup"><span data-stu-id="71d6a-108">**Developers**: permissions choices and resource usage limits</span></span>
    
- <span data-ttu-id="71d6a-109">**管理者**: パフォーマンスのしきい値を設定する特権</span><span class="sxs-lookup"><span data-stu-id="71d6a-109">**Administrators**: privileges to set performance thresholds</span></span>
    

## <a name="permissions-model"></a><span data-ttu-id="71d6a-110">アクセス許可モデル</span><span class="sxs-lookup"><span data-stu-id="71d6a-110">Permissions model</span></span>

<span data-ttu-id="71d6a-p101">お客様のアドインのセキュリティの認知度がアドインの導入に影響する可能性があるため、Outlook アドインのセキュリティは階層化されたアクセス許可モデルに依存します。Outlook アドインは、アドインがお客様のメールボックス データに対して実行可能なアクセスとアクションを特定した上で、必要なアクセス許可レベルを開示します。</span><span class="sxs-lookup"><span data-stu-id="71d6a-p101">Because customers' perception of add-in security can affect add-in adoption, Outlook add-in security relies on a tiered permissions model. An Outlook add-in would disclose the level of permissions it needs, identifying the possible access and actions that the add-in can make on the customer's mailbox data.</span></span> 

<span data-ttu-id="71d6a-113">マニフェスト スキーマのバージョン 1.1 には、4 つのレベルのアクセス許可が含まれています。</span><span class="sxs-lookup"><span data-stu-id="71d6a-113">Manifest schema version 1.1 includes four levels of permissions.</span></span> 


<span data-ttu-id="71d6a-114">**表 1.アドインのアクセス許可レベル**</span><span class="sxs-lookup"><span data-stu-id="71d6a-114">**Table 1. Add-in permission levels**</span></span>

|<span data-ttu-id="71d6a-115">**アクセス許可レベル**</span><span class="sxs-lookup"><span data-stu-id="71d6a-115">**Permission level**</span></span>|<span data-ttu-id="71d6a-116">**Outlook アドインのマニフェストの値**</span><span class="sxs-lookup"><span data-stu-id="71d6a-116">**Value in Outlook add-in manifest**</span></span>|
|:-----|:-----|
|<span data-ttu-id="71d6a-117">Restricted</span><span class="sxs-lookup"><span data-stu-id="71d6a-117">Restricted</span></span>|<span data-ttu-id="71d6a-118">Restricted</span><span class="sxs-lookup"><span data-stu-id="71d6a-118">Restricted</span></span>|
|<span data-ttu-id="71d6a-119">アイテムの読み取り</span><span class="sxs-lookup"><span data-stu-id="71d6a-119">Read item</span></span>|<span data-ttu-id="71d6a-120">ReadItem</span><span class="sxs-lookup"><span data-stu-id="71d6a-120">ReadItem</span></span>|
|<span data-ttu-id="71d6a-121">アイテムの読み取り/書き込み</span><span class="sxs-lookup"><span data-stu-id="71d6a-121">Read/write item</span></span>|<span data-ttu-id="71d6a-122">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="71d6a-122">ReadWriteItem</span></span>|
|<span data-ttu-id="71d6a-123">メールボックスの読み取り/書き込み</span><span class="sxs-lookup"><span data-stu-id="71d6a-123">Read/write mailbox</span></span>|<span data-ttu-id="71d6a-124">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="71d6a-124">ReadWriteMailbox</span></span>|

<span data-ttu-id="71d6a-125">アクセス許可の 4 つのレベルは累積的です。**メールボックス読み取り/書き込み**アクセス許可には**アイテム読み取り/書き込み**、**アイテム読み取り**、および**制限付き**が含まれており、**アイテム読み取り/書き込み**には**アイテム読み取り**と**制限付き**が含まれており、また**アイテム読み取り**アクセス許可には**制限付き**が含まれています。</span><span class="sxs-lookup"><span data-stu-id="71d6a-125">The four levels of permissions are cumulative: the **read/write mailbox** permission includes the permissions of **read/write item**, **read item** and **restricted**, **read/write item** includes **read item** and **restricted**, and the **read item** permission includes **restricted**.</span></span> 

<span data-ttu-id="71d6a-126">次の図は、アクセス許可の 4 つのレベルを示しています。また、各層でエンド ユーザー、開発者、および管理者に提供される機能が示されています。</span><span class="sxs-lookup"><span data-stu-id="71d6a-126">The following figure shows the four levels of permissions and describes the capabilities offered to the end user, developer, and administrator by each tier.</span></span> <span data-ttu-id="71d6a-127">これらのアクセス許可の詳細については、「[エンド ユーザー: プライバシーとパフォーマンスについて](#end-users-privacy-and-performance-concerns)」、「[開発者: アクセス許可の選択とリソース使用の制限](#developers-permission-choices-and-resource-usage-limits)」、および「[Outlook アドインのアクセス許可について](understanding-outlook-add-in-permissions.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="71d6a-127">For more information about these permissions, see [End users: privacy and performance concerns](#end-users-privacy-and-performance-concerns), [Developers: permission choices and resource usage limits](#developers-permission-choices-and-resource-usage-limits), and [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md).</span></span> 


<span data-ttu-id="71d6a-128">**4 層のアクセス許可モデルとエンド ユーザー、開発者、および管理者の関連性**</span><span class="sxs-lookup"><span data-stu-id="71d6a-128">**Relating the four-tier permission model to the end user, developer, and administrator**</span></span>

![メール アプリ スキーマ v1.1 の 4 層アクセス許可モデル](../images/add-in-permission-tiers.png)


## <a name="appsource-add-in-integrity"></a><span data-ttu-id="71d6a-130">AppSource: アドインの整合性</span><span class="sxs-lookup"><span data-stu-id="71d6a-130">AppSource: add-in integrity</span></span>

<span data-ttu-id="71d6a-131">[AppSource](https://appsource.microsoft.com) は、エンド ユーザーと管理者がインストールできるアドインをホストします。</span><span class="sxs-lookup"><span data-stu-id="71d6a-131">[AppSource](https://appsource.microsoft.com) hosts add-ins that can be installed by end users and administrators.</span></span> <span data-ttu-id="71d6a-132">AppSource は、これらの Outlook アドインの整合性を維持するために次の手段を適用します。</span><span class="sxs-lookup"><span data-stu-id="71d6a-132">AppSource enforces the following measures to maintain the integrity of these Outlook add-ins:</span></span>

- <span data-ttu-id="71d6a-133">アドインのホスト サーバーは必ず Secure Socket Layer (SSL) を使用して通信する必要があります。</span><span class="sxs-lookup"><span data-stu-id="71d6a-133">Requires the host server of an add-in to always use Secure Socket Layer (SSL) to communicate.</span></span>
    
- <span data-ttu-id="71d6a-134">開発者はアドインを提出する際に、ID の証明、契約上の合意、および法規制に準拠したプライバシー ポリシーを提供する必要があります。</span><span class="sxs-lookup"><span data-stu-id="71d6a-134">Requires a developer to provide proof of identity, a contractual agreement, and a compliant privacy policy to submit add-ins.</span></span> 
    
- <span data-ttu-id="71d6a-135">アドインを読み取り専用モードでアーカイブします。</span><span class="sxs-lookup"><span data-stu-id="71d6a-135">Archives add-ins in read-only mode.</span></span>
    
- <span data-ttu-id="71d6a-136">使用可能なアドインに対するユーザーレビュー システムをサポートしてコミュニティの自己管理を促します。</span><span class="sxs-lookup"><span data-stu-id="71d6a-136">Supports a user-review system for available add-ins to promote a self-policing community.</span></span>
    

## <a name="end-users-privacy-and-performance-concerns"></a><span data-ttu-id="71d6a-137">エンド ユーザー: プライバシーとパフォーマンスの問題</span><span class="sxs-lookup"><span data-stu-id="71d6a-137">End users: privacy and performance concerns</span></span>

<span data-ttu-id="71d6a-138">セキュリティ モデルによって、エンド ユーザーのセキュリティ、プライバシー、およびパフォーマンスの問題に次のような方法で対処できます。</span><span class="sxs-lookup"><span data-stu-id="71d6a-138">The security model addresses security, privacy, and performance concerns of end users in the following ways:</span></span>

- <span data-ttu-id="71d6a-139">Outlook の IRM (Information Rights Management) で保護されているエンド ユーザーのメッセージは、Outlook アドインとやり取りしません。</span><span class="sxs-lookup"><span data-stu-id="71d6a-139">End user's messages that are protected by Outlook's Information Rights Management (IRM) do not interact with Outlook add-ins.</span></span>
    
  > [!IMPORTANT]
  > <span data-ttu-id="71d6a-140">Windows の Outlook ビルド13120.1000 から、アドインは IRM で保護されたアイテムに対して有効になるようになりました。</span><span class="sxs-lookup"><span data-stu-id="71d6a-140">Starting with Outlook build 13120.1000 on Windows, add-ins can now activate on items protected by IRM.</span></span> <span data-ttu-id="71d6a-141">この機能のプレビューの詳細については、「[Information Rights Management (IRM) で保護されているアイテムのアドインのアクティブ化](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="71d6a-141">For more information about this feature in preview, see [Add-in activation on items protected by Information Rights Management (IRM)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm).</span></span>

- <span data-ttu-id="71d6a-142">AppSource からアドインをインストールする前に、エンド ユーザーは、そのアドインが自分のデータに対して実行可能なアクセスとアクションを確認して、先に進むことを明示的に確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="71d6a-142">Before installing an add-in from AppSource, end users can see the access and actions that the add-in can make on their data and must explicitly confirm to proceed.</span></span> <span data-ttu-id="71d6a-143">Outlook アドインは、ユーザーまたは管理者による手動検証なしでクライアント コンピューター上に自動的にインストールされることはありません。</span><span class="sxs-lookup"><span data-stu-id="71d6a-143">No Outlook add-in is automatically pushed onto a client computer without manual validation by the user or administrator.</span></span>
    
- <span data-ttu-id="71d6a-p106">
            \*\*制限付き\*\*のアクセス許可を与えると、Outlook アドインは現在のアイテムでのみ制限付きでアクセスできるようになります。\*\*アイテムの読み取り\*\*のアクセス許可を与えると、Outlook アドインは送信者と受信者の名前やメール アドレスなど、個人を特定できる情報に現在のアイテムでのみアクセスできるようになります。</span><span class="sxs-lookup"><span data-stu-id="71d6a-p106">Granting the **restricted** permission allows the Outlook add-in to have limited access on only the current item. Granting the **read item** permission allows the Outlook add-in to access personal identifiable information, such as sender and recipient names and email addresses, on only the current item,.</span></span>
    
- <span data-ttu-id="71d6a-p107">エンド ユーザーは、自分だけが使用する Outlook アドインをインストールできます。組織に影響を与える Outlook アドインは管理者がインストールします。</span><span class="sxs-lookup"><span data-stu-id="71d6a-p107">An end user can install an Outlook add-in for only himself or herself. Outlook add-ins that affect an organization are installed by an administrator.</span></span>
    
- <span data-ttu-id="71d6a-148">エンド ユーザーは、ユーザーのセキュリティ リスクを最小限に抑えながら、ユーザーにとって魅力的な状況依存のシナリオを実現する Outlook アドインをインストールできます。</span><span class="sxs-lookup"><span data-stu-id="71d6a-148">End users can install Outlook add-ins that enable context-sensitive scenarios that are compelling to users while minimizing the users' security risks.</span></span>
    
- <span data-ttu-id="71d6a-149">インストールされた Outlook アドインのマニフェスト ファイルは、ユーザーの電子メール アカウントに安全に保管されます。</span><span class="sxs-lookup"><span data-stu-id="71d6a-149">Manifest files of installed Outlook add-ins are secured in the user's email account.</span></span>
    
- <span data-ttu-id="71d6a-150">Office アドインをホストするサーバーと通信するデータは、Secure Socket Layer (SSL) プロトコルで常に暗号化されます。</span><span class="sxs-lookup"><span data-stu-id="71d6a-150">Data communicated with servers hosting Office Add-ins is always encrypted according to the Secure Socket Layer (SSL) protocol.</span></span>
    
- <span data-ttu-id="71d6a-151">Outlook リッチ クライアントのみ: Outlook リッチ クライアントは、インストールされた Outlook アドインのパフォーマンスを監視し、ガバナンス制御を実施し、次の領域で制限を超えている Outlook アドインを無効にします。</span><span class="sxs-lookup"><span data-stu-id="71d6a-151">Applicable to only the Outlook rich clients: The Outlook rich clients monitor the performance of installed Outlook add-ins, exercise governance control, and disable those Outlook add-ins that exceed limits in the following areas:</span></span>
    
  - <span data-ttu-id="71d6a-152">アクティブ化までの応答時間</span><span class="sxs-lookup"><span data-stu-id="71d6a-152">Response time to activate</span></span>
    
  - <span data-ttu-id="71d6a-153">アクティブ化または再アクティブ化に失敗した回数</span><span class="sxs-lookup"><span data-stu-id="71d6a-153">Number of failures to activate or reactivate</span></span>
    
  - <span data-ttu-id="71d6a-154">メモリ使用量</span><span class="sxs-lookup"><span data-stu-id="71d6a-154">Memory usage</span></span>
    
  - <span data-ttu-id="71d6a-155">CPU 使用率</span><span class="sxs-lookup"><span data-stu-id="71d6a-155">CPU usage</span></span>  

  <span data-ttu-id="71d6a-p108">ガバナンスはサービス拒否攻撃を阻止し、アドインのパフォーマンスを適度なレベルに維持します。エンド ユーザーには、このようなガバナンス制御に基づいて、Outlook リッチ クライアントが該当の Outlook アドインを無効にしたという通知がビジネス バーに表示されます。</span><span class="sxs-lookup"><span data-stu-id="71d6a-p108">Governance deters denial-of-service attacks and maintains add-in performance at a reasonable level. The Business Bar alerts end users about Outlook add-ins that the Outlook rich client has disabled based on such governance control.</span></span>

- <span data-ttu-id="71d6a-158">エンド ユーザーは、いつでも Exchange 管理センターで、インストールした Outlook アドインから要求されたアクセス許可を確認したり、Outlook アドインを無効にしたり、その後で有効にしたりできます。</span><span class="sxs-lookup"><span data-stu-id="71d6a-158">At any time, end users can verify the permissions requested by installed Outlook add-ins, and disable or subsequently enable any Outlook add-in in the Exchange Admin Center.</span></span>


## <a name="developers-permission-choices-and-resource-usage-limits"></a><span data-ttu-id="71d6a-159">開発者: アクセス許可の選択とリソース使用量の制限</span><span class="sxs-lookup"><span data-stu-id="71d6a-159">Developers: permission choices and resource usage limits</span></span>

<span data-ttu-id="71d6a-160">開発者は、セキュリティ モデルで規定されたきめ細かいレベルのアクセス許可を選択し、厳密なパフォーマンス ガイドラインを守る必要があります。</span><span class="sxs-lookup"><span data-stu-id="71d6a-160">The security model provides developers granular levels of permissions to choose from, and strict performance guidelines to observe.</span></span>

### <a name="tiered-permissions-increases-transparency"></a><span data-ttu-id="71d6a-161">階層化された許可で透過性が向上</span><span class="sxs-lookup"><span data-stu-id="71d6a-161">Tiered permissions increases transparency</span></span>

<span data-ttu-id="71d6a-162">開発者は階層化された許可モデルに従うことにより、透明性を提供しつつ、アドインがデータとメールボックスに対して実行可能なアクションに対するユーザーの懸念を緩和し、アドインの導入を間接的に促進できます。</span><span class="sxs-lookup"><span data-stu-id="71d6a-162">Developers should follow the tiered permissions model to provide transparency and alleviate users' concern about what add-ins can do to their data and mailbox, indirectly promoting add-in adoption:</span></span>

- <span data-ttu-id="71d6a-163">開発者は、Outlook アドインがアクティブ化される方法、およびメール アドインがアイテムの特定のプロパティを読み書きする必要性や、アイテムを作成および送信する必要性に基づいて、Outlook アドインの適切なレベルのアクセス許可を要求します。</span><span class="sxs-lookup"><span data-stu-id="71d6a-163">Developers request an appropriate level of permission for an Outlook add-in, based on how the Outlook add-in should be activated, and its need to read or write certain properties of an item, or to create and send an item.</span></span>

- <span data-ttu-id="71d6a-164">開発者は、Outlook アドインのマニフェストの [Permissions](../reference/manifest/permissions.md) 要素を使用して、**Restricted**、**ReadItem**、**ReadWriteItem** または **ReadWriteMailbox** の値を必要に応じて割り当ててアクセス許可を要求します。</span><span class="sxs-lookup"><span data-stu-id="71d6a-164">Developers request permission by using the [Permissions](../reference/manifest/permissions.md) element in the manifest of the Outlook add-in, by assigning a value of **Restricted**, **ReadItem**, **ReadWriteItem** or **ReadWriteMailbox**, as appropriate.</span></span>

  > [!NOTE]
  > <span data-ttu-id="71d6a-165">**ReadWriteItem** のアクセス許可は、マニフェスト スキーマ v1.1 以降で利用できます。</span><span class="sxs-lookup"><span data-stu-id="71d6a-165">Note that the **ReadWriteItem** permission is available starting in manifest schema v1.1.</span></span>

  <span data-ttu-id="71d6a-166">次の例では、**アイテムの読み取り**のアクセス許可を要求しています。</span><span class="sxs-lookup"><span data-stu-id="71d6a-166">The following example requests the **read item** permission.</span></span>

  ```XML
    <Permissions>ReadItem</Permissions>
  ```

- <span data-ttu-id="71d6a-167">特定の種類の Outlook アイテム (予定やメッセージ)、またはアイテムの件名や本文から抽出された特定のエンティティ (電話番号、住所、URL) に対して Outlook アドインをアクティブ化する場合、開発者は**制限付き**のアクセス許可を要求できます。</span><span class="sxs-lookup"><span data-stu-id="71d6a-167">Developers can request the **restricted** permission if the Outlook add-in activates on a specific type of Outlook items (appointment or message), or on specific extracted entities (phone number, address, URL) being present in the item's subject or body.</span></span> <span data-ttu-id="71d6a-168">たとえば、次のルールは、現在のメッセージの件名または本文に電話番号、郵送先住所、URL の 3 つのエンティティのうち 1 つ以上のエンティティが見つかった場合に Outlook アドインをアクティブ化します。</span><span class="sxs-lookup"><span data-stu-id="71d6a-168">For example, the following rule activates the Outlook add-in if one or more of three entities - phone number, postal address, or URL - are found in the subject or body of the current message.</span></span>
    
  ```XML
    <Permissions>Restricted</Permissions>
        <Rule xsi:type="RuleCollection" Mode="And">
        <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
        <Rule xsi:type="RuleCollection" Mode="Or">
            <Rule xsi:type="ItemHasKnownEntity" EntityType="PhoneNumber" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
        </Rule>
    </Rule>
  ```

- <span data-ttu-id="71d6a-169">Outlook アドインで、現在のアイテムの既定の抽出されたエンティティ以外のプロパティを読み取る必要がある場合や、現在のアイテムにアドインが設定するカスタム プロパティを書き込む必要がある場合に、その他のアイテムに対する読み取りや書き込み、またはユーザーのメールボックスのメッセージの作成や送信が不要な場合、開発者は**アイテムの読み取り**のアクセス許可を要求します。</span><span class="sxs-lookup"><span data-stu-id="71d6a-169">Developers should request the **read item** permission if the Outlook add-in needs to read properties of the current item other than the default extracted entities, or write custom properties set by the add-in on the current item, but does not require reading or writing to other items, or creating or sending a message in the user's mailbox.</span></span> <span data-ttu-id="71d6a-170">たとえば、Outlook アドインでアイテムの件名または本文に含まれる会議開催の提案、タスクの提案、メール アドレス、連絡先名などのエンティティを検索する必要がある場合や、アクティブ化に正規表現を使用する必要がある場合は、**アイテムの読み取り**のアクセス許可を要求します。</span><span class="sxs-lookup"><span data-stu-id="71d6a-170">For example, a developer should request **read item** permission if an Outlook add-in needs to look for an entity like a meeting suggestion, task suggestion, email address, or contact name in the item's subject or body, or uses a regular expression to activate.</span></span>

- <span data-ttu-id="71d6a-171">Outlook アドインが新規作成アイテムのプロパティ (受信者名、メールアドレス、本文、件名など) を書き込む必要がある場合、またはアイテムの添付ファイルを追加または削除する必要がある場合、開発者は**アイテムの読み取り/書き込み**許可を要求します。</span><span class="sxs-lookup"><span data-stu-id="71d6a-171">Developers should request the **read/write item** permission if the Outlook add-in needs to write to properties of the composed item, such as recipient names, email addresses, body, and subject, or needs to add or remove item attachments.</span></span>

- <span data-ttu-id="71d6a-172">開発者は、Outlook アドインで [mailbox.makeEWSRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッドを使用して次のいずれか 1 つ以上の処理を実行する必要がある場合にのみ、**メールボックスの読み取り/書き込み**のアクセス許可を要求します。</span><span class="sxs-lookup"><span data-stu-id="71d6a-172">Developers request the **read/write mailbox** permission only if the Outlook add-in needs to do one or more of the following actions by using the [mailbox.makeEWSRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method:</span></span>

  - <span data-ttu-id="71d6a-173">メールボックスのアイテムのプロパティに対する読み取りまたは書き込み。</span><span class="sxs-lookup"><span data-stu-id="71d6a-173">Read or write to properties of items in the mailbox.</span></span>
  - <span data-ttu-id="71d6a-174">メールボックスのアイテムの作成、読み取り、書き込み、または送信。</span><span class="sxs-lookup"><span data-stu-id="71d6a-174">Create, read, write, or send items in the mailbox.</span></span>
  - <span data-ttu-id="71d6a-175">メールボックスのフォルダーの作成、読み取り、または書き込み。</span><span class="sxs-lookup"><span data-stu-id="71d6a-175">Create, read, or write to folders in the mailbox.</span></span>


### <a name="resource-usage-tuning"></a><span data-ttu-id="71d6a-176">リソース使用量の調整</span><span class="sxs-lookup"><span data-stu-id="71d6a-176">Resource usage tuning</span></span>

<span data-ttu-id="71d6a-p111">パフォーマンスの良くないアドインがホストのサービスを拒否する事態を減らすため、開発者はアクティブ化におけるリソース使用量の限度を意識し、開発ワークフローにパフォーマンスの調整を組み込む必要があります。また、「 [Outlook アドインのアクティブ化と JavaScript API の制限](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)」に記載するとおり、アクティブ化ルールの設計ガイドラインに従うことをお勧めします。Outlook アドインを Outlook リッチ クライアント上で実行する予定がある場合、開発者はアドインがリソース使用量の制限内で動作することを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="71d6a-p111">Developers should be aware of resource usage limits for activation, incorporate performance tuning in their development workflow, so as to reduce the chance of a poorly performing add-in denying service of the host. Developers should follow the guidelines in designing activation rules as described in [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md). If an Outlook add-in is intended to run on an Outlook rich client, then developers should verify that the add-in performs within the resource usage limits.</span></span>


### <a name="other-measures-to-promote-user-security"></a><span data-ttu-id="71d6a-179">ユーザーのセキュリティを高めるその他の方法</span><span class="sxs-lookup"><span data-stu-id="71d6a-179">Other measures to promote user security</span></span>

<span data-ttu-id="71d6a-180">開発者は、以下の点についても意識し、計画する必要があります。</span><span class="sxs-lookup"><span data-stu-id="71d6a-180">Developers should be aware of and plan for the following as well:</span></span>

- <span data-ttu-id="71d6a-181">ActiveX コントロールはサポートされていないため、アドインで ActiveX コントロールを使用することはできません。</span><span class="sxs-lookup"><span data-stu-id="71d6a-181">Developers cannot use ActiveX controls in add-ins because they are not supported.</span></span>
    
- <span data-ttu-id="71d6a-182">開発者は AppSource に Outlook アドインを提出する際に、次の作業を実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="71d6a-182">Developers should do the following when submitting an Outlook add-in to AppSource:</span></span>
    
  - <span data-ttu-id="71d6a-183">ID の証明として Extended Validation (EV) SSL 証明書を生成する。</span><span class="sxs-lookup"><span data-stu-id="71d6a-183">Produce an Extended Validation (EV) SSL certificate as a proof of identity.</span></span>
    
  - <span data-ttu-id="71d6a-184">SSL をサポートする Web サーバーで、提出するアドインをホストする。</span><span class="sxs-lookup"><span data-stu-id="71d6a-184">Host the add-in they are submitting on a web server that supports SSL.</span></span>
    
  - <span data-ttu-id="71d6a-185">準拠したプライバシー ポリシーを生成する。</span><span class="sxs-lookup"><span data-stu-id="71d6a-185">Produce a compliant privacy policy.</span></span>
    
  - <span data-ttu-id="71d6a-186">アドインの提出時に契約合意書に署名する。</span><span class="sxs-lookup"><span data-stu-id="71d6a-186">Be ready to sign a contractual agreement upon submitting the add-in.</span></span>
    

## <a name="administrators-privileges"></a><span data-ttu-id="71d6a-187">管理者: 特権</span><span class="sxs-lookup"><span data-stu-id="71d6a-187">Administrators: privileges</span></span>

<span data-ttu-id="71d6a-188">セキュリティ モデルによって、管理者に次の権利と責任が与えられます。</span><span class="sxs-lookup"><span data-stu-id="71d6a-188">The security model provides the following rights and responsibilities to administrators:</span></span>

- <span data-ttu-id="71d6a-189">AppSource のアドインを含めて、エンド ユーザーが Outlook アドインをインストールできないようにすることができます。</span><span class="sxs-lookup"><span data-stu-id="71d6a-189">Can prevent end users from installing any Outlook add-in, including add-ins from AppSource.</span></span>
    
- <span data-ttu-id="71d6a-190">Exchange 管理センターで Outlook アドインを無効または有効にできます。</span><span class="sxs-lookup"><span data-stu-id="71d6a-190">Can disable or enable any Outlook add-in on the Exchange Admin Center.</span></span>
    
- <span data-ttu-id="71d6a-191">Windows 版 Outlook のみ: GPO レジストリ設定を使用して、パフォーマンスのしきい値の設定を無効にすることができます。</span><span class="sxs-lookup"><span data-stu-id="71d6a-191">Applicable to only Outlook on Windows: Can override performance threshold settings by GPO registry settings.</span></span>
    


## <a name="see-also"></a><span data-ttu-id="71d6a-192">関連項目</span><span class="sxs-lookup"><span data-stu-id="71d6a-192">See also</span></span>

- [<span data-ttu-id="71d6a-193">Office アドインのプライバシーとセキュリティ</span><span class="sxs-lookup"><span data-stu-id="71d6a-193">Privacy and security for Office Add-ins</span></span>](../develop/privacy-and-security.md)    
- [<span data-ttu-id="71d6a-194">Outlook アドインの API</span><span class="sxs-lookup"><span data-stu-id="71d6a-194">Outlook add-in APIs</span></span>](apis.md)    
- [<span data-ttu-id="71d6a-195">Outlook アドインのアクティブ化と JavaScript API の制限</span><span class="sxs-lookup"><span data-stu-id="71d6a-195">Limits for activation and JavaScript API for Outlook add-ins</span></span>](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
    
